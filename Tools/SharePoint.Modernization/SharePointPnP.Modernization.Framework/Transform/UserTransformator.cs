﻿using Microsoft.SharePoint.Client;
using SharePointPnP.Modernization.Framework.Cache;
using SharePointPnP.Modernization.Framework.Entities;
using SharePointPnP.Modernization.Framework.Telemetry;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace SharePointPnP.Modernization.Framework.Transform
{
    public class UserTransformator: BaseTransform
    {

        private ClientContext _sourceContext;
        private ClientContext _targetContext;
        private List<UserMappingEntity> _userMapping;
        private bool _useOriginalValuesOnNoMatch;
        private string _ldapSpecifiedByUser;
        private SPVersion _sourceVersion;

        /// <summary>
        /// Determine if the user transforming according to mapped file
        /// </summary>
        public bool IsUserMappingSpecified { get
            {
                return (this._userMapping != default);
            } 
        }

        #region Construction
        
        /// <summary>
        /// User Transformator constructor
        /// </summary>
        /// <param name="baseTransformationInformation">Transformation configuration settings</param>
        /// <param name="sourceContext">Source Context</param>
        /// <param name="targetContext">Target Context</param>
        /// <param name="logObservers">Logging</param>
        /// <param name="useOriginalValuesOnNoMatch"></param>
        public UserTransformator(BaseTransformationInformation baseTransformationInformation, ClientContext sourceContext, ClientContext targetContext, IList<ILogObserver> logObservers = null, bool useOriginalValuesOnNoMatch = true)
        {
            // Hookup logging
            if (logObservers != null)
            {
                foreach (var observer in logObservers)
                {
                    base.RegisterObserver(observer);
                }
            }

            // Ensure source and target context are set
            if (sourceContext == null && targetContext != null)
            {
                sourceContext = targetContext;
            }

            if (targetContext == null && sourceContext != null)
            {
                targetContext = sourceContext;
            }
                        
            this._sourceContext = sourceContext;
            this._targetContext = targetContext;

            // Load the User mapping file
            if (!string.IsNullOrEmpty(baseTransformationInformation?.UserMappingFile))
            {
                this._userMapping = CacheManager.Instance.GetUserMapping(baseTransformationInformation.UserMappingFile, logObservers);
            }
            else
            {
                this._userMapping = default; //Pass through if there is no mapping
            }

            _useOriginalValuesOnNoMatch = useOriginalValuesOnNoMatch;

            _ldapSpecifiedByUser = baseTransformationInformation?.LDAPQuery ?? string.Empty;

            _sourceVersion = baseTransformationInformation?.SourceVersion ?? SPVersion.SPO; // SPO Fall Back
        }

        #endregion

        /*
         *  User Cases
         *      SME running PnP Transform not connected to AD but can connect to both On-Prem SP and SPO.
         *      SME running PnP Transform connected to AD and can connect to both On-Prem SP and SPO.
         *      SME running PnP Transform connected to AD and can connect to both On-Prem SP and SPO not AD Synced.
         *      Majority of the target functions already perform checking against the target context
         *      
         *  Design Notes
         *  
         *    	 Executing transform computer IS NOT on the domain
		 *          Answer: Specify Domain (assumes that Computer can talk to domain)
	     *       Executing transform computer is on the domain
		 *          Use c#: System.DirectoryServices.ActiveDirectory.Domain.GetComputerDomain()
	     *       SharePoint and Office 365 is AD Connected
		 *          Auto-resolution via UPN
	     *       SharePoint and Office 365 is NOT AD Connected
		 *          Mapping only, unless credentials from connection can also query domain controllers
         *       Owner, Member, Reader auto mapping
         */

        /// <summary>
        /// Remap principal to alternative mapping
        /// </summary>
        /// <param name="principal"></param>
        /// <returns></returns>
        public string RemapPrincipal(string principalInput)
        {
            LogDebug($"Principal Input: {principalInput}", LogStrings.Heading_UserTransform);

            // Mapping Provided
            // Allow all types of platforms
            if(this.IsUserMappingSpecified)
            {
                // Find Mapping
                // We dont like mulitple matches
                // There are token added to the source address that may need to be replaced
                // When matching, do with and without the tokens   
                var result = principalInput;
                var firstCheck = this._userMapping.Where(o => o.SourceUser.Equals(principalInput, StringComparison.InvariantCultureIgnoreCase));
                if(firstCheck.Count() == 0)
                {
                    //Second check
                    if (principalInput.Contains("|"))
                    {
                        var tokenSplit = principalInput.Split('|');
                        var secondCheck = this._userMapping.Where(o => o.SourceUser.Equals(tokenSplit[1], StringComparison.InvariantCultureIgnoreCase));

                        if(secondCheck.Count() > 0)
                        {
                            result = secondCheck.First().TargetUser;

                            // Log Result
                            if (secondCheck.Count() > 1)
                            {
                                // Log Warning, only first user replaced
                                LogWarning(string.Format(LogStrings.Warning_MultipleMatchFound, result), 
                                    LogStrings.Heading_UserTransform);
                            }
                            else
                            {
                                LogInfo(string.Format(LogStrings.UserTransformSuccess, principalInput, result), 
                                    LogStrings.Heading_UserTransform);
                            }   
                        }
                        else
                        {
                            //Not Found Logging, let method pass-through with original value
                            LogDebug(string.Format(LogStrings.UserTransformMappingNotFound, principalInput),
                                LogStrings.Heading_UserTransform);
                        }
                    }
                }
                else
                {
                    //Found Match
                    result = firstCheck.First().TargetUser;

                    if (firstCheck.Count() > 1)
                    {
                        // Log Warning, only first user replaced
                        LogWarning(string.Format(LogStrings.Warning_MultipleMatchFound, result),
                            LogStrings.Heading_UserTransform);
                    }
                    else
                    {
                        LogInfo(string.Format(LogStrings.UserTransformSuccess, principalInput, result),
                            LogStrings.Heading_UserTransform);
                    }
                }

                return result;

            }
            else
            {
                // If not then default user transformation from on-premises only.
                if(_sourceVersion != SPVersion.SPO && IsExecutingTransformOnDomain())
                {
                    LogDebug($"Default remapping of user {principalInput}", LogStrings.Heading_UserTransform);

                    // If a group, remove the domain element if specified
                    // this assumes that groups are named the same in SharePoint Online
                    var basicPrincipal = StripUserPrefixTokenAndDomain(principalInput);
                    var principalResult = SearchSourceDomainForUPN(AccountType.User, basicPrincipal);

                    if (string.IsNullOrEmpty(principalResult))
                    {
                        // If a user, replace with the UPN
                        principalResult = SearchSourceDomainForUPN(AccountType.Group, basicPrincipal);
                    }

                    if (!string.IsNullOrEmpty(principalResult))
                    {
                        LogInfo($"Remapped user {principalInput} with {principalResult}", LogStrings.Heading_UserTransform);
                        
                        // Resolve group SID or name
                        principalInput = principalResult;
                        
                    }
                }
                else
                {
                    LogInfo($"Not remapping user {principalInput}", LogStrings.Heading_UserTransform);
                }
            }

            //Returns original input to pass through where re-mapping is not required
            return principalInput;
        }
        
        /// <summary>
        /// Determine if the transform is running on a computer on the domain
        /// </summary>
        /// <returns></returns>
        internal bool IsExecutingTransformOnDomain()
        {
            try
            {
                if (_sourceContext != null && _sourceContext.Credentials is NetworkCredential)
                {
                    //Assumes the connection domain to SP is the same domain as the user
                    var credential = _sourceContext.Credentials as NetworkCredential;
                    return (credential.Domain == System.Environment.UserDomainName);
                }
            }
            catch
            {
                // Cannot be sure the user is on the domain for the auto-resolution
                LogWarning("Failed to detect if user is part of the domain, please use mapping instead.", LogStrings.Heading_UserTransform);
            }

            return false;
        }

        /// <summary>
        /// Gets the transform executing domain
        /// </summary>
        /// <returns></returns>
        internal string GetFriendlyComputerDomain()
        {
            try
            {
                //System.DirectoryServices.ActiveDirectory.Domain.GetComputerDomain() - can fail if AD system unstable
                //System.Environment.UserDomainName
                //System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties().DomainName;

                if(_sourceContext != null && _sourceContext.Credentials is NetworkCredential)
                {
                    //Assumes the connection domain to SP is the same domain as the user
                    var credential = _sourceContext.Credentials as NetworkCredential; 
                    return credential.Domain;
                }
                else
                {
                    return System.Environment.UserDomainName;
                }
 
            }
            catch
            {
                LogWarning("Cannot get current domain", LogStrings.Heading_UserTransform);
            }

            return string.Empty;
        }

        /// <summary>
        /// Get LDAP Connection string
        /// </summary>
        /// <returns></returns>
        internal string GetLDAPConnectionString()
        {
            if (!string.IsNullOrEmpty(this._ldapSpecifiedByUser))
            {
                return _ldapSpecifiedByUser;
            }
            else
            {
                // Example from test rig
                /*
                    Forest                  : AlphaDelta.Local
                    DomainControllers       : {AD.AlphaDelta.Local}
                    Children                : {}
                    DomainMode              : Unknown
                    DomainModeLevel         : 7
                    Parent                  :
                    PdcRoleOwner            : AD.AlphaDelta.Local
                    RidRoleOwner            : AD.AlphaDelta.Local
                    InfrastructureRoleOwner : AD.AlphaDelta.Local
                    Name                    : AlphaDelta.Local
                */

                // User Provided with the base transformation information

                // Auto Detect and calculate
                var friendlyDomainName = GetFriendlyComputerDomain();
                var fqdn = ResolveFriendlyDomainToLdapDomain(friendlyDomainName);

                if (!string.IsNullOrEmpty(fqdn))
                {
                    StringBuilder builder = new StringBuilder();
                    builder.Append("LDAP://");
                    foreach (var part in fqdn.Split('.'))
                    {
                        builder.Append($"DC={part},");
                    }

                    return builder.ToString().TrimEnd(',');
                }

                return string.Empty;
            }            
        }

        /// <summary>
        /// Search the source domain for a UPN
        /// </summary>
        /// <param name="accountType"></param>
        /// <param name="samAccountName"></param>
        internal string SearchSourceDomainForUPN(AccountType accountType, string samAccountName)
        {
            try
            {

                //reference: https://github.com/SharePoint/PnP-Transformation/blob/master/InfoPath/Migration/PeoplePickerRemediation.Console/PeoplePickerRemediation.Console/PeoplePickerRemediation.cs#L613

                //e.g. LDAP://DC=onecity,DC=corp,DC=fabrikam,DC=com
                string ldapQuery = GetLDAPConnectionString();

                if (!string.IsNullOrEmpty(ldapQuery))
                {

                    // Bind to the users container.
                    DirectoryEntry entry = new DirectoryEntry(ldapQuery);

                    // Create a DirectorySearcher object.
                    DirectorySearcher mySearcher = new DirectorySearcher(entry);
                    // Create a SearchResultCollection object to hold a collection of SearchResults
                    // returned by the FindAll method.
                    mySearcher.PageSize = 500;

                    string strFilter = string.Empty;
                    if (accountType == AccountType.User)
                    {
                        strFilter = string.Format("(&(objectCategory=User)(| (SAMAccountName={0})(cn={0})))", samAccountName);
                    }
                    else if (accountType == AccountType.Group)
                    {
                        strFilter = string.Format("(&(objectCategory=Group)(objectClass=group)(| (objectsid={0})(name={0})))", samAccountName);
                    }

                    var propertiesToLoad = new[] { "SAMAccountName", "userprincipalname", "sid" };

                    mySearcher.PropertiesToLoad.AddRange(propertiesToLoad);
                    mySearcher.Filter = strFilter;
                    mySearcher.CacheResults = false;

                    SearchResultCollection result = mySearcher.FindAll(); //Consider FindOne

                    if (result != null && result.Count > 0)
                    {
                        if (accountType == AccountType.User)
                        {
                            return GetProperty(result[0], "userprincipalname");
                        }

                        if (accountType == AccountType.Group)
                        {
                            return GetProperty(result[0], "samaccountname"); // This will only confirm existance
                        }
                    }

                }
                else
                {
                    LogWarning("Cann use the LDAP Query to connect to domain", LogStrings.Heading_UserTransform);
                }
            }
            catch(Exception ex)
            {
                LogError("Error Searching Source Domain For UPN", LogStrings.Heading_UserTransform, ex);
            }

            return string.Empty;
        }

        /// <summary>
        /// Get a property from resulting AD query
        /// </summary>
        /// <param name="searchResult"></param>
        /// <param name="PropertyName"></param>
        /// <returns></returns>
        private static string GetProperty(SearchResult searchResult, string PropertyName)
        {
            if (searchResult.Properties.Contains(PropertyName))
            {
                return searchResult.Properties[PropertyName][0].ToString();
            }
            else
            {
                return string.Empty;
            }
        }
        
        /// <summary>
        /// Resolves friendly domain name to Fully Qualified Domain Name
        /// </summary>
        /// <param name="friendlyDomainName"></param>
        /// <returns></returns>
        internal string ResolveFriendlyDomainToLdapDomain(string friendlyDomainName)
        {
            //Reference and credit: https://www.codeproject.com/Articles/18102/Howto-Almost-Everything-In-Active-Directory-via-C#13 

            string ldapPath = string.Empty;

            try
            {
                DirectoryContext objContext = new DirectoryContext(
                    DirectoryContextType.Domain, friendlyDomainName);
                Domain objDomain = Domain.GetDomain(objContext);
                ldapPath = objDomain.Name;
                return ldapPath;
            }
            catch (Exception ex)
            {
                LogError("Error Resolving Friendly Domain To Ldap Domain", LogStrings.Heading_UserTransform, ex);
            }
            return string.Empty;
        }

        /// <summary>
        /// Strip User Prefix Token And Domain
        /// </summary>
        /// <param name="principal"></param>
        /// <returns></returns>
        public string StripUserPrefixTokenAndDomain(string principal)
        {
            var cleanerString = principal;

            if (principal.Contains('|'))
            {
                cleanerString = principal.Split('|')[1];
            }

            if (principal.Contains('\\')){
                cleanerString = principal.Split('\\')[1];
            }

            return cleanerString;
        }

    }

    /// <summary>
    /// Simple class for value for account type
    /// </summary>
    public enum AccountType
    {
        User,
        Group
    }
}
