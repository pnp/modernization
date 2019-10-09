using Microsoft.SharePoint.Client;
using SharePointPnP.Modernization.Framework.Cache;
using SharePointPnP.Modernization.Framework.Entities;
using SharePointPnP.Modernization.Framework.Telemetry;
using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Linq;
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

        #region Construction
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

        }

        #endregion

        /*
         *  User Cases
         *      SME running PnP Transform not connected to AD but can connect to both On-Prem SP and SPO.
         *      SME running PnP Transform connected to AD and can connect to both On-Prem SP and SPO.
         *      SME running PnP Transform connected to AD and can connect to both On-Prem SP and SPO not AD Synced.
         *      
         */

        /// <summary>
        /// Remap principal to alternative mapping
        /// </summary>
        /// <param name="principal"></param>
        /// <returns></returns>
        public string RemapPrincipal(string principal)
        {
            if(this._userMapping != default)
            {
                // Find Mapping
                // We dont like mulitple matches

                // Replace value from input

                // Validate target samAccountName

            }

            return principal;
        }
        
        /// <summary>
        /// Get On-Premises AD UPN value from SID reference
        /// </summary>
        /// <param name="sidInput"></param>
        /// <returns></returns>
        public string GetOnPremUPN(string accountType, string samAccountName)
        {
            throw new NotImplementedException();
            //OK This has got to be a better way of doing this...
            //DO We need this???
            //string ldapQuery = "LDAP://DC=test,DC=contoso,DC=com";

            

            //// Bind to the users container.
            //DirectoryEntry entry = new DirectoryEntry(ldapQuery);
            //// Create a DirectorySearcher object.
            //DirectorySearcher mySearcher = new DirectorySearcher(entry);
            //// Create a SearchResultCollection object to hold a collection of SearchResults
            //// returned by the FindAll method.
            //mySearcher.PageSize = 500;  // ADD THIS LINE HERE !

            //string strFilter = string.Empty;
            //if (accountType.ToLower().Equals("user"))
            //    strFilter = string.Format("(&(objectCategory=User)(SAMAccountName={0}))", samAccountName);
            //else if (accountType.ToLower().Contains("group"))
            //    strFilter = string.Format("(&(objectCategory=Group)(sid={0}))", samAccountName);

            //var propertiesToLoad = new[] { "SAMAccountName", "userprincipalname", "sid" };
            //mySearcher.PropertiesToLoad.AddRange(propertiesToLoad);
            //mySearcher.Filter = strFilter;
            //mySearcher.CacheResults = false;

            //SearchResultCollection result = mySearcher.FindAll();

            //if (result != null && result.Count > 0)
            //{
            //    return GetProperty(result[0], "userprincipalname");
            //}

            //return string.Empty;
           
        }


        /// <summary>
        /// Input contains SID reference
        /// </summary>
        /// <param name="mappingInput"></param>
        /// <returns></returns>
        internal bool ContainsSID(string mappingInput)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Input contains a valid UPN
        /// </summary>
        /// <param name="userMapping"></param>
        /// <returns></returns>
        internal bool IsValidUpn(string userMapping)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Get Property from AD Search Result
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

    }
}
