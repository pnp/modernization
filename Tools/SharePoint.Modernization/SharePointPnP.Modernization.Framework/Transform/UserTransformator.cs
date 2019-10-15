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

        /// <summary>
        /// Determine if the user transforming according to mapped file
        /// </summary>
        public bool IsUserTranforming { get
            {
                return (this._userMapping != default);
            } 
        }

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
         *      Majority of the target functions already perform checking against the target context
         */

        /// <summary>
        /// Remap principal to alternative mapping
        /// </summary>
        /// <param name="principal"></param>
        /// <returns></returns>
        public string RemapPrincipal(string principalInput)
        {
            if(this.IsUserTranforming)
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

            return principalInput;
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
              

    }
}
