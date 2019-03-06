using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using OfficeDevPnP.Core.Pages;
using SharePointPnP.Modernization.Framework.Entities;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SharePointPnP.Modernization.Framework.Cache
{
    /// <summary>
    /// Caching manager, singleton
    /// Important: don't cache SharePoint Client objects as these are tied to a specific client context and hence will fail when there's context switching!
    /// </summary>
    public sealed class CacheManager
    {
        private static readonly Lazy<CacheManager> _lazyInstance = new Lazy<CacheManager>(() => new CacheManager());
        private ConcurrentDictionary<string, List<ClientSideComponent>> clientSideComponents;
        private ConcurrentDictionary<Guid, string> siteToComponentMapping;
        private ConcurrentDictionary<string, List<FieldData>> fieldsToCopy;
        private OfficeDevPnP.Core.Framework.Provisioning.Model.ProvisioningTemplate baseTemplate;

        /// <summary>
        /// Get's the single cachemanager instance, singleton pattern
        /// </summary>
        public static CacheManager Instance
        {
            get
            {
                return _lazyInstance.Value;
            }
        }

        private CacheManager()
        {
            // place for instance initialization code
            clientSideComponents = new ConcurrentDictionary<string, List<ClientSideComponent>>(10, 10);
            siteToComponentMapping = new ConcurrentDictionary<Guid, string>(10, 100);
            baseTemplate = null;
            fieldsToCopy = new ConcurrentDictionary<string, List<FieldData>>(10, 10);
            AssetsTransfered = new List<AssetTransferredEntity>();
        }

        #region Asset Transfer
        /// <summary>
        /// List of assets transferred from source to destination
        /// </summary>
        public List<AssetTransferredEntity> AssetsTransfered { get; set; }
        #endregion

        #region Client Side Components
        /// <summary>
        /// Get's the clientside components from cache or if needed retrieves and caches them
        /// </summary>
        /// <param name="page">Page to grab the components for</param>
        /// <returns></returns>
        public List<ClientSideComponent> GetClientSideComponents(ClientSidePage page)
        {
            Guid webId = page.Context.Web.Id;

            if (siteToComponentMapping.ContainsKey(webId))
            {
                // Components are cached for this site, get the component key
                if (siteToComponentMapping.TryGetValue(webId, out string componentKey))
                {
                    if (clientSideComponents.TryGetValue(componentKey, out List<ClientSideComponent> componentList))
                    {
                        return componentList;
                    }
                }
            }

            // Ok, so nothing in cache so it seems, so let's get the components
            var componentsToAdd = page.AvailableClientSideComponents().ToList();

            // calculate the componentkey
            string componentKeyToCache = Sha256(JsonConvert.SerializeObject(componentsToAdd));

            // store the retrieved data in cache
            if (siteToComponentMapping.TryAdd(webId, componentKeyToCache))
            {
                // Since the components list is big and often the same across webs we only store it in cache if it's different
                if (!clientSideComponents.ContainsKey(componentKeyToCache))
                {
                    clientSideComponents.TryAdd(componentKeyToCache, componentsToAdd);
                }
            }

            return componentsToAdd;
        }

        /// <summary>
        /// Clear the clientside component cache
        /// </summary>
        public void ClearClientSideComponents()
        {
            clientSideComponents.Clear();
            siteToComponentMapping.Clear();
        }
        #endregion

        #region Base template
        /// <summary>
        /// Get's the base template that will be used to filter out "OOB" fields
        /// </summary>
        /// <param name="web">web to operate against</param>
        /// <returns>Provisioning template of the base template of STS#0</returns>
        public OfficeDevPnP.Core.Framework.Provisioning.Model.ProvisioningTemplate GetBaseTemplate(Web web)
        {
            if (this.baseTemplate == null)
            {
                this.baseTemplate = web.GetBaseTemplate("STS", 0);

                // Ensure certain new fields are there
                var sitePagesInBaseTemplate = baseTemplate.Lists.Where(p => p.Url == "SitePages").FirstOrDefault();
                if (sitePagesInBaseTemplate != null)
                {
                    AddFieldRef(sitePagesInBaseTemplate.FieldRefs, new Guid("ccc1037f-f65e-434a-868e-8c98af31fe29"), "_ComplianceFlags");
                    AddFieldRef(sitePagesInBaseTemplate.FieldRefs, new Guid("d4b6480a-4bed-4094-9a52-30181ea38f1d"), "_ComplianceTag");
                    AddFieldRef(sitePagesInBaseTemplate.FieldRefs, new Guid("92be610e-ddbb-49f4-b3b1-5c2bc768df8f"), "_ComplianceTagWrittenTime");
                    AddFieldRef(sitePagesInBaseTemplate.FieldRefs, new Guid("418d7676-2d6f-42cf-a16a-e43d2971252a"), "_ComplianceTagUserId");
                    AddFieldRef(sitePagesInBaseTemplate.FieldRefs, new Guid("8382d247-72a9-44b1-9794-7b177edc89f3"), "_IsRecord");
                    AddFieldRef(sitePagesInBaseTemplate.FieldRefs, new Guid("d307dff3-340f-44a2-9f4b-fbfe1ba07459"), "_CommentCount");
                    AddFieldRef(sitePagesInBaseTemplate.FieldRefs, new Guid("db8d9d6d-dc9a-4fbd-85f3-4a753bfdc58c"), "_LikeCount");
                    AddFieldRef(sitePagesInBaseTemplate.FieldRefs, new Guid("3a6b296c-3f50-445c-a13f-9c679ea9dda3"), "ComplianceAssetId");
                    AddFieldRef(sitePagesInBaseTemplate.FieldRefs, new Guid("9de685c5-fdf5-4319-b987-3edf55efb36f"), "_SPSitePageFlags");
                    AddFieldRef(sitePagesInBaseTemplate.FieldRefs, new Guid("1a53ab5a-11f9-4b92-a377-8cfaaf6ba7be"), "_DisplayName");
                }
            }

            return this.baseTemplate;
        }

        /// <summary>
        /// Clear base template cache
        /// </summary>
        public void ClearBaseTemplate()
        {
            this.baseTemplate = null;
        }

        /// <summary>
        /// Get the list of fields that need to be copied from cache. If cache is empty the list will be calculated
        /// </summary>
        /// <param name="web">Web to operate against</param>
        /// <param name="pagesLibrary">Pages library instance</param>
        /// <returns>List of fields that need to be copied</returns>
        public List<FieldData> GetFieldsToCopy(Web web, List pagesLibrary)
        {
            List<FieldData> fieldsToCopyRetrieved = new List<FieldData>();

            // Did we already do the calculation for this sitepages library? If so then return from cache
            if (fieldsToCopy.ContainsKey(pagesLibrary.Id.ToString()))
            {
                if (fieldsToCopy.TryGetValue(pagesLibrary.Id.ToString(), out List<FieldData> fields))
                {
                    return fields;
                }
            }
            else
            {
                // calculate the fields to copy
                var baseTemplate = GetBaseTemplate(web);
                if (baseTemplate != null)
                {
                    var sitePagesInBaseTemplate = baseTemplate.Lists.Where(p => p.Url == "SitePages").FirstOrDefault();

                    // Compare site pages list fields
                    foreach (var sitePagesField in pagesLibrary.Fields.Where(p => p.Hidden == false).ToList())
                    {
                        // Skip OOB fields
                        if (!OfficeDevPnP.Core.Enums.BuiltInFieldId.Contains(sitePagesField.Id))
                        {
                            if (sitePagesInBaseTemplate != null)
                            {
                                var fieldFoundInBaseSitePages = sitePagesInBaseTemplate.FieldRefs.Where(p => p.Name == sitePagesField.StaticName).FirstOrDefault();
                                if (fieldFoundInBaseSitePages == null)
                                {
                                    // copy metadata for this field
                                    FieldData fieldToAdd = new FieldData()
                                    {
                                        FieldName = sitePagesField.StaticName,
                                        FieldId = sitePagesField.Id,
                                        FieldType = sitePagesField.TypeAsString,
                                    };

                                    fieldsToCopyRetrieved.Add(fieldToAdd);
                                }
                            }
                        }
                    }

                    // Add to cache
                    if (fieldsToCopy.TryAdd(pagesLibrary.Id.ToString(), fieldsToCopyRetrieved))
                    {
                        return fieldsToCopyRetrieved;
                    }
                }
            }

            // We should not get here...
            return null;
        }

        /// <summary>
        /// Clear the fields to copy cache
        /// </summary>
        public void ClearFieldsToCopy()
        {
            this.fieldsToCopy.Clear();
        }

        #endregion

        #region Generic methods
        public void ClearAllCaches()
        {
            this.AssetsTransfered.Clear();
            ClearClientSideComponents();
            ClearBaseTemplate();
            ClearFieldsToCopy();
        }
        #endregion

        private static string Sha256(string randomString)
        {
            var crypt = new System.Security.Cryptography.SHA256Managed();
            var hash = new StringBuilder();
            byte[] crypto = crypt.ComputeHash(Encoding.UTF8.GetBytes(randomString));
            foreach (byte theByte in crypto)
            {
                hash.Append(theByte.ToString("x2"));
            }
            return hash.ToString();
        }

        private void AddFieldRef(OfficeDevPnP.Core.Framework.Provisioning.Model.FieldRefCollection fieldRefs, Guid Id, string name)
        {
            if (fieldRefs.Where(p => p.Id.Equals(Id)).FirstOrDefault() == null)
            {
                fieldRefs.Add(new OfficeDevPnP.Core.Framework.Provisioning.Model.FieldRef(name) { Id = Id });
            }
        }

    }
}
