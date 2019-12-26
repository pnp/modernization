using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using OfficeDevPnP.Core.Pages;
using SharePointPnP.Modernization.Framework.Entities;
using SharePointPnP.Modernization.Framework.Extensions;
using SharePointPnP.Modernization.Framework.Publishing;
using SharePointPnP.Modernization.Framework.Telemetry;
using SharePointPnP.Modernization.Framework.Transform;
using SharePointPnP.Modernization.Framework.Utilities;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace SharePointPnP.Modernization.Framework.Cache
{
    /// <summary>
    /// Caching manager, singleton
    /// Important: don't cache SharePoint Client objects as these are tied to a specific client context and hence will fail when there's context switching!
    /// </summary>
    public sealed class CacheManager
    {
        private static readonly Lazy<CacheManager> _lazyInstance = new Lazy<CacheManager>(() => new CacheManager());
        private OfficeDevPnP.Core.Framework.Provisioning.Model.ProvisioningTemplate baseTemplate;
        private ConcurrentDictionary<string, List<ClientSideComponent>> clientSideComponents;
        private ConcurrentDictionary<Guid, string> siteToComponentMapping;
        private ConcurrentDictionary<string, List<FieldData>> fieldsToCopy;
        private ConcurrentDictionary<uint, string> publishingPagesLibraryNames;
        private ConcurrentDictionary<uint, string> blogListNames;
        private ConcurrentDictionary<string, string> webType;
        private ConcurrentDictionary<string, Dictionary<uint, string>> resourceStrings;
        private ConcurrentDictionary<string, PageLayout> generatedPageLayoutMappings;
        private ConcurrentDictionary<string, Dictionary<int, UserEntity>> userJsonStrings;
        private ConcurrentDictionary<string, Dictionary<string, UserEntity>> userJsonStringsViaUpn;
        private ConcurrentDictionary<string, Dictionary<string, ResolvedUser>> ensuredUsers;
        private ConcurrentDictionary<string, string> contentTypes;
        private ConcurrentDictionary<string, List<FieldData>> publishingContentTypeFields;
        private ConcurrentDictionary<Uri, Guid> aadTenantId;
        private ConcurrentDictionary<Uri, SPVersion> sharepointVersions;
        private ConcurrentDictionary<Uri, string> exactSharepointVersions;
        private BasePageTransformator lastUsedTransformator;
        private List<UrlMapping> urlMapping;
        private List<UserMappingEntity> userMappings;
        private List<AssetTransferredEntity> assetsTransfered;
        private Dictionary<string, string> mappedUsers;

        private static readonly string Publishing = "publishing";
        private static readonly string Blog = "blog";

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

        #region Construction
        private CacheManager()
        {
            // place for instance initialization code
            clientSideComponents = new ConcurrentDictionary<string, List<ClientSideComponent>>();
            siteToComponentMapping = new ConcurrentDictionary<Guid, string>();
            baseTemplate = null;
            fieldsToCopy = new ConcurrentDictionary<string, List<FieldData>>();
            assetsTransfered = new List<AssetTransferredEntity>();
            publishingPagesLibraryNames = new ConcurrentDictionary<uint, string>();
            blogListNames = new ConcurrentDictionary<uint, string>();
            webType = new ConcurrentDictionary<string, string>();
            resourceStrings = new ConcurrentDictionary<string, Dictionary<uint, string>>();
            generatedPageLayoutMappings = new ConcurrentDictionary<string, PageLayout>();
            userJsonStrings = new ConcurrentDictionary<string, Dictionary<int, UserEntity>>();
            userJsonStringsViaUpn = new ConcurrentDictionary<string, Dictionary<string, UserEntity>>();
            ensuredUsers = new ConcurrentDictionary<string, Dictionary<string, ResolvedUser>>();
            mappedUsers = new Dictionary<string, string>();
            contentTypes = new ConcurrentDictionary<string, string>();
            publishingContentTypeFields = new ConcurrentDictionary<string, List<FieldData>>();
            sharepointVersions = new ConcurrentDictionary<Uri, SPVersion>();
            exactSharepointVersions = new ConcurrentDictionary<Uri, string>();
            aadTenantId = new ConcurrentDictionary<Uri, Guid>();
        }
        #endregion

        #region SharePoint Versions and AAD
        /// <summary>
        /// Get's the cached SharePoint version for a given site
        /// </summary>
        /// <param name="site">Site to get the SharePoint version for</param>
        /// <returns>Found SharePoint version or "Unknown" if not found in cache</returns>
        public SPVersion GetSharePointVersion(Uri site)
        {
            if (this.sharepointVersions.ContainsKey(site))
            {
                return this.sharepointVersions[site];
            }

            return SPVersion.Unknown;
        }

        /// <summary>
        /// Sets the SharePoint version in cache
        /// </summary>
        /// <param name="site">Site to the set the SharePoint version for</param>
        /// <param name="version">SharePoint version of the site</param>
        public void SetSharePointVersion(Uri site, SPVersion version)
        {
            if (!this.sharepointVersions.ContainsKey(site))
            {
                this.sharepointVersions.TryAdd(site, version);
            }
        }

        /// <summary>
        /// Get's the exact SharePoint version from cache
        /// </summary>
        /// <param name="site">Site to get the exact version for</param>
        /// <returns>Exact version from cache</returns>
        public string GetExactSharePointVersion(Uri site)
        {
            if (this.exactSharepointVersions.ContainsKey(site))
            {
                return this.exactSharepointVersions[site];
            }

            return null;
        }

        /// <summary>
        /// Adds exact SharePoint version for a given site to cache
        /// </summary>
        /// <param name="site">Site to add the SharePoint version for to cache</param>
        /// <param name="version">Version to add</param>
        public void SetExactSharePointVersion(Uri site, string version)
        {
            if (!this.exactSharepointVersions.ContainsKey(site))
            {
                this.exactSharepointVersions.TryAdd(site, version);    
            }
        }

        /// <summary>
        /// Returns the used AzureAD tenant id
        /// </summary>
        /// <param name="site">Url of the site</param>
        /// <returns>Azure AD tenant id</returns>
        public Guid GetAADTenantId(Uri site)
        {
            if (this.aadTenantId.ContainsKey(site))
            {
                return this.aadTenantId[site];
            }
            else
            {
                return Guid.Empty;
            }
        }

        /// <summary>
        /// Sets the Azure AD tenant Id in cache
        /// </summary>
        /// <param name="tenantId">Tenant Id</param>
        /// <param name="site">Site url</param>
        public void SetAADTenantId(Guid tenantId, Uri site)
        {
            if (!this.aadTenantId.ContainsKey(site))
            {
                this.aadTenantId.TryAdd(site, tenantId);
            }
        }
        #endregion

        #region Asset Transfer
        public List<AssetTransferredEntity> GetAssetsTransferred()
        {
            return this.assetsTransfered;
        }

        public void AddAssetTransferredEntity(AssetTransferredEntity asset)
        {
            if (!this.assetsTransfered.Contains(asset))
            {
                this.assetsTransfered.Add(asset);
            }
        }
        #endregion

        #region Client Side Components
        /// <summary>
        /// Get's the clientside components from cache or if needed retrieves and caches them
        /// </summary>
        /// <param name="page">Page to grab the components for</param>
        /// <returns></returns>
        public List<ClientSideComponent> GetClientSideComponents(ClientSidePage page)
        {
            Guid webId = page.Context.Web.EnsureProperty(o => o.Id);

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

        #region Base template and metadata
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
        /// <param name="sourceLibrary">Pages library instance</param>
        /// <returns>List of fields that need to be copied</returns>
        public List<FieldData> GetFieldsToCopy(Web web, List sourceLibrary, string pageType)
        {
            List<FieldData> fieldsToCopyRetrieved = new List<FieldData>();

            // Did we already do the calculation for this sitepages library? If so then return from cache
            if (fieldsToCopy.ContainsKey(sourceLibrary.Id.ToString()))
            {
                if (fieldsToCopy.TryGetValue(sourceLibrary.Id.ToString(), out List<FieldData> fields))
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
                    foreach (var sourceField in sourceLibrary.Fields.Where(p => p.Hidden == false).ToList())
                    {
                        // Skip OOB fields
                        if (!IsBuiltInField(pageType.Equals("BlogPage", StringComparison.InvariantCultureIgnoreCase), sourceField.Id))
                        {
                            if (sitePagesInBaseTemplate != null)
                            {
                                var fieldFoundInBaseSitePages = sitePagesInBaseTemplate.FieldRefs.Where(p => p.Name == sourceField.StaticName).FirstOrDefault();
                                if (fieldFoundInBaseSitePages == null)
                                {
                                    // copy metadata for this field
                                    FieldData fieldToAdd = new FieldData()
                                    {
                                        FieldName = sourceField.StaticName,
                                        FieldId = sourceField.Id,
                                        FieldType = sourceField.TypeAsString,
                                    };

                                    fieldsToCopyRetrieved.Add(fieldToAdd);
                                }
                            }
                        }
                    }

                    // Add to cache
                    if (fieldsToCopy.TryAdd(sourceLibrary.Id.ToString(), fieldsToCopyRetrieved))
                    {
                        return fieldsToCopyRetrieved;
                    }
                }
            }

            // We should not get here...
            return null;
        }

        /// <summary>
        /// Get field information of a content type field
        /// </summary>
        /// <param name="pagesLibrary">Pages library list</param>
        /// <param name="contentTypeId">ID of the content type</param>
        /// <param name="fieldName">Name of the field to get information from</param>
        /// <returns>FieldData object holding field information</returns>
        public FieldData GetPublishingContentTypeField(List pagesLibrary, string contentTypeId, string fieldName)
        {
            // Try to get from cache
            if (this.publishingContentTypeFields.TryGetValue(contentTypeId, out List<FieldData> fieldsFromCache))
            {
                // return field if found
                return fieldsFromCache.Where(p => p.FieldName.Equals(fieldName, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
            }

            ClientContext context = pagesLibrary.Context as ClientContext;

            // Load all fields of the list
            FieldCollection fields = pagesLibrary.Fields;
            context.Load(fields, fs => fs.Include(f => f.Id, f => f.TypeAsString, f => f.InternalName));
            context.ExecuteQueryRetry();

            List<FieldData> contentTypeFieldsInList = new List<FieldData>();
            foreach (var field in fields)
            {
                contentTypeFieldsInList.Add(new FieldData()
                {
                    FieldId = field.Id,
                    FieldName = field.InternalName,
                    FieldType = field.TypeAsString,
                });
            }

            var authorField = new FieldData()
            {
                FieldId = new Guid("1df5e554-ec7e-46a6-901d-d85a3881cb18"),
                FieldName = "Author",
                FieldType = "User",
            };

            var editorField = new FieldData()
            {
                FieldId = new Guid("d31655d1-1d5b-4511-95a1-7a09e9b75bf2"),
                FieldName = "Editor",
                FieldType = "User",
            };

            if (!contentTypeFieldsInList.Contains(authorField))
            {
                contentTypeFieldsInList.Add(authorField);
            }

            if (!contentTypeFieldsInList.Contains(editorField))
            {
                contentTypeFieldsInList.Add(editorField);
            }

            // Store in cache
            this.publishingContentTypeFields.TryAdd(contentTypeId, contentTypeFieldsInList);

            // Return field, if found
            return contentTypeFieldsInList.Where(p => p.FieldName.Equals(fieldName, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
        }

        /// <summary>
        /// Clear the fields to copy cache
        /// </summary>
        public void ClearFieldsToCopy()
        {
            this.fieldsToCopy.Clear();
            this.publishingContentTypeFields.Clear();
        }

        #endregion

        #region Web type handling
        /// <summary>
        /// Marks this web as a publishing web
        /// </summary>
        /// <param name="webUrl">Url of the web</param>
        public void SetPublishingWeb(string webUrl)
        {
            if (!this.webType.ContainsKey(webUrl))
            {
                this.webType.TryAdd(webUrl, CacheManager.Publishing);
            }
        }

        /// <summary>
        /// Marks this web as a blog web
        /// </summary>
        /// <param name="webUrl">Url of the web</param>
        public void SetBlogWeb(string webUrl)
        {
            if (!this.webType.ContainsKey(webUrl))
            {
                this.webType.TryAdd(webUrl, CacheManager.Blog);
            }
        }

        /// <summary>
        /// Checks if this is publishing web
        /// </summary>
        /// <param name="webUrl">Web url to check</param>
        /// <returns>True if publishing, false otherwise</returns>
        public bool IsPublishingWeb(string webUrl)
        {
            if (this.webType.ContainsKey(webUrl))
            {
                if (this.webType.TryGetValue(webUrl, out string type))
                {
                    if (type.Equals(CacheManager.Publishing))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Checks if this is blog web
        /// </summary>
        /// <param name="webUrl">Web url to check</param>
        /// <returns>True if blog, false otherwise</returns>
        public bool IsBlogWeb(string webUrl)
        {
            if (this.webType.ContainsKey(webUrl))
            {
                if (this.webType.TryGetValue(webUrl, out string type))
                {
                    if (type.Equals(CacheManager.Blog))
                    {
                        return true;
                    }
                }
            }

            return false;
        }
        #endregion

        #region Publishing Pages Library
        /// <summary>
        /// Get translation for the publishing pages library
        /// </summary>
        /// <param name="context">Context of the site</param>
        /// <returns>Translated name of the pages library</returns>
        public string GetPublishingPagesLibraryName(ClientContext context)
        {
            // Simplier implementation - Get the Pages library then get the relative URL of the rootfolder of the library

            //Keys: 
            //  Web Property: __PagesListId
            //  Found in 2010, SPO

            string pagesLibraryName = "pages";

            if (context == null)
            {
                return pagesLibraryName;
            }

            uint lcid = context.Web.EnsureProperty(p => p.Language);

            var propertyBagKey = Constants.WebPropertyKeyPagesListId;

            if (publishingPagesLibraryNames.ContainsKey(lcid))
            {
                if (publishingPagesLibraryNames.TryGetValue(lcid, out string name))
                {
                    return name;
                }
                else
                {
                    // let's fallback to the default...we should never get here unless there's some threading issue
                    return pagesLibraryName;
                }
            }
            else
            {
                if (BaseTransform.GetVersion(context) == SPVersion.SP2010)
                {
                    var keyVal = context.Web.GetPropertyBagValueString(propertyBagKey, string.Empty);
                    if (!string.IsNullOrEmpty(keyVal))
                    {
                        var list = context.Web.GetListById(Guid.Parse(keyVal), o => o.RootFolder.ServerRelativeUrl);
                        var webServerRelativeUrl = context.Web.EnsureProperty(w => w.ServerRelativeUrl);

                        pagesLibraryName = list.RootFolder.ServerRelativeUrl.Replace(webServerRelativeUrl, "").Trim('/').ToLower();

                        // add to cache
                        publishingPagesLibraryNames.TryAdd(lcid, pagesLibraryName);

                        return pagesLibraryName;
                    }
                }
                else
                {
                    // Fall back to older logic
                    ClientResult<string> result = Microsoft.SharePoint.Client.Utilities.Utility.GetLocalizedString(context, "$Resources:List_Pages_UrlName", "osrvcore", int.Parse(lcid.ToString()));
                    context.ExecuteQueryRetry();

                    var altPagesLibraryName = new Regex(@"['´`]").Replace(result.Value, "");

                    if (string.IsNullOrEmpty(altPagesLibraryName))
                    {
                        return pagesLibraryName;
                    }

                    // add to cache
                    publishingPagesLibraryNames.TryAdd(lcid, altPagesLibraryName.ToLower());

                    return altPagesLibraryName.ToLower();
                }
            }

            return pagesLibraryName;
        }

        #endregion

        #region Blog list name
        /// <summary>
        /// Get translation for the blog list name
        /// </summary>
        /// <param name="context">Context of the site</param>
        /// <returns>Translated name of the blog list</returns>
        public string GetBlogListName(ClientContext context)
        {
            string blogListName = "posts";

            if (context == null)
            {
                return blogListName;
            }

            uint lcid = context.Web.EnsureProperty(p => p.Language);
            if (blogListNames.ContainsKey(lcid))
            {
                if (blogListNames.TryGetValue(lcid, out string name))
                {
                    return name;
                }
                else
                {
                    // let's fallback to the default...we should never get here unless there's some threading issue
                    return blogListName;
                }
            }
            else
            {
                string altBlogListName = null;
                try
                {
                    ClientResult<string> result = Microsoft.SharePoint.Client.Utilities.Utility.GetLocalizedString(context, "$Resources:blogpost_Folder", "core", int.Parse(lcid.ToString()));
                    context.ExecuteQueryRetry();
                    altBlogListName = new Regex(@"['´`]").Replace(result.Value, "");
                }
                catch
                {
                    // Use "simple" method, which also works for SharePoint 2010
                    altBlogListName = PostsTranslation(lcid);
                }

                if (string.IsNullOrEmpty(altBlogListName))
                {
                    return blogListName;
                }

                // add to cache
                blogListNames.TryAdd(lcid, altBlogListName.ToLower());

                return altBlogListName.ToLower();
            }
        }
        #endregion

        #region Resource strings
        /// <summary>
        /// Returns the translated value for a resource string
        /// </summary>
        /// <param name="context">Context of the site</param>
        /// <param name="resource">Key of the resource (e.g. $Resources:core,ScriptEditorWebPartDescription;) </param>
        /// <returns>Translated string</returns>
        public string GetResourceString(ClientContext context, string resource)
        {
            uint lcid = context.Web.EnsureProperty(p => p.Language);

            if (resourceStrings.ContainsKey(resource))
            {
                if (resourceStrings.TryGetValue(resource, out Dictionary<uint, string> resourceValues))
                {
                    if (resourceValues.ContainsKey(lcid))
                    {
                        if (resourceValues.TryGetValue(lcid, out string resourceValue))
                        {
                            return resourceValue;
                        }
                    }
                }
            }

            // If we got here then we need to still add the resource translation
            var resourceString = resource.Replace("$Resources:", "").Replace(";", "");
            var splitResourceString = resourceString.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            string resourceFile = "core";
            string resourceKey = null;
            if (splitResourceString.Length == 2)
            {
                resourceFile = splitResourceString[0];
                resourceKey = splitResourceString[1];
            }
            else
            {
                resourceKey = splitResourceString[0];
            }

            ClientResult<string> result = Microsoft.SharePoint.Client.Utilities.Utility.GetLocalizedString(context, $"$Resources:{resourceKey}", resourceFile, int.Parse(lcid.ToString()));
            context.ExecuteQueryRetry();

            if (result == null)
            {
                return resource;
            }

            if (resourceStrings.ContainsKey(resource))
            {
                if (resourceStrings.TryGetValue(resource, out Dictionary<uint, string> resourceValues))
                {
                    if (!resourceValues.ContainsKey(lcid))
                    {
                        // Add translations in existing array
                        Dictionary<uint, string> newResourceValues = new Dictionary<uint, string>(resourceValues)
                        {
                            { lcid, result.Value }
                        };
                        resourceStrings.TryUpdate(resource, newResourceValues, resourceValues);
                    }
                }
            }
            else
            {
                // No translations were already retrieved in this language
                Dictionary<uint, string> translations = new Dictionary<uint, string>
                {
                    { lcid, result.Value }
                };

                resourceStrings.TryAdd(resource, translations);
            }

            return result.Value;
        }
        #endregion

        #region PageLayout mappings
        /// <summary>
        /// Generate pagelayout mapping file for given publishing page
        /// </summary>
        /// <param name="page">Publishing page</param>
        /// <returns>Page layout mapping model</returns>
        public PageLayout GetPageLayoutMapping(ListItem page)
        {
            string key = page.PageLayoutFile();

            // Try get the page layout from cache
            if (generatedPageLayoutMappings.TryGetValue(key, out PageLayout pageLayoutFromCache))
            {
                return pageLayoutFromCache;
            }

            PageLayoutAnalyser pageLayoutAnalyzer = new PageLayoutAnalyser(page.Context as ClientContext);

            // Let's try to generate a 'basic' model and use that...not optimal, but better than bailing out.
            var newPageLayoutMapping = pageLayoutAnalyzer.AnalysePageLayoutFromPublishingPage(page);

            // Add to cache for future reuse
            generatedPageLayoutMappings.TryAdd(key, newPageLayoutMapping);

            // Return to requestor
            return newPageLayoutMapping;
        }

        #endregion

        #region Generic methods
        /// <summary>
        /// Clear all the caches
        /// </summary>
        public void ClearAllCaches()
        {
            this.assetsTransfered.Clear();
            ClearClientSideComponents();
            ClearBaseTemplate();

            this.urlMapping?.Clear();
            this.userMappings?.Clear();
            this.ensuredUsers?.Clear();

            ClearFieldsToCopy();
            ClearSharePointVersions();
            ClearGeneratedPageLayoutMappings();
        }

        /// <summary>
        /// Clears Cached SharePoint versions
        /// </summary>
        public void ClearSharePointVersions()
        {
            this.sharepointVersions.Clear();
        }

        /// <summary>
        /// Clears the cache of generated page layout mappings
        /// </summary>
        public void ClearGeneratedPageLayoutMappings()
        {
            this.generatedPageLayoutMappings.Clear();
        }

        #endregion

        #region Users
        /// <summary>
        /// Mapped users
        /// </summary>
        /// <returns>A dictionary of mapped users</returns>
        public Dictionary<string, string> GetMappedUsers()
        {
            return this.mappedUsers;
        }

        /// <summary>
        /// Adds a user to the dictionary of mapped users
        /// </summary>
        /// <param name="principal">Principal to map</param>
        /// <param name="user">mapped user</param>
        public void AddMappedUser(string principal, string user)
        {
            if (!this.mappedUsers.ContainsKey(principal))
            {
                this.mappedUsers.Add(principal, user);
            }
        }

        /// <summary>
        /// Run and cache the output value of EnsureUser for a given user
        /// </summary>
        /// <param name="context">ClientContext to operate on</param>
        /// <param name="userValue">User name of user to ensure</param>
        /// <returns>ResolvedUser instance holding information about the ensured user</returns>
        public ResolvedUser GetEnsuredUser(ClientContext context, string userValue)
        {
            if (string.IsNullOrEmpty(userValue))
            {
                return null;
            }

            string key = context.Web.GetUrl();

            if (this.ensuredUsers.TryGetValue(key, out Dictionary<string, ResolvedUser> ensuredUsersFromCache))
            {
                if (ensuredUsersFromCache.TryGetValue(userValue, out ResolvedUser userLoginName))
                {
                    return userLoginName;
                }
            }

            try
            {
                using (var clonedContext = context.Clone(context.Web.GetUrl()))
                {
                    var userToResolve = clonedContext.Web.EnsureUser(userValue);
                    clonedContext.Load(userToResolve);
                    clonedContext.ExecuteQueryRetry();

                    ResolvedUser resolvedUser = new ResolvedUser()
                    {
                        LoginName = userToResolve.LoginName,
                        Id = userToResolve.Id,
                    };

                    // Store in cache
                    if (ensuredUsersFromCache != null)
                    {
                        // We already has a user list, simply add this one
                        Dictionary<string, ResolvedUser> newEnsuredUsersFromCache = new Dictionary<string, ResolvedUser>(ensuredUsersFromCache)
                            {
                                { userValue, resolvedUser }
                            };

                        this.ensuredUsers.TryUpdate(key, newEnsuredUsersFromCache, ensuredUsersFromCache);
                    }
                    else
                    {
                        // First user for this key (= web)
                        Dictionary<string, ResolvedUser> newEnsuredUsersFromCache = new Dictionary<string, ResolvedUser>()
                            {
                                { userValue, resolvedUser }
                            };

                        this.ensuredUsers.TryAdd(key, newEnsuredUsersFromCache);
                    }

                    return resolvedUser;
                }
            }
            catch
            {
                // Logging is not needed as an "empty" ensured user is handled by the callers of this method
            }

            return null;
        }

        /// <summary>
        /// Lookup a user from the site's user list based upon the user's upn
        /// </summary>
        /// <param name="context">Context of the web holding the user list</param>
        /// <param name="userUpn">Upn of the user to fetch</param>
        /// <returns>A UserEntity instance holding information about the user</returns>
        public UserEntity GetUserFromUserList(ClientContext context, string userUpn)
        {
            if (string.IsNullOrEmpty(userUpn))
            {
                return null;
            }

            string key = context.Web.GetUrl();

            if (!userUpn.StartsWith("i:0#.f|membership|"))
            {
                userUpn = $"i:0#.f|membership|{userUpn}";
            }

            if (this.userJsonStringsViaUpn.TryGetValue(key, out Dictionary<string, UserEntity> userListFromCache))
            {
                if (userListFromCache.TryGetValue(userUpn, out UserEntity userJsonFromCache))
                {
                    return userJsonFromCache;
                }
            }

            try
            {
                string CAMLQueryByName = @"
                <View Scope='Recursive'>
                  <Query>
                    <Where>
                      <Eq>
                        <FieldRef Name='Name'/>
                        <Value Type='Text'>{0}</Value>
                      </Eq>
                    </Where>
                  </Query>
                </View>";

                List siteUserInfoList = context.Web.SiteUserInfoList;
                CamlQuery query = new CamlQuery
                {
                    ViewXml = string.Format(CAMLQueryByName, userUpn)
                };
                var loadedUsers = context.LoadQuery(siteUserInfoList.GetItems(query));
                context.ExecuteQueryRetry();

                UserEntity author = null;
                if (loadedUsers != null)
                {
                    var loadedUser = loadedUsers.FirstOrDefault();
                    if (loadedUser != null)
                    {
                        // Does not work for groups
                        if (loadedUser["Name"] == null)
                        {
                            return null;
                        }

                        bool isGroup = loadedUser["Name"].ToString().StartsWith("c:0t.c|tenant|");
                        string userUpnValue = loadedUser["Name"].ToString().GetUserName();

                        author = new UserEntity()
                        {
                            Upn = userUpnValue,
                            Name = loadedUser["Title"] != null ? loadedUser["Title"].ToString() : "",
                            Role = loadedUser["JobTitle"] != null ? loadedUser["JobTitle"].ToString() : "",
                            LoginName = loadedUser["Name"] != null ? loadedUser["Name"].ToString() : "",
                            IsGroup = isGroup || IsGroup(userUpnValue),
                        };

                        author.Id = $"i:0#.f|membership|{author.Upn}";

                        // Store in cache
                        if (userListFromCache != null)
                        {
                            // We already has a user list, simply add this one
                            Dictionary<string, UserEntity> newUserListToCache = new Dictionary<string, UserEntity>(userListFromCache)
                            {
                                { userUpn, author }
                            };

                            this.userJsonStringsViaUpn.TryUpdate(key, newUserListToCache, userListFromCache);
                        }
                        else
                        {
                            // First user for this key (= web)
                            Dictionary<string, UserEntity> newUserListToCache = new Dictionary<string, UserEntity>()
                            {
                                { userUpn, author }
                            };

                            this.userJsonStringsViaUpn.TryAdd(key, newUserListToCache);
                        }

                        // return 
                        return author;
                    }
                }
            }
            catch
            {
                // Logging is not needed as an "empty" ensured user is handled by the callers of this method
            }

            return null;
        }

        /// <summary>
        /// Lookup a user from the site's user list based upon the user's id
        /// </summary>
        /// <param name="context">Context of the web holding the user list</param>
        /// <param name="userListId">Id of the user to fetch</param>
        /// <returns>A UserEntity instance holding information about the user</returns>
        public UserEntity GetUserFromUserList(ClientContext context, int userListId)
        {
            string key = context.Web.GetUrl();

            if (this.userJsonStrings.TryGetValue(key, out Dictionary<int, UserEntity> userListFromCache))
            {
                if (userListFromCache.TryGetValue(userListId, out UserEntity userJsonFromCache))
                {
                    return userJsonFromCache;
                }
            }

            try
            {
                string CAMLQueryByName = @"
                <View Scope='Recursive'>
                  <Query>
                    <Where>
                      <Contains>
                        <FieldRef Name='ID'/>
                        <Value Type='Integer'>{0}</Value>
                      </Contains>
                    </Where>
                  </Query>
                </View>";

                List siteUserInfoList = context.Web.SiteUserInfoList;
                CamlQuery query = new CamlQuery
                {
                    ViewXml = string.Format(CAMLQueryByName, userListId)
                };
                var loadedUsers = context.LoadQuery(siteUserInfoList.GetItems(query));
                context.ExecuteQueryRetry();

                UserEntity author = null;
                if (loadedUsers != null)
                {
                    var loadedUser = loadedUsers.FirstOrDefault();
                    if (loadedUser != null)
                    {
                        if (loadedUser["Name"] == null)
                        {
                            return null;
                        }

                        bool isGroup = loadedUser["Name"].ToString().StartsWith("c:0t.c|tenant|");
                        string userUpnValue = loadedUser["Name"].ToString().GetUserName();

                        author = new UserEntity()
                        {
                            Upn = userUpnValue,
                            Name = loadedUser["Title"] != null ? loadedUser["Title"].ToString() : "",
                            Role = loadedUser["JobTitle"] != null ? loadedUser["JobTitle"].ToString() : "",
                            LoginName = loadedUser["Name"] != null ? loadedUser["Name"].ToString() : "",
                            IsGroup = isGroup || IsGroup(userUpnValue),
                        };

                        author.Id = $"i:0#.f|membership|{author.Upn}";

                        // Store in cache
                        if (userListFromCache != null)
                        {
                            // We already has a user list, simply add this one
                            Dictionary<int, UserEntity> newUserListToCache = new Dictionary<int, UserEntity>(userListFromCache)
                            {
                                { userListId, author }
                            };

                            this.userJsonStrings.TryUpdate(key, newUserListToCache, userListFromCache);
                        }
                        else
                        {
                            // First user for this key (= web)
                            Dictionary<int, UserEntity> newUserListToCache = new Dictionary<int, UserEntity>()
                            {
                                { userListId, author }
                            };

                            this.userJsonStrings.TryAdd(key, newUserListToCache);
                        }

                        // return 
                        return author;
                    }
                }
            }
            catch
            {
                // Logging is not needed as an "empty" ensured user is handled by the callers of this method
            }

            return null;
        }
        #endregion

        #region Content types
        /// <summary>
        /// Get's the ID of a contenttype
        /// </summary>
        /// <param name="pagesLibrary">Pages library holding the content type</param>
        /// <param name="contentTypeName">Name of the content type</param>
        /// <returns>ID of the content type</returns>
        public string GetContentTypeId(List pagesLibrary, string contentTypeName)
        {
            string contentTypeId = null;

            // try to get from cache
            this.contentTypes.TryGetValue(contentTypeName, out string contentTypeIdFromCache);
            if (!string.IsNullOrEmpty(contentTypeIdFromCache))
            {
                return contentTypeIdFromCache;
            }

            // Load content type
            var ctCol = pagesLibrary.ContentTypes;
            var results = pagesLibrary.Context.LoadQuery(ctCol.Where(item => item.Name == contentTypeName));
            pagesLibrary.Context.ExecuteQueryRetry();

            if (results.FirstOrDefault() != null)
            {
                contentTypeId = results.FirstOrDefault().StringId;

                // We only allow content types that inherit from the OOB Site Page content type
                if (!contentTypeId.StartsWith(Constants.ModernPageContentTypeId, StringComparison.InvariantCultureIgnoreCase))
                {
                    return null;
                }

                // add to cache
                this.contentTypes.TryAdd(contentTypeName, contentTypeId);
            }

            return contentTypeId;
        }
        #endregion

        #region Last used transformator
        /// <summary>
        /// Caches the last used page transformator instance, needed to postpone log writing when transforming multiple pages
        /// </summary>
        /// <param name="transformator"></param>
        public void SetLastUsedTransformator(BasePageTransformator transformator)
        {
            this.lastUsedTransformator = transformator;
        }

        /// <summary>
        /// Gets the last used page transformator instance
        /// </summary>
        /// <returns></returns>
        public BasePageTransformator GetLastUsedTransformator()
        {
            return this.lastUsedTransformator;
        }
        #endregion

        #region URL rewriting
        /// <summary>
        /// Returns a list of url mappings
        /// </summary>
        /// <param name="urlMappingFile">File with url mappings</param>
        /// <param name="logObservers">Attached list of log observers</param>
        /// <returns>List of url mappings</returns>
        public List<UrlMapping> GetUrlMapping(string urlMappingFile, IList<ILogObserver> logObservers = null)
        {
            if (this.urlMapping != null && this.urlMapping.Count > 0)
            {
                return this.urlMapping;
            }

            FileManager fileManager = new FileManager(logObservers);
            this.urlMapping = fileManager.LoadUrlMappingFile(urlMappingFile);

            return this.urlMapping;
        }
        #endregion

        #region Get User Mapping
        /// <summary>
        /// Gets the list of user mappings, if first time file will be laoded
        /// </summary>
        /// <param name="userMappingFile">File with the user mappings</param>
        /// <param name="logObservers">Attached list of log observers</param>
        /// <returns>List of user mappings</returns>
        public List<UserMappingEntity> GetUserMapping(string userMappingFile, IList<ILogObserver> logObservers = null)
        {
            if (this.userMappings != null && this.userMappings.Count > 0)
            {
                return this.userMappings;
            }

            FileManager fileManager = new FileManager(logObservers);
            this.userMappings = fileManager.LoadUserMappingFile(userMappingFile);

            return this.userMappings;
        }
        #endregion

        #region Helper methods

        private static bool IsGroup(string loginName)
        {
            // Possible input
            // c:0t.c|tenant|b0f984d9-e9d5-432a-bec9-896f910254ba (group in SPO)
            // S-5-1-76-1812374880-3438888550-261701130-6117 (group in SPO on-premises)

            if (loginName.StartsWith("c:0t.c|tenant|") || IsSID(loginName))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private static bool IsSID(string loginName)
        {
            return Regex.IsMatch(loginName.ToUpper(), @"^S-\d-\d+-(\d+-){1,14}\d+$");
        }

        private static string PostsTranslation(uint lcid)
        {
            // See https://capacreative.co.uk/resources/reference-sharepoint-online-languages-ids/ for list of language id's
            switch (lcid)
            {
                case 1033: return "Posts";
                case 1027: return "Publicacions";
                case 1029: return "Prispevky";
                case 1030: return "Blogmeddelelser";
                case 1031: return "Beitraege";
                case 1035: return "Viestit";
                case 1036: return "Billets";
                case 1038: return "Bejegyzesek";
                case 1040: return "Post";
                case 1043: return "Berichten";
                case 1044: return "Innlegg";
                case 1046: return "Postagens";
                case 1050: return "Clanci";
                case 1053: return "Anslag";
                case 1057: return "Pos";
                case 1060: return "Obvestila";
                case 1061: return "Postitused";
                case 1069: return "Blog-sarrerak";
                case 1086: return "Catatan";
                case 1106: return "Postiadau";
                case 1110: return "Mensaxes";
                case 2070: return "Artigos";
                case 2074: return "ObjavljenePoruke";
                case 2108: return "Poist";
                case 3082: return "EntradasDeBlog";
                case 5146: return "Objave";
                case 9424: return "ObjavljenePoruke";
                default: return "Posts";
            }
        }

        private static string Sha256(string randomString)
        {
            SHA256CryptoServiceProvider provider = new SHA256CryptoServiceProvider();
            byte[] hash = provider.ComputeHash(Encoding.Unicode.GetBytes(JsonConvert.SerializeObject(randomString)));
            string componentKeyToCache = BitConverter.ToString(hash).Replace("-", "");
            return componentKeyToCache;
        }

        private void AddFieldRef(OfficeDevPnP.Core.Framework.Provisioning.Model.FieldRefCollection fieldRefs, Guid Id, string name)
        {
            if (fieldRefs.Where(p => p.Id.Equals(Id)).FirstOrDefault() == null)
            {
                fieldRefs.Add(new OfficeDevPnP.Core.Framework.Provisioning.Model.FieldRef(name) { Id = Id });
            }
        }

        private bool IsBuiltInField(bool isBlog, Guid fieldId)
        {
            if (OfficeDevPnP.Core.Enums.BuiltInFieldId.Contains(fieldId))
            {
                if (isBlog)
                {
                    // Always allow the PostCategory field
                    if (fieldId.Equals(Constants.PostCategory))
                    {
                        return false;
                    }
                    else
                    {
                        return true;
                    }
                }
                else
                {
                    return true;
                }
            }
            else
            {
                return false;
            }
        }
        #endregion
    }
}
