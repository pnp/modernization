using Newtonsoft.Json;
using OfficeDevPnP.Core.Pages;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointPnP.Modernization.Framework.Cache
{
    /// <summary>
    /// Caching manager, singleton
    /// </summary>
    public sealed class CacheManager
    {
        private static readonly Lazy<CacheManager> _lazyInstance = new Lazy<CacheManager>(() => new CacheManager());
        private ConcurrentDictionary<string, List<ClientSideComponent>> clientSideComponents;
        private ConcurrentDictionary<Guid, string> siteToComponentMapping;

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
        }

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

        public void ClearClientSideComponents()
        {
            clientSideComponents.Clear();
            siteToComponentMapping.Clear();
        }

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

    }
}
