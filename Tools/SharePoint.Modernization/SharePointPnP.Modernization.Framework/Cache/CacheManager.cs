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
    public sealed class CacheManager
    {
        private static readonly Lazy<CacheManager> _lazyInstance = new Lazy<CacheManager>(() => new CacheManager());
        private ConcurrentDictionary<string, List<ClientSideComponent>> clientSideComponents;
        private ConcurrentDictionary<Guid, string> siteToComponentMapping;

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

        public List<ClientSideComponent> GetClientSideComponents(ClientSidePage page)
        {
            Guid siteId = page.Context.Site.Id;

            if (siteToComponentMapping.ContainsKey(siteId))
            {
                // Components are cached for this site, get the component key
                if (siteToComponentMapping.TryGetValue(siteId, out string componentKey))
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
            if (siteToComponentMapping.TryAdd(siteId, componentKeyToCache))
            {
                if (clientSideComponents.ContainsKey(componentKeyToCache))
                {
                    return componentsToAdd;
                }
                else if (clientSideComponents.TryAdd(componentKeyToCache, componentsToAdd))
                {
                    return componentsToAdd;
                }
            }

            throw new Exception("Failed to load client side compenents from the page or from cache");
        }

        static string Sha256(string randomString)
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
