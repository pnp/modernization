using System;
using System.Linq;

namespace SharePointPnP.Modernization.Framework.Extensions
{
    public static class StringExtensions
    {
        /// <summary>
        /// Determines if a string exists in another string regardless of casing
        /// </summary>
        /// <param name="value">original string</param>
        /// <param name="comparedWith">string to compare with</param>
        /// <param name="stringComparison">optional comparison mode</param>
        /// <returns></returns>
        public static bool ContainsIgnoringCasing(this string value, string comparedWith, StringComparison stringComparison = StringComparison.InvariantCultureIgnoreCase)
        {
            return value.IndexOf(comparedWith, stringComparison) >= 0;
        }

        /// <summary>
        /// Prepends string to another including null checking
        /// </summary>
        /// <param name="value"></param>
        /// <param name="prependString"></param>
        /// <returns></returns>
        public static string PrependIfNotNull(this string value, string prependString)
        {
            if (!string.IsNullOrEmpty(value) && !value.ContainsIgnoringCasing(prependString))
            {
                value = prependString + value;
            }

            return value; // Fall back
        }


        /// <summary>
        /// Removes a relative section of by string where context not available
        /// </summary>
        /// <param name="value"></param>
        /// <param name="seperator"></param>
        /// <param name="instanceFrom"></param>
        /// <returns></returns>
        public static string StripRelativeUrlSectionString(this string value)
        {
            if (!string.IsNullOrEmpty(value))
            {
                var sitesColl = "/sites/";
                var teamsColl = "/teams/";
                var containsSites = value.IndexOf(sitesColl, StringComparison.InvariantCultureIgnoreCase);
                var containsTeams = value.IndexOf(teamsColl, StringComparison.InvariantCultureIgnoreCase);
                if (containsSites > -1 || containsTeams > -1)
                {
                    if (containsSites > -1)
                    {
                        var result = value.TrimStart(sitesColl.ToCharArray());
                        if (result.IndexOf('/') > -1)
                        {
                            return result.Substring(result.IndexOf('/'));
                        }
                    }
                    else if (containsTeams > -1)
                    {
                        var result = value.TrimStart(teamsColl.ToCharArray());
                        if (result.IndexOf('/') > -1)
                        {
                            return result.Substring(result.IndexOf('/'));
                        }
                    }
                }
            }

            return value;
        }

        /// <summary>
        /// Gets base url from string
        /// </summary>
        /// <param name="sourceSite"></param>
        /// <returns></returns>
        public static string GetBaseUrl(this string url)
        {
            try
            {
                if (!string.IsNullOrEmpty(url) && (url.ContainsIgnoringCasing("https://") || url.ContainsIgnoringCasing("http://")))
                {
                    Uri siteUri = new Uri(url);
                    string host = $"{siteUri.Scheme}://{siteUri.DnsSafeHost}";
                    return host;
                }
            }
            catch (Exception)
            {
                //Swallow
            }

            return string.Empty;
        }

        /// <summary>
        /// Get type in short form
        /// </summary>
        /// <param name="typeValue"></param>
        /// <returns></returns>
        public static string GetTypeShort(this string typeValue)
        {
            string name = typeValue;
            var typeSplit = typeValue.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            if (typeSplit.Length > 0)
            {
                name = typeSplit[0];
            }

            return $"{name}";
        }

        /// <summary>
        /// Gets classname from type
        /// </summary>
        /// <param name="typeName"></param>
        /// <returns></returns>
        public static string InferClassNameFromNameSpace(this string typeName)
        {
            string shortType = typeName;
            string className = string.Empty;
            if (typeName.Contains(","))
            {
                shortType = typeName.GetTypeShort();
            }

            var typeShortSplit = shortType.Split(new string[] { "." }, StringSplitOptions.RemoveEmptyEntries);
            if (typeShortSplit.Length > 0)
            {
                className = typeShortSplit.Last();
            }

            return className;
        }
    }
}
