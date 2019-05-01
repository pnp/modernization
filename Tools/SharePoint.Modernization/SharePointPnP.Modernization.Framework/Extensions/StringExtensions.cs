using System;

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
            if (!value.ContainsIgnoringCasing(prependString) && !string.IsNullOrEmpty(value))
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
            var siteColl = "/sites/";
            var containsSites = value.IndexOf(siteColl, StringComparison.InvariantCultureIgnoreCase);
            if (containsSites > -1)
            {

                var result = value.TrimStart(siteColl.ToCharArray());
                if (result.IndexOf('/') > -1)
                {
                    return result.Substring(result.IndexOf('/'));
                }

            }

            return value;
        }
    }
}
