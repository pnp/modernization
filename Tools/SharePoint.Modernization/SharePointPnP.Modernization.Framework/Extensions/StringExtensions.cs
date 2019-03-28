using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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

    }
}
