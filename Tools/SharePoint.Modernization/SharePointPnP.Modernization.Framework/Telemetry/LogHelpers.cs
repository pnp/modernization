using SharePointPnP.Modernization.Framework.Transform;
using System;
using System.Collections.Generic;
using System.Text;

namespace SharePointPnP.Modernization.Framework.Telemetry
{
    public static class LogHelpers
    {
        /// <summary>
        /// Converts boolean value to Yes/No string
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static string ToYesNoString(this bool value)
        {
            return value ? "Yes" : "No";
        }

        /// <summary>
        /// Formats a string that has the format ThisIsAClassName and formats in a friendly way
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static string FormatAsFriendlyTitle(this string value)
        {
            var charArr = value.ToCharArray();
            var result = new StringBuilder();
            for (var i = 0; i < charArr.Length; i++)
            {
                if (char.IsUpper(charArr[i]))
                {
                    result.Append($" {charArr[i]}");
                }
                else
                {
                    result.Append(charArr[i]);
                }
            }

            // Convert to string and remove space at start
            return result.ToString().TrimStart(' ');
        }

        /// <summary>
        /// Use reflection to read the object properties and detail the values
        /// </summary>
        /// <param name="pti">PageTransformationInformation object</param>
        /// <returns></returns>
        public static List<LogEntry> DetailSettingsAsLogEntries(this PageTransformationInformation pti)
        {
            List<LogEntry> logs = new List<LogEntry>();

            try
            {

                var properties = pti.GetType().GetProperties();
                foreach (var property in properties)
                {
                    if (property.PropertyType == typeof(String) ||
                        property.PropertyType == typeof(bool))
                    {
                        logs.Add(new LogEntry() { Heading = LogStrings.Heading_PageTransformationInfomation,
                            Message = $"{property.Name.FormatAsFriendlyTitle()} {LogStrings.KeyValueSeperatorToken} {property.GetValue(pti)}" });
                    }
                }
            }
            catch (Exception ex)
            {
                logs.Add(new LogEntry() { Message = "Failed to convert object properties for reporting", Exception = ex, Heading = LogStrings.Heading_PageTransformationInfomation });
            }
            
            return logs;

        }
    }
}
