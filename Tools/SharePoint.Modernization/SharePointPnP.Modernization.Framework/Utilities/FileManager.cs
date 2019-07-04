using SharePointPnP.Modernization.Framework.Entities;
using SharePointPnP.Modernization.Framework.Telemetry;
using SharePointPnP.Modernization.Framework.Transform;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SharePointPnP.Modernization.Framework.Utilities
{
    /// <summary>
    /// Class that's responsible for loading (mapping) files
    /// </summary>
    public class FileManager: BaseTransform
    {

        #region Construction
        public FileManager(IList<ILogObserver> logObservers = null): base()
        {
            //Register any existing observers
            if (logObservers != null)
            {
                foreach (var observer in logObservers)
                {
                    base.RegisterObserver(observer);
                }
            }
        }
        #endregion

        /// <summary>
        /// Loads a URL mapping file
        /// </summary>
        /// <param name="mappingFile">Path to the mapping file</param>
        /// <returns>A collection of URLMapping objects</returns>
        public List<UrlMapping> LoadUrlMappingFile(string mappingFile)
        {
            List<UrlMapping> urlMappings = new List<UrlMapping>();

            LogInfo(string.Format(LogStrings.LoadingUrlMappingFile, mappingFile), LogStrings.Heading_UrlRewriter);

            if (System.IO.File.Exists(mappingFile))
            {
                var lines = System.IO.File.ReadLines(mappingFile);

                if (lines.Count() > 0)
                {
                    string delimiter = this.DetectDelimiter(lines);

                    foreach(var line in lines)
                    {
                        var split = line.Split(new string[] { delimiter }, StringSplitOptions.RemoveEmptyEntries);

                        if (split.Length == 2)
                        {
                            string fromUrl = split[0];
                            string toUrl = split[1];

                            if (!string.IsNullOrEmpty(fromUrl) && !string.IsNullOrEmpty(toUrl))
                            {
                                urlMappings.Add(new UrlMapping() { SourceUrl = fromUrl, TargetUrl = toUrl });
                                LogDebug(string.Format(LogStrings.UrlMappingLoaded, fromUrl, toUrl), LogStrings.Heading_UrlRewriter);
                            }
                        }
                    }
                }
            }
            else
            {
                LogError(string.Format(LogStrings.Error_UrlMappingFileNotFound, mappingFile), LogStrings.Heading_UrlRewriter);
                throw new Exception(string.Format(LogStrings.Error_UrlMappingFileNotFound, mappingFile));
            }

            return urlMappings;
        }

        #region Helper methods
        private string DetectDelimiter(IEnumerable<string> lines)
        {
            if (lines.First().IndexOf(',') > 0)
            {
                return ",";
            }
            else if (lines.First().IndexOf(';') > 0)
            {
                return ";";
            }

            return "";
        }
        #endregion
    }
}
