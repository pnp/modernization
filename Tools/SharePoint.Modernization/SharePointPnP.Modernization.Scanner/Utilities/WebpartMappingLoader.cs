using System.IO;

namespace SharePoint.Modernization.Scanner.Utilities
{
    /// <summary>
    /// Class to support the embedded webpartmapping.xml file
    /// </summary>
    public static class WebpartMappingLoader
    {
        /// <summary>
        /// Load the webpartmapping file from the embedded resources
        /// </summary>
        /// <param name="fileName">Fully qualified path to file</param>
        /// <returns>String contents</returns>
        public static string LoadFile(string fileName)
        {
            var fileContent = "";
            using (Stream stream = typeof(WebpartMappingLoader).Assembly.GetManifestResourceStream(fileName))
            {
                using (StreamReader reader = new StreamReader(stream))
                {
                    fileContent = reader.ReadToEnd();
                }
            }

            return fileContent;
        }

        /// <summary>
        /// Transforms a string into a stream
        /// </summary>
        /// <param name="s">String to transform</param>
        /// <returns>Stream</returns>
        public static Stream GenerateStreamFromString(string s)
        {
            var stream = new MemoryStream();
            var writer = new StreamWriter(stream);
            writer.Write(s);
            writer.Flush();
            stream.Position = 0;
            return stream;
        }
    }
}
