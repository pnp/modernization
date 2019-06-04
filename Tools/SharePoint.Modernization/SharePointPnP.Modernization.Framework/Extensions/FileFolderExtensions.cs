using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointPnP.Modernization.Framework.Extensions
{
    public static class FileFolderExtensions
    {

        /// <summary>
        /// Returns a file as string
        /// </summary>
        /// <param name="web">The Web to process</param>
        /// <param name="serverRelativeUrl">The server relative URL to the file</param>
        /// <returns>The file contents as a string</returns>
        /// <remarks>#
        /// 
        ///     Based on https://github.com/SharePoint/PnP-Sites-Core/blob/master/Core/OfficeDevPnP.Core/Extensions/FileFolderExtensions.cs
        ///     Modified to force onpremises support
        ///     
        /// </remarks>
        public static string GetFileByServerRelativeUrlAsString(this Web web, string serverRelativeUrl)
        {

            var file = web.GetFileByServerRelativeUrl(serverRelativeUrl);
            
            web.Context.Load(file);
            web.Context.ExecuteQueryRetry();

            ClientResult<Stream> stream = file.OpenBinaryStream();

            web.Context.ExecuteQueryRetry();

            string returnString = string.Empty;

            using (Stream memStream = new MemoryStream())
            {
                CopyStream(stream.Value, memStream);
                memStream.Position = 0;

                StreamReader reader = new StreamReader(memStream);
                returnString = reader.ReadToEnd();
            }

            return returnString;
        }

        private static void CopyStream(Stream source, Stream destination)
        {
            byte[] buffer = new byte[32768];
            int bytesRead;

            do
            {
                bytesRead = source.Read(buffer, 0, buffer.Length);
                destination.Write(buffer, 0, bytesRead);

            } while (bytesRead != 0);

        }
    }
}
