using Microsoft.SharePoint.Client;
using SharePointPnP.Modernization.Framework.Entities;
using System;
using System.IO;
using System.Linq;

namespace SharePointPnP.Modernization.Framework.Transform
{
    /// <summary>
    /// Class for operations for transferring the assets over to the target site collection
    /// </summary>
    public class AssetTransfer
    {
        //Plan:
        //  Detect for referenced assets within the web parts
        //  Referenced assets should only be files e.g. not aspx pages and located in the pages, site pages libraries
        //  Ensure the referenced assets exist within the same site collection/web according to the level of transformation
        //  With the modern destination, locate assets in the site assets library with in a folder using the same naming convention as SharePoint Comm Sites
        //  Add copy assets method to transfer the files to target site collection
        //  Store a dictionary of copied assets to update the URLs of the transferred web parts
        //  Phased approach for this: 
        //      Image Web Parts
        //      Text Web Parts with inline images (need to determine how they are handled)
        //      TBC - expanded as testing progresses

        private ClientContext _sourceClientContext;
        private ClientContext _targetClientContext;

        /// <summary>
        /// Constructor for the asset transfer class
        /// </summary>
        /// <param name="source">Source connection to SharePoint</param>
        /// <param name="target">Target connection to SharePoint</param>
        public AssetTransfer(ClientContext source, ClientContext target)
        {
            if (source == null || target == null)
            {
                throw new ArgumentNullException("One or more client context is null");
            }

            _sourceClientContext = source;
            _targetClientContext = target;
        }

        /// <summary>
        /// Main entry point to perform the series of operations to transfer related assets
        /// </summary>
        public string TransferAsset(string sourceAssetRelativeUrl, string pageFileName)
        {

            // Deep validation of urls
            var isValid = ValidateAssetInSupportedLocation(sourceAssetRelativeUrl) && !string.IsNullOrEmpty(pageFileName);

            // Check the string is not null
            if (!string.IsNullOrEmpty(sourceAssetRelativeUrl) && isValid)
            {

                // Check the target library exists
                string targetFolderServerRelativeUrl = EnsureDestination(pageFileName);
                // Read in a preferred location

                // Check that the target location, a file does not already exist

                // Copy the asset file
                string newLocationUrl = CopyAssetToTargetLocation(sourceAssetRelativeUrl, targetFolderServerRelativeUrl);

                // Store a reference in the cache manager - ensure a test exists with multiple identical web parts
                // Return the URL for the new location
                return newLocationUrl;
            }
            
            // Fall back to send back the same link
            return sourceAssetRelativeUrl;
        }
               

        /// <summary>
        /// Checks if the URL is located in a supported location
        /// </summary>
        public bool ValidateAssetInSupportedLocation(string sourceUrl)
        {
            //  Referenced assets should only be files e.g. 
            //      not aspx pages 
            //      located in the pages, site pages libraries
            
            var fileExtension = Path.GetExtension(sourceUrl);
            var containsBlockedExtension = Constants.BlockedAssetFileExtensions.Any(o => o == fileExtension.Replace(".",""));
            if (containsBlockedExtension)
            {
                return false;
            }

            //  Ensure the referenced assets exist within the same site collection/web according to the level of transformation
            var sourceWebUrl = _sourceClientContext.Web.EnsureProperty(w => w.ServerRelativeUrl);
            var targetWebUrl = _targetClientContext.Web.EnsureProperty(w => w.ServerRelativeUrl);

            if ( sourceWebUrl == targetWebUrl)
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// Ensure the site assets and page sub-folder exists in the target location
        /// </summary>
        public string EnsureDestination(string pageFileName)
        {
            // In this method we need to calculate the target location from the following factors
            //  Target Site Context + Site Assets Library + Folder (if located in or calculate based on SP method)
            //  Check the libary and folder exists in the target site collection
            //  Currently this method ignores anything from the source, will probabily need an override or params for target location

            // Ensure the Site Assets library exists
            var siteAssetsLibrary = this.EnsureSiteAssetsLibrary();
            var sitePagesFolder = siteAssetsLibrary.RootFolder.EnsureFolder("SitePages");

            var friendlyFolder = ConvertFileToFolderFriendlyName(pageFileName);
            var pageFolder = sitePagesFolder.EnsureFolder(friendlyFolder);

            return pageFolder.EnsureProperty(o => o.ServerRelativeUrl);
            
        }

        /// <summary>
        /// Create a site assets library
        /// </summary>
        public List EnsureSiteAssetsLibrary()
        {
            // Use a PnP Provisioning template to create a site assets library
            // We cannot assume the SiteAssets library exists, in the case of vanilla communication sites - provision a new library if none exists
            // If a site assets library exist, add a folder, into the library using the same format as SharePoint uses for creating sub folders for pages

            //Ensure that the Site Assets library is created using the out of the box creation mechanism
            //Site Assets that are created using the EnsureSiteAssetsLibrary method slightly differ from
            //default Document Libraries. See issue 512 (https://github.com/SharePoint/PnP-Sites-Core/issues/512)
            //for details about the issue fixed by this approach.
            var createdList = this._targetClientContext.Web.Lists.EnsureSiteAssetsLibrary();
            //Check that Title and Description have the correct values
            this._targetClientContext.Web.Context.Load(createdList, l => l.Title, l => l.RootFolder);
            this._targetClientContext.Web.Context.ExecuteQueryRetry();

            return createdList;
        }

        /// <summary>
        /// Copy the file from the source to the target location
        /// </summary>
        /// <param name="sourceFileUrl"></param>
        /// <param name="targetLocationUrl"></param>
        /// <remarks>
        ///     Based on the documentation: https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/upload-large-files-sample-app-for-sharepoint
        /// </remarks>
        public string CopyAssetToTargetLocation(string sourceFileUrl, string targetLocationUrl, int fileChunkSizeInMB = 3)
        {
            // This copies the latest version of the asset to the target site collection
            // Going to need to add a bunch of checks to ensure the target file exists
            // TODO: Need to add a ton of error and logging here

            // Each sliced upload requires a unique ID.
            Guid uploadId = Guid.NewGuid();
            // Calculate block size in bytes.
            int blockSize = fileChunkSizeInMB * 1024 * 1024;
            bool fileOverwrite = true;

            // Get the file from SharePoint
            var sourceAssetFile = _sourceClientContext.Web.GetFileByServerRelativeUrl(sourceFileUrl);
            ClientResult<System.IO.Stream> sourceAssetFileData = sourceAssetFile.OpenBinaryStream();

            _sourceClientContext.Load(sourceAssetFile);
            _sourceClientContext.ExecuteQueryRetry();

            using (Stream sourceFileStream = sourceAssetFileData.Value)
            {

                string fileName = sourceAssetFile.Name;

                // New File object.
                Microsoft.SharePoint.Client.File uploadFile;

                // Get the information about the folder that will hold the file.
                // Add the file to the target site
                Folder targetFolder = _targetClientContext.Web.GetFolderByServerRelativeUrl(targetLocationUrl);
                _targetClientContext.Load(targetFolder);
                _targetClientContext.ExecuteQueryRetry();

                // Get the file size
                long fileSize = sourceFileStream.Length;

                // Process with two approaches
                if (fileSize <= blockSize)
                {

                    // Use regular approach.

                    FileCreationInformation fileInfo = new FileCreationInformation();
                    fileInfo.ContentStream = sourceFileStream;
                    fileInfo.Url = fileName;
                    fileInfo.Overwrite = fileOverwrite;

                    uploadFile = targetFolder.Files.Add(fileInfo);
                    _targetClientContext.Load(uploadFile);
                    _targetClientContext.ExecuteQuery();

                    // Return the file object for the uploaded file.
                    return uploadFile.EnsureProperty(o => o.ServerRelativeUrl);

                }
                else
                {
                    // Use large file upload approach.
                    ClientResult<long> bytesUploaded = null;

                    using (BinaryReader br = new BinaryReader(sourceFileStream))
                    {
                        byte[] buffer = new byte[blockSize];
                        Byte[] lastBuffer = null;
                        long fileoffset = 0;
                        long totalBytesRead = 0;
                        int bytesRead;
                        bool first = true;
                        bool last = false;

                        // Read data from file system in blocks. 
                        while ((bytesRead = br.Read(buffer, 0, buffer.Length)) > 0)
                        {
                            totalBytesRead = totalBytesRead + bytesRead;

                            // You've reached the end of the file.
                            if (totalBytesRead == fileSize)
                            {
                                last = true;
                                // Copy to a new buffer that has the correct size.
                                lastBuffer = new byte[bytesRead];
                                Array.Copy(buffer, 0, lastBuffer, 0, bytesRead);
                            }

                            if (first)
                            {
                                using (MemoryStream contentStream = new MemoryStream())
                                {
                                    // Add an empty file.
                                    FileCreationInformation fileInfo = new FileCreationInformation();
                                    fileInfo.ContentStream = contentStream;
                                    fileInfo.Url = fileName;
                                    fileInfo.Overwrite = fileOverwrite;
                                    uploadFile = targetFolder.Files.Add(fileInfo);

                                    // Start upload by uploading the first slice. 
                                    using (MemoryStream s = new MemoryStream(buffer))
                                    {
                                        // Call the start upload method on the first slice.
                                        bytesUploaded = uploadFile.StartUpload(uploadId, s);
                                        _targetClientContext.ExecuteQueryRetry();
                                        // fileoffset is the pointer where the next slice will be added.
                                        fileoffset = bytesUploaded.Value;
                                    }

                                    // You can only start the upload once.
                                    first = false;
                                }
                            }
                            else
                            {
                                // Get a reference to your file.
                                var fileUrl = targetFolder.ServerRelativeUrl + System.IO.Path.AltDirectorySeparatorChar + fileName;
                                uploadFile = _targetClientContext.Web.GetFileByServerRelativeUrl(fileUrl);

                                if (last)
                                {
                                    // Is this the last slice of data?
                                    using (MemoryStream s = new MemoryStream(lastBuffer))
                                    {
                                        // End sliced upload by calling FinishUpload.
                                        uploadFile = uploadFile.FinishUpload(uploadId, fileoffset, s);
                                        _targetClientContext.ExecuteQuery();

                                        // Return the file object for the uploaded file.
                                        return fileUrl;
                                    }
                                }
                                else
                                {
                                    using (MemoryStream s = new MemoryStream(buffer))
                                    {
                                        // Continue sliced upload.
                                        bytesUploaded = uploadFile.ContinueUpload(uploadId, fileoffset, s);
                                        _targetClientContext.ExecuteQuery();
                                        // Update fileoffset for the next slice.
                                        fileoffset = bytesUploaded.Value;
                                    }
                                }
                            }
                        }
                    }

                }

            }

            return null;
        }

        /// <summary>
        /// Stores an asset transfer reference
        /// </summary>
        /// <param name="assetTransferReferenceEntity"></param>
        /// <param name="update"></param>
        public void StoreAssetTransferReference(AssetTransferReferenceEntity assetTransferReferenceEntity, bool? update)
        {
            // Using the Cache Manager store the asset transfer references
            // If update - treat the source URL as unique, if multiple web parts reference to this, then it will still refer to the single resource
            throw new NotImplementedException();
        }

        /// <summary>
        /// Get all asset transfer references
        /// </summary>
        public void GetAssetTransferReferences()
        {
            // Using the Cache Manager retrieve asset transfer references (all)
            throw new NotImplementedException();
        }

        /// <summary>
        /// Converts the file name into a friendly format
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public string ConvertFileToFolderFriendlyName(string fileName)
        {
            // This is going to need some heavy testing
            var justFileName = Path.GetFileNameWithoutExtension(fileName);
            var friendlyName = justFileName.Replace(" ", "-");
            return friendlyName;
        }

    }
}
