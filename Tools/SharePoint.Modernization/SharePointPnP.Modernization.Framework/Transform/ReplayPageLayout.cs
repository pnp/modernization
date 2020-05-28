using Microsoft.SharePoint.Client;
using SharePointPnP.Modernization.Framework.Telemetry;
using OfficeDevPnP.Core.Pages;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SharePointPnP.Modernization.Framework.Cache;
using SharePointPnP.Modernization.Framework.Entities;

namespace SharePointPnP.Modernization.Framework.Transform
{
    /// <summary>
    /// 
    /// </summary>
    public class ReplayPageLayout : BaseTransform
    {
        public const string TextWebPart = "Text";
        private ClientContext _targetContext;
        private BaseTransformationInformation _transformationInformation;
        private bool _isPageCapture;
        private bool _isPageReplay;
        private ReplayPageCaptureData _replayPageCaptureData;

        /// <summary>
        /// Constructor for the replay page layout class
        /// </summary>
        /// <param name="transformationInformation"></param>
        /// <param name="targetContext"></param>
        /// <param name="logObservers"></param>
        public ReplayPageLayout(BaseTransformationInformation transformationInformation, ClientContext targetContext, string pageLayoutName, IList<ILogObserver> logObservers = null)
        {
            // Hookup logging
            if (logObservers != null)
            {
                foreach (var observer in logObservers)
                {
                    base.RegisterObserver(observer);
                }
            }

            this._targetContext = targetContext;
            this._transformationInformation = transformationInformation;
            this._isPageCapture = transformationInformation.IsReplayCapture;
            this._isPageReplay = transformationInformation.IsReplayLayout;
            this._replayPageCaptureData = new ReplayPageCaptureData()
            {
                PageName = this._transformationInformation.TargetPageName,
                PageUrl = $"{this._transformationInformation.Folder}{this._transformationInformation.TargetPageName}",
                PageLayoutName = pageLayoutName
            };
        }

        /// <summary>
        /// Check and loads the capture data from the cache, if the transformation includes page layout replay mode
        /// </summary>
        public void CheckAndLoadCaptureDataFromCache()
        {
            if (this._isPageReplay)
            {
                var captureData = CacheManager.Instance.GetReplayCaptureData();
                if(captureData != null && !string.IsNullOrEmpty(captureData.PageUrl))
                {
                    this._replayPageCaptureData = ScanTargetPageWebPartLocationsForChanges(captureData);
                }
                else
                {
                    //Cancel replay mode
                    this._isPageReplay = false;
                    LogWarning("You cannot replay without first recording the capture transform", LogStrings.Heading_ReplayPageLayout);
                }
            }
        }

        /// <summary>
        /// Store capture data in the cache - essential this is triggered after save operation
        /// </summary>
        public void StoreCaptureData()
        {
            // To Cache
            // Future option to save to JSON file and persist :-)
            if (_isPageCapture)
            {
                //Store the capture data in memory
                CacheManager.Instance.SetReplayCaptureData(this._replayPageCaptureData);
            }
        }

        /// <summary>
        /// This method is used to capture the positions and locations during the initial transform
        /// </summary>
        /// <param name="location"></param>
        public void StoreInitialWebPartLocations(ReplayWebPartLocation location)
        {
            if (_isPageCapture)
            {
                this._replayPageCaptureData.ReplayWebPartLocations.Add(location);
            }
        }

        /// <summary>
        /// This method will check the modified target page for chnages in web part location, and deletions
        /// </summary>
        /// <remarks>
        ///     Design: the content transformator records the type, order, row and column. If we recommend the row/column are 0.
        ///     We can match the web part instance by the type and order. 
        ///     In capture mode, we record the changes from the original transformed location to the new location. 
        ///     On replay when the same combinations appear for the order and for the page the layout target is updated.
        ///     If we limit the transform to only the source page layout, tag the file, then we can protect against other layouts 
        ///     or vastly different combinations of pages.
        /// </remarks>
        public ReplayPageCaptureData ScanTargetPageWebPartLocationsForChanges(ReplayPageCaptureData previousReplayCaptureData)
        {
            
            // Connect to target page - this action must happen AFTER save and user has updated the page
            // OR log the marked scan file - then replay based on that file 
            
            if(previousReplayCaptureData != null && !string.IsNullOrEmpty(previousReplayCaptureData.PageUrl))
            {
                if(previousReplayCaptureData.PageLayoutName == this._replayPageCaptureData.PageLayoutName)
                {
                    // Get the page
                    // TODO: Check the scenario where the page is in a folder
                    ClientSidePage previousClientSidepage = ClientSidePage.Load(this._targetContext, previousReplayCaptureData.PageUrl);

                    int sectionIndex = 0;

                    // The answer to get the positional data within teh page is to loop through the sections, 
                    // then the columns to find the controls, not controls then sections...

                    // Find the components - by ID
                    foreach (var section in previousClientSidepage.Sections)
                    {   
                        int columnIndex = 0;
                        

                        foreach(var column in section.Columns)
                        {
                            foreach(var control in column.Controls)
                            {
                                // This should compensate if the user removes the web part or transform cleanup

                                var location = previousReplayCaptureData.ReplayWebPartLocations.Where(o => o.TargetWebPartInstanceId == control.InstanceId).FirstOrDefault();
                                if (location != default)
                                {
                                    location.MovedToOrder = control.Order;
                                    location.MovedToRow = sectionIndex;
                                    location.MovedToColumn = columnIndex;
                                    location.MovedToColumnFactor = column.ColumnFactor;
                                    location.MovedToIsVerticalColumn = column.IsVerticalSectionColumn;
                                    location.MovedToRowZoneEmphesis = section.ZoneEmphasis;
                                }
                            }
                            
                            columnIndex++;
                        }

                        sectionIndex++;
                    }

                    return previousReplayCaptureData; //Includes the location updates
                }
                else
                {
                    //Previous capture data page layout incorrect, the position could be different, therefore reject
                    LogWarning("Previous capture data page layout incorrect, the position could be different", LogStrings.Heading_ReplayPageLayout);
                }
            }
            else
            {
                //Log that you cannot replay without first recording the referenced transform
                LogWarning("You cannot replay without first recording the capture transform", LogStrings.Heading_ReplayPageLayout);
            }

            return default;
        }

        /// <summary>
        /// This method will check for changes in the layout by the user and adjust the planned mapped locations
        /// </summary>
        public ReplayWebPartLocation GetLayoutUpdatedPositionForWebPart(string sourceWebPartType, string targetTypeId, int plannedRow, int plannedColumn, int plannedOrder)
        {
            if (this._isPageReplay)
            {
                // Use the next page web part source type, and target type, transform co-ordinates, if there is an adjustment then
                // New Replay Target page may need sections built prior to adjusting the web part locations
                // return the adjusted co-ordinates.
                var location = this._replayPageCaptureData.ReplayWebPartLocations.Where(o => o.TargetWebPartTypeId == targetTypeId &&
                    o.SourceWebPartType == sourceWebPartType && o.Order == plannedOrder && o.Row == plannedRow && o.Column == plannedColumn).FirstOrDefault();

                //TODO: Switch out the location data

                return location;
            }

            return default;

        }
    }
}
