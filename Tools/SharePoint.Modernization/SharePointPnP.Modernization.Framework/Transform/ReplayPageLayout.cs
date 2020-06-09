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
        private List<ReplayWebPartLocation> _currentPageLocations;
        private bool _hasLayoutChangedFromPrevious;

        /// <summary>
        /// Check if the layout has changed from previous
        /// </summary>
        public bool HasLayoutChangedFromPrevious { get {
                return _hasLayoutChangedFromPrevious;
            } 
        }

        /// <summary>
        /// Is in Page Replay Mode
        /// </summary>
        public bool IsPageReplayMode { get {
                return _isPageReplay;
            } 
        }

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

            _currentPageLocations = new List<ReplayWebPartLocation>();
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

                    if(this._replayPageCaptureData.ReplayWebPartLocations.Any(o=>o.ColumnFactor != o.MovedToColumnFactor))
                    {
                        this._hasLayoutChangedFromPrevious = true;
                    }
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
        /// Temporary store current transformation web part locations
        /// </summary>
        /// <param name="sourceWebPartType"></param>
        /// <param name="targetTypeId"></param>
        /// <param name="plannedRow"></param>
        /// <param name="plannedColumn"></param>
        /// <param name="plannedOrder"></param>
        /// <param name="sourceWebPartTitle"></param>
        /// <param name="sourceWebPartGroup"></param>
        public void StoreReplayWebParts(string sourceWebPartType, string targetTypeId, Guid targetInstanceId, int plannedRow, int plannedColumn, int plannedOrder, string sourceWebPartTitle, string sourceWebPartGroup)
        {
            _currentPageLocations.Add(new ReplayWebPartLocation
            {
                SourceWebPartType = sourceWebPartType,
                TargetWebPartTypeId = targetTypeId,
                Row = plannedRow,
                Column = plannedColumn,
                Order = plannedOrder,
                SourceWebPartTitle = sourceWebPartTitle,
                SourceGroupName = sourceWebPartGroup,
                TargetWebPartInstanceId = targetInstanceId
            });
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
        public ReplayWebPartLocation GetLayoutUpdatedPositionForWebPart(List<ReplayWebPartLocation> availableLocations, string sourceWebPartType, string targetTypeId, int plannedRow, int plannedColumn, int plannedOrder, string sourceWebPartTitle, string sourceGroupName)
        {
            if (this._isPageReplay)
            {
                // Use the next page web part source type, and target type, transform co-ordinates, if there is an adjustment then
                // New Replay Target page may need sections built prior to adjusting the web part locations
                // return the adjusted co-ordinates.
                var location = availableLocations.Where(
                    o => o.TargetWebPartTypeId == targetTypeId &&
                    o.SourceWebPartType == sourceWebPartType && 
                    o.Order == plannedOrder && 
                    o.Row == plannedRow && 
                    o.Column == plannedColumn && o.SourceWebPartTitle == sourceWebPartTitle && o.SourceGroupName == sourceGroupName).FirstOrDefault();

                //TODO: This needs to be smarter to encounter a block similar to this. e.g. Instance of if order not exact...
                // This implementation is likely to be unstable
                if(location == null)
                {
                    location = availableLocations.Where(
                        o => o.TargetWebPartTypeId == targetTypeId &&
                        o.SourceWebPartType == sourceWebPartType && 
                        o.Row == plannedRow && 
                        o.Column == plannedColumn && 
                        o.SourceGroupName == sourceGroupName && o.SourceWebPartTitle == sourceWebPartTitle).FirstOrDefault();
                }
                
                

                return location;
            }

            return default;

        }

        /// <summary>
        /// Apply any changes made to the captured page an apply this to the current transform
        /// </summary>
        public void ApplyLocationChanges(ClientSidePage clientSidePage)
        {
            if (this._isPageReplay) {

                var availableLocations = this._replayPageCaptureData.ReplayWebPartLocations;

                foreach (var location in _currentPageLocations)
                {
                    var result = GetLayoutUpdatedPositionForWebPart(availableLocations, location.SourceWebPartType, location.TargetWebPartTypeId, location.Row, location.Column, 
                        location.Order, location.SourceWebPartTitle, location.SourceGroupName);
                    if (result != null && result.CanUseMoveToLocation)
                    {
                        // This web part is found, remove from the left over web parts to detect changes to
                        availableLocations.Remove(result);
                       

                        var row = result.MovedToRow;
                        var column = result.MovedToColumn;
                        //currentOrder = result.MovedToOrder;
                        //var lastColumnOrder = LastColumnOrder(row, column);
                        var currentOrder = LastColumnOrder(clientSidePage, row, column);
                        var control = clientSidePage.Controls.FirstOrDefault(c => c.InstanceId == location.TargetWebPartInstanceId);
                        if (control != null)
                        {
                            var section = clientSidePage.Sections[row];
                            if (section != control.Section)
                            {
                                control.Move(section);
                            }
                            var canvasColumn = section.Columns[column];
                            if (canvasColumn != control.Column)
                            {
                                control.Move(canvasColumn);
                            }
                        }
                    }
                }

            }


            
        }

        /// <summary>
        /// Duplicates the page layout based on the previous page
        /// </summary>
        /// <param name="previousReplayCaptureData"></param>
        /// <param name="currentPage"></param>
        /// <returns>If the page has been duplicated</returns>
        public bool DuplicatePageLayout(ClientSidePage currentPage)
        {
            if (this._isPageReplay)
            {

                var previousReplayCaptureData = this._replayPageCaptureData;

                if (previousReplayCaptureData != null && !string.IsNullOrEmpty(previousReplayCaptureData.PageUrl))
                {
                    if (previousReplayCaptureData.PageLayoutName == this._replayPageCaptureData.PageLayoutName)
                    {
                        // Get the page
                        // TODO: Check the scenario where the page is in a folder
                        ClientSidePage previousClientSidepage = ClientSidePage.Load(this._targetContext, previousReplayCaptureData.PageUrl);
                        
                        // First drop all sections, ensure the sections are gone
                        currentPage.Sections.Clear();

                        foreach (var section in previousClientSidepage.Sections)
                        {
                            //Ensure an empty layout
                            currentPage.AddSection(section.Type, section.Order, section.ZoneEmphasis, section.VerticalSectionColumn?.VerticalSectionEmphasis);
                        }
                                                
                        return true;
                    }
                    else
                    {
                        //Previous capture data page layout incorrect, the position could be different, therefore reject
                        LogWarning("Cannot duplicate page - previous capture data page layout incorrect, the position could be different", LogStrings.Heading_ReplayPageLayout);
                    }
                }
                else
                {
                    //Log that you cannot replay without first recording the referenced transform
                    LogWarning("Cannot duplicate page - you cannot replay without first recording the capture transform", LogStrings.Heading_ReplayPageLayout);
                }
            }

            return false;
        }

        /// <summary>
        /// Get the last column order
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        private Int32 LastColumnOrder(ClientSidePage page, int row, int col)
        {
            var lastControl = page.Sections[row].Columns[col].Controls.OrderBy(p => p.Order).LastOrDefault();
            if (lastControl != null)
            {
                return lastControl.Order;
            }
            else
            {
                return -1;
            }
        }
    }
}
