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

        private bool _isReplayEnabled;
        private string _referencePageName;
        private ClientContext _targetContext;
        private BaseTransformationInformation _transformationInformation;
        private bool _isPageCapture;

        // Constructor
        public ReplayPageLayout(BaseTransformationInformation transformationInformation, ClientContext targetContext, IList<ILogObserver> logObservers = null)
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
            this._referencePageName = transformationInformation.ReplayLayoutChangeBasedOn;
            this._transformationInformation = transformationInformation;
            this._isPageCapture = transformationInformation.IsReplayCapture;
            if (!string.IsNullOrEmpty(_referencePageName))
            {
                _isReplayEnabled = true;
            }
        }

        public void StoreLocation(ReplayWebPartLocation location)
        {
            if (_isPageCapture)
            {
                var pageName = $"{this._transformationInformation.Folder}{this._transformationInformation.TargetPageName}";
                location.PageUrl = pageName;

                //TODO: Capture the layout mapping name.
                CacheManager.Instance.StoreWebPartLocationsForTargetPage(location);
            }
        }

        public void RetrieveTargetPageWebPartLayout()
        {
            throw new NotImplementedException();

            // Design: the content transformator records the type, order, row and column. If we recommend the row/column are 0.
            // We can match the web part instance by the type and order. 
            // In capture mode, we record the changes from the original transformed location to the new location. 
            // On replay when the same combinations appear for the order and for the page the layout target is updated.
            // If we limit the transform to only the source page layout, tag the file, then we can protect against other layouts 
            // or vastly different combinations of pages.
        }

        public void GetLayoutPositionForWebPart()
        {
            throw new NotImplementedException();
        }

    }
}
