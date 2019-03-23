using Microsoft.SharePoint.Client;
using SharePointPnP.Modernization.Framework.Functions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Pages;
using SharePointPnP.Modernization.Framework.Telemetry;

namespace SharePointPnP.Modernization.Framework.Publishing
{
    public class PublishingBuiltIn: FunctionsBase
    {
        private ClientContext sourceClientContext;

        #region Construction
        /// <summary>
        /// Instantiates the base builtin function library
        /// </summary>
        /// <param name="pageClientContext">ClientContext object for the site holding the page being transformed</param>
        /// <param name="sourceClientContext">The ClientContext for the source </param>
        /// <param name="clientSidePage">Reference to the client side page</param>
        public PublishingBuiltIn(ClientContext sourceClientContext, IList<ILogObserver> logObservers = null) : base(sourceClientContext)
        {
            // This is an optional property, in cross site transfer the two contexts would be different.
            this.sourceClientContext = sourceClientContext;

            if (logObservers != null)
            {
                foreach (var observer in logObservers)
                {
                    base.RegisterObserver(observer);
                }
            }
        }
        #endregion

        #region Image functions
        public string ToImageUrl(string htmlImage)
        {
            return "/sites/modernizationtestportal/PublishingImages/extra6.jpg";
        }

        public string ToImageAltText(string htmlImage)
        {
            return "";
        }
        #endregion
    }
}
