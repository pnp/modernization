using Microsoft.SharePoint.Client;
using SharePointPnP.Modernization.Framework.Functions;
using System;
using OfficeDevPnP.Core.Pages;
using SharePointPnP.Modernization.Framework.Transform;
using SharePointPnP.Modernization.Framework.Telemetry;
using System.Collections.Generic;
using Microsoft.Graph;

namespace SharePointPnP.Modernization.Framework.SampleAddOn
{
    public class MyCustomFunctions: FunctionsBase
    {
        private ClientContext sourceClientContext;
        private ClientSidePage clientSidePage;
        private BaseTransformationInformation baseTransformationInformation;
        private UrlTransformator urlTransformator;
        private UserTransformator userTransformator;

        #region Construction
        public MyCustomFunctions(BaseTransformationInformation baseTransformationInformation, ClientContext pageClientContext, ClientContext sourceClientContext,ClientSidePage clientSidePage, IList<ILogObserver> logObservers) : base(pageClientContext)
        {
            if (logObservers != null)
            {
                foreach (var observer in logObservers)
                {
                    base.RegisterObserver(observer);
                }
            }

            // This is an optional property, in cross site transfer the two contexts would be different.
            this.sourceClientContext = sourceClientContext;
            this.clientSidePage = clientSidePage;
            this.baseTransformationInformation = baseTransformationInformation;
            this.urlTransformator = new UrlTransformator(baseTransformationInformation, this.sourceClientContext, this.clientContext, base.RegisteredLogObservers);
            this.userTransformator = new UserTransformator(baseTransformationInformation, this.sourceClientContext, this.clientContext, base.RegisteredLogObservers);
        }
        #endregion

        public string MyListAddServerRelativeUrl(Guid listId)
        {
            try { 
            if (listId == Guid.Empty)
            {
                return "";
            }
            else
            {
                var list = this.clientContext.Web.GetListById(listId);
                list.EnsureProperty(p => p.RootFolder).EnsureProperty(p => p.ServerRelativeUrl);
                return list.RootFolder.ServerRelativeUrl;
            }
            }
            catch (Exception ex)
            {
                LogError("MyListAddServerRelativeUrl", "MyCustomFunctions", ex);
                return string.Empty;
            }
        }

        public string SplitInput(string input, string sectionNumber)
        {
            // Parse the received html content and return the part you need
            return $"Part {sectionNumber}";
        }


    }
}
