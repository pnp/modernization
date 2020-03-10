using Microsoft.SharePoint.Client;
using SharePointPnP.Modernization.Framework.Functions;
using System;

namespace SharePointPnP.Modernization.Framework.SampleAddOn
{
    public class MyCustomFunctions: FunctionsBase
    {
        #region Construction
        public MyCustomFunctions(ClientContext clientContext) : base(clientContext)
        {
        }
        #endregion

        public string MyListAddServerRelativeUrl(Guid listId)
        {
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

        public string SplitInput(string input, string sectionNumber)
        {
            // Parse the received html content and return the part you need
            return $"Part {sectionNumber}";
        }


    }
}
