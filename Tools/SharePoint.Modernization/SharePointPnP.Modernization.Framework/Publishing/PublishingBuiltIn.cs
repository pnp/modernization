using AngleSharp.Dom;
using AngleSharp.Parser.Html;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using SharePointPnP.Modernization.Framework.Functions;
using SharePointPnP.Modernization.Framework.Telemetry;
using System.Collections.Generic;
using System.Linq;

namespace SharePointPnP.Modernization.Framework.Publishing
{
    public class PublishingBuiltIn: FunctionsBase
    {
        private ClientContext sourceClientContext;
        private ClientContext targetClientContext;
        private HtmlParser parser;

        #region Construction
        /// <summary>
        /// Instantiates the base builtin function library
        /// </summary>
        /// <param name="pageClientContext">ClientContext object for the site holding the page being transformed</param>
        /// <param name="sourceClientContext">The ClientContext for the source </param>
        /// <param name="clientSidePage">Reference to the client side page</param>
        public PublishingBuiltIn(ClientContext sourceClientContext, ClientContext targetClientContext, IList<ILogObserver> logObservers = null) : base(sourceClientContext)
        {
            this.sourceClientContext = sourceClientContext;
            this.targetClientContext = targetClientContext;
            this.parser = new HtmlParser();

            if (logObservers != null)
            {
                foreach (var observer in logObservers)
                {
                    base.RegisterObserver(observer);
                }
            }
        }
        #endregion

        #region Text functions
        /// <summary>
        /// Returns an empty string
        /// </summary>
        /// <returns>Empty string</returns>
        [FunctionDocumentation(Description = "Returns an empty string",
                               Example = "EmptyString()")]
        [OutputDocumentation(Name = "return value", Description = "Empty string")]
        public string EmptyString()
        {
            return "";
        }


        /// <summary>
        /// Returns an the (static) string provided as input
        /// </summary>
        /// <returns>String provided as input</returns>
        [FunctionDocumentation(Description = "Returns an the (static) string provided as input",
                               Example = "EmptyString('static string')")]
        [InputDocumentation(Name = "'static string'", Description = "Static input string")]
        [OutputDocumentation(Name = "return value", Description = "String provided as input")]
        public string StaticString(string staticString)
        {
            return staticString;
        }
        #endregion

        #region Image functions
        /// <summary>
        /// Returns the server relative image url of a Publishing Image field value
        /// </summary>
        /// <param name="htmlImage">Publishing Image field value</param>
        /// <returns>Server relative image url</returns>
        [FunctionDocumentation(Description = "Returns the server relative image url of a Publishing Image field value.",
                       Example = "ToImageUrl({PublishingPageImage})")]
        [InputDocumentation(Name = "{PublishingPageImage}", Description = "Publishing Image field value")]
        [OutputDocumentation(Name = "return value", Description = "Server relative image url")]
        public string ToImageUrl(string htmlImage)
        {
            // Sample input: <img alt="" src="/sites/devportal/PublishingImages/page-travel-instructions.jpg?RenditionID=2" style="BORDER: 0px solid; ">
            var htmlDoc = parser.Parse(htmlImage);
            var imgElement = htmlDoc.QuerySelectorAll("img").FirstOrDefault();

            string imageUrl = "";

            if (imgElement != null && imgElement != default(IElement) && imgElement.HasAttribute("src"))
            {
                imageUrl = imgElement.GetAttribute("src");

                // drop of url params (if any)
                if (imageUrl.Contains("?"))
                {
                    imageUrl = imageUrl.Substring(0, imageUrl.IndexOf("?"));
                }
            }

            return imageUrl;
        }

        /// <summary>
        /// Returns the image alternate text of a Publishing Image field value.
        /// </summary>
        /// <param name="htmlImage">PublishingPageImage</param>
        /// <returns>Image alternate text</returns>
        [FunctionDocumentation(Description = "Returns the image alternate text of a Publishing Image field value.",
                       Example = "ToImageAltText({PublishingPageImage})")]
        [InputDocumentation(Name = "{PublishingPageImage}", Description = "Publishing Image field value")]
        [OutputDocumentation(Name = "return value", Description = "Image alternate text")]
        public string ToImageAltText(string htmlImage)
        {
            // Sample input: <img alt="bla" src="/sites/devportal/PublishingImages/page-travel-instructions.jpg?RenditionID=2" style="BORDER: 0px solid; ">
            var htmlDoc = parser.Parse(htmlImage);
            var imgElement = htmlDoc.QuerySelectorAll("img").FirstOrDefault();

            string imageAltText = "";

            if (imgElement != null && imgElement != default(IElement) && imgElement.HasAttribute("alt"))
            {
                imageAltText = imgElement.GetAttribute("alt");
            }

            return imageAltText;
        }
        #endregion

        #region Person functions
        public string ToAuthors(string userId)
        {
            if (int.TryParse(userId, out int userIdInt))
            {
                var author = Cache.CacheManager.Instance.GetUserFromUserList(this.sourceClientContext, userIdInt);

                if (author != null)
                {
                    // Don't serialize null values
                    var jsonSerializerSettings = new JsonSerializerSettings()
                    {
                        MissingMemberHandling = MissingMemberHandling.Ignore,
                        NullValueHandling = NullValueHandling.Ignore
                    };

                    var json = JsonConvert.SerializeObject(author, jsonSerializerSettings);
                    return json;
                }
            }

            return "";
        }
        #endregion
    }
}
