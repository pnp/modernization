using AngleSharp.Dom;
using AngleSharp.Parser.Html;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using SharePointPnP.Modernization.Framework.Functions;
using SharePointPnP.Modernization.Framework.Telemetry;
using SharePointPnP.Modernization.Framework.Transform;
using System.Collections.Generic;
using System.Linq;

namespace SharePointPnP.Modernization.Framework.Publishing
{
    public class PublishingBuiltIn: FunctionsBase
    {
        private ClientContext sourceClientContext;
        private ClientContext targetClientContext;
        private HtmlParser parser;
        private BuiltIn builtIn;
        private BaseTransformationInformation baseTransformationInformation;


        #region Construction
        /// <summary>
        /// Instantiates the base builtin function library
        /// </summary>
        /// <param name="pageClientContext">ClientContext object for the site holding the page being transformed</param>
        /// <param name="sourceClientContext">The ClientContext for the source </param>
        /// <param name="clientSidePage">Reference to the client side page</param>
        public PublishingBuiltIn(BaseTransformationInformation baseTransformationInformation, ClientContext sourceClientContext, ClientContext targetClientContext, IList<ILogObserver> logObservers = null) : base(sourceClientContext)
        {
            if (logObservers != null)
            {
                foreach (var observer in logObservers)
                {
                    base.RegisterObserver(observer);
                }
            }

            this.sourceClientContext = sourceClientContext;
            this.targetClientContext = targetClientContext;
            this.baseTransformationInformation = baseTransformationInformation;
            this.parser = new HtmlParser();
            this.builtIn = new BuiltIn(this.baseTransformationInformation, targetClientContext, sourceClientContext, logObservers: logObservers);
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
                               Example = "StaticString('static string')")]
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
            // If the image string is not a html image representation then simply return the trimmed value. If an image has a link it's wrapped in an anchor tag
            if (string.IsNullOrEmpty(htmlImage) || !(htmlImage.Trim().StartsWith("<img", System.StringComparison.InvariantCultureIgnoreCase) || htmlImage.Trim().StartsWith("<a", System.StringComparison.InvariantCultureIgnoreCase)))
            {
                return htmlImage;
            }

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
            // If the image string is not a html image representation then simply return the trimmed value. If an image has a link it's wrapped in an anchor tag
            if (string.IsNullOrEmpty(htmlImage) || !(htmlImage.Trim().StartsWith("<img", System.StringComparison.InvariantCultureIgnoreCase) || htmlImage.Trim().StartsWith("<a", System.StringComparison.InvariantCultureIgnoreCase)))
            {
                return htmlImage;
            }

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

        /// <summary>
        /// Returns the image anchor url of a Publishing Image field value
        /// </summary>
        /// <param name="htmlImage">Publishing Image field value</param>
        /// <returns>Image anchor url</returns>
        [FunctionDocumentation(Description = "Returns the image anchor url of a Publishing Image field value.",
                       Example = "ToImageAnchor({PublishingPageImage})")]
        [InputDocumentation(Name = "{PublishingPageImage}", Description = "Publishing Image field value")]
        [OutputDocumentation(Name = "return value", Description = "Image anchor url")]
        public string ToImageAnchor(string htmlImage)
        {
            // If the image string is not a html image representation then simply return the trimmed value. If an image has a link it's wrapped in an anchor tag
            if (string.IsNullOrEmpty(htmlImage) || !(htmlImage.Trim().StartsWith("<img", System.StringComparison.InvariantCultureIgnoreCase) || htmlImage.Trim().StartsWith("<a", System.StringComparison.InvariantCultureIgnoreCase)))
            {
                return htmlImage;
            }

            // Sample input: <img alt="" src="/sites/devportal/PublishingImages/page-travel-instructions.jpg?RenditionID=2" style="BORDER: 0px solid; ">
            var htmlDoc = parser.Parse(htmlImage);
            var anchorElement = htmlDoc.QuerySelectorAll("a").FirstOrDefault();

            string imageAnchor = "";

            if (anchorElement != null && anchorElement != default(IElement) && anchorElement.HasAttribute("href"))
            {
                imageAnchor = anchorElement.GetAttribute("href");

                // drop of url params (if any)
                if (imageAnchor.Contains("?"))
                {
                    imageAnchor = imageAnchor.Substring(0, imageAnchor.IndexOf("?"));
                }
            }

            return imageAnchor;
        }

        /// <summary>
        /// Returns the image caption of a Publishing Html image caption field
        /// </summary>
        /// <param name="htmlField">Publishing Html image caption field value</param>
        /// <returns>Image caption</returns>
        [FunctionDocumentation(Description = "Returns the image caption of a Publishing Html image caption field",
                       Example = "ToImageCaption({PublishingImageCaption})")]
        [InputDocumentation(Name = "{PublishingImageCaption}", Description = "Publishing Html image caption field value")]
        [OutputDocumentation(Name = "return value", Description = "Image caption")]
        public string ToImageCaption(string htmlField)
        {
            // If the image string is not a html image representation then simply return the trimmed value. If an image has a link it's wrapped in an anchor tag
            if (string.IsNullOrEmpty(htmlField))
            {
                return "";
            }

            // Sample input: <p>Some caption<BR></p> 
            try
            {
                var htmlDoc = parser.Parse(htmlField);

                string imageCaption = null;

                if (htmlDoc.FirstElementChild != null)
                {
                    imageCaption = htmlDoc.FirstElementChild.TextContent;
                }

                if (!string.IsNullOrEmpty(imageCaption))
                {
                    return imageCaption;
                }
            }
            catch
            {
                // No need to fail for this reason...
            }

            return "";
        }

        /// <summary>
        /// Returns a page preview image url
        /// </summary>
        /// <param name="image">A publishing image field value or a string containing a server relative image path</param>
        /// <returns>A formatted preview image url</returns>
        [FunctionDocumentation(Description = "Returns a page preview image url.",
                                   Example = "ToPreviewImageUrl({PreviewImage})")]
        [InputDocumentation(Name = "{PreviewImage}", Description = "A publishing image field value or a string containing a server relative image path")]
        [OutputDocumentation(Name = "return value", Description = "A formatted preview image url")]
        public string ToPreviewImageUrl(string image)
        {
            if (string.IsNullOrEmpty(image))
            {
                return "";
            }

            // If the image string is a html image representation
            if (image.Trim().StartsWith("<img", System.StringComparison.InvariantCultureIgnoreCase) || image.Trim().StartsWith("<a", System.StringComparison.InvariantCultureIgnoreCase))
            {
                image = ToImageUrl(image);
            }

            // The image string should now be a server relative path...trigger asset transfer if needed by calling the builtin function ReturnCrossSiteRelativePath
            var previewServerRelativeUrl = this.builtIn.ReturnCrossSiteRelativePath(image);

            // Lookup the image properties by calling the builtin function ImageLookup
            var imageProperties = this.builtIn.ImageLookup(previewServerRelativeUrl);

            // Construct preview image url
            string siteIdString = this.targetClientContext.Site.EnsureProperty(p => p.Id).ToString().Replace("-", "");
            string webIdString = this.targetClientContext.Web.EnsureProperty(p => p.Id).ToString().Replace("-", "");
            if (imageProperties.TryGetValue("ImageUniqueId", out string uniqueIdString))
            {
                uniqueIdString = uniqueIdString.Replace("-", "");
                string extension = System.IO.Path.GetExtension(previewServerRelativeUrl);
                if (!string.IsNullOrEmpty(extension))
                {
                    extension = extension.Replace(".", "");
                }

                if (!string.IsNullOrEmpty(siteIdString) && !string.IsNullOrEmpty(webIdString) && !string.IsNullOrEmpty(uniqueIdString) && !string.IsNullOrEmpty(extension))
                {
                    return $"{this.targetClientContext.Web.EnsureProperty(p => p.Url)}/_layouts/15/getpreview.ashx?guidSite={siteIdString}&guidWeb={webIdString}&guidFile={uniqueIdString}&ext={extension}";
                }
            }

            // Something went wrong...leave preview image url blank so that the default logic during page save can still pick up a nice preview image
            return "";
        }
        #endregion

        #region Person functions
        [FunctionDocumentation(Description = "Looksup user information for passed user id",
                                   Example = "ToAuthors({PublishingContact})")]
        [InputDocumentation(Name = "{userId}", Description = "The id (int) of a user")]
        [OutputDocumentation(Name = "return value", Description = "A formatted json blob describing the user's details")]

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
