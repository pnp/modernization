using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointPnP.Modernization.Framework.Publishing.Layouts
{
    /// <summary>
    /// Contains a central point for defaults for publishing page processing
    /// </summary>
    public static class PublishingDefaults
    {
       /// <summary>
       /// OOB Page Layout defaults
       /// </summary>
        public static List<PageLayoutOOBEntity> OOBPageLayouts = new List<PageLayoutOOBEntity>()
        {
            new PageLayoutOOBEntity(){ Layout = OOBLayout.ArticleLeft, Name = "ArticleLeft", PageLayoutTemplate = "TwoColumnsWithHeader", PageHeader = "Custom" },
            new PageLayoutOOBEntity(){ Layout = OOBLayout.WelcomeLinks, Name = "WelcomeLinks", PageLayoutTemplate = "", PageHeader = "" },
            new PageLayoutOOBEntity(){ Layout = OOBLayout.ArticleLinks, Name = "ArticleLinks", PageLayoutTemplate = "", PageHeader = "" },
            new PageLayoutOOBEntity(){ Layout = OOBLayout.ArticleRight, Name = "ArticleRight", PageLayoutTemplate = "", PageHeader = "" },
            new PageLayoutOOBEntity(){ Layout = OOBLayout.PageFromDocLayout, Name = "PageFromDocLayout", PageLayoutTemplate = "", PageHeader = "" },
            new PageLayoutOOBEntity(){ Layout = OOBLayout.WelcomeSplash, Name = "WelcomeSplash", PageLayoutTemplate = "", PageHeader = "" },
            new PageLayoutOOBEntity(){ Layout = OOBLayout.RedirectPageLayout, Name = "RedirectPageLayout", PageLayoutTemplate = "", PageHeader = "" },
            new PageLayoutOOBEntity(){ Layout = OOBLayout.BlankWebPartPage, Name = "BlankWebPartPage", PageLayoutTemplate = "", PageHeader = "" },
            new PageLayoutOOBEntity(){ Layout = OOBLayout.ErrorLayout, Name = "ErrorLayout", PageLayoutTemplate = "", PageHeader = "" },
            new PageLayoutOOBEntity(){ Layout = OOBLayout.CatalogArticle, Name = "CatalogArticle", PageLayoutTemplate = "", PageHeader = "" },
            new PageLayoutOOBEntity(){ Layout = OOBLayout.CatalogWelcome, Name = "CatalogWelcome", PageLayoutTemplate = "", PageHeader = "" },
            new PageLayoutOOBEntity(){ Layout = OOBLayout.EnterpriseWiki, Name = "EnterpriseWiki", PageLayoutTemplate = "", PageHeader = "" },
            new PageLayoutOOBEntity(){ Layout = OOBLayout.ProjectPage, Name = "ProjectPage", PageLayoutTemplate = "", PageHeader = "" }
        };


        /// <summary>
        /// Web Part Zone Row/Columns for mappings
        /// </summary>
        public static Dictionary<string, string> FieldToTargetWebParts = new Dictionary<string, string>()
        {
            { "PublishingPageImage", "SharePointPnP.Modernization.WikiImagePart"},
            { "PublishingPageContent", "SharePointPnP.Modernization.WikiTextPart"},
            { "SummaryLinks", "Microsoft.SharePoint.Publishing.WebControls.SummaryLinkWebPart"}
        };

        
        /// <summary>
        /// Field Control Defaults for mappings
        /// </summary>
        public static List<PageLayoutFieldControlEntity> FieldControlProperties = new List<PageLayoutFieldControlEntity>()
        {
            new PageLayoutFieldControlEntity() { TargetWebPart = "SharePointPnP.Modernization.WikiImagePart", FieldName = "PublishingPageImage",  Name="ImageUrl", FieldType="String", ProcessFunction = "ToImageUrl({PublishingPageImage})" },
            new PageLayoutFieldControlEntity() { TargetWebPart = "SharePointPnP.Modernization.WikiImagePart", FieldName = "PublishingPageImage",  Name="AlternativeText", FieldType="String", ProcessFunction = "ToImageAltText({PublishingPageImage})" },

            new PageLayoutFieldControlEntity() { TargetWebPart = "Microsoft.SharePoint.Publishing.WebControls.SummaryLinkWebPart", FieldName = "SummaryLinks", Name = "SummaryLinkStore", FieldType="string"},

            new PageLayoutFieldControlEntity() { TargetWebPart = "SharePointPnP.Modernization.WikiTextPart", FieldName = "PublishingPageContent", Name="Text", FieldType="string" },
            
        };

        /// <summary>
        /// Metadata field default mappings
        /// </summary>
        public static List<PageLayoutMetadataEntity> MetaDataFieldToTargetMappings = new List<PageLayoutMetadataEntity>()
        {
            new PageLayoutMetadataEntity(){ FieldName = "Title", TargetFieldName="Title", Functions = "" },
            
        };

        /// <summary>
        /// Field to header mappings
        /// </summary>
        public static List<PageLayoutHeaderFieldEntity> PageLayoutHeaderMetadata = new List<PageLayoutHeaderFieldEntity>()
        {
            new PageLayoutHeaderFieldEntity() { HeaderType = "FullWidthImage", FieldName = "PublishingRollupImage", FieldHeaderProperty = "ImageServerRelativeUrl", FieldFunctions = "ToImageUrl({PublishingRollupImage})" },
            new PageLayoutHeaderFieldEntity() { HeaderType = "FullWidthImage", FieldName="ArticleByLine", FieldHeaderProperty = "TopicHeader", FieldFunctions = "" },
        };

        /// <summary>
        /// List of metadata fields in content types to ignore in mappings
        /// </summary>
        public static List<string> IgnoreMetadataFields = new List<string>()
        {
            "ContentType",
            "FileLeafRef",
            "RobotsNoIndex",
            "SeoBrowserTitle",
            "SeoMetaDescription",
            "SeoKeywords",
            "PublishingPageLayout"
        };

        
    }
}
