using System;

namespace SharePointPnP.Modernization.Framework
{
    /// <summary>
    /// Constants used 
    /// </summary>
    public static class Constants
    {
        // Fields
        public const string FileRefField = "FileRef";
        public const string FileDirRefField = "FileDirRef";
        public const string FileLeafRefField = "FileLeafRef";
        public const string FileTitleField = "Title";
        public const string ClientSideApplicationIdField = "ClientSideApplicationId";
        public const string HtmlFileTypeField = "HTML_x0020_File_x0020_Type";
        public const string WikiField = "WikiField";
        public const string ModifiedField = "Modified";
        public const string ModifiedByField = "Editor";
        public const string PublishingPageLayoutField = "PublishingPageLayout";
        public const string AudienceField = "Audience";
        public const string PublishingRollupImageField = "PublishingRollupImage";
        public const string TitleField = "Title";

        // Features
        public static readonly Guid FeatureId_Web_ModernPage = new Guid("B6917CB1-93A0-4B97-A84D-7CF49975D4EC");

        // Queries
        public const string CAMLQueryByExtension = @"
                <View Scope='Recursive'>
                  <Query>
                    <Where>
                      <Contains>
                        <FieldRef Name='File_x0020_Type'/>
                        <Value Type='text'>aspx</Value>
                      </Contains>
                    </Where>
                  </Query>
                </View>";
        public const string CAMLQueryByExtensionAndName = @"
                <View Scope='Recursive'>
                  <Query>
                    <Where>
                      <And>
                        <Contains>
                          <FieldRef Name='File_x0020_Type'/>
                          <Value Type='text'>aspx</Value>
                        </Contains>
                        <BeginsWith>
                          <FieldRef Name='FileLeafRef'/>
                          <Value Type='text'>{0}</Value>
                        </BeginsWith>
                      </And>
                    </Where>
                  </Query>
                </View>";


        // Cross site assets transfers that are currently allowed
        public static string[] AllowedAssetFileExtensions = new string[] { "png", "jpg", "gif", "mp4", "mpeg" };
        // Cross site assets transfer blocked files that are referred to
        public static string[] BlockedAssetFileExtensions = new string[] { "aspx" };
    }
}
