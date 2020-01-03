# Modernization Framework release notes

## How to get and use

### How to get

- Get binaries via nuget: https://www.nuget.org/packages/SharePointPnPModernizationOnline. You can rename the .nupkg package to .zip and extract the binaries from it
- Optionally self-service compile the binaries by pulling down the https://github.com/SharePoint/sp-dev-modernization/tree/master/Tools/SharePoint.Modernization/SharePointPnP.Modernization.Framework project from GitHub

### How to use

- From .Net: see https://docs.microsoft.com/en-us/sharepoint/dev/transform/modernize-userinterface-site-pages-dotnet as nice sample to start with
- From PnP PowerShell: see https://docs.microsoft.com/en-us/sharepoint/dev/transform/modernize-userinterface-site-pages-powershell for a sample

## [Unreleased]

### Added

### Changed

- Fix: Improved v3 (e.g. XSLTListView) web part handling for SP2010 - now webpart properties, zoneId and controlId are correct loaded #384
- Fix: Correctly handle images hosted in the server side _layouts folder when doing cross site transformations #390
- Fix: Allow to force a page to be created in the root of the SitePages library by setting `TargetPageFolder = "<root>"` and `TargetPageFolderOverridesDefaultFolder = true`
- Bumped to .Net 4.6.1 as minimal .Net runtime version
- Caching now uses IDistributedCache as cache store model, allowing efficient caching in distributed systems
- The default webpartmapping.xml file is available as embedded resource
- Target site can be a developer site (DEV web template) or document center (BDR web template) #400

## [December release - version 1.0.1912.0]

### Added

- Added support for transforming Delve blog pages
- Added TargetPageFolderOverridesDefaultFolder option to have the provided target folder override the default generated folder #366
- SharePoint Add-In part properties can now be used in web part mappings #374

### Changed

- Fix: Double quotes used in the "alt" property of an img do break the web part properties later on #365
- Fix: Include page transformation version number in the log #362
- Fix: Update logging to show correct user names when performing user mapping #357
- Fix: Find username also works for SPO user names when doing SPO to SPO user mapping
- Fix: Sections with a vertical column are now also "cleaned" when they do contain empty columns #359
- Fix: Empty table rows are now correctly processed as empty rows #358
- Fix: Webpart page layouts HeaderFooter2Columns4Rows and HeaderFooter4ColumnsTopRow threw "out of range" error #371
- Fix: Return result of page transformation an an URL encoded url #370
- Fix: ClearCache doesnt clear page layout mappings #364
- Fix: Minor logging message misleading during user transform #363 
- Fix: Doing http requests (e.g. ASMX service calls) fail if cookie based authentication (e.g. when using ADFS) information was not attached to those requests #372
- Fix: TargetUrl is not populated in log (Nov release regression) #375

## [November release - version 1.0.1911.0]

### Added

- Mapping from on-premises users/groups to SPO users/groups #161 [pkbullock]
- Section emphasis (=section background) can be defined in your page layout mapping files via the SectionEmphasis element in the PageLayout element
- Publishing pages can be transformed to a target page that uses a vertical column. Add IncludeVerticalColumn="true" to the PageLayout element in your page layout mapping file
- Blog page transformation now also works for SharePoint 2010 blog pages
- Use the TargetPageFolder setting to specify the target folder in which the modern page will be created #260
- Url rewriting now is applied to anchor tags of transformed images. If the original anchor and image are the same then they're kept the same after transformation

### Changed

- Fix: Cannot assign limited access permission #338 [pkbullock]
- Fix: Web Part page layout is now correctly detected for non English sites #342
- Fix: Keep casing of original folder and file names during transformation #343
- Additional logging to show asset transfer progress #346 [pkbullock]

## [October intermediate release - version 1.0.1910.1]

### Changed

- Fix: On-premises web part handling refactoring, SP2010 now uses different flow so that v3 web parts (sp2013+) are also loaded correctly
- Fix: Handle web part loading of web parts in wiki/blog pages when transforming from SP2010/2013/2016 to SPO #320
- Fix: Additional check to determine whether source file exists or not #318 [gautamdsheth]
- Allow overriding the target page name for web part pages living outside of a library
- Fix: Improved wiki page layout detection, ensure we always return a layout type with the correct amount of column to avoid index errors #304
- Fix: Only populate page header author information when transforming from on-premises. Pages with bad author information cannot be edited 
- Fix: The swap pages (TargetPageTakesSourcePageName option) now uses File.Move instead of File.Copy. This fixes the issue described in #275

## [October release - version 1.0.1910.0]

### Added

- Support for doing cross site classic blog to modern page transformation
- KeepPageCreationModificationInformation option to tell page transformation to keep the source page's author, editor, creation and modification dates
- PostAsNews option to post the created page as news. Setting this option will also set PublishCreatedPage to true
- Added error handling and logging points to the publishing page analyser #287 [pkbullock]
- Reporting Enhancements #293 [pkbullock]
- Added support for wiki page metadata copy across site collections #281 [gautamdsheth]
- Support to populate the author shown in the modern page header with author information of the original page

### Changed

- Fix: Handle web part loading of web parts embedded in publishing content fields when transforming from SP2010/2013/2016 #265
- Fix: Returned modern page url is containing /pages/ when transforming publishing pages the from root site collection #262
- Prevent unneeded intermediate versions of the created page
- Support publishing page layout mapping files without a MetaData section (in case no metadata needs to be taken over)
- Fix: Logging shows class name instead of data #273
- Fix: Handle case where the publishing page header fields are not populated due to missing data in the source page
- Fix: Option to insert 'hard coded' html content on the created target page now also works when transforming from SP2010/2013/2016
- Log a warning when a metadata field defined in the used page layout mapping file does not exist in the source page
- Transform to the newer, more versatile, News web part versus the NewsReel #263
- Fix: Don't try to transfer CDN images from OnPrem root site collections #278 [gautamdsheth]
- Fix: Remove asp prefixes from XML block when processing on-premises pages #280 [pkbullock]
- Fix: Fetching the fields from the list instead of the content types for more robustness in metadata copy of publishing pages #279 [gautamdsheth]
- Fix: Changed version request credentials to use clientContext credentials #289 [thechriskent]
- Fix: If we've set a custom thumbnail value then we need to update the page html to mark the isDefaultThumbnail pageslicer property to false #290
- Fix: Transfer of assets on root site collections now also works for sub sites of that root site collection #301 [pkbullock]
- Changed default value from AddTableListImageAsImageWebPart to true 
- Fix: Sanitized blog page names to remove special characters #309 [gautamdsheth]
- When a user is not resolving via EnsureUser (e.g. because the user account was deleted) then a warning is logged during page metadata copying
- Fix: Calendar lists shown with the "All Events" and "Current Events" views are now correctly recognized as calendar #312

## [September release - version 1.0.1909.0]

### Added

- A page layout mapping can be reused for multiple, similar, page layouts by specifying the additional page layouts as a semi colon separated list in the AlsoAppliesTo attribute #217
- SkipDefaultUrlRewrite pageTransformationInformation property that allows one to skip the default URL rewriting logic while still applying a possible provided custom URL mapping #219
- Option to insert 'hard coded' html content on the created target page (e.g. hard coded text in your page layout) by using below construct in your page layout mapping file. Running functions on fields targetting the "SharePointPnP.Modernization.WikiTextPart" web part is not also possible.

```XML
<WebParts>
  ...
  <Field Name="ID" TargetWebPart="SharePointPnP.Modernization.WikiTextPart" Row="1" Column="3">
    <Property Name="Text" Type="string" Functions="StaticString('&lt;H1&gt;This is some extra text&lt;/H1&gt;')" />
  </Field>
  ...
</WebParts>
```

### Changed

- Changed default to not insert a placeholder message anymore above an image inside a table/list as nowadays images are not dropped from the editor anymore ==> default web part mapping file bumped to version 1.0.1909.0
- Images embedded in a table/list are not added as separate image web parts anymore, you can use the AddTableListImageAsImageWebPart PageTransformationInformation property if you still require want this to happen
- Fixed a bug where web service not handling empty web part returns related to #232 [pkbullock]
- Fix: Don't transform closed web parts #236 [gautamdsheth]
- Fix: Handle the case where a calendar web part is a source page and the calendar list is not available in the target site collection #239
- Fix: Publishing of a page was "undone" by possible metadata copies or item level security handling. Publish part now is now moved to the very end of the transformation flow #242
- Fix: Log which page layout mapping is used for the given publishing page #241
- CBS/CBQ transforms now uses the custom KQL/CAML query option of the highlighted content web part whenever a KQL/CAML query is available #238
- Fix: In the publishing page flow the created modern page can have the original page author/editor and page creation/edit date #246
- Fix: Ensure correct usage of EnsureProperty method #248 [gautamdsheth]
- Fix: web part handling for SP2013/2016, now uses the 2010 flow #234 [pkbullock]
- Fix: added trailing slash to support transforming pages in the root site collection #252 [gautamdsheth]
- Fix: when multiple add in web parts (provider hosted, SharePoint hosted) are available on a site then the web part match up could fail
- Mark created pages as "MigratedFromServerRendered" via the _SPSitePageFlags field
- Add tenantid in telemetry data
- Configured the ClientSideWebPart and ClientWebPart as cross site supported ==> if the SPFX app or add in delivered web part is available on the target it will be put on the page #245
- Fix: handle web part page layouts that originated from SP2010 on-premises #261

## [August release - version 1.0.1908.0]

### Added

- SharePoint 2010 Preview Support #204, #209 [pkbullock]

### Changed

- Url Rewrite - Issue with root addresses #205 [pkbullock]
- Fix item level permission copy for cross site collection scenarios (groups are only assigned when the group already exists in the target site) #212

## [July release - version 1.0.1907.0]

### Added

- Custom URL mapping logic: provide a csv file with source and target values and these will be used by the url mapper #135
- Option to override default QuickLinks configuration in publishing page transformation scenarios #191
- Option to "map" web parts inside a web part zone #167
- Support for multiple source field "name" values in page layout mapping files, allows to define "overrides" if a given field is not populated in the source page #201 [MartinHatch]

### Changed

- User fields are now correctly copied over in cross site publishing transformation scenarios #184
- Mapping Files Version Change Notice #188
- Set TargetPageName is now used to construct the return URL value #194
- Correctly detect 'empty' text parts #192

## [June release - version 1.0.1906.0]

### Added

- Preview On-Premises publishing page to SharePoint Online modern page support #165 [pkbullock]
- Support for transforming web part pages living outside of a library (so in the root folder of the site)
- Support for provisioning the Page Properties web part on a page (only for publishing page transformation). #171

### Changed

- Reporting improvements for on-premises as source + correct log level for some log entries #169 [pkbullock]
- Logic added that disables item level permissions copy in cross-farm scenarios #178 [pkbullock]
- In publishing scenarios it's common to not have all fields defined in the page layout mapping filled. By default we'll not map empty fields as that will result in empty web parts which impact the page look and feel. Using the RemoveEmptySectionsAndColumns flag this behaviour can be turned off. #156

## [May release (prod) - version 1.0.1905.3]

- Intermediate release due to needed intermediate release of the used PnP Sites Core library

## [May release (prod) - version 1.0.1905.2]

- Intermediate release due to needed intermediate release of the used PnP Sites Core library

## [May release (prod) - version 1.0.1905.1]

- Intermediate release due to needed intermediate release of the used PnP Sites Core library

## [May release (prod) - version 1.0.1905.0]

### Added

- Support for static parameter values in function definitions. Use the new StaticString function, e.g. StaticString('your static string') to define a static value. Fixes #119
- Support for running a function on MetaData field mappings (single function can be added per field, not supported for taxonomy fields)
- Added ToPreviewImageUrl built in function which allows to control the page preview image via either a dynamic value (field of the source page list item) or a static string
- Added mapping for image anchor (wiki and publishing) and image caption (publishing) to modern image web part
- Simple URL rewrite engine for publishing page transformation
- Summary report generation #141 [pkbullock]

### Changed

- Added filter for ASPX files and additional error handling #102 [pkbullock]
- Duplicate analyser mappings produced #100 [pkbullock]
- Amended Page Transformator to check target site for existing file #117 [pkbullock]
- Fixed CDATA handling in page layout analyzer
- Support table structures with mixed TD and TH cells in the same TR
- Content Editor web part title is taking over when the web part's "ChromeType" differs from "None" or "Border-only" #120
- Fix: Table size detection issues #123 and #124
- Fix: Performance improvements around asset transfer #125 (fixes #111) [pkbullock]
- Fix: Take over text alignment in table cells #104
- Fix: Integrate "enhanced" processing of content editor text content also in web part page analyzer flow #106
- Fix: capacity was less than the current size. #130
- PageLayoutAnalyser.AnalyseAll now can skip OOB page layouts
- Fix: Page layout analyzer can handle fields specified by id instead of name #131 and #133 [pkbullock]
- Fix: Web part title (when the web part's "ChromeType" differs from "None" or "Border-only") is retained when summarylinks are transformed to html #137
- Fix: switch to FIPS compliant hash method #146
- Fix: order can be set also for web part zones and fields that transform into web parts allowing to now set order for all visible components on the page #148

## [April release (prod) - version 1.0.1904.0]

### Added

- Publishing page support!
- Transformation Reporting. Get (verbose) logging as md file, page in SharePoint, console or a combination of these #82 [pkbullock]
- Page Layout Mapping  #86 [pkbullock]
- UserDocsWebPart transformation support: this web part is transformed to the highlighted content web part showing the current user's active pages
- PublishCreatedPage configuration option: allows to define if a page needs to be published or not
- DisablePageComments configuration option: allows to define if page comments needs to be disabled or not

### Changed

- Transforming Summary Links to Quick Links json encoding Bug #74
- Tables with col/row spans: split cells and put the content in the first cell of the split #77
- Transform nested tables as individual tables #75
- Support transformation from pages living outside of the sitepages library #80
- Content editor: if content is recognized as transformable html (so no script) then it will be treated as wiki content, hence embedded images and videos will be created as separate image and video web part + in cross site scenarios the images are copied over to the target site

## [March release (prod) - version 1.0.1903.0]

### Added

- Support for creating the modern site pages in another site collection. Does support asset transfer to the target site collection for a limited set of web parts decorated with the CrossSiteTransformationSupported="true" attribute. #59, #65, #66 and #71 [pkbullock]
- Support for using Boolean as return type of functions used in the web part transformation model
- XSLTListView transformation: map the web part toolbar configuration to the hideCommandBar property
- Transformation support for ContentBySearchWebPart and ResultScriptWebPart
- Drop "empty" text parts...text parts with html tags without visual presentation are useless. Wiki pages, especially with multi section/column layouts, tend to have these
- Drop empty sections and columns to optimize the screen real estate - also better aligns with how web part pages and wiki pages behave in classic. This behavior is on by default, but can be turned off via the RemoveEmptySectionsAndColumns flag in the PageTransformationInformation class
- ExcelWebRenderer transformation: take over the configured named item (table, chart, range)
- SummaryLinks transformation: new default is transform to QuickLinks, optionally you still transform to text by setting the SummaryLinksToQuickLinks mapping property to false
- ContactFieldControl transformation support: this web part transforms to the People web part
- Support for defining functions on a mapping: this allows to execute code only when a specific mapping was chosen

### Changed

- ContentEditor transformation: when not using 3rd party script editor embedded and file contents without script references is not treated as text
- Content by query transformation:
  - Support for site collection and sub site scoped queries, including filters and sorting for those type of queries
  - Specific support for SitePages library queries in the list scoped query handling
  - More detailed content type filter handling
  - Switched to version 2.2 of data model
- SummaryLinks transformation: links without heading are now correctly transformed to html
- Mapping properties allow for mapping based up on configuration: the UseCommunityScriptEditor property can be set to use the community script editor, no need for changing mapping files to support this scenario
- MembersWebPart transformation: now shows a text making users aware of the OOB Site Permissions feature that replaces this web part's functionality

## [February release (prod) - version 1.0.1902.0]

### Added

- Support for pages living inside a folder (issue #34)
- Support for copying of metadata (issue #35)

### Changed

- Wiki page parser: check if this element is nested in another already processed element...this needs to be skipped to avoid content duplication and possible processing errors (issue #37)
- Improved item level permission copy logic
- Check for proper permissions before attempting item level copy, if insufficient permissions the item level permissions are not copied but the transformation will still succeed
- Only transform content editor web part pointing to .aspx file to contentembed, .html files result in a file download instead of a file load

## [January release (prod) - version 1.0.1901.1]

### Changed

- Fix for issue #30 to enable page transformation for pages in tenant root web in combination with XSLTListView web part transformation

## [January release (prod) - version 1.0.1901.0]

### Added

- Support for new 1st party web parts: these can now be included in your webpartmapping.xml files

### Changed

- Massive performance improvements (double as fast) for page transformation. Also improves performance of the (publishing) page scanner components

## [December release (prod) - version 1.0.1812.1]

### Changed

- Using December 2018 PnP Sites core package

## [December release (prod) - version 1.0.1812.0]

### Added

- Return site relative URL of the created modern page

### Changed

- Compiled using .Net Framework version 4.5 instead of 4.5.1 to allow inclusion in 4.5 projects (like PnP PowerShell)

## [November release (prod) - version 1.0.1811.2]

### Changed

- Support for putting a banner web part on the created pages. See https://github.com/SharePoint/sp-dev-modernization/tree/dev/Solutions/PageTransformationUI for more details

## [November release (prod) - version 1.0.1811.1]

### Changed

- Added Page Propertybag entry with version stamp of modernization framework used to generate the page
- Added Azure AppInsights based telemetry, only anonymous data is sent

## [November release (prod) - version 1.0.1811.0]

### Changed

- Updates when transforming wiki html:
  - H4 to H6 elements now retain their formatting when converted to text
  - Combining italic/underline/bold in combination with other type of formatting now works stable
  - Strip out the "zero width space characters"
  - Drop wiki font information
  - Handle additional styles (ms-rteStyle-Quote,ms-rteStyle-IntenseQuote,ms-rteStyle-Emphasis,ms-rteStyle-IntenseEmphasis,ms-rteStyle-References,ms-rteStyle-IntenseReference,ms-rteStyle-Accent1,ms-rteStyle-Accent2)
  - Better handling complex nested styles
  - Full rewrite of indent handling: now supports complex formatting inside indents, indenting of blocks and unlimited indent depth
  - Switch default table style to borderHeaderTableStyleNeutral - this allows highlighted text to show as highlighted, plain table style suppresses this
  - Assume a table width of 800px and spread evenly across available columns
  - Improved reliability in detecting images/videos inside wiki text fragments
  - Clean wiki html before/after processing to drop nodes which are not support in RTE
  - Full rewrite of wiki splitting...better reliability, better results and better performance

## [October release (prod) - version 1.0.1810.2]

### Changed

- Fixed issue with default page layout transformation for "One Column with Sidebar" wiki pages

## [October release (prod) - version 1.0.1810.1]

### Added

- Support for adding a web part of choice as banner on all generated pages. Used to give end users an option to accept/decline the generated page

### Changed

- Lowered minimal .Net framework version from 4.7 to 4.5.1
- Expose the swap pages logic so that it can be used by folks using the page transformation engine

## [October release (prod) - version 1.0.1810.0]

### Added

- Wiki text handling: Headers (H1 to H3), STRONG and EM tags with custom formatting do retain their formatting
- Supported formatting in table cells is retained when the table html is rewritten

### Changed

- Approach to give newly created modern page the same name as the source page has been fixed: now url's to these pages in navigation or other pages are not rewritten

## [Beta release - version 0.1.1808.0]

### Added

- Header (H1 to H4) alignment is retained when transforming wiki text
- Combined styles (e.g. forecolor with strike-through and font size) are now correctly handled when transforming wiki text
- Documentation for functions and selectors is now autogenerated (https://docs.microsoft.com/en-us/sharepoint/dev/transform/modernize-userinterface-site-pages-api)
- Support added for having text before and after the web part but inside the div surrounding the web part
- Theme colors are transformed now
- Source page item level permissions are copied to the target page (can be optionally turned off)

### Changed

- Page title handling got improved
- Improved handing of BR tags
- Improved reliability in handling image URL's outside of the current web
- Fixed layout transformation for HeaderRightColumnBody and HeaderLeftColumnBody web part page layouts
- Fixed "duplicate key" issue when transforming multiple pages in sequence
- Fixed ListId datatype in model
- Calendar is now transformed to the Events web part
- Tasks web part is not transformed anymore
- Correctly identify a discussion board
