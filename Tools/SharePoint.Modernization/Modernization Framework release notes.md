# Modernization Framework release notes

## How to get and use

### How to get

- Get binaries via nuget: https://www.nuget.org/packages/SharePointPnPModernizationOnline. You can rename the .nupkg package to .zip and extract the binaries from it
- Optionally self-service compile the binaries by pulling down the https://github.com/SharePoint/sp-dev-modernization/tree/master/Tools/SharePoint.Modernization/SharePointPnP.Modernization.Framework project from GitHub

### How to use

- From .Net: see https://docs.microsoft.com/en-us/sharepoint/dev/transform/modernize-userinterface-site-pages-dotnet as nice sample to start with
- From PnP PowerShell: see https://docs.microsoft.com/en-us/sharepoint/dev/transform/modernize-userinterface-site-pages-powershell for a sample

## [March release (prod) - unreleased]

### Added

- Experimental support for creating the modern site collection in another site collection. Currently does not yet support web parts with references to the source site + copy of page metadata #59 [pkbullock]
- Support for using Boolean as return type of functions used in the web part transformation model
- XSLTListView transformation: map the web part toolbar configuration to the hideCommandBar property

### Changed

- ContentEditor transformation: when not using 3rd party script editor embedded and file contents without script references is not treated as text

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
