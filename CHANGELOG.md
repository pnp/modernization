# PnP Changelog
*Please do not commit changes to this file, it is maintained by the repo owner.*

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](http://keepachangelog.com/en/1.0.0/).

## [December 2018]

### Added

- SharePoint Modernization Framework production release (1.0.1812.1):
	- Return site relative URL of the created modern page

- SharePoint Page Transformation UI (preview release):
	- First public release

- SharePoint.Modernization.Scanner v2.3:
	- Built-in check to see if there's a newer version available
	- Generation of SitesWithCustomizations.csv

### Changed

- SharePoint Modernization Framework production release (1.0.1812.1):
	- Using December 2018 PnP Sites core package
	- Compiled using .Net Framework version 4.5 instead of 4.5.1 to allow inclusion in 4.5 projects (like PnP PowerShell)
	- Support for putting a banner web part on the created pages. See https://github.com/SharePoint/sp-dev-modernization/tree/dev/Solutions/PageTransformationUI for more details
	- Added Page Propertybag entry with version stamp of modernization framework used to generate the page
	- Added Azure AppInsights based telemetry, only anonymous data is sent

- SharePoint.Modernization.Scanner v2.3:
	- Using latest PnP Sites Core library with updated throttling implementation
	- Several small reliability improvements
	- Export SiteId in ModernizationSiteScanResults.csv
