# Modernization Scanner release notes

## How to get and use

See https://aka.ms/sppnp-modernizationscanner

## [Unreleased]

## Added

- Delve blogs are scanned as part of the blog scan component
- Added option to use a certificate stored in certificate store (next to the already existing option of providing via pfx)
- Office 365 Group connection report and csv's will now also list if a site has a Teams team (only when using Azure AD auth and when the Groups.Read.All permission was granted)

## Changed

- Refactored the scanner into a core scanner library and a consumer (.exe). Core scanner library uses streams for all the file manipulation, the consumer is responsible for providing/persisting files. This will make the core scan component easier to re-use

## [Version 2.6]

### Added

- Blog site/page scanning: provides you the needed information on blog usage in your environment
- Option to run scanner with Sites.Read.All permission when using Azure AD App-Only. Note that this implies that the SkipUserInformation will be automatically turned on and that workflow scanning is skipped.
- Workflow report now contains an "upgradability" score based upon the mapping of workflow to Flow actions
- Workflow report now also allows to filter on last change date of a workflow definition, this can be used to identify the recently changed workflows 
- Page report contains a column to identify "uncustomized" STS#0 home pages
- New parameter (-q) and UI option to configure the date format to be used in the exported CSV files

### Changed

- Tenant root site collection is now included if you scan for a complete tenant
- Add tenantid in telemetry data
- InfoPath scanner detects form libraries which are created by adding the Form content type as default
- InfoPath scanner detects libraries which have a Form content type attached
- DateTime values are outputted as Date strings based upon the chosen date format

## [Version 2.5]

### Added

- Publishing page report now shows "web part transformation compatibility" graphs
- New scan component: classic workflow inventory as preparation for Microsoft Flow migrations
- New scan component: InfoPath usage inventory as preparation for Microsoft PowerApps migrations

### Changed

- Improved report generation from multiple individual scans (-g parameter):
  - Publishing portal now is correctly aggregated
  - Handle the scenario where there's certain scan files missing because that scan component was not selected
  - Handle the use of relative paths
- Sites based upon CMSPUBLISHING#0 are counted as publishing sites
- Drop mappings for scripteditor and htmlform web parts as these depend on the community script editor. This change is required to align with the new default webpartmapping.xml file
- Add "Unmapped web parts" column to the publishing page scan CSV file (ModernizationPublishingPageScanResults.csv)

## [Version 2.4]

### Added

- Built-in webpartmapping.xml file, no need to deploy scanner.exe + webpartmapping.xml file together

### Changed

- Performance tuning: version 2.4 is 40% faster for a "full" run than 2.3

## [Version 2.3]

### Added

- Built-in check to see if there's a newer version available
- Generation of SitesWithCustomizations.csv

### Changed

- Using latest PnP Sites Core library with updated throttling implementation
- Several small reliability improvements
- Export SiteId in ModernizationSiteScanResults.csv