# Modernization Scanner release notes

## How to get and use

See https://aka.ms/sppnp-modernizationscanner

## [Version 2.18]

## Added

## Changed

- Workflows on list content types are now included

## [Version 2.17]

## Added

## Changed

- Prevent scanning the root site from issuing a search query that would return all sites
- Removed the Delve blog scanning feature as Delve has been removed from SPO

## [Version 2.16]

## Added

## Changed

- Previous workflow versions are now correclty excluded, also when the site was created using a non-Enlish language #518
- Prevent rapid exit without waiting for the user to read error (if anything throws) #511 [victorbutuza]
- Also process 2010 WF activities from the Microsoft.SharePoint.WorkflowActions.WithKey namespace

## [Version 2.15]

## Added

## Changed

- Fixed bug in search code that resulted in an infinite search loop [Chipzter]

## [Version 2.14]

## Added

- Option to scan for workflows without analyzing the workflows (increases performance)

## Changed

- Scanner now outputs CSV results each minute

## [Version 2.13]

### Added

- Support to use Azure AD based authentication in US Government, Germany and China clouds

### Changed

- Change urls parameter help #500 [KoenZomers]

## [Version 2.12]

### Added

- Workflow and InfoPath reports now also contain the Admins and Owners of the site collection (copy from the information of the sitescan results) #483
- Add option to wizard to specify tenant admin center url when using a CSV file input

### Changed

- Skip site collections created for Team private channels, no point in scanning those #487

## [Version 2.11]

### Added

### Changed

- Fix: Fixed list view threshold error when using Sites.Read.All permissions in combination with a tenant having a lot of sites
- Fix: Use the CanModernizeHomepage API call to understand whether a home page will be modernized automatically by the planned home page modernization effort

## [Version 2.10]

### Added

### Changed

- Fix: Checking for home.aspx now takes in account the locale of the site
- The definition of an uncustomized home page changed, added the check on publishing features and master page
- Fix crash in workflow scanning component #432

## [Version 2.9]

### Added

- Site and web search center url setting is included in the ModernizationSiteScanResults.csv and ModernizationWebScanResults.csv files
- The scanner uses modern username/password auth, there's no dependency anymore on legacy auth being enabled on the tenant
- The scanner supports multi-factor authentication via an interactive login prompt

## [Version 2.8]

### Added

- Home page only page scan mode
- Creation of a SitesWithUncustomizedHomePages.csv file listing all the home pages which are uncustomized

### Changed

- Bumped to .Net 4.6.1 as minimal .Net runtime version

## [Version 2.7]

### Added

- Delve blogs are scanned as part of the blog scan component
- Added option to use a certificate stored in certificate store (next to the already existing option of providing via pfx)
- Office 365 Group connection report and csv's will now also list if a site has a Teams team (only when using Azure AD auth and when the Groups.Read.All permission was granted)

### Changed

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