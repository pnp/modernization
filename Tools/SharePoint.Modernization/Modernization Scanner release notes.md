# Modernization Scanner release notes

## How to get and use

See https://aka.ms/sppnp-modernizationscanner

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