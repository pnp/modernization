# Modernization Scanner release notes

## How to get and use

See https://aka.ms/sppnp-modernizationscanner

## [Version 2.5 - Unreleased]

### Added

### Changed

- Improved report generation from multiple individual scans (-g parameter:
  - Publishing portal now is correctly aggregated
  - Handle the scenario where there's certain scan files missing because that scan component was not selected
  - Handle the use of relative paths
- Sites based upon CMSPUBLISHING#0 are counted as publishing sites

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