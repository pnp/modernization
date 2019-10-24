# SharePoint Modernization scanner #

### Summary ###

Using this scanner you can prepare your classic sites for modernization via connecting these sites to an Office 365 group (the "groupify" process), modernizing the pages and in case of a publishing portal designing a modern publishing portal. This scanner is a key tool to use if you want to prepare for modernizing your classic sites.

> **Important**
> Checkout the [Modernize your classic sites](https://aka.ms/sppnp-modernize) article series on docs.microsoft.com to learn more about modernization. [Connect to an Office 365 group](https://docs.microsoft.com/en-us/sharepoint/dev/transform/modernize-connect-to-office365-group) and [Transform classic pages to modern client-side pages](https://docs.microsoft.com/en-us/sharepoint/dev/transform/modernize-userinterface-site-pages) are articles that refer to this scanner.

### Applies to ###

- SharePoint Online

### Solution ###

Solution | Author(s)
---------|----------
SharePoint.Modernization.Scanner | Bert Jansen (**Microsoft**)

### Version history ###

Version  | Date | Comments
---------| -----| --------
2.6 | October 22th 2019 | Updated Workflow and InfoPath scan components + new blog scan component + updates to allow running with Sites.Read.All Azure AD App-Only permission
2.5 | June 27th 2019 | Beta release of the Workflow and InfoPath scan components + various small improvements and fixes
2.4 | February 16th 2019 | Performance improvements (40% faster for full scan) + embedded webpartmapping.xml file. Scanner distribution is now a single .exe
2.3 | December 12th 2018 | Using latest PnP Sites Core library with updated throttling implementation, several small reliability improvements, export SiteId in ModernizationSiteScanResults.csv, built-in check to see if there's a newer version available, generation of SitesWithCustomizations.csv
2.2 | November 9th 2018 | Updated complexity calculation for publishing portals, updated list of modern capable lists (850), improved telemetry and error handling, bug fixing
2.1 | October 24th 2018 | Publishing portal analysis improvements: detect custom page layouts, classify publishing portals in simple/medium/complex
2.0 | October 15th 2018 | Built-in wizard will help you configure the scan parameters...you can forget about these long complex command lines!
1.7 | October 14th 2018 | Integrated SharePoint UI Experience scanner results + bug fixes
1.6 | September 25th 2018 | Added support for scanning classic Publishing Portals, simplified Group Connection dashboard, Compiled as X64 to avoid memory constraints during large scans
1.5 | June 1st 2018 | Added generation of Excel based reports which make it easier to consume the generated data
1.4 | May 5th 2018 | Added web part mapping percentage in page scan + by default raw web part data is not exported + allow to skip search query for site/page usage information
1.3 | March 16th 2018 | Added site usage information
1.2 | March 7th 2018 | Reliability improvements
1.1 | January 31st 2018 | Performance and stability improvements + Page scanner component integrated
1.0 | January 19th 2018 | First main version

### Disclaimer ###

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

----------

See https://aka.ms/sppnp-modernizationscanner for all details on the scanner usage, generated reports and tips and tricks. To learn more about each Modernization scanner release check [the release notes](./Modernization&#32;Scanner&#32;release&#32;notes.md).

<img src="https://telemetry.sharepointpnp.com/sp-dev-modernization/Tools/sharepoint-modernizationscanner" />