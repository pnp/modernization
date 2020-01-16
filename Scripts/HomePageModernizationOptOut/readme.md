# Modern List and library scripts

## Summary

Uncustomized classic home pages of team sites will be automatically upgraded to a modern home page. Use this script to do a bulk enabling/disabling of a web scoped feature that prevents this automatic upgrade from happening.

## Applies to

- Office 365 Multi-Tenant (MT)

## Prerequisites

- SharePoint PnP PowerShell

## Solution

Solution|Author(s)
--------|---------
HomePageModernizationOptOut.ps1 | Bert Jansen (**Microsoft**)

## Version history

Version|Date|Comments
-------|----|--------
1.0 | January 16th 2020 | Initial commit

## Disclaimer

THIS CODE IS PROVIDED AS IS WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANT ABILITY, OR NON-INFRINGEMENT.

---

## HomePageModernizationOptOut.ps1

This scripts will enable or disable the automatic upgrade of uncustomized classic home pages to a modern home page. The script can handle a single site or a list of sites provided via a CSV file. To get the CSV file you can run the Modernization Scanner, version 2.8 or higher, and use the "Wiki/Webpart Page transformation readiness (home pages)" mode (see https://aka.ms/sppnp-modernizationscanner) or alternatively create the file yourselves. Simply run the script and provide the needed input on the command prompt.

### Structure for the CSV file

The CSV file is simple list of site URL's without a header as shown in below sample:

```Text
"https://contoso.sharepoint.com/sites/siteA"
"https://contoso.sharepoint.com/sites/siteB"
"https://contoso.sharepoint.com/sites/siteB/subsite1"
"https://contoso.sharepoint.com/sites/siteB/subsite2"
```
