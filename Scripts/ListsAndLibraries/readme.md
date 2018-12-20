# Modern List and library scripts

## Summary

Modern lists and libraries are enabled for most tenants and that's typically the best option. You however might want to disable the modern list and library experience for certain site collections, which can be done by enabling or disabling a site collection modern list blocking feature.

## Applies to

- Office 365 Multi-Tenant (MT)

## Prerequisites

- SharePoint PnP PowerShell

## Solution

Solution|Author(s)
--------|---------
SetModernListUsage.ps1 | Bert Jansen (**Microsoft**)

## Version history

Version|Date|Comments
-------|----|--------
1.0 | December 12th 2018 | Initial commit

## Disclaimer

THIS CODE IS PROVIDED AS IS WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANT ABILITY, OR NON-INFRINGEMENT.

---

## SetModernListUsage.ps1

This scripts will enable or disable the modern list and library experience for the given site collection(s). The script can be run for either a single site collection or for list of site collections provided via a CSV file. Simply run the script and provide the needed input on the command prompt.

### Structure for the CSV file

The CSV file is simple list of site collection URL's without a header as shown in below sample:

```Text
"https://contoso.sharepoint.com/sites/siteA"
"https://contoso.sharepoint.com/sites/siteB"
"https://contoso.sharepoint.com/sites/siteC"
```

### How do I know for which site collections it makes sense to disable the modern list and library experience?

Most customers have modern lists and libraries enabled across the board, which will give them the best experience. However some customers might be using incompatible customizations in their lists, which would be a reason to keep the site collections using these customizations in classic. If you do know the site collections having incompatible customizations by heart then you can manually craft the needed CSV file, but often it's better to run the [SharePoint Modernization scanner](https://aka.ms/sppnp-modernizationscanner) in the **"Modern list experience readiness"** mode as that will output all site collections holding incompatible customizations into a CSV file named **SitesWithCustomizations.csv**. You can then use this file to drive the script to enable/disable the modern list and library experience. See also [https://aka.ms/sppnp-modernlistoptout](https://aka.ms/sppnp-modernlistoptout) for an end-to-end view on options to opt out lists and libraries from the modern experience.