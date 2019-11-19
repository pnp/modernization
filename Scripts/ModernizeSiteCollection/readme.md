# Sample that modernizes site collections

## Summary

These scripts shows how to modernize a site collections by group connecting the site, modernizing the pages, cleaning up branding, creating a Teams team and much more.

## Windows Credential Manager can be used for automation scenarios

The sample uses [Windows Credential Manager](https://github.com/SharePoint/PnP-PowerShell/wiki/How-to-use-the-Windows-Credential-Manager-to-ease-authentication-with-PnP-PowerShell) for getting a credential in scenarios when full automation is needed. If label is not found in the Windows Credential Manager then prompt would ask for tenant admin email and password.

## Applies to

- Office 365 Multi-Tenant (MT)

## Prerequisites

- SharePoint PnP PowerShell
- Azure PowerShell

## Solution

Solution|Author(s)
--------|---------
ModernizeSitecollections.ps1 and ValidateSiteCollectionsInput.ps1 | Bert Jansen (**Microsoft**)

## Version history

Version|Date|Comments
-------|----|--------
2.0 | November 19th 2019 | Updated version that performs a bulk site collection modernization
1.1 | December 18th 2018 | Updated to use the PnP PowerShell option for transforming pages
1.0 | November 26th 2018 | Initial commit

## Disclaimer

THIS CODE IS PROVIDED AS IS WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANT ABILITY, OR NON-INFRINGEMENT.

---

## Minimal Path to Awesome

### I want to upgrade a single site collection

- Open a PowerShell session and run ModernizeSiteCollections.ps1 and provide the requested input

### I want to upgrade multiple site collections

- Create a CSV file like shown in the sample sitecollections.csv file
- Open a PowerShell session and run the ValidateSiteCollectionsInput.ps1 script and provide the needed input
- If the result of the validation is good then continue to the next step
- Run the ModernizeSiteCollections.ps1 script and provide the path to the csv file holding the site collections to modernize

<img src="https://telemetry.sharepointpnp.com/sp-dev-modernization/scripts/modernizesitecollection" />