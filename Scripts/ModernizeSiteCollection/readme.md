# Sample that modernizes a site

## Summary

This script shows how to modernize a site by group connecting the site, modernizing the pages, configuring customizations and much more.

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
modernizesitecollection.ps1 | Bert Jansen (**Microsoft**)

## Version history

Version|Date|Comments
-------|----|--------
1.1 | December 18th 2018 | Updated to use the PnP PowerShell option for transforming pages
1.0 | November 26th 2018 | Initial commit

## Disclaimer

THIS CODE IS PROVIDED AS IS WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANT ABILITY, OR NON-INFRINGEMENT.

---

## Minimal Path to Awesome

- Update $tenantAdminUrl with your tenant admin URL
- Update $credentialManagerCredentialToUse with your credential manager entry...if not specified tenant admin credentials are asked by the script

<img src="https://telemetry.sharepointpnp.com/sp-dev-modernization/scripts/modernizesitecollection" />