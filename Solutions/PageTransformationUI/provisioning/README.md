# Provisioning scripts

This folder contains the script you need to provision and configure the Page Transformation UI solution. When you use the .xml template you'll also need the assets folder present as the .xml version of the template refers to assets. It's often easier to use the provided .pnp template or package your own version of the .pnp file via:

```Powershell
$t = Read-PnPTenantTemplate modernization.xml
Save-PnPTenantTemplate -Template $t -Out modernization.pnp
```