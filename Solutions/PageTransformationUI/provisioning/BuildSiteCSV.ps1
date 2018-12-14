<#
.SYNOPSIS
Shows how to obtain a list of site collections based upon a filter and then export them to a CSV file

.EXAMPLE
PS C:\> .\BuildSiteCSV.ps1
#>

# List of sites that will be exported to CSV
$array = @()
$csvFileName = ".\sites.csv"

# Get a list of sites, e.g. filter on Template
$sites = get-pnptenantsite | where-object {$_.Template -eq "STS#0" -and $_.Url -like "*espc*"}

foreach($site in $sites)
{
    #if ($site.Url -like "*espctest*")
    #{        
        $array += New-Object PsObject -property @{'SiteCollectionUrl' = $site.Url }
    #}
}

$array | export-csv -NoTypeInformation -Path $csvFileName