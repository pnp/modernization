
#IMPORTANT: this requires the PnP PowerShell version 3.5.1812 (December 2018) or higher to work!

# Connect to the web holding the pages to modernize
Connect-PnPOnline -Url https://bertonline.sharepoint.com/sites/modernizationtestpages -Credentials bertonline -Verbose

# Get all the pages in the site pages library
$pages = Get-PnPListItem -List sitepages

# Iterate over the pages
foreach($page in $pages) 
{ 
    # Optionally filter the pages you want to modernize
    if ($page.FieldValues["FileLeafRef"].StartsWith(("t")))
    {
        # No need to convert modern pages again
        if ($page.FieldValues["ClientSideApplicationId"] -eq "b6917cb1-93a0-4b97-a84d-7cf49975d4ec" ) 
        { 
            Write-Host `Page $page.FieldValues["FileLeafRef"] is modern, no need to modernize it again`
        } 
        else 
        { 
            # Create a modern version of this page
            Write-Host `Modernizing $page.FieldValues["FileLeafRef"]...`
            $modernPage = ConvertTo-PnPClientSidePage -Identity $page.FieldValues["FileLeafRef"] -Overwrite
            Write-Host "Done" -ForegroundColor Green
        }
    }  
}


