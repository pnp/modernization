<#
.SYNOPSIS
Enables page transformation UI integration for the currently connected site collection

.EXAMPLE
PS C:\> .\Enable-PageTransformation.ps1
#>

Param (
    [Parameter(Mandatory=$false, Position=0)]
    [String] $ModernizationCenterUrl    
)

# check modernization center url
if ($ModernizationCenterUrl -eq $null -or $ModernizationCenterUrl -eq "")
{
  $ModernizationCenterUrl = "/sites/modernizationcenter"
}

# Load the ServerRelativeUrl as we need to use that to embed the site collection url in the custom action
$site = Get-PnPSite -Includes ServerRelativeUrl

# Ensure the modern page feature is enabled
Write-Host "Enabling the modern page feature"
Enable-PnPFeature -Identity "B6917CB1-93A0-4B97-A84D-7CF49975D4EC" -Scope Web -Force

Write-Host "Enabling page transformation for $($site.ServerRelativeUrl)" -ForegroundColor White

# Add the site page library extensions
$command = '<CommandUIExtension><CommandUIDefinitions>
<CommandUIDefinition Location="Ribbon.Documents.Copies.Controls._children">
  <Button
    Id="Ribbon.Documents.Copies.ModernizePage"
    Command="SharePointPnP.Cmd.ModernizePage"
    Image16by16="{modernizationcenter}/siteassets/modernize16x16.png"
    Image32by32="{modernizationcenter}/siteassets/modernize32x32.png"
    LabelText="Create modern version"
    Description="Create a modern version of this page."
    ToolTipTitle="Create modern version"
    ToolTipDescription="Create a modern version of this page."
    TemplateAlias="o1"
    Sequence="15"/>
</CommandUIDefinition>
</CommandUIDefinitions>
<CommandUIHandlers>
<CommandUIHandler
  Command="SharePointPnP.Cmd.ModernizePage"
  CommandAction="{modernizationcenter}/SitePages/modernize.aspx?SiteUrl={SiteCollection}&amp;ListId={SelectedListId}&amp;ItemId={SelectedItemId}"
  EnabledScript="javascript:SP.ListOperation.Selection.getSelectedItems().length == 1;" />
</CommandUIHandlers></CommandUIExtension>'
$command = $command.Replace("{SiteCollection}", $site.ServerRelativeUrl);
$command = $command.Replace("{modernizationcenter}", $ModernizationCenterUrl);

Add-PnPCustomAction -Scope Site -Name "CA_PnP_Modernize_SitePages_RIBBON" -Title "Create modern version" -Description "Create a modern version of this page." `
                    -Location "CommandUI.Ribbon" -RegistrationType 1 -RegistrationId "119" -Rights EditListItems ` -Group " " `
                    -CommandUIExtension $command

$script = '{modernizationcenter}/SitePages/modernize.aspx?SiteUrl={SiteCollection}&ListId={ListId}&ItemId={ItemId}'
$script = $script.Replace("{SiteCollection}", $site.ServerRelativeUrl);
$script = $script.Replace("{modernizationcenter}", $ModernizationCenterUrl);

Add-PnPCustomAction -Scope Site -Name "CA_PnP_Modernize_SitePages_ECB" -Title "Create modern version" -Description "Create a modern version of this page." `
                    -Location "EditControlBlock" -RegistrationType 1 -RegistrationId "119" -Rights EditListItems ` -Group " " `
                    -Url $script

# Add the wiki page library ribbon
$command = '<CommandUIExtension>
<CommandUIDefinitions>
  <CommandUIDefinition Location="Ribbon.WikiPageTab.PageActions.Controls._children">
    <Button
      Id="Ribbon.WikiPageTab.PageActions.ModernizeWikiPage"
      Command="SharePointPnP.Cmd.ModernizeWikiPage"
      Image16by16="{modernizationcenter}/siteassets/modernize16x16.png"
      Image32by32="{modernizationcenter}/siteassets/modernize32x32.png"
      LabelText="Create modern version"
      Description="Create a modern version of this page."
      ToolTipTitle="Create modern version"
      ToolTipDescription="Create a modern version of this page."
      TemplateAlias="o1"
      Sequence="1500"/>
  </CommandUIDefinition>
</CommandUIDefinitions>
<CommandUIHandlers>
  <CommandUIHandler
    Command="SharePointPnP.Cmd.ModernizeWikiPage"
    CommandAction="javascript:function redirect(){ var url = ''{modernizationcenter}/SitePages/modernize.aspx?SiteUrl={SiteCollection}&#038;ListId='' + _spPageContextInfo.listId + ''&#038;ItemId='' + _spPageContextInfo.pageItemId; window.location = url; } redirect();" />
</CommandUIHandlers>
</CommandUIExtension>'
$command = $command.Replace("{SiteCollection}", $site.ServerRelativeUrl);
$command = $command.Replace("{modernizationcenter}", $ModernizationCenterUrl);

Add-PnPCustomAction -Scope Site -Name "CA_PnP_Modernize_WikiPage_RIBBON" -Title "Create modern version" -Description "Create a modern version of this page." `
                    -Location "CommandUI.Ribbon" -Rights EditListItems ` -Group " " `
                    -CommandUIExtension $command

# Add the web part page library ribbon
$command = '<CommandUIExtension>
<CommandUIDefinitions>
  <CommandUIDefinition Location="Ribbon.WebPartPage.Actions.Controls._children">
    <Button
      Id="Ribbon.WebPartPage.Actions.ModernizeWebPartPage"
      Command="SharePointPnP.Cmd.ModernizeWebPartPage"
      Image16by16="{modernizationcenter}/siteassets/modernize16x16.png"
      Image32by32="{modernizationcenter}/siteassets/modernize32x32.png"
      LabelText="Create modern version"
      Description="Create a modern version of this page."
      ToolTipTitle="Create modern version"
      ToolTipDescription="Create a modern version of this page."
      TemplateAlias="o1"
      Sequence="1500"/>
  </CommandUIDefinition>
</CommandUIDefinitions>
<CommandUIHandlers>
  <CommandUIHandler
    Command="SharePointPnP.Cmd.ModernizeWebPartPage"
    CommandAction="javascript:function redirect(){ var url = ''{modernizationcenter}/SitePages/modernize.aspx?SiteUrl={SiteCollection}&#038;ListId='' + _spPageContextInfo.listId + ''&#038;ItemId='' + _spPageContextInfo.pageItemId; window.location = url; } redirect();" />
</CommandUIHandlers>
</CommandUIExtension>'
$command = $command.Replace("{SiteCollection}", $site.ServerRelativeUrl);
$command = $command.Replace("{modernizationcenter}", $ModernizationCenterUrl);


Add-PnPCustomAction -Scope Site -Name "CA_PnP_Modernize_WebPartPage_RIBBON" -Title "Create modern version" -Description "Create a modern version of this page." `
                    -Location "CommandUI.Ribbon" -Rights EditListItems ` -Group " " `
                    -CommandUIExtension $command

# Add the script that will show a banner on the classic pages to indicate presence of a modern version
$command = '{modernizationcenter}/SiteAssets/pnppagetransformationclassicbanner.js?rev=beta.1'
$command = $command.Replace("{modernizationcenter}", $ModernizationCenterUrl);

Add-PnPJavaScriptLink -Scope Site -Key "CA_PnP_Modernize_ClassicBanner" `
                      -Sequence 1000 -Url $command

Write-Host "Done" -ForegroundColor Green

