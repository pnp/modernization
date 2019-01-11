import * as React from 'react';
import styles from './PageTransformatorAdmin.module.scss';
import { IPageTransformatorAdminProps, IPageTransformatorAdminState } from './IPageTransformatorAdminProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp, Site } from '@pnp/sp';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { DefaultButton, PrimaryButton, MessageBarButton } from 'office-ui-fabric-react/lib/Button';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import { UserCustomAction } from '@pnp/sp/src/usercustomactions';

export default class PageTransformatorAdmin extends React.Component<IPageTransformatorAdminProps, IPageTransformatorAdminState> {

  private SITEPAGESFEATUREID: string = 'B6917CB1-93A0-4B97-A84D-7CF49975D4EC';
  private CAPNPMODERNIZESITEPAGESECB: string = "CA_PnP_Modernize_SitePages_ECB";
  private CAPNPMODERNIZESITEPAGES: string = "CA_PnP_Modernize_SitePages_RIBBON";
  private CAPNPMODERNIZEWIKIPAGE: string = "CA_PnP_Modernize_WikiPage_RIBBON";
  private CAPPNPMODERNIZEWEBPARTPAGE: string = "CA_PnP_Modernize_WebPartPage_RIBBON";
  private CAPNPCLASSICPAGEBANNER: string = "CA_PnP_Modernize_ClassicBanner";

  constructor(props: IPageTransformatorAdminProps, state: IPageTransformatorAdminState) {
    super(props);

    this.state = {
      siteUrl: "",
      buttonsDisabled: true,
      resultMessage: null,
      resultMessageType: MessageBarType.info
    };

    sp.setup({
      spfxContext: this.props.context,
    });

    this.onSiteUrlChange = this.onSiteUrlChange.bind(this);
  }

  private enableClick = (): void => {
    const centerUrl = this.props.context.pageContext.site.serverRelativeUrl;
    let site = new Site(this.getTenantUrl() + this.state.siteUrl);

    const sitePageLibraryCA: any = {
      Description: "Create a modern version of this page.",
      Name: this.CAPNPMODERNIZESITEPAGESECB,
      Location: "EditControlBlock",
      RegistrationType: 1,
      RegistrationId: "119",
      Rights: { High: "0", Low: "4" },
      Title: "Create modern version",
      Url: `${centerUrl}/SitePages/modernize.aspx?SiteUrl=/${this.state.siteUrl}&ListId={ListId}&ItemId={ItemId}`
    };

    // Wiki Page library user custom actions
    const wikiPageLibraryCA: any = {
      Description: "Create a modern version of this page.",
      Name: this.CAPNPMODERNIZESITEPAGES,
      Location: "CommandUI.Ribbon",
      RegistrationType: 1,
      RegistrationId: "119",
      Rights: { High: "0", Low: "4" },
      Title: "Create modern version",
      CommandUIExtension: `<CommandUIExtension><CommandUIDefinitions>
      <CommandUIDefinition Location="Ribbon.Documents.Copies.Controls._children">
        <Button
          Id="Ribbon.Documents.Copies.ModernizePage"
          Command="SharePointPnP.Cmd.ModernizePage"
          Image16by16="${centerUrl}/siteassets/modernize16x16.png"
          Image32by32="${centerUrl}/siteassets/modernize32x32.png"
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
        CommandAction="${centerUrl}/SitePages/modernize.aspx?SiteUrl=/${this.state.siteUrl}&amp;ListId={SelectedListId}&amp;ItemId={SelectedItemId}"
        EnabledScript="javascript:SP.ListOperation.Selection.getSelectedItems().length == 1;" />
    </CommandUIHandlers></CommandUIExtension>`
    };

    // Wiki page ribbon user custom action
    const wikiPageRibbonCA: any = {
      Description: "Create a modern version of this page.",
      Name: this.CAPNPMODERNIZEWIKIPAGE,
      Location: "CommandUI.Ribbon",
      Title: "Create modern version",
      Rights: { High: "0", Low: "4" },
      CommandUIExtension: `<CommandUIExtension>
      <CommandUIDefinitions>
        <CommandUIDefinition Location="Ribbon.WikiPageTab.PageActions.Controls._children">
          <Button
            Id="Ribbon.WikiPageTab.PageActions.ModernizeWikiPage"
            Command="SharePointPnP.Cmd.ModernizeWikiPage"
            Image16by16="${centerUrl}/siteassets/modernize16x16.png"
            Image32by32="${centerUrl}/siteassets/modernize32x32.png"
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
          CommandAction="javascript:function redirect(){ var url = '${centerUrl}/SitePages/modernize.aspx?SiteUrl=/${this.state.siteUrl}&#038;ListId=' + _spPageContextInfo.listId + '&#038;ItemId=' + _spPageContextInfo.pageItemId; window.location = url; } redirect();" />
      </CommandUIHandlers>
    </CommandUIExtension>`
    };

    // Web part page ribbon user custom action
    const webpartPageRibbonCA: any = {
      Description: "Create a modern version of this page.",
      Name: this.CAPPNPMODERNIZEWEBPARTPAGE,
      Location: "CommandUI.Ribbon",
      Title: "Create modern version",
      Rights: { High: "0", Low: "4" },
      CommandUIExtension: `<CommandUIExtension>
      <CommandUIDefinitions>
        <CommandUIDefinition Location="Ribbon.WebPartPage.Actions.Controls._children">
          <Button
            Id="Ribbon.WebPartPage.Actions.ModernizeWebPartPage"
            Command="SharePointPnP.Cmd.ModernizeWebPartPage"
            Image16by16="${centerUrl}/siteassets/modernize16x16.png"
            Image32by32="${centerUrl}/siteassets/modernize32x32.png"
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
          CommandAction="javascript:function redirect(){ var url = '${centerUrl}/SitePages/modernize.aspx?SiteUrl=/${this.state.siteUrl}&#038;ListId=' + _spPageContextInfo.listId + '&#038;ItemId=' + _spPageContextInfo.pageItemId; window.location = url; } redirect();" />
      </CommandUIHandlers>
    </CommandUIExtension>`
    };

    // Classic page banner user custom action
    const classicPageBannerCA: any = {
      Description: "Shows a banner on the classic pages.",
      Name: this.CAPNPCLASSICPAGEBANNER,
      Location: "ScriptLink",
      //RegistrationType: 1,
      Title: "Shows a banner on the classic pages",
      ScriptSrc: `${this.getTenantUrl()}${centerUrl}/SiteAssets/pnppagetransformationclassicbanner.js?rev=beta.1`
    };

    Promise.all([
      this.activateSitePagesFeatureIfNeeded(site),
      this.addCustomActionIfNeeded(site, sitePageLibraryCA),
      this.addCustomActionIfNeeded(site, wikiPageLibraryCA),
      this.addCustomActionIfNeeded(site, wikiPageRibbonCA),
      this.addCustomActionIfNeeded(site, webpartPageRibbonCA),
      this.addCustomActionIfNeeded(site, classicPageBannerCA)
    ]).then(() => {
      this.setState((current) => ({ ...current, resultMessage: "Modernization functionality added to site", resultMessageType: MessageBarType.success }));
    }).catch(() => {
      this.setState((current) => ({ ...current, resultMessage: "Unable to add mo0dernization functionality to site", resultMessageType: MessageBarType.error }));
    });
  }

  private disableClick = (): void => {
    const centerUrl = this.props.context.pageContext.site.serverRelativeUrl;
    let site = new Site(this.getTenantUrl() + this.state.siteUrl);

    Promise.all([
      this.activateSitePagesFeatureIfNeeded(site),
      this.removeCustomActionIfNeeded(site, this.CAPNPMODERNIZESITEPAGESECB),
      this.removeCustomActionIfNeeded(site, this.CAPNPMODERNIZESITEPAGES),
      this.removeCustomActionIfNeeded(site, this.CAPNPMODERNIZEWIKIPAGE),
      this.removeCustomActionIfNeeded(site, this.CAPPNPMODERNIZEWEBPARTPAGE),
      this.removeCustomActionIfNeeded(site, this.CAPNPCLASSICPAGEBANNER)
    ]).then(() => {
      this.setState((current) => ({ ...current, resultMessage: "Modernization functionality removed from site", resultMessageType: MessageBarType.success }));
    }).catch(() => {
      this.setState((current) => ({ ...current, resultMessage: "Unable to remove modernization functionality from site", resultMessageType: MessageBarType.error }));
    });
  }

  private activateSitePagesFeatureIfNeeded(site: Site): Promise<void> {
    if (!site.rootWeb.features.getById(this.SITEPAGESFEATUREID)) {
      site.rootWeb.features.add(this.SITEPAGESFEATUREID);
    }
    return new Promise<void>((resolve) => { resolve(); });
  }

  private addCustomActionIfNeeded(site: Site, customAction: any): Promise<void> {
    return new Promise<void>((resolve) => {
      site.userCustomActions.filter(`Name eq '${customAction.Name}'`).get().then((customActions: any[]) => {
        if (customActions.length == 0) {
          site.userCustomActions.add(customAction).then(_ => { resolve(); });
        } else {
          resolve();
        }
      });
    });
  }

  private removeCustomActionIfNeeded(site: Site, name: string): Promise<void> {
    return new Promise<void>((resolve) => {
      site.userCustomActions.filter(`Name eq '${name}'`).get().then((customActions: any[]) => {
        if (customActions.length == 1) {
          site.userCustomActions.getById(customActions[0].Id).delete().then(_ => { resolve(); });
        } else {
          resolve();
        }
      });
      resolve();
    });
  }

  private getTenantUrl(): string {
    const siteUrl = this.props.context.pageContext.site.absoluteUrl;
    const slashPos = siteUrl.indexOf("/", 9);
    if (slashPos > 0) {
      return siteUrl.substring(0, slashPos) + "/";
    } else {
      return siteUrl + "/";
    }
  }

  private onSiteUrlChange = (text: string): void => {
    this.setState({ siteUrl: text, buttonsDisabled: this.state.siteUrl === "", resultMessage: null });
  }

  public render(): React.ReactElement<IPageTransformatorAdminProps> {
    return (
      <div className={styles.pageTransformatorAdmin}>
        <span className="ms-font-xxl">Enable Site for Page Transformation</span>
        <TextField label="Url of Site" prefix={this.getTenantUrl()} value={this.state.siteUrl} onChanged={this.onSiteUrlChange} />
        <div className={styles.actions}>
          <PrimaryButton text="Enable" onClick={this.enableClick} disabled={this.state.buttonsDisabled}></PrimaryButton>
          <PrimaryButton text="Disable" onClick={this.disableClick} disabled={this.state.buttonsDisabled}></PrimaryButton>
        </div>
        <div className={styles.messages}>
          {this.state.resultMessage != null && <MessageBar
            messageBarType={this.state.resultMessageType}
            isMultiline={false}
          >{this.state.resultMessage}</MessageBar>}
          </div>
      </div >
    );
  }
}
