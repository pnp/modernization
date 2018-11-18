import * as React from 'react';
import styles from './PageBanner.module.scss';
import { IPageBannerProps } from './IPageBannerProps';
import { IPageBannerState } from './IPageBannerState';
import { escape } from '@microsoft/sp-lodash-subset';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { sp, ClientSidePage, NavigationNode, spODataEntityArray } from '@pnp/sp';
import PostFeedbackDialog from './PostFeedback';
import { AppInsights } from "applicationinsights-js";


export default class PageBanner extends React.Component<IPageBannerProps, IPageBannerState> {

  constructor(props: IPageBannerProps) {
    super(props);

    this.KeepPage = this.KeepPage.bind(this);
    this.DiscardPage = this.DiscardPage.bind(this);

    this.state = {
      progressMessage: undefined,
      errorString: undefined
    };
  }

  public render(): React.ReactElement<IPageBannerProps> {

    const progressToDisplay: string = this.state.progressMessage;
    
    let pageName: string = '';
    if (this.props.sourcePage != undefined)
    {
      pageName = this.props.sourcePage.substring(this.props.sourcePage.lastIndexOf('/') + 1);
    }

    return (
      <div className={ styles.pageBanner }>
        <div className={ styles.container }>
          
          <div className={ styles.row }>
            <div className={ styles.column }>
              <div className={ styles.left }>
                <i className={`${styles.lefticon} ms-Icon ms-Icon--Lightbulb`} aria-hidden="true"></i>
                <DefaultButton description='Replace old page with this one' onClick={this.KeepPage}>Keep this page</DefaultButton>
              </div>
              <div className={ styles.left }>
                <DefaultButton description='This page is not kept' onClick={this.DiscardPage}>Discard this page</DefaultButton>
              </div>
              <div className={ styles.description }>This page was generated from <a target="_blank" data-interception="off" rel="noopener noreferrer" href={this.props.sourcePage}>{escape(pageName)}</a>. To learn more <a target="_blank" rel="noopener noreferrer" href={this.props.learnMoreUrl}>click here</a>.</div>
            </div>
          </div>

          <div className={ styles.row }>
            <div className={ styles.column }>
              <div className="ms-font-l">
                <div>
                  {(progressToDisplay) && <span className="ms-font-l">{progressToDisplay}</span>}
                </div>
              </div>
            </div>
          </div>

        </div>
      </div>
    );
  }

  private KeepPage() {
    this.setState({
      progressMessage: "Processing started..."
    });

    this.KeepPageOperation();
  }

  private DiscardPage() {
    const root = this.props.pageContext.site.absoluteUrl.replace(this.props.pageContext.site.serverRelativeUrl, '');
    const sourcePageName = this.props.sourcePage.replace(root, '');

    if (this.props.feedbackList)
    {
      // Show a dialog to collect feedback about why the page is not good
      const dialog: PostFeedbackDialog = new PostFeedbackDialog();
      dialog.FileName = sourcePageName;  
      dialog.ModernizationCenterUrl = this.props.modernizationCenterUrl;
      dialog.FeedbackList = this.props.feedbackList;

      dialog.show().then(() => {
        if (dialog.OKHit)
        {
          // If the user pressed OK on the dialog continue with the requested page deletion        
          this.discardPageOperation();
        }
      });
    }
    else
    {
      console.log("Collecting feedback was skipped because there's on feedback list configured");
      this.discardPageOperation();
    }
  }

  private  async KeepPageOperation()
  {
    const root = this.props.pageContext.site.absoluteUrl.replace(this.props.pageContext.site.serverRelativeUrl, '');
    const sourcePageName = this.props.sourcePage.replace(root, '');
    const sourceFileName = sourcePageName.substring(sourcePageName.lastIndexOf('/') + 1);
    const sourcePath = sourcePageName.substring(0, sourcePageName.lastIndexOf('/'));
    const targetPageName = this.props.targetPage.replace(root, '');
    const newSourcePage = `${sourcePath}/Old_${sourceFileName}`;

    console.log(`root: ${root}`);
    console.log(`sourcePageName: ${sourcePageName}`);
    console.log(`sourceFileName: ${sourceFileName}`);
    console.log(`sourcePath: ${sourcePath}`);
    console.log(`newSourcePage: ${newSourcePage}`);
    console.log(`targetPageName: ${targetPageName}`);

    // https://github.com/Microsoft/ApplicationInsights-JS/blob/master/API-reference.md
    AppInsights.trackEvent("TransformationService.KeepPage", {}, {});
    AppInsights.trackMetric("TransformationService.PagesKept", 1);
  
    // STEP1: First copy the source page to a new name. We on purpose use CopyTo as we want to avoid that "linked" url's get 
    //        patched up during a MoveTo operation as that would also patch the url's in our new modern page
    let sourcePage = await sp.web.getFileByServerRelativeUrl(sourcePageName);    
    await sourcePage.copyTo(newSourcePage, true);

    //Load the created target page
    let targetPage = await sp.web.getFileByServerRelativeUrl(targetPageName);

    // STEP2: Fix possible navigation entries to point to the "copied" source page first
    // Rename the target page to the original source page name
    // CopyTo and MoveTo with option to overwrite first internally delete the file to overwrite, which
    // results in all page navigation nodes pointing to this file to be deleted. Hence let's point these
    // navigation entries first to the copied version of the page we just created    
    let navWasFixed: boolean = false;
    let nodes = await sp.web.navigation.quicklaunch.filter(encodeURIComponent(`Url eq '${sourcePageName}'`)).get<NavigationNode[]>(spODataEntityArray(NavigationNode));
    if (nodes && nodes.length > 0)
    {
      navWasFixed = true;
      for (let node of nodes)
      {
        await (node as NavigationNode).update({Url: newSourcePage});
      }
    }
    nodes = await sp.web.navigation.topNavigationBar.filter(encodeURIComponent(`Url eq '${sourcePageName}'`)).get<NavigationNode[]>(spODataEntityArray(NavigationNode));
    if (nodes && nodes.length > 0)
    {
      navWasFixed = true;
      for (let node of nodes)
      {
        await (node as NavigationNode).update({Url: newSourcePage});
      }
    }

    // STEP3: Now copy the created modern page over the original source page, at this point the new page has the same name as the original page had before transformation
    await targetPage.copyTo(sourcePageName, true);

    // STEP4: Finish with restoring the page navigation: update the navlinks to point back the original page name
    if (navWasFixed)
    {
      nodes = await sp.web.navigation.quicklaunch.filter(encodeURIComponent(`Url eq '${newSourcePage}'`)).get<NavigationNode[]>(spODataEntityArray(NavigationNode));
      if (nodes && nodes.length > 0)
      {
        navWasFixed = true;
        for (let node of nodes)
        {
          await (node as NavigationNode).update({Url: sourcePageName});
        }
      }
      nodes = await sp.web.navigation.topNavigationBar.filter(encodeURIComponent(`Url eq '${newSourcePage}'`)).get<NavigationNode[]>(spODataEntityArray(NavigationNode));
      if (nodes && nodes.length > 0)
      {
        navWasFixed = true;
        for (let node of nodes)
        {
          await (node as NavigationNode).update({Url: sourcePageName});
        }
      }
    }

    //STEP5: Continue with deleting the originally created modern page as we did copy that already in step 3
    await targetPage.delete();

    //STEP6: Conclude with removing the banner web part from the page and then reload the page
    // Load the created page and remove the banner web part
    let page = await ClientSidePage.fromFile(sp.web.getFileByServerRelativeUrl(sourcePageName));      
    if (page)
    {
      // Banner web part should live in it's own top level section
      if (page.sections[0] && page.sections[0].defaultColumn.controls.length == 1)
      {
        // Drop the section
        // TODO: check for web part and only remove the section when there's only a banner web part
        page.sections[0].remove();
        await page.save();

        // Redirect to this final new page:
        window.location.href = `${root}${sourcePageName}`;
      }
      else
      {
        const error: string = `The banner web part on ${sourcePageName} could not be removed`;
        console.log(`Error during deletion of page: ${error}`);
        this.setState({ 
          errorString: error
        });
      }
    }

  }

  private async discardPageOperation()
  {
    const root = this.props.pageContext.site.absoluteUrl.replace(this.props.pageContext.site.serverRelativeUrl, '');
    const targetPageName = this.props.targetPage.replace(root, '');

    console.log(`discarding page: ${targetPageName}`);

    AppInsights.trackEvent("TransformationService.DiscardPage", {}, {});
    AppInsights.trackMetric("TransformationService.PagesDiscarded", 1);

    // grab the page to delete
    let targetPage = await sp.web.getFileByServerRelativeUrl(targetPageName);
    
    targetPage.delete()
    .then(() => {
      // Navigate back to the page library
      window.location.href = this.props.sourcePage.substring(0, this.props.sourcePage.lastIndexOf('/'));
    })
    .catch((error) => {
      console.log(`Error during deletion of page: ${error}`);
      this.setState({ 
        errorString: error
      });
    });
  }

}
