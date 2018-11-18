import * as React from 'react';
import styles from './PageTransformator.module.scss';
import { IPageTransformatorProps } from './IPageTransformatorProps';
import { IPageTransformatorState } from './IPageTransformatorState';
import { escape } from '@microsoft/sp-lodash-subset';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { AadHttpClient, HttpClientResponse, HttpClient } from '@microsoft/sp-http';


export default class PageTransformator extends React.Component<IPageTransformatorProps, IPageTransformatorState> {

  private modernizationApiEndPoint: string = '';

  constructor(props: IPageTransformatorProps) {
    super(props);

    // setup state
    this.state = { editMode: false, errorString: undefined };

    // bind events
    this.goBackOrClose = this.goBackOrClose.bind(this);

    // construct API url
    this.modernizationApiEndPoint = props.modernizationApi + `?SiteUrl=${encodeURIComponent(this.props.siteUrl)}&PageUrl=${encodeURIComponent(this.props.pageUrl)}`;
    console.log(`Url to hit: ${this.modernizationApiEndPoint}`);
  }

  public render(): React.ReactElement<IPageTransformatorProps> {  
    
    const errorToDisplay: string = this.state.errorString;

    if (this.props.editMode === true)
    {
      return(
        <div className={styles.pageTransformator}>
          <div className={styles.container}>
            <div className={`ms-Grid-row ms-fontColor-black ${styles.row}`}>
                <div className="ms-Grid-col ms-lg10 ms-xl8 ms-xlPush2 ms-lgPush1">
                  <span className="ms-font-xl ms-fontColor-black">Edit mode...nothing to see right now</span>
                </div>
            </div>
          </div>
        </div>        
      );
    } 
    else if (!this.props.storageEntitiesSet)
    {
      let azureADApp = "";
      let functionHost = "";
      let pageTransformationEndpoint = "";
      if (this.props.storageEntities)
      {
        azureADApp = this.props.storageEntities.azureADApp;
        functionHost = this.props.storageEntities.functionHost;
        pageTransformationEndpoint = this.props.storageEntities.pageTransformationEndpoint;
      }

      return(
        <div className={styles.pageTransformator}>
          <div className={styles.container}>
            <div className={`ms-Grid-row ms-fontColor-black ${styles.row}`}>
                <div className="ms-Grid-col ms-lg10 ms-xl8 ms-xlPush2 ms-lgPush1">
                  <span className="ms-font-xl ms-fontColor-black">Not all parameters are configured, can't execute. Please run the setup steps.</span>
                  <p><span className="ms-font-l ms-fontColor-black">AzureAD application ID: {escape(azureADApp)}</span></p>
                  <p><span className="ms-font-l ms-fontColor-black">AzureAD function host: {escape(functionHost)}</span></p>
                  <p><span className="ms-font-l ms-fontColor-black">Page transformation endpoint: {escape(pageTransformationEndpoint)}</span></p>
                </div>
            </div>
          </div>
        </div>        
      );
    }
    else if (this.props.pageUrl == null)
    {
      // Page was already a modern page, hence no modernization needed
      return(
        <div className={styles.pageTransformator}>
          <div className={styles.container}>
            <div className={`ms-Grid-row ms-fontColor-black ${styles.row}`}>
                <div className="ms-Grid-col ms-lg10 ms-xl8 ms-xlPush2 ms-lgPush1">
                  <span className="ms-font-xl ms-fontColor-black">This already is a modern page, no need to modernize it again.</span>
                  <p className="ms-font-l ms-fontColor-white"></p>
                  <DefaultButton description='Go back' onClick={this.goBackOrClose}>Go back</DefaultButton>
                </div>
            </div>
          </div>
        </div>
      );    
    }
    else
    {
      // Inform the user that the page is being modernized
      console.log(`in Render(). Errorstring = ${errorToDisplay}`);
      return (
        <div className={styles.pageTransformator}>
          <div className={styles.container}>                            
            <div className={`ms-Grid-row ms-fontColor-black ${styles.row}`}>
                <div className="ms-Grid-col ms-lg10 ms-xl8 ms-xlPush2 ms-lgPush1">
                    {(!errorToDisplay) && <Spinner size={SpinnerSize.large} />}
                    <div className={styles.centeralign}>
                      <span className="ms-font-m ms-fontColor-black">Busy generating a modern version of {escape(this.props.pageUrl.substring(this.props.pageUrl.lastIndexOf('/') + 1))}...</span>
                      {(errorToDisplay) && <div className="ms-font-m ms-fontColor-red">Error: <span>{escape(errorToDisplay)}</span></div>}
                    </div>
                </div>
            </div>
          </div>
        </div>
      );
    }
  }

  public componentDidMount(): void {

    // if pageUrl = null then the page is modern page, no need to call the modernization endpoint
    if (this.props.pageUrl != null)
    {
      // Launch the page modernization, only when there was no error 
      if (this.state.errorString == undefined)
      {
        this.modernizePageServiceCall();
      }
    }
  }

  private goBackOrClose(): void
  {
    // If a new tab was opened then close the tab
    if (window.history.length == 1)
    {
      window.close();
    }
    else
    {
      // Else go back to the previous page in the history
      window.history.go(-1);
    }    
  }

  private modernizePageServiceCall(): void
  {
    this.props.modernizationClient.get(this.modernizationApiEndPoint, AadHttpClient.configurations.v1)    
    .then((response: HttpClientResponse) => {
      response.json().then((responseString: string) => {
        if (response.ok) {
          window.location.href = responseString;
        }
        else {
          console.log(`modernizePageServiceCall() failed: ${responseString}`);
          this.setState({
            errorString: responseString
          });    
        }
      })
      .catch((error: any) => {
        console.log(`modernizePageServiceCall() failed: ${error}`);
        this.setState({
          errorString: error
        });    
      });
    });
  }  

}
