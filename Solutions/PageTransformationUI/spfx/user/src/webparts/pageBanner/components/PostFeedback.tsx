import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import styles from './PageBanner.module.scss';
import {
    autobind,
    PrimaryButton,
    CommandButton,
    TextField,
    Dropdown,
    IDropdownOption,
    DialogFooter,
    DialogContent
  } from 'office-ui-fabric-react';
import { ItemAddResult, Web } from '@pnp/sp';

interface IPostFeedbackDialogContentProps {
    fileName: string;
    close: () => void;
    submit: (rejectionCategory: string, rejectionComments: string) => void;
}
  
interface IPostFeedbackDialogContentState {
    rejectionCategory: string;
    rejectionComments: string;
}


class PostFeedbackDialogContent extends React.Component<IPostFeedbackDialogContentProps, IPostFeedbackDialogContentState> {

    constructor(props) {
        super(props);

        this.state = {
            rejectionCategory: "The text formatting is wrong",
            rejectionComments: "",
        };
    }

    public render(): JSX.Element {
        return (<div className={ styles.postFeedbackDialogRoot } >
            <DialogContent
                title={ "Provide feedback" }
                subText={ "Please let us know why the generated page is not good by filling out below form" }
                onDismiss={ this.props.close }
                showCloseButton={ true }                                 
                >

                <div className={ styles.postFeedbackDialogContent }>
                    <div className="ms-Grid-row">
                        <div className="ms-Grid-col ms-u-lg12">
                        <Dropdown
                            label={ "Select a reason" }
                            required={ true }
                            defaultSelectedKey='1'
                            options={
                            [
                                { key: '1', text: 'The text formatting is wrong' },
                                { key: '2', text: 'There are missing web parts on this page' },
                                { key: '3', text: 'The web parts are not configured correctly' },
                                { key: '4', text: 'Other reasons' },
                            ]
                            }
                            onChanged={ this._onChangedRejectionReason }
                        />
                        </div>
                    </div> 
                        <div className="ms-Grid-row">
                            <div className={'ms-Grid-col ms-u-lg12'}>
                                <TextField 
                                    label={ "Optionally provide more feedback" } 
                                    required={ false } 
                                    value={ this.state.rejectionComments }
                                    multiline = { true }                                
                                    onChanged={ this._onChangedFeedback }     
                                />
                            </div>
                        </div>
                </div>    

                <DialogFooter>
                    <CommandButton text='Cancel' title='Cancel' onClick={ this.props.close } />
                    <PrimaryButton text='OK' title='OK' onClick={() => { this.props.submit(this.state.rejectionCategory, this.state.rejectionComments); }} />
                </DialogFooter>
            </DialogContent>
        </div>);
    }

    @autobind
    private _onChangedFeedback(feedback: string): void {
      this.setState({
        rejectionComments: feedback,
      });
    }

    @autobind
    private _onChangedRejectionReason(option: IDropdownOption, index?: number): void {
      this.setState({
        rejectionCategory: option.text
      });
    }
}

export default class PostFeedbackDialog extends BaseDialog {

    public FileName: string;
    public ModernizationCenterUrl: string;
    public FeedbackList: string;
    public OKHit: boolean = false;

    protected onBeforeOpen(): Promise<void>{
        return super.onBeforeOpen().then(_ => {  

        });
    }

    public render(): void {
        ReactDOM.render(<PostFeedbackDialogContent
          fileName={ this.FileName }
          close={ this.close }
          submit={ this._submit }
        />, this.domElement);
      }

      public getConfig(): IDialogConfiguration {
        return {
          isBlocking: false,
        };
      }

      // Workaround for being able to open the dialog multiple times with SPFX 1.7 / React 16
      protected onAfterClose(): void {
        super.onAfterClose();
        ReactDOM.unmountComponentAtNode(this.domElement);
      }

      @autobind
      private async _submit(rejectionCategory: string, rejectionComments: string): Promise<void> {
        console.log(`rejectionCategory = ${rejectionCategory} rejectionComments = ${rejectionComments}`);

        // Write results to a central list  
        const url = `${PostFeedbackDialog.getAbsoluteDomainUrl()}${this.ModernizationCenterUrl}`;
        console.log(`Connecting to ${url}`);
        let modernizationCenterWeb = new Web(url);

        modernizationCenterWeb.get().then(w => {

            // Get current month and day in UTC as we're structuring the feedback like that
            const year: number = new Date().getUTCFullYear();
            const month: number = new Date().getUTCMonth();
            const day: number = new Date().getUTCDay();
            
            modernizationCenterWeb.lists.getByTitle(this.FeedbackList).items.add(
                {
                    Title: "PageTransformation",
                    Year: year,
                    Month: month,
                    Day: day,
                    FeedbackCategory: rejectionCategory,
                    Feedback: rejectionComments,
                    ModernizationSubject: this.FileName
                }
            )
            .then((iar: ItemAddResult) => {                
                console.log("done");
            })
            .catch((error: any) => {
                console.log(`Error writing feedback: ${error}`);
            });
        });
        
        // close the dialog after we've saved our feedback
        this.OKHit = true;
        this.close();
      }

      private static getAbsoluteDomainUrl(): string 
      {
        if (window
            && "location" in window
            && "protocol" in window.location
            && "host" in window.location) {
            return window.location.protocol + "//" + window.location.host;
        }
        return null;
      }      
}
