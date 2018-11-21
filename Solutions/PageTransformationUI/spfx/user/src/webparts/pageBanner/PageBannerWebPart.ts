import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneLabel,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'PageBannerWebPartStrings';
import PageBanner from './components/PageBanner';
import { IPageBannerProps } from './components/IPageBannerProps';
import { sp } from '@pnp/sp';
import { IStorageEntities } from './components/IStorageEntities';
import { AppInsights } from "applicationinsights-js";

export interface IPageBannerWebPartProps {
  sourcePage: string;
  targetPage: string;
}

export default class PageBannerWebPart extends BaseClientSideWebPart<IPageBannerWebPartProps> {
  private debug: boolean = true;
  private modernizationCenterUrl: string = "";
  private feedbackList: string = "";
  private learnMoreUrl: string = "";

  public onInit(): Promise<void> {
    
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      
      // setup telemetry
      AppInsights.downloadAndSetup({ instrumentationKey: "373400f5-a9cc-48f3-8298-3fd7f4c063d6" });

      // Configure SPPnPJS
      sp.setup({
        spfxContext: this.context
      });

      // Load storage entities
      this.LoadConfiguration().then((settings: IStorageEntities)=> {
          this.modernizationCenterUrl = settings.modernizationCenterUrl.Value;
          this.feedbackList = settings.feedbackList.Value;

          if (settings.learnMoreUrl.Value)
          {
            this.learnMoreUrl = settings.learnMoreUrl.Value;
          }
          else
          {
            this.learnMoreUrl = "https://aka.ms/sppnp-pagetransformationui-manual";
          }

          resolve();
        }, err => reject(err));
      });

    }

  public render(): void {
    const element: React.ReactElement<IPageBannerProps> = React.createElement(
      PageBanner,
      {
        pageContext: this.context.pageContext,
        sourcePage: this.properties.sourcePage,
        targetPage: this.properties.targetPage,
        modernizationCenterUrl: this.modernizationCenterUrl,
        feedbackList: this.feedbackList,
        learnMoreUrl: this.learnMoreUrl,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private async LoadConfiguration(): Promise<IStorageEntities>
  {
    // Set-PnPStorageEntity -Key "Modernization_CenterUrl" -Value "/sites/modernizationcenter" -Description "Site relative URL for the modernization center in this tenant"
    // Set-PnPStorageEntity -Key "Modernization_FeedbackList" -Value "ModernizationFeedback" -Description "Name of the created feedback list"    
    // Set-PnPStorageEntity -Key "Modernization_LearnMoreUrl" -Value "https://aka.ms/sppnp-pagetransformationui-manual" -Description "Url shown in the learn more link"  
    const result: IStorageEntities = { 
      modernizationCenterUrl :  await sp.web.getStorageEntity("Modernization_CenterUrl"),
      feedbackList : await sp.web.getStorageEntity("Modernization_FeedbackList"),
      learnMoreUrl:  await sp.web.getStorageEntity("Modernization_LearnMoreUrl")
    };

    // Ensure all promises are fullfilled before returning
    Promise.all([result.modernizationCenterUrl, result.feedbackList, result.learnMoreUrl]);

    return result;
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {

    if (this.debug)
    {
      return {
        pages: [
          {
            header: {
              description: strings.PropertyPaneDescription
            },
            groups: [
              {
                groupName: strings.BasicGroupName,
                groupFields: [
                  PropertyPaneTextField('sourcePage', {
                    label: strings.SourcePageFieldLabel
                  }),
                  PropertyPaneTextField('targetPage', {
                    label: strings.TargetPageFieldLabel
                  })
                ]
              }
            ]
          }
        ]
      };

    }
    else
    {
      return {
        pages: [
          {
            header: {
              description: strings.PropertyPaneDescription
            },
            groups: [
              {
                groupName: strings.BasicGroupName,
                groupFields: [
                  PropertyPaneLabel('sourcePage', {
                    text: this.properties.sourcePage             
                  }),
                  PropertyPaneLabel('targetPage', {
                    text: this.properties.targetPage
                  })
                ]
              }
            ]
          }
        ]
      };
    }
  }

}
