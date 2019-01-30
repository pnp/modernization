import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, Environment, EnvironmentType, DisplayMode } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'PageTransformatorWebPartStrings';
import PageTransformator from './components/PageTransformator';
import { IPageTransformatorProps } from './components/IPageTransformatorProps';
import { AadHttpClient, HttpClient } from '@microsoft/sp-http';
import { UrlQueryParameterCollection } from '@microsoft/sp-core-library';
import { IStorageEntity } from './components/IStorageEntity';
import { IStorageEntities } from './components/IStorageEntities';


export interface IPageTransformatorWebPartProps {
  description: string;
}

export default class PageTransformatorWebPart extends BaseClientSideWebPart<IPageTransformatorWebPartProps> {
  private modernizationClient: AadHttpClient = null;
  private siteUrl: string = null;
  private pageUrl: string = null;
  private modernizationAPI: string = '';
  private editMode: boolean = false;
  private storageEntitiesSet: boolean = false;
  private storageEntities: IStorageEntities;

  protected onInit(): Promise<void> {

    // Grab parameters from url
    var queryParms = new UrlQueryParameterCollection(window.location.href);
    var siteUrlQueryParam = queryParms.getValue("SiteUrl").toLowerCase();

    // extend siteUrl value with sub site string if needed
    var subSitePart = document.referrer.toLowerCase().replace(`${PageTransformatorWebPart.getAbsoluteDomainUrl().toLowerCase()}${siteUrlQueryParam}`, "");
    subSitePart = subSitePart.substring(0, subSitePart.toLowerCase().indexOf("/sitepages"));

    if (subSitePart)
    {
      siteUrlQueryParam = siteUrlQueryParam + subSitePart;
    }

    var listIdQueryParam = queryParms.getValue("ListId");
    const itemIdQueryParam: number = +queryParms.getValue("ItemId");
    this.editMode = this.isInEditMode();

    // Initializes the needed configuration data and preps for the call to the Azure AD secured endpoint
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {

      // no point in trying to initialize stuff when in edit mode
      if (this.editMode || listIdQueryParam == null)
      {
        resolve();
        return;
      }

      this.loadStorageEntities().then((storageEntitiesValue) => {

        // Get settings (stored as storage entities)
        // Load storage entities, only needed if we've found the page to transform
        // Set-PnPStorageEntity -Key "Modernization_AzureADApp" -Value "277a116f-db6f-45f2-9f7a-f4f866424442" -Description "ID of the Azure AD app is used for page transformation"
        // Set-PnPStorageEntity -Key "Modernization_FunctionHost" -Value "https://sharepointpnpmodernization.azurewebsites.net" -Description "Host of the SharePoint PnP Modernization service"
        // Set-PnPStorageEntity -Key "Modernization_PageTransformationEndpoint" -Value "api/ModernizePage -Description "Api endpoint for page transformation"
        
        // store storage entities for later use
        this.storageEntities = storageEntitiesValue;    
        
        // trigger bad config for testing
        // this.storageEntities.azureADApp = "";

        // Check is all storage entities are set
        if (this.storageEntities.azureADApp && this.storageEntities.functionHost && this.storageEntities.pageTransformationEndpoint)
        {
          this.storageEntitiesSet = true;  
        }

        // no point in trying to continue when the storage entities are not set
        if (!this.storageEntitiesSet)
        {
          resolve();
          return;
        }

        // Instantiate AADHttpClient
        this.context.aadHttpClientFactory
          .getClient(this.storageEntities.azureADApp)
          .then((client: AadHttpClient): void => {
            this.modernizationClient = client;

            // Load page to transform
            this.loadPageToModernize(siteUrlQueryParam, listIdQueryParam, itemIdQueryParam)
            .then((value: any):void => {
              
              // this is already a modern page, no need to transform it
              if (value.ClientSideApplicationId != null)
              {
                  this.pageUrl = null;
                  this.siteUrl = null;
              }
              else
              {
                var host: string = PageTransformatorWebPart.getAbsoluteDomainUrl();
                this.siteUrl = `${host}${siteUrlQueryParam}`;
                this.pageUrl = `${host}${value.FileRef}`; 
              }
            }).then(() => {
              
              if (this.pageUrl != null)
              {

                    // Append trailing slash if needed
                    let hostUrl: string = this.storageEntities.functionHost;                
                    if (hostUrl.substr(-1) != '/') {
                      hostUrl += '/';
                    }

                    this.modernizationAPI = `${hostUrl}${this.storageEntities.pageTransformationEndpoint}`; 
                    console.log(`Modernization API: ${this.modernizationAPI}`);
                    resolve();       
              }
              else
              {
                resolve();
              }  
            });
            
          }, err => reject(err)); 


      });
                
    });
  }

  public render(): void {
    const element: React.ReactElement<IPageTransformatorProps > = React.createElement(
      PageTransformator,
      {
        modernizationClient: this.modernizationClient,
        modernizationApi: this.modernizationAPI,
        siteUrl: this.siteUrl,
        pageUrl: this.pageUrl,
        editMode: this.editMode,
        storageEntitiesSet: this.storageEntitiesSet,
        storageEntities: this.storageEntities
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  //#region Helper methods
  private async loadPageToModernize(siteUrl: string, listId: string, itemId: number)
  {
    const restApiUrl = `${siteUrl}/_api/web/lists('${listId}')/items(${itemId})?$select=FileRef,ClientSideApplicationId`;
    return await this.context.httpClient.get(restApiUrl, HttpClient.configurations.v1,
      {
          headers: {
              'Accept': 'application/json;odata=nometadata',
              'odata-version': ''
          }
      })
      .then(data => data.json())
      .then((data) => {
        if (data) {
          return data;
        }                                                                     
        return null;
      });
  }

  private async getStorageEntity(key: string)
  {
    const restApiUrl = `${this.context.pageContext.site.absoluteUrl}/_api/web/GetStorageEntity('${key}')`;
    return await this.context.httpClient.get(restApiUrl, HttpClient.configurations.v1,
      {
          headers: {
              'Accept': 'application/json;odata=nometadata',
              'odata-version': ''
          }
      })
      .then(data => data.json())
      .then((data: IStorageEntity) => {
        if (data && data.Value) {
          return data.Value;
        }

        return null;
      });
  }

  private async loadStorageEntities(): Promise<IStorageEntities>
  {
    const result: IStorageEntities = { 
      azureADApp: await this.getStorageEntity("Modernization_AzureADApp"), 
      functionHost: await this.getStorageEntity("Modernization_FunctionHost"), 
      pageTransformationEndpoint: await this.getStorageEntity("Modernization_PageTransformationEndpoint")
    };
    // Wait for all results to be loaded
    Promise.all([result.azureADApp, result.functionHost, result.pageTransformationEndpoint]);
    return result;
  }

  private isInEditMode(): boolean
  {
    if (Environment.type == EnvironmentType.ClassicSharePoint)
    {
      // For now return false, not plannin to host this anyhow on a classic page
      return false;
    }
    else if(Environment.type == EnvironmentType.SharePoint)
    {
      if (this.displayMode == DisplayMode.Edit)
      {
        // Modern SharePoint in Edit Mode
        return true;
      } 
      else if (this.displayMode == DisplayMode.Read)
      {
        // Modern SharePoint in Read Mode
        return false;
      }
    }
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
  //#endregion
}
