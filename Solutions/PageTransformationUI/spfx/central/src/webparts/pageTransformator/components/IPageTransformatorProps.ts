import { AadHttpClient } from '@microsoft/sp-http';
import { IStorageEntities } from './IStorageEntities';

export interface IPageTransformatorProps {
  modernizationClient: AadHttpClient;
  modernizationApi: string;
  siteUrl: string;
  pageUrl: string;
  editMode: boolean;
  storageEntitiesSet: boolean;
  storageEntities: IStorageEntities;
}
