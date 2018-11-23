import { WebPartContext } from "@microsoft/sp-webpart-base";
import { MessageBarType } from "office-ui-fabric-react/lib/MessageBar";

export interface IPageTransformatorAdminProps {
  description: string;
  context: WebPartContext;
}

export interface IPageTransformatorAdminState {
  siteUrl: string;
  buttonsDisabled: boolean;
  resultMessage :string;
  resultMessageType: MessageBarType;
}