import { IWebPartContext } from "@microsoft/sp-webpart-base";

export interface IExtensionManagerProps {
  webPartContext: IWebPartContext;
}

export interface IExtensionManagerState {
  dataLoaded: boolean;
}