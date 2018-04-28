/**
 * IExtensionManager.types
 */
import { IWebPartContext } from "@microsoft/sp-webpart-base";
import { IUserCustomAction } from "../services";

export interface IExtensionManagerProps {
  webPartContext: IWebPartContext;
}

export interface IExtensionManagerState {
  dataLoaded: boolean;
  selection?: IUserCustomAction[];
}