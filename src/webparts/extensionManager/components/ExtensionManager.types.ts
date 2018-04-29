/**
 * IExtensionManager.types
 */
import { IWebPartContext } from "@microsoft/sp-webpart-base";
import { IUserCustomAction } from "../services";
import { IColumn, IContextualMenuProps, IGroup } from "office-ui-fabric-react";

export interface IExtensionManagerProps {
  webPartContext: IWebPartContext;
}

export interface IExtensionManagerState {
  dataLoaded: boolean;
  selection?: IUserCustomAction[];
  sortedItems: IUserCustomAction[];
  columns: IColumn[];
  loading: boolean;
  contextualMenuProps?: IContextualMenuProps;
  selectionCount: number;
  showPane: boolean;
  hideDeleteDialog: boolean;
}