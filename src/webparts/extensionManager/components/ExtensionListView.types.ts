import { IUserCustomAction } from "../services";

export interface IExtensionListViewProps {
    items: IUserCustomAction[];
    defaultSelection: IUserCustomAction[];
}

export interface IExtensionListViewState {
    items: IUserCustomAction[];
}