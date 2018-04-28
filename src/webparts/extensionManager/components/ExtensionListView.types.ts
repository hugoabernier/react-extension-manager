import { IContextualMenuProps } from "office-ui-fabric-react/lib/ContextualMenu";
import { IColumn } from "office-ui-fabric-react/lib/DetailsList";
import { IUserCustomAction } from "../services";

export interface IExtensionListViewProps {
    items: IUserCustomAction[];
    defaultSelection: IUserCustomAction[];
    // onSelectionChanged?: (item?: IUserCustomAction[]) => void;
}

export interface IExtensionListViewState {
    sortedItems: IUserCustomAction[];
    columns:IColumn[];
    loading:boolean;
    contextualMenuProps?: IContextualMenuProps;
    selectionCount: number;
}