import { CommandBar } from "office-ui-fabric-react/lib/CommandBar";
import * as strings from "ExtensionManagerWebPartStrings";
import * as React from "react";

export interface IExtensionCommandBarProps {
    selectionCount:number;
}

export interface IExtensionCommandBarState { }

const newItems: any[] = [
    {
        key: "newItem",
        name: strings.NewButton,
        icon: "Add",
        ["data-automation-id"]: "newItemMenu",
    },
    {
        key: "upload",
        name: strings.UploadButton,
        icon: "Upload",
        ["data-automation-id"]: "uploadButton"
    },
];

const editItems: any[] = [
    {
        key: "edit",
        name: strings.EditButton,
        icon: "Edit",
        ["data-automation-id"]: "editButton"
    },
];

const deleteItems: any[] = [
    {
        key: "delete",
        name: strings.DeleteButton,
        icon: "Delete",
        ["data-automation-id"]: "deleteButton"
    },
];

export const farItems: any = [
    {
        key: "info",
        name: strings.InfoButton,
        icon: "Info",
        title: strings.InfoButton,
        iconOnly: true,
        onClick: () => { return; }
    }
];

export default class ExtensionCommandBar extends React.Component<IExtensionCommandBarProps, IExtensionCommandBarState> {
    public render(): React.ReactElement<IExtensionCommandBarProps> {
        var items:any[] = [];
        const count:number = this.props.selectionCount;
        if (count === 0) {
            items = newItems;
        } else if (count === 1) {
            items = items.concat(editItems, deleteItems);
        } else {
            items = deleteItems;
        }

        return (
            <CommandBar
                isSearchBoxVisible={false}
                items={items}
                farItems={farItems}
            />
        );
    }
}
