/**
 * ExtensionCommandBar
 */
import { CommandBar } from "office-ui-fabric-react";
import * as strings from "ExtensionManagerWebPartStrings";
import * as React from "react";

export interface IExtensionCommandBarProps {
    selectionCount: number;
    onToggleInfoPane: () => void;
}

export interface IExtensionCommandBarState { }

export class ExtensionCommandBar extends React.Component<IExtensionCommandBarProps, IExtensionCommandBarState> {
    private newItems: any[] = [
        {
            key: "newItem",
            name: strings.NewButton,
            icon: "Add",
            ["data-automation-id"]: "newItemMenu"
        },
        {
            key: "upload",
            name: strings.UploadButton,
            icon: "Upload",
            ["data-automation-id"]: "uploadButton"
        }
    ];

    private editItems: any[] = [
        {
            key: "edit",
            name: strings.EditButton,
            icon: "Edit",
            ["data-automation-id"]: "editButton"
        }
    ];

    private deleteItems: any[] = [
        {
            key: "delete",
            name: strings.DeleteButton,
            icon: "Delete",
            ["data-automation-id"]: "deleteButton"
        }
    ];

    private farItems: any = [
        {
            key: "info",
            name: strings.InfoButton,
            icon: "Info",
            title: strings.InfoButton,
            iconOnly: true,
            onClick: () => { this.props.onToggleInfoPane(); }
        }
    ];

    public render(): React.ReactElement<IExtensionCommandBarProps> {
        let items: any[] = [];

        // get the number of items currently selected
        const count: number = this.props.selectionCount;

        // combine menu items to create a toolbar that changes according to selection
        // to mimic the behaviour found in (modern) SharePoint lists
        if (count === 0) {
            // no items selected, show the New and Upload options
            items = this.newItems;
        } else if (count === 1) {
            // 1 item selected, allow editing and deleting
            items = items.concat(this.editItems, this.deleteItems);
        } else {
            // more than 1 item, only allow deleting
            items = this.deleteItems;
        }

        return (
            <CommandBar
                isSearchBoxVisible={false}
                items={items}
                farItems={this.farItems}
            />
        );
    }
}
