import { cloneDeep, findIndex, has, isEqual, sortBy } from "@microsoft/sp-lodash-subset";
import * as strings from "ExtensionManagerWebPartStrings";
import {
    DetailsList,
    DetailsListLayoutMode,
    IColumn,
    IGroup,
    IObjectWithKey,
    Selection,
    SelectionMode
} from "office-ui-fabric-react/lib/DetailsList";
import { MarqueeSelection } from "office-ui-fabric-react/lib/MarqueeSelection";
import * as React from "react";
import { IExtensionListViewProps, IExtensionListViewState } from "./ExtensionListView.types";

export class ExtensionListView extends React.Component<IExtensionListViewProps, IExtensionListViewState> {
    private _selection: Selection;
    private _columns: IColumn[] = [
        {
            key: "titleColumn",
            name: strings.TitleHeader,
            fieldName: "Title",
            minWidth: 100,
            maxWidth: 300,
            isResizable: true,
        },
        {
            key: "scopeColumn",
            name: strings.ScopeHeader,
            fieldName: "Scope",
            minWidth: 30,
            maxWidth: 50,
            isResizable: true,
            onRender: this._renderScopeColumn
        },
        {
            key: "rtColumn",
            name: strings.RegistrationTypeHeader,
            fieldName: "RegistrationType",
            minWidth: 50,
            maxWidth: 100,
            isResizable: true,
            onRender: this._renderRegistrationTypeColumn
        },
        {
            key: "locationColumn",
            name: strings.LocationHeader,
            fieldName: "Location",
            minWidth: 10,
            maxWidth: 200,
            isResizable: true,
            onRender: this._renderLocationColumn
        },
    ];
    constructor(props: IExtensionListViewProps) {
        super(props);

        this._selection = new Selection({
            // onSelectionChanged: () => this.setState({ selectionDetails: this._getSelectionDetails() })
        });

        // initialize state
        this.state = {
            items: this.props.items,
        };
    }

    public render(): React.ReactElement<IExtensionListViewProps> {
        const { items } = this.state;
        console.log("render", items);

        return (
            <div>
                <MarqueeSelection selection={this._selection}>
                    <DetailsList
                        items={items}
                        columns={this._columns}
                        setKey="key"
                        layoutMode={DetailsListLayoutMode.fixedColumns}
                        selection={this._selection}
                        selectionPreservedOnEmptyClick={false}
                    />
                </MarqueeSelection>
            </div>
        );
    }

    public componentDidUpdate(prevProps: IExtensionListViewProps, prevState: IExtensionListViewState): void {
        // select default items
        this._setSelectedItems();

        if (!isEqual(prevProps, this.props)) {
            // this._selection.setItems(this.props.items, true);
            this.setState({
                items: this.props.items
            });
        }
    }

    private _setSelectedItems(): void {
        if (this.props.items &&
            this.props.items.length > 0 &&
            this.props.defaultSelection &&
            this.props.defaultSelection.length > 0) {
            this.props.defaultSelection.forEach((element: any, index: number) => {
                if (index > -1) {
                    this._selection.setIndexSelected(index, true, false);
                }
            });

        }
    }
    private _renderScopeColumn(item: any, index: number, column: IColumn): JSX.Element {
        const fieldContent: any = item[column.fieldName];
        let scopeLabel: string;
        switch (fieldContent) {
            case 0:
                scopeLabel = strings.UnknowScopeLabel;
                break;
            case 2:
                scopeLabel = strings.SiteScopeLabel;
                break;
            case 3:
                scopeLabel = strings.WebScopeLabel;
                break;
            case 4:
                scopeLabel = strings.ListScopeLabel;
                break;
            default:
                scopeLabel = strings.NAScopeLabel;
        }

        return <span>{scopeLabel}</span>;
    }

    private _renderRegistrationTypeColumn(item: any, index: number, column: IColumn): JSX.Element {
        const fieldContent: any = item[column.fieldName];

        let registrationTypeLabel: string;
        switch (fieldContent) {
            case 0:
                registrationTypeLabel = strings.NoneRegistrationTypeLabel;
                break;
            case 1:
                registrationTypeLabel = strings.ListRegistrationTypeLabel;
                break;
            case 2:
                registrationTypeLabel = strings.ContentTypeRegistrationTypeLabel;
                break;
            case 3:
                registrationTypeLabel = strings.ProgIdRegistrationTypeLabel;
                break;
            case 4:
                registrationTypeLabel = strings.FileTypeRegistrationTypeLabel;
                break;
            default:
                registrationTypeLabel = strings.NAScopeLabel;
        }

        return <span>{registrationTypeLabel}</span>;
    }

    private _renderLocationColumn(item: any, index: number, column: IColumn): JSX.Element {
        const fieldContent: any = item[column.fieldName];

        let locationLabel: string;
        switch (fieldContent) {
            case "ClientSideExtension.ApplicationCustomizer":
                locationLabel = strings.ApplicationCustomizerLocation;
                break;
            case "ClientSideExtension.ListViewCommandSet.CommandBar":
                locationLabel = strings.CommandBarLocation;
                break;
            case "ClientSideExtension.ListViewCommandSet.ContextMenu":
                locationLabel = strings.ContextMenuLocation;
                break;
            case "ClientSideExtension.ListViewCommandSet":
                locationLabel = strings.ListViewLocation;
                break;
            case "EditControlBlock":
                locationLabel = strings.ECBLocation;
                break;
            default:
                locationLabel = fieldContent;
        }

        return <span>{locationLabel}</span>;
    }

    private _getRegistrationTypeLabel(registrationTypeValue: number): string {
        // translate values according to https://msdn.microsoft.com/en-us/library/office/dn531432.aspx#bk_UserCustomActionProperties
        switch (registrationTypeValue) {
            case 0:
                return strings.NoneRegistrationTypeLabel;
            case 1:
                return strings.ListRegistrationTypeLabel;
            case 2:
                return strings.ContentTypeRegistrationTypeLabel;
            case 3:
                return strings.ProgIdRegistrationTypeLabel;
            case 4:
                return strings.FileTypeRegistrationTypeLabel;
            default:
                return strings.NAScopeLabel;
        }
    }

    private _sortItems(items: any[], columnName: string, descending: boolean = false): any[] {
        // sort the items
        const ascItems: any[] = sortBy(items, [columnName]);
        const sortedItems: any[] = descending ? ascItems.reverse() : ascItems;

        // check if selection needs to be updated
        const selection: IObjectWithKey[] = this._selection.getSelection();
        if (selection && selection.length > 0) {
            // clear selection
            this._selection.setItems([], true);
            setTimeout(() => {
                // find new index
                const idxs: number[] = selection.map((item: IObjectWithKey) => findIndex(sortedItems, item));
                idxs.forEach((idx: number) => this._selection.setIndexSelected(idx, true, false));
            }, 0);
        }

        // return the sorted items list
        return sortedItems;
    }
}