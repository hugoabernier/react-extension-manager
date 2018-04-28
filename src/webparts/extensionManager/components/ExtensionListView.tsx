/**
 * ExtensionListView
 */
import { cloneDeep, findIndex, has, isEqual, sortBy } from "@microsoft/sp-lodash-subset";
import { IUserCustomAction } from "../services";
import { ExtensionCommandBar } from "./ExtensionCommandBar";
import { ExtensionPanel } from "./ExtensionPanel";
import {
    ContextualMenu,
    DirectionalHint,
    IContextualMenuItem,
    IContextualMenuProps,
    buildColumns,
    ColumnActionsMode,
    DetailsList,
    DetailsListLayoutMode,
    IColumn,
    IGroup,
    IObjectWithKey,
    MarqueeSelection,
    Selection,
    SelectionMode
} from "office-ui-fabric-react";
import * as strings from "ExtensionManagerWebPartStrings";
import * as React from "react";
import { IExtensionListViewProps, IExtensionListViewState } from "./ExtensionListView.types";

let items: any[];

export class ExtensionListView extends React.Component<IExtensionListViewProps, IExtensionListViewState> {
    private selection: Selection;
    private columns: IColumn[] = [
        {
            key: "titleColumn",
            name: strings.TitleHeader,
            fieldName: "Title",
            minWidth: 100,
            maxWidth: 300,
            isResizable: true,
            columnActionsMode: ColumnActionsMode.hasDropdown
        },
        {
            key: "scopeColumn",
            name: strings.ScopeHeader,
            fieldName: "Scope",
            minWidth: 30,
            maxWidth: 50,
            isResizable: true,
            onRender: this.renderScopeColumn,
            columnActionsMode: ColumnActionsMode.hasDropdown
        },
        {
            key: "rtColumn",
            name: strings.RegistrationTypeHeader,
            fieldName: "RegistrationType",
            minWidth: 50,
            maxWidth: 100,
            isResizable: true,
            onRender: this.renderRegistrationTypeColumn,
            columnActionsMode: ColumnActionsMode.hasDropdown
        },
        {
            key: "locationColumn",
            name: strings.LocationHeader,
            fieldName: "Location",
            minWidth: 10,
            maxWidth: 200,
            isResizable: true,
            onRender: this.renderLocationColumn,
            columnActionsMode: ColumnActionsMode.hasDropdown
        }
    ];
    constructor(props: IExtensionListViewProps) {
        super(props);

        this.selection = new Selection({
            onSelectionChanged: () => this.getSelectionDetails()
        });

        items = this.props.items;

        // initialize state
        this.state = {
            sortedItems: items,
            columns: this.columns,
            loading: true,
            contextualMenuProps: undefined,
            selectionCount: this.selection.getSelectedCount(),
            showPane: false
        };

        this.onColumnClick = this.onColumnClick.bind(this);
        this.onItemInvoked = this.onItemInvoked.bind(this);
        this.onColumnHeaderContextMenu = this.onColumnHeaderContextMenu.bind(this);
    }

    public render(): React.ReactElement<IExtensionListViewProps> {
        const {
            sortedItems,
            columns,
            loading,
            contextualMenuProps
        } = this.state;

        return (
            <div>
                <ExtensionCommandBar
                    selectionCount={this.state.selectionCount}
                    onToggleInfoPane={this.onToggleInfoPane}
                />
                <MarqueeSelection selection={this.selection}>
                    <DetailsList
                        items={sortedItems}
                        columns={this.columns}
                        setKey="key"
                        layoutMode={DetailsListLayoutMode.fixedColumns}
                        selection={this.selection}
                        selectionPreservedOnEmptyClick={false}
                        onColumnHeaderClick={this.onColumnClick}
                        onItemInvoked={this.onItemInvoked}
                        onColumnHeaderContextMenu={this.onColumnHeaderContextMenu}
                    />
                </MarqueeSelection>
                {contextualMenuProps && (
                    <ContextualMenu {...contextualMenuProps} />
                )}
                <ExtensionPanel
                isOpen={this.state.showPane}
                onDismiss={this.onDismissPane}
                />
            </div>
        );
    }

    public componentDidUpdate(prevProps: IExtensionListViewProps, prevState: IExtensionListViewState): void {
        // select default items
        this.setSelectedItems();

        if (!isEqual(prevProps, this.props)) {
            // this._selection.setItems(this.props.items, true);
            items = this.props.items;
        }
    }

    private getSelectionDetails(): void {
        const selected: IUserCustomAction[] = this.selection.getSelection() as IUserCustomAction[];
        this.setState(
            {
                selectionCount: this.selection.getSelectedCount()
            });
      }

    private setSelectedItems(): void {
        if (this.props.items &&
            this.props.items.length > 0 &&
            this.props.defaultSelection &&
            this.props.defaultSelection.length > 0) {
            this.props.defaultSelection.forEach((element: any, index: number) => {
                if (index > -1) {
                    this.selection.setIndexSelected(index, true, false);
                }
            });

        }
    }

    private renderScopeColumn(item: any, index: number, column: IColumn): JSX.Element {
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

    private renderRegistrationTypeColumn(item: any, index: number, column: IColumn): JSX.Element {
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

    private renderLocationColumn(item: any, index: number, column: IColumn): JSX.Element {
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

    private getRegistrationTypeLabel(registrationTypeValue: number): string {
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

    // private _sortItems(items: any[], columnName: string, descending: boolean = false): any[] {
    //     console.log("_sortItems", columnName, descending);
    //     // sort the items
    //     const ascItems: any[] = sortBy(items, [columnName]);
    //     const sortedItems: any[] = descending ? ascItems.reverse() : ascItems;

    //     // check if selection needs to be updated
    //     const selection: IObjectWithKey[] = this._selection.getSelection();
    //     if (selection && selection.length > 0) {
    //         // clear selection
    //         this._selection.setItems([], true);
    //         setTimeout(() => {
    //             // find new index
    //             const idxs: number[] = selection.map((item: IObjectWithKey) => findIndex(sortedItems, item));
    //             idxs.forEach((idx: number) => this._selection.setIndexSelected(idx, true, false));
    //         }, 0);
    //     }

    //     // return the sorted items list
    //     return sortedItems;
    // }

    private onColumnHeaderContextMenu(column: IColumn | undefined, ev: React.MouseEvent<HTMLElement> | undefined): void {
        if (column.columnActionsMode !== ColumnActionsMode.disabled) {
            this.setState({
                contextualMenuProps: this.getContextualMenuProps(ev, column)
            });
        }
    }

    private onItemInvoked(item: any, index: number | undefined): void {
        alert(`Item ${item.name} at index ${index} has been invoked.`);
    }

    private onColumnClick(ev: React.MouseEvent<HTMLElement>, column: IColumn): void {
        if (column.columnActionsMode !== ColumnActionsMode.disabled) {
            this.setState({
                contextualMenuProps: this.getContextualMenuProps(ev, column)
            });
        }
    }

    private sortItems(column: IColumn, isSortedDescending: boolean): void {
        let { sortedItems } = this.state;
        const { columns } = this.state;

        // if we've sorted this column, flip it.
        if (column.isSorted) {
            isSortedDescending = !isSortedDescending;
        }

        // sort the items.
        sortedItems = sortedItems!.concat([]).sort((a: IUserCustomAction, b: IUserCustomAction) => {
            const firstValue: any = a[column.fieldName];
            const secondValue: any = b[column.fieldName];
            if (isSortedDescending) {
                return firstValue > secondValue ? -1 : 1;
            } else {
                return firstValue > secondValue ? 1 : -1;
            }
        });

        // reset the items and columns to match the state.
        this.setState({
            sortedItems: sortedItems,
            columns: columns!.map((col: IColumn) => {
                col.isSorted = (col.key === column.key);
                if (col.isSorted) {
                    col.isSortedDescending = isSortedDescending;
                }
                return col;
            })
        });
    }

    private getContextualMenuProps(ev: React.MouseEvent<HTMLElement>, column: IColumn): IContextualMenuProps {
        const menuItems: IContextualMenuItem[] = [
            {
                key: "aToZ",
                name: "A to Z",
                canCheck: true,
                checked: column.isSorted && !column.isSortedDescending,
                onClick: () => this.sortItems(column, false)
            },
            {
                key: "zToA",
                name: "Z to A",
                canCheck: true,
                checked: column.isSorted && column.isSortedDescending,
                onClick: () => this.sortItems(column, true)
            }
        ];
        // if (isGroupable(column.key)) {
        //     menuItems.push({
        //         key: "groupBy",
        //         name: "Group By " + column.name,
        //         icon: "GroupedDescending",
        //         canCheck: true,
        //         checked: column.isGrouped,
        //         onClick: () => this._onGroupByColumn(column)
        //     });
        // }
        return {
            items: menuItems,
            target: ev.currentTarget as HTMLElement,
            directionalHint: DirectionalHint.bottomLeftEdge,
            gapSpace: 10,
            isBeakVisible: true,
            onDismiss: this.onContextualMenuDismissed
        };
    }

    private onContextualMenuDismissed = (): void => {
        this.setState({
            contextualMenuProps: undefined
        });
    }

    private onDismissPane = (): void => {
        this.setState({
            showPane: false
        });
    }

    private onToggleInfoPane = (): void => {
        this.setState({
            showPane: !this.state.showPane
        });
    }
}