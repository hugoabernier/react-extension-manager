/**
 * Renders a list containing all extensions registered against a site
 */
import {
  Environment,
  EnvironmentType
} from "@microsoft/sp-core-library";
import {
  SPHttpClient,
  SPHttpClientResponse
} from "@microsoft/sp-http";
import { isEqual } from "@microsoft/sp-lodash-subset";
import * as strings from "ExtensionManagerWebPartStrings";
import * as React from "react";
import {
  ExtensionService,
  IExtensionService,
  IUserCustomAction,
  IUserCustomActionCollection,
  MockExtensionService
} from "../services";
import styles from "./ExtensionManager.module.scss";
import {
  IExtensionManagerProps,
  IExtensionManagerState
} from "./ExtensionManager.types";
import { ExtensionPanel } from "./ExtensionPanel";
import {
  CommandBar,
  ContextualMenu,
  ContextualMenuItemType,
  DirectionalHint,
  IContextualMenuItem,
  IContextualMenuProps,
  buildColumns,
  ColumnActionsMode,
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  IObjectWithKey,
  MarqueeSelection,
  Selection,
  SelectionMode,
  Spinner,
  DialogType,
  Dialog,
  DialogFooter,
  PrimaryButton,
  DefaultButton
} from "office-ui-fabric-react";
import KeyHandler from "react-key-handler";

export class ExtensionManager extends React.Component<IExtensionManagerProps, IExtensionManagerState> {

  private maxResults: number = 1000;
  // private extensionItems: IUserCustomAction[] = [];
  private selectedExtensionItems: IUserCustomAction[] = [];

  private selection: Selection;
  private columns: IColumn[] = [
    {
      key: "Title",
      name: strings.TitleHeader,
      fieldName: "Title",
      minWidth: 100,
      maxWidth: 300,
      isResizable: true,
      columnActionsMode: ColumnActionsMode.hasDropdown
    },
    {
      key: "Scope",
      name: strings.ScopeHeader,
      fieldName: "Scope",
      minWidth: 30,
      maxWidth: 70,
      isResizable: true,
      onRender: this.renderScopeColumn,
      columnActionsMode: ColumnActionsMode.hasDropdown
    },
    {
      key: "RegistrationType",
      name: strings.RegistrationTypeHeader,
      fieldName: "RegistrationType",
      minWidth: 50,
      maxWidth: 120,
      isResizable: true,
      onRender: this.renderRegistrationTypeColumn,
      columnActionsMode: ColumnActionsMode.hasDropdown
    },
    {
      key: "Location",
      name: strings.LocationHeader,
      fieldName: "Location",
      minWidth: 10,
      maxWidth: 200,
      isResizable: true,
      onRender: this.renderLocationColumn,
      columnActionsMode: ColumnActionsMode.hasDropdown
    }
  ];

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
      onClick: () => { this.showDeleteConfirmation(); },
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
      onClick: () => { this.onToggleInfoPane(); }
    }
  ];

  constructor(props: IExtensionManagerProps) {
    super(props);

    this.selection = new Selection({
      onSelectionChanged: () => this.getSelectionDetails()
    });

    this.state = {
      dataLoaded: false,
      sortedItems: [],
      columns: this.columns,
      loading: true,
      contextualMenuProps: undefined,
      selectionCount: this.selection.getSelectedCount(),
      showPane: false,
      hideDeleteDialog: true
    };
  }

  public async componentDidMount(): Promise<void> {
    this.props.webPartContext.statusRenderer.displayLoadingIndicator(
      document.getElementsByClassName(styles.extensionManager)[0], strings.LoadingLabel);

    const extensionItems: IUserCustomAction[] = await this.getExtensionItems();
    this.setState({
      dataLoaded: true,
      sortedItems: extensionItems
    });
    this.props.webPartContext.statusRenderer.clearLoadingIndicator(
      document.getElementsByClassName(styles.extensionManager)[0]);
  }

  public render(): React.ReactElement<IExtensionManagerProps> {
    const {
      sortedItems,
      columns,
      loading,
      contextualMenuProps
    } = this.state;

    const loadingSpinner: JSX.Element =
      this.state.loading ? <div style={{ margin: "0 auto" }}><Spinner label={"Loading extensions..."} /></div> : <div />;

    // const error: JSX.Element = this.state.error ? <div><strong>Error: </strong> {this.state.error}</div> : <div/>;
    const commandBar: JSX.Element = this.renderCommandBar();
    const detailsList: JSX.Element =
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
      </MarqueeSelection>;

    const panel: JSX.Element =
      <ExtensionPanel
        isOpen={this.state.showPane}
        onDismiss={this.onDismissPane}
      />;

      const deleteDialog: JSX.Element = this.renderDeleteDialog();
    return (
      <div className={styles.extensionManager}>
        {this.state.dataLoaded &&
          <div>
            {/*
            I wish I could have used Office Fabric's FocusTrapZone, but I couldn't get it to work
            */}
            <KeyHandler keyEventName="keydown" keyValue="Delete" onKeyHandle={this.showDeleteConfirmation} />

            {commandBar}
            {detailsList}
            {contextualMenuProps && (
              <ContextualMenu {...contextualMenuProps} />
            )}
            {panel}
            {deleteDialog}
          </div>
        }
      </div>
    );
  }

  public renderCommandBar(): JSX.Element {
    let menuItems: any[] = [];

    // get the number of items currently selected
    const { selectionCount } = this.state;

    // combine menu items to create a toolbar that changes according to selection
    // to mimic the behaviour found in (modern) SharePoint lists
    if (selectionCount === 0) {
      // no items selected, show the New and Upload options
      menuItems = this.newItems;
    } else if (selectionCount === 1) {
      // 1 item selected, allow editing and deleting
      menuItems = menuItems.concat(this.editItems, this.deleteItems);
    } else {
      // more than 1 item, only allow deleting
      menuItems = this.deleteItems;
    }

    return (
      <CommandBar
        isSearchBoxVisible={false}
        items={menuItems}
        farItems={this.farItems}
      />
    );
  }

  public renderDeleteDialog(): JSX.Element {
    return (
      <Dialog
      hidden={ this.state.hideDeleteDialog }
      onDismiss={this.closeDeleteDialog }
      dialogContentProps={ {
        type: DialogType.normal,
        title: strings.DeleteDialogTitle,
        subText: strings.DeleteDialogDescription
      } }
      modalProps={ {
        titleAriaId: "myLabelId",
        subtitleAriaId: "mySubTextId",
        isBlocking: false,
        containerClassName: "ms-dialogMainOverride"
      } }
    >
      <DialogFooter>
        <PrimaryButton onClick={ this.closeDeleteDialog } text="Remove" />
        <DefaultButton onClick={ this.closeDeleteDialog } text="Cancel" />
      </DialogFooter>
    </Dialog>

    );

  }

  private showDeleteConfirmation = (): void => {
    this.setState({
      hideDeleteDialog: false
    });
  }
  private closeDeleteDialog = (): void => {
    this.setState({
      hideDeleteDialog: true
    });
  }

  private getSelectionDetails(): void {
    const selected: IUserCustomAction[] = this.selection.getSelection() as IUserCustomAction[];
    this.setState(
      {
        selectionCount: this.selection.getSelectedCount()
      });
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

  private sortByColumn(column: IColumn, isSortedDescending: boolean): void {
    let { sortedItems } = this.state;
    const { columns } = this.state;

    // don't do anything if we're already sorting by this column
    if (column.isSorted && column.isSortedDescending === isSortedDescending) {
      return;
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
        name: strings.AscendingSort,
        canCheck: true,
        checked: column.isSorted && !column.isSortedDescending,
        onClick: () => this.sortByColumn(column, false)
      },
      {
        key: "zToA",
        name: strings.DescendingSort,
        canCheck: true,
        checked: column.isSorted && column.isSortedDescending,
        onClick: () => this.sortByColumn(column, true)
      }
    ];

    return {
      items: menuItems,
      target: ev.currentTarget as HTMLElement,
      directionalHint: DirectionalHint.bottomLeftEdge,
      gapSpace: 10,
      isBeakVisible: true,
      onDismiss: this.onContextualMenuDismissed
    };
  }

  private async getExtensionItems(): Promise<IUserCustomAction[]> {
    const dataService: IExtensionService = (Environment.type === EnvironmentType.Test || Environment.type === EnvironmentType.Local) ?
      new MockExtensionService() :
      new ExtensionService(this.props.webPartContext);
    return dataService.getExtensions();
  }

  private onColumnHeaderContextMenu = (column: IColumn | undefined, ev: React.MouseEvent<HTMLElement> | undefined): void => {
    if (column.columnActionsMode !== ColumnActionsMode.disabled) {
      this.setState({
        contextualMenuProps: this.getContextualMenuProps(ev, column)
      });
    }
  }

  private onItemInvoked = (item: any, index: number | undefined): void => {
    alert(`Item ${item.name} at index ${index} has been invoked.`);
  }

  private onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
    if (column.columnActionsMode !== ColumnActionsMode.disabled) {
      this.setState({
        contextualMenuProps: this.getContextualMenuProps(ev, column)
      });
    }
  }

  private onSelectionChanged = (selection: IUserCustomAction[]): void => {
    this.setState({
      selection: selection
    });
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