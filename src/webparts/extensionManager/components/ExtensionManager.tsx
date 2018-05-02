/**
 * Renders a list containing all extensions registered against a site
 */
import {
  Environment,
  EnvironmentType
} from "@microsoft/sp-core-library";
import * as strings from "ExtensionManagerWebPartStrings";
import * as React from "react";
import {
  ExtensionService,
  IExtensionService,
  IUserCustomAction,
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
  DirectionalHint,
  IContextualMenuItem,
  IContextualMenuProps,
  ColumnActionsMode,
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  Link,
  MarqueeSelection,
  Selection,
  Spinner,
  DialogType,
  Dialog,
  DialogFooter,
  PrimaryButton,
  DefaultButton
} from "office-ui-fabric-react";
import KeyHandler from "react-key-handler";

export class ExtensionManager extends React.Component<IExtensionManagerProps, IExtensionManagerState> {

  // private _extensionItems: IUserCustomAction[] = [];
  private _selection: Selection;
  private _columns: IColumn[] = [
    {
      key: "Title",
      name: strings.TitleHeader,
      fieldName: "Title",
      minWidth: 100,
      maxWidth: 300,
      isResizable: true,
      columnActionsMode: ColumnActionsMode.hasDropdown,
      onRender: this._renderTitleColumn
    },
    {
      key: "Scope",
      name: strings.ScopeHeader,
      fieldName: "Scope",
      minWidth: 30,
      maxWidth: 70,
      isResizable: true,
      onRender: this._renderScopeColumn,
      columnActionsMode: ColumnActionsMode.hasDropdown
    },
    {
      key: "RegistrationType",
      name: strings.RegistrationTypeHeader,
      fieldName: "RegistrationType",
      minWidth: 50,
      maxWidth: 120,
      isResizable: true,
      onRender: this._renderRegistrationTypeColumn,
      columnActionsMode: ColumnActionsMode.hasDropdown
    },
    {
      key: "Location",
      name: strings.LocationHeader,
      fieldName: "Location",
      minWidth: 10,
      maxWidth: 200,
      isResizable: true,
      onRender: this._renderLocationColumn,
      columnActionsMode: ColumnActionsMode.hasDropdown
    }
  ];

  private _newItems: IContextualMenuItem[] = [
    {
      key: "newItem",
      name: strings.NewButton,
      icon: "Add",
      ariaLabel: strings.NewButtonAriaLabel,
      ["data-automation-id"]: "newItemMenu"
    },
    {
      key: "upload",
      name: strings.UploadButton,
      icon: "Upload",
      ariaLabel: strings.UploadButtonAriaLabel,
      ["data-automation-id"]: "uploadButton"
    }
  ];

  private _editItems: IContextualMenuItem[] = [
    {
      key: "edit",
      name: strings.EditButton,
      icon: "Edit",
      ariaLabel: strings.EditButtonLabel,
      ["data-automation-id"]: "editButton"
    }
  ];

  private _deleteItems: IContextualMenuItem[] = [
    {
      key: "delete",
      name: strings.DeleteButton,
      icon: "Delete",
      onClick: () => { this._showDeleteConfirmation(); },
      ["data-automation-id"]: "deleteButton"
    }
  ];

  private _farItems: any = [
    {
      key: "info",
      name: strings.InfoButton,
      icon: "Info",
      title: strings.InfoButton,
      iconOnly: true,
      onClick: () => { this._onShowPane(); }
    }
  ];

  constructor(props: IExtensionManagerProps) {
    super(props);

    this._selection = new Selection({
      onSelectionChanged: () => this._getSelectionDetails()
    });

    this.state = {
      loading: true,
      sortedItems: [],
      columns: this._columns,
      contextualMenuProps: undefined,
      selectionCount: this._selection.getSelectedCount(),
      showPane: false,
      hideDeleteDialog: true
    };
  }

  public async componentDidMount(): Promise<void> {
    // this.props.webPartContext.statusRenderer.displayLoadingIndicator(
    //   document.getElementsByClassName(styles.extensionManager)[0], strings.LoadingLabel);

    const extensionItems: IUserCustomAction[] = await this._getExtensionItems();
    this.setState({
      loading: false,
      sortedItems: extensionItems
    });
    // this.props.webPartContext.statusRenderer.clearLoadingIndicator(
    //   document.getElementsByClassName(styles.extensionManager)[0]);
  }

  public render(): React.ReactElement<IExtensionManagerProps> {
    const {
      sortedItems,
      loading,
      contextualMenuProps
    } = this.state;

    const loadingSpinner: JSX.Element =
      loading ? <div className={styles.spinner}><Spinner label={"Loading extensions..."} /></div> : <div />;

    // const error: JSX.Element = this.state.error ? <div><strong>Error: </strong> {this.state.error}</div> : <div/>;
    const commandBar: JSX.Element = this.renderCommandBar();
    const detailsList: JSX.Element =
      <MarqueeSelection selection={this._selection}>
        <DetailsList
          className={styles.extensionsDetailsList}
          items={sortedItems}
          columns={this._columns}
          setKey="key"
          layoutMode={DetailsListLayoutMode.fixedColumns}
          selection={this._selection}
          selectionPreservedOnEmptyClick={false}
          onColumnHeaderClick={this._onColumnClick}
          onItemInvoked={this._onItemInvoked}
          onColumnHeaderContextMenu={this._onColumnHeaderContextMenu}
        />
      </MarqueeSelection>;

    const panel: JSX.Element = this.state.showPane &&
      <ExtensionPanel
        isOpen={this.state.showPane}
        onDismiss={this._onDismissPane}
      />;

    const deleteDialog: JSX.Element = this.renderDeleteDialog();
    return (
      <div className={styles.extensionManager}>
        {/*
            I wish I could have used Office Fabric's FocusTrapZone, but I couldn't get it to work
            */}
        <KeyHandler keyEventName="keydown" keyValue="Delete" onKeyHandle={this._showDeleteConfirmation} />
        <KeyHandler keyEventName="keydown" keyValue="Escape" onKeyHandle={this._onClearSelection} />

        {commandBar}

        {this.state.loading ? loadingSpinner : detailsList}
        {contextualMenuProps && (
          <ContextualMenu {...contextualMenuProps} />
        )}
        {panel}
        {deleteDialog}

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
      menuItems = this._newItems;
    } else if (selectionCount === 1) {
      // 1 item selected, allow editing and deleting
      menuItems = menuItems.concat(this._editItems, this._deleteItems);
    } else {
      // more than 1 item, only allow deleting
      menuItems = this._deleteItems;
    }

    const farItems: IContextualMenuItem[] = this._farItems;

    if (selectionCount > 0) {
      farItems.push(
        {
          key: "clearSelection",
          name: strings.ClearSelectionButton.replace("{0}", `${selectionCount}`),
          icon: "Clear",
          title: strings.ClearSelectionButtonTitle,
          iconOnly: false,
          ariaLabel: strings.ClearSelectionButtonAriaLabel.replace("{0}", `${selectionCount}`),
          onClick: () => this._onClearSelection(),
          className: styles.isFlipped
        });
    }

    return (
      <CommandBar
        isSearchBoxVisible={false}
        items={menuItems}
        farItems={farItems}
      />
    );
  }

  public renderDeleteDialog(): JSX.Element {
    return (
      <Dialog
        hidden={this.state.hideDeleteDialog}
        onDismiss={this._closeDeleteDialog}
        dialogContentProps={{
          type: DialogType.normal,
          title: strings.DeleteDialogTitle,
          subText: strings.DeleteDialogDescription
        }}
        modalProps={{
          titleAriaId: "myLabelId",
          subtitleAriaId: "mySubTextId",
          isBlocking: false,
          containerClassName: "ms-dialogMainOverride"
        }}
      >
        <DialogFooter>
          <PrimaryButton onClick={this._closeDeleteDialog} text="Remove" />
          <DefaultButton onClick={this._closeDeleteDialog} text="Cancel" />
        </DialogFooter>
      </Dialog>

    );

  }

  private _showDeleteConfirmation = (): void => {
    this.setState({
      hideDeleteDialog: false
    });
  }
  private _closeDeleteDialog = (): void => {
    this.setState({
      hideDeleteDialog: true
    });
  }

  private _getSelectionDetails(): void {
    this.setState(
      {
        selectionCount: this._selection.getSelectedCount()
      });
  }

  private _renderTitleColumn(item: any, index: number, column: IColumn): JSX.Element {
    const fieldContent: any = item[column.fieldName];
    return <Link>{ fieldContent }</Link>;
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

  private _sortByColumn(column: IColumn, isSortedDescending: boolean): void {
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

  private _getContextualMenuProps(ev: React.MouseEvent<HTMLElement>, column: IColumn): IContextualMenuProps {
    const menuItems: IContextualMenuItem[] = [
      {
        key: "aToZ",
        name: strings.AscendingSort,
        canCheck: true,
        checked: column.isSorted && !column.isSortedDescending,
        onClick: () => this._sortByColumn(column, false)
      },
      {
        key: "zToA",
        name: strings.DescendingSort,
        canCheck: true,
        checked: column.isSorted && column.isSortedDescending,
        onClick: () => this._sortByColumn(column, true)
      }
    ];

    return {
      items: menuItems,
      target: ev.currentTarget as HTMLElement,
      directionalHint: DirectionalHint.bottomLeftEdge,
      gapSpace: 10,
      isBeakVisible: true,
      onDismiss: this._onContextualMenuDismissed
    };
  }

  private async _getExtensionItems(): Promise<IUserCustomAction[]> {
    const dataService: IExtensionService = (Environment.type === EnvironmentType.Test || Environment.type === EnvironmentType.Local) ?
      new MockExtensionService() :
      new ExtensionService(this.props.webPartContext);
    return dataService.getExtensions();
  }

  private _onColumnHeaderContextMenu = (column: IColumn | undefined, ev: React.MouseEvent<HTMLElement> | undefined): void => {
    if (column.columnActionsMode !== ColumnActionsMode.disabled) {
      this.setState({
        contextualMenuProps: this._getContextualMenuProps(ev, column)
      });
    }
  }

  private _onItemInvoked = (item: any, index: number | undefined): void => {
    alert(`Item ${item.name} at index ${index} has been invoked.`);
  }

  private _onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
    if (column.columnActionsMode !== ColumnActionsMode.disabled) {
      this.setState({
        contextualMenuProps: this._getContextualMenuProps(ev, column)
      });
    }
  }

  private _onClearSelection = (): void => {
    this._selection.setAllSelected(false);
    this.setState({
      selection: []
    });
  }

  private _onContextualMenuDismissed = (): void => {
    this.setState({
      contextualMenuProps: undefined
    });
  }

  private _onDismissPane = (): void => {
    this.setState({
      showPane: false
    });
  }

  private _onShowPane = (): void => {
    this.setState({
      showPane: true
    });
  }
}