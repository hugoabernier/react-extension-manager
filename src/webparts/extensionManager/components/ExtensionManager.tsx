import {
  Environment,
  EnvironmentType
} from "@microsoft/sp-core-library";
import {
  SPHttpClient,
  SPHttpClientResponse
} from "@microsoft/sp-http";
import { escape } from "@microsoft/sp-lodash-subset";
import * as strings from "ExtensionManagerWebPartStrings";
import * as React from "react";
import {
  ExtensionService,
  IExtensionService,
  IUserCustomAction,
  IUserCustomActionCollection,
  MockExtensionService
} from "../services";
import { ExtensionListView } from "./ExtensionListView";
import styles from "./ExtensionManager.module.scss";
import {
  IExtensionManagerProps,
  IExtensionManagerState,
} from "./IExtensionManager.types";

export default class ExtensionManager extends React.Component<IExtensionManagerProps, IExtensionManagerState> {
  private _maxResults = 1000;
  private _extensionItems: IUserCustomAction[] = [];
  private _selectedExtensionItems: IUserCustomAction[] = [];

  constructor(props: IExtensionManagerProps) {
    super(props);

    this.state = {
      dataLoaded: false,
    };

    this._onSelectionChanged = this._onSelectionChanged.bind(this);
  }

  public async componentDidMount(): Promise<void> {
    this.props.webPartContext.statusRenderer.displayLoadingIndicator(
      document.getElementsByClassName(styles.extensionManager)[0], strings.LoadingLabel);

    this._extensionItems = await this._getExtensionItems();
    this.setState({
      dataLoaded: true
    });
    this.props.webPartContext.statusRenderer.clearLoadingIndicator(
      document.getElementsByClassName(styles.extensionManager)[0]);
  }

  public render(): React.ReactElement<IExtensionManagerProps> {
    console.log("items", this._extensionItems);

    return (
      <div className={styles.extensionManager}>
        {this.state.dataLoaded &&
          <ExtensionListView
            items={this._extensionItems}
            defaultSelection={[]}
            // onSelectionChanged={this._onSelectionChanged}
          />
        }
      </div>
    );
  }

  private async _getExtensionItems(): Promise<IUserCustomAction[]> {
    const dataService: IExtensionService = (Environment.type === EnvironmentType.Test || Environment.type === EnvironmentType.Local) ?
      new MockExtensionService() :
      new ExtensionService(this.props.webPartContext);
    return dataService.getExtensions();
  }

  private _onSelectionChanged(selection:IUserCustomAction[]): void {
    console.log("SelectionChanged", selection.length);
    this.setState({
      selection: selection
    });
  }
}
