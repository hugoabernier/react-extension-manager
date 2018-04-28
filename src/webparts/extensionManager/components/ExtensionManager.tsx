/**
 * ExtensionManager
 */
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
  IExtensionManagerState
} from "./IExtensionManager.types";

export class ExtensionManager extends React.Component<IExtensionManagerProps, IExtensionManagerState> {
  private maxResults: number = 1000;
  private extensionItems: IUserCustomAction[] = [];
  private selectedExtensionItems: IUserCustomAction[] = [];

  constructor(props: IExtensionManagerProps) {
    super(props);

    this.state = {
      dataLoaded: false
    };

    this.onSelectionChanged = this.onSelectionChanged.bind(this);
  }

  public async componentDidMount(): Promise<void> {
    this.props.webPartContext.statusRenderer.displayLoadingIndicator(
      document.getElementsByClassName(styles.extensionManager)[0], strings.LoadingLabel);

    this.extensionItems = await this.getExtensionItems();
    this.setState({
      dataLoaded: true
    });
    this.props.webPartContext.statusRenderer.clearLoadingIndicator(
      document.getElementsByClassName(styles.extensionManager)[0]);
  }

  public render(): React.ReactElement<IExtensionManagerProps> {
    return (
      <div className={styles.extensionManager}>
        {this.state.dataLoaded &&
          <ExtensionListView
            items={this.extensionItems}
            defaultSelection={[]}
            // onSelectionChanged={this._onSelectionChanged}
          />
        }
      </div>
    );
  }

  private async getExtensionItems(): Promise<IUserCustomAction[]> {
    const dataService: IExtensionService = (Environment.type === EnvironmentType.Test || Environment.type === EnvironmentType.Local) ?
      new MockExtensionService() :
      new ExtensionService(this.props.webPartContext);
    return dataService.getExtensions();
  }

  private onSelectionChanged(selection: IUserCustomAction[]): void {
    this.setState({
      selection: selection
    });
  }
}
