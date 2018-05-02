/**
 * Renders a web part that calls the extension manager component
 */
import { Version } from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
} from "@microsoft/sp-webpart-base";
import * as React from "react";
import * as ReactDom from "react-dom";
import { ExtensionManager } from "./components/ExtensionManager";
import { IExtensionManagerProps } from "./components/ExtensionManager.types";

export interface IExtensionManagerWebPartProps {
  // empty
}

export default class ExtensionManagerWebPart extends BaseClientSideWebPart<IExtensionManagerWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IExtensionManagerProps> = React.createElement(
      ExtensionManager,
      {
        webPartContext: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }
}
