import { Version } from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from "@microsoft/sp-webpart-base";
import * as strings from "ExtensionManagerWebPartStrings";
import * as React from "react";
import * as ReactDom from "react-dom";
import ExtensionManager from "./components/ExtensionManager";
import { IExtensionManagerProps } from "./components/IExtensionManager.types";

export interface IExtensionManagerWebPartProps {
  description: string;
}

export default class ExtensionManagerWebPart extends BaseClientSideWebPart<IExtensionManagerWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IExtensionManagerProps > = React.createElement(
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

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
