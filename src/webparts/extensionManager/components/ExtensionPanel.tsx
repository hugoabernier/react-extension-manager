/**
 * Responsible for showing the edit panel for a componentgit
 */
import * as React from "react";
import * as strings from "ExtensionManagerWebPartStrings";
import {
    CommandBar,
    DefaultButton,
    Panel,
    PanelType,
    PrimaryButton,
    IContextualMenuItem
} from "office-ui-fabric-react";
import { unescape } from "@microsoft/sp-lodash-subset";
import AceEditor from "react-ace";
import * as ace from "brace";
import "brace/mode/json";
import "brace/theme/github";
export interface IExtensionPanelProps {
    isOpen: boolean;
    onDismiss: () => void;
}

export interface IExtensionPanelState {
    //
}

const sampleObject: string = "{&quot;sampleTextOne&quot;:&quot;One item is selected in the list.&quot;,"
+ "&quot;sampleTextTwo&quot;:&quot;This command is always visible.&quot;}";

export class ExtensionPanel extends React.Component<IExtensionPanelProps, IExtensionPanelState> {
    private _paneCommands: IContextualMenuItem[] = [
        {
          key: "saveItem",
          name: strings.SaveButton,
          icon: "Save",
          ariaLabel: strings.SaveButtonAriaLabel,
          ["data-automation-id"]: "saveButton",
          onClick: this.props.onDismiss,
        },
        {
          key: "cancelItem",
          name: strings.CancelButton,
          icon: "Cancel",
          ariaLabel: strings.CancelButtonAriaLabel,
          ["data-automation-id"]: "cancelButton",
          onClick: this.props.onDismiss,
        }
      ];

    public render(): React.ReactElement<IExtensionPanelProps> {

        // automatically convert json string to an object so that we can format the string
        let sampleObjectClean: string = unescape(sampleObject);
        console.log("CleanedObject", sampleObjectClean);
        const jsonObject: any = JSON.parse(sampleObjectClean);
        const jsonString: string = JSON.stringify(jsonObject, null, "\t");

        return (
            <Panel
                isOpen={this.props.isOpen}
                onDismiss={this.props.onDismiss}
                type={PanelType.medium}
                onRenderNavigation={ this._onRenderNavigation }
                onRenderFooterContent={ this._onRenderFooter }
                headerText="Custom Panel with custom 888px width"
            >
                <span>Content goes here.</span>
                <AceEditor
  mode="json"
  theme="github"
  name="blah2"
  onChange={this._handleJsonChange}
  fontSize={14}
  showPrintMargin={true}
  showGutter={true}
  highlightActiveLine={true}
  value={jsonString}
  setOptions={{
  enableBasicAutocompletion: true,
  enableLiveAutocompletion: true,
  enableSnippets: false,
  showLineNumbers: true,
  tabSize: 2,
  }}/>
            </Panel>
        );
    }

    private _handleJsonChange = (): void => {
        // do nothing
    }

    private _onRenderNavigation = (): JSX.Element => {
        return (
            <CommandBar
        isSearchBoxVisible={false}
        items={this._paneCommands}
      />
        );
      }

      private _onRenderFooter = (): JSX.Element => {
                  return (
         <div>
            <PrimaryButton
              onClick={ this.props.onDismiss }
            >
              {strings.SaveButton}
            </PrimaryButton>
            <DefaultButton
              onClick={ this.props.onDismiss }
            >
              {strings.CancelButton}
            </DefaultButton>
          </div>
        );
      }
}
