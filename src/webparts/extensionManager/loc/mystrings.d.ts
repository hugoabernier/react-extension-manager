/**
 * mystrings.d.ts
 */
declare interface IExtensionManagerWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  LoadingLabel: string;
  UnknowScopeLabel: string;
  SiteScopeLabel: string;
  WebScopeLabel: string;
  ListScopeLabel: string;
  NAScopeLabel: string;
  NoneRegistrationTypeLabel: string;
  ListRegistrationTypeLabel: string;
  ContentTypeRegistrationTypeLabel: string;
  ProgIdRegistrationTypeLabel: string;
  FileTypeRegistrationTypeLabel: string;
  NARegistrationTypeLabel: string;
  ScopeHeader: string;
  TitleHeader: string;
  RegistrationTypeHeader: string;
  LocationHeader: string;
  ApplicationCustomizerLocation: string;
  CommandBarLocation: string;
  ContextMenuLocation: string;
  ListViewLocation: string;
  ECBLocation: string;
  NewButton: string;
  NewButtonAriaLabel: string;
  EditButton: string;
  EditButtonLabel: string;
  DeleteButton: string;
  UploadButton: string;
  UploadButtonAriaLabel: string;
  InfoButton: string;
  EditButton: string;
  GroupByMenu: string;
  AscendingSort: string;
  DescendingSort: string;
  DeleteDialogTitle: string;
  DeleteDialogDescription: string;
  ClearSelectionButton: string;
  ClearSelectionButtonTitle: string;
  ClearSelectionButtonAriaLabel: string;
  SaveButton: string;
  SaveButtonAriaLabel: string;
  CancelButton: string;
  CancelButtonAriaLabel: string;
}

declare module "ExtensionManagerWebPartStrings" {
  const strings: IExtensionManagerWebPartStrings;
  export = strings;
}
