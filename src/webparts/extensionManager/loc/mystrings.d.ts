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
  ECBLocation:string;
}

declare module 'ExtensionManagerWebPartStrings' {
  const strings: IExtensionManagerWebPartStrings;
  export = strings;
}
