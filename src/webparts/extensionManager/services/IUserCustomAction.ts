
export interface IUserCustomAction {
    "@odata.type": string;
    "@odata.id": string;
    "@odata.editLink": string;
    ClientSideComponentId: string;
    ClientSideComponentProperties: string;
    CommandUIExtension: string;
    Description: string;
    Group: string;
    Id: string;
    ImageUrl: string;
    Location: string;
    Name: string;
    RegistrationId: number;
    RegistrationType: number;
    Rights: IRights;
    Scope: number;
    ScriptBlock: string;
    ScriptSrc: string;
    Sequence: number;
    Title: string;
    Url: string;
    VersionOfUserCustomAction: string;
}

export interface IRights {
    High: number;
    Low: number;
}
