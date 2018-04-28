import {
    ISPHttpClientOptions,
    SPHttpClient,
    SPHttpClientResponse
} from "@microsoft/sp-http";
import {
    IWebPartContext
} from "@microsoft/sp-webpart-base";
import { IExtensionService } from "./IExtensionService";
import { IUserCustomAction } from "./IUserCustomAction";
import { IUserCustomActionCollection } from "./IUserCustomActionCollection";

export class ExtensionService implements IExtensionService {
    constructor(private context: IWebPartContext) {
        //
    }

    public async getExtensions(): Promise<IUserCustomAction[]> {
        console.log("getExtension", this.context);
        const webUrl: string = this.context.pageContext.web.absoluteUrl;
        return this.getExtensionsByUrl(webUrl);
    }

    public async getExtensionsByUrl(url:string): Promise<IUserCustomAction[]> {
        const apiUrl: string = `${url}/_api/web/UserCustomActions`;

        try {
            // get tasks
            return await this.context.spHttpClient.get(
                apiUrl,
                SPHttpClient.configurations.v1)
                .then((data: SPHttpClientResponse) => data.json())
                .then((data: IUserCustomActionCollection) => {
                    console.log("Response from onInit", data);
                    return data.value;
                });
        } catch (error) {
            console.log("Error loading extensions: ", error);
        }
    }
}