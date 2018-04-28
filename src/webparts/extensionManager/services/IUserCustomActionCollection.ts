/**
 * IUserCustomActionCollection
 */
import { IUserCustomAction } from "./IUserCustomAction";

export interface IUserCustomActionCollection {
    "@odata.context": string;
    value: IUserCustomAction[];
}