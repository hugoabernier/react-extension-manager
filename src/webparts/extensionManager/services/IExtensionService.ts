/**
 * IExtensionService
 */
import { IUserCustomAction } from "./IUserCustomAction";
import { IUserCustomActionCollection } from "./IUserCustomActionCollection";

export interface IExtensionService {
  getExtensions: () => Promise<IUserCustomAction[]>;
  getExtensionsByUrl: (url: string) => Promise<IUserCustomAction[]>;
}