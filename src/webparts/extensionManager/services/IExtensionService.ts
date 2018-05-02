/**
 * IExtensionService
 */
import { IUserCustomAction } from "./IUserCustomAction";

export interface IExtensionService {
  getExtensions: () => Promise<IUserCustomAction[]>;
  getExtensionsByUrl: (url: string) => Promise<IUserCustomAction[]>;
}