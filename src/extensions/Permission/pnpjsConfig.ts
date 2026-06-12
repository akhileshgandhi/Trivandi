import { BaseComponentContext } from "@microsoft/sp-component-base";
import { spfi, SPFI, SPFx } from "@pnp/sp";
import { LogLevel, PnPLogging } from "@pnp/logging";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/batching";
import "@pnp/sp/site-users/web";
import "@pnp/sp/profiles";
import "@pnp/sp/items/get-all";
import "@pnp/sp/folders";
import "@pnp/sp/files/folder";
import "@pnp/sp/fields";
import "@pnp/sp/files";
import "@pnp/sp/security";
import "@pnp/sp/presets/all";

let _sp: SPFI;
export const getSP = (context?: BaseComponentContext | any): SPFI => {
  if (context !== null && context !== undefined) {
    _sp = spfi().using(SPFx(context)).using(PnPLogging(LogLevel.Warning));
  }
  return _sp;
};

let _spurl: SPFI;
export const getSPContext = (context?: BaseComponentContext | any): SPFI => {
  if (context !== null && context !== undefined) {
    _spurl = spfi().using(SPFx(context)).using(PnPLogging(LogLevel.Warning));
  }
  return _spurl;
};
