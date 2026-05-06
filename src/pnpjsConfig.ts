import { WebPartContext } from "@microsoft/sp-webpart-base";
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
let _context: WebPartContext;

export const getSP = (context?: WebPartContext): SPFI => {
  if (context !== null && context !== undefined) {
    _context = context;
  }

  if (_context && (_sp === undefined || _sp === null)) {
    _sp = spfi().using(SPFx(_context)).using(PnPLogging(LogLevel.Warning));
  }
  return _sp;
};

export const getContext = (): WebPartContext => {
  return _context;
};

