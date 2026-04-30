import { WebPartContext } from "@microsoft/sp-webpart-base";
// import pnp and pnp logging system
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
export const getSP = (context?: WebPartContext): SPFI => {
 
   
  if (context !== null && (_sp === undefined ||_sp === null)) {
    //You must add the @pnp/logging package to include the PnPLogging behavior it is no longer a peer dependency
    // The LogLevel set's at what level a message will be written to the console
    _sp = spfi().using(SPFx(context)).using(PnPLogging(LogLevel.Warning));
 
   
  }
  return _sp;
}
let _spurl: SPFI;
export const getSPContext = (context?: WebPartContext): SPFI=> {
  if (context !== null && (_spurl === undefined ||_spurl === null)) {
    //You must add the @pnp/logging package to include the PnPLogging behavior it is no longer a peer dependency
    // The LogLevel set's at what level a message will be written to the console
   _spurl = spfi().using(SPFx(context)).using(PnPLogging(LogLevel.Warning));
 
   
  }
  return _spurl;
 
}
// export const getContext=(context?: WebPartContext)=>
// {
// const _context = context
// return _context;
// }
 
 