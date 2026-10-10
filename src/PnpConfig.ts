import { ISPFXContext, spfi,SPFI,SPFx } from "@pnp/sp";
import "@pnp/sp"
import "@pnp/sp/webs"
import "@pnp/sp/lists"
import "@pnp/sp/items"
import "@pnp/sp/fields"

let sp:SPFI;
export const get_SP=(context?:ISPFXContext):SPFI=>{
    if(!sp && context){
        sp=spfi().using(SPFx(context))
    }
    return sp
}