import { 
  SPRest  
} from "@pnp/sp";

export interface IPnPProps {  
  pageUrl: string;
  spRest:SPRest
  userLoginName:string
}
