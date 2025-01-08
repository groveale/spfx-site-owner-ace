import { ISiteItem } from "../models/ISiteItem";

export interface ISiteOwnerService {  
    getSiteItems(siteUrl: string, listTitle: string, userEmail: string): Promise<ISiteItem[]>;
}