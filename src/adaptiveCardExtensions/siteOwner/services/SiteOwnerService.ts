import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { ISiteItem } from "../models/ISiteItem";
import { ISiteOwnerService } from "./ISiteOwnerService";
import { SPHttpClient } from '@microsoft/sp-http'

export class SiteOwnerService implements ISiteOwnerService {

    public static readonly serviceKey: ServiceKey<ISiteOwnerService> =
        ServiceKey.create<ISiteOwnerService>('Groveale.SiteOwnerService', SiteOwnerService);

    private _client: SPHttpClient;

    constructor(serviceScope: ServiceScope) {
        serviceScope.whenFinished(() => {
            this._client = serviceScope.consume(SPHttpClient.serviceKey);
        });
    }

    getSiteItems(siteUrl: string, listTitle: string, userEmail: string): Promise<ISiteItem[]> {
        // Check if the SiteOwnerService has been initialized.
        if (this._client === undefined) {
            throw new Error('SiteOwnerService not initialized!')
        }

        console.log("SiteOwnerService: getSiteItems called");

        // Step 1: Use the SharePoint HTTP client to send a GET request to retrieve data from the specified SharePoint list.
        // The URL for the request is built using the SharePoint API with the list title from the properties.
        return this._client.get(
            `${siteUrl}` +
            `/_api/web/lists/getByTitle('${listTitle}')/items?$filter=field_1 eq '${userEmail}'`,
            SPHttpClient.configurations.v1
        )
            // Step 2: After getting a response from the server, convert it to JSON format.
            .then((response) => response.json())
            // Step 3: Map the JSON response to a new array of objects representing the menu items.
            .then((jsonResponse) => jsonResponse.value.map(
                (item: any) => {
                    // Step 4: Extract specific properties (Title, Description, Day, ImageUrl) from each item in the JSON response.
                    // Return a new object for each item with the extracted properties.
                    return {
                        SiteUrl: item.Title,
                        Email: item.field_1,
                        PermissionSource: item.field_2,
                        UserType: item.field_3,
                        DomainList: item.field_4,
                        FileCount: item.field_5,
                        LastContentModifiedDate: item.field_6,
                        SiteName: item.field_7,
                        Notes: item.field_8,
                        LogTime: item.field_9,
                        GroupName: item.field_10,
                        LockStatus: item.field_11,
                        SiteTemplate: item.field_12,
                        GroupId: item.field_13,
                        SharingCapability: item.field_14
                    };
                }))
    }
}