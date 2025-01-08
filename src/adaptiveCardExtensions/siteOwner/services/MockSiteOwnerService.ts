import { ISiteOwnerService } from "./ISiteOwnerService";
import { ServiceKey } from "@microsoft/sp-core-library";
import { ISiteItem } from "../models/ISiteItem";

export class MockSiteOwnerService implements ISiteOwnerService {

    public static readonly serviceKey: ServiceKey<ISiteOwnerService> =
            ServiceKey.create<ISiteOwnerService>('Groveale.MockSiteOwnerService', MockSiteOwnerService);

    public getSiteItems(siteUrl: string, listTitle: string, userEmail: string): Promise<ISiteItem[]> {
        // Mock data
        console.log("MockSiteOwnerService: getSiteItems called");
        const mockData: ISiteItem[] = [
            {
                SiteUrl: "https://m365cpi77517573.sharepoint.com/sites/BoschAccount",
                Email: "admin@M365CPI77517573.onmicrosoft.com",
                UserType: "Owner",
                FileCount: 10,
                LastUserActivityDate: new Date("2025-01-01T14:35:36"),
                SiteName: "Bosch Account",
                LogTime: new Date("2025-01-01T14:35:36"),
                LockStatus: "Unlock",
                SiteTemplate: "Team Site",
                SharingCapability: "ExternalUserSharingOnly",
                SensitivityLabel: "General",
                CreatedDate: new Date("2024-07-01T14:35:36"),
                Visibility: "Private",
                HasTeam: true
            },
            // Add more mock items as needed
        ];

        // Filter mock data by userEmail
        const filteredData = mockData.filter(item => item.Email === userEmail);

        return Promise.resolve(filteredData);
    }
}