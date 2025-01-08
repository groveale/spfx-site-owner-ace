export interface ISiteItem {
    SiteUrl: string;
    Email: string;
    UserType: string;
    FileCount: number | null;
    LastUserActivityDate: Date | null;
    SiteName: string;
    LogTime: Date;
    LockStatus: string;
    SiteTemplate: string | null;
    SharingCapability: string;
    SensitivityLabel: string;
    CreatedDate: Date;
    Visibility: string;
    HasTeam: boolean;
}