export interface ISiteItem {
    SiteUrl: string;
    Email: string;
    PermissionSource: string;
    UserType: string;
    DomainList: string;
    FileCount: number | null;
    LastContentModifiedDate: Date | null;
    SiteName: string;
    Notes: string | null;
    LogTime: Date;
    GroupName: string;
    LockStatus: string;
    SiteTemplate: string | null;
    GroupId: string;
    SharingCapability: string;
}