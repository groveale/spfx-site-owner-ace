import type { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { QuickView } from './quickView/QuickView';
import { SiteOwnerPropertyPane } from './SiteOwnerPropertyPane';
import { SPHttpClient } from '@microsoft/sp-http'
import { ISiteItem } from './models/ISiteItem';



export interface ISiteOwnerAdaptiveCardExtensionProps {
  title: string;
  listTitle: string;
  siteUrl: string;
}

export interface ISiteOwnerAdaptiveCardExtensionState {
  siteItems: ISiteItem[];
}

const CARD_VIEW_REGISTRY_ID: string = 'SiteOwner_CARD_VIEW';
export const QUICK_VIEW_REGISTRY_ID: string = 'SiteOwner_QUICK_VIEW';

export default class SiteOwnerAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  ISiteOwnerAdaptiveCardExtensionProps,
  ISiteOwnerAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: SiteOwnerPropertyPane;

  public onInit(): Promise<void> {
    this.state = {
      siteItems: []
     };

    // registers the card view to be shown in a dashboard
    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    // registers the quick view to open via QuickView action
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());

    //return Promise.resolve();
    return this._fetchData();
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'SiteOwner-property-pane'*/
      './SiteOwnerPropertyPane'
    )
      .then(
        (component) => {
          this._deferredPropertyPane = new component.SiteOwnerPropertyPane();
        }
      );
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane?.getPropertyPaneConfiguration();
  }

  private _fetchData(): Promise<void> {
    console.log('fetching data');
    // log url
    console.log(`${this.properties.siteUrl}` +
    `/_api/web/lists/getByTitle('${this.properties.listTitle}')/items`);
    // Step 1: Use the SharePoint HTTP client to send a GET request to retrieve data from the specified SharePoint list.
    // The URL for the request is built using the SharePoint API with the list title from the properties.
    return this.context.spHttpClient.get(
      `${this.properties.siteUrl}` +
      `/_api/web/lists/getByTitle('${this.properties.listTitle}')/items?$filter=field_1 eq '${this.context.pageContext.user.email}'`,
      SPHttpClient.configurations.v1
    )
    // Step 2: After getting a response from the server, convert it to JSON format.
    .then((response) => response.json())
    // Step 3: Map the JSON response to a new array of objects representing the menu items.
    .then((jsonResponse) => jsonResponse.value.map(
      (item: any) => 
      { 
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
    // Step 5: After mapping the JSON response to menu items, update the component's state with the retrieved menu items.
    .then((items) => this.setState(
      { 
        siteItems: items 
      }));
  }
}
