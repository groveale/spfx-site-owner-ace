import type { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { QuickView } from './quickView/QuickView';
import { SiteOwnerPropertyPane } from './SiteOwnerPropertyPane';
import { ISiteOwnerService } from './services/ISiteOwnerService';
import { SiteOwnerService } from './services/SiteOwnerService';
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
  private _client: ISiteOwnerService;

  public onInit(): Promise<void> {
    this.state = {
      siteItems: []
    };

    // registers the card view to be shown in a dashboard
    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    // registers the quick view to open via QuickView action
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());

    // consume the service
    this._client = this.context.serviceScope.consume(SiteOwnerService.serviceKey);

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
    return this._client.getSiteItems(this.properties.siteUrl, this.properties.listTitle, this.context.pageContext.user.email)
      // Step 5: After mapping the JSON response to menu items, update the component's state with the retrieved menu items.
      .then((items) => this.setState(
        {
          siteItems: items
        }));
  }
}
