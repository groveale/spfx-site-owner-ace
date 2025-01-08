import {
  BaseComponentsCardView,
  ComponentsCardViewParameters,
  BasicCardView,
  IExternalLinkCardAction,
  IQuickViewCardAction
} from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'SiteOwnerAdaptiveCardExtensionStrings';
import {
  ISiteOwnerAdaptiveCardExtensionProps,
  ISiteOwnerAdaptiveCardExtensionState,
  QUICK_VIEW_REGISTRY_ID
} from '../SiteOwnerAdaptiveCardExtension';

export class CardView extends BaseComponentsCardView<
  ISiteOwnerAdaptiveCardExtensionProps,
  ISiteOwnerAdaptiveCardExtensionState,
  ComponentsCardViewParameters
> {
  public get cardViewParameters(): ComponentsCardViewParameters {
    console.log(this.state.siteItems.length);
    console.log(this.state.siteItems[0]);
    console.log(this.context.pageContext.user.email);
    return BasicCardView({
      cardBar: {
        componentName: 'cardBar',
        title: this.properties.title
      },
      header: {
        componentName: 'text',
        text: this.state.siteItems.length > 0 ? `${this.state.siteItems.length} site(s)` : 'No sites'
      },
      footer: {
        componentName: 'cardButton',
        title: strings.QuickViewButton,
        action: {
          type: 'QuickView',
          parameters: {
            view: QUICK_VIEW_REGISTRY_ID
          }
        }
      }
    });
  }

  public get onCardSelection(): IQuickViewCardAction | IExternalLinkCardAction | undefined {
    return {
      type: 'ExternalLink',
      parameters: {
        target: 'https://www.bing.com'
      }
    };
  }
}
