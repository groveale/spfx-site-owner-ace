import { ISPFxAdaptiveCard, BaseAdaptiveCardQuickView } from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'SiteOwnerAdaptiveCardExtensionStrings';
import {
  ISiteOwnerAdaptiveCardExtensionProps,
  ISiteOwnerAdaptiveCardExtensionState
} from '../SiteOwnerAdaptiveCardExtension';
import { ISiteItem } from '../models/ISiteItem';

export interface IQuickViewData {
  subTitle: string;
  title: string;
  siteItems: ISiteItem[];
}

export class QuickView extends BaseAdaptiveCardQuickView<
  ISiteOwnerAdaptiveCardExtensionProps,
  ISiteOwnerAdaptiveCardExtensionState,
  IQuickViewData
> {
  public get data(): IQuickViewData {
    return {
      subTitle: strings.SubTitle,
      title: strings.Title,
      siteItems: this.state.siteItems
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return require('./template/QuickViewTemplate.json');
  }
}
