import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import * as strings from 'SiteOwnerAdaptiveCardExtensionStrings';

export class SiteOwnerPropertyPane {
  public getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                }),
                PropertyPaneTextField('listTitle', {
                  label: strings.ListTitleFieldLabel
                }),
                PropertyPaneTextField('siteUrl', {
                  label: strings.SiteUrlFieldLabel
                }),
                PropertyPaneTextField('useMock', {
                  label: strings.UseMockFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
