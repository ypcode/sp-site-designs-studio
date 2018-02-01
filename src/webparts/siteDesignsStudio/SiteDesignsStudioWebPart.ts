import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version , ServiceScope} from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'SiteDesignsStudioWebPartStrings';
import SiteDesignsStudio from './components/SiteDesignsStudio';
import { ISiteDesignsStudioProps } from './components/ISiteDesignsStudioProps';
import { AppStartup } from './AppStartup';

export interface ISiteDesignsStudioWebPartProps {

}

export default class SiteDesignsStudioWebPart extends BaseClientSideWebPart<ISiteDesignsStudioWebPartProps> {
  private usedServiceScope: ServiceScope;


  public onInit(): Promise<any> {
    return super.onInit()
      // Set the global configuration of the application
      // This is where we will define the proper services according to the context (Local, Test, Prod,...)
      // or according to specific settings
      .then(_ => AppStartup.configureServices(this.context))
      // When configuration is done, we get the instances of the services we want to use
      .then(serviceScope => {
        this.usedServiceScope = serviceScope;
        // Get services instance references here
        // this.dataService = serviceScope.consume(DataServiceKey);
        // this.config = serviceScope.consume(ConfigurationServiceKey);
      });
  }

  public render(): void {
    const element: React.ReactElement<ISiteDesignsStudioProps > = React.createElement(
      SiteDesignsStudio,
      {
        serviceScope: this.usedServiceScope
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [

              ]
            }
          ]
        }
      ]
    };
  }
}
