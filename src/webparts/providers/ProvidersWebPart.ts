import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import { MSGraphClient, HttpClient } from '@microsoft/sp-http';

import * as strings from 'ProvidersWebPartStrings';
import Providers from './components/Providers';
import { IProvidersProp } from './components/IProvidersProps';

export interface IProvidersWebPartProps {
  description: string;
}

export default class ProvidersWebPart extends BaseClientSideWebPart<IProvidersWebPartProps> {

  public render(): void {

    var currentContext = this.context;


    this.context.msGraphClientFactory.getClient()
      .then((_graphClient: MSGraphClient): void => {
        const element: React.ReactElement<IProvidersProp> = React.createElement(
          Providers,
          {
            description: this.properties.description,
            currentContext: currentContext,
            siteUrl: this.context.pageContext.web.absoluteUrl,
            graphClient: _graphClient,
          }
        );

        ReactDom.render(element, this.domElement);
      });
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
