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
            providerAssignedHTML: `<html lang="en">
            <head>
              <meta charset="UTF-8" />
              <meta name="viewport" content="width=device-width, initial-scale=1.0" />
              <title>Second</title>
              <link
                href="https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;500;700&display=swap"
                rel="stylesheet"
              />
            </head>
            <body style="font-family: 'Noto Sans JP', sans-serif;">
              <p>Hi,</p>
          
              <p>Broward Behavioral Health has granted you access to your Provider SharePoint Library so you can upload your agreement documentation.</p>
              <ol>
                <li>
                To login visit the BBHC Provider SharePoint Site ->
                   <a href="http://provider.bbhcflorida.org/"
                     >http://provider.bbhcflorida.org/</a
                   >.
                 </li>
               </ol>
             </body>
           </html>
           `
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
