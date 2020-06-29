import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ProviderDocumentsWebPartStrings';
import ProviderDocuments from './components/ProviderDocuments';
import { IProviderDocumentsProps } from './components/IProviderDocumentsProps';

export interface IProviderDocumentsWebPartProps {
  description: string;
}

export default class ProviderDocumentsWebPart extends BaseClientSideWebPart<IProviderDocumentsWebPartProps> {


  public render(): void {

    var currentContext = this.context;

    const element: React.ReactElement<IProviderDocumentsProps> = React.createElement(
      ProviderDocuments,
      {
        description: this.properties.description,
        currentContext: currentContext,
        siteUrl: this.context.pageContext.web.absoluteUrl
      }
    );

    ReactDom.render(element, this.domElement);
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
