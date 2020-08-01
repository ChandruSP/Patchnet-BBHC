import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import { MSGraphClient, HttpClient } from '@microsoft/sp-http';

import * as strings from 'VisitorsGroupWebPartStrings';
import VisitorsGroup from './components/VisitorsGroup';
import { IVisitorsGroupProps } from './components/IVisitorsGroupProps';

export interface IVisitorsGroupWebPartProps {
  description: string;
}

export default class VisitorsGroupWebPart extends BaseClientSideWebPart<IVisitorsGroupWebPartProps> {

  public render(): void {

    var currentContext = this.context;

    this.context.msGraphClientFactory.getClient()
      .then((_graphClient: MSGraphClient): void => {

        const element: React.ReactElement<IVisitorsGroupProps> = React.createElement(
          VisitorsGroup,
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
