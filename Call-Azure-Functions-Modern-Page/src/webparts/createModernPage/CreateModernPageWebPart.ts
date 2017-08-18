import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'createModernPageStrings';
import CreateModernPage from './components/CreateModernPage';
import { ICreateModernPageProps } from './components/ICreateModernPageProps';
import { ICreateModernPageWebPartProps } from './ICreateModernPageWebPartProps';

export default class CreateModernPageWebPart extends BaseClientSideWebPart<ICreateModernPageWebPartProps> {


  public render(): void {
    const element: React.ReactElement<ICreateModernPageProps > = React.createElement(
      CreateModernPage,
      {
        siteUrl: "https://companynetcloud.sharepoint.com/sites/provenbase-dev/RayTest",//this.context.pageContext.site.absoluteUrl,
        functionUrl:"https://raytestfunctions.azurewebsites.net/api/CreateModernPages?code=IbDHZ3HVVyo4hPCyzf3aVNEfWMagarO8UAcQg19s3bnyGcigI9hCiw==",
        httpClient : this.context.httpClient
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
