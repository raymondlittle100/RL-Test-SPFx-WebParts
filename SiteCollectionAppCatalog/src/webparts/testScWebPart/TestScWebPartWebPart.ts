import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'TestScWebPartWebPartStrings';
import TestScWebPart from './components/TestScWebPart';
import { ITestScWebPartProps } from './components/ITestScWebPartProps';

export interface ITestScWebPartWebPartProps {
  description: string;
}

export default class TestScWebPartWebPart extends BaseClientSideWebPart<ITestScWebPartWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ITestScWebPartProps > = React.createElement(
      TestScWebPart,
      {
        description: this.properties.description
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
