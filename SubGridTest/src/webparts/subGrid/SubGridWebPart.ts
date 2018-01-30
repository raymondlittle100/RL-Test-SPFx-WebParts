import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'SubGridWebPartStrings';
import SubGrid from './components/SubGrid';
// import ReactListSubGrid from './components/ReactListSubGrid';

import { ISubGridProps, ISubGridWebPartProps } from './InterfaceFiles';



export default class SubGridWebPart extends BaseClientSideWebPart<ISubGridWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISubGridProps > = React.createElement(
      SubGrid,
      {
        webUrl: this.context.pageContext.web.absoluteUrl,
        projectListName: this.properties.projectListName,
        dateFormat:this.properties.dateFormat
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
                PropertyPaneTextField('projectListName', {
                  label: strings.ProjectListNameFieldLabel
                }),
                PropertyPaneTextField('dateFormat', {
                  label: strings.DateFormatFieldLabel
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
