import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'PnPWebPartStrings';
import PnP from './components/PnP';
import { IPnPProps } from './components/IPnPProps';

import { sp } from "@pnp/sp";

export interface IPnPWebPartProps {
  description: string;
}

export default class PnPWebPart extends BaseClientSideWebPart<IPnPWebPartProps> {

  public onInit(): Promise<void> {

    return super.onInit().then(_ => {
  
      // other init code may be present
  
      sp.setup({
        spfxContext: this.context
      });
    });
  }

  public render(): void {
    
    let pageUrl:string = "";
    if(this.context.pageContext.listItem != null && 
      this.context.pageContext.listItem != undefined)
    {
      pageUrl = window.location.pathname;
    }
    else
    {
      pageUrl = this.context.pageContext.web.serverRelativeUrl + "/SitePages/Page.aspx";
    }

    const element: React.ReactElement<IPnPProps > = React.createElement(
      PnP,
      {        
        pageUrl:pageUrl,
        spRest:sp,
        userLoginName:this.context.pageContext.user.loginName
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
