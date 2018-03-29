import * as React from 'react';
import styles from './PnP.module.scss';
import { IPnPProps } from './IPnPProps';
import { escape } from '@microsoft/sp-lodash-subset';

import {
  ClientSidePage,
  Item,
  ListItemFormUpdateValue
} from '@pnp/sp';

export default class PnP extends React.Component<IPnPProps, {}> 
{

  constructor(props)
  {
    super(props)

    this.loadPageDetails = this.loadPageDetails.bind(this);
    this.loadTestListDetails = this.loadTestListDetails.bind(this);
  }

  public componentDidMount()
  {
    //this.loadPageDetails();
    this.loadTestListDetails()
  } 

  public async loadTestListDetails()
  {
    try
    {
      // const page = await ClientSidePage.fromFile(
      //   this.props.spRest.web.getFileByServerRelativeUrl(this.props.pageUrl));        
      // const pageItem = await page.getItem("Title", "CommentsDisabled", "PromotedState", "UserOne");

      const listItem = await this.props.spRest.web.getList("/sites/Finance-Dev/Lists/test%20list").items.getById(1).select("Title","UserOne");
            
      const fieldNames = await this.getPageFieldNames();

      const userId= await this.getUserByLoginName();
      const pageSaveResults:ListItemFormUpdateValue[] = await this.savePageMetadata(listItem, userId);
    }
    catch(error)
    {
      console.log(`Error ${error}`);
      console.log(error);
      let message = error.data.responseBody["odata.error"].message.value;
      console.log(message);
    }
  }

  public async loadPageDetails()
  {
    try
    {
      const page = await ClientSidePage.fromFile(
        this.props.spRest.web.getFileByServerRelativeUrl(this.props.pageUrl));        
      const pageItem = await page.getItem("Title", "CommentsDisabled", "PromotedState", "UserOne");
      
      const fieldNames = await this.getPageFieldNames();

      const userId= await this.getUserByLoginName();
      const pageSaveResults:ListItemFormUpdateValue[] = await this.savePageMetadata(pageItem, userId);
    }
    catch(error)
    {
      console.log(`Error ${error}`);
      console.log(error);
      let message = error.data.responseBody["odata.error"].message.value;
      console.log(message);
    }
  }

  public async getUserByLoginName() : Promise<any[]>
  {
    const currentUser = await this.props.spRest.web.currentUser.get();
    console.log(currentUser);
    return currentUser.Id;
  }

  public async getPageFieldNames() : Promise<any[]>
  {
    // const fieldDetails = await this.props.spRest.web.lists
    // .getByTitle('Site Pages')
    // .fields
    // .select('Title, EntityPropertyName')
    // .filter(`Hidden eq false and Title eq 'UserOne'`)
    // .get();    

    const fieldDetails = await this.props.spRest.web.lists
    .getByTitle('test list')
    .fields
    .select('Title, EntityPropertyName')
    .filter(`Hidden eq false and Title eq 'UserOne'`)
    .get();   
    
    console.log(fieldDetails.map(field => {
      return {
        Title: field.Title,
        EntityPropertyName: field.EntityPropertyName
      };
    }));

    return fieldDetails
  }

  public async savePageMetadata(pageListItem:Item, userId:any):Promise<ListItemFormUpdateValue[]>
  {
    //my list of objects to pass to the function
    let fieldValues = this.createItemUpdateValues(userId);
    console.log(fieldValues);
    const pageSaveResults = await pageListItem.validateUpdateListItem(fieldValues);
    
    console.log(`Saved`);
    console.log(pageSaveResults)
    return pageSaveResults;
  }

  private createItemUpdateValues(userId:any):ListItemFormUpdateValue[]
  {
    //setup array of object
    let pageValues:ListItemFormUpdateValue[] =[];

    const userSaveValue = userId.toString();
    //update with ID field name
    const ownerId:ListItemFormUpdateValue =
    {
      FieldName:"UserOneId",
      FieldValue:userSaveValue
    };

     //update with entity property name
     const owner:ListItemFormUpdateValue =
     {
       FieldName:"UserOne",
       FieldValue:userSaveValue
     };     

    //add static values to array 
    //pageValues.push(ownerId);         
    pageValues.push(owner);    
    
    return pageValues;
  }

  public render(): React.ReactElement<IPnPProps> {
    return (
      <div className={ styles.pnP }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>              
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
