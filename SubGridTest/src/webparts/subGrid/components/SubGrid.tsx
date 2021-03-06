import * as React from 'react';
import styles from './SubGrid.module.scss';
import { ISubGridProps, IStepState, IProjectStepsState } from '../InterfaceFiles';
import { escape } from '@microsoft/sp-lodash-subset';

import * as pnp from 'sp-pnp-js';
import * as moment from 'moment';

// Import React Table
import ReactTable from "react-table";
import "react-table/react-table.css";

export default class SubGrid extends React.Component<ISubGridProps, IProjectStepsState> {

  //declare private variables
  private web;
  private projectStepsList;

  constructor(props) {
    super(props);

    //set the values of the private variables
    this.web = new pnp.Web(this.props.webUrl);
    this.projectStepsList = this.web.lists.getByTitle(this.props.projectListName);
    
    this.state = {
      steps: [],
      loading:true,
      error:""
    };

    this.renderEditable = this.renderEditable.bind(this);
    this.getProjectSteps = this.getProjectSteps.bind(this);
  }

  public componentDidMount(): void {
    this.getProjectSteps();    
  } 

  private getProjectSteps() {
    try {
      let query = {
        "ViewXml": "<View>" +
                     "<ViewFields>" +
                        "<FieldRef Name='Title' />" +                        
                        "<FieldRef Name='Grouping' />" +   
                        "<FieldRef Name='StepStartDate' />" +   
                        "<FieldRef Name='ProjectDueDate' />" +   
                        "<FieldRef Name='StepEstimatedDuration' />" +   
                        "<FieldRef Name='StepNumber' />" +   
                      "</ViewFields>" +
                      "<OrderBy>" +
                        "<FieldRef Name='StepNumber' />" +
                      "</OrderBy>" +
                    "</View>"
      };
      this.projectStepsList.getItemsByCAMLQuery(query).then((items) => {        
        let steps: IStepState[] = [];
        if(items.length > 0) {
          items.map(item=>{
            let startDate:string="";
            let endDate:string="";
            if(item.StepStartDate != null && item.StepStartDate != "")
            {
              const date:Date =new Date(item.StepStartDate);
              startDate = moment(date).format(this.props.dateFormat);
            }

            if(item.ProjectDueDate != null && item.ProjectDueDate != "")
            {
              const date:Date =new Date(item.ProjectDueDate);
              endDate = moment(date).format(this.props.dateFormat);
            }

            steps.push(
            {
              id: item.Id, 
              title: item.Title,
              grouping:item.Grouping,
              startDate:startDate,
              endDate:endDate,
              effort:item.StepEstimatedDuration 
            });        
          });          
        }
        this.setState({ steps : steps, loading:false });
      }).catch(error=>{        
        console.log("Query error with getsteps: ", error.message);    
        this.setState({ error: error.message });            
      });      
    } 
    catch (error) {
      console.log("Catch error with getsteps: " + error.message);
      this.setState({ error: error.message });            
    }
  }

  renderEditable(cellInfo) {
    return (
      <div
        style={{ backgroundColor: "#fafafa" }}
        contentEditable
        suppressContentEditableWarning
        onBlur={e => {
          const data = [...this.state.steps];
          data[cellInfo.index][cellInfo.column.id] = e.currentTarget.innerHTML;
          this.setState({ steps: data });
        }}
        dangerouslySetInnerHTML={{
          __html: this.state.steps[cellInfo.index][cellInfo.column.id]
        }}
      />
    );
  }

  public render(): React.ReactElement<ISubGridProps> {     
    if(this.state.error != null && this.state.error !="") 
    {
      return <div>Error: {this.state.error}</div>;
    }
    else
    {
      return (
        <ReactTable
        data={this.state.steps}
        loading={this.state.loading}
        columns={[
          {          
            columns: [
              {
                Header: "Title",
                accessor: "title",
                Cell: this.renderEditable,
                minWidth:150,
                maxWidth:250,
                resizable:true
              },
              {
                Header: "Grouping",
                accessor: "grouping",          
                minWidth:150,
                maxWidth:250,
                resizable:true,
                headerStyle:{"text-align":"left"},
                headerClassName:"balls"
              },
              {
                Header: "Start Date",
                accessor: "startDate",
                resizable:false,
                maxWidth:110
              },
              {
                Header: "End Date",
                accessor: "endDate",  
                resizable:false,
                maxWidth:110          
              },
              {
                Header: "Effort",
                accessor: "effort",  
                Cell: this.renderEditable,
                resizable:false,
                maxWidth:70       
              }
            ]
          }
        ]}
        sortable={false}
        filterable={false}
        resizable={true}
        noDataText="No steps could be found"
        { ...( this.state.steps.length <40 && {minRows: this.state.steps.length+1})}                
        pageSize={40}        
        className="-striped -highlight test"
        showPageSizeOptions={false}            
      />  
      );
    }
  }
}
