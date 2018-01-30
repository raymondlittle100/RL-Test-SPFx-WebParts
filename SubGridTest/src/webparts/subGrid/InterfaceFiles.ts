export interface ISubGridWebPartProps {    
    projectListName: string;
    dateFormat:string;
}

export interface ISubGridProps {
    webUrl:string;
    projectListName: string;
    dateFormat:string;
}

//table component
export interface IProjectStepsState {    
    steps:IStepState[];
    loading:boolean;
    error:string;    
}

//Fabric component
// export interface IProjectStepsState {
//     filteredSteps?:IStepState[];
//     allSteps:IStepState[];
//     loading:boolean;
//     error:string;
//     filterText:string;
// }

export interface IStepState {
    id: number;
    title:string;
    grouping:string;
    startDate:string;
    endDate:string;
    effort:number;
}
