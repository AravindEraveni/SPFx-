export interface IStates{
    Title:string,
    Age:number,
    isActive:boolean,
    Dept:string,
    Date:Date,
    Sal:number,
    ID:number

}
export interface IFinalStates{
    emps:IStates[],
    Title:string,
    Age:number,
    isActive:boolean,
    Dept:string,
    DeptOptions?:string[]
    Date?:Date,
    Sal:number,
    showForm:boolean,
    submitMsg:string,
    emptyData:string,
    dltMsg:string,
    newEdit:string,
    editId:string,
    searchText: string,
    
}