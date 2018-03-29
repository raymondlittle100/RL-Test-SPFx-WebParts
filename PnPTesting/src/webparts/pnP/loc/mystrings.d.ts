declare interface IPnPWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
}

declare module 'PnPWebPartStrings' {
  const strings: IPnPWebPartStrings;
  export = strings;
}
