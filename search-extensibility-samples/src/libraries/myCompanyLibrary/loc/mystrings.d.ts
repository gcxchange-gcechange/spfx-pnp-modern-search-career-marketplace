declare interface IMyCompanyLibraryLibraryStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  CustomQueryModifier: {
    GroupName:string;
    PrefixLabel:string;
    PrefixDescription:string;
    PrefixPlaceholder:string;
    SuffixLabel:string;
    SuffixDescription:string;
    SuffixPlaceholder:string;
  }
  classificationLevel: string;
  opportunityType: string;
  duration: string;
  location: string;
  deadline: string;
  view: string;
  apply: string;
  results: string;
  resultsFor: string;
}

declare module 'MyCompanyLibraryLibraryStrings' {
  const strings: IMyCompanyLibraryLibraryStrings;
  export = strings;
}
