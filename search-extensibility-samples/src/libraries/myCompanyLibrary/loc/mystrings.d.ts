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
  MyOpportunitiesQueryModifier: {
    GroupName:string;
  },
  classificationLevel: string;
  opportunityType: string;
  duration: string;
  description: string;
  location: string;
  deadline: string;
  view: string;
  apply: string;
  results: string;
  resultsFor: string;
  viewAria: string;
  applyAria: string;
}

declare module 'MyCompanyLibraryLibraryStrings' {
  const strings: IMyCompanyLibraryLibraryStrings;
  export = strings;
}
