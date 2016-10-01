declare interface ISimpleweatherStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  LocationFieldLabel: string;
  NumberOfDaysFieldLabel: string;
}

declare module 'simpleweatherStrings' {
  const strings: ISimpleweatherStrings;
  export = strings;
}
