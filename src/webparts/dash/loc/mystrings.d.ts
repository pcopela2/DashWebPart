declare interface IDashStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
}

declare module 'dashStrings' {
  const strings: IDashStrings;
  export = strings;
}
