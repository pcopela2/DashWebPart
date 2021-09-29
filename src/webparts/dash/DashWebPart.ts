import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Environment, Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'dashStrings';
import Dash from './components/Dash';
import { IDashProps } from './components/IDashProps';
import { SPHttpClient, SPHttpClientResponse, SPHttpClientConfiguration } from '@microsoft/sp-http'; 
import SharePointService from '../../services/SharePoint/SharePointServices';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';


export interface IDashWebPartProps{
  listId: string;
  selectedFields: string;
  chartType: string;
  chartTitle: string;
  color1: string;
  color2: string;
  color3: string;
}

export default class DashWebPart extends BaseClientSideWebPart<IDashWebPartProps> {

//Field Options state for dropdown
private fieldOptions: IPropertyPaneDropdownOption[];
private fieldOptionsLoading: boolean = false;


  public render(): void {
    const element: React.ReactElement<IDashProps > = React.createElement(
      Dash,
      {
        listId: this.properties.listId,
        selectedFields: this.properties.selectedFields.split(','),
        chartType: this.properties.chartType,
        chartTitle: this.properties.chartTitle,
        colors: [
          this.properties.color1,
          this.properties.color2,
          this.properties.color3,
        ],
      }
    );

    ReactDom.render(element, this.domElement);
  }

  public onInit(): Promise<void> {
    return super.onInit().then(() => {
      SharePointService.setup(this.context, Environment.type);
    });
  }

  public onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Dash Settings'
          },
          groups: [
            {
              groupName: 'Chart Data',
              groupFields: [
                PropertyFieldListPicker('listId', {
                  label: 'Select a list',
                  selectedList: this.properties.listId,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'listPickerFieldId'
                }),
                PropertyPaneTextField('selectedFields', {
                  label: 'Selected Fields',
                }),
              ]
            },
            {
              groupName: 'Chart Settings',
              groupFields: [
              PropertyPaneDropdown('chartType', {
                label: 'Chart Type',
                options: [
                  {key: 'Bar', text: 'Bar'},
                  {key: 'HorizontalBar', text: 'HorizontalBar'},
                  {key: 'Line', text: 'Line'},
                  {key: 'Pie', text: 'Pie'},
                  {key: 'Doughnut', text: 'Doughnut'},
                  {key: 'Radar', text: 'Radar'},
                  {key: 'Polar', text: 'Polar'},
                ],
              }),
              PropertyPaneTextField('chartTitle', {
                label: 'Chart Title'
              }),
            ],
            },
            {
              groupName: 'Chart Style',
              groupFields: [
                PropertyPaneDropdown('color1', {
                  label: 'Color 1',
                  selectedKey: this.properties.color1,
                  options: [
                    {key: '#0078d4', text: 'Blue'},
                    {key: '#e50000', text: 'Red'},
                    {key: '#bad80a', text: 'Yellow'},
                    {key: '#00b294', text: 'Green'},
                    {key: '#5c2d91', text: 'Purple'},
                  ],
                }),
                PropertyPaneDropdown('color2', {
                  label: 'Color 2',
                  options: [
                    {key: '#0078d4', text: 'Blue'},
                    {key: '#e50000', text: 'Red'},
                    {key: '#bad80a', text: 'Yellow'},
                    {key: '#00b294', text: 'Green'},
                    {key: '#5c2d91', text: 'Purple'},
                  ],
                }),
                PropertyPaneDropdown('color3', {
                  label: 'Color 3',
                  options: [
                    {key: '#0078d4', text: 'Blue'},
                    {key: '#e50000', text: 'Red'},
                    {key: '#bad80a', text: 'Yellow'},
                    {key: '#00b294', text: 'Green'},
                    {key: '#5c2d91', text: 'Purple'},
                  ],
                }),
              ],
            }
          ]
        }
      ]
    };
  }
}