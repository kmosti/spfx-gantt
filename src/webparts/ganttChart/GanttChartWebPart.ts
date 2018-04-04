import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-webpart-base';

import * as strings from 'GanttChartWebPartStrings';
import GanttChart from './components/GanttChart';
import { IGanttChartProps } from './components/IGanttChartProps';
import pnp from "sp-pnp-js";
import * as moment from 'moment';
import { SPComponentLoader } from '@microsoft/sp-loader';

export interface IGanttChartWebPartProps {
  description: string;
  listTitle: string;
  zoom: string;
}

export default class GanttChartWebPart extends BaseClientSideWebPart<IGanttChartWebPartProps> {

  public onInit(): Promise<void> {
    SPComponentLoader.loadCss('https://ateastavanger.azureedge.net/dhtmlx/codebase/dhtmlxcombo.css');

    // Init the moment JS library locale globally
    const currentLocale = this.context.pageContext.cultureInfo.currentCultureName;
    moment.locale(currentLocale);

    return super.onInit().then(_ => {

      pnp.setup({
        spfxContext: this.context
      });

      pnp.sp.web.lists.filter("BaseTemplate eq 171").select("Title").get().then( lists => {
        //console.dir(lists);
        this._dropdownOptions = lists.map( list => {
          return {
            key: list.Title,
            text: list.Title
          }
        });
      });

    });
  }

  public render(): void {
    const element: React.ReactElement<IGanttChartProps > = React.createElement(
      GanttChart,
      {
        description: this.properties.description,
        context: this.context,
        zoom: this.properties.zoom,
        listTitle:  this.properties.listTitle
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneDropdown('listTitle', {
                  label: 'List Title',
                  options: this._dropdownOptions
                }),
                PropertyPaneDropdown('zoom', {
                  label: 'Default zoom',
                  options: this._zoomOptions,
                  selectedKey: "Days"
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private _dropdownOptions: IPropertyPaneDropdownOption[] = [];
  private _zoomOptions: IPropertyPaneDropdownOption[] = [
    {
      key: "Hours",
      text: "Hours"
    },
    {
      key: "Days",
      text: "Days"
    },
    {
      key: "Months",
      text: "Months"
    }
  ];
}
