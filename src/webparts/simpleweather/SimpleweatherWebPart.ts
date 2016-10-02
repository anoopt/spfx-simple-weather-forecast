import {
  BaseClientSideWebPart,
  IPropertyPaneSettings,
  IWebPartContext,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneDropdown
} from '@microsoft/sp-client-preview';

import { EnvironmentType } from '@microsoft/sp-client-base';
import styles from './Simpleweather.module.scss';
import * as strings from 'simpleweatherStrings';
import { ISimpleweatherWebPartProps } from './ISimpleweatherWebPartProps';
import MockHttpClient from './MockHttpClient';

import * as $ from 'jquery';
require('simpleWeather');
import * as pnp from 'sp-pnp-js';

interface IPropertyPaneDropdownOption{
  key: string;
  text: string;
}

export interface ILocation {
  Title: string;
}

interface ILocations {
  value: ILocation[];
}

export default class SimpleweatherWebPart extends BaseClientSideWebPart<ISimpleweatherWebPartProps> {
  private container: JQuery;
  public constructor(context: IWebPartContext) {
    super(context);
  }

  public render(): void {
    if (this.renderedOnce === false) {
      this.domElement.innerHTML = `<div class="${styles.simpleweather}"></div>`;
    }

    this.renderContents();
  }

  private renderContents(): void {
    this.container = $(`.${styles.simpleweather}`, this.domElement);

    var location: string = this.properties.locationDropdown;

    if(this.properties.locationDropdown === "None"){
      location = this.properties.location;
    }

    if (!location || location.length === 0) {
      this.container.html('<p>Please specify a location</p>');
      return;
    }

    const webPart: SimpleweatherWebPart = this;

    ($ as any).simpleWeather({
      location: location,
      woeid: '',
      unit: 'c',
      success: (weather: any): void => {
        var html: string = `<h2><i class="icon${weather.code}"></i> ${weather.temp}&deg;${weather.units.temp}</h2>`;
        html += `<ul><li><p>${weather.city}, ${weather.region}</p></li>`;
        html += `<li class="currently"><p>${weather.currently}</p></li></ul><br/>`;
        html += `<ul>`;

        for(var i=0;i<this.properties.numberOfDays;i++) {

          html += `<li><div class="ms-FacePile"><i class="icon${weather.forecast[i+1].code}"></i>
                  <p>${weather.forecast[i+1].day}: ${weather.forecast[i].high}&deg;C </p></div></li>`;
        }

        html += `</ul>`;

        webPart.container.html(html)
        .removeAttr('style')
        .css('background','#e1e1e1');

      },
      error: (error: any): void => {
        webPart.container.html(`<p>${error.message}</p>`).removeAttr('style');
      }
    });
  }

  private _locations: IPropertyPaneDropdownOption[] = [];

  public onInit<T>(): Promise<T> {
    this.fetchOptions().then((data) => {
        this._locations = data;
    });

    return Promise.resolve();
  }

  private fetchOptions(): Promise<IPropertyPaneDropdownOption[]> {

    if (this.context.environment.type === EnvironmentType.Local) {
        return this.fetchMockLocations().then((response) => {
          return this.fetchOptionsFromResponse(response.value);
        });
    }
    else {
      //var url = this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('Location')/items`;
      //return this.fetchLocations(url).then((response) => {
         // return this.fetchOptionsFromResponse(response.value);
      //});
      /***********************************/
      /* OR using PnP JS*/
      /***********************************/
      return pnp.sp.web.lists.getByTitle('Location')
      .items.select('Title')
      .get().then((response) => {
          return this.fetchOptionsFromResponse(response);
      });
    }
  }

  private fetchOptionsFromResponse(locations: ILocation[]): IPropertyPaneDropdownOption[]{
    var options: Array<IPropertyPaneDropdownOption> = new Array<IPropertyPaneDropdownOption>();
    options.push( { key: "None", text: "Specify in the text box" });
    locations.forEach((location: ILocation) => {
              console.log("Found location with title = " + location.Title);
              options.push( { key: location.Title, text: location.Title });
          });
    return options;
  }

  private fetchLocations(url: string) : Promise<ILocations> {
    return this.context.httpClient.get(url).then((response: Response) => {
        if (response.ok) {
          return response.json();
        } else {
          console.log("WARNING - failed to hit URL " + url + ". Error = " + response.statusText);
          return null;
        }
      });
  }

  private fetchMockLocations(): Promise<ILocations> {
    return MockHttpClient.get(this.context.pageContext.web.absoluteUrl)
            .then((data: ILocation[]) => {
                 var locationData: ILocations = { value: data };
                 return locationData;
             }) as Promise<ILocations>;
  }



  protected get propertyPaneSettings(): IPropertyPaneSettings {
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
                PropertyPaneDropdown('locationDropdown', {
                  label: 'Select a location',
                  isDisabled: false,
                  options: this._locations
                }),
                 PropertyPaneTextField('location', {
                  label: strings.LocationFieldLabel
                }),
                PropertyPaneSlider('numberOfDays', {
                  label: strings.NumberOfDaysFieldLabel,
                  min: 1,
                  max: 5,
                  step: 1
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
