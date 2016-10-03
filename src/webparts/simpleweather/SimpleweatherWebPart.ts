import {
  BaseClientSideWebPart,
  IPropertyPaneSettings,
  IWebPartContext,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneDropdown,
  PropertyPaneChoiceGroup
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

export interface IUserProperty {
  Key: string;
  Value: string;
}

export default class SimpleweatherWebPart extends BaseClientSideWebPart<ISimpleweatherWebPartProps> {
  private container: JQuery;
  private _userLocation: IUserProperty;

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

    //Get location from textbox
    if(location === "None"){
      location = this.properties.location;
      this.showWeather(location);
    }
    //Get location from user profile
    else if(location === "UserLocation"){
      this.fetchLocationFromProfile().then((data) => {
        location = data.Value;
        if (!location || location.length === 0) {
          this.container.html('<p>Please specify a location in the user profile field</p>');
          return;
        }
        this.showWeather(location);
      });
    }
    //Pick one of the locations from the dropdown
    else{
      this.showWeather(location);
    }
  }

  private showWeather(location: string): void{
    if (!location || location.length === 0) {
      this.container.html(`<p class="ms-font-l ${styles.black}">Please specify a location</p>`);
      return;
    }

    const webPart: SimpleweatherWebPart = this;

    ($ as any).simpleWeather({
      location: location,
      woeid: '',
      unit: 'c',
      success: (weather: any): void => {
        var html: string = `<h2><i class="icon${weather.code}"></i> ${weather.temp}&deg;${weather.units.temp}</h2>`;
        html += `<ul><li><p>${weather.city}, ${weather.country}</p></li>`;
        html += `<li class="currently"><p>${weather.currently}</p></li>`;
        html += `<li><p>${weather.wind.direction} ${weather.wind.speed} ${weather.units.speed}</p></li></ul>`;
        html += `<ul>`;
        html += `<p class="ms-font-l ${styles.black}">Forecast</p>`;
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

  //Get location from the 'Office' property of the current user's profile
  private fetchLocationFromProfile(): Promise<IUserProperty>{
     if (this.context.environment.type === EnvironmentType.Local) {
        return this.fetchMockUserLocation().then((response) => {
          return response;
        });
    }
    else{
      return pnp.sp.profiles.myProperties.get().then((response) => {
          var userLocation: IUserProperty = {Key: "", Value: ""};
          var allUserProps: IUserProperty[] = response.UserProfileProperties;
          //Find if there is a efficient way
          allUserProps.forEach((userProp: IUserProperty) => {
              if(userProp.Key === 'Office'){
                console.log("Found property with key = " + userProp.Key);
                userLocation = {Key: userProp.Key, Value: userProp.Value}
              }
          });
          return userLocation;
      });
    }
  }

  //Get dummy location
  private fetchMockUserLocation(): Promise<IUserProperty> {
    return MockHttpClient.getUserLocation(this.context.pageContext.web.absoluteUrl)
            .then((data: IUserProperty) => {
                 var locationData: IUserProperty = { Key: data.Key, Value:data.Value };
                 return locationData;
             }) as Promise<IUserProperty>;
  }

  //Get options for the location - either mock or from the location list dropdown property
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
    options.push( { key: "UserLocation", text: "Choose user location" });
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

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }
}
