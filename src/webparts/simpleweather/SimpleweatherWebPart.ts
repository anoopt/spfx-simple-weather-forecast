import {
  BaseClientSideWebPart,
  IPropertyPaneSettings,
  IWebPartContext,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneDropdown,
  PropertyPaneChoiceGroup,
  PropertyPaneToggle
} from '@microsoft/sp-client-preview';
import { PropertyFieldColorPicker } from 'sp-client-custom-fields/lib/PropertyFieldColorPicker';

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
    this.onPropertyChange = this.onPropertyChange.bind(this);
  }

  public render(): void {
    if (this.renderedOnce === false) {
      this.domElement.innerHTML = `<div class="${styles.simpleweather}"></div>`;
    }
    this.renderContents();
  }

  private renderContents(): void {
    this.container = $(`.${styles.simpleweather}`, this.domElement);
    var selectedOptionChoice: string = this.properties.locationOptionChoice;

    var location: string = this.properties.location;

    //Get location from textbox
    if(selectedOptionChoice === "None"){
      location = this.properties.location;
      this.showWeather(location);
    }
    //Get location from user profile
    else if(selectedOptionChoice === "UserLocation"){
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
      location = this.properties.locationDropdown;
      this.showWeather(location);
    }
  }

  private showWeather(location: string): void{
    if (!location || location.length === 0) {
      this.container.html(`<p class="ms-font-l" style="color:${this.properties.fontColour}">Please specify a location</p>`);
      return;
    }
    var topText: string = this.properties.webpartTopText.replace("%location%",location);
    var html: string = `<p class="ms-font-l" style="color:${this.properties.fontColour}">${topText}</p>`;
    const webPart: SimpleweatherWebPart = this;

    ($ as any).simpleWeather({
      location: location,
      woeid: '',
      unit: 'c',
      success: (weather: any): void => {
        html += `<h2 style="color:${this.properties.fontColour};background:${this.properties.textBgColour}"><i class="icon${weather.code}" style="color:${this.properties.fontColour}"></i> ${weather.temp}&deg;${weather.units.temp}</h2>`;
        html += `<ul>`; //<li style="background:${this.properties.textBgColour}"><p style="color:${this.properties.fontColour}">${weather.city}, ${weather.country}</p></li>
        html += `<li class="currently" style="background:${this.properties.textBgColour}"><p style="color:${this.properties.fontColour}">${weather.currently}</p></li>`;
        html += `<li style="background:${this.properties.textBgColour}"><p style="color:${this.properties.fontColour}">${weather.wind.direction} ${weather.wind.speed} ${weather.units.speed}</p></li></ul>`;
        html += `<ul>`;
        html += `<p class="ms-font-l ${styles.forecast}" style="color:${this.properties.fontColour}">Forecast</p>`;
        for(var i=0;i<this.properties.numberOfDays;i++) {
          html += `<li style="background:${this.properties.textBgColour}"><div class="ms-FacePile"><i class="icon${weather.forecast[i+1].code}" style="color:${this.properties.fontColour}"></i>
                  <p style="color:${this.properties.fontColour}">${weather.forecast[i+1].day}: ${weather.forecast[i].high}&deg;C </p></div></li>`;
        }
        html += `</ul>`;
        webPart.container.html(html)
        .removeAttr('style')
        .css('background',this.properties.backgroundColour);
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
    //options.push( { key: "None", text: "Specify in the text box" });
    //options.push( { key: "UserLocation", text: "Choose user location" });
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

    let templateProperty: any;
    if (this.properties.locationOptionChoice == "None") {
      templateProperty = PropertyPaneTextField('location', {
                  label: strings.LocationFieldLabel
                });
    }
    else if (this.properties.locationOptionChoice == "PreConfig"){
      templateProperty = PropertyPaneDropdown('locationDropdown', {
                  label: 'Select a location',
                  isDisabled: false,
                  options: this._locations
                });
    }
    else {
      templateProperty = PropertyPaneToggle('userLocationSelected', {
                label: 'Pick location form user\'s profile',
                disabled: true,
                checked: true
              });
    }
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          displayGroupsAsAccordion: true,
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('webpartTopText', {
                  label: "Top text"
                }),
                PropertyPaneChoiceGroup('locationOptionChoice', {
                  label: 'Choose one of the following',
                  options: [
                    { key: 'None', text: 'Specify a location' },
                    { key: 'UserLocation', text: 'Pick location from user\'s profile' },
                    { key: 'PreConfig', text: 'Select one from pre-configured locations' }
                  ]
                }),
                templateProperty,
                PropertyPaneSlider('numberOfDays', {
                  label: strings.NumberOfDaysFieldLabel,
                  min: 1,
                  max: 5,
                  step: 1
                })
              ]
            },
            {
              groupName: "Colours",
              groupFields: [
                PropertyFieldColorPicker('textBgColour', {
                  label: "Text background colour",
                  initialColor: this.properties.textBgColour,
                  onPropertyChange: this.onPropertyChange
                }),
                PropertyFieldColorPicker('fontColour', {
                  label: "Font colour",
                  initialColor: this.properties.fontColour,
                  onPropertyChange: this.onPropertyChange
                }),
                PropertyFieldColorPicker('backgroundColour', {
                  label: "Background colour",
                  initialColor: this.properties.backgroundColour,
                  onPropertyChange: this.onPropertyChange
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
