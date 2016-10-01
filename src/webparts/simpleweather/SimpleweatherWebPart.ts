import {
  BaseClientSideWebPart,
  IPropertyPaneSettings,
  IWebPartContext,
  PropertyPaneTextField,
  PropertyPaneSlider
} from '@microsoft/sp-client-preview';

import styles from './Simpleweather.module.scss';
import * as strings from 'simpleweatherStrings';
import { ISimpleweatherWebPartProps } from './ISimpleweatherWebPartProps';

import * as $ from 'jquery';
require('simpleWeather');

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

    const location: string = this.properties.location;

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

        html += `<ul>`

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
                PropertyPaneTextField('location', {
                  label: strings.LocationFieldLabel
                }),
                PropertyPaneSlider('numberOfDays', {
                  label: strings.NumberOfDaysFieldLabel,
                  min: 1,
                  max: 10,
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
