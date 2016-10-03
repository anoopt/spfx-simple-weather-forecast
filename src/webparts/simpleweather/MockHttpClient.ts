import { ILocation } from './SimpleWeatherWebPart';
import { IUserProperty } from './SimpleWeatherWebPart'

export default class MockHttpClient {

  private static _items: ILocation[] = [
    { Title: 'London' },
    { Title: 'Bangalore' },
    { Title: 'Perth' }];

  public static get(restUrl: string, options?: any): Promise<ILocation[]> {
    return new Promise<ILocation[]>((resolve) => {
      resolve(MockHttpClient._items);
    });
  }

  private static _userLocation: IUserProperty = { Key:"Office", Value:"Belgaum" };

  public static getUserLocation(restUrl: string, options?: any): Promise<IUserProperty> {
    return new Promise<IUserProperty>((resolve) => {
      resolve(MockHttpClient._userLocation);
    });
  }
}