import { ILocation } from './SimpleWeatherWebPart';

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
}