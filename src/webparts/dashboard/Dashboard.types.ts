import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IDashboardProps {
  userDisplayName: string;
  // MS Graph (toplantılar, kullanıcılar) ve SharePoint REST (Duyurular,
  // Belgeler listeleri) çağrıları için gereken context. Web part'ınızın
  // render metodunda <Dashboard context={this.context} ... /> şeklinde geçin.
  context: WebPartContext;
  // Property Pane'den girilen OpenWeatherMap API anahtarı — bkz.
  // DashboardWebPart.ts içindeki IDashboardWebPartProps.weatherApiKey yorumu.
  weatherApiKey: string;
}