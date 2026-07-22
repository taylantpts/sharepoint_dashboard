import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
    BaseClientSideWebPart,
    IPropertyPaneConfiguration,
    PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import Dashboard from './components/Dashboard';
import { IDashboardProps } from './Dashboard.types';
// Yerel/gitignore'lu yedek anahtar — bkz. weatherApiKey.local.example.ts.
// Property Pane'den bir değer girilmişse HER ZAMAN o kullanılır; bu sadece
// panel boşken (ör. yerel workbench testinde) devreye giren bir kolaylıktır.
import { WEATHER_API_KEY_LOCAL_FALLBACK } from './weatherApiKey.local';

/* ------------------------------------------------------------------ *
 *  DİKKAT: Bu interface web part'ın property pane (ayarlar paneli)
 *  alanlarını tanımlar — Dashboard bileşeninin props'u DEĞİLDİR.
 *  İkisini karıştırmak "Property 'description' does not exist on type
 *  'IDashboardProps'" hatasının tam sebebiydi.
 * ------------------------------------------------------------------ */
export interface IDashboardWebPartProps {
    description: string;
    // OpenWeatherMap API anahtarı — KASITLI OLARAK kaynak koduna gömülmüyor
    // (herkese açık JS paketinde herkes tarafından okunabilir olurdu).
    // Bunun yerine site yöneticisi bu değeri Property Pane'den (web part
    // düzenleme paneli) girer; sadece bu web part'ın property bag'inde saklanır.
    // openweathermap.org anahtar panelinden HTTP referrer kısıtlaması da
    // eklenmesi önerilir (ek savunma katmanı).
    weatherApiKey: string;
}

export default class DashboardWebPart extends BaseClientSideWebPart<IDashboardWebPartProps> {

    public render(): void {
        const element: React.ReactElement<IDashboardProps> = React.createElement(
            Dashboard,
            {
                userDisplayName: this.context.pageContext.user.displayName,
                // Graph ve SharePoint REST çağrıları (toplantılar, rehber,
                // duyurular, belgeler) bu context olmadan çalışmaz.
                context: this.context,
                weatherApiKey: this.properties.weatherApiKey || WEATHER_API_KEY_LOCAL_FALLBACK
            }
        );

        ReactDom.render(element, this.domElement);
    }

    protected onDispose(): void {
        ReactDom.unmountComponentAtNode(this.domElement);
    }

    protected get dataVersion(): Version {
        return Version.parse('1.0');
    }

    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
        return {
            pages: [
                {
                    header: { description: 'Dashboard Ayarları' },
                    groups: [
                        {
                            groupName: 'Genel',
                            groupFields: [
                                PropertyPaneTextField('description', {
                                    label: 'Açıklama'
                                }),
                                PropertyPaneTextField('weatherApiKey', {
                                    label: 'OpenWeatherMap API Anahtarı',
                                    description: 'openweathermap.org üzerinden alınan ücretsiz API anahtarı. ' +
                                        'Hava Durumu widget\'ının çalışması için gereklidir.'
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    }
}