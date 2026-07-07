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

/* ------------------------------------------------------------------ *
 *  DİKKAT: Bu interface web part'ın property pane (ayarlar paneli)
 *  alanlarını tanımlar — Dashboard bileşeninin props'u DEĞİLDİR.
 *  İkisini karıştırmak "Property 'description' does not exist on type
 *  'IDashboardProps'" hatasının tam sebebiydi.
 * ------------------------------------------------------------------ */
export interface IDashboardWebPartProps {
    description: string;
}

export default class DashboardWebPart extends BaseClientSideWebPart<IDashboardWebPartProps> {

    public render(): void {
        const element: React.ReactElement<IDashboardProps> = React.createElement(
            Dashboard,
            {
                userDisplayName: this.context.pageContext.user.displayName,
                // Graph ve SharePoint REST çağrıları (toplantılar, rehber,
                // duyurular, belgeler) bu context olmadan çalışmaz.
                context: this.context
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
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    }
}