import * as React from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { HR_ADMIN_GROUP_NAME, ISG_GROUP_NAME } from '../constants';

export interface IPermissions {
    /** Duyurular ve Etkinlikler listelerine "+" ile içerik ekleyebilir mi. */
    canManageAnnouncements: boolean;
    /** İSG Takvimi'ne yeni dosya yükleyebilir mi. */
    canManageISGCalendar: boolean;
}

const DEFAULT_PERMISSIONS: IPermissions = {
    canManageAnnouncements: false,
    canManageISGCalendar: false
};

/**
 * Oturum açan kullanıcının SharePoint gruplarını (`_api/web/currentuser/groups`)
 * bir kez çekip widget "+" butonlarının görünürlüğünü buna göre belirler.
 * ÖNEMLİ: bu hook sadece arayüzü gizler — gerçek yetkilendirme, ilgili
 * listelerde (Duyurular/Etkinlikler/ISGTakvimi) kırılan izin kalıtımı ve
 * bu gruplara verilen liste-bazlı Katılım (Contribute) izniyle sağlanır.
 * Grup bulunamazsa (istek hatası, grup üyeliği yok) güvenli taraf seçilir:
 * buton hiç gösterilmez.
 */
export const usePermissions = (context: WebPartContext): IPermissions => {
    const [permissions, setPermissions] = React.useState<IPermissions>(DEFAULT_PERMISSIONS);

    React.useEffect(() => {
        let isMounted = true;

        const load = async (): Promise<void> => {
            try {
                const webUrl = context.pageContext.web.absoluteUrl;
                const endpoint = `${webUrl}/_api/web/currentuser/groups?$select=Title`;
                const response: SPHttpClientResponse = await context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);

                if (!response.ok) {
                    console.error(`[usePermissions] Kullanıcı grupları okunamadı (HTTP ${response.status})`);
                    return;
                }

                const body: { value: { Title: string }[] } = await response.json();
                const groupTitles = (body.value ?? []).map((g) => g.Title);

                if (isMounted) {
                    setPermissions({
                        canManageAnnouncements: groupTitles.indexOf(HR_ADMIN_GROUP_NAME) !== -1,
                        canManageISGCalendar: groupTitles.indexOf(ISG_GROUP_NAME) !== -1
                    });
                }
            } catch (error) {
                console.error('[usePermissions] Kullanıcı grupları çekilirken beklenmeyen hata:', error);
            }
        };

        load().catch(() => { /* hata zaten yukarıda loglandı */ });

        return () => {
            isMounted = false;
        };
    }, [context]);

    return permissions;
};
