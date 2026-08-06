import * as React from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { HR_ADMIN_GROUP_NAME, ISG_GROUP_NAME, IT_GROUP_NAME, ACCOUNTING_GROUP_NAME } from '../constants';

export interface IPermissions {
    /** Duyurular ve Etkinlikler listelerine "+" ile içerik ekleyebilir mi. */
    canManageAnnouncements: boolean;
    /** İSG Takvimi'ne yeni dosya yükleyebilir mi. */
    canManageISGCalendar: boolean;
    /** Katılış/Ayrılış takibinde "+" (yeni kayıt) açabilir mi — SADECE İK. */
    canManageOnboarding: boolean;
    /** Katılış/Ayrılış kayıtlarındaki kalem/düzenle ikonunu görebilir mi — İK, BT veya Muhasebe. */
    canEditOnboarding: boolean;
    /**
     * Katılış & Ayrılış Takibi widget'ının KENDİSİNİ (tüm tabloyu, tüm
     * çalışan kayıtlarını) görebilir mi — İK, BT veya Muhasebe DIŞINDAKİ
     * hiçbir kullanıcı bu widget'ı hiç görmemeli (önceden sadece "+" ve
     * kalem/düzenle ikonları gizleniyordu, tablo herkese açıktı).
     */
    canViewOnboarding: boolean;
}

const DEFAULT_PERMISSIONS: IPermissions = {
    canManageAnnouncements: false,
    canManageISGCalendar: false,
    canManageOnboarding: false,
    canEditOnboarding: false,
    canViewOnboarding: false
};

/**
 * Oturum açan kullanıcının SharePoint gruplarını (`_api/web/currentuser/groups`)
 * bir kez çekip widget "+" butonlarının görünürlüğünü buna göre belirler.
 * ÖNEMLİ: bu hook sadece arayüzü gizler — gerçek yetkilendirme, ilgili
 * listelerde (Duyurular/Etkinlikler/ISGTakvimi) kırılan izin kalıtımı ve
 * bu gruplara verilen liste-bazlı Katılım (Contribute) izniyle sağlanır.
 * Grup bulunamazsa (istek hatası, grup üyeliği yok) güvenli taraf seçilir:
 * buton hiç gösterilmez.
 *
 * NOT (Katılış/Ayrılış için önemli): grup üyeliği her zaman bu widget'ın
 * BARINDIĞI portal sitesinden (context.pageContext.web) okunur — Katılış/
 * Ayrılış listelerinin GERÇEKTE yaşadığı uzak sitelerden (YazilimTeknoloji /
 * Turquality-BilgiTeknolojileri) DEĞİL. Yani "BT Personeli" ve "Muhasebe
 * Personeli" gruplarının da (İK ve İSG grupları gibi) bu PORTAL sitesinde
 * oluşturulmuş olması gerekir — buton görünürlüğü buradan belirlenir, uzak
 * listelerdeki GERÇEK yazma izni ise o sitelerin kendi izin yapısına bağlıdır.
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
                const isHr = groupTitles.indexOf(HR_ADMIN_GROUP_NAME) !== -1;
                const isIt = groupTitles.indexOf(IT_GROUP_NAME) !== -1;
                const isAccounting = groupTitles.indexOf(ACCOUNTING_GROUP_NAME) !== -1;

                if (isMounted) {
                    const canSeeOnboarding = isHr || isIt || isAccounting;
                    setPermissions({
                        canManageAnnouncements: isHr,
                        canManageISGCalendar: groupTitles.indexOf(ISG_GROUP_NAME) !== -1,
                        canManageOnboarding: isHr,
                        canEditOnboarding: canSeeOnboarding,
                        canViewOnboarding: canSeeOnboarding
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
