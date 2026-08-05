import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

const ANNOUNCEMENTS_LIST_TITLE = 'Duyurular';
const EVENTS_LIST_TITLE = 'Etkinlikler';
// Katılış/Ayrılış listeleri dashboard'un kendi sitesinde değil, Teams'teki
// ilgili takımların sitelerinde yaşıyor — bkz. OnboardingService.ts.
const KATILIS_SITE_URL = 'https://yorpas.sharepoint.com/sites/YazilimTeknoloji';
const KATILIS_LIST_TITLE = 'Çalışan katılışı';
const AYRILIS_SITE_URL = 'https://yorpas.sharepoint.com/sites/Turquality-BilgiTeknolojileri';
const AYRILIS_LIST_TITLE = 'Ayrılışlar';

export type NotificationCategory = 'announcements' | 'events' | 'katilis' | 'ayrilis';

export const NOTIFICATION_CATEGORY_ORDER: NotificationCategory[] = ['announcements', 'events', 'katilis', 'ayrilis'];

export const NOTIFICATION_CATEGORY_LABELS: Record<NotificationCategory, string> = {
    announcements: 'Yeni duyuru var',
    events: 'Yeni etkinlik var',
    katilis: 'Yeni katılış kaydı var',
    ayrilis: 'Yeni ayrılış kaydı var'
};

export const NOTIFICATION_CATEGORY_ICONS: Record<NotificationCategory, string> = {
    announcements: 'Megaphone',
    events: 'Calendar',
    katilis: 'AddFriend',
    ayrilis: 'UserRemove'
};

const CATEGORY_LIST: Record<NotificationCategory, { siteUrl: string; listTitle: string }> = {
    announcements: { siteUrl: '', listTitle: ANNOUNCEMENTS_LIST_TITLE },
    events: { siteUrl: '', listTitle: EVENTS_LIST_TITLE },
    katilis: { siteUrl: KATILIS_SITE_URL, listTitle: KATILIS_LIST_TITLE },
    ayrilis: { siteUrl: AYRILIS_SITE_URL, listTitle: AYRILIS_LIST_TITLE }
};

export type LatestIds = Record<NotificationCategory, number>;

/**
 * "Yenilik" tespiti, tarih/saat karşılaştırması yerine listenin en büyük
 * Id'sine bakarak yapılıyor — SharePoint Id'leri her zaman artan sırada
 * olduğu için bu, saat dilimi/senkronizasyon riski taşımayan basit ve
 * güvenilir bir yöntem. Her liste için tek satırlık, hafif bir istek.
 */
const getLatestId = async (context: WebPartContext, siteUrl: string, listTitle: string): Promise<number> => {
    const effectiveSiteUrl = siteUrl || context.pageContext.web.absoluteUrl;
    const endpoint = `${effectiveSiteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items?$top=1&$orderby=Id desc&$select=Id`;
    try {
        const response: SPHttpClientResponse = await context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
        if (!response.ok) {
            return 0;
        }
        const body: { value: { Id: number }[] } = await response.json();
        return body.value?.[0]?.Id ?? 0;
    } catch (error) {
        console.error(`[NotificationService] "${listTitle}" için son kayıt alınamadı:`, error);
        return 0;
    }
};

export const getLatestIds = async (context: WebPartContext): Promise<LatestIds> => {
    const entries = await Promise.all(
        NOTIFICATION_CATEGORY_ORDER.map(async (category) => {
            const { siteUrl, listTitle } = CATEGORY_LIST[category];
            const id = await getLatestId(context, siteUrl, listTitle);
            return [category, id] as const;
        })
    );
    return entries.reduce((acc, [category, id]) => {
        acc[category] = id;
        return acc;
    }, {} as LatestIds);
};

const storageKey = (loginName: string): string => `yorpas-dashboard-notif-seen-${loginName}`;

/** Kullanıcı bazlı — aynı tarayıcıyı paylaşan farklı hesaplar birbirinin "görüldü" durumunu ezmesin diye. */
export const getSeenIds = (loginName: string): LatestIds | undefined => {
    try {
        const raw = window.localStorage.getItem(storageKey(loginName));
        return raw ? (JSON.parse(raw) as LatestIds) : undefined;
    } catch {
        return undefined;
    }
};

export const setSeenIds = (loginName: string, ids: LatestIds): void => {
    try {
        window.localStorage.setItem(storageKey(loginName), JSON.stringify(ids));
    } catch {
        /* localStorage kullanılamıyorsa (ör. gizli sekme kısıtlaması) sessizce yok say — bildirim kritik değil. */
    }
};
