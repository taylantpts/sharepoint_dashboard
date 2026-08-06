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
// bkz. SharePointService.ts IKINCI_EL_LIST_TITLE — aynı liste adı, iki
// dosyada da sabit olarak tutuluyor (NotificationService liste adını
// SharePointService'ten import ETMİYOR — döngüsel bağımlılık kurmamak için
// bu projede zaten kurulu desen: her iki dosya da liste adını kendi sabiti
// olarak taşır).
const IKINCI_EL_LIST_TITLE = 'İkinci El İlanlar';

export type NotificationCategory = 'announcements' | 'events' | 'katilis' | 'ayrilis' | 'ikinciEl';

export const NOTIFICATION_CATEGORY_ORDER: NotificationCategory[] =
    ['announcements', 'events', 'katilis', 'ayrilis', 'ikinciEl'];

export const NOTIFICATION_CATEGORY_LABELS: Record<NotificationCategory, string> = {
    announcements: 'Yeni duyuru var',
    events: 'Yeni etkinlik var',
    katilis: 'Yeni katılış kaydı var',
    ayrilis: 'Yeni ayrılış kaydı var',
    ikinciEl: 'Yeni ilan var'
};

export const NOTIFICATION_CATEGORY_ICONS: Record<NotificationCategory, string> = {
    announcements: 'Megaphone',
    events: 'Calendar',
    katilis: 'AddFriend',
    ayrilis: 'UserRemove',
    ikinciEl: 'ShoppingCart'
};

/**
 * Bildirime tıklanınca sayfada kaydırılacak (scroll) hedef — Dashboard.tsx
 * ilgili widget'ı saran div'e bu id'yi veriyor. "katilis"/"ayrilis" kasıtlı
 * olarak "Katılış & Ayrılış Takibi" (yönetim/düzenleme widget'ı) DEĞİL,
 * "Aramıza Katılanlar"/"Aramızdan Ayrılanlar" (salt-okunur son-5 özeti)
 * widget'larına gidiyor — bildirimin kendisi de o listelerdeki YENİ kaydı
 * haber verdiği için en doğrudan/anlamlı hedef bu.
 */
export const NOTIFICATION_CATEGORY_ANCHOR_ID: Record<NotificationCategory, string> = {
    announcements: 'dashboard-anchor-announcements',
    events: 'dashboard-anchor-events',
    katilis: 'dashboard-anchor-katilis',
    ayrilis: 'dashboard-anchor-ayrilis',
    ikinciEl: 'dashboard-anchor-ikinciel'
};

const CATEGORY_LIST: Record<NotificationCategory, { siteUrl: string; listTitle: string }> = {
    announcements: { siteUrl: '', listTitle: ANNOUNCEMENTS_LIST_TITLE },
    events: { siteUrl: '', listTitle: EVENTS_LIST_TITLE },
    katilis: { siteUrl: KATILIS_SITE_URL, listTitle: KATILIS_LIST_TITLE },
    ayrilis: { siteUrl: AYRILIS_SITE_URL, listTitle: AYRILIS_LIST_TITLE },
    ikinciEl: { siteUrl: '', listTitle: IKINCI_EL_LIST_TITLE }
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

/**
 * Kullanıcı bazlı — aynı tarayıcıyı paylaşan farklı hesaplar birbirinin
 * "görüldü" durumunu ezmesin diye.
 * ÖNCEKİ HATA (ileriye dönük): yeni bir NotificationCategory eklendiğinde
 * (ör. "ikinciEl"), tarayıcısında ESKİDEN KAYITLI seenIds objesi bu yeni
 * alanı hiç TAŞIMAZ — `seenIds.ikinciEl` `undefined` olur ve
 * `latestId > undefined` karşılaştırması JS'te her zaman `false` döner, yani
 * o kategori GERÇEKTEN yeni bir kayıt olsa bile asla "yeni" görünmez. Eksik
 * her kategori burada 0 ile dolduruluyor ki yeni kategoriler eski
 * kullanıcılarda da doğru çalışsın.
 */
export const getSeenIds = (loginName: string): LatestIds | undefined => {
    try {
        const raw = window.localStorage.getItem(storageKey(loginName));
        if (!raw) {
            return undefined;
        }
        const parsed = JSON.parse(raw) as Partial<LatestIds>;
        return NOTIFICATION_CATEGORY_ORDER.reduce((acc, category) => {
            acc[category] = parsed[category] ?? 0;
            return acc;
        }, {} as LatestIds);
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

/**
 * Zil menüsünde görünen TEK TEK bildirim kayıtları — ÖNCEKİ MODEL sadece
 * "bu kategoride görülmemiş bir şey var mı" (evet/hayır) tutuyordu, bu
 * yüzden bir bildirime tıklamak (ya da menüyü kapatmak) o kategoriyi
 * anında listeden SİLİYORDU. Kullanıcı bunun yerine kalıcı, en fazla
 * MAX_HISTORY kayıtlık bir GEÇMİŞ istedi: tıklanan bildirim silinmeyip
 * "okundu" görünümüne dönüşecek, sadece 6. yeni bildirim geldiğinde en
 * eskisi (okunma durumu ne olursa olsun) düşecek.
 */
export interface INotificationEntry {
    id: string;
    category: NotificationCategory;
    itemId: number;
    read: boolean;
    timestamp: number;
}

const MAX_HISTORY = 5;

const historyStorageKey = (loginName: string): string => `yorpas-dashboard-notif-history-${loginName}`;

export const getHistory = (loginName: string): INotificationEntry[] => {
    try {
        const raw = window.localStorage.getItem(historyStorageKey(loginName));
        if (!raw) {
            return [];
        }
        const parsed: unknown = JSON.parse(raw);
        return Array.isArray(parsed) ? (parsed as INotificationEntry[]) : [];
    } catch {
        return [];
    }
};

const setHistory = (loginName: string, entries: INotificationEntry[]): void => {
    try {
        window.localStorage.setItem(historyStorageKey(loginName), JSON.stringify(entries));
    } catch {
        /* bkz. setSeenIds'teki not — kritik değil, sessizce yok say. */
    }
};

/** Tek bir bildirimi "okundu" işaretler — listeden SİLMEZ, sadece görünümünü değiştirir. */
export const markNotificationRead = (loginName: string, entryId: string): INotificationEntry[] => {
    const updated = getHistory(loginName).map((entry) => (entry.id === entryId ? { ...entry, read: true } : entry));
    setHistory(loginName, updated);
    return updated;
};

/**
 * Bir kategorideki TÜM bildirimleri "okundu" işaretler — WelcomeHeader'ın
 * özet şeridindeki bir çipe (ör. "X yeni ilan") tıklanınca kullanılır: o
 * kategoriye ait ne kadar bildirim varsa hepsi görülmüş sayılır, sayaç
 * sıfıra döner (zil menüsündeki tek tek "okundu" işaretlemeyle tutarlı).
 */
export const markCategoryRead = (loginName: string, category: NotificationCategory): INotificationEntry[] => {
    const updated = getHistory(loginName).map((entry) => (entry.category === category ? { ...entry, read: true } : entry));
    setHistory(loginName, updated);
    return updated;
};

/**
 * Sayfa her açıldığında bir kez çağrılır: taban (baseline) id'lerle şu anki
 * en güncel id'leri karşılaştırır. Artış gösteren HER kategori için yeni,
 * OKUNMAMIŞ bir geçmiş kaydı listenin EN BAŞINA eklenir (aynı kategoriden
 * önceki kayıtlar SİLİNMEZ — kasıtlı: "3 gün önceki duyuru" ile "bugünkü
 * duyuru" ayrı ayrı bildirim olarak kalabilsin). Liste en fazla
 * MAX_HISTORY (5) kayıt tutar; taşan en eski kayıt, okunma durumu ne
 * olursa olsun düşer. Taban, üretilen her kayıt için hemen ileri alınır —
 * aksi halde bir sonraki sayfa yüklemesinde AYNI artış tekrar yeni bir
 * bildirim üretirdi.
 */
export const syncNotificationHistory = async (
    context: WebPartContext,
    loginName: string
): Promise<{ history: INotificationEntry[]; latestIds: LatestIds }> => {
    const latestIds = await getLatestIds(context);
    const baseline = getSeenIds(loginName);

    if (!baseline) {
        // İlk çalıştırma: var olan içeriği "yeni" saymadan taban al.
        setSeenIds(loginName, latestIds);
        return { history: getHistory(loginName), latestIds };
    }

    let history = getHistory(loginName);
    const nextBaseline: LatestIds = { ...baseline };
    let changed = false;

    NOTIFICATION_CATEGORY_ORDER.forEach((category) => {
        if (latestIds[category] > baseline[category]) {
            history = [
                {
                    id: `${category}-${latestIds[category]}`,
                    category,
                    itemId: latestIds[category],
                    read: false,
                    timestamp: Date.now()
                },
                ...history
            ];
            nextBaseline[category] = latestIds[category];
            changed = true;
        }
    });

    if (changed) {
        history = history.slice(0, MAX_HISTORY);
        setHistory(loginName, history);
        setSeenIds(loginName, nextBaseline);
    }

    return { history, latestIds };
};
