import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

/**
 * Katılış/Ayrılış listeleri, bu widget'ın barındığı portal sitesinde DEĞİL,
 * Teams'teki "Yazılım ve Teknoloji Ekibi" takımının kendi SharePoint
 * sitelerinde yaşıyor (Teams'teki "Katılış"/"Ayrılış" liste sekmelerinden
 * "SharePoint'te Aç" ile doğrulandı). Bu yüzden context.pageContext.web
 * yerine buradaki SABİT site adresleri kullanılıyor — SPHttpClient aynı
 * kiracı (tenant) içindeki başka bir site koleksiyonuna GET/POST atabilir,
 * tek fark POST için o sitenin KENDİ request digest'inin ayrıca alınması
 * gerekmesi (bkz. getDigest).
 */
const KATILIS_SITE_URL = 'https://yorpas.sharepoint.com/sites/YazilimTeknoloji';
const KATILIS_LIST_TITLE = 'Çalışan katılışı';
const AYRILIS_SITE_URL = 'https://yorpas.sharepoint.com/sites/Turquality-BilgiTeknolojileri';
const AYRILIS_LIST_TITLE = 'Ayrılışlar';

export type OnboardingKind = 'katilis' | 'ayrilis';

export interface IChecklistItem {
    /** Listedeki GERÇEK dahili sütun adı — PATCH gönderirken bu kullanılır. */
    internalName: string;
    label: string;
    done: boolean;
}

export interface IOnboardingRecord {
    id: number;
    /** Çalışan/Ayrılan Personel adı. */
    name: string;
    /** Sadece katılışta dolu (unvan). */
    title?: string;
    /** Sadece katılışta dolu (yönetici). */
    manager?: string;
    location: string;
    dateLabel: string;
    /** "Evet" | "Hayır" | boş. */
    transfer?: string;
    status: string;
    description?: string;
    checklist: IChecklistItem[];
}

interface ISPItem {
    [key: string]: unknown;
}

/** Ham API alan adını (ör. "A_x00e7__x0131_klama") okunabilir bir etikete eşler. */
const KATILIS_CHECKLIST_FIELDS: { internalName: string; label: string }[] = [
    { internalName: 'Complete', label: 'Bilgisayar' },
    { internalName: 'Telefon', label: 'Telefon' },
    { internalName: 'Hat', label: 'Şirket Hattı' },
    { internalName: 'Hesap', label: 'Kullanıcı Hesabı' },
    { internalName: 'Zimmet', label: 'Zimmet' },
    { internalName: 'Email', label: 'E-Posta' },
    { internalName: 'ENVANTEREGIRILDIMI', label: 'Envanter' },
    { internalName: 'Tablet', label: 'Tablet' },
    { internalName: 'ERPHesap', label: 'ERP Hesabı' },
    // Muhasebe grubu için — bu üç sütun kullanıcı isteğiyle doğrudan
    // "Çalışan katılışı" listesine (Boolean alan) eklendi.
    { internalName: 'Bordro_x0020_Sistemi', label: 'Bordro Sistemi' },
    { internalName: 'Banka_x0020_Hesab_x0131_', label: 'Banka Hesabı' },
    { internalName: 'Muhasebe_x0020_Kayd_x0131_', label: 'Muhasebe Kaydı' }
];

const AYRILIS_CHECKLIST_FIELDS: { internalName: string; label: string }[] = [
    { internalName: 'Complete', label: 'Zimmet' },
    { internalName: 'Envanter', label: 'Envanter' },
    { internalName: 'Hesaplar', label: 'Kullanıcı Hesapları' },
    { internalName: 'E_x002d_Posta', label: 'E-Posta' },
    { internalName: 'ERP', label: 'ERP Hesapları' },
    { internalName: 'Hat', label: 'Şirket Hattı' },
    { internalName: 'MailGruplar_x0131_', label: 'Mail Grupları' },
    { internalName: 'Lisans_x0130_ptali', label: 'Lisans İptali' }
];

const getListConfig = (kind: OnboardingKind): { siteUrl: string; listTitle: string; checklist: { internalName: string; label: string }[] } =>
    kind === 'katilis'
        ? { siteUrl: KATILIS_SITE_URL, listTitle: KATILIS_LIST_TITLE, checklist: KATILIS_CHECKLIST_FIELDS }
        : { siteUrl: AYRILIS_SITE_URL, listTitle: AYRILIS_LIST_TITLE, checklist: AYRILIS_CHECKLIST_FIELDS };

/** Yönetici kullanıcı aramasının hangi site koleksiyonuna karşı yapılacağını widget'a verir. */
export const getOnboardingSiteUrl = (kind: OnboardingKind): string => getListConfig(kind).siteUrl;

const extractSPErrorDetail = async (response: SPHttpClientResponse): Promise<string> => {
    try {
        const raw = await response.text();
        const parsed = JSON.parse(raw) as { 'odata.error'?: { message?: { value?: string } } };
        return parsed['odata.error']?.message?.value ?? `HTTP ${response.status}`;
    } catch {
        return `HTTP ${response.status}`;
    }
};

/**
 * Date.toISOString() UTC'ye çevirirken yerel gece yarısını (ör. TR, UTC+3)
 * bir önceki güne kaydırabiliyor (05.08 seçilince 04.08 olarak kaydediliyordu).
 * Bunun yerine sadece seçilen takvim gününü, saat bileşeni olmadan, UTC
 * öğlen (12:00) sabitleyerek gönderiyoruz — hiçbir saat dilimi bunu bir
 * önceki/sonraki güne kaydıramaz.
 */
const toDateOnlyIso = (date: Date): string => {
    const year = date.getFullYear();
    const month = date.getMonth() + 1 < 10 ? `0${date.getMonth() + 1}` : `${date.getMonth() + 1}`;
    const day = date.getDate() < 10 ? `0${date.getDate()}` : `${date.getDate()}`;
    return `${year}-${month}-${day}T12:00:00Z`;
};

const toDateLabel = (iso?: string): string => {
    if (!iso) {
        return '—';
    }
    const match = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso);
    if (!match) {
        return '—';
    }
    return `${match[3]}.${match[2]}.${match[1]}`;
};

const mapListItem = (
    kind: OnboardingKind,
    checklist: { internalName: string; label: string }[],
    item: ISPItem
): IOnboardingRecord => {
    // Ayrılışlar listesinde "Ayrılan Personel" temiz görünen ad, "Title"
    // ise "AD SOYAD / jira-linki" gibi karışık — bu yüzden varsa o
    // önceliklidir. Katılışta ise çalışan adı doğrudan Title'da.
    const name = kind === 'ayrilis'
        ? (item.Ayr_x0131_lanPersonel as string) || (item.Title as string) || ''
        : (item.Title as string) || '';

    return {
        id: item.Id as number,
        name,
        title: kind === 'katilis' ? (item.Unvan as string) : undefined,
        manager: kind === 'katilis' ? (item.Y_x00f6_netici as string) : undefined,
        location: kind === 'katilis' ? (item.Lokasyon as string) : (item.Completeby as string),
        dateLabel: toDateLabel(item.Completedon as string | undefined),
        transfer: item.Devir as string | undefined,
        status: (item.Durum as string) || '',
        description: kind === 'katilis' ? (item.A_x00e7__x0131_klama as string | undefined) : undefined,
        checklist: checklist.map((field) => ({
            internalName: field.internalName,
            label: field.label,
            done: !!item[field.internalName]
        }))
    };
};

export const getOnboardingRecords = async (context: WebPartContext, kind: OnboardingKind): Promise<IOnboardingRecord[]> => {
    const { siteUrl, listTitle, checklist } = getListConfig(kind);
    const endpoint = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items?$top=2000&$orderby=Completedon desc`;

    const response: SPHttpClientResponse = await context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
    if (!response.ok) {
        const detail = await extractSPErrorDetail(response);
        console.error(`[OnboardingService] "${listTitle}" listesi okunamadı (${siteUrl}) — ${detail}`);
        throw new Error(`"${listTitle}" listesi okunamadı (${detail})`);
    }

    const body: { value: ISPItem[] } = await response.json();
    return (body.value ?? []).map((item) => mapListItem(kind, checklist, item));
};

/**
 * Herkese açık, izin gerektirmeyen "son N kişi" widget'ları (Aramıza
 * Katılanlar / Aramızdan Ayrılanlar) için — tüm listeyi çekmek yerine
 * $top ile sunucu tarafında sadece son N kaydı ister.
 */
export const getRecentOnboardingRecords = async (
    context: WebPartContext,
    kind: OnboardingKind,
    count: number = 5
): Promise<IOnboardingRecord[]> => {
    const { siteUrl, listTitle, checklist } = getListConfig(kind);
    const endpoint = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items?$top=${count}&$orderby=Completedon desc`;

    const response: SPHttpClientResponse = await context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
    if (!response.ok) {
        const detail = await extractSPErrorDetail(response);
        console.error(`[OnboardingService] "${listTitle}" listesi (son kayıtlar) okunamadı (${siteUrl}) — ${detail}`);
        throw new Error(`"${listTitle}" listesi okunamadı (${detail})`);
    }

    const body: { value: ISPItem[] } = await response.json();
    return (body.value ?? []).map((item) => mapListItem(kind, checklist, item));
};

export interface IOrgUser {
    displayName: string;
    email: string;
}

/**
 * "Yönetici" alanı SharePoint listesinde düz metin (Text) — gerçek bir
 * Person alanına çevirmek liste şemasını değiştirmeyi gerektirir. Bunun
 * yerine kullanıcı deneyimini iyileştirmek için: SharePoint'in tenant
 * genelindeki kullanıcı dizininde (Azure AD) arama yapan standart
 * ClientPeoplePickerWebServiceInterface'i kullanıyoruz, kullanıcı seçince
 * sadece görünen adını metin olarak Y_x00f6_netici alanına yazıyoruz.
 */
export const searchOrgUsers = async (context: WebPartContext, siteUrl: string, queryText: string): Promise<IOrgUser[]> => {
    if (!queryText.trim()) {
        return [];
    }
    try {
        const digest = await getRequestDigest(context, siteUrl);
        const endpoint = `${siteUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`;
        // NOT: context.spHttpClient.post bu uç noktada tutarlı biçimde HTTP 400
        // döndürüyordu (muhtemelen odata=verbose gövdesini kendi istek
        // ardışık düzeninde başka türlü işlemesinden) — aynı çağrı ham fetch
        // ile (tarayıcının zaten mevcut SharePoint oturum çerezleriyle)
        // sorunsuz çalışıyor, bu yüzden burada ona geri dönüldü.
        const response = await fetch(endpoint, {
            method: 'POST',
            credentials: 'include',
            headers: {
                Accept: 'application/json;odata=nometadata',
                'Content-type': 'application/json;odata=verbose',
                'X-RequestDigest': digest
            },
            body: JSON.stringify({
                queryParams: {
                    '__metadata': { type: 'SP.UI.ApplicationPages.ClientPeoplePickerQueryParameters' },
                    AllowEmailAddresses: true,
                    AllowMultipleEntities: false,
                    MaximumEntitySuggestions: 8,
                    PrincipalSource: 15,
                    PrincipalType: 1,
                    QueryString: queryText
                }
            })
        });
        if (!response.ok) {
            return [];
        }
        const body: { value: string } = await response.json();
        const parsed = JSON.parse(body.value) as { DisplayText: string; EntityData?: { Email?: string } }[];
        return parsed.map((entity) => ({
            // ÖNCEKİ HATA: DisplayText genelde "Ad Soyad | ŞUBE-KODU" şeklinde
            // geliyor (ör. "Taylan TOPTAŞ | BITECH") — bu uzun metin hem
            // seçim rozetinde (chip) X butonuyla çakışıp ismi yarım
            // gösteriyordu, hem de "Yönetici" metin alanına gereksiz bir
            // şube kodu yazıyordu. Sadece " | " öncesini (temiz adı) alıyoruz.
            displayName: entity.DisplayText.split(' | ')[0].trim(),
            email: entity.EntityData?.Email ?? ''
        }));
    } catch (error) {
        console.error('[OnboardingService] Kullanıcı araması başarısız:', error);
        return [];
    }
};

/** Hedef sitenin KENDİ form digest'ini alır — cross-site POST için SPHttpClient bunu otomatik yapmaz. */
const getRequestDigest = async (context: WebPartContext, siteUrl: string): Promise<string> => {
    const response = await context.spHttpClient.post(`${siteUrl}/_api/contextinfo`, SPHttpClient.configurations.v1, {
        headers: { Accept: 'application/json;odata=nometadata' }
    });
    if (!response.ok) {
        throw new Error(`Site erişim belirteci (digest) alınamadı (HTTP ${response.status})`);
    }
    const body: { FormDigestValue: string } = await response.json();
    return body.FormDigestValue;
};

export interface IOnboardingSubmitResult {
    success: boolean;
    errorMessage?: string;
}

export interface INewOnboardingInput {
    name: string;
    title?: string;
    manager?: string;
    location: string;
    date: Date;
    transfer: string;
    description?: string;
}

/** Yeni bir Katılış/Ayrılış kaydı açar — SADECE İK grubu için (bkz. usePermissions.canManageOnboarding). */
export const createOnboardingRecord = async (
    context: WebPartContext,
    kind: OnboardingKind,
    input: INewOnboardingInput
): Promise<IOnboardingSubmitResult> => {
    const { siteUrl, listTitle, checklist } = getListConfig(kind);
    const endpoint = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items`;

    try {
        const digest = await getRequestDigest(context, siteUrl);

        const fields: Record<string, unknown> = {
            Title: input.name,
            Completedon: toDateOnlyIso(input.date),
            Devir: input.transfer,
            Durum: 'DEVAM EDİYOR'
        };
        // ÖNCEKİ HATA: checklist alanları (Bilgisayar, Telefon, E-Posta, ...)
        // hiç gönderilmiyordu — bu sütunların SharePoint'teki varsayılan
        // değeri "Evet" olduğu için yeni kayıt hep TAMAMEN İŞARETLİ
        // görünüyordu. Artık hepsi açıkça false gönderiliyor.
        checklist.forEach((field) => {
            fields[field.internalName] = false;
        });
        if (kind === 'katilis') {
            fields.Unvan = input.title ?? '';
            fields.Y_x00f6_netici = input.manager ?? '';
            fields.Lokasyon = input.location;
            fields.A_x00e7__x0131_klama = input.description ?? '';
        } else {
            fields.Ayr_x0131_lanPersonel = input.name;
            fields.Completeby = input.location;
        }

        const response = await context.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, {
            headers: {
                Accept: 'application/json;odata=nometadata',
                'Content-type': 'application/json;odata=nometadata',
                'X-RequestDigest': digest
            },
            body: JSON.stringify(fields)
        });

        if (!response.ok) {
            const detail = await extractSPErrorDetail(response);
            console.error(`[OnboardingService] "${listTitle}" kaydı oluşturulamadı — ${detail}`);
            return { success: false, errorMessage: detail };
        }
        return { success: true };
    } catch (error) {
        console.error('[OnboardingService] Kayıt oluşturulurken beklenmeyen hata:', error);
        return { success: false, errorMessage: (error as Error).message };
    }
};

/** Bir kaydın kontrol listesi (checklist) alanlarını ve durumunu günceller — İK/BT/Muhasebe için (bkz. usePermissions.canEditOnboarding). */
export const updateOnboardingRecord = async (
    context: WebPartContext,
    kind: OnboardingKind,
    itemId: number,
    checklist: IChecklistItem[],
    status: string
): Promise<IOnboardingSubmitResult> => {
    const { siteUrl, listTitle } = getListConfig(kind);
    const endpoint = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items(${itemId})`;

    try {
        const digest = await getRequestDigest(context, siteUrl);

        const fields: Record<string, unknown> = { Durum: status };
        checklist.forEach((item) => {
            fields[item.internalName] = item.done;
        });

        const response = await context.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, {
            headers: {
                Accept: 'application/json;odata=nometadata',
                'Content-type': 'application/json;odata=nometadata',
                'X-RequestDigest': digest,
                'IF-MATCH': '*',
                'X-HTTP-Method': 'MERGE'
            },
            body: JSON.stringify(fields)
        });

        if (!response.ok) {
            const detail = await extractSPErrorDetail(response);
            console.error(`[OnboardingService] "${listTitle}" öğesi (${itemId}) güncellenemedi — ${detail}`);
            return { success: false, errorMessage: detail };
        }
        return { success: true };
    } catch (error) {
        console.error('[OnboardingService] Kayıt güncellenirken beklenmeyen hata:', error);
        return { success: false, errorMessage: (error as Error).message };
    }
};

/** Bir kaydı kalıcı olarak siler — SADECE İK grubu için (bkz. usePermissions.canManageOnboarding). */
export const deleteOnboardingRecord = async (
    context: WebPartContext,
    kind: OnboardingKind,
    itemId: number
): Promise<IOnboardingSubmitResult> => {
    const { siteUrl, listTitle } = getListConfig(kind);
    const endpoint = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items(${itemId})`;

    try {
        const digest = await getRequestDigest(context, siteUrl);

        const response = await context.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, {
            headers: {
                Accept: 'application/json;odata=nometadata',
                'X-RequestDigest': digest,
                'IF-MATCH': '*',
                'X-HTTP-Method': 'DELETE'
            }
        });

        if (!response.ok) {
            const detail = await extractSPErrorDetail(response);
            console.error(`[OnboardingService] "${listTitle}" öğesi (${itemId}) silinemedi — ${detail}`);
            return { success: false, errorMessage: detail };
        }
        return { success: true };
    } catch (error) {
        console.error('[OnboardingService] Kayıt silinirken beklenmeyen hata:', error);
        return { success: false, errorMessage: (error as Error).message };
    }
};
