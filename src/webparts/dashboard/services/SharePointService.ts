import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IBirthdayItem {
    id: number;
    name: string;
    department: string;
    dateLabel: string;
    /** Ayın günü (1-31) — sıralama için. */
    day: number;
    /** Doğum günü bugün mü — "İyi ki doğdun" rozetini tetiklemek için. */
    isToday: boolean;
}

export interface IFeedbackSubmitResult {
    success: boolean;
    errorMessage?: string;
}

export interface IAnnouncementItem {
    id: number;
    title: string;
    /** Kart önizlemesi için kısaltılmış, düz metne çevrilmiş özet. */
    excerpt: string;
    /** Modal/Dialog içinde HTML olarak render edilecek ham içerik (zengin metin korunur). */
    bodyHtml: string;
    dateLabel: string;
    /** Öğeye eklenmiş ilk EK (attachment) dosyasının SharePoint URL'i — yoksa undefined. */
    imageUrl?: string;
}

export interface IUpcomingEventItem {
    id: number;
    title: string;
    dateLabel: string;
    location: string;
    /** Takvim yaprağı için kısaltılmış BÜYÜK HARF ay adı (ör. "TEM"). */
    monthShort: string;
    /** Takvim yaprağı için ayın günü (1-31). */
    day: number;
    /** Etkinliğin saati (ör. "09:00") — EventDate'in saat bileşeni. */
    timeLabel: string;
    /** Öğeye eklenmiş ilk EK (attachment) dosyasının SharePoint URL'i — yoksa undefined. */
    imageUrl?: string;
    /**
     * ÖNCEKİ HATA: "Etkinlik Detayı" formda kaydediliyordu (bkz. createEvent'in
     * Description alanı) ama bu arayüz hiç bir slot taşımıyordu — kullanıcı
     * girdiği açıklama sunucuya yazılsa bile hiçbir yerde tekrar
     * GÖSTERİLMİYORDU. Artık okunuyor (bkz. getUpcomingEvents).
     */
    description?: string;
}

// Liste adları — tenant'taki gerçek SharePoint listeleriyle birebir eşleşmeli.
const BIRTHDAYS_LIST_TITLE = 'Birthdays';
const FEEDBACK_LIST_TITLE = 'Feedbacks';
const ANNOUNCEMENTS_LIST_TITLE = 'Duyurular';
const EVENTS_LIST_TITLE = 'Etkinlikler';

// Ad Soyad ve Doğum Tarihi sütunlarının DAHİLİ (internal) adları — SharePoint,
// görünen adındaki boşlukları "_x0020_" ile kodlar (ör. "Ad Soyad" ->
// "Ad_x0020_Soyad"). Bu değerler ekran görüntülerinden teyit edilmiştir.
const BIRTHDAY_NAME_FIELD = 'Ad_x0020_Soyad';
const DATE_OF_BIRTH_FIELD = 'Dogum_x0020_Tarihi';

// Geri bildirim metninin yazılacağı GERÇEK yorum sütunu (görünen adı "Comment").
const FEEDBACK_COMMENT_FIELD = 'Comment';

// Gönderen kimliği ve tarih sütunlarının TAHMİNİ dahili adları — bu tam
// olarak tutmayabilir (görünen ad ile dahili ad farklı olabilir), bu yüzden
// submitFeedback bu isimleri önce dener, tutmazsa listenin gerçek şemasından
// (Title/TypeAsString) doğru alanı otomatik bulur. "Worker" bir Kişi/Grup
// alanı OLMAYABİLİR (düz metin de olabilir) — TypeAsString'e göre hem
// Person-Id hem düz metin gönderimi ayrı ayrı destekleniyor.
const FEEDBACK_PERSON_FIELD = 'Worker';
const FEEDBACK_DATE_FIELD = 'Tarih';

// Duyuru gövdesi/içeriği için beklenen dahili sütun adı — yoksa özet boş kalır
// ve konsola uyarı basılır.
const ANNOUNCEMENT_BODY_FIELD = 'Duyuru';

// "Etkinlikler" standart bir SharePoint Takvim (Calendar) listesiyse bu iki
// alan OOTB dahili adlardır. Özel bir liste kullanıyorsanız güncelleyin.
const EVENT_DATE_FIELD = 'EventDate';
const EVENT_LOCATION_FIELD = 'Location';

interface ISPListItem {
    Id: number;
    Title: string;
    Created?: string;
    // Gerçek liste, burada tanımlanmayan başka alanlar da içerebilir — teşhis
    // (diagnostic) amaçlı Object.keys() ile okunabilmesi için index signature
    // bırakıyoruz.
    [key: string]: unknown;
}

interface ISPListField {
    InternalName: string;
    Title: string;
    TypeAsString: string;
    Required: boolean;
}

interface ISPODataError {
    'odata.error'?: {
        code?: string;
        message?: { lang?: string; value?: string };
    };
    error?: { message?: string };
}

/**
 * SharePoint REST'in HTTP 400/403/... hata gövdesini okunabilir bir metne çevirir.
 * Body sadece bir kez okunabildiği için (stream), tek seferde .text() ile alıp
 * ardından JSON.parse deniyoruz — böylece hem odata.error.message.value hem de
 * ham metin (JSON olmayan bir hata sayfası dönerse) yakalanmış olur.
 */
const extractSPErrorDetail = async (response: SPHttpClientResponse): Promise<string> => {
    let raw = '';
    try {
        raw = await response.text();
    } catch {
        return `HTTP ${response.status} (yanıt gövdesi okunamadı)`;
    }

    try {
        const parsed: ISPODataError = JSON.parse(raw);
        const detail = parsed['odata.error']?.message?.value ?? parsed.error?.message;
        if (detail) {
            return `HTTP ${response.status} — ${detail}`;
        }
    } catch {
        // Yanıt JSON değil (ör. HTML hata sayfası) — ham metni kullan.
    }

    return `HTTP ${response.status} — ${raw.substring(0, 500)}`;
};

/** Basit HTML etiketi temizleyici — zengin metin (rich text) alanlarından kısa özet çıkarmak için. */
const stripHtml = (html: string): string => html.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim();

const TURKISH_MONTH_NAMES = [
    'Ocak', 'Şubat', 'Mart', 'Nisan', 'Mayıs', 'Haziran',
    'Temmuz', 'Ağustos', 'Eylül', 'Ekim', 'Kasım', 'Aralık'
];

// "Takvim yaprağı" kartındaki kısaltılmış BÜYÜK HARF ay adları (ör. "TEM").
const TURKISH_MONTH_SHORT = ['OCA', 'ŞUB', 'MAR', 'NİS', 'MAY', 'HAZ', 'TEM', 'AĞU', 'EYL', 'EKİ', 'KAS', 'ARA'];

interface ISPAttachment {
    FileName: string;
    ServerRelativeUrl: string;
}

/**
 * Bir liste öğesine eklenmiş EK (attachment) dosyalarını çeker ve varsa
 * İLKİNİN URL'ini döner — widget'ta "kapak görseli" olarak kullanılır. Öğenin
 * standart "Attachments" alanı false ise ağ isteği hiç yapılmaz (gereksiz
 * REST çağrısından kaçınmak için). Herhangi bir hata durumunda sessizce
 * undefined döner — kapak görseli göstermemek, widget'ı kilitlemekten iyidir.
 */
const getFirstAttachmentUrl = async (
    context: WebPartContext,
    listTitle: string,
    itemId: number
): Promise<string | undefined> => {
    try {
        const webUrl = context.pageContext.web.absoluteUrl;
        const endpoint = `${webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items(${itemId})/AttachmentFiles`;
        const response = await context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
        if (!response.ok) {
            return undefined;
        }
        const body: { value: ISPAttachment[] } = await response.json();
        return body.value?.[0]?.ServerRelativeUrl;
    } catch {
        return undefined;
    }
};

/**
 * Yeni oluşturulan bir liste öğesine kullanıcının seçtiği görseli EK
 * (attachment) olarak yükler. Görsel opsiyonel bir "iyi olursa güzel olur"
 * özelliği olduğu için bu fonksiyon ASLA fırlatmaz (throw) — yükleme
 * başarısız olsa bile öğenin kendisi (başlık/içerik/tarih) zaten kaydedilmiş
 * olur, kullanıcı sadece kapak görseli olmadan devam eder.
 */
const uploadItemImage = async (
    context: WebPartContext,
    listTitle: string,
    itemId: number,
    file: File
): Promise<void> => {
    try {
        const webUrl = context.pageContext.web.absoluteUrl;
        // Dosya adı SharePoint URL'inde sorun çıkarabilecek karakterlerden
        // (ör. #, %) arındırılıyor; uzantı korunuyor.
        const safeName = file.name.replace(/[^a-zA-Z0-9._-]/g, '_');
        const endpoint =
            `${webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items(${itemId})` +
            `/AttachmentFiles/add(FileName='${encodeURIComponent(safeName)}')`;
        const fileBuffer = await file.arrayBuffer();
        const response = await context.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, {
            headers: { Accept: 'application/json;odata=nometadata' },
            body: fileBuffer
        });
        if (!response.ok) {
            console.warn(`[SharePointService] Görsel eki yüklenemedi ("${listTitle}", öğe ${itemId}):`, await extractSPErrorDetail(response));
        }
    } catch (error) {
        console.warn(`[SharePointService] Görsel eki yüklenirken beklenmeyen hata ("${listTitle}", öğe ${itemId}):`, error);
    }
};

/**
 * SharePoint'in döndürdüğü ISO tarih string'inden (ör. "2026-07-15T00:00:00Z")
 * ay/gün bileşenlerini DOĞRUDAN, tarayıcının yerel saat dilimine hiç
 * uğramadan okur. new Date(iso).getDate() KULLANMIYORUZ çünkü UTC gece
 * yarısı bir tarih, tarayıcının yerel saat dilimine göre bir önceki/sonraki
 * güne kayabilir (ör. UTC-5 bir tarayıcıda 15 Temmuz 00:00 UTC, yerel saatte
 * hâlâ 14 Temmuz'dur) — bu da "bugün doğanlar" karşılaştırmasını gün
 * kaydırmasıyla bozar. Doğum tarihi gibi saatten bağımsız, sabit bir takvim
 * günü için ham string'ten okumak tek doğru yoldur.
 */
const getMonthDayFromIso = (iso: string): { month: number; day: number } | undefined => {
    const match = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso);
    if (!match) {
        return undefined;
    }
    return { month: parseInt(match[2], 10) - 1, day: parseInt(match[3], 10) };
};

/** Konsola, listede beklenen bir sütunun bulunamadığını ve gerçek sütun adlarını basar. */
const warnMissingField = (listTitle: string, fieldName: string, sampleItem: ISPListItem): void => {
    console.warn(
        `[SharePointService] "${fieldName}" sütunu "${listTitle}" listesinde bulunamadı. ` +
        `Bu listedeki gerçek sütun adları: ${Object.keys(sampleItem).join(', ')}. ` +
        'SharePointService.ts içindeki ilgili alan sabitini gerçek sütun adına göre güncelleyin.'
    );
};

/**
 * Verilen listeden ham öğeleri çeker. $select KASITLI olarak KULLANILMIYOR:
 * sütun adı uyuşmazlığı bu sayede HTTP 400'e değil, çağıran fonksiyondaki
 * client-side eşlemede sessiz/eksik sonuca düşer ve konsoldaki uyarıdan
 * teşhis edilebilir. Sıralama/filtreleme de OData yerine JS tarafında yapılır.
 */
const getListItems = async (context: WebPartContext, listTitle: string, top: number): Promise<ISPListItem[]> => {
    const webUrl = context.pageContext.web.absoluteUrl;
    const endpoint = `${webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items?$top=${top}`;

    const response: SPHttpClientResponse = await context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);

    if (!response.ok) {
        const detail = await extractSPErrorDetail(response);
        console.error(`[SharePointService] "${listTitle}" listesi okunamadı — istek: ${endpoint}`, detail);
        throw new Error(`"${listTitle}" listesi okunamadı (${detail})`);
    }

    const body: { value: ISPListItem[] } = await response.json();
    return body.value ?? [];
};

/**
 * Verilen listenin TÜM öğelerini, sayfalama (odata.nextLink) takip ederek
 * çeker — tek istekte $top ile sınırlı KALMAZ. ÖNCEKİ HATA (Doğum Günleri
 * widget'ı): getListItems'a sabit $top=500 veriliyordu; "Birthdays" listesi
 * aslında 500'den FAZLA satır içeriyordu (ör. 859), SharePoint de tam olarak
 * istenen 500'ü döndürüp gerisini sessizce görmezden geliyordu. Sonuç: ID'si
 * ilk 500'ün dışında kalan kişiler — doğum günleri ay içinde/yaklaşan olsa
 * bile — HİÇ ÇEKİLMİYORDU (filtre/sıralama sorunu değil, veri hiç gelmiyordu).
 * Bu fonksiyon liste boyutundan bağımsız olarak HER satırı garanti eder.
 */
const getAllListItems = async (context: WebPartContext, listTitle: string): Promise<ISPListItem[]> => {
    const webUrl = context.pageContext.web.absoluteUrl;
    let endpoint: string | undefined =
        `${webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items?$top=2000`;
    const all: ISPListItem[] = [];

    while (endpoint) {
        const response: SPHttpClientResponse = await context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);

        if (!response.ok) {
            const detail = await extractSPErrorDetail(response);
            console.error(`[SharePointService] "${listTitle}" listesi okunamadı — istek: ${endpoint}`, detail);
            throw new Error(`"${listTitle}" listesi okunamadı (${detail})`);
        }

        const body: { value: ISPListItem[]; 'odata.nextLink'?: string } = await response.json();
        all.push(...(body.value ?? []));
        endpoint = body['odata.nextLink'];
    }

    return all;
};

/**
 * Listenin GERÇEK sütun (alan) şemasını çeker — dahili ad tahminimiz tutmadığında
 * doğru sütunu tipe göre (ör. Tarih/Saat) otomatik bulabilmek için kullanılır.
 * İstek başarısız olursa (yetki vb.) sessizce boş dizi döner; bu fonksiyon asla
 * fırlatmaz, sadece otomatik-algılama devre dışı kalır.
 */
const getListFieldsMeta = async (context: WebPartContext, listTitle: string): Promise<ISPListField[]> => {
    try {
        const webUrl = context.pageContext.web.absoluteUrl;
        // NOT: "Hidden eq false" filtresi KASITLI OLARAK kullanılmıyor — bir
        // sütun formlardan gizlenmiş olsa bile (ör. "Worker" gizli bir alan
        // olabilir) REST üzerinden yine de yazılabilir olmalı; filtre olsaydı
        // böyle bir alan şemada hiç görünmez ve otomatik algılama sessizce
        // atlardı.
        const endpoint = `${webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/fields` +
            '?$select=InternalName,Title,TypeAsString,Required';
        const response = await context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
        if (!response.ok) {
            return [];
        }
        const body: { value: ISPListField[] } = await response.json();
        return body.value ?? [];
    } catch {
        return [];
    }
};

/**
 * "Birthdays" listesinden bugünden itibaren (bugün dahil) bu ay içindeki
 * YAKLAŞAN doğum günlerini çeker. Geçmiş günler KASITLI OLARAK listeye
 * dahil edilmez (kullanıcı isteği). Ay/gün karşılaştırması tamamen JS
 * tarafında yapılır (OData $filter tarih karşılaştırmasında sorun çıkardığı
 * için kullanılmaz).
 *
 * KENDİ KENDİNİ ONARAN ALAN ALGILAMA: BIRTHDAY_NAME_FIELD / DATE_OF_BIRTH_FIELD
 * sabitleri gerçek liste öğesinde yoksa (dahili ad tahmini yanlışsa), listenin
 * GERÇEK şeması (/fields) sorgulanır ve tarih için TypeAsString==='DateTime'
 * olan ilk özel sütun otomatik kullanılır; isim için Title'a düşülür. Böylece
 * dahili ad uyuşmazlığı doğum günlerinin sessizce boş görünmesine yol açmaz.
 */
export const getBirthdaysThisMonth = async (context: WebPartContext): Promise<IBirthdayItem[]> => {
    const items = await getAllListItems(context, BIRTHDAYS_LIST_TITLE);

    let nameField = BIRTHDAY_NAME_FIELD;
    let dateField = DATE_OF_BIRTH_FIELD;

    if (items.length > 0 && (!(dateField in items[0]) || !(nameField in items[0]))) {
        const fields = await getListFieldsMeta(context, BIRTHDAYS_LIST_TITLE);

        if (!(dateField in items[0])) {
            const autoDateField = fields.filter((f) =>
                f.TypeAsString === 'DateTime' && f.InternalName !== 'Created' && f.InternalName !== 'Modified'
            )[0];
            if (autoDateField && autoDateField.InternalName in items[0]) {
                console.warn(
                    `[SharePointService] "${dateField}" sütunu bulunamadı, otomatik algılanan tarih sütunu ` +
                    `kullanılıyor: "${autoDateField.InternalName}" (${autoDateField.Title}).`
                );
                dateField = autoDateField.InternalName;
            } else {
                warnMissingField(BIRTHDAYS_LIST_TITLE, dateField, items[0]);
            }
        }

        if (!(nameField in items[0])) {
            console.warn(`[SharePointService] "${nameField}" sütunu bulunamadı, "Title" alanına geri dönülüyor.`);
            nameField = 'Title';
        }
    }

    const today = new Date();
    const currentMonth = today.getMonth();
    const currentDay = today.getDate();

    return items
        .map((item) => {
            const raw = item[dateField] as string | undefined;
            const parsed = raw ? getMonthDayFromIso(raw) : undefined;
            return { item, parsed };
        })
        // Sadece bu ay VE bugün ya da sonrası — geçmiş günler listeye eklenmez.
        .filter((entry): entry is { item: ISPListItem; parsed: { month: number; day: number } } =>
            !!entry.parsed && entry.parsed.month === currentMonth && entry.parsed.day >= currentDay
        )
        .map(({ item, parsed }) => ({
            id: item.Id,
            name: (item[nameField] as string) ?? item.Title,
            department: (item.Department as string) ?? '',
            day: parsed.day,
            isToday: parsed.day === currentDay,
            dateLabel: `${parsed.day} ${TURKISH_MONTH_NAMES[parsed.month]}`
        }))
        .sort((a, b) => a.day - b.day);
};

/**
 * "Duyurular" listesinden güncel şirket haberlerini çeker. Sıralama Created
 * alanına göre (her SharePoint listesinde standart olarak bulunur) JS
 * tarafında yapılır, en yeni önce.
 */
export const getAnnouncements = async (context: WebPartContext): Promise<IAnnouncementItem[]> => {
    const items = await getListItems(context, ANNOUNCEMENTS_LIST_TITLE, 30);

    if (items.length > 0 && !(ANNOUNCEMENT_BODY_FIELD in items[0])) {
        warnMissingField(ANNOUNCEMENTS_LIST_TITLE, ANNOUNCEMENT_BODY_FIELD, items[0]);
    }

    const createdTime = (item: ISPListItem): number => (item.Created ? new Date(item.Created).getTime() : 0);
    const sorted = items.slice().sort((a, b) => createdTime(b) - createdTime(a));

    // Ekleri (attachment) SADECE "Attachments" alanı true olan öğeler için,
    // hepsi PARALEL olarak çekiliyor — her öğe için sırayla beklemek widget'ı
    // gereksiz yavaşlatırdı.
    const imageUrls = await Promise.all(
        sorted.map((item) => (item.Attachments ? getFirstAttachmentUrl(context, ANNOUNCEMENTS_LIST_TITLE, item.Id) : Promise.resolve(undefined)))
    );

    return sorted.map((item, index) => {
        const rawBody = item[ANNOUNCEMENT_BODY_FIELD];
        const bodyHtml = typeof rawBody === 'string' ? rawBody : '';
        const excerptSource = bodyHtml ? stripHtml(bodyHtml) : '';
        return {
            id: item.Id,
            title: item.Title,
            excerpt: excerptSource.length > 180 ? `${excerptSource.substring(0, 180)}…` : excerptSource,
            bodyHtml,
            dateLabel: item.Created
                ? new Date(item.Created).toLocaleDateString('tr-TR', { day: 'numeric', month: 'long', year: 'numeric' })
                : '',
            imageUrl: imageUrls[index]
        };
    });
};

/**
 * "Etkinlikler" listesinden (standart bir SharePoint Takvim listesi varsayılır)
 * bugünden itibaren yaklaşan etkinlikleri çeker. Tarih filtresi ve sıralaması
 * JS tarafında yapılır.
 */
export const getUpcomingEvents = async (context: WebPartContext): Promise<IUpcomingEventItem[]> => {
    const items = await getListItems(context, EVENTS_LIST_TITLE, 200);

    if (items.length > 0 && !(EVENT_DATE_FIELD in items[0])) {
        warnMissingField(EVENTS_LIST_TITLE, EVENT_DATE_FIELD, items[0]);
    }

    const startOfToday = new Date();
    startOfToday.setHours(0, 0, 0, 0);

    const upcoming = items
        .filter((item) => {
            const raw = item[EVENT_DATE_FIELD] as string | undefined;
            if (!raw) {
                return false;
            }
            const parsed = new Date(raw);
            return !isNaN(parsed.getTime()) && parsed.getTime() >= startOfToday.getTime();
        })
        .sort((a, b) =>
            new Date(a[EVENT_DATE_FIELD] as string).getTime() - new Date(b[EVENT_DATE_FIELD] as string).getTime())
        .slice(0, 6);

    // Ekleri (attachment) SADECE "Attachments" alanı true olan öğeler için,
    // hepsi PARALEL olarak çekiliyor.
    const imageUrls = await Promise.all(
        upcoming.map((item) => (item.Attachments ? getFirstAttachmentUrl(context, EVENTS_LIST_TITLE, item.Id) : Promise.resolve(undefined)))
    );

    return upcoming.map((item, index) => {
        const eventDate = new Date(item[EVENT_DATE_FIELD] as string);
        return {
            id: item.Id,
            title: item.Title,
            dateLabel: eventDate.toLocaleDateString('tr-TR', { day: 'numeric', month: 'long' }),
            location: (item[EVENT_LOCATION_FIELD] as string) ?? '',
            monthShort: TURKISH_MONTH_SHORT[eventDate.getMonth()],
            day: eventDate.getDate(),
            timeLabel: eventDate.toLocaleTimeString('tr-TR', { hour: '2-digit', minute: '2-digit' }),
            imageUrl: imageUrls[index],
            description: (item.Description as string) || undefined
        };
    });
};

/**
 * Oturum sahibinin hesabından "taylan.toptas" biçiminde bir kullanıcı adı
 * türetir (e-posta @ işaretinden önceki kısım). Feedback listesindeki
 * "Worker" sütunu gerçek bir Kişi/Grup alanı değilse (düz metinse), oraya
 * kullanıcının kendi yazdığı bir şey değil, hesabından otomatik türetilen bu
 * kimlik yazılır.
 */
const getUsernameFromContext = (context: WebPartContext): string => {
    const email = context.pageContext.user.email;
    if (email && email.indexOf('@') > -1) {
        return email.split('@')[0];
    }
    const loginName = context.pageContext.user.loginName;
    const claimsMatch = /\|([^|]+)$/.exec(loginName ?? '');
    const rawLogin = claimsMatch ? claimsMatch[1] : loginName;
    if (rawLogin && rawLogin.indexOf('@') > -1) {
        return rawLogin.split('@')[0];
    }
    return context.pageContext.user.displayName;
};

/** Oturum sahibinin SharePoint kullanıcı ID'sini döner (Person/Group alanına yazmak için gerekir). */
const getCurrentUserId = async (context: WebPartContext): Promise<number | undefined> => {
    try {
        const webUrl = context.pageContext.web.absoluteUrl;
        const response = await context.spHttpClient.get(`${webUrl}/_api/web/currentuser`, SPHttpClient.configurations.v1);
        if (!response.ok) {
            return undefined;
        }
        const data: { Id: number } = await response.json();
        return data.Id;
    } catch {
        return undefined;
    }
};

/**
 * Kullanıcı geri bildirimini "Feedbacks" listesine yeni bir öğe olarak yazar.
 *
 * Üç alan ayrı ayrı doldurulur:
 *   - Yorum metni  -> "Comment" (görünen/dahili ad eşleşmesiyle bulunur; hiç
 *     bulunamazsa Title'a düşülür).
 *   - Gönderen kimliği -> "Worker". Bu sütun GERÇEK bir Kişi/Grup alanı
 *     olabileceği gibi düz metin de olabilir (TypeAsString'e bakılarak
 *     ayrımı yapılır): Kişi/Grup ise oturum sahibinin ID'siyle ("WorkerId"),
 *     düz metinse hesaptan türetilen kullanıcı adıyla ("taylan.toptas" gibi,
 *     e-postanın @ öncesi kısmı) dolduruluyor — kullanıcının formda yazdığı
 *     serbest metin ASLA bu alana yazılmaz.
 *   - Tarih -> "Tarih". Site Bölgesel Ayarları UTC+3 (İstanbul) yerine ham
 *     UTC gösteriyorsa ekranda doğru saat çıksın diye getTurkeyTimeAsUtcIso()
 *     ile telafi edilmiş bir değer gönderilir.
 *
 * Sütunların dahili adı tahminimizle uyuşmuyorsa (SharePoint bir sütun
 * yeniden adlandırıldığında ilk oluşturulduğu andaki dahili adı korur),
 * listenin gerçek şeması (/fields) üzerinden görünen ada (Title) göre
 * eşleştirme yapılır.
 */
export const submitFeedback = async (context: WebPartContext, message: string): Promise<IFeedbackSubmitResult> => {
    const webUrl = context.pageContext.web.absoluteUrl;
    const endpoint = `${webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(FEEDBACK_LIST_TITLE)}')/items`;

    const postItem = (fields: Record<string, unknown>): Promise<SPHttpClientResponse> =>
        context.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, {
            headers: {
                Accept: 'application/json;odata=nometadata',
                'Content-type': 'application/json;odata=nometadata'
            },
            body: JSON.stringify(fields)
        });

    try {
        const fieldsMeta = await getListFieldsMeta(context, FEEDBACK_LIST_TITLE);

        const byInternalName = (name: string): ISPListField | undefined =>
            fieldsMeta.filter((f) => f.InternalName === name)[0];
        // Boşluk/büyük-küçük harf farklarına (ör. "Worker ", "WORKER") karşı
        // toleranslı, görünen ada (Title) göre eşleştirme.
        const byDisplayTitle = (title: string): ISPListField | undefined =>
            fieldsMeta.filter((f) => f.Title.trim().toLowerCase() === title.trim().toLowerCase())[0];

        // Yorum metninin yazılacağı GERÇEK sütun.
        const commentFieldName =
            (byInternalName(FEEDBACK_COMMENT_FIELD) ?? byDisplayTitle('Comment'))?.InternalName ?? 'Title';

        // Gönderen kimliği için kullanılacak sütun (varsa). İsim eşleşmesi
        // (ne dahili ad "Worker" ne görünen ad "Worker") başarısız olursa SON
        // ÇARE olarak şemadaki ilk Kişi/Grup (TypeAsString === 'User') alanına
        // düşülür — Author/Editor (Oluşturan/Değiştiren) hariç.
        const workerField =
            byInternalName(FEEDBACK_PERSON_FIELD) ??
            byDisplayTitle('Worker') ??
            fieldsMeta.filter(
                (f) => f.TypeAsString === 'User' && f.InternalName !== 'Author' && f.InternalName !== 'Editor'
            )[0];

        // Tarih sütunu: önce tahmini ad, sonra görünen ad, sonra TypeAsString
        // ile otomatik algılama.
        const dateField =
            byInternalName(FEEDBACK_DATE_FIELD) ??
            byDisplayTitle('Tarih') ??
            fieldsMeta.filter(
                (f) => f.TypeAsString === 'DateTime' && f.InternalName !== 'Created' && f.InternalName !== 'Modified'
            )[0];

        console.info(
            `[SharePointService] Feedback alan eşlemesi — Yorum sütunu: "${commentFieldName}", ` +
            `Gönderen (Worker) sütunu: ${workerField ? `"${workerField.InternalName}" (görünen ad: "${workerField.Title}", tip: ${workerField.TypeAsString})` : 'BULUNAMADI'}, ` +
            `Tarih sütunu: ${dateField ? `"${dateField.InternalName}"` : 'BULUNAMADI'}.`
        );

        const baseFields: Record<string, unknown> = { [commentFieldName]: message };

        if (dateField) {
            baseFields[dateField.InternalName] = new Date().toISOString();
        }

        // Title zorunlu bir alansa ve mesaj başka bir sütuna (Comment) gittiyse,
        // Title boş kalıp SharePoint'in "required field" hatası vermesini
        // önlemek için kısa bir özetle dolduruyoruz.
        const titleField = byInternalName('Title');
        if (titleField?.Required && commentFieldName !== 'Title') {
            baseFields.Title = message.length > 80 ? `${message.substring(0, 77)}...` : message;
        }

        const username = getUsernameFromContext(context);

        let response: SPHttpClientResponse;
        let attemptLabel: string;

        if (workerField) {
            if (workerField.TypeAsString === 'User') {
                const userId = await getCurrentUserId(context);
                if (userId !== undefined) {
                    // 1) Tekli değer biçimi — "Allow multiple selections: No" olan
                    // Kişi/Grup alanlarının standart REST biçimi.
                    response = await postItem({ ...baseFields, [`${workerField.InternalName}Id`]: userId });
                    attemptLabel = `${workerField.InternalName}Id (Kişi/Grup alanı — tekli)`;

                    if (!response.ok) {
                        const singleAttemptDetail = await extractSPErrorDetail(response);
                        console.warn(
                            `[SharePointService] Feedback "${workerField.InternalName}Id" (tekli) ile gönderilemedi ` +
                            `(${singleAttemptDetail}); çoklu-değer biçimi deneniyor.`
                        );

                        // 2) Çoklu değer biçimi — alan "Allow multiple selections: Yes"
                        // ile oluşturulmuşsa REST bir "results" dizisi bekler; tekli
                        // sayı gönderilirse hata vermeden sessizce boş kalabiliyor.
                        response = await postItem({
                            ...baseFields,
                            [`${workerField.InternalName}Id`]: { results: [userId] }
                        });
                        attemptLabel = `${workerField.InternalName}Id (Kişi/Grup alanı — çoklu)`;

                        if (!response.ok) {
                            const multiAttemptDetail = await extractSPErrorDetail(response);
                            console.warn(
                                `[SharePointService] Feedback "${workerField.InternalName}Id" (çoklu) ile de ` +
                                `gönderilemedi (${multiAttemptDetail}); "${workerField.InternalName}" düz metin ` +
                                'alanıyla tekrar deneniyor.'
                            );
                            response = await postItem({ ...baseFields, [workerField.InternalName]: username });
                            attemptLabel = `${workerField.InternalName} (metin alanı, Id denemeleri başarısız olduktan sonra)`;
                        }
                    }
                } else {
                    response = await postItem({ ...baseFields, [workerField.InternalName]: username });
                    attemptLabel = `${workerField.InternalName} (metin alanı — kullanıcı ID alınamadı)`;
                }
            } else {
                response = await postItem({ ...baseFields, [workerField.InternalName]: username });
                attemptLabel = `${workerField.InternalName} (düz metin — hesaptan türetilen kullanıcı adı)`;
            }
        } else {
            console.warn(
                `[SharePointService] "${FEEDBACK_LIST_TITLE}" listesinde bir "Worker" sütunu bulunamadı — ` +
                'gönderen bilgisi olmadan devam edilecek.'
            );
            response = await postItem(baseFields);
            attemptLabel = 'Gönderen sütunu olmadan';
        }

        if (!response.ok) {
            const detail = await extractSPErrorDetail(response);
            console.error(`[SharePointService] Feedback gönderilemedi (son deneme: ${attemptLabel}) — istek: ${endpoint}`, detail);
            return { success: false, errorMessage: detail };
        }
        console.info(
            `[SharePointService] Feedback gönderildi — gönderen alanı: "${attemptLabel}", kullanıcı adı: "${username}".`
        );
        return { success: true };
    } catch (error) {
        console.error('[SharePointService] Feedback gönderilirken beklenmeyen hata:', error);
        return { success: false, errorMessage: (error as Error).message };
    }
};

export interface ISPWriteResult {
    success: boolean;
    errorMessage?: string;
}

/**
 * "Duyurular" listesine GERÇEK bir öğe ekler (SharePoint REST POST) — yerel/
 * geçici bir state DEĞİLDİR: sayfa yenilense de, başka bir kullanıcı açsa da
 * kalıcı olarak görünür. Tarih/saat kullanıcıdan İSTENMEZ, sunucuya isteğin
 * gönderildiği an (new Date().toISOString()) otomatik yazılır — Duyurular
 * listesinde ayrı bir tarih sütunu yoksa (yalnızca standart "Created" alanı
 * varsa) bu, listenin zaten Created alanına göre sıraladığı davranışla
 * (bkz. getAnnouncements) tutarlıdır ve fazladan bir yazım denemesi yapmaz.
 */
export const createAnnouncement = async (
    context: WebPartContext,
    title: string,
    bodyHtml: string,
    imageFile?: File
): Promise<ISPWriteResult> => {
    const webUrl = context.pageContext.web.absoluteUrl;
    const endpoint = `${webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(ANNOUNCEMENTS_LIST_TITLE)}')/items`;

    try {
        const response = await context.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, {
            headers: {
                Accept: 'application/json;odata=nometadata',
                'Content-type': 'application/json;odata=nometadata'
            },
            body: JSON.stringify({
                Title: title,
                [ANNOUNCEMENT_BODY_FIELD]: bodyHtml
            })
        });

        if (!response.ok) {
            const detail = await extractSPErrorDetail(response);
            console.error(`[SharePointService] Duyuru eklenemedi — istek: ${endpoint}`, detail);
            return { success: false, errorMessage: detail };
        }

        // Görsel OPSİYONEL — asıl öğe zaten kaydedildi (yukarıdaki response.ok).
        // Görsel yükleme başarısız olsa bile bu, "duyuru eklenemedi" anlamına
        // GELMEZ; uploadItemImage kendi içinde hatayı yutar, asla fırlatmaz.
        if (imageFile) {
            const created: { Id: number } = await response.json();
            await uploadItemImage(context, ANNOUNCEMENTS_LIST_TITLE, created.Id, imageFile);
        }

        return { success: true };
    } catch (error) {
        console.error('[SharePointService] Duyuru eklenirken beklenmeyen hata:', error);
        return { success: false, errorMessage: (error as Error).message };
    }
};

/**
 * "Etkinlikler" listesine GERÇEK bir öğe ekler (SharePoint REST POST).
 * Liste standart bir Takvim (Calendar) listesiyse EventDate/EndDate çifti
 * genelde beklenir — bu yüzden önce listenin GERÇEK şeması (/fields)
 * sorgulanıp "EndDate" alanı varsa başlangıçtan 1 saat sonrasıyla otomatik
 * dolduruluyor; yoksa (özel/basit bir liste ise) gönderilmiyor, hataya yol
 * açmıyor.
 */
export const createEvent = async (
    context: WebPartContext,
    title: string,
    startDate: Date,
    description: string,
    location?: string,
    imageFile?: File
): Promise<ISPWriteResult> => {
    const webUrl = context.pageContext.web.absoluteUrl;
    const endpoint = `${webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(EVENTS_LIST_TITLE)}')/items`;

    try {
        const fieldsMeta = await getListFieldsMeta(context, EVENTS_LIST_TITLE);
        const hasField = (name: string): boolean => fieldsMeta.some((f) => f.InternalName === name);

        const fields: Record<string, unknown> = {
            Title: title,
            [EVENT_DATE_FIELD]: startDate.toISOString()
        };

        if (hasField('EndDate')) {
            const endDate = new Date(startDate.getTime() + 60 * 60 * 1000);
            fields.EndDate = endDate.toISOString();
        }
        if (location && hasField(EVENT_LOCATION_FIELD)) {
            fields[EVENT_LOCATION_FIELD] = location;
        }
        // "Etkinlik Detayı" için bu listede özel bir alan tanımlı değil (sadece
        // Title/EventDate/Location okunuyor, bkz. getUpcomingEvents) — açıklama
        // metni, alan gerçekten varsa "Description"a, yoksa sessizce atlanır.
        if (description && hasField('Description')) {
            fields.Description = description;
        }

        const response = await context.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, {
            headers: {
                Accept: 'application/json;odata=nometadata',
                'Content-type': 'application/json;odata=nometadata'
            },
            body: JSON.stringify(fields)
        });

        if (!response.ok) {
            const detail = await extractSPErrorDetail(response);
            console.error(`[SharePointService] Etkinlik eklenemedi — istek: ${endpoint}`, detail);
            return { success: false, errorMessage: detail };
        }

        // Görsel OPSİYONEL — bkz. createAnnouncement'taki aynı yorum.
        if (imageFile) {
            const created: { Id: number } = await response.json();
            await uploadItemImage(context, EVENTS_LIST_TITLE, created.Id, imageFile);
        }

        return { success: true };
    } catch (error) {
        console.error('[SharePointService] Etkinlik eklenirken beklenmeyen hata:', error);
        return { success: false, errorMessage: (error as Error).message };
    }
};

/**
 * Verilen listeden TEK bir öğeyi kalıcı olarak siler (SharePoint REST POST +
 * "X-HTTP-Method: DELETE" override — tarayıcılar/SPHttpClient doğrudan DELETE
 * body/header kombinasyonunu her zaman güvenilir göndermediği için bu,
 * SharePoint REST API'sinin resmi önerdiği yöntemdir). "IF-MATCH: *" ile
 * eşzamanlılık (concurrency) kontrolü atlanır — burada bir öğenin son halini
 * koruma değil, silme amaçlanıyor.
 */
const deleteListItem = async (context: WebPartContext, listTitle: string, itemId: number): Promise<ISPWriteResult> => {
    const webUrl = context.pageContext.web.absoluteUrl;
    const endpoint = `${webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items(${itemId})`;

    try {
        const response = await context.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, {
            headers: {
                Accept: 'application/json;odata=nometadata',
                'IF-MATCH': '*',
                'X-HTTP-Method': 'DELETE'
            }
        });

        if (!response.ok) {
            const detail = await extractSPErrorDetail(response);
            console.error(`[SharePointService] "${listTitle}" öğesi (${itemId}) silinemedi — istek: ${endpoint}`, detail);
            return { success: false, errorMessage: detail };
        }

        return { success: true };
    } catch (error) {
        console.error(`[SharePointService] "${listTitle}" öğesi silinirken beklenmeyen hata:`, error);
        return { success: false, errorMessage: (error as Error).message };
    }
};

/** "Duyurular" listesinden tek bir duyuruyu kalıcı olarak siler. */
export const deleteAnnouncement = async (context: WebPartContext, itemId: number): Promise<ISPWriteResult> =>
    deleteListItem(context, ANNOUNCEMENTS_LIST_TITLE, itemId);

/** "Etkinlikler" listesinden tek bir etkinliği kalıcı olarak siler. */
export const deleteEvent = async (context: WebPartContext, itemId: number): Promise<ISPWriteResult> =>
    deleteListItem(context, EVENTS_LIST_TITLE, itemId);
