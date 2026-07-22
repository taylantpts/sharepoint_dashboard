import { WebPartContext } from '@microsoft/sp-webpart-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';

export interface IDirectoryUser {
    id: string;
    displayName: string;
    jobTitle: string;
    department: string;
    mail: string;
    /** Cep telefonu varsa o, yoksa iş telefonlarının ilki — yoksa boş string. */
    phone: string;
}

export interface ITodayEvent {
    id: string;
    subject: string;
    startLabel: string;
    endLabel: string;
    location: string;
}

export interface ITodoTask {
    id: string;
    title: string;
    isCompleted: boolean;
    /** Görevin ait olduğu To Do liste ID'si — tamamlandı işaretlemek için gerekir. */
    listId: string;
}

interface IGraphUser {
    id: string;
    displayName: string;
    jobTitle?: string;
    department?: string;
    mail?: string;
    mobilePhone?: string;
    businessPhones?: string[];
    accountEnabled?: boolean;
    /** On-prem AD'den senkronize edilen tam LDAP DN'i — hangi OU'da olduğunu içerir. */
    onPremisesDistinguishedName?: string;
    signInActivity?: {
        lastSignInDateTime?: string;
        lastNonInteractiveSignInDateTime?: string;
    };
}

// Rehberde ASLA görünmemesi gereken OU — DISABLED_FOLDER içindeki (devre dışı
// bırakılmış/ayrılmış personel) hesapları arama sonuçlarından tamamen eler.
const DISABLED_OU_MARKER = 'ou=disabled_folder';

const isInDisabledOu = (user: IGraphUser): boolean =>
    (user.onPremisesDistinguishedName ?? '').toLowerCase().includes(DISABLED_OU_MARKER);

/**
 * "Boş profil" filtresi — sistem hesapları, dış/guest kullanıcılar ve eksik
 * profiller (departmanı/unvanı hiç doldurulmamış hesaplar) rehberde ve arama
 * sonuçlarında GÖRÜNMEMELİ. Bir kullanıcının listelenebilmesi için HEM
 * Department HEM DE jobTitle alanının dolu olması ZORUNLU — biri bile boşsa
 * (veya sadece boşluk karakterlerinden oluşuyorsa) o hesap tamamen elenir.
 */
const hasCompleteProfile = (user: IGraphUser): boolean =>
    !!user.department?.trim() && !!user.jobTitle?.trim();

const RECENT_ACTIVITY_WINDOW_DAYS = 15;

interface IGraphEvent {
    id: string;
    subject: string;
    start: { dateTime: string };
    end: { dateTime: string };
    location?: { displayName?: string };
}

interface IGraphTodoTask {
    id: string;
    title: string;
    status: string;
}

const timeLabel = (iso: string): string =>
    new Date(iso).toLocaleTimeString('tr-TR', { hour: '2-digit', minute: '2-digit' });

const toDirectoryUsers = (users: IGraphUser[]): IDirectoryUser[] =>
    users
        .filter((u) => !isInDisabledOu(u))
        .filter(hasCompleteProfile)
        .map((u) => ({
            id: u.id,
            displayName: u.displayName,
            jobTitle: u.jobTitle ?? '',
            department: u.department ?? '',
            mail: u.mail ?? '',
            phone: u.mobilePhone ?? u.businessPhones?.[0] ?? ''
        }));

/**
 * Gelişmiş sorgu: $filter (accountEnabled eq true) + $search + signInActivity
 * $select birlikte. Bunun için tenant'ta User.Read.All/Directory.Read.All ve
 * (15 günlük filtre için) AuditLog.Read.All gerekir — sadece User.Read.Basic.All
 * ile bu istek Graph'tan 400/403 döner ve fırlatılır (çağıran taraf basit
 * aramaya düşer).
 */
const searchDirectoryAdvanced = async (client: MSGraphClientV3, escapedQuery: string): Promise<IDirectoryUser[]> => {
    const response: { value: IGraphUser[] } = await client
        .api('/users')
        .header('ConsistencyLevel', 'eventual')
        .filter('accountEnabled eq true')
        .search(`"displayName:${escapedQuery}"`)
        .select(
            'id,displayName,jobTitle,department,mail,mobilePhone,businessPhones,accountEnabled,' +
            'onPremisesDistinguishedName,signInActivity'
        )
        .top(15)
        .get();

    const users = response.value ?? [];
    const hasActivityData = users.some((u) => !!u.signInActivity?.lastSignInDateTime);

    let activeUsers = users;
    if (hasActivityData) {
        const cutoff = Date.now() - RECENT_ACTIVITY_WINDOW_DAYS * 24 * 60 * 60 * 1000;
        activeUsers = users.filter((u) => {
            const last = u.signInActivity?.lastSignInDateTime;
            if (!last) {
                return false;
            }
            return new Date(last).getTime() >= cutoff;
        });
    } else {
        console.warn(
            '[GraphService] signInActivity verisi alınamadı (AuditLog.Read.All izni gerekebilir) — ' +
            '15 günlük son aktiflik filtresi uygulanmadan tüm aktif hesaplar gösteriliyor.'
        );
    }

    return toDirectoryUsers(activeUsers);
};

/**
 * Basit sorgu: yalnızca $search, $filter/accountEnabled/signInActivity YOK.
 * User.Read.Basic.All gibi kısıtlı izinlerle de çalışır — gelişmiş sorgu
 * yetki hatasıyla düşerse bu, rehberin en azından temel haliyle çalışmasını
 * garanti eden yedek yoldur.
 */
const searchDirectoryBasic = async (client: MSGraphClientV3, escapedQuery: string): Promise<IDirectoryUser[]> => {
    const response: { value: IGraphUser[] } = await client
        .api('/users')
        .header('ConsistencyLevel', 'eventual')
        .search(`"displayName:${escapedQuery}"`)
        .select('id,displayName,jobTitle,department,mail,mobilePhone,businessPhones,onPremisesDistinguishedName')
        .top(15)
        .get();

    return toDirectoryUsers(response.value ?? []);
};

/**
 * Graph /users endpoint'inde isim, unvan veya departmana göre arama yapar.
 * Önce accountEnabled + signInActivity (son 15 gün) filtreli GELİŞMİŞ sorgu
 * denenir; bu sorgu tenant'ın izin seviyesine göre (bkz. searchDirectoryAdvanced
 * yorumu) 400/403 ile başarısız olabilir — bu durumda OTOMATİK olarak filtresiz
 * BASİT bir aramaya düşülür, böylece rehber izin yetersizliğinde de tamamen boş
 * kalmaz. Her iki deneme de başarısız olursa hata konsola detaylı basılıp
 * tekrar fırlatılır — widget katmanı bunu yakalayıp kullanıcıya sade/profesyonel
 * bir uyarı gösterir (bkz. CompanyDirectory.tsx).
 */
export const searchDirectory = async (context: WebPartContext, query: string): Promise<IDirectoryUser[]> => {
    const trimmed = query.trim();
    if (!trimmed) {
        return [];
    }

    const client = await context.msGraphClientFactory.getClient('3');
    const escaped = trimmed.replace(/"/g, '\\"');

    try {
        return await searchDirectoryAdvanced(client, escaped);
    } catch (advancedError) {
        console.warn(
            '[GraphService] Gelişmiş rehber sorgusu (accountEnabled + signInActivity) başarısız, ' +
            'basit aramaya düşülüyor:', advancedError
        );
        try {
            return await searchDirectoryBasic(client, escaped);
        } catch (basicError) {
            console.error('[GraphService] Şirket rehberi araması başarısız (/users):', basicError);
            throw basicError;
        }
    }
};

/**
 * Oturum sahibinin bugünkü takvim etkinliklerini getirir (/me/calendarView).
 */
export const getTodayEvents = async (context: WebPartContext): Promise<ITodayEvent[]> => {
    const client = await context.msGraphClientFactory.getClient('3');

    const now = new Date();
    const startOfDay = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0);
    const endOfDay = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59);

    const response: { value: IGraphEvent[] } = await client
        .api('/me/calendarView')
        .header('Prefer', 'outlook.timezone="Turkey Standard Time"')
        .query({
            startDateTime: startOfDay.toISOString(),
            endDateTime: endOfDay.toISOString()
        })
        .select('id,subject,start,end,location')
        .orderby('start/dateTime')
        .top(10)
        .get();

    const events = response.value ?? [];
    return events.map((e) => ({
        id: e.id,
        subject: e.subject,
        startLabel: timeLabel(e.start.dateTime),
        endLabel: timeLabel(e.end.dateTime),
        location: e.location?.displayName ?? ''
    }));
};

// Görev listesi seçiminde tercih edilen görünen adlar (bulunamazsa ilk liste
// kullanılır). Endpoint tam olarak GraphService dokümantasyonundaki gibidir:
// GET /me/todo/lists  ->  GET /me/todo/lists/{listId}/tasks
const PREFERRED_TASK_LIST_NAMES = ['Tasks', 'Görevler'];

export interface ITodoTasksResult {
    tasks: ITodoTask[];
    /**
     * Dolu ise Graph isteği gerçekten başarısız olmuştur (ör. Tasks.Read izni
     * onaylanmamış) — widget bunu "görev yok" ile KARIŞTIRMAMALI, kullanıcıya
     * görünür bir uyarı göstermeli. Bu, izin/erişim sorunlarının sessizce
     * "boş liste" gibi görünmesini engellemek için var.
     */
    errorDetail?: string;
    /**
     * Seçilen hedef listenin ID'si — liste BOŞ olsa bile (0 görev) doludur.
     * Yeni görev eklerken (createTodoTask) gerekir: tek bir görevin listId'si
     * üzerinden okumak, liste boşken hiç görev olmadığı için mümkün değildir.
     */
    listId?: string;
}

/** Graph SDK hatasından (veya herhangi bir throw edilen değerden) okunabilir bir metin çıkarır. */
const describeGraphError = (error: unknown): string => {
    if (error instanceof Error) {
        return error.message;
    }
    if (typeof error === 'object' && error !== null) {
        const withMessage = error as { message?: string; statusCode?: number };
        if (withMessage.message) {
            return withMessage.statusCode ? `HTTP ${withMessage.statusCode} — ${withMessage.message}` : withMessage.message;
        }
    }
    return String(error);
};

/**
 * Oturum sahibinin Microsoft To Do listesindeki tamamlanmamış görevlerini
 * MSGraphClientV3 ile getirir. Önce /me/todo/lists ile tüm listeler çekilir,
 * displayName'i "Tasks" veya "Görevler" olan liste seçilir (yoksa ilk liste
 * kullanılır); ardından o listenin ID'siyle /me/todo/lists/{listId}/tasks
 * çağrılır.
 *
 * NOT: /me/todo/lists/{id}/tasks üzerinde $filter kullanmak bazı tenant'larda
 * "Invalid request" (HTTP 400) hatası verebiliyor (desteklenen filtrelenebilir
 * alanlar sınırlı). Bu yüzden $filter GÖNDERMİYORUZ — ham listeyi çekip
 * "tamamlanmamış" ayrımını JavaScript tarafında yapıyoruz.
 *
 * Graph isteği başarısız olursa (en sık neden: Tasks.Read izni
 * config/package-solution.json'da tanımlı ama SharePoint Yönetim Merkezi'nde
 * henüz ONAYLANMAMIŞ) errorDetail dolu döner ve hata konsola da basılır —
 * widget bunu "hiç görev yok" durumundan AYRI göstermelidir, aksi halde bir
 * izin sorunu sonsuza kadar sessizce boş liste gibi görünür.
 */
export const getTodoTasks = async (context: WebPartContext): Promise<ITodoTasksResult> => {
    const client = await context.msGraphClientFactory.getClient('3');

    let lists: Array<{ id: string; displayName: string }>;
    try {
        const listsResponse: { value: Array<{ id: string; displayName: string }> } = await client
            .api('/me/todo/lists')
            .get();
        lists = listsResponse.value ?? [];
    } catch (error) {
        const detail = describeGraphError(error);
        console.error('[GraphService] Yapılacaklar LİSTELERİ alınamadı (/me/todo/lists):', error);
        return { tasks: [], errorDetail: `Liste adımı: ${detail}` };
    }

    console.info(`[GraphService] To Do listeleri bulundu (${lists.length}):`, lists.map((l) => l.displayName));

    if (lists.length === 0) {
        return {
            tasks: [],
            errorDetail:
                'Hesabınızda henüz bir Microsoft To Do listesi yok. to-do.office.com veya To Do uygulamasını bir ' +
                'kez açıp en az bir görev eklerseniz burada görünmeye başlar.'
        };
    }

    const preferredList = lists.filter(
        (l) => PREFERRED_TASK_LIST_NAMES.indexOf(l.displayName) !== -1
    )[0];
    const targetList = preferredList ?? lists[0];

    // To Do liste ID'leri base64 tabanlı olup "/", "+", "=" gibi URL'de özel
    // anlamı olan karakterler içerebilir — kodlanmadan doğrudan yola eklenirse
    // Graph "Invalid request" (400) döner. encodeURIComponent ile bu
    // karakterler güvenli hale getiriliyor.
    const encodedListId = encodeURIComponent(targetList.id);

    try {
        const tasksResponse: { value: IGraphTodoTask[] } = await client
            .api(`/me/todo/lists/${encodedListId}/tasks`)
            .get();

        const tasks = tasksResponse.value ?? [];
        console.info(`[GraphService] "${targetList.displayName}" listesinde ${tasks.length} görev bulundu.`);

        return {
            tasks: tasks
                .filter((t) => t.status !== 'completed')
                .slice(0, 10)
                .map((t) => ({
                    id: t.id,
                    title: t.title,
                    isCompleted: t.status === 'completed',
                    listId: targetList.id
                })),
            listId: targetList.id
        };
    } catch (error) {
        const detail = describeGraphError(error);
        console.error(`[GraphService] Görevler alınamadı (/me/todo/lists/${targetList.id}/tasks):`, error);
        return { tasks: [], errorDetail: `Görev adımı ("${targetList.displayName}"): ${detail}`, listId: targetList.id };
    }
};

export interface ITodoTaskActionResult {
    success: boolean;
    errorMessage?: string;
}

/**
 * Widget'taki görevi işaretlemeyi GERÇEK Microsoft To Do hesabına yazar
 * (PATCH /me/todo/lists/{listId}/tasks/{taskId}, status: "completed") —
 * yalnızca yerel/görsel bir işaretleme değildir, to-do.office.com'da ve
 * Outlook/Teams To Do'da da tamamlanmış olarak görünür.
 */
/**
 * Widget'taki "+" butonuyla girilen yeni görevi GERÇEK Microsoft To Do
 * hesabına yazar (POST /me/todo/lists/{listId}/tasks) — to-do.office.com'da
 * ve Outlook/Teams To Do'da da anında görünür, yerel/geçici bir liste
 * öğesi değildir.
 */
export const createTodoTask = async (
    context: WebPartContext,
    listId: string,
    title: string
): Promise<ITodoTaskActionResult> => {
    try {
        const client = await context.msGraphClientFactory.getClient('3');
        await client
            .api(`/me/todo/lists/${encodeURIComponent(listId)}/tasks`)
            .post({ title });
        return { success: true };
    } catch (error) {
        const detail = describeGraphError(error);
        console.error('[GraphService] Görev eklenemedi:', error);
        return { success: false, errorMessage: detail };
    }
};

export const completeTodoTask = async (
    context: WebPartContext,
    listId: string,
    taskId: string
): Promise<ITodoTaskActionResult> => {
    try {
        const client = await context.msGraphClientFactory.getClient('3');
        await client
            .api(`/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}`)
            .patch({ status: 'completed' });
        return { success: true };
    } catch (error) {
        const detail = describeGraphError(error);
        console.error('[GraphService] Görev tamamlandı olarak işaretlenemedi:', error);
        return { success: false, errorMessage: detail };
    }
};
