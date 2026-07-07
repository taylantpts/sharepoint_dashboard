import * as React from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import styles from './Dashboard.module.scss';
import { IDashboardProps } from '../Dashboard.types';

/* ------------------------------------------------------------------ *
 *  KURULUM NOTLARI
 * ------------------------------------------------------------------
 *  1) SharePoint listesi "Duyurular" oluşturun:
 *       - Title            (Tek satır metin)
 *       - Aciklama         (Çok satırlı metin)
 *     Danışmanlarınız/İK bu listeye öğe eklediğinde otomatik burada çıkar.
 *
 *  2) Bir doküman kütüphanesi kullanın (varsayılan "Belgeler" / "Documents").
 *     Danışmanlarınızın bayilere göndereceği kurulum dokümanlarını buraya
 *     yüklemesi yeterli — widget en son değiştirilen 6 dosyayı gösterir.
 *     Kütüphane adınız farklıysa aşağıdaki DOCUMENT_LIBRARY_TITLE sabitini
 *     güncelleyin.
 *
 *  3) package-solution.json içine şu izin isteklerini ekleyip SharePoint
 *     Yönetim Merkezi > Gelişmiş > API Erişimi bölümünden onaylayın,
 *     yoksa Graph çağrıları 401/403 döner:
 *       "webApiPermissionRequests": [
 *         { "resource": "Microsoft Graph", "scope": "Calendars.Read" },
 *         { "resource": "Microsoft Graph", "scope": "User.Read.All" },
 *         { "resource": "Microsoft Graph", "scope": "AuditLog.Read.All" },
 *         { "resource": "Microsoft Graph", "scope": "Tasks.Read" } 
 *       ]
 * ------------------------------------------------------------------ */

const DOCUMENT_LIBRARY_TITLE = 'Belgeler';

// Hava durumu için merkez ofis enlem/boylam bilgisi
const WEATHER_LAT = 41.0082;
const WEATHER_LON = 28.9784;

// Konumlar
const LOCATIONS: { name: string; query: string }[] = [
    { name: 'Merkez Ofis', query: 'Komagene Merkez Ofis, İstanbul' },
    { name: 'Fabrika', query: 'Komagene Fabrika, Kocaeli' },
];

// Hızlı Bağlantılar Menüsü
const QUICK_LINKS = [
    { title: 'Yardım Merkezi', icon: '🛟', url: 'https://destek.komagene.com.tr' },
    { title: 'Sözlük', icon: '📖', url: '/sites/YORPASA/Lists/Komagene_Sozluk/Szlk%20Galerisi.aspx' },
    { title: 'Eğitim Platformu', icon: '🎓', url: 'https://egitim.komagene.com.tr/login/index.php' },
    { title: 'Bayi Sipariş', icon: '🛒', url: 'https://genegenekomagene.com/Admin/Login.aspx' },
    { title: 'Envanter', icon: '📦', url: 'https://asm.komagene.com.tr/' },
    { title: 'Web Portal', icon: '🌐', url: 'https://web.genegenekomagene.com/Admin/Login?ReturnUrl=/Admin/Default' },
    { title: 'Masraff', icon: '💸', url: 'https://admin.masraff.co/#/access/login' },
    { title: 'Entegre Yön.', icon: '🔄', url: 'https://qdms.komagene.com.tr/QDMSNET/BSAT/Logon.aspx' },
    { title: 'Org. Şeması', icon: '👥', url: '/SiteAssets/SitePages/Yörpaş/Organizasyon-Şeması_OCAK2026.pptx' },
];

interface IMeeting { 
    subject: string; 
    start: { dateTime: string }; 
    end: { dateTime: string }; 
    isOnlineMeeting: boolean;
    webLink?: string; 
    onlineMeeting?: { joinUrl: string }; 
}
interface IContact { name: string; title: string; mail: string; department: string; }
interface IAnnouncement { title: string; body: string; date: string; }
interface IDocument { name: string; url: string; editor: string; modified: string; }
interface IWeather { temp: number; emoji: string; label: string; }
interface IBirthday { name: string; date: Date; department: string; } 
interface IToDo { title: string; dueDate: Date | null; isCompleted: boolean; }

export interface IDashboardState {
    meetings: IMeeting[];
    loadingMeetings: boolean;
    errorMeetings: string | null;

    contacts: IContact[];
    departments: string[];
    selectedDepartment: string;
    searchQuery: string;
    loadingContacts: boolean;
    errorContacts: string | null;

    announcements: IAnnouncement[];
    activeAnnouncement: number;
    loadingAnnouncements: boolean;

    documents: IDocument[];
    loadingDocuments: boolean;

    weather: IWeather | null;
    loadingWeather: boolean;
    currentTime: string;

    activeLocation: string;

    birthdays: IBirthday[];
    loadingBirthdays: boolean;

    todos: IToDo[];
    loadingTodos: boolean;
    errorTodos: string | null;
}

function describeWeather(code: number): { emoji: string; label: string } {
    if (code === 0) { return { emoji: '☀️', label: 'Açık' }; }
    if (code <= 3) { return { emoji: '⛅', label: 'Parçalı Bulutlu' }; }
    if (code === 45 || code === 48) { return { emoji: '🌫️', label: 'Sisli' }; }
    if (code >= 51 && code <= 67) { return { emoji: '🌦️', label: 'Yağmurlu' }; }
    if (code >= 71 && code <= 77) { return { emoji: '🌨️', label: 'Karlı' }; }
    if (code >= 80 && code <= 82) { return { emoji: '🌧️', label: 'Sağanak' }; }
    if (code >= 95) { return { emoji: '⛈️', label: 'Fırtınalı' }; }
    return { emoji: '🌡️', label: '—' };
}

function describeError(error: any, context: string): string {
    const status = error?.statusCode || error?.status || error?.code;
    if (status === 401 || status === 403 || status === 'Forbidden' || status === 'Unauthorized') {
        return `${context}: Erişim reddedildi. SharePoint Yönetim Merkezi > Gelişmiş > API Erişimi bölümünden ilgili Graph izninin onaylandığından emin olun.`;
    }
    if (status === 404) {
        return `${context}: Bulunamadı (404). Liste/kütüphane adı doğru mu kontrol edin.`;
    }
    if (error?.message && /InvalidAuthenticationToken|token/i.test(error.message)) {
        return `${context}: Oturum/izin sorunu. Sayfayı yenileyip tekrar deneyin.`;
    }
    return `${context}: Veri alınamadı.`;
}

function getInitials(fullName: string): string {
    return fullName.split(' ').map((p) => p.charAt(0)).join('').slice(0, 2).toUpperCase();
}

function fileIcon(fileName: string): string {
    const ext = (fileName.split('.').pop() || '').toLowerCase();
    if (ext === 'pdf') { return '📕'; }
    if (ext === 'xlsx' || ext === 'xls') { return '📗'; }
    if (ext === 'pptx' || ext === 'ppt') { return '📘'; }
    if (ext === 'docx' || ext === 'doc') { return '📄'; }
    return '🗂️';
}

const AVATAR_TONES = [styles.tone1, styles.tone2, styles.tone3, styles.tone4];

export default class Dashboard extends React.Component<IDashboardProps, IDashboardState> {
    private clockTimer: number | undefined;
    private bleedTimer: number | undefined;
    private containerRef = React.createRef<HTMLDivElement>();

    constructor(props: IDashboardProps) {
        super(props);
        this.state = {
            meetings: [], loadingMeetings: true, errorMeetings: null,
            contacts: [], departments: ['Tümü'], selectedDepartment: 'Tümü', searchQuery: '', loadingContacts: true, errorContacts: null,
            announcements: [], activeAnnouncement: 0, loadingAnnouncements: true,
            documents: [], loadingDocuments: true,
            weather: null, loadingWeather: true,
            currentTime: Dashboard.formatTime(new Date()),
            activeLocation: LOCATIONS[0].name,
            birthdays: [], loadingBirthdays: true,
            todos: [], loadingTodos: true, errorTodos: null,
        };
    }

    public componentDidMount(): void {
        this.fetchMeetings().catch(() => undefined);
        this.fetchContacts().catch(() => undefined);
        this.fetchAnnouncements().catch(() => undefined);
        this.fetchDocuments().catch(() => undefined);
        this.fetchWeather().catch(() => undefined);
        this.fetchBirthdays().catch(() => undefined); 
        this.fetchTodos().catch(() => undefined);

        this.clockTimer = window.setInterval(() => {
            this.setState({ currentTime: Dashboard.formatTime(new Date()) });
        }, 30000);

        this.applyFullBleed();
        window.addEventListener('resize', this.applyFullBleed);
        this.bleedTimer = window.setTimeout(this.applyFullBleed, 400);
    }

    public componentWillUnmount(): void {
        if (this.clockTimer) { window.clearInterval(this.clockTimer); }
        if (this.bleedTimer) { window.clearTimeout(this.bleedTimer); }
        window.removeEventListener('resize', this.applyFullBleed);
    }

    private applyFullBleed = (): void => {
        const el = this.containerRef.current;
        if (!el) { return; }
        el.style.marginLeft = '';
        el.style.marginRight = '';
        el.style.width = '';
        const rect = el.getBoundingClientRect();
        const viewportWidth = document.documentElement.clientWidth;
        el.style.width = `${viewportWidth}px`;
        el.style.marginLeft = `${-rect.left}px`;
    };

    private static formatTime(date: Date): string {
        return date.toLocaleTimeString('tr-TR', { hour: '2-digit', minute: '2-digit' });
    }

    private static formatDate(date: Date): string {
        return date.toLocaleDateString('tr-TR', { day: 'numeric', month: 'long', weekday: 'long' });
    }

    /* ---------------------------- Graph: Toplantılar ---------------------------- */
    private fetchMeetings = async (): Promise<void> => {
        try {
            const client = await this.props.context.msGraphClientFactory.getClient('3');
            const start = new Date();
            const end = new Date(Date.now() + 86400000);
            const response = await client
                .api(`/me/calendarView?startDateTime=${start.toISOString()}&endDateTime=${end.toISOString()}`)
                .select('subject,start,end,isOnlineMeeting,webLink,onlineMeeting')
                .orderby('start/dateTime')
                .top(5)
                .get();
            this.setState({ meetings: response.value || [], loadingMeetings: false });
        } catch (error) {
            console.error('Toplantılar alınamadı:', error);
            this.setState({ loadingMeetings: false, errorMeetings: describeError(error, 'Toplantılar') });
        }
    };

    /* ---------------------------- Graph: Şirket Rehberi (Güvenli Sorgu) ---------------------------- */
    private fetchContacts = async (): Promise<void> => {
        try {
            const client = await this.props.context.msGraphClientFactory.getClient('3');
            
            const response = await client
                .api('/users')
                .select('displayName,jobTitle,mail,department')
                .filter('accountEnabled eq true')
                .top(999)
                .get();

            const activeContacts: IContact[] = (response.value || [])
                .filter((u: any) => u.mail)
                .sort((a: any, b: any) => a.displayName.localeCompare(b.displayName, 'tr'))
                .map((u: any) => ({
                    name: u.displayName,
                    title: u.jobTitle || '—',
                    mail: u.mail,
                    department: u.department || 'Diğer',
                }));

            const departments = ['Tümü', ...Array.from(new Set(activeContacts.map((c) => c.department))).sort((a, b) => a.localeCompare(b, 'tr'))];

            this.setState({ contacts: activeContacts, departments, loadingContacts: false });
        } catch (error) {
            console.error('Şirket rehberi alınamadı:', error);
            this.setState({ loadingContacts: false, errorContacts: describeError(error, 'Şirket rehberi') });
        }
    };

    /* ---------------------------- SP REST: Duyurular ---------------------------- */
    private fetchAnnouncements = async (): Promise<void> => {
        try {
            const webUrl = this.props.context.pageContext.web.absoluteUrl;
            const endpoint = `${webUrl}/_api/web/lists/getbytitle('Duyurular')/items` +
                `?$select=Title,Aciklama,Created&$orderby=Created desc&$top=6`;
            const res: SPHttpClientResponse = await this.props.context.spHttpClient.get(
                endpoint,
                SPHttpClient.configurations.v1,
                { headers: { Accept: 'application/json;odata=nometadata' } }
            );
            if (!res.ok) { throw new Error(`Duyurular listesi okunamadı (${res.status})`); }
            const data = await res.json();
            const items: IAnnouncement[] = (data.value || []).map((it: any) => ({
                title: it.Title,
                body: it.Aciklama || '',
                date: it.Created,
            }));
            this.setState({ announcements: items, loadingAnnouncements: false });
        } catch (error) {
            console.error('Duyurular alınamadı:', error);
            this.setState({ loadingAnnouncements: false });
        }
    };

    /* ---------------------------- SP REST: Belgeler ---------------------------- */
    private fetchDocuments = async (): Promise<void> => {
        try {
            const webUrl = this.props.context.pageContext.web.absoluteUrl;
            const endpoint = `${webUrl}/_api/web/lists/getbytitle('${DOCUMENT_LIBRARY_TITLE}')/items` +
                `?$select=FileLeafRef,FileRef,Modified,Editor/Title&$expand=Editor` +
                `&$filter=FSObjType eq 0&$orderby=Modified desc&$top=6`;
            const res: SPHttpClientResponse = await this.props.context.spHttpClient.get(
                endpoint,
                SPHttpClient.configurations.v1,
                { headers: { Accept: 'application/json;odata=nometadata' } }
            );
            if (!res.ok) { throw new Error(`Belge kütüphanesi okunamadı (${res.status})`); }
            const data = await res.json();
            const items: IDocument[] = (data.value || []).map((it: any) => ({
                name: it.FileLeafRef,
                url: `${window.location.origin}${it.FileRef}`,
                editor: it.Editor?.Title || '—',
                modified: it.Modified,
            }));
            this.setState({ documents: items, loadingDocuments: false });
        } catch (error) {
            console.error('Belgeler alınamadı:', error);
            this.setState({ loadingDocuments: false });
        }
    };

    /* ---------------------------- SP REST: Doğum Günleri ---------------------------- */
    private fetchBirthdays = async (): Promise<void> => {
        try {
            const webUrl = this.props.context.pageContext.web.absoluteUrl;
            
            const endpoint = `${webUrl}/_api/web/lists/getbytitle('Birthdays')/items?$select=Title,field_2&$top=200`;
            const res: SPHttpClientResponse = await this.props.context.spHttpClient.get(
                endpoint,
                SPHttpClient.configurations.v1,
                { headers: { Accept: 'application/json;odata=nometadata' } }
            );
            
            if (!res.ok) { throw new Error('Doğum günleri listesi okunamadı.'); }
            const data = await res.json();
            
            const today = new Date();
            const tMonth = today.getMonth();
            const tDay = today.getDate();

            let items: IBirthday[] = (data.value || []).map((it: any) => {
                if (!it.field_2) return null;

                const bDate = new Date(it.field_2);
                const bMonth = bDate.getMonth();
                const bDay = bDate.getDate();
                
                bDate.setFullYear(today.getFullYear());
                
                if (bMonth < tMonth || (bMonth === tMonth && bDay < tDay)) {
                    bDate.setFullYear(today.getFullYear() + 1);
                }
                
                return {
                    name: it.Title, 
                    department: 'İyi Ki Doğdun! 🎉', 
                    date: bDate
                };
            }).filter(Boolean); 

            items.sort((a: any, b: any) => a.date.getTime() - b.date.getTime());
            
            this.setState({ birthdays: items.slice(0, 5), loadingBirthdays: false });
        } catch (error) {
            console.error('Doğum günleri alınamadı:', error);
            this.setState({ loadingBirthdays: false });
        }
    };

    /* ---------------------------- Graph: To-Do (Görevlerim) ---------------------------- */
    private fetchTodos = async (): Promise<void> => {
        try {
            const client = await this.props.context.msGraphClientFactory.getClient('3');
            
            const listsRes = await client.api('/me/todo/lists').get();
            const lists = listsRes.value || [];
            
            if (lists.length === 0) {
                this.setState({ todos: [], loadingTodos: false });
                return;
            }

            const defaultList = lists.find((l: any) => l.wellknownListName === 'defaultList') || lists[0];

            const tasksRes = await client.api(`/me/todo/lists/${defaultList.id}/tasks`)
                .filter("status ne 'completed'")
                .select('title,status,dueDateTime')
                .top(5)
                .get();

            const items: IToDo[] = (tasksRes.value || []).map((t: any) => ({
                title: t.title,
                dueDate: t.dueDateTime ? new Date(t.dueDateTime.dateTime) : null,
                isCompleted: t.status === 'completed'
            }));

            this.setState({ todos: items, loadingTodos: false });
        } catch (error) {
            console.error('Görevler alınamadı:', error);
            this.setState({ loadingTodos: false, errorTodos: describeError(error, 'Görevlerim') });
        }
    };

    /* ---------------------------- Hava Durumu ---------------------------- */
    private fetchWeather = async (): Promise<void> => {
        try {
            const url = `https://api.open-meteo.com/v1/forecast?latitude=${WEATHER_LAT}&longitude=${WEATHER_LON}` +
                `&current=temperature_2m,weather_code&timezone=Europe%2FIstanbul`;
            const res = await fetch(url);
            if (!res.ok) { throw new Error('Hava durumu alınamadı'); }
            const data = await res.json();
            const { emoji, label } = describeWeather(data.current.weather_code);
            this.setState({
                weather: { temp: Math.round(data.current.temperature_2m), emoji, label },
                loadingWeather: false,
            });
        } catch (error) {
            console.error('Hava durumu alınamadı:', error);
            this.setState({ loadingWeather: false });
        }
    };

    /* ---------------------------- UI Event Handlers ---------------------------- */
    private goToAnnouncement = (index: number): void => {
        const total = this.state.announcements.length;
        if (total === 0) { return; }
        this.setState({ activeAnnouncement: ((index % total) + total) % total });
    };

    private onSearchChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
        this.setState({ searchQuery: e.target.value });
    };

    private onDepartmentChange = (e: React.ChangeEvent<HTMLSelectElement>): void => {
        this.setState({ selectedDepartment: e.target.value });
    };

    private setLocation = (name: string): void => {
        this.setState({ activeLocation: name });
    };

    private get filteredContacts(): IContact[] {
        const { contacts, searchQuery, selectedDepartment } = this.state;
        const q = searchQuery.trim().toLowerCase();
        return contacts.filter((c) => {
            const matchesDept = selectedDepartment === 'Tümü' || c.department === selectedDepartment;
            const matchesQuery = !q ||
                c.name.toLowerCase().indexOf(q) !== -1 ||
                c.title.toLowerCase().indexOf(q) !== -1 ||
                c.department.toLowerCase().indexOf(q) !== -1;
            return matchesDept && matchesQuery;
        });
    }

    /* ---------------------------- Destek Bileti Yönlendirmesi ---------------------------- */
    private handleTicketClick = (e: React.MouseEvent<HTMLAnchorElement>): void => {
        e.preventDefault(); 
        
        const email = "destek@komagene.com.tr";
        const desktopAppUrl = `mailto:${email}`; 
        const webFallbackUrl = `https://outlook.office.com/mail/deeplink/compose?to=${email}`;

        window.location.href = desktopAppUrl;

        setTimeout(() => {
            if (document.hasFocus()) {
                window.open(webFallbackUrl, '_blank');
            }
        }, 800);
    };

    public render(): React.ReactElement<IDashboardProps> {
        const {
            meetings, loadingMeetings, errorMeetings,
            departments, selectedDepartment, searchQuery, loadingContacts, errorContacts,
            announcements, activeAnnouncement, loadingAnnouncements,
            documents, loadingDocuments,
            weather, loadingWeather, currentTime,
            activeLocation,
            birthdays, loadingBirthdays,
            todos, loadingTodos, errorTodos // Eksik olan To-Do verileri eklendi
        } = this.state;

        const filteredContacts = this.filteredContacts;
        const announcement = announcements[activeAnnouncement];
        const currentLocation = LOCATIONS.find((l) => l.name === activeLocation) || LOCATIONS[0];
        const firstName = this.props.userDisplayName.split(' ')[0];

        return (
            <div className={styles.dashboardContainer} ref={this.containerRef}>
                <div className={styles.bentoGrid}>

                    {/* HERO */}
                    <div className={`${styles.card} ${styles.hero}`}>
                        <div className={styles.heroText}>
                            <span className={styles.heroEyebrow}>KOMAGENE İNTRANET</span>
                            <h1>Merhaba, {firstName}! 👋</h1>
                            <p>{Dashboard.formatDate(new Date())} — ihtiyacın olan her şey burada.</p>
                        </div>
                        <div className={styles.heroStatus}>
                            <div className={styles.statusChip}>
                                <span className={styles.statusLabel}>Saat</span>
                                <span className={styles.statusValue}>{currentTime}</span>
                            </div>
                            <div className={styles.statusChip}>
                                <span className={styles.statusLabel}>Hava</span>
                                <span className={styles.statusValue}>
                                    {loadingWeather ? '…' : weather ? `${weather.temp}° ${weather.emoji}` : '—'}
                                </span>
                            </div>
                        </div>
                    </div>
                    
                    {/* HIZLI BAĞLANTILAR */}
                    <div className={`${styles.card} ${styles.spanFull} ${styles.quickLinksCard}`}>
                        <div className={styles.quickLinksContainer}>
                            {QUICK_LINKS.map((link, i) => (
                                <a key={i} href={link.url} target="_blank" rel="noreferrer" className={styles.quickLinkItem}>
                                    <span className={styles.quickLinkIcon}>{link.icon}</span>
                                    <span className={styles.quickLinkText}>{link.title}</span>
                                </a>
                            ))}
                        </div>
                    </div>

                    {/* TOPLANTILARIM */}
                    <div className={`${styles.card} ${styles.span2}`}>
                        <h3><span className={styles.cardIcon}>📅</span>Toplantılarım</h3>
                        {loadingMeetings && <p className={styles.emptyState}>Yükleniyor…</p>}
                        {!loadingMeetings && errorMeetings && (
                            <p className={styles.errorState}>{errorMeetings}</p>
                        )}
                        {!loadingMeetings && !errorMeetings && meetings.length === 0 && (
                            <p className={styles.emptyState}>Bugün planlanmış toplantın yok 🎉</p>
                        )}
                        {!loadingMeetings && !errorMeetings && meetings.length > 0 && (
                            <ul className={styles.meetingList}>
                                {meetings.map((m, i) => {
                                    const utcDate = m.start.dateTime + (m.start.dateTime.includes('Z') ? '' : 'Z');
                                    const startDate = new Date(utcDate);
                                    const formattedTime = startDate.toLocaleTimeString('tr-TR', { hour: '2-digit', minute: '2-digit' });

                                    const meetingLink = m.isOnlineMeeting && m.onlineMeeting?.joinUrl ? m.onlineMeeting.joinUrl : m.webLink;
                                    
                                    return (
                                        <li key={i}>
                                            <span className={styles.meetingTime}>{formattedTime}</span>
                                            
                                            {meetingLink ? (
                                                <a 
                                                    href={meetingLink} 
                                                    target="_blank" 
                                                    rel="noreferrer" 
                                                    className={styles.meetingTitle}
                                                    style={{ textDecoration: 'none', color: 'inherit', cursor: 'pointer', transition: 'color 0.2s ease' }}
                                                    onMouseOver={(e) => e.currentTarget.style.color = '#E31E24'}
                                                    onMouseOut={(e) => e.currentTarget.style.color = 'inherit'}
                                                >
                                                    {m.subject}
                                                </a>
                                            ) : (
                                                <span className={styles.meetingTitle}>{m.subject}</span>
                                            )}

                                            <span className={styles.meetingLocation}>{m.isOnlineMeeting ? 'Teams' : 'Ofis'}</span>
                                        </li>
                                    );
                                })}
                            </ul>
                        )}
                    </div>

                    {/* GÖREVLERİM (TO-DO) - Yeni Eklenen HTML Tasarım Bloğu */}
                    <div className={`${styles.card} ${styles.span2}`}>
                        <h3><span className={styles.cardIcon}>✅</span>Görevlerim</h3>
                        {loadingTodos && <p className={styles.emptyState}>Yükleniyor…</p>}
                        {!loadingTodos && errorTodos && (
                            <p className={styles.errorState}>{errorTodos}</p>
                        )}
                        {!loadingTodos && !errorTodos && todos.length === 0 && (
                            <p className={styles.emptyState}>Harika! Bekleyen göreviniz yok 🎉</p>
                        )}
                        {!loadingTodos && !errorTodos && todos.length > 0 && (
                            <ul className={styles.todoList}>
                                {todos.map((todo, i) => (
                                    <li key={i} className={styles.todoItem}>
                                        <div className={styles.todoCheck}></div>
                                        <span className={styles.todoTitle}>{todo.title}</span>
                                        {todo.dueDate && (
                                            <span className={styles.todoDate}>
                                                {todo.dueDate.toLocaleDateString('tr-TR', { day: 'numeric', month: 'short' })}
                                            </span>
                                        )}
                                    </li>
                                ))}
                            </ul>
                        )}
                        
                        <a 
                            href="https://to-do.office.com/tasks/" 
                            target="_blank" 
                            rel="noreferrer" 
                            style={{ display: 'block', marginTop: '15px', fontSize: '13px', color: '#D9A44A', fontWeight: 'bold', textDecoration: 'none' }}
                        >
                            TÜM GÖREVLERİ GÖR →
                        </a>
                    </div>

                    {/* DUYURULAR */}
                    <div className={`${styles.card} ${styles.span2}`}>
                        <h3><span className={styles.cardIcon}>📢</span>Duyurular</h3>
                        {loadingAnnouncements && <p className={styles.emptyState}>Yükleniyor…</p>}
                        {!loadingAnnouncements && announcements.length === 0 && (
                            <p className={styles.emptyState}>Henüz duyuru yok.</p>
                        )}
                        {!loadingAnnouncements && announcement && (
                            <>
                                <div className={styles.announceBody}>
                                    <p className={styles.announceTitle}>{announcement.title}</p>
                                    <p className={styles.announceText}>{announcement.body}</p>
                                </div>
                                <div className={styles.dots}>
                                    {announcements.map((_, i) => (
                                        <button
                                            key={i}
                                            aria-label={`Duyuru ${i + 1}`}
                                            className={`${styles.dot} ${i === activeAnnouncement ? styles.dotActive : ''}`}
                                            onClick={() => this.goToAnnouncement(i)}
                                        />
                                    ))}
                                    <span className={styles.dotsCounter}>{activeAnnouncement + 1}/{announcements.length}</span>
                                </div>
                            </>
                        )}
                    </div>

                    {/* YENİ PAYLAŞILAN BELGELER */}
                    <div className={`${styles.card} ${styles.span2}`}>
                        <h3><span className={styles.cardIcon}>🗂️</span>Yeni Paylaşılan Belgeler</h3>
                        {loadingDocuments && <p className={styles.emptyState}>Yükleniyor…</p>}
                        {!loadingDocuments && documents.length === 0 && (
                            <p className={styles.emptyState}>Henüz belge paylaşılmamış.</p>
                        )}
                        {!loadingDocuments && documents.length > 0 && (
                            <ul className={styles.fileList}>
                                {documents.map((doc, i) => (
                                    <li key={i}>
                                        <span className={styles.fileIcon}>{fileIcon(doc.name)}</span>
                                        <a href={doc.url} target="_blank" rel="noreferrer" className={styles.fileName}>{doc.name}</a>
                                        <span className={styles.fileMeta}>{doc.editor}</span>
                                    </li>
                                ))}
                            </ul>
                        )}
                        
                        <a 
                            href={`${this.props.context.pageContext.web.absoluteUrl}/${DOCUMENT_LIBRARY_TITLE}`} 
                            target="_blank" 
                            rel="noreferrer" 
                            style={{ display: 'block', marginTop: '15px', fontSize: '13px', color: '#D9A44A', fontWeight: 'bold', textDecoration: 'none' }}
                        >
                            TÜM KLASÖRLERİ GÖR →
                        </a>
                    </div>

                    {/* KONUM / GOOGLE MAPS */}
                    <div className={`${styles.card} ${styles.span2}`}>
                        <h3><span className={styles.cardIcon}>📍</span>Konumlarımız</h3>
                        <div className={styles.locationTabs}>
                            {LOCATIONS.map((loc) => (
                                <button
                                    key={loc.name}
                                    className={`${styles.locationTab} ${activeLocation === loc.name ? styles.locationTabActive : ''}`}
                                    onClick={() => this.setLocation(loc.name)}
                                >
                                    {loc.name}
                                </button>
                            ))}
                        </div>
                        <iframe
                            className={styles.mapFrame}
                            title={currentLocation.name}
                            src={`https://www.google.com/maps?q=${encodeURIComponent(currentLocation.query)}&output=embed`}
                            loading="lazy"
                        />
                        <a
                            className={styles.mapLink}
                            href={`https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(currentLocation.query)}`}
                            target="_blank"
                            rel="noreferrer"
                        >
                            Haritada Aç →
                        </a>
                    </div>

                    {/* ŞİRKET REHBERİ */}
                    <div className={`${styles.card} ${styles.span2}`}>
                        <h3><span className={styles.cardIcon}>👥</span>Şirket Rehberi</h3>
                        <div className={styles.filterRow}>
                            <input
                                type="text"
                                className={styles.contactSearch}
                                placeholder="İsim, unvan veya departman ara…"
                                value={searchQuery}
                                onChange={this.onSearchChange}
                            />
                            <select className={styles.filterSelect} value={selectedDepartment} onChange={this.onDepartmentChange}>
                                {departments.map((d) => <option key={d} value={d}>{d}</option>)}
                            </select>
                        </div>
                        {loadingContacts && <p className={styles.emptyState}>Yükleniyor…</p>}
                        {!loadingContacts && errorContacts && (
                            <p className={styles.errorState}>{errorContacts}</p>
                        )}
                        {!loadingContacts && !errorContacts && (
                            <>
                                <ul className={`${styles.peopleList} ${styles.scrollList}`}>
                                    {filteredContacts.slice(0, 8).map((c, i) => (
                                        <li key={i}>
                                            <a href={`mailto:${c.mail}`} className={`${styles.avatar} ${AVATAR_TONES[i % 4]}`}>
                                                {getInitials(c.name)}
                                            </a>
                                            <span className={styles.peopleInfo}>
                                                <span className={styles.peopleName}>{c.name}</span>
                                                <span className={styles.peopleSub}>{c.title} · {c.department}</span>
                                            </span>
                                        </li>
                                    ))}
                                    {filteredContacts.length === 0 && (
                                        <li className={styles.contactEmpty}>Sonuç bulunamadı</li>
                                    )}
                                </ul>
                                {filteredContacts.length > 8 && (
                                    <p className={styles.dotsCounter}>+{filteredContacts.length - 8} kişi daha — aramayı daraltın</p>
                                )}
                            </>
                        )}
                    </div>

                    {/* DOĞUM GÜNLERİ */}
                    <div className={`${styles.card} ${styles.span2}`}>
                        <h3><span className={styles.cardIcon}>🎂</span>Yaklaşan Doğum Günleri</h3>
                        {loadingBirthdays && <p className={styles.emptyState}>Yükleniyor…</p>}
                        {!loadingBirthdays && birthdays.length === 0 && (
                            <p className={styles.emptyState}>Yakın zamanda listelenmiş bir doğum günü yok.</p>
                        )}
                        {!loadingBirthdays && birthdays.length > 0 && (
                            <ul className={styles.peopleList}>
                                {birthdays.map((b, i) => (
                                    <li key={i}>
                                        <div className={`${styles.avatar} ${AVATAR_TONES[i % 4]}`}>
                                            {getInitials(b.name)}
                                        </div>
                                        <span className={styles.peopleInfo}>
                                            <span className={styles.peopleName}>{b.name}</span>
                                            <span className={styles.peopleSub}>{b.department}</span>
                                        </span>
                                        <span style={{ marginLeft: 'auto', fontFamily: 'monospace', color: '#E31E24', fontWeight: 600, fontSize: '12.5px' }}>
                                            {b.date.toLocaleDateString('tr-TR', { day: 'numeric', month: 'long' })}
                                        </span>
                                    </li>
                                ))}
                            </ul>
                        )}
                    </div>

                    {/* YARDIM MERKEZİ */}
                    <div className={`${styles.card} ${styles.span2} ${styles.ticketCard}`}>
                        
                        <div className={styles.ticketTop}>
                            <h3><span className={styles.cardIcon}>🎫</span> IT Destek Merkezi</h3>
                            <p>Teknik bir sorunla mı karşılaştınız? Hemen bir destek talebi oluşturun, yardımcı olalım.</p>
                        </div>

                        <div className={styles.ticketDivider}></div>

                        <div className={styles.ticketBottom}>
                            <a
                                href="#"
                                onClick={this.handleTicketClick}
                                className={styles.ticketAction}
                            >
                                DESTEK BİLETİ OLUŞTUR →
                            </a>
                        </div>
                        
                    </div>

                </div>
            </div>
        );
    }
}