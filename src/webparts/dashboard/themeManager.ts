import * as React from 'react';

export type TimeBucket = 'morning' | 'day' | 'evening';

export interface IHeaderVariant {
    /** Cam panelin zemini — üç zaman diliminde de FARKLI (beyaz/krem/lacivert), sadece tint değil. */
    glassBackground: string;
    glassBorder: string;
    /** "GÜNAYDIN" gibi üst etiketin rengi. */
    mutedText: string;
    /** İsim (kalın, "Hoş geldin, Ad") rengi. */
    nameText: string;
    clockText: string;
    clockSecondsText: string;
    dateText: string;
    /** Fluent ikon adı — selamlama etiketinin yanındaki ikon. */
    greetingIcon: string;
    /** Süsleme (ışıltı) karakteri — zaman dilimine göre farklı bir "kimlik". */
    decorationChar: string;
    decorationColor: string;
}

export interface ITimeTheme {
    label: string;
    /** Cam panelin arkasında hafifçe süzülen, dağınık "mesh" gradyanının 5 durağı. */
    meshColors: [string, string, string, string, string];
    /** Selamlama/saat metinlerinin rengi — header'ın cam zemininde okunaklı bir ton. */
    text: string;
    /** Sayfanın TAMAMININ (widget grid'inin arkası) 3 duraklı zemin degradesi. */
    pageBackground: [string, string, string];
    /** Karşılama header'ının zaman dilimine göre TAMAMEN farklı kimliği (renk + ikon + süsleme). */
    header: IHeaderVariant;
}

/**
 * Zaman bazlı sayfa zemini + KARŞILAMA HEADER'I artık üç zaman diliminde
 * BİLİNÇLİ OLARAK BİRBİRİNDEN TAMAMEN FARKLI: sabah berrak/soğuk mavi-beyaz
 * cam, öğlen sıcak/krem "relax" cam (mavi durağı YOK — morning'den kasıtlı
 * olarak uzak), akşam ise gerçekten KOYU lacivert bir cam (beyaz camın
 * zamana göre sadece tonlanması değil). Her üçü de kendi ikonuna (Sunny /
 * Coffee / ClearNight) ve kendi süsleme rengine/karakterine sahip — bkz.
 * WelcomeHeader.tsx'in `header` alanlarını nasıl kullandığı. Widget kartları
 * (theme.ts → yorpasTheme) BİLİNÇLİ OLARAK sabit/açık kalıyor; SAYFA ZEMİNİ
 * (Dashboard.module.scss → .dashboardRoot) pageBackground'dan besleniyor ve
 * sabah/öğle açık, akşam GERÇEKTEN koyu.
 */
const THEMES: Record<TimeBucket, ITimeTheme> = {
    morning: {
        label: 'Fresh Morning',
        // 08:00-12:00 — berrak, soğuk kurumsal mavi. Neredeyse beyaza yakın
        // duraklar + belirgin bir gök mavisi durağı ("uyanış" hissi).
        meshColors: ['#EAF2FC', '#F7FBFF', '#D6E7FA', '#B7D6F3', '#E3EFFB'],
        text: '#334155',
        pageBackground: ['#FFFFFF', '#EAF3FC', '#FFFFFF'],
        header: {
            // Beyaz/berrak cam — sabahın "temiz sayfa" hissi.
            glassBackground: 'rgba(255, 255, 255, 0.72)',
            glassBorder: 'rgba(255, 255, 255, 0.5)',
            mutedText: '#64748B',
            nameText: '#0F172A',
            clockText: '#1E293B',
            clockSecondsText: 'rgba(30,41,59,0.45)',
            dateText: 'rgba(51,65,85,0.72)',
            greetingIcon: 'Sunny',
            decorationChar: '✦',
            decorationColor: '#F0A868'
        }
    },
    day: {
        label: 'Relax Afternoon',
        // 12:01-16:00 — ÖNCEKİ HATA: bu palette mavi durakları (EAF2FC,
        // F7FBFF, D6E7FA) sızıyordu, yani "sabah"tan yeterince ayrışmıyordu.
        // Artık TAMAMEN sıcak/krem bir aile (mavi durağı YOK) — "relax"
        // hissi versin diye camın kendisi de krem tonunda (glassBackground),
        // sabahın beyaz/berrak camından kasıtlı olarak uzak.
        meshColors: ['#FDF3E1', '#FFFDF8', '#F7E2C2', '#F0D6A8', '#FBEEDA'],
        text: '#5C4632',
        pageBackground: ['#FFFEFA', '#F6EDD8', '#FFFFFF'],
        header: {
            glassBackground: 'rgba(255, 247, 235, 0.78)',
            glassBorder: 'rgba(255, 241, 219, 0.55)',
            mutedText: '#8A6D4E',
            nameText: '#3E2F20',
            clockText: '#4A3624',
            clockSecondsText: 'rgba(74,54,36,0.45)',
            dateText: 'rgba(90,70,50,0.72)',
            greetingIcon: 'Coffee',
            decorationChar: '❀',
            decorationColor: '#C98A54'
        }
    },
    evening: {
        label: 'Evening Corporate',
        // 16:01-07:59 — ÖNCEKİ HATA: burada turuncu→mor→mavi bir "gün batımı"
        // karışımı vardı; marka mavisiyle alakasız, tutarsız ve "kurumsal"
        // olmaktan uzak duruyordu. Artık aynı mavi ailenin daha doygun/derin
        // durakları kullanılıyor. Header'ın camı da artık GERÇEKTEN koyu
        // lacivert (glassBackground) — sabah/öğlenin aksine beyaz/krem bir
        // camın altında değil, kendisi karanlık.
        meshColors: ['#D9E6F7', '#C3D9F2', '#EAF2FC', '#9FBEE6', '#7C9FD4'],
        text: '#E2E8F0',
        pageBackground: ['#0B1A2E', '#122A47', '#0B1A2E'],
        header: {
            glassBackground: 'rgba(15, 23, 42, 0.82)',
            glassBorder: 'rgba(148, 163, 184, 0.25)',
            mutedText: 'rgba(226,232,240,0.65)',
            nameText: '#FFFFFF',
            clockText: '#F1F5F9',
            clockSecondsText: 'rgba(241,245,249,0.5)',
            dateText: 'rgba(226,232,240,0.68)',
            greetingIcon: 'ClearNight',
            decorationChar: '✧',
            decorationColor: '#9FBEE6'
        }
    }
};

/**
 * 08:00-12:00 Sabah (açık) · 12:01-16:00 Öğlen (relax/ılık) ·
 * 16:01-07:59 Akşam/Gece (koyu, gece yarısını sarar).
 */
export const getTimeBucket = (date: Date = new Date()): TimeBucket => {
    const totalMinutes = date.getHours() * 60 + date.getMinutes();
    const MORNING_START = 8 * 60;   // 08:00
    const MORNING_END = 12 * 60;    // 12:00 (dahil)
    const DAY_END = 16 * 60;        // 16:00 (dahil)

    if (totalMinutes >= MORNING_START && totalMinutes <= MORNING_END) {
        return 'morning';
    }
    if (totalMinutes > MORNING_END && totalMinutes <= DAY_END) {
        return 'day';
    }
    return 'evening';
};

export const getTimeTheme = (bucket: TimeBucket): ITimeTheme => THEMES[bucket];

// SharePoint sayfasındaki başka bir web part'ın kendi CSS değişkenleriyle
// çakışmasın diye tüm değişken adları bu önekle yazılıyor.
const CSS_VAR_PREFIX = '--yorpas-theme-';

/** :root üzerine o anki zaman dilimine ait CSS değişkenlerini yazar. */
export const applyTimeThemeVariables = (bucket: TimeBucket): void => {
    const theme = getTimeTheme(bucket);
    const root = document.documentElement.style;
    root.setProperty(`${CSS_VAR_PREFIX}mesh-1`, theme.meshColors[0]);
    root.setProperty(`${CSS_VAR_PREFIX}mesh-2`, theme.meshColors[1]);
    root.setProperty(`${CSS_VAR_PREFIX}mesh-3`, theme.meshColors[2]);
    root.setProperty(`${CSS_VAR_PREFIX}mesh-4`, theme.meshColors[3]);
    root.setProperty(`${CSS_VAR_PREFIX}mesh-5`, theme.meshColors[4]);
    root.setProperty(`${CSS_VAR_PREFIX}text`, theme.text);
    root.setProperty(`${CSS_VAR_PREFIX}page-bg-1`, theme.pageBackground[0]);
    root.setProperty(`${CSS_VAR_PREFIX}page-bg-2`, theme.pageBackground[1]);
    root.setProperty(`${CSS_VAR_PREFIX}page-bg-3`, theme.pageBackground[2]);
    // Karşılama header'ının zaman dilimine göre TAMAMEN farklı kimliği —
    // WelcomeHeader.tsx bu değişkenleri `var(...)` ile okuyup geçişleri
    // (transition: background/color 1.5s ease) yumuşak hâle getiriyor.
    root.setProperty(`${CSS_VAR_PREFIX}header-glass-bg`, theme.header.glassBackground);
    root.setProperty(`${CSS_VAR_PREFIX}header-glass-border`, theme.header.glassBorder);
    root.setProperty(`${CSS_VAR_PREFIX}header-muted`, theme.header.mutedText);
    root.setProperty(`${CSS_VAR_PREFIX}header-name`, theme.header.nameText);
    root.setProperty(`${CSS_VAR_PREFIX}header-clock`, theme.header.clockText);
    root.setProperty(`${CSS_VAR_PREFIX}header-clock-seconds`, theme.header.clockSecondsText);
    root.setProperty(`${CSS_VAR_PREFIX}header-date`, theme.header.dateText);
    root.setProperty(`${CSS_VAR_PREFIX}header-decoration`, theme.header.decorationColor);
};

/**
 * "ThemeManager" hook'u — bileşen ilk yüklendiğinde VE her dakika (saat
 * başı geçişini kaçırmamak için) o anki zaman dilimini kontrol edip
 * :root CSS değişkenlerini günceller. Değişken adları --yorpas-theme-*
 * öneki taşıdığı için sayfadaki başka bir bileşenle çakışmaz. Bucket
 * gerçekten değişmediği sürece state güncellemesi (ve dolayısıyla
 * yeniden render) TETİKLENMEZ.
 */
export const useTimeAwareTheme = (): TimeBucket => {
    const [bucket, setBucket] = React.useState<TimeBucket>(() => getTimeBucket());

    React.useEffect(() => {
        applyTimeThemeVariables(bucket);
    }, [bucket]);

    React.useEffect(() => {
        const interval = window.setInterval(() => {
            const next = getTimeBucket();
            setBucket((prev) => (prev === next ? prev : next));
        }, 60 * 1000);
        return () => window.clearInterval(interval);
    }, []);

    return bucket;
};
