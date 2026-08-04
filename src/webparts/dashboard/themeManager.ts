import * as React from 'react';

export type TimeBucket = 'morning' | 'day' | 'evening';

export interface IHeaderVariant {
    glassBackground: string;
    glassBorder: string;
    /** "GÜNAYDIN" gibi üst etiketin rengi. */
    mutedText: string;
    /** İsim (kalın, "Hoş geldin, Ad") rengi. */
    nameText: string;
    clockText: string;
    clockSecondsText: string;
    dateText: string;
    /** Fluent ikon adı — selamlama etiketinin yanındaki ikon (saate göre değişir, RENK değişmez). */
    greetingIcon: string;
    /** Süsleme (ışıltı) karakteri — saate göre değişir (RENK değişmez, bkz. decorationColor). */
    decorationChar: string;
    decorationColor: string;
}

export interface ITimeTheme {
    label: string;
    meshColors: [string, string, string, string, string];
    text: string;
    pageBackground: [string, string, string];
    header: IHeaderVariant;
    cardAmbient: string;
}

// ÖNCEKİ HATA (birden fazla iterasyon): sabah/öğle/akşam için üç ayrı renk
// paleti vardı ("Fresh Morning" mavi, "Relax Afternoon" krem, "Evening
// Corporate" lacivert) — kullanıcı bunu "renkli/dağınık" bulup TEK, sabit bir
// KURUMSAL tema istedi (referans: Klarinet Velocity — lacivert/beyaz zemin +
// tek bir vurgu rengi, widget'lara özel renk yok). Bu dosya artık zaman
// dilimine göre HİÇBİR RENK değiştirmiyor — tek bir CORPORATE_PALETTE var.
// Zaman dilimi (TimeBucket) sadece WelcomeHeader'daki selamlama METNİNİ
// ("Günaydın" / "İyi günler" / "İyi akşamlar") ve yanındaki ikonu/süsleme
// karakterini belirlemek için hâlâ kullanılıyor — bu salt kişiselleştirme
// amaçlı bir metin/ikon seçimi, bir RENK teması değil.
const CORPORATE_PALETTE = {
    // Çok hafif, neredeyse fark edilmeyen mavi-gri mesh — Velocity'nin
    // sade beyaz zeminine yakın, dikkat dağıtmayan bir doku.
    meshColors: ['#EEF3FA', '#FAFCFF', '#E6EEF8', '#DCE8F5', '#F2F6FB'] as [string, string, string, string, string],
    text: '#334155',
    pageBackground: ['#FFFFFF', '#F5F8FC', '#FFFFFF'] as [string, string, string],
    header: {
        glassBackground: 'rgba(255, 255, 255, 0.85)',
        glassBorder: 'rgba(226, 232, 240, 0.9)',
        mutedText: '#64748B',
        nameText: '#0F172A',
        clockText: '#1E293B',
        clockSecondsText: 'rgba(30,41,59,0.45)',
        dateText: 'rgba(51,65,85,0.72)',
        // Kurumsal marka mavisiyle (yorpasTheme #0078d4) birebir aynı —
        // widget kartlarındaki ikon rozeti/üst şeritle de aynı ton, tüm
        // portalda TEK bir vurgu rengi.
        decorationColor: '#0078D4'
    },
    // Widget kartlarının zemin tonu — sabit, çok hafif mavi-gri bir "wash".
    cardAmbient: '#EBF1FA'
};

const GREETING_BY_BUCKET: Record<TimeBucket, { label: string; greetingIcon: string; decorationChar: string }> = {
    morning: { label: 'Sabah', greetingIcon: 'Sunny', decorationChar: '✦' },
    day: { label: 'Öğlen', greetingIcon: 'Brightness', decorationChar: '✦' },
    evening: { label: 'Akşam', greetingIcon: 'ClearNight', decorationChar: '✦' }
};

const buildTimeTheme = (bucket: TimeBucket): ITimeTheme => {
    const greeting = GREETING_BY_BUCKET[bucket];
    return {
        label: greeting.label,
        meshColors: CORPORATE_PALETTE.meshColors,
        text: CORPORATE_PALETTE.text,
        pageBackground: CORPORATE_PALETTE.pageBackground,
        header: {
            ...CORPORATE_PALETTE.header,
            greetingIcon: greeting.greetingIcon,
            decorationChar: greeting.decorationChar
        },
        cardAmbient: CORPORATE_PALETTE.cardAmbient
    };
};

/**
 * 08:00-12:00 Sabah · 12:01-16:00 Öğlen · 16:01-07:59 Akşam/Gece — sadece
 * selamlama metni/ikonu için, renk paletini ETKİLEMEZ (bkz. üstteki not).
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

export const getTimeTheme = (bucket: TimeBucket): ITimeTheme => buildTimeTheme(bucket);

// SharePoint sayfasındaki başka bir web part'ın kendi CSS değişkenleriyle
// çakışmasın diye tüm değişken adları bu önekle yazılıyor.
const CSS_VAR_PREFIX = '--yorpas-theme-';

/** :root üzerine (sabit) kurumsal tema değişkenlerini yazar. */
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
 * "ThemeManager" hook'u — bileşen ilk yüklendiğinde VE her dakika o anki
 * zaman dilimini (SADECE selamlama metni/ikonu için) kontrol edip :root CSS
 * değişkenlerini günceller. Renk değerleri sabit olduğu için bucket
 * değiştiğinde CSS değişkenlerinin GÖRSEL sonucu değişmez, ama
 * WelcomeHeader'ın "Günaydın/İyi günler/İyi akşamlar" metni bucket state'i
 * üzerinden güncellenmeye devam eder.
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
