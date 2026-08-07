import * as React from 'react';
import { useTheme, mergeStyleSets } from '@fluentui/react';
import WidgetCard from '../WidgetCard';

interface IStretchCategory {
    key: string;
    label: string;
    exerciseName: string;
    instruction: string;
    color: string;
}

/**
 * ÖNCEKİ SÜRÜM (ince çizgili "çöp adam" SVG'ler) kullanıcı tarafından
 * "dandik" bulundu — ince stroke'lu iskelet çizimler kurumsal bir panoda
 * amatör/taslak gibi duruyordu. Bunun yerine QuickLinksPanel/WidgetCard'ta
 * zaten kullanılan "renkli degrade daire + beyaz ikon" diliyle aynı
 * mantıkta, KALIN yuvarlak uçlu (round line cap) beyaz figür + doygun
 * degrade daire arka plan kullanılıyor — "ahşap eklemli manken" tarzı sade
 * ama dolgun bir siluet, ince çöp-adam çizgilerinden çok daha "tasarlanmış"
 * görünüyor. Yine tamamen kendi SVG çizimimiz — dış görsele bağımlılık yok.
 */
const CATEGORIES: IStretchCategory[] = [
    {
        key: 'neck',
        label: 'Boyun',
        exerciseName: 'Boyun Eğme',
        instruction: 'Kulağınızı yavaşça omzunuza yaklaştırın, iki tarafı da yapın.',
        color: '#0078d4'
    },
    {
        key: 'shoulder',
        label: 'Omuz',
        exerciseName: 'Omuz ve Göğüs Gerdirme',
        instruction: 'Kolunuzu arkanızda bir duvara ya da kapı kenarına yaslayıp gövdenizi hafifçe öne çevirin.',
        color: '#107c10'
    },
    {
        key: 'back',
        label: 'Sırt',
        exerciseName: 'Gövde Çevirme',
        instruction: 'Ayakta kollarınızı yana açın, belinizden gövdenizi yavaşça sağa ve sola çevirin.',
        color: '#d83b01'
    },
    {
        key: 'arm',
        label: 'Kol',
        exerciseName: 'Kol Arkası (Triceps) Gerinmesi',
        instruction: 'Bir kolunuzu başınızın üzerinden arkaya doğru katlayıp diğer elinizle dirseğinizi nazikçe destekleyin.',
        color: '#5c2d91'
    }
];

// Dört hareketin de gövde/bacak "iskeleti" ortak (aynı omuz/kalça/bacak
// koordinatları) — sadece kol pozu ve baş konumu değişiyor. Kalın (stroke
// 10), yuvarlak uçlu tek-parça çizgiler eklem yerlerinde doğal olarak
// birleşip dolgun bir siluet oluşturuyor; ayrıca ince/donuk bir hareket
// yayı (dashed, opacity 0.55) — eski kalın/keskin ok ikonları kaldırıldı.
const StretchIllustration: React.FunctionComponent<{ pose: string }> = ({ pose }) => {
    const legsAndHips = (
        <>
            <line x1="32" y1="36" x2="68" y2="36" />
            <line x1="50" y1="34" x2="50" y2="68" />
            <line x1="40" y1="68" x2="60" y2="68" />
            <line x1="40" y1="68" x2="34" y2="95" />
            <line x1="60" y1="68" x2="66" y2="95" />
        </>
    );

    return (
        <svg viewBox="0 0 100 100" fill="none" stroke="#ffffff" strokeWidth={10} strokeLinecap="round" strokeLinejoin="round">
            {pose === 'neck' && (
                <>
                    <path d="M50 34 Q42 26 37 22" strokeWidth={10} />
                    <circle cx="35" cy="18" r="11" fill="#ffffff" stroke="none" />
                    <line x1="32" y1="36" x2="24" y2="58" />
                    <line x1="68" y1="36" x2="76" y2="58" />
                    <path d="M20 6 A22 22 0 0 1 50 6" strokeWidth={3} strokeDasharray="1 7" opacity={0.55} />
                    {legsAndHips}
                </>
            )}
            {pose === 'shoulder' && (
                <>
                    <circle cx="50" cy="20" r="11" fill="#ffffff" stroke="none" />
                    <line x1="32" y1="36" x2="24" y2="58" />
                    <path d="M68 36 L84 40 L88 52" />
                    <line x1="90" y1="18" x2="90" y2="58" strokeWidth={3} opacity={0.55} />
                    {legsAndHips}
                </>
            )}
            {pose === 'back' && (
                <>
                    <circle cx="50" cy="20" r="11" fill="#ffffff" stroke="none" />
                    <line x1="18" y1="42" x2="82" y2="30" />
                    <path d="M24 52 A28 28 0 1 0 30 72" strokeWidth={3} strokeDasharray="1 7" opacity={0.55} />
                    {legsAndHips}
                </>
            )}
            {pose === 'arm' && (
                <>
                    <circle cx="50" cy="20" r="11" fill="#ffffff" stroke="none" />
                    <line x1="32" y1="36" x2="26" y2="54" />
                    <line x1="26" y1="54" x2="40" y2="60" />
                    <path d="M68 36 Q82 44 70 12 Q62 0 55 14" />
                    {legsAndHips}
                </>
            )}
        </svg>
    );
};

const DeskExerciseWidget: React.FunctionComponent = () => {
    const theme = useTheme();
    const [selectedKey, setSelectedKey] = React.useState(CATEGORIES[0].key);
    const selected = CATEGORIES.filter((c) => c.key === selectedKey)[0];

    const styles = mergeStyleSets({
        root: {
            display: 'flex',
            flexDirection: 'column',
            height: '100%'
        },
        // "Döviz/Coin" segmentli geçişiyle aynı görsel dil (bkz.
        // ExchangeRatesWidget.tsx modeToggleRow) — tek fark iki yerine dört
        // seçenek olması.
        tabRow: {
            display: 'flex',
            background: theme.palette.neutralLighterAlt,
            borderRadius: 10,
            padding: 3,
            marginBottom: 16
        },
        tabButton: {
            flexGrow: 1,
            border: 'none',
            background: 'transparent',
            borderRadius: 8,
            padding: '6px 0',
            fontSize: 12,
            fontWeight: 600,
            color: theme.semanticColors.bodySubtext,
            cursor: 'pointer',
            transition: 'background 0.15s ease, color 0.15s ease'
        },
        tabButtonActive: {
            background: theme.palette.white,
            color: theme.palette.themePrimary,
            boxShadow: '0 1px 4px rgba(0,0,0,0.12)'
        },
        contentCol: {
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            justifyContent: 'center',
            flexGrow: 1,
            minHeight: 0
        },
        // Renkli degrade daire arka plan — QuickLinksPanel.tsx'teki iconWrap
        // ile AYNI mantık (linear-gradient 135deg + renkli box-shadow),
        // sadece burada tüm görsel alanı kaplayacak kadar büyük bir daire.
        imageWrap: {
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            width: '100%',
            maxWidth: 200,
            aspectRatio: '1 / 1',
            borderRadius: '50%',
            flexShrink: 1,
            minHeight: 0,
            overflow: 'hidden',
            marginBottom: 16,
            transition: 'background 0.2s ease, box-shadow 0.2s ease'
        },
        image: {
            width: '62%',
            height: '62%'
        },
        textCol: {
            minWidth: 0,
            textAlign: 'center'
        },
        exerciseName: {
            fontSize: 14,
            fontWeight: 700,
            color: theme.semanticColors.bodyText,
            marginBottom: 4
        },
        instruction: {
            fontSize: 12,
            lineHeight: '17px',
            color: theme.semanticColors.bodySubtext,
            maxWidth: 280,
            margin: '0 auto'
        }
    });

    return (
        <WidgetCard title="Esneme Molası" subtitle="Kısa bir rahatlatma hareketi seçin" iconName="Health">
            <div className={styles.root}>
                <div className={styles.tabRow}>
                    {CATEGORIES.map((cat) => (
                        <button
                            key={cat.key}
                            type="button"
                            className={`${styles.tabButton} ${cat.key === selectedKey ? styles.tabButtonActive : ''}`}
                            onClick={() => setSelectedKey(cat.key)}
                        >
                            {cat.label}
                        </button>
                    ))}
                </div>
                <div className={styles.contentCol}>
                    <div
                        className={styles.imageWrap}
                        style={{
                            background: `linear-gradient(135deg, ${selected.color} 0%, ${selected.color}cc 100%)`,
                            boxShadow: `0 8px 20px ${selected.color}55`
                        }}
                    >
                        <div className={styles.image}>
                            <StretchIllustration pose={selected.key} />
                        </div>
                    </div>
                    <div className={styles.textCol}>
                        <div className={styles.exerciseName}>{selected.exerciseName}</div>
                        <div className={styles.instruction}>{selected.instruction}</div>
                    </div>
                </div>
            </div>
        </WidgetCard>
    );
};

export default DeskExerciseWidget;
