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
 * ÖNCEKİ SÜRÜMLER wger.de'nin (açık kaynaklı, üçüncü kişilerce yüklenen)
 * egzersiz görsellerini kullanıyordu — kullanıcı bunları "dandik" ve
 * birbirinden farklı temada buldu (bazıları gerçek fotoğraf, bazıları 3D
 * render, biri hatta hareketle alakasız bir logo çıktı). Dışarıdan hiçbir
 * görsele bağımlı olmamak için 4 hareket artık kendi SVG çizimimizle
 * (StretchIllustration) çiziliyor — tek elden, tek çizgi stiliyle, her
 * zaman aynı temada ve her zaman erişilebilir (ağ isteği/CORS/bozuk link
 * riski yok).
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

// Dört hareketin de gövde/kol/bacak "iskeleti" ortak — sadece kol pozu ve
// varsa hareket ipucu (ok/duvar çizgisi) değişiyor. Böylece dört çizim de
// aynı orantı ve çizgi kalınlığıyla, tek bir elden çıkmış gibi görünüyor.
const StretchIllustration: React.FunctionComponent<{ pose: string; color: string }> = ({ pose, color }) => {
    const common = (
        <>
            <line x1="38" y1="72" x2="62" y2="72" />
            <line x1="38" y1="72" x2="33" y2="98" />
            <line x1="62" y1="72" x2="67" y2="98" />
        </>
    );

    return (
        <svg viewBox="0 0 100 100" fill="none" stroke={color} strokeWidth={4.5} strokeLinecap="round" strokeLinejoin="round">
            {pose === 'neck' && (
                <>
                    <circle cx="40" cy="20" r="9" fill={color} stroke="none" />
                    <path d="M46 27 L50 33" />
                    <line x1="35" y1="33" x2="65" y2="33" />
                    <line x1="50" y1="33" x2="50" y2="72" />
                    <line x1="35" y1="33" x2="28" y2="58" />
                    <line x1="65" y1="33" x2="72" y2="58" />
                    <path d="M22 16 Q17 20 22 26" strokeWidth={3.5} />
                    <path d="M58 10 Q63 14 58 20" strokeWidth={3.5} />
                    {common}
                </>
            )}
            {pose === 'shoulder' && (
                <>
                    <circle cx="50" cy="18" r="9" fill={color} stroke="none" />
                    <line x1="50" y1="27" x2="50" y2="33" />
                    <line x1="35" y1="33" x2="65" y2="33" />
                    <line x1="50" y1="33" x2="50" y2="72" />
                    <line x1="35" y1="33" x2="28" y2="58" />
                    <path d="M65 33 L82 30 L88 40" />
                    <line x1="90" y1="14" x2="90" y2="50" strokeWidth={3} />
                    {common}
                </>
            )}
            {pose === 'back' && (
                <>
                    <circle cx="50" cy="18" r="9" fill={color} stroke="none" />
                    <line x1="50" y1="27" x2="50" y2="33" />
                    <line x1="20" y1="38" x2="80" y2="30" />
                    <line x1="50" y1="33" x2="50" y2="72" />
                    <path d="M28 50 A26 26 0 1 1 34 68" strokeWidth={3.5} />
                    <path d="M31 62 L34 68 L40 65" strokeWidth={3.5} />
                    {common}
                </>
            )}
            {pose === 'arm' && (
                <>
                    <circle cx="50" cy="18" r="9" fill={color} stroke="none" />
                    <line x1="50" y1="27" x2="50" y2="33" />
                    <line x1="35" y1="33" x2="65" y2="33" />
                    <line x1="50" y1="33" x2="50" y2="72" />
                    <line x1="35" y1="33" x2="30" y2="52" />
                    <line x1="30" y1="52" x2="42" y2="58" />
                    <path d="M65 33 Q78 40 68 14 Q60 4 55 15" />
                    {common}
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
        // Widget artık İSG Takvimi ile aynı en/boy oranına yakın (bkz.
        // Dashboard.module.scss .areaExercise — yarım genişlik) — çizim de
        // bu kareye uyum sağlasın diye flex-grow ile kalan tüm dikey alanı
        // dolduruyor, taban çizgisi hep ortada kalıyor.
        contentCol: {
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            justifyContent: 'center',
            flexGrow: 1,
            minHeight: 0
        },
        imageWrap: {
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            width: '100%',
            maxWidth: 220,
            aspectRatio: '1 / 1',
            borderRadius: 16,
            background: theme.palette.neutralLighterAlt,
            flexShrink: 1,
            minHeight: 0,
            overflow: 'hidden',
            marginBottom: 16
        },
        image: {
            width: '78%',
            height: '78%'
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
                    <div className={styles.imageWrap}>
                        <div className={styles.image}>
                            <StretchIllustration pose={selected.key} color={selected.color} />
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
