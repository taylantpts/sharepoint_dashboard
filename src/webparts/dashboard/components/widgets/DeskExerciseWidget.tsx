import * as React from 'react';
import { useTheme, mergeStyleSets } from '@fluentui/react';
import WidgetCard from '../WidgetCard';

interface IStretchCategory {
    key: string;
    label: string;
    exerciseName: string;
    instruction: string;
    imageUrl: string;
}

/**
 * ÖNCEKİ İKİ SÜRÜM (wger.de'nin karışık-stil görselleri, sonra kendi
 * çizdiğimiz "manken" SVG'ler) kullanıcı tarafından sırasıyla "dandik" ve
 * "iğrenç" bulundu. Bu sürüm free-exercise-db (github.com/yuhonas/
 * free-exercise-db, Unlicense/kamu malı) veri setinden — TEK bir stüdyoda,
 * TEK bir modelle, aynı ışık/arka planla çekilmiş GERÇEK fotoğraflar
 * kullanıyor; bu yüzden dördü de birbiriyle tutarlı, profesyonel görünüyor.
 * jsDelivr CDN üzerinden <img> ile gösteriliyor (CORS/fetch riski yok, sabit
 * URL'ler) — API anahtarı gerekmiyor.
 */
const CATEGORIES: IStretchCategory[] = [
    {
        key: 'neck',
        label: 'Boyun',
        exerciseName: 'Boyun Yan Gerinmesi',
        instruction: 'Elinizle başınızı yavaşça omzunuza doğru çekin, iki tarafı da yapın.',
        imageUrl: 'https://cdn.jsdelivr.net/gh/yuhonas/free-exercise-db@main/exercises/Side_Neck_Stretch/1.jpg'
    },
    {
        key: 'shoulder',
        label: 'Omuz',
        exerciseName: 'Omuz Gerdirme',
        instruction: 'Bir kolunuzu göğsünüzün önünden karşı yöne doğru çekip diğer kolunuzla nazikçe destekleyin.',
        imageUrl: 'https://cdn.jsdelivr.net/gh/yuhonas/free-exercise-db@main/exercises/Shoulder_Stretch/1.jpg'
    },
    {
        key: 'back',
        label: 'Sırt',
        exerciseName: 'Ayakta Yan Gerinme',
        instruction: 'Ayakta bir kolunuzu başınızın üzerinden karşı yöne uzatıp gövdenizi o yöne doğru esnetin.',
        imageUrl: 'https://cdn.jsdelivr.net/gh/yuhonas/free-exercise-db@main/exercises/Standing_Lateral_Stretch/1.jpg'
    },
    {
        key: 'arm',
        label: 'Kol',
        exerciseName: 'Kol Arkası (Triceps) Gerinmesi',
        instruction: 'Bir kolunuzu başınızın üzerinden arkaya doğru katlayıp diğer elinizle dirseğinizi nazikçe destekleyin.',
        imageUrl: 'https://cdn.jsdelivr.net/gh/yuhonas/free-exercise-db@main/exercises/Triceps_Stretch/1.jpg'
    }
];

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
        // GERÇEK fotoğraf — oranı ne olursa olsun TAM KARE bir çerçevede
        // (aspectRatio 1/1) gösteriliyor, objectFit "cover" ile kırpılmadan/
        // gerilmeden kareyi dolduruyor; object-position üstte tutuyor ki
        // hareketin asıl göründüğü üst gövde/kollar kadraj dışı kalmasın.
        imageWrap: {
            width: '100%',
            maxWidth: 220,
            aspectRatio: '1 / 1',
            borderRadius: 16,
            overflow: 'hidden',
            marginBottom: 16,
            background: theme.palette.neutralLighterAlt,
            boxShadow: '0 4px 14px rgba(0,0,0,0.12)'
        },
        image: {
            width: '100%',
            height: '100%',
            objectFit: 'cover',
            objectPosition: 'center 15%'
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
        },
        credit: {
            fontSize: 10,
            color: theme.semanticColors.bodySubtext,
            textAlign: 'right',
            marginTop: 10
        },
        creditLink: {
            color: theme.semanticColors.bodySubtext
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
                        <img src={selected.imageUrl} alt={selected.exerciseName} className={styles.image} />
                    </div>
                    <div className={styles.textCol}>
                        <div className={styles.exerciseName}>{selected.exerciseName}</div>
                        <div className={styles.instruction}>{selected.instruction}</div>
                    </div>
                </div>
                <div className={styles.credit}>
                    Görsel:{' '}
                    <a href="https://yuhonas.github.io/free-exercise-db/" target="_blank" rel="noopener noreferrer" className={styles.creditLink}>
                        free-exercise-db
                    </a>
                </div>
            </div>
        </WidgetCard>
    );
};

export default DeskExerciseWidget;
