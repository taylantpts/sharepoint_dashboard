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
 * Görseller tek tek elle doğrulandı (wger.de'de her kayıt farklı kişi
 * tarafından yüklendiği için stil/doğruluk garantisi yok — bazı kayıtlarda
 * hareketle alakasız görseller bile var, ör. eski "Omuz Silkme" kaydı
 * aslında bir logo görseliydi). Şu an kullanılan 4 görsel elle kontrol
 * edildi: 3'ü aynı "düz gri/beyaz fonlu, kırmızı kas vurgulu 3D anatomik
 * çizim" ailesinden (1233, 1377), biri (1230) Healthwise tıbbi illüstrasyon
 * setinden — hepsi düz fonlu, sade illüstrasyon; karışık fotoğraf/çizim
 * karışımı yok.
 */
const CATEGORIES: IStretchCategory[] = [
    {
        key: 'neck',
        label: 'Boyun',
        exerciseName: 'Boyun Eğme',
        instruction: 'Kulağınızı yavaşça omzunuza yaklaştırın, iki tarafı da yapın.',
        imageUrl: 'https://wger.de/media/exercise-images/1018/5bbd3879-b6fc-4aaa-9e8e-33ae9a688112.png'
    },
    {
        key: 'shoulder',
        label: 'Omuz',
        exerciseName: 'Omuz ve Göğüs Gerdirme',
        instruction: 'Kolunuzu arkanızda bir duvara ya da kapı kenarına yaslayıp gövdenizi hafifçe öne çevirin.',
        imageUrl: 'https://wger.de/media/exercise-images/1233/d7d6f9e1-7834-4cca-bd3b-f9def33ff44d.png'
    },
    {
        key: 'back',
        label: 'Sırt',
        exerciseName: 'Gövde Çevirme',
        instruction: 'Ayakta kollarınızı yana açın, belinizden gövdenizi yavaşça sağa ve sola çevirin.',
        imageUrl: 'https://wger.de/media/exercise-images/1377/12e7a231-d36a-4992-bf57-ff7bfe0f3ae4.jpg'
    },
    {
        key: 'arm',
        label: 'Kol',
        exerciseName: 'Kol Arkası (Triceps) Gerinmesi',
        instruction: 'Bir kolunuzu başınızın üzerinden arkaya doğru katlayıp diğer elinizle dirseğinizi nazikçe destekleyin.',
        imageUrl: 'https://wger.de/media/exercise-images/1230/9fd1e2fd-f2c4-432d-b3ae-5e5f24085777.webp'
    }
];

const DeskExerciseWidget: React.FunctionComponent = () => {
    const theme = useTheme();
    const [selectedKey, setSelectedKey] = React.useState(CATEGORIES[0].key);
    const selected = CATEGORIES.filter((c) => c.key === selectedKey)[0];

    const styles = mergeStyleSets({
        // "Döviz/Coin" segmentli geçişiyle aynı görsel dil (bkz.
        // ExchangeRatesWidget.tsx modeToggleRow) — tek fark iki yerine dört
        // seçenek olması.
        tabRow: {
            display: 'flex',
            background: theme.palette.neutralLighterAlt,
            borderRadius: 10,
            padding: 3,
            marginBottom: 12
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
        // Görsel, ORANI NE OLURSA OLSUN her zaman TAM KARE bir çerçevede
        // gösteriliyor (aspectRatio: '1 / 1') — dikdörtgen bir kutuya sıkışıp
        // yamulan/kırpılan görsel sorunu böylece ortadan kalkıyor. objectFit
        // "contain" ile görsel bu karenin içine, kırpılmadan/gerilmeden sığdırılıyor.
        contentCol: {
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center'
        },
        imageWrap: {
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            width: '100%',
            maxWidth: 168,
            aspectRatio: '1 / 1',
            borderRadius: 12,
            background: theme.palette.white,
            border: `1px solid ${theme.palette.neutralLight}`,
            flexShrink: 0,
            overflow: 'hidden',
            marginBottom: 12
        },
        image: {
            maxWidth: '92%',
            maxHeight: '92%',
            objectFit: 'contain'
        },
        textCol: {
            minWidth: 0,
            textAlign: 'center'
        },
        exerciseName: {
            fontSize: 13,
            fontWeight: 700,
            color: theme.semanticColors.bodyText,
            marginBottom: 3
        },
        instruction: {
            fontSize: 11,
            lineHeight: '15px',
            color: theme.semanticColors.bodySubtext,
            maxWidth: 260,
            margin: '0 auto'
        },
        credit: {
            fontSize: 10,
            color: theme.semanticColors.bodySubtext,
            textAlign: 'right',
            marginTop: 8
        },
        creditLink: {
            color: theme.semanticColors.bodySubtext
        }
    });

    return (
        <WidgetCard title="Esneme Molası" subtitle="Kısa bir rahatlatma hareketi seçin" iconName="Health">
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
                <a href="https://wger.de" target="_blank" rel="noopener noreferrer" className={styles.creditLink}>
                    wger.de
                </a>
            </div>
        </WidgetCard>
    );
};

export default DeskExerciseWidget;
