import * as React from 'react';
import { Icon, IconButton, useTheme, mergeStyleSets } from '@fluentui/react';
import WidgetCard from '../WidgetCard';

interface IExercise {
    name: string;
    instruction: string;
    imageUrl: string;
}

/**
 * Sabit, elle seçilmiş bir masa başı egzersizi listesi — GÖRSELLER
 * wger.de'nin (açık kaynaklı, anahtarsız/ücretsiz) egzersiz veritabanından
 * geliyor (CC lisanslı, bkz. altındaki atıf satırı). wger'in 834 kayıtlık
 * tüm veritabanı gym/ekipman ağırlıklı olduğu için OLDUĞU GİBİ çekilmiyor —
 * "equipment=none" + boyun/omuz/kol/bacak kategorilerinde GERÇEKTEN
 * görseli olan ~50 kayıt arasından, masa başında oturarak/ayakta kolayca
 * yapılabilecek 11 tanesi elle seçildi; ad/açıklama de buraya Türkçe
 * olarak elle yazıldı (API'de her kaydın Türkçe çevirisi yok).
 *
 * Görseller <img> ile gösterildiği için CORS/fetch riski YOK — sadece bir
 * resim URL'i, tıpkı harici bir web sitesindeki resmi göstermek gibi.
 */
const EXERCISES: IExercise[] = [
    {
        name: 'Boyun Çevirme',
        instruction: 'Başınızı yavaşça sağa çevirin, 5 saniye bekleyin, sonra sola çevirin.',
        imageUrl: 'https://wger.de/media/exercise-images/1007/757846d3-78e4-4068-bbca-62e567372c94.png'
    },
    {
        name: 'Boyun Eğme',
        instruction: 'Kulağınızı yavaşça omzunuza yaklaştırın, iki tarafı da yapın.',
        imageUrl: 'https://wger.de/media/exercise-images/1018/5bbd3879-b6fc-4aaa-9e8e-33ae9a688112.png'
    },
    {
        name: 'Omuz Silkme',
        instruction: 'Omuzlarınızı kulaklarınıza doğru kaldırın, 3 saniye tutup gevşetin.',
        imageUrl: 'https://wger.de/media/exercise-images/570/68b4a33f-40f1-4dda-b56c-a2e20ed13903.jpg'
    },
    {
        name: 'Yan Gerinme',
        instruction: 'Bir kolunuzu yukarı kaldırıp gövdenizi yavaşça yana doğru esnetin.',
        imageUrl: 'https://wger.de/media/exercise-images/1861/0ffe4e99-71ad-47fb-b98c-1f243faa0499.png'
    },
    {
        name: 'Gövde Çevirme',
        instruction: 'Otururken üst gövdenizi yavaşça sağa, sonra sola çevirin.',
        imageUrl: 'https://wger.de/media/exercise-images/1377/12e7a231-d36a-4992-bf57-ff7bfe0f3ae4.jpg'
    },
    {
        name: 'Sol Pazı Gerinmesi',
        instruction: 'Sol kolunuzu arkaya uzatıp avucunuzu içe çevirerek gerinin.',
        imageUrl: 'https://wger.de/media/exercise-images/1232/2b6de046-5806-49e3-bf36-b6fae16af021.png'
    },
    {
        name: 'Sağ Pazı Gerinmesi',
        instruction: 'Sağ kolunuzu arkaya uzatıp avucunuzu içe çevirerek gerinin.',
        imageUrl: 'https://wger.de/media/exercise-images/1233/d7d6f9e1-7834-4cca-bd3b-f9def33ff44d.png'
    },
    {
        name: 'Sol Kol Arkası (Triceps) Gerinmesi',
        instruction: 'Sol kolunuzu başınızın üzerinden arkaya doğru katlayıp diğer elle destekleyin.',
        imageUrl: 'https://wger.de/media/exercise-images/1230/9fd1e2fd-f2c4-432d-b3ae-5e5f24085777.webp'
    },
    {
        name: 'Sağ Kol Arkası (Triceps) Gerinmesi',
        instruction: 'Sağ kolunuzu başınızın üzerinden arkaya doğru katlayıp diğer elle destekleyin.',
        imageUrl: 'https://wger.de/media/exercise-images/1231/b10457ce-5fa5-4d20-a32f-3c7100c6a9d9.webp'
    },
    {
        name: 'Kalça Ön Gerinmesi',
        instruction: 'Ayağa kalkıp bir adım öne atarak kalça ön kaslarınızı yavaşça gerin.',
        imageUrl: 'https://wger.de/media/exercise-images/1867/767631e5-10d2-46b8-b03f-cc298f96963b.png'
    },
    {
        name: 'Uyluk (Quadriceps) Gerinmesi',
        instruction: 'Ayağa kalkıp bir ayağınızı arkaya doğru katlayarak elinizle tutun.',
        imageUrl: 'https://wger.de/media/exercise-images/1873/c0ed299b-6d87-4d90-885d-bb3b5d85f1eb.png'
    }
];

/** Sayfa her açıldığında AYNI ilk egzersizle karşılaşılmasın diye başlangıç rastgele seçiliyor. */
const getRandomStartIndex = (): number => Math.floor(Math.random() * EXERCISES.length);

const DeskExerciseWidget: React.FunctionComponent = () => {
    const theme = useTheme();
    const [index, setIndex] = React.useState(getRandomStartIndex);

    const exercise = EXERCISES[index];
    const goPrev = (): void => setIndex((i) => (i === 0 ? EXERCISES.length - 1 : i - 1));
    const goNext = (): void => setIndex((i) => (i === EXERCISES.length - 1 ? 0 : i + 1));

    const styles = mergeStyleSets({
        imageWrap: {
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            height: 180,
            borderRadius: 12,
            background: theme.palette.neutralLighterAlt,
            marginBottom: 14,
            overflow: 'hidden'
        },
        image: {
            maxWidth: '100%',
            maxHeight: '100%',
            objectFit: 'contain'
        },
        name: {
            fontSize: 15,
            fontWeight: 700,
            color: theme.semanticColors.bodyText,
            marginBottom: 6
        },
        instruction: {
            fontSize: 12,
            lineHeight: '18px',
            color: theme.semanticColors.bodySubtext,
            marginBottom: 14
        },
        // NOT: "gap" flex özelliği burada BİLİNÇLİ OLARAK KULLANILMIYOR — bu
        // sayfanın render edildiği (kurumsal/eski) tarayıcı ortamında flex "gap"
        // desteklenmiyor. Diğer widget'lardaki (Personel İlanları slider'ı,
        // Resmi Tatiller sayfalaması) aynı düzen deseni.
        navRow: {
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center'
        },
        navPrevButton: {
            marginRight: 8
        },
        counter: {
            fontSize: 11,
            color: theme.semanticColors.bodySubtext,
            minWidth: 46,
            textAlign: 'center',
            marginRight: 8
        },
        credit: {
            fontSize: 10,
            color: theme.semanticColors.bodySubtext,
            textAlign: 'center',
            marginTop: 10
        },
        creditLink: {
            color: theme.semanticColors.bodySubtext
        }
    });

    return (
        <WidgetCard title="Masa Başı Molası" subtitle="Kısa bir gerinme molası verin" iconName="Health">
            <div className={styles.imageWrap}>
                <img src={exercise.imageUrl} alt={exercise.name} className={styles.image} />
            </div>
            <div className={styles.name}>{exercise.name}</div>
            <div className={styles.instruction}>{exercise.instruction}</div>
            <div className={styles.navRow}>
                <IconButton
                    className={styles.navPrevButton}
                    iconProps={{ iconName: 'ChevronLeft' }}
                    ariaLabel="Önceki hareket"
                    onClick={goPrev}
                />
                <span className={styles.counter}>{index + 1} / {EXERCISES.length}</span>
                <IconButton
                    iconProps={{ iconName: 'ChevronRight' }}
                    ariaLabel="Sonraki hareket"
                    onClick={goNext}
                />
            </div>
            <div className={styles.credit}>
                <Icon iconName="Info" style={{ marginRight: 4, fontSize: 10 }} />
                Görseller{' '}
                <a href="https://wger.de" target="_blank" rel="noopener noreferrer" className={styles.creditLink}>
                    wger.de
                </a>{' '}
                açık kaynaklı egzersiz veritabanından alınmıştır.
            </div>
        </WidgetCard>
    );
};

export default DeskExerciseWidget;
