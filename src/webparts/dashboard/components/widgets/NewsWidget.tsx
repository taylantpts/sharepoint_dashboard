import * as React from 'react';
import { Icon, Spinner, SpinnerSize, MessageBar, MessageBarType, useTheme, mergeStyleSets } from '@fluentui/react';
import WidgetCard from '../WidgetCard';
import { DATA_UNAVAILABLE_MESSAGE } from '../../constants';

// Kaynak RSS akışları — kod değiştirmeden başka/ek bir haber kaynağına geçmek
// için sadece bu listeyi güncellemek yeterli. rss2json.com anahtarsız/ücretsiz
// katmanı (günlük istek sınırı var — burada widget başına N istek yapıldığı
// için limit tek akışa göre N kat daha hızlı dolar) RSS'i tarayıcıdan CORS ile
// okunabilir JSON'a çeviriyor; trafik arttığında rss2json'da ücretli bir
// plana/kendi anahtarınıza geçmeniz gerekebilir.
// ÖNCEKİ HATA (v1): "gundem" (genel gündem) akışı zaman zaman siyasi/parti
// içerikli başlıklar da içeriyordu.
// ÖNCEKİ HATA (v2): tek başına "Yaşam" kategorisine geçildi — siyasetten
// bağımsızdı ama tamamen magazin/insan hikâyesi ağırlıklıydı, "kaliteli/genel
// haber" beklentisini karşılamıyordu.
// v3: Aynı güvenilir kaynağın (NTV) SİYASİ RİSKİ DÜŞÜK, ciddi/genel birden
// fazla kategorisi birleştiriliyor (Ekonomi, Teknoloji, Sağlık, Eğitim,
// Yaşam). "Gündem", "Türkiye" ve "Dünya" kategorileri BİLİNÇLİ OLARAK dışarıda
// bırakıldı — bunlar hükümet/parti/seçim/dış politika haberlerinin en yoğun
// olduğu kategoriler. Ek bir güvenlik katmanı olarak POLITICAL_KEYWORDS ile
// başlık bazlı bir filtre de uygulanıyor (ör. ekonomi haberinde bir bakanın
// adının geçmesi gibi sızıntıları yakalamak için).
const NEWS_CATEGORIES: ReadonlyArray<{ url: string; label: string }> = [
    { url: 'https://www.ntv.com.tr/ekonomi.rss', label: 'Ekonomi' },
    { url: 'https://www.ntv.com.tr/teknoloji.rss', label: 'Teknoloji' },
    { url: 'https://www.ntv.com.tr/saglik.rss', label: 'Sağlık' },
    { url: 'https://www.ntv.com.tr/egitim.rss', label: 'Eğitim' },
    { url: 'https://www.ntv.com.tr/yasam.rss', label: 'Yaşam' }
];

const rss2JsonEndpoint = (rssUrl: string): string =>
    `https://api.rss2json.com/v1/api.json?rss_url=${encodeURIComponent(rssUrl)}`;

// Kategori seçimi siyasi içeriği büyük ölçüde eliyor, ama bu son bir güvenlik
// ağı: "ciddi" kategorilerde (ör. Ekonomi) bile ara sıra bir bakan/parti
// açıklamasına atıf geçebiliyor — böyle başlıklar listeye hiç girmesin diye
// başlık metninde bu kelimelerden biri geçen haberler baştan eleniyor.
const POLITICAL_KEYWORDS: ReadonlyArray<string> = [
    'cumhurbaşkan', 'başbakan', 'bakan', 'meclis', 'parti', 'seçim',
    'milletvekili', 'vekil', 'ittifak', 'muhalefet', 'iktidar', 'siyaset',
    'siyasi', 'chp', 'akp', 'mhp', 'iyi parti', 'dem parti'
];

const containsPoliticalKeyword = (title: string): boolean => {
    const normalized = title.toLocaleLowerCase('tr-TR');
    return POLITICAL_KEYWORDS.some((keyword) => normalized.indexOf(keyword) !== -1);
};

const MAX_ITEMS = 6;

interface IRss2JsonItem {
    title: string;
    link: string;
    pubDate: string;
    thumbnail?: string;
    description?: string;
}

interface IRss2JsonResponse {
    status: string;
    items: IRss2JsonItem[];
}

interface INewsItem {
    title: string;
    link: string;
    dateLabel: string;
    thumbnail?: string;
    category: string;
    pubDate: string;
}

type LoadState = 'loading' | 'loaded' | 'error';

const stripHtml = (html: string): string => html.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim();

// RSS başlıkları HTML-kodlanmış karakterler taşıyabilir (ör. tırnak için
// "&quot;", kesme işareti için "&#39;") — bunlar sade metin olarak render
// edildiği için (dangerouslySetInnerHTML KULLANILMIYOR) React'in kendisi
// bunları asla çözmez, ekranda literal "&quot;" yazar. <textarea>.innerHTML
// bu kodlamayı ÇÖZER ama İÇİNE KONAN İÇERİĞİ ASLA HTML OLARAK ÇALIŞTIRMAZ
// (textarea'nın içeriği her zaman düz metin muamelesi görür) — bu yüzden
// DOMParser/innerHTML'in aksine XSS riski taşımadan güvenle kullanılabilir.
const decodeHtmlEntities = (text: string): string => {
    const el = document.createElement('textarea');
    el.innerHTML = text;
    return el.value;
};

// rss2json her haberde "thumbnail" alanını doldurmuyor (RSS'te <enclosure>/
// <media:thumbnail> yoksa boş geliyor) — bu durumda haberin HTML açıklaması
// (description) içine gömülü ilk <img>'in src'si son çare olarak kullanılıyor.
const extractFirstImageUrl = (html?: string): string | undefined => {
    if (!html) {
        return undefined;
    }
    const match = /<img[^>]+src=["']([^"']+)["']/i.exec(html);
    return match ? match[1] : undefined;
};

const NewsWidget: React.FunctionComponent = () => {
    const theme = useTheme();
    const [state, setState] = React.useState<LoadState>('loading');
    const [news, setNews] = React.useState<INewsItem[]>([]);
    // Bazı thumbnail URL'leri (kırık kaynak/404) tarayıcıda hiç yüklenmeyebilir
    // — bu durumda kırık resim ikonu yerine yer tutucu ikona düşülür.
    const [brokenThumbs, setBrokenThumbs] = React.useState<Set<string>>(new Set());

    React.useEffect(() => {
        let isMounted = true;

        // Promise.allSettled bu projenin TS lib hedefinde yok — aynı "bir
        // kategori başarısız olsa da diğerleri listeyi doldursun" davranışını
        // her fetch'i kendi içinde yakalayıp hatada boş dizi döndürerek elde
        // ediyoruz; Promise.all bu durumda asla reject olmaz.
        const fetchCategory = async (category: { url: string; label: string }): Promise<INewsItem[]> => {
            try {
                const response = await fetch(rss2JsonEndpoint(category.url));
                if (!response.ok) {
                    throw new Error(`rss2json isteği başarısız (${category.label}, HTTP ${response.status})`);
                }
                const data: IRss2JsonResponse = await response.json();
                if (data.status !== 'ok') {
                    throw new Error(`rss2json beklenmeyen durum döndürdü (${category.label}): ${data.status}`);
                }

                return data.items
                    .filter((item) => !containsPoliticalKeyword(item.title))
                    .map((item) => ({
                        title: decodeHtmlEntities(stripHtml(item.title)),
                        link: item.link,
                        dateLabel: new Date(item.pubDate).toLocaleDateString('tr-TR', { day: 'numeric', month: 'short' }),
                        thumbnail: item.thumbnail || extractFirstImageUrl(item.description) || undefined,
                        category: category.label,
                        pubDate: item.pubDate
                    }));
            } catch (error) {
                console.error('[NewsWidget] Kategori alınamadı:', category.label, error);
                return [];
            }
        };

        const load = async (): Promise<void> => {
            const perCategory = await Promise.all(NEWS_CATEGORIES.map(fetchCategory));
            const merged = perCategory.reduce<INewsItem[]>((all, items) => all.concat(items), []);

            merged.sort((a, b) => new Date(b.pubDate).getTime() - new Date(a.pubDate).getTime());

            if (isMounted) {
                if (merged.length === 0) {
                    setState('error');
                } else {
                    setNews(merged.slice(0, MAX_ITEMS));
                    setState('loaded');
                }
            }
        };

        load().catch(() => { /* load kendi içinde hatayı yönetir */ });

        return () => {
            isMounted = false;
        };
    }, []);

    const styles = mergeStyleSets({
        list: {
            display: 'flex',
            flexDirection: 'column'
        },
        row: {
            display: 'flex',
            alignItems: 'center',
            padding: '8px',
            marginBottom: 8,
            borderRadius: 10,
            textDecoration: 'none',
            transition: 'background 0.15s ease',
            selectors: {
                ':hover': { background: theme.palette.neutralLighterAlt }
            }
        },
        thumbWrap: {
            width: 56,
            height: 56,
            borderRadius: 8,
            overflow: 'hidden',
            flexShrink: 0,
            marginRight: 12,
            background: theme.palette.neutralLighterAlt,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center'
        },
        thumbImg: {
            width: '100%',
            height: '100%',
            objectFit: 'cover'
        },
        placeholderIcon: {
            fontSize: 20,
            color: theme.palette.neutralTertiary
        },
        detailGroup: {
            display: 'flex',
            flexDirection: 'column',
            minWidth: 0
        },
        title: {
            fontSize: 13,
            fontWeight: 600,
            color: theme.semanticColors.bodyText,
            lineHeight: '18px',
            display: '-webkit-box',
            WebkitLineClamp: 2,
            WebkitBoxOrient: 'vertical',
            overflow: 'hidden',
            marginBottom: 4
        },
        date: {
            fontSize: 11,
            color: theme.semanticColors.bodySubtext
        }
    });

    return (
        <WidgetCard title="Haberler" subtitle="Ekonomi, teknoloji, sağlık ve daha fazlası" iconName="News">
            {state === 'loading' && <Spinner size={SpinnerSize.medium} label="Haberler yükleniyor..." />}
            {state === 'error' && <MessageBar messageBarType={MessageBarType.error}>{DATA_UNAVAILABLE_MESSAGE}</MessageBar>}
            {state === 'loaded' && (
                <div className={styles.list}>
                    {news.map((item) => (
                        <a key={item.link} href={item.link} target="_blank" rel="noopener noreferrer" className={styles.row}>
                            <div className={styles.thumbWrap}>
                                {item.thumbnail && !brokenThumbs.has(item.thumbnail) ? (
                                    <img
                                        src={item.thumbnail}
                                        alt=""
                                        className={styles.thumbImg}
                                        onError={() => {
                                            const url = item.thumbnail as string;
                                            setBrokenThumbs((prev) => (prev.has(url) ? prev : new Set(prev).add(url)));
                                        }}
                                    />
                                ) : (
                                    <Icon iconName="News" className={styles.placeholderIcon} />
                                )}
                            </div>
                            <div className={styles.detailGroup}>
                                <div className={styles.title}>{item.title}</div>
                                <div className={styles.date}>{item.category} · {item.dateLabel}</div>
                            </div>
                        </a>
                    ))}
                </div>
            )}
        </WidgetCard>
    );
};

export default NewsWidget;
