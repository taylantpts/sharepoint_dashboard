import * as React from 'react';
import { Icon, Spinner, SpinnerSize, MessageBar, MessageBarType, useTheme, mergeStyleSets } from '@fluentui/react';
import WidgetCard from '../WidgetCard';
import { DATA_UNAVAILABLE_MESSAGE } from '../../constants';

// Kaynak RSS akışı — kod değiştirmeden başka bir haber kaynağına geçmek için
// sadece bu satırı güncellemek yeterli. rss2json.com anahtarsız/ücretsiz
// katmanı (günlük istek sınırı var) RSS'i tarayıcıdan CORS ile okunabilir
// JSON'a çeviriyor; trafik arttığında rss2json'da ücretli bir plana/kendi
// anahtarınıza geçmeniz gerekebilir.
// ÖNCEKİ HATA: "gundem" (genel gündem) akışı zaman zaman siyasi/parti içerikli
// başlıklar da içeriyordu — kurumsal bir intranet için istenmeyen bir durum.
// Aynı güvenilir kaynağın (NTV) SADECE "Yaşam" kategorisine geçildi: bu akış
// tamamen gündelik/magazin/insan hikâyesi ağırlıklı, siyasetten bağımsız.
const NEWS_RSS_FEED_URL = 'https://www.ntv.com.tr/yasam.rss';
const RSS2JSON_ENDPOINT = `https://api.rss2json.com/v1/api.json?rss_url=${encodeURIComponent(NEWS_RSS_FEED_URL)}`;

const MAX_ITEMS = 5;

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
}

type LoadState = 'loading' | 'loaded' | 'error';

const stripHtml = (html: string): string => html.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim();

const NewsWidget: React.FunctionComponent = () => {
    const theme = useTheme();
    const [state, setState] = React.useState<LoadState>('loading');
    const [news, setNews] = React.useState<INewsItem[]>([]);

    React.useEffect(() => {
        let isMounted = true;

        const load = async (): Promise<void> => {
            try {
                const response = await fetch(RSS2JSON_ENDPOINT);
                if (!response.ok) {
                    throw new Error(`rss2json isteği başarısız (HTTP ${response.status})`);
                }
                const data: IRss2JsonResponse = await response.json();
                if (data.status !== 'ok') {
                    throw new Error(`rss2json beklenmeyen durum döndürdü: ${data.status}`);
                }

                const items: INewsItem[] = data.items.slice(0, MAX_ITEMS).map((item) => ({
                    title: stripHtml(item.title),
                    link: item.link,
                    dateLabel: new Date(item.pubDate).toLocaleDateString('tr-TR', { day: 'numeric', month: 'short' }),
                    thumbnail: item.thumbnail || undefined
                }));

                if (isMounted) {
                    setNews(items);
                    setState('loaded');
                }
            } catch (error) {
                console.error('[NewsWidget] Haberler alınamadı:', error);
                if (isMounted) {
                    setState('error');
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
        <WidgetCard title="Haberler" subtitle="Gündemden önemli başlıklar" iconName="News" accentColor="#3f6fb0">
            {state === 'loading' && <Spinner size={SpinnerSize.medium} label="Haberler yükleniyor..." />}
            {state === 'error' && <MessageBar messageBarType={MessageBarType.error}>{DATA_UNAVAILABLE_MESSAGE}</MessageBar>}
            {state === 'loaded' && (
                <div className={styles.list}>
                    {news.map((item) => (
                        <a key={item.link} href={item.link} target="_blank" rel="noopener noreferrer" className={styles.row}>
                            <div className={styles.thumbWrap}>
                                {item.thumbnail ? (
                                    <img src={item.thumbnail} alt="" className={styles.thumbImg} />
                                ) : (
                                    <Icon iconName="News" className={styles.placeholderIcon} />
                                )}
                            </div>
                            <div className={styles.detailGroup}>
                                <div className={styles.title}>{item.title}</div>
                                <div className={styles.date}>{item.dateLabel}</div>
                            </div>
                        </a>
                    ))}
                </div>
            )}
        </WidgetCard>
    );
};

export default NewsWidget;
