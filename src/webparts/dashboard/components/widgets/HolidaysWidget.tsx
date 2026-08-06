import * as React from 'react';
import { Spinner, SpinnerSize, MessageBar, MessageBarType, useTheme, mergeStyleSets } from '@fluentui/react';
import WidgetCard from '../WidgetCard';
import { DATA_UNAVAILABLE_MESSAGE } from '../../constants';

interface INagerHoliday {
    date: string;
    localName: string;
    name: string;
}

interface IHolidayItem {
    date: string;
    name: string;
    monthShort: string;
    day: number;
    dayOfWeek: string;
    countdownLabel: string;
}

type LoadState = 'loading' | 'loaded' | 'error';

const TURKISH_MONTH_SHORT = ['OCA', 'ŞUB', 'MAR', 'NİS', 'MAY', 'HAZ', 'TEM', 'AĞU', 'EYL', 'EKİ', 'KAS', 'ARA'];
const TURKISH_DAY_OF_WEEK = ['Pazar', 'Pazartesi', 'Salı', 'Çarşamba', 'Perşembe', 'Cuma', 'Cumartesi'];

// ÖNCEKİ HATA: sadece 5 tatil gösteriliyordu ve her satır ("takvim yaprağı"
// + bol padding) çok yer kaplıyordu — kullanıcı satırların daraltılıp daha
// FAZLA tatilin gösterilmesini istedi. Satırlar kompakt tek-satırlık bir
// listeye dönüştürüldükten sonra 5 yerine 14 tatil rahatça sığıyor (~3-4 ay).
const MAX_ITEMS = 14;

const fetchHolidaysForYear = async (year: number): Promise<INagerHoliday[]> => {
    const response = await fetch(`https://date.nager.at/api/v3/PublicHolidays/${year}/TR`);
    if (!response.ok) {
        throw new Error(`Nager.Date isteği başarısız (HTTP ${response.status})`);
    }
    return response.json();
};

const buildCountdownLabel = (daysUntil: number): string => {
    if (daysUntil === 0) {
        return 'Bugün';
    }
    if (daysUntil === 1) {
        return 'Yarın';
    }
    return `${daysUntil} gün kaldı`;
};

const HolidaysWidget: React.FunctionComponent = () => {
    const theme = useTheme();
    const [state, setState] = React.useState<LoadState>('loading');
    const [holidays, setHolidays] = React.useState<IHolidayItem[]>([]);

    React.useEffect(() => {
        let isMounted = true;

        const load = async (): Promise<void> => {
            try {
                const today = new Date();
                today.setHours(0, 0, 0, 0);
                const currentYear = today.getFullYear();

                let raw = await fetchHolidaysForYear(currentYear);
                let upcoming = raw.filter((h) => new Date(`${h.date}T00:00:00`).getTime() >= today.getTime());

                // Yıl sonuna yaklaşıldığında bu yıldan az tatil kalıyorsa, listeyi
                // doldurmak için bir sonraki yılın tatilleri de eklenir.
                if (upcoming.length < MAX_ITEMS) {
                    const nextYearRaw = await fetchHolidaysForYear(currentYear + 1);
                    upcoming = upcoming.concat(nextYearRaw);
                }

                upcoming.sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime());

                const items: IHolidayItem[] = upcoming.slice(0, MAX_ITEMS).map((h) => {
                    const holidayDate = new Date(`${h.date}T00:00:00`);
                    const daysUntil = Math.round((holidayDate.getTime() - today.getTime()) / 86400000);
                    return {
                        date: h.date,
                        name: h.localName,
                        monthShort: TURKISH_MONTH_SHORT[holidayDate.getMonth()],
                        day: holidayDate.getDate(),
                        dayOfWeek: TURKISH_DAY_OF_WEEK[holidayDate.getDay()],
                        countdownLabel: buildCountdownLabel(daysUntil)
                    };
                });

                if (isMounted) {
                    setHolidays(items);
                    setState('loaded');
                }
            } catch (error) {
                console.error('[HolidaysWidget] Resmi tatiller alınamadı:', error);
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
        // ÖNCEKİ HATA: her satır 52px genişliğinde iki katlı bir "takvim
        // yaprağı" + bol padding taşıyordu (~72px yükseklik/satır). Artık tek
        // satırlık, kompakt bir liste: küçük bir tarih rozeti + isim + gün adı
        // yan yana, satır başına ~40px.
        row: {
            display: 'flex',
            alignItems: 'center',
            padding: '6px 4px',
            borderRadius: 8,
            borderBottom: `1px solid ${theme.palette.neutralLighter}`,
            selectors: {
                ':last-child': { borderBottom: 'none' }
            }
        },
        dateBadge: {
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            justifyContent: 'center',
            width: 40,
            height: 36,
            borderRadius: 8,
            background: '#7a2e26',
            color: '#ffffff',
            flexShrink: 0,
            marginRight: 12
        },
        dateBadgeDay: {
            fontSize: 15,
            fontWeight: 800,
            lineHeight: '17px'
        },
        dateBadgeMonth: {
            fontSize: 9,
            fontWeight: 700,
            letterSpacing: 0.4,
            lineHeight: '11px'
        },
        detailGroup: {
            display: 'flex',
            flexDirection: 'column',
            minWidth: 0,
            flexGrow: 1
        },
        name: {
            fontSize: 13,
            fontWeight: 700,
            color: theme.semanticColors.bodyText,
            overflow: 'hidden',
            textOverflow: 'ellipsis',
            whiteSpace: 'nowrap'
        },
        // NOT: "gap" burada kullanılmıyor (flex "gap" bu render ortamında
        // desteklenmiyor) — dayOfWeek'e marginRight verildi.
        metaRow: {
            display: 'flex',
            alignItems: 'center',
            fontSize: 11,
            color: theme.semanticColors.bodySubtext
        },
        dayOfWeek: {
            marginRight: 8
        },
        countdown: {
            flexShrink: 0,
            marginLeft: 10,
            fontSize: 11,
            fontWeight: 700,
            color: theme.palette.themePrimary,
            whiteSpace: 'nowrap'
        },
        emptyHint: {
            fontSize: 12,
            color: theme.semanticColors.bodySubtext
        }
    });

    return (
        <WidgetCard title="Resmi Tatiller" subtitle="Yaklaşan resmi tatiller" iconName="Flag">
            {state === 'loading' && <Spinner size={SpinnerSize.medium} label="Tatiller yükleniyor..." />}
            {state === 'error' && <MessageBar messageBarType={MessageBarType.error}>{DATA_UNAVAILABLE_MESSAGE}</MessageBar>}
            {state === 'loaded' && holidays.length === 0 && (
                <div className={styles.emptyHint}>Yaklaşan resmi tatil bulunmuyor.</div>
            )}
            {state === 'loaded' && holidays.length > 0 && (
                <div className={styles.list}>
                    {holidays.map((h) => (
                        <div key={h.date} className={styles.row}>
                            <div className={styles.dateBadge}>
                                <span className={styles.dateBadgeDay}>{h.day}</span>
                                <span className={styles.dateBadgeMonth}>{h.monthShort}</span>
                            </div>
                            <div className={styles.detailGroup}>
                                <div className={styles.name}>{h.name}</div>
                                <div className={styles.metaRow}>
                                    <span className={styles.dayOfWeek}>{h.dayOfWeek}</span>
                                </div>
                            </div>
                            <div className={styles.countdown}>{h.countdownLabel}</div>
                        </div>
                    ))}
                </div>
            )}
        </WidgetCard>
    );
};

export default HolidaysWidget;
