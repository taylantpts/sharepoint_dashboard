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
    countdownLabel: string;
}

type LoadState = 'loading' | 'loaded' | 'error';

const TURKISH_MONTH_SHORT = ['OCA', 'ŞUB', 'MAR', 'NİS', 'MAY', 'HAZ', 'TEM', 'AĞU', 'EYL', 'EKİ', 'KAS', 'ARA'];

const MAX_ITEMS = 5;

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
        row: {
            display: 'flex',
            alignItems: 'center',
            padding: '10px 12px',
            marginBottom: 10,
            borderRadius: 10
        },
        leaf: {
            display: 'flex',
            flexDirection: 'column',
            width: 52,
            borderRadius: 8,
            overflow: 'hidden',
            boxShadow: '0 2px 6px rgba(0,0,0,0.14)',
            flexShrink: 0,
            marginRight: 14
        },
        leafMonth: {
            background: '#7a2e26',
            color: '#ffffff',
            fontSize: 11,
            fontWeight: 700,
            textAlign: 'center',
            padding: '5px 0',
            lineHeight: '14px',
            letterSpacing: 0.6
        },
        leafDay: {
            background: '#ffffff',
            color: theme.semanticColors.bodyText,
            fontSize: 20,
            fontWeight: 800,
            textAlign: 'center',
            padding: '5px 0 7px',
            lineHeight: '26px'
        },
        detailGroup: {
            display: 'flex',
            flexDirection: 'column',
            minWidth: 0,
            flexGrow: 1
        },
        name: {
            fontSize: 14,
            fontWeight: 700,
            color: theme.semanticColors.bodyText,
            overflow: 'hidden',
            textOverflow: 'ellipsis',
            whiteSpace: 'nowrap',
            marginBottom: 2
        },
        countdown: {
            fontSize: 12,
            color: theme.semanticColors.bodySubtext
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
                            <div className={styles.leaf}>
                                <div className={styles.leafMonth}>{h.monthShort}</div>
                                <div className={styles.leafDay}>{h.day}</div>
                            </div>
                            <div className={styles.detailGroup}>
                                <div className={styles.name}>{h.name}</div>
                                <div className={styles.countdown}>{h.countdownLabel}</div>
                            </div>
                        </div>
                    ))}
                </div>
            )}
        </WidgetCard>
    );
};

export default HolidaysWidget;
