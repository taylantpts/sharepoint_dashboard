import * as React from 'react';
import { Persona, PersonaSize, Spinner, SpinnerSize, MessageBar, MessageBarType, Text, IconButton, useTheme, mergeStyleSets } from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import WidgetCard from '../WidgetCard';
import { getBirthdaysThisMonth, IBirthdayItem } from '../../services/SharePointService';
import { DATA_UNAVAILABLE_MESSAGE } from '../../constants';

export interface IBirthdaysWidgetProps {
    context: WebPartContext;
}

type LoadState = 'loading' | 'loaded' | 'error';

const PAGE_SIZE = 10;

const BirthdaysWidget: React.FunctionComponent<IBirthdaysWidgetProps> = (props) => {
    const { context } = props;
    const theme = useTheme();

    const [birthdays, setBirthdays] = React.useState<IBirthdayItem[]>([]);
    const [state, setState] = React.useState<LoadState>('loading');
    const [page, setPage] = React.useState(0);

    React.useEffect(() => {
        let isMounted = true;
        getBirthdaysThisMonth(context)
            .then((items) => {
                if (isMounted) {
                    setBirthdays(items);
                    setState('loaded');
                }
            })
            .catch((error: Error) => {
                console.error('[BirthdaysWidget] Doğum günleri alınamadı:', error);
                if (isMounted) {
                    setState('error');
                }
            });
        return () => {
            isMounted = false;
        };
    }, [context]);

    const pageCount = Math.max(1, Math.ceil(birthdays.length / PAGE_SIZE));
    const paged = birthdays.slice(page * PAGE_SIZE, page * PAGE_SIZE + PAGE_SIZE);

    const styles = mergeStyleSets({
        emptyHint: {
            fontSize: 12,
            color: theme.semanticColors.bodySubtext
        },
        // NOT: "gap" flex özelliği burada BİLİNÇLİ OLARAK KULLANILMIYOR — bu
        // sayfanın render edildiği (kurumsal/eski) tarayıcı ortamında flex "gap"
        // desteklenmiyor. Avatar kendi marginRight'ını taşıyor.
        personRow: {
            display: 'flex',
            alignItems: 'center',
            padding: '5px 0'
        },
        personRowAvatar: {
            marginRight: 10
        },
        badge: {
            marginLeft: 'auto',
            fontSize: 10,
            fontWeight: 700,
            padding: '2px 8px',
            borderRadius: 10,
            background: theme.palette.neutralLighter,
            color: theme.semanticColors.bodySubtext,
            whiteSpace: 'nowrap'
        },
        todayBadge: {
            marginLeft: 'auto',
            fontSize: 10,
            fontWeight: 700,
            padding: '2px 8px',
            borderRadius: 10,
            background: theme.palette.themePrimary,
            color: '#ffffff',
            whiteSpace: 'nowrap'
        },
        pagination: {
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            marginTop: 10
        },
        paginationPrevButton: {
            marginRight: 8
        },
        pageLabel: {
            fontSize: 11,
            color: theme.semanticColors.bodySubtext,
            minWidth: 60,
            textAlign: 'center',
            marginRight: 8
        }
    });

    return (
        <WidgetCard title="Doğum Günleri" subtitle="Bu ay kutlama zamanı" iconName="Balloons">
            {state === 'loading' && <Spinner size={SpinnerSize.small} label="Yükleniyor..." />}
            {state === 'error' && <MessageBar messageBarType={MessageBarType.error}>{DATA_UNAVAILABLE_MESSAGE}</MessageBar>}
            {state === 'loaded' && birthdays.length === 0 && (
                <Text className={styles.emptyHint}>Bu ay içinde yaklaşan doğum günü bulunmuyor.</Text>
            )}
            {state === 'loaded' && paged.map((b) => (
                <div key={b.id} className={styles.personRow}>
                    <Persona text={b.name} secondaryText={b.department} size={PersonaSize.size32} className={styles.personRowAvatar} />
                    {b.isToday ? (
                        <span className={styles.todayBadge}>🎉 İyi ki doğdun!</span>
                    ) : (
                        <span className={styles.badge}>{b.dateLabel}</span>
                    )}
                </div>
            ))}
            {state === 'loaded' && birthdays.length > PAGE_SIZE && (
                <div className={styles.pagination}>
                    <IconButton
                        iconProps={{ iconName: 'ChevronLeft' }}
                        ariaLabel="Önceki sayfa"
                        disabled={page === 0}
                        onClick={() => setPage((p) => Math.max(0, p - 1))}
                        className={styles.paginationPrevButton}
                    />
                    <span className={styles.pageLabel}>Sayfa {page + 1} / {pageCount}</span>
                    <IconButton
                        iconProps={{ iconName: 'ChevronRight' }}
                        ariaLabel="Sonraki sayfa"
                        disabled={page >= pageCount - 1}
                        onClick={() => setPage((p) => Math.min(pageCount - 1, p + 1))}
                    />
                </div>
            )}
        </WidgetCard>
    );
};

export default BirthdaysWidget;
