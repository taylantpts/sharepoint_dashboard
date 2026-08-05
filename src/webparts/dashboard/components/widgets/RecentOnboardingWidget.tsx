import * as React from 'react';
import { Persona, PersonaSize, Spinner, SpinnerSize, MessageBar, MessageBarType, Text, useTheme, mergeStyleSets } from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import WidgetCard from '../WidgetCard';
import { getRecentOnboardingRecords, OnboardingKind, IOnboardingRecord } from '../../services/OnboardingService';
import { DATA_UNAVAILABLE_MESSAGE } from '../../constants';

export interface IRecentOnboardingWidgetProps {
    context: WebPartContext;
    kind: OnboardingKind;
    title: string;
    subtitle: string;
    iconName: string;
    emptyHint: string;
}

type LoadState = 'loading' | 'loaded' | 'error';

const RECENT_COUNT = 5;

/**
 * Herkese açık, salt-okunur "son 5 kişi" özeti — tüm çalışma listesini,
 * checklist'i veya düzenleme aracını göstermez (bunlar için bkz.
 * OnboardingTrackerWidget, sadece İK/BT/Muhasebe grupları erişebiliyor).
 */
const RecentOnboardingWidget: React.FunctionComponent<IRecentOnboardingWidgetProps> = (props) => {
    const { context, kind, title, subtitle, iconName, emptyHint } = props;
    const theme = useTheme();

    const [records, setRecords] = React.useState<IOnboardingRecord[]>([]);
    const [state, setState] = React.useState<LoadState>('loading');

    React.useEffect(() => {
        let isMounted = true;
        getRecentOnboardingRecords(context, kind, RECENT_COUNT)
            .then((result) => {
                if (isMounted) {
                    setRecords(result);
                    setState('loaded');
                }
            })
            .catch((error: Error) => {
                console.error('[RecentOnboardingWidget] Kayıtlar alınamadı:', error);
                if (isMounted) {
                    setState('error');
                }
            });
        return () => {
            isMounted = false;
        };
    }, [context, kind]);

    const styles = mergeStyleSets({
        emptyHint: {
            fontSize: 12,
            color: theme.semanticColors.bodySubtext
        },
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
        }
    });

    return (
        <WidgetCard title={title} subtitle={subtitle} iconName={iconName}>
            {state === 'loading' && <Spinner size={SpinnerSize.small} label="Yükleniyor..." />}
            {state === 'error' && <MessageBar messageBarType={MessageBarType.error}>{DATA_UNAVAILABLE_MESSAGE}</MessageBar>}
            {state === 'loaded' && records.length === 0 && <Text className={styles.emptyHint}>{emptyHint}</Text>}
            {state === 'loaded' && records.map((record) => (
                <div key={record.id} className={styles.personRow}>
                    <Persona text={record.name} secondaryText={record.location || undefined} size={PersonaSize.size32} className={styles.personRowAvatar} />
                    <span className={styles.badge}>{record.dateLabel}</span>
                </div>
            ))}
        </WidgetCard>
    );
};

export default RecentOnboardingWidget;
