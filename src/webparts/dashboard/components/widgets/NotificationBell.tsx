import * as React from 'react';
import { Icon, Callout, DirectionalHint, useTheme, mergeStyleSets } from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import {
    getLatestIds, getSeenIds, setSeenIds,
    LatestIds, NotificationCategory,
    NOTIFICATION_CATEGORY_ORDER, NOTIFICATION_CATEGORY_LABELS, NOTIFICATION_CATEGORY_ICONS
} from '../../services/NotificationService';

export interface INotificationBellProps {
    context: WebPartContext;
}

/**
 * Karşılama header'ının sağ üst köşesindeki zil — Duyurular, Etkinlikler,
 * Katılış ve Ayrılış listelerinden herhangi birine yeni bir kayıt
 * eklenmişse kırmızı bir rozet gösterir (bkz. NotificationService —
 * "yenilik", tarih değil listenin en büyük Id'sine bakılarak tespit
 * edilir). "Görüldü" durumu kullanıcının TARAYICISINDA (localStorage)
 * saklanır; ayrı bir SharePoint listesi/altyapısı gerektirmez, ama bu
 * yüzden cihaza özeldir (başka bilgisayarda tekrar "yeni" sayılabilir).
 */
const NotificationBell: React.FunctionComponent<INotificationBellProps> = ({ context }) => {
    const theme = useTheme();
    const loginName = context.pageContext.user.loginName;
    const buttonRef = React.useRef<HTMLButtonElement>(null);

    const [latestIds, setLatestIds] = React.useState<LatestIds | undefined>(undefined);
    const [seenIds, setSeenIdsState] = React.useState<LatestIds | undefined>(undefined);
    const [isOpen, setIsOpen] = React.useState(false);

    React.useEffect(() => {
        let isMounted = true;
        getLatestIds(context)
            .then((ids) => {
                if (!isMounted) {
                    return;
                }
                setLatestIds(ids);
                const stored = getSeenIds(loginName);
                if (!stored) {
                    // İlk çalıştırma: var olan tüm geçmişi "yeni" saymadan,
                    // şu anki durumu taban (baseline) al — sadece bundan
                    // SONRA eklenenler bildirim üretir.
                    setSeenIds(loginName, ids);
                    setSeenIdsState(ids);
                } else {
                    setSeenIdsState(stored);
                }
            })
            .catch((error: Error) => {
                console.error('[NotificationBell] Son kayıtlar alınamadı:', error);
            });
        return () => {
            isMounted = false;
        };
    }, [context, loginName]);

    const newCategories: NotificationCategory[] = latestIds && seenIds
        ? NOTIFICATION_CATEGORY_ORDER.filter((category) => latestIds[category] > seenIds[category])
        : [];

    const handleDismiss = (): void => {
        setIsOpen(false);
        if (latestIds) {
            setSeenIds(loginName, latestIds);
            setSeenIdsState(latestIds);
        }
    };

    const styles = mergeStyleSets({
        button: {
            position: 'relative',
            width: 38,
            height: 38,
            borderRadius: '50%',
            border: '1px solid rgba(15,23,42,0.08)',
            background: 'rgba(255,255,255,0.55)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            cursor: 'pointer',
            padding: 0,
            transition: 'background 0.15s ease, transform 0.15s ease',
            selectors: {
                ':hover': { background: 'rgba(255,255,255,0.9)', transform: 'translateY(-1px)' }
            }
        },
        icon: {
            fontSize: 16,
            color: 'var(--yorpas-theme-header-clock, #1E293B)'
        },
        badge: {
            position: 'absolute',
            top: -2,
            right: -2,
            minWidth: 17,
            height: 17,
            borderRadius: 9,
            background: '#E5484D',
            color: '#ffffff',
            fontSize: 10,
            fontWeight: 700,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            padding: '0 4px',
            border: '2px solid #ffffff',
            boxSizing: 'border-box'
        },
        calloutBody: {
            padding: '16px 18px',
            minWidth: 260,
            maxWidth: 300
        },
        calloutTitle: {
            fontSize: 14,
            fontWeight: 700,
            color: theme.semanticColors.bodyText,
            marginBottom: 12
        },
        row: {
            display: 'flex',
            alignItems: 'center',
            padding: '7px 0'
        },
        rowIconWrap: {
            width: 30,
            height: 30,
            borderRadius: 9,
            background: theme.palette.themeLighter,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            marginRight: 10,
            flexShrink: 0
        },
        rowIcon: {
            fontSize: 14,
            color: theme.palette.themePrimary
        },
        rowLabel: {
            fontSize: 13,
            color: theme.semanticColors.bodyText
        },
        emptyHint: {
            fontSize: 12,
            color: theme.semanticColors.bodySubtext,
            lineHeight: 1.5
        }
    });

    return (
        <>
            <button
                ref={buttonRef}
                type="button"
                className={styles.button}
                onClick={() => setIsOpen(true)}
                aria-label="Bildirimler"
                title="Bildirimler"
            >
                <Icon iconName="Ringer" className={styles.icon} />
                {newCategories.length > 0 && <span className={styles.badge}>{newCategories.length}</span>}
            </button>
            {isOpen && (
                <Callout
                    target={buttonRef.current}
                    onDismiss={handleDismiss}
                    directionalHint={DirectionalHint.bottomRightEdge}
                    gapSpace={10}
                >
                    <div className={styles.calloutBody}>
                        <div className={styles.calloutTitle}>Bildirimler</div>
                        {newCategories.length === 0 ? (
                            <div className={styles.emptyHint}>Yeni bir şey yok, her şeyi görmüşsün.</div>
                        ) : (
                            newCategories.map((category) => (
                                <div key={category} className={styles.row}>
                                    <div className={styles.rowIconWrap}>
                                        <Icon iconName={NOTIFICATION_CATEGORY_ICONS[category]} className={styles.rowIcon} />
                                    </div>
                                    <span className={styles.rowLabel}>{NOTIFICATION_CATEGORY_LABELS[category]}</span>
                                </div>
                            ))
                        )}
                    </div>
                </Callout>
            )}
        </>
    );
};

export default NotificationBell;
