import * as React from 'react';
import { Icon, Callout, DirectionalHint, useTheme, mergeStyleSets } from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import {
    getLatestIds, getSeenIds, setSeenIds,
    LatestIds, NotificationCategory,
    NOTIFICATION_CATEGORY_ORDER, NOTIFICATION_CATEGORY_LABELS, NOTIFICATION_CATEGORY_ICONS,
    NOTIFICATION_CATEGORY_ANCHOR_ID
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

    /**
     * Bir bildirim satırına tıklanınca: callout kapanır (görüldü sayılır,
     * handleDismiss ile aynı), sayfa ilgili widget'a kaydırılır (bkz.
     * Dashboard.tsx'teki id'ler ve NOTIFICATION_CATEGORY_ANCHOR_ID) ve o
     * widget'ın etrafında kısa süreli bir "ışıma" halkası belirip sönerek
     * kullanıcıya TAM OLARAK hangi karta gittiğini gösterir — sadece
     * kaydırmak, özellikle birbirine yakın/benzer görünen kartlar arasında
     * hangisinin hedef olduğunu belli etmiyordu.
     * ÖNCEKİ HATA: scrollIntoView({behavior:'smooth'}) bu render ortamında
     * SESSİZCE hiçbir şey yapmıyordu (aynı "modern CSS/davranış seçeneği bu
     * ortamda çalışmıyor" kalıbı — bkz. lineHeight/flex-direction notları) —
     * fonksiyon hatasız çalışıyor, hedef doğru bulunuyor ama sayfa YERİNDE
     * kalıyordu. 'auto' (anlık atlama) burada güvenilir şekilde çalışıyor.
     */
    const handleNotificationClick = (category: NotificationCategory): void => {
        handleDismiss();
        const target = document.getElementById(NOTIFICATION_CATEGORY_ANCHOR_ID[category]);
        if (!target) {
            return;
        }
        target.scrollIntoView({ behavior: 'auto', block: 'center' });
        const prevTransition = target.style.transition;
        const prevBoxShadow = target.style.boxShadow;
        const prevBorderRadius = target.style.borderRadius;
        target.style.transition = 'box-shadow 0.3s ease';
        target.style.borderRadius = target.style.borderRadius || '22px';
        target.style.boxShadow = `0 0 0 3px ${theme.palette.themePrimary}, 0 8px 24px ${theme.palette.themePrimary}55`;
        window.setTimeout(() => {
            target.style.boxShadow = prevBoxShadow;
            window.setTimeout(() => {
                target.style.transition = prevTransition;
                target.style.borderRadius = prevBorderRadius;
            }, 320);
        }, 1400);
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
        // ÖNCEKİ HATA: satırlar sade, tıklanamaz bir <div> idi — kullanıcı bir
        // bildirime tıklayınca ilgili widget'a kaydırılmasını istedi. Artık
        // tam genişlikte bir <button>: cursor:pointer + hover zemini, buranın
        // tıklanabilir bir kontrol olduğunu görsel olarak da anlatıyor.
        row: {
            display: 'flex',
            alignItems: 'center',
            width: '100%',
            padding: '7px 8px',
            margin: '0 -8px',
            border: 'none',
            background: 'transparent',
            borderRadius: 8,
            cursor: 'pointer',
            textAlign: 'left',
            font: 'inherit',
            transition: 'background 0.15s ease',
            selectors: {
                ':hover': { background: theme.palette.neutralLighterAlt }
            }
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
            // ÖNCEKİ HATA: birimsiz sayı (1.5) bu render ortamında "1.5px" olarak
            // hesaplanıyor (bkz. DetailModal.tsx'teki aynı not) — yüzde string'i kullanılıyor.
            lineHeight: '150%'
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
                                <button
                                    key={category}
                                    type="button"
                                    className={styles.row}
                                    onClick={() => handleNotificationClick(category)}
                                >
                                    <div className={styles.rowIconWrap}>
                                        <Icon iconName={NOTIFICATION_CATEGORY_ICONS[category]} className={styles.rowIcon} />
                                    </div>
                                    <span className={styles.rowLabel}>{NOTIFICATION_CATEGORY_LABELS[category]}</span>
                                </button>
                            ))
                        )}
                    </div>
                </Callout>
            )}
        </>
    );
};

export default NotificationBell;
