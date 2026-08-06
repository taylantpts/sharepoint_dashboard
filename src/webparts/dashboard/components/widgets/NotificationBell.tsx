import * as React from 'react';
import { Icon, Callout, DirectionalHint, useTheme, mergeStyleSets } from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import {
    INotificationEntry,
    NOTIFICATION_CATEGORY_LABELS, NOTIFICATION_CATEGORY_ICONS,
    NOTIFICATION_CATEGORY_ANCHOR_ID
} from '../../services/NotificationService';
import { scrollToAnchorWithHighlight } from '../../utils/scrollToAnchor';

export interface INotificationBellProps {
    context: WebPartContext;
    /** Bildirim geçmişi ARTIK burada değil, üst bileşende (WelcomeHeader,
     * bkz. useNotificationHistory) TEK bir yerde tutuluyor — hem zil hem
     * özet şeridi çipleri AYNI veriyi okusun diye (bkz. o hook'taki not:
     * önceden ikisi bağımsız senkronize oluyordu ve bu bir yarış durumuna
     * yol açıyordu). NotificationBell artık salt görüntüleyici. */
    history: INotificationEntry[];
    onEntryRead: (entryId: string) => void;
}

/**
 * Karşılama header'ının sağ üst köşesindeki zil — Duyurular, Etkinlikler,
 * Katılış, Ayrılış ve İkinci El listelerinden herhangi birine yeni bir
 * kayıt eklenmişse bir bildirim kaydı üretir (bkz. NotificationService —
 * "yenilik", tarih değil listenin en büyük Id'sine bakılarak tespit
 * edilir). Son MAX_HISTORY (5) bildirim KALICI olarak (okunmuş olsa bile)
 * listede tutulur; bir bildirime tıklamak onu SİLMEZ, sadece "okundu"
 * görünümüne çevirir — 6. yeni bildirim geldiğinde en eski kayıt düşer.
 * Bu geçmiş kullanıcının TARAYICISINDA (localStorage) saklanır; ayrı bir
 * SharePoint listesi/altyapısı gerektirmez, ama bu yüzden cihaza özeldir.
 */
const NotificationBell: React.FunctionComponent<INotificationBellProps> = ({ history, onEntryRead }) => {
    const theme = useTheme();
    const buttonRef = React.useRef<HTMLButtonElement>(null);

    const [isOpen, setIsOpen] = React.useState(false);

    const unreadCount = history.filter((entry) => !entry.read).length;

    const handleDismiss = (): void => {
        setIsOpen(false);
    };

    /**
     * Bir bildirim satırına tıklanınca: o kayıt "okundu" işaretlenir (listeden
     * SİLİNMEZ), sayfa ilgili widget'a YUMUŞAKÇA kaydırılır ve o widget'ın
     * etrafında kısa süreli bir "ışıma" halkası belirip söner (bkz.
     * utils/scrollToAnchor.ts — WelcomeHeader'daki özet şeridiyle de
     * paylaşılan ortak yardımcı).
     *
     * ÖNCEKİ HATA: menü kapanması İLE AYNI ANDA (senkron) tetikleniyordu.
     * Fluent'in Callout'u varsayılan olarak kapanınca focus'u tetikleyici
     * butona (zile) GERİ VERİR (shouldRestoreFocus, bkz. Callout.types.d.ts)
     * — tarayıcı da focus alan bir elemanı otomatik görünüre kaydırır. Bu,
     * header'daki zil sayfanın EN ÜSTÜNDE olduğu için, tam da bizim
     * kaydırma animasyonumuz başlarken sayfayı GERİ en tepeye çekiyor ve
     * "bildirime tıklayınca ilgili widget'a kaymıyor" hissi yaratıyordu.
     * Çözüm iki parça: Callout'a shouldRestoreFocus={false} verildi VE
     * menü kapatma, kaydırma animasyonu bitene kadar ERTELENDİ.
     */
    const handleNotificationClick = (entry: INotificationEntry): void => {
        onEntryRead(entry.id);
        scrollToAnchorWithHighlight(NOTIFICATION_CATEGORY_ANCHOR_ID[entry.category], theme.palette.themePrimary);
        window.setTimeout(() => setIsOpen(false), 650);
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
        // Okunmamış bildirimler renkli/canlı, okunanlar soluk/gri — kullanıcının
        // "tıklananlar farklı görünecek, yeni bildirimler farklı görünecek"
        // isteğinin karşılığı. Satır ARKA PLANI da hafifçe farklı (okunmamışta
        // çok soluk bir vurgu zemini) ki liste taranırken göz hemen ayırt etsin.
        rowUnread: {
            background: theme.palette.themeLighter + '55'
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
        rowIconWrapRead: {
            background: theme.semanticColors.disabledBackground
        },
        rowIcon: {
            fontSize: 14,
            color: theme.palette.themePrimary
        },
        rowIconRead: {
            color: theme.semanticColors.disabledText
        },
        rowLabel: {
            flexGrow: 1,
            fontSize: 13,
            fontWeight: 600,
            color: theme.semanticColors.bodyText
        },
        rowLabelRead: {
            fontWeight: 400,
            color: theme.semanticColors.bodySubtext
        },
        unreadDot: {
            width: 7,
            height: 7,
            borderRadius: '50%',
            background: theme.palette.themePrimary,
            marginLeft: 8,
            flexShrink: 0
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
                {unreadCount > 0 && <span className={styles.badge}>{unreadCount}</span>}
            </button>
            {isOpen && (
                <Callout
                    target={buttonRef.current}
                    onDismiss={handleDismiss}
                    directionalHint={DirectionalHint.bottomRightEdge}
                    gapSpace={10}
                    shouldRestoreFocus={false}
                >
                    <div className={styles.calloutBody}>
                        <div className={styles.calloutTitle}>Bildirimler</div>
                        {history.length === 0 ? (
                            <div className={styles.emptyHint}>Yeni bir şey yok, her şeyi görmüşsün.</div>
                        ) : (
                            history.map((entry) => (
                                <button
                                    key={entry.id}
                                    type="button"
                                    className={`${styles.row} ${!entry.read ? styles.rowUnread : ''}`}
                                    onClick={() => handleNotificationClick(entry)}
                                >
                                    <div className={`${styles.rowIconWrap} ${entry.read ? styles.rowIconWrapRead : ''}`}>
                                        <Icon
                                            iconName={NOTIFICATION_CATEGORY_ICONS[entry.category]}
                                            className={`${styles.rowIcon} ${entry.read ? styles.rowIconRead : ''}`}
                                        />
                                    </div>
                                    <span className={`${styles.rowLabel} ${entry.read ? styles.rowLabelRead : ''}`}>
                                        {NOTIFICATION_CATEGORY_LABELS[entry.category]}
                                    </span>
                                    {!entry.read && <span className={styles.unreadDot} aria-hidden="true" />}
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
