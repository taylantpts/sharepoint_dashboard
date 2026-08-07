import * as React from 'react';
import { Icon, useTheme, mergeStyleSets } from '@fluentui/react';
import WidgetCard from '../WidgetCard';

const SUPPORT_MAILTO = 'mailto:destek@komagene.com.tr?subject=Destek%20Talebi';

const ITSupportCenter: React.FunctionComponent = () => {
    const theme = useTheme();

    const styles = mergeStyleSets({
        // NOT: "gap" flex özelliği burada BİLİNÇLİ OLARAK KULLANILMIYOR — bu
        // sayfanın render edildiği (kurumsal/eski) tarayıcı ortamında flex "gap"
        // desteklenmiyor. `flexWrap: 'wrap'` olduğu için iconWrap/text kendi
        // marginRight + marginBottom'larını taşıyor; son çocuk (ticket) dokunulmadan kalıyor.
        // ÖNCEKİ HATA: WidgetCard'ın gövdesi (.body, bkz. WidgetCard.tsx) zaten
        // kartın tüm kalan yüksekliğini dolduruyor (flexGrow:1) — ama banner
        // kendi DOĞAL/içerik yüksekliğinde kalıyordu, bu yüzden komşu widget
        // (aynı grid satırındaki) daha uzunsa altında boş beyaz alan kalıyordu.
        // height:'100%' vermek, zaten var olan alignItems:'center' sayesinde
        // içeriği otomatik olarak dikey ortalıyor — "Zimmetlerim" ile aynı
        // kod/tasarım olduğu için SADECE bu widget'a özgü bir görsel fark
        // kalmıyor.
        banner: {
            display: 'flex',
            alignItems: 'center',
            height: '100%',
            boxSizing: 'border-box',
            padding: '28px 32px',
            borderRadius: 12,
            background: `linear-gradient(135deg, ${theme.palette.themeLighterAlt} 0%, ${theme.palette.themeLighter} 100%)`,
            flexWrap: 'wrap'
        },
        iconWrap: {
            width: 68,
            height: 68,
            borderRadius: '50%',
            background: theme.palette.neutralLighterAlt,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            flexShrink: 0,
            boxShadow: '0 1.6px 3.6px rgba(0,0,0,0.08)',
            marginRight: 20,
            marginBottom: 20
        },
        icon: {
            fontSize: 32,
            color: theme.palette.themePrimary
        },
        text: {
            flexGrow: 1,
            minWidth: 200,
            marginRight: 20,
            marginBottom: 20
        },
        title: {
            fontSize: 18,
            fontWeight: 700,
            color: theme.palette.themeDark
        },
        subtitle: {
            fontSize: 12,
            color: theme.semanticColors.bodySubtext,
            marginTop: 4
        },
        // "Ticket" (bilet) görünümü: ikon bölmesi | kesikli çizgi (perforasyon) | metin bölmesi.
        ticket: {
            display: 'flex',
            alignItems: 'stretch',
            textDecoration: 'none',
            background: `linear-gradient(135deg, ${theme.palette.themePrimary} 0%, ${theme.palette.themeDark} 100%)`,
            borderRadius: 10,
            overflow: 'hidden',
            boxShadow: '0 4px 14px rgba(0,0,0,0.20)',
            cursor: 'pointer',
            transition: 'transform 0.15s ease, box-shadow 0.15s ease',
            selectors: {
                ':hover': {
                    transform: 'translateY(-2px)',
                    boxShadow: '0 6px 18px rgba(0,0,0,0.26)'
                }
            }
        },
        ticketIconWrap: {
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            width: 56,
            flexShrink: 0,
            background: 'rgba(255,255,255,0.14)'
        },
        ticketIcon: {
            fontSize: 22,
            color: '#ffffff'
        },
        ticketDivider: {
            borderLeft: '2px dashed rgba(255,255,255,0.45)',
            margin: '10px 0'
        },
        ticketLabel: {
            display: 'flex',
            flexDirection: 'column',
            justifyContent: 'center',
            padding: '10px 22px',
            color: '#ffffff'
        },
        ticketLabelMain: {
            fontSize: 14,
            fontWeight: 700
        },
        ticketLabelSub: {
            fontSize: 11,
            color: 'rgba(255,255,255,0.82)',
            marginTop: 2
        }
    });

    return (
        <WidgetCard title="BT Destek Merkezi" subtitle="Sorunlarınızı tek tıkla iletin" iconName="Chat">
            <div className={styles.banner}>
                <div className={styles.iconWrap}>
                    <Icon iconName="Chat" className={styles.icon} />
                </div>
                <div className={styles.text}>
                    <div className={styles.title}>Bir sorunuz mu var?</div>
                    <div className={styles.subtitle}>Arıza, donanım talebi veya herhangi bir BT sorunu için tek tıkla destek ekibine ulaşın.</div>
                </div>
                <a href={SUPPORT_MAILTO} className={styles.ticket}>
                    <span className={styles.ticketIconWrap}>
                        <Icon iconName="Ticket" className={styles.ticketIcon} />
                    </span>
                    <span className={styles.ticketDivider} />
                    <span className={styles.ticketLabel}>
                        <span className={styles.ticketLabelMain}>Destek Talebi Oluştur</span>
                        <span className={styles.ticketLabelSub}>Tek tıkla BT ekibine ulaşın</span>
                    </span>
                </a>
            </div>
        </WidgetCard>
    );
};

export default ITSupportCenter;
