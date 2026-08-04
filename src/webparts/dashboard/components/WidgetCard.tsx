import * as React from 'react';
import { Icon, useTheme, mergeStyleSets } from '@fluentui/react';
import { useTimeAwareTheme, getTimeTheme } from '../themeManager';

export interface IWidgetCardProps {
    title: string;
    subtitle?: string;
    iconName?: string;
    headerAction?: React.ReactNode;
    children?: React.ReactNode;
    className?: string;
}

const hexToRgb = (hex: string): [number, number, number] => {
    const clean = hex.replace('#', '');
    const bigint = parseInt(clean.length === 3 ? clean.split('').map((c) => c + c).join('') : clean, 16);
    return [(bigint >> 16) & 255, (bigint >> 8) & 255, bigint & 255];
};

const parseColorToRgb = (color: string): [number, number, number] => {
    const rgbMatch = /rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/.exec(color);
    if (rgbMatch) {
        return [parseInt(rgbMatch[1], 10), parseInt(rgbMatch[2], 10), parseInt(rgbMatch[3], 10)];
    }
    return hexToRgb(color);
};

/** "#0078d4" -> "rgba(0,120,212,alpha)" — ikon rozetindeki renkli parlama (glow) gölgesi için. */
const hexToRgba = (hex: string, alpha: number): string => {
    const [r, g, b] = parseColorToRgb(hex);
    return `rgba(${r}, ${g}, ${b}, ${alpha})`;
};

/** İki rengi OPAK (alfasız) bir tona karıştırır — kartın zemin degradesinin ilk durağı için kullanılıyor. */
const mixHex = (hex: string, baseHex: string, weight: number): string => {
    const [r1, g1, b1] = parseColorToRgb(hex);
    const [r2, g2, b2] = parseColorToRgb(baseHex);
    const mix = (a: number, b: number): number => Math.round(a * weight + b * (1 - weight));
    return `rgb(${mix(r1, r2)}, ${mix(g1, g2)}, ${mix(b1, b2)})`;
};

/**
 * Tüm widget'ların ortak kabuğu. ÖNCEKİ HATA (birden fazla iterasyon):
 * her widget kendi `accentColor` prop'unu geçiyordu (Duyurular turuncu,
 * Etkinlikler mor, ...) ve bu, saat dilimine göre değişen bir "mevsim"
 * rengiyle harmanlanıyordu — kullanıcı bunu "renkli/dağınık" bulup TEK,
 * sabit bir KURUMSAL tema istedi (referans: Klarinet Velocity — lacivert/
 * beyaz zemin + tek bir vurgu rengi). Artık `accentColor` prop'u YOK: ikon
 * rozeti, üst aksan şeridi ve hover halkası HER widget'ta aynı kurumsal mavi
 * (`theme.palette.themePrimary`, #0078d4) — tek fark, kartın hangi ikonu ve
 * başlığı taşıdığı.
 */
const WidgetCard: React.FunctionComponent<IWidgetCardProps> = (props) => {
    const { title, subtitle, iconName, headerAction, children, className } = props;
    const theme = useTheme();
    const accent = theme.palette.themePrimary;

    // useTimeAwareTheme hâlâ çağrılıyor çünkü :root CSS değişkenlerini
    // (WelcomeHeader'ın okuduğu) güncel tutuyor — ama getTimeTheme artık
    // HER bucket için aynı sabit cardAmbient'i döndürüyor (bkz.
    // themeManager.ts), yani bu satırlar kartın rengini zamana göre
    // DEĞİŞTİRMEZ; sadece tek, sabit kurumsal tonu okur.
    const timeBucket = useTimeAwareTheme();
    const cardAmbient = getTimeTheme(timeBucket).cardAmbient;
    const ambientBase = mixHex(cardAmbient, theme.palette.neutralLighterAlt, 0.28);

    const styles = mergeStyleSets({
        card: {
            position: 'relative',
            background: `linear-gradient(165deg, ${mixHex(accent, ambientBase, 0.14)} 0%, ${ambientBase} 42%)`,
            backdropFilter: 'blur(14px)',
            WebkitBackdropFilter: 'blur(14px)',
            border: 'none',
            borderRadius: 22,
            boxShadow: '0 8px 24px rgba(15,23,42,0.10)',
            padding: 24,
            display: 'flex',
            flexDirection: 'column',
            height: '100%',
            boxSizing: 'border-box',
            selectors: {
                '::before': {
                    content: '""',
                    position: 'absolute',
                    top: 0,
                    left: 14,
                    right: 14,
                    height: 4,
                    borderRadius: '0 0 4px 4px',
                    background: `linear-gradient(90deg, ${accent} 0%, ${hexToRgba(accent, 0.5)} 100%)`
                }
            }
        },
        header: {
            marginBottom: 16
        },
        // NOT: "gap" flex özelliği burada da BİLİNÇLİ OLARAK KULLANILMIYOR — bu
        // sayfanın render edildiği (kurumsal/eski) tarayıcı ortamında flex "gap"
        // desteklenmiyor. titleGroup, titleRow'un headerAction'dan önce gelen
        // TEK çocuğu olduğu için buraya doğrudan marginRight veriliyor.
        titleRow: {
            display: 'flex',
            alignItems: 'flex-start',
            justifyContent: 'space-between'
        },
        titleGroup: {
            display: 'flex',
            alignItems: 'center',
            marginRight: 12
        },
        iconWrap: {
            width: 34,
            height: 34,
            borderRadius: 9,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            background: `linear-gradient(135deg, ${accent} 0%, ${hexToRgba(accent, 0.72)} 100%)`,
            boxShadow: `0 4px 10px ${hexToRgba(accent, 0.35)}`,
            flexShrink: 0,
            // titleGroup içindeki bir sonraki çocuktan (isim/unvan bloğu) önce
            // gelen boşluk — "gap" yerine.
            marginRight: 10
        },
        icon: {
            fontSize: 16,
            color: '#ffffff'
        },
        title: {
            fontWeight: 600,
            fontSize: 15,
            color: theme.semanticColors.bodyText
        },
        subtitle: {
            fontSize: 12,
            color: theme.semanticColors.bodySubtext,
            marginTop: 2
        },
        body: {
            flexGrow: 1,
            minWidth: 0
        }
    });

    return (
        <section className={[styles.card, className].filter(Boolean).join(' ')}>
            <div className={styles.header}>
                <div className={styles.titleRow}>
                    <div className={styles.titleGroup}>
                        {iconName && (
                            <div className={styles.iconWrap}>
                                <Icon iconName={iconName} className={styles.icon} />
                            </div>
                        )}
                        <div>
                            <div className={styles.title}>{title}</div>
                            {subtitle && <div className={styles.subtitle}>{subtitle}</div>}
                        </div>
                    </div>
                    {headerAction}
                </div>
            </div>
            <div className={styles.body}>{children}</div>
        </section>
    );
};

export default WidgetCard;
