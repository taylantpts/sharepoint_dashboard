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
    // ÖNCEKİ HATA: kart zemini, mavi vurgu rengiyle karışık görünür bir
    // gradyandı (accent %14 + cardAmbient %28) — kullanıcı bunu hâlâ "mavi
    // gradyan" olarak algılayıp minimale çekilmesini istedi. Artık zemin
    // DÜZ (gradyansız) ve neredeyse beyaz (%6 gibi çok hafif bir "wash") —
    // "kaliteli/dikkat çekici" hissi artık renkten değil, İNCE detaylardan
    // geliyor: net bir kenarlık, iki katmanlı yumuşak gölge, ve üst aksan
    // şeridinin altındaki hafif "ışıma" (glow). Renk sadece üst şeritte ve
    // ikon rozetinde — tüm kart yüzeyine YAYILMIYOR.
    const timeBucket = useTimeAwareTheme();
    const cardAmbient = getTimeTheme(timeBucket).cardAmbient;
    const surface = mixHex(cardAmbient, theme.palette.white, 0.06);

    const styles = mergeStyleSets({
        card: {
            position: 'relative',
            background: surface,
            border: '1px solid rgba(15,23,42,0.06)',
            borderRadius: 22,
            // İki katmanlı gölge — biri kartın hemen altında ince/keskin
            // (contact shadow), diğeri daha geniş/yumuşak (ambient shadow).
            // Tek büyük/bulanık gölge yerine bu ikili, kartın "kalitesini"
            // asıl belirleyen ince detay: hafif ama net bir yükseklik hissi.
            boxShadow: '0 1px 2px rgba(15,23,42,0.04), 0 10px 24px rgba(15,23,42,0.06)',
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
                    height: 3,
                    borderRadius: '0 0 4px 4px',
                    background: `linear-gradient(90deg, ${accent} 0%, ${hexToRgba(accent, 0.5)} 100%)`,
                    // Şeridin altında hafif bir "ışıma" — kartın tamamını
                    // boyamadan, tek bir ince çizgide dikkat çekici bir
                    // premium detay.
                    boxShadow: `0 2px 8px ${hexToRgba(accent, 0.30)}`
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
