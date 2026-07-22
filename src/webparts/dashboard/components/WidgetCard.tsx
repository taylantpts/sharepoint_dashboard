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
    /** Kartın üst kenarına ve ikon rozetine renk kimliği verir — verilmezse tema ana rengi kullanılır. */
    accentColor?: string;
}

const hexToRgb = (hex: string): [number, number, number] => {
    const clean = hex.replace('#', '');
    const bigint = parseInt(clean.length === 3 ? clean.split('').map((c) => c + c).join('') : clean, 16);
    return [(bigint >> 16) & 255, (bigint >> 8) & 255, bigint & 255];
};

/** "#0078d4" -> "rgba(0,120,212,alpha)" — ikon rozetindeki renkli parlama (glow) gölgesi için. */
const hexToRgba = (hex: string, alpha: number): string => {
    const [r, g, b] = hexToRgb(hex);
    return `rgba(${r}, ${g}, ${b}, ${alpha})`;
};

/**
 * İki rengi OPAK (alfasız) bir tona karıştırır — kartın zemin degradesinin
 * ilk durağı için kullanılıyor. ÖNCEKİ HATA: burada hexToRgba(accent, 0.10)
 * gibi %90 şeffaf bir renk vardı ve bunun "buzlu cam" gibi durması TAMAMEN
 * backdrop-filter: blur()'a bağlıydı; bu CSS özelliğini desteklemeyen
 * tarayıcılarda (bu portalın render edildiği ortamda TAM OLARAK bu oluyor —
 * flex "gap" ile aynı sınıf sorun) şeffaf katmanın ARKASINDAKİ koyu sayfa
 * zemini neredeyse ham hâliyle sızıyor, kartın üst-sol köşesini karanlık/
 * bulanık, okunaksız bir yama gibi gösteriyordu (bkz. ekran görüntüsü).
 * mixHex, aynı "hafif renkli soluk" hissini backdrop-filter'a hiç ihtiyaç
 * duymadan, %100 opak bir renk üreterek verir — hangi tarayıcıda olursa
 * olsun HER ZAMAN aynı görünür.
 */
const mixHex = (hex: string, baseHex: string, weight: number): string => {
    const [r1, g1, b1] = hexToRgb(hex);
    const [r2, g2, b2] = hexToRgb(baseHex);
    const mix = (a: number, b: number): number => Math.round(a * weight + b * (1 - weight));
    return `rgb(${mix(r1, r2)}, ${mix(g1, g2)}, ${mix(b1, b2)})`;
};

/**
 * Tüm widget'ların ortak kabuğu — ferah, aydınlık kurumsal kart görünümü:
 * hafif saydam (rgba) zemin + backdrop-filter: blur ile modern bir "buzlu
 * cam" hissi, yumuşak/derin gölge. Yükseklik bilinçli olarak "auto" — kart,
 * satırdaki diğer kartlara eşitlenmek yerine kendi içeriğine göre esner
 * (sıkışık/gereksiz boş alan oluşmaz). Üstte ince renkli bir aksan çizgisi
 * ve hover'da hafif yükselme efekti, aynı gri-beyaz tonlardaki kartların
 * tekdüze durmaması için eklendi.
 */
const WidgetCard: React.FunctionComponent<IWidgetCardProps> = (props) => {
    const { title, subtitle, iconName, headerAction, children, className, accentColor } = props;
    const theme = useTheme();
    const accent = accentColor ?? theme.palette.themePrimary;

    // Kullanıcı isteği: "3 tema da birbirinden farklı olacak — arkaplan,
    // widgetlar ve özellikle karşılama header'ı FULL birbirinden farklı renk
    // paletleri kullanacak." Ancak kartların KENDİSİ (zemin) daha önce
    // kasıtlı olarak sabit/açık bırakılmıştı (bkz. yukarıdaki not — tam koyu
    // widget teması denenip kullanıcı geri bildirimiyle geri alınmıştı) ve
    // kullanıcı bunu onayladı: kart zemini AÇIK kalsın, sadece "vurgular"
    // (üst şerit + ikon rozeti) zaman dilimine göre değişsin. Bu yüzden her
    // widget'ın KENDİ kimlik rengi (accent, örn. Duyurular turuncu, Yaklaşan
    // Etkinlikler mor) o anki zaman temasının süsleme rengiyle (headerVariant.
    // decorationColor) harmanlanıyor — widget'lar birbirinden hâlâ ayırt
    // edilebilir kalırken, günün o anki paleti de her karttan hissediliyor.
    const timeBucket = useTimeAwareTheme();
    const headerVariant = getTimeTheme(timeBucket).header;
    const flairAccent = mixHex(headerVariant.decorationColor, accent, 0.45);

    const styles = mergeStyleSets({
        card: {
            position: 'relative',
            // Nötr açık zemin (theme.palette.neutralLighterAlt) + kartın kendi
            // vurgu renginden hafif bir soluk (wash) — tekdüze beyaz yerine.
            // İKİ DURAK DA OPAK (mixHex) — backdrop-filter desteklenmeyen
            // tarayıcılarda dahi kart HER ZAMAN aydınlık/okunaklı kalır.
            background: `linear-gradient(165deg, ${mixHex(accent, theme.palette.neutralLighterAlt, 0.14)} 0%, ${theme.palette.neutralLighterAlt} 42%)`,
            backdropFilter: 'blur(14px)',
            WebkitBackdropFilter: 'blur(14px)',
            border: 'none',
            borderRadius: 22,
            // Nötr, slate tonlu "floating" gölge — WelcomeHeader'daki aynı
            // yumuşak/kurumsal gölge diliyle tutarlı. ÖNCEKİ HATA: burada
            // rgba(0,0,0,0.35)/0.45 gibi çok koyu, saf siyah bir gölge vardı —
            // "ferah/aydınlık" kart zemininin üzerinde her kartın etrafında
            // kirli/koyu bir hale oluşturuyordu (header'ın 0.08 alfa'sıyla
            // yan yana bariz tutarsızdı). Artık aynı slate (15,23,42) tonunda,
            // çok daha hafif bir alfa kullanılıyor.
            boxShadow: '0 8px 24px rgba(15,23,42,0.10)',
            padding: 24,
            display: 'flex',
            flexDirection: 'column',
            height: '100%',
            boxSizing: 'border-box',
            transition: 'all 0.3s ease',
            // transition zaten yukarıda 'all' — üst şerit/gölge rengi zaman
            // dilimi değiştiğinde de (flairAccent) yumuşak geçiş yapar.
            selectors: {
                '::before': {
                    content: '""',
                    position: 'absolute',
                    top: 0,
                    left: 14,
                    right: 14,
                    height: 4,
                    borderRadius: '0 0 4px 4px',
                    background: `linear-gradient(90deg, ${flairAccent} 0%, ${hexToRgba(flairAccent, 0.5)} 100%)`,
                    transition: 'background 1.5s ease'
                },
                ':hover': {
                    transform: 'translateY(-6px)',
                    boxShadow: `0 16px 34px rgba(15,23,42,0.16), 0 0 0 1px ${hexToRgba(flairAccent, 0.25)}`
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
            background: `linear-gradient(135deg, ${flairAccent} 0%, ${hexToRgba(flairAccent, 0.72)} 100%)`,
            boxShadow: `0 4px 10px ${hexToRgba(flairAccent, 0.35)}`,
            flexShrink: 0,
            transition: 'background 1.5s ease, box-shadow 1.5s ease',
            // titleGroup içindeki bir sonraki çocuktan (isim/unvan bloğu) önce
            // gelen boşluk — "gap" yerine.
            marginRight: 10
        },
        icon: {
            fontSize: 16,
            // Literal beyaz — bu ikon her zaman renkli bir rozetin üzerinde
            // duruyor, tema aydınlık/karanlık olsa da rozet zemini hep canlı
            // bir vurgu rengi olduğu için ikon hep beyaz kalmalı.
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
