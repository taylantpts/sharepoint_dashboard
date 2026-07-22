import * as React from 'react';
import { DefaultButton, IconButton, Icon, TooltipHost, useTheme, mergeStyleSets } from '@fluentui/react';
import WidgetCard from '../WidgetCard';

export interface ILocationCardProps {
    label: string;
    address: string;
    iconName: string;
    accentColor: string;
    /**
     * Haritada TEK bir nokta işaretlensin diye kullanılan, sadeleştirilmiş
     * arama sorgusu. Kat/ofis no gibi ayrıntılar Google'ın bina içindeki
     * BİRDEN FAZLA işletmeyi (plaza vb.) eşleştirip birden çok nokta/sponsorlu
     * sonuç göstermesine yol açabiliyor — verilmezse tam `address` kullanılır.
     * `coordinates` verilmişse bu alan yalnızca "Haritada Gör" linkinde yedek
     * olarak kullanılır.
     */
    mapQuery?: string;
    /**
     * Bina/plaza gibi metin tabanlı adreslerde Google'ın civardaki İŞLETME
     * sonuçlarını (sponsorlu dahil) göstermesini TAMAMEN engellemenin tek
     * garantili yolu: metin araması yerine doğrudan enlem/boylam ile tek bir
     * ham nokta işaretlemek. Verilirse mapQuery/address yerine bu kullanılır.
     */
    coordinates?: { lat: number; lon: number };
}

const getGoogleMapsUrl = (address: string, coordinates?: { lat: number; lon: number }): string =>
    coordinates
        ? `https://www.google.com/maps/search/?api=1&query=${coordinates.lat},${coordinates.lon}`
        : `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(address)}`;

// API anahtarı gerektirmeyen genel Google Maps gömme (embed) biçimi. Koordinat
// verilmişse ham enlem/boylamla ("q=lat,lon") tek bir nokta işaretlenir —
// metin aramasının aksine civardaki işletme/sponsorlu sonuçları GÖSTERMEZ.
const getGoogleMapsEmbedUrl = (query: string, coordinates?: { lat: number; lon: number }): string =>
    coordinates
        ? `https://www.google.com/maps?q=${coordinates.lat},${coordinates.lon}&z=17&output=embed`
        : `https://www.google.com/maps?q=${encodeURIComponent(query)}&output=embed`;

const openInGoogleMaps = (address: string, coordinates?: { lat: number; lon: number }): void => {
    window.open(getGoogleMapsUrl(address, coordinates), '_blank', 'noopener,noreferrer');
};

/**
 * Tek bir lokasyonu (Merkez Ofis, Fabrika vb.) gösteren bağımsız kart —
 * Dashboard'da "Gerekli Dosyalar" ile aynı satırda, ayrı sütunlarda
 * kullanılabilmesi için her lokasyon kendi WidgetCard'ına sahip.
 */
const LocationCard: React.FunctionComponent<ILocationCardProps> = (props) => {
    const { label, address, iconName, accentColor, mapQuery, coordinates } = props;
    const theme = useTheme();
    const query = mapQuery ?? address;
    const [copied, setCopied] = React.useState(false);

    const handleCopyAddress = (): void => {
        if (!navigator.clipboard) {
            return;
        }
        navigator.clipboard
            .writeText(address)
            .then(() => {
                setCopied(true);
                window.setTimeout(() => setCopied(false), 1800);
            })
            .catch(() => { /* pano erişimi engellenmiş olabilir, sessizce yok say */ });
    };

    const styles = mergeStyleSets({
        // Merkez Ofis'in adresi Fabrika'nınkinden çok daha uzun (daha fazla
        // satıra sarıyor) — bu yüzden içerik sabit bir dikey akış yerine
        // tam yükseklikte bir flex-column: adres+harita üstte, "Haritada
        // Gör" butonu ise `spacer` (flexGrow:1) sayesinde HER ZAMAN kartın
        // altına yapışık kalır. İki lokasyon kartı da grid'de eşit
        // yüksekliğe stretch edildiği için (bkz. Dashboard.module.scss
        // align-items:stretch), bu sayede iki buton da aynı hizada durur —
        // adres uzunluğu farkı artık buton konumunu KAYDIRMIYOR.
        content: {
            display: 'flex',
            flexDirection: 'column',
            height: '100%'
        },
        spacer: {
            flexGrow: 1,
            minHeight: 12
        },
        // Adres, düz bir metin satırı yerine kendi çerçeveli "kutusunda"
        // duruyor — hem daha OKUNAKLI (satırlar birbirine yapışmıyor) hem de
        // yanına kopyala butonu eklenebiliyor.
        //
        // SABİT minHeight KASITLI: Merkez Ofis'in adresi (~3 satıra sarıyor)
        // Fabrika'nınkinden (~1-2 satır) belirgin şekilde uzun. minHeight
        // olmasaydı adres kutusu her kartta FARKLI yükseklikte biterdi ve
        // altındaki HARİTA (iframe) iki kartta da FARKLI bir Y konumunda
        // başlardı — sadece "Haritada Gör" butonu değil, haritanın kendisi
        // de hizasız görünürdü. Değer, mevcut Merkez Ofis adresinin bu kart
        // genişliğinde kapladığı ~3 satıra göre ayarlandı; adresler önemli
        // ölçüde uzarsa bu değerin yeniden ayarlanması gerekebilir.
        // NOT: "gap" flex özelliği burada BİLİNÇLİ OLARAK KULLANILMIYOR — bu
        // sayfanın render edildiği (kurumsal/eski) tarayıcı ortamında flex "gap"
        // desteklenmiyor. Bunun yerine pinIcon/addressText kendi marginRight'ını
        // taşıyor (bkz. aşağıda), son çocuk (kopyala butonu) dokunulmadan kalıyor.
        addressRow: {
            display: 'flex',
            alignItems: 'flex-start',
            minHeight: 84,
            boxSizing: 'border-box',
            background: theme.palette.neutralLighterAlt,
            borderRadius: 12,
            boxShadow: '0 2px 8px rgba(0,0,0,0.04)',
            padding: '10px 10px 10px 12px',
            marginBottom: 14
        },
        pinIcon: {
            fontSize: 13,
            color: accentColor,
            marginTop: 3,
            flexShrink: 0,
            marginRight: 8
        },
        addressText: {
            flexGrow: 1,
            minWidth: 0,
            fontSize: 12,
            lineHeight: '19px',
            color: theme.semanticColors.bodyText,
            wordBreak: 'break-word',
            whiteSpace: 'normal',
            marginRight: 8
        },
        copyButton: {
            flexShrink: 0,
            marginTop: -2
        },
        mapFrame: {
            width: '100%',
            height: 170,
            border: 'none',
            display: 'block',
            borderRadius: 8
        },
        button: {
            width: '100%'
        }
    });

    return (
        <WidgetCard title={label} subtitle="Adres ve konum" iconName={iconName} accentColor={accentColor}>
            <div className={styles.content}>
                <div className={styles.addressRow}>
                    <Icon iconName="MapPin" className={styles.pinIcon} />
                    <div className={styles.addressText}>{address}</div>
                    <TooltipHost content={copied ? 'Kopyalandı!' : 'Adresi kopyala'}>
                        <IconButton
                            iconProps={{ iconName: copied ? 'CheckMark' : 'Copy' }}
                            className={styles.copyButton}
                            onClick={handleCopyAddress}
                            ariaLabel="Adresi kopyala"
                        />
                    </TooltipHost>
                </div>
                <iframe
                    className={styles.mapFrame}
                    title={`${label} - Google Haritalar`}
                    src={getGoogleMapsEmbedUrl(query, coordinates)}
                    loading="lazy"
                    referrerPolicy="no-referrer-when-downgrade"
                />
                <div className={styles.spacer} />
                <DefaultButton
                    iconProps={{ iconName: 'OpenInNewWindow' }}
                    text="Haritada Gör"
                    onClick={() => openInGoogleMaps(query, coordinates)}
                    className={styles.button}
                />
            </div>
        </WidgetCard>
    );
};

export default LocationCard;
