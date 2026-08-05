import * as React from 'react';
import { Icon, Persona, PersonaSize, mergeStyleSets, keyframes } from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { useTimeAwareTheme, getTimeTheme } from '../../themeManager';
import NotificationBell from './NotificationBell';

export interface IWelcomeHeaderProps {
    userDisplayName: string;
    context: WebPartContext;
}

const getFirstName = (displayName: string): string => {
    if (!displayName) {
        return '';
    }
    return displayName.split(' ')[0];
};

// "Mesh Gradient" — o anki zaman dilimine ait 5 durak arasında yavaşça akan,
// çok hafif/dağınık bir arka plan. Renkler doğrudan yazılmıyor, CSS
// değişkenlerinden (--yorpas-theme-mesh-*) okunuyor — ThemeManager bunları
// güncelledikçe gradyan otomatik yeni paletle boyanır. Bu katman, üstündeki
// yarı şeffaf beyaz cam katmanının (glassLayer) ALTINDA kalır — bu yüzden
// kaynak renkler pastel olsa da sonuçta iyice soluklaşmış/zarif bir "enerji"
// hissi olarak süzülür.
const meshFlow = keyframes({
    '0%': { backgroundPosition: '0% 50%' },
    '50%': { backgroundPosition: '100% 50%' },
    '100%': { backgroundPosition: '0% 50%' }
});

// Camın üzerinden yavaşça geçen ince ışık huzmesi — bir cam yüzeyde ışığın
// kırılması gibi, dikkat çekici ama rahatsız etmeyen bir parıltı.
const shimmerSweep = keyframes({
    '0%': { transform: 'translateX(-60%) rotate(10deg)' },
    '100%': { transform: 'translateX(160%) rotate(10deg)' }
});

// "Hoş geldin" yanındaki 👋 emojisi hafifçe yukarıda ve eğik duruyor
// (bkz. waveEmoji'deki statik transform); bu animasyon o eğik duruştan
// sürekli bir el sallama hareketi üretir.
const wave = keyframes({
    '0%, 100%': { transform: 'translateY(-6px) rotate(-8deg)' },
    '50%': { transform: 'translateY(-6px) rotate(14deg)' }
});

// Avatarın etrafındaki yeşil "Status Ring" için yumuşak bir nabız/pulse —
// çevrimiçi olduğunu belli belirsiz ama sürekli hatırlatan bir detay.
const avatarGlow = keyframes({
    '0%, 100%': { boxShadow: '0 0 0 0 rgba(52,199,89,0.35), 0 4px 14px rgba(15,23,42,0.10)' },
    '50%': { boxShadow: '0 0 0 6px rgba(52,199,89,0.12), 0 4px 14px rgba(15,23,42,0.10)' }
});

// İsmin altında sayfa yüklendiğinde soldan sağa "çizilen" ince aksan çizgisi
// — yıldız/ışıltı süslemelerinin (kullanıcı geri bildirimiyle kaldırıldı)
// yerini alan, daha sade/kurumsal ama hâlâ dikkat çekici bir "reveal" detayı.
// Bir kerelik oynar (bkz. greetingName'deki animationIterationCount: 1).
const growUnderline = keyframes({
    '0%': { transform: 'scaleX(0)', opacity: 0 },
    '100%': { transform: 'scaleX(1)', opacity: 1 }
});

// Avatarın arkasında yavaşça büyüyüp küçülen, zaman dilimi rengindeki bulanık
// "halo" — karşılama header'ına daha dikkat çekici, "spotlight" hissi veren
// odak noktası. Işıltılarla aynı ailede ama daha büyük/yumuşak bir hareket.
const haloPulse = keyframes({
    '0%, 100%': { transform: 'translate(-50%, -50%) scale(1)', opacity: 0.55 },
    '50%': { transform: 'translate(-50%, -50%) scale(1.2)', opacity: 0.85 }
});

// Header'ın dış kenarında zaman dilimi rengiyle yavaşça nefes alan bir
// ışıma — sabit gölgeye ek, ikinci (ayrı) bir animasyon katmanı olarak.
const edgeGlow = keyframes({
    '0%, 100%': { boxShadow: '0 10px 40px -10px rgba(0,0,0,0.08), 0 0 0px 0px var(--yorpas-theme-header-decoration, transparent)' },
    '50%': { boxShadow: '0 10px 40px -10px rgba(0,0,0,0.08), 0 0 30px 2px var(--yorpas-theme-header-decoration, transparent)' }
});

// Sayfa yüklendiğinde header'ın yavaşça belirmesi (fade-in-up) — bir kerelik,
// sonsuz döngülü DEĞİL (bkz. root'taki animationIterationCount: 1).
const fadeInUp = keyframes({
    '0%': { opacity: 0, transform: 'translateY(16px)' },
    '100%': { opacity: 1, transform: 'translateY(0)' }
});

const getTodayLabel = (): string => {
    const label = new Date().toLocaleDateString('tr-TR', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });
    return label.charAt(0).toLocaleUpperCase('tr-TR') + label.slice(1);
};

/** Saat:dakika ve saniye AYRI döner — saniye, saatten daha küçük/soluk bir "widget" detayı olarak ayrı stillenebilsin diye. */
const getClockParts = (): { main: string; seconds: string } => {
    const now = new Date();
    const main = now.toLocaleTimeString('tr-TR', { hour: '2-digit', minute: '2-digit' });
    const seconds = now.toLocaleTimeString('tr-TR', { second: '2-digit' });
    return { main, seconds };
};

interface IHeaderClockProps {
    clockRowClassName: string;
    clockMainClassName: string;
    clockSecondsClassName: string;
}

/**
 * Saat, saniyede bir kendi state'ini güncellemesi gereken TEK parça —
 * bu yüzden ayrı bir bileşene çıkarıldı. Önceden bu tik WelcomeHeader'ın
 * kendi state'indeydi; bu da mergeStyleSets ile üretilen ~30 kuralı (7
 * CSS keyframe animasyonu dahil) SANİYEDE BİR yeniden hesaplatıyor ve
 * arkadaki mesh gradyanı animasyonunun sürekli kesilip yeniden
 * başlamasına ("kasma") yol açıyordu. Artık sadece bu küçük alt bileşen
 * saniyede bir render oluyor, header'ın geri kalanı (ve ağır stil/
 * animasyon katmanları) saat tikinden tamamen etkilenmiyor.
 */
const HeaderClock: React.FunctionComponent<IHeaderClockProps> = (props) => {
    const [clock, setClock] = React.useState(getClockParts());

    React.useEffect(() => {
        const interval = window.setInterval(() => setClock(getClockParts()), 1000);
        return () => window.clearInterval(interval);
    }, []);

    return (
        <div className={props.clockRowClassName}>
            <span className={props.clockMainClassName}>{clock.main}</span>
            <span className={props.clockSecondsClassName}>:{clock.seconds}</span>
        </div>
    );
};

/** Saate göre "Günaydın / İyi günler / İyi akşamlar" — motivasyon sözü değil, sade bir kişiselleştirme. */
const getGreetingPrefix = (bucket: string): string => {
    if (bucket === 'morning') { return 'Günaydın'; }
    if (bucket === 'day') { return 'İyi günler'; }
    return 'İyi akşamlar';
};

/**
 * Sayfanın en üstünde tam genişlikte "yüzen" karşılama şeridi — Premium
 * Kurumsal / Light Glass tasarımı: sabit yarı şeffaf beyaz bir cam paneli
 * (glassLayer) + arkasında zaman dilimine göre değişen, çok hafif/dağınık
 * bir mesh gradyanı (meshLayer, bkz. themeManager.ts). `useTimeAwareTheme()`
 * sayfa yüklendiğinde ve her dakika o anki zaman dilimini kontrol edip
 * :root üzerindeki --yorpas-theme-* değişkenlerini günceller; bu bileşen o
 * değişkenleri `var(...)` ile OKUR. Zaman dilimi değiştiğinde geçişler
 * `transition: background 1.5s ease` ile yumuşak olur.
 */
const WelcomeHeader: React.FunctionComponent<IWelcomeHeaderProps> = (props) => {
    const { userDisplayName, context } = props;
    const firstName = getFirstName(userDisplayName);
    const timeBucket = useTimeAwareTheme();
    const headerVariant = getTimeTheme(timeBucket).header;

    // Kullanıcının kendi profil fotoğrafı — Graph izni/onay beklemeye gerek
    // kalmadan, her SharePoint sitesinde hazır bulunan userphoto.aspx sistem
    // sayfasından (aynı origin, ekstra izin gerektirmez) çekiliyor. Persona
    // bileşeni fotoğraf yüklenemezse (404) otomatik olarak baş harfli rozete
    // düşer — ayrıca bir onError yönetimine gerek yok.
    const loginName = context.pageContext.user.loginName;
    const photoUrl = loginName
        ? `${context.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?size=L&accountname=${encodeURIComponent(loginName)}`
        : undefined;

    // headerVariant (dolayısıyla timeBucket) DIŞINDA hiçbir şeye bağlı değil —
    // useMemo olmadan bu obje HER render'da (ör. üst bileşenin re-render'ında)
    // yeniden kurulur ve mergeStyleSets'in ürettiği class adları değişebilir,
    // bu da üzerlerindeki CSS animasyonlarının sıfırdan yeniden başlamasına
    // (görsel "kasma") yol açar. Artık sadece dakikada bir zaman dilimi
    // değiştiğinde yeniden hesaplanıyor.
    const styles = React.useMemo(() => mergeStyleSets({
        root: {
            position: 'relative',
            overflow: 'hidden',
            width: '100%',
            borderRadius: 20,
            // Sabit, nötr "floating" gölge — koyu temadaki renkli ışıma yerine
            // açık/kurumsal bir zeminde süzülüyormuş hissi veren yumuşak gölge.
            boxShadow: '0 10px 40px -10px rgba(0,0,0,0.08)',
            minHeight: 128,
            boxSizing: 'border-box',
            // İki AYRI animasyon aynı anda: fadeInUp bir kerelik giriş,
            // edgeGlow ise header açık kaldığı sürece sonsuz döngüde nefes
            // alan zaman-dilimi renginde bir dış ışıma.
            animationName: `${fadeInUp}, ${edgeGlow}`,
            animationDuration: '0.6s, 4.5s',
            animationTimingFunction: 'ease-out, ease-in-out',
            animationIterationCount: '1, infinite',
            animationFillMode: 'both, none'
        },
        // Katman 1: çok hafif/dağınık mesh gradyanı — zaman dilimine göre değişir.
        meshLayer: {
            position: 'absolute',
            inset: 0,
            background:
                'linear-gradient(45deg, var(--yorpas-theme-mesh-1, #FFD9A6), var(--yorpas-theme-mesh-2, #FFF4E3), ' +
                'var(--yorpas-theme-mesh-3, #DCF1FF), var(--yorpas-theme-mesh-4, #A9D9F5), ' +
                'var(--yorpas-theme-mesh-5, #FFE2C2))',
            backgroundSize: '300% 300%',
            animationName: meshFlow,
            animationDuration: '16s',
            animationTimingFunction: 'ease',
            animationIterationCount: 'infinite',
            transition: 'background 1.5s ease'
        },
        // Katman 1b: camın üzerinden yavaşça geçen ince ışık huzmesi. ZIndex
        // BİLİNÇLİ OLARAK glassLayer'ın (2) ÜSTÜNDE (3) — önceki sürümde bu
        // katman zIndex:1 idi, yani glassLayer'ın (zIndex:2) ALTINDA kalıyordu
        // ve buzlu/opak cam panelinin arkasında neredeyse tamamen görünmez
        // oluyordu; "camın üzerinden geçen ışık huzmesi" tarif edildiği hâlde
        // fiilen hiç görünmüyordu. content ile aynı seviyede olsa da DOM
        // sırası content'ten önce geldiği için metnin altında/gerisinde kalır.
        shimmerLayer: {
            position: 'absolute',
            top: '-60%',
            left: 0,
            width: '35%',
            height: '220%',
            background: 'linear-gradient(90deg, rgba(255,255,255,0) 0%, rgba(255,255,255,0.5) 50%, rgba(255,255,255,0) 100%)',
            animationName: shimmerSweep,
            animationDuration: '7s',
            animationTimingFunction: 'ease-in-out',
            animationIterationCount: 'infinite',
            pointerEvents: 'none',
            zIndex: 3
        },
        // Katman 2: Glassmorphism camı — ÖNCEKİ HATA: bu her zaman sabit beyaz
        // tutuluyordu, yani üç zaman dilimi de aslında aynı "beyaz cam
        // üzerinde farklı mesh" görünümüne sahipti — sabah/öğlen/akşam
        // "tamamen farklı" hissettirmiyordu. Artık camın kendisi de zaman
        // dilimine göre değişiyor (sabah beyaz/berrak, öğlen krem/"relax",
        // akşam GERÇEKTEN koyu lacivert) — bkz. themeManager.ts → header.
        // Yüksek opaklık (0.72-0.82) BİLİNÇLİ: backdrop-filter'ı desteklemeyen
        // tarayıcılarda (bu portalda doğrulandı — flex "gap" ile aynı sınıf
        // sorun) dahi bulanıklık OLMADAN, sadece iki opak/yarı-opak rengin
        // düz alfa karışımıyla bile doğru/okunaklı görünür; blur sadece bir
        // "bonus" ince detay, doğruluk ona bağlı değil.
        glassLayer: {
            position: 'absolute',
            inset: 0,
            zIndex: 2,
            background: 'var(--yorpas-theme-header-glass-bg, rgba(255, 255, 255, 0.72))',
            backdropFilter: 'blur(24px)',
            WebkitBackdropFilter: 'blur(24px)',
            border: '1px solid var(--yorpas-theme-header-glass-border, rgba(255, 255, 255, 0.4))',
            borderRadius: 20,
            boxSizing: 'border-box',
            transition: 'background 1.5s ease, border-color 1.5s ease'
        },
        // NOT: "gap" flex özelliği burada BİLİNÇLİ OLARAK KULLANILMIYOR — bu
        // sayfanın render edildiği (kurumsal/eski) tarayıcı ortamında flex "gap"
        // desteklenmiyor. `flexWrap: 'wrap'` olduğu için mainRow'a hem
        // marginRight hem marginBottom veriliyor (dar ekranda alt alta sarınca).
        content: {
            position: 'relative',
            zIndex: 3,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            flexWrap: 'wrap',
            padding: '28px 40px',
            boxSizing: 'border-box',
            minHeight: 128,
            selectors: {
                '@media (max-width: 520px)': {
                    justifyContent: 'flex-start',
                    padding: '20px 24px'
                }
            }
        },
        mainRow: {
            display: 'flex',
            alignItems: 'center',
            // content'in (üstteki) headerAction'a benzer bir sonraki kardeşi
            // (dateTimeBlock) olduğu için "gap" yerine burada kendi payına
            // düşen marginRight/marginBottom taşınıyor.
            marginRight: 16,
            marginBottom: 16
        },
        // Avatarın etrafındaki "Status Ring" — çevrimiçi olduğunu gösteren
        // zarif, yeşil tonlu bir halka (avatarGlow ile yavaşça nabız atar).
        personaRing: {
            position: 'relative',
            borderRadius: '50%',
            padding: 3,
            background: 'rgba(52,199,89,0.16)',
            border: '2px solid rgba(52,199,89,0.55)',
            transition: 'transform 0.2s ease',
            cursor: 'default',
            // mainRow'daki "gap" yerine — textBlock'tan önceki boşluk.
            marginRight: 18,
            animationName: avatarGlow,
            animationDuration: '2.8s',
            animationTimingFunction: 'ease-in-out',
            animationIterationCount: 'infinite',
            selectors: {
                ':hover': {
                    transform: 'translateY(-4px)'
                }
            }
        },
        // Avatarın ARKASINDA, zaman dilimi rengiyle yavaşça büyüyüp küçülen
        // bulanık bir "halo" — header'a daha dikkat çekici bir odak noktası
        // katar. avatarContent'in (Persona + rozet) GERİSİNDE kalması için
        // zIndex 0, avatarContent ise açıkça zIndex 1.
        avatarHalo: {
            position: 'absolute',
            top: '50%',
            left: '50%',
            width: 122,
            height: 122,
            borderRadius: '50%',
            background: 'radial-gradient(circle, var(--yorpas-theme-header-decoration, #F0A868) 0%, transparent 68%)',
            filter: 'blur(10px)',
            zIndex: 0,
            pointerEvents: 'none',
            transition: 'background 1.5s ease',
            animationName: haloPulse,
            animationDuration: '3.6s',
            animationTimingFunction: 'ease-in-out',
            animationIterationCount: 'infinite'
        },
        avatarContent: {
            position: 'relative',
            zIndex: 1
        },
        // Avatarın sağ altındaki küçük "çevrimiçi" rozeti — Status Ring'e ek,
        // ikinci bir görsel doğrulama (yaygın bir UI konvansiyonu).
        presenceDot: {
            position: 'absolute',
            right: 1,
            bottom: 1,
            width: 15,
            height: 15,
            borderRadius: '50%',
            background: '#2ecc71',
            border: '2.5px solid rgba(255,255,255,0.95)',
            boxShadow: '0 0 6px rgba(46,204,113,0.6)',
            zIndex: 1
        },
        // Header'ın alt kenarındaki ince degrade çizgi — o anki zaman dilimi
        // paletinden (mesh-1/3/5) beslenir, WidgetCard'ların üst aksan
        // çizgisiyle aynı görsel dili header'a da taşır.
        bottomAccent: {
            position: 'absolute',
            left: 14,
            right: 14,
            bottom: 3,
            height: 3,
            borderRadius: 2,
            background:
                'linear-gradient(90deg, var(--yorpas-theme-mesh-1, #FFD9A6) 0%, ' +
                'var(--yorpas-theme-mesh-3, #DCF1FF) 50%, var(--yorpas-theme-mesh-5, #FFE2C2) 100%)',
            opacity: 0.9,
            zIndex: 2,
            pointerEvents: 'none',
            transition: 'background 1.5s ease'
        },
        textBlock: {},
        // "GÜNAYDIN" — ince, tamamen büyük harf, geniş harf aralığı, soluk slate.
        // NOT: "gap" burada da kullanılmıyor (flex "gap" desteklenmiyor) — ikon
        // için ayrı bir sınıf (greetingPrefixIcon) marginRight taşıyor.
        greetingPrefix: {
            display: 'flex',
            alignItems: 'center',
            fontSize: 12,
            fontWeight: 500,
            color: 'var(--yorpas-theme-header-muted, #64748B)',
            textTransform: 'uppercase',
            letterSpacing: 2,
            margin: 0,
            transition: 'color 1.5s ease'
        },
        greetingPrefixIcon: {
            marginRight: 6
        },
        // İsim satırı: "Hoş geldin," daha yumuşak, isim ise kalın/keskin/koyu slate.
        // NOT: "gap" burada kullanılmıyor — ilk çocuk düz bir metin (text node)
        // olduğu için ona margin verilemiyor; bunun yerine SONRAKİ elemanlara
        // (greetingName, waveEmoji) marginLeft veriliyor — aynı görsel sonucu üretir.
        greeting: {
            display: 'flex',
            alignItems: 'center',
            fontFamily: "'Segoe UI', -apple-system, sans-serif",
            fontSize: 32,
            fontWeight: 400,
            color: 'var(--yorpas-theme-text, #334155)',
            margin: '6px 0 0',
            letterSpacing: 0.1
        },
        greetingName: {
            position: 'relative',
            display: 'inline-block',
            fontWeight: 800,
            color: 'var(--yorpas-theme-header-name, #0F172A)',
            marginLeft: 8,
            transition: 'color 1.5s ease',
            selectors: {
                '::after': {
                    content: '""',
                    position: 'absolute',
                    left: 0,
                    right: 0,
                    bottom: -4,
                    height: 3,
                    borderRadius: 2,
                    background: 'var(--yorpas-theme-header-decoration, #0078D4)',
                    transformOrigin: 'left center',
                    animationName: growUnderline,
                    animationDuration: '0.7s',
                    animationDelay: '0.5s',
                    animationTimingFunction: 'ease-out',
                    animationFillMode: 'both',
                    animationIterationCount: 1
                }
            }
        },
        // Eğik ve hafifçe yukarıda duran el sallama emojisi (statik konum bu
        // transform'da, sallanma hareketi wave animasyonundan gelir). marginLeft
        // = eski "gap" (8) + kendi ince konum düzeltmesi (2).
        waveEmoji: {
            display: 'inline-block',
            fontSize: 24,
            marginLeft: 10,
            transformOrigin: '70% 70%',
            animationName: wave,
            animationDuration: '1.6s',
            animationTimingFunction: 'ease-in-out',
            animationIterationCount: 'infinite'
        },
        // Sağ taraf: saat kendi başına bir "widget" gibi — büyük/ince ana
        // kısım + küçük/soluk saniye, tarih altında ince bir hizada.
        dateTimeBlock: {
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'flex-end',
            selectors: {
                '@media (max-width: 520px)': {
                    alignItems: 'flex-start'
                }
            }
        },
        // NOT: "gap" burada kullanılmıyor — clockMain'e doğrudan marginRight verildi.
        clockRow: {
            display: 'flex',
            alignItems: 'baseline'
        },
        clockMain: {
            fontFamily: "'Segoe UI Light', 'Segoe UI', sans-serif",
            fontSize: 36,
            fontWeight: 200,
            color: 'var(--yorpas-theme-header-clock, #1E293B)',
            fontVariantNumeric: 'tabular-nums',
            letterSpacing: 0.5,
            lineHeight: '40px',
            marginRight: 2,
            transition: 'color 1.5s ease'
        },
        clockSeconds: {
            fontFamily: "'Segoe UI Light', 'Segoe UI', sans-serif",
            fontSize: 16,
            fontWeight: 300,
            color: 'var(--yorpas-theme-header-clock-seconds, rgba(30,41,59,0.45))',
            fontVariantNumeric: 'tabular-nums',
            lineHeight: '20px',
            transition: 'color 1.5s ease'
        },
        // NOT: "gap" burada kullanılmıyor — Icon'a ayrı bir sınıf (dateTextIcon)
        // üzerinden marginRight veriliyor.
        dateText: {
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'flex-end',
            fontSize: 13,
            fontWeight: 500,
            color: 'var(--yorpas-theme-header-date, rgba(51,65,85,0.72))',
            lineHeight: '18px',
            marginTop: 3,
            transition: 'color 1.5s ease'
        },
        dateTextIcon: {
            marginRight: 6
        },
        // Zil + saat/tarih bloğu yan yana — NOT: "gap" burada da kullanılmıyor,
        // zil kendi marginRight'ını taşıyor.
        topRightRow: {
            display: 'flex',
            alignItems: 'flex-start'
        },
        notificationBellWrap: {
            marginRight: 16,
            marginTop: 2
        }
    }), [headerVariant]);

    return (
        <div className={styles.root}>
            <span className={styles.meshLayer} />
            <span className={styles.shimmerLayer} />
            <span className={styles.glassLayer} />
            <span className={styles.bottomAccent} aria-hidden="true" />
            <div className={styles.content}>
                <div className={styles.mainRow}>
                    <div className={styles.personaRing}>
                        <span className={styles.avatarHalo} aria-hidden="true" />
                        <div className={styles.avatarContent}>
                            <Persona
                                imageUrl={photoUrl}
                                text={userDisplayName}
                                size={PersonaSize.size72}
                                hidePersonaDetails
                                imageAlt={userDisplayName}
                            />
                            <span className={styles.presenceDot} aria-hidden="true" />
                        </div>
                    </div>
                    <div className={styles.textBlock}>
                        <p className={styles.greetingPrefix}><Icon iconName={headerVariant.greetingIcon} className={styles.greetingPrefixIcon} /> {getGreetingPrefix(timeBucket)}</p>
                        <p className={styles.greeting}>
                            Hoş geldin, <span className={styles.greetingName}>{firstName || 'Kullanıcı'}</span>
                            <span className={styles.waveEmoji} aria-hidden="true">👋</span>
                        </p>
                    </div>
                </div>
                <div className={styles.topRightRow}>
                    <div className={styles.notificationBellWrap}>
                        <NotificationBell context={context} />
                    </div>
                    <div className={styles.dateTimeBlock}>
                        <HeaderClock
                            clockRowClassName={styles.clockRow}
                            clockMainClassName={styles.clockMain}
                            clockSecondsClassName={styles.clockSeconds}
                        />
                        <div className={styles.dateText}>
                            <Icon iconName="Calendar" className={styles.dateTextIcon} />
                            {getTodayLabel()}
                        </div>
                    </div>
                </div>
            </div>
        </div>
    );
};

export default WelcomeHeader;
