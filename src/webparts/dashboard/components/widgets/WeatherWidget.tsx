import * as React from 'react';
import { Icon, Spinner, SpinnerSize, MessageBar, MessageBarType, useTheme, mergeStyleSets } from '@fluentui/react';
import WidgetCard from '../WidgetCard';
import AnimatedWeatherIcon, { WeatherCondition } from './AnimatedWeatherIcon';
import { DATA_UNAVAILABLE_MESSAGE } from '../../constants';

interface IWeatherStat {
    iconName: string;
    label: string;
    value: string;
}

interface IForecastDay {
    dayLabel: string;
    high: number;
    low: number;
    variant: WeatherCondition;
}

interface IHourlyForecast {
    timeLabel: string;
    tempC: number;
    variant: WeatherCondition;
}

interface IWeatherData {
    city: string;
    tempC: number;
    feelsLikeC: number;
    condition: string;
    variant: WeatherCondition;
    humidity: number;
    windKmh: number;
    pressureHpa: number;
    forecast: IForecastDay[];
    hourly: IHourlyForecast[];
}

interface IOpenWeatherResponse {
    weather?: Array<{ description: string; icon: string }>;
    main?: { temp: number; feels_like: number; humidity: number; pressure: number };
    wind?: { speed: number };
    name?: string;
    cod?: number | string;
    message?: string;
}

interface IOpenWeatherForecastEntry {
    dt: number;
    dt_txt: string;
    main: { temp: number };
    weather: Array<{ icon: string }>;
}

interface IOpenWeatherForecastResponse {
    list?: IOpenWeatherForecastEntry[];
}

const TURKISH_DAY_ABBR = ['Paz', 'Pzt', 'Sal', 'Çar', 'Per', 'Cum', 'Cmt'];

// OpenWeatherMap ikon kodları: 01 açık, 02/03 az/parçalı bulutlu, 04 kapalı,
// 09/10 yağmurlu, 11 gök gürültülü fırtına, 13 kar, 50 sis (bulutlu ile
// gösteriliyor). https://openweathermap.org/weather-conditions
const mapOwmIconToCondition = (icon: string): WeatherCondition => {
    const code = icon.substring(0, 2);
    if (code === '01') {
        return 'sunny';
    }
    if (code === '02' || code === '03') {
        return 'partlyCloudy';
    }
    if (code === '04' || code === '50') {
        return 'cloudy';
    }
    if (code === '09' || code === '10') {
        return 'rainy';
    }
    if (code === '11') {
        return 'stormy';
    }
    if (code === '13') {
        return 'snowy';
    }
    return 'cloudy';
};

/**
 * OpenWeatherMap'in ücretsiz "5 day / 3 hour" tahmin uç noktasından gelen
 * 3'er saatlik kayıtları GÜNE göre gruplar; her gün için en yüksek/en düşük
 * sıcaklığı ve öğlene (12:00) en yakın kaydın ikonunu temsilci olarak alır.
 * Bugünün kalan saatleri hariç tutulur, sonraki 5 gün döner.
 */
const buildDailyForecast = (entries: IOpenWeatherForecastEntry[]): IForecastDay[] => {
    const todayKey = new Date().toISOString().substring(0, 10);
    const byDay = new Map<string, IOpenWeatherForecastEntry[]>();

    entries.forEach((entry) => {
        const dayKey = entry.dt_txt.substring(0, 10);
        if (dayKey === todayKey) {
            return;
        }
        const existing = byDay.get(dayKey);
        if (existing) {
            existing.push(entry);
        } else {
            byDay.set(dayKey, [entry]);
        }
    });

    return Array.from(byDay.entries())
        .slice(0, 5)
        .map(([dayKey, dayEntries]) => {
            const temps = dayEntries.map((e) => e.main.temp);
            const middayEntry = dayEntries.reduce((closest, entry) => {
                const hour = parseInt(entry.dt_txt.substring(11, 13), 10);
                const closestHour = parseInt(closest.dt_txt.substring(11, 13), 10);
                return Math.abs(hour - 12) < Math.abs(closestHour - 12) ? entry : closest;
            }, dayEntries[0]);

            const dayIndex = new Date(`${dayKey}T00:00:00`).getDay();

            return {
                dayLabel: TURKISH_DAY_ABBR[dayIndex],
                high: Math.round(Math.max(...temps)),
                low: Math.round(Math.min(...temps)),
                variant: mapOwmIconToCondition(middayEntry.weather[0]?.icon ?? '01d')
            };
        });
};

/**
 * Bugünün KALAN saatleri — buildDailyForecast bunları bilerek dışlıyor
 * (günlük özet bugünü hiç göstermiyordu). Aynı ham "5 day / 3 hour" liste
 * üzerinden, sadece bugüne ait kayıtları saat etiketiyle birlikte döner —
 * ekstra bir API isteği gerekmez.
 */
const buildTodayHourly = (entries: IOpenWeatherForecastEntry[]): IHourlyForecast[] => {
    const todayKey = new Date().toISOString().substring(0, 10);
    return entries
        .filter((entry) => entry.dt_txt.substring(0, 10) === todayKey)
        .slice(0, 6)
        .map((entry) => ({
            timeLabel: entry.dt_txt.substring(11, 16),
            tempC: Math.round(entry.main.temp),
            variant: mapOwmIconToCondition(entry.weather[0]?.icon ?? '01d')
        }));
};

type LoadState = 'loading' | 'loaded' | 'error' | 'missingKey';

export interface IWeatherWidgetProps {
    // openweathermap.org üzerinden alınan API anahtarı — KASITLI OLARAK
    // kaynak koduna gömülmez (bkz. DashboardWebPart.ts), site yöneticisi
    // tarafından Property Pane'den girilir ve buraya prop olarak akar.
    apiKey: string;
}

const capitalize = (text: string): string => (text ? text.charAt(0).toUpperCase() + text.slice(1) : text);

const WeatherWidget: React.FunctionComponent<IWeatherWidgetProps> = ({ apiKey }) => {
    const theme = useTheme();
    const [state, setState] = React.useState<LoadState>('loading');
    const [weather, setWeather] = React.useState<IWeatherData | undefined>(undefined);

    React.useEffect(() => {
        let isMounted = true;

        const fetchWeather = async (lat: number, lon: number): Promise<void> => {
            if (!apiKey) {
                console.error(
                    '[WeatherWidget] Konum başarıyla alındı ama bir API anahtarı tanımlı değil — ' +
                    'web part düzenleme panelindeki "OpenWeatherMap API Anahtarı" alanını ' +
                    'openweathermap.org üzerinden alınan (ücretsiz) bir anahtarla doldurun.'
                );
                if (isMounted) {
                    setState('missingKey');
                }
                return;
            }

            try {
                const currentUrl = `https://api.openweathermap.org/data/2.5/weather?lat=${lat}&lon=${lon}&appid=${apiKey}&units=metric&lang=tr`;
                const response = await fetch(currentUrl);

                if (!response.ok) {
                    const detail = await response.text();
                    console.error('[WeatherWidget] OpenWeatherMap isteği başarısız:', response.status, detail);
                    if (isMounted) {
                        setState('error');
                    }
                    return;
                }

                const data: IOpenWeatherResponse = await response.json();
                const conditionInfo = data.weather?.[0];

                if (!data.main || !conditionInfo) {
                    console.error('[WeatherWidget] Beklenmeyen OpenWeatherMap yanıtı:', data);
                    if (isMounted) {
                        setState('error');
                    }
                    return;
                }

                // Haftalık tahmin AYRI ve İYİMSER (best-effort) çekiliyor — bu
                // istek her ne sebeple olursa olsun (ağ hatası, geçici 4xx/5xx,
                // vb.) başarısız olsa bile ana anlık durum ASLA etkilenmez.
                // ÖNCEKİ HATA: iki istek Promise.all ile birlikte bekleniyordu;
                // forecast isteği reddedilince (reject) TÜM widget "veri
                // ulaşılamıyor" hatasına düşüyordu, anlık durum çoktan
                // alınmış olsa bile.
                let forecast: IForecastDay[] = [];
                let hourly: IHourlyForecast[] = [];
                try {
                    const forecastUrl = `https://api.openweathermap.org/data/2.5/forecast?lat=${lat}&lon=${lon}&appid=${apiKey}&units=metric&lang=tr`;
                    const forecastResponse = await fetch(forecastUrl);
                    if (forecastResponse.ok) {
                        const forecastData: IOpenWeatherForecastResponse = await forecastResponse.json();
                        forecast = buildDailyForecast(forecastData.list ?? []);
                        hourly = buildTodayHourly(forecastData.list ?? []);
                    } else {
                        console.error('[WeatherWidget] Haftalık tahmin alınamadı:', forecastResponse.status);
                    }
                } catch (forecastError) {
                    console.error('[WeatherWidget] Haftalık tahmin alınırken hata (ana veri etkilenmedi):', forecastError);
                }

                if (isMounted) {
                    setWeather({
                        city: data.name ?? 'Konumunuz',
                        tempC: Math.round(data.main.temp),
                        feelsLikeC: Math.round(data.main.feels_like),
                        condition: capitalize(conditionInfo.description),
                        variant: mapOwmIconToCondition(conditionInfo.icon),
                        humidity: data.main.humidity,
                        windKmh: Math.round((data.wind?.speed ?? 0) * 3.6),
                        pressureHpa: data.main.pressure,
                        forecast,
                        hourly
                    });
                    setState('loaded');
                }
            } catch (error) {
                console.error('[WeatherWidget] Hava durumu alınamadı:', error);
                if (isMounted) {
                    setState('error');
                }
            }
        };

        if (!navigator.geolocation) {
            console.error('[WeatherWidget] Bu tarayıcı navigator.geolocation desteklemiyor.');
            setState('error');
            return undefined;
        }

        navigator.geolocation.getCurrentPosition(
            (position) => {
                fetchWeather(position.coords.latitude, position.coords.longitude).catch(() => { /* handled inside */ });
            },
            (geoError) => {
                console.error('[WeatherWidget] Konum alınamadı:', geoError);
                if (isMounted) {
                    setState('error');
                }
            },
            { timeout: 10000 }
        );

        return () => {
            isMounted = false;
        };
    }, [apiKey]);

    const stats: IWeatherStat[] = weather
        ? [
            { iconName: 'Drop', label: 'Nem', value: `%${weather.humidity}` },
            { iconName: 'WindDirection', label: 'Rüzgar', value: `${weather.windKmh} km/s` },
            { iconName: 'Info', label: 'Hissedilen', value: `${weather.feelsLikeC}°C` },
            // ÖNCEKİ HATA: 'Gauge' Fluent'in ikon setinde HİÇ YOKTU (font-icons-mdl2
            // paketinde böyle bir isim tanımlı değil) - bu yüzden ikon rozeti
            // bomboş görünüyordu ("basınç kısmında sembol yok" geri bildirimi).
            // 'CompassNW' gerçekten var olan, gösterge/kadran görünümlü bir ikon.
            { iconName: 'CompassNW', label: 'Basınç', value: `${weather.pressureHpa} hPa` }
        ]
        : [];

    const statColors = ['#0ea5c4', '#4f9d90', '#e08e45', '#8b6bd1'];

    const styles = mergeStyleSets({
        // NOT: "gap" flex özelliği burada BİLİNÇLİ OLARAK KULLANILMIYOR — bu
        // sayfanın render edildiği (kurumsal/eski) tarayıcı ortamında flex "gap"
        // desteklenmiyor. hero'ya marginBottom veriliyor; forecastRow ise
        // KOŞULLU render edildiği (weather.forecast.length > 0) için kendi
        // marginTop'unu taşıyor (statsRow'un olası son çocuk olmasına bağlı kalmadan).
        root: {
            display: 'flex',
            flexDirection: 'column',
            width: '100%',
            height: '100%'
        },
        // "Hero" paneli — mevcut alanı doldursun, düz/sıkıcı durmasın diye
        // gökyüzü tonlarında yumuşak bir gradyan zemin + ikonun arkasında
        // hafif bir parlama (glow) katmanı eklendi.
        hero: {
            position: 'relative',
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            width: '100%',
            padding: '22px 16px 18px',
            borderRadius: 16,
            overflow: 'hidden',
            background: `linear-gradient(160deg, ${theme.palette.themeLighter} 0%, ${theme.palette.neutralLighterAlt} 65%)`,
            marginBottom: 16
        },
        glow: {
            position: 'absolute',
            top: -30,
            left: '50%',
            transform: 'translateX(-50%)',
            width: 140,
            height: 140,
            borderRadius: '50%',
            background: 'radial-gradient(circle, rgba(95,168,172,0.35) 0%, rgba(95,168,172,0) 70%)',
            pointerEvents: 'none'
        },
        iconRow: {
            position: 'relative',
            zIndex: 1,
            display: 'flex',
            justifyContent: 'center'
        },
        // İkonun ALTINDA: solda derece, sağda konum — aralarında sabit bir
        // boşluk (gap) olduğu için metinler asla iç içe geçemez.
        infoRow: {
            position: 'relative',
            zIndex: 1,
            display: 'grid',
            gridTemplateColumns: '1fr 1fr',
            alignItems: 'baseline',
            columnGap: 16,
            width: '100%',
            marginTop: 6
        },
        temp: {
            display: 'block',
            fontSize: 42,
            fontWeight: 700,
            lineHeight: '46px',
            color: theme.semanticColors.bodyText,
            whiteSpace: 'nowrap',
            textAlign: 'left',
            overflow: 'hidden',
            textOverflow: 'ellipsis'
        },
        city: {
            display: 'block',
            fontSize: 16,
            fontWeight: 700,
            color: theme.palette.themePrimary,
            lineHeight: '20px',
            whiteSpace: 'nowrap',
            textAlign: 'right',
            overflow: 'hidden',
            textOverflow: 'ellipsis'
        },
        condition: {
            position: 'relative',
            zIndex: 1,
            display: 'block',
            width: '100%',
            fontSize: 13,
            fontWeight: 600,
            color: theme.semanticColors.bodySubtext,
            whiteSpace: 'nowrap',
            overflow: 'hidden',
            textOverflow: 'ellipsis',
            textAlign: 'center',
            marginTop: 6
        },
        statsRow: {
            display: 'grid',
            gridTemplateColumns: 'repeat(auto-fit, minmax(92px, 1fr))',
            gap: 12
        },
        sectionLabel: {
            fontSize: 11,
            fontWeight: 700,
            color: theme.semanticColors.bodySubtext,
            textTransform: 'uppercase',
            letterSpacing: 0.4,
            marginTop: 18,
            marginBottom: 8
        },
        // Bugünün saatlik tahmini — gün içinde 3-4 taneden fazla kayıt
        // olabileceği için grid yerine yatay kaydırılabilir bir şerit.
        hourlyRow: {
            display: 'flex',
            overflowX: 'auto',
            paddingBottom: 2
        },
        hourlySlot: {
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            flexShrink: 0,
            width: 56,
            padding: '8px 0'
        },
        hourlyTime: {
            fontSize: 10,
            fontWeight: 700,
            color: theme.semanticColors.bodySubtext,
            marginBottom: 6
        },
        hourlyIconWrap: {
            marginBottom: 6
        },
        hourlyTemp: {
            fontSize: 12,
            fontWeight: 700,
            color: theme.semanticColors.bodyText
        },
        forecastRow: {
            display: 'grid',
            gridTemplateColumns: 'repeat(5, 1fr)',
            gap: 8
        },
        // NOT: "gap" flex özelliği burada BİLİNÇLİ OLARAK KULLANILMIYOR — flex
        // "gap" desteklenmiyor. forecastLabel/forecastHigh kendi marginBottom'ını
        // taşıyor; AnimatedWeatherIcon className kabul etmediği için JSX'te bir
        // <span> ile sarmalanıp o span'a marginBottom veriliyor (forecastIconWrap).
        forecastDay: {
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            padding: '10px 2px',
            borderRadius: 12,
            minWidth: 0,
            background: theme.palette.neutralLighterAlt
        },
        forecastLabel: {
            fontSize: 10,
            fontWeight: 700,
            color: theme.semanticColors.bodySubtext,
            textTransform: 'uppercase',
            marginBottom: 4
        },
        forecastIconWrap: {
            marginBottom: 4
        },
        forecastHigh: {
            fontSize: 13,
            fontWeight: 700,
            color: theme.semanticColors.bodyText,
            marginBottom: 4
        },
        forecastLow: {
            fontSize: 10,
            color: theme.semanticColors.bodySubtext
        },
        // NOT: "gap" burada da kullanılmıyor — statIconWrap'e marginRight verildi.
        statChip: {
            display: 'flex',
            alignItems: 'center',
            padding: '10px 12px',
            borderRadius: 12,
            background: theme.palette.neutralLighterAlt,
            minWidth: 0
        },
        statIconWrap: {
            width: 30,
            height: 30,
            borderRadius: '50%',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            flexShrink: 0,
            marginRight: 10
        },
        statIcon: {
            fontSize: 14,
            color: '#ffffff',
            flexShrink: 0
        },
        statTextGroup: {
            display: 'flex',
            flexDirection: 'column',
            rowGap: 3,
            minWidth: 0
        },
        statValue: {
            display: 'block',
            fontSize: 14,
            fontWeight: 700,
            color: theme.semanticColors.bodyText,
            lineHeight: '16px',
            whiteSpace: 'nowrap',
            overflow: 'hidden',
            textOverflow: 'ellipsis'
        },
        statLabel: {
            display: 'block',
            fontSize: 10,
            color: theme.semanticColors.bodySubtext,
            lineHeight: '12px',
            whiteSpace: 'nowrap'
        },
        loadingWrap: {
            display: 'flex',
            alignItems: 'center',
            minHeight: 70
        }
    });

    return (
        <WidgetCard title="Hava Durumu" subtitle="Anlık konumunuza göre koşullar" iconName="Sunny">
            {state === 'loading' && (
                <div className={styles.loadingWrap}>
                    <Spinner size={SpinnerSize.medium} label="Konumunuz ve hava durumu alınıyor..." />
                </div>
            )}

            {state === 'error' && <MessageBar messageBarType={MessageBarType.error}>{DATA_UNAVAILABLE_MESSAGE}</MessageBar>}

            {state === 'missingKey' && (
                <MessageBar messageBarType={MessageBarType.warning}>
                    Konumunuz alındı, ancak hava durumu servisi için bir API anahtarı tanımlı değil.
                    Bu web part&apos;ı düzenleyip özellik panelinden (Property Pane) openweathermap.org
                    üzerinden alınan ücretsiz bir &quot;OpenWeatherMap API Anahtarı&quot; girmeniz gerekiyor.
                </MessageBar>
            )}

            {state === 'loaded' && weather && (
                <div className={styles.root}>
                    <div className={styles.hero}>
                        <span className={styles.glow} />
                        <div className={styles.iconRow}>
                            <AnimatedWeatherIcon condition={weather.variant} size={76} />
                        </div>
                        <div className={styles.infoRow}>
                            <div className={styles.temp}>{weather.tempC}°C</div>
                            <div className={styles.city}>{weather.city}</div>
                        </div>
                        <div className={styles.condition}>{weather.condition}</div>
                    </div>

                    <div className={styles.statsRow}>
                        {stats.map((stat, index) => (
                            <div key={stat.label} className={styles.statChip}>
                                <div
                                    className={styles.statIconWrap}
                                    style={{ background: statColors[index % statColors.length] }}
                                >
                                    <Icon iconName={stat.iconName} className={styles.statIcon} />
                                </div>
                                <div className={styles.statTextGroup}>
                                    <div className={styles.statValue}>{stat.value}</div>
                                    <div className={styles.statLabel}>{stat.label}</div>
                                </div>
                            </div>
                        ))}
                    </div>

                    {weather.hourly.length > 0 && (
                        <>
                            <div className={styles.sectionLabel}>Bugün Saatlik</div>
                            <div className={styles.hourlyRow}>
                                {weather.hourly.map((hour) => (
                                    <div key={hour.timeLabel} className={styles.hourlySlot}>
                                        <span className={styles.hourlyTime}>{hour.timeLabel}</span>
                                        <span className={styles.hourlyIconWrap}>
                                            <AnimatedWeatherIcon condition={hour.variant} size={24} />
                                        </span>
                                        <span className={styles.hourlyTemp}>{hour.tempC}°</span>
                                    </div>
                                ))}
                            </div>
                        </>
                    )}

                    {weather.forecast.length > 0 && (
                        <>
                            <div className={styles.sectionLabel}>5 Günlük Tahmin</div>
                            <div className={styles.forecastRow}>
                                {weather.forecast.map((day) => (
                                    <div key={day.dayLabel + day.high} className={styles.forecastDay}>
                                        <span className={styles.forecastLabel}>{day.dayLabel}</span>
                                        <span className={styles.forecastIconWrap}>
                                            <AnimatedWeatherIcon condition={day.variant} size={26} />
                                        </span>
                                        <span className={styles.forecastHigh}>{day.high}°</span>
                                        <span className={styles.forecastLow}>{day.low}°</span>
                                    </div>
                                ))}
                            </div>
                        </>
                    )}
                </div>
            )}
        </WidgetCard>
    );
};

export default WeatherWidget;
