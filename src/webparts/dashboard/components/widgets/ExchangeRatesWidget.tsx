import * as React from 'react';
import { Spinner, SpinnerSize, MessageBar, MessageBarType, useTheme, mergeStyleSets } from '@fluentui/react';
import WidgetCard from '../WidgetCard';
import { DATA_UNAVAILABLE_MESSAGE } from '../../constants';

// fxratesapi.com — anahtarsız/ücretsiz, hem güncel hem GEÇMİŞ (timeseries) veri
// veriyor, ayrıca altın/gümüş (XAU/XAG) de aynı API'de "para birimi" gibi
// destekleniyor — tek kaynaktan hepsi çekilebiliyor. ÖNCEKİ HATA: frankfurter.app
// PowerShell'den çalışıyor göründüğü halde GERÇEK tarayıcıdan (muhtemelen bir
// proxy/SSL katmanı) tutarsız bir CORS başlığı yüzünden her zaman "Failed to
// fetch" veriyordu — bu API doğrudan tarayıcıdan test edilip doğrulandı.
const TROY_OUNCE_IN_GRAMS = 31.1034768;
const HISTORY_DAYS = 30;
const CHART_WIDTH = 300;
const CHART_HEIGHT = 120;
const CHART_PADDING = 8;

interface IAssetDef {
    code: string;
    label: string;
    unit: string;
    isMetal?: boolean;
    chartColor: string;
}

const ASSETS: IAssetDef[] = [
    { code: 'USD', label: 'Dolar', unit: '$', chartColor: '#2f6f76' },
    { code: 'EUR', label: 'Euro', unit: '€', chartColor: '#4a6fa5' },
    { code: 'GBP', label: 'Sterlin', unit: '£', chartColor: '#7b8fd4' },
    { code: 'XAU', label: 'Altın (gr)', unit: '₺/gr', isMetal: true, chartColor: '#c9a227' },
    { code: 'XAG', label: 'Gümüş (gr)', unit: '₺/gr', isMetal: true, chartColor: '#9099a3' }
];

interface ISeriesPoint {
    date: string;
    value: number;
}

interface IAssetData {
    current: number;
    series: ISeriesPoint[];
    changePct?: number;
    min: number;
    max: number;
}

interface IFxRatesLatestResponse {
    success: boolean;
    rates: Record<string, number>;
}

interface IFxRatesTimeseriesResponse {
    success: boolean;
    rates: Record<string, Record<string, number>>;
}

type LoadState = 'loading' | 'loaded' | 'error';

const toIsoDate = (date: Date): string => date.toISOString().substring(0, 10);

const formatTry = (value: number): string =>
    value.toLocaleString('tr-TR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

const formatChartDate = (iso: string): string =>
    new Date(iso).toLocaleDateString('tr-TR', { day: 'numeric', month: 'short' });

// fxratesapi TRY tabanlı döner (1 TRY = X birim) — göstermek istediğimiz
// "1 birim = Y TRY" için tersini (ve metal ise gram'a çevirmek için ons
// bölücüsünü) uygularız.
const toTryValue = (tryBasedRate: number, isMetal?: boolean): number => {
    const perUnit = 1 / tryBasedRate;
    return isMetal ? perUnit / TROY_OUNCE_IN_GRAMS : perUnit;
};

/** Bir seri için 0..width / 0..height viewBox koordinatlarına ölçekleyen yardımcı — hem büyük grafik hem mini sparkline'lar tarafından paylaşılır. */
const buildScaler = (points: ISeriesPoint[], width: number, height: number, padding: number) => {
    const values = points.map((p) => p.value);
    const min = Math.min(...values);
    const max = Math.max(...values);
    const range = max - min || 1;
    return (index: number, value: number): [number, number] => [
        (index / Math.max(1, points.length - 1)) * width,
        height - padding - ((value - min) / range) * (height - padding * 2)
    ];
};

const MiniSparkline: React.FunctionComponent<{ points: ISeriesPoint[]; color: string }> = ({ points, color }) => {
    if (points.length < 2) {
        return null;
    }
    const width = 64;
    const height = 22;
    const toXY = buildScaler(points, width, height, 2);
    const linePoints = points.map((p, i) => toXY(i, p.value).join(',')).join(' ');
    return (
        <svg viewBox={`0 0 ${width} ${height}`} width={width} height={height} preserveAspectRatio="none">
            <polyline points={linePoints} fill="none" stroke={color} strokeWidth={1.6} strokeLinejoin="round" strokeLinecap="round" />
        </svg>
    );
};

const ExchangeRatesWidget: React.FunctionComponent = () => {
    const theme = useTheme();
    const [state, setState] = React.useState<LoadState>('loading');
    const [dataByCode, setDataByCode] = React.useState<Record<string, IAssetData>>({});
    const [selectedCode, setSelectedCode] = React.useState<string>('USD');
    const [hoverIndex, setHoverIndex] = React.useState<number | undefined>(undefined);

    React.useEffect(() => {
        let isMounted = true;
        const codes = ASSETS.map((a) => a.code).join(',');

        const load = async (): Promise<void> => {
            try {
                const today = new Date();
                const startDate = new Date();
                startDate.setDate(startDate.getDate() - HISTORY_DAYS);

                const [latestResponse, seriesResponse] = await Promise.all([
                    fetch(`https://api.fxratesapi.com/latest?base=TRY&currencies=${codes}`),
                    fetch(
                        `https://api.fxratesapi.com/timeseries?base=TRY&currencies=${codes}` +
                            `&start_date=${toIsoDate(startDate)}&end_date=${toIsoDate(today)}`
                    )
                ]);

                if (!latestResponse.ok || !seriesResponse.ok) {
                    throw new Error(`fxratesapi isteği başarısız (HTTP ${latestResponse.status}/${seriesResponse.status})`);
                }

                const latest: IFxRatesLatestResponse = await latestResponse.json();
                const series: IFxRatesTimeseriesResponse = await seriesResponse.json();
                if (!latest.success || !series.success) {
                    throw new Error('fxratesapi beklenmeyen bir sonuç döndürdü.');
                }

                const sortedDates = Object.keys(series.rates).sort();

                const next: Record<string, IAssetData> = {};
                ASSETS.forEach((asset) => {
                    const points: ISeriesPoint[] = sortedDates
                        .filter((d) => series.rates[d][asset.code] !== undefined)
                        .map((d) => ({ date: d, value: toTryValue(series.rates[d][asset.code], asset.isMetal) }));

                    const values = points.map((p) => p.value);
                    const changePct =
                        points.length > 1 ? ((points[points.length - 1].value - points[0].value) / points[0].value) * 100 : undefined;

                    next[asset.code] = {
                        current: toTryValue(latest.rates[asset.code], asset.isMetal),
                        series: points,
                        changePct,
                        min: values.length ? Math.min(...values) : 0,
                        max: values.length ? Math.max(...values) : 0
                    };
                });

                if (isMounted) {
                    setDataByCode(next);
                    setState('loaded');
                }
            } catch (error) {
                console.error('[ExchangeRatesWidget] Döviz/kıymetli maden kurları alınamadı:', error);
                if (isMounted) {
                    setState('error');
                }
            }
        };

        load().catch(() => { /* load kendi içinde hatayı yönetir */ });

        return () => {
            isMounted = false;
        };
    }, []);

    // Varlık değiştiğinde önceki hover durumu anlamsız kalır (farklı bir
    // seriye ait bir indeksi işaret ediyor olabilir).
    React.useEffect(() => {
        setHoverIndex(undefined);
    }, [selectedCode]);

    const styles = mergeStyleSets({
        tileRow: {
            display: 'grid',
            gridTemplateColumns: 'repeat(auto-fit, minmax(96px, 1fr))',
            gap: 10,
            marginBottom: 20
        },
        tile: {
            position: 'relative',
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            padding: '14px 8px 10px',
            borderRadius: 14,
            border: `1px solid ${theme.palette.neutralLight}`,
            background: theme.palette.neutralLighterAlt,
            cursor: 'pointer',
            font: 'inherit',
            overflow: 'hidden',
            transition: 'transform 0.15s ease, box-shadow 0.15s ease',
            selectors: {
                ':hover': { transform: 'translateY(-3px)', boxShadow: '0 6px 16px rgba(0,0,0,0.12)' }
            }
        },
        tileTopBar: {
            position: 'absolute',
            top: 0,
            left: 0,
            right: 0,
            height: 3
        },
        tileCode: {
            fontSize: 11,
            fontWeight: 700,
            letterSpacing: 0.4,
            color: theme.semanticColors.bodySubtext,
            marginBottom: 6
        },
        tileValue: {
            fontSize: 15,
            fontWeight: 800,
            color: theme.semanticColors.bodyText,
            marginBottom: 6
        },
        chartSection: {
            borderTop: `1px solid ${theme.palette.neutralLighter}`,
            paddingTop: 16
        },
        chartHeader: {
            display: 'flex',
            alignItems: 'flex-end',
            justifyContent: 'space-between',
            marginBottom: 12
        },
        chartTitleGroup: {
            display: 'flex',
            flexDirection: 'column'
        },
        chartTitle: {
            fontSize: 12,
            fontWeight: 600,
            color: theme.semanticColors.bodySubtext,
            marginBottom: 2
        },
        chartBigValue: {
            fontSize: 26,
            fontWeight: 800,
            color: theme.semanticColors.bodyText,
            fontVariantNumeric: 'tabular-nums'
        },
        chartChange: {
            fontSize: 13,
            fontWeight: 700,
            padding: '4px 10px',
            borderRadius: 20
        },
        chartContainer: {
            position: 'relative',
            width: '100%',
            height: CHART_HEIGHT,
            cursor: 'crosshair'
        },
        chartSvg: {
            position: 'absolute',
            inset: 0,
            width: '100%',
            height: '100%',
            display: 'block'
        },
        // Kılavuz çizgisi, işaretçi noktası ve tooltip artık SVG İÇİNDE DEĞİL,
        // yüzde bazlı konumlandırılmış normal HTML katmanları olarak çiziliyor.
        // ÖNCEKİ HATA: SVG viewBox'ı (300x120) ile gerçek render edilen konteyner
        // en/boy oranı (ör. 580x120) ÖRTÜŞMÜYORDU — preserveAspectRatio="none"
        // ile yatay eksende orantısız (non-uniform) bir gerilme oluyor, bu da
        // <text>/<circle> gibi SVG içeriğinin yamulmuş/kesik görünmesine yol
        // açıyordu. Çizgi/alan grafiği için bu gerilme (bir eğrinin yatayda
        // esnemesi) sorun değil, ama okunması gereken metin ve nokta için HTML
        // katmanı kullanmak metnin HER ZAMAN doğal/net boyutunda kalmasını sağlar.
        guideLine: {
            position: 'absolute',
            top: 0,
            bottom: 0,
            width: 0,
            borderLeft: '1px dashed',
            opacity: 0.5,
            pointerEvents: 'none'
        },
        markerDot: {
            position: 'absolute',
            width: 10,
            height: 10,
            borderRadius: '50%',
            background: '#ffffff',
            transform: 'translate(-50%, -50%)',
            pointerEvents: 'none',
            boxShadow: '0 1px 3px rgba(0,0,0,0.25)'
        },
        tooltipBubble: {
            position: 'absolute',
            top: 4,
            transform: 'translateY(0)',
            padding: '4px 8px',
            borderRadius: 8,
            background: theme.palette.white,
            boxShadow: '0 2px 8px rgba(0,0,0,0.18)',
            pointerEvents: 'none',
            whiteSpace: 'nowrap'
        },
        tooltipValue: {
            fontSize: 12,
            fontWeight: 700,
            color: theme.semanticColors.bodyText,
            lineHeight: '15px'
        },
        tooltipDate: {
            fontSize: 10,
            color: theme.semanticColors.bodySubtext,
            lineHeight: '13px'
        },
        statsRow: {
            display: 'flex',
            justifyContent: 'space-between',
            marginTop: 10,
            paddingTop: 10,
            borderTop: `1px dashed ${theme.palette.neutralLighter}`
        },
        statItem: {
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            flexGrow: 1
        },
        statLabel: {
            fontSize: 10,
            color: theme.semanticColors.bodySubtext,
            textTransform: 'uppercase',
            letterSpacing: 0.4,
            marginBottom: 2
        },
        statValue: {
            fontSize: 13,
            fontWeight: 700,
            color: theme.semanticColors.bodyText
        }
    });

    const selectedAsset = ASSETS.filter((a) => a.code === selectedCode)[0];
    const selectedData = dataByCode[selectedCode];

    const handleChartMouseMove = (event: React.MouseEvent<HTMLDivElement>, pointCount: number): void => {
        const bounds = event.currentTarget.getBoundingClientRect();
        const ratio = (event.clientX - bounds.left) / bounds.width;
        const index = Math.round(ratio * (pointCount - 1));
        setHoverIndex(Math.max(0, Math.min(pointCount - 1, index)));
    };

    // Elle çizilmiş, bağımlılıksız bir SVG alan grafiği — çizgi/alan SVG'de
    // (yatayda esnemesi sorun değil), ama fareyle en yakın noktayı vurgulayan
    // kılavuz çizgisi/nokta/tooltip yüzde bazlı normal HTML katmanları (bkz.
    // guideLine/markerDot/tooltipBubble stilleri) — metin/nokta HER ZAMAN
    // doğal boyutunda, gerilmeden kalır.
    const renderChart = (asset: IAssetDef, data: IAssetData): React.ReactNode => {
        const points = data.series;
        if (points.length < 2) {
            return null;
        }
        const toXY = buildScaler(points, CHART_WIDTH, CHART_HEIGHT, CHART_PADDING);
        const linePoints = points.map((p, i) => toXY(i, p.value).join(',')).join(' ');
        const areaPath = `M0,${CHART_HEIGHT} L${linePoints} L${CHART_WIDTH},${CHART_HEIGHT} Z`;

        const activeIndex = hoverIndex ?? points.length - 1;
        const xPercent = (activeIndex / (points.length - 1)) * 100;
        const [, activeYUnits] = toXY(activeIndex, points[activeIndex].value);
        const yPercent = (activeYUnits / CHART_HEIGHT) * 100;
        const labelOnRight = xPercent > 65;

        return (
            <div
                className={styles.chartContainer}
                onMouseMove={(e) => handleChartMouseMove(e, points.length)}
                onMouseLeave={() => setHoverIndex(undefined)}
            >
                <svg viewBox={`0 0 ${CHART_WIDTH} ${CHART_HEIGHT}`} className={styles.chartSvg} preserveAspectRatio="none">
                    <defs>
                        <linearGradient id={`fill-${asset.code}`} x1="0" y1="0" x2="0" y2="1">
                            <stop offset="0%" stopColor={asset.chartColor} stopOpacity={0.38} />
                            <stop offset="100%" stopColor={asset.chartColor} stopOpacity={0} />
                        </linearGradient>
                    </defs>
                    <path d={areaPath} fill={`url(#fill-${asset.code})`} stroke="none" />
                    <polyline points={linePoints} fill="none" stroke={asset.chartColor} strokeWidth={2.5} strokeLinejoin="round" />
                </svg>
                <span className={styles.guideLine} style={{ left: `${xPercent}%`, borderColor: asset.chartColor }} />
                <span className={styles.markerDot} style={{ left: `${xPercent}%`, top: `${yPercent}%`, border: `2.5px solid ${asset.chartColor}` }} />
                <div
                    className={styles.tooltipBubble}
                    style={
                        labelOnRight
                            ? { right: `${100 - xPercent}%`, transform: 'translateX(-12px)' }
                            : { left: `${xPercent}%`, transform: 'translateX(12px)' }
                    }
                >
                    <div className={styles.tooltipValue}>{formatTry(points[activeIndex].value)}</div>
                    <div className={styles.tooltipDate}>{formatChartDate(points[activeIndex].date)}</div>
                </div>
            </div>
        );
    };

    return (
        <WidgetCard title="Döviz & Kıymetli Maden" subtitle="TRY karşılığı güncel kurlar, dokunarak inceleyin" iconName="Bank">
            {state === 'loading' && <Spinner size={SpinnerSize.medium} label="Kurlar yükleniyor..." />}
            {state === 'error' && <MessageBar messageBarType={MessageBarType.error}>{DATA_UNAVAILABLE_MESSAGE}</MessageBar>}
            {state === 'loaded' && (
                <>
                    <div className={styles.tileRow}>
                        {ASSETS.map((asset) => {
                            const assetData = dataByCode[asset.code];
                            const isSelected = asset.code === selectedCode;
                            return (
                                <button
                                    key={asset.code}
                                    type="button"
                                    className={styles.tile}
                                    style={
                                        isSelected
                                            ? { borderColor: asset.chartColor, background: `${asset.chartColor}14` }
                                            : undefined
                                    }
                                    onClick={() => setSelectedCode(asset.code)}
                                >
                                    <span className={styles.tileTopBar} style={{ background: asset.chartColor, opacity: isSelected ? 1 : 0.35 }} />
                                    <div className={styles.tileCode}>{asset.code}</div>
                                    <div className={styles.tileValue}>{assetData ? formatTry(assetData.current) : '—'}</div>
                                    {assetData && <MiniSparkline points={assetData.series} color={asset.chartColor} />}
                                </button>
                            );
                        })}
                    </div>

                    {selectedAsset && selectedData && (
                        <div className={styles.chartSection}>
                            <div className={styles.chartHeader}>
                                <div className={styles.chartTitleGroup}>
                                    <span className={styles.chartTitle}>{selectedAsset.label} — Son {HISTORY_DAYS} Gün</span>
                                    <span className={styles.chartBigValue}>{formatTry(selectedData.current)} <small>{selectedAsset.unit}</small></span>
                                </div>
                                {selectedData.changePct !== undefined && (
                                    <span
                                        className={styles.chartChange}
                                        style={{
                                            color: selectedData.changePct >= 0 ? '#0f7b3c' : '#c4314b',
                                            background: selectedData.changePct >= 0 ? 'rgba(15,123,60,0.10)' : 'rgba(196,49,75,0.10)'
                                        }}
                                    >
                                        {selectedData.changePct >= 0 ? '▲' : '▼'} %{Math.abs(selectedData.changePct).toFixed(2)}
                                    </span>
                                )}
                            </div>
                            {renderChart(selectedAsset, selectedData)}
                            <div className={styles.statsRow}>
                                <div className={styles.statItem}>
                                    <span className={styles.statLabel}>30G Düşük</span>
                                    <span className={styles.statValue}>{formatTry(selectedData.min)}</span>
                                </div>
                                <div className={styles.statItem}>
                                    <span className={styles.statLabel}>30G Yüksek</span>
                                    <span className={styles.statValue}>{formatTry(selectedData.max)}</span>
                                </div>
                                <div className={styles.statItem}>
                                    <span className={styles.statLabel}>Güncelleme</span>
                                    <span className={styles.statValue}>{formatChartDate(selectedData.series[selectedData.series.length - 1].date)}</span>
                                </div>
                            </div>
                        </div>
                    )}
                </>
            )}
        </WidgetCard>
    );
};

export default ExchangeRatesWidget;
