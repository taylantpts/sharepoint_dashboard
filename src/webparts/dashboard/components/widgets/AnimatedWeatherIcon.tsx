import * as React from 'react';
import { keyframes, mergeStyleSets } from '@fluentui/react';

export type WeatherCondition = 'sunny' | 'partlyCloudy' | 'cloudy' | 'rainy' | 'stormy' | 'snowy';

export interface IAnimatedWeatherIconProps {
    condition: WeatherCondition;
    size?: number;
}

const spin = keyframes({
    from: { transform: 'rotate(0deg)' },
    to: { transform: 'rotate(360deg)' }
});

const pulse = keyframes({
    '0%, 100%': { opacity: 0.55, transform: 'scale(1)' },
    '50%': { opacity: 0.9, transform: 'scale(1.08)' }
});

const drift = keyframes({
    '0%, 100%': { transform: 'translateX(-3px)' },
    '50%': { transform: 'translateX(4px)' }
});

const fall = keyframes({
    '0%': { transform: 'translateY(0)', opacity: 0.9 },
    '100%': { transform: 'translateY(9px)', opacity: 0 }
});

const flash = keyframes({
    '0%, 100%': { opacity: 0 },
    '48%, 52%': { opacity: 0 },
    '50%': { opacity: 1 }
});

const sway = keyframes({
    '0%, 100%': { transform: 'translate(0, 0)' },
    '50%': { transform: 'translate(2px, 6px)' }
});

/**
 * Bağımlılık gerektirmeyen, saf SVG + CSS keyframes tabanlı animasyonlu hava
 * durumu ikonu (güneş yavaşça döner/parlar, bulut hafifçe sürüklenir). Harici
 * bir paket (ör. react-weather-widget) kuramadığımız için — bu ortamda
 * npm install çalıştırma izni yok — aynı görsel etkiyi bağımlılıksız üretir.
 */
const AnimatedWeatherIcon: React.FunctionComponent<IAnimatedWeatherIconProps> = (props) => {
    const { condition, size = 44 } = props;

    const styles = mergeStyleSets({
        root: {
            position: 'relative',
            width: size,
            height: size,
            flexShrink: 0,
            overflow: 'hidden'
        },
        rays: {
            position: 'absolute',
            inset: 0,
            animationName: spin,
            animationDuration: '18s',
            animationTimingFunction: 'linear',
            animationIterationCount: 'infinite'
        },
        glow: {
            position: 'absolute',
            inset: '18%',
            borderRadius: '50%',
            background: 'radial-gradient(circle, #ffd166 0%, rgba(255,209,102,0) 70%)',
            animationName: pulse,
            animationDuration: '3s',
            animationTimingFunction: 'ease-in-out',
            animationIterationCount: 'infinite'
        },
        cloud: {
            position: 'absolute',
            right: '0%',
            bottom: '0%',
            width: '62%',
            animationName: drift,
            animationDuration: '6s',
            animationTimingFunction: 'ease-in-out',
            animationIterationCount: 'infinite'
        },
        drop: {
            animationName: fall,
            animationDuration: '1.1s',
            animationTimingFunction: 'linear',
            animationIterationCount: 'infinite'
        },
        bolt: {
            animationName: flash,
            animationDuration: '2.4s',
            animationTimingFunction: 'ease-in-out',
            animationIterationCount: 'infinite'
        },
        flake: {
            animationName: sway,
            animationDuration: '2.6s',
            animationTimingFunction: 'ease-in-out',
            animationIterationCount: 'infinite'
        }
    });

    const showSun = condition === 'sunny' || condition === 'partlyCloudy';
    const showCloud =
        condition === 'partlyCloudy' ||
        condition === 'cloudy' ||
        condition === 'rainy' ||
        condition === 'stormy' ||
        condition === 'snowy';
    const showRain = condition === 'rainy' || condition === 'stormy';
    const showBolt = condition === 'stormy';
    const showSnow = condition === 'snowy';

    return (
        <div className={styles.root} aria-hidden="true">
            {showSun && (
                <>
                    <div className={styles.glow} />
                    <svg viewBox="0 0 100 100" className={styles.rays}>
                        <g stroke="#f2a900" strokeWidth="4" strokeLinecap="round">
                            <line x1="50" y1="4" x2="50" y2="16" />
                            <line x1="50" y1="84" x2="50" y2="96" />
                            <line x1="4" y1="50" x2="16" y2="50" />
                            <line x1="84" y1="50" x2="96" y2="50" />
                            <line x1="17" y1="17" x2="25" y2="25" />
                            <line x1="75" y1="75" x2="83" y2="83" />
                            <line x1="17" y1="83" x2="25" y2="75" />
                            <line x1="75" y1="25" x2="83" y2="17" />
                        </g>
                    </svg>
                    <svg viewBox="0 0 100 100" style={{ position: 'absolute', inset: 0 }}>
                        <circle cx="50" cy="50" r="24" fill="#ffb900" />
                    </svg>
                </>
            )}
            {showCloud && (
                <svg viewBox="0 0 64 54" className={styles.cloud}>
                    <path
                        d="M18 30c-6 0-11-4.7-11-10.5S12 9 18 9c1.6 0 3.1.3 4.5 1C24.5 5.4 29.4 2 35 2c7.6 0 13.8 5.9 14.4 13.4 5.4 1 9.6 5.6 9.6 11.1 0 6.3-5.2 11.5-11.6 11.5H18z"
                        fill={showRain || showSnow ? '#e4ebf3' : '#ffffff'}
                        stroke="#c8d6e5"
                        strokeWidth="1.5"
                    />
                    {showRain && (
                        <g stroke="#4a90d9" strokeWidth="2.4" strokeLinecap="round">
                            <line x1="19" y1="33" x2="16" y2="42" className={styles.drop} style={{ animationDelay: '0s' }} />
                            <line x1="31" y1="33" x2="28" y2="42" className={styles.drop} style={{ animationDelay: '0.35s' }} />
                            <line x1="43" y1="33" x2="40" y2="42" className={styles.drop} style={{ animationDelay: '0.7s' }} />
                        </g>
                    )}
                    {showBolt && (
                        <path
                            d="M33 30 24 44h7l-3 10 12-16h-8z"
                            fill="#f2c94c"
                            stroke="#c98f0a"
                            strokeWidth="0.5"
                            strokeLinejoin="round"
                            className={styles.bolt}
                        />
                    )}
                    {showSnow && (
                        <g fill="#8fb8e8">
                            <circle cx="19" cy="38" r="2.2" className={styles.flake} style={{ animationDelay: '0s' }} />
                            <circle cx="31" cy="41" r="2.2" className={styles.flake} style={{ animationDelay: '0.5s' }} />
                            <circle cx="43" cy="38" r="2.2" className={styles.flake} style={{ animationDelay: '1s' }} />
                        </g>
                    )}
                </svg>
            )}
        </div>
    );
};

export default AnimatedWeatherIcon;
