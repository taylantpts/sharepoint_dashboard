import * as React from 'react';
import { Icon, useTheme, mergeStyleSets } from '@fluentui/react';
import WidgetCard from '../WidgetCard';

const DAILY_GOAL = 8; // bardak (~250 ml/bardak, ~2000 ml/gün)
const ML_PER_GLASS = 250;
const STORAGE_PREFIX = 'yorpas-dashboard-water-intake-';

const todayKey = (): string => {
    const now = new Date();
    return `${STORAGE_PREFIX}${now.getFullYear()}-${now.getMonth() + 1}-${now.getDate()}`;
};

// localStorage'a yazılıyor — sunucu tarafında/tüm çalışanlarda ortak bir
// sayaç DEĞİL, sadece bu tarayıcıda "bugün kaç bardak içtim" hatırlatıcısı.
// Anahtar güne göre değiştiği için (todayKey) gece yarısı otomatik sıfırlanmış
// gibi davranır — ayrıca elle "Sıfırla" ile de sıfırlanabilir.
const readCount = (): number => {
    const raw = window.localStorage.getItem(todayKey());
    const parsed = raw ? parseInt(raw, 10) : 0;
    return Number.isNaN(parsed) ? 0 : parsed;
};

const CIRCLE_R = 42;
const CIRCUMFERENCE = 2 * Math.PI * CIRCLE_R;

const WaterIntakeWidget: React.FunctionComponent = () => {
    const theme = useTheme();
    const [count, setCount] = React.useState<number>(readCount);

    const persist = (next: number): void => {
        setCount(next);
        window.localStorage.setItem(todayKey(), String(next));
    };

    const progress = Math.min(count / DAILY_GOAL, 1);
    const dashOffset = CIRCUMFERENCE * (1 - progress);
    const accent = theme.palette.themePrimary;

    const styles = mergeStyleSets({
        root: {
            display: 'flex',
            flexDirection: 'column',
            height: '100%'
        },
        contentCol: {
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            justifyContent: 'center',
            flexGrow: 1,
            minHeight: 0
        },
        ringWrap: {
            position: 'relative',
            width: '100%',
            maxWidth: 200,
            aspectRatio: '1 / 1',
            marginBottom: 18
        },
        ringSvg: {
            width: '100%',
            height: '100%',
            transform: 'rotate(-90deg)'
        },
        ringCenter: {
            position: 'absolute',
            top: 0,
            left: 0,
            right: 0,
            bottom: 0,
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            justifyContent: 'center'
        },
        ringCount: {
            fontSize: 34,
            fontWeight: 700,
            color: theme.semanticColors.bodyText,
            lineHeight: '38px'
        },
        ringGoal: {
            fontSize: 12,
            color: theme.semanticColors.bodySubtext,
            marginTop: 2
        },
        mlHint: {
            fontSize: 12,
            color: theme.semanticColors.bodySubtext,
            marginBottom: 16
        },
        buttonRow: {
            display: 'flex',
            alignItems: 'center'
        },
        addButton: {
            border: 'none',
            borderRadius: 10,
            padding: '10px 20px',
            fontSize: 13,
            fontWeight: 600,
            color: '#ffffff',
            cursor: 'pointer',
            background: `linear-gradient(135deg, ${accent} 0%, ${accent}cc 100%)`,
            boxShadow: `0 6px 16px ${accent}45`,
            display: 'flex',
            alignItems: 'center',
            marginRight: 12,
            selectors: {
                ':disabled': {
                    opacity: 0.5,
                    cursor: 'default',
                    boxShadow: 'none'
                }
            }
        },
        addIcon: {
            fontSize: 13,
            marginRight: 6
        },
        resetLink: {
            border: 'none',
            background: 'transparent',
            fontSize: 12,
            color: theme.semanticColors.bodySubtext,
            cursor: 'pointer',
            textDecoration: 'underline',
            padding: 0
        }
    });

    return (
        <WidgetCard title="Su İçme Takibi" subtitle="Bugün ne kadar içtin?" iconName="DropShapeSolid">
            <div className={styles.root}>
                <div className={styles.contentCol}>
                    <div className={styles.ringWrap}>
                        <svg viewBox="0 0 100 100" className={styles.ringSvg}>
                            <circle cx="50" cy="50" r={CIRCLE_R} fill="none" stroke={theme.palette.neutralLighterAlt} strokeWidth={10} />
                            <circle
                                cx="50"
                                cy="50"
                                r={CIRCLE_R}
                                fill="none"
                                stroke={accent}
                                strokeWidth={10}
                                strokeLinecap="round"
                                strokeDasharray={CIRCUMFERENCE}
                                strokeDashoffset={dashOffset}
                                style={{ transition: 'stroke-dashoffset 0.3s ease' }}
                            />
                        </svg>
                        <div className={styles.ringCenter}>
                            <div className={styles.ringCount}>{count}</div>
                            <div className={styles.ringGoal}>/ {DAILY_GOAL} bardak</div>
                        </div>
                    </div>
                    <div className={styles.mlHint}>
                        ≈ {count * ML_PER_GLASS} ml / {DAILY_GOAL * ML_PER_GLASS} ml
                    </div>
                    <div className={styles.buttonRow}>
                        <button type="button" className={styles.addButton} onClick={() => persist(count + 1)}>
                            <Icon iconName="Add" className={styles.addIcon} />
                            Bardak Ekle
                        </button>
                        <button type="button" className={styles.resetLink} onClick={() => persist(0)}>
                            Sıfırla
                        </button>
                    </div>
                </div>
            </div>
        </WidgetCard>
    );
};

export default WaterIntakeWidget;
