import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { getUpcomingEvents } from '../services/SharePointService';

export interface IDashboardSummary {
    upcomingEventsCount: number;
}

/**
 * Duyuru ve ilan "yeni" sayıları artık useNotificationHistory'nin paylaşılan
 * geçmişinden geliyor (bkz. o dosyadaki not) — bu hook sadece "yaklaşan
 * etkinlik" ÇİPİ için gerekli TOPLAM (yeni değil, mevcut) etkinlik sayısını
 * sağlıyor; bu ikisi kasıtlı olarak farklı anlamlar taşıyor (çip metni de
 * "X yeni etkinlik" değil "X yaklaşan etkinlik" der).
 */
export const useDashboardSummary = (context: WebPartContext): IDashboardSummary | undefined => {
    const [summary, setSummary] = React.useState<IDashboardSummary | undefined>(undefined);

    React.useEffect(() => {
        let isMounted = true;
        getUpcomingEvents(context)
            .then((events) => {
                if (isMounted) {
                    setSummary({ upcomingEventsCount: events.length });
                }
            })
            .catch(() => { /* özet şeridi opsiyonel bir detaydır, sessizce boş kalır */ });
        return () => {
            isMounted = false;
        };
    }, [context]);

    return summary;
};
