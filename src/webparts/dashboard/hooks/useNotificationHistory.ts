import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import {
    syncNotificationHistory, markNotificationRead, markCategoryRead,
    INotificationEntry, NotificationCategory
} from '../services/NotificationService';

export interface IUseNotificationHistory {
    history: INotificationEntry[];
    /** İlk senkronizasyon tamamlanana kadar false — sayaçların gerçek veriden
     * ÖNCE kısaca "0" gösterip sonra sıçramasını (flaş) önlemek için. */
    isLoaded: boolean;
    markRead: (entryId: string) => void;
    markCategoryAsRead: (category: NotificationCategory) => void;
}

/**
 * NotificationBell (zil) VE WelcomeHeader'ın özet şeridi (çipler) AYNI
 * bildirim geçmişini göstermesi gerekiyor — ÖNCEKİ HATA: bu ikisi
 * BİRBİRİNDEN BAĞIMSIZ, kendi kendine "yenilik" hesaplıyordu
 * (NotificationBell kendi syncNotificationHistory'sini çağırıyor, özet
 * şeridi ayrıca getSeenIds okuyordu). syncNotificationHistory bir delta
 * tespit eder etmez taban (baseline) id'yi HEMEN ileri aldığı için, zil
 * senkronizasyonu özet şeridinin okumasından ÖNCE tamamlanırsa özet şeridi
 * "0 yeni X" görüyordu — kullanıcı yeni bir ilan eklediğinde bile. Tek bir
 * paylaşılan hook (tek bir syncNotificationHistory çağrısı, WelcomeHeader
 * tarafından sahiplenilir) bu yarış durumunu tamamen ortadan kaldırır: hem
 * zil hem özet şeridi AYNI `history` state'ini okur.
 */
export const useNotificationHistory = (context: WebPartContext): IUseNotificationHistory => {
    const loginName = context.pageContext.user.loginName;
    const [history, setHistory] = React.useState<INotificationEntry[]>([]);
    const [isLoaded, setIsLoaded] = React.useState(false);

    React.useEffect(() => {
        let isMounted = true;
        syncNotificationHistory(context, loginName)
            .then(({ history: synced }) => {
                if (isMounted) {
                    setHistory(synced);
                    setIsLoaded(true);
                }
            })
            .catch((error: Error) => {
                console.error('[useNotificationHistory] Bildirim geçmişi güncellenemedi:', error);
                if (isMounted) {
                    setIsLoaded(true);
                }
            });
        return () => {
            isMounted = false;
        };
    }, [context, loginName]);

    const markRead = React.useCallback((entryId: string): void => {
        setHistory(markNotificationRead(loginName, entryId));
    }, [loginName]);

    const markCategoryAsRead = React.useCallback((category: NotificationCategory): void => {
        setHistory(markCategoryRead(loginName, category));
    }, [loginName]);

    return { history, isLoaded, markRead, markCategoryAsRead };
};
