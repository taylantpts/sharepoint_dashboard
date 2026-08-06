/**
 * ÖNCEKİ HATA: native scrollIntoView({behavior:'smooth'}) bu render ortamında
 * SESSİZCE hiçbir şey yapmıyordu (canlı ortamda doğrulandı — bkz.
 * NotificationBell.tsx'teki eski not) — bu yüzden 'auto' (anlık atlama)
 * kullanılmıştı. Kullanıcı GERÇEKTEN yumuşak bir kaydırma istiyor; bu dosya
 * tarayıcının bozuk native smooth-scroll'una GÜVENMEDEN, requestAnimationFrame
 * ile KENDİ kaydırma animasyonumuzu üretir.
 *
 * Hangi elemanın gerçekte kaydığı (window mu, yoksa SharePoint workbench'in
 * kendi bir sarmalayıcı div'i mi) bu ortamda garanti değil — bu yüzden HER
 * adımda olası TÜM adaylara (window, documentElement, body, ve hedefin en
 * yakın "gerçekten taşan" atası) aynı anda scrollTop yazılır. Kaydırmayı
 * gerçekten kontrol eden hangisiyse o hareket eder, diğerlerine yazmak
 * zararsız bir no-op'tur.
 */

const findScrollableAncestor = (el: HTMLElement): HTMLElement | undefined => {
    let node = el.parentElement;
    while (node) {
        if (node.scrollHeight > node.clientHeight + 1) {
            return node;
        }
        node = node.parentElement;
    }
    return undefined;
};

const getCurrentScrollTop = (container: HTMLElement | undefined): number =>
    container ? container.scrollTop : (window.scrollY || document.documentElement.scrollTop || document.body.scrollTop);

const setScrollTop = (container: HTMLElement | undefined, y: number): void => {
    if (container) {
        container.scrollTop = y;
    }
    window.scrollTo(0, y);
    document.documentElement.scrollTop = y;
    document.body.scrollTop = y;
};

const easeInOutCubic = (t: number): number => (t < 0.5 ? 4 * t * t * t : 1 - Math.pow(-2 * t + 2, 3) / 2);

/** Hedef elemanı, yumuşak (kendi ürettiğimiz) bir animasyonla dikey ortaya yaklaşacak şekilde kaydırır. */
export const smoothScrollToElement = (target: HTMLElement, durationMs = 550): void => {
    const container = findScrollableAncestor(target);
    const startY = getCurrentScrollTop(container);
    const targetRectTop = target.getBoundingClientRect().top;
    const targetHeight = target.getBoundingClientRect().height;
    const desiredViewportTop = window.innerHeight / 2 - targetHeight / 2;
    const endY = Math.max(0, startY + (targetRectTop - desiredViewportTop));

    if (Math.abs(endY - startY) < 2) {
        return;
    }

    const startTime = performance.now();

    const step = (now: number): void => {
        const progress = Math.min(1, (now - startTime) / durationMs);
        setScrollTop(container, startY + (endY - startY) * easeInOutCubic(progress));
        if (progress < 1) {
            window.requestAnimationFrame(step);
        }
    };
    window.requestAnimationFrame(step);
};

// Aynı widget'a KISA SÜRE İÇİNDE birden fazla kez ışıma tetiklenebilir (ör.
// hem bildirim zilinden hem WelcomeHeader özet şeridinden aynı kategoriye
// art arda tıklanması — ikisi de aynı anchor'ı hedefler). Bu elemanlar
// üzerinde bekleyen zamanlayıcıları izlemeden yeni bir çağrı başlatmak,
// ESKİ çağrının "orijinal" diye sakladığı değerin aslında YENİ çağrının
// glow'unu yakalamasına yol açar — sonuçta ışıma kaldırılırken önceki
// glow rengine "geri dönülür" ve halka KALICI OLARAK takılı kalır. Bunun
// yerine: aynı elemana yeni bir çağrı gelince önceki bekleyen
// zamanlayıcılar iptal edilir ve her zaman bilinen/sabit bir "orijinal"
// değere (boş string — bu anchor sarmalayıcı div'lerinin kendi satır içi
// box-shadow'u hiç olmaz) dönülür.
const highlightTimers = new WeakMap<HTMLElement, number[]>();

/**
 * Hedef elemanın etrafında kısa süreli bir "ışıma" halkası belirip söner —
 * kullanıcıya TAM OLARAK hangi widget'a kaydırıldığını gösterir (bkz.
 * NotificationBell.tsx'teki orijinal desen — buraya taşındı ki hem
 * NotificationBell hem de WelcomeHeader'daki yeni özet şeridi aynı efekti
 * tekrar yazmadan paylaşabilsin).
 */
export const highlightElement = (target: HTMLElement, accentColor: string): void => {
    const pending = highlightTimers.get(target);
    if (pending) {
        pending.forEach((id) => window.clearTimeout(id));
    }

    target.style.transition = 'box-shadow 0.3s ease';
    target.style.borderRadius = target.style.borderRadius || '22px';
    target.style.boxShadow = `0 0 0 3px ${accentColor}, 0 8px 24px ${accentColor}55`;

    const revertTimeout = window.setTimeout(() => {
        target.style.boxShadow = '';
        const cleanupTimeout = window.setTimeout(() => {
            target.style.transition = '';
            target.style.borderRadius = '';
            highlightTimers.delete(target);
        }, 320);
        highlightTimers.set(target, [cleanupTimeout]);
    }, 1400);

    highlightTimers.set(target, [revertTimeout]);
};

/** Bir anchor id'sine yumuşakça kaydırır ve ışıma efektini uygular — id bulunamazsa sessizce hiçbir şey yapmaz. */
export const scrollToAnchorWithHighlight = (anchorId: string, accentColor: string): void => {
    const target = document.getElementById(anchorId);
    if (!target) {
        return;
    }
    smoothScrollToElement(target);
    highlightElement(target, accentColor);
};
