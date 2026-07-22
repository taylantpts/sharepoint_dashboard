// SharePoint zengin metin (rich text) alanlarından gelen HTML, ekrana
// dangerouslySetInnerHTML ile basılmadan ÖNCE bu fonksiyondan geçirilmelidir.
// İçerik normalde site içeriği yöneticileri tarafından girilir, ama OWASP
// "defence in depth" ilkesi gereği İSTEMCİ TARAFINDA da temizleniyoruz —
// tehlikeli etiketler (script/iframe/object/embed/style/form/base/link/meta)
// ve tehlikeli öznitelikler (on* olay dinleyicileri, javascript:/data: URI'ler)
// DOM ağacından tamamen çıkarılıyor.

const DANGEROUS_TAGS = ['script', 'iframe', 'object', 'embed', 'link', 'meta', 'style', 'form', 'base'];

// "javascript:" öneki burada bir string KARŞILAŞTIRMASI için yazılıyor,
// çalıştırılmıyor — eslint'in no-script-url kuralı bunu yanlış pozitif
// olarak işaretliyor, bu yüzden bilinçli olarak devre dışı bırakıldı.
// eslint-disable-next-line no-script-url
const DANGEROUS_URL_PREFIXES = ['javascript:', 'data:text/html'];

const isDangerousUrl = (value: string): boolean => {
    const normalized = value.trim().toLowerCase();
    return DANGEROUS_URL_PREFIXES.some((prefix) => normalized.startsWith(prefix));
};

/**
 * Ham HTML'i tarayıcının kendi DOMParser'ı ile ayrıştırıp tehlikeli
 * etiket/öznitelikleri temizler, güvenli HTML string'i döner. Ayrıştırma
 * herhangi bir sebeple başarısız olursa (bozuk içerik vb.) çökmek yerine
 * boş string döner — bileşen asla bozuk/güvensiz içerik göstermez.
 */
export const sanitizeHtml = (html: string): string => {
    if (!html) {
        return '';
    }

    try {
        const parser = new DOMParser();
        const doc = parser.parseFromString(html, 'text/html');

        DANGEROUS_TAGS.forEach((tag) => {
            const elements = doc.body.querySelectorAll(tag);
            elements.forEach((el) => el.parentNode?.removeChild(el));
        });

        const allElements = doc.body.querySelectorAll('*');
        allElements.forEach((el) => {
            // Statik anlık görüntü — kaldırma işlemi sırasında canlı
            // NamedNodeMap üzerinde atlama yapılmasını önler.
            Array.from(el.attributes).forEach((attr) => {
                const name = attr.name.toLowerCase();
                const value = attr.value;
                if (name.startsWith('on')) {
                    el.removeAttribute(attr.name);
                    return;
                }
                if ((name === 'href' || name === 'src' || name === 'action') && isDangerousUrl(value)) {
                    el.removeAttribute(attr.name);
                }
            });
        });

        return doc.body.innerHTML;
    } catch (error) {
        console.error('[sanitizeHtml] İçerik ayrıştırılamadı, güvenlik gereği boş gösteriliyor:', error);
        return '';
    }
};
