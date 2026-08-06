import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IRecentFileItem {
    /** Sıralama/anahtar için — Path zaten benzersiz. */
    id: string;
    title: string;
    /** Dosyayı açmak/indirmek için tam URL. */
    path: string;
    extension: string;
    /** Dosyanın yaşadığı site — kullanıcıya "nerede" olduğunu hatırlatır (arama tüm kiracıyı tarıyor). */
    siteTitle: string;
    modifiedLabel: string;
}

interface ISearchCell {
    Key: string;
    Value: string;
}

interface ISearchRow {
    Cells: ISearchCell[];
}

/**
 * SharePoint Search REST API'sinin (_api/search/query) yanıtı OData'nın
 * ALIŞILDIK { value: [...] } biçiminde DEĞİL — her satır bir "Cells"
 * anahtar/değer çifti dizisi olarak geliyor (CSOM'dan miras kalan eski bir
 * biçim). Bu yardımcı, bir satırı düz bir { PropertyName: value } nesnesine
 * çeviriyor ki geri kalan kod normal bir obje gibi okuyabilsin.
 */
const rowToObject = (row: ISearchRow): Record<string, string> =>
    row.Cells.reduce((acc, cell) => {
        acc[cell.Key] = cell.Value;
        return acc;
    }, {} as Record<string, string>);

const getExtension = (name: string): string => {
    const idx = name.lastIndexOf('.');
    return idx === -1 ? '' : name.substring(idx + 1).toLowerCase();
};

/**
 * Oturum sahibinin YAZARI YA DA SON DÜZENLEYENİ olduğu belgeleri, en son
 * değiştirilenden en eskiye doğru getirir — "son/sık kullandığım dosyalar"
 * alanının veri kaynağı.
 *
 * NEDEN Microsoft Graph /me/insights/used veya /me/drive/recent DEĞİL:
 * ikisi de Microsoft Learn'e göre kullanımdan kaldırıldı ve Kasım 2026'dan
 * itibaren veri döndürmeyi durduracak — bunun üzerine yeni bir özellik inşa
 * etmek birkaç ay içinde bozulurdu. Bunun yerine, KULLANIMDAN KALDIRILMAMIŞ
 * klasik SharePoint Search REST API'si kullanılıyor: aynı context.spHttpClient
 * ile çalışıyor, YENİ BİR GRAPH İZNİ/YÖNETİCİ ONAYI GEREKTİRMİYOR (bkz.
 * package-solution.json'daki webApiPermissionRequests notu) ve kiracı
 * genelinde (kullanıcının erişimi olan HER site) izin-budanmış (permission-
 * trimmed) sonuç döndürüyor — tek bir kütüphaneyle sınırlı kalmıyor.
 *
 * Semantik fark: bu "AÇTIĞIM" dosyalar değil, "YAZARI/SON DÜZENLEYENİ BEN
 * OLAN" dosyalar — ama bir iş portalı için "üzerinde çalıştığım belgeler"
 * anlamına gelen makul ve dürüst bir karşılık.
 */
export const getMyRecentFiles = async (context: WebPartContext, top = 6): Promise<IRecentFileItem[]> => {
    const loginName = context.pageContext.user.loginName;
    const escapedLoginName = loginName.replace(/"/g, '\\"');
    const queryText = `IsDocument:1 (AuthorOWSUSER:"${escapedLoginName}" OR EditorOWSUSER:"${escapedLoginName}")`;
    const selectProperties = 'Title,Path,FileExtension,LastModifiedTime,SiteTitle';

    const endpoint =
        `${context.pageContext.web.absoluteUrl}/_api/search/query` +
        `?querytext='${encodeURIComponent(queryText)}'` +
        `&selectproperties='${encodeURIComponent(selectProperties)}'` +
        `&sortlist='LastModifiedTime:descending'` +
        `&trimduplicates=true&rowlimit=${top}`;

    // ÖNCEKİ HATA (canlı ortamda doğrulandı): SPHttpClient.configurations.v1
    // GET isteklerine varsayılan olarak "OData-Version: 4.0" ve
    // "Accept: application/json;odata.metadata=minimal" başlıklarını EKLİYOR
    // (bkz. @microsoft/sp-http-base SPHttpClientHelper.applyDefaultHeaders —
    // v1 yapılandırması defaultODataVersion:v4 kullanıyor). Klasik Search
    // REST API (_api/search/query) bu OData v4 sözleşmesini TANIMIYOR ve bu
    // başlıklarla her zaman genel/anlamsız bir "-1, UnknownError" 500 hatası
    // döndürüyor — sorgunun kendisiyle hiçbir ilgisi yok. Bu iki başlığı
    // AÇIKÇA v3 değerleriyle geçmek (SPHttpClient yalnızca eksik başlıkları
    // varsayılanla dolduruyor) sorunu tamamen çözüyor.
    const response: SPHttpClientResponse = await context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1, {
        headers: {
            Accept: 'application/json;odata=nometadata',
            'OData-Version': '3.0'
        }
    });
    if (!response.ok) {
        throw new Error(`Arama isteği başarısız oldu, HTTP ${response.status}`);
    }

    const body = await response.json();
    const rows: ISearchRow[] = body?.PrimaryQueryResult?.RelevantResults?.Table?.Rows ?? [];

    return rows.map((row) => {
        const props = rowToObject(row);
        const name = props.Title || props.Path?.split('/').pop() || 'Adsız dosya';
        return {
            id: props.Path,
            title: name,
            path: props.Path,
            extension: getExtension(props.Path || name),
            siteTitle: props.SiteTitle || '',
            modifiedLabel: props.LastModifiedTime
                ? new Date(props.LastModifiedTime).toLocaleDateString('tr-TR', { day: 'numeric', month: 'short', year: 'numeric' })
                : ''
        };
    });
};
