import * as React from 'react';
import {
    IconButton,
    Icon,
    Spinner,
    SpinnerSize,
    MessageBar,
    MessageBarType,
    Text,
    useTheme,
    mergeStyleSets
} from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import WidgetCard from '../WidgetCard';
import DetailModal from '../DetailModal';
import { DATA_UNAVAILABLE_MESSAGE } from '../../constants';

export interface IRequiredDocumentsProps {
    context: WebPartContext;
}

interface IDocCategory {
    label: string;
    /** "Shared Documents" kütüphanesinin altındaki gerçek alt klasör adı. */
    folderName: string;
    iconName: string;
    accentColor: string;
}

interface IDocFile {
    name: string;
    serverRelativeUrl: string;
    extension: string;
    sizeLabel: string;
    modifiedLabel: string;
}

interface ISPFileEntry {
    Name: string;
    ServerRelativeUrl: string;
    Length: string;
    TimeLastModified: string;
}

// 10 kategori — SharePoint'teki "Shared Documents" (görünen adıyla
// "Belgeler") kütüphanesinin altındaki gerçek alt klasör adlarıyla birebir
// eşleşmeli. ÖNEMLİ: "Belgeler" SADECE görünen (display) addır — kütüphanenin
// GERÇEK/dahili (system) adı "Shared Documents"tır. İlk 5'i mevcut/onaylı
// klasörler; son 5'i her departmanı ilgilendiren genel alanlar için eklendi —
// bu 5 klasörün de "Shared Documents" altında AYNI adlarla oluşturulması
// gerekir, yoksa widget boş dosya listesiyle açılır (konsola uyarı basar).
//
// Renkler BİLİNÇLİ OLARAK tek bir aile içinde tutuluyor: hepsi orta
// doygunlukta/yumuşatılmış tonlar (parlak/neon yok) ve renk çemberinde
// birbirinden düzenli aralıklarla ayrılıyor — bitişik iki kategori asla
// aynı ton grubunda olmuyor. Önceki sürümde Bilgi İşlem'in canlı camgöbeği
// ve Üretim'in parlak macentası diğer 8 yumuşak tonun arasında "bağırıyordu";
// ikisi de burada aynı aileye çekildi.
const CATEGORIES: IDocCategory[] = [
    { label: 'Franchise', folderName: 'Franchise', iconName: 'CityNext', accentColor: '#5c8fc4' },
    { label: 'Kalite', folderName: 'Kalite', iconName: 'Ribbon', accentColor: '#6fa87a' },
    { label: 'İSG', folderName: 'ISG', iconName: 'Shield', accentColor: '#c9635c' },
    { label: 'Satın Alma', folderName: 'SatinAlma', iconName: 'ShoppingCart', accentColor: '#b8945a' },
    { label: 'Marka', folderName: 'Marka', iconName: 'Tag', accentColor: '#9a8bc7' },
    { label: 'İnsan Kaynakları', folderName: 'InsanKaynaklari', iconName: 'People', accentColor: '#4f9d90' },
    { label: 'Finans', folderName: 'Finans', iconName: 'Calculator', accentColor: '#4a6fa5' },
    { label: 'Bilgi İşlem', folderName: 'BilgiIslem', iconName: 'Server', accentColor: '#5b90a8' },
    { label: 'Pazarlama', folderName: 'Pazarlama', iconName: 'BullseyeTarget', accentColor: '#c17ba0' },
    { label: 'Üretim', folderName: 'Uretim', iconName: 'Manufacturing', accentColor: '#b8763f' }
];

const formatFileSize = (bytes: number): string => {
    if (!bytes) {
        return '0 KB';
    }
    if (bytes < 1024 * 1024) {
        return `${Math.max(1, Math.round(bytes / 1024))} KB`;
    }
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
};

const getExtension = (name: string): string => {
    const idx = name.lastIndexOf('.');
    return idx === -1 ? '' : name.substring(idx + 1).toLowerCase();
};

const getFileIconName = (extension: string): string => {
    switch (extension) {
        case 'pdf': return 'PDF';
        case 'doc':
        case 'docx': return 'WordDocument';
        case 'xls':
        case 'xlsx': return 'ExcelDocument';
        case 'ppt':
        case 'pptx': return 'PowerPointDocument';
        default: return 'TextDocument';
    }
};

const getFileIconColor = (extension: string): string => {
    switch (extension) {
        case 'pdf': return '#d83b01';
        case 'doc':
        case 'docx': return '#185ABD';
        case 'xls':
        case 'xlsx': return '#107c10';
        case 'ppt':
        case 'pptx': return '#b7472a';
        default: return '#a99d8b';
    }
};

// Dosyaya (satıra) tıklanınca AÇAR — SharePoint dosya türüne göre kendi
// önizleme/görüntüleyicisine (Office Online, PDF görüntüleyici vb.) yönlendirir.
const viewFile = (serverRelativeUrl: string): void => {
    window.open(encodeURI(serverRelativeUrl), '_blank', 'noopener,noreferrer');
};

// İndir butonuna basılınca ise "?download=1" ile önizlemeyi atlayıp
// doğrudan indirmeyi tetikler.
const downloadFile = (serverRelativeUrl: string): void => {
    window.open(`${encodeURI(serverRelativeUrl)}?download=1`, '_blank', 'noopener,noreferrer');
};

type LoadState = 'loading' | 'loaded' | 'error';

/**
 * "Gerekli Dosyalar" widget'ı — bir kategoriye tıklandığında SharePoint'e
 * yönlendirmek yerine, o klasörün içindeki dosyaları GERÇEK ZAMANLI çekip
 * ekranın ortasında açılan bir pop-up'ta (DetailModal) listeler. Kullanıcı
 * SharePoint'e hiç gitmeden dosyayı doğrudan popup içinden indirebilir.
 */
const RequiredDocuments: React.FunctionComponent<IRequiredDocumentsProps> = (props) => {
    const { context } = props;
    const theme = useTheme();

    const [selectedCategory, setSelectedCategory] = React.useState<IDocCategory | undefined>(undefined);
    const [filesState, setFilesState] = React.useState<LoadState>('loading');
    const [files, setFiles] = React.useState<IDocFile[]>([]);

    React.useEffect(() => {
        if (!selectedCategory) {
            return;
        }
        let isMounted = true;
        setFilesState('loading');
        setFiles([]);

        const fetchFiles = async (): Promise<void> => {
            try {
                const webUrl = context.pageContext.web.absoluteUrl;
                const webServerRelativeUrl = context.pageContext.web.serverRelativeUrl;
                const folderServerRelativeUrl = `${webServerRelativeUrl}/Shared Documents/${selectedCategory.folderName}`;
                const endpoint =
                    `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURI(folderServerRelativeUrl)}')/Files` +
                    '?$select=Name,ServerRelativeUrl,Length,TimeLastModified';

                const response = await context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
                if (!response.ok) {
                    console.error('[RequiredDocuments] Dosyalar alınamadı, HTTP', response.status);
                    if (isMounted) {
                        setFilesState('error');
                    }
                    return;
                }

                const body: { value: ISPFileEntry[] } = await response.json();
                if (isMounted) {
                    setFiles(
                        (body.value ?? []).map((f) => ({
                            name: f.Name,
                            serverRelativeUrl: f.ServerRelativeUrl,
                            extension: getExtension(f.Name),
                            sizeLabel: formatFileSize(parseInt(f.Length, 10) || 0),
                            modifiedLabel: new Date(f.TimeLastModified).toLocaleDateString('tr-TR', {
                                day: 'numeric',
                                month: 'short',
                                year: 'numeric'
                            })
                        }))
                    );
                    setFilesState('loaded');
                }
            } catch (error) {
                console.error('[RequiredDocuments] Dosyalar alınırken hata oluştu:', error);
                if (isMounted) {
                    setFilesState('error');
                }
            }
        };

        fetchFiles().catch(() => { /* fetchFiles kendi içinde hatayı yönetir */ });

        return () => {
            isMounted = false;
        };
    }, [selectedCategory, context]);

    const styles = mergeStyleSets({
        // 5'ten 10 kategoriye çıkınca sabit 2 sütun widget'ı gereksiz uzatıyordu —
        // auto-fit ile kart genişledikçe 3-4 sütuna kadar esner, dar ekranda 2'ye düşer.
        grid: {
            display: 'grid',
            gridTemplateColumns: 'repeat(auto-fit, minmax(112px, 1fr))',
            gap: 12
        },
        // Eski sürüm her kategoriyi DÜZ RENKLE dolu, tam genişlikte bir buton
        // olarak gösteriyordu — 10 kategoriyle bu, yan yana 10 farklı canlı
        // rengin çakıştığı gürültülü bir görünüme yol açıyordu. Artık renk
        // sadece küçük bir ikon rozetini boyuyor (QuickLinksPanel'deki dille
        // aynı), kartın geri kalanı nötr — palet ne kadar geniş olursa olsun
        // sakin duruyor.
        // NOT: "gap" flex özelliği burada BİLİNÇLİ OLARAK KULLANILMIYOR — bu
        // sayfanın render edildiği (kurumsal/eski) tarayıcı ortamında flex "gap"
        // desteklenmiyor. Bunun yerine categoryIconWrap kendi marginBottom'ını
        // taşıyor (categoryLabel her zaman son çocuk olduğu için margin almıyor).
        categoryTile: {
            position: 'relative',
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            justifyContent: 'center',
            width: '100%',
            padding: '16px 10px',
            borderRadius: 14,
            border: `1px solid ${theme.palette.neutralLight}`,
            background: theme.palette.neutralLighterAlt,
            cursor: 'pointer',
            textAlign: 'center',
            font: 'inherit',
            transition: 'transform 0.2s ease, box-shadow 0.2s ease',
            boxShadow: '0 1px 3px rgba(0,0,0,0.06)',
            selectors: {
                ':hover': {
                    transform: 'translateY(-4px)'
                }
            }
        },
        categoryIconWrap: {
            width: 40,
            height: 40,
            borderRadius: 11,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            flexShrink: 0,
            marginBottom: 10
        },
        categoryIcon: {
            fontSize: 18,
            color: '#ffffff'
        },
        categoryLabel: {
            fontSize: 12,
            fontWeight: 700,
            lineHeight: '15px',
            color: theme.semanticColors.bodyText
        },
        // NOT: "gap" flex özelliği burada da BİLİNÇLİ OLARAK KULLANILMIYOR — bkz.
        // yukarıdaki categoryTile açıklaması. fileIcon ve fileTextGroup kendi
        // marginRight'larını taşıyor, son çocuk (indir butonu) dokunulmadan kalıyor.
        fileRow: {
            display: 'flex',
            alignItems: 'center',
            width: '100%',
            padding: '10px 8px',
            borderRadius: 10,
            border: 'none',
            background: 'transparent',
            textAlign: 'left',
            font: 'inherit',
            cursor: 'pointer',
            transition: 'background 0.15s ease',
            selectors: {
                ':hover': {
                    background: theme.palette.neutralLighterAlt
                }
            }
        },
        fileIcon: {
            fontSize: 22,
            flexShrink: 0,
            marginRight: 12
        },
        // DetailModal'ın gövdesi lineHeight:1.6 uyguluyor — bu satırlar
        // kendi SABİT lineHeight'lerini açıkça tanımlamazsa, o miras alınan
        // değerle birlikte iki satır birbirinin üzerine binip metnin "kesilmiş"
        // görünmesine yol açabiliyordu. Açık flex-column + sabit lineHeight
        // bunu yapısal olarak imkansız hale getiriyor.
        fileTextGroup: {
            display: 'flex',
            flexDirection: 'column',
            rowGap: 4,
            flexGrow: 1,
            minWidth: 0,
            marginRight: 12
        },
        fileName: {
            display: 'block',
            fontSize: 13,
            fontWeight: 600,
            lineHeight: '18px',
            color: theme.semanticColors.bodyText,
            overflow: 'hidden',
            textOverflow: 'ellipsis',
            whiteSpace: 'nowrap'
        },
        fileMeta: {
            display: 'block',
            fontSize: 11,
            lineHeight: '15px',
            color: theme.semanticColors.bodySubtext
        },
        emptyHint: {
            fontSize: 12,
            color: theme.semanticColors.bodySubtext,
            padding: '12px 0'
        }
    });

    return (
        <WidgetCard title="Gerekli Dosyalar" subtitle="Sık aranan doküman ve formlar" iconName="DocumentSet" accentColor="#8a97a0">
            <div className={styles.grid}>
                {CATEGORIES.map((cat) => (
                    <button
                        key={cat.folderName}
                        type="button"
                        className={styles.categoryTile}
                        onClick={() => setSelectedCategory(cat)}
                        title={`${cat.label} dosyalarını görüntüle`}
                        onMouseEnter={(e) => {
                            e.currentTarget.style.boxShadow = `0 10px 22px rgba(0,0,0,0.28), 0 0 0 1px ${cat.accentColor}80`;
                        }}
                        onMouseLeave={(e) => {
                            e.currentTarget.style.boxShadow = '0 1px 3px rgba(0,0,0,0.06)';
                        }}
                    >
                        <div
                            className={styles.categoryIconWrap}
                            style={{
                                background: `linear-gradient(135deg, ${cat.accentColor} 0%, ${cat.accentColor}cc 100%)`,
                                boxShadow: `0 4px 12px ${cat.accentColor}55`
                            }}
                        >
                            <Icon iconName={cat.iconName} className={styles.categoryIcon} />
                        </div>
                        <span className={styles.categoryLabel}>{cat.label}</span>
                    </button>
                ))}
            </div>

            <DetailModal
                isOpen={!!selectedCategory}
                title={selectedCategory ? `${selectedCategory.label} Dosyaları` : ''}
                onDismiss={() => setSelectedCategory(undefined)}
            >
                {filesState === 'loading' && <Spinner size={SpinnerSize.medium} label="Dosyalar yükleniyor..." />}
                {filesState === 'error' && <MessageBar messageBarType={MessageBarType.error}>{DATA_UNAVAILABLE_MESSAGE}</MessageBar>}
                {filesState === 'loaded' && files.length === 0 && (
                    <Text className={styles.emptyHint}>Bu klasörde dosya bulunmuyor.</Text>
                )}
                {filesState === 'loaded' && files.map((file) => (
                    <button
                        key={file.serverRelativeUrl}
                        type="button"
                        className={styles.fileRow}
                        onClick={() => viewFile(file.serverRelativeUrl)}
                        title={`${file.name} aç`}
                    >
                        <Icon
                            iconName={getFileIconName(file.extension)}
                            className={styles.fileIcon}
                            style={{ color: getFileIconColor(file.extension) }}
                        />
                        <div className={styles.fileTextGroup}>
                            <div className={styles.fileName}>{file.name}</div>
                            <div className={styles.fileMeta}>{file.modifiedLabel} · {file.sizeLabel}</div>
                        </div>
                        <IconButton
                            iconProps={{ iconName: 'Download' }}
                            ariaLabel={`${file.name} indir`}
                            title="İndir"
                            onClick={(e) => {
                                // Satırın kendi tıklama (görüntüleme) davranışını
                                // tetiklemesin diye olay yukarı taşınmıyor.
                                e.stopPropagation();
                                downloadFile(file.serverRelativeUrl);
                            }}
                        />
                    </button>
                ))}
            </DetailModal>
        </WidgetCard>
    );
};

export default RequiredDocuments;
