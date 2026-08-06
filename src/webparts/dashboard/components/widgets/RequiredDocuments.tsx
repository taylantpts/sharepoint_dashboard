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
import { WebPartContext } from '@microsoft/sp-webpart-base';
import WidgetCard from '../WidgetCard';
import { DATA_UNAVAILABLE_MESSAGE } from '../../constants';
import { getMyRecentFiles, IRecentFileItem } from '../../services/SearchService';

export interface IRequiredDocumentsProps {
    context: WebPartContext;
}

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
const viewFile = (path: string): void => {
    window.open(encodeURI(path), '_blank', 'noopener,noreferrer');
};

// İndir butonuna basılınca ise "?download=1" ile önizlemeyi atlayıp
// doğrudan indirmeyi tetikler.
const downloadFile = (path: string): void => {
    window.open(`${encodeURI(path)}?download=1`, '_blank', 'noopener,noreferrer');
};

type LoadState = 'loading' | 'loaded' | 'error';

/**
 * "Gerekli Dosyalar" widget'ı — ÖNCEKİ SÜRÜM sabit 10 kategori/klasöre
 * (Franchise, Kalite, İSG...) tıklayıp o klasörün içeriğini listeleten bir
 * tasarımdı; kullanıcı geri bildirimiyle kaldırıldı çünkü çoğu klasör
 * SharePoint'te hiç oluşturulmamıştı ve widget "işlevsiz" hissettiriyordu.
 * Yerine, kullanıcının YAZARI YA DA SON DÜZENLEYENİ olduğu belgeleri en
 * yeniden eskiye listeleyen bir "Dosyalarım" akışı geldi (bkz.
 * SearchService.ts — Microsoft Graph'ın "son kullanılanlar" API'leri
 * kullanımdan kaldırıldığı için SharePoint Search REST API'si kullanılıyor).
 */
const RequiredDocuments: React.FunctionComponent<IRequiredDocumentsProps> = (props) => {
    const { context } = props;
    const theme = useTheme();

    const [state, setState] = React.useState<LoadState>('loading');
    const [files, setFiles] = React.useState<IRecentFileItem[]>([]);

    React.useEffect(() => {
        let isMounted = true;
        setState('loading');

        getMyRecentFiles(context)
            .then((items) => {
                if (isMounted) {
                    setFiles(items);
                    setState('loaded');
                }
            })
            .catch((error: Error) => {
                console.error('[RequiredDocuments] Son dosyalar alınamadı:', error);
                if (isMounted) {
                    setState('error');
                }
            });

        return () => {
            isMounted = false;
        };
    }, [context]);

    const styles = mergeStyleSets({
        // NOT: "gap" flex özelliği burada BİLİNÇLİ OLARAK KULLANILMIYOR — bu
        // sayfanın render edildiği (kurumsal/eski) tarayıcı ortamında flex "gap"
        // desteklenmiyor. fileIcon ve fileTextGroup kendi marginRight'larını
        // taşıyor, son çocuk (indir butonu) dokunulmadan kalıyor.
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
            color: theme.semanticColors.bodySubtext,
            overflow: 'hidden',
            textOverflow: 'ellipsis',
            whiteSpace: 'nowrap'
        },
        emptyHint: {
            fontSize: 12,
            color: theme.semanticColors.bodySubtext,
            // ÖNCEKİ HATA: birimsiz sayı (1.5) bu render ortamında "1.5px" olarak
            // hesaplanıyor (bkz. DetailModal.tsx'teki aynı not) — yüzde string'i kullanılıyor.
            lineHeight: '150%',
            padding: '4px 0'
        }
    });

    return (
        <WidgetCard title="Dosyalarım" subtitle="Son düzenlediğiniz veya oluşturduğunuz belgeler" iconName="History">
            {state === 'loading' && <Spinner size={SpinnerSize.medium} label="Dosyalar yükleniyor..." />}
            {state === 'error' && <MessageBar messageBarType={MessageBarType.error}>{DATA_UNAVAILABLE_MESSAGE}</MessageBar>}
            {state === 'loaded' && files.length === 0 && (
                <Text className={styles.emptyHint}>Henüz üzerinde çalıştığınız bir belge bulunmuyor.</Text>
            )}
            {state === 'loaded' && files.map((file) => (
                <button
                    key={file.id}
                    type="button"
                    className={styles.fileRow}
                    onClick={() => viewFile(file.path)}
                    title={`${file.title} aç`}
                >
                    <Icon
                        iconName={getFileIconName(file.extension)}
                        className={styles.fileIcon}
                        style={{ color: getFileIconColor(file.extension) }}
                    />
                    <div className={styles.fileTextGroup}>
                        <div className={styles.fileName}>{file.title}</div>
                        <div className={styles.fileMeta}>
                            {file.siteTitle ? `${file.siteTitle} · ` : ''}{file.modifiedLabel}
                        </div>
                    </div>
                    <IconButton
                        iconProps={{ iconName: 'Download' }}
                        ariaLabel={`${file.title} indir`}
                        title="İndir"
                        onClick={(e) => {
                            // Satırın kendi tıklama (görüntüleme) davranışını
                            // tetiklemesin diye olay yukarı taşınmıyor.
                            e.stopPropagation();
                            downloadFile(file.path);
                        }}
                    />
                </button>
            ))}
        </WidgetCard>
    );
};

export default RequiredDocuments;
