import * as React from 'react';
import {
    Spinner, SpinnerSize, MessageBar, MessageBarType, Icon, IconButton, PrimaryButton,
    useTheme, mergeStyleSets
} from '@fluentui/react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import WidgetCard from '../WidgetCard';
import DetailModal from '../DetailModal';
import { DATA_UNAVAILABLE_MESSAGE, SUBMIT_UNAVAILABLE_MESSAGE, IS_ADMIN_MOCK } from '../../constants';

export interface IISGCalendarWidgetProps {
    context: WebPartContext;
}

// İSG personeli her ay bu klasöre widget üzerindeki "+" ile yeni takvimi
// yükler — SharePoint doküman kütüphanesiyle hiç uğraşmaz. Klasör yoksa
// otomatik oluşturulur; yeni bir dosya yüklendiğinde eskisi otomatik silinir
// (klasörde her zaman TEK bir dosya kalır).
const CALENDAR_FOLDER_NAME = 'ISGTakvimi';
const IMAGE_EXTENSIONS = ['jpg', 'jpeg', 'png', 'gif', 'webp'];
const PDF_EXTENSIONS = ['pdf'];
const ALLOWED_EXTENSIONS = [...IMAGE_EXTENSIONS, ...PDF_EXTENSIONS];

interface ISPFileEntry {
    Name: string;
    ServerRelativeUrl: string;
    TimeLastModified: string;
}

interface ICalendarFile {
    name: string;
    fileUrl: string;
    isPdf: boolean;
}

type LoadState = 'loading' | 'loaded' | 'empty' | 'error';
type UploadState = 'idle' | 'sending' | 'error';

const getExtension = (name: string): string => {
    const idx = name.lastIndexOf('.');
    return idx === -1 ? '' : name.substring(idx + 1).toLowerCase();
};

const extractErrorDetail = async (response: SPHttpClientResponse): Promise<string> => {
    try {
        const raw = await response.text();
        const parsed = JSON.parse(raw) as { 'odata.error'?: { message?: { value?: string } } };
        return parsed['odata.error']?.message?.value ?? `HTTP ${response.status}`;
    } catch {
        return `HTTP ${response.status}`;
    }
};

const ISGCalendarWidget: React.FunctionComponent<IISGCalendarWidgetProps> = ({ context }) => {
    const theme = useTheme();
    const [state, setState] = React.useState<LoadState>('loading');
    const [file, setFile] = React.useState<ICalendarFile | undefined>(undefined);
    const [isPreviewOpen, setIsPreviewOpen] = React.useState(false);

    const [isUploadOpen, setIsUploadOpen] = React.useState(false);
    const [newFile, setNewFile] = React.useState<File | undefined>(undefined);
    const [uploadState, setUploadState] = React.useState<UploadState>('idle');
    const [uploadError, setUploadError] = React.useState<string | undefined>(undefined);

    const folderServerRelativeUrl = `${context.pageContext.web.serverRelativeUrl}/Shared Documents/${CALENDAR_FOLDER_NAME}`;

    const loadCalendar = React.useCallback((onlyIfMounted?: () => boolean): void => {
        setState('loading');

        const run = async (): Promise<void> => {
            try {
                const webUrl = context.pageContext.web.absoluteUrl;
                const endpoint =
                    `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURI(folderServerRelativeUrl)}')/Files` +
                    '?$select=Name,ServerRelativeUrl,TimeLastModified';

                const response = await context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);

                if (!response.ok) {
                    // Klasör henüz oluşturulmamışsa SharePoint 404 döner — bu bir hata
                    // değil, sadece içerik editörünün henüz dosya yüklemediği anlamına
                    // gelir.
                    if (response.status === 404) {
                        if (!onlyIfMounted || onlyIfMounted()) {
                            setState('empty');
                        }
                        return;
                    }
                    console.error('[ISGCalendarWidget] Klasör okunamadı, HTTP', response.status);
                    if (!onlyIfMounted || onlyIfMounted()) {
                        setState('error');
                    }
                    return;
                }

                const body: { value: ISPFileEntry[] } = await response.json();
                const candidates = (body.value ?? [])
                    .filter((f) => ALLOWED_EXTENSIONS.indexOf(getExtension(f.Name)) > -1)
                    .sort((a, b) => new Date(b.TimeLastModified).getTime() - new Date(a.TimeLastModified).getTime());

                if (candidates.length === 0) {
                    if (!onlyIfMounted || onlyIfMounted()) {
                        setState('empty');
                    }
                    return;
                }

                const latest = candidates[0];
                if (!onlyIfMounted || onlyIfMounted()) {
                    setFile({
                        name: latest.Name,
                        fileUrl: latest.ServerRelativeUrl,
                        isPdf: PDF_EXTENSIONS.indexOf(getExtension(latest.Name)) > -1
                    });
                    setState('loaded');
                }
            } catch (error) {
                console.error('[ISGCalendarWidget] İSG takvimi alınırken hata oluştu:', error);
                if (!onlyIfMounted || onlyIfMounted()) {
                    setState('error');
                }
            }
        };

        run().catch(() => { /* run kendi içinde hatayı yönetir */ });
    }, [context, folderServerRelativeUrl]);

    React.useEffect(() => {
        let isMounted = true;
        loadCalendar(() => isMounted);
        return () => {
            isMounted = false;
        };
    }, [loadCalendar]);

    const closeUploadModal = (): void => {
        setIsUploadOpen(false);
        setNewFile(undefined);
        setUploadState('idle');
        setUploadError(undefined);
    };

    const handleUploadSubmit = async (): Promise<void> => {
        if (!newFile) {
            return;
        }
        setUploadState('sending');
        setUploadError(undefined);

        const webUrl = context.pageContext.web.absoluteUrl;

        try {
            // 1) Klasörü garanti et — yoksa oluştur. Zaten varsa SharePoint bir hata
            // döner, bu BEKLENEN bir durumdur ve yutulur.
            try {
                await context.spHttpClient.post(
                    `${webUrl}/_api/web/folders/add('${encodeURIComponent(folderServerRelativeUrl)}')`,
                    SPHttpClient.configurations.v1,
                    { headers: { Accept: 'application/json;odata=nometadata' } }
                );
            } catch {
                // Klasör zaten vardı — sorun değil.
            }

            // 2) Klasördeki mevcut dosyaları listele ve HEPSİNİ geri dönüşüme
            // gönder — "eskisi iptal olsun" isteği: klasörde her zaman tek bir
            // takvim dosyası kalır.
            const listEndpoint =
                `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURI(folderServerRelativeUrl)}')/Files` +
                '?$select=ServerRelativeUrl';
            const listResponse = await context.spHttpClient.get(listEndpoint, SPHttpClient.configurations.v1);
            if (listResponse.ok) {
                const listBody: { value: Array<{ ServerRelativeUrl: string }> } = await listResponse.json();
                await Promise.all(
                    (listBody.value ?? []).map((f) =>
                        context.spHttpClient
                            .post(
                                `${webUrl}/_api/web/GetFileByServerRelativeUrl('${encodeURI(f.ServerRelativeUrl)}')/recycle()`,
                                SPHttpClient.configurations.v1,
                                { headers: { Accept: 'application/json;odata=nometadata' } }
                            )
                            .catch((error) => console.warn('[ISGCalendarWidget] Eski dosya silinemedi:', f.ServerRelativeUrl, error))
                    )
                );
            }

            // 3) Yeni dosyayı yükle.
            const safeName = newFile.name.replace(/[^a-zA-Z0-9._-]/g, '_');
            const uploadEndpoint =
                `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURI(folderServerRelativeUrl)}')` +
                `/Files/add(url='${encodeURIComponent(safeName)}',overwrite=true)`;
            const fileBuffer = await newFile.arrayBuffer();
            const uploadResponse = await context.spHttpClient.post(uploadEndpoint, SPHttpClient.configurations.v1, {
                headers: { Accept: 'application/json;odata=nometadata' },
                body: fileBuffer
            });

            if (!uploadResponse.ok) {
                const detail = await extractErrorDetail(uploadResponse);
                console.error('[ISGCalendarWidget] Takvim yüklenemedi:', detail);
                setUploadState('error');
                setUploadError(detail);
                return;
            }

            closeUploadModal();
            loadCalendar();
        } catch (error) {
            console.error('[ISGCalendarWidget] Takvim yüklenirken beklenmeyen hata:', error);
            setUploadState('error');
            setUploadError((error as Error).message);
        }
    };

    const styles = mergeStyleSets({
        preview: {
            display: 'block',
            width: '100%',
            border: 'none',
            padding: 0,
            background: 'transparent',
            cursor: 'pointer',
            textAlign: 'left'
        },
        image: {
            width: '100%',
            borderRadius: 12,
            display: 'block',
            boxShadow: '0 4px 14px rgba(0,0,0,0.14)'
        },
        docCard: {
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            textAlign: 'center',
            padding: '28px 12px',
            borderRadius: 12,
            background: theme.palette.neutralLighterAlt
        },
        docIcon: {
            fontSize: 40,
            color: '#d83b01',
            marginBottom: 12
        },
        docName: {
            fontSize: 13,
            fontWeight: 600,
            color: theme.semanticColors.bodyText,
            wordBreak: 'break-word'
        },
        docHint: {
            fontSize: 11,
            color: theme.semanticColors.bodySubtext,
            marginTop: 4
        },
        emptyState: {
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            textAlign: 'center',
            padding: '28px 12px',
            color: theme.semanticColors.bodySubtext
        },
        emptyIcon: {
            fontSize: 28,
            marginBottom: 10,
            color: theme.palette.neutralTertiary
        },
        emptyText: {
            fontSize: 13,
            lineHeight: '18px'
        },
        modalImage: {
            width: '100%',
            display: 'block',
            borderRadius: 8
        },
        modalPdfFrame: {
            width: '100%',
            height: 520,
            border: 'none',
            borderRadius: 8
        },
        modalLink: {
            display: 'inline-block',
            marginTop: 14,
            fontSize: 13
        },
        dropzoneWrapper: {
            display: 'block',
            cursor: 'pointer'
        },
        imageDropzone: {
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            justifyContent: 'center',
            padding: '24px 20px',
            textAlign: 'center',
            background: 'rgba(248,250,252,0.8)',
            border: '2px dashed #CBD5E1',
            borderRadius: 12,
            transition: 'border-color 0.2s ease',
            selectors: {
                ':hover': { borderColor: '#3B82F6' }
            }
        },
        imageDropzoneIconWrap: {
            width: 40,
            height: 40,
            borderRadius: '50%',
            background: 'rgba(59,130,246,0.10)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            marginBottom: 12
        },
        imageDropzoneIcon: {
            color: '#3B82F6',
            fontSize: 20
        },
        imageDropzoneLabel: {
            fontSize: 13,
            fontWeight: 600,
            color: theme.semanticColors.bodyText,
            lineHeight: '18px',
            marginBottom: 4
        },
        imageDropzoneHint: {
            fontSize: 11,
            color: theme.semanticColors.bodySubtext,
            lineHeight: '16px'
        },
        imageDropzoneFileName: {
            fontSize: 12,
            color: theme.semanticColors.bodySubtext,
            overflow: 'hidden',
            textOverflow: 'ellipsis',
            whiteSpace: 'nowrap',
            maxWidth: '100%',
            marginTop: 8
        },
        hiddenFileInput: {
            display: 'none'
        },
        formErrorBar: {
            marginBottom: 16
        },
        formActions: {
            marginTop: 20,
            display: 'flex',
            justifyContent: 'flex-end'
        }
    });

    return (
        <WidgetCard
            title="İSG Takvimi"
            subtitle="Güncel iş sağlığı ve güvenliği takvimi"
            iconName="Shield"
            accentColor="#c9635c"
            headerAction={IS_ADMIN_MOCK && (
                <IconButton
                    iconProps={{ iconName: 'Add' }}
                    ariaLabel="Yeni takvim yükle"
                    title="Yeni takvim yükle"
                    onClick={() => setIsUploadOpen(true)}
                />
            )}
        >
            {state === 'loading' && <Spinner size={SpinnerSize.medium} label="Takvim yükleniyor..." />}
            {state === 'error' && <MessageBar messageBarType={MessageBarType.error}>{DATA_UNAVAILABLE_MESSAGE}</MessageBar>}
            {state === 'empty' && (
                <div className={styles.emptyState}>
                    <Icon iconName="CalendarWeek" className={styles.emptyIcon} />
                    <div className={styles.emptyText}>
                        Henüz bir İSG takvimi yüklenmedi.<br />
                        Sağ üstteki "+" ile görsel veya PDF olarak yükleyebilirsiniz.
                    </div>
                </div>
            )}
            {state === 'loaded' && file && (
                <button type="button" className={styles.preview} onClick={() => setIsPreviewOpen(true)}>
                    {file.isPdf ? (
                        <div className={styles.docCard}>
                            <Icon iconName="PDF" className={styles.docIcon} />
                            <div className={styles.docName}>{file.name}</div>
                            <div className={styles.docHint}>Büyük görüntülemek için tıklayın</div>
                        </div>
                    ) : (
                        <img src={file.fileUrl} alt={file.name} title="Büyük görüntülemek için tıklayın" className={styles.image} />
                    )}
                </button>
            )}

            <DetailModal isOpen={isPreviewOpen} title="İSG Takvimi" onDismiss={() => setIsPreviewOpen(false)}>
                {file && file.isPdf && (
                    <>
                        <iframe src={encodeURI(file.fileUrl)} title={file.name} className={styles.modalPdfFrame} />
                        <a
                            href={encodeURI(file.fileUrl)}
                            target="_blank"
                            rel="noopener noreferrer"
                            className={styles.modalLink}
                        >
                            Yeni sekmede aç / indir
                        </a>
                    </>
                )}
                {file && !file.isPdf && <img src={file.fileUrl} alt={file.name} className={styles.modalImage} />}
            </DetailModal>

            <DetailModal isOpen={isUploadOpen} title="Yeni İSG Takvimi Yükle" onDismiss={closeUploadModal}>
                <div>
                    {uploadState === 'error' && (
                        <MessageBar messageBarType={MessageBarType.error} className={styles.formErrorBar}>
                            {uploadError ?? SUBMIT_UNAVAILABLE_MESSAGE}
                        </MessageBar>
                    )}
                    <label className={styles.dropzoneWrapper} htmlFor="isg-calendar-file-input">
                        <div className={styles.imageDropzone}>
                            <div className={styles.imageDropzoneIconWrap}>
                                <Icon iconName="CloudUpload" className={styles.imageDropzoneIcon} />
                            </div>
                            <span className={styles.imageDropzoneLabel}>Görsel veya PDF Seç</span>
                            <span className={styles.imageDropzoneHint}>Yüklenince önceki takvim otomatik silinir</span>
                            {newFile && <span className={styles.imageDropzoneFileName}>{newFile.name}</span>}
                        </div>
                        <input
                            id="isg-calendar-file-input"
                            type="file"
                            accept="image/*,.pdf"
                            className={styles.hiddenFileInput}
                            disabled={uploadState === 'sending'}
                            onChange={(e) => setNewFile(e.target.files?.[0])}
                        />
                    </label>
                    <div className={styles.formActions}>
                        {uploadState === 'sending' ? (
                            <Spinner size={SpinnerSize.small} label="Yükleniyor..." />
                        ) : (
                            <PrimaryButton
                                text="Kaydet"
                                onClick={() => { handleUploadSubmit().catch(() => { /* handled in handleUploadSubmit */ }); }}
                                disabled={!newFile}
                            />
                        )}
                    </div>
                </div>
            </DetailModal>
        </WidgetCard>
    );
};

export default ISGCalendarWidget;
