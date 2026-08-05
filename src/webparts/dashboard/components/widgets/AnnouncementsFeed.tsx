import * as React from 'react';
import {
    Icon, Spinner, SpinnerSize, MessageBar, MessageBarType, Text, IconButton, TextField, PrimaryButton,
    DefaultButton, Dialog, DialogType, DialogFooter,
    useTheme, mergeStyleSets, ITextFieldStyles
} from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import WidgetCard from '../WidgetCard';
import DetailModal from '../DetailModal';
import { getAnnouncements, createAnnouncement, deleteAnnouncement, IAnnouncementItem } from '../../services/SharePointService';
import { DATA_UNAVAILABLE_MESSAGE } from '../../constants';
import { sanitizeHtml } from '../../utils/sanitizeHtml';
import { usePermissions } from '../../hooks/usePermissions';

export interface IAnnouncementsFeedProps {
    context: WebPartContext;
}

type LoadState = 'loading' | 'loaded' | 'error';
type SubmitState = 'idle' | 'sending' | 'error';

// Duyurular ASLA otomatik silinmez (bkz. getAnnouncements) — liste zamanla
// uzayabileceği için sayfa başına sabit bir adet gösterilip altta ileri/geri
// oklarıyla gezinilir. 5: Hava Durumu widget'ıyla aynı satırdaki kartın
// gereksiz boş/şişkin durmaması için Doğum Günleri ile aynı sayfa boyutu.
const PAGE_SIZE = 5;

// Form alanlarının ortak "yumuşak" görünümü — hafif gri zemin, ince/zarif
// kenarlık ve odaklanınca mavi bir halka. Tüm TextField'lar bunu paylaşır.
const inputFieldStyles: Partial<ITextFieldStyles> = {
    // minHeight: her input/textarea'nın kendi sabit bir alt sınırı olsun ki
    // Fluent'in içerik uzunluğuna göre değişen otomatik yüksekliği, formdaki
    // diğer alanlarla dikey hizayı bozmasın.
    // root: formContainer'daki "gap" yerine — bu bileşen formContainer'ın
    // doğrudan çocuğu olduğu için altındaki boşluğu kendisi taşıyor.
    root: { marginBottom: 20 },
    // ÖNCEKİ HATA: kenarlık (#E2E8F0) ve zemin o kadar açık/soluktu ki alan,
    // beyaz modal zemininin üzerinde neredeyse görünmez oluyor, "devre dışı"
    // bir kutu gibi duruyordu — kullanıcıya buranın tıklanabilir bir giriş
    // alanı olduğunu anlatmıyordu. Kenarlık koyulaştırıldı + hafif bir iç
    // gölge (inset shadow) eklendi, gerçek bir "input kutusu" hissi versin.
    fieldGroup: {
        minHeight: 36,
        background: '#F8FAFC',
        border: '1px solid #CBD5E1',
        borderRadius: 12,
        boxShadow: 'inset 0 1px 2px rgba(15,23,42,0.05)',
        selectors: {
            ':after': { borderRadius: 12, border: '2px solid #3B82F6', boxShadow: '0 0 0 4px rgba(59,130,246,0.15)' }
        }
    },
    field: { padding: '12px 16px', fontSize: 14 },
    // ÖNCEKİ HATA: etiket (ör. "Duyuru Başlığı") ile altındaki kutu arasında
    // hiç boşluk YOKTU — Label'ın varsayılan alt boşluğu kutunun artık
    // belirgin bir kenarlığı/gölgesi olmasıyla (bkz. fieldGroup) görsel
    // olarak yetersiz kalıp ikisi iç içe geçmiş gibi görünüyordu.
    subComponentStyles: {
        label: { root: { marginBottom: 8 } }
    }
};

const AnnouncementsFeed: React.FunctionComponent<IAnnouncementsFeedProps> = (props) => {
    const { canManageAnnouncements } = usePermissions(props.context);
    const { context } = props;
    const theme = useTheme();

    const [items, setItems] = React.useState<IAnnouncementItem[]>([]);
    const [state, setState] = React.useState<LoadState>('loading');
    const [selected, setSelected] = React.useState<IAnnouncementItem | undefined>(undefined);
    const [page, setPage] = React.useState(0);

    const [isAddOpen, setIsAddOpen] = React.useState(false);
    const [newTitle, setNewTitle] = React.useState('');
    const [newBody, setNewBody] = React.useState('');
    const [newImage, setNewImage] = React.useState<File | undefined>(undefined);
    const [submitState, setSubmitState] = React.useState<SubmitState>('idle');
    const [submitError, setSubmitError] = React.useState<string | undefined>(undefined);

    const [isDeleteConfirmOpen, setIsDeleteConfirmOpen] = React.useState(false);
    const [deleteState, setDeleteState] = React.useState<SubmitState>('idle');

    // Seçilen görsel için canlı bir önizleme (dosya adı yerine) — obje URL'i
    // sadece bu bileşenin belleğinde yaşar, seçim değişince/kapanınca
    // (temizleme fonksiyonu) serbest bırakılır, bellek sızıntısı olmaz.
    const [newImagePreviewUrl, setNewImagePreviewUrl] = React.useState<string | undefined>(undefined);
    React.useEffect(() => {
        if (!newImage) {
            setNewImagePreviewUrl(undefined);
            return;
        }
        const url = URL.createObjectURL(newImage);
        setNewImagePreviewUrl(url);
        return () => URL.revokeObjectURL(url);
    }, [newImage]);

    const loadAnnouncements = React.useCallback((onlyIfMounted?: () => boolean): void => {
        setState('loading');
        getAnnouncements(context)
            .then((result) => {
                if (!onlyIfMounted || onlyIfMounted()) {
                    setItems(result);
                    setState('loaded');
                }
            })
            .catch((error: Error) => {
                console.error('[AnnouncementsFeed] Duyurular alınamadı:', error);
                if (!onlyIfMounted || onlyIfMounted()) {
                    setState('error');
                }
            });
    }, [context]);

    React.useEffect(() => {
        let isMounted = true;
        loadAnnouncements(() => isMounted);
        return () => {
            isMounted = false;
        };
    }, [loadAnnouncements]);

    const closeAddModal = (): void => {
        setIsAddOpen(false);
        setNewTitle('');
        setNewBody('');
        setNewImage(undefined);
        setSubmitState('idle');
        setSubmitError(undefined);
    };

    /**
     * "Kaydet" — girilen veri SharePoint'teki GERÇEK "Duyurular" listesine
     * yazılır (yerel bir state DEĞİL); başarılı olursa liste sunucudan tekrar
     * çekilir (loadAnnouncements) — böylece ekranda görünen, gerçekten
     * kaydedilmiş veridir, iyimser/sahte bir öğe değil (görsel dahil — kapak
     * görseli de sunucudan geri okunan gerçek ek URL'idir). Tarih/saat
     * kullanıcıdan istenmiyor; SharePoint'in kendi "Created" alanı zaten
     * ekleme anını taşır.
     */
    const handleAddSubmit = async (): Promise<void> => {
        if (!newTitle.trim() || !newBody.trim()) {
            return;
        }
        setSubmitState('sending');
        setSubmitError(undefined);

        try {
            const bodyHtml = newBody.trim().split('\n').join('<br/>');
            const result = await createAnnouncement(context, newTitle.trim(), bodyHtml, newImage);
            if (result.success) {
                closeAddModal();
                setPage(0);
                loadAnnouncements();
            } else {
                setSubmitError(result.errorMessage);
                setSubmitState('error');
            }
        } catch (error) {
            setSubmitError((error as Error).message);
            setSubmitState('error');
        }
    };

    /**
     * Silme, "İK ve İdari İşler Personeli" grubundakiler için detay popup'ında
     * çöp kutusu ikonuyla tetiklenir; asıl kalıcı silme öncesi (geri alınamaz
     * bir işlem olduğu için) her zaman bir onay diyaloğu araya giriyor.
     */
    const handleDeleteConfirm = async (): Promise<void> => {
        if (!selected) {
            return;
        }
        setDeleteState('sending');
        try {
            const result = await deleteAnnouncement(context, selected.id);
            if (result.success) {
                setIsDeleteConfirmOpen(false);
                setSelected(undefined);
                setDeleteState('idle');
                loadAnnouncements();
            } else {
                setDeleteState('error');
            }
        } catch {
            setDeleteState('error');
        }
    };

    const styles = mergeStyleSets({
        // Geniş sütuna oturan, dikey haber-akışı tarzı LİSTE — kart/karo
        // grid'i yerine her duyuru kendi tam genişlikte satırında. İnce alt
        // çizgilerle (row'daki borderBottom) ayrılıyor, son satırda çizgi yok
        // (bkz. lastRow) — salt boşluğa güvenmek yerine net bir liste hissi.
        list: {
            display: 'flex',
            flexDirection: 'column'
        },
        // NOT: "gap" burada kullanılmıyor (flex "gap" bu render ortamında
        // desteklenmiyor) — sağdaki boşluk, ilk çocuk olan iconWrap/thumb'a
        // marginRight olarak taşınıyor (textGroup her zaman son çocuk).
        row: {
            display: 'flex',
            alignItems: 'flex-start',
            width: '100%',
            padding: '14px 4px',
            border: 'none',
            borderBottom: `1px solid ${theme.palette.neutralLighter}`,
            borderRadius: 12,
            background: 'transparent',
            cursor: 'pointer',
            textAlign: 'left',
            font: 'inherit',
            transition: 'background 0.15s ease, box-shadow 0.15s ease, transform 0.15s ease',
            selectors: {
                ':hover': {
                    background: theme.palette.neutralLighterAlt,
                    boxShadow: '0 4px 12px rgba(0,0,0,0.05)',
                    transform: 'translateX(2px)'
                }
            }
        },
        lastRow: {
            borderBottom: 'none'
        },
        // ÖNCEKİ HATA: düz/tek tonlu mavi kare, kartın geri kalanındaki
        // (WidgetCard başlık ikonu, WidgetCard iconWrap) gradyanlı/gölgeli
        // rozetlerin yanında yassı ve "ucuz" duruyordu — artık aynı görsel
        // dili paylaşıyor (gradyan + renkli glow gölgesi).
        iconWrap: {
            width: 40,
            height: 40,
            borderRadius: 11,
            background: `linear-gradient(135deg, ${theme.palette.themePrimary} 0%, ${theme.palette.themeDarkAlt} 100%)`,
            boxShadow: `0 4px 10px ${theme.palette.themePrimary}59`,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            flexShrink: 0,
            marginTop: 2,
            marginRight: 14
        },
        icon: {
            color: '#ffffff',
            fontSize: 17
        },
        // Görseli olan duyurular için: ikon rozeti yerine küçük, köşeleri
        // yuvarlatılmış bir kapak görseli (object-fit: cover — hiçbir zaman
        // taşmaz/bozulmaz), ince bir kenarlık + gölgeyle "yapıştırılmış"
        // değil "yerleştirilmiş" hissi verir.
        thumb: {
            width: 54,
            height: 54,
            borderRadius: 11,
            objectFit: 'cover',
            flexShrink: 0,
            marginTop: 2,
            marginRight: 14,
            border: '1px solid rgba(15,23,42,0.06)',
            boxShadow: '0 2px 6px rgba(15,23,42,0.08)'
        },
        textGroup: {
            minWidth: 0,
            flexGrow: 1,
            paddingTop: 2
        },
        // Premium isteği: başlıklar daha büyük ve kalın (bold).
        // NOT: "-webkit-line-clamp" ile iki satıra sınırlama denendi ama bu
        // render ortamında (flex "gap" desteklemeyen aynı eski/kurumsal
        // tarayıcı) metni neredeyse görünmez hale getirdi — bu yüzden
        // sade/normal satır kaydırmaya geri dönüldü.
        title: {
            fontSize: 15,
            fontWeight: 700,
            color: theme.semanticColors.bodyText,
            lineHeight: 1.35
        },
        // Tarih başlığın altında, silik (muted) bir fontla, küçük bir takvim ikonuyla.
        date: {
            display: 'flex',
            alignItems: 'center',
            fontSize: 12,
            color: theme.semanticColors.bodySubtext,
            marginTop: 6
        },
        dateIcon: {
            fontSize: 11,
            marginRight: 5
        },
        emptyHint: {
            fontSize: 12,
            color: theme.semanticColors.bodySubtext
        },
        modalDate: {
            fontSize: 12,
            color: theme.semanticColors.bodySubtext,
            marginBottom: 16
        },
        // Duyuru detayında, görseli varsa metnin ÜSTÜNDE tam genişlikte,
        // köşeleri sadece üstten yuvarlatılmış (rounded-t-lg) şekilde durur.
        modalImage: {
            width: '100%',
            maxHeight: 260,
            objectFit: 'cover',
            borderRadius: '12px 12px 0 0',
            marginBottom: 16,
            display: 'block'
        },
        modalBody: {
            fontSize: 14,
            lineHeight: 1.7,
            color: theme.semanticColors.bodyText,
            // Duyuru zengin metin (rich text) alanından geldiği için temel
            // biçimlendirme (kalın, liste, bağlantı) burada da görünsün.
            selectors: {
                'img': { maxWidth: '100%' },
                'a': { color: theme.palette.themePrimary }
            }
        },
        // Form kapsayıcısı — elemanlar arasında nefes alan, düzenli bir dikey ritim.
        // NOT: "gap" burada kullanılmıyor (flex "gap" desteklenmiyor) — her
        // çocuğa (formErrorBar, inputFieldStyles.root, dropzoneWrapper) kendi
        // marginBottom'ı veriliyor; her zaman son çocuk olan formActions'a
        // gerek yok.
        formContainer: {
            display: 'flex',
            flexDirection: 'column'
        },
        formErrorBar: {
            marginBottom: 20
        },
        // "Görsel Ekle" — dış <label> SADECE tıklama alanını/erişilebilirlik
        // bağını (htmlFor) taşır; gerçek flex/hizalama düzeni İÇTEKİ <div>
        // (imageDropzone) üzerinde. Native <label> etiketleri bazı global
        // sayfa/tema CSS'lerinde (ör. Office UI Fabric Core'un form reset'i)
        // "display" için kendi kuralına sahip olabiliyor ve mergeStyleSets'in
        // ürettiği tek sınıfı ezerek ikon ile metnin üst üste binmesine yol
        // açabiliyor — bu yüzden flex layout'u hiçbir zaman <label>'a değil,
        // her zaman sıradan bir <div>'e uygulanıyor.
        dropzoneWrapper: {
            display: 'block',
            cursor: 'pointer',
            marginBottom: 20
        },
        // NOT: "gap" flex özelliği burada BİLİNÇLİ OLARAK KULLANILMIYOR — bu
        // sayfanın render edildiği (kurumsal/eski) tarayıcı ortamında flex
        // "gap" desteklenmiyor ve tüm çocuklar sıfır boşlukla üst üste
        // biniyordu (bir önceki denemede negatif margin bunu daha da
        // kötüleştirdi). Bunun yerine HER çocuğa kendi marginBottom'ı
        // veriliyor — bu, gap desteklemeyen tarayıcılarda dahi güvenilir.
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
            fontWeight: 400,
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
        // Görsel seçildiğinde dropzone, ortalanmış ikon+etiket yerine soldan
        // hizalı bir "önizleme satırı"na dönüşür — kullanıcı ne yüklediğini
        // metinden değil gerçek küçük resimden görür.
        imageDropzoneFilled: {
            flexDirection: 'row',
            justifyContent: 'flex-start',
            textAlign: 'left',
            padding: '10px 14px'
        },
        imagePreviewThumb: {
            width: 44,
            height: 44,
            borderRadius: 8,
            objectFit: 'cover',
            flexShrink: 0,
            marginRight: 12
        },
        imagePreviewTextGroup: {
            minWidth: 0,
            overflow: 'hidden'
        },
        hiddenFileInput: {
            display: 'none'
        },
        formActions: {
            display: 'flex',
            justifyContent: 'flex-end'
        },
        cancelButton: {
            marginRight: 12
        },
        deleteConfirmBody: {
            fontSize: 14,
            lineHeight: 1.6,
            color: theme.semanticColors.bodyText
        },
        formErrorDetail: {
            fontSize: 11,
            color: theme.semanticColors.bodySubtext,
            marginTop: 6,
            fontFamily: 'Consolas, monospace',
            wordBreak: 'break-word'
        },
        pagination: {
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            marginTop: 8
        },
        // NOT: "gap" burada kullanılmıyor — önceki-sayfa butonuna ayrı bir
        // sınıf (paginationPrevButton) ile marginRight veriliyor.
        paginationPrevButton: {
            marginRight: 8
        },
        pageLabel: {
            fontSize: 11,
            color: theme.semanticColors.bodySubtext,
            minWidth: 70,
            textAlign: 'center',
            marginRight: 8
        }
    });

    const pageCount = Math.max(1, Math.ceil(items.length / PAGE_SIZE));
    const pagedItems = items.slice(page * PAGE_SIZE, page * PAGE_SIZE + PAGE_SIZE);

    return (
        <WidgetCard
            title="Duyurular"
            subtitle="Şirket içi güncel haber akışı"
            iconName="News"
            headerAction={canManageAnnouncements && (
                <IconButton
                    iconProps={{ iconName: 'Add' }}
                    ariaLabel="Yeni duyuru ekle"
                    title="Yeni duyuru ekle"
                    onClick={() => setIsAddOpen(true)}
                />
            )}
        >
            {state === 'loading' && <Spinner size={SpinnerSize.small} label="Duyurular yükleniyor..." />}
            {state === 'error' && <MessageBar messageBarType={MessageBarType.error}>{DATA_UNAVAILABLE_MESSAGE}</MessageBar>}
            {state === 'loaded' && items.length === 0 && (
                <Text className={styles.emptyHint}>Görüntülenecek duyuru bulunmuyor.</Text>
            )}
            {state === 'loaded' && items.length > 0 && (
                <>
                    <div className={styles.list}>
                        {pagedItems.map((a, index) => (
                            <button
                                key={a.id}
                                type="button"
                                className={`${styles.row} ${index === pagedItems.length - 1 ? styles.lastRow : ''}`}
                                onClick={() => setSelected(a)}
                            >
                                {/* Görsel varsa küçük kapak resmi, yoksa eski/sade ikon rozeti —
                                    iki durum arasında boşluk veya kırık resim ikonu YOK. */}
                                {a.imageUrl ? (
                                    <img src={a.imageUrl} alt="" className={styles.thumb} />
                                ) : (
                                    <div className={styles.iconWrap}>
                                        <Icon iconName="Megaphone" className={styles.icon} />
                                    </div>
                                )}
                                <div className={styles.textGroup}>
                                    <div className={styles.title}>{a.title}</div>
                                    <div className={styles.date}>
                                        <Icon iconName="Calendar" className={styles.dateIcon} />
                                        {a.dateLabel}
                                    </div>
                                </div>
                            </button>
                        ))}
                    </div>
                    {/* Duyurular asla otomatik silinmez — liste zamanla uzar, bu yüzden
                        sayfa başına sabit (PAGE_SIZE) adet + ileri/geri oklarıyla gezinme. */}
                    {items.length > PAGE_SIZE && (
                        <div className={styles.pagination}>
                            <IconButton
                                className={styles.paginationPrevButton}
                                iconProps={{ iconName: 'ChevronLeft' }}
                                ariaLabel="Önceki sayfa"
                                disabled={page === 0}
                                onClick={() => setPage((p) => Math.max(0, p - 1))}
                            />
                            <span className={styles.pageLabel}>Sayfa {page + 1} / {pageCount}</span>
                            <IconButton
                                iconProps={{ iconName: 'ChevronRight' }}
                                ariaLabel="Sonraki sayfa"
                                disabled={page >= pageCount - 1}
                                onClick={() => setPage((p) => Math.min(pageCount - 1, p + 1))}
                            />
                        </div>
                    )}
                </>
            )}

            <DetailModal
                isOpen={!!selected}
                title={selected?.title}
                onDismiss={() => setSelected(undefined)}
                onDeleteClick={canManageAnnouncements ? () => setIsDeleteConfirmOpen(true) : undefined}
                deleteAriaLabel="Duyuruyu sil"
            >
                {selected && (
                    <>
                        {selected.imageUrl && (
                            <img src={selected.imageUrl} alt="" className={styles.modalImage} />
                        )}
                        <div className={styles.modalDate}>{selected.dateLabel}</div>
                        {selected.bodyHtml && (
                            // Duyuru içeriği SharePoint'in zengin metin (rich text) editöründen gelir;
                            // yine de OWASP XSS önlemi gereği ham HTML render edilmeden ÖNCE
                            // sanitizeHtml() ile tehlikeli etiket/öznitelikler temizleniyor.
                            <div className={styles.modalBody} dangerouslySetInnerHTML={{ __html: sanitizeHtml(selected.bodyHtml) }} />
                        )}
                    </>
                )}
            </DetailModal>

            <Dialog
                hidden={!isDeleteConfirmOpen}
                onDismiss={() => { setIsDeleteConfirmOpen(false); setDeleteState('idle'); }}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: 'Duyuru silinsin mi?',
                    subText: `"${selected?.title ?? ''}" başlıklı duyuru kalıcı olarak silinecek. Bu işlem geri alınamaz.`
                }}
            >
                {deleteState === 'error' && (
                    <MessageBar messageBarType={MessageBarType.error}>Duyuru silinemedi. Lütfen tekrar deneyin.</MessageBar>
                )}
                <DialogFooter>
                    {deleteState === 'sending' ? (
                        <Spinner size={SpinnerSize.small} label="Siliniyor..." />
                    ) : (
                        <>
                            <PrimaryButton
                                text="Evet, Sil"
                                styles={{ root: { background: '#B91C1C', border: 'none' }, rootHovered: { background: '#991B1B' } }}
                                onClick={() => { handleDeleteConfirm().catch(() => { /* handled in handleDeleteConfirm */ }); }}
                            />
                            <DefaultButton text="Vazgeç" onClick={() => { setIsDeleteConfirmOpen(false); setDeleteState('idle'); }} />
                        </>
                    )}
                </DialogFooter>
            </Dialog>

            <DetailModal isOpen={isAddOpen} title="Yeni Duyuru Ekle" onDismiss={closeAddModal}>
                <div className={styles.formContainer}>
                    {submitState === 'error' && (
                        <MessageBar className={styles.formErrorBar} messageBarType={MessageBarType.error} onDismiss={() => setSubmitState('idle')}>
                            Duyuru eklenemedi.
                            {submitError && <div className={styles.formErrorDetail}>{submitError}</div>}
                        </MessageBar>
                    )}
                    <TextField
                        label="Duyuru Başlığı"
                        value={newTitle}
                        onChange={(_, v) => setNewTitle(v ?? '')}
                        disabled={submitState === 'sending'}
                        styles={inputFieldStyles}
                    />
                    <TextField
                        label="Duyuru İçeriği"
                        multiline
                        rows={5}
                        value={newBody}
                        onChange={(_, v) => setNewBody(v ?? '')}
                        disabled={submitState === 'sending'}
                        styles={inputFieldStyles}
                    />
                    <label className={styles.dropzoneWrapper} htmlFor="announcement-image-input">
                        <div className={`${styles.imageDropzone} ${newImagePreviewUrl ? styles.imageDropzoneFilled : ''}`}>
                            {newImagePreviewUrl ? (
                                <>
                                    <img src={newImagePreviewUrl} alt="" className={styles.imagePreviewThumb} />
                                    <div className={styles.imagePreviewTextGroup}>
                                        <div className={styles.imageDropzoneFileName} style={{ marginTop: 0 }}>{newImage?.name}</div>
                                        <span className={styles.imageDropzoneHint}>Değiştirmek için tıklayın</span>
                                    </div>
                                </>
                            ) : (
                                <>
                                    <div className={styles.imageDropzoneIconWrap}>
                                        <Icon iconName="Photo2Add" className={styles.imageDropzoneIcon} />
                                    </div>
                                    <span className={styles.imageDropzoneLabel}>Görsel Ekle</span>
                                    <span className={styles.imageDropzoneHint}>Opsiyonel</span>
                                </>
                            )}
                        </div>
                        <input
                            id="announcement-image-input"
                            type="file"
                            accept="image/*"
                            className={styles.hiddenFileInput}
                            disabled={submitState === 'sending'}
                            onChange={(e) => setNewImage(e.target.files?.[0])}
                        />
                    </label>
                    <div className={styles.formActions}>
                        {submitState === 'sending' ? (
                            <Spinner size={SpinnerSize.small} label="Kaydediliyor..." />
                        ) : (
                            <>
                                <DefaultButton className={styles.cancelButton} text="Vazgeç" onClick={closeAddModal} />
                                <PrimaryButton
                                    text="Kaydet"
                                    onClick={() => { handleAddSubmit().catch(() => { /* handled in handleAddSubmit */ }); }}
                                    disabled={!newTitle.trim() || !newBody.trim()}
                                />
                            </>
                        )}
                    </div>
                </div>
            </DetailModal>
        </WidgetCard>
    );
};

export default AnnouncementsFeed;
