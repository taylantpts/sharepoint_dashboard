import * as React from 'react';
import {
    Icon, Spinner, SpinnerSize, MessageBar, MessageBarType, Text, IconButton, TextField, PrimaryButton,
    DefaultButton, Dropdown, IDropdownOption, useTheme, mergeStyleSets, ITextFieldStyles
} from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import WidgetCard from '../WidgetCard';
import DetailModal from '../DetailModal';
import {
    getIkinciElListings, createIkinciElListing, deleteIkinciElListing, getCurrentUserId,
    IIkinciElItem, IKINCI_EL_CATEGORIES
} from '../../services/SharePointService';
import { DATA_UNAVAILABLE_MESSAGE } from '../../constants';

export interface IIkinciElWidgetProps {
    context: WebPartContext;
}

type LoadState = 'loading' | 'loaded' | 'error';
type SubmitState = 'idle' | 'sending' | 'error';

// Kullanıcı: "satış sohbet vs olmayacak sadece ilan koyacak" — bu yüzden
// widget PAYLAŞMA/SATIN ALMA akışı İÇERMİYOR, sadece ilan görüntüleme +
// posterin e-postasıyla (mevcut CompanyDirectory'deki gibi mailto) iletişim.
// Kategori DEĞERLERİ (STORED) SharePoint'teki Kategori (Choice) sütunuyla
// birebir aynı olmak zorunda (bkz. SharePointService.ts IKINCI_EL_CATEGORIES)
// — bu yüzden ASCII tutuluyor; burada sadece GÖRÜNEN etiketler Türkçeleştirilir.
const CATEGORY_DISPLAY_LABELS: Record<string, string> = {
    'Elektronik': 'Elektronik',
    'Ev ve Yasam': 'Ev & Yaşam',
    'Giyim ve Aksesuar': 'Giyim & Aksesuar',
    'Arac ve Vasita': 'Araç & Vasıta',
    'Kitap ve Hobi': 'Kitap & Hobi',
    'Diger': 'Diğer'
};

const PAGE_SIZE = 4;

const inputFieldStyles: Partial<ITextFieldStyles> = {
    root: { marginBottom: 20 },
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
    subComponentStyles: {
        label: { root: { marginBottom: 8 } }
    }
};

const IkinciElWidget: React.FunctionComponent<IIkinciElWidgetProps> = (props) => {
    const { context } = props;
    const theme = useTheme();

    const [listings, setListings] = React.useState<IIkinciElItem[]>([]);
    const [state, setState] = React.useState<LoadState>('loading');
    const [selected, setSelected] = React.useState<IIkinciElItem | undefined>(undefined);
    const [selectedImageIndex, setSelectedImageIndex] = React.useState(0);
    const [page, setPage] = React.useState(0);
    const [currentUserId, setCurrentUserId] = React.useState<number | undefined>(undefined);

    const [isAddOpen, setIsAddOpen] = React.useState(false);
    const [newTitle, setNewTitle] = React.useState('');
    const [newCategory, setNewCategory] = React.useState<string | undefined>(undefined);
    const [newPrice, setNewPrice] = React.useState('');
    const [newDescription, setNewDescription] = React.useState('');
    const [newImages, setNewImages] = React.useState<File[]>([]);
    const [submitState, setSubmitState] = React.useState<SubmitState>('idle');
    const [submitError, setSubmitError] = React.useState<string | undefined>(undefined);

    const [isDeleteConfirmOpen, setIsDeleteConfirmOpen] = React.useState(false);
    const [deleteState, setDeleteState] = React.useState<SubmitState>('idle');

    // Seçilen görseller için canlı önizlemeler — bkz. AnnouncementsFeed'deki
    // aynı desen, burada ÇOKLU dosya için diziye genişletildi.
    const [newImagePreviewUrls, setNewImagePreviewUrls] = React.useState<string[]>([]);
    React.useEffect(() => {
        if (newImages.length === 0) {
            setNewImagePreviewUrls([]);
            return;
        }
        const urls = newImages.map((f) => URL.createObjectURL(f));
        setNewImagePreviewUrls(urls);
        return () => urls.forEach((u) => URL.revokeObjectURL(u));
    }, [newImages]);

    const loadListings = React.useCallback((onlyIfMounted?: () => boolean): void => {
        setState('loading');
        getIkinciElListings(context)
            .then((result) => {
                if (!onlyIfMounted || onlyIfMounted()) {
                    setListings(result);
                    setState('loaded');
                }
            })
            .catch((error: Error) => {
                console.error('[IkinciElWidget] İlanlar alınamadı:', error);
                if (!onlyIfMounted || onlyIfMounted()) {
                    setState('error');
                }
            });
    }, [context]);

    React.useEffect(() => {
        let isMounted = true;
        loadListings(() => isMounted);
        getCurrentUserId(context).then((id) => {
            if (isMounted) {
                setCurrentUserId(id);
            }
        }).catch(() => { /* kendi ilanını silme kontrolü sessizce devre dışı kalır */ });
        return () => {
            isMounted = false;
        };
    }, [loadListings, context]);

    const closeAddModal = (): void => {
        setIsAddOpen(false);
        setNewTitle('');
        setNewCategory(undefined);
        setNewPrice('');
        setNewDescription('');
        setNewImages([]);
        setSubmitState('idle');
        setSubmitError(undefined);
    };

    const handleAddSubmit = async (): Promise<void> => {
        if (!newTitle.trim() || !newCategory) {
            return;
        }
        setSubmitState('sending');
        setSubmitError(undefined);

        try {
            const result = await createIkinciElListing(
                context, newTitle.trim(), newDescription.trim(), newCategory, newPrice.trim(), newImages
            );
            if (result.success) {
                closeAddModal();
                setPage(0);
                loadListings();
            } else {
                setSubmitError(result.errorMessage);
                setSubmitState('error');
            }
        } catch (error) {
            setSubmitError((error as Error).message);
            setSubmitState('error');
        }
    };

    const handleDeleteConfirm = async (): Promise<void> => {
        if (!selected) {
            return;
        }
        setDeleteState('sending');
        try {
            const result = await deleteIkinciElListing(context, selected.id);
            if (result.success) {
                setIsDeleteConfirmOpen(false);
                setSelected(undefined);
                setDeleteState('idle');
                loadListings();
            } else {
                setDeleteState('error');
            }
        } catch {
            setDeleteState('error');
        }
    };

    const openDetail = (item: IIkinciElItem): void => {
        setSelected(item);
        setSelectedImageIndex(0);
    };

    const styles = mergeStyleSets({
        grid: {
            display: 'grid',
            gridTemplateColumns: 'repeat(2, 1fr)',
            gap: 14
        },
        card: {
            display: 'flex',
            flexDirection: 'column',
            width: '100%',
            background: '#ffffff',
            border: `1px solid ${theme.palette.neutralLighter}`,
            borderRadius: 14,
            overflow: 'hidden',
            cursor: 'pointer',
            textAlign: 'left',
            font: 'inherit',
            padding: 0,
            transition: 'box-shadow 0.15s ease, transform 0.15s ease',
            selectors: {
                ':hover': { boxShadow: '0 8px 20px rgba(15,23,42,0.10)', transform: 'translateY(-2px)' }
            }
        },
        cardImage: {
            width: '100%',
            height: 110,
            objectFit: 'cover',
            display: 'block',
            background: theme.palette.neutralLighterAlt
        },
        cardImagePlaceholder: {
            width: '100%',
            height: 110,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            background: theme.palette.neutralLighterAlt,
            color: theme.palette.neutralTertiary,
            fontSize: 28
        },
        cardBody: {
            padding: '10px 12px'
        },
        cardCategory: {
            display: 'inline-block',
            fontSize: 10,
            fontWeight: 700,
            color: theme.palette.themePrimary,
            background: theme.palette.themeLighterAlt,
            borderRadius: 6,
            padding: '2px 7px',
            marginBottom: 6,
            textTransform: 'uppercase',
            letterSpacing: 0.3
        },
        cardTitle: {
            fontSize: 13,
            fontWeight: 700,
            color: theme.semanticColors.bodyText,
            overflow: 'hidden',
            textOverflow: 'ellipsis',
            whiteSpace: 'nowrap',
            marginBottom: 2
        },
        // NOT: "gap" burada kullanılmıyor — cardPoster'a marginRight yerine
        // ikisi ayrı satır (flex column) olduğu için ihtiyaç yok.
        cardMetaRow: {
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            marginTop: 4
        },
        cardPrice: {
            fontSize: 12,
            fontWeight: 700,
            color: theme.semanticColors.bodyText
        },
        cardPoster: {
            fontSize: 10,
            color: theme.semanticColors.bodySubtext,
            overflow: 'hidden',
            textOverflow: 'ellipsis',
            whiteSpace: 'nowrap'
        },
        emptyHint: {
            fontSize: 12,
            color: theme.semanticColors.bodySubtext
        },
        pagination: {
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            marginTop: 12
        },
        paginationPrevButton: {
            marginRight: 8
        },
        pageLabel: {
            fontSize: 11,
            color: theme.semanticColors.bodySubtext,
            minWidth: 70,
            textAlign: 'center',
            marginRight: 8
        },
        // --- Detay modalı ---
        modalMainImage: {
            width: '100%',
            height: 260,
            objectFit: 'contain',
            background: theme.palette.neutralLighterAlt,
            borderRadius: '12px 12px 0 0',
            marginBottom: 12,
            display: 'block'
        },
        modalMainImagePlaceholder: {
            width: '100%',
            height: 200,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            background: theme.palette.neutralLighterAlt,
            color: theme.palette.neutralTertiary,
            fontSize: 48,
            borderRadius: '12px 12px 0 0',
            marginBottom: 12
        },
        // NOT: "gap" burada kullanılmıyor — thumbnail'lara marginRight veriliyor.
        thumbnailRow: {
            display: 'flex',
            marginBottom: 16,
            overflowX: 'auto'
        },
        thumbnail: {
            width: 48,
            height: 48,
            objectFit: 'cover',
            borderRadius: 8,
            marginRight: 8,
            cursor: 'pointer',
            flexShrink: 0,
            border: '2px solid transparent',
            opacity: 0.6,
            transition: 'opacity 0.15s ease, border-color 0.15s ease'
        },
        thumbnailActive: {
            opacity: 1,
            borderColor: theme.palette.themePrimary
        },
        modalCategory: {
            display: 'inline-block',
            fontSize: 11,
            fontWeight: 700,
            color: theme.palette.themePrimary,
            background: theme.palette.themeLighterAlt,
            borderRadius: 6,
            padding: '3px 9px',
            marginBottom: 10,
            textTransform: 'uppercase',
            letterSpacing: 0.3
        },
        modalPrice: {
            fontSize: 20,
            fontWeight: 800,
            color: theme.semanticColors.bodyText,
            marginBottom: 10
        },
        modalDescription: {
            fontSize: 14,
            lineHeight: '160%',
            color: theme.semanticColors.bodyText,
            marginBottom: 16
        },
        modalPosterRow: {
            display: 'flex',
            alignItems: 'center',
            fontSize: 13,
            color: theme.semanticColors.bodySubtext,
            marginBottom: 16
        },
        modalPosterIcon: {
            marginRight: 8
        },
        // --- Form ---
        formContainer: {
            display: 'flex',
            flexDirection: 'column'
        },
        formErrorBar: {
            marginBottom: 20
        },
        dropzoneWrapper: {
            display: 'block',
            cursor: 'pointer',
            marginBottom: 20
        },
        imageDropzone: {
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            justifyContent: 'center',
            padding: '20px 20px',
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
            marginBottom: 10
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
        // NOT: "gap" burada kullanılmıyor — previewThumb'lara marginRight/
        // marginBottom veriliyor.
        previewGrid: {
            display: 'flex',
            flexWrap: 'wrap',
            marginTop: 12
        },
        previewThumbWrap: {
            position: 'relative',
            width: 56,
            height: 56,
            marginRight: 8,
            marginBottom: 8
        },
        previewThumb: {
            width: '100%',
            height: '100%',
            objectFit: 'cover',
            borderRadius: 8
        },
        previewRemoveButton: {
            position: 'absolute',
            top: -6,
            right: -6,
            width: 20,
            height: 20,
            minWidth: 20,
            borderRadius: '50%',
            background: '#B91C1C',
            color: '#ffffff',
            selectors: {
                ':hover': { background: '#991B1B', color: '#ffffff' }
            }
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
        formErrorDetail: {
            fontSize: 11,
            color: theme.semanticColors.bodySubtext,
            marginTop: 6,
            fontFamily: 'Consolas, monospace',
            wordBreak: 'break-word'
        }
    });

    const pageCount = Math.max(1, Math.ceil(listings.length / PAGE_SIZE));
    const pagedListings = listings.slice(page * PAGE_SIZE, page * PAGE_SIZE + PAGE_SIZE);

    const categoryOptions: IDropdownOption[] = IKINCI_EL_CATEGORIES.map((c) => ({
        key: c, text: CATEGORY_DISPLAY_LABELS[c] ?? c
    }));

    const isOwnListing = !!selected && !!currentUserId && selected.posterId === currentUserId;

    return (
        <WidgetCard
            title="İkinci El İlanlar"
            subtitle="Çalışanlar arası eşya ilan panosu"
            iconName="ShoppingCart"
            headerAction={(
                <IconButton
                    iconProps={{ iconName: 'Add' }}
                    ariaLabel="Yeni ilan ekle"
                    title="Yeni ilan ekle"
                    onClick={() => setIsAddOpen(true)}
                />
            )}
        >
            {state === 'loading' && <Spinner size={SpinnerSize.small} label="İlanlar yükleniyor..." />}
            {state === 'error' && <MessageBar messageBarType={MessageBarType.error}>{DATA_UNAVAILABLE_MESSAGE}</MessageBar>}
            {state === 'loaded' && listings.length === 0 && (
                <Text className={styles.emptyHint}>Henüz bir ilan bulunmuyor. İlk ilanı sen ekle!</Text>
            )}
            {state === 'loaded' && listings.length > 0 && (
                <>
                    <div className={styles.grid}>
                        {pagedListings.map((item) => (
                            <button key={item.id} type="button" className={styles.card} onClick={() => openDetail(item)}>
                                {item.imageUrls[0] ? (
                                    <img src={item.imageUrls[0]} alt="" className={styles.cardImage} />
                                ) : (
                                    <div className={styles.cardImagePlaceholder}>
                                        <Icon iconName="Tag" />
                                    </div>
                                )}
                                <div className={styles.cardBody}>
                                    <span className={styles.cardCategory}>{CATEGORY_DISPLAY_LABELS[item.category] ?? item.category}</span>
                                    <div className={styles.cardTitle}>{item.title}</div>
                                    <div className={styles.cardMetaRow}>
                                        <span className={styles.cardPrice}>{item.price || '—'}</span>
                                        <span className={styles.cardPoster}>{item.posterName}</span>
                                    </div>
                                </div>
                            </button>
                        ))}
                    </div>
                    {listings.length > PAGE_SIZE && (
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
                onDeleteClick={isOwnListing ? () => setIsDeleteConfirmOpen(true) : undefined}
                deleteAriaLabel="İlanı sil"
            >
                {selected && (
                    <>
                        {selected.imageUrls.length > 0 ? (
                            <img src={selected.imageUrls[selectedImageIndex]} alt="" className={styles.modalMainImage} />
                        ) : (
                            <div className={styles.modalMainImagePlaceholder}>
                                <Icon iconName="Tag" />
                            </div>
                        )}
                        {selected.imageUrls.length > 1 && (
                            <div className={styles.thumbnailRow}>
                                {selected.imageUrls.map((url, index) => (
                                    <img
                                        key={url}
                                        src={url}
                                        alt=""
                                        className={`${styles.thumbnail} ${index === selectedImageIndex ? styles.thumbnailActive : ''}`}
                                        onClick={() => setSelectedImageIndex(index)}
                                    />
                                ))}
                            </div>
                        )}
                        <span className={styles.modalCategory}>{CATEGORY_DISPLAY_LABELS[selected.category] ?? selected.category}</span>
                        {selected.price && <div className={styles.modalPrice}>{selected.price}</div>}
                        {selected.description && (
                            <div className={styles.modalDescription}>
                                {selected.description.split('\n').map((line, index) => (
                                    <div key={index}>{line || ' '}</div>
                                ))}
                            </div>
                        )}
                        <div className={styles.modalPosterRow}>
                            <Icon iconName="Contact" className={styles.modalPosterIcon} />
                            {selected.posterName} · {selected.dateLabel}
                        </div>
                        {selected.posterEmail && (
                            <DefaultButton
                                iconProps={{ iconName: 'Mail' }}
                                text="E-postayla İletişime Geç"
                                href={`mailto:${selected.posterEmail}?subject=${encodeURIComponent(selected.title)}`}
                            />
                        )}
                    </>
                )}
            </DetailModal>

            <DetailModal
                isOpen={isDeleteConfirmOpen}
                title="İlan silinsin mi?"
                onDismiss={() => { setIsDeleteConfirmOpen(false); setDeleteState('idle'); }}
            >
                <Text>&quot;{selected?.title ?? ''}&quot; başlıklı ilan kalıcı olarak silinecek. Bu işlem geri alınamaz.</Text>
                {deleteState === 'error' && (
                    <MessageBar messageBarType={MessageBarType.error} styles={{ root: { marginTop: 12 } }}>
                        İlan silinemedi. Lütfen tekrar deneyin.
                    </MessageBar>
                )}
                <div className={styles.formActions} style={{ marginTop: 20 }}>
                    {deleteState === 'sending' ? (
                        <Spinner size={SpinnerSize.small} label="Siliniyor..." />
                    ) : (
                        <>
                            <DefaultButton
                                className={styles.cancelButton}
                                text="Vazgeç"
                                onClick={() => { setIsDeleteConfirmOpen(false); setDeleteState('idle'); }}
                            />
                            <PrimaryButton
                                text="Evet, Sil"
                                styles={{ root: { background: '#B91C1C', border: 'none' }, rootHovered: { background: '#991B1B' } }}
                                onClick={() => { handleDeleteConfirm().catch(() => { /* handled in handleDeleteConfirm */ }); }}
                            />
                        </>
                    )}
                </div>
            </DetailModal>

            <DetailModal isOpen={isAddOpen} title="Yeni İlan Ekle" onDismiss={closeAddModal}>
                <div className={styles.formContainer}>
                    {submitState === 'error' && (
                        <MessageBar className={styles.formErrorBar} messageBarType={MessageBarType.error} onDismiss={() => setSubmitState('idle')}>
                            İlan eklenemedi.
                            {submitError && <div className={styles.formErrorDetail}>{submitError}</div>}
                        </MessageBar>
                    )}
                    <TextField
                        label="İlan Başlığı"
                        value={newTitle}
                        onChange={(_, v) => setNewTitle(v ?? '')}
                        disabled={submitState === 'sending'}
                        styles={inputFieldStyles}
                    />
                    <Dropdown
                        label="Kategori"
                        placeholder="Kategori seçin"
                        options={categoryOptions}
                        selectedKey={newCategory}
                        onChange={(_, option) => setNewCategory(option?.key as string)}
                        disabled={submitState === 'sending'}
                        styles={{ root: { marginBottom: 20 }, dropdown: { borderRadius: 12 }, title: { borderRadius: 12, background: '#F8FAFC', border: '1px solid #CBD5E1', height: 36, lineHeight: '34px' } }}
                    />
                    <TextField
                        label="Fiyat (opsiyonel)"
                        placeholder="ör. 500 TL, Pazarlıklı"
                        value={newPrice}
                        onChange={(_, v) => setNewPrice(v ?? '')}
                        disabled={submitState === 'sending'}
                        styles={inputFieldStyles}
                    />
                    <TextField
                        label="Açıklama"
                        multiline
                        rows={4}
                        value={newDescription}
                        onChange={(_, v) => setNewDescription(v ?? '')}
                        disabled={submitState === 'sending'}
                        styles={inputFieldStyles}
                    />
                    <label className={styles.dropzoneWrapper} htmlFor="ikincel-image-input">
                        <div className={styles.imageDropzone}>
                            <div className={styles.imageDropzoneIconWrap}>
                                <Icon iconName="Photo2Add" className={styles.imageDropzoneIcon} />
                            </div>
                            <span className={styles.imageDropzoneLabel}>Görsel Ekle</span>
                            <span className={styles.imageDropzoneHint}>Birden fazla seçebilirsiniz — opsiyonel</span>
                        </div>
                        <input
                            id="ikincel-image-input"
                            type="file"
                            accept="image/*"
                            multiple
                            className={styles.hiddenFileInput}
                            disabled={submitState === 'sending'}
                            onChange={(e) => setNewImages(e.target.files ? Array.from(e.target.files) : [])}
                        />
                    </label>
                    {newImagePreviewUrls.length > 0 && (
                        <div className={styles.previewGrid}>
                            {newImagePreviewUrls.map((url, index) => (
                                <div key={url} className={styles.previewThumbWrap}>
                                    <img src={url} alt="" className={styles.previewThumb} />
                                    <IconButton
                                        className={styles.previewRemoveButton}
                                        iconProps={{ iconName: 'Cancel' }}
                                        ariaLabel="Görseli kaldır"
                                        onClick={() => setNewImages((prev) => prev.filter((_, i) => i !== index))}
                                    />
                                </div>
                            ))}
                        </div>
                    )}
                    <div className={styles.formActions} style={{ marginTop: 20 }}>
                        {submitState === 'sending' ? (
                            <Spinner size={SpinnerSize.small} label="Kaydediliyor..." />
                        ) : (
                            <>
                                <DefaultButton className={styles.cancelButton} text="Vazgeç" onClick={closeAddModal} />
                                <PrimaryButton
                                    text="Kaydet"
                                    onClick={() => { handleAddSubmit().catch(() => { /* handled in handleAddSubmit */ }); }}
                                    disabled={!newTitle.trim() || !newCategory}
                                />
                            </>
                        )}
                    </div>
                </div>
            </DetailModal>
        </WidgetCard>
    );
};

export default IkinciElWidget;
