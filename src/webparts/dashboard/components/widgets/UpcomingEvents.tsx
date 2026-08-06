import * as React from 'react';
import {
    Icon, Spinner, SpinnerSize, MessageBar, MessageBarType, Text, IconButton, TextField, PrimaryButton,
    DefaultButton, Dialog, DialogType, DialogFooter,
    DatePicker, DayOfWeek, useTheme, mergeStyleSets, ITextFieldStyles
} from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import WidgetCard from '../WidgetCard';
import DetailModal from '../DetailModal';
import { getUpcomingEvents, createEvent, deleteEvent, IUpcomingEventItem } from '../../services/SharePointService';
import { DATA_UNAVAILABLE_MESSAGE } from '../../constants';
import { usePermissions } from '../../hooks/usePermissions';

export interface IUpcomingEventsProps {
    context: WebPartContext;
}

type LoadState = 'loading' | 'loaded' | 'error';
type SubmitState = 'idle' | 'sending' | 'error';

const DAY_PICKER_STRINGS = {
    months: ['Ocak', 'Şubat', 'Mart', 'Nisan', 'Mayıs', 'Haziran', 'Temmuz', 'Ağustos', 'Eylül', 'Ekim', 'Kasım', 'Aralık'],
    shortMonths: ['Oca', 'Şub', 'Mar', 'Nis', 'May', 'Haz', 'Tem', 'Ağu', 'Eyl', 'Eki', 'Kas', 'Ara'],
    days: ['Pazar', 'Pazartesi', 'Salı', 'Çarşamba', 'Perşembe', 'Cuma', 'Cumartesi'],
    shortDays: ['Pz', 'Pt', 'Sa', 'Ça', 'Pe', 'Cu', 'Ct'],
    goToToday: 'Bugüne git',
    prevMonthAriaLabel: 'Önceki ay',
    nextMonthAriaLabel: 'Sonraki ay',
    prevYearAriaLabel: 'Önceki yıl',
    nextYearAriaLabel: 'Sonraki yıl',
    closeButtonAriaLabel: 'Kapat'
};

// Form alanlarının ortak "yumuşak" görünümü — hafif gri zemin, ince/zarif
// kenarlık ve odaklanınca mavi bir halka.
// ÖNCEKİ HATA: kenarlık (#E2E8F0) ve zemin o kadar açık/soluktu ki alan,
// beyaz modal zemininin üzerinde neredeyse görünmez oluyor, "devre dışı"
// bir kutu gibi duruyordu — kullanıcıya buranın tıklanabilir bir giriş
// alanı olduğunu anlatmıyordu. Kenarlık koyulaştırıldı + hafif bir iç
// gölge (inset shadow) eklendi, gerçek bir "input kutusu" hissi versin.
const inputFieldStyles: Partial<ITextFieldStyles> = {
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
    // ÖNCEKİ HATA: etiket (ör. "Etkinlik Başlığı") ile altındaki kutu
    // arasında hiç boşluk YOKTU — bkz. AnnouncementsFeed.tsx'teki aynı not.
    subComponentStyles: {
        label: { root: { marginBottom: 8 } }
    }
};

// formContainer'ın DOĞRUDAN çocuğu olan TextField'lar (Etkinlik Başlığı/
// Detayı) için — inputFieldStyles'ın KENDİSİNE root marginBottom eklenemez,
// çünkü aynı sabit DatePicker'ın İÇ textField'ı tarafından da paylaşılıyor
// (formRow'un içinde, formContainer'ın gap'inden bağımsız bir bağlamda) —
// oraya sızarsa Tarih/Saat sütunlarının hizasını bozar.
const topLevelFieldStyles: Partial<ITextFieldStyles> = {
    ...inputFieldStyles,
    root: { marginBottom: 20 }
};

// ÖNCEKİ HATA: DatePicker'ın kendi içindeki TextField, inputFieldStyles'ı
// (fieldGroup: minHeight:36) paylaşıyordu — ama DatePicker'ın takvim ikonu
// kutuyu Saat alanından (native <input type=time>, height:36 SABİT) daha
// UZUN büyütüyordu (muhtemelen ikonun kendi doğal yüksekliği + flex
// stretch), bu da tarih metninin kutunun altına yapışmış görünmesine yol
// açıyordu ("textbox'ın en altına yapışıyor" geri bildirimi). minHeight
// yerine SABİT height:36 ile — Saat alanıyla birebir aynı yükseklik ve
// dikey ortalama.
const dateFieldStyles: Partial<ITextFieldStyles> = {
    fieldGroup: {
        height: 36,
        minHeight: 36,
        background: '#F8FAFC',
        border: '1px solid #CBD5E1',
        borderRadius: 12,
        boxShadow: 'inset 0 1px 2px rgba(15,23,42,0.05)',
        selectors: {
            ':after': { borderRadius: 12, border: '2px solid #3B82F6', boxShadow: '0 0 0 4px rgba(59,130,246,0.15)' }
        }
    },
    field: { padding: '0 16px', fontSize: 14, lineHeight: '34px' },
    // ÖNCEKİ HATA: Fluent'in varsayılan <Label> bileşeni kendi "padding: 5px 0"
    // stilini taşıyor (marginBottom:8 override'ı bunu SIFIRLAMIYOR) — bu da
    // DatePicker'ın "Tarih" etiketini yandaki "Saat" etiketinden (padding'i
    // olmayan sade bir <label>) 10px daha uzun yapıp altındaki kutuyu 10px
    // aşağı itiyordu, Tarih/Saat kutuları aynı satırda hizasız görünüyordu.
    subComponentStyles: {
        label: { root: { marginBottom: 8, padding: 0 } }
    }
};

// Kullanıcı isteğiyle: her zaman en yakın 2 etkinlik görünür, yenisi
// eklenince en uzak (eski) olan otomatik olarak sonraki sayfaya kayar
// (getUpcomingEvents zaten tarihe göre en yakından en uzağa sıralı
// döndürüyor) — bkz. AnnouncementsFeed.tsx'teki aynı desen.
const PAGE_SIZE = 2;

const UpcomingEvents: React.FunctionComponent<IUpcomingEventsProps> = (props) => {
    const { canManageAnnouncements } = usePermissions(props.context);
    const { context } = props;
    const theme = useTheme();

    const [events, setEvents] = React.useState<IUpcomingEventItem[]>([]);
    const [state, setState] = React.useState<LoadState>('loading');
    const [selected, setSelected] = React.useState<IUpcomingEventItem | undefined>(undefined);
    const [page, setPage] = React.useState(0);

    const [isAddOpen, setIsAddOpen] = React.useState(false);
    const [newTitle, setNewTitle] = React.useState('');
    const [newDate, setNewDate] = React.useState<Date | undefined>(undefined);
    const [newTime, setNewTime] = React.useState('09:00');
    const [newDescription, setNewDescription] = React.useState('');
    const [newImage, setNewImage] = React.useState<File | undefined>(undefined);
    const [submitState, setSubmitState] = React.useState<SubmitState>('idle');
    const [submitError, setSubmitError] = React.useState<string | undefined>(undefined);

    const [isDeleteConfirmOpen, setIsDeleteConfirmOpen] = React.useState(false);
    const [deleteState, setDeleteState] = React.useState<SubmitState>('idle');

    // Seçilen görsel için canlı önizleme — bkz. AnnouncementsFeed'deki aynı desen.
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

    const loadEvents = React.useCallback((onlyIfMounted?: () => boolean): void => {
        setState('loading');
        getUpcomingEvents(context)
            .then((result) => {
                if (!onlyIfMounted || onlyIfMounted()) {
                    setEvents(result);
                    setState('loaded');
                }
            })
            .catch((error: Error) => {
                console.error('[UpcomingEvents] Etkinlikler alınamadı:', error);
                if (!onlyIfMounted || onlyIfMounted()) {
                    setState('error');
                }
            });
    }, [context]);

    React.useEffect(() => {
        let isMounted = true;
        loadEvents(() => isMounted);
        return () => {
            isMounted = false;
        };
    }, [loadEvents]);

    const closeAddModal = (): void => {
        setIsAddOpen(false);
        setNewTitle('');
        setNewDate(undefined);
        setNewTime('09:00');
        setNewDescription('');
        setNewImage(undefined);
        setSubmitState('idle');
        setSubmitError(undefined);
    };

    /**
     * "Kaydet" — girilen veri SharePoint'teki GERÇEK "Etkinlikler" listesine
     * yazılır (yerel bir state DEĞİL); başarılı olursa liste sunucudan tekrar
     * çekilir (loadEvents) — ekranda görünen, gerçekten kaydedilmiş veridir.
     * Tarih (DatePicker) + saat (native time input) tek bir Date'e birleştirilir.
     */
    const handleAddSubmit = async (): Promise<void> => {
        if (!newTitle.trim() || !newDate) {
            return;
        }
        setSubmitState('sending');
        setSubmitError(undefined);

        try {
            const [hourStr, minuteStr] = newTime.split(':');
            const combined = new Date(newDate);
            combined.setHours(parseInt(hourStr, 10) || 0, parseInt(minuteStr, 10) || 0, 0, 0);

            const result = await createEvent(context, newTitle.trim(), combined, newDescription.trim(), undefined, newImage);
            if (result.success) {
                closeAddModal();
                setPage(0);
                loadEvents();
            } else {
                setSubmitError(result.errorMessage);
                setSubmitState('error');
            }
        } catch (error) {
            setSubmitError((error as Error).message);
            setSubmitState('error');
        }
    };

    /** Bkz. AnnouncementsFeed.handleDeleteConfirm — aynı onay-sonra-sil deseni. */
    const handleDeleteConfirm = async (): Promise<void> => {
        if (!selected) {
            return;
        }
        setDeleteState('sending');
        try {
            const result = await deleteEvent(context, selected.id);
            if (result.success) {
                setIsDeleteConfirmOpen(false);
                setSelected(undefined);
                setDeleteState('idle');
                loadEvents();
            } else {
                setDeleteState('error');
            }
        } catch {
            setDeleteState('error');
        }
    };

    const styles = mergeStyleSets({
        // NOT: "gap" burada kullanılmıyor (flex "gap" desteklenmiyor) — satırlar
        // arası boşluk, tekrarlanan `row` sınıfının kendi marginBottom'ıyla veriliyor.
        list: {
            display: 'flex',
            flexDirection: 'column'
        },
        // "Takvim Yaprağı + Detay" kartı — eski düz metin + mavi hap tasarımının
        // yerini alıyor. Zeminde çok hafif bir gri (hover'da biraz koyulaşır) +
        // yumuşak bir geçiş. NOT: "gap" burada da kullanılmıyor — calendarLeaf'e
        // marginRight veriliyor (detailGroup ile chevron arası zaten chevron'un
        // marginLeft:'auto'suyla ayrılıyor).
        row: {
            display: 'flex',
            alignItems: 'center',
            padding: '10px 12px',
            marginBottom: 10,
            borderRadius: 10,
            background: 'transparent',
            cursor: 'pointer',
            border: 'none',
            width: '100%',
            textAlign: 'left',
            font: 'inherit',
            transition: 'background 0.18s ease',
            selectors: {
                ':hover': { background: '#F9FAFB' }
            }
        },
        // Takvim yaprağı: üstte renkli ay şeridi, altta büyük gün rakamı,
        // arkasında hafif bir gölge — bir masa takvimi yaprağını andırır.
        // Ay/gün satırları arasında net bir sınır oluşsun diye her ikisinde
        // de SABİT piksel line-height kullanılıyor (relative (1.1 gibi)
        // line-height, kalın/büyük punto rakamlarda fontun iç metrikleriyle
        // birlikte satırı olduğundan dar hesaplayıp üstteki ay şeridine çok
        // yakın/bindirilmiş görünmesine yol açabiliyordu).
        calendarLeaf: {
            display: 'flex',
            flexDirection: 'column',
            width: 52,
            borderRadius: 8,
            overflow: 'hidden',
            boxShadow: '0 2px 6px rgba(0,0,0,0.14)',
            flexShrink: 0,
            marginRight: 14
        },
        calendarLeafMonth: {
            // Kurumsal mavi kimlikle (yorpasTheme #0078d4 ailesi) tutarlı
            // koyu lacivert — önceki rastgele/marka dışı bordo (#9a3b47)
            // yerine.
            background: '#12395E',
            color: '#ffffff',
            fontSize: 11,
            fontWeight: 700,
            textAlign: 'center',
            padding: '5px 0',
            lineHeight: '14px',
            letterSpacing: 0.6
        },
        calendarLeafDay: {
            background: '#ffffff',
            color: theme.semanticColors.bodyText,
            fontSize: 20,
            fontWeight: 800,
            textAlign: 'center',
            padding: '5px 0 7px',
            lineHeight: '26px'
        },
        detailGroup: {
            display: 'flex',
            flexDirection: 'column',
            minWidth: 0,
            flexGrow: 1
        },
        // NOT: "gap" burada kullanılmıyor — eventTitleText'e marginBottom verildi.
        eventTitleText: {
            fontSize: 14,
            fontWeight: 700,
            color: theme.semanticColors.bodyText,
            overflow: 'hidden',
            textOverflow: 'ellipsis',
            whiteSpace: 'nowrap',
            marginBottom: 2
        },
        eventTimeText: {
            fontSize: 12,
            color: theme.semanticColors.bodySubtext
        },
        chevron: {
            marginLeft: 'auto',
            color: theme.semanticColors.bodySubtext,
            fontSize: 12,
            flexShrink: 0
        },
        emptyHint: {
            fontSize: 12,
            color: theme.semanticColors.bodySubtext
        },
        // Bkz. AnnouncementsFeed.tsx'teki aynı pagination/paginationPrevButton/
        // pageLabel deseni — sayfa başına PAGE_SIZE adet + ileri/geri oklar.
        pagination: {
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            marginTop: 8
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
        modalDate: {
            fontSize: 12,
            color: theme.semanticColors.bodySubtext,
            marginBottom: 4
        },
        modalLocation: {
            fontSize: 13,
            color: theme.semanticColors.bodyText,
            display: 'flex',
            alignItems: 'center'
        },
        // NOT: "gap" burada kullanılmıyor — ikondan sonra gelen metin düz bir
        // text node olduğu için ikona doğrudan marginRight veriliyor.
        modalLocationIcon: {
            marginRight: 6
        },
        // ÖNCEKİ HATA: "Etkinlik Detayı" hiç gösterilmiyordu — bkz. dosyanın
        // başındaki description alanı notu.
        modalDescription: {
            fontSize: 14,
            // ÖNCEKİ HATA: birimsiz (unitless) 1.6 bu render ortamında "1.6px" olarak
            // hesaplanıyor (birim otomatik ekleniyor) — satırlar neredeyse sıfır
            // yükseklikte üst üste biniyordu. Yüzde string'i açık birim taşıdığı için
            // bu hataya düşmüyor (bkz. DetailModal.tsx'teki aynı not).
            lineHeight: '160%',
            color: theme.semanticColors.bodyText,
            marginTop: 12
        },
        // ÖNCEKİ HATA: objectFit:'cover' + maxHeight kullanılıyordu — görsel
        // kutunun oranına uymadığında (ör. dikey/kare bir fotoğraf) kenarlarından
        // KIRPILIYORDU ("etkinliğe görsel koyunca kesiliyor" geri bildirimi).
        // 'contain' ile görsel HİÇBİR ZAMAN kırpılmıyor — kutuya sığacak şekilde
        // küçültülüyor, oranı farklıysa üstte/altta veya yanlarda kalan boşluk
        // düz bir zeminle (theme.palette.neutralLighterAlt, kartın diğer
        // "chip" zeminleriyle aynı ton) dolduruluyor.
        modalImage: {
            width: '100%',
            height: 220,
            objectFit: 'contain',
            background: theme.palette.neutralLighterAlt,
            borderRadius: '12px 12px 0 0',
            marginBottom: 16,
            display: 'block'
        },
        // NOT: "gap" burada kullanılmıyor — her çocuğa (formErrorBar,
        // inputFieldStyles.root, formRow, dropzoneWrapper) kendi marginBottom'ı
        // veriliyor; her zaman son çocuk olan formActions'a gerek yok.
        formContainer: {
            display: 'flex',
            flexDirection: 'column'
        },
        formErrorBar: {
            marginBottom: 20
        },
        // alignItems: 'flex-start' BİLİNÇLİ — DatePicker (kendi label'ını
        // kendi render eder) ile yandaki manuel <Label>+<input> grubu farklı
        // iç yüksekliklere sahip olabilir; 'stretch' (varsayılan) bu durumda
        // biri diğerini büyütüp hizaların/kutuların birbirine yapışmış gibi
        // görünmesine yol açabiliyordu. flex-start her sütunu SADECE kendi
        // içeriği kadar yükseklikte tutar, iki sütun da kendi doğal boyunda
        // üstten hizalı durur. NOT: "gap" burada da kullanılmıyor — her iki
        // formRowItem'a (Tarih/Saat sütunları) marginRight veriliyor.
        formRow: {
            display: 'flex',
            alignItems: 'flex-start',
            marginBottom: 20
        },
        formRowItem: {
            display: 'flex',
            flexDirection: 'column',
            flexGrow: 1,
            minWidth: 0,
            marginRight: 14
        },
        // "Saat" — DatePicker'ın kendi label'ıyla AYNI görsel ölçülerde
        // (fontSize 14, fontWeight 600, alt boşluk) — iki sütunun etiketleri
        // farklı boyda olursa alttaki kutular hizasız/çakışık görünüyordu.
        timeLabel: {
            fontSize: 14,
            fontWeight: 600,
            color: '#323130',
            marginBottom: 8,
            display: 'block'
        },
        // Fluent'in temel paketinde bir TimePicker bileşeni yok — ek bir
        // bağımlılık eklemek yerine yerel <input type="time"> kullanılıyor.
        timeInput: {
            width: '100%',
            boxSizing: 'border-box',
            height: 36,
            borderRadius: 12,
            background: '#F8FAFC',
            border: '1px solid #CBD5E1',
            boxShadow: 'inset 0 1px 2px rgba(15,23,42,0.05)',
            padding: '0 16px',
            fontSize: 14,
            fontFamily: 'inherit',
            selectors: {
                ':focus': { outline: 'none', border: '2px solid #3B82F6', boxShadow: '0 0 0 4px rgba(59,130,246,0.15)' }
            }
        },
        // Dış <label> SADECE tıklama alanını/erişilebilirlik bağını (htmlFor)
        // taşır; gerçek flex/hizalama düzeni İÇTEKİ <div> (imageDropzone)
        // üzerinde. Native <label> etiketleri bazı global sayfa/tema
        // CSS'lerinde (ör. Office UI Fabric Core'un form reset'i) "display"
        // için kendi kuralına sahip olabiliyor ve mergeStyleSets'in ürettiği
        // tek sınıfı ezerek ikon ile metnin üst üste binmesine yol açabiliyor
        // — bu yüzden flex layout'u hiçbir zaman <label>'a değil, her zaman
        // sıradan bir <div>'e uygulanıyor.
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
        formErrorDetail: {
            fontSize: 11,
            color: theme.semanticColors.bodySubtext,
            marginTop: 6,
            fontFamily: 'Consolas, monospace',
            wordBreak: 'break-word'
        }
    });

    const pageCount = Math.max(1, Math.ceil(events.length / PAGE_SIZE));
    const pagedEvents = events.slice(page * PAGE_SIZE, page * PAGE_SIZE + PAGE_SIZE);

    return (
        <WidgetCard
            title="Yaklaşan Etkinlikler"
            subtitle="Şirket etkinlik takvimi"
            iconName="Calendar"
            headerAction={canManageAnnouncements && (
                <IconButton
                    iconProps={{ iconName: 'Add' }}
                    ariaLabel="Yeni etkinlik ekle"
                    title="Yeni etkinlik ekle"
                    onClick={() => setIsAddOpen(true)}
                />
            )}
        >
            {state === 'loading' && <Spinner size={SpinnerSize.small} label="Etkinlikler yükleniyor..." />}
            {state === 'error' && <MessageBar messageBarType={MessageBarType.error}>{DATA_UNAVAILABLE_MESSAGE}</MessageBar>}
            {state === 'loaded' && events.length === 0 && (
                <Text className={styles.emptyHint}>Yaklaşan bir etkinlik bulunmuyor.</Text>
            )}
            {state === 'loaded' && events.length > 0 && (
                <>
                    <div className={styles.list}>
                        {pagedEvents.map((ev) => (
                            <button key={ev.id} type="button" className={styles.row} onClick={() => setSelected(ev)}>
                                <div className={styles.calendarLeaf}>
                                    <div className={styles.calendarLeafMonth}>{ev.monthShort}</div>
                                    <div className={styles.calendarLeafDay}>{ev.day}</div>
                                </div>
                                <div className={styles.detailGroup}>
                                    <div className={styles.eventTitleText}>{ev.title}</div>
                                    <div className={styles.eventTimeText}>{ev.timeLabel}</div>
                                </div>
                                <Icon iconName="ChevronRight" className={styles.chevron} />
                            </button>
                        ))}
                    </div>
                    {/* Kullanıcı isteğiyle: sayfa başına sabit (PAGE_SIZE) adet + ileri/geri
                        oklarıyla gezinme — bkz. AnnouncementsFeed.tsx'teki aynı desen. */}
                    {events.length > PAGE_SIZE && (
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
                deleteAriaLabel="Etkinliği sil"
            >
                {selected && (
                    <>
                        {selected.imageUrl && (
                            <img src={selected.imageUrl} alt="" className={styles.modalImage} />
                        )}
                        <div className={styles.modalDate}>{selected.dateLabel} · {selected.timeLabel}</div>
                        {selected.location && (
                            <div className={styles.modalLocation}>
                                <Icon iconName="MapPin" className={styles.modalLocationIcon} />
                                {selected.location}
                            </div>
                        )}
                        {selected.description && (
                            <div className={styles.modalDescription}>
                                {/* ÖNCEKİ HATA: white-space:pre-wrap ile tek bir <div> içine
                                    konan çok satırlı metin, bu render ortamında satırları
                                    üst üste bindiriyordu (\n karakterleri düzgün satır
                                    sonu üretmiyordu). Her satırı AYRI bir <div> (garanti
                                    blok düzeyinde) olarak basmak bu sorunu kesin çözüyor. */}
                                {selected.description.split('\n').map((line, index) => (
                                    <div key={index}>{line || ' '}</div>
                                ))}
                            </div>
                        )}
                    </>
                )}
            </DetailModal>

            <Dialog
                hidden={!isDeleteConfirmOpen}
                onDismiss={() => { setIsDeleteConfirmOpen(false); setDeleteState('idle'); }}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: 'Etkinlik silinsin mi?',
                    subText: `"${selected?.title ?? ''}" başlıklı etkinlik kalıcı olarak silinecek. Bu işlem geri alınamaz.`
                }}
            >
                {deleteState === 'error' && (
                    <MessageBar messageBarType={MessageBarType.error}>Etkinlik silinemedi. Lütfen tekrar deneyin.</MessageBar>
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

            <DetailModal isOpen={isAddOpen} title="Yeni Etkinlik Ekle" onDismiss={closeAddModal}>
                <div className={styles.formContainer}>
                    {submitState === 'error' && (
                        <MessageBar className={styles.formErrorBar} messageBarType={MessageBarType.error} onDismiss={() => setSubmitState('idle')}>
                            Etkinlik eklenemedi.
                            {submitError && <div className={styles.formErrorDetail}>{submitError}</div>}
                        </MessageBar>
                    )}
                    <TextField
                        label="Etkinlik Başlığı"
                        value={newTitle}
                        onChange={(_, v) => setNewTitle(v ?? '')}
                        disabled={submitState === 'sending'}
                        styles={topLevelFieldStyles}
                    />
                    <div className={styles.formRow}>
                        <div className={styles.formRowItem}>
                            <DatePicker
                                label="Tarih"
                                value={newDate}
                                onSelectDate={(d) => setNewDate(d ?? undefined)}
                                firstDayOfWeek={DayOfWeek.Monday}
                                strings={DAY_PICKER_STRINGS}
                                formatDate={(d) => (d ? d.toLocaleDateString('tr-TR') : '')}
                                disabled={submitState === 'sending'}
                                textField={{ styles: dateFieldStyles }}
                            />
                        </div>
                        <div className={styles.formRowItem}>
                            {/* Fluent'in <Label> bileşeni className prop'unu render'a
                                yansıtmıyordu (DOM'da hep varsayılan "ms-Label root-418"
                                çıkıyordu, özel stilim hiç uygulanmamıştı) — bu yüzden
                                garanti çalışan düz bir <label> kullanılıyor. */}
                            <label className={styles.timeLabel}>Saat</label>
                            <input
                                type="time"
                                className={styles.timeInput}
                                value={newTime}
                                onChange={(e) => setNewTime(e.target.value)}
                                disabled={submitState === 'sending'}
                            />
                        </div>
                    </div>
                    <TextField
                        label="Etkinlik Detayı"
                        multiline
                        rows={4}
                        value={newDescription}
                        onChange={(_, v) => setNewDescription(v ?? '')}
                        disabled={submitState === 'sending'}
                        styles={topLevelFieldStyles}
                    />
                    <label className={styles.dropzoneWrapper} htmlFor="event-image-input">
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
                            id="event-image-input"
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
                                    disabled={!newTitle.trim() || !newDate}
                                />
                            </>
                        )}
                    </div>
                </div>
            </DetailModal>
        </WidgetCard>
    );
};

export default UpcomingEvents;
