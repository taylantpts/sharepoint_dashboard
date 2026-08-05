import * as React from 'react';
import {
    IconButton, Spinner, SpinnerSize, MessageBar, MessageBarType, TextField, PrimaryButton,
    DefaultButton, Dialog, DialogType, DialogFooter, DatePicker, DayOfWeek, Checkbox, Dropdown, IDropdownOption,
    useTheme, mergeStyleSets, ITextFieldStyles
} from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import WidgetCard from '../WidgetCard';
import DetailModal from '../DetailModal';
import { usePermissions } from '../../hooks/usePermissions';
import { DATA_UNAVAILABLE_MESSAGE, SUBMIT_UNAVAILABLE_MESSAGE } from '../../constants';
import {
    getOnboardingRecords, createOnboardingRecord, updateOnboardingRecord,
    OnboardingKind, IOnboardingRecord, IChecklistItem
} from '../../services/OnboardingService';

export interface IOnboardingTrackerWidgetProps {
    context: WebPartContext;
}

type LoadState = 'loading' | 'loaded' | 'error';
type SubmitState = 'idle' | 'sending' | 'error';

const PAGE_SIZE = 8;

const LOCATION_OPTIONS: IDropdownOption[] = [
    { key: 'Maltepe Ofis', text: 'Maltepe Ofis' },
    { key: 'Gebze Fabrika', text: 'Gebze Fabrika' },
    { key: 'Franchise', text: 'Franchise' }
];

const TRANSFER_OPTIONS: IDropdownOption[] = [
    { key: 'Evet', text: 'Evet' },
    { key: 'Hayır', text: 'Hayır' }
];

const STATUS_OPTIONS: IDropdownOption[] = [
    { key: 'DEVAM EDİYOR', text: 'Devam Ediyor' },
    { key: 'TAMAMLANDI', text: 'Tamamlandı' }
];

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

const inputFieldStyles: Partial<ITextFieldStyles> = {
    root: { marginBottom: 16 },
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

const OnboardingTrackerWidget: React.FunctionComponent<IOnboardingTrackerWidgetProps> = (props) => {
    const { context } = props;
    const theme = useTheme();
    const { canManageOnboarding, canEditOnboarding } = usePermissions(context);

    const [kind, setKind] = React.useState<OnboardingKind>('katilis');
    const [state, setState] = React.useState<LoadState>('loading');
    const [records, setRecords] = React.useState<IOnboardingRecord[]>([]);
    const [page, setPage] = React.useState(0);

    const [isAddOpen, setIsAddOpen] = React.useState(false);
    const [newName, setNewName] = React.useState('');
    const [newTitle, setNewTitle] = React.useState('');
    const [newManager, setNewManager] = React.useState('');
    const [newLocation, setNewLocation] = React.useState<string>('Maltepe Ofis');
    const [newDate, setNewDate] = React.useState<Date | undefined>(undefined);
    const [newTransfer, setNewTransfer] = React.useState<string>('Hayır');
    const [newDescription, setNewDescription] = React.useState('');
    const [submitState, setSubmitState] = React.useState<SubmitState>('idle');
    const [submitError, setSubmitError] = React.useState<string | undefined>(undefined);

    const [selected, setSelected] = React.useState<IOnboardingRecord | undefined>(undefined);
    const [editChecklist, setEditChecklist] = React.useState<IChecklistItem[]>([]);
    const [editStatus, setEditStatus] = React.useState<string>('DEVAM EDİYOR');
    const [editState, setEditState] = React.useState<SubmitState>('idle');

    const loadRecords = React.useCallback((activeKind: OnboardingKind, onlyIfMounted?: () => boolean): void => {
        setState('loading');
        getOnboardingRecords(context, activeKind)
            .then((result) => {
                if (!onlyIfMounted || onlyIfMounted()) {
                    setRecords(result);
                    setState('loaded');
                    setPage(0);
                }
            })
            .catch((error: Error) => {
                console.error('[OnboardingTrackerWidget] Kayıtlar alınamadı:', error);
                if (!onlyIfMounted || onlyIfMounted()) {
                    setState('error');
                }
            });
    }, [context]);

    React.useEffect(() => {
        let isMounted = true;
        loadRecords(kind, () => isMounted);
        return () => {
            isMounted = false;
        };
    }, [kind, loadRecords]);

    const closeAddModal = (): void => {
        setIsAddOpen(false);
        setNewName('');
        setNewTitle('');
        setNewManager('');
        setNewLocation('Maltepe Ofis');
        setNewDate(undefined);
        setNewTransfer('Hayır');
        setNewDescription('');
        setSubmitState('idle');
        setSubmitError(undefined);
    };

    const handleAddSubmit = async (): Promise<void> => {
        if (!newName.trim() || !newDate) {
            return;
        }
        setSubmitState('sending');
        setSubmitError(undefined);
        try {
            const result = await createOnboardingRecord(context, kind, {
                name: newName.trim(),
                title: newTitle.trim(),
                manager: newManager.trim(),
                location: newLocation,
                date: newDate,
                transfer: newTransfer,
                description: newDescription.trim()
            });
            if (result.success) {
                closeAddModal();
                loadRecords(kind);
            } else {
                setSubmitError(result.errorMessage);
                setSubmitState('error');
            }
        } catch (error) {
            setSubmitError((error as Error).message);
            setSubmitState('error');
        }
    };

    const openEdit = (record: IOnboardingRecord): void => {
        setSelected(record);
        setEditChecklist(record.checklist.map((c) => ({ ...c })));
        setEditStatus(record.status || 'DEVAM EDİYOR');
        setEditState('idle');
    };

    const closeEdit = (): void => {
        setSelected(undefined);
        setEditState('idle');
    };

    const handleEditSave = async (): Promise<void> => {
        if (!selected) {
            return;
        }
        setEditState('sending');
        try {
            const result = await updateOnboardingRecord(context, kind, selected.id, editChecklist, editStatus);
            if (result.success) {
                closeEdit();
                loadRecords(kind);
            } else {
                setEditState('error');
            }
        } catch {
            setEditState('error');
        }
    };

    const styles = mergeStyleSets({
        modeToggleRow: {
            display: 'flex',
            background: theme.palette.neutralLighterAlt,
            borderRadius: 10,
            padding: 3,
            marginBottom: 16,
            maxWidth: 320
        },
        modeButton: {
            flexGrow: 1,
            border: 'none',
            background: 'transparent',
            borderRadius: 8,
            padding: '8px 0',
            fontSize: 13,
            fontWeight: 600,
            color: theme.semanticColors.bodySubtext,
            cursor: 'pointer',
            transition: 'background 0.15s ease, color 0.15s ease'
        },
        modeButtonActive: {
            background: theme.palette.white,
            color: theme.palette.themePrimary,
            boxShadow: '0 1px 4px rgba(0,0,0,0.12)'
        },
        table: {
            width: '100%',
            borderCollapse: 'collapse'
        },
        headerRow: {
            display: 'flex',
            alignItems: 'center',
            padding: '0 12px 10px',
            borderBottom: `1px solid ${theme.palette.neutralLighter}`,
            marginBottom: 4
        },
        row: {
            display: 'flex',
            alignItems: 'center',
            padding: '10px 12px',
            borderRadius: 10,
            transition: 'background 0.15s ease',
            selectors: {
                ':hover': { background: theme.palette.neutralLighterAlt }
            }
        },
        colName: { flex: '1 1 220px', minWidth: 0, fontSize: 13, fontWeight: 600, color: theme.semanticColors.bodyText, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' },
        colSub: { flex: '1 1 200px', minWidth: 0, fontSize: 12, color: theme.semanticColors.bodySubtext, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' },
        colLocation: { flex: '0 0 130px', fontSize: 12 },
        colDate: { flex: '0 0 100px', fontSize: 12, color: theme.semanticColors.bodySubtext },
        colStatus: { flex: '0 0 130px' },
        colEdit: { flex: '0 0 36px', display: 'flex', justifyContent: 'flex-end' },
        headerLabel: { fontSize: 11, fontWeight: 700, color: theme.semanticColors.bodySubtext, textTransform: 'uppercase', letterSpacing: 0.4 },
        locationBadge: {
            display: 'inline-block',
            padding: '3px 10px',
            borderRadius: 20,
            fontSize: 11,
            fontWeight: 600,
            background: theme.palette.neutralLighterAlt,
            color: theme.semanticColors.bodyText
        },
        statusBadge: {
            display: 'inline-block',
            padding: '3px 10px',
            borderRadius: 20,
            fontSize: 11,
            fontWeight: 700
        },
        emptyHint: {
            fontSize: 12,
            color: theme.semanticColors.bodySubtext,
            padding: '12px'
        },
        pagination: {
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            marginTop: 12
        },
        paginationPrevButton: { marginRight: 8 },
        pageLabel: { fontSize: 11, color: theme.semanticColors.bodySubtext, minWidth: 70, textAlign: 'center', marginRight: 8 },
        formContainer: { display: 'flex', flexDirection: 'column' },
        formErrorBar: { marginBottom: 16 },
        formRow: { display: 'flex', alignItems: 'flex-start' },
        formRowItem: { flexGrow: 1, minWidth: 0, marginRight: 14 },
        formActions: { display: 'flex', justifyContent: 'flex-end' },
        cancelButton: { marginRight: 12 },
        checklistGrid: {
            display: 'grid',
            gridTemplateColumns: 'repeat(2, 1fr)',
            gap: 12,
            marginBottom: 20
        },
        editMeta: {
            fontSize: 12,
            color: theme.semanticColors.bodySubtext,
            marginBottom: 16
        }
    });

    const statusStyle = (status: string): React.CSSProperties =>
        status === 'TAMAMLANDI'
            ? { color: '#0f7b3c', background: 'rgba(15,123,60,0.10)' }
            : { color: '#b45309', background: 'rgba(180,83,9,0.10)' };

    const pageCount = Math.max(1, Math.ceil(records.length / PAGE_SIZE));
    const pagedRecords = records.slice(page * PAGE_SIZE, page * PAGE_SIZE + PAGE_SIZE);

    return (
        <WidgetCard
            title="Katılış & Ayrılış Takibi"
            subtitle="Yeni işe giriş ve ayrılış süreçlerinin durumu"
            iconName="People"
            headerAction={canManageOnboarding && (
                <IconButton
                    iconProps={{ iconName: 'Add' }}
                    ariaLabel={kind === 'katilis' ? 'Yeni katılış ekle' : 'Yeni ayrılış ekle'}
                    title={kind === 'katilis' ? 'Yeni katılış ekle' : 'Yeni ayrılış ekle'}
                    onClick={() => setIsAddOpen(true)}
                />
            )}
        >
            <div className={styles.modeToggleRow}>
                <button
                    type="button"
                    className={`${styles.modeButton} ${kind === 'katilis' ? styles.modeButtonActive : ''}`}
                    onClick={() => setKind('katilis')}
                >
                    Katılış
                </button>
                <button
                    type="button"
                    className={`${styles.modeButton} ${kind === 'ayrilis' ? styles.modeButtonActive : ''}`}
                    onClick={() => setKind('ayrilis')}
                >
                    Ayrılış
                </button>
            </div>

            {state === 'loading' && <Spinner size={SpinnerSize.medium} label="Kayıtlar yükleniyor..." />}
            {state === 'error' && <MessageBar messageBarType={MessageBarType.error}>{DATA_UNAVAILABLE_MESSAGE}</MessageBar>}
            {state === 'loaded' && records.length === 0 && <div className={styles.emptyHint}>Görüntülenecek kayıt bulunmuyor.</div>}
            {state === 'loaded' && records.length > 0 && (
                <>
                    <div className={styles.headerRow}>
                        <span className={`${styles.colName} ${styles.headerLabel}`}>{kind === 'katilis' ? 'Çalışan' : 'Ayrılan Personel'}</span>
                        <span className={`${styles.colSub} ${styles.headerLabel}`}>{kind === 'katilis' ? 'Unvan / Yönetici' : 'Devir mi?'}</span>
                        <span className={`${styles.colLocation} ${styles.headerLabel}`}>Lokasyon</span>
                        <span className={`${styles.colDate} ${styles.headerLabel}`}>Tarih</span>
                        <span className={`${styles.colStatus} ${styles.headerLabel}`}>Durum</span>
                        <span className={styles.colEdit} />
                    </div>
                    {pagedRecords.map((record) => (
                        <div key={record.id} className={styles.row}>
                            <span className={styles.colName}>{record.name}</span>
                            <span className={styles.colSub}>
                                {kind === 'katilis' ? [record.title, record.manager].filter(Boolean).join(' · ') : record.transfer}
                            </span>
                            <span className={styles.colLocation}><span className={styles.locationBadge}>{record.location || '—'}</span></span>
                            <span className={styles.colDate}>{record.dateLabel}</span>
                            <span className={styles.colStatus}>
                                <span className={styles.statusBadge} style={statusStyle(record.status)}>
                                    {record.status === 'TAMAMLANDI' ? 'Tamamlandı' : 'Devam Ediyor'}
                                </span>
                            </span>
                            <span className={styles.colEdit}>
                                {canEditOnboarding && (
                                    <IconButton
                                        iconProps={{ iconName: 'Edit' }}
                                        ariaLabel="Kaydı düzenle"
                                        title="Kaydı düzenle"
                                        onClick={() => openEdit(record)}
                                    />
                                )}
                            </span>
                        </div>
                    ))}
                    {records.length > PAGE_SIZE && (
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

            <DetailModal isOpen={isAddOpen} title={kind === 'katilis' ? 'Yeni Katılış Ekle' : 'Yeni Ayrılış Ekle'} onDismiss={closeAddModal}>
                <div className={styles.formContainer}>
                    {submitState === 'error' && (
                        <MessageBar className={styles.formErrorBar} messageBarType={MessageBarType.error} onDismiss={() => setSubmitState('idle')}>
                            {SUBMIT_UNAVAILABLE_MESSAGE}
                            {submitError && <div>{submitError}</div>}
                        </MessageBar>
                    )}
                    <TextField
                        label={kind === 'katilis' ? 'Çalışan Adı Soyadı' : 'Ayrılan Personel'}
                        value={newName}
                        onChange={(_, v) => setNewName(v ?? '')}
                        disabled={submitState === 'sending'}
                        styles={inputFieldStyles}
                    />
                    {kind === 'katilis' && (
                        <>
                            <TextField
                                label="Unvan"
                                value={newTitle}
                                onChange={(_, v) => setNewTitle(v ?? '')}
                                disabled={submitState === 'sending'}
                                styles={inputFieldStyles}
                            />
                            <TextField
                                label="Yönetici"
                                value={newManager}
                                onChange={(_, v) => setNewManager(v ?? '')}
                                disabled={submitState === 'sending'}
                                styles={inputFieldStyles}
                            />
                        </>
                    )}
                    <div className={styles.formRow}>
                        <div className={styles.formRowItem}>
                            <Dropdown
                                label="Lokasyon"
                                selectedKey={newLocation}
                                options={LOCATION_OPTIONS}
                                onChange={(_, option) => option && setNewLocation(option.key as string)}
                                disabled={submitState === 'sending'}
                                styles={{ root: { marginBottom: 16 } }}
                            />
                        </div>
                        <div className={styles.formRowItem} style={{ marginRight: 0 }}>
                            <Dropdown
                                label="Devir mi?"
                                selectedKey={newTransfer}
                                options={TRANSFER_OPTIONS}
                                onChange={(_, option) => option && setNewTransfer(option.key as string)}
                                disabled={submitState === 'sending'}
                                styles={{ root: { marginBottom: 16 } }}
                            />
                        </div>
                    </div>
                    <DatePicker
                        label={kind === 'katilis' ? 'Başlama Tarihi' : 'Ayrılma Tarihi'}
                        value={newDate}
                        onSelectDate={(d) => setNewDate(d ?? undefined)}
                        firstDayOfWeek={DayOfWeek.Monday}
                        strings={DAY_PICKER_STRINGS}
                        formatDate={(d) => (d ? d.toLocaleDateString('tr-TR') : '')}
                        disabled={submitState === 'sending'}
                        textField={{ styles: inputFieldStyles }}
                    />
                    {kind === 'katilis' && (
                        <TextField
                            label="Açıklama"
                            multiline
                            rows={3}
                            value={newDescription}
                            onChange={(_, v) => setNewDescription(v ?? '')}
                            disabled={submitState === 'sending'}
                            styles={inputFieldStyles}
                        />
                    )}
                    <div className={styles.formActions}>
                        {submitState === 'sending' ? (
                            <Spinner size={SpinnerSize.small} label="Kaydediliyor..." />
                        ) : (
                            <>
                                <DefaultButton className={styles.cancelButton} text="Vazgeç" onClick={closeAddModal} />
                                <PrimaryButton
                                    text="Kaydet"
                                    onClick={() => { handleAddSubmit().catch(() => { /* handled in handleAddSubmit */ }); }}
                                    disabled={!newName.trim() || !newDate}
                                />
                            </>
                        )}
                    </div>
                </div>
            </DetailModal>

            <Dialog
                hidden={!selected}
                onDismiss={closeEdit}
                minWidth={480}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: selected?.name,
                    subText: kind === 'katilis' ? [selected?.title, selected?.manager].filter(Boolean).join(' · ') : selected?.location
                }}
            >
                {editState === 'error' && (
                    <MessageBar messageBarType={MessageBarType.error} styles={{ root: { marginBottom: 12 } }}>
                        {SUBMIT_UNAVAILABLE_MESSAGE}
                    </MessageBar>
                )}
                <div className={styles.editMeta}>Tamamlanan işlemleri işaretleyin ve durumu güncelleyin.</div>
                <div className={styles.checklistGrid}>
                    {editChecklist.map((item, index) => (
                        <Checkbox
                            key={item.internalName}
                            label={item.label}
                            checked={item.done}
                            onChange={(_, checked) => {
                                setEditChecklist((prev) => prev.map((c, i) => (i === index ? { ...c, done: !!checked } : c)));
                            }}
                            disabled={editState === 'sending'}
                        />
                    ))}
                </div>
                <Dropdown
                    label="Durum"
                    selectedKey={editStatus}
                    options={STATUS_OPTIONS}
                    onChange={(_, option) => option && setEditStatus(option.key as string)}
                    disabled={editState === 'sending'}
                    styles={{ root: { marginBottom: 8 } }}
                />
                <DialogFooter>
                    {editState === 'sending' ? (
                        <Spinner size={SpinnerSize.small} label="Kaydediliyor..." />
                    ) : (
                        <>
                            <PrimaryButton text="Kaydet" onClick={() => { handleEditSave().catch(() => { /* handled */ }); }} />
                            <DefaultButton text="Vazgeç" onClick={closeEdit} />
                        </>
                    )}
                </DialogFooter>
            </Dialog>
        </WidgetCard>
    );
};

export default OnboardingTrackerWidget;
