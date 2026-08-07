import * as React from 'react';
import {
    Icon, Spinner, SpinnerSize, MessageBar, MessageBarType, Checkbox, IconButton, TextField, PrimaryButton,
    useTheme, mergeStyleSets, ITextFieldStyles
} from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import WidgetCard from '../WidgetCard';
import DetailModal from '../DetailModal';
import {
    getTodayEvents, getTodoTasks, completeTodoTask, createTodoTask, ITodayEvent, ITodoTask
} from '../../services/GraphService';
import { DATA_UNAVAILABLE_MESSAGE } from '../../constants';

export interface IPersonalAssistantProps {
    context: WebPartContext;
}

type LoadState = 'loading' | 'loaded' | 'error';
type SubmitState = 'idle' | 'sending' | 'error';

// Form alanının ortak "yumuşak" görünümü — diğer ekleme modalleriyle (Duyuru/
// Etkinlik) aynı görsel dil.
const inputFieldStyles: Partial<ITextFieldStyles> = {
    // formContainer'daki "gap" yerine — formActions'tan önceki boşluk.
    root: { marginBottom: 20 },
    fieldGroup: {
        minHeight: 36,
        background: 'rgba(248,250,252,0.8)',
        border: '1px solid #E2E8F0',
        borderRadius: 12,
        selectors: {
            ':after': { borderRadius: 12, border: '2px solid #3B82F6', boxShadow: '0 0 0 4px rgba(59,130,246,0.15)' }
        }
    },
    field: { padding: '12px 16px', fontSize: 14 }
};

const PersonalAssistant: React.FunctionComponent<IPersonalAssistantProps> = (props) => {
    const { context } = props;
    const theme = useTheme();

    const [events, setEvents] = React.useState<ITodayEvent[]>([]);
    const [eventsState, setEventsState] = React.useState<LoadState>('loading');
    // Bitiş saati geçmiş toplantılar listeden düşsün diye "şu an" dakikada bir
    // yenileniyor — sayfa açık kaldığı sürece toplantı bitince otomatik kalkar,
    // sayfayı yenilemek gerekmez. HeaderClock'taki saniyelik tik deseniyle aynı
    // mantık, ama burada dakikalık hassasiyet yeterli (gereksiz re-render yok).
    const [now, setNow] = React.useState(() => Date.now());
    React.useEffect(() => {
        const interval = window.setInterval(() => setNow(Date.now()), 60000);
        return () => window.clearInterval(interval);
    }, []);
    const visibleEvents = events.filter((ev) => new Date(ev.endDateTime).getTime() > now);

    const [tasks, setTasks] = React.useState<ITodoTask[]>([]);
    const [tasksState, setTasksState] = React.useState<LoadState>('loading');
    // Boş "görev yok" ile gerçek bir izin/erişim hatası birbirine karışmasın
    // diye ayrı tutuluyor — dolu ise ekranda görünür bir uyarı gösterilir.
    const [tasksErrorDetail, setTasksErrorDetail] = React.useState<string | undefined>(undefined);
    // Liste BOŞ olsa bile (0 görev) yeni görev eklemek için gereken hedef
    // liste ID'si — tek bir görevin listId'sinden okumak mümkün olmayabilir.
    const [todoListId, setTodoListId] = React.useState<string | undefined>(undefined);
    // Tamamlanma isteği Graph'a giderken checkbox'ı devre dışı bırakmak için.
    const [completingIds, setCompletingIds] = React.useState<Set<string>>(new Set());

    const [isAddOpen, setIsAddOpen] = React.useState(false);
    const [newTaskTitle, setNewTaskTitle] = React.useState('');
    const [submitState, setSubmitState] = React.useState<SubmitState>('idle');
    const [submitError, setSubmitError] = React.useState<string | undefined>(undefined);

    const loadTasks = React.useCallback((onlyIfMounted?: () => boolean): void => {
        getTodoTasks(context)
            .then((result) => {
                if (!onlyIfMounted || onlyIfMounted()) {
                    setTasks(result.tasks);
                    setTasksErrorDetail(result.errorDetail);
                    setTodoListId(result.listId);
                    setTasksState('loaded');
                }
            })
            .catch(() => { /* getTodoTasks kendi içinde hatayı yönetir, buraya asla ulaşmaz */ });
    }, [context]);

    React.useEffect(() => {
        let isMounted = true;

        getTodayEvents(context)
            .then((result) => {
                if (isMounted) {
                    setEvents(result);
                    setEventsState('loaded');
                }
            })
            .catch((error: Error) => {
                console.error('[PersonalAssistant] Toplantılar alınamadı:', error);
                if (isMounted) {
                    setEventsState('error');
                }
            });

        loadTasks(() => isMounted);

        return () => {
            isMounted = false;
        };
    }, [context, loadTasks]);

    /**
     * Checkbox işaretlendiğinde GERÇEK Microsoft To Do hesabına yazar. İyimser
     * (optimistic) güncelleme: görev anında listeden kaybolur; Graph isteği
     * başarısız olursa görev geri eklenir ve konsola hata basılır.
     */
    const handleCompleteTask = (task: ITodoTask): void => {
        setCompletingIds((prev) => new Set(prev).add(task.id));
        setTasks((prev) => prev.filter((t) => t.id !== task.id));

        // completeTodoTask kendi içinde hatayı yakalar, asla reddetmez (reject).
        completeTodoTask(context, task.listId, task.id)
            .then((result) => {
                if (!result.success) {
                    console.error('[PersonalAssistant] Görev tamamlanamadı, listeye geri eklendi:', result.errorMessage);
                    setTasks((prev) => [task, ...prev]);
                }
                setCompletingIds((prev) => {
                    const next = new Set(prev);
                    next.delete(task.id);
                    return next;
                });
            })
            .catch(() => { /* completeTodoTask kendi içinde hatayı yönetir, buraya asla ulaşmaz */ });
    };

    const closeAddModal = (): void => {
        setIsAddOpen(false);
        setNewTaskTitle('');
        setSubmitState('idle');
        setSubmitError(undefined);
    };

    /**
     * "Kaydet" — GERÇEK Microsoft To Do hesabına yazar (POST /me/todo/lists/
     * {listId}/tasks), yerel bir state değil. Başarılı olursa liste Graph'tan
     * tekrar çekilir (loadTasks) ki eklenen görev anında, sunucudan doğrulanmış
     * hâliyle görünsün.
     */
    const handleAddTask = async (): Promise<void> => {
        if (!newTaskTitle.trim() || !todoListId) {
            return;
        }
        setSubmitState('sending');
        setSubmitError(undefined);

        try {
            const result = await createTodoTask(context, todoListId, newTaskTitle.trim());
            if (result.success) {
                closeAddModal();
                loadTasks();
            } else {
                setSubmitError(result.errorMessage);
                setSubmitState('error');
            }
        } catch (error) {
            setSubmitError((error as Error).message);
            setSubmitState('error');
        }
    };

    const styles = mergeStyleSets({
        columns: {
            display: 'grid',
            gridTemplateColumns: '1fr 1fr',
            gap: 20,
            selectors: {
                '@media (max-width: 600px)': {
                    gridTemplateColumns: '1fr'
                }
            }
        },
        // NOT: "gap" flex özelliği burada BİLİNÇLİ OLARAK KULLANILMIYOR — bu
        // sayfanın render edildiği (kurumsal/eski) tarayıcı ortamında flex "gap"
        // desteklenmiyor. columnTitle iki farklı bağlamda kullanılıyor (bazen tek
        // çocuklu, bazen "Yeni görev ekle" butonuyla iki çocuklu) — bu yüzden
        // marginRight, tüm kullanımlarda ortak olan columnTitleLabel'a verildi.
        columnTitle: {
            fontSize: 12,
            fontWeight: 600,
            color: theme.semanticColors.bodyText,
            marginBottom: 10,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between'
        },
        columnTitleLabel: {
            display: 'flex',
            alignItems: 'center',
            marginRight: 6
        },
        columnTitleIcon: {
            marginRight: 6
        },
        // NOT: "gap" burada da kullanılmıyor — eventTime'a marginRight verildi.
        // Toplantı satırı artık tıklanabilir bir <button> — Teams'te "katıl"
        // linki varsa oraya, yoksa etkinliği Outlook'ta açan webLink'e gider
        // (bkz. GraphService.ts ITodayEvent.link).
        eventRow: {
            display: 'flex',
            width: '100%',
            padding: '8px 0',
            borderWidth: '0 0 1px 0',
            borderStyle: 'solid',
            borderColor: theme.semanticColors.variantBorder,
            background: 'transparent',
            textAlign: 'left',
            font: 'inherit',
            cursor: 'pointer',
            borderRadius: 6,
            transition: 'background 0.15s ease',
            selectors: {
                ':hover': { background: theme.palette.neutralLighterAlt }
            }
        },
        eventTime: {
            fontSize: 11,
            fontWeight: 600,
            color: theme.palette.themePrimary,
            minWidth: 80,
            marginRight: 10
        },
        eventSubject: {
            fontSize: 12,
            fontWeight: 600,
            color: theme.semanticColors.bodyText
        },
        eventLocation: {
            fontSize: 11,
            color: theme.semanticColors.bodySubtext
        },
        // NOT: "gap" burada kullanılmıyor — taskRow'un tek çocuğu (Checkbox)
        // olduğu için zaten etkisizdi, doğrudan kaldırıldı.
        taskRow: {
            display: 'flex',
            alignItems: 'center',
            padding: '6px 0',
            borderBottom: `1px solid ${theme.semanticColors.variantBorder}`
        },
        taskTitle: {
            fontSize: 12,
            color: theme.semanticColors.bodyText
        },
        taskHint: {
            fontSize: 11,
            color: theme.semanticColors.bodySubtext,
            marginTop: 6,
            marginBottom: 4
        },
        emptyHint: {
            fontSize: 12,
            color: theme.semanticColors.bodySubtext
        },
        taskErrorHint: {
            fontSize: 11,
            color: theme.semanticColors.bodySubtext,
            marginTop: 6,
            fontFamily: 'Consolas, monospace'
        },
        // NOT: "gap" burada kullanılmıyor — hata MessageBar'ı KOŞULLU olarak
        // ilk sırada render edilebildiği için kendi marginBottom'ını
        // (formMessageBar) taşıyor; TextField ise inputFieldStyles.root üzerinden
        // marginBottom alıyor (her zaman render edilen, formActions'tan önceki çocuk).
        formContainer: {
            display: 'flex',
            flexDirection: 'column'
        },
        formMessageBar: {
            marginBottom: 20
        },
        formActions: {
            display: 'flex',
            justifyContent: 'flex-end'
        },
        formErrorDetail: {
            fontSize: 11,
            color: theme.semanticColors.bodySubtext,
            marginTop: 6,
            fontFamily: 'Consolas, monospace',
            wordBreak: 'break-word'
        }
    });

    return (
        <WidgetCard title="Kişisel Asistan" subtitle="Bugünkü toplantılarınız ve yapılacaklar listeniz" iconName="Calendar">
            <div className={styles.columns}>
                <div>
                    <div className={styles.columnTitle}>
                        <span className={styles.columnTitleLabel}><Icon iconName="Calendar" className={styles.columnTitleIcon} /> Bugünkü Toplantılar</span>
                    </div>
                    {eventsState === 'loading' && <Spinner size={SpinnerSize.small} label="Yükleniyor..." />}
                    {eventsState === 'error' && (
                        <MessageBar messageBarType={MessageBarType.error}>{DATA_UNAVAILABLE_MESSAGE}</MessageBar>
                    )}
                    {eventsState === 'loaded' && visibleEvents.length === 0 && (
                        <span className={styles.emptyHint}>Bugün için planlanmış toplantınız yok.</span>
                    )}
                    {eventsState === 'loaded' && visibleEvents.map((ev) => (
                        <button
                            key={ev.id}
                            type="button"
                            className={styles.eventRow}
                            onClick={() => window.open(ev.link, '_blank', 'noopener,noreferrer')}
                            title={`${ev.subject} — toplantıya git`}
                        >
                            <span className={styles.eventTime}>{ev.startLabel} - {ev.endLabel}</span>
                            <div>
                                <div className={styles.eventSubject}>{ev.subject}</div>
                                {ev.location && <div className={styles.eventLocation}>{ev.location}</div>}
                            </div>
                        </button>
                    ))}
                </div>

                <div>
                    <div className={styles.columnTitle}>
                        <span className={styles.columnTitleLabel}><Icon iconName="CheckList" className={styles.columnTitleIcon} /> Yapılacaklar</span>
                        {tasksState === 'loaded' && todoListId && (
                            <IconButton
                                iconProps={{ iconName: 'Add' }}
                                ariaLabel="Yeni görev ekle"
                                title="Yeni görev ekle"
                                onClick={() => setIsAddOpen(true)}
                            />
                        )}
                    </div>
                    {tasksState === 'loading' && <Spinner size={SpinnerSize.small} label="Yükleniyor..." />}
                    {tasksState === 'loaded' && tasksErrorDetail && (
                        <MessageBar messageBarType={MessageBarType.warning}>
                            Görevler alınamadı. Aşağıdaki teknik detayı destek ekibine iletin:
                            <div className={styles.taskErrorHint}>{tasksErrorDetail}</div>
                        </MessageBar>
                    )}
                    {tasksState === 'loaded' && !tasksErrorDetail && tasks.length === 0 && (
                        <span className={styles.emptyHint}>Açık göreviniz bulunmuyor.</span>
                    )}
                    {tasksState === 'loaded' && !tasksErrorDetail && tasks.length > 0 && (
                        <div className={styles.taskHint}>İşaretlediğinizde Microsoft To Do&apos;da da tamamlanır.</div>
                    )}
                    {tasksState === 'loaded' && tasks.map((task) => (
                        <div key={task.id} className={styles.taskRow}>
                            <Checkbox
                                label={task.title}
                                disabled={completingIds.has(task.id)}
                                onChange={() => handleCompleteTask(task)}
                                styles={{ text: styles.taskTitle }}
                            />
                        </div>
                    ))}
                </div>
            </div>

            <DetailModal isOpen={isAddOpen} title="Yeni Görev Ekle" onDismiss={closeAddModal}>
                <div className={styles.formContainer}>
                    {submitState === 'error' && (
                        <MessageBar
                            messageBarType={MessageBarType.error}
                            onDismiss={() => setSubmitState('idle')}
                            className={styles.formMessageBar}
                        >
                            Görev eklenemedi.
                            {submitError && <div className={styles.formErrorDetail}>{submitError}</div>}
                        </MessageBar>
                    )}
                    <TextField
                        label="Görev Adı"
                        value={newTaskTitle}
                        onChange={(_, v) => setNewTaskTitle(v ?? '')}
                        disabled={submitState === 'sending'}
                        styles={inputFieldStyles}
                    />
                    <div className={styles.formActions}>
                        {submitState === 'sending' ? (
                            <Spinner size={SpinnerSize.small} label="Kaydediliyor..." />
                        ) : (
                            <PrimaryButton
                                text="Kaydet"
                                onClick={() => { handleAddTask().catch(() => { /* handled in handleAddTask */ }); }}
                                disabled={!newTaskTitle.trim()}
                            />
                        )}
                    </div>
                </div>
            </DetailModal>
        </WidgetCard>
    );
};

export default PersonalAssistant;
