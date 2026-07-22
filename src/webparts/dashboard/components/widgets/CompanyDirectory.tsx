import * as React from 'react';
import { SearchBox, Persona, PersonaSize, Spinner, SpinnerSize, MessageBar, MessageBarType, Text, Icon, DefaultButton, IconButton, useTheme, mergeStyleSets } from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import WidgetCard from '../WidgetCard';
import DetailModal from '../DetailModal';
import { searchDirectory, IDirectoryUser } from '../../services/GraphService';
import { getBirthdaysThisMonth, IBirthdayItem } from '../../services/SharePointService';
import { DATA_UNAVAILABLE_MESSAGE } from '../../constants';

export interface ICompanyDirectoryProps {
    context: WebPartContext;
}

type LoadState = 'idle' | 'loading' | 'loaded' | 'error';

const SEARCH_DEBOUNCE_MS = 350;
const BIRTHDAYS_PAGE_SIZE = 10;

/**
 * SharePoint'in her sitede hazır bulunan userphoto.aspx sistem sayfasından
 * (aynı origin, ekstra Graph izni gerektirmez) bir kullanıcının M365 profil
 * fotoğrafı URL'ini üretir — bkz. WelcomeHeader.tsx'teki aynı desen (orada
 * oturum sahibinin kendi fotoğrafı için kullanılıyor, burada rehberde
 * seçilen HERHANGİ BİR kullanıcı için). E-posta yoksa undefined döner;
 * Persona bileşeni imageUrl'siz kaldığında zaten kendiliğinden baş harfli
 * rozete düşer.
 */
const getUserPhotoUrl = (webAbsoluteUrl: string, mail: string | undefined): string | undefined =>
    mail ? `${webAbsoluteUrl}/_layouts/15/userphoto.aspx?size=L&accountname=${encodeURIComponent(mail)}` : undefined;

const CompanyDirectory: React.FunctionComponent<ICompanyDirectoryProps> = (props) => {
    const { context } = props;
    const theme = useTheme();

    const [query, setQuery] = React.useState('');
    const [results, setResults] = React.useState<IDirectoryUser[]>([]);
    const [searchState, setSearchState] = React.useState<LoadState>('idle');
    const [selectedUser, setSelectedUser] = React.useState<IDirectoryUser | undefined>(undefined);

    const [birthdays, setBirthdays] = React.useState<IBirthdayItem[]>([]);
    const [birthdaysState, setBirthdaysState] = React.useState<LoadState>('loading');
    const [birthdayPage, setBirthdayPage] = React.useState(0);

    React.useEffect(() => {
        let isMounted = true;
        getBirthdaysThisMonth(context)
            .then((items) => {
                if (isMounted) {
                    setBirthdays(items);
                    setBirthdaysState('loaded');
                }
            })
            .catch((error: Error) => {
                console.error('[CompanyDirectory] Doğum günleri alınamadı:', error);
                if (isMounted) {
                    setBirthdaysState('error');
                }
            });
        return () => {
            isMounted = false;
        };
    }, [context]);

    React.useEffect(() => {
        const trimmed = query.trim();
        if (!trimmed) {
            setResults([]);
            setSearchState('idle');
            return undefined;
        }

        setSearchState('loading');

        const handle = setTimeout(() => {
            searchDirectory(context, trimmed)
                .then((users) => {
                    setResults(users);
                    setSearchState('loaded');
                })
                .catch((error: Error) => {
                    console.error('[CompanyDirectory] Rehber araması başarısız:', error);
                    setSearchState('error');
                });
        }, SEARCH_DEBOUNCE_MS);

        return () => clearTimeout(handle);
    }, [query, context]);

    const birthdayPageCount = Math.max(1, Math.ceil(birthdays.length / BIRTHDAYS_PAGE_SIZE));
    const pagedBirthdays = birthdays.slice(
        birthdayPage * BIRTHDAYS_PAGE_SIZE,
        birthdayPage * BIRTHDAYS_PAGE_SIZE + BIRTHDAYS_PAGE_SIZE
    );

    const styles = mergeStyleSets({
        // NOT: "gap" flex özelliği burada BİLİNÇLİ OLARAK KULLANILMIYOR — bu
        // sayfanın render edildiği (kurumsal/eski) tarayıcı ortamında flex "gap"
        // desteklenmiyor. .map() ile tekrar eden resultRow kendi marginBottom'ını
        // taşıyor (son öğede zararsız bir alt boşluk kalır — kabul edilebilir).
        resultList: {
            marginTop: 10,
            display: 'flex',
            flexDirection: 'column',
            maxHeight: 220,
            overflowY: 'auto'
        },
        resultRow: {
            display: 'flex',
            flexDirection: 'column',
            width: '100%',
            background: 'none',
            border: 'none',
            padding: '4px 0',
            borderRadius: 6,
            cursor: 'pointer',
            textAlign: 'left',
            font: 'inherit',
            transition: 'background 0.12s ease',
            marginBottom: 4,
            selectors: {
                ':hover': { background: theme.palette.themeLighterAlt }
            }
        },
        resultMeta: {
            fontSize: 11,
            color: theme.semanticColors.bodySubtext,
            marginLeft: 44,
            marginTop: -4
        },
        emptyHint: {
            fontSize: 12,
            color: theme.semanticColors.bodySubtext,
            marginTop: 10
        },
        // NOT: "gap" burada kullanılmıyor — sectionTitle'ın (bu dosyada) tek
        // çocuğu olduğu için zaten etkisizdi, doğrudan kaldırıldı. sectionTitleLabel
        // içindeki ikon+metin boşluğu ise sectionTitleIcon'un marginRight'ıyla veriliyor.
        sectionTitle: {
            fontSize: 12,
            fontWeight: 600,
            color: theme.semanticColors.bodyText,
            marginTop: 16,
            marginBottom: 6,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between'
        },
        sectionTitleLabel: {
            display: 'flex',
            alignItems: 'center'
        },
        sectionTitleIcon: {
            marginRight: 6
        },
        // personRow, hem tek çocuklu (arama sonucu satırı) hem iki çocuklu
        // (doğum günü satırı: Persona + rozet) bağlamlarda kullanılıyor — "gap"
        // yerine ortak avatar sınıfı (personRowAvatar) marginRight taşıyor; tek
        // çocuklu kullanımda bu marginRight zararsız (sağda görünmeyen boşluk).
        personRow: {
            display: 'flex',
            alignItems: 'center',
            padding: '3px 0'
        },
        personRowAvatar: {
            marginRight: 8
        },
        badge: {
            marginLeft: 'auto',
            fontSize: 10,
            fontWeight: 700,
            padding: '2px 8px',
            borderRadius: 10,
            background: theme.palette.neutralLighter,
            color: theme.semanticColors.bodySubtext,
            whiteSpace: 'nowrap'
        },
        todayBadge: {
            marginLeft: 'auto',
            fontSize: 10,
            fontWeight: 700,
            padding: '2px 8px',
            borderRadius: 10,
            background: theme.palette.themePrimary,
            color: '#ffffff',
            whiteSpace: 'nowrap'
        },
        // NOT: "gap" burada kullanılmıyor — "önceki sayfa" butonuna marginRight
        // veriliyor, pageLabel'e de sonraki (next) buton için marginRight.
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
            minWidth: 60,
            textAlign: 'center',
            marginRight: 8
        },
        // Kişi kartı (modal) içeriği — Avatar + metin bloğunu kapsayan en dış
        // kapsayıcı: temiz bir flex satırı, position:absolute/float YOK.
        // NOT: "gap" kullanılmıyor — Persona'ya (cardHeaderAvatar) marginRight verildi.
        cardHeader: {
            display: 'flex',
            alignItems: 'center',
            marginBottom: 16
        },
        cardHeaderAvatar: {
            marginRight: 16
        },
        // İsim + unvan grubu. GERÇEK HATA burada değil, Persona bileşimindeydi:
        // `hidePersonaDetails` verilmeden kullanılan bir <Persona text=.../>
        // KENDİ İÇİNDE de aynı ismi render ediyordu — bu yüzden buradaki metin
        // Persona'nın kendi metniyle ÇAKIŞIYOR/üst üste biniyormuş gibi
        // görünüyordu. Aşağıdaki JSX'te artık hidePersonaDetails var, Persona
        // SADECE fotoğrafı/rozeti gösteriyor, isim/unvan TEK yerden (burada)
        // geliyor. lineHeight'ler de DetailModal'ın miras bıraktığı 1.6'yı
        // ezmek için hâlâ açıkça tanımlı (bkz. RequiredDocuments.tsx'teki aynı
        // desen) — flex-column + sabit lineHeight, taşma/binişmeyi yapısal
        // olarak imkansız hale getiriyor.
        cardTextGroup: {
            display: 'flex',
            flexDirection: 'column',
            justifyContent: 'center',
            minWidth: 0
        },
        cardName: {
            display: 'block',
            fontWeight: 700,
            fontSize: 18, // 1.125rem
            lineHeight: '22px',
            color: '#1f2937',
            overflow: 'hidden',
            textOverflow: 'ellipsis',
            whiteSpace: 'nowrap'
        },
        cardJobTitle: {
            display: 'block',
            fontSize: 14, // 0.875rem
            lineHeight: '18px',
            color: '#6b7280',
            marginTop: 2
        },
        // NOT: "gap" kullanılmıyor — cardFieldIcon'a marginRight verildi.
        cardField: {
            display: 'flex',
            alignItems: 'center',
            padding: '8px 0',
            borderTop: `1px solid ${theme.palette.neutralLighter}`,
            fontSize: 13,
            color: theme.semanticColors.bodyText
        },
        cardFieldIcon: {
            color: theme.palette.themePrimary,
            fontSize: 14,
            width: 16,
            marginRight: 8
        },
        // cardActions'ın çocukları (E-posta/Teams/Ara butonları) HER BİRİ
        // ayrı ayrı koşullu render ediliyor — hangisinin "son" olacağı derleme
        // zamanında belli değil. Bu yüzden "gap" yerine, görünen HER butona
        // (cardActionButton) marginRight veriliyor; en sonda kalan buton için bu
        // zararsız bir sağ boşluk olarak kalır (kabul edilebilir tradeoff).
        cardActions: {
            marginTop: 16,
            display: 'flex'
        },
        cardActionButton: {
            marginRight: 8
        }
    });

    return (
        <WidgetCard title="Şirket Rehberi" subtitle="Kim Kimdir?" iconName="ContactCard" accentColor="#5c8fc4">
            <SearchBox
                placeholder="İsim, departman veya unvan ara..."
                value={query}
                onChange={(_, newValue) => setQuery(newValue ?? '')}
            />

            {searchState === 'loading' && <Spinner size={SpinnerSize.small} label="Aranıyor..." styles={{ root: { marginTop: 10 } }} />}
            {searchState === 'error' && (
                <MessageBar messageBarType={MessageBarType.error} styles={{ root: { marginTop: 10 } }}>
                    {DATA_UNAVAILABLE_MESSAGE}
                </MessageBar>
            )}

            {searchState === 'loaded' && (
                <div className={styles.resultList}>
                    {results.length === 0 && <Text className={styles.emptyHint}>Sonuç bulunamadı.</Text>}
                    {results.map((user) => (
                        <button
                            key={user.id}
                            type="button"
                            className={styles.resultRow}
                            onClick={() => setSelectedUser(user)}
                        >
                            <div className={styles.personRow}>
                                <Persona
                                    imageUrl={getUserPhotoUrl(context.pageContext.web.absoluteUrl, user.mail)}
                                    imageAlt={user.displayName}
                                    text={user.displayName}
                                    secondaryText={user.jobTitle}
                                    size={PersonaSize.size32}
                                    className={styles.personRowAvatar}
                                />
                            </div>
                            <span className={styles.resultMeta}>{user.department} · {user.mail}</span>
                        </button>
                    ))}
                </div>
            )}

            <div className={styles.sectionTitle}>
                <span className={styles.sectionTitleLabel}><Icon iconName="Balloons" className={styles.sectionTitleIcon} /> Yaklaşan Doğum Günleri</span>
            </div>
            {birthdaysState === 'loading' && <Spinner size={SpinnerSize.small} label="Yükleniyor..." />}
            {birthdaysState === 'error' && (
                <MessageBar messageBarType={MessageBarType.error}>{DATA_UNAVAILABLE_MESSAGE}</MessageBar>
            )}
            {birthdaysState === 'loaded' && birthdays.length === 0 && (
                <Text className={styles.emptyHint}>Bu ay içinde yaklaşan doğum günü bulunmuyor.</Text>
            )}
            {birthdaysState === 'loaded' && pagedBirthdays.map((b) => (
                <div key={b.id} className={styles.personRow}>
                    <Persona text={b.name} secondaryText={b.department} size={PersonaSize.size24} className={styles.personRowAvatar} />
                    {b.isToday ? (
                        <span className={styles.todayBadge}>🎉 İyi ki doğdun!</span>
                    ) : (
                        <span className={styles.badge}>{b.dateLabel}</span>
                    )}
                </div>
            ))}
            {birthdaysState === 'loaded' && birthdays.length > BIRTHDAYS_PAGE_SIZE && (
                <div className={styles.pagination}>
                    <IconButton
                        iconProps={{ iconName: 'ChevronLeft' }}
                        ariaLabel="Önceki sayfa"
                        disabled={birthdayPage === 0}
                        onClick={() => setBirthdayPage((p) => Math.max(0, p - 1))}
                        className={styles.paginationPrevButton}
                    />
                    <span className={styles.pageLabel}>Sayfa {birthdayPage + 1} / {birthdayPageCount}</span>
                    <IconButton
                        iconProps={{ iconName: 'ChevronRight' }}
                        ariaLabel="Sonraki sayfa"
                        disabled={birthdayPage >= birthdayPageCount - 1}
                        onClick={() => setBirthdayPage((p) => Math.min(birthdayPageCount - 1, p + 1))}
                    />
                </div>
            )}

            <DetailModal isOpen={!!selectedUser} title="Kişi Kartı" onDismiss={() => setSelectedUser(undefined)}>
                {selectedUser && (
                    <>
                        <div className={styles.cardHeader}>
                            <Persona
                                imageUrl={getUserPhotoUrl(context.pageContext.web.absoluteUrl, selectedUser.mail)}
                                imageAlt={selectedUser.displayName}
                                text={selectedUser.displayName}
                                size={PersonaSize.size72}
                                hidePersonaDetails
                                className={styles.cardHeaderAvatar}
                            />
                            <div className={styles.cardTextGroup}>
                                <div className={styles.cardName}>{selectedUser.displayName}</div>
                                {selectedUser.jobTitle && <div className={styles.cardJobTitle}>{selectedUser.jobTitle}</div>}
                            </div>
                        </div>
                        {selectedUser.department && (
                            <div className={styles.cardField}>
                                <Icon iconName="Org" className={styles.cardFieldIcon} />
                                {selectedUser.department}
                            </div>
                        )}
                        {selectedUser.mail && (
                            <div className={styles.cardField}>
                                <Icon iconName="Mail" className={styles.cardFieldIcon} />
                                {selectedUser.mail}
                            </div>
                        )}
                        {selectedUser.phone && (
                            <div className={styles.cardField}>
                                <Icon iconName="Phone" className={styles.cardFieldIcon} />
                                {selectedUser.phone}
                            </div>
                        )}
                        <div className={styles.cardActions}>
                            {selectedUser.mail && (
                                <DefaultButton
                                    iconProps={{ iconName: 'Mail' }}
                                    text="E-posta Gönder"
                                    href={`mailto:${selectedUser.mail}`}
                                    className={styles.cardActionButton}
                                />
                            )}
                            {selectedUser.mail && (
                                <DefaultButton
                                    iconProps={{ iconName: 'Chat' }}
                                    text="Teams'de Sohbet"
                                    href={`https://teams.microsoft.com/l/chat/0/0?users=${encodeURIComponent(selectedUser.mail)}`}
                                    target="_blank"
                                    rel="noopener noreferrer"
                                    className={styles.cardActionButton}
                                />
                            )}
                            {selectedUser.phone && (
                                <DefaultButton
                                    iconProps={{ iconName: 'Phone' }}
                                    text="Ara"
                                    href={`tel:${selectedUser.phone}`}
                                    className={styles.cardActionButton}
                                />
                            )}
                        </div>
                    </>
                )}
            </DetailModal>
        </WidgetCard>
    );
};

export default CompanyDirectory;
