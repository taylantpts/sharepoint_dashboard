import * as React from 'react';
import { SearchBox, Persona, PersonaSize, Spinner, SpinnerSize, MessageBar, MessageBarType, Text, Icon, DefaultButton, useTheme, mergeStyleSets } from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import WidgetCard from '../WidgetCard';
import DetailModal from '../DetailModal';
import { searchDirectory, IDirectoryUser } from '../../services/GraphService';
import { DATA_UNAVAILABLE_MESSAGE } from '../../constants';

export interface ICompanyDirectoryProps {
    context: WebPartContext;
}

type LoadState = 'idle' | 'loading' | 'loaded' | 'error';

const SEARCH_DEBOUNCE_MS = 350;

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
        // personRow, arama sonucu satırındaki Persona'yı sarmalıyor — "gap"
        // yerine avatar sınıfı (personRowAvatar) marginRight taşıyor.
        personRow: {
            display: 'flex',
            alignItems: 'center',
            padding: '3px 0'
        },
        personRowAvatar: {
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
