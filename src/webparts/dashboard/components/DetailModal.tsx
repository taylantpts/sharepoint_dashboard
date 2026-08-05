import * as React from 'react';
import { Modal, IconButton, useTheme, mergeStyleSets } from '@fluentui/react';

export interface IDetailModalProps {
    isOpen: boolean;
    title?: string;
    onDismiss: () => void;
    children?: React.ReactNode;
    /** Verilirse kapatma butonunun solunda kırmızı bir çöp kutusu butonu render edilir — yetkili kullanıcılar için silme aksiyonu. */
    onDeleteClick?: () => void;
    deleteAriaLabel?: string;
}

/**
 * Duyuru/Etkinlik gibi liste öğelerinin detayını göstermek için kullanılan
 * ortak, ekranın TAM ORTASINDA açılan, arkayı karartan (overlay) Modal
 * kabuğu — Fluent UI'ın Modal bileşeni bunu varsayılan olarak (isBlocking)
 * sağlıyor. Geniş bir kutu, büyük/belirgin başlık, rahat okunur (line-height:
 * 1.6) gövde metni ve sağ üstte belirgin bir kapatma (X) butonu içerir.
 */
const DetailModal: React.FunctionComponent<IDetailModalProps> = (props) => {
    const { isOpen, title, onDismiss, children, onDeleteClick, deleteAriaLabel } = props;
    const theme = useTheme();

    const styles = mergeStyleSets({
        container: {
            width: 640,
            maxWidth: '92vw',
            borderRadius: 12,
            overflow: 'hidden',
            boxShadow: '0 24px 60px rgba(0,0,0,0.22)'
        },
        // NOT: "gap" flex özelliği burada BİLİNÇLİ OLARAK KULLANILMIYOR — bu
        // sayfanın render edildiği (kurumsal/eski) tarayıcı ortamında flex "gap"
        // desteklenmiyor. header'ın tek gerçek çocuğu title olduğu için (X butonu
        // öncesinde) marginRight doğrudan title'a veriliyor.
        header: {
            display: 'flex',
            alignItems: 'flex-start',
            justifyContent: 'space-between',
            padding: '28px 28px 20px'
        },
        title: {
            fontSize: 22,
            fontWeight: 700,
            color: theme.semanticColors.bodyText,
            // ÖNCEKİ HATA: burada birimsiz (unitless) bir sayı (1.3) kullanılıyordu —
            // bu render ortamında merge-styles bunu "1.3px" olarak (birim EKLEYEREK)
            // hesaplıyor, satırın gerçek yüksekliğini neredeyse sıfıra indirip metnin
            // üst üste binmesine yol açıyor. Yüzde string'i ('130%') açık bir CSS
            // birimi taşıdığı için bu hataya düşmüyor.
            lineHeight: '130%',
            marginRight: 16
        },
        // NOT: "gap" burada da kullanılmıyor — iki buton yan yana geldiğinde
        // aralarındaki boşluk, silme butonuna verilen marginRight ile sağlanıyor.
        headerActions: {
            display: 'flex',
            alignItems: 'center',
            flexShrink: 0
        },
        closeButton: {
            flexShrink: 0,
            width: 36,
            height: 36,
            borderRadius: 8,
            background: theme.palette.neutralLighterAlt,
            selectors: {
                ':hover': {
                    background: theme.palette.neutralLighter
                }
            }
        },
        deleteButton: {
            flexShrink: 0,
            width: 36,
            height: 36,
            borderRadius: 8,
            marginRight: 8,
            color: '#B91C1C',
            background: 'rgba(185,28,28,0.08)',
            selectors: {
                ':hover': {
                    color: '#B91C1C',
                    background: 'rgba(185,28,28,0.16)'
                }
            }
        },
        body: {
            padding: '0 28px 28px',
            // ÖNCEKİ HATA: bkz. yukarıdaki "title" notu — aynı birimsiz-sayı hatası.
            // Bu değer, kendi lineHeight'ını ayrıca belirtmeyen tüm alt bileşenler
            // (ör. RequiredDocuments) tarafından miras alınıyor, bu yüzden düzeltmek
            // site genelinde birçok detay modalını etkiliyor.
            lineHeight: '160%'
        }
    });

    return (
        <Modal isOpen={isOpen} onDismiss={onDismiss} isBlocking={false} containerClassName={styles.container}>
            <div className={styles.header}>
                <div className={styles.title}>{title}</div>
                <div className={styles.headerActions}>
                    {onDeleteClick && (
                        <IconButton
                            iconProps={{ iconName: 'Delete' }}
                            ariaLabel={deleteAriaLabel ?? 'Sil'}
                            title={deleteAriaLabel ?? 'Sil'}
                            onClick={onDeleteClick}
                            className={styles.deleteButton}
                        />
                    )}
                    <IconButton
                        iconProps={{ iconName: 'Cancel' }}
                        ariaLabel="Kapat"
                        onClick={onDismiss}
                        className={styles.closeButton}
                    />
                </div>
            </div>
            <div className={styles.body}>{children}</div>
        </Modal>
    );
};

export default DetailModal;
