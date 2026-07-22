import * as React from 'react';
import { Modal, IconButton, useTheme, mergeStyleSets } from '@fluentui/react';

export interface IDetailModalProps {
    isOpen: boolean;
    title?: string;
    onDismiss: () => void;
    children?: React.ReactNode;
}

/**
 * Duyuru/Etkinlik gibi liste öğelerinin detayını göstermek için kullanılan
 * ortak, ekranın TAM ORTASINDA açılan, arkayı karartan (overlay) Modal
 * kabuğu — Fluent UI'ın Modal bileşeni bunu varsayılan olarak (isBlocking)
 * sağlıyor. Geniş bir kutu, büyük/belirgin başlık, rahat okunur (line-height:
 * 1.6) gövde metni ve sağ üstte belirgin bir kapatma (X) butonu içerir.
 */
const DetailModal: React.FunctionComponent<IDetailModalProps> = (props) => {
    const { isOpen, title, onDismiss, children } = props;
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
            lineHeight: 1.3,
            marginRight: 16
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
        body: {
            padding: '0 28px 28px',
            lineHeight: 1.6
        }
    });

    return (
        <Modal isOpen={isOpen} onDismiss={onDismiss} isBlocking={false} containerClassName={styles.container}>
            <div className={styles.header}>
                <div className={styles.title}>{title}</div>
                <IconButton
                    iconProps={{ iconName: 'Cancel' }}
                    ariaLabel="Kapat"
                    onClick={onDismiss}
                    className={styles.closeButton}
                />
            </div>
            <div className={styles.body}>{children}</div>
        </Modal>
    );
};

export default DetailModal;
