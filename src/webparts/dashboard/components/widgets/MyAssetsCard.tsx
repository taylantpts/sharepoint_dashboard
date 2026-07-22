import * as React from 'react';
import { DefaultButton, Icon, useTheme, mergeStyleSets } from '@fluentui/react';
import WidgetCard from '../WidgetCard';

const ASSET_SYSTEM_URL = 'https://asm.komagene.com.tr/';

const MyAssetsCard: React.FunctionComponent = () => {
    const theme = useTheme();

    const styles = mergeStyleSets({
        // NOT: "gap" flex özelliği burada BİLİNÇLİ OLARAK KULLANILMIYOR — bu
        // sayfanın render edildiği (kurumsal/eski) tarayıcı ortamında flex "gap"
        // desteklenmiyor. Satır `flexWrap: 'wrap'` olduğu için her çocuğa hem
        // marginRight (aynı satırdaki boşluk) hem marginBottom (sarılan satırlar
        // arası boşluk) veriliyor.
        card: {
            display: 'flex',
            alignItems: 'center',
            padding: '18px 20px',
            borderRadius: 10,
            border: `2px solid ${theme.palette.themeLight}`,
            background: theme.palette.themeLighterAlt,
            flexWrap: 'wrap'
        },
        iconWrap: {
            width: 48,
            height: 48,
            borderRadius: 10,
            background: theme.palette.themePrimary,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            flexShrink: 0,
            marginRight: 16,
            marginBottom: 16
        },
        icon: {
            fontSize: 22,
            color: '#ffffff'
        },
        text: {
            flexGrow: 1,
            minWidth: 180,
            marginRight: 16,
            marginBottom: 16
        },
        title: {
            fontSize: 14,
            fontWeight: 700,
            color: theme.semanticColors.bodyText
        },
        desc: {
            fontSize: 12,
            color: theme.semanticColors.bodySubtext,
            marginTop: 2
        },
        ssoNote: {
            fontSize: 11,
            color: theme.semanticColors.bodySubtext,
            marginTop: 10
        }
    });

    return (
        <WidgetCard title="Zimmetlerim" subtitle="Üzerinize kayıtlı cihaz ve envanter" iconName="Devices3" accentColor="#6fa87a">
            <div className={styles.card}>
                <div className={styles.iconWrap}>
                    <Icon iconName="Devices3" className={styles.icon} />
                </div>
                <div className={styles.text}>
                    <div className={styles.title}>Cihaz ve zimmet kayıtlarınız</div>
                    <div className={styles.desc}>Zimmet kayıtları ayrı bir varlık yönetim sisteminde tutulur.</div>
                </div>
                <DefaultButton
                    iconProps={{ iconName: 'OpenInNewWindow' }}
                    text="Zimmetlerimi Görüntüle"
                    href={ASSET_SYSTEM_URL}
                    target="_blank"
                    rel="noopener noreferrer"
                    styles={{ root: { fontWeight: 600 } }}
                />
            </div>
            <div className={styles.ssoNote}>Microsoft hesabınızla (SSO) otomatik giriş yapılır — ayrı bir kullanıcı adı/şifre gerekmez.</div>
        </WidgetCard>
    );
};

export default MyAssetsCard;
