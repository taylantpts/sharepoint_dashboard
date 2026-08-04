import * as React from 'react';
import { Icon, useTheme, mergeStyleSets } from '@fluentui/react';
import WidgetCard from '../WidgetCard';

interface IQuickLink {
    id: string;
    label: string;
    iconName: string;
    href: string;
    color: string;
}

// Renkler bilinçli olarak Fluent UI'ın canlı kurumsal paletinden seçildi —
// her kutucuk farklı bir vurgu rengiyle görsel çeşitlilik sağlar.
const QUICK_LINKS: IQuickLink[] = [
    { id: 'ql-1', label: 'Sözlük', iconName: 'Dictionary', href: 'https://yorpas.sharepoint.com/sites/YORPASA/Lists/Komagene_Sozluk/Szlk%20Galerisi.aspx', color: '#0078d4' },
    { id: 'ql-2', label: 'Eğitim', iconName: 'Education', href: 'https://egitim.komagene.com.tr/login/index.php', color: '#107c10' },
    { id: 'ql-3', label: 'Bayi Sipariş', iconName: 'ShoppingCart', href: 'https://genegenekomagene.com/Admin/Login.aspx', color: '#d83b01' },
    { id: 'ql-4', label: 'Komagene Portal', iconName: 'Website', href: 'https://web.genegenekomagene.com/Admin/Login?ReturnUrl=/Admin/Default', color: '#038387' },
    { id: 'ql-5', label: 'Masraff', iconName: 'Money', href: 'https://admin.masraff.co/#/access/login', color: '#5c2d91' },
    { id: 'ql-6', label: 'Entegre Yönetim', iconName: 'Settings', href: 'https://qdms.komagene.com.tr/QDMSNET/BSAT/Logon.aspx', color: '#b4009e' },
    { id: 'ql-7', label: 'Organizasyon Şeması', iconName: 'Org', href: 'https://yorpas.sharepoint.com/SiteAssets/SitePages/Y%C3%B6rpa%C5%9F/Organizasyon-%C5%9Eemas%C4%B1_OCAK2026.pptx?web=1', color: '#004578' }
];

const QuickLinksPanel: React.FunctionComponent = () => {
    const theme = useTheme();

    const styles = mergeStyleSets({
        grid: {
            display: 'grid',
            gridTemplateColumns: 'repeat(auto-fit, minmax(140px, 1fr))',
            gap: 14
        },
        // NOT: "gap" flex özelliği burada BİLİNÇLİ OLARAK KULLANILMIYOR — bu
        // sayfanın render edildiği (kurumsal/eski) tarayıcı ortamında flex "gap"
        // desteklenmiyor. externalIcon position:absolute olduğu için normal akışın
        // (ve gap'in) dışında kalıyor; gerçek akıştaki tek gerekli boşluk
        // iconWrap → label arasında, bu yüzden marginBottom sadece iconWrap'te.
        link: {
            position: 'relative',
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'flex-start',
            padding: '18px 16px',
            borderRadius: 14,
            background: theme.palette.neutralLighterAlt,
            border: `1px solid ${theme.palette.neutralLight}`,
            textDecoration: 'none',
            transition: 'transform 0.2s ease, box-shadow 0.2s ease, border-color 0.2s ease',
            selectors: {
                ':hover': {
                    transform: 'translateY(-4px)',
                    borderColor: 'transparent'
                },
                // Vurgu rengi her karta göre değiştiği için hover gölgesi CSS
                // yerine onMouseEnter/onMouseLeave ile (aşağıda) veriliyor.
                ':hover .ql-externalIcon': {
                    opacity: 1,
                    transform: 'translate(0, 0)'
                }
            }
        },
        iconWrap: {
            width: 42,
            height: 42,
            borderRadius: 11,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            flexShrink: 0,
            marginBottom: 12
        },
        icon: {
            fontSize: 20,
            color: '#ffffff'
        },
        label: {
            fontSize: 13,
            fontWeight: 700,
            color: theme.semanticColors.bodyText,
            lineHeight: '17px'
        },
        // Sağ üst köşede, sadece hover'da beliren "yeni sekmede açılır" ipucu —
        // her kartın harici bir bağlantı olduğunu ince bir şekilde hatırlatır.
        externalIcon: {
            position: 'absolute',
            top: 12,
            right: 12,
            fontSize: 11,
            color: theme.semanticColors.bodySubtext,
            opacity: 0,
            transform: 'translate(-2px, 2px)',
            transition: 'opacity 0.2s ease, transform 0.2s ease'
        }
    });

    return (
        <WidgetCard title="Hızlı Erişim" subtitle="Sık kullanılan uygulama ve portallar" iconName="Rocket">
            <div className={styles.grid}>
                {QUICK_LINKS.map((link) => (
                    <a
                        key={link.id}
                        href={link.href}
                        target="_blank"
                        rel="noopener noreferrer"
                        className={styles.link}
                        style={{ boxShadow: `0 1px 3px rgba(0,0,0,0.06)` }}
                        onMouseEnter={(e) => {
                            e.currentTarget.style.boxShadow = `0 10px 22px rgba(0,0,0,0.30), 0 0 0 1px ${link.color}80`;
                        }}
                        onMouseLeave={(e) => {
                            e.currentTarget.style.boxShadow = '0 1px 3px rgba(0,0,0,0.06)';
                        }}
                    >
                        <Icon iconName="OpenInNewWindow" className={`${styles.externalIcon} ql-externalIcon`} />
                        <div
                            className={styles.iconWrap}
                            style={{
                                background: `linear-gradient(135deg, ${link.color} 0%, ${link.color}cc 100%)`,
                                boxShadow: `0 4px 12px ${link.color}55`
                            }}
                        >
                            <Icon iconName={link.iconName} className={styles.icon} />
                        </div>
                        <span className={styles.label}>{link.label}</span>
                    </a>
                ))}
            </div>
        </WidgetCard>
    );
};

export default QuickLinksPanel;
