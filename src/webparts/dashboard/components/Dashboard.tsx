import * as React from 'react';
import { ThemeProvider } from '@fluentui/react';
import { IDashboardProps } from '../Dashboard.types';
import { yorpasTheme } from '../theme';
import styles from './Dashboard.module.scss';
import WelcomeHeader from './widgets/WelcomeHeader';
import WeatherWidget from './widgets/WeatherWidget';
import QuickLinksPanel from './widgets/QuickLinksPanel';
import CompanyDirectory from './widgets/CompanyDirectory';
import AnnouncementsFeed from './widgets/AnnouncementsFeed';
import UpcomingEvents from './widgets/UpcomingEvents';
import ITSupportCenter from './widgets/ITSupportCenter';
import HRPortalCard from './widgets/HRPortalCard';
import MyAssetsCard from './widgets/MyAssetsCard';
import RequiredDocuments from './widgets/RequiredDocuments';
import LocationCard from './widgets/LocationCard';
import FeedbackForm from './widgets/FeedbackForm';
import PersonalAssistant from './widgets/PersonalAssistant';
import HolidaysWidget from './widgets/HolidaysWidget';
import ExchangeRatesWidget from './widgets/ExchangeRatesWidget';
import NewsWidget from './widgets/NewsWidget';
import ISGCalendarWidget from './widgets/ISGCalendarWidget';
import BirthdaysWidget from './widgets/BirthdaysWidget';
import OnboardingTrackerWidget from './widgets/OnboardingTrackerWidget';
import RecentOnboardingWidget from './widgets/RecentOnboardingWidget';

const Dashboard: React.FunctionComponent<IDashboardProps> = (props) => {
    return (
        <ThemeProvider theme={yorpasTheme} className={styles.dashboardRoot}>
            <div className={styles.headerRow}>
                <WelcomeHeader userDisplayName={props.userDisplayName} context={props.context} />
            </div>

            <div className={styles.grid}>
                <div className={styles.areaAnnouncements}>
                    <AnnouncementsFeed context={props.context} />
                </div>
                <div className={styles.areaWeather}>
                    <WeatherWidget apiKey={props.weatherApiKey} />
                </div>

                {/* Hızlı Erişim/Rehber/Etkinlikler/Asistan artık 4 dar sütun yerine
                    kendi 2x2 grid'inde ferah biçimde yerleşiyor (bkz. .quadGrid). */}
                <div className={styles.areaQuad}>
                    <div className={styles.quadGrid}>
                        <div className={styles.quadItem}>
                            <QuickLinksPanel />
                        </div>
                        <div className={styles.quadItem}>
                            <CompanyDirectory context={props.context} />
                        </div>
                        <div className={styles.quadItem}>
                            <BirthdaysWidget context={props.context} />
                        </div>
                        <div className={styles.quadItem}>
                            <UpcomingEvents context={props.context} />
                        </div>
                        <div className={styles.quadItem}>
                            <PersonalAssistant context={props.context} />
                        </div>
                    </div>
                </div>

                <div className={styles.areaHolidays}>
                    <HolidaysWidget />
                </div>
                <div className={styles.areaExchange}>
                    <ExchangeRatesWidget />
                </div>
                {/* Gerekli Dosyalar ve Haberler kasıtlı olarak yer değiştirdi: Gerekli
                    Dosyalar artık BT Destek Merkezi ile, Haberler ise İSG Takvimi ile
                    aynı satırda — grid alanlarının (areaNews/areaDocuments) KENDİSİ
                    değişmedi, sadece hangi widget'ın hangi alanda göründüğü değişti. */}
                <div className={styles.areaNews}>
                    <RequiredDocuments context={props.context} />
                </div>

                <div className={styles.areaSupport}>
                    <ITSupportCenter />
                </div>
                <div className={styles.areaHR}>
                    <HRPortalCard />
                </div>
                <div className={styles.areaAssets}>
                    <MyAssetsCard />
                </div>

                <div className={styles.areaDocuments}>
                    <NewsWidget />
                </div>
                <div className={styles.areaISGCalendar}>
                    <ISGCalendarWidget context={props.context} />
                </div>
                <div className={styles.areaRecentJoiners}>
                    <RecentOnboardingWidget
                        context={props.context}
                        kind="katilis"
                        title="Aramıza Katılanlar"
                        subtitle="Son katılan 5 çalışan"
                        iconName="AddFriend"
                        emptyHint="Henüz görüntülenecek katılış kaydı bulunmuyor."
                    />
                </div>
                <div className={styles.areaRecentLeavers}>
                    <RecentOnboardingWidget
                        context={props.context}
                        kind="ayrilis"
                        title="Aramızdan Ayrılanlar"
                        subtitle="Son ayrılan 5 çalışan"
                        iconName="UserRemove"
                        emptyHint="Henüz görüntülenecek ayrılış kaydı bulunmuyor."
                    />
                </div>
                <div className={styles.areaOnboarding}>
                    <OnboardingTrackerWidget context={props.context} />
                </div>

                <div className={styles.areaLocationOffice}>
                    <LocationCard
                        label="Merkez Ofis"
                        address="Ofisim İstanbul Plazaları, Cevizli Mah. Tugay Yolu Cad. No:20 A Blok Kat:6 Ofis No:29, 34846 Maltepe/İstanbul, Türkiye"
                        coordinates={{ lat: 40.9183137, lon: 29.1669627 }}
                        iconName="CityNext"
                    />
                </div>
                <div className={styles.areaLocationFactory}>
                    <LocationCard
                        label="Fabrika"
                        address="İnönü Mah. Atatürk Blv. No:19, 41400 Gebze/Kocaeli"
                        iconName="Manufacturing"
                    />
                </div>

                <div className={styles.areaFeedback}>
                    <FeedbackForm context={props.context} />
                </div>
            </div>
        </ThemeProvider>
    );
};

export default Dashboard;
