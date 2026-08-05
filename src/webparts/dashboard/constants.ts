/** Tüm widget'larda veri çekme hatalarında gösterilen ortak, kullanıcı dostu mesaj. */
export const DATA_UNAVAILABLE_MESSAGE = 'Veri şu an ulaşılamıyor. Lütfen daha sonra tekrar deneyin.';

/** Gönderim (POST) hatalarında gösterilen ortak mesaj. */
export const SUBMIT_UNAVAILABLE_MESSAGE = 'Şu an gönderilemiyor. Lütfen daha sonra tekrar deneyin.';

/**
 * Duyuru/Etkinlik ekleme yetkisi olan SharePoint grubunun adı. Site İzinleri >
 * Gelişmiş İzin Ayarları'ndan oluşturulmuş gerçek gruptur; Duyurular ve
 * Etkinlikler listelerinde bu gruba Katılım (Contribute) izni verilmiştir.
 */
export const HR_ADMIN_GROUP_NAME = 'İK ve İdari İşler Personeli';

/**
 * İSG Takvimi yükleme yetkisi olan SharePoint grubunun adı. ISGTakvimi
 * klasöründe bu gruba Katılım (Contribute) izni verilmiştir.
 */
export const ISG_GROUP_NAME = 'İSG Personeli';

/**
 * Katılış/Ayrılış takibinde güncelleme yapabilecek ek gruplar — BT ve
 * Muhasebe. Kullanıcı isteğine göre: "+" (yeni kayıt açma) SADECE
 * HR_ADMIN_GROUP_NAME'e özel kalıyor (bkz. OnboardingTrackerWidget.tsx),
 * ama satırdaki kalem/düzenle ikonu HR_ADMIN_GROUP_NAME + IT_GROUP_NAME +
 * ACCOUNTING_GROUP_NAME üçüne birden açık.
 */
export const IT_GROUP_NAME = 'BT Personeli';

/** Katılış/Ayrılış kayıtlarını düzenleyebilecek Muhasebe grubunun adı. */
export const ACCOUNTING_GROUP_NAME = 'Muhasebe Personeli';
