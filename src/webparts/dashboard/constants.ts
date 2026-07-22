/** Tüm widget'larda veri çekme hatalarında gösterilen ortak, kullanıcı dostu mesaj. */
export const DATA_UNAVAILABLE_MESSAGE = 'Veri şu an ulaşılamıyor. Lütfen daha sonra tekrar deneyin.';

/** Gönderim (POST) hatalarında gösterilen ortak mesaj. */
export const SUBMIT_UNAVAILABLE_MESSAGE = 'Şu an gönderilemiyor. Lütfen daha sonra tekrar deneyin.';

/**
 * GEÇİCİ MOCK — Duyuru/Etkinlik ekleme butonlarının (+) kimlere görüneceğini
 * belirler. Şu an sabit `true`: yetki kontrolü henüz gerçek bir mekanizmaya
 * (ör. kullanıcının belirli bir SharePoint grubunda olup olmadığının
 * `_api/web/currentuser/groups` ile kontrolü) bağlanmadı — bu iş üretime
 * alınmadan önce yapılmalı, aksi halde bu butonlar TÜM kullanıcılara görünür.
 */
export const IS_ADMIN_MOCK = true;
