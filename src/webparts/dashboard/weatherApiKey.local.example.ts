// ŞABLON DOSYA — bunu "weatherApiKey.local.ts" adıyla AYNI klasöre kopyalayıp
// aşağıdaki boş değeri kendi (ücretsiz) OpenWeatherMap API anahtarınızla
// doldurun. "weatherApiKey.local.ts" .gitignore'dadır — bu yüzden anahtarınız
// asla commit'e/repoya girmez, sadece bu makinede kalır.
//
// Bu değer yalnızca web part'ın Property Pane'inden ("OpenWeatherMap API
// Anahtarı" alanı) hiçbir anahtar girilmemişse devreye giren bir YEDEKTİR —
// bkz. DashboardWebPart.ts. Property Pane'den girilen değer HER ZAMAN bunun
// önüne geçer. Üretim/prod dağıtımı için önerilen yol hâlâ Property Pane'dir;
// bu dosya sadece yerel geliştirme/workbench testinde tekrar tekrar panel
// doldurmak zorunda kalmamak için bir kolaylıktır.
export const WEATHER_API_KEY_LOCAL_FALLBACK = '';
