import { createTheme, ITheme } from '@fluentui/react';

/**
 * Standart aydınlık/mavi kurumsal tema (SharePoint'in varsayılan mavi
 * #0078d4 ana rengi). Daha önce burada "Vardiya Merkezi" adlı, tamamen ters
 * çevrilmiş (isInverted: true) koyu kömür/kahve bir tema deneniyordu — ancak
 * bu tema `useTheme()` çağıran TÜM bileşenlere (kart zeminleri dahil)
 * yayıldığı için tüm portal kahverengi görünüyordu. Kullanıcı bunu
 * istemediğini belirtti, o yüzden geri alındı. Neutral/theme* durakları
 * BİLİNÇLİ OLARAK belirtilmiyor — Fluent'in kendi renk üretim algoritması
 * themePrimary'den tutarlı, iyi test edilmiş bir aydınlık ramp üretir.
 */
export const yorpasTheme: ITheme = createTheme({
  palette: {
    themePrimary: '#0078d4'
  }
});
