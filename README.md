# Active Directory Analiz Paneli

<p align="center">
  <kbd><img src="img1.png" width="48%" alt="Kullanıcı Paneli" /></kbd>
  <kbd><img src="img2.png" width="48%" alt="Cihaz Paneli" /></kbd>
  <br>
  <sub><b>Görsel 1:</b> Kullanıcı Detay Paneli &nbsp; | &nbsp; <b>Görsel 2:</b> Cihaz Envanter Paneli</sub>
</p>

Active Directory ortamındaki kullanıcı ve cihaz verilerini interaktif bir web raporuna dönüştüren geniş kapsamlı bir PowerShell aracıdır.

## Başlıca Özellikler
- **Modern Görselleştirme:** Yan menü tabanlı dashboard yapısı ile kullanıcı, cihaz ve departman verilerinin anlık gösterimi.
- **Dinamik Filtreleme:** Kullanıcı ve cihaz listelerinde sütun bazlı arama ve `>90`, `<30` gibi mantıksal gün aralığı sorgulama desteği.
- **Güvenlik Analizi:** Yönetici yetkisine sahip kullanıcıların (Domain Admins, Schema Admins vb.) ve kilitli hesapların otomatik tespiti.
- **Kullanım Takibi:** Atıl (Son 90 gün içinde oturum açmamış) hesapların ve parolası dolmuş profillerin raporlanması.
- **Envanter Yönetimi:** Bilgisayar nesnelerinin işletim sistemi dağılımı ve son oturum açma tarihlerinin takibi.
- **Karakter Güvenliği:** Türkçe karakter sorunu oluşturmayan, tüm tarayıcılarla uyumlu HTML5 yapısı.

## Kullanım Rehberi
1. `AD-Analiz-Paneli.ps1` dosyasını sağ tıklayarak "PowerShell ile çalıştır" seçeneğini kullanın.
2. İşlem tamamlandığında masaüstünde `AD-Analiz-Raporu_*.html` isimli rapor dosyası oluşturulacaktır.
3. Oluşan bu dosyayı herhangi bir tarayıcı (Chrome, Edge vb.) ile açabilirsiniz.

*Hazırlayan: Safak Can Bav*
