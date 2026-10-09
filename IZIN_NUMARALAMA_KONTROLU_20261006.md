# İzin Evrakı Numaralama Ayarı

Tarih: 06.10.2026

## Yeni Talepler

- Oluşturulurken resmi sayı veya resmi kayıt tarihi ayrılmaz.
- Sabit takip numarası `TALEP-oluşturma_yılı-kayıt_kimliği` biçimindedir.
- Taslak kaydetme, red, iade, yeniden sunma ve ıslak imza çıktısı alma resmi
  sayacı tüketmez. Takip numarası değişmez.
- Resmi sayı ve kayıt tarihi, başarılı son imza işleminin parçası olarak
  özlük aktarımından önce hazırlanır. Sayaç yılı son imzanın yılıdır.
- Sayı verme, son onay, işlem geçmişi ve özlük aktarımı aynı atomik işlemle
  tamamlanır. Bakiye veya aktarım hatasında sayı da geri alınır.
- Son imzanın tekrar gönderilmesi yeni sayı veya tekrar kullanım oluşturmaz.
- İlk parafta sabitlenen personel ve imzacı JSON'u numaralama için değişmez.

## Görünüm ve Arama

Formda Takip No ve Resmi Evrak No ayrı alanlardır. Resmi sayı boşken
"Son imzada verilecek" durumu gösterilir. Görevlerde ve listelerde resmi
sayı yoksa takip numarası gösterilir. PDF/Excel, imza öncesinde Takip No,
imza sonrasında Evrak Tarih ve Sayısı başlığını kullanır.

Resmi sayı verildikten sonra eski takip numarasıyla da arama yapılabilir.
Sayısı ve resmi tarihi boş yeni evraklar tarih filtresinden kaybolmaz.
Takip/sayı sıralaması sayfalama öncesinde veritabanında uygulanır.

## Mevcut Kayıtlar

Önceden ayrılmış sayılar ve mevcut kayıt tarihleri korunur; bekleyen eski
evraklar da yeniden numaralandırılmaz. Sayaç sıfırlanmaz, eski ayrılmış
sayılar yeniden kullanılmaz. Mevcut tamamlanmış evraklara veri güncellemesi
yapılmadı. Bu ayar yeni model alanı veya migrasyon gerektirmiyor.

## Doğrulama

- İzin ortamında toplam 139 test geçti: 12,481 saniye.
- 26 odaklı numaralama testi PDF metni okuyucu bulunan ortamda da geçti.
- Chrome'da yeni taslak > ilk paraf > kontrol > amir reddi > yeniden sunma
  > bütün paraflar > son imza akışı doğrulandı.
- İmza öncesi/sonrası PDF metni ve Excel başlığı gerçek HTTP indirmelerinde
  kontrol edildi. Önceki takip numarasıyla arama çalıştı.
- Masaüstü taslak, son imza öncesi, mobil bekleyen ve onaylı evrak olmak
  üzere 4 ekran görüntüsü kontrol edildi; yatay taşma bulunmadı.
- Tarayıcı konsol veya statik dosya hatası bulunmadı.
- Django sistem kontrolü temiz; yeni migrasyon yok; diff boşluk kontrolü temiz.
- Gerçek personel/izin kayıtlarıyla onay denemesi yapılmadı. Testler ayrı
  veritabanlarında çalıştı. Eşzamanlı üretim yük testi yapılmadı; tekrar
  gönderim, hata geri alma ve SQLite kilit hatası testlerle doğrulandı.

Kanıtlar kardeş maaş projesinin `cikti` klasöründedir:
`izin_numara_tum_test_20261006.log`, `izin_numara_pdf_test_20261006.log`,
`izin_numara_tarayici_sonuc_20261006.json`, `izin_numara_*.png`.

Mevcut sunuculara dokunmadan güncel yerel önizleme başlatıldı:
http://127.0.0.1:8008/izin/
