# İzin Sistemi: Yetki ve Evrak Akışı Kontrolü

Tarih: 05.10.2026

## Kapsam

İzin portalının oluşturma, kaydetme, onaya sunma, ilk paraf, kontrol parafı,
sıralı amir parafları, son imza, red, sahibine iade, geri çekme, yeniden gönderim,
bildirim, ek dosya, özlük aktarımı ve yetki yönetimi yolları incelendi.
Mevcut kullanıcı değişiklikleri korundu. Bu çalışma yeni veritabanı alanı veya
migrasyon gerektirmiyor. Otomatik testler gerçek personel ve izin kayıtlarını
değiştirmeyen ayrı test veritabanlarında çalıştırıldı.

## Akış Kuralları

| Aşama | İşlem Yapabilen | Sonuç |
| --- | --- | --- |
| Taslak / red / sahibine iade | Evrakın sahibi | Önceden girilmiş tarih ve süreler korunur; düzenlenebilir ve yeniden sunulabilir. |
| İlk paraf | Evrakın sahibi | Aktif personel, kontrolcü, gerekli bütün imzacılar ve izin bakiyesi doğrulanır. Belge bilgileri sabitlenir. |
| Kontrol parafı | İlgili personel tipinin atanmış kontrolcüsü | Sıradaki amire aktarılır; henüz özlükte izin kullanımı oluşturulmaz. |
| Amir parafları | O adımın yetkili ve aktif imzacısı | Yalnızca bir sonraki adım açılır. Yönetici olmak başka kişinin yerine imzalama hakkı vermez. |
| Son imza | Son sıradaki yetkili imzacı | Bakiye tekrar doğrulanır; onay, işlem geçmişi ve özlük aktarımı birlikte kaydedilir. |
| Red | Aktif adımın yetkilisi | Red nedeni zorunludur. Evrak sahibine döner; kontrolcüye yalnızca kapatılabilir bilgi bildirimi gider. |
| Sahibine iade | İlgili genel veya tip bazlı iade yetkisi olan kişi | Evrak sahibinde düzenlemeye açılır. |
| Geri çekme | İlgili geri çekme yetkisi olan kişi | Yalnızca geçerli önceki adıma dönülebilir. İleri adıma atlama engellenir. Sahibine dönüş düzenlenebilir iadedir. |
| Tamamlanmış evrak | Görüntüleme kapsamına giren kişiler | Sonraki unvan, birim, izin türü ve imzacı değişiklikleri belgeyi yeniden yazmaz. Normal düzenleme, iade ve geri çekme kapalıdır. |

İlk paraftan sonra kontrolcü veya geçerli vekalet yetkisi değişir ve kayıtlı
imzacıyla çelişirse belge JSON'u değiştirilmez. Bekleyen evrak açıklama ve
denetim kaydıyla sahibine iade edilir. Yeniden sunma ve ilk paraf, güncel
bilgilerle yeni bir tur oluşturur. Tamamlanmış evraklar bu işlemden etkilenmez.

Önceki gönderim turunun imzaları, yeni turun henüz paraflanmamış evrakına
erişim sağlamaz. Belgedeki paraf göstergeleri de geçerli tur ve adıma göre
hesaplanır; geri çekilmiş sonraki imzalar geçerliymiş gibi gösterilmez.

## Yetki Yönetimi

- Yetki ekranı ve detay uç noktası artık açıkça tam yönetim yetkisi ister.
  Anonim veya yalnızca `is_staff` işaretli hesaplar kendilerine yetki veremez.
- Aktif süper kullanıcı sistem yöneticisi olarak genel erişimini korur;
  kontrolcü ataması ve başka kişinin yerine imza verme bundan bağımsızdır.
- Memur, işçi ve şirket için görüntüleme/arama, kontrolcü, geri çekme, iade,
  silme ve bildirim hakları ayrı yönetilir.
- Genel personel aktarımı, birim/imzacı ve vekalet yönetimi hakları ayrıdır.
  Dar kapsamlı yönetici yalnızca yetkili olduğu yönetim sayfasına erişir.
- Tek kullanıcı ve toplu kullanıcı için ekleme, seçili hakları kaldırma,
  değiştirme ve bütün panel haklarını kaldırma işlemleri vardır.
- Bir işlemde en fazla 100 aktif kullanıcı seçilebilir. Kullanıcı listesi
  arama ve sayfalama içerir. Panel haklarını kaldırmak süper kullanıcı
  işaretini veya Django'nun bağımsız admin yetkilerini kaldırmaz.
- Personel tipi başına tek kontrolcü korunur. Kontrolcü devri açık onay
  ister; aynı kontrolcü hakkı topluca birden fazla kişiye verilemez.
- Yetki formu sürüm kontrolü kullanır. Başka bir işlemle değişen tek kayıt
  bile varsa toplu işlem bütünüyle reddedilir; kısmi uygulama yapılmaz.
- Değişen haklar, işlemi yapan ve hedef kişiyle denetim geçmişine yazılır.
- Katalog, `YonetimYetkisi` modelindeki Boolean yetki alanlarından üretilir.
  Yeni alan panelde görünür; yeni bir sayfa/işlem eklenirken sunucu tarafındaki
  erişim kontrolü de uygun yetkiyle bağlanmalıdır. Yeni URL tek başına güvenli
  yetkilendirme üretmez.

Mevcut veritabanında 4 aktif `staff` hesabının açık portal rolü bulunmuyor.
Bu hesaplara kendiliğinden hak verilmedi. Yönetim işlerini yapmaları
gerekiyorsa sistem yöneticisi panelden uygun hakları atamalıdır.

## Veri ve İşlem Güvenceleri

- Özlük aktarımı yalnızca başarılı son imzada ve bir kez yapılır.
- Yıllık izin veya tanımlı mesai bakiyesi son imzada yetersizse onay ve
  aktarım birlikte geri alınır; bakiye dışı bir kullanım kaydı oluşturulmaz.
- Başlangıçta karşılığı olmayan mesai izni elle süreyle kullanılabilir ve
  özlüğe kendi türüyle kaydedilir; yıllık izinden veya sonradan tanımlanan
  mesai hakkından düşülmez.
- Eski tarayıcı formu sürümü paraf/düzenleme işleminde reddedilir.
- SQLite kilitlenmesi halinde atomik işlem geri alındıktan sonra tekrar
  deneme bildirimi gösterilir. JSON istemcisi 409 alır. Diğer veritabanı
  hataları yoğunluk mesajının arkasına gizlenmez.
- Eksik veya pasif imzacı, bozuk imza sırası ve geçersiz kayıt kimliği
  evrakı otomatik tamamlayamaz.
- Ek dosyalar yalnızca sahibinin düzenlenebilir evrakında değiştirilebilir.
  İndirme evrak erişim kapsamını doğrular; bulunmayan dosya 404 verir.
- Django admin ekranından izin durumu elle değiştirilerek imza akışı
  atlanamaz. Portalın mevcut özel tamamlanmış evrak silme yolu, aktif
  `erenaygun` süper kullanıcısı ve parola doğrulamasıyla sınırlıdır.

## Doğrulama Sonuçları

| Kontrol | Sonuç |
| --- | --- |
| `manage.py test portal --noinput` | 113 test geçti; 6,486 saniye. |
| `manage.py check` | Sorun bulunmadı. |
| `manage.py makemigrations --check --dry-run` | Yeni model değişikliği yok. |
| `git diff --check` | Boşluk hatası yok; mevcut LF/CRLF uyarıları dışında temiz. |
| Chrome: toplu yetki verme | Gerçek form gönderimi ve veritabanı sonucu doğrulandı. |
| Chrome: sahip > kontrolcü > amir > red > yeniden gönderim > son imza | Akış ve son imzada tek aktarım doğrulandı. |
| PDF ve Excel | Gerçek indirme isteği 200; içerik üretildi. |
| Ekran görüntüleri | Yetki masaüstü/tablet/mobil, ilk paraf, tamamlanan evrak ve mobil evrak olmak üzere 6 görünüm kontrol edildi. |
| Mobil taşma / işlem çubuğu | Sağa taşma giderildi; işlem çubuğu artık belge üzerine sabitlenmiyor. |
| Tarayıcı hataları | Konsol hatası ve başarısız statik dosya isteği yok. |

500 onaylı test evrakıyla, yerel Django test istemcisindeki üç ölçümün
medyanı (tarayıcı çizimi ve ağ süresi dahil değildir):

| Sayfa | Süre | SQL Sorgusu |
| --- | ---: | ---: |
| Kişisel sayfa | 14,3 ms | 20 |
| İzin arama / liste | 32,3 ms | 28 |
| Kullanıcı yetkileri | 9,4 ms | 14 |
| Birim / imzacı yönetimi | 7,8 ms | 14 |

Varsayılan liste sıralaması veritabanında sayfalanır; tüm evraklar önce
Python'a yüklenmez. Serbest metin ve özel sütun sıralama yollarının mevcut
davranışı korunmuştur; yukarıdaki ölçüm bu yolların yük testi değildir.

## Sınırlar

Bu kontroller tüm olası üretim koşullarında hata olmayacağı garantisi değildir.
SQLite yoğunluğu hata enjeksiyonuyla doğrulandı; eşzamanlı çok kullanıcılı
üretim yük testi yapılmadı. PostgreSQL gibi farklı bir veritabanı denenmedi.

Eski istemcilerle uyum için evrak sürüm alanı eksik istekler bütünüyle
yasaklanmadı; portalın güncel formları bu alanı gönderir. Tek aktarım ve
işlem geri alma kontrolleri ayrıca uygulanır.

Eski sabit belgelerde asıl/vekil kimliği bulunmayabilir. Bu çalışmada
mevcut tek bekleyen evrakta vekil rol veya ilişkili vekalet kaydı olmadığı
okuma sorgusuyla kontrol edildi. Eski, vekil bilgisi eksik belgeler topluca
yeniden yazılmadı. Yeni paraf turları asıl imzacı kimliğini ayrıca saklar.

## Önizleme ve Kanıtlar

Mevcut 8002 sunucusuna dokunmadan yerel önizleme başlatıldı:

- Giriş: http://127.0.0.1:8007/giris/
- Yetkiler: http://127.0.0.1:8007/yonetim/kullanici-yetkileri/

Test logları ve ekran görüntüleri kardeş maaş projesinin `cikti` klasöründe:

- `izin_yetki_final_test_20261005.log`
- `izin_tarayici_sonuc_20261005.json`
- `izin_tarayici_log_20261005.log`
- `izin_yetki_desktop.png`, `izin_yetki_tablet.png`, `izin_yetki_mobile.png`
- `izin_ilk_paraf.png`, `izin_tamamlanan_evrak.png`, `izin_evrak_mobile.png`
