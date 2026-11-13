# İzin Programları

Bu depo Django izin portalını ve `excel_izin` klasöründe Excel memur izin programının güncel VBA kaynaklarını içerir.

## Web izin portalı

`calistir.bat` yerel ortamı başlatır. Bağımlılıklar `requirements.txt` dosyasındadır. Web izin portalı maaş uygulamasıyla aynı SQLite veritabanını kullanır; yeni bilgisayarda `MAAS_PROJESI_DB` ortam değişkenini maaş projesinin `db.sqlite3` dosyasına ayarlayın. Canlı ortamda `DJANGO_SECRET_KEY` ve ilgili güvenlik ayarları ayrıca tanımlanmalıdır.

## Excel izin programı

Kaynaklar ve VBE için karakter kodlama hazırlama aracı `excel_izin` klasöründedir. İzin/rapor geçmişi yıl/tür filtresiyle sayfalanır; yeni yıl açılışı yedek alır ve önceki yılın kalanlarını taşır.

Personel kayıtları açık biçimde bu depoya konulmaz. Çalışır Excel ve ortak veritabanı, maaş deposundaki `guvenli_yedekler` klasöründe şifreli tam yedek olarak saklanır. Ayrı kurtarma anahtarı Git'e yüklenmez; geri yükleme adımları [maaş deposundaki belgede](https://github.com/Erenaygun61/Maas_Projesi/blob/main/docs/GIT_YEDEKLEME.md) açıklanır.
