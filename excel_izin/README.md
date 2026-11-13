# Excel memur izin programı

Güncel VBA kaynakları bu klasördedir. İzin/rapor geçmişi sayfalıdır; yıl ve türle süzülür. Yeni yıl işlemi önce kapanış yedeği alır, önceki yılın kalanlarını taşır ve yeni hak alanını boş açar.

Kaynaklar UTF-8 biçimindedir. VBE içe aktarmasında Türkçe karakterleri korumak için `python vbe_import_hazirla.py` çalıştırın; Windows-1254 çıktıları `import_icin` klasörüne oluşur. `Module1.bas` standart modüldür; belge sınıflarının kodu Excel'de aynı belge modüllerine yerleştirilir.

Çalışır dosyanın düzeni ve kişisel verileri şifreli tam yedeğin `excel/Memur_Izin_Takibi.xlsm` dosyasındadır. Tam yedek maaş deposundaki `guvenli_yedekler` klasöründe saklanır. Kurtarma anahtarı ayrı tutulur; Git'e yüklenmez. Geri yükleme açıklaması maaş deposundaki `docs/GIT_YEDEKLEME.md` dosyasındadır.

Çalışma sayfası koruma şifresi: `2022-tiski`. Bu yalnız sayfa düzenleme korumasıdır; tam yedeğin şifreleme anahtarı değildir.

13.11.2026: İzin kaydedilince form ve personel araması temizlenir. Seçili personelin toplam izni ve girilen yıllık izin sonrası kalan izin önizlemesi gösterilir.
