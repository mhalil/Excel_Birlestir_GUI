# Excel Birleştirme Aracı (Excel Merger GUI) v2.0

Excel Birleştirme Aracı, belirlediğiniz bir klasördeki birden fazla Excel dosyasını (`.xls`, `.xlsx`, `.xlsm`) güçlü filtreleme ve veri formatlama kuralları uygulayarak tek bir Excel tablosunda birleştirmenizi sağlayan kullanıcı dostu bir masaüstü uygulamasıdır. 

PySide6 ve Pandas kütüphaneleri kullanılarak geliştirilmiş modern bir arayüze sahiptir.

## Özellikler

Bu uygulama, standart birleştirme araçlarının ötesine geçerek detaylı parametreler sunar:

- **Hızlı ve Kolay Klasör Seçimi**: Seçilen klasördeki tüm uyumlu Excel dosyalarını otomatik tespit eder.
- **Özel Sayfa (Sheet) Belirleme**: Dosyalardaki varsayılan ilk sayfa yerine, adına (örn: `Sayfa1`) veya indeksine (örn: `0`) göre özel bir sayfadan verileri çekebilirsiniz.
- **Esnek Başlık ve Veri Satırı Ayarı**: Başlık satırının ve asıl verilerin kaçıncı satırdan başladığını belirleyebilirsiniz (Özellikle üst kısmında açıklama veya logolar bulunan Excel şablonları için idealdir).
- **Özel Sütun Seçimi**: Tüm tabloyu almak yerine sadece belirlediğiniz sütun aralıklarını (Örn: `B:G`) veya belirli sütunları (Örn: `A, C, D`) kopyalayabilirsiniz.
- **Veri Kaybını Önleme (Metin Olarak Okuma)**: TCKN, IBAN veya telefon numarası gibi uzun sayısal verilerin bilimsel formata (`1,5E+16` vb.) dönüşmesini veya başındaki `0`'ların silinmesini önlemek için belirli sütunları doğrudan "Metin (String)" olarak okutabilirsiniz.
- **Ataşman Cetveli (Örüntü) Desteği**: Standart olmayan veya tekrar eden ara toplam/imza bloklarına sahip dosyalarda, döngüsel olarak belirli sayıda satırı tutup belirli sayıda satırı silme (Örn: "7 satır veri al, sonraki 5 satırı atla") işlemi yapabilirsiniz.
- **Kaynak Dosya Takibi**: Birleştirilen ana tabloda otomatik olarak bir `Dosya ADI` sütunu oluşturularak, hangi satırın hangi dosyadan geldiği kolayca takip edilebilir.
- **Arka Plan İşlemleri (Threading)**: Büyük dosyalarda işlem yapılırken program arayüzü donmaz, güncel durum "İşlem Kayıtları" (Log) ekranından anlık olarak izlenebilir.

## Gereksinimler

Projenin çalışması için bilgisayarınızda Python yüklü olmalı ve aşağıdaki kütüphanelerin kurulu olması gerekmektedir:

- `pandas`
- `PySide6`
- `openpyxl` (`.xlsx` dosyalarını okuma/yazma için)
- `xlwt` / `xlrd` (Eski nesil `.xls` dosyaları için)

Gerekli kütüphaneleri kurmak için komut satırında şu komutu çalıştırabilirsiniz:

```bash
pip install pandas PySide6 openpyxl xlwt xlrd
```

## Kullanım Rehberi

1. **Çalıştırma**: Uygulamayı `python birlestir_detayli.py` komutuyla başlatın.
2. **Klasör Seçimi**: "Klasör Seç..." butonuna tıklayarak birleştirilecek dosyaların bulunduğu dizini belirleyin.
3. **Parametre Ayarları (İsteğe Bağlı)**:
    - **Özel Sayfa Adı Belirt**: Tüm sayfalardan değil, sadece belirteceğiniz adlı sayfadan (veya indeksten) veri çeker.
    - **Özel Başlık/Veri Satırı**: Başlığın 1. satırda olmadığı durumlarda ilgili satır numaralarını girin.
    - **Özel Sütunları Kopyala**: Sadece belli sütunları birleştirmek için sütun harflerini (Örn: `A:F`) yazın.
    - **Metin Olarak Okunacak Sütunlar**: TCKN veya IBAN olan sütunların harf/indeks değerini girin.
    - **Tekrarlayan Alt Satırları Sil**: Her sayfada tekrar eden alt bilgi/imza blokları varsa, verinin ve atlanacak kısmın satır sayılarını belirterek örüntü oluşturun.
4. **Birleştir ve Kaydet**: İşlem parametrelerinizi belirledikten sonra "Dosyaları Birleştir ve Kaydet" butonuna tıklayın. Karşınıza çıkacak pencerede oluşturulacak birleştirilmiş dosyanın adını ve nereye kaydedileceğini seçin.
5. İşlem bitiminde oluşan Excel dosyasını seçtiğiniz konumda bulabilirsiniz.

## Geliştirici ve Lisans

Bu yazılım **Vibe Coder: Mustafa Halil GÖRENTAŞ** tarafından **GPL Lisansı** altında (2026) dağıtılmaktadır. Vibe Coding metodolojisi ve Google Antigravity platformu kullanılarak geliştirilmiştir.

- **Kaynak Kod**: [github.com/mhalil/Excel_Birlestir_GUI](https://github.com/mhalil/Excel_Birlestir_GUI)
