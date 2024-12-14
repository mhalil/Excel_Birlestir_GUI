from ttkbootstrap import Style
from tkinter import ttk, filedialog
import glob
import pandas as pd
from PIL import ImageTk

style = Style()
style = Style(theme='litera')

pencere = style.master
pencere.title("Excel Birleştir")
pencere.geometry("500x500+500+100")

excel_dosyalari = []

### FONKSİYONLAR
def klasor_sec():
	global excel_dosyalari
	klasor_adi = filedialog.askdirectory()
	xls = klasor_adi + "/*.xls*"
	excel_dosyalari = glob.glob(xls)
	### Seçilen dosyalar hakkında bilgi veren bir diyalog açılabilir


def birlestir():
	global excel_dosyalari
	for i in range(len(excel_dosyalari)):
		a = "df_" + str(i)
		print(a + "Veri Çerçevesi:")
		a = pd.read_excel(excel_dosyalari[i])
		print(a)

def yardim():
    image = ImageTk.PhotoImage(file = "parametreler.png")
    etiket_resim.configure(image = image)
    etiket_resim.image = image
    pencere.geometry("1400x700+200+100")



### Klasör seçici
ttk.Label(pencere, text='Excel dosyalarını seç:').grid(row=0, column=0, pady=5, padx=20)
ttk.Button(pencere, text="Birleştirilecek dosyaları seçin...", style='primary.TButton', command=klasor_sec).grid(row=0, column=1, pady=5, padx=25)

### Çerçeve
cerceve_parametreler = ttk.LabelFrame(
    pencere,
    width=400,
    height=450,
    text="Parametreler")
cerceve_parametreler.grid(row=1, column=0, pady=5, padx=25, columnspan=2)


### Parametreler
sayfa_belirt = ttk.Checkbutton(cerceve_parametreler, text='Sayfa Adı Belirt', style='primary.Roundtoggle.Toolbutton').grid(row=0, column=0, pady=5, padx=25)
entry_sayfa_adi = ttk.Entry(cerceve_parametreler).grid(row=0, column=1, pady=5, padx=25)

etiket_baslik_satir_no = ttk.Label(cerceve_parametreler, text='Başlık Satırı Numarası:').grid(row=1, column=0, pady=5, padx=25)
entry_baslik_satiri = ttk.Entry(cerceve_parametreler).grid(row=1, column=1, pady=5, padx=25)

etiket_ilk_veri_satiri = ttk.Label(cerceve_parametreler, text='İlk Veri Satırı Numarası:').grid(row=2, column=0, pady=5, padx=25)
entry_ilk_veri_satiri = ttk.Entry(cerceve_parametreler).grid(row=2, column=1, pady=5, padx=25)

etiket_kopyalanacak_satir = ttk.Label(cerceve_parametreler, text='Kopyalanacak Satır Sayısı:').grid(row=3, column=0, pady=5, padx=25)
entry_kopyalanacak_satir = ttk.Entry(cerceve_parametreler).grid(row=3, column=1, pady=5, padx=25)

etiket_atlanacak_satir = ttk.Label(cerceve_parametreler, text='Atlanacak Satırı Numarası:').grid(row=4, column=0, pady=5, padx=25)
entry_atlanacak_satir = ttk.Entry(cerceve_parametreler).grid(row=4, column=1, pady=5, padx=25)

etiket_kopyalanacak_sutun = ttk.Label(cerceve_parametreler, text='Kopyalanacak Sütunlar:').grid(row=5, column=0, pady=5, padx=25)
entry_kopyalanacak_sutun = ttk.Entry(cerceve_parametreler).grid(row=5, column=1, pady=5, padx=25)

etiket_dongu_sayisi = ttk.Label(cerceve_parametreler, text='Döngü Sayısı:').grid(row=6, column=0, pady=5, padx=25)
entry_dongu_sayisi = ttk.Entry(cerceve_parametreler).grid(row=6, column=1, pady=5, padx=25)

kayit_dosya_adi = ttk.Checkbutton(cerceve_parametreler, text='Kayıt için Dosya Adı Belirt', style='primary.Roundtoggle.Toolbutton').grid(row=7, column=0, pady=5, padx=25)
entry_kayit_dosya_adi = ttk.Entry(cerceve_parametreler).grid(row=7, column=1, pady=5, padx=25)
#####

### Alt Butonlar
ttk.Button(pencere, text="Görsel Yardım", style='info.TButton', command=yardim).grid(row=2, column=0, pady=5, padx=25)
ttk.Button(pencere, text="Dosyaları Birleştir", style='primary.TButton', command=birlestir).grid(row=2, column=1, pady=5, padx=25)

### Durum Çubuğu yerine bilgi metni
bilgi = ttk.Label(pencere, text="Bilgi: Program birleştirme işlemi için hazır...", anchor="w").grid(row=3, column=0, columnspan=2)

etiket_resim = ttk.Label(pencere)
etiket_resim.grid(row=0, column=2, rowspan=4)

pencere.mainloop()
