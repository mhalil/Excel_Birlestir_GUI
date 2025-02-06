from ttkbootstrap import Style
from tkinter import ttk, filedialog, Tk, Toplevel, Label, PhotoImage, IntVar, messagebox
from glob import glob
from pandas import read_excel, DataFrame, concat, ExcelFile
from os import name		# bu satır da silinip kod düzeltilecek. "/" ya da  "\" ibaresi aranacak

style = Style()
style = Style(theme = 'litera')

pencere = style.master
pencere.title(".:: Excel Dosyalarını Birleştir (Merge Excel Files) ::.")
pencere.geometry("500x410+500+250")
pencere.resizable(width = False, height = False)

excel_dosyalari = []
yardim = False

########### FONKSİYONLAR ###########

### Klasör Seçme Fonksiyonu
def klasor_sec():
	global excel_dosyalari
	klasor_adi = filedialog.askdirectory()

	if klasor_adi:
		xls = klasor_adi + "/*.xls*"
		excel_dosyalari = glob(xls)	# liste veri yapısı. Tam yol ve dosya adı icerir.

	### Dosya sayısını bilgi etiketi üzerinde göster
	if excel_dosyalari:
		bilgi.config(text = f"Seçilen klasörde Toplam {len(excel_dosyalari)} adet Excel dosyası bulundu.")

	else:
		bilgi.config(text = "Seçilen klasörde hiç Excel dosyası bulunamadı.")


### excel_dosyalari listesini, belirtilen sayfa_adı'na göre guncelle
def excel_sayfa_adlari(excel_dosyasi):
	sayfa_adlari = ExcelFile(excel_dosyasi).sheet_names
	return sayfa_adlari

### Kayıt için Klasör ve Dosya Seçme Fonksiyonu
def kayit_icin_sec():
	kayit_dosya_adi = filedialog.asksaveasfile(filetypes = [('Excel 2007-365', '*.xlsx'), ('Excel 97-2003', '*.xls'), ('All Files', '*.*')])

	if kayit_dosya_adi:
		return kayit_dosya_adi.name


### Yardım penceresi görüntüleme Fonksiyonu
def yardim():
	yardim_penceresi = Toplevel()
	yardim_penceresi.geometry("850x700+600+100")
	yardim_penceresi.title(".:: GÖRSEL YARDIM ::.")
	resim = PhotoImage(file = "resimler/parametreler.png")
	etiket_resim = Label(yardim_penceresi, image = resim).pack()
	yardim_penceresi.mainloop()


### Sayfa Adı belirtme/belirtmeme Fonksiyonu
def sayfa_adi_belirt():
	if kontrol_sayfa_adi.get():
		entry_sayfa_adi.config(state = 'normal')

	else:
		entry_sayfa_adi.config(state = 'disabled')


########### ARABİRİM OLUSTURULUYOR###########

### Klasör seçici
ttk.Label(pencere, text = 'Excel dosyalarını seç:').grid(row = 0, column = 0, pady = 5, padx = 20)
ttk.Button(pencere,
			text = "Birleştirilecek dosyaları seçin...",
			style = 'primary.TButton',
			command = klasor_sec).grid(row = 0, column = 1, pady = 5, padx = 25)

### Çerçeve (Frame)
cerceve_parametreler = ttk.LabelFrame(pencere,
										width = 400,
										height = 400,
										text = "Parametreler")
cerceve_parametreler.grid(row = 1, column = 0, pady = 5, padx = 25, columnspan = 2)

### Parametreler
kontrol_sayfa_adi = IntVar()
kontrol_sayfa_belirt = ttk.Checkbutton(cerceve_parametreler,
										text = 'Sayfa Adı Belirt',
										style = 'primary.Roundtoggle.Toolbutton',
										variable = kontrol_sayfa_adi,
										command = sayfa_adi_belirt)
kontrol_sayfa_belirt.grid(row = 0, column = 0, pady = 5, padx = 25)

entry_sayfa_adi = ttk.Entry(cerceve_parametreler)
# # entry_sayfa_adi.insert(string = 0, index = 0)
entry_sayfa_adi.config(state = "disabled")
entry_sayfa_adi.grid(row = 0, column = 1, pady = 5, padx = 25)

etiket_baslik_satir_no = ttk.Label(cerceve_parametreler,
									text = 'Başlık Satırı Numarası:')
etiket_baslik_satir_no.grid(row = 1, column = 0, pady = 5, padx = 25, sticky = 'w')

entry_baslik_satiri = ttk.Entry(cerceve_parametreler)
entry_baslik_satiri.insert(string = 4, index = 0)
entry_baslik_satiri.grid(row = 1, column = 1, pady = 5, padx = 25)

etiket_ilk_veri_satiri = ttk.Label(cerceve_parametreler,
									text = 'İlk Veri Satırı Numarası:')
etiket_ilk_veri_satiri.grid(row = 2, column = 0, pady = 5, padx = 25, sticky = 'w')

entry_ilk_veri_satiri = ttk.Entry(cerceve_parametreler)
entry_ilk_veri_satiri.insert(string = 5, index = 0)
entry_ilk_veri_satiri.grid(row = 2, column = 1, pady = 5, padx = 25)

etiket_kopyalanacak_satir = ttk.Label(cerceve_parametreler,
										text = 'Veri Barındıran Satır Sayısı:')
etiket_kopyalanacak_satir.grid(row = 3, column = 0, pady = 5, padx = 25, sticky = 'w')

entry_kopyalanacak_satir = ttk.Entry(cerceve_parametreler)
entry_kopyalanacak_satir.insert(string = 7, index = 0)
entry_kopyalanacak_satir.grid(row = 3, column = 1, pady = 5, padx = 25)

etiket_atlanacak_satir = ttk.Label(cerceve_parametreler,
									text = 'Silinecek Satır Sayısı:')
etiket_atlanacak_satir.grid(row = 4, column = 0, pady = 5, padx = 25, sticky = 'w')

entry_atlanacak_satir = ttk.Entry(cerceve_parametreler)
entry_atlanacak_satir.insert(string = 5, index = 0)
entry_atlanacak_satir.grid(row = 4, column = 1, pady = 5, padx = 25)

etiket_kopyalanacak_sutun = ttk.Label(cerceve_parametreler,
										text = 'Kopyalanacak Sütunlar:')
etiket_kopyalanacak_sutun.grid(row = 5, column = 0, pady = 5, padx = 25, sticky = 'w')

entry_kopyalanacak_sutun = ttk.Entry(cerceve_parametreler)
entry_kopyalanacak_sutun.insert(string = "B:G", index = 0)
entry_kopyalanacak_sutun.grid(row = 5, column = 1, pady = 5, padx = 25)

#####

### BASLİK LİSTESİ FONKSİYONU
def baslik(dosya):  # Başlık belirlemek için kullanılan fonksiyon. Fonksiyondaki hata ChatGPT ile cozuldu.
# # def baslik(dosya = "/home/halil/Documents/GitHub/Excel_Birlestir_GUI/D1.xlsx"):  # silinecek
	sayfa_adi = entry_sayfa_adi.get()
	kopyalanacak_sutun = entry_kopyalanacak_sutun.get().strip() if entry_kopyalanacak_sutun.get().strip() else None
	baslik_satiri = int(entry_baslik_satiri.get()) - 2
	baslik = list()

	try:
		if kontrol_sayfa_adi.get() == 0:
			df_g = read_excel(dosya,
								sheet_name = 0,
								usecols = kopyalanacak_sutun)
			# # print("BASLIK SATIRI", baslik_satiri)		# silinecek
			baslik = list(df_g.iloc[baslik_satiri])

		elif kontrol_sayfa_adi.get() == 1 and sayfa_adi in excel_sayfa_adlari(dosya):
			df_g = read_excel(dosya,
								sheet_name = 0 if sayfa_adi == "0" else sayfa_adi,
								usecols = kopyalanacak_sutun)
			baslik = list(df_g.iloc[baslik_satiri])
		print(dosya, "Dosyasına ait B A S L I K:", baslik)		# silinecek
		return baslik

	except:
		messagebox.showwarning(title = "Sayfa Adı Hatası",
									message = f"HATA: {dosya} Dosyasında {sayfa_adi} sayfa adi  belirtilmemiş ya da mevcut olmayabilir. Veyahut ilk sayfada veri olmayabilir.")

### SADECE BİR DOSYA İÇERİSİNDEKİ VERİLERİ TOPLAYAN FONKSİYON
def dosya_verileri(dosya_adi):
# # def dosya_verileri(dosya_adi = "/home/halil/Documents/GitHub/Excel_Birlestir_GUI/D2.xlsx"):		# silinecek
	sayfa_adi = entry_sayfa_adi.get()
	ilk_veri_satiri = int(entry_ilk_veri_satiri.get())
	kopyalanacak_sutun = entry_kopyalanacak_sutun.get().strip()

	if (kontrol_sayfa_adi.get() == 1):
		try:
			df_g = read_excel(dosya_adi,
								header = None,
								names = baslik(dosya_adi),
								sheet_name = 0 if sayfa_adi == "0" else sayfa_adi,
								skiprows = range(0, ilk_veri_satiri - 1),
								usecols = kopyalanacak_sutun)	# ***** sayfa_adi olmayan secenek te eklenecek - revize
		except:
			messagebox.showwarning(title = "Sayfa Adı Hatası",
									message = f"{dosya_adi} dosyası içerisinde '{sayfa_adi}' isimli sayfa bulunmamaktadır")
			df_g = DataFrame()

	else:
		df_g = read_excel(dosya_adi,
							header = None,
							names = baslik(dosya_adi),
							skiprows = range(0, ilk_veri_satiri - 1),
							usecols = kopyalanacak_sutun)	# ***** sayfa_adi olmayan secenek te eklenecek - revize

	dosyanin_adi = ""
	isl_sistemi = name

	if isl_sistemi == "posix":
		dosyanin_adi = dosya_adi.rsplit("/", 1)[1]

	else:
		dosyanin_adi = dosya_adi.rsplit("\\", 1)[1]

	df_g["Dosya ADI"] = dosyanin_adi

	### SİLİNECEK SATIR NUMARALARINI TESPİT ET.
	df_satir_sayisi_liste = list(range(df_g.shape[0]))
	silinecek_satirlar = []

	kopyala = int(entry_kopyalanacak_satir.get())
	atla = int(entry_atlanacak_satir.get())

	i = int(entry_kopyalanacak_satir.get())
	while i < len(df_satir_sayisi_liste):
		silinecek_satirlar.extend(df_satir_sayisi_liste[i:i+atla])
		i += (kopyala + atla)

	### TESPİT EDİLEN SATİRLAR SİL.
	df_g.drop(silinecek_satirlar, axis = 0, inplace = True)

	print("VERİ ÇERÇEVESİ:\n", df_g)
	return df_g

### EXCELLERİ BİRLEŞTİRME FONKSİYONU
def birlestir():
	bayrak = True
	df = DataFrame()	# bos bir veri cercevesi

	if len(excel_dosyalari) > 0:
		for excel_dosyasi in excel_dosyalari:
			if bayrak:
				df = dosya_verileri(excel_dosyasi)
				bayrak = False
			else:
				df_g = dosya_verileri(excel_dosyasi)
				df = concat([df, df_g])

		### DOSYAYI KAYDET
		ad = kayit_icin_sec()
		if ad:
			df.to_excel(ad)
		print("İşlem gerçekleşti / iptal edildi")

	else:
		messagebox.showwarning(title = "Dosya Seçimi Hatası",
									message = "Seçili Excel dosyaları bulunmamaktadır.\nÖncelikle Birleştirilecek dosyaları seçmelisiniz.")

### ALT BUTONLAR
buton_yardim = ttk.Button(pencere,
							text = "Görsel Yardımı Aç",
							style = 'info.TButton',
							command = yardim)
buton_yardim.grid(row = 2, column = 0, pady = 5, padx = 25)

ttk.Button(pencere,
			text = "Dosyaları Birleştir ve Kaydet...",
			style = 'primary.TButton',
			# # command = baslik).grid(row = 2, column = 1, pady = 5, padx = 25)
			# # command = dosya_verileri).grid(row = 2, column = 1, pady = 5, padx = 25)
			command = birlestir).grid(row = 2, column = 1, pady = 5, padx = 25)

### BİLGİ ETKETİ
bilgi = Label(pencere,
				text = "Bilgi: Program birleştirme işlemi için hazır...")
bilgi.grid(row = 3, column = 0, columnspan = 2)


pencere.mainloop()

"""
Eklenecekler:
* baslik satırı en üstte olursa, kodda ne değişecek incele.
* birden fazla sayfa adı ile birleştirme yapmak istenirse ...
"""
