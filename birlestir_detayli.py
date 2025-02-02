from ttkbootstrap import Style
from tkinter import ttk, filedialog, Tk, Toplevel, Label, PhotoImage, IntVar, messagebox
from glob import glob
from pandas import read_excel, DataFrame, concat, ExcelFile
from os import name		# bu satır da silinip kod düzeltilecek. "/" ya da  "\" ibaresi aranacak

style = Style()
style = Style(theme='litera')

pencere = style.master
pencere.title(".:: Excel Dosyalarını Birleştir (Merge Excel Files) ::. [ Mustafa Halil GORENTAS]")
pencere.geometry("500x410+500+250")
pencere.resizable(width=False, height=False)

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
		bilgi.config(text=f"Seçilen klasörde Toplam {len(excel_dosyalari)} adet Excel dosyası bulundu.")

	else:
		bilgi.config(text="Seçilen klasörde hiç Excel dosyası bulunamadı.")


# # ### excel_dosyalari listesini, belirtilen sayfa_adı'na göre guncelle
# # def excel_dosyalari_guncelle(sayfa_adi):
	# # global excel_dosyalari

	# # for dosya in excel_dosyalari:
		# # sayfa_adlari = ExcelFile(dosya).sheet_names
		# # if sayfa_adi not in sayfa_adlari:
			# # excel_dosyalari.remove(dosya)


### Kayıt için Klasör ve Dosya Seçme Fonksiyonu
def kayit_icin_sec():
	kayit_dosya_adi = filedialog.asksaveasfile(filetypes=[('Excel 2007-365', '*.xlsx'), ('Excel 97-2003', '*.xls'), ('All Files', '*.*')])

	if kayit_dosya_adi:
		return kayit_dosya_adi.name


### Yardım penceresi görüntüleme Fonksiyonu
def yardim():
	yardim_penceresi = Toplevel()
	yardim_penceresi.geometry("850x700+600+100")
	yardim_penceresi.title(".:: GÖRSEL YARDIM ::.")
	resim = PhotoImage(file="resimler/parametreler.png")
	etiket_resim = Label(yardim_penceresi, image=resim).pack()
	yardim_penceresi.mainloop()


### Sayfa Adı belirtme/belirtmeme Fonksiyonu
def sayfa_adi_belirt():
	if kontrol_sayfa_adi.get():
		entry_sayfa_adi.config(state='normal')
		# # entry_sayfa_adi.insert(string="birden fazla sayfa için aralarına - koyarak yazın", index=0)
	else:
		entry_sayfa_adi.config(state='disabled')


########### ARABİRİM OLUSTURULUYOR###########

### Klasör seçici
ttk.Label(pencere, text='Excel dosyalarını seç:').grid(row=0, column=0, pady=5, padx=20)
ttk.Button(pencere, text="Birleştirilecek dosyaları seçin...",
					style='primary.TButton',
					command=klasor_sec).grid(row=0, column=1, pady=5, padx=25)

### Çerçeve (Frame)
cerceve_parametreler = ttk.LabelFrame(
    pencere,
    width=400,
    height=400,
    text="Parametreler")
cerceve_parametreler.grid(row=1, column=0, pady=5, padx=25, columnspan=2)

### Parametreler
kontrol_sayfa_adi = IntVar()
kontrol_sayfa_belirt = ttk.Checkbutton(cerceve_parametreler,
										text='Sayfa Adı Belirt',
										style='primary.Roundtoggle.Toolbutton',
										variable=kontrol_sayfa_adi,
										command=sayfa_adi_belirt)
kontrol_sayfa_belirt.grid(row=0, column=0, pady=5, padx=25)

entry_sayfa_adi = ttk.Entry(cerceve_parametreler, state='enable')
entry_sayfa_adi.insert(string=0, index=0)
entry_sayfa_adi.config(state="disabled")
entry_sayfa_adi.grid(row=0, column=1, pady=5, padx=25)

etiket_baslik_satir_no = ttk.Label(cerceve_parametreler,
									text='Başlık Satırı Numarası:')
etiket_baslik_satir_no.grid(row=1, column=0, pady=5, padx=25, sticky='w')

entry_baslik_satiri = ttk.Entry(cerceve_parametreler)
entry_baslik_satiri.insert(string=4, index=0)
entry_baslik_satiri.grid(row=1, column=1, pady=5, padx=25)

etiket_ilk_veri_satiri = ttk.Label(cerceve_parametreler,
									text='İlk Veri Satırı Numarası:')
etiket_ilk_veri_satiri.grid(row=2, column=0, pady=5, padx=25, sticky='w')

entry_ilk_veri_satiri = ttk.Entry(cerceve_parametreler)
entry_ilk_veri_satiri.insert(string=5, index=0)
entry_ilk_veri_satiri.grid(row=2, column=1, pady=5, padx=25)

etiket_kopyalanacak_satir = ttk.Label(cerceve_parametreler,
										text='Veri Barındıran Satır Sayısı:')
etiket_kopyalanacak_satir.grid(row=3, column=0, pady=5, padx=25, sticky='w')

entry_kopyalanacak_satir = ttk.Entry(cerceve_parametreler)
entry_kopyalanacak_satir.insert(string=7, index=0)
entry_kopyalanacak_satir.grid(row=3, column=1, pady=5, padx=25)

etiket_atlanacak_satir = ttk.Label(cerceve_parametreler,
									text='Silinecek Satır Sayısı:')
etiket_atlanacak_satir.grid(row=4, column=0, pady=5, padx=25, sticky='w')

entry_atlanacak_satir = ttk.Entry(cerceve_parametreler)
entry_atlanacak_satir.insert(string=5, index=0)
entry_atlanacak_satir.grid(row=4, column=1, pady=5, padx=25)

etiket_kopyalanacak_sutun = ttk.Label(cerceve_parametreler,
										text='Kopyalanacak Sütunlar:')
etiket_kopyalanacak_sutun.grid(row=5, column=0, pady=5, padx=25, sticky='w')

entry_kopyalanacak_sutun = ttk.Entry(cerceve_parametreler)
entry_kopyalanacak_sutun.insert(string="B:G", index=0)
entry_kopyalanacak_sutun.grid(row=5, column=1, pady=5, padx=25)

#####

########## Gerekli Bilgiler   	##########
# # sayfa_adi = entry_sayfa_adi.get()								# kopyalanacak verilerin bulunduğu sayfa adı
baslik_satiri = int(entry_baslik_satiri.get())					# başlık olarak kullanılacak satır numarası. tamsayı değeri olmalı
ilk_veri_satiri_orj = int(entry_ilk_veri_satiri.get())			# kopyalanacak ilk verinin bulunduğu satır numarası. tamsayı değeri olmalı
satir_kopyala_orj = int(entry_kopyalanacak_satir.get())			# kopyalanacak verilerin bulunduğu satır sayısı. tamsayı değeri olmalı
sutun_kopyala = entry_kopyalanacak_sutun.get()					# kopyalanacak verilerin bulunduğu sütun aralığı. Örneğin "A:K"
atlanacak_satir_sayisi_orj = int(entry_atlanacak_satir.get()) 	# ilk veri grubu kopyalandıktan sonra ikinci veri grubuna erişmek için atlanacak satır sayısı.tamsayı değeri olmalı
# # ####################################################################################
ilk_veri_satiri = ilk_veri_satiri_orj
satir_kopyala = satir_kopyala_orj
atlanacak_satir_sayisi = atlanacak_satir_sayisi_orj
####################################################################################

### BASLİK LİSTESİ FONKSİYONU
def baslik():  # Başlık belirlemek için kullanılan fonksiyon. Fonksiyondaki hata ChatGPT ile cozuldu.
	sutun_kopyala = entry_kopyalanacak_sutun.get().strip()
	# # print("SAYFA ADI:", entry_sayfa_adi.get(), "TURU:", type(entry_sayfa_adi.get()))
	if ":" not in sutun_kopyala:
		bilgi.config(text="Hata: Kopyalanacak sütun aralığını doğru formatta belirtin (örnek: B:G).")
		return []

	secili_sutun = sutun_kopyala if sutun_kopyala else None

	df_g = read_excel(excel_dosyalari[0],
						sheet_name= 0 if entry_sayfa_adi.get() == "0" else entry_sayfa_adi.get(),
						usecols=secili_sutun)		# ****** burası da revize edilecek. sayfa_adi olmayan secenek te eklenecek

	return list(df_g.iloc[baslik_satiri - 2])


### SADECE BİR DOSYA İÇERİSİNDEKİ VERİLERİ TOPLAYAN FONKSİYON
def dosya_verileri(dosya_adi):

	if (kontrol_sayfa_adi.get() == 1):
		try:
			df_g = read_excel(dosya_adi,
								header=None,
								names=baslik(),
								sheet_name= 0 if entry_sayfa_adi.get() == "0" else entry_sayfa_adi.get() ,
								skiprows=range(0,ilk_veri_satiri-1),
								usecols=sutun_kopyala)	# ***** sayfa_adi olmayan secenek te eklenecek - revize
		except:
			messagebox.showwarning(title="Sayfa Adı Hatası",
									message=f"{dosya_adi} dosyası içerisinde '{entry_sayfa_adi.get()}' isimli sayfa bulunmamaktadır")
			df_g = DataFrame()

	else:
		df_g = read_excel(dosya_adi,
							header=None,
							names=baslik(),
							skiprows=range(0,ilk_veri_satiri-1),
							usecols=sutun_kopyala)	# ***** sayfa_adi olmayan secenek te eklenecek - revize

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

	kopyala = satir_kopyala
	atla = atlanacak_satir_sayisi

	i = satir_kopyala
	while i < len(df_satir_sayisi_liste):
		silinecek_satirlar.extend(df_satir_sayisi_liste[i:i+atla])
		# # print(silinecek_satirlar)
		i += (kopyala + atla)

	### TESPİT EDİLEN SATİRLAR SİL.
	df_g.drop(silinecek_satirlar, axis = 0, inplace = True)

	return df_g

### EXCELLERİ BİRLEŞTİRME FONKSİYONU
def birlestir():
	bayrak = True
	df = DataFrame()	# bos bir veri cercevesi

	for excel in excel_dosyalari:
		if bayrak:
			df = dosya_verileri(excel)
			bayrak = False
		else:
			df_g = dosya_verileri(excel)
			df = concat([df, df_g])

	### DOSYAYI KAYDET
	ad = kayit_icin_sec()
	if ad:
		df.to_excel(ad)
	print("İşlem gerçekleşti / iptal edildi")


### ALT BUTONLAR
buton_yardim = ttk.Button(pencere,
							text="Görsel Yardımı Aç",
							style='info.TButton',
							command=yardim)
buton_yardim.grid(row=2, column=0, pady=5, padx=25)

ttk.Button(pencere,
			text="Dosyaları Birleştir ve Kaydet...",
			style='primary.TButton',
			command=birlestir).grid(row=2, column=1, pady=5, padx=25)

### BİLGİ ETKETİ
bilgi = Label(pencere,
				text="Bilgi: Program birleştirme işlemi için hazır...")
bilgi.grid(row=3, column=0, columnspan=2)


pencere.mainloop()

"""
EKLENECEK ÖZELLİKLER:
* Birden fazla sayfa için aralarına ; koyarak yazın.
* widget ve değişkenleri kontrol et kullanılmayanları sil
* entry'lere girilen değerleri kontrol eden yapıyı kur / kontrol et
"""
