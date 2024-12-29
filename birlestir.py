from ttkbootstrap import Style
from tkinter import ttk, filedialog, Tk, Toplevel, Label, PhotoImage, IntVar
import glob
import pandas as pd

style = Style()
style = Style(theme='litera')

pencere = style.master
pencere.title("Excel Birleştir")
pencere.geometry("500x500+200+100")

excel_dosyalari = []
yardim = False

########### FONKSİYONLAR ###########

### Klasör Seçme Fonksiyonu
def klasor_sec():
	global excel_dosyalari
	klasor_adi = filedialog.askdirectory()

	if not klasor_adi:
		print("Hiçbir klasör seçilmedi.")
		return

	xls = klasor_adi + "/*.xls*"
	excel_dosyalari = glob.glob(xls)

	# Dosya sayısını bilgi etiketi üzerinde göster
	if excel_dosyalari:
		bilgi.config(text=f"Seçilen klasörde Toplam {len(excel_dosyalari)} adet Excel dosyası bulundu.")
	else:
		bilgi.config(text="Seçilen klasörde hiç Excel dosyası bulunamadı.")

### Excelleri birleştirme Fonksiyonu
def birlestir():
	global excel_dosyalari
	for i in range(len(excel_dosyalari)):
		a = "df_" + str(i)
		# # print(a + "Veri Çerçevesi:")
		a = pd.read_excel(excel_dosyalari[i])
		# # print(a)

### Yardım penceresi görüntüleme Fonksiyonu
def yardim():
	yardim_penceresi = Toplevel()
	yardim_penceresi.geometry("850x700+700+100")
	yardim_penceresi.title(".:: GÖRSEL YARDIM ::.")
	resim = PhotoImage(file="resimler/parametreler.png")
	etiket_resim = Label(yardim_penceresi, image=resim).pack()
	yardim_penceresi.mainloop()


def sayfa_adi_belirt():
	if kontrol_sayfa_adi.get():
		entry_sayfa_adi.config(state='normal')
	else:
		entry_sayfa_adi.config(state='disabled')

def dosya_adi_belirt():
	if kontrol_dosya_adi.get():
		entry_kayit_dosya_adi.config(state='normal')
	else:
		entry_kayit_dosya_adi.config(state='disabled')

########### ARABİRİM OLUSTURULUYOR###########

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
kontrol_sayfa_adi = IntVar()
sayfa_belirt = ttk.Checkbutton(cerceve_parametreler, text='Sayfa Adı Belirt', style='primary.Roundtoggle.Toolbutton', variable=kontrol_sayfa_adi, command=sayfa_adi_belirt)
sayfa_belirt.grid(row=0, column=0, pady=5, padx=25)

entry_sayfa_adi = ttk.Entry(cerceve_parametreler, state='disabled')
entry_sayfa_adi.grid(row=0, column=1, pady=5, padx=25)

etiket_baslik_satir_no = ttk.Label(cerceve_parametreler, text='Başlık Satırı Numarası:').grid(row=1, column=0, pady=5, padx=25)
entry_baslik_satiri = ttk.Entry(cerceve_parametreler).grid(row=1, column=1, pady=5, padx=25)

etiket_ilk_veri_satiri = ttk.Label(cerceve_parametreler, text='İlk Veri Satırı Numarası:').grid(row=2, column=0, pady=5, padx=25)
entry_ilk_veri_satiri = ttk.Entry(cerceve_parametreler).grid(row=2, column=1, pady=5, padx=25)

etiket_kopyalanacak_satir = ttk.Label(cerceve_parametreler, text='Kopyalanacak Satır Sayısı:').grid(row=3, column=0, pady=5, padx=25)
entry_kopyalanacak_satir = ttk.Entry(cerceve_parametreler).grid(row=3, column=1, pady=5, padx=25)

etiket_atlanacak_satir = ttk.Label(cerceve_parametreler, text='Atlanacak Satırı Sayısı:').grid(row=4, column=0, pady=5, padx=25)
entry_atlanacak_satir = ttk.Entry(cerceve_parametreler).grid(row=4, column=1, pady=5, padx=25)

etiket_kopyalanacak_sutun = ttk.Label(cerceve_parametreler, text='Kopyalanacak Sütunlar:').grid(row=5, column=0, pady=5, padx=25)
entry_kopyalanacak_sutun = ttk.Entry(cerceve_parametreler).grid(row=5, column=1, pady=5, padx=25)

etiket_dongu_sayisi = ttk.Label(cerceve_parametreler, text='Döngü Sayısı:').grid(row=6, column=0, pady=5, padx=25)
entry_dongu_sayisi = ttk.Entry(cerceve_parametreler).grid(row=6, column=1, pady=5, padx=25)

kontrol_dosya_adi = IntVar()
kayit_dosya_adi = ttk.Checkbutton(cerceve_parametreler, text='Kayıt için Dosya Adı Belirt', style='primary.Roundtoggle.Toolbutton', variable=kontrol_dosya_adi, command=dosya_adi_belirt).grid(row=7, column=0, pady=5, padx=25)
entry_kayit_dosya_adi = ttk.Entry(cerceve_parametreler, state="disabled")
entry_kayit_dosya_adi.grid(row=7, column=1, pady=5, padx=25)
#####

### Alt Butonlar
buton_yardim = ttk.Button(pencere, text="Görsel Yardımı Aç", style='info.TButton', command=yardim)
buton_yardim.grid(row=2, column=0, pady=5, padx=25)
ttk.Button(pencere, text="Dosyaları Birleştir", style='primary.TButton', command=birlestir).grid(row=2, column=1, pady=5, padx=25)

### Bilgi Etketi
bilgi = Label(pencere, text="Bilgi: Program birleştirme işlemi için hazır...")
bilgi.grid(row=3, column=0, columnspan=2)

pencere.mainloop()
