from ttkbootstrap import Style
from tkinter import ttk

style = Style()
style = Style(theme='litera')

pencere = style.master
pencere.title("Excel Birleştir")
pencere.geometry("500x500")

### FONKSİYONLAR
def dosya_sec():
	print("dosya secme diyalogu acilacak")


def birlestir():
	print("Birlestirme islemi basladı ve bitti")

def yardim():
	print("Görsel parametre yardımı goruntulenecek")

### Dsoya / Klasör seçici
ttk.Label(pencere, text='Excel dosyalarını seç:').grid(row=0, column=0, pady=5, padx=20)
ttk.Button(pencere, text="...", style='primary.TButton', command=dosya_sec).grid(row=0, column=1, pady=5, padx=25)

cerceve_parametreler = ttk.LabelFrame(
    pencere,
    relief="ridge",
    width=400,
    height=400,
    text="Parametreler")
cerceve_parametreler.grid(row=1, column=0, pady=5, padx=25, columnspan=2)


### Parametreler
sayfa_belirt = ttk.Checkbutton(cerceve_parametreler, text='Sayfa Adı Belirt', style='primary.Roundtoggle.Toolbutton').grid(row=0, column=0, pady=5, padx=25)
entry_sayfa_adi = ttk.Entry(cerceve_parametreler).grid(row=0, column=1, pady=5, padx=25)

etiket_baslik_satir_no = ttk.Label(cerceve_parametreler, text='Başlık Satırı Numarası:').grid(row=1, column=0, pady=5, padx=25)
entry_baslik_satiri = ttk.Entry(cerceve_parametreler).grid(row=1, column=1, pady=5, padx=25)

etiket_baslik_satir_no = ttk.Label(cerceve_parametreler, text='İlk Veri Satırı Numarası:').grid(row=2, column=0, pady=5, padx=25)
entry_baslik_satiri = ttk.Entry(cerceve_parametreler).grid(row=2, column=1, pady=5, padx=25)

etiket_baslik_satir_no = ttk.Label(cerceve_parametreler, text='Kopyalanacak Satır Sayısı:').grid(row=3, column=0, pady=5, padx=25)
entry_baslik_satiri = ttk.Entry(cerceve_parametreler).grid(row=3, column=1, pady=5, padx=25)

etiket_baslik_satir_no = ttk.Label(cerceve_parametreler, text='Atlanacak Satırı Numarası:').grid(row=4, column=0, pady=5, padx=25)
entry_baslik_satiri = ttk.Entry(cerceve_parametreler).grid(row=4, column=1, pady=5, padx=25)

etiket_baslik_satir_no = ttk.Label(cerceve_parametreler, text='Kopyalanacak Sütunlar:').grid(row=5, column=0, pady=5, padx=25)
entry_baslik_satiri = ttk.Entry(cerceve_parametreler).grid(row=5, column=1, pady=5, padx=25)

etiket_baslik_satir_no = ttk.Label(cerceve_parametreler, text='Döngü Sayısı:').grid(row=6, column=0, pady=5, padx=25)
entry_baslik_satiri = ttk.Entry(cerceve_parametreler).grid(row=6, column=1, pady=5, padx=25)

sayfa_belirt = ttk.Checkbutton(cerceve_parametreler, text='Dosya Adı Belirt', style='primary.Roundtoggle.Toolbutton').grid(row=7, column=0, pady=5, padx=25)
entry_sayfa_adi = ttk.Entry(cerceve_parametreler).grid(row=7, column=1, pady=5, padx=25)
#####

ttk.Button(pencere, text="Görsel Yardım", style='info.TButton', command=yardim).grid(row=2, column=0, pady=5, padx=25)
ttk.Button(pencere, text="Dosyaları Birleştir", style='primary.TButton', command=birlestir).grid(row=2, column=1, pady=5, padx=25)
pencere.mainloop()
