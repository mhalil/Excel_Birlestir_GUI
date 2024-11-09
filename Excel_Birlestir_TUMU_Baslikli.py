#!/usr/bin/env python
# coding: utf-8

# Aynı klasörde bulunan excel dosyalarını birleştiren python kodu.
# Tüm dosya içeriği birleştirildikten sonsa "TUMU_Baslikli.xlsx" adında yeni bir dosyaya kaydedilecek.
# Python Kodunun çalışması için bilgisayarınızda "Pandas", "openpyxl" ve "xlrd." kütüphanelerinin / modüllerinin yüklü olması gerekir.

import pandas as pd
import os

########## Gerekli Bilgileri  Düzenleyin ##########
sayfa_adi = "Sayfa1"						# Veri çerçevesi oluşturulurken excel dosyasındaki hangi sayfadaki (sekmedeki)  verilerin seçimine dair sayfa adi
satir_atla_orj = 2							# Veri çerçevesi oluşturulurken kaç satır atlamak (seçmemek) istiyorsun? tamsayi değeri olmalı
satir_sec_orj = 7							# Veri çerçevesi oluşturulurken kaç satırlık veri seçilsin? tamsayi değeri olmalı
sutun_sec = "B:G"							# Veri çerçevesi oluşturulurken hangi sütun aralığı veri seçilsin?
ikinci_secim_oncesi_atlanacak_satir_orj = 4 # atlanacak (örneğin İmzacı kısımları) satır sayısı
dongu_orj = 20									# Döngüyü kaç kez tekrarlamak istediğinizi belirtin. tamsayi değeri olmalı
baslik_satiri = 2
####################################################################################
satir_atla = satir_atla_orj
satir_sec = satir_sec_orj
ikinci_secim_oncesi_atlanacak_satir = ikinci_secim_oncesi_atlanacak_satir_orj
dongu = dongu_orj
####################################################################################

dosyalar = os.listdir()     # "birlestir.py" dosyasının bulunduğu dizindeki (klasördeki) TÜM DOSYA isimlerini, uzantıları ile birlikte al, "dosyalar" isimli listeye ekle / ata.
dosyalar.sort()             # dosyalar listesindeki öğeleri (dosya isimlerini) alfabetik olarak sırala.

if "TUMU_Baslikli.xlsx" in dosyalar:         # Klasör içinde "TUMU_Baslikli.xlsx" dosyasının olup olmadığını kontrol et, varsa aşağıdaki kodları çalıştır.
    os.remove("TUMU_Baslikli.xlsx")          # Klasör içindeki "TUMU_Baslikli.xlsx" isimli dosyayı sil.
    dosyalar.remove("TUMU_Baslikli.xlsx")    # "TUMU_Baslikli.xlsx" isimli öğeyi "dosyalar" listesinden çıkar.

excel_dosyalari= []			# ".xlsx", ".xls" ya da ".ods" uzantılı dosyaların toplanacağı boş liste oluştur.

for i in dosyalar:          # Dizindeki tüm dosya isimlerini kontrol et, ".xlsx", ".xls" ya da ".ods" uzantılı dosyaları "dosya_isimleri" isimli listeye ekle.
    if ((i[-5:] == ".xlsx") or (i[-4:] == ".xls") or (i[-4:] == ".ods") ):     # dosya uzantılarını kontrol et.
        excel_dosyalari.append(i)
print("\nExcel dosyalari:\n", excel_dosyalari)

def baslik(dosya_adi, say_adi=sayfa_adi, sat_atla=baslik_satiri, sat_sec=1, sut_sec=sutun_sec):	# Baslik belirlemek icin kullanilan fonksiyon.
	global sayfa_adi, satir_atla, satir_sec, sutun_sec
	return pd.read_excel(dosya_adi, sheet_name=say_adi, header=None, skiprows=range(0,sat_atla), nrows=satir_sec, usecols=sut_sec)

df_baslik = baslik(excel_dosyalari[0])		# Basligi tespit etmek icin olusturulan df.
# # print(df_baslik)
baslik = (list(df_baslik.iloc[0]))
# # print("Baslik listesi:\n", baslik)

df = pd.DataFrame(columns = baslik)
# # print("Baslikli BOS df:\n", df)

def VeriCercevesi(dosya_adi, say_adi=sayfa_adi, sat_atla=satir_atla, sat_sec=satir_sec-1, sut_sec=sutun_sec):      # Belirtilen dosya adına göre, dosya içeriğini Başlıksız DataFrame'e çeviren fonksiyon.
	global sayfa_adi, satir_atla, satir_sec, sutun_sec, baslik
	g = pd.read_excel(dosya_adi, sheet_name=say_adi, names=baslik, skiprows=range(0,sat_atla), nrows=satir_sec, usecols=sut_sec)
	g["Dosya Adi"] = dosya_adi
	return g
# # print("VeriCercevesi Fonksiyonu calisti, Sonuc:\n", VeriCercevesi(DosyaAdi))

def tum_veriler(dosya_adi):
	global sayfa_adi, satir_atla, satir_sec, sutun_sec, ikinci_secim_oncesi_atlanacak_satir, baslik, dongu, df
	for _ in range(dongu):
		try:
			gecici_df = VeriCercevesi(dosya_adi, say_adi=sayfa_adi, sat_atla=satir_atla, sat_sec=satir_sec, sut_sec=sutun_sec)
			# # print(gecici_df)
			df = pd.concat([df, gecici_df])
			# # print("\ndongu sonrası df:\n", df)
			satir_artir = satir_sec + ikinci_secim_oncesi_atlanacak_satir
			satir_atla += satir_artir
			df.to_excel("TUMU_Baslikli.xlsx")    # Tüm dosyalar birleştirildikten sonra sonuç "TUMU_Baslikli.xlsx" ismi ile kaydedilir.
		except:
			print(f"Dongu tekrarlandı ancak {dosya_adi} dosyasında dongu sayısı kadar veri bulunmuyor")
			dongu -= 1
	print(f"Dosyada toplam {dongu} dongu var.")

for dosya in excel_dosyalari:
	tum_veriler(dosya)
	satir_atla = satir_atla_orj							# Baslangic degerlerine geri don
	satir_sec = satir_sec_orj							# Baslangic degerlerine geri don
	ikinci_secim_oncesi_atlanacak_satir = ikinci_secim_oncesi_atlanacak_satir_orj # Baslangic degerlerine geri don
	dongu = dongu_orj

print("\n\nBİRLEŞİM SONRASI VERİ ÇERÇEVESİ:\n\n", df)
