#!/usr/bin/env python
# coding: utf-8

# Aynı klasörde bulunan excel dosyalarını birleştiren python kodu.
# Tüm dosya içeriği birleştirildikten sonsa "TUMU.xlsx" adında yeni bir dosyaya kaydedilecek.
# Python Kodunun çalışması için bilgisayarınızda "Pandas", "openpyxl" ve "xlrd." kütüphanelerinin / modüllerinin yüklü olması gerekir.

import pandas as pd
import os

dosyalar = os.listdir()     # Bu Python dosyasının bulunduğu dizindeki (klasördeki) TÜM DOSYA isimlerini, uzantıları ile birlikte al, "dosyalar" isimli listeye ekle / ata.
dosyalar.sort()             # dosyalar listesindeki öğeleri (dosya isimlerini) alfabetik olarak sırala.
# # print(dosyalar)			# Kontrol amacli

def temizle():
	if "TUMU.xlsx" in dosyalar:         # Klasör içinde "TUMU.xlsx" dosyasının olup olmadığını kontrol et, varsa aşağıdaki kodları çalıştır.
		os.remove("TUMU.xlsx")          # Klasör içindeki "TUMU.xlsx" isimli dosyayı sil.
		dosyalar.remove("TUMU.xlsx")    # "TUMU.xlsx" isimli öğeyi "dosyalar" listesinden çıkar.

temizle()					# temizle isimli fonksiyonu çalıştır

excel_dosya_isimleri= []	# ".xlsx", ".xls" ya da ".ods" uzantılı dosyaların toplanacağı boş liste oluştur.

def excel_dosyalari():
	global dosyalar
	for i in dosyalar:              # Dizindeki tüm dosya isimlerini kontrol et, ".xlsx", ".xls" ya da ".ods" uzantılı dosyaları "excel_dosya_isimleri" isimli listeye ekle.
		if ((i[-5:] == ".xlsx") or (i[-4:] == ".xls") or (i[-4:] == ".ods") ):     # dosya uzantılarını kontrol et.
			excel_dosya_isimleri.append(i)

excel_dosyalari()					# excel_dosyalari isimli fonksiyonu çalıştır
# # print(excel_dosya_isimleri)			# Kontrol amacli

def VeriCercevesi():       	# Boş bir DataFrame oluşturan fonksiyon.
    return pd.DataFrame()

df = VeriCercevesi()		# Birleşim için hazır bekleyen boş veri çerçevesi
# # print("BOŞ VERİ ÇERÇEVESİ:\n", df)					# Kontrol amacli

sayfa_adi = "Sayfa1"			# Veri çerçevesi oluşturulurken excel dosyasındaki hangi sayfadaki (sekmedeki)  verilerin seçimine dair sayfa adi
satir_atla = 3					# Veri çerçevesi oluşturulurken kaç satır atlamak (seçmemek) istiyorsun? tamsayi değeri olmalı
satir_sec = 7					# Veri çerçevesi oluşturulurken kaç satırlık veri seçilsin? tamsayi değeri olmalı
sutun_sec = "B:G"				# Veri çerçevesi oluşturulurken hangi sütun aralığı veri seçilsin?
ikinci_secim_oncesi_atlanacak_satir = 4
dongu = 20						# Döngüyü kaç kez tekrarlamak istediğinizi belirtin. tamsayi değeri olmalı
DosyaAdi = "D3.ods"


def VeriCercevesiBasliksiz(dosya_adi, say_adi=sayfa_adi, sat_atla=satir_atla, sat_sec=satir_sec, sut_sec=sutun_sec):      # Belirtilen dosya adına göre, dosya içeriğini Başlıksız DataFrame'e çeviren fonksiyon.
	global sayfa_adi, satir_atla, satir_sec, sutun_sec
	g = pd.read_excel(dosya_adi, sheet_name=say_adi, header=None, skiprows=range(0,sat_atla), nrows=satir_sec, usecols=sut_sec)
	g["Dosya Adi"] = dosya_adi
	return g

def tum_veriler(dosya_adi):
	global sayfa_adi, satir_atla, satir_sec, sutun_sec, ikinci_secim_oncesi_atlanacak_satir, df
	for _ in range(dongu):
		gecici_df = VeriCercevesiBasliksiz(dosya_adi, say_adi=sayfa_adi, sat_atla=satir_atla, sat_sec=satir_sec, sut_sec=sutun_sec)
		df = pd.concat([df, gecici_df])
		satir_artir = satir_sec + ikinci_secim_oncesi_atlanacak_satir
		satir_atla += satir_artir

tum_veriler(DosyaAdi)

print("\n\nBİRLEŞİM SONRASI df VERİ ÇERÇEVESİ:\n\n", df)

df.to_excel("Basliksiz_" + DosyaAdi)    # Tüm dosyalar birleştirildikten sonra sonuç "TUMU.xlsx" ismi ile kaydedilir.
