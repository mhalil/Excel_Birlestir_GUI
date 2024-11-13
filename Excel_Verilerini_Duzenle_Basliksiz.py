#!/usr/bin/env python
# coding: utf-8

# Aynı klasörde bulunan excel dosyalarını birleştiren python kodu.
# Tüm dosya içeriği birleştirildikten sonsa "TUMU.xlsx" adında yeni bir dosyaya kaydedilecek.
# Python Kodunun çalışması için bilgisayarınızda "Pandas", "openpyxl" ve "xlrd." kütüphanelerinin / modüllerinin yüklü olması gerekir.

import pandas as pd
import os

##### Excel dosyasında yapılacak seçime ait parametrelerin belirlendiği bölüm #####
DosyaAdi = "D1.xlsx"			# kopyalanacak verilerin bulunduğu dosya adı
sayfa_adi = "Sayfa1"			# kopyalanacak verilerin bulunduğu sayfa adı
ilk_veri_satiri = 5				# kopyalanacak ilk verinin bulunduğu satır numarası. tamsayı değeri olmalı
satir_kopyala = 7				# kopyalanacak verilerin bulunduğu satır sayısı. tamsayı değeri olmalı
sutun_kopyala = "B:G"			# kopyalanacak verilerin bulunduğu sütun aralığı. Örneğin "A:K"
atlanacak_satir_sayisi = 5		# ilk veri grubu kopyalandıktan sonra ikinci veri grubuna erişmek için atlanacak satır sayısı.tamsayı değeri olmalı
dongu = 20						# veri kopyalarken dosya içerisinde döngüyü kaç kez tekrarlamak istediğinizi belirtin. tamsayı değeri olmalı
#################################

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

def VeriCercevesiBasliksiz(dosya_adi, say_adi=sayfa_adi, ilk_veri=ilk_veri_satiri, sat_sec=satir_kopyala, sut_sec=sutun_kopyala):      # Belirtilen dosya adına göre, dosya içeriğini Başlıksız DataFrame'e çeviren fonksiyon.
	global sayfa_adi, ilk_veri_satiri, satir_kopyala, sutun_kopyala
	g = pd.read_excel(dosya_adi, sheet_name=say_adi, header=None, skiprows=range(0,ilk_veri-1), nrows=satir_kopyala, usecols=sut_sec)
	g["Dosya Adi"] = dosya_adi
	return g

def tum_veriler(dosya_adi):
	global sayfa_adi, ilk_veri_satiri, satir_kopyala, sutun_kopyala, atlanacak_satir_sayisi, df
	for _ in range(dongu):
		gecici_df = VeriCercevesiBasliksiz(dosya_adi, say_adi=sayfa_adi, ilk_veri=ilk_veri_satiri, sat_sec=satir_kopyala, sut_sec=sutun_kopyala)
		df = pd.concat([df, gecici_df])
		satir_artir = satir_kopyala + atlanacak_satir_sayisi
		ilk_veri_satiri += satir_artir

tum_veriler(DosyaAdi)

print("\n\nBİRLEŞİM SONRASI df VERİ ÇERÇEVESİ:\n\n", df)

df.to_excel("Basliksiz_" + DosyaAdi)    # Tüm veriler alt alta toplandıktan sonra sonuç "Basliksiz_DosyaAdi" ismi ile kaydedilir.
