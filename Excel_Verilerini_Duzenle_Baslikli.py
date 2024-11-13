#!/usr/bin/env python
# coding: utf-8

# Aynı klasörde bulunan excel dosyalarını birleştiren python kodu.
# Python Kodunun çalışması için bilgisayarınızda "Pandas", "openpyxl" ve "xlrd." kütüphanelerinin / modüllerinin yüklü olması gerekir.

import pandas as pd

########## Gerekli Bilgileri  Düzenleyin ##########
DosyaAdi = "D2.ods"				# kopyalanacak verilerin bulunduğu dosya adı
sayfa_adi = "Sayfa1"			# kopyalanacak verilerin bulunduğu sayfa adı
baslik_satiri = 3				# başlık olarak kullanılacak satır numarası. tamsayı değeri olmalı
ilk_veri_satiri = 5				# kopyalanacak ilk verinin bulunduğu satır numarası. tamsayı değeri olmalı
satir_kopyala = 7				# kopyalanacak verilerin bulunduğu satır sayısı. tamsayı değeri olmalı
sutun_kopyala = "B:G"			# kopyalanacak verilerin bulunduğu sütun aralığı. Örneğin "A:K"
atlanacak_satir_sayisi = 5 		# ilk veri grubu kopyalandıktan sonra ikinci veri grubuna erişmek için atlanacak satır sayısı.tamsayı değeri olmalı
dongu = 15						# veri kopyalarken dosya içerisinde döngüyü kaç kez tekrarlamak istediğinizi belirtin. tamsayı değeri olmalı
######################################################################################################################################################

def baslik(dosya_adi, say_adi=sayfa_adi, ilk_veri=baslik_satiri, sat_sec=1, sut_sec=sutun_kopyala):	# Baslik belirlemek icin kullanilan fonksiyon.
	global sayfa_adi, ilk_veri_satiri, satir_kopyala, sutun_kopyala
	return pd.read_excel(dosya_adi, sheet_name=say_adi, header=None, skiprows=range(0,ilk_veri), nrows=satir_kopyala, usecols=sut_sec)

df_baslik = baslik(DosyaAdi)		# Basligi tespit etmek icin olusturulan df.
baslik = (list(df_baslik.iloc[0]))	# basligin liste biçimi
# # print("Baslik listesi:\n", baslik)

df = pd.DataFrame(columns = baslik)
# # print("Baslikli BOS df:\n", df)

def VeriCercevesi(dosya_adi, say_adi=sayfa_adi, ilk_veri=ilk_veri_satiri, sat_sec=satir_kopyala, sut_sec=sutun_kopyala):      # Belirtilen dosya adına göre, dosya içeriğini Başlıksız DataFrame'e çeviren fonksiyon.
	global sayfa_adi, ilk_veri_satiri, satir_kopyala, sutun_kopyala, baslik
	gecici = pd.read_excel(dosya_adi, sheet_name=say_adi, names=baslik, skiprows=range(1,ilk_veri-1), nrows=satir_kopyala, usecols=sut_sec)
	gecici["Dosya Adi"] = dosya_adi
	return gecici
# # print("VeriCercevesi Fonksiyonu calisti, Sonuc:\n", VeriCercevesi(DosyaAdi))

def tum_veriler(dosya_adi):
	global sayfa_adi, ilk_veri_satiri, satir_kopyala, sutun_kopyala, atlanacak_satir_sayisi, baslik, dongu, df
	for _ in range(dongu):
		try:
			gecici_df = VeriCercevesi(dosya_adi, say_adi=sayfa_adi, ilk_veri=ilk_veri_satiri, sat_sec=satir_kopyala, sut_sec=sutun_kopyala)
			df = pd.concat([df, gecici_df])
			satir_artir = satir_kopyala + atlanacak_satir_sayisi
			ilk_veri_satiri += satir_artir
			df.to_excel("Baslikli_" + DosyaAdi)    # Tüm dosyalar birleştirildikten sonra sonuç "Baslikli_+ DosyaAdi.xlsx" ismi ile kaydedilir.
		except:
			print(f"Dongu tekrarlandı ancak dosyada dongu sayısı kadar veri bulunmuyor. Dongu {dongu}")
			dongu -= 1

tum_veriler(DosyaAdi)

print("\n\nBİRLEŞİM SONRASI df VERİ ÇERÇEVESİ:\n\n", df)
