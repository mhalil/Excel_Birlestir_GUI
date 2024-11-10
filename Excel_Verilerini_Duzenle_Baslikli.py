#!/usr/bin/env python
# coding: utf-8

# Aynı klasörde bulunan excel dosyalarını birleştiren python kodu.
# Tüm dosya içeriği birleştirildikten sonsa "TUMU.xlsx" adında yeni bir dosyaya kaydedilecek.
# Python Kodunun çalışması için bilgisayarınızda "Pandas", "openpyxl" ve "xlrd." kütüphanelerinin / modüllerinin yüklü olması gerekir.

import pandas as pd

########## Gerekli Bilgileri  Düzenleyin ##########
sayfa_adi = "Sayfa1"					# Veri çerçevesi oluşturulurken excel dosyasındaki hangi sayfadaki (sekmedeki)  verilerin seçimine dair sayfa adi
satir_atla = 2							# Veri çerçevesi oluşturulurken kaç satır atlamak (seçmemek) istiyorsun? tamsayi değeri olmalı
satir_sec = 7							# Veri çerçevesi oluşturulurken kaç satırlık veri seçilsin? tamsayi değeri olmalı
sutun_sec = "B:G"						# Veri çerçevesi oluşturulurken hangi sütun aralığı veri seçilsin?
ikinci_secim_oncesi_atlanacak_satir = 4 # atlanacak (örneğin İmzacı kısımları) satır sayısı
dongu = 10								# Döngüyü kaç kez tekrarlamak istediğinizi belirtin. tamsayi değeri olmalı
DosyaAdi = "D2.ods"					# Düzenlenecek Excel dosyasının adını, uzantısıyla birlikte yazın.
baslik_satiri = 2
######################################################################################################################################################

def baslik(dosya_adi, say_adi=sayfa_adi, sat_atla=baslik_satiri, sat_sec=1, sut_sec=sutun_sec):	# Baslik belirlemek icin kullanilan fonksiyon.
	global sayfa_adi, satir_atla, satir_sec, sutun_sec
	return pd.read_excel(dosya_adi, sheet_name=say_adi, header=None, skiprows=range(0,sat_atla), nrows=satir_sec, usecols=sut_sec)

df_baslik = baslik(DosyaAdi)		# Basligi tespit etmek icin olusturulan df.
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
			satir_artir = satir_sec + ikinci_secim_oncesi_atlanacak_satir
			satir_atla += satir_artir
			df.to_excel("D_" + DosyaAdi)    # Tüm dosyalar birleştirildikten sonra sonuç "D_+ DosyaAdi.xlsx" ismi ile kaydedilir.
		except:
			print(f"Dongu tekrarlandı ancak dosyada dongu sayısı kadar veri bulunmuyor")
			dongu -= 1
	print(f"Dosyada toplam {dongu} dongu var.")

tum_veriler(DosyaAdi)

print("\n\nBİRLEŞİM SONRASI df VERİ ÇERÇEVESİ:\n\n", df)

# # df.to_excel("D_" + DosyaAdi)    # Tüm dosyalar birleştirildikten sonra sonuç "TUMU.xlsx" ismi ile kaydedilir.
