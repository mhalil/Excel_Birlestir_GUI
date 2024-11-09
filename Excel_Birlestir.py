#!/usr/bin/env python
# coding: utf-8

import pandas as pd
import os

sayfa_adi = "Sayfa1"		# Veri çerçevesi oluşturulurken excel dosyasındaki hangi sayfadaki (sekmedeki)  verilerin seçimine dair sayfa adi

dosyalar = os.listdir()     # "birlestir.py" dosyasının bulunduğu dizindeki (klasördeki) TÜM DOSYA isimlerini, uzantıları ile birlikte al, "dosyalar" isimli listeye ekle / ata.
dosyalar.sort()             # dosyalar listesindeki öğeleri (dosya isimlerini) alfabetik olarak sırala.

if "TUMU.xlsx" in dosyalar:         # Klasör içinde "TUMU.xlsx" dosyasının olup olmadığını kontrol et, varsa aşağıdaki kodları çalıştır.
    os.remove("TUMU.xlsx")          # Klasör içindeki "TUMU.xlsx" isimli dosyayı sil.
    dosyalar.remove("TUMU.xlsx")    # "TUMU.xlsx" isimli öğeyi "dosyalar" listesinden çıkar.

excel_dosyalari= []			# ".xlsx", ".xls" ya da ".ods" uzantılı dosyaların toplanacağı boş liste oluştur.

for i in dosyalar:          # Dizindeki tüm dosya isimlerini kontrol et, ".xlsx", ".xls" ya da ".ods" uzantılı dosyaları "dosya_isimleri" isimli listeye ekle.
    if ((i[-5:] == ".xlsx") or (i[-4:] == ".xls") or (i[-4:] == ".ods") ):     # dosya uzantılarını kontrol et.
        excel_dosyalari.append(i)

print("Excel dosyalari:\n", excel_dosyalari)

df = pd.DataFrame()			# Birleştirme için olusturulan bos veri cercevesi.
print("\nBoş df:\n", df)

for dosya in excel_dosyalari:
	g_df = pd.read_excel(dosya, sheet_name=sayfa_adi)
	g_df["Dosya Adi"] = dosya
	df = pd.concat([df, g_df])

print("\nBirleştirme sonrası df:\n", df)

df.to_excel("TUMU.xlsx")
