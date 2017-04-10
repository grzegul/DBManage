#! /usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import imp
import shutil
import datetime
import win32api # dialog box

from openpyxl.drawing.image import Image
from openpyxl import Workbook
from openpyxl import load_workbook

source = u'T:\Projekty\Indukcyjności\Zxx SZABLON'
sciezka = u'T:\Projekty\Indukcyjności'

# -----------------------------	TWORZENIE NOWEGO KLIENTA
def n_k():
	klient = raw_input(u'Podaj nazwę nowego klienta: '.encode(sys.stdout.encoding)) #bez 'encode...' pluje się o 'u'
	klient = klient.upper()
	nazwa_proj = raw_input(u'Podaj nazwę projektu: '.encode(sys.stdout.encoding))

	# zwracanie nowego numeru klienta
	def new_number(src):
		klienci = []
		for item in os.listdir(src):
			temp = item.split("\\")
			x = temp[len(temp)-1]
			try:
				x = int(x[2:5])
			except ValueError:	# Oops, it wasn't an int, and that's fine
				pass
			else:				# It was an int, and now we have the int value
				klienci.append(x)
		return max(klienci)+1


	nowy_nr = new_number(sciezka)
	nr_proj = '01'
	directory = u'%s\PT%d %s\PT%d_%s %s\\' %(sciezka, nowy_nr, klient, nowy_nr, nr_proj, nazwa_proj)

	# kopiowanie katalogu z całą podstrukturą
	COPTYREE_path = 'J:drop/program/PythonMyScripts/klasy_pomocnicze/COPYTREE.py'
	try: 
		imp.load_source('COPTYREE', COPTYREE_path).copytree(source, directory)
	except:
		win32api.MessageBox(0, "Coś poszlo nie tak", "Tworzenie katalogu", 0x00001000)# pops on TOP
		raise

	# -------------------------------------- AKTUALIZACJA KATALOGU
	AKT_KAT_path = 'J:drop/program/PythonMyScripts/klasy_pomocnicze/AKTUALIZACJA_KATALOGU.py'
	try: 
		imp.load_source('AKTUALIZACJA_KATALOGU', AKT_KAT_path).akt_kat(directory)
	except:
		win32api.MessageBox(0, "Coś poszlo nie tak", "Aktualizacja katalogu", 0x00001000)# pops on TOP
		raise