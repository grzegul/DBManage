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
sciezka = u'T:\Projekty\Indukcyjności\\' + "\\".join((os.path.dirname(sys.argv[0])).split("\\")[3:])
#os.chdir(os.path.dirname(sys.argv[0]))	# wykonanie w katalogu, w którym się znajduje

# -----------------------------	TWORZENIE NOWEGO PROKEKTU
def n_p_d():
	nazwa_proj = raw_input(u'Podaj nazwę projektu: '.encode(sys.stdout.encoding))

	# zwracanie nowego numeru klienta
	def new_number(src):
		projekty = []
		for item in os.listdir(src):
			temp = item.split("\\")
			x = temp[len(temp)-1]
			try:
				x = int(x[6:9])
			except ValueError:	# Oops, it wasn't an int, and that's fine
				pass
			else:				# It was an int, and now we have the int value
				projekty.append(x)
		return max(projekty)+1


	nr_proj = str(new_number(sciezka)).zfill(2)
	nr_kli = (sciezka.split("\\")[len(sciezka.split("\\"))-1]).split(" ")[0]
	kli = (" ".join((sciezka.split("\\")[len(sciezka.split("\\"))-1]).split(" ")[1:])).strip()

	directory = u'%s\%s_%s %s\\' %(sciezka, nr_kli, nr_proj, nazwa_proj)

	# kopiowanie katalogu z całą podstrukturą
	COPYTREE_path = 'J:drop/program/PythonMyScripts/klasy_pomocnicze/COPYTREE.py'
	try: 
		imp.load_source('COPYTREE', COPYTREE_path).copytree(source, directory)
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