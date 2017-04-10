#! /usr/bin/env python
# -*- coding: utf-8 -*-


import imp
import win32api # dialog box

N_P_path = 'J:drop/program/PythonMyScripts/klasy_pomocnicze/NOWY_PROJEKT.py'
try: 
	imp.load_source('NOWY PROJEKT', N_P_path).n_p()
except:
	win32api.MessageBox(0, "Coś poszlo nie tak", "Nowy Klient", 0x00001000)# pops on TOP
	raise