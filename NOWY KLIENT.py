#! /usr/bin/env python
# -*- coding: utf-8 -*-


import imp
import win32api # dialog box

N_K_path = 'J:drop/program/PythonMyScripts/klasy_pomocnicze/NOWY_KLIENT.py'
try: 
	imp.load_source('NOWY KLIENT', N_K_path).n_k()
except:
	win32api.MessageBox(0, "Coś poszlo nie tak", "Nowy Klient", 0x00001000)# pops on TOP
	raise