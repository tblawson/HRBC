# -*- coding: utf-8 -*-
"""
Created on Wed Jul 29 13:21:22 2015

@author: t.lawson
"""
# GMHstuff.py - provides access to dll functions for the GMH sensors

import os
import ctypes as ct
# Change PATH to C:\GMH\GMHdll\
os.environ['GMHPATH'] = 'C:\Software\High Resistance\HRBC\GMHdll'
gmhpath = os.environ['GMHPATH']
GMHLIB = ct.windll.LoadLibrary(os.path.join(gmhpath,'GMH3x32E'))
