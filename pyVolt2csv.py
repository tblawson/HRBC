# -*- coding: utf-8 -*-
"""
pyVolt2csv.py

Created on Wed Aug 21 09:28:50 2019

@author: t.lawson
"""

logfile = 'pyVOLT_2019-08-20.log'
with open(logfile, 'r') as log:
    lines = log.readlines()

print len(lines)
print lines[10]
