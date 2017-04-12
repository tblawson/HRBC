# -*- coding: utf-8 -*-
"""
Created on Fri Nov 27 13:34:33 2015

@author: t.lawson
"""

import visa
import time

t = 0.1

RM = visa.ResourceManager()
print RM,'\n'


i = RM.open_resource('GPIB0::22::INSTR') # 'GPIB0::22::INSTR','ASRL5::INSTR'
print 'Opened session',i.session,'\n'
i.write_termination = '\r\n' # carriage return, line feed
i.read_termination = '\r\n' # carriage return, line feed
#i.timeout = 1000 # default 1 s timeout

print 'Reading :',i.read(),'\n'
time.sleep(t)

s = "ID?"
print 'Querying with %s :'%s,i.query(s),'\n'
time.sleep(t)

print 'Reading :',i.read(),'\n'
time.sleep(t)

s = "DCV,10"
print 'Sending %s :'%s,i.write(s),'\n'
time.sleep(t)
s = "LFREQ LINE"
print 'Sending %s :'%s,i.write(s),'\n'
time.sleep(t)

print 'Reading :',i.read(),'\n'
time.sleep(t)

print 'Closing session',i.session
i.close()

RM.close()