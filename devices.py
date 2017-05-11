# -*- coding: utf-8 -*-
"""
devices.py

Gathers together all info required to use external instruments.

All device data is collected in a 'dictionary of dictionaries' - INSTR_DATA
Each piece of information is accessed as:
INSTR_DATA[<instrument description>][<parameter>].
E.g. the 'set function' string for the HP3458(s/n518) is:
INSTR_DATA['DVM: HP3458A, s/n518']['setfn_str'],
which retrieves the string 'FUNC OHMF;OCOMP ON'.
Note that some of the strings are LISTS of strings (i.e. multiple commands)

Created on Fri Mar 17 13:52:15 2017

@author: t.lawson
"""

import numpy as np
import os
import ctypes as ct
import visa


INSTR_DATA = {} # Dictionary of instrument parameter dictionaries, keyed by description
DESCR = []
sublist = []
ROLES_WIDGETS = {} # Dictionary of GUI widgets keyed by role
ROLES_INSTR = {} # Dictionary of GMH_sensor or Instrument objects keyed by role

"""
VISA-specific stuff:
Only ONE VISA resource manager is required at any time -
All comunications for all GPIB and RS232 instruments (except GMH)
are handled by RM.
"""
RM = visa.ResourceManager()

# Switchbox
SWITCH_CONFIGS = {'V1':'A','Vd1':'C','Vd2':'D','V2':'B'}

T_Sensors = ('none','Pt','SR104t','thermistor')

"""
---------------------------------------------------------------
GMH-specific stuff:
GMH probe communications are handled by low-level routines in GMHdll.dll.
"""
os.environ['GMHPATH'] = 'I:\MSL\Private\Electricity\Staff\TBL\Python\High_Res_Bridge\GMHdll'  # 'C:\Software\High Resistance\HRBC\GMHdll'
gmhpath = os.environ['GMHPATH']
GMHLIB = ct.windll.LoadLibrary(os.path.join(gmhpath,'GMH3x32E'))
GMH_DESCR = ('GMH, s/n627',
             'GMH, s/n628')
             
'''--------------------------------------------------------------'''

class device():
    """
    A generic external device or instrument
    """
    def __init__(self,demo = True):
        self.demo = demo
    
    def open(self):
        pass
    
    def close(self):
        pass



class GMH_Sensor(device):
    """
    A class to wrap around the low-level functions of GMH3x32E.dll. 
    For use with most Greisinger GMH devices.
    """
    def __init__(self,descr,demo = True):
        self.Descr = descr
        self.demo = demo
        
        self.addr = INSTR_DATA[self.Descr]['addr'] # COM port-number assigned to USB 3100N adapter cable
        self.str_addr = INSTR_DATA[self.Descr]['str_addr']
        self.role = INSTR_DATA[self.Descr]['role']        
        
        self.Prio = ct.c_short()
        self.flData = ct.c_double() # Don't change this type!! It's the exactly right one!
        self.intData = ct.c_long()
        self.meas_str = ct.create_string_buffer(30)
        self.unit_str = ct.create_string_buffer(10)
        self.lang_offset = ct.c_int16(4096) # English language-offset
        self.MeasFn = ct.c_short(180) # GetMeasCode()
        self.UnitFn = ct.c_int16(178) # GetUnitCode()
        self.ValFn = ct.c_short(0) # GetValue()
        self.error_msg = ct.create_string_buffer(70)
        self.meas_alias = {'T':'Temperature',
                          'P':'Absolute Pressure',
                          'RH':'Rel. Air Humidity',
                          'T_dew':'Dewpoint Temperature',
                          'T_wb':'Wet Bulb Temperature',
                          'H_atm':'Atmospheric Humidity',
                          'H_abs':'Absolute Humidity'}
        self.Open()
        self.info = self.GetSensorInfo()
        print 'devices.GMH_Sensor:\n',self.info


    def Open(self):
        """
        Use COM port number to open device
        Returns return-code of GMH_OpenCom() function
        """
        
        err_code = GMHLIB.GMH_OpenCom(self.addr).value
        print 'open() port', self.addr,'. Return code:',err_code
        if err_code in range(0,4):
            self.demo = False
        else:
            self.demo = True
        return err_code


    def Init(self):
        pass
     
       
    def Close(self):
        """
        Closes all / any GMH devices that are currently open.
        Returns return-code of GMH_CloseCom() function
        """
        if self.demo == True:
            return 1
        else:  
            err_code = GMHLIB.GMH_CloseCom()
        return err_code
 
   
    def Transmit(self,Addr,Func):
        """
        A wrapper for the general-purpose interrogation function GMH_Transmit().
        """
        if self.demo == True:
            return 1
        else:
            # Call GMH_Transmit():
            err_code = GMHLIB.GMH_Transmit(Addr,Func,ct.byref(self.Prio),ct.byref(self.flData),ct.byref(self.intData))
            
            # Translate return code into error message and store in self.error_msg
            self.error_code = ct.c_int16(err_code + self.lang_offset.value)
            GMHLIB.GMH_GetErrorMessageRet(self.error_code, ct.byref(self.error_msg))

            return self.error_code.value
 
   
    def GetSensorInfo(self):
        """
        Interrogates GMH sensor.
        Returns a dictionary keyed by measurement string.
        Values are tuples: (<address>, <measurement unit>),
        where <address> is an int and <measurement unit> is a string.
        
        The address corressponds with a unique measurement function within the device.
        """

        addresses = [] # Between 1 and 99
        measurements = [] # E.g. 'Temperature', 'Absolute Pressure', 'Rel. Air Humidity',...
        units = [] # E.g. 'deg C', 'hPascal', '%RH',...
        
        for Address in range(1,100):
            Addr = ct.c_short(Address)
            self.error_code = self.Transmit(Addr,self.MeasFn) # Writes result to self.intData
            if self.intData.value == 0:
                break # Bail-out if we run out of measurement functions
            addresses.append(Address)
    
            meas_code = ct.c_int16(self.intData.value + self.lang_offset.value)
            GMHLIB.GMH_GetMeasurement(meas_code, ct.byref(self.meas_str)) # Writes result to self.meas_str
            measurements.append(self.meas_str.value)
    
            self.Transmit(Addr,self.UnitFn) # Writes result to self.intData
                                     
            unit_code = ct.c_int16(self.intData.value + self.lang_offset.value)
            GMHLIB.GMH_GetUnit(unit_code, ct.byref(self.unit_str)) # Writes result to self.unit_str
            units.append(self.unit_str.value)
        
        return dict(zip(measurements,zip(addresses,units)))


    def Measure(self, meas):
        """
        Measure either temperature, pressure or %RH, based on parameter meas
        (see keys of self.meas_alias for options).
        Returns a tuple: (<Temperature/Pressure/RH as int>, <unit as string>)
        """
        if self.demo == True:
            return np.random.normal(20.5,0.1)
        else:
            assert self.demo == False,'GMH sensor in demo mode.'
            assert len(self.info.keys()) > 0,'No measurement functions available from this GMH sensor.'
            assert self.meas_alias.has_key(meas),'This measurement function is not available.'
            Address = self.info[self.meas_alias[meas]][0]
            Addr = ct.c_short(Address)
            self.Transmit(Addr,self.ValFn)
            return (self.flData.value, self.info[self.meas_alias[meas]][1])


    def Test(self, meas):
        """ Used to test that the device is functioning. """
        return self.Measure(meas)


'''
###############################################################################
'''        
        
class instrument(device):
    '''
    A class for associating instrument data with a VISA instance of that instrument
    '''
    def __init__(self, descr, demo=True): # Default to demo mode
        self.Descr = descr
        
        assert INSTR_DATA.has_key(self.Descr),'Unknown instrument - check instrument data is loaded from Excel Parameters sheet.'
        
        self.addr = INSTR_DATA[self.Descr]['addr']
        self.str_addr = INSTR_DATA[self.Descr]['str_addr']
        self.role = INSTR_DATA[self.Descr]['role']

        if INSTR_DATA[self.Descr].has_key('init_str'):
            self.InitStr = INSTR_DATA[self.Descr]['init_str'] # a tuple of strings
        else:
            self.InitStr = ('',) # a tuple of empty strings
        if INSTR_DATA[self.Descr].has_key('setfn_str'):
            self.SetFnStr = INSTR_DATA[self.Descr]['setfn_str']
        else:
            self.SetFnStr = '' # an empty string
        if INSTR_DATA[self.Descr].has_key('oper_str'):
            self.OperStr = INSTR_DATA[self.Descr]['oper_str']
        else:
            self.OperStr = '' # an empty string
        if INSTR_DATA[self.Descr].has_key('stby_str'):
            self.StbyStr = INSTR_DATA[self.Descr]['stby_str']
        else:
            self.StbyStr = ''
        if INSTR_DATA[self.Descr].has_key('chk_err_str'):
            self.ChkErrStr = INSTR_DATA[self.Descr]['chk_err_str']
        else:
            self.ChkErrStr = ('',)
        if INSTR_DATA[self.Descr].has_key('setV_str'):
		self.VStr = INSTR_DATA[self.Descr]['setV_str'] # a tuple of strings
        else:
            self.VStr = ''


    def Open(self):
        try:
            self.instr = RM.open_resource(self.str_addr)
            if '3458A' in self.Descr:
                self.instr.read_termination = '\r\n' # carriage return,line feed
                self.instr.write_termination = '\r\n' # carriage return,line feed
            self.instr.timeout = 2000 # default 2 s timeout
            INSTR_DATA[self.Descr]['demo'] = False # A real working instrument
            self.Demo = False # A real working instrument ONLY on Open() success
            print 'devices.instrument.Open():',self.Descr,'session handle=',self.instr.session
        except visa.VisaIOError:
            self.instr = None
            self.Demo = True # default to demo mode if can't open
            INSTR_DATA[self.Descr]['demo'] = True
            print 'devices.instrument.Open() failed:',self.Descr,'opened in demo mode'
        return self.instr


    def Close(self):
        # Close comms with instrument
        if self.Demo == True:
            print 'devices.instrument.Close():',self.Descr,'in demo mode - nothing to close'
        if self.instr is not None:
            print 'devices.instrument.Close():',self.Descr,'session handle=',self.instr.session
            self.instr.close()
        else:
            print 'devices.instrument.Close():',self.Descr,'is "None" or already closed'


    def Init(self):
        # Send initiation string
        if self.Demo == True:
            print 'devices.instrument.Init():',self.Descr,'in demo mode - no initiation necessary'
            return 1
        else:
            reply = 1
            for s in self.InitStr:
                if s != '': # instrument has an initiation string
                    try:
                        self.instr.write(s)
                    except visa.VisaIOError:
                        print'Failed to write "%s" to %s'%(s,self.Descr)
                        reply = -1
                        return reply
            print 'devices.instrument.Init():',self.Descr,'initiated with cmd:',s
        return reply


    def SetV(self,V):
        # set output voltage (SRC) or input range (DVM)
        if self.Demo == True:
		return 1
        elif 'SRC:' in self.Descr:
            # Set voltage-source to V
            s = str(V).join(self.VStr)
            print'devices.instrument.SetV():',self.Descr,'s=',s
            try:
                self.instr.write(s)
            except visa.VisaIOError:
                print'Failed to write "%s" to %s,via handle %s'%(s,self.Descr,self.instr.session)
                return -1
            return 1
        elif 'DVM:' in self.Descr:
            # Set DVM range to V
            s = str(V).join(self.VStr)
            self.instr.write(s)
            return 1
        else : # 'none' in self.Descr, (or something odd has happened)
            print 'Invalid function for instrument', self.Descr
            return -1


    def SetFn(self):
        # Set DVM function
        if self.Demo == True:
            return 1
        if 'DVM' in self.Descr:
            s = self.SetFnStr
            if s != '':
                self.instr.write(s)
            print'devices.instrument.SetFn():',self.Descr,'- OK.'
            return 1
        else:
            print'devices.instrument.SetFn(): Invalid function for',self.Descr
            return -1


    def Oper(self):
        # Enable O/P terminals
        # For V-source instruments only
        if self.Demo == True:
            return 1
        if 'SRC' in self.Descr:
            s = self.OperStr
            if s != '':
                try:
                    self.instr.write(s)
                except visa.VisaIOError:
                    print'Failed to write "%s" to %s'%(s,self.Descr)
                    return -1
            print'devices.instrument.Oper():',self.Descr,'output ENABLED.'
            return 1
        else:
            print'devices.instrument.Oper(): Invalid function for',self.Descr
            return -1


    def Stby(self):
        # Disable O/P terminals
        # For V-source instruments only
        if self.Demo == True:
            return 1
        if 'SRC' in self.Descr:
            s = self.StbyStr
            if s != '':
                self.instr.write(s) # was: query(s)
            print'devices.instrument.Stby():',self.Descr,'output DISABLED.'
            return 1
        else:
            print'devices.instrument.Stby(): Invalid function for',self.Descr
            return -1


    def CheckErr(self):
        # Get last error string and clear error queue
        # For V-source instruments only (F5520A)
        if self.Demo == True:
            return 1
        if 'F5520A' in self.Descr:
            s = self.ChkErrStr
            if s != ('',):
                reply = self.instr.query(s[0]) # read error message
                self.instr.write(s[1]) # clear registers
            return reply
        else:
            print'devices.instrument.CheckErr(): Invalid function for',self.Descr
            return -1


    def SendCmd(self,s):
        demo_reply = 'SendCmd(): DEMO resp. to '+s
        reply = 1
        if self.role == 'switchbox': # update icb
            pass # may need an event here...
        if self.Demo == True:
            print 'devices.instrument.SendCmd(): returning',demo_reply
            return demo_reply
        # Check if s contains '?' or 'X' or is an empty string
        # ... in which case a response is expected
        if any(x in s for x in'?X'):
            print'devices.instrument.SendCmd(): Query(%s) to %s'%(s,self.Descr)
            reply = self.instr.query(s)
            return reply
        elif s == '':
            reply = self.instr.read()
            print'devices.instrument.SendCmd(): Read()',reply,'from',self.Descr
            return reply
        else:
            print'devices.instrument.SendCmd(): Write(%s) to %s'%(s,self.Descr)
            self.instr.write(s)
            return reply


    def Read(self):
        reply = 0
        if self.Demo == True:
            return reply
        if 'DVM' in self.Descr:
            print'devices.instrument.Read(): from',self.Descr
            if '3458A' in self.Descr:
                reply = self.instr.read()
                return reply
            else:
                reply = self.instr.query('READ?')
                return reply
        else:
            print 'devices.instrument.Read(): Invalid function for',self.Descr
            return reply
        
        
    def Test(self,s):
        """ Used to test that the instrument is functioning. """
        return self.SendCmd(s)
#__________________________________________
