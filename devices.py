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
LANG_OFFSET = 4096            
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
        
        self.addr = int(INSTR_DATA[self.Descr]['addr']) # COM port-number assigned to USB 3100N adapter cable
        self.str_addr = INSTR_DATA[self.Descr]['str_addr']
        self.role = INSTR_DATA[self.Descr]['role']        
        
        self.Prio = ct.c_short()
        self.flData = ct.c_double() # Don't change this type!! It's the exactly right one!
        self.intData = ct.c_long()
        self.meas_str = ct.create_string_buffer(30)
        self.unit_str = ct.create_string_buffer(10)
        self.lang_offset = ct.c_int16(LANG_OFFSET) # English language-offset
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
        self.info = {}


    def Open(self):
        """
        Use COM port number to open device
        Returns 1 if successful, 0 if not
        """
        print'\ndevices.GMH_Sensor.Open(): Trying port',repr(self.addr)
        self.error_code = ct.c_int16(GMHLIB.GMH_OpenCom(self.addr))
        self.GetErrMsg() # Get self.error_msg
        
        if self.error_code.value in range(0,4) or self.error_code.value == -2:
            print 'devices.GMH_Sensor.Open(): ',self.str_addr,'is open.'
            
            # We're not there yet - test device responsiveness
            self.Transmit(1,self.ValFn)
            self.GetErrMsg()
            if self.error_code.value in range(0,4): # Sensor responds...
                if len(self.info) == 0: # No device info yet
                    print 'devices.GMH_Sensor.Open(): Getting sensor info...'
                    self.GetSensorInfo()
                    self.demo = False # If we've got this far we're probably OK
                    return True
                else: # Already have device measurement info
                    print'devices.GMH_Sensor.Open(): Instrument ready - demo=False.'
                    self.demo = False # If we've got this far we're probably OK
                    return True
            else: # No response
                print 'devices.GMH_Sensor.Open():',self.error_msg.value
                self.Close()
                self.demo = True
                return False
   
        else: # Com open failed
            print'devices.GMH_Sensor.Open() FAILED:',self.Descr
            self.Close()
            self.demo = True
            return False


    def Init(self):
        print'devices.GMH_Sensor.Init():',self.Descr,'initiated (nothing happens here).'
        pass
     
       
    def Close(self):
        """
        Closes all / any GMH devices that are currently open.
        """
        print'\ndevices.GMH_Sensor.Close(): Setting demo=True and Closing all GMH sensors ...'
        self.demo = True
        self.error_code = ct.c_int16(GMHLIB.GMH_CloseCom())
#        self.GetErrMsg()
        print 'devices.GMH_Sensor.Close(): CloseCom err_msg:',self.error_msg.value
        return 1

  
    def Transmit(self,Addr,Func):
        """
        A wrapper for the general-purpose interrogation function GMH_Transmit().
        """
        self.error_code = ct.c_int16(GMHLIB.GMH_Transmit(Addr,Func,ct.byref(self.Prio),ct.byref(self.flData),ct.byref(self.intData)))
        self.GetErrMsg()
        if self.error_code.value < 0:
            print'\ndevices.GMH_Sensor.Transmit():FAIL'
            return False
        else:
            print'\ndevices.GMH_Sensor.Transmit():PASS'
            return True
 

    def GetErrMsg(self):
        """
        Translate return code into error message and store in self.error_msg.
        """
        error_code_ENG = ct.c_int16(self.error_code.value + self.lang_offset.value)
        GMHLIB.GMH_GetErrorMessageRet(error_code_ENG, ct.byref(self.error_msg))
        if self.error_code.value in range(0,4): # Correct message_0
            self.error_msg.value = 'Success'
        return 1
 
  
    def GetSensorInfo(self):
        """
        Interrogates GMH sensor.
        Returns a dictionary keyed by measurement string.
        Values are tuples: (<address>, <measurement unit>),
        where <address> is an int and <measurement unit> is a string.
        
        The address corressponds with a unique measurement function within the device.
        It's assumed the measurement functions are at consecutive addresses starting at 1.
        """
        addresses = [] # Between 1 and 99
        measurements = [] # E.g. 'Temperature', 'Absolute Pressure', 'Rel. Air Humidity',...
        units = [] # E.g. 'deg C', 'hPascal', '%RH',...
        self.info.clear()
        
        for Address in range(1,100):
            Addr = ct.c_short(Address)
            if self.Transmit(Addr,self.MeasFn): # Writes result to self.intData
                # Transmit() was successful
                addresses.append(Address)
                
                meas_code = ct.c_int16(self.intData.value + self.lang_offset.value)
                GMHLIB.GMH_GetMeasurement(meas_code, ct.byref(self.meas_str)) # Writes result to self.meas_str
                measurements.append(self.meas_str.value)
                
                self.Transmit(Addr,self.UnitFn) # Writes result to self.intData
                
                unit_code = ct.c_int16(self.intData.value + self.lang_offset.value)
                GMHLIB.GMH_GetUnit(unit_code, ct.byref(self.unit_str)) # Writes result to self.unit_str
                units.append(self.unit_str.value)
                
                print'Found',self.meas_str.value,'(',self.unit_str.value,')','at address',Address
            else:
                print'devices.GMH_Sensor.GetSensorInfo(): Exhausted addresses at',Address
                if Address > 1: # Don't let the last address tried screw it up.
                    self.error_code.value = 0
                    self.demo = False
                else:
                    self.demo = True
                break # Assumes all functions are in a contiguous address range starting at 1

        self.info = dict(zip(measurements,zip(addresses,units)))
        print 'devices.GMH_Sensor.GetSensorInfo():\n',self.info,'demo =',self.demo
        return len(self.info)


    def Measure(self, meas):
        """
        Measure either temperature, pressure or humidity, based on parameter meas
        Returns a float.
        meas is one of: 'T', 'P', 'RH', 'T_dew', 't_wb', 'H_atm' or 'H_abs'.\
        
        NOTE that because GMH_CloseCom() acts on ALL open GMH devices it makes
        sense to only have a device open when communicating with it and to
        immediately close it afterwards. This way the default state is closed
        and the open state is treated as a special case. Hence an Open()-Close()
        'bracket' surrounds the Measure() function.
        """
        
        self.flData.value = 0
        if self.Open(): # port and device open success
            assert self.demo == False,'Illegal access to demo device!'
            Address = self.info[self.meas_alias[meas]][0]
            Addr = ct.c_short(Address)
            self.Transmit(Addr,self.ValFn)
            self.Close()
            
            print'devices.Measure():',self.meas_alias[meas],'=',self.flData.value
            return self.flData.value
        else:
            assert self.demo == True,'Illegal denial to demo device!'
            print'devices.GMH_Sensor.Measure(): Returning demo-value.'
            demo_rtn = {'T':(20.5,0.2),'P':(1013,5),'RH':(50,10)}
            return np.random.normal(*demo_rtn[meas])


    def Test(self, meas):
        """ Used to test that the device is functioning. """
        print'\ndevices.GMH_Sensor.Test()...'
        result = self.Measure(meas)
        return result


'''
###############################################################################
'''        
        
class instrument(device):
    '''
    A class for associating instrument data with a VISA instance of that instrument
    '''
    def __init__(self, descr, demo=True): # Default to demo mode
        self.Descr = descr
        self.demo = demo
        self.is_open = 0
        self.is_operational = 0
        
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
            self.is_open = 1
            if '3458A' in self.Descr:
                self.instr.read_termination = '\r\n' # carriage return,line feed
                self.instr.write_termination = '\r\n' # carriage return,line feed
            self.instr.timeout = 2000 # default 2 s timeout
            INSTR_DATA[self.Descr]['demo'] = False # A real working instrument
            self.demo = False # A real working instrument ONLY on Open() success
            print 'devices.instrument.Open():',self.Descr,'session handle=',self.instr.session
        except visa.VisaIOError:
            self.instr = None
            self.demo = True # default to demo mode if can't open
            INSTR_DATA[self.Descr]['demo'] = True
            print 'devices.instrument.Open() failed:',self.Descr,'opened in demo mode'
        return self.instr


    def Close(self):
        # Close comms with instrument
        if self.demo == True:
            print 'devices.instrument.Close():',self.Descr,'in demo mode - nothing to close'
        if self.instr is not None:
            print 'devices.instrument.Close():',self.Descr,'session handle=',self.instr.session
            self.instr.close()
        else:
            print 'devices.instrument.Close():',self.Descr,'is "None" or already closed'
        self.is_open = 0


    def Init(self):
        # Send initiation string
        if self.demo == True:
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
        if self.demo == True:
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
        if self.demo == True:
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
        if self.demo == True:
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
        if self.demo == True:
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
        if self.demo == True:
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
        if self.demo == True:
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
        if self.demo == True:
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
