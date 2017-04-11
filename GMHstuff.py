# -*- coding: utf-8 -*-
"""
Created on Wed Jul 29 13:21:22 2015

@author: t.lawson
"""
# GMHstuff.py - required to access dll functions for GMH probes

import os
import ctypes as ct

# Change PATH to C:\GMH\GMHdll\ (or wherever you've put GMH3x32E.dll):
os.environ['GMHPATH'] = 'C:\Documents and Settings\\t.lawson\My Documents\Python Scripts\High_Res_Bridge\GMHdll'
gmhpath = os.environ['GMHPATH']
GMHLIB = ct.windll.LoadLibrary(os.path.join(gmhpath,'GMH3x32E'))


class GMH_Sensor():
    """
    A class to wrap around the low-level functions of GMH3x32E.dll. 
    For use with most Greisinger GMH devices.
    """
    def __init__(self,port,demo = True):
        self.demo = demo
        self.port = port # COM port of USB 3100N adapter cable
        self.c_Prio = ct.c_short()
        self.c_flData = ct.c_double()
        self.c_intData = ct.c_long()
        self.c_meas_str = ct.create_string_buffer(30)
        self.c_unit_str = ct.create_string_buffer(10)
        self.c_lang_offset = ct.c_int16(4096) # English language-offset
        self.c_MeasFn = ct.c_short(180) # GetMeasCode()
        self.c_UnitFn = ct.c_int16(178) # GetUnitCode()
        self.c_ValFn = ct.c_short(0) # GetValue()
        self.c_error_msg = ct.create_string_buffer(70)
        self.meas_alias = {'T':'Temperature',
                          'P':'Absolute Pressure',
                          'RH':'Rel. Air Humidity',
                          'T_dew':'Dewpoint Temperature',
                          'T_wb':'Wet Bulb Temperature',
                          'H_atm':'Atmospheric Humidity',
                          'H_abs':'Absolute Humidity'}
        self.error_msg = ''
        self.error_code = self.Open()
        self.info = self.GetSensorInfo()


    def Open(self):
        if self.demo == True:
            return 1
        else:  
            c_err_code = ct.c_int16(GMHLIB.GMH_OpenCom(self.port))
            self.c_error_code = ct.c_int16(c_err_code.value + self.c_lang_offset.value)
            GMHLIB.GMH_GetErrorMessageRet(self.c_error_code, ct.byref(self.c_error_msg))
            if c_err_code.value < 0:
                self.error_msg += self.c_error_msg.value
            
            print 'open() port', self.port,'...', self.error_msg
            return self.c_error_code.value
 
       
    def Close(self):
        if self.demo == True:
            return 1
        else:  
            GMHLIB.GMH_CloseCom()
        return 1
 
   
    def Transmit(self,Addr,Func):
        """
        A wrapper for the GMH general-purpose interrogation function GMH_Transmit().
        """
        if self.demo == True:
            return 1
        else:
            err_code = GMHLIB.GMH_Transmit(Addr,Func,ct.byref(self.c_Prio),ct.byref(self.c_flData),ct.byref(self.c_intData))
            
            self.c_error_code = ct.c_int16(err_code + self.c_lang_offset.value)
            GMHLIB.GMH_GetErrorMessageRet(self.c_error_code, ct.byref(self.c_error_msg))

            return self.c_error_code.value
 
   
    def GetSensorInfo(self):
        """
        Interrogates GMH sensor.
        Returns a dictionary keyed by measurement string.
        Values are tuples: (<address>, <measurement unit>),
        where <address> is an int and <measurement unit> is a string.
        """

        addresses = [] # Between 1 and 99
        measurements = [] # E.g. 'Temperature', 'Absolute Pressure', 'Rel. Air Humidity',...
        units = [] # E.g. 'deg C', 'hPascal', '%RH',...
        
        for Address in range(1,100):
            c_Addr = ct.c_short(Address)
            self.error_code = self.Transmit(c_Addr,self.c_MeasFn) # Writes result to self.c_intData
            if self.c_intData.value == 0:
                break # Bail-out if we run out of measurement functions
            addresses.append(Address)
    
            c_meas_code = ct.c_int16(self.c_intData.value + self.c_lang_offset.value)
            GMHLIB.GMH_GetMeasurement(c_meas_code, ct.byref(self.c_meas_str)) # Writes result to self.c_meas_str
            measurements.append(self.c_meas_str.value)
    
            self.Transmit(c_Addr,self.c_UnitFn) # Writes result to self.c_intData
                                     
            c_unit_code = ct.c_int16(self.c_intData.value + self.c_lang_offset.value)
            GMHLIB.GMH_GetUnit(c_unit_code, ct.byref(self.c_unit_str)) # Writes result to self.c_unit_str
            units.append(self.c_unit_str.value)

        return dict(zip(measurements,zip(addresses,units)))


    def Measure(self, meas):
        """
        Measure parameter meas (temperature, pressure or humidity).
        Returns a tuple: (<Temperature/Pressure/RH as int>, <unit as string>)
        meas is one of: 'T', 'P', 'RH', 'T_dew', 't_wb', 'H_atm' or 'H_abs'.
        """
        if len(self.info)==0:
            print 'Measure(): No measurements available! - Check sensor is connected and ON.'
            return (0,'')
        else:
            Address = self.info[self.meas_alias[meas]][0]
            Addr = ct.c_short(Address)
            self.Transmit(Addr,self.c_ValFn)
            return (self.c_flData.value, self.info[self.meas_alias[meas]][1])