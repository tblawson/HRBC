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
import GMHstuff as Gmh  # GMH probe coms are handled by low-level routines in GMH3x32E.dll.

INSTR_DATA = {}  # Dict of instrument parameter dicts, keyed by description
DESCR = []
sublist = []
ROLES_WIDGETS = {}  # Dictionary of GUI widgets keyed by role
ROLES_INSTR = {}  # Dictionary of GMH_sensor or Instr objects keyed by role

"""
VISA-specific stuff:
Only ONE VISA resource manager is required at any time -
All comunications for all GPIB and RS232 instruments (except GMH)
are handled by RM.
"""
RM = visa.ResourceManager()

# Switchbox
SWITCH_CONFIGS = {'V1': 'A', 'Vd1': 'C', 'Vd2': 'D', 'V2': 'B'}

T_Sensors = ('none', 'Pt', 'SR104t', 'thermistor')


class GMHDevice(Gmh.GMHSensor):
    """
    A derived class of GMHstuff.GMHSensor with additional functionality.
    On creation, an instance needs a description string 'descr'.
    """
    def __init__(self, descr):
        self.descr = descr
        self.addr = int(INSTR_DATA[self.descr]['addr'])
        super().__init__(self.addr)
        self.demo = True

    def test(self, meas):
        """
        Test that the device is functioning.

        :argument meas (str) - an alias for the measurement type:
            'T', 'P', 'RH', 'T_dew', 't_wb', 'H_atm' or 'H_abs'.

        :returns measurement tuple: (<value>, <unit string>)
        """
        print('\ndevices.GMHDvice.Test()...')
        result = self.measure(meas)
        return result

    def init(self):
        """
        A dummy method - mirrors the init method of Instrument class,
        but not needed for GMH sensors.
        """
        pass


class Device(object):
    """
    A generic external device or instrument
    """
    def __init__(self, demo=True):
        self.demo = demo

    def open(self):
        pass

    def close(self):
        pass


'''
###############################################################################
'''


class Instrument(Device):
    """
    A class for associating instrument data with a VISA instance of
    that instrument.
    """
    def __init__(self, descr, demo=True):  # Default to demo mode
        self.instr = None
        self.descr = descr
        self.demo = demo
        self.is_open = 0
        self.is_operational = 0

        msg = 'Unknown instrument - check instrument data is loaded from \
Excel Parameters sheet.'
        assert self.descr in INSTR_DATA, msg

        self.addr = INSTR_DATA[self.descr]['addr']
        self.str_addr = INSTR_DATA[self.descr]['str_addr']
        self.role = INSTR_DATA[self.descr]['role']

        if 'init_str' in INSTR_DATA[self.descr]:
            self.InitStr = INSTR_DATA[self.descr]['init_str']  # tuple of str
        else:
            self.InitStr = ('',)  # a tuple of empty strings

        if 'setfn_str' in INSTR_DATA[self.descr]:
            self.SetFnStr = INSTR_DATA[self.descr]['setfn_str']
        else:
            self.SetFnStr = ''  # an empty string

        if 'oper_str' in INSTR_DATA[self.descr]:
            self.OperStr = INSTR_DATA[self.descr]['oper_str']
        else:
            self.OperStr = ''  # an empty string

        if 'stby_str' in INSTR_DATA[self.descr]:
            self.StbyStr = INSTR_DATA[self.descr]['stby_str']
        else:
            self.StbyStr = ''

        if 'chk_err_str' in INSTR_DATA[self.descr]:
            self.ChkErrStr = INSTR_DATA[self.descr]['chk_err_str']
        else:
            self.ChkErrStr = ('',)

        if 'setV_str' in INSTR_DATA[self.descr]:
            self.V_str = INSTR_DATA[self.descr]['setV_str']  # tuple of str
        else:
            self.V_str = ''

        if 'range_str' in INSTR_DATA[self.descr]:  # Unique to Transmille calibrators.
            self.RangeStr = INSTR_DATA[self.descr]['range_str']
        else:
            self.RangeStr = ''

        if 'cmd_sep' in INSTR_DATA[self.descr]:  # Unique to Transmille calibrators.
            self.CmdSep = INSTR_DATA[self.descr]['cmd_sep']
        else:
            self.CmdSep = ''

        transmille_dcv_ranges = {}

    def open(self):
        m = 'devices.instrument.Open():'
        try:
            self.instr = RM.open_resource(self.str_addr)
            self.is_open = 1
            if '3458A' in self.descr:
                self.instr.read_termination = '\r\n'  # carriage ret, l-feed
                self.instr.write_termination = '\r\n'  # carriage ret, l-feed
            self.instr.timeout = 2000  # default 2 s timeout
            INSTR_DATA[self.descr]['demo'] = False  # A real working instr
            self.demo = False  # A real working instr ONLY on Open() success
            print(m, self.descr, 'session handle={}.'.format(self.instr.session))
        except visa.VisaIOError:
            self.instr = None
            self.demo = True  # default to demo mode if can't open
            INSTR_DATA[self.descr]['demo'] = True
            print(m, self.descr, 'opened in demo mode.')
        return self.instr

    def close(self):
        # Close comms with instrument
        m = 'devices.instrument.Close():'
        if self.demo is True:
            print(m, self.descr, 'in demo mode - nothing to close.')
        if self.instr is not None:
            print(m, self.descr, 'session handle=', self.instr.session)
            self.instr.close()
        else:
            print(m, self.descr, 'is "None" or already closed.')
        self.is_open = 0

    def init(self):
        # Send initiation string
        m = 'devices.instrument.Init():'
        if self.demo is True:
            print(m, self.descr, 'in demo mode - no initiation necessary.')
            return 1
        else:
            reply = 1
            for s in self.InitStr:
                if s != '':  # instrument has an initiation string.
                    try:
                        self.instr.write(s)
                    except visa.VisaIOError:
                        print('Failed to write "{}" to {}'.format(s, self.descr))
                        reply = -1
                        return reply
            print(m, self.descr, 'initiated with cmd: {}'.format(s))
        return reply

    def set_V(self, V):
        # set output voltage (SRC) or input range (DVM)
        if self.demo is True:
            return 1
        elif 'SRC:' in self.descr:
            # Set voltage-source to V
            if 'T3310A' in self.descr:  # Transmille calibrator
                s = self.Transmille_V_str(V)
            else:
                s = str(V).join(self.V_str)
            print('devices.instrument.SetV():', self.descr, 's=', s)
            try:
                self.instr.write(s)
            except visa.VisaIOError:
                m = 'Failed to write "{}" to {}, via handle {}.'
                print(m.format(s, self.descr, self.instr.session))
                return -1
            return 1
        elif 'DVM:' in self.descr:
            # Set DVM range to V
            s = str(V).join(self.V_str)
            self.instr.write(s)
            return 1
        else:  # 'none' in self.Descr, (or something odd has happened)
            print('Invalid function for instrument {}.'.format(self.descr))
            return -1

    def Transmille_V_str(self, V):
        ranges = [0.2, 2, 20, 200, 1000]
        v_str = str(V)
        r_str = ''
        for i, r in enumerate(ranges):
            if i == 0:  # Lowest range - convert to mV
                v_str = str(1000*V)
            else:
                v_str = str(V)
            if V > r:
                continue
            else:
                r_str = 'R'+str(i+1)
                break
        cmd_seq = [r_str, 'O'+v_str, 'S0']
        return self.CmdSep.join(cmd_seq)

    def set_fn(self):
        # Set DVM function
        if self.demo is True:
            return 1
        if 'DVM' in self.descr:
            s = self.SetFnStr
            if s != '':
                self.instr.write(s)
            print('devices.instrument.SetFn():', self.descr, '- OK.')
            return 1
        else:
            print('devices.instrument.SetFn(): Invalid function for', self.descr)
            return -1

    def oper(self):
        # Enable O/P terminals.
        # For V-source instruments only!
        if self.demo is True:
            return 1
        if 'SRC' in self.descr:
            s = self.OperStr
            if s != '':
                try:
                    self.instr.write(s)
                except visa.VisaIOError:
                    print('Failed to write "{}" to {}.'.format(s, self.descr))
                    return -1
            print('devices.instrument.Oper():', self.descr, 'output ENABLED.')
            return 1
        else:
            print('devices.instrument.Oper(): Invalid function for {}.'.format(self.descr))
            return -1

    def stby(self):
        # Disable O/P terminals.
        # For V-source instruments only!
        if self.demo is True:
            return 1
        if 'SRC' in self.descr:
            s = self.StbyStr
            if s != '':
                try:
                    self.instr.write(s)
                except visa.VisaIOError:
                    print('Failed to write "{}" to {}.'.format(s, self.descr))
                    return -1
            print('devices.instrument.Stby():', self.descr, 'output DISABLED.')
            return 1
        else:
            print('devices.instrument.Stby(): Invalid function for {}.'.format(self.descr))
            return -1

    def check_err(self):
        # Get last error string and clear error queue
        # For V-source instruments only (F5520A)
        if self.demo is True:
            return 1
        if 'F5520A' in self.descr:
            s = self.ChkErrStr
            if s != ('',):
                reply = self.instr.query(s[0])  # read error message
                self.instr.write(s[1])  # clear registers
            return reply
        else:
            m = 'devices.instrument.CheckErr(): Invalid function for {}.'
            print(m.format(self.descr))
            return -1

    def send_cmd(self, s):
        m = 'devices.instrument.SendCmd(): '
        demo_reply = 'SendCmd(): DEMO resp. to {}'.format(s)
        reply = ''
        if self.role == 'switchbox':  # update icb
            pass  # may need an event here...
        if self.demo is True:
            print(m, 'returning {}.'.format(demo_reply))
            return demo_reply
        # Check if s contains '?' or 'X' or is an empty string
        # ... in which case a response is expected.
        if any(x in s for x in'?X'):
            print(m, 'Query({}) to {}'.format(s, self.descr))
            reply = self.instr.query(s)
            return reply
        elif s == '':
            reply = self.instr.read()
            print(m, 'Read("{}") from {}.'.format(reply, self.descr))
            return reply
        else:
            print(m, 'Write({}) to {}'.format(s, self.descr))
            self.instr.write(s)
            return reply

    def read(self):
        reply = 0
        if self.demo is True:
            return reply
        if 'DVM' in self.descr:
            print('devices.instrument.Read(): from {}.'.format(self.descr))
            if '3458A' in self.descr:
                reply = self.instr.read()
                return reply
            else:
                reply = self.instr.query('READ?')
                return reply
        else:
            print('devices.instrument.Read(): Invalid function for {}.'.format(self.descr))
            return reply

    def test(self, s):
        """ Used to test that the instrument is functioning. """
        return self.send_cmd(s)
# __________________________________________
