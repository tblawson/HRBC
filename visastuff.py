# -*- coding: utf-8 -*-
""" visastuff.py
Created on Tue Jun 16 13:51:52 2015

DEVELOPMENT VERSION

@author: t.lawson
"""

"""
This is the place where all information about the instruments available for use
in the procedure should be recorded. Each dictionary key (description) accesses a
particular instrument and the corresponding dictionary value is a lower-level dictionary
of key:value pairs, eg: GPIB address, initiation command string
(which may or may not result in output),Voltage-setting command, data read
command-string (which should cause readings to appear at the instrument's
output buffer), etc.

There is also an instrumant class that includes an instance method SendCmd() that
allows any command string to be sent to the instrument and any response to be recorded.
"""

import visa


# Only ONE resource manager is required at any time -
# All comunications for all GPIB and RS232 instruments
# are handled by RM.
RM = visa.ResourceManager()

# Switchbox
SWITCH_CONFIGS = {'V1':'A','Vd1':'C','Vd2':'D','V2':'B'}

# GMH probe communications are handled by low-level routines in
# GMHdll.dll
GMH_DESCR = ('GMH, s/n627',
             'GMH, s/n628')

T_Sensors = ('none','Pt','SR104t','thermistor')



# All the data is collected in a 'dictionary of dictionaries' - INSTR_DATA
# Each piece of information is accessed as:
# INSTR_DATA[<instrument description>][<parameter>].
# E.g. the 'set function' string for the HP3458(s/n518) is:
# INSTR_DATA['DVM: HP3458A, s/n518']['setfn_str'],
# which retrieves the string 'FUNC OHMF;OCOMP ON'.
# Note that some of the strings are LISTS of strings (i.e. multiple commands)
#INSTR_DATA = dict(zip(DESCR,sublist))

INSTR_DATA = {}
DESCR = []
sublist = []

#-------------------------------------------
# Roles dictionaries - for disseminating role-instrument info

ROLES_WIDGETS = {} # Dictionary of GUI widgets keyed by role
ROLES_INSTR = {} # Dictionary of visa instrument objects keyed by role

#--------------------------------------
class instrument():
    '''
    A class for associating instrument data with a VISA instance of that instrument
    '''
    def __init__(self, descr, demo=True): # Default to demo mode
        self.Descr = descr
        
        INSTR_DATA[descr]['demo'] = demo # update demo state, if neccesary
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
        if INSTR_DATA[self.Descr].has_key('hw_addr'):
		self.hw_addr = INSTR_DATA[self.Descr]['hw_addr'] 
        else:
		self.hw_addr = 0

    def Open(self):
        try:
            self.instr = RM.open_resource(self.str_addr)
            if '3458A' in self.Descr:
                self.instr.read_termination = '\r\n' # carriage return,line feed
                self.instr.write_termination = '\r\n' # carriage return,line feed
            self.instr.timeout = 2000 # default 2 s timeout
            INSTR_DATA[self.Descr]['demo'] = False # A real working instrument
            self.Demo = False # A real working instrument ONLY on Open() success
            print 'visastuff.instrument.Open():',self.Descr,'session handle=',self.instr.session
                
        except visa.VisaIOError:
            self.instr = None
            self.Demo = True # default to demo mode if can't open
            INSTR_DATA[self.Descr]['demo'] = True
            print 'visastuff.instrument.Open() failed:',self.Descr,'opened in demo mode'
        return self.instr

    def Close(self):
        # Close comms with instrument
        if self.Demo == True:
            print 'visastuff.instrument.Close():',self.Descr,'in demo mode - nothing to close'
        if self.instr is not None:
            print 'visastuff.instrument.Close():',self.Descr,'session handle=',self.instr.session
            self.instr.close()
        else:
            print 'visastuff.instrument.Close():',self.Descr,'is "None" or already closed'

    def Init(self):
        # Send initiation string
        if self.Demo == True:
            print 'visastuff.instrument.Init():',self.Descr,'in demo mode - no initiation necessary'
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
            print 'visastuff.instrument.Init():',self.Descr,'initiated with cmd:',s
        return reply

    def SetV(self,V):
        # set output voltage (SRC) or input range (DVM)
        if self.Demo == True:
		return 1
        elif 'SRC:' in self.Descr:
            # Set voltage-source to V
            s = str(V).join(self.VStr)
            print'visastuff.instrument.SetV():',self.Descr,'s=',s
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
            print'visastuff.instrument.SetFn():',self.Descr,'- OK.'
            return 1
        else:
            print'visastuff.instrument.SetFn(): Invalid function for',self.Descr
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
            print'visastuff.instrument.Oper():',self.Descr,'output ENABLED.'
            return 1
        else:
            print'visastuff.instrument.Oper(): Invalid function for',self.Descr
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
            print'visastuff.instrument.Stby():',self.Descr,'output DISABLED.'
            return 1
        else:
            print'visastuff.instrument.Stby(): Invalid function for',self.Descr
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
            print'visastuff.instrument.CheckErr(): Invalid function for',self.Descr
            return -1

    def SendCmd(self,s):
        demo_reply = 'SendCmd(): DEMO resp. to '+s
        reply = 1
        if self.role == 'switchbox': # update icb
            pass # may need an event here...
        if self.Demo == True:
            print 'visastuff.instrument.SendCmd(): returning',demo_reply
            return demo_reply
        # Check if s contains '?' or 'X' or is an empty string
        # ... in which case a response is expected
        if any(x in s for x in'?X'):
            print'visastuff.instrument.SendCmd(): Query(%s) to %s'%(s,self.Descr)
            reply = self.instr.query(s)
            return reply
        elif s == '':
            reply = self.instr.read()
            print'visastuff.instrument.SendCmd(): Read()',reply,'from',self.Descr
            return reply
        else:
            print'visastuff.instrument.SendCmd(): Write(%s) to %s'%(s,self.Descr)
            self.instr.write(s)
            return reply

    def Read(self):
        reply = 0
        if self.Demo == True:
            return reply
        if 'DVM' in self.Descr:
            print'visastuff.instrument.Read(): from',self.Descr
            if '3458A' in self.Descr:
                reply = self.instr.read()
                return reply
            else:
                reply = self.instr.query('READ?')
                return reply
        else:
            print 'visastuff.instrument.Read(): Invalid function for',self.Descr
            return reply
#__________________________________________
