# -*- coding: utf-8 -*-
"""
Created on Wed Jun 24 09:36:42 2015

DEVELOPMENT VERSION

@author: t.lawson
"""

"""
acquisition.py:
Thread class that executes processing.
Contains definitions for usual __init__() and run() methods
 AND an abort() method. The Run() method forms the core of the
 procedure - any changes to the way the measurements are taken
 should be made here, and within included subroutines.
"""
import wx
from threading import Thread
import datetime as dt
import time
#import os.path
#os.environ['XLPATH'] = 'C:\Documents and Settings\\t.lawson\My Documents\Python Scripts\High_Res_Bridge'

import numpy as np

from openpyxl import load_workbook
from openpyxl.styles import Font,Border,Side

import HighRes_events as evts
import visastuff
import GMHstuff as GMH

class AqnThread(Thread):
    """Acquisition Thread Class."""
    def __init__(self, parent):
        # This runs when an instance of the class is created
        Thread.__init__(self)
        self.RunPage = parent
        self.SetupPage = self.RunPage.GetParent().GetPage(0)
        self.PlotPage = self.RunPage.GetParent().GetPage(2)
        self.TopLevel = self.RunPage.GetTopLevelParent()
        self.Comment = self.RunPage.Comment.GetValue()
        self._want_abort = 0

        print'Role -> Instrument:'
        print'------------------------------'
        # Print all GPIB instrument objects
        for r in visastuff.ROLES_WIDGETS.keys():
            d = visastuff.ROLES_WIDGETS[r]['icb'].GetValue()
            # For 'switchbox' role, d is actually the setting (V1, Vd1,...) not the instrument description.
            
            print'%s -> %s'%(visastuff.INSTR_DATA[d]['role'],d)
            if r != visastuff.INSTR_DATA[d]['role']:
                visastuff.INSTR_DATA[d]['role'] = r
                print'Role data corrected to:',r,'->',d

        # Get filename of Excel file
        self.xlfilename = self.SetupPage.XLFile.GetValue()

        # open existing workbook
        self.wb_io = load_workbook(self.xlfilename,data_only=True) # 'data_only=True' ensures we read cell value, NOT formula
        self.ws = self.wb_io.get_sheet_by_name('Data')

        # read start/stop row numbers from Excel file
        self.start_row = self.ws['B1'].value
        self.stop_row = self.ws['B2'].value
        strt_ev = evts.StartRowEvent(row = self.start_row)
        wx.PostEvent(self.RunPage,strt_ev)
        stp_ev = evts.StopRowEvent(row = self.stop_row)
        wx.PostEvent(self.RunPage,stp_ev)

        self.settle_time = self.RunPage.SettleDel.GetValue()

        # Local record of GMH ports and addresses
        self.GMH1Demo_status = visastuff.INSTR_DATA[self.SetupPage.GMH1Probes.GetValue()]['demo']
        self.GMH2Demo_status = visastuff.INSTR_DATA[self.SetupPage.GMH2Probes.GetValue()]['demo']
        self.GMH1Port = self.SetupPage.GMH1Ports.GetValue().replace('COM','')
        self.GMH2Port = self.SetupPage.GMH2Ports.GetValue().replace('COM','')
        self.GMH1Addr = visastuff.INSTR_DATA[self.SetupPage.GMH1Probes.GetValue()]['hw_addr']
        self.GMH2Addr = visastuff.INSTR_DATA[self.SetupPage.GMH2Probes.GetValue()]['hw_addr']
        
        print 'GMH1 on port %s, hw_addr %d' % (self.GMH1Port, self.GMH1Addr)
        print 'GMH2 on port %s, hw_addr %d' % (self.GMH2Port, self.GMH2Addr)
        
        self.start() # Starts the thread running on creation

    def run(self):
        # Run Worker Thread. This is where all the important stuff goes, in a repeated cycle

        # Clear plots
        clr_plot_ev = evts.ClearPlotEvent()
        wx.PostEvent(self.PlotPage, clr_plot_ev)

        # Column headings
        Head_row = self.start_row-2 # Main headings
        sub_row = self.start_row-1 # Sub-headings
        # Write unique id for this run - used to pair measurement data with RLink data
        self.ws['A'+str(sub_row)] = 'Run Id:'
        self.ws['B'+str(sub_row)].font = Font(b=True)
        self.ws['B'+str(sub_row)] = str(self.RunPage.run_id)
        self.ws['A'+str(Head_row)] = 'V1_set'
        self.ws['B'+str(Head_row)] = 'V2_set'
        self.ws['C'+str(Head_row)] = 'n'
        self.ws['D'+str(Head_row)] = 'Start/xl del.'
        self.ws['E'+str(Head_row)] = 'AZ1 del.'
        self.ws['F'+str(Head_row)] = 'Range del.'
        self.ws['G'+str(Head_row)] = 'V2'
        self.ws['G'+str(sub_row)] = 't'
        self.ws['H'+str(sub_row)] = 'V'
        self.ws['I'+str(sub_row)] = 'sd(V)'
        # miss columns j,k,l
        self.ws['M'+str(Head_row)] = 'Vd1'
        self.ws['M'+str(sub_row)] = 't'
        self.ws['N'+str(sub_row)] = 'V'
        self.ws['O'+str(sub_row)] = 'sd(V)'
        self.ws['P'+str(Head_row)] = 'V1'
        self.ws['P'+str(sub_row)] = 't'
        self.ws['Q'+str(sub_row)] = 'V'
        self.ws['R'+str(sub_row)] = 'sd(V)'
        self.ws['S'+str(Head_row)] = 'dvm_T1'
        self.ws['T'+str(Head_row)] = 'dvm_T2'
        self.ws['U'+str(Head_row)] = 'GMH_T1'
        self.ws['V'+str(Head_row)] = 'GMH_T2'
        self.ws['W'+str(Head_row)] = 'Ambient Conditions'
        self.ws['W'+str(sub_row)] = 'T'
        self.ws['X'+str(sub_row)] = 'P(mbar)'
        self.ws['Y'+str(sub_row)] = '%RH'
        self.ws['Z'+str(Head_row)] = 'Comment'
        self.ws['AC'+str(Head_row)] = 'Role'
        self.ws['AD'+str(Head_row)] = 'Instrument descr.'


        stat_ev = evts.StatusEvent(msg='AqnThread.run():',field = 0)
        wx.PostEvent(self.TopLevel, stat_ev)
        stat_ev = evts.StatusEvent(msg='Waiting to settle...',field = 1)
        wx.PostEvent(self.TopLevel, stat_ev)

        time.sleep(self.settle_time)

        # Initialise all instruments
        self.initialise()

        stat_ev = evts.StatusEvent(msg='', field='b') # write to both status fields
        wx.PostEvent(self.TopLevel, stat_ev)

        stat_ev = evts.StatusEvent(msg='Post-initialise delay...', field=1)
        wx.PostEvent(self.TopLevel, stat_ev)
        time.sleep(3) # 3

        # Get some initial temperatures...
        self.TR1 = self.ReadGMH(int(self.GMH1Port),int(self.GMH1Addr),self.GMH1Demo_status)
        self.TR2 = self.ReadGMH(int(self.GMH2Port),int(self.GMH2Addr),self.GMH2Demo_status)
        self.ws['U'+str(self.start_row-1)] = self.TR1
        self.ws['V'+str(self.start_row-1)] = self.TR2

        # Record ALL POSSIBLE roles and corresponding instrument descriptions in XL sheet
        role_row = self.start_row
        bord_tl = Border(top = Side(style='thin'), left = Side(style='thin'))
        bord_tr = Border(top = Side(style='thin'), right = Side(style='thin'))
        bord_l = Border(left = Side(style='thin'))
        bord_r = Border(right = Side(style='thin'))
        bord_bl = Border(bottom = Side(style='thin'), left = Side(style='thin'))
        bord_br = Border(bottom = Side(style='thin'), right = Side(style='thin'))
        for r in visastuff.ROLES_WIDGETS.keys():
            if role_row == self.start_row: # 1st row
                self.ws['AC'+str(role_row)].border = bord_tl
                self.ws['AD'+str(role_row)].border = bord_tr
            elif role_row == self.start_row + 8: # last row
                self.ws['AC'+str(role_row)].border = bord_bl
                self.ws['AD'+str(role_row)].border = bord_br
            else: # in-between rows
                self.ws['AC'+str(role_row)].border = bord_l
                self.ws['AD'+str(role_row)].border = bord_r
            self.ws['AC'+str(role_row)] = r
            d = visastuff.ROLES_WIDGETS[r]['icb'].GetValue() # descr
            self.ws['AD'+str(role_row)] = d
            role_row += 1

        row = self.start_row
        pbar = 1

        # loop over xl rows..
        while row <= self.stop_row:
            if self._want_abort:
                self.AbortRun()
                return
            #self.role_list['DVM12'].SendCmd('DCV,100') # dvmV1V2:'DCV100'-REDUNDANT?

            if self._want_abort:
                self.AbortRun()
                return
            stat_ev = evts.StatusEvent(msg='AqnThread.run():', field=0)
            wx.PostEvent(self.TopLevel, stat_ev)

            stat_ev = evts.StatusEvent(msg='Short delay 1...', field=1)
            wx.PostEvent(self.TopLevel, stat_ev)
            time.sleep(0.1) # 5

            self.SetUpMeasThisRow(row)

            row_ev = evts.RowEvent(r = row)
            wx.PostEvent(self.RunPage, row_ev)

            #  V1...
            visastuff.ROLES_INSTR['DVM12'].SendCmd('LFREQ LINE') # dvmV1V2:'LFREQ LINE'
            time.sleep(0.5)
            visastuff.ROLES_INSTR['DVM12'].SendCmd('DCV,'+str(int(self.V1_set))) # dvmV1V2:'DCV'+str(self.V1_set)
            if self._want_abort:
                self.AbortRun()
                return

            stat_ev = evts.StatusEvent(msg='AqnThread.run():', field=0)
            wx.PostEvent(self.TopLevel, stat_ev)
            stat_ev = evts.StatusEvent(msg='Short delay 2...', field=1)
            wx.PostEvent(self.TopLevel, stat_ev)
            time.sleep(3) # 3

            # Set RS232 to V1
            visastuff.ROLES_INSTR['switchbox'].SendCmd(visastuff.SWITCH_CONFIGS['V1'])
            self.SetupPage.Switchbox.SetValue('V1') # update switchbox configuration icb
            visastuff.ROLES_INSTR['DVM12'].SendCmd('AZERO ON') # dvmV1V2: 'AZERO ON'
            if  self._want_abort:
                self.AbortRun()
                return

            stat_ev = evts.StatusEvent(msg='Range delay...', field=1)
            wx.PostEvent(self.TopLevel, stat_ev)
            time.sleep(self.range_del)

            stat_ev = evts.StatusEvent(msg='Measuring V1', field=1)
            wx.PostEvent(self.TopLevel, stat_ev)
            visastuff.ROLES_INSTR['DVM12'].Read()# junk = ...dvmV1V2
            visastuff.ROLES_INSTR['DVM12'].Read()# junk = ...dvmV1V2
            for i in range(self.n_readings):
                self.MeasureV1()
            self.T1 = self.ReadGMH(int(self.GMH1Port),int(self.GMH1Addr),self.GMH1Demo_status)

            # Update run displays on Run page via a DataEvent:
            t1 = str(dt.datetime.fromtimestamp(np.mean(self.V1Times)).strftime("%d/%m/%Y %H:%M:%S"))
            V1m = str(np.mean(self.V1Data))
            print 'AqnThread.run(): V1m =',V1m
            V1sd = str(np.std(self.V1Data,ddof=1))
            P = 100.0*pbar/(1 + self.stop_row - self.start_row) # % progress
            update_ev = evts.DataEvent(t=t1, Vm=V1m, Vsd=V1sd, P=P, r=row, flag='1')
            wx.PostEvent(self.RunPage,update_ev)

            #  V2...
            # Set RS232 to V2 BEFORE changing DVM range
            visastuff.ROLES_INSTR['switchbox'].SendCmd(visastuff.SWITCH_CONFIGS['V2'])
            self.SetupPage.Switchbox.SetValue('V2') # update switchbox configuration icb
            
            visastuff.ROLES_INSTR['DVM12'].SendCmd('DCV,'+str(self.V2_set)) # Reset DVM range
            if self._want_abort:
                self.AbortRun()
                return
            time.sleep(0.5) # was 0.1
            visastuff.ROLES_INSTR['DVM12'].SendCmd('LFREQ LINE') # dvmV1V2:'LFREQ LINE'
            
            stat_ev = evts.StatusEvent(msg='AqnThread.run():', field=0)
            wx.PostEvent(self.TopLevel, stat_ev)
            stat_ev = evts.StatusEvent(msg='Short delay 3...', field=1)
            wx.PostEvent(self.TopLevel, stat_ev)
            if self._want_abort:
                self.AbortRun()
                return
            time.sleep(3) # 3

            stat_ev = evts.StatusEvent(msg='Range delay...', field=1)
            wx.PostEvent(self.TopLevel, stat_ev)
            if self._want_abort:
                self.AbortRun()
                return
            time.sleep(self.range_del)

            stat_ev = evts.StatusEvent(msg='Measuring V2', field=1)
            wx.PostEvent(self.TopLevel, stat_ev)

            visastuff.ROLES_INSTR['DVM12'].Read() # dvmV1V2 (why these 2 unused reads?)
            visastuff.ROLES_INSTR['DVM12'].Read()# dvmV1V2
            for i in range(self.n_readings):
                self.MeasureV2()
            self.T2 = self.ReadGMH(int(self.GMH2Port),int(self.GMH2Addr),self.GMH2Demo_status)

            # Update displays on Run page via a DataEvent:
            t2 = str(dt.datetime.fromtimestamp(np.mean(self.V2Times)).strftime("%d/%m/%Y %H:%M:%S"))
            V2m = str(np.mean(self.V2Data))
            print 'AqnThread.run(): V2m =',V2m
            V2sd = str(np.std(self.V2Data,ddof=1))
            P = 100.0*pbar/(1 + self.stop_row - self.start_row) # % progress
            update_ev = evts.DataEvent(t=t2, Vm=V2m, Vsd=V2sd, P=P, r=row, flag='2')
            wx.PostEvent(self.RunPage, update_ev)

            #  Vd...
            # Set RS232 to Vd1
            visastuff.ROLES_INSTR['switchbox'].SendCmd(visastuff.SWITCH_CONFIGS['Vd1'])
            self.SetupPage.Switchbox.SetValue('Vd1') # update switchbox configuration icb
            visastuff.ROLES_INSTR['DVMd'].SendCmd('RANGE AUTO') # dvmVd:'RANGE AUTO'
            if self._want_abort:
                self.AbortRun()
                return
            stat_ev = evts.StatusEvent(msg='Range delay...', field=1)
            wx.PostEvent(self.TopLevel, stat_ev)
            time.sleep(self.range_del)

            stat_ev = evts.StatusEvent(msg='Measuring Vd', field=1)
            wx.PostEvent(self.TopLevel, stat_ev)
            visastuff.ROLES_INSTR['DVMd'].SendCmd('LFREQ LINE') # dvmVd   'LFREQ LINE'
            visastuff.ROLES_INSTR['DVMd'].Read() # dummy read
            for i in range(self.n_readings):
                self.MeasureVd()
            # Update displays on Run page via a DataEvent:
            td = str(dt.datetime.fromtimestamp(np.mean(self.VdTimes)).strftime("%d/%m/%Y %H:%M:%S"))
            Vdm = str(np.mean(self.VdData))
            print 'AqnThread.run(): Vdm =',Vdm
            Vdsd = str(np.std(self.VdData,ddof=1))
            P = 100.0*pbar/(1 + self.stop_row - self.start_row) # % progress
            update_ev = evts.DataEvent(t=td, Vm=Vdm, Vsd=Vdsd, P=P, r=row, flag='d')
            wx.PostEvent(self.RunPage, update_ev)

            self.WriteDataThisRow(row)

            # Plot data
            VdDates = []
            V1Dates = []
            V2Dates = []
            for d in self.VdTimes:
                VdDates.append(dt.datetime.fromtimestamp(d))
            for d in self.V1Times:
                V1Dates.append(dt.datetime.fromtimestamp(d))
            for d in self.V2Times:
                V2Dates.append(dt.datetime.fromtimestamp(d))
            clear_plot = 0
            if row == self.start_row:
                clear_plot = 1 # start each run with a clear plot 
            plot_ev = evts.PlotEvent(td=VdDates, t1=V1Dates, t2=V2Dates,
                                     Vd=self.VdData, V1=self.V1Data, V2=self.V2Data, clear=clear_plot)
            wx.PostEvent(self.PlotPage, plot_ev)
#            print'acquisition.run(): V1Data:',self.V1Data
            pbar += 1
            row += 1

        # (end of while loop):
        self.FinishRun()
        return

    def initialise(self):
        # This is a Dascon (%RH) PLACEHOLDER for now - replace with some actual code...
        # Set Dascon Outlets 1,3 to 'On' and initialise (room T & RH)

        stat_ev = evts.StatusEvent(msg='Initialising instruments...', field=0)
        wx.PostEvent(self.TopLevel, stat_ev)

        for r in visastuff.ROLES_INSTR.keys():
            d = visastuff.ROLES_WIDGETS[r]['icb'].GetValue()
            stat_ev = evts.StatusEvent(msg=d, field=1)
            wx.PostEvent(self.TopLevel, stat_ev)
            visastuff.ROLES_INSTR[r].Init()
            time.sleep(1)
        stat_ev = evts.StatusEvent(msg='Done', field=0)
        wx.PostEvent(self.TopLevel, stat_ev)


    def SetUpMeasThisRow(self,row):
        d = visastuff.ROLES_INSTR['SRC2'].Descr
        if d.endswith('F5520A'):
            err = visastuff.ROLES_INSTR['SRC2'].CheckErr() # srcV2  'ERR?', '*CLS'
            print 'Cleared F5520A error:',err
        time.sleep(3) # Wait 3 s after checking error
        # Get V1,V2 setting, n, delays from spreadsheet
        self.V1_set = self.ws.cell(row=row,column=1).value
        self.RunPage.V1Setting.SetValue(str(self.V1_set))
        if self._want_abort:
                self.AbortRun()
                return
        time.sleep(5) # wait 5 s after setting voltage
        self.V2_set = self.ws.cell(row=row,column=2).value
        self.RunPage.V2Setting.SetValue(str(self.V2_set))
        self.start_del = self.ws.cell(row=row,column=4).value
        if self._want_abort:
                self.AbortRun()
                return
        time.sleep(self.start_del)
        self.n_readings = self.ws.cell(row=row,column=3).value
        self.AZ1_del = self.ws.cell(row=row,column=5).value
        self.range_del = self.ws.cell(row=row,column=6).value
        del_ev = evts.DelaysEvent(n = self.n_readings,
                                  s = self.start_del,
                                  AZ1 = self.AZ1_del,
                                  r = self.range_del)
        wx.PostEvent(self.RunPage, del_ev)

        del self.V1Data[:]
        del self.V2Data[:]
        del self.VdData[:]
        del self.V1Times[:]
        del self.V2Times[:]
        del self.VdTimes[:]


    def MeasureV1(self):
        self.V1Times.append(time.time())
        if visastuff.ROLES_INSTR['DVM12'].Demo == True:
            dvmOP = np.random.normal(self.V1_set,1.0e-5*abs(self.V1_set))
            self.V1Data.append(dvmOP)
        else:
            # lfreq line, azero once,range auto, wait for settle
            dvmOP = visastuff.ROLES_INSTR['DVM12'].Read()# dvmV1V2
            self.V1Data.append(float(filter(self.filt,dvmOP)))

    def MeasureV2(self):
        self.V2Times.append(time.time())
        if visastuff.ROLES_INSTR['DVM12'].Demo == True:
            dvmOP = np.random.normal(self.V2_set,1.0e-5*abs(self.V2_set))
            self.V2Data.append(dvmOP)
        else:
            dvmOP = visastuff.ROLES_INSTR['DVM12'].Read() # dvmV1V2
            self.V2Data.append(float(filter(self.filt,dvmOP)))

    def MeasureVd(self):
        self.VdTimes.append(time.time())
        if self.AZ1_del > 0:
            visastuff.ROLES_INSTR['DVMd'].SendCmd('AZERO ONCE') # dvmVd: AZERO ONCE
            time.sleep(self.AZ1_del)
        if visastuff.ROLES_INSTR['DVMd'].Demo == True:
            dvmOP = np.random.normal(0.0,1.0e-6)
            self.VdData.append(dvmOP)
        else:
            dvmOP = visastuff.ROLES_INSTR['DVMd'].Read() # dvmVd
            self.VdData.append(float(filter(self.filt,dvmOP)))

    def ReadGMH(self,port,addr,demo_stat):
        com = GMH.ct.c_short(port)
        GMH.GMHLIB.GMH_OpenCom(com)
        Prio = GMH.ct.c_short()
        flData = GMH.ct.c_double() # Don't change this type!! It's the exactly right one!
        intData = GMH.ct.c_long()
        ValFunc = GMH.ct.c_short(0) # GetValue()
        GMH.GMHLIB.GMH_Transmit(addr,ValFunc,GMH.ct.byref(Prio),GMH.ct.byref(flData),GMH.ct.byref(intData))
        GMH.GMHLIB.GMH_CloseCom()
        if demo_stat == True:
            return np.random.normal(20.5,0.1)
        else:
            return flData.value


    def WriteDataThisRow(self,row):
        stat_ev = evts.StatusEvent(msg='AqnThread.WriteDataThisRow():', field=0)
        wx.PostEvent(self.TopLevel, stat_ev)
        stat_ev = evts.StatusEvent(msg='Row '+str(row), field=1)
        wx.PostEvent(self.TopLevel, stat_ev)

        self.ws['P'+str(row)] = str(dt.datetime.fromtimestamp(np.mean(self.V1Times)).strftime("%d/%m/%Y %H:%M:%S"))
        self.ws['Q'+str(row)] = np.mean(self.V1Data)
        self.ws['R'+str(row)] = np.std(self.V1Data,ddof=1)
        self.ws['G'+str(row)] = str(dt.datetime.fromtimestamp(np.mean(self.V2Times)).strftime("%d/%m/%Y %H:%M:%S"))
        self.ws['H'+str(row)] = np.mean(self.V2Data)
        self.ws['I'+str(row)] = np.std(self.V2Data,ddof=1)
        self.ws['M'+str(row)] = str(dt.datetime.fromtimestamp(np.mean(self.VdTimes)).strftime("%d/%m/%Y %H:%M:%S"))
        self.ws['N'+str(row)] = np.mean(self.VdData)
        self.ws['O'+str(row)] = np.std(self.VdData,ddof=1)

        if visastuff.ROLES_INSTR['DVMT1'].Demo == True:
            T1dvmOP = np.random.normal(108.0,1.0e-2)
            self.ws['S'+str(row)] = T1dvmOP
        else:
            T1dvmOP = visastuff.ROLES_INSTR['DVMT1'].SendCmd('READ?')
            self.ws['S'+str(row)] = float(filter(self.filt,T1dvmOP))

        if visastuff.ROLES_INSTR['DVMT2'].Demo == True:
            T2dvmOP = np.random.normal(108.0,1.0e-2)
            self.ws['T'+str(row)] = T2dvmOP
        else:
            T2dvmOP = visastuff.ROLES_INSTR['DVMT2'].SendCmd('READ?')
            self.ws['T'+str(row)] = float(filter(self.filt,T2dvmOP))

        self.ws['U'+str(row)] = self.T1
        self.ws['V'+str(row)] = self.T2
        # measure RH - write to col 23(W)
        self.ws['Z'+str(row)] = self.Comment
        
        """
        Need to add ambient parameters (T, P, %RH)
        to columns W, X, Y respectively.
        """

        self.wb_io.save(self.xlfilename)

    def AbortRun(self):
        # prematurely end run
        self.Standby() # Set sources to 0V and leave system safe

        stat_ev = evts.StatusEvent(msg='AbortRun(): Run stopped', field=0)
        wx.PostEvent(self.TopLevel, stat_ev)

        stop_ev = evts.DataEvent(t='-', Vm='-', Vsd='-', P=0, r='-',flag='E') # End
        wx.PostEvent(self.RunPage, stop_ev)

        self.RunPage.StartBtn.Enable(True)

    def FinishRun(self):
        # Run complete - leave system safe and final xl save
        self.wb_io.save(self.xlfilename)

        self.Standby() # Set sources to 0V and leave system safe

        stat_ev = evts.StatusEvent(msg='Closing instruments...', field=0)
        wx.PostEvent(self.TopLevel, stat_ev)

        stop_ev = evts.DataEvent(t='-', Vm='-', Vsd='-', P=0, r='-', flag='F') # Finished
        wx.PostEvent(self.RunPage, stop_ev)
        stat_ev = evts.StatusEvent(msg='RUN COMPLETED', field=0)
        wx.PostEvent(self.TopLevel, stat_ev)
        stat_ev = evts.StatusEvent(msg='', field=1)
        wx.PostEvent(self.TopLevel, stat_ev)

        self.RunPage.StartBtn.Enable(True)

    def Standby(self):
        # Set sources to 0V and disable outputs
        visastuff.ROLES_INSTR['SRC1'].SendCmd('R0=') # srcV1  'R0='
        self.RunPage.V1Setting.SetValue(str(0))
        self.RunPage.V2Setting.SetValue(str(0))
        
    def abort(self):
        """abort worker thread."""
        # Method for use by main thread to signal an abort
        stat_ev = evts.StatusEvent(msg='abort(): Run aborted', field=0)
        wx.PostEvent(self.TopLevel, stat_ev)
        self._want_abort = 1

    def filt(self,char):
        # A helper function to filter rubbish from DVM o/p unicode string
        # ...and retain any number (as a string)
        accept_str = u'-0.12345678eE9'
        return char in accept_str # Returns 'True' or 'False'

    # Only store 10 readings per line, and then clear
    V1Data=[]
    V2Data=[]
    VdData=[]
    V1Times=[]
    V2Times=[]
    VdTimes=[]
    TR1 = TR2 = 0

"""--------------End of Thread class definition-------------------"""