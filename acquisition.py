# -*- coding: utf-8 -*-
"""
Created on Wed Jun 24 09:36:42 2015

DEVELOPMENT VERSION

@author: t.lawson

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

import numpy as np

from openpyxl.styles import Font, Border, Side

import HighRes_events as evts
import devices


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

        self.V1Data = []
        self.V2Data = []
        self.VdData = []
        self.V1Times = []
        self.V2Times = []
        self.VdTimes = []

        self.log = self.SetupPage.log

        self.Range_Mode = {True: 'AUTO', False: 'FIXED'}

        print'Role -> Instrument:'
        print >>self.log, 'Role -> Instrument:'
        print'------------------------------'
        print >>self.log, '------------------------------'
        # Print all GPIB instrument objects
        for r in devices.ROLES_WIDGETS.keys():
            d = devices.ROLES_WIDGETS[r]['icb'].GetValue()
            # For 'switchbox' role, d is actually the setting (V1, Vd1,...)
            # not the instrument description.

            print'%s -> %s' % (devices.INSTR_DATA[d]['role'], d)
            print >>self.log, '%s -> %s' % (devices.INSTR_DATA[d]['role'], d)
            if r != devices.INSTR_DATA[d]['role']:
                devices.INSTR_DATA[d]['role'] = r
                print'Role data corrected to:', r, '->', d
                print >>self.log, 'Role data corrected to:', r, '->', d

        # Get filename of Excel file
        self.xlfilename = self.SetupPage.XLFile.GetValue()  # Full path
        self.path_components = self.xlfilename.split('\\')
        self.directory = '\\'.join(self.path_components[0:-1])

        # Find existing workbook
        self.wb_io = self.SetupPage.wb
        self.ws = self.wb_io.get_sheet_by_name('Data')

        # read start/stop row numbers from Excel file
        self.start_row = self.ws['B1'].value
        self.stop_row = self.ws['B2'].value
        assert self.stop_row >= self.start_row, 'Stop row must be > start row!'
        strt_ev = evts.StartRowEvent(row=self.start_row)
        wx.PostEvent(self.RunPage, strt_ev)
        stp_ev = evts.StopRowEvent(row=self.stop_row)
        wx.PostEvent(self.RunPage, stp_ev)

        self.settle_time = self.RunPage.SettleDel.GetValue()

        # Local record of GMH ports and addresses
        self.GMH1Demo_status = devices.ROLES_INSTR['GMH1'].demo
        self.GMH2Demo_status = devices.ROLES_INSTR['GMH2'].demo
        self.GMH1Port = devices.ROLES_INSTR['GMH1'].addr
        self.GMH2Port = devices.ROLES_INSTR['GMH2'].addr

        self.start()  # Starts the thread running on creation

    def run(self):
        '''
        Run Worker Thread. This is where all the important stuff goes,
        in a repeated cycle
        '''
        # Set button availability
        self.RunPage.StopBtn.Enable(True)
        self.RunPage.StartBtn.Enable(False)
        self.RunPage.RLinkBtn.Enable(False)

        # Clear plots
        clr_plot_ev = evts.ClearPlotEvent()
        wx.PostEvent(self.PlotPage, clr_plot_ev)

        # Column headings
        Head_row = self.start_row-2  # Main headings
        sub_row = self.start_row-1  # Sub-headings

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
        self.ws['AE'+str(Head_row)] = 'Range mode'

        stat_ev = evts.StatusEvent(msg='AqnThread.run():', field=0)
        wx.PostEvent(self.TopLevel, stat_ev)
        stat_ev = evts.StatusEvent(msg='Waiting to settle...', field=1)
        wx.PostEvent(self.TopLevel, stat_ev)

        time.sleep(self.settle_time)

        # Initialise all instruments (doesn't open GMH sensors yet)
        self.initialise()

        stat_ev = evts.StatusEvent(msg='', field='b')  # write to both fields
        wx.PostEvent(self.TopLevel, stat_ev)

        stat_ev = evts.StatusEvent(msg='Post-initialise delay...', field=1)
        wx.PostEvent(self.TopLevel, stat_ev)
        time.sleep(3)  # 3

        # Get some initial temperatures...
        TR1 = devices.ROLES_INSTR['GMH1'].Measure('T')
        TR2 = devices.ROLES_INSTR['GMH2'].Measure('T')
        self.ws['U'+str(self.start_row-1)] = TR1
        self.ws['V'+str(self.start_row-1)] = TR2

        # Record ALL roles and corresponding instrument descriptions
        role_row = self.start_row
        bord_tl = Border(top=Side(style='thin'), left=Side(style='thin'))
        bord_tr = Border(top=Side(style='thin'), right=Side(style='thin'))
        bord_l = Border(left=Side(style='thin'))
        bord_r = Border(right=Side(style='thin'))
        bord_bl = Border(bottom=Side(style='thin'), left=Side(style='thin'))
        bord_br = Border(bottom=Side(style='thin'), right=Side(style='thin'))
        for r in devices.ROLES_WIDGETS.keys():
            if role_row == self.start_row:  # 1st row
                self.ws['AC'+str(role_row)].border = bord_tl
                self.ws['AD'+str(role_row)].border = bord_tr
            elif role_row == self.start_row + 9:  # last row
                self.ws['AC'+str(role_row)].border = bord_bl
                self.ws['AD'+str(role_row)].border = bord_br
            else:  # in-between rows
                self.ws['AC'+str(role_row)].border = bord_l
                self.ws['AD'+str(role_row)].border = bord_r
            self.ws['AC'+str(role_row)] = r
            d = devices.ROLES_WIDGETS[r]['icb'].GetValue()
            self.ws['AD'+str(role_row)] = d
            if r == 'DVM12':
                range_mode = self.Range_Mode[self.RunPage.RangeTBtn.GetValue()]
                self.ws['AE'+str(role_row)] = range_mode
            role_row += 1

        row = self.start_row
        pbar = 1

        # loop over xl rows..
        while row <= self.stop_row:
            if self._want_abort:
                self.AbortRun()
                return

            if self._want_abort:
                self.AbortRun()
                return
            stat_ev = evts.StatusEvent(msg='AqnThread.run():', field=0)
            wx.PostEvent(self.TopLevel, stat_ev)

            stat_ev = evts.StatusEvent(msg='Short delay 1...', field=1)
            wx.PostEvent(self.TopLevel, stat_ev)
            time.sleep(5)

            self.SetUpMeasThisRow(row)

            row_ev = evts.RowEvent(r=row)
            wx.PostEvent(self.RunPage, row_ev)

            #  V1...
            devices.ROLES_INSTR['DVM12'].SendCmd('LFREQ LINE')
            time.sleep(0.5)
            devices.ROLES_INSTR['DVM12'].SendCmd('DCV,'+str(int(self.V1_set)))
            if self._want_abort:
                self.AbortRun()
                return

            stat_ev = evts.StatusEvent(msg='AqnThread.run():', field=0)
            wx.PostEvent(self.TopLevel, stat_ev)
            stat_ev = evts.StatusEvent(msg='Short delay 2...', field=1)
            wx.PostEvent(self.TopLevel, stat_ev)
            time.sleep(3)  # 3

            # Set RS232 to V1
            devices.ROLES_INSTR['switchbox'].SendCmd(devices.SWITCH_CONFIGS['V1'])
            self.SetupPage.Switchbox.SetValue('V1')
            devices.ROLES_INSTR['DVM12'].SendCmd('AZERO ON')
            if self._want_abort:
                self.AbortRun()
                return

            stat_ev = evts.StatusEvent(msg='Range delay...', field=1)
            wx.PostEvent(self.TopLevel, stat_ev)
            time.sleep(self.range_del)

            stat_ev = evts.StatusEvent(msg='Measuring V1', field=1)
            wx.PostEvent(self.TopLevel, stat_ev)
            devices.ROLES_INSTR['DVM12'].Read()
            devices.ROLES_INSTR['DVM12'].Read()
            for i in range(self.n_readings):
                self.MeasureV('V1')
            self.T1 = devices.ROLES_INSTR['GMH1'].Measure('T')

            # Update run displays on Run page via a DataEvent:
            t1 = dt.datetime.fromtimestamp(np.mean(self.V1Times)).strftime("%d/%m/%Y %H:%M:%S")
            V1m = np.mean(self.V1Data)
            print 'AqnThread.run(): V1m =', V1m
            print >>self.log, 'AqnThread.run(): V1m =', V1m
            assert len(self.V1Data) > 1, "Can't take SD of one or less items!"
            V1sd = np.std(self.V1Data, ddof=1)
            P = 100.0*pbar/(1 + self.stop_row - self.start_row)  # % progress
            update_ev = evts.DataEvent(t=t1, Vm=V1m, Vsd=V1sd,
                                       P=P, r=row, flag='1')
            wx.PostEvent(self.RunPage, update_ev)

            #  V2...
            # Set RS232 to V2 BEFORE changing DVM range
            devices.ROLES_INSTR['switchbox'].SendCmd(devices.SWITCH_CONFIGS['V2'])
            self.SetupPage.Switchbox.SetValue('V2')

            # If running with fixed range set range to 'str(self.V1_set)':
            if self.RunPage.RangeTBtn.GetValue() == True:
                range2 = self.V2_set
            else:
                range2 = self.V1_set
            devices.ROLES_INSTR['DVM12'].SendCmd('DCV,'+str(range2))
            if self._want_abort:
                self.AbortRun()
                return
            time.sleep(0.5)  # was 0.1
            devices.ROLES_INSTR['DVM12'].SendCmd('LFREQ LINE')

            stat_ev = evts.StatusEvent(msg='AqnThread.run():', field=0)
            wx.PostEvent(self.TopLevel, stat_ev)
            stat_ev = evts.StatusEvent(msg='Short delay 3...', field=1)
            wx.PostEvent(self.TopLevel, stat_ev)
            if self._want_abort:
                self.AbortRun()
                return
            time.sleep(3)  # 3

            stat_ev = evts.StatusEvent(msg='Range delay...', field=1)
            wx.PostEvent(self.TopLevel, stat_ev)
            if self._want_abort:
                self.AbortRun()
                return
            time.sleep(self.range_del)

            stat_ev = evts.StatusEvent(msg='Measuring V2', field=1)
            wx.PostEvent(self.TopLevel, stat_ev)

            devices.ROLES_INSTR['DVM12'].Read()
            devices.ROLES_INSTR['DVM12'].Read()
            for i in range(self.n_readings):
                self.MeasureV('V2')
            self.T2 = devices.ROLES_INSTR['GMH2'].Measure('T')

            # Update displays on Run page via a DataEvent:
            t2 = dt.datetime.fromtimestamp(np.mean(self.V2Times)).strftime("%d/%m/%Y %H:%M:%S")
            V2m = np.mean(self.V2Data)
            print 'AqnThread.run(): V2m =', V2m
            print >>self.log, 'AqnThread.run(): V2m =', V2m
            assert len(self.V2Data) > 1, "Can't take SD of one or less items!"
            V2sd = np.std(self.V2Data, ddof=1)
            P = 100.0*pbar/(1 + self.stop_row - self.start_row)  # % progress
            update_ev = evts.DataEvent(t=t2, Vm=V2m, Vsd=V2sd,
                                       P=P, r=row, flag='2')
            wx.PostEvent(self.RunPage, update_ev)

            #  Vd...
            # Set RS232 to Vd1
            devices.ROLES_INSTR['switchbox'].SendCmd(devices.SWITCH_CONFIGS['Vd1'])
            self.SetupPage.Switchbox.SetValue('Vd1')
            devices.ROLES_INSTR['DVMd'].SendCmd('RANGE AUTO')
            if self._want_abort:
                self.AbortRun()
                return
            stat_ev = evts.StatusEvent(msg='Range delay...', field=1)
            wx.PostEvent(self.TopLevel, stat_ev)
            time.sleep(self.range_del)

            stat_ev = evts.StatusEvent(msg='Measuring Vd', field=1)
            wx.PostEvent(self.TopLevel, stat_ev)
            devices.ROLES_INSTR['DVMd'].SendCmd('LFREQ LINE')
            devices.ROLES_INSTR['DVMd'].Read()  # dummy read
            for i in range(self.n_readings):
                self.MeasureV('Vd')
            # Update displays on Run page via a DataEvent:
            td = dt.datetime.fromtimestamp(np.mean(self.VdTimes)).strftime("%d/%m/%Y %H:%M:%S")
            Vdm = np.mean(self.VdData)
            print 'AqnThread.run(): Vdm =', Vdm
            print >>self.log, 'AqnThread.run(): Vdm =', Vdm
            assert len(self.VdData) > 1, "Can't take SD of one or less items!"
            Vdsd = np.std(self.VdData, ddof=1)
            P = 100.0*pbar/(1 + self.stop_row - self.start_row)
            update_ev = evts.DataEvent(t=td, Vm=Vdm, Vsd=Vdsd,
                                       P=P, r=row, flag='d')
            wx.PostEvent(self.RunPage, update_ev)

            # Record room conditions
            if devices.ROLES_INSTR['GMHroom'].demo is False:
                self.Troom = devices.ROLES_INSTR['GMHroom'].Measure('T')
                self.Proom = devices.ROLES_INSTR['GMHroom'].Measure('P')
                self.RHroom = devices.ROLES_INSTR['GMHroom'].Measure('RH')
            else:
                self.Troom = self.Proom = self.RHroom = 0.0

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
                clear_plot = 1  # start each run with a clear plot
            plot_ev = evts.PlotEvent(td=VdDates, t1=V1Dates, t2=V2Dates,
                                     Vd=self.VdData, V1=self.V1Data,
                                     V2=self.V2Data, clear=clear_plot)
            wx.PostEvent(self.PlotPage, plot_ev)
            pbar += 1
            row += 1

        # (end of while loop):
        self.FinishRun()
        return

    def initialise(self):
        # This is a Dascon (%RH) PLACEHOLDER for now - replace with some
        # actual code...
        # Set Dascon Outlets 1,3 to 'On' and initialise (room T & RH)

        stat_ev = evts.StatusEvent(msg='Initialising instruments...', field=0)
        wx.PostEvent(self.TopLevel, stat_ev)

        for r in devices.ROLES_INSTR.keys():
            d = devices.ROLES_WIDGETS[r]['icb'].GetValue()
            # Open non-GMH devices:
            if 'GMH' not in devices.ROLES_INSTR[r].Descr:
                print'AqnThread.initialise(): Opening', d
                print >>self.log, 'AqnThread.initialise(): Opening', d
                devices.ROLES_INSTR[r].Open()
            else:
                print'AqnThread.initialise(): %s already open' % d
                print >>self.log, 'AqnThread.initialise(): %s already open' % d

            stat_ev = evts.StatusEvent(msg=d, field=1)
            wx.PostEvent(self.TopLevel, stat_ev)
            devices.ROLES_INSTR[r].Init()
            time.sleep(1)
        stat_ev = evts.StatusEvent(msg='Done', field=0)
        wx.PostEvent(self.TopLevel, stat_ev)

    def SetUpMeasThisRow(self, row):
        d = devices.ROLES_INSTR['SRC2'].Descr
        if d.endswith('F5520A'):
            err = devices.ROLES_INSTR['SRC2'].CheckErr()  # Clear error buffer
            print 'Cleared F5520A error:', err
            print >>self.log, 'Cleared F5520A error:', err
        time.sleep(3)  # Wait 3 s after checking error
        # Get V1,V2 setting, n, delays from spreadsheet
        self.V1_set = self.ws.cell(row=row, column=1).value
        self.RunPage.V1Setting.SetValue(str(self.V1_set))
        if self._want_abort:
                self.AbortRun()
                return
        time.sleep(5)  # wait 5 s after setting voltage
        self.V2_set = self.ws.cell(row=row, column=2).value
        self.RunPage.V2Setting.SetValue(str(self.V2_set))
        self.start_del = self.ws.cell(row=row, column=4).value
        if self._want_abort:
                self.AbortRun()
                return
        time.sleep(self.start_del)
        self.n_readings = self.ws.cell(row=row, column=3).value
        self.AZ1_del = self.ws.cell(row=row, column=5).value
        self.range_del = self.ws.cell(row=row, column=6).value
        del_ev = evts.DelaysEvent(n=self.n_readings,
                                  s=self.start_del,
                                  AZ1=self.AZ1_del,
                                  r=self.range_del)
        wx.PostEvent(self.RunPage, del_ev)

        del self.V1Data[:]
        del self.V2Data[:]
        del self.VdData[:]
        del self.V1Times[:]
        del self.V2Times[:]
        del self.VdTimes[:]


    def MeasureV(self, node):
        assert node in ('V1', 'V2', 'Vd'), 'Unknown argument to MeasureV().'
        if node == 'V1':
            self.V1Times.append(time.time())
            if devices.ROLES_INSTR['DVM12'].demo is True:
                dvmOP = np.random.normal(self.V1_set, 1.0e-5*abs(self.V1_set))
                self.V1Data.append(dvmOP)
            else:
                # lfreq line, azero once,range auto, wait for settle
                dvmOP = devices.ROLES_INSTR['DVM12'].Read()
                self.V1Data.append(float(filter(self.filt, dvmOP)))
        elif node == 'V2':
            self.V2Times.append(time.time())
            if devices.ROLES_INSTR['DVM12'].demo is True:
                dvmOP = np.random.normal(self.V2_set, 1.0e-5*abs(self.V2_set))
                self.V2Data.append(dvmOP)
            else:
                dvmOP = devices.ROLES_INSTR['DVM12'].Read()
                self.V2Data.append(float(filter(self.filt, dvmOP)))
        elif node == 'Vd':
            self.VdTimes.append(time.time())
            if self.AZ1_del > 0:
                devices.ROLES_INSTR['DVMd'].SendCmd('AZERO ONCE')
                time.sleep(self.AZ1_del)
            if devices.ROLES_INSTR['DVMd'].demo is True:
                dvmOP = np.random.normal(0.0, 1.0e-6)
                self.VdData.append(dvmOP)
            else:
                dvmOP = devices.ROLES_INSTR['DVMd'].Read()
                self.VdData.append(float(filter(self.filt, dvmOP)))
            return 1

    def WriteDataThisRow(self, row):
        stat_ev = evts.StatusEvent(msg='AqnThread.WriteDataThisRow():',
                                   field=0)
        wx.PostEvent(self.TopLevel, stat_ev)
        stat_ev = evts.StatusEvent(msg='Row '+str(row), field=1)
        wx.PostEvent(self.TopLevel, stat_ev)

        self.ws['P'+str(row)] = str(dt.datetime.fromtimestamp(np.mean(self.V1Times)).strftime("%d/%m/%Y %H:%M:%S"))
        print >>self.log, 'WriteDataThisRow(): cell', 'P'+str(row), ':', str(dt.datetime.fromtimestamp(np.mean(self.V1Times)).strftime("%d/%m/%Y %H:%M:%S"))
        self.ws['Q'+str(row)] = np.mean(self.V1Data)
        print >>self.log, 'WriteDataThisRow(): cell', 'Q'+str(row), ':', np.mean(self.V1Data)
        self.ws['R'+str(row)] = np.std(self.V1Data, ddof=1)
        print >>self.log, 'WriteDataThisRow(): cell', 'R'+str(row), np.std(self.V1Data, ddof=1)
        self.ws['G'+str(row)] = str(dt.datetime.fromtimestamp(np.mean(self.V2Times)).strftime("%d/%m/%Y %H:%M:%S"))
        print >>self.log, 'WriteDataThisRow(): cell', 'G'+str(row), ':', str(dt.datetime.fromtimestamp(np.mean(self.V2Times)).strftime("%d/%m/%Y %H:%M:%S"))
        self.ws['H'+str(row)] = np.mean(self.V2Data)
        print >>self.log, 'WriteDataThisRow(): cell', 'H'+str(row), ':', np.mean(self.V2Data)
        self.ws['I'+str(row)] = np.std(self.V2Data, ddof=1)
        print >>self.log, 'WriteDataThisRow(): cell', 'I'+str(row), ':', np.std(self.V2Data, ddof=1)
        self.ws['M'+str(row)] = str(dt.datetime.fromtimestamp(np.mean(self.VdTimes)).strftime("%d/%m/%Y %H:%M:%S"))
        print >>self.log, 'WriteDataThisRow(): cell', 'M'+str(row), ':', str(dt.datetime.fromtimestamp(np.mean(self.VdTimes)).strftime("%d/%m/%Y %H:%M:%S"))
        self.ws['N'+str(row)] = np.mean(self.VdData)
        print >>self.log, 'WriteDataThisRow(): cell', 'N'+str(row), ':', np.mean(self.VdData)
        self.ws['O'+str(row)] = np.std(self.VdData, ddof=1)
        print >>self.log, 'WriteDataThisRow(): cell', 'O'+str(row), ':', np.std(self.VdData, ddof=1)

        if devices.ROLES_INSTR['DVMT1'].demo is True:
            T1dvmOP = np.random.normal(108.0, 1.0e-2)
            self.ws['S'+str(row)] = T1dvmOP
            print >>self.log, 'WriteDataThisRow(): cell', 'S'+str(row), ':', T1dvmOP
        else:
            T1dvmOP = devices.ROLES_INSTR['DVMT1'].SendCmd('READ?')
            self.ws['S'+str(row)] = float(filter(self.filt, T1dvmOP))
            print >>self.log, 'WriteDataThisRow(): cell', 'S'+str(row), ':', float(filter(self.filt, T1dvmOP))

        if devices.ROLES_INSTR['DVMT2'].demo is True:
            T2dvmOP = np.random.normal(108.0, 1.0e-2)
            self.ws['T'+str(row)] = T2dvmOP
            print >>self.log, 'WriteDataThisRow(): cell', 'T'+str(row), ':', T2dvmOP
        else:
            T2dvmOP = devices.ROLES_INSTR['DVMT2'].SendCmd('READ?')
            self.ws['T'+str(row)] = float(filter(self.filt, T2dvmOP))
            print >>self.log, 'WriteDataThisRow(): cell', 'T'+str(row), ':', float(filter(self.filt, T2dvmOP))

        self.ws['U'+str(row)] = self.T1
        print >>self.log, 'WriteDataThisRow(): cell', 'U'+str(row), ':', self.T1
        self.ws['V'+str(row)] = self.T2
        print >>self.log, 'WriteDataThisRow(): cell', 'V'+str(row), ':', self.T2
        self.ws['W'+str(row)] = self.Troom
        print >>self.log, 'WriteDataThisRow(): cell', 'W'+str(row), ':', self.Troom
        self.ws['X'+str(row)] = self.Proom
        print >>self.log, 'WriteDataThisRow(): cell', 'X'+str(row), ':', self.Proom
        self.ws['Y'+str(row)] = self.RHroom
        print >>self.log, 'WriteDataThisRow(): cell', 'Y'+str(row), ':', self.RHroom
        self.ws['Z'+str(row)] = self.Comment
        print >>self.log, 'WriteDataThisRow(): cell', 'Z'+str(row), ':', self.Comment

        # Save after every row
        self.wb_io.save(self.xlfilename)

    def AbortRun(self):
        # prematurely end run, prompted by regular checks of _want_abort flag
        self.Standby()  # Set sources to 0V and leave system safe

        stop_ev = evts.DataEvent(t='-', Vm='-', Vsd='-', P=0, r='-',flag='E')
        wx.PostEvent(self.RunPage, stop_ev)

        self.RunPage.StartBtn.Enable(True)
        self.RunPage.RLinkBtn.Enable(True)
        self.RunPage.StopBtn.Enable(False)

    def FinishRun(self):
        '''
        Run complete - leave system safe and final xl save
        '''
        self.wb_io.save(self.xlfilename)

        self.Standby()  # Set sources to 0V and leave system safe

        stop_ev = evts.DataEvent(t='-', Vm='-', Vsd='-', P=0, r='-', flag='F')
        wx.PostEvent(self.RunPage, stop_ev)
        stat_ev = evts.StatusEvent(msg='RUN COMPLETED', field=0)
        wx.PostEvent(self.TopLevel, stat_ev)
        stat_ev = evts.StatusEvent(msg='', field=1)
        wx.PostEvent(self.TopLevel, stat_ev)

        self.RunPage.StartBtn.Enable(True)
        self.RunPage.RLinkBtn.Enable(True)
        self.RunPage.StopBtn.Enable(False)

    def Standby(self):
        # Set sources to 0V and disable outputs
        devices.ROLES_INSTR['SRC1'].SendCmd('R0=')
        self.RunPage.V1Setting.SetValue(str(0))
        self.RunPage.V2Setting.SetValue(str(0))

    def abort(self):
        """abort worker thread."""
        # Method for use by main thread to signal an abort
        stat_ev = evts.StatusEvent(msg='abort(): Run aborted', field=0)
        wx.PostEvent(self.TopLevel, stat_ev)
        self._want_abort = 1

    def filt(self, char):
        # A helper function to filter rubbish from DVM o/p unicode string
        # ...and retain any number (as a string)
        accept_str = u'-0.12345678eE9'
        return char in accept_str  # Returns 'True' or 'False'

"""--------------End of Thread class definition-------------------"""
