# -*- coding: utf-8 -*-
"""
Created on Wed Jun 24 09:36:42 2015

WORKING VERSION

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
# import os.path
# os.environ['XLPATH'] = 'C:\Documents and Settings\\t.lawson\My Documents\Python Scripts\High_Res_Bridge'

import numpy as np

# from openpyxl import load_workbook # WEDNESDAY
from openpyxl.styles import Font, Border, Side

import HighRes_events as evts
import devices  # visastuff
# import devices as GMH


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

#        self.V1Data = []
#        self.V2Data = []
#        V1Times
        self.VData = {'V1': [], 't1': [], 'V2': [], 't2': [],
                      'Va': [], 'ta': [], 'Vb': [], 'tb': [],
                      'Vc': [], 'tc': [], 'Vd': [], 'td': []}
        self.Times = []

        self.log = self.SetupPage.log

        self.Range_Mode = {True: 'AUTO', False: 'FIXED'}

        print'Role -> Instrument:'
        print >>self.log, 'Role -> Instrument:'
        print'------------------------------'
        print >>self.log, '------------------------------'
        # Print all GPIB instrument objects
        for r in devices.ROLES_WIDGETS.keys():
            d = devices.ROLES_WIDGETS[r]['icb'].GetValue()
            '''
            For 'switchbox' role, d is actually the setting (V1, Vd1,...)
            not the instrument description.
            '''
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

        # open existing workbook
        # 'data_only=True' ensures we read cell value, NOT formula
#        self.wb_io = load_workbook(self.xlfilename,data_only=True)
#        self.ws = self.wb_io.get_sheet_by_name('Data')

        # Find existing workbook
        self.wb_io = self.SetupPage.wb
        self.ws = self.wb_io.get_sheet_by_name('Data')

        # read start/stop row numbers from Excel file
        self.start_row = self.ws['B1'].value
        self.stop_row = self.ws['B2'].value
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
        Run Worker Thread.
        This is where all the important stuff goes, in a repeated cycle.
        '''

        # Set button availability
        self.RunPage.StopBtn.Enable(True)
        self.RunPage.StartBtn.Enable(False)
#        self.RunPage.RLinkBtn.Enable(False)

        # Clear plots
        clr_plot_ev = evts.ClearPlotEvent()
        wx.PostEvent(self.PlotPage, clr_plot_ev)

        # Column headings
        Head_row = self.start_row-2  # Main headings
        sub_row = self.start_row-1  # Sub-headings
        # Write unique id for this run.
        self.ws['A'+str(Head_row)] = 'Run Id:'
        self.ws['B'+str(Head_row)].font = Font(b=True)
        self.ws['B'+str(Head_row)] = str(self.RunPage.run_id)

        self.ws['A'+str(sub_row)] = 'V1_set'
        self.ws['B'+str(sub_row)] = 'V2_set'
        self.ws['C'+str(sub_row)] = 'n'
        self.ws['D'+str(sub_row)] = 'Row del.'
        self.ws['E'+str(sub_row)] = 'AZ1 del.'
        self.ws['F'+str(sub_row)] = 'Range del.'

        self.ws['H'+str(Head_row)] = 'V1'
        self.ws['H'+str(sub_row)] = 'Vmeas'
        self.ws['I'+str(sub_row)] = 'sd(V)'

        self.ws['J'+str(Head_row)] = 'V2'
        self.ws['J'+str(sub_row)] = 'Vmeas'
        self.ws['K'+str(sub_row)] = 'sd(V)'

        self.ws['L'+str(Head_row)] = 'Va'
        self.ws['L'+str(sub_row)] = 'Vmeas'
        self.ws['M'+str(sub_row)] = 'sd(V)'

        self.ws['N'+str(Head_row)] = 'Vb'
        self.ws['N'+str(sub_row)] = 'Vmeas'
        self.ws['O'+str(sub_row)] = 'sd(V)'

        self.ws['P'+str(Head_row)] = 'Vc'
        self.ws['P'+str(sub_row)] = 'Vmeas'
        self.ws['Q'+str(sub_row)] = 'sd(V)'

        self.ws['R'+str(Head_row)] = 'Vd'
        self.ws['R'+str(sub_row)] = 'Vmeas'
        self.ws['S'+str(sub_row)] = 'sd(V)'

        self.ws['T'+str(Head_row)] = 'dvm_T1'
        self.ws['T'+str(sub_row)] = '(Ohm)'
        self.ws['U'+str(Head_row)] = 'dvm_T2'
        self.ws['U'+str(sub_row)] = '(Ohm)'
        self.ws['V'+str(Head_row)] = 'GMH_T1'
        self.ws['V'+str(sub_row)] = '(deg C)'
        self.ws['W'+str(Head_row)] = 'GMH_T2'
        self.ws['W'+str(sub_row)] = '(deg C)'
        self.ws['X'+str(Head_row)] = 'Ambient Conditions'
        self.ws['X'+str(sub_row)] = 'T'
        self.ws['Y'+str(sub_row)] = 'P(mbar)'
        self.ws['Z'+str(sub_row)] = '%RH'

        self.ws['AA'+str(Head_row)] = 'Time'
        self.ws['AB'+str(Head_row)] = 'Comment'
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

        stat_ev = evts.StatusEvent(msg='',
                                   field='b')  # write to both status fields
        wx.PostEvent(self.TopLevel, stat_ev)

        stat_ev = evts.StatusEvent(msg='Post-initialise delay...', field=1)
        wx.PostEvent(self.TopLevel, stat_ev)
        time.sleep(3)

        # Get some initial temperatures...
#        self.ws['U'+str(self.start_row-1)] =\
#            devices.ROLES_INSTR['GMH1'].Measure('T')  # self.TR1
#        self.ws['V'+str(self.start_row-1)] =\
#            devices.ROLES_INSTR['GMH2'].Measure('T')  # self.TR2

        # Record ALL roles & corresponding instrument descriptions in XL sheet
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
            elif role_row == self.start_row + 8:  # last row
                self.ws['AC'+str(role_row)].border = bord_bl
                self.ws['AD'+str(role_row)].border = bord_br
            else:  # in-between rows
                self.ws['AC'+str(role_row)].border = bord_l
                self.ws['AD'+str(role_row)].border = bord_r
            self.ws['AC'+str(role_row)] = r
            d = devices.ROLES_WIDGETS[r]['icb'].GetValue()  # descr
            self.ws['AD'+str(role_row)] = d
            if r == 'DVM':
                self.ws['AE'+str(role_row)] =\
                    self.Range_Mode[self.RunPage.RangeTBtn.GetValue()]
            role_row += 1

        row = self.start_row
        pbar = 1

        # loop over xl rows..
        while row <= self.stop_row:

            stat_ev = evts.StatusEvent(msg='AqnThread.run():', field=0)
            wx.PostEvent(self.TopLevel, stat_ev)

#            stat_ev = evts.StatusEvent(msg='Row Delay...', field=1)
#            wx.PostEvent(self.TopLevel, stat_ev)
#            if self._want_abort:
#                self.AbortRun()
#                return
#            time.sleep(self.row_del)
#            if self._want_abort:
#                self.AbortRun()
#                return
#            time.sleep(5)
#
            self.SetUpMeasThisRow(row)
#            if self._want_abort:
#                self.AbortRun()
#                return

            row_ev = evts.RowEvent(r=row)
            wx.PostEvent(self.RunPage, row_ev)

# V1 ~~~~~~~~~~~~~~~~~~~~~~~~~~~
            devices.ROLES_INSTR['DVM'].SendCmd('LFREQ LINE')
            time.sleep(0.5)
            devices.ROLES_INSTR['DVM'].SendCmd('DCV,'+str(int(self.V1_set)))
            stat_ev = evts.StatusEvent(msg='AqnThread.run():', field=0)
            wx.PostEvent(self.TopLevel, stat_ev)
            stat_ev = evts.StatusEvent(msg='Range(DVM) delay...',
                                       field=1)
            wx.PostEvent(self.TopLevel, stat_ev)
            if self._want_abort:
                self.AbortRun()
                return
            time.sleep(self.range_del)

            # Set RS232 to V1 (and null relay to Vd1) AFTER DVM range-setting
            devices.ROLES_INSTR['switchbox'].\
                SendCmd(devices.SWITCH_CONFIGS['Vd1'])
            time.sleep(0.5)
            devices.ROLES_INSTR['switchbox'].\
                SendCmd(devices.SWITCH_CONFIGS['V1'])
            self.SetupPage.Switchbox.SetValue('V1')  # update sw-box display
            if self._want_abort:
                self.AbortRun()
                return
            stat_ev = evts.StatusEvent(msg='Relay delay...', field=1)
            wx.PostEvent(self.TopLevel, stat_ev)
            time.sleep(self.relay_del)

            devices.ROLES_INSTR['DVM'].SendCmd('AZERO ON')
            if self._want_abort:
                self.AbortRun()
                return
            stat_ev = evts.StatusEvent(msg='V1 Range delay...', field=1)
            wx.PostEvent(self.TopLevel, stat_ev)
            time.sleep(self.range_del)

            stat_ev = evts.StatusEvent(msg='Measuring V1', field=1)
            wx.PostEvent(self.TopLevel, stat_ev)
            devices.ROLES_INSTR['DVM'].Read()  # junk
            devices.ROLES_INSTR['DVM'].Read()  # junk
            for i in range(self.n_readings):
                self.MeasureV('V1')
            self.T1 = devices.ROLES_INSTR['GMH1'].Measure()  # Default is 'T'

            # Update run displays on Run page via a DataEvent:
            t_av = np.mean(self.VData['t1'])
            t1 = dt.datetime.fromtimestamp(t_av).\
                strftime("%d/%m/%Y %H:%M:%S")
            V1m = np.mean(self.VData['V1'])
            print 'AqnThread.run(): V1m =', V1m
            print >>self.log, 'AqnThread.run(): V1m =', V1m
            assert len(self.VData['V1']) > 1,\
                "Can't take SD of one or less items!"
            V1sd = np.std(self.VData['V1'], ddof=1)
            P = 100.0*pbar/(1 + self.stop_row - self.start_row)  # % progress
            update_ev = evts.DataEvent(t=t1, Vm=V1m, Vsd=V1sd, P=P,
                                       r=row, flag='1')
            wx.PostEvent(self.RunPage, update_ev)

# V2 ~~~~~~~~~~~~~~~~~~~~~~~~~~~
            # Set RS232 to V2 BEFORE changing DVM range
            devices.ROLES_INSTR['switchbox'].\
                SendCmd(devices.SWITCH_CONFIGS['V2'])
            self.SetupPage.Switchbox.SetValue('V2')  # update sw-box display

            devices.ROLES_INSTR['DVM'].SendCmd('LFREQ LINE')
            time.sleep(0.5)  # was 0.1

            # If running with fixed range set range to 'str(self.V1_set)':
            if self.RunPage.RangeTBtn.GetValue() == True:
                range2 = self.V2_set  # Auto-range
            else:
                range2 = self.V1_set  # Fixed range
            devices.ROLES_INSTR['DVM'].SendCmd('DCV,'+str(range2))
            if self._want_abort:
                self.AbortRun()
                return
            stat_ev = evts.StatusEvent(msg='V2 Range delay...', field=1)
            wx.PostEvent(self.TopLevel, stat_ev)
            time.sleep(self.range_del)

#            stat_ev = evts.StatusEvent(msg='AqnThread.run():', field=0)
#            wx.PostEvent(self.TopLevel, stat_ev)
#            stat_ev = evts.StatusEvent(msg='LFREQ LINE delay (3s)...',field=1)
#            wx.PostEvent(self.TopLevel, stat_ev)
#            if self._want_abort:
#                self.AbortRun()
#                return
#            time.sleep(3)

#            stat_ev = evts.StatusEvent(msg='Range delay...', field=1)
#            wx.PostEvent(self.TopLevel, stat_ev)
#            if self._want_abort:
#                self.AbortRun()
#                return
#            time.sleep(self.range_del)

            stat_ev = evts.StatusEvent(msg='Measuring V2', field=1)
            wx.PostEvent(self.TopLevel, stat_ev)

            devices.ROLES_INSTR['DVM'].Read()  # junk
            devices.ROLES_INSTR['DVM'].Read()  # junk
            for i in range(self.n_readings):
                self.MeasureV('V2')
            self.T2 = devices.ROLES_INSTR['GMH2'].Measure('T')

            # Update displays on Run page via a DataEvent:
            t2 = dt.datetime.fromtimestamp(np.mean(self.VData['t2'])).\
                strftime("%d/%m/%Y %H:%M:%S")
            V2m = np.mean(self.VData['V2'])
            print 'AqnThread.run(): V2m =', V2m
            print >>self.log, 'AqnThread.run(): V2m =', V2m
            assert len(self.VData['V2']) > 1,\
                "Can't take SD of one or less items!"
            V2sd = np.std(self.VData['V2'], ddof=1)
            P = 100.0*pbar/(1 + self.stop_row - self.start_row)  # % progress
            update_ev = evts.DataEvent(t=t2, Vm=V2m, Vsd=V2sd, P=P,
                                       r=row, flag='2')
            wx.PostEvent(self.RunPage, update_ev)

# Vc ~~~~~~~~~~~~~~~~~~~~~~~~~~~
            # Set RS232 to Vc BEFORE changing DVM range
            devices.ROLES_INSTR['switchbox'].\
                SendCmd(devices.SWITCH_CONFIGS['Vd1'])
            self.SetupPage.Switchbox.SetValue('Vd1')  # update sw-box display

            devices.ROLES_INSTR['DVM'].SendCmd('RANGE AUTO')
            if self._want_abort:
                self.AbortRun()
                return
            stat_ev = evts.StatusEvent(msg='Vd1 Range delay...', field=1)
            wx.PostEvent(self.TopLevel, stat_ev)
            time.sleep(self.range_del)

            stat_ev = evts.StatusEvent(msg='Measuring Vc', field=1)
            wx.PostEvent(self.TopLevel, stat_ev)

            devices.ROLES_INSTR['DVM'].SendCmd('LFREQ LINE')
            time.sleep(0.5)

            devices.ROLES_INSTR['DVM'].Read()  # dummy read
            for i in range(self.n_readings):
                self.MeasureV('Vc')

            # Update displays on Run page via a DataEvent:
            td = dt.datetime.fromtimestamp(np.mean(self.VData['tc'])).\
                strftime("%d/%m/%Y %H:%M:%S")
            Vdm = np.mean(self.VData['Vc'])
            print 'AqnThread.run(): Vdm =', Vdm
            print >>self.log, 'AqnThread.run(): Vdm =', Vdm
            assert len(self.VData['Vc']) > 1,\
                "Can't take SD of one or less items!"
            Vdsd = np.std(self.VData['Vc'], ddof=1)
            P = 100.0*pbar/(1 + self.stop_row - self.start_row)  # % progress
            update_ev = evts.DataEvent(t=td, Vm=Vdm, Vsd=Vdsd, P=P,
                                       r=row, flag='d')
            wx.PostEvent(self.RunPage, update_ev)

# Vd ~~~~~~~~~~~~~~~~~~~~~~~~~~~
            # Set RS232 to Vd BEFORE changing DVM range
            devices.ROLES_INSTR['switchbox'].\
                SendCmd(devices.SWITCH_CONFIGS['Vd2'])
            self.SetupPage.Switchbox.SetValue('Vd2')  # update sw-box display

            devices.ROLES_INSTR['DVM'].SendCmd('RANGE AUTO')
            if self._want_abort:
                self.AbortRun()
                return
            stat_ev = evts.StatusEvent(msg='Vd2 Range delay...', field=1)
            wx.PostEvent(self.TopLevel, stat_ev)
            time.sleep(self.range_del)

            stat_ev = evts.StatusEvent(msg='Measuring Vd', field=1)
            wx.PostEvent(self.TopLevel, stat_ev)

            devices.ROLES_INSTR['DVM'].SendCmd('LFREQ LINE')
            time.sleep(0.5)

            devices.ROLES_INSTR['DVM'].Read()  # dummy read
            for i in range(self.n_readings):
                self.MeasureV('Vd')

            # Update displays on Run page via a DataEvent:
            td = dt.datetime.fromtimestamp(np.mean(self.VData['td'])).\
                strftime("%d/%m/%Y %H:%M:%S")
            Vdm = np.mean(self.VData['Vd'])
            print 'AqnThread.run(): Vdm =', Vdm
            print >>self.log, 'AqnThread.run(): Vdm =', Vdm
            assert len(self.VData['Vd']) > 1,\
                "Can't take SD of one or less items!"
            Vdsd = np.std(self.VData['Vd'], ddof=1)
            P = 100.0*pbar/(1 + self.stop_row - self.start_row)  # % progress
            update_ev = evts.DataEvent(t=td, Vm=Vdm, Vsd=Vdsd, P=P,
                                       r=row, flag='d')
            wx.PostEvent(self.RunPage, update_ev)

# Vb ~~~~~~~~~~~~~~~~~~~~~~~~~~~
            # Set RS232 to Vb BEFORE changing DVM range
            devices.ROLES_INSTR['switchbox'].\
                SendCmd(devices.SWITCH_CONFIGS['V1'])
            time.sleep(0.5)
            devices.ROLES_INSTR['switchbox'].\
                SendCmd(devices.SWITCH_CONFIGS['Vd2'])
            self.SetupPage.Switchbox.SetValue('Vd2')  # update sw-box display

            devices.ROLES_INSTR['DVM'].SendCmd('RANGE AUTO')
            if self._want_abort:
                self.AbortRun()
                return
            stat_ev = evts.StatusEvent(msg='Vd2 Range delay...', field=1)
            wx.PostEvent(self.TopLevel, stat_ev)
            time.sleep(self.range_del)

            stat_ev = evts.StatusEvent(msg='Measuring Vb', field=1)
            wx.PostEvent(self.TopLevel, stat_ev)

            devices.ROLES_INSTR['DVM'].SendCmd('LFREQ LINE')
            time.sleep(0.5)

            devices.ROLES_INSTR['DVM'].Read()  # dummy read
            for i in range(self.n_readings):
                self.MeasureV('Vb')

            # Update displays on Run page via a DataEvent:
            td = dt.datetime.fromtimestamp(np.mean(self.VData['tb'])).\
                strftime("%d/%m/%Y %H:%M:%S")
            Vdm = np.mean(self.VData['Vb'])
            print 'AqnThread.run(): Vdm =', Vdm
            print >>self.log, 'AqnThread.run(): Vdm =', Vdm
            assert len(self.VData['Vb']) > 1,\
                "Can't take SD of one or less items!"
            Vdsd = np.std(self.VData['Vb'], ddof=1)
            P = 100.0*pbar/(1 + self.stop_row - self.start_row)  # % progress
            update_ev = evts.DataEvent(t=td, Vm=Vdm, Vsd=Vdsd, P=P,
                                       r=row, flag='d')
            wx.PostEvent(self.RunPage, update_ev)

# Va ~~~~~~~~~~~~~~~~~~~~~~~~~~~

            # Set RS232 to Va BEFORE changing DVM range
            devices.ROLES_INSTR['switchbox'].\
                SendCmd(devices.SWITCH_CONFIGS['Vd1'])
            self.SetupPage.Switchbox.SetValue('Vd1')  # update sw-box display

            devices.ROLES_INSTR['DVM'].SendCmd('RANGE AUTO')
            if self._want_abort:
                self.AbortRun()
                return
            stat_ev = evts.StatusEvent(msg='Vd1 Range delay...', field=1)
            wx.PostEvent(self.TopLevel, stat_ev)
            time.sleep(self.range_del)

            stat_ev = evts.StatusEvent(msg='Measuring Va', field=1)
            wx.PostEvent(self.TopLevel, stat_ev)

            devices.ROLES_INSTR['DVM'].SendCmd('LFREQ LINE')
            time.sleep(0.5)

            devices.ROLES_INSTR['DVM'].Read()  # dummy read
            for i in range(self.n_readings):
                self.MeasureV('Va')

            # Update displays on Run page via a DataEvent:
            td = dt.datetime.fromtimestamp(np.mean(self.VData['ta'])).\
                strftime("%d/%m/%Y %H:%M:%S")
            Vdm = np.mean(self.VData['Va'])
            print 'AqnThread.run(): Vdm =', Vdm
            print >>self.log, 'AqnThread.run(): Vdm =', Vdm
            assert len(self.VData['Va']) > 1,\
                "Can't take SD of one or less items!"
            Vdsd = np.std(self.VData['Va'], ddof=1)
            P = 100.0*pbar/(1 + self.stop_row - self.start_row)  # % progress
            update_ev = evts.DataEvent(t=td, Vm=Vdm, Vsd=Vdsd, P=P,
                                       r=row, flag='d')
# End of measurement sequence ~~~~~

            # Record room conditions
            if devices.ROLES_INSTR['GMHroom'].demo is False:
                self.Troom = devices.ROLES_INSTR['GMHroom'].Measure('T')
                self.Proom = devices.ROLES_INSTR['GMHroom'].Measure('P')
                self.RHroom = devices.ROLES_INSTR['GMHroom'].Measure('RH')
            else:
                self.Troom = self.Proom = self.RHroom = 0.0

#            self.Times.append(time.time())
            self.WriteDataThisRow(row)

            # Plot data
            Dates = []
            # (Use data for Vd as a proxy for all null data)
            for d in self.VData['td']:
                Dates.append(dt.datetime.fromtimestamp(d))

            clear_plot = 0
            if row == self.start_row:
                clear_plot = 1  # start each run with a clear plot
            plot_ev = evts.PlotEvent(td=Dates, t1=Dates, t2=Dates,
                                     Vd=self.VData['Vd'], V1=self.VData['V1'],
                                     V2=self.VData['V2'], clear=clear_plot)
            wx.PostEvent(self.PlotPage, plot_ev)

            pbar += 1
            row += 1

        # (end of while loop):
        self.FinishRun()
        return

    def initialise(self):
        # This is a Dascon (%RH) PLACEHOLDER for now...
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
            err = devices.ROLES_INSTR['SRC2'].CheckErr()
            print 'Cleared F5520A error:', err
            print >>self.log, 'Cleared F5520A error:', err
        if self._want_abort:
                self.AbortRun()
                return
        time.sleep(3)  # Wait 3 s after checking error

        # Get V1,V2 setting, n, delays from spreadsheet
        self.V1_set = self.ws.cell(row=row, column=1).value
        self.RunPage.V1Setting.SetValue(str(self.V1_set))

        self.V2_set = self.ws.cell(row=row, column=2).value
        self.RunPage.V2Setting.SetValue(str(self.V2_set))

        # Read delay settings from XL sheet
        self.n_readings = self.ws.cell(row=row, column=3).value
        self.row_del = self.ws.cell(row=row, column=4).value   # start_del
        self.AZ1_del = self.ws.cell(row=row, column=5).value
        self.range_del = self.ws.cell(row=row, column=6).value
        self.relay_del = self.ws.cell(row=row, column=7).value
        del_ev = evts.DelaysEvent(n=self.n_readings,
                                  s=self.row_del,
                                  AZ1=self.AZ1_del,
                                  r=self.range_del,
                                  rel=self.relay_del)
        wx.PostEvent(self.RunPage, del_ev)

        for k in self.VData:
            del self.VData[k][:]

        stat_ev = evts.StatusEvent(msg='Row Delay...', field=1)
        wx.PostEvent(self.TopLevel, stat_ev)
        if self._want_abort:
                self.AbortRun()
                return
        time.sleep(self.row_del)

    def MeasureV(self, node):
        assert node in ('V1', 'V2', 'Va', 'Vb', 'Vc', 'Vd'),\
            'Unknown argument to MeasureV().'
#        self.Times.append(time.time())
        self.VData['t' + node[-1]].append(time.time())
        if node in ('V1', 'V2'):

            if devices.ROLES_INSTR['DVM'].demo is True:
                dvmOP = np.random.normal(self.V1_set, 1.0e-5*abs(self.V1_set))
                self.VData[node].append(dvmOP)
            else:
                # lfreq line, azero once, range auto, wait for settle
                dvmOP = devices.ROLES_INSTR['DVM'].Read()
                self.VData[node].append(float(filter(self.filt, dvmOP)))
#        elif node == 'V2':
#            if devices.ROLES_INSTR['DVM12'].demo is True:
#                dvmOP = np.random.normal(self.V2_set, 1.0e-5*abs(self.V2_set))
#                self.V2Data.append(dvmOP)
#            else:
#                dvmOP = devices.ROLES_INSTR['DVM'].Read()
#                self.V2Data.append(float(filter(self.filt, dvmOP)))
        elif node in ('Va', 'Vb', 'Vc', 'Vd'):
            if self.AZ1_del > 0:
                devices.ROLES_INSTR['DVM'].SendCmd('AZERO ONCE')
                time.sleep(self.AZ1_del)
            if devices.ROLES_INSTR['DVM'].demo is True:
                dvmOP = np.random.normal(0.0, 1.0e-6)
                self.VData[node].append(dvmOP)
            else:
                dvmOP = devices.ROLES_INSTR['DVM'].Read()
                self.VData[node].append(float(filter(self.filt, dvmOP)))
            return 1

    def WriteDataThisRow(self, row):
        stat_ev = evts.StatusEvent(msg='AqnThread.WriteDataThisRow():',
                                   field=0)
        wx.PostEvent(self.TopLevel, stat_ev)
        stat_ev = evts.StatusEvent(msg='Row '+str(row), field=1)
        wx.PostEvent(self.TopLevel, stat_ev)

        self.ws['H'+str(row)] = np.mean(self.VData['V1'])
        print >>self.log, 'WriteDataThisRow(): cell', 'H'+str(row), ':',\
            np.mean(self.VData['V1'])
        self.ws['I'+str(row)] = np.std(self.VData['V1'], ddof=1)
        print >>self.log, 'WriteDataThisRow(): cell', 'I'+str(row),\
            np.std(self.VData['V1'], ddof=1)

        self.ws['J'+str(row)] = np.mean(self.VData['V2'])
        print >>self.log, 'WriteDataThisRow(): cell', 'J'+str(row), ':',\
            np.mean(self.VData['V2'])
        self.ws['K'+str(row)] = np.std(self.VData['V2'], ddof=1)
        print >>self.log, 'WriteDataThisRow(): cell', 'K'+str(row), ':',\
            np.std(self.VData['V2'], ddof=1)

        self.ws['L'+str(row)] = np.mean(self.VData['Va'])
        print >>self.log, 'WriteDataThisRow(): cell', 'L'+str(row), ':',\
            np.mean(self.VData['Va'])
        self.ws['M'+str(row)] = np.std(self.VData['Va'], ddof=1)
        print >>self.log, 'WriteDataThisRow(): cell', 'M'+str(row), ':',\
            np.std(self.VData['Va'], ddof=1)

        self.ws['N'+str(row)] = np.mean(self.VData['Vb'])
        print >>self.log, 'WriteDataThisRow(): cell', 'N'+str(row), ':',\
            np.mean(self.VData['Va'])
        self.ws['O'+str(row)] = np.std(self.VData['Vb'], ddof=1)
        print >>self.log, 'WriteDataThisRow(): cell', 'O'+str(row), ':',\
            np.std(self.VData['Va'], ddof=1)

        self.ws['P'+str(row)] = np.mean(self.VData['Vc'])
        print >>self.log, 'WriteDataThisRow(): cell', 'P'+str(row), ':',\
            np.mean(self.VData['Va'])
        self.ws['Q'+str(row)] = np.std(self.VData['Vc'], ddof=1)
        print >>self.log, 'WriteDataThisRow(): cell', 'Q'+str(row), ':',\
            np.std(self.VData['Va'], ddof=1)

        self.ws['R'+str(row)] = np.mean(self.VData['Vd'])
        print >>self.log, 'WriteDataThisRow(): cell', 'R'+str(row), ':',\
            np.mean(self.VData['Va'])
        self.ws['S'+str(row)] = np.std(self.VData['Vd'], ddof=1)
        print >>self.log, 'WriteDataThisRow(): cell', 'S'+str(row), ':',\
            np.std(self.VData['Va'], ddof=1)

        if devices.ROLES_INSTR['DVMT1'].demo is True:
            T1dvmOP = np.random.normal(108.0, 1.0e-2)
            self.ws['S'+str(row)] = T1dvmOP
            print >>self.log, 'WriteDataThisRow(): cell', 'T'+str(row), ':',\
                T1dvmOP
        else:
            T1dvmOP = devices.ROLES_INSTR['DVMT1'].SendCmd('READ?')
            self.ws['S'+str(row)] = float(filter(self.filt, T1dvmOP))
            print >>self.log, 'WriteDataThisRow(): cell', 'T'+str(row), ':',\
                float(filter(self.filt, T1dvmOP))

        if devices.ROLES_INSTR['DVMT2'].demo is True:
            T2dvmOP = np.random.normal(108.0, 1.0e-2)
            self.ws['U'+str(row)] = T2dvmOP
            print >>self.log, 'WriteDataThisRow(): cell', 'U'+str(row), ':',\
                T2dvmOP
        else:
            T2dvmOP = devices.ROLES_INSTR['DVMT2'].SendCmd('READ?')
            self.ws['U'+str(row)] = float(filter(self.filt, T2dvmOP))
            print >>self.log, 'WriteDataThisRow(): cell', 'U'+str(row), ':',\
                float(filter(self.filt, T2dvmOP))

        self.ws['V'+str(row)] = self.T1
        print >>self.log, 'WriteDataThisRow(): cell', 'V'+str(row), ':',\
            self.T1
        self.ws['W'+str(row)] = self.T2
        print >>self.log, 'WriteDataThisRow(): cell', 'W'+str(row), ':',\
            self.T2
        self.ws['X'+str(row)] = self.Troom
        print >>self.log, 'WriteDataThisRow(): cell', 'X'+str(row), ':',\
            self.Troom
        self.ws['Y'+str(row)] = self.Proom
        print >>self.log, 'WriteDataThisRow(): cell', 'Y'+str(row), ':',\
            self.Proom
        self.ws['Z'+str(row)] = self.RHroom
        print >>self.log, 'WriteDataThisRow(): cell', 'Z'+str(row), ':',\
            self.RHroom

        t_av = np.mean(self.VData['t1'] + self.VData['t2'] +
                       self.VData['ta'] + self.VData['tb'] +
                       self.VData['tc'] + self.VData['td'])
        timestamp = dt.datetime.fromtimestamp(t_av).\
            strftime("%d/%m/%Y %H:%M:%S")
        self.ws['AA'+str(row)] = str(timestamp)
        print >>self.log, 'WriteDataThisRow(): cell', 'AA'+str(row), ':',\
            str(timestamp)

        self.ws['AB'+str(row)] = self.Comment
        print >>self.log, 'WriteDataThisRow(): cell', 'AB'+str(row), ':',\
            self.Comment

        # Save after every row
        self.wb_io.save(self.xlfilename)

    def AbortRun(self):
        # prematurely end run, prompted by regular checks of _want_abort flag
        self._want_abort = 1
        self.Standby()  # Set sources to 0V and leave system safe

        stat_ev = evts.StatusEvent(msg='', field='b')
        wx.PostEvent(self.TopLevel, stat_ev)
        stat_ev = evts.StatusEvent(msg='AbortRun(): Run stopped', field=0)
        wx.PostEvent(self.TopLevel, stat_ev)

        stop_ev = evts.DataEvent(t='-', Vm='-', Vsd='-', P=0, r='-',
                                 flag='E')  # End
        wx.PostEvent(self.RunPage, stop_ev)

#        for r in devices.ROLES_INSTR.keys():
#            d = devices.ROLES_INSTR[r].Descr
#            if devices.ROLES_INSTR[r].demo == False:
#                print'AqnThread.AbortRun(): Closing',d
#                print >>self.log,'AqnThread.AbortRun(): Closing',d
#                devices.ROLES_INSTR[r].Close()
#            else:
#                print'AqnThread.AbortRun(): %s already closed'%d
#                print>>self.log,'AqnThread.AbortRun(): %s already closed'%d

        self.RunPage.StartBtn.Enable(True)
#        self.RunPage.RLinkBtn.Enable(True)
        self.RunPage.StopBtn.Enable(False)

    def FinishRun(self):
        # Run complete - leave system safe and final xl save
        self.wb_io.save(self.xlfilename)

        self.Standby()  # Set sources to 0V and leave system safe

#        stat_ev = evts.StatusEvent(msg='Closing instruments...', field=0)
#        wx.PostEvent(self.TopLevel, stat_ev)

        stop_ev = evts.DataEvent(t='-', Vm='-', Vsd='-', P=0, r='-',
                                 flag='F')  # Finished
        wx.PostEvent(self.RunPage, stop_ev)
        stat_ev = evts.StatusEvent(msg='RUN COMPLETED', field=0)
        wx.PostEvent(self.TopLevel, stat_ev)
        stat_ev = evts.StatusEvent(msg='', field=1)
        wx.PostEvent(self.TopLevel, stat_ev)

#        for r in devices.ROLES_INSTR.keys():
#            d = devices.ROLES_INSTR[r].Descr
#            if devices.ROLES_INSTR[r].demo == False:
#                print'AqnThread.FinishRun(): Closing',d
#                devices.ROLES_INSTR[r].Close()
#            else:
#                print'AqnThread.FinishRun(): %s already closed'%d

        self.RunPage.StartBtn.Enable(True)
#        self.RunPage.RLinkBtn.Enable(True)
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
