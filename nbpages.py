# -*- coding: utf-8 -*-
""" nbpages.py - Defines individual notebook pages as panel-like objects

WORKING VERSION

Created on Tue Jun 30 10:10:16 2015

@author: t.lawson
"""

import os

import wx
from wx.lib.masked import NumCtrl
import datetime as dt
import time

import matplotlib
matplotlib.use('WXAgg')  # Agg renderer for drawing on a wx canvas
from matplotlib.backends.backend_wxagg import FigureCanvasWxAgg as FigureCanvas
#from matplotlib.backends.backend_wx import NavigationToolbar2Wx
from matplotlib.figure import Figure
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib.ticker as mtick
from openpyxl import load_workbook, utils  # , cell

import HighRes_events as evts
import acquisition as acq
import RLink as rl
import devices

matplotlib.rc('lines', linewidth=1, color='blue')

# os.environ['XLPATH'] = 'C:\Documents and Settings\\t.lawson\My Documents\Python Scripts\High_Res_Bridge'
'''
------------------------
# Setup Page definition:
------------------------
'''


class SetupPage(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)

        # Event bindings
        self.Bind(evts.EVT_FILEPATH, self.update_filepath)

        self.status = self.GetTopLevelParent().sb

        self.SRC_COMBO_CHOICE = ['none']
        self.DVM_COMBO_CHOICE = ['none']
        self.GMH_COMBO_CHOICE = ['none']
        self.SB_COMBO_CHOICE = list(devices.SWITCH_CONFIGS.keys())
        self.T_SENSOR_CHOICE = devices.T_Sensors
        self.cbox_addr_COM = []
        self.cbox_addr_GPIB = []
        self.cbox_instr_SRC = []
        self.cbox_instr_DVM = []
        self.cbox_instr_GMH = []

#        self.BuildComboChoices()

        self.GMH1Addr = self.GMH2Addr = 0  # invalid initial address as default

        self.ResourceList = []
        self.ComList = []
        self.GPIBList = []
        self.GPIBAddressList = ['addresses', 'GPIB0::0']  # Initial dummy vals
        self.COMAddressList = ['addresses', 'COM0']  # Initial dummy vals

        self.test_btns = []  # list of test buttons

        # Instruments
        src1_lbl = wx.StaticText(self, label='V1 source (SRC 1):', id=wx.ID_ANY)
        self.V1Sources = wx.ComboBox(self, wx.ID_ANY,
                                     choices=self.SRC_COMBO_CHOICE,
                                     size=(150, 10), style=wx.CB_DROPDOWN)
        self.V1Sources.Bind(wx.EVT_COMBOBOX, self.update_instr)
        self.cbox_instr_SRC.append(self.V1Sources)
        src2_lbl = wx.StaticText(self, label='V2 source (SRC 2):', id=wx.ID_ANY)
        self.V2Sources = wx.ComboBox(self, wx.ID_ANY,
                                     choices=self.SRC_COMBO_CHOICE,
                                     style=wx.CB_DROPDOWN)
        self.V2Sources.Bind(wx.EVT_COMBOBOX, self.update_instr)
        self.cbox_instr_SRC.append(self.V2Sources)
        dvm_v1v2_lbl = wx.StaticText(self, label='V1,V2 DVM (DVM12):',
                                     id=wx.ID_ANY)
        self.v1v2_dvms = wx.ComboBox(self, wx.ID_ANY,
                                     choices=self.DVM_COMBO_CHOICE,
                                     style=wx.CB_DROPDOWN)
        self.v1v2_dvms.Bind(wx.EVT_COMBOBOX, self.update_instr)
        self.cbox_instr_DVM.append(self.v1v2_dvms)
        dvm_vd_lbl = wx.StaticText(self, label='Vd DVM (DVMd):', id=wx.ID_ANY)
        self.VdDvms = wx.ComboBox(self, wx.ID_ANY,
                                  choices=self.DVM_COMBO_CHOICE,
                                  style=wx.CB_DROPDOWN)
        self.VdDvms.Bind(wx.EVT_COMBOBOX, self.update_instr)
        self.cbox_instr_DVM.append(self.VdDvms)
        T1_dvm_lbl = wx.StaticText(self, label='R1 T-probe DVM (DVMT1):',
                                   id=wx.ID_ANY)
        self.T1Dvms = wx.ComboBox(self, wx.ID_ANY,
                                  choices=self.DVM_COMBO_CHOICE,
                                  style=wx.CB_DROPDOWN)
        self.T1Dvms.Bind(wx.EVT_COMBOBOX, self.update_instr)
        self.cbox_instr_DVM.append(self.T1Dvms)
        T2_dvm_lbl = wx.StaticText(self, label='R2 T-probe DVM (DVMT2):',
                                   id=wx.ID_ANY)
        self.T2Dvms = wx.ComboBox(self, wx.ID_ANY,
                                  choices=self.DVM_COMBO_CHOICE,
                                  style=wx.CB_DROPDOWN)
        self.T2Dvms.Bind(wx.EVT_COMBOBOX, self.update_instr)
        self.cbox_instr_DVM.append(self.T2Dvms)

        gmh1_lbl = wx.StaticText(self, label='R1 GMH probe (GMH1):',
                                id=wx.ID_ANY)
        self.GMH1Probes = wx.ComboBox(self, wx.ID_ANY,
                                      choices=self.GMH_COMBO_CHOICE,
                                      style=wx.CB_DROPDOWN)
        self.GMH1Probes.Bind(wx.EVT_COMBOBOX, self.build_comment_str)
        self.cbox_instr_GMH.append(self.GMH1Probes)
        gmh2_lbl = wx.StaticText(self, label='R2 GMH probe (GMH2):',
                                id=wx.ID_ANY)
        self.GMH2Probes = wx.ComboBox(self, wx.ID_ANY,
                                      choices=self.GMH_COMBO_CHOICE,
                                      style=wx.CB_DROPDOWN)
        self.GMH2Probes.Bind(wx.EVT_COMBOBOX, self.build_comment_str)
        self.cbox_instr_GMH.append(self.GMH2Probes)

        gmh_room_lbl = wx.StaticText(self,
                                     label='Room conds. GMH probe (GMHroom):',
                                     id=wx.ID_ANY)
        self.GMHroomProbes = wx.ComboBox(self, wx.ID_ANY,
                                         choices=self.GMH_COMBO_CHOICE,
                                         style=wx.CB_DROPDOWN)
        self.GMHroomProbes.Bind(wx.EVT_COMBOBOX, self.update_instr)
        self.cbox_instr_GMH.append(self.GMHroomProbes)

        switchbox_lbl = wx.StaticText(self, label='Switchbox configuration:',
                                      id=wx.ID_ANY)
        self.Switchbox = wx.ComboBox(self, wx.ID_ANY,
                                     choices=self.SB_COMBO_CHOICE,
                                     style=wx.CB_DROPDOWN)
        self.Switchbox.Bind(wx.EVT_COMBOBOX, self.update_instr)

        # Addresses
        self.V1SrcAddr = wx.ComboBox(self, wx.ID_ANY,
                                     choices=self.GPIBAddressList,
                                     size=(150, 10), style=wx.CB_DROPDOWN)
        self.cbox_addr_GPIB.append(self.V1SrcAddr)
        self.V1SrcAddr.Bind(wx.EVT_COMBOBOX, self.update_addr)
        self.V2SrcAddr = wx.ComboBox(self, wx.ID_ANY,
                                     choices=self.GPIBAddressList,
                                     style=wx.CB_DROPDOWN)
        self.cbox_addr_GPIB.append(self.V2SrcAddr)
        self.V2SrcAddr.Bind(wx.EVT_COMBOBOX, self.update_addr)
        self.V1V2DvmAddr = wx.ComboBox(self, wx.ID_ANY,
                                       choices=self.GPIBAddressList,
                                       style=wx.CB_DROPDOWN)
        self.cbox_addr_GPIB.append(self.V1V2DvmAddr)
        self.V1V2DvmAddr.Bind(wx.EVT_COMBOBOX, self.update_addr)
        self.VdDvmAddr = wx.ComboBox(self, wx.ID_ANY,
                                     choices=self.GPIBAddressList,
                                     style=wx.CB_DROPDOWN)
        self.cbox_addr_GPIB.append(self.VdDvmAddr)
        self.VdDvmAddr.Bind(wx.EVT_COMBOBOX, self.update_addr)
        self.T1DvmAddr = wx.ComboBox(self, wx.ID_ANY,
                                     choices=self.GPIBAddressList,
                                     style=wx.CB_DROPDOWN)
        self.cbox_addr_GPIB.append(self.T1DvmAddr)
        self.T1DvmAddr.Bind(wx.EVT_COMBOBOX, self.update_addr)
        self.T2DvmAddr = wx.ComboBox(self, wx.ID_ANY,
                                     choices=self.GPIBAddressList,
                                     style=wx.CB_DROPDOWN)
        self.cbox_addr_GPIB.append(self.T2DvmAddr)
        self.T2DvmAddr.Bind(wx.EVT_COMBOBOX, self.update_addr)

        self.GMH1Ports = wx.ComboBox(self, wx.ID_ANY,
                                     choices=self.COMAddressList,
                                     style=wx.CB_DROPDOWN)
        self.cbox_addr_COM.append(self.GMH1Ports)
        self.GMH1Ports.Bind(wx.EVT_COMBOBOX, self.update_addr)
        self.GMH2Ports = wx.ComboBox(self, wx.ID_ANY,
                                     choices=self.COMAddressList,
                                     style=wx.CB_DROPDOWN)
        self.cbox_addr_COM.append(self.GMH2Ports)
        self.GMH2Ports.Bind(wx.EVT_COMBOBOX, self.update_addr)

        self.GMHroomPorts = wx.ComboBox(self, wx.ID_ANY,
                                        choices=self.COMAddressList,
                                        style=wx.CB_DROPDOWN)
        self.cbox_addr_COM.append(self.GMHroomPorts)
        self.GMHroomPorts.Bind(wx.EVT_COMBOBOX, self.update_addr)

        self.SwitchboxAddr = wx.ComboBox(self, wx.ID_ANY,
                                         choices=self.COMAddressList,
                                         style=wx.CB_DROPDOWN)
        self.cbox_addr_COM.append(self.SwitchboxAddr)

        # Filename
        file_lbl = wx.StaticText(self, label='Excel file full path:',
                                 id=wx.ID_ANY)
        self.XLFile = wx.TextCtrl(self, id=wx.ID_ANY,
                                  value=self.GetTopLevelParent().excelpath)

        # Resistors
        self.R1Name = wx.TextCtrl(self, id=wx.ID_ANY, value='R1 Name')
        self.R1Name.Bind(wx.EVT_TEXT, self.build_comment_str)
        R_names_lbl = wx.StaticText(self, label='Resistor Names:', id=wx.ID_ANY)
        self.R2Name = wx.TextCtrl(self, id=wx.ID_ANY, value='R2 Name')
        self.R2Name.Bind(wx.EVT_TEXT, self.build_comment_str)

        # Autopopulate btn
        self.AutoPop = wx.Button(self, id=wx.ID_ANY, label='AutoPopulate')
        self.AutoPop.Bind(wx.EVT_BUTTON, self.on_autopop)

        # Test buttons
        self.VisaList = wx.Button(self, id=wx.ID_ANY, label='List Visa res')
        self.VisaList.Bind(wx.EVT_BUTTON, self.on_visa_list)
        self.ResList = wx.TextCtrl(self, id=wx.ID_ANY,
                                   value='Available Visa resources',
                                   style=wx.TE_READONLY | wx.TE_MULTILINE)
        self.S1Test = wx.Button(self, id=wx.ID_ANY, label='Test')
        self.S1Test.Bind(wx.EVT_BUTTON, self.on_test)
        self.S2Test = wx.Button(self, id=wx.ID_ANY, label='Test')
        self.S2Test.Bind(wx.EVT_BUTTON, self.on_test)
        self.D12Test = wx.Button(self, id=wx.ID_ANY, label='Test')
        self.D12Test.Bind(wx.EVT_BUTTON, self.on_test)
        self.DdTest = wx.Button(self, id=wx.ID_ANY, label='Test')
        self.DdTest.Bind(wx.EVT_BUTTON, self.on_test)
        self.DT1Test = wx.Button(self, id=wx.ID_ANY, label='Test')
        self.DT1Test.Bind(wx.EVT_BUTTON, self.on_test)
        self.DT2Test = wx.Button(self, id=wx.ID_ANY, label='Test')
        self.DT2Test.Bind(wx.EVT_BUTTON, self.on_test)

        self.GMH1Test = wx.Button(self, id=wx.ID_ANY, label='Test')
        self.GMH1Test.Bind(wx.EVT_BUTTON, self.on_test)
        self.GMH2Test = wx.Button(self, id=wx.ID_ANY, label='Test')
        self.GMH2Test.Bind(wx.EVT_BUTTON, self.on_test)

        self.GMHroomTest = wx.Button(self, id=wx.ID_ANY, label='Test')
        self.GMHroomTest.Bind(wx.EVT_BUTTON, self.on_test)

        self.SwitchboxTest = wx.Button(self, id=wx.ID_ANY, label='Test')
        self.SwitchboxTest.Bind(wx.EVT_BUTTON, self.on_switch_test)

        response_lbl = wx.StaticText(self, label='Instrument Test Response:',
                                    id=wx.ID_ANY)
        self.Response = wx.TextCtrl(self, id=wx.ID_ANY, value='',
                                    style=wx.TE_READONLY)
        gb_sizer = wx.GridBagSizer()

        # Instruments
        gb_sizer.Add(src1_lbl, pos=(0, 0), span=(1, 1), flag=wx.ALL | wx.EXPAND,
                    border=5)
        gb_sizer.Add(self.V1Sources, pos=(0, 1), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(src2_lbl, pos=(1, 0), span=(1, 1), flag=wx.ALL | wx.EXPAND,
                    border=5)
        gb_sizer.Add(self.V2Sources, pos=(1, 1), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(dvm_v1v2_lbl, pos=(2, 0), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.v1v2_dvms, pos=(2, 1), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(dvm_vd_lbl, pos=(3, 0), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.VdDvms, pos=(3, 1), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(T1_dvm_lbl, pos=(4, 0), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.T1Dvms, pos=(4, 1), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(T2_dvm_lbl, pos=(5, 0), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.T2Dvms, pos=(5, 1), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(gmh1_lbl, pos=(6, 0), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.GMH1Probes, pos=(6, 1), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(gmh2_lbl, pos=(7, 0), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.GMH2Probes, pos=(7, 1), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(gmh_room_lbl, pos=(8, 0), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.GMHroomProbes, pos=(8, 1), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(switchbox_lbl, pos=(9, 0), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.Switchbox, pos=(9, 1), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        # Addresses
        gb_sizer.Add(self.V1SrcAddr, pos=(0, 2), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.V2SrcAddr, pos=(1, 2), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.V1V2DvmAddr, pos=(2, 2), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.VdDvmAddr, pos=(3, 2), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.T1DvmAddr, pos=(4, 2), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.T2DvmAddr, pos=(5, 2), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.GMH1Ports, pos=(6, 2), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.GMH2Ports, pos=(7, 2), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.GMHroomPorts, pos=(8, 2), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.SwitchboxAddr, pos=(9, 2), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        # R Name
        gb_sizer.Add(self.R1Name, pos=(6, 4), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.R2Name, pos=(7, 4), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(R_names_lbl, pos=(5, 4), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        # Filename
        gb_sizer.Add(file_lbl, pos=(11, 0), span=(1, 1), flag=wx.ALL | wx.EXPAND,
                    border=5)
        gb_sizer.Add(self.XLFile, pos=(11, 1), span=(1, 5),
                    flag=wx.ALL | wx.EXPAND, border=5)
        # Test buttons
        gb_sizer.Add(self.S1Test, pos=(0, 3), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.S2Test, pos=(1, 3), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.D12Test, pos=(2, 3), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.DdTest, pos=(3, 3), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.DT1Test, pos=(4, 3), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.DT2Test, pos=(5, 3), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.GMH1Test, pos=(6, 3), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.GMH2Test, pos=(7, 3), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.GMHroomTest, pos=(8, 3), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.SwitchboxTest, pos=(9, 3), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)

        gb_sizer.Add(response_lbl, pos=(3, 4), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.Response, pos=(4, 4), span=(1, 3),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.VisaList, pos=(0, 5), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.ResList, pos=(0, 4), span=(3, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)

        # Autopopulate btn
        gb_sizer.Add(self.AutoPop, pos=(2, 5), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)

        self.SetSizerAndFit(gb_sizer)

        # Roles and corresponding comboboxes/test btns are associated here:
        devices.ROLES_WIDGETS = {'SRC1': {'icb': self.V1Sources,
                                          'acb': self.V1SrcAddr,
                                          'tbtn': self.S1Test}}
        devices.ROLES_WIDGETS.update({'SRC2': {'icb': self.V2Sources,
                                               'acb': self.V2SrcAddr,
                                               'tbtn': self.S2Test}})
        devices.ROLES_WIDGETS.update({'DVM12': {'icb': self.v1v2_dvms,
                                                'acb': self.V1V2DvmAddr,
                                                'tbtn': self.D12Test}})
        devices.ROLES_WIDGETS.update({'DVMd': {'icb': self.VdDvms,
                                               'acb': self.VdDvmAddr,
                                               'tbtn': self.DdTest}})
        devices.ROLES_WIDGETS.update({'DVMT1': {'icb': self.T1Dvms,
                                                'acb': self.T1DvmAddr,
                                                'tbtn': self.DT1Test}})
        devices.ROLES_WIDGETS.update({'DVMT2': {'icb': self.T2Dvms,
                                                'acb': self.T2DvmAddr,
                                                'tbtn': self.DT2Test}})
        devices.ROLES_WIDGETS.update({'GMH1': {'icb': self.GMH1Probes,
                                               'acb': self.GMH1Ports,
                                               'tbtn': self.GMH1Test}})
        devices.ROLES_WIDGETS.update({'GMH2': {'icb': self.GMH2Probes,
                                               'acb': self.GMH2Ports,
                                               'tbtn': self.GMH2Test}})
        devices.ROLES_WIDGETS.update({'GMHroom': {'icb': self.GMHroomProbes,
                                                  'acb': self.GMHroomPorts,
                                                  'tbtn': self.GMHroomTest}})
        devices.ROLES_WIDGETS.update({'switchbox': {'icb': self.Switchbox,
                                                    'acb': self.SwitchboxAddr,
                                                    'tbtn': self.SwitchboxTest}})

        self.log = None
        self.wb = None
        self.ws_params = None
        self.instrument_choice = {}
        self.res_addr_list = []

    def build_combo_choices(self):
        """
        Populate combobox choices from known list of available instruments.
        Called from update_filepath().
        """
        for d in devices.INSTR_DATA.keys():
            if 'SRC:' in d:
                self.SRC_COMBO_CHOICE.append(d)
            elif 'DVM:' in d:
                self.DVM_COMBO_CHOICE.append(d)
            elif 'GMH:' in d:
                self.GMH_COMBO_CHOICE.append(d)

        # Strip redundant entries:
        self.SRC_COMBO_CHOICE = list(set(self.SRC_COMBO_CHOICE))
        self.DVM_COMBO_CHOICE = list(set(self.DVM_COMBO_CHOICE))
        self.GMH_COMBO_CHOICE = list(set(self.GMH_COMBO_CHOICE))

        # Re-build combobox choices from list of SRC's
        for cbox in self.cbox_instr_SRC:
            cbox.Clear()
            cbox.AppendItems(self.SRC_COMBO_CHOICE)

        # Re-build combobox choices from list of DVM's
        for cbox in self.cbox_instr_DVM:
            cbox.Clear()
            cbox.AppendItems(self.DVM_COMBO_CHOICE)

        # Re-build combobox choices from list of GMH's
        for cbox in self.cbox_instr_GMH:
            cbox.Clear()
            cbox.AppendItems(self.GMH_COMBO_CHOICE)

    def update_filepath(self, e):
        """
        Called when a new Excel file has been selected.
        """
        self.XLFile.SetValue(e.XLpath)

        # Open logfile
        logname = 'HRBCv'+str(e.v)+'_'+str(dt.date.today())+'.log'
        logfile = os.path.join(e.d, logname)
        self.log = open(logfile, 'a')

        # Read parameters sheet - gather instrument info:
        # Need cell VALUE, not FORMULA, so set data_only = True
        self.wb = load_workbook(self.XLFile.GetValue(), data_only=True)
        self.ws_params = self.wb.get_sheet_by_name('Parameters')

        headings = (None, u'description', u'Instrument Info:', u'parameter',
                    u'value', u'uncert', u'dof', u'label')

        # Determine colummn indices from column letters:
        col_I = utils.cell.column_index_from_string('I') - 1
        col_J = utils.cell.column_index_from_string('J') - 1
        col_K = utils.cell.column_index_from_string('K') - 1
        col_L = utils.cell.column_index_from_string('L') - 1
        col_M = utils.cell.column_index_from_string('M') - 1
        col_N = utils.cell.column_index_from_string('N') - 1

        params = []
        values = []

        for r in self.ws_params.rows:  # a tuple of row objects
            descr = r[col_I].value  # cell.value
            param = r[col_J].value  # cell.value
            # 'v_u_d_l' is short for: 'value, uncert, dof, label':
            v_u_d_l = [r[col_K].value, r[col_L].value, r[col_M].value,
                       r[col_N].value]

            if descr in headings and param in headings:
                continue  # Skip this row
            else:  # not header
                params.append(param)
                if v_u_d_l[1] is None:  # single-valued (no uncert)
                    values.append(v_u_d_l[0])  # append value as next item
                    print('{} : {} = {}'.format(descr, param, v_u_d_l[0]))
                    print('{} : {} = {}'.format(descr, param, v_u_d_l[0]), file=self.log)
                else:  # multi-valued
                    while v_u_d_l[-1] is None:  # remove empty cells
                        del v_u_d_l[-1]  # v_u_d_l.pop()
                    values.append(v_u_d_l)  # append value-list as next item
                    print('{} : {} = {}'.format(descr, param, v_u_d_l))
                    print('{} : {} = {}'.format(descr, param, v_u_d_l), file=self.log)  # self.log.write(logline)

                if param == u'test':  # last parameter for this description
                    devices.DESCR.append(descr)  # build description list
                    devices.sublist.append(dict(zip(params,values)))  # adds parameter dictionary to sublist
                    del params[:]
                    del values[:]

        print('----END OF PARAMETER LIST----')
        print('----END OF PARAMETER LIST----', file=self.log)

        # Compile into a dictionary that lives in devices.py...  
        devices.INSTR_DATA = dict(zip(devices.DESCR, devices.sublist))
        self.build_combo_choices()

    def on_autopop(self, e):
        """
        Pre-select instrument and address comboboxes -
        Choose from instrument descriptions listed in devices.DESCR
        (Uses address assignments in devices.INSTR_DATA)
        """
        self.instrument_choice = {'SRC1': 'SRC: D4808',
                                  'SRC2': 'SRC: F5520A',
                                  'DVM12': 'DVM: HP3458A, s/n452',
                                  'DVMd': 'DVM: HP3458A, s/n382',
                                  'DVMT1': 'none',  # 'DVM: HP34401A, s/n976'
                                  'DVMT2': 'none',  # 'DVM: HP34420A, s/n130'
                                  'GMH1': 'GMH: s/n627',
                                  'GMH2': 'GMH: s/n628',
                                  'GMHroom': 'GMH: s/n367',
                                  'switchbox': 'V1'}
        for r in self.instrument_choice.keys():
            d = self.instrument_choice[r]
            devices.ROLES_WIDGETS[r]['icb'].SetValue(d)  # Update i_cb
            self.create_instr(d, r)
        if 'Name' in self.R1Name.GetValue(): 
            self.R1Name.SetValue('CHANGE_THIS! 1G')
        if 'Name' in self.R2Name.GetValue():
            self.R2Name.SetValue('CHANGE_THIS! 1M')

    def update_instr(self, e):
        """
        An instrument was selected for a role.
        Find description d and role r, then pass to create_instr()
        """
        d = e.GetString()
        r = ''
        for r in devices.ROLES_WIDGETS.keys():  # Cycle through roles
            if devices.ROLES_WIDGETS[r]['icb'] == e.GetEventObject():
                break  # stop looking when find the right instrument & role
        self.create_instr(d, r)

    def create_instr(self, d, r):
        """
        Called by both on_autopop() and update_instr()
        Create each instrument in software & open visa session (for GPIB)
        For GMH instruments, use GMH dll not visa.
        """
        if 'GMH' in r:  # Changed from d to r
            # create and open a GMH instrument instance
            print('\nnbpages.SetupPage.CreateInstr(): Creating GMH device ({} -> {}).'.format(d, r))
            devices.ROLES_INSTR.update({r: devices.GMHDevice(d)})
        else:
            # create a visa instrument instance
            print('\nnbpages.SetupPage.CreateInstr(): Creating VISA device ({} -> {}).'.format(d, r))
            devices.ROLES_INSTR.update({r: devices.Instrument(d)})
            devices.ROLES_INSTR[r].open()
        self.set_instr(d, r)

    @staticmethod
    def set_instr(d, r):
        """
        Called by create_instr().
        Updates internal info (INSTR_DATA) and Enables/disables testbuttons
        as necessary.
        """
        # print 'nbpages.SetupPage.SetInstr():',d,'assigned to role',r,'demo mode:',devices.ROLES_INSTR[r].demo
        assert d in devices.INSTR_DATA, 'Unknown instrument: {} - check Excel file is loaded.'.format(d)
        assert 'role' in devices.INSTR_DATA[d],\
            'Unknown instrument parameter - check Excel Parameters sheet is populated.'
        devices.INSTR_DATA[d]['role'] = r  # update default role
        
        # Set the address cb to correct value (according to devices.INSTR_DATA)
        a_cb = devices.ROLES_WIDGETS[r]['acb']
        a_cb.SetValue((devices.INSTR_DATA[d]['str_addr']))
        if d == 'none':
            devices.ROLES_WIDGETS[r]['tbtn'].Enable(False)
        else:
            devices.ROLES_WIDGETS[r]['tbtn'].Enable(True)

    def update_addr(self, e):
        """
        An address was manually selected, so change INSTR_DATA...
            1st, we'll need instrument description d...
        """
        d = 'none'
        r = ''
        addr = 0
        acb = e.GetEventObject()  # 'a'ddress 'c'ombo 'b'ox
        for r in devices.ROLES_WIDGETS.keys():
            if devices.ROLES_WIDGETS[r]['acb'] == acb:
                d = devices.ROLES_WIDGETS[r]['icb'].GetValue()
                break  # stop looking when find the right instrument descr
        a = e.GetString()  # address string, eg 'COM5' or 'GPIB0::23'
        if (a not in self.GPIBAddressList) or (a not in self.COMAddressList):  # Ignore dummy values, like 'NO_ADDRESS'
            devices.INSTR_DATA[d]['str_addr'] = a
            devices.ROLES_INSTR[r].str_addr = a
            addr = int(a.lstrip('COMGPIB0:'))  # leave only numeric part of address string
            devices.INSTR_DATA[d]['addr'] = addr
            devices.ROLES_INSTR[r].addr = addr
        print('update_addr(): {} using {} set to addr {} ({})'.format(r, d, addr, a))

    def on_test(self, e):
        """Called when a 'test' button is clicked"""
        d = 'none'
        r = ''
        for r in devices.ROLES_WIDGETS.keys():  # check every role
            if devices.ROLES_WIDGETS[r]['tbtn'] == e.GetEventObject():
                d = devices.ROLES_WIDGETS[r]['icb'].GetValue()
                break  # stop looking when find the right instrument descr
        print('\nnbpages.SetupPage.OnTest():', d)
        assert 'test' in devices.INSTR_DATA[d], 'No test exists for this device.'
        test_str = devices.INSTR_DATA[d]['test']  # test string
        print('\tTest string:', test_str)
        self.Response.SetValue(str(devices.ROLES_INSTR[r].test(test_str)))
        self.status.SetStatusText('Testing %s with cmd %s' % (d, test_str), 0)

    def on_switch_test(self, e):
        resource = self.SwitchboxAddr.GetValue()
        config = str(devices.SWITCH_CONFIGS[self.Switchbox.GetValue()])
        try:
            instr = devices.RM.open_resource(resource)
            instr.write(config)
        except devices.visa.VisaIOError:
            self.Response.SetValue('Couldn\'t open visa resource for switchbox!')

    def build_comment_str(self, e):
        """Called by a change in GMH probe selection, or resistor name"""
        d = e.GetString()
        r = ''
        if 'GMH' in d:  # A GMH probe selection changed
            # Find the role associated with the selected instrument description
            for r in devices.ROLES_WIDGETS.keys():
                if devices.ROLES_WIDGETS[r]['icb'].GetValue() == d:
                    break
            # Update our knowledge of role <-> instr. descr. association
            self.create_instr(d, r)
        run_page = self.GetParent().GetPage(1)
        params={'R1': self.R1Name.GetValue(), 'TR1': self.GMH1Probes.GetValue(),
                'R2': self.R2Name.GetValue(), 'TR2': self.GMH2Probes.GetValue()}
        joinstr = ' monitored by '
        commstr = 'R1: ' + params['R1'] + joinstr + params['TR1'] + '. R2: ' + params['R2'] + joinstr + params['TR2']
        evt = evts.UpdateCommentEvent(str=commstr)
        wx.PostEvent(run_page, evt)

    def on_visa_list(self, e):
        res_list = devices.RM.list_resources()
        del self.ResourceList[:]  # list of COM ports & GPIB addresses
        del self.ComList[:]  # list of COM ports (numbers only)
        del self.GPIBList[:]  # list of GPIB addresses (numbers only)
        for item in res_list:
            self.ResourceList.append(item.replace('ASRL', 'COM'))
        for item in self.ResourceList:
            addr = item.replace('::INSTR', '')
            if 'COM' in item:
                self.ComList.append(addr)
            elif 'GPIB' in item:
                self.GPIBList.append(addr)

        # Re-build combobox choices from list of COM ports
        for cbox in self.cbox_addr_COM:
            cbox.Clear()
            cbox.AppendItems(self.ComList)

        # Re-build combobox choices from list of GPIB addresses
        for cbox in self.cbox_addr_GPIB:
            cbox.Clear()
            cbox.AppendItems(self.GPIBList)

        # Add resources to ResList TextCtrl widget
        self.res_addr_list = '\n'.join(self.ResourceList)
        self.ResList.SetValue(self.res_addr_list)

'''
____________________________________________
#-------------- End of Setup Page -----------
____________________________________________
'''
'''
----------------------
# Run Page definition:
----------------------
'''


class RunPage(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)

        self.status = self.GetTopLevelParent().sb
        self.version = self.GetTopLevelParent().version
        self.run_id = 'none'

        # Event bindings
        self.Bind(evts.EVT_UPDATE_COM_STR, self.update_comment)
        self.Bind(evts.EVT_DATA, self.update_data)
        self.Bind(evts.EVT_DELAYS, self.update_dels)
        self.Bind(evts.EVT_START_ROW, self.update_start_row)
        self.Bind(evts.EVT_STOP_ROW, self.update_stop_row)

        self.RunThread = None
        self.RLinkThread = None

        # Comment widgets
        comment_lbl = wx.StaticText(self, id=wx.ID_ANY, label='Comment:')
        self.Comment = wx.TextCtrl(self, id=wx.ID_ANY, size=(600, 20))
        self.Comment.Bind(wx.EVT_TEXT, self.on_comment)
        comtip_a = 'This is auto-generated from data on the Setup page.'
        comtip_b = ' Other notes may be added manually.'
        self.Comment.SetToolTip(comtip_a + comtip_b)  # changed from deprecated SetToolTipString()

        self.NewRunIDBtn = wx.Button(self, id=wx.ID_ANY,
                                     label='Create new run id')
        idcomtip = 'New id used to link subsequent Rlink and measurement data.'
        self.NewRunIDBtn.SetToolTip(idcomtip)  # changed from deprecated SetToolTipString()
        self.NewRunIDBtn.Bind(wx.EVT_BUTTON, self.on_new_run_id)
        self.RunID = wx.TextCtrl(self, id=wx.ID_ANY, size=(500, 20))

        # Voltage source widgets
        v1_src_lbl = wx.StaticText(self, id=wx.ID_ANY, style=wx.ALIGN_RIGHT,
                                 label='Set V1:')
        self.V1Setting = NumCtrl(self, id=wx.ID_ANY, integerWidth=3,
                                 fractionWidth=8, groupDigits=True)
        self.V1Setting.Bind(wx.lib.masked.EVT_NUM, self.on_v1_set)

        v2_src_lbl = wx.StaticText(self, id=wx.ID_ANY, style=wx.ALIGN_RIGHT,
                                 label='Set V2:')
        self.V2Setting = NumCtrl(self, id=wx.ID_ANY, integerWidth=3,
                                 fractionWidth=8, groupDigits=True)
        self.V2Setting.Bind(wx.lib.masked.EVT_NUM, self.on_v2_set)

        zero_volts_btn = wx.Button(self, id=wx.ID_ANY, label='Set zero volts')
        zero_volts_btn.Bind(wx.EVT_BUTTON, self.on_zero_volts)

        self.RangeTBtn = wx.ToggleButton(self, id=wx.ID_ANY,
                                         label='DVM12 Range mode')
        self.RangeTBtn.Bind(wx.EVT_TOGGLEBUTTON, self.on_range_mode)

        # Delay widgets
        settle_del_lbl = wx.StaticText(self, id=wx.ID_ANY, label='Settle delay:')
        self.SettleDel = wx.SpinCtrl(self, id=wx.ID_ANY, value='0',
                                     min=0, max=600)
        start_del_lbl = wx.StaticText(self, id=wx.ID_ANY, label='Start delay:')
        self.StartDel = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        azero1_del_lbl = wx.StaticText(self, id=wx.ID_ANY,
                                     label='AZERO_ONCE delay:')
        self.AZERO1Del = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        range_del_lbl = wx.StaticText(self, id=wx.ID_ANY, label='Range delay:')
        self.RangeDel = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        n_samples_lbl = wx.StaticText(self, id=wx.ID_ANY,
                                    label='Number of samples:')
        self.NSamples = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)

        #  Run control and progress widgets
        self.StartRow = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        start_row_lbl = wx.StaticText(self, id=wx.ID_ANY, label='Start row:')
        self.StopRow = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        stop_row_lbl = wx.StaticText(self, id=wx.ID_ANY, label='Stop row:')
        row_lbl = wx.StaticText(self, id=wx.ID_ANY, label='Current row:')
        self.Row = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        time_lbl = wx.StaticText(self, id=wx.ID_ANY, label='Timestamp:')
        self.Time = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)

        v_av_lbl = wx.StaticText(self, id=wx.ID_ANY, label='Mean voltage(V):')
        # self.Vav = wx.TextCtrl(self, id = wx.ID_ANY, style = wx.TE_READONLY)
        self.Vav = NumCtrl(self, id=wx.ID_ANY, integerWidth=3, fractionWidth=9,
                           groupDigits=True)
        v_sd_lbl = wx.StaticText(self, id=wx.ID_ANY, label='Stdev(voltage):')
        # self.Vsd = wx.TextCtrl(self, id = wx.ID_ANY, style = wx.TE_READONLY)
        self.Vsd = NumCtrl(self, id=wx.ID_ANY, integerWidth=3, fractionWidth=9,
                           groupDigits=True)

        self.StartBtn = wx.Button(self, id=wx.ID_ANY, label='Start run')
        self.StartBtn.Bind(wx.EVT_BUTTON, self.on_start)
        self.StopBtn = wx.Button(self, id=wx.ID_ANY, label='Abort run')
        self.StopBtn.Bind(wx.EVT_BUTTON, self.on_abort)
        self.StopBtn.Enable(False)
        self.RLinkBtn = wx.Button(self, id=wx.ID_ANY, label='Measure R-link')
        self.RLinkBtn.Bind(wx.EVT_BUTTON, self.on_rlink)

        progress_lbl = wx.StaticText(self, id=wx.ID_ANY, style=wx.ALIGN_RIGHT,
                                    label='Run progress:')
        self.Progress = wx.Gauge(self, id=wx.ID_ANY, range=100,
                                 name='Progress')

        gb_sizer = wx.GridBagSizer()

        # Comment widgets
        gb_sizer.Add(comment_lbl,pos=(0, 0), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.Comment, pos=(0, 1), span=(1, 6),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.NewRunIDBtn, pos=(1, 0), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.RunID, pos=(1, 1), span=(1, 6),
                    flag=wx.ALL | wx.EXPAND, border=5)
        # gbSizer.Add(self.h_sep1, pos=(2,0), span=(1,5), flag=wx.ALL|wx.EXPAND, border=5)

        # Voltage source widgets
        gb_sizer.Add(zero_volts_btn, pos=(2, 0), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(v1_src_lbl, pos=(2, 1), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.V1Setting, pos=(2, 2), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(v2_src_lbl, pos=(2, 3), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.V2Setting, pos=(2, 4), span=(1, 1),
                    flag=wx.ALL, border=5)
        gb_sizer.Add(self.RangeTBtn, pos=(2, 5), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)

        # Delay widgets
        gb_sizer.Add(settle_del_lbl, pos=(3, 0), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.SettleDel, pos=(4, 0), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(start_del_lbl, pos=(3, 1), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.StartDel, pos=(4, 1), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(azero1_del_lbl, pos=(3, 2), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.AZERO1Del, pos=(4, 2), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(range_del_lbl, pos=(3, 3), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.RangeDel, pos=(4, 3), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)

        gb_sizer.Add(n_samples_lbl, pos=(3, 4), span=(1, 1),
                    flag=wx.ALL, border=5)
        gb_sizer.Add(self.NSamples, pos=(4, 4), span=(1, 1),
                    flag=wx.ALL, border=5)
        #gbSizer.Add(self.h_sep3, pos=(7,0), span=(1,5), flag=wx.ALL|wx.EXPAND, border=5)

        #  Run control and progress widgets
        gb_sizer.Add(start_row_lbl, pos=(5, 0), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.StartRow, pos=(6, 0), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(stop_row_lbl, pos=(5, 1), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.StopRow, pos=(6, 1), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(row_lbl, pos=(5, 2), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.Row, pos=(6, 2), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(time_lbl, pos=(5, 3), span=(1, 2),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.Time, pos=(6, 3), span=(1, 2),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(v_av_lbl, pos=(5, 5), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.Vav, pos=(6, 5), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(v_sd_lbl, pos=(5, 6), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.Vsd, pos=(6, 6), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)

        gb_sizer.Add(self.RLinkBtn, pos=(7, 0), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.StartBtn, pos=(7, 1), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.StopBtn, pos=(7, 2), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(progress_lbl, pos=(7, 3), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.Progress, pos=(7, 4), span=(1, 3),
                    flag=wx.ALL | wx.EXPAND, border=5)

        self.SetSizerAndFit(gb_sizer)

        self.autocomstr = ''
        self.manstr = ''
        self.fullstr = ''

    def on_range_mode(self, e):
        state = e.GetEventObject().GetValue()
        print('OnRangeMode(): Range toggle button value =', state)
        if state is True:
            e.GetEventObject().SetLabel("AUTO-range DVM12")
        else:
            e.GetEventObject().SetLabel("FIXED-range DVM12")

    def on_new_run_id(self, e):
        start = self.fullstr.find('R1: ')
        end = self.fullstr.find(' monitored', start)
        r1name = self.fullstr[start+4:end]
        start = self.fullstr.find('R2: ')
        end = self.fullstr.find(' monitored', start)
        r2name = self.fullstr[start+4:end]
        self.run_id = str('HRBC.v' + self.version + ' ' + r1name + ':' +
                          r2name + ' ' +
                          dt.datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
        self.status.SetStatusText('Id for subsequent runs:', 0)
        self.status.SetStatusText(str(self.run_id), 1)
        self.RunID.SetValue(str(self.run_id))

    def update_comment(self, e):
        # writes combined auto-comment and manual comment when
        # auto-generated comment is re-built
        self.autocomstr = e.str  # store a copy of auto-generated comment
        self.Comment.SetValue(e.str+self.manstr)

    def on_comment(self, e):
        # Called when comment emits EVT_TEXT (i.e. whenever it's changed)
        # Make sure comment field (with extra manually-entered notes)
        # isn't overwritten
        self.fullstr = self.Comment.GetValue()  # store a copy of full comment
        # Extract last part of comment (the manually-inserted bit)
        # - assume we manually added extra notes to END
        self.manstr = self.fullstr[len(self.autocomstr):]

    def update_data(self, e):
        # Triggered by an 'update data' event
        # event params:(t,Vm,Vsd,r,P,flag['1','2','d' or 'E'])
        if e.flag in 'EF':  # finished
            self.RunThread = None
            self.StartBtn.Enable(True)
            self.Progress.SetToolTip(str(0)+'%')  # Changed from deprecated SetToolTipString()
        else:
            self.Time.SetValue(str(e.t))
            self.Vav.SetValue(str(e.Vm))
            self.Vsd.SetValue(str(e.Vsd))
            self.Row.SetValue(str(e.r))
            self.Progress.SetValue(e.P)
            self.Progress.SetToolTip(str(e.P)+'%')  # Changed from deprecated SetToolTipString()

    def update_dels(self, e):
        # Triggered by an 'update delays' event
        self.StartDel.SetValue(str(e.s))
        self.NSamples.SetValue(str(e.n))
        self.AZERO1Del.SetValue(str(e.AZ1))
        self.RangeDel.SetValue(str(e.r))

    def update_start_row(self, e):
        # Triggered by an 'update startrow' event
        self.StartRow.SetValue(str(e.row))

    def update_stop_row(self, e):
        # Triggered by an 'update stoprow' event
        self.StopRow.SetValue(str(e.row))

    def on_v1_set(self, e):
        # Called by change in value (manually OR by software!)
        v1 = e.GetValue()
        src1 = devices.ROLES_INSTR['SRC1']
        src1.set_V(v1)  # 'M+0R0='
        time.sleep(0.5)
        if v1 == 0:
            src1.stby()
        else:
            src1.oper()
        time.sleep(0.5)

    def on_v2_set(self, e):
        # Called by change in value (manually OR by software!)
        v2 = e.GetValue()
        src2 = devices.ROLES_INSTR['SRC2']
        src2.set_V(v2)
        time.sleep(0.5)
        if v2 == 0:
            src2.stby()
        else:
            src2.oper()
        time.sleep(0.5)

    def on_zero_volts(self, e):
        # V1:
        src1 = devices.ROLES_INSTR['SRC1']
        if self.V1Setting.GetValue() == 0:
            print('RunPage.OnZeroVolts(): Zero/Stby directly (not via V1 display)')
            src1.set_V(0)
            src1.stby()
        else:
            self.V1Setting.SetValue('0')  # Calls OnV1Set() ONLY IF VALUE CHANGES
            print('RunPage.OnZeroVolts():  Zero/Stby via V1 display')

        # V2:
        src2 = devices.ROLES_INSTR['SRC2']
        if self.V2Setting.GetValue() == 0:
            print('RunPage.OnZeroVolts(): Zero/Stby directly (not via V2 display)')
            src2.set_V(0)
            src2.stby()
        else:
            self.V2Setting.SetValue('0')  # Calls OnV2Set() ONLY IF VALUE CHANGES
            print('RunPage.OnZeroVolts(): Zero/Stby via V2 display')

    def on_start(self, e):
        self.Progress.SetValue(0)
        self.RunThread = None
        self.status.SetStatusText('', 1)
        self.status.SetStatusText('Starting run', 0)
        if self.RunThread is None:
            self.StopBtn.Enable(True)  # Enable Stop button
            self.StartBtn.Enable(False)  # Disable Start button
            # start acquisition thread here
            self.RunThread = acq.AqnThread(self)

    def on_abort(self, e):
        self.StartBtn.Enable(True)
        self.StopBtn.Enable(False)  # Disable Stop button
        self.RLinkBtn.Enable(True)  # Enable Start button
        if self.RunThread:
            self.RunThread.abort()
        elif self.RLinkThread:
            self.RLinkThread.abort()

    def on_rlink(self, e):
        self.Progress.SetValue(0)
        self.RLinkThread = None
        self.status.SetStatusText('', 1)
        self.status.SetStatusText('Starting R-link measurement', 0)
        if self.RLinkThread is None:
            self.StopBtn.Enable(True)  # Enable Stop button
            self.RLinkBtn.Enable(False)
            self.RLinkThread = rl.RLThread(self)

'''
__________________________________________
#-------------- End of Run Page ----------
__________________________________________
'''
'''
-----------------------
# Plot Page definition:
-----------------------
'''


class PlotPage(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)

        self.Bind(evts.EVT_PLOT, self.update_plot)
        self.Bind(evts.EVT_CLEARPLOT, self.clear_plot)

        self.figure = Figure()

        # 0.3" height space between subplots:
        self.figure.subplots_adjust(hspace=0.3)

        self.Vdax = self.figure.add_subplot(3, 1, 3)  # 3high x 1wide, 3rd plot down 
        self.Vdax.ticklabel_format(style='sci', useOffset=False, axis='y',
                                   scilimits=(2, -2))  # Auto-centre on data
        # Use scientific notation for Vd y-axis:
        self.Vdax.yaxis.set_major_formatter(mtick.ScalarFormatter(useMathText=True, useOffset=False))
        self.Vdax.autoscale(enable=True, axis='y', tight=False)  # Autoscale with 'buffer' around data extents
        self.Vdax.set_xlabel('time')
        self.Vdax.set_ylabel('Vd')

        self.V1ax = self.figure.add_subplot(3, 1, 1, sharex=self.Vdax)  # 3high x 1wide, 1st plot down 
        self.V1ax.ticklabel_format(useOffset=False, axis='y')  # Auto offset to centre on data
        self.V1ax.autoscale(enable=True, axis='y', tight=False)  # Autoscale with 'buffer' around data extents
        plt.setp(self.V1ax.get_xticklabels(), visible=False)  # Hide x-axis labels
        self.V1ax.set_ylabel('V1')
        self.V1ax.set_ylim(auto=True)
        v1_y_offset = self.V1ax.get_xaxis().get_offset_text()
        v1_y_offset.set_visible(False)

        self.V2ax = self.figure.add_subplot(3, 1, 2, sharex=self.Vdax)  # 3high x 1wide, 2nd plot down 
        self.V2ax.ticklabel_format(useOffset=False, axis='y') # Auto offset to centre on data
        self.V2ax.autoscale(enable=True, axis='y', tight=False)  # Autoscale with 'buffer' around data extents
        plt.setp(self.V2ax.get_xticklabels(), visible=False)  # Hide x-axis labels
        self.V2ax.set_ylabel('V2')
        self.V2ax.set_ylim(auto=True)
        v2_y_offset = self.V2ax.get_xaxis().get_offset_text()
        v2_y_offset.set_visible(False)

        self.canvas = FigureCanvas(self, wx.ID_ANY, self.figure)
        self.sizer = wx.BoxSizer(wx.VERTICAL)
        self.sizer.Add(self.canvas, 1, wx.LEFT | wx.TOP | wx.GROW)
        self.SetSizerAndFit(self.sizer)

    def update_plot(self, e):
        # six event attributes: td, t1, t2 (list of n times),
        # and Vd, V1, V2 (list of n voltages) plus clear_plot flag
        self.V1ax.plot_date(e.t1, e.V1, 'bo')
        self.V2ax.plot_date(e.t2, e.V2, 'go')
        self.Vdax.plot_date(e.td, e.Vd, 'ro')
        self.figure.autofmt_xdate()  # default settings
        self.Vdax.fmt_xdata = mdates.DateFormatter('%d-%m-%Y, %H:%M:%S')
        self.canvas.draw()
        self.canvas.Refresh()

    def clear_plot(self, e):
        self.V1ax.cla()
        self.V2ax.cla()
        self.Vdax.cla()
        self.Vdax.set_ylabel('Vd')
        self.V1ax.set_ylabel('V1')
        self.V2ax.set_ylabel('V2')
        self.canvas.draw()
        self.canvas.Refresh()
