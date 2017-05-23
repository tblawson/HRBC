# -*- coding: utf-8 -*-
""" nbpages.py - Defines individual notebook pages as panel-like objects

DEVELOPMENT VERSION

Created on Tue Jun 30 10:10:16 2015

@author: t.lawson
"""

import os

import wx
from wx.lib.masked import NumCtrl
import datetime as dt
import time

import matplotlib
matplotlib.use('WXAgg') # Agg renderer for drawing on a wx canvas
from matplotlib.backends.backend_wxagg import FigureCanvasWxAgg as FigureCanvas
#from matplotlib.backends.backend_wx import NavigationToolbar2Wx
from matplotlib.figure import Figure
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib.ticker as mtick
from openpyxl import load_workbook, cell

import HighRes_events as evts
import acquisition as acq
import RLink as rl
import devices

matplotlib.rc('lines', linewidth=1, color='blue')

os.environ['XLPATH'] = 'C:\Documents and Settings\\t.lawson\My Documents\Python Scripts\High_Res_Bridge'
'''
------------------------
# Setup Page definition:
------------------------
'''

class SetupPage(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)
        
        # Event bindings
        self.Bind(evts.EVT_FILEPATH, self.UpdateFilepath)

        self.status = self.GetTopLevelParent().sb

        self.SRC_COMBO_CHOICE = ['none']
        self.DVM_COMBO_CHOICE = ['none']
        self.GMH_COMBO_CHOICE = ['none'] # devices.GMH_DESCR # ('GMH s/n628', 'GMH s/n627')
        self.SB_COMBO_CHOICE =  devices.SWITCH_CONFIGS.keys()
        self.T_SENSOR_CHOICE = devices.T_Sensors
        self.cbox_addr_COM = []
        self.cbox_addr_GPIB = []
        self.cbox_instr_SRC = []
        self.cbox_instr_DVM = []
        self.cbox_instr_GMH = []

        self.BuildComboChoices()

        self.GMH1Addr = self.GMH2Addr = 0 # invalid initial address as default

        self.ResourceList = []
        self.ComList = []
        self.GPIBList = []
        self.GPIBAddressList = ['addresses','GPIB0::0'] # dummy values for starters...
        self.COMAddressList = ['addresses','COM0'] # dummy values for starters...

        self.test_btns = [] # list of test buttons

        # Instruments
        Src1Lbl = wx.StaticText(self, label='V1 source (SRC 1):', id = wx.ID_ANY)
        self.V1Sources = wx.ComboBox(self,wx.ID_ANY, choices = self.SRC_COMBO_CHOICE, size = (150,10), style = wx.CB_DROPDOWN)
        self.V1Sources.Bind(wx.EVT_COMBOBOX, self.UpdateInstr)
        self.cbox_instr_SRC.append(self.V1Sources)
        Src2Lbl = wx.StaticText(self, label='V2 source (SRC 2):', id = wx.ID_ANY)
        self.V2Sources = wx.ComboBox(self,wx.ID_ANY, choices = self.SRC_COMBO_CHOICE, style=wx.CB_DROPDOWN)
        self.V2Sources.Bind(wx.EVT_COMBOBOX, self.UpdateInstr)
        self.cbox_instr_SRC.append(self.V2Sources)
        DVM_V1V2Lbl = wx.StaticText(self, label='V1,V2 DVM (DVM12):', id = wx.ID_ANY)
        self.V1V2Dvms = wx.ComboBox(self,wx.ID_ANY, choices = self.DVM_COMBO_CHOICE, style=wx.CB_DROPDOWN)
        self.V1V2Dvms.Bind(wx.EVT_COMBOBOX, self.UpdateInstr)
        self.cbox_instr_DVM.append(self.V1V2Dvms)
        DVM_VdLbl = wx.StaticText(self, label='Vd DVM (DVMd):', id = wx.ID_ANY)
        self.VdDvms = wx.ComboBox(self,wx.ID_ANY, choices = self.DVM_COMBO_CHOICE, style=wx.CB_DROPDOWN)
        self.VdDvms.Bind(wx.EVT_COMBOBOX, self.UpdateInstr)
        self.cbox_instr_DVM.append(self.VdDvms)
        T1DvmLbl = wx.StaticText(self, label='R1 T-probe DVM (DVMT1):', id = wx.ID_ANY)
        self.T1Dvms = wx.ComboBox(self,wx.ID_ANY, choices = self.DVM_COMBO_CHOICE, style=wx.CB_DROPDOWN)
        self.T1Dvms.Bind(wx.EVT_COMBOBOX, self.UpdateInstr)
        self.cbox_instr_DVM.append(self.T1Dvms)
        T2DvmLbl = wx.StaticText(self, label='R2 T-probe DVM (DVMT2):', id = wx.ID_ANY)
        self.T2Dvms = wx.ComboBox(self,wx.ID_ANY, choices = self.DVM_COMBO_CHOICE, style=wx.CB_DROPDOWN)
        self.T2Dvms.Bind(wx.EVT_COMBOBOX, self.UpdateInstr)
        self.cbox_instr_DVM.append(self.T2Dvms)
        
        GMH1Lbl = wx.StaticText(self, label='R1 GMH probe (GMH1):', id = wx.ID_ANY)
        self.GMH1Probes = wx.ComboBox(self,wx.ID_ANY, choices = self.GMH_COMBO_CHOICE, style=wx.CB_DROPDOWN)
        self.GMH1Probes.Bind(wx.EVT_COMBOBOX, self.BuildCommStr)
        self.cbox_instr_GMH.append(self.GMH1Probes)
        GMH2Lbl = wx.StaticText(self, label='R2 GMH probe (GMH2):', id = wx.ID_ANY)
        self.GMH2Probes = wx.ComboBox(self,wx.ID_ANY, choices = self.GMH_COMBO_CHOICE, style=wx.CB_DROPDOWN)
        self.GMH2Probes.Bind(wx.EVT_COMBOBOX, self.BuildCommStr)
        self.cbox_instr_GMH.append(self.GMH2Probes)
        
        GMHroomLbl = wx.StaticText(self, label='Room conds. GMH probe (GMHroom):', id = wx.ID_ANY)
        self.GMHroomProbes = wx.ComboBox(self,wx.ID_ANY, choices = self.GMH_COMBO_CHOICE, style=wx.CB_DROPDOWN)
        self.GMHroomProbes.Bind(wx.EVT_COMBOBOX, self.UpdateInstr)
        self.cbox_instr_GMH.append(self.GMHroomProbes)
        
        SwitchboxLbl = wx.StaticText(self, label='Switchbox configuration:', id = wx.ID_ANY)
        self.Switchbox = wx.ComboBox(self,wx.ID_ANY, choices = self.SB_COMBO_CHOICE, style=wx.CB_DROPDOWN)
        self.Switchbox.Bind(wx.EVT_COMBOBOX, self.UpdateInstr)

        # Addresses
        self.V1SrcAddr = wx.ComboBox(self,wx.ID_ANY, choices = self.GPIBAddressList, size = (150,10), style=wx.CB_DROPDOWN)
        self.cbox_addr_GPIB.append(self.V1SrcAddr)
        self.V1SrcAddr.Bind(wx.EVT_COMBOBOX, self.UpdateAddr)
        self.V2SrcAddr = wx.ComboBox(self,wx.ID_ANY, choices = self.GPIBAddressList, style=wx.CB_DROPDOWN)
        self.cbox_addr_GPIB.append(self.V2SrcAddr)
        self.V2SrcAddr.Bind(wx.EVT_COMBOBOX, self.UpdateAddr)
        self.V1V2DvmAddr = wx.ComboBox(self,wx.ID_ANY, choices = self.GPIBAddressList, style=wx.CB_DROPDOWN)
        self.cbox_addr_GPIB.append(self.V1V2DvmAddr)
        self.V1V2DvmAddr.Bind(wx.EVT_COMBOBOX, self.UpdateAddr)
        self.VdDvmAddr = wx.ComboBox(self,wx.ID_ANY, choices = self.GPIBAddressList, style=wx.CB_DROPDOWN)
        self.cbox_addr_GPIB.append(self.VdDvmAddr)
        self.VdDvmAddr.Bind(wx.EVT_COMBOBOX, self.UpdateAddr)
        self.T1DvmAddr = wx.ComboBox(self,wx.ID_ANY, choices = self.GPIBAddressList, style=wx.CB_DROPDOWN)
        self.cbox_addr_GPIB.append(self.T1DvmAddr)
        self.T1DvmAddr.Bind(wx.EVT_COMBOBOX, self.UpdateAddr)
        self.T2DvmAddr = wx.ComboBox(self,wx.ID_ANY, choices = self.GPIBAddressList, style=wx.CB_DROPDOWN)
        self.cbox_addr_GPIB.append(self.T2DvmAddr)
        self.T2DvmAddr.Bind(wx.EVT_COMBOBOX, self.UpdateAddr)
        
        self.GMH1Ports = wx.ComboBox(self,wx.ID_ANY, choices = self.COMAddressList, style=wx.CB_DROPDOWN)
        self.cbox_addr_COM.append(self.GMH1Ports)
        self.GMH1Ports.Bind(wx.EVT_COMBOBOX, self.UpdateAddr)
        self.GMH2Ports = wx.ComboBox(self,wx.ID_ANY, choices = self.COMAddressList, style=wx.CB_DROPDOWN)
        self.cbox_addr_COM.append(self.GMH2Ports)
        self.GMH2Ports.Bind(wx.EVT_COMBOBOX, self.UpdateAddr)
        
        self.GMHroomPorts = wx.ComboBox(self,wx.ID_ANY, choices = self.COMAddressList, style=wx.CB_DROPDOWN)
        self.cbox_addr_COM.append(self.GMHroomPorts)
        self.GMHroomPorts.Bind(wx.EVT_COMBOBOX, self.UpdateAddr)
        
        self.SwitchboxAddr = wx.ComboBox(self,wx.ID_ANY, choices = self.COMAddressList, style=wx.CB_DROPDOWN)
        self.cbox_addr_COM.append(self.SwitchboxAddr)
        

        # Filename
        FileLbl = wx.StaticText(self, label='Excel file full path:', id = wx.ID_ANY)
        self.XLFile = wx.TextCtrl(self, id = wx.ID_ANY, value=self.GetTopLevelParent().ExcelPath)
        
        # Resistors
        self.R1Name = wx.TextCtrl(self, id = wx.ID_ANY, value= 'R1 Name')
        self.R1Name.Bind(wx.EVT_TEXT, self.BuildCommStr)
        RNamesLbl = wx.StaticText(self, label='Resistor Names:', id = wx.ID_ANY)
        self.R2Name = wx.TextCtrl(self, id = wx.ID_ANY, value= 'R2 Name')
        self.R2Name.Bind(wx.EVT_TEXT, self.BuildCommStr)

        # Autopopulate btn
        self.AutoPop = wx.Button(self,id = wx.ID_ANY, label='AutoPopulate')
        self.AutoPop.Bind(wx.EVT_BUTTON, self.OnAutoPop)
        
        # Test buttons
        self.VisaList = wx.Button(self,id = wx.ID_ANY, label='List Visa res')
        self.VisaList.Bind(wx.EVT_BUTTON, self.OnVisaList)
        self.ResList = wx.TextCtrl(self, id = wx.ID_ANY, value = 'Available Visa resources',
                                   style = wx.TE_READONLY|wx.TE_MULTILINE)
        self.S1Test = wx.Button(self,id = wx.ID_ANY, label='Test')
        self.S1Test.Bind(wx.EVT_BUTTON, self.OnTest)
        self.S2Test = wx.Button(self,id = wx.ID_ANY, label='Test')
        self.S2Test.Bind(wx.EVT_BUTTON, self.OnTest)
        self.D12Test = wx.Button(self,id = wx.ID_ANY, label='Test')
        self.D12Test.Bind(wx.EVT_BUTTON, self.OnTest)
        self.DdTest = wx.Button(self,id = wx.ID_ANY, label='Test')
        self.DdTest.Bind(wx.EVT_BUTTON, self.OnTest)
        self.DT1Test = wx.Button(self,id = wx.ID_ANY, label='Test')
        self.DT1Test.Bind(wx.EVT_BUTTON, self.OnTest)
        self.DT2Test = wx.Button(self,id = wx.ID_ANY, label='Test')
        self.DT2Test.Bind(wx.EVT_BUTTON, self.OnTest)
        
        self.GMH1Test = wx.Button(self,id = wx.ID_ANY, label='Test')
        self.GMH1Test.Bind(wx.EVT_BUTTON, self.OnTest)
        self.GMH2Test = wx.Button(self,id = wx.ID_ANY, label='Test')
        self.GMH2Test.Bind(wx.EVT_BUTTON, self.OnTest)
        
        self.GMHroomTest = wx.Button(self,id = wx.ID_ANY, label='Test')
        self.GMHroomTest.Bind(wx.EVT_BUTTON, self.OnTest)
        
        self.SwitchboxTest = wx.Button(self,id = wx.ID_ANY, label='Test')
        self.SwitchboxTest.Bind(wx.EVT_BUTTON, self.OnSwitchTest)
        
        ResponseLbl = wx.StaticText(self, label='Instrument Test Response:', id = wx.ID_ANY)
        self.Response = wx.TextCtrl(self, id = wx.ID_ANY, value= '', style = wx.TE_READONLY)
        
        self.TR1 = wx.TextCtrl(self, id = wx.ID_ANY, value = 'T(R1)', style = wx.TE_READONLY)
        self.TR2 = wx.TextCtrl(self, id = wx.ID_ANY, value = 'T(R2)', style = wx.TE_READONLY)

        gbSizer = wx.GridBagSizer()

        # Instruments
        gbSizer.Add(Src1Lbl, pos=(0,0), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.V1Sources, pos=(0,1), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(Src2Lbl, pos=(1,0), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.V2Sources, pos=(1,1), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(DVM_V1V2Lbl, pos=(2,0), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.V1V2Dvms, pos=(2,1), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(DVM_VdLbl, pos=(3,0), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.VdDvms, pos=(3,1), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(T1DvmLbl, pos=(4,0), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.T1Dvms, pos=(4,1), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(T2DvmLbl, pos=(5,0), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.T2Dvms, pos=(5,1), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(GMH1Lbl, pos=(6,0), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.GMH1Probes, pos=(6,1), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(GMH2Lbl, pos=(7,0), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.GMH2Probes, pos=(7,1), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(GMHroomLbl, pos=(8,0), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.GMHroomProbes, pos=(8,1), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(SwitchboxLbl, pos=(9,0), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.Switchbox, pos=(9,1), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        # Addresses
        gbSizer.Add(self.V1SrcAddr, pos=(0,2), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.V2SrcAddr, pos=(1,2), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.V1V2DvmAddr, pos=(2,2), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.VdDvmAddr, pos=(3,2), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.T1DvmAddr, pos=(4,2), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.T2DvmAddr, pos=(5,2), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.GMH1Ports, pos=(6,2), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.GMH2Ports, pos=(7,2), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.GMHroomPorts, pos=(8,2), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)    
        gbSizer.Add(self.SwitchboxAddr, pos=(9,2), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        # R Name
        gbSizer.Add(self.R1Name, pos=(6,5), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.R2Name, pos=(7,5), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(RNamesLbl, pos=(5,5), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        # Filename
        gbSizer.Add(FileLbl, pos=(11,0), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.XLFile, pos=(11,1), span=(1,5), flag=wx.ALL|wx.EXPAND, border=5)
        # Test buttons
        gbSizer.Add(self.S1Test, pos=(0,3), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.S2Test, pos=(1,3), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.D12Test, pos=(2,3), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.DdTest, pos=(3,3), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.DT1Test, pos=(4,3), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.DT2Test, pos=(5,3), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.GMH1Test, pos=(6,3), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.GMH2Test, pos=(7,3), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.GMHroomTest, pos=(8,3), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.SwitchboxTest, pos=(9,3), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        
        gbSizer.Add(ResponseLbl, pos=(1,4), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.Response, pos=(2,4), span=(1,3), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.TR1, pos=(6,4), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.TR2, pos=(7,4), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.VisaList, pos=(3,5), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.ResList, pos=(3,4), span=(3,1), flag=wx.ALL|wx.EXPAND, border=5)

        # Autopopulate btn
        gbSizer.Add(self.AutoPop, pos=(0,4), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)

        self.SetSizerAndFit(gbSizer)

        # Roles and corresponding comboboxes/test btns are associated here:
        devices.ROLES_WIDGETS = {'SRC1':{'icb':self.V1Sources,'acb':self.V1SrcAddr,'tbtn':self.S1Test}}
        devices.ROLES_WIDGETS.update({'SRC2':{'icb':self.V2Sources,'acb':self.V2SrcAddr,'tbtn':self.S2Test}})
        devices.ROLES_WIDGETS.update({'DVM12':{'icb':self.V1V2Dvms,'acb':self.V1V2DvmAddr,'tbtn':self.D12Test}})
        devices.ROLES_WIDGETS.update({'DVMd':{'icb':self.VdDvms,'acb':self.VdDvmAddr,'tbtn':self.DdTest}})
        devices.ROLES_WIDGETS.update({'DVMT1':{'icb':self.T1Dvms,'acb':self.T1DvmAddr,'tbtn':self.DT1Test}})
        devices.ROLES_WIDGETS.update({'DVMT2':{'icb':self.T2Dvms,'acb':self.T2DvmAddr,'tbtn':self.DT2Test}})
        devices.ROLES_WIDGETS.update({'GMH1':{'icb':self.GMH1Probes,'acb':self.GMH1Ports,'tbtn':self.GMH1Test}})
        devices.ROLES_WIDGETS.update({'GMH2':{'icb':self.GMH2Probes,'acb':self.GMH2Ports,'tbtn':self.GMH2Test}})
        devices.ROLES_WIDGETS.update({'GMHroom':{'icb':self.GMHroomProbes,'acb':self.GMHroomPorts,'tbtn':self.GMHroomTest}})
        devices.ROLES_WIDGETS.update({'switchbox':{'icb':self.Switchbox,'acb':self.SwitchboxAddr,'tbtn':self.SwitchboxTest}})


    def BuildComboChoices(self):
        for d in devices.INSTR_DATA.keys():
            if 'SRC:' in d:
                self.SRC_COMBO_CHOICE.append(d)
            elif 'DVM:' in d:
                self.DVM_COMBO_CHOICE.append(d)
            elif 'GMH:' in d:
                self.GMH_COMBO_CHOICE.append(d)

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


    def UpdateFilepath(self, e):
        self.XLFile.SetValue(e.path)
    
        # Read parameters sheet - gather instrument info:
        wb = load_workbook(self.XLFile.GetValue())
        ws_params = wb.get_sheet_by_name('Parameters')
        
        headings = (None, u'description',u'Instrument Info:',u'parameter',u'value',u'uncert',u'dof',u'label')
        
        # Determine colummn indices from column letters:
        col_I = cell.column_index_from_string('I') - 1
        col_J = cell.column_index_from_string('J') - 1
        col_K = cell.column_index_from_string('K') - 1
        col_L = cell.column_index_from_string('L') - 1
        col_M = cell.column_index_from_string('M') - 1
        col_N = cell.column_index_from_string('N') - 1

        params = []
        values = []
        
        for r in ws_params.rows: # a tuple of row objects
            descr = r[col_I].value # cell.value
            param = r[col_J].value # cell.value
            v_u_d_l = [r[col_K].value, r[col_L].value, r[col_M].value, r[col_N].value] # value,uncert,dof,label
        
            if descr in headings and param in headings:
                continue # Skip this row
            else: # not header
                params.append(param)
                if v_u_d_l[1] is None: # single-valued (no uncert)
                    values.append(v_u_d_l[0]) # append value as next item 
                    print descr,' : ',param,' = ',v_u_d_l[0]
                else: #multi-valued
                    while v_u_d_l[-1] is None: # remove empty cells
                        del v_u_d_l[-1] # v_u_d_l.pop()
                    values.append(v_u_d_l) # append value-list as next item 
                    print descr,' : ',param,' = ',v_u_d_l
                
                if param == u'test': # last parameter for this description
                    devices.DESCR.append(descr) # build description list
                    devices.sublist.append(dict(zip(params,values))) # adds parameter dictionary to sublist
                    del params[:]
                    del values[:] 
    
        print'----END OF PARAMETER LIST----'            
        # Compile into a dictionary that lives in devices.py...  
        devices.INSTR_DATA = dict(zip(devices.DESCR,devices.sublist))
        self.BuildComboChoices()


    def OnAutoPop(self, e):
        # Pre-select instrument and address comboboxes -
        # Choose from instrument descriptions listed in devices.DESCR
        # (Uses address assignments in devices.INSTR_DATA)
        self.instrument_choice = {'SRC1':'SRC: D4808',
                                  'SRC2':'SRC: F5520A',
                                  'DVM12':'DVM: HP3458A, s/n452',
                                  'DVMd':'DVM: HP3458A, s/n230',
                                  'DVMT1':'none',#'DVM: HP34401A, s/n976'
                                  'DVMT2':'none',#'DVM: HP34420A, s/n130'
                                  'GMH1':'GMH: s/n627',
                                  'GMH2':'GMH: s/n628',
                                  'GMHroom':'GMH: s/n367',
                                  'switchbox':'V1'}
        for r in self.instrument_choice.keys():
            d = self.instrument_choice[r]
            devices.ROLES_WIDGETS[r]['icb'].SetValue(d) # Update i_cb
            self.CreateInstr(d,r)
        self.R1Name.SetValue('CHANGE_THIS! 1G')
        self.R2Name.SetValue('CHANGE_THIS! 1M')


    def UpdateInstr(self, e):
        # An instrument was selected for a role.
        # Find description d and role r, then pass to CreatInstr()
        d = e.GetString()
        for r in devices.ROLES_WIDGETS.keys(): # Cycle through roles
            if devices.ROLES_WIDGETS[r]['icb'] == e.GetEventObject():
                break # stop looking when we've found the right instrument, role
        self.CreateInstr(d,r)


    def CreateInstr(self,d,r):
        # Called by both OnAutoPop() and UpdateInstr()
        # Create each instrument in software & open visa session (for GPIB instruments)
        # For GMH instruments, use GMH dll not visa

        if 'GMH' in d:
            # create and open a GMH instrument instance
            print'\nnbpages.SetupPage.CreateInstr(): Creating GMH device (%s -> %s).'%(d,r)
            devices.ROLES_INSTR.update({r:devices.GMH_Sensor(d)})
        else:
            # create a visa instrument instance
            print'\nnbpages.SetupPage.CreateInstr(): Creating VISA device (%s -> %s).'%(d,r)
            devices.ROLES_INSTR.update({r:devices.instrument(d)})
        self.SetInstr(d,r)


    def SetInstr(self,d,r):
        """
        Called by CreateInstr().
        Updates internal info (INSTR_DATA) and Enables/disables testbuttons as necessary.
        """
#        print 'nbpages.SetupPage.SetInstr():',d,'assigned to role',r,'demo mode:',devices.ROLES_INSTR[r].demo
        assert devices.INSTR_DATA.has_key(d),'Unknown instrument: %s - check Excel file is loaded.'%d
        assert devices.INSTR_DATA[d].has_key('role'),'Unknown instrument parameter - check Excel Parameters sheet is populated.'
        devices.INSTR_DATA[d]['role'] = r # update default role
        
        # Set the address cb to correct value (according to devices.INSTR_DATA)
        a_cb = devices.ROLES_WIDGETS[r]['acb']
        a_cb.SetValue((devices.INSTR_DATA[d]['str_addr']))
        if d == 'none':
            devices.ROLES_WIDGETS[r]['tbtn'].Enable(False)
        else:
            devices.ROLES_WIDGETS[r]['tbtn'].Enable(True)


    def UpdateAddr(self, e):
        # An address was manually selected
        # Change INSTR_DATA ...
        # 1st, we'll need instrument description d...
        d = 'none'
        acb = e.GetEventObject() # 'a'ddress 'c'ombo 'b'ox
        for r in devices.ROLES_WIDGETS.keys():
            if devices.ROLES_WIDGETS[r]['acb'] == acb:
                d = devices.ROLES_WIDGETS[r]['icb'].GetValue()
                break # stop looking when we've found the right instrument description
        a = e.GetString() # address string, eg 'COM5' or 'GPIB0::23'
        if (a not in self.GPIBAddressList) or (a not in self.COMAddressList): # Ignore dummy values, like 'NO_ADDRESS'
            devices.INSTR_DATA[d]['str_addr'] = a
            devices.ROLES_INSTR[r].str_addr = a
            addr = a.lstrip('COMGPIB0:') # leave only numeric part of address string
            devices.INSTR_DATA[d]['addr'] = int(addr)
            devices.ROLES_INSTR[r].addr = int(addr)
        print'UpdateAddr():',r,'using',d,'set to addr',addr,'(',a,')'
            

    def OnTest(self, e):
        # Called when a 'test' button is clicked
        d = 'none'
        for r in devices.ROLES_WIDGETS.keys(): # check every role
            if devices.ROLES_WIDGETS[r]['tbtn'] == e.GetEventObject():
                d = devices.ROLES_WIDGETS[r]['icb'].GetValue()
                break # stop looking when we've found the right instrument description
        print'\nnbpages.SetupPage.OnTest():',d
        assert devices.INSTR_DATA[d].has_key('test'), 'No test exists for this device.'
        test = devices.INSTR_DATA[d]['test'] # test string
        print '\tTest string:',test
        self.Response.SetValue(str(devices.ROLES_INSTR[r].Test(test)))
        self.status.SetStatusText('Testing %s with cmd %s' % (d,test),0)


    def OnSwitchTest(self, e):
        resource = self.SwitchboxAddr.GetValue()
        config = str(devices.SWITCH_CONFIGS[self.Switchbox.GetValue()])
        try:
            instr = devices.RM.open_resource(resource)
            instr.write(config)
        except devices.visa.VisaIOError:
            self.Response.SetValue('Couldn\'t open visa resource for switchbox!')


    def BuildCommStr(self,e):
    # Called by a change in GMH probe selection, or resistor name
        d = e.GetString()
        if 'GMH' in d: # A GMH probe selection changed
            # Find the role associated with the selected instrument description
            for r in devices.ROLES_WIDGETS.keys():
                if devices.ROLES_WIDGETS[r]['icb'].GetValue() == d:
                    break
            # Update our knowledge of role <-> instr. descr. association
            self.CreateInstr(d,r)
        RunPage = self.GetParent().GetPage(1)
        params={'R1':self.R1Name.GetValue(),'TR1':self.GMH1Probes.GetValue(),
                'R2':self.R2Name.GetValue(),'TR2':self.GMH2Probes.GetValue()}
        joinstr = ' monitored by '
        commstr = 'R1: ' + params['R1'] + joinstr + params['TR1'] + '. R2: ' + params['R2'] + joinstr + params['TR2']
        evt = evts.UpdateCommentEvent(str=commstr)
        wx.PostEvent(RunPage,evt)


    def OnVisaList(self, e):
        res_list = devices.RM.list_resources()
        del self.ResourceList[:] # list of COM ports ('COM X') & GPIB addresses
        del self.ComList[:] # list of COM ports (numbers only)
        del self.GPIBList[:] # list of GPIB addresses (numbers only)
        for item in res_list:
            self.ResourceList.append(item.replace('ASRL','COM'))
        for item in self.ResourceList:
            addr = item.replace('::INSTR','')
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
        self.Bind(evts.EVT_UPDATE_COM_STR, self.UpdateComment)
        self.Bind(evts.EVT_DATA, self.UpdateData)
        self.Bind(evts.EVT_DELAYS, self.UpdateDels)
        self.Bind(evts.EVT_START_ROW, self.UpdateStartRow)
        self.Bind(evts.EVT_STOP_ROW, self.UpdateStopRow)

        self.RunThread = None
        self.RLinkThread = None

        # Comment widgets
        CommentLbl = wx.StaticText(self,id = wx.ID_ANY, label = 'Comment:')
        self.Comment = wx.TextCtrl(self, id = wx.ID_ANY, size=(600,20))
        self.Comment.Bind(wx.EVT_TEXT,self.OnComment)
        comtip = 'This is auto-generated from data on the Setup page. Other notes may be added manually.'
        self.Comment.SetToolTipString(comtip)
        
        self.NewRunIDBtn = wx.Button(self, id = wx.ID_ANY, label='Create new run id')
        idcomtip = 'New id used to link subsequent Rlink and measurement data.'
        self.NewRunIDBtn.SetToolTipString(idcomtip)
        self.NewRunIDBtn.Bind(wx.EVT_BUTTON, self.OnNewRunID)
        self.RunID = wx.TextCtrl(self, id = wx.ID_ANY, size=(500,20))

        # Voltage source widgets
        V1SrcLbl = wx.StaticText(self,id = wx.ID_ANY, style=wx.ALIGN_RIGHT, label = 'Set V1:')
        self.V1Setting = NumCtrl(self, id = wx.ID_ANY, integerWidth=3, fractionWidth=8, groupDigits=True)
        self.V1Setting.Bind(wx.lib.masked.EVT_NUM, self.OnV1Set)

        V2SrcLbl = wx.StaticText(self,id = wx.ID_ANY, style=wx.ALIGN_RIGHT, label = 'Set V2:')
        self.V2Setting = NumCtrl(self, id = wx.ID_ANY, integerWidth=3, fractionWidth=8, groupDigits=True)
        self.V2Setting.Bind(wx.lib.masked.EVT_NUM , self.OnV2Set)

        ZeroVoltsBtn = wx.Button(self, id = wx.ID_ANY, label='Set zero volts')
        ZeroVoltsBtn.Bind(wx.EVT_BUTTON, self.OnZeroVolts)

        # Delay widgets
        SettleDelLbl = wx.StaticText(self,id = wx.ID_ANY, label = 'Settle delay:')
        self.SettleDel = wx.SpinCtrl(self,id = wx.ID_ANY,value ='0', min = 0, max=600)
        StartDelLbl = wx.StaticText(self,id = wx.ID_ANY, label = 'Start delay:')
        self.StartDel = wx.TextCtrl(self, id = wx.ID_ANY, style = wx.TE_READONLY)
        AZERO1DelLbl = wx.StaticText(self,id = wx.ID_ANY, label = 'AZERO_ONCE delay:')
        self.AZERO1Del = wx.TextCtrl(self, id = wx.ID_ANY, style = wx.TE_READONLY)
        RangeDelLbl = wx.StaticText(self,id = wx.ID_ANY, label = 'Range delay:')
        self.RangeDel = wx.TextCtrl(self, id = wx.ID_ANY, style = wx.TE_READONLY) 
        NSamplesLbl = wx.StaticText(self,id = wx.ID_ANY, label = 'Number of samples:')
        self.NSamples= wx.TextCtrl(self, id = wx.ID_ANY, style = wx.TE_READONLY)
        
        #  Run control and progress widgets
        self.StartRow = wx.TextCtrl(self, id = wx.ID_ANY, style = wx.TE_READONLY)
        StartRowLbl = wx.StaticText(self,id = wx.ID_ANY, label = 'Start row:')
        self.StopRow = wx.TextCtrl(self, id = wx.ID_ANY, style = wx.TE_READONLY)
        StopRowLbl = wx.StaticText(self,id = wx.ID_ANY, label = 'Stop row:')
        RowLbl =  wx.StaticText(self,id = wx.ID_ANY, label = 'Current row:')
        self.Row = wx.TextCtrl(self, id = wx.ID_ANY, style = wx.TE_READONLY)
        TimeLbl = wx.StaticText(self,id = wx.ID_ANY, label = 'Timestamp:')
        self.Time = wx.TextCtrl(self, id = wx.ID_ANY, style = wx.TE_READONLY)
        
        VavLbl = wx.StaticText(self,id = wx.ID_ANY, label = 'Mean voltage(V):')
        #self.Vav = wx.TextCtrl(self, id = wx.ID_ANY, style = wx.TE_READONLY)
        self.Vav = NumCtrl(self, id = wx.ID_ANY, integerWidth=3, fractionWidth=9, groupDigits=True)
        VsdLbl = wx.StaticText(self,id = wx.ID_ANY, label = 'Stdev(voltage):')
        #self.Vsd = wx.TextCtrl(self, id = wx.ID_ANY, style = wx.TE_READONLY)
        self.Vsd = NumCtrl(self, id = wx.ID_ANY, integerWidth=3, fractionWidth=9, groupDigits=True)
        
        self.StartBtn = wx.Button(self, id = wx.ID_ANY, label='Start run')
        self.StartBtn.Bind(wx.EVT_BUTTON, self.OnStart)
        self.StopBtn = wx.Button(self, id = wx.ID_ANY, label='Abort run')
        self.StopBtn.Bind(wx.EVT_BUTTON, self.OnAbort)
        self.RLinkBtn = wx.Button(self, id = wx.ID_ANY, label='Measure R-link')
        self.RLinkBtn.Bind(wx.EVT_BUTTON, self.OnRLink)
        
        ProgressLbl = wx.StaticText(self,id = wx.ID_ANY, style=wx.ALIGN_RIGHT, label = 'Run progress:')
        self.Progress = wx.Gauge(self,id = wx.ID_ANY,range=100, name='Progress')

        gbSizer = wx.GridBagSizer()

        # Comment widgets
        gbSizer.Add(CommentLbl,pos=(0,0), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.Comment, pos=(0,1), span=(1,6), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.NewRunIDBtn, pos=(1,0), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.RunID, pos=(1,1), span=(1,6), flag=wx.ALL|wx.EXPAND, border=5)
        #gbSizer.Add(self.h_sep1, pos=(2,0), span=(1,5), flag=wx.ALL|wx.EXPAND, border=5)

        # Voltage source widgets
        gbSizer.Add(ZeroVoltsBtn, pos=(2,0), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(V1SrcLbl,pos=(2,1), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.V1Setting,pos=(2,2), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(V2SrcLbl,pos=(2,3), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.V2Setting,pos=(2,4), span=(1,1), flag=wx.ALL, border=5)
        #gbSizer.Add(self.h_sep2, pos=(4,0), span=(1,5), flag=wx.ALL|wx.EXPAND, border=5)
        
        # Delay widgets
        gbSizer.Add(SettleDelLbl, pos=(3,0), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.SettleDel, pos=(4,0), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(StartDelLbl, pos=(3,1), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.StartDel, pos=(4,1), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(AZERO1DelLbl, pos=(3,2), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.AZERO1Del, pos=(4,2), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(RangeDelLbl, pos=(3,3), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.RangeDel, pos=(4,3), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        
        gbSizer.Add(NSamplesLbl, pos=(3,4), span=(1,1), flag=wx.ALL, border=5)
        gbSizer.Add(self.NSamples, pos=(4,4), span=(1,1), flag=wx.ALL, border=5)
        #gbSizer.Add(self.h_sep3, pos=(7,0), span=(1,5), flag=wx.ALL|wx.EXPAND, border=5)
        
        #  Run control and progress widgets
        gbSizer.Add(StartRowLbl, pos=(5,0), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.StartRow, pos=(6,0), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(StopRowLbl, pos=(5,1), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.StopRow, pos=(6,1), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(RowLbl, pos=(5,2), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.Row, pos=(6,2), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(TimeLbl, pos=(5,3), span=(1,2), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.Time, pos=(6,3), span=(1,2), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(VavLbl, pos=(5,5), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.Vav, pos=(6,5), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(VsdLbl, pos=(5,6), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.Vsd, pos=(6,6), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        
        gbSizer.Add(self.RLinkBtn, pos=(7,0), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.StartBtn, pos=(7,1), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.StopBtn, pos=(7,2), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(ProgressLbl, pos=(7,3), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.Progress, pos=(7,4), span=(1,3), flag=wx.ALL|wx.EXPAND, border=5)
        
        self.SetSizerAndFit(gbSizer)

        self.autocomstr = ''
        self.manstr = ''

    def OnNewRunID(self,e):
        start = self.fullstr.find('R1: ')
        end = self.fullstr.find(' monitored',start)
        R1name = self.fullstr[start+4:end]
        start = self.fullstr.find('R2: ')
        end = self.fullstr.find(' monitored',start)
        R2name = self.fullstr[start+4:end]
        self.run_id = str('HRBC.v' + self.version + ' ' + R1name + ':' + R2name + ' ' +
                          dt.datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
        self.status.SetStatusText('Id for subsequent runs:',0)
        self.status.SetStatusText(str(self.run_id),1)
        self.RunID.SetValue(str(self.run_id))


    def UpdateComment(self,e):
        # writes combined auto-comment and manual comment when
        # auto-generated comment is re-built
        self.autocomstr = e.str # store a copy of automtically-generated comment
        self.Comment.SetValue(e.str+self.manstr)

    def OnComment(self,e):
        # Called when comment emits EVT_TEXT (i.e. whenever it's changed)
        # Make sure comment field (with extra manually-entered notes) isn't overwritten
        self.fullstr = self.Comment.GetValue() # store a copy of full comment
        # Extract last part of comment (the manually-inserted bit)
        # - assume we manually added extra notes to END
        self.manstr = self.fullstr[len(self.autocomstr):]

    def UpdateData(self,e):
        # Triggered by an 'update data' event
        # event params:(t,Vm,Vsd,r,P,flag['1','2','d' or 'E'])
        if e.flag in 'EF':# finished
            self.RunThread = None
            self.StartBtn.Enable(True)
        else:
            self.Time.SetValue(str(e.t))
            self.Vav.SetValue(str(e.Vm))
            self.Vsd.SetValue(str(e.Vsd))
            self.Row.SetValue(str(e.r))
            self.Progress.SetValue(e.P)

    def UpdateDels(self,e):
        # Triggered by an 'update delays' event
        self.StartDel.SetValue(str(e.s))
        self.NSamples.SetValue(str(e.n))
        self.AZERO1Del.SetValue(str(e.AZ1))
        self.RangeDel.SetValue(str(e.r))

    def UpdateStartRow(self,e):
        # Triggered by an 'update startrow' event
        self.StartRow.SetValue(str(e.row))

    def UpdateStopRow(self,e):
        # Triggered by an 'update stoprow' event
        self.StopRow.SetValue(str(e.row))

    def OnV1Set(self,e):
        # Called by change in value (manually OR by software!)
        V1 = e.GetValue()
        src1 = devices.ROLES_INSTR['SRC1']
        src1.SetV(V1) #'M+0R0='
        time.sleep(0.5)
        if V1 == 0:
            src1.Stby()
        else:
            src1.Oper()
        time.sleep(0.5)


    def OnV2Set(self,e):
        # Called by change in value (manually OR by software!)
        V2 = e.GetValue()
        src2 = devices.ROLES_INSTR['SRC2']
        src2.SetV(V2)
        time.sleep(0.5)
        if V2 == 0:
            src2.Stby()
        else:
            src2.Oper()
        time.sleep(0.5)

    def OnZeroVolts(self,e):
        # V1:
        src1 = devices.ROLES_INSTR['SRC1']
        if self.V1Setting.GetValue() == '0':
            print'RunPage.OnZeroVolts(): Zero/Stby directly (not via V1 display)'
            src1.SetV(0)
            src1.Stby()
        else:
            self.V1Setting.SetValue('0') # Calls OnV1Set() ONLY IF VALUE CHANGES
            print'RunPage.OnZeroVolts():  Zero/Stby via V1 display'

        # V2:
        src2 = devices.ROLES_INSTR['SRC2']
        if self.V2Setting.GetValue() == '0':
            print'RunPage.OnZeroVolts(): Zero/Stby directly (not via V2 display)'
            src2.SetV(0)
            src2.Stby()
        else:
            self.V2Setting.SetValue('0') # Calls OnV2Set() ONLY IF VALUE CHANGES
            print'RunPage.OnZeroVolts():  Zero/Stby via V2 display'

    def OnStart(self,e):
        self.Progress.SetValue(0)
        self.RunThread = None
        self.status.SetStatusText('',1)
        self.status.SetStatusText('Starting run',0)
        if self.RunThread is None:
            self.StopBtn.Enable(True) # Enable Stop button
            self.StartBtn.Enable(False) # Disable Start button
            # start acquisition thread here
            self.RunThread = acq.AqnThread(self)

    def OnAbort(self,e):
        if self.RunThread:
            self.StartBtn.Enable(True)
            self.StopBtn.Enable(False) # Disable Stop button
            self.RunThread.abort()
        elif self.RLinkThread:
            self.RLinkBtn.Enable(True) # Enable Start button
            self.StopBtn.Enable(False) # Disable Stop button
            self.RLinkThread.abort()

    def OnRLink(self,e):
        self.Progress.SetValue(0)
        self.RLinkThread = None
        self.status.SetStatusText('',1)
        self.status.SetStatusText('Starting R-link measurement',0)
        if self.RLinkThread is None:
            self.StopBtn.Enable(True) # Enable Stop button
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

        self.Bind(evts.EVT_PLOT,self.UpdatePlot)
        self.Bind(evts.EVT_CLEARPLOT,self.ClearPlot)

        self.figure = Figure()

        self.figure.subplots_adjust(hspace = 0.3) # 0.3" height space between subplots
        
        self.Vdax = self.figure.add_subplot(3,1,3) # 3high x 1wide, 3rd plot down 
        self.Vdax.ticklabel_format(style='sci', useOffset=False, axis='y', scilimits=(2,-2)) # Auto offset to centre on data
        self.Vdax.yaxis.set_major_formatter(mtick.ScalarFormatter(useMathText=True, useOffset=False)) # Scientific notation .
        self.Vdax.autoscale(enable=True, axis='y', tight=False) # Autoscale with 'buffer' around data extents
        self.Vdax.set_xlabel('time')
        self.Vdax.set_ylabel('Vd')

        self.V1ax = self.figure.add_subplot(3,1,1, sharex=self.Vdax) # 3high x 1wide, 1st plot down 
        self.V1ax.ticklabel_format(useOffset=False, axis='y') # Auto offset to centre on data
        self.V1ax.autoscale(enable=True, axis='y', tight=False) # Autoscale with 'buffer' around data extents
        plt.setp(self.V1ax.get_xticklabels(), visible=False) # Hide x-axis labels
        self.V1ax.set_ylabel('V1')
        self.V1ax.set_ylim(auto=True)
        V1_y_ost = self.V1ax.get_xaxis().get_offset_text()
        V1_y_ost.set_visible(False)

        self.V2ax = self.figure.add_subplot(3,1,2, sharex=self.Vdax) # 3high x 1wide, 2nd plot down 
        self.V2ax.ticklabel_format(useOffset=False, axis='y') # Auto offset to centre on data
        self.V2ax.autoscale(enable=True, axis='y', tight=False) # Autoscale with 'buffer' around data extents
        plt.setp(self.V2ax.get_xticklabels(), visible=False) # Hide x-axis labels
        self.V2ax.set_ylabel('V2')
        self.V2ax.set_ylim(auto=True)
        V2_y_ost = self.V2ax.get_xaxis().get_offset_text()
        V2_y_ost.set_visible(False)

        self.canvas = FigureCanvas(self, wx.ID_ANY, self.figure)
        self.sizer = wx.BoxSizer(wx.VERTICAL)
        self.sizer.Add(self.canvas, 1, wx.LEFT | wx.TOP | wx.GROW)
        self.SetSizerAndFit(self.sizer)


    def UpdatePlot(self, e):
        # six event attributes: td, t1, t2 (list of n times),
        # and Vd, V1, V2 (list of n voltages) plus clear_plot flag
        self.V1ax.plot_date(e.t1, e.V1, 'bo')
        self.V2ax.plot_date(e.t2, e.V2, 'go')
        self.Vdax.plot_date(e.td, e.Vd, 'ro')
        self.figure.autofmt_xdate() # default settings
        self.Vdax.fmt_xdata = mdates.DateFormatter('%d-%m-%Y, %H:%M:%S')
        self.canvas.draw()
        self.canvas.Refresh()

    def ClearPlot(self, e):
        self.V1ax.cla()
        self.V2ax.cla()
        self.Vdax.cla()
        self.Vdax.set_ylabel('Vd')
        self.V1ax.set_ylabel('V1')
        self.V2ax.set_ylabel('V2')
        self.canvas.draw()
        self.canvas.Refresh()
