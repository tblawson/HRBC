#!python
# -*- coding: utf-8 -*-
"""
WORKING VERSION

Created on Mon Jun 29 11:36:13 2015

@author: t.lawson
"""

"""
High_R_Bridge.py - Version 0.3
A Python version of the high resistance bridge TestPoint application.
This app is intended to offer the same functionality as the original
TestPoint version but avoiding the clutter. It uses a wxPython notebook,
with separate pages (tabs) dedicated to:
* Instrument / file setup,
* Run controls and
* Plotting.

The same data input/output protocol as the original will be used, i.e.
initiation parameters will be read from the same spreadsheet as the results
are output to.
"""

import os

#os.environ['GMHPATH'] = 'C:\Users\\t.lawson\Documents\Python Scripts\High_Res_Bridge\GMHdll'
#os.environ['XLPATH'] = 'C:\Users\\t.lawson\Documents\Python Scripts\High_Res_Bridge'

import wx
#import wx.lib.inspection
import nbpages as page
import HighRes_events as evts
import devices

VERSION = "2.0"

print 'HRBC', VERSION


class MainFrame(wx.Frame):
    """
    MainFrame Definition:
    holds the MainPanel in which the appliction runs
    """
    def __init__(self, *args, **kwargs):
        wx.Frame.__init__(self, size=(900, 500), *args, **kwargs)
        self.version = VERSION
        self.ExcelPath = ""

        # Event bindings
        self.Bind(evts.EVT_STAT, self.UpdateStatus)

        self.sb = self.CreateStatusBar()
        self.sb.SetFieldsCount(2)

        MenuBar = wx.MenuBar()
        FileMenu = wx.Menu()

        About = FileMenu.Append(wx.ID_ABOUT, text='&About',
                                help='About HighResBridgeControl (HRBC)')
        self.Bind(wx.EVT_MENU, self.OnAbout, About)

        Open = FileMenu.Append(wx.ID_OPEN, text='&Open',
                               help='Open an Excel file')
        self.Bind(wx.EVT_MENU, self.OnOpen, Open)

        Save = FileMenu.Append(wx.ID_SAVE, text='&Save',
                               help='Save data to an Excel file - ' +
                               'this usually happens automatically ' +
                               'during a run.')
        self.Bind(wx.EVT_MENU, self.OnSave, Save)

        FileMenu.AppendSeparator()

        Quit = FileMenu.Append(wx.ID_EXIT, text='&Quit',
                               help='Exit HighResBridge')
        self.Bind(wx.EVT_MENU, self.OnQuit, Quit)

        MenuBar.Append(FileMenu, "&File")
        self.SetMenuBar(MenuBar)

        # Create a panel to hold the NoteBook...
        self.MainPanel = wx.Panel(self)
        # ... and a Notebook to hold some pages
        self.NoteBook = wx.Notebook(self.MainPanel)

        # Create the page windows as children of the notebook
        self.page1 = page.SetupPage(self.NoteBook)
        self.page2 = page.RunPage(self.NoteBook)
        self.page3 = page.PlotPage(self.NoteBook)

        # Add the pages to the notebook with the label to show on the tab
        self.NoteBook.AddPage(self.page1, "Setup")
        self.NoteBook.AddPage(self.page2, "Run")
        self.NoteBook.AddPage(self.page3, "Plots")

        # Finally, put the notebook in a sizer for the panel to manage
        # the layout
        sizer = wx.BoxSizer()
        sizer.Add(self.NoteBook, 1, wx.EXPAND)
        self.MainPanel.SetSizer(sizer)

    def UpdateStatus(self, e):
        if e.field == 'b':
            self.sb.SetStatusText(e.msg, 0)
            self.sb.SetStatusText(e.msg, 1)
        else:
            self.sb.SetStatusText(e.msg, e.field)

    def OnAbout(self, event=None):
        '''
        A message dialog with 'OK' button. wx.OK is a standard wxWidgets ID.
        '''
        dlg_description = "HRBC v" + VERSION +\
            ": A Python'd version of the TestPoint" +\
            "High Resistance Bridge program."
        dlg_title = "About HighResBridge"
        dlg = wx.MessageDialog(self, dlg_description, dlg_title, wx.OK)
        dlg.ShowModal()  # Show dialog.
        dlg.Destroy()  # Destroy when done.

    def OnSave(self, event=None):
        XLfilename = self.page1.XLFile.GetValue()
        if XLfilename:  # Not [Empty string or None]
            print 'Saving', XLfilename, '...'
            self.page1.wb.save(XLfilename)
            self.page1.log.close()

    def OnOpen(self, event=None):
        dlg = wx.FileDialog(self, message="Select data file",
                            defaultDir=os.getcwd(), defaultFile="",
                            wildcard="*", style=wx.OPEN | wx.CHANGE_DIR)
        if dlg.ShowModal() == wx.ID_OK:
            self.ExcelPath = dlg.GetPath()
            self.directory = dlg.GetDirectory()
            print self.directory
            print self.ExcelPath
            file_evt = evts.FilePathEvent(XLpath=self.ExcelPath,
                                          d=self.directory, v=VERSION)
            wx.PostEvent(self.page1, file_evt)
        dlg.Destroy()

    def CloseInstrSessions(self, event=None):
        for r in devices.ROLES_INSTR.keys():
            devices.ROLES_INSTR[r].Close()
        devices.RM.close()
        print'Main.CloseInstrSessions():' +\
            ' closed VISA resource manager and GMH instruments.'

    def OnQuit(self, event=None):
        self.CloseInstrSessions()
        self.OnSave()
        self.Close()

"""_______________________________________________"""


class MainApp(wx.App):
    """Class MainApp."""
    def OnInit(self):
        """Initiate Main App."""
        self.frame = MainFrame(None, wx.ID_ANY)
        self.frame.Show(True)
        self.SetTopWindow(self.frame)
        self.frame.SetTitle("High Resistance Bridge Control v"+VERSION)
        return True

if __name__ == '__main__':
    app = MainApp(0)
    # wx.lib.inspection.InspectionTool().Show()
    app.MainLoop()
