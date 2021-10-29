import wx
import wx.dataview as dv
import wx.lib.scrolledpanel as sc
import textwrap
import sqlite3
import re
import datetime
import os
import webbrowser
from PyPDF4 import PdfFileMerger
from bs4 import BeautifulSoup
from matplotlib.backends.backend_wxagg import FigureCanvasWxAgg as FigureCanvas
from matplotlib.figure import Figure


# this initializes the database and the cursor
def connect_db(filename):
    global cursr, db
    db = sqlite3.connect(filename)
    with db:
        cursr = db.cursor()
        cursr.execute('PRAGMA foreign_keys=ON')


class StrUpFrm(wx.Frame):
    # the main stsart up form which begins with the selection of the database
    # followed with the login and then selection of the user screens
    def __init__(self):

        super(StrUpFrm, self).__init__(None, wx.ID_ANY,
                                       "Pipe Specification Start Up Screen",
                                       size=(400, 350),
                                       style=wx.DEFAULT_FRAME_STYLE &
                                       ~(wx.MAXIMIZE_BOX | wx.MINIMIZE_BOX))

        self.Bind(wx.EVT_CLOSE, self.OnClosePrt)
        self.InitUI()

    def InitUI(self):
        self.go_value = False
        self.currentDirectory = os.getcwd()
        # create the buttons and bindings
        self.openFileDlgBtn = wx.Button(self, label="Open DataBase")
        self.openFileDlgBtn.Bind(wx.EVT_BUTTON, self.onOpenFile)
        self.openFileDlgBtn.SetBackgroundColour("Light Green")

        self.openPaswrdBtn = wx.Button(self, label='Sign Into Database')
        self.openPaswrdBtn.Bind(wx.EVT_BUTTON, self.Login_Frm)
        self.openPaswrdBtn.Enable(False)

        rdBox1Choices = ['None',
                         'Review Support Tables',
                         'Pipe Wall and Hydro Test Calculations',
                         'Review Commodities and Pipe Specifications',
                         'Manage Passwords and Users']

        bxSzr1 = wx.BoxSizer(wx.VERTICAL)
        self.rdBox1 = wx.RadioBox(self, wx.ID_ANY,
                                  label='Select One of the Following',
                                  choices=rdBox1Choices,
                                  majorDimension=1,
                                  style=wx.RA_SPECIFY_COLS)
        self.rdBox1.SetSelection(0)
        self.rdBox1.Bind(wx.EVT_RADIOBOX, self.OnRadioBox1)
        self.rdBox1.Enable(False)
        bxSzr1.Add(self.rdBox1, 0, wx.ALL, 5)

        self.b1 = wx.Button(self, label="Exit", size=(50, 30))
        self.Bind(wx.EVT_BUTTON, self.OnClosePrt, self.b1)
        self.b1.SetForegroundColour((255, 0, 0))

        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(self.openFileDlgBtn, 0, wx.ALL | wx.CENTER, 5)
        sizer.Add(self.openPaswrdBtn, 0, wx.ALL | wx.CENTER, 5)
        sizer.Add(bxSzr1, 1, wx.ALL | wx.CENTER, 5)
        sizer.Add(self.b1, 0, wx.ALL | wx.CENTER, 5)
        self.SetSizer(sizer)

    def onOpenFile(self, evt):
        dlg = wx.FileDialog(
            self, message="Choose a file",
            defaultDir=self.currentDirectory,
            defaultFile="",
            wildcard="SQLite file (*.db)|*.db",
            style=wx.FD_OPEN | wx.FD_MULTIPLE | wx.FD_CHANGE_DIR
            )
        if dlg.ShowModal() == wx.ID_OK:
            self.path = dlg.GetPaths()[0]
            self.openPaswrdBtn.Enable()

        dlg.Destroy()
        self.openFileDlgBtn.SetBackgroundColour("Yellow")
        self.openPaswrdBtn.SetBackgroundColour("Light Green")

    def Login_Frm(self, evt):
        dlg = LoginFrm(self)
        self.go_value = dlg.ShowModal()
        if self.go_value:
            self.rdBox1.Enable()
            self.openPaswrdBtn.SetBackgroundColour("Yellow")

    def OnRadioBox1(self, evt):
        call_btn = self.rdBox1.GetSelection()
        if call_btn == 1:
            SupportFrms(self)
        elif call_btn == 2:
            CalcFrm(self)
        elif call_btn == 3:
            SpecFrm(self)
        elif call_btn == 4:
            CmbLst1(self, 'Password')

    def OnClosePrt(self, evt):
        if self.go_value is True:
            cursr.close()
            db.close
        self.Destroy()


class LoginFrm(wx.Dialog):
    # the login dialog form
    def __init__(self, parent):

        super(LoginFrm, self).__init__(parent, title='Database Password',
                                       size=(350, 200),
                                       style=wx.DEFAULT_FRAME_STYLE &
                                       ~(wx.RESIZE_BORDER | wx.MAXIMIZE_BOX
                                         | wx.MINIMIZE_BOX))
        self.db_file = parent.path
        self.InitUI()

    def InitUI(self):

        self.Sizer = wx.BoxSizer(wx.VERTICAL)
        usersizer = wx.BoxSizer(wx.HORIZONTAL)
        user_lbl = wx.StaticText(self, -1, size=(160, -1),
                                 style=wx.ALIGN_RIGHT)
        user_lbl.SetLabel('UserName:')
        self.user_box = wx.TextCtrl(self, size=(120, -1), style=wx.TE_LEFT)
        self.user_box.ChangeValue('user')
        usersizer.Add(user_lbl, 0, wx.ALL, 5)
        usersizer.Add(self.user_box, 0)

        passsizer = wx.BoxSizer(wx.HORIZONTAL)
        pass_lbl = wx.StaticText(self, -1, size=(160, -1),
                                 style=wx.ALIGN_RIGHT)
        pass_lbl.SetLabel('Database Password:')
        self.pass_box = wx.TextCtrl(self, size=(120, -1), style=wx.TE_LEFT)
        self.pass_box.ChangeValue('password')
        passsizer.Add(pass_lbl, 0, wx.ALL, 5)
        passsizer.Add(self.pass_box, 0)

        lognsizer = wx.BoxSizer(wx.VERTICAL)
        self.lognbtn = wx.Button(self, label='Login')
        self.lognbtn.SetForegroundColour((255, 0, 0))
        self.Bind(wx.EVT_BUTTON, self.OnLogin, self.lognbtn)
        self.lognbtn.Enable()
        lognsizer.Add(self.lognbtn, 0, wx.ALIGN_CENTER)

        self.Sizer.Add(usersizer, 0, wx.ALIGN_LEFT | wx.TOP, 10)
        self.Sizer.Add(passsizer, 0, wx.ALIGN_LEFT | wx.ALL, 10)
        self.Sizer.Add((10, 20))
        self.Sizer.Add(lognsizer, 0, wx.ALIGN_CENTER)

    def OnLogin(self, evt):
        if (self.user_box.GetValue() or self.pass_box.GetValue()) != '':
            connect_db(self.db_file)
            qry = ('SELECT Password, UserID FROM Password WHERE UserID = "' +
                   str(self.user_box.GetValue()) + '"')
            cursr.execute(qry)
            dta = cursr.fetchall()
            n = 0
            for grp in dta:
                if dta[n][0] == str(self.pass_box.GetValue()):
                    self.EndModal(True)
                else:
                    self.EndModal(False)
                n += 1
        else:
            self.EndModal(False)
        self.Destroy()


class SpecFrm(wx.Frame):
    # this is the form to build the pipe specification based
    # on a selected commodity property
    def __init__(self, parent):

        txt1 = 'Commodity Properties and Related Piping Specification.'
        txt1 += '\t\tADMINISTRATOR USE ONLY'

        super(SpecFrm, self).__init__(parent, title=txt1)

        #  add as this is call up form
        self.parent = parent
        self.lstLft = []

        self.Maximize(True)
        self.SetSizeHints(minW=1125, minH=750)
        self.Bind(wx.EVT_CLOSE, self.OnClose)

        self.pnl = BldFrm(self)

        sizer_1 = wx.BoxSizer(wx.VERTICAL)
        sizer_1.Add(self.pnl, 1, wx.EXPAND | wx.ALL, 0)
        self.SetSizer(sizer_1)
        self.Layout()

        menuBar = wx.MenuBar()
        # File Menu
        fyl = wx.Menu()
        html = fyl.Append(wx.ID_ANY, "Open HTML file")
        self.Bind(wx.EVT_MENU, self.pnl.OnHTML, html)
        pdf = fyl.Append(wx.ID_ANY, "Open PDF file")
        self.Bind(wx.EVT_MENU, self.pnl.OnPDF, pdf)
        mrg = fyl.Append(wx.ID_ANY, "Merge PDF files")
        self.Bind(wx.EVT_MENU, self.pnl.OnMerge, mrg)
        html2pdf = fyl.Append(wx.ID_ANY, "Convert HTML to PDF")
        self.Bind(wx.EVT_MENU, self.pnl.OnConvert, html2pdf)
        dlet = fyl.Append(wx.ID_ANY, "Delete Document")
        self.Bind(wx.EVT_MENU, self.pnl.OnDelete, dlet)
        exxt = fyl.Append(wx.ID_EXIT, "E&xit")
        self.Bind(wx.EVT_MENU, self.OnClose, exxt)
        menuBar.Append(fyl, "&File")

        forms = wx.Menu()
        sow = forms.Append(200, 'Scope of Work')
        self.Bind(wx.EVT_MENU, self.pnl.OnForm, sow)
        htr = forms.Append(201, 'Hydro Test Report')
        self.Bind(wx.EVT_MENU, self.pnl.OnForm, htr)
        htw = forms.Append(202, 'Hydro Test Waiver')
        self.Bind(wx.EVT_MENU, self.pnl.OnForm, htw)
        ncr = forms.Append(203, 'Nonconformance Report')
        self.Bind(wx.EVT_MENU, self.pnl.OnForm, ncr)
        msr = forms.Append(204, 'Material Substitution Request')
        self.Bind(wx.EVT_MENU, self.pnl.OnForm, msr)
        rns = forms.Append(205, 'Request for New Specification')
        self.Bind(wx.EVT_MENU, self.pnl.OnForm, rns)
        its = forms.Append(206, 'Inspection Travel Sheet')
        self.Bind(wx.EVT_MENU, self.pnl.OnITS, its)
        menuBar.Append(forms, 'F&orms')

        dtaimp = wx.Menu()
        ncrimp = dtaimp.Append(400, 'Nonconformance HTML Data')
        self.Bind(wx.EVT_MENU, self.pnl.OnNCRImp, ncrimp)
        msrimp = dtaimp.Append(401, 'Material Substitution HTML Data')
        self.Bind(wx.EVT_MENU, self.pnl.OnMSRImp, msrimp)
        rnsimp = dtaimp.Append(402, 'Request New Spec HTML Data')
        self.Bind(wx.EVT_MENU, self.pnl.OnRNSImp, rnsimp)
        menuBar.Append(dtaimp, '&Import')

        frmhelp = wx.Menu()
        calc = frmhelp.Append(300, 'Misc Calculations')
        self.Bind(wx.EVT_MENU, self.pnl.OnCalcs, calc)
        hlp = frmhelp.Append(301, 'Use of Form')
        self.Bind(wx.EVT_MENU, self.pnl.OnHelp, hlp)
        abut = frmhelp.Append(302, 'About')
        self.Bind(wx.EVT_MENU, self.pnl.OnAbout, abut)
        menuBar.Append(frmhelp, '&Help')

        self.SetMenuBar(menuBar)

        # add these following lines since this is a call up form
        self.CenterOnParent()
        self.GetParent().Enable(False)
        self.Show(True)
        self.__eventLoop = wx.GUIEventLoop()
        self.__eventLoop.Run()

    def OnClose(self, evt):
        self.GetParent().Enable(True)   # add this line for child
        self.__eventLoop.Exit()     # add this line for child
        self.Destroy()


class BldFrm(wx.Panel):
    # The panel holding the widgets for the SpecFrm
    def __init__(self, parent, model=None):
        self.model = model

        super(BldFrm, self).__init__(parent)

        self.btns = []
        self.lctrls = []
        self.tblname = 'CommodityProperties'
        self.Srch = False
        self.NewSpec = False
        self.ComdPrtyID = None

        self.InitUI()

    def InitUI(self):
        font1 = wx.Font(12, wx.DEFAULT, wx.NORMAL, wx.FONTWEIGHT_BOLD)

        self.data = self.LoadComdData()

        # set up the table column names, width and if
        # column can be edited ie primary autoincrement
        tblinfo = []

        tblinfo = Dbase(self.tblname).Fld_Size_Type()
        ID_col = tblinfo[1]
        autoincrement = tblinfo[2]

        self.columnames = ['ID', 'Commodity\nCode', 'Commodity Description',
                           'Pipe Material\nSpecification', 'Fluid Category',
                           'Design\nPressure', 'Min. Design\nTemperature',
                           'Max. Design\nTemperature', 'End Connections',
                           'Specification\nCode', 'Pending', 'Note']

        self.colwdth = [6, 10, 35, 15, 17, 10, 11, 11, 18, 11, 6, 30]

        # Create a dataview control
        self.dvc = dv.DataViewCtrl(self, wx.ID_ANY, wx.DefaultPosition,
                                   wx.Size(500, 300),
                                   style=wx.BORDER_THEME
                                   | dv.DV_ROW_LINES
                                   | dv.DV_VERT_RULES
                                   | dv.DV_HORIZ_RULES
                                   | dv.DV_SINGLE
                                   | dv.DV_VARIABLE_LINE_HEIGHT)

        # if autoincrement is false then the data can be sorted based on ID_col
        if autoincrement == 0:
            self.data.sort(key=lambda tup: tup[ID_col])

        # use the sorted data to load the dataviewlistcontrol
        if self.model is None:
            self.model = DataMods(self.tblname, self.data)
        self.dvc.AssociateModel(self.model)

        n = 0
        for colname in self.columnames:
            self.dvc.AppendTextColumn(colname, n,
                                      width=wx.LIST_AUTOSIZE_USEHEADER,
                                      mode=dv.DATAVIEW_CELL_INERT)
            n += 1

        # make columns not sortable and but reorderable.
        for c in self.dvc.Columns:
            c.Sortable = False
            c.Reorderable = True
            c.Resizeable = True

        self.dvc.Columns[(1)].Sortable = True
        self.dvc.Columns[(3)].Sortable = True

        # change to not let the ID col be moved.
        self.dvc.Columns[(ID_col)].Reorderable = False
        self.dvc.Columns[(ID_col)].Resizeable = False

        # Bind some events so we can see what the DVC sends us
        self.Bind(dv.EVT_DATAVIEW_ITEM_ACTIVATED, self.OnGridSelect, self.dvc)

        # set the Sizer property (same as SetSizer)
        self.Sizer = wx.BoxSizer(wx.VERTICAL)

        # set up first row of combo boxesand label
        self.lblsizer1 = wx.BoxSizer(wx.HORIZONTAL)
        self.cmbsizer1 = wx.BoxSizer(wx.HORIZONTAL)
        self.lblsizer2 = wx.BoxSizer(wx.HORIZONTAL)
        self.cmbsizer2 = wx.BoxSizer(wx.HORIZONTAL)
        self.dvsizer = wx.BoxSizer(wx.HORIZONTAL)

        self.dvsizer.Add(10, -1, 0)
        self.dvsizer.Add(self.dvc, 1, wx.ALL | wx.EXPAND, 5)
        self.dvsizer.Add(10, -1, 0)

        self.Codelbl = wx.StaticText(self, -1, style=wx.ALIGN_RIGHT)
        self.Codelbl.SetLabel('Commodity\nCode')
        self.Codelbl.SetForegroundColour((255, 0, 0))

        self.addCode = wx.Button(self, label='+', size=(35, -1))
        self.addCode.SetForegroundColour((255, 0, 0))
        self.addCode.SetFont(font1)
        self.Bind(wx.EVT_BUTTON, self.OnAddCode, self.addCode)

        self.Speclbl = wx.StaticText(self, -1, style=wx.ALIGN_RIGHT)
        self.Speclbl.SetLabel("Piping\nMat'r Spec")
        self.Speclbl.SetForegroundColour((255, 0, 0))

        self.addSpec = wx.Button(self, label='+', size=(35, -1))
        self.addSpec.SetForegroundColour((255, 0, 0))
        self.addSpec.SetFont(font1)
        self.Bind(wx.EVT_BUTTON, self.OnAddSpec, self.addSpec)

        self.Fluidlbl = wx.StaticText(self, -1, style=wx.ALIGN_RIGHT)
        self.Fluidlbl.SetLabel('Fluid\nCategory')
        self.Fluidlbl.SetForegroundColour((255, 0, 0))

        self.addFld = wx.Button(self, label='+', size=(35, -1))
        self.addFld.SetForegroundColour((255, 0, 0))
        self.addFld.SetFont(font1)
        self.Bind(wx.EVT_BUTTON, self.OnAddFld, self.addFld)

        self.DPlbl = wx.StaticText(self, -1, style=wx.ALIGN_CENTER)
        self.DPlbl.SetLabel('Design\nPressure')
        self.DPlbl.SetForegroundColour((255, 0, 0))

        self.DTlbl = wx.StaticText(self, -1, style=wx.ALIGN_CENTER)
        self.DTlbl.SetLabel(' Design Temperature\nMin.         Max.')
        self.DTlbl.SetForegroundColour((255, 0, 0))

        self.Endlbl = wx.StaticText(self, -1, style=wx.ALIGN_RIGHT)
        self.Endlbl.SetLabel('Commodity End\nConnections')
        self.Endlbl.SetForegroundColour((255, 0, 0))

        self.addEnd = wx.Button(self, label='+', size=(35, -1))
        self.addEnd.SetForegroundColour((255, 0, 0))
        self.addEnd.SetFont(font1)
        self.Bind(wx.EVT_BUTTON, self.OnAddEnd, self.addEnd)

        self.lblsizer1.Add(self.Codelbl, 0, wx.ALL, 5)
        self.lblsizer1.Add(self.addCode, 0, wx.ALL, 5)
        self.lblsizer1.Add(10, -1, 0)
        self.lblsizer1.Add(self.Speclbl, 0, wx.ALL, 5)
        self.lblsizer1.Add(self.addSpec, 0, wx.ALL, 5)
        self.lblsizer1.Add(35, -1, 0)
        self.lblsizer1.Add(self.Fluidlbl, 0, wx.ALL, 5)
        self.lblsizer1.Add(self.addFld, 0, wx.ALL, 5)
        self.lblsizer1.Add(80, -1, 0)
        self.lblsizer1.Add(self.DPlbl, 0, wx.ALL, 5)
        self.lblsizer1.Add(15, -1, 0)
        self.lblsizer1.Add(self.DTlbl, 0, wx.ALL, 5)
        self.lblsizer1.Add(120, -1, 0)
        self.lblsizer1.Add(self.Endlbl, 0, wx.ALL, 5)
        self.lblsizer1.Add(self.addEnd, 0, wx.ALL, 5)
        self.lblsizer1.Add(110, -1, 0)

        self.textDP = wx.TextCtrl(self, size=(80, -1), style=wx.TE_CENTER)
        self.textDP.SetHint('psig')

        self.textMinT = wx.TextCtrl(self, size=(80, -1), style=wx.TE_CENTER)
        self.textMinT.SetHint('Deg F')

        self.textMaxT = wx.TextCtrl(self, size=(80, -1), style=wx.TE_CENTER)
        self.textMaxT.SetHint('Deg F')

        # Start the generation of the required combo boxes
        # using a dictionary of column names and table names

        self.cmbCode = wx.ComboCtrl(self, size=(80, -1), style=wx.CB_READONLY)
        self.Bind(wx.EVT_TEXT, self.OnSelectC, self.cmbCode)
        self.Bind(wx.EVT_COMBOBOX_CLOSEUP, self.OnSelectC, self.cmbCode)
        self.cmbCode.SetPopupControl(ListCtrlComboPopup(
            'CommodityCodes', showcol=0, PupQuery='', lctrls=self.lctrls))

        self.cmbPipe = wx.ComboCtrl(self, size=(90, -1), style=wx.CB_READONLY)
        self.Bind(wx.EVT_COMBOBOX_CLOSEUP, self.OnSelectP, self.cmbPipe)
        self.Bind(wx.EVT_TEXT, self.OnSelectP, self.cmbPipe)
        self.cmbPipe.SetPopupControl(ListCtrlComboPopup(
            'PipeMaterialSpec', showcol=1, PupQuery='', lctrls=self.lctrls))

        self.cmbFluid = wx.ComboCtrl(self, size=(150, -1),
                                     style=wx.CB_READONLY)
        self.Bind(wx.EVT_COMBOBOX_CLOSEUP, self.OnSelectF, self.cmbFluid)
        self.Bind(wx.EVT_TEXT, self.OnSelectF, self.cmbFluid)
        self.cmbFluid.SetPopupControl(ListCtrlComboPopup(
            'FluidCategory', showcol=1, PupQuery='', lctrls=self.lctrls))

        self.cmbEnd = wx.ComboCtrl(self, size=(300, -1), style=wx.CB_READONLY)
        self.Bind(wx.EVT_COMBOBOX_CLOSEUP, self.OnSelectE, self.cmbEnd)
        self.Bind(wx.EVT_TEXT, self.OnSelectE, self.cmbEnd)
        self.cmbEnd.SetPopupControl(ListCtrlComboPopup(
            'CommodityEnds', showcol=1, PupQuery='', lctrls=self.lctrls))

        self.txtDscrpt = wx.TextCtrl(self, size=(300, -1),
                                     style=wx.TE_CENTER)

        self.textSpec = wx.TextCtrl(self, size=(120, -1),
                                    style=wx.TE_CENTER)

        self.chkPend = wx.CheckBox(self)
        self.chkPend.SetValue(False)
        self.chkPend.SetForegroundColour(wx.Colour(255, 0, 0))

        self.notes = wx.TextCtrl(self, size=(700, 40), value='',
                                 style=wx.TE_MULTILINE | wx.TE_LEFT)

        self.cmbsizer1.Add(self.cmbCode, 0)
        self.cmbsizer1.Add((55, 5))
        self.cmbsizer1.Add(self.cmbPipe, 0)
        self.cmbsizer1.Add((45, 5))
        self.cmbsizer1.Add(self.cmbFluid, 0)
        self.cmbsizer1.Add((45, 5))
        self.cmbsizer1.Add(self.textDP, 0)
        self.cmbsizer1.Add(self.textMinT, 0)
        self.cmbsizer1.Add(self.textMaxT, 0)
        self.cmbsizer1.Add((45, 5))
        self.cmbsizer1.Add(self.cmbEnd, 0)

        self.Dscrptlbl = wx.StaticText(self, -1, style=wx.ALIGN_CENTER)
        self.Dscrptlbl.SetLabel('Commodity\nDescription')
        self.Dscrptlbl.SetForegroundColour((255, 0, 0))

        self.Speclbl = wx.StaticText(self, -1, style=wx.ALIGN_CENTER)
        self.Speclbl.SetLabel('Piping\nSpecificaion')
        self.Speclbl.SetForegroundColour((255, 0, 0))

        self.Pendlbl = wx.StaticText(self, -1, style=wx.ALIGN_CENTER)
        self.Pendlbl.SetLabel('Specificaion\nPending Approval')
        self.Pendlbl.SetForegroundColour((255, 0, 0))

        self.lblsizer2.Add(self.Dscrptlbl, 0, wx.TOP, 15)
        self.lblsizer2.Add(330, -1, 0)
        self.lblsizer2.Add(self.Speclbl, 0, wx.TOP, 15)
        self.lblsizer2.Add(25, -1, 0)
        self.lblsizer2.Add(self.Pendlbl, 0, wx.TOP, 15)
        self.lblsizer2.Add(10, -1, 0)

        # add a button to call main form to search combo list data
        self.b6 = wx.Button(self, label="<-- Search\nDescription")
        self.Bind(wx.EVT_BUTTON, self.OnSearch, self.b6)

        self.b5 = wx.Button(self, label="Reset")
        self.Bind(wx.EVT_BUTTON, self.OnRestoreBoxs, self.b5)

        self.cmbsizer2.Add(15, -1, 0)
        self.cmbsizer2.Add(self.txtDscrpt, 0, wx.ALIGN_CENTRE)
        self.cmbsizer2.Add(10, -1, 0)
        self.cmbsizer2.Add(self.b6, 0, wx.BOTTOM | wx.ALIGN_CENTRE, 5)
        self.cmbsizer2.Add(45, -1, 0)
        self.cmbsizer2.Add(self.textSpec, 0, wx.BOTTOM | wx.CENTER, 5)
        self.cmbsizer2.Add(40, -1, 0)
        self.cmbsizer2.Add(self.chkPend, 0, wx.BOTTOM, 40)
        self.cmbsizer2.Add(40, -1, 0)
        self.cmbsizer2.Add(self.b5, 0, wx.BOTTOM | wx.ALIGN_CENTER, 5)

        self.Notelbl = wx.StaticText(self, -1, style=wx.ALIGN_CENTER)
        self.Notelbl.SetLabel('Fluid Specific Notes')
        self.Notelbl.SetForegroundColour((255, 0, 0))

        self.notebox = wx.BoxSizer(wx.HORIZONTAL)
        self.notebox.Add(self.Notelbl, 1, wx.ALL | wx.ALIGN_CENTER, 25)
        self.notebox.Add(self.notes, 0, wx.ALIGN_CENTER, 5)

        # Add some buttons
        self.b1 = wx.Button(self, label="Print All\nCommodity\nProperties")
        self.Bind(wx.EVT_BUTTON, self.PrintFile, self.b1)

        self.b2 = wx.Button(self, label="Save New\nor Update")
        self.Bind(wx.EVT_BUTTON, self.OnDevSpec, self.b2)

        self.b3 = wx.Button(self, label="Delete Row")
        self.Bind(wx.EVT_BUTTON, self.OnDeleteRow, self.b3)

        # add a button box and place the buttons
        self.btnbox = wx.BoxSizer(wx.HORIZONTAL)
        self.btnbox.Add(self.b1, 0, wx.ALL | wx.ALIGN_CENTER)
        self.btnbox.Add((20, 10))
        self.btnbox.Add(self.b2, 0, wx.ALIGN_CENTER)
        self.btnbox.Add((20, 10))
        self.btnbox.Add(self.b3, 0, wx.ALIGN_CENTER)
        self.btnbox.Add((60, 10))

        # add static label to explain how to add / edit data
        self.editlbl = wx.StaticText(self, -1, style=wx.ALIGN_CENTER)
        txt = ('''To Add a new Commodity Property or Edit and exiting double\
 click on the corresponding Commodity Code or click the Save\
 New/Update and build\na new Commodity property from scratch.\
   When done click Save New/Update, then save either as new Property\
 or update the existing.''')
        self.editlbl.SetLabel(txt)
        self.editlbl.SetForegroundColour((255, 0, 0))

        # add the frame to hold the various form selection buttons
        self.optstb = wx.StaticBox(self, -1, 'Specification Details')
        self.optsizer = wx.StaticBoxSizer(self.optstb, wx.HORIZONTAL)

        self.optbox1 = wx.BoxSizer(wx.VERTICAL)
        self.btns.append(wx.Button(self, size=(80, -1),
                         label="Pipe &&\nNipples"))
        self.Bind(wx.EVT_BUTTON, self.OnPipe, self.btns[0])
        self.btns.append(wx.Button(self, size=(80, -1), label="Gaskets"))
        self.Bind(wx.EVT_BUTTON, self.OnGskt, self.btns[1])
        self.btns.append(wx.Button(self, size=(80, -1), label="Clamps"))
        self.Bind(wx.EVT_BUTTON, self.OnClmp, self.btns[2])
        self.optbox1.Add(self.btns[0], 0, wx.ALL, 5)
        self.optbox1.Add(self.btns[1], 0, wx.ALL, 5)
        self.optbox1.Add(self.btns[2], 0, wx.ALL, 5)

        self.optbox2 = wx.BoxSizer(wx.VERTICAL)
        self.btns.append(wx.Button(self, size=(80, -1), label="Branch\nChart"))
        self.Bind(wx.EVT_BUTTON, self.OnBrch, self.btns[3])
        self.btns.append(wx.Button(self, size=(95, -1), label="Plug Valves"))
        self.Bind(wx.EVT_BUTTON, self.OnPG, self.btns[4])
        self.btns.append(wx.Button(self, size=(95, -1), label="Tubing"))
        self.Bind(wx.EVT_BUTTON, self.OnTube, self.btns[5])
        self.optbox2.Add(self.btns[3], 0, wx.ALL | wx.LEFT, 5)
        self.optbox2.Add(self.btns[4], 0, wx.ALL | wx.LEFT, 5)
        self.optbox2.Add(self.btns[5], 0, wx.ALL | wx.LEFT, 5)

        self.optbox3 = wx.BoxSizer(wx.VERTICAL)
        self.btns.append(wx.Button(self, size=(80, -1), label="Unions"))
        self.Bind(wx.EVT_BUTTON, self.OnUnion, self.btns[6])
        self.btns.append(wx.Button(self, size=(95, -1), label="Gate\nValves"))
        self.Bind(wx.EVT_BUTTON, self.OnGV, self.btns[7])
        self.btns.append(wx.Button(self, size=(90, -1), label="Bolting"))
        self.Bind(wx.EVT_BUTTON, self.OnBolt, self.btns[8])
        self.optbox3.Add(self.btns[6], 0, wx.ALL | wx.LEFT, 5)
        self.optbox3.Add(self.btns[7], 0, wx.ALL | wx.LEFT, 5)
        self.optbox3.Add(self.btns[8], 0, wx.ALL | wx.LEFT, 5)

        self.optbox4 = wx.BoxSizer(wx.VERTICAL)
        self.btns.append(wx.Button(self, size=(80, -1), label="Fittings"))
        self.Bind(wx.EVT_BUTTON, self.OnFtg, self.btns[9])
        self.btns.append(wx.Button(self, size=(100, -1),
                         label="Swing Check\nValves"))
        self.Bind(wx.EVT_BUTTON, self.OnSWC, self.btns[10])
        self.btns.append(wx.Button(self, size=(90, -1), label="Welding"))
        self.Bind(wx.EVT_BUTTON, self.OnWeld, self.btns[11])
        self.optbox4.Add(self.btns[9], 0, wx.ALL | wx.LEFT, 5)
        self.optbox4.Add(self.btns[10], 0, wx.ALL | wx.LEFT, 5)
        self.optbox4.Add(self.btns[11], 0, wx.ALL | wx.LEFT, 5)

        self.optbox5 = wx.BoxSizer(wx.VERTICAL)
        self.btns.append(wx.Button(self, size=(80, -1), label="O-Lets"))
        self.Bind(wx.EVT_BUTTON, self.OnOlet, self.btns[12])
        self.btns.append(wx.Button(self, size=(100, -1),
                         label="Piston Check\nValves"))
        self.Bind(wx.EVT_BUTTON, self.OnPC, self.btns[13])
        self.btns.append(wx.Button(self, size=(90, -1), label="Insulation"))
        self.Bind(wx.EVT_BUTTON, self.OnInsul, self.btns[14])
        self.optbox5.Add(self.btns[12], 0, wx.ALL | wx.LEFT, 5)
        self.optbox5.Add(self.btns[13], 0, wx.ALL | wx.LEFT, 5)
        self.optbox5.Add(self.btns[14], 0, wx.ALL | wx.LEFT, 5)

        self.optbox6 = wx.BoxSizer(wx.VERTICAL)
        self.btns.append(wx.Button(self, size=(80, -1), label="Flanges"))
        self.Bind(wx.EVT_BUTTON, self.OnFlg, self.btns[15])
        self.btns.append(wx.Button(self, size=(80, -1),
                         label="Butterfly\nValves"))
        self.Bind(wx.EVT_BUTTON, self.OnBU, self.btns[16])
        self.btns.append(wx.Button(self, size=(90, -1), label="Paint"))
        self.Bind(wx.EVT_BUTTON, self.OnPaint, self.btns[17])
        self.optbox6.Add(self.btns[15], 0, wx.ALL | wx.LEFT, 5)
        self.optbox6.Add(self.btns[16], 0, wx.ALL | wx.LEFT, 5)
        self.optbox6.Add(self.btns[17], 0, wx.ALL | wx.LEFT, 5)

        self.optbox7 = wx.BoxSizer(wx.VERTICAL)
        self.btns.append(wx.Button(self, size=(80, -1),
                         label="Orifice\nFlanges"))
        self.Bind(wx.EVT_BUTTON, self.OnOFlg, self.btns[18])
        self.btns.append(wx.Button(self, size=(95, -1), label="Ball Valves"))
        self.Bind(wx.EVT_BUTTON, self.OnBL, self.btns[19])
        self.btns.append(wx.Button(self, size=(90, -1), label="Inspection"))
        self.Bind(wx.EVT_BUTTON, self.OnInspct, self.btns[20])
        self.optbox7.Add(self.btns[18], 0, wx.ALL | wx.LEFT, 5)
        self.optbox7.Add(self.btns[19], 0, wx.ALL | wx.LEFT, 5)
        self.optbox7.Add(self.btns[20], 0, wx.ALL | wx.LEFT, 5)

        self.optbox8 = wx.BoxSizer(wx.VERTICAL)
        self.btns.append(wx.Button(self, size=(95, -1), label="Globe\nValves"))
        self.Bind(wx.EVT_BUTTON, self.OnGB, self.btns[21])
        self.btns.append(wx.Button(self, size=(90, -1), label="Specials"))
        self.Bind(wx.EVT_BUTTON, self.OnSpecl, self.btns[22])
        self.btns.append(wx.Button(self, size=(90, -1), label="Notes"))
        self.Bind(wx.EVT_BUTTON, self.OnNotes, self.btns[23])
        self.optbox8.Add(self.btns[21], 0, wx.ALL | wx.LEFT, 5)
        self.optbox8.Add(self.btns[22], 0, wx.ALL | wx.LEFT, 5)
        self.optbox8.Add(self.btns[23], 0, wx.ALL | wx.LEFT, 5)

        self.optsizer.Add((5, 10))
        self.optsizer.Add(self.optbox1, 0)
        self.optsizer.Add(self.optbox2, 0)
        self.optsizer.Add(self.optbox3, 0)
        self.optsizer.Add(self.optbox4, 0)
        self.optsizer.Add(self.optbox5, 0)
        self.optsizer.Add(self.optbox6, 0)
        self.optsizer.Add(self.optbox7, 0)
        self.optsizer.Add(self.optbox8, 0)
        self.optsizer.Add((5, 10))

        self.Sizer.Add((20, 15))
        self.Sizer.Add(self.lblsizer1, 0, wx.ALIGN_CENTER)
        self.Sizer.Add(self.cmbsizer1, 0, wx.ALIGN_CENTER)
        self.Sizer.Add(self.lblsizer2, 0, wx.ALIGN_CENTER)
        self.Sizer.Add(self.cmbsizer2, 0, wx.ALIGN_CENTER)
        self.Sizer.Add(self.dvsizer, 1, wx.EXPAND)
        self.Sizer.Add(self.notebox, 0, wx.ALIGN_CENTER)
        self.Sizer.Add(self.editlbl, 0, wx.ALIGN_CENTER | wx.BOTTOM, 10)
        self.Sizer.Add(self.btnbox, 0, wx.ALIGN_CENTER)
        self.Sizer.Add(self.optsizer, 0, wx.ALL | wx.ALIGN_CENTER, 5)
        self.Sizer.Add((20, 30))

        if self.ComdPrtyID is None:
            for i in range(24):
                self.btns[i].Enable(False)

    def LoadComdData(self):
        # provides the actual data from the table
        # note the order of the fields is important and reflects the order set
        # up in the SQLite database there should be a more secure way of doing
        # this that does not depend on the database structure
        self.Dsql = ('''SELECT v.CommodityPropertyID, a.Commodity_Code,
         a.Commodity_Description, b.Pipe_Material_Spec,
         d.Designation, v.Design_Pressure, v.Minimum_Design_Temperature,
         v.Maximum_Design_Temperature, c.Commodity_End, v.Pipe_Code,
         v.Pending, v.Note''')

        self.Dsql = self.Dsql + ' FROM ' + self.tblname + ' v' + '\n'

        # the index fields are returned from the PRAGMA statement tbldata
        # in reverse order so the alpha characters need to count up not down
        # join above list of grid columns + INNER JOIN Frg_Tbl 'alpha' ON
        # v.LinkField = 'alpha'.Frg_Fld where alpha is incremented
        n = 0
        tbldata = Dbase().Dtbldata(self.tblname)
        for element in tbldata:
            alpha = chr(96-n+len(tbldata))
            self.Dsql = (self.Dsql + ' INNER JOIN ' + element[2] + ' ' + alpha
                         + ' ON v.' + element[3] + ' = ' + alpha + '.' +
                         element[4] + '\n')
            n += 1
        data = Dbase().Dsqldata(self.Dsql)
        return data

    def OnPipe(self, evt):
        BldConduit(self, 'Piping', ComdPrtyID=self.ComdPrtyID)

    def OnForm(self, evt):
        slectID = evt.GetId()
        fileid = str(slectID)[-1]
        filenames = ['Original_SOW.html', 'Original_HTR.html',
                     'Original_HTW.html', 'Original_NCR.html',
                     'Original_MSR.html', 'Original_RPS.html']
        html_path = ''

        # this should be the default location for the files
        os.chdir('..')
        filename = ('file:' + os.sep*2 + os.getcwd() + os.sep + 'Forms'
                    + os.sep + filenames[int(fileid)])

        # if default location does not work then open directory finder
        if os.path.isfile(filename) is False:
            dlg = wx.DirDialog(self,
                               'Select the template forms Directory:',
                               style=wx.DD_DEFAULT_STYLE)

            if dlg.ShowModal() == wx.ID_OK:
                html_path = dlg.GetPath()
            dlg.Destroy()

        if html_path != '':
            filename = ('file:' + os.sep*2 + html_path
                        + os.sep + filenames[int(fileid)])

        if filename == '':
            wx.MessageBox('Problem Locating HTML File', 'Error', wx.OK)
        else:
            '''brwsr_lst = list(webbrowser._browsers.keys())

            all_brwsrs = ['firefox', 'safari', 'chrome', 'opera',
                        'netscape', 'google-chrome', 'lynx',
                        'mozilla', 'galeon', 'chromium',
                        'chromium-browser', 'windows-default', 'w3m']

            if list(set(brwsr_lst) & set(all_brwsrs)) != []:
                select_brwsr = list(set(brwsr_lst) & set(all_brwsrs))[0]
                webbrowser.get(select_brwsr).open(filename, new=2)
            else:'''
            try:
                webbrowser.get(using=None).open(filename, new=2)
            except Exception:
                wx.MessageBox('Problem Locating Web Browser', 'Error', wx.OK)

    def OnHTML(self, evt):
        # show the html file in browser
        wildcard = "HTML file (*.html)|*.html"
        msg = 'Select HTML File to View'
        filename = self.FylDilog(wildcard, msg, wx.FD_OPEN |
                                 wx.FD_CHANGE_DIR)
        if filename != 'No File':
            try:
                webbrowser.get(using=None).open(filename, new=2)
            except Exception:
                wx.MessageBox('Problem Locating Web Browser', 'Error', wx.OK)
            '''brwsr_lst = list(webbrowser._browsers.keys())
            all_brwsrs = ['firefox', 'safari', 'chrome', 'opera',
                          'netscape', 'google-chrome', 'lynx',
                          'mozilla', 'galeon', 'chromium',
                          'chromium-browser', 'windows-default', 'w3m']
            select_brwsr = list(set(brwsr_lst) & set(all_brwsrs))[0]

            if select_brwsr != '':
                webbrowser.get(select_brwsr).open(filename, new=2)
            else:
                wx.MessageBox('Problem Locating Web Browser',
                              'Error', wx.OK)'''

    def OnPDF(self, evt):
        PDFFrm(self)

    def OnMerge(self, evt):
        MergeFrm(self)

    def OnITS(self, evt):
        BldTrvlSht(self)

    def OnConvert(self, evt):
        # select the html file to convert
        wildcard = "HTML file (*.html)|*.html"
        msg = 'Select HTML to Convert'
        filename = self.FylDilog(wildcard, msg, wx.FD_OPEN |
                                 wx.FD_FILE_MUST_EXIST)

        if filename != 'No File':

            with open(filename) as f:
                soup = BeautifulSoup(f, "html.parser")

            id_txt = []
            vl_txt = []
            id_txtar = []
            vl_txtar = []
            chk_chkd = []

            # get the name of the form for proper selection of pdf formate
            if [ttl.get_text() for ttl in soup.select('h1')] == []:
                msg = 'Program can only convert the following HTML Documents:\
                \n\t\tHydrostatic Test Report,\n\
                Hydrostatic Test Waiver,\n\
                Nonconformance Report,\n\
                Scope Of Work,\n\
                Material Substitution or\n\
                Request For New Specification'
                dlg = wx.MessageDialog(self, message=msg,
                                       caption="Invalid Document",
                                       style=wx.OK)
                dlg.ShowModal()
                dlg.Destroy()
                return
            else:
                titl = [ttl.get_text() for ttl in soup.select('h1')][0]

            # get all the text input values
            for item in soup.find_all("input", {"type": "text"}):
                id_txt.append(item.get('id'))
                if item.get('value') is None:
                    vl_txt.append('TBD')
                else:
                    vl_txt.append(item.get('value'))
            txt_inpt = dict(zip(id_txt, vl_txt))
            txt_boxes = txt_inpt

            # get all the text area input
            for item in soup.find_all('textarea'):
                id_txtar.append(item.get('id'))
                if item.contents != []:
                    vl_txtar.append(item.contents[0])
                else:
                    vl_txtar.append('')
            txtar_inpt = dict(zip(id_txtar, vl_txtar))
            txt_area = txtar_inpt

            # collect the checkboxes which are checked only
            for item in soup.find_all('input', checked=True):
                chk_chkd.append(item.get('id'))
            chkd_boxes = chk_chkd

            # open dialog to specify new pdf file name
            wildcrd = "PDF file (*.pdf)|*.pdf"
            mssg = 'Save new PDF'
            pdf_file = self.FylDilog(wildcrd, mssg, wx.FD_SAVE)
            # if pdf extention is not specified add it to the file name
            if pdf_file.find(".pdf") == -1:
                pdf_file = pdf_file + '.pdf'

            # use the form title to select the proper convertion program
            if titl.find('Nonconformance') != -1:
                from NCR_pdf import NCRForm
                NCRForm(pdf_file, txt_boxes, txt_area, chkd_boxes).create()
            elif titl.find('Specification') != -1:
                from RPS_pdf import RPSForm
                RPSForm(pdf_file, txt_boxes, txt_area, chkd_boxes).create()
            elif titl.find('Scope') != -1:
                from SOW_pdf import SOWForm
                SOWForm(pdf_file, txt_boxes, txt_area, chkd_boxes).create()
            elif titl.find('Material') != -1:
                from MSR_pdf import MSRForm
                MSRForm(pdf_file, txt_boxes, txt_area, chkd_boxes).create()
            elif titl.find('Waiver') != -1:
                from HTW_pdf import HTWForm
                HTWForm(pdf_file, txt_boxes, txt_area, chkd_boxes).create()
            elif titl.find('Hydrostatic') != -1:
                from HTR_pdf import HTRForm
                HTRForm(pdf_file, txt_boxes, txt_area, chkd_boxes).create()
            else:
                msg = 'Program can only convert the following HTML Documents:\n\
                Hydrostatic Test Report,\n\
                Hydrostatic Test Waiver,\n\
                Nonconformance Report,\n\
                Scope Of Work,\n\
                Material Substitution or\n\
                Request For New Specification'
                dlg = wx.MessageDialog(self, message=msg,
                                       caption="Invalid Document",
                                       style=wx.OK)
                dlg.ShowModal()
                dlg.Destroy()

    def OnDelete(self, evt):
        wildcard = "HTML file (*.html)|*.html|"\
                   "PDF files (*.pdf)|*.pdf"
        msg = 'Select File to Delete'
        filename = self.FylDilog(wildcard, msg, wx.FD_OPEN |
                                 wx.FD_FILE_MUST_EXIST)
        if filename != 'No File':
            msg = 'Confirm deletion of file?\n' + filename
            dlg = wx.MessageDialog(
                self, message=msg, caption="Question",
                style=wx.OK | wx.CANCEL | wx.ICON_QUESTION)
            if dlg.ShowModal() == wx.ID_OK:
                os.remove(filename)

    def OnHelp(self, evt):
        msg = ('''
    \tThis form is designed to review, modify and build new pipe
     specifications for a given commodity property.\n
    \tIt is tended to be used only by the administrator(s) of the
    pipe specification data.  If you are not the administrator you
    should close this form and request access to the user interface.\n
    \tTo help search for a specific commodity property the sort
    order of the Commodity Code and Pipe Material Spec columns
    can be changed by double clicking the table header.  You can
    also search based on words in the commodity description.
    Simply type word or phrase in the Commodity Description box then
    press Search Description.
    ''')

        dlg = wx.MessageDialog(self, message=msg,
                               caption='Use of Form',
                               style=wx.ICON_INFORMATION | wx.STAY_ON_TOP
                               | wx.CENTRE)
        dlg.ShowModal()
        dlg.Destroy()

    def OnAbout(self, evt):
        msg = ('''
    This program was written as open source, the program code can
    be used downloaded and distributed freely.\n
    Please note; all data was entered for demonstration use only.\n
    If the user wishes data can be input at a cost either as data
    supplied by the user or developed on a fee base.\n
    kphprojects@gmail.com''')
        wx.MessageBox(msg, 'Title', wx.OK)

    def OnRNSImp(self, evt):
        RPSImport(self)

    def OnMSRImp(self, evt):
        MSRImport(self)

    def OnNCRImp(self, evt):
        NCRImport(self)

    def OnCalcs(self, evt):
        CalcFrm(self)

    def FylDilog(self, wildcard, msg, stl):
        self.currentDirectory = os.getcwd()
        path = 'No File'
        dlg = wx.FileDialog(
            self, message=msg,
            defaultDir=self.currentDirectory,
            defaultFile="",
            wildcard=wildcard,
            style=stl
            )
        if dlg.ShowModal() == wx.ID_OK:
            path = dlg.GetPaths()[0]
        dlg.Destroy()

        return path

    def merge_pdf(self, targetfile):
        import glob
        PDF_merge = PdfFileMerger()

        sorc = glob.glob('PDFtmp*.pdf')
        sorc.sort()
        for pdffile in sorc:
            PDF_merge.append(pdffile)
            os.remove(pdffile)

        with open(targetfile, 'wb') as fileobj:
            PDF_merge.write(fileobj)

    def OnGV(self, evt):
        BldValve(self, 'GateValve', ComdPrtyID=self.ComdPrtyID)

    def OnPG(self, evt):
        BldValve(self, 'PlugValve', ComdPrtyID=self.ComdPrtyID)

    def OnBL(self, evt):
        BldValve(self, 'BallValve', ComdPrtyID=self.ComdPrtyID)

    def OnGB(self, evt):
        BldValve(self, 'GlobeValve', ComdPrtyID=self.ComdPrtyID)

    def OnBU(self, evt):
        BldValve(self, 'ButterflyValve', ComdPrtyID=self.ComdPrtyID)

    def OnSWC(self, evt):
        BldValve(self, 'SwingCheckValve', ComdPrtyID=self.ComdPrtyID)

    def OnPC(self, evt):
        BldValve(self, 'PistonCheckValve', ComdPrtyID=self.ComdPrtyID)

    def OnFtg(self, evt):
        BldFtgs(self, 'Fittings', ComdPrtyID=self.ComdPrtyID)

    def OnBolt(self, evt):
        BldFst(self, 'Fasteners', ComdPrtyID=self.ComdPrtyID)

    def OnPaint(self, evt):
        BldLvl3(self, 'PaintSpec', ComdPrtyID=self.ComdPrtyID)

    def OnTube(self, evt):
        BldConduit(self, 'Tubing', ComdPrtyID=self.ComdPrtyID)

    def OnWeld(self, evt):
        BldWeld(self, 'WeldRequirements', ComdPrtyID=self.ComdPrtyID)

    def OnGskt(self, evt):
        BldLvl3(self, 'GasketPacks', ComdPrtyID=self.ComdPrtyID)

    def OnBrch(self, evt):
        BrchFrm(self, ComdPrtyID=self.ComdPrtyID)

    def OnInspct(self, evt):
        BldLvl3(self, 'InspectionPacks', ComdPrtyID=self.ComdPrtyID)

    def OnInsul(self, evt):
        BldInsul(self, 'Insulation', ComdPrtyID=self.ComdPrtyID)

    def OnOFlg(self, evt):
        BldFtgs(self, 'OrificeFlanges', ComdPrtyID=self.ComdPrtyID)

    def OnFlg(self, evt):
        BldFtgs(self, 'Flanges', ComdPrtyID=self.ComdPrtyID)

    def OnOlet(self, evt):
        BldFtgs(self, 'OLets', ComdPrtyID=self.ComdPrtyID)

    def OnUnion(self, evt):
        BldFtgs(self, 'Unions', ComdPrtyID=self.ComdPrtyID)

    def OnClmp(self, evt):
        BldFtgs(self, 'GrooveClamps', ComdPrtyID=self.ComdPrtyID)

    def OnSpecl(self, evt):
        BldSpc_Nts(self, 'Specials', ComdPrtyID=self.ComdPrtyID)

    def OnNotes(self, evt):
        BldSpc_Nts(self, 'Notes', ComdPrtyID=self.ComdPrtyID)

    def PrintFile(self, evt):
        import Report_General
        # function to call up the dialog to generate the file name
        # this will priint out all the commodity properties
        wildcard = 'PDF (*.pdf)|*.pdf'
        msg = 'Save Report as PDF.'

        filename = self.FylDilog(wildcard, msg,
                                 wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT)
        if filename != 'No File':
            if filename.find(".pdf") == -1:
                filename = filename + '.pdf'

            num_col = 8
            num_rows = 5
            Report_General.Report(self.tblname, self.data, self.colwdth,
                                  self.columnames, num_rows, num_col,
                                  filename).create_pdf()

    def Search_Restore(self):
        if self.Srch is True:
            self.b6.SetLabel("<-- Search\nDescription")
            self.Bind(wx.EVT_BUTTON, self.OnSearch, self.b6)
            self.RestoreBoxs()
            self.Srch = False
        else:
            self.b6.SetLabel("Restore\nTable")
            self.Bind(wx.EVT_BUTTON, self.OnRestoreTbl, self.b6)
            self.Srch = True

    def RestoreBoxs(self):
        self.cmbCode.ChangeValue('')
        self.cmbPipe.ChangeValue('')
        self.cmbFluid.ChangeValue('')
        self.textDP.ChangeValue('')
        self.textMinT.ChangeValue('')
        self.textMaxT.ChangeValue('')
        self.cmbEnd.ChangeValue('')
        self.txtDscrpt.ChangeValue('')
        self.textSpec.ChangeValue('')
        self.chkPend.SetValue(False)
        self.notes.ChangeValue('')

        self.textMaxT.SetHint('Deg F')
        self.textMinT.SetHint('Deg F')
        self.textDP.SetHint('psig')

        for i in range(24):
            self.btns[i].Enable(False)

        self.Bind(wx.EVT_BUTTON, self.OnDevSpec, self.b2)
        self.NewSpec = False

    def RestoreTbl(self):
        self.data = Dbase().Dsqldata(self.Dsql)
        self.model = DataMods(self.tblname, self.data)
        self.dvc.AssociateModel(self.model)
        self.dvc.Refresh
        self.Search_Restore()

    def ValData(self):
        digstr = []
        DataStrg = []
        NoData = 0
        DialogStr = ''

        # before saving record check that all the
        # required data has been entered
        if self.cmbCode.GetValue() == '':
            # each record is assigned value
            # of binary number 000001 = 1 for no data
            NoData = 1
        else:
            # develop list of each data item
            DataStrg.append(self.cmbCode.GetValue())
        # next binary number level 000010 = 2 etc.
        if self.txtDscrpt.GetValue() == '':
            NoData = 2 + NoData
        else:
            DataStrg.append(self.txtDscrpt.GetValue())

        if self.cmbPipe.GetValue() == '':
            NoData = 4 + NoData
        else:
            DataStrg.append(self.cmbPipe.GetValue())

        if self.cmbFluid.GetValue() == '':
            NoData = 8 + NoData
        else:
            DataStrg.append(self.cmbFluid.GetValue())

        if self.textDP.GetValue() == '':
            NoData = 16 + NoData
        else:
            DataStrg.append(self.textDP.GetValue())

        if self.textMinT.GetValue() == '':
            NoData = 32 + NoData
        else:
            DataStrg.append(self.textMinT.GetValue())

        if self.textMaxT.GetValue() == '':
            NoData = 64 + NoData
        else:
            DataStrg.append(self.textMaxT.GetValue())

        if self.cmbEnd.GetValue() == '':
            NoData = 128 + NoData
        else:
            DataStrg.append(self.cmbEnd.GetValue())

        if self.textSpec.GetValue() == '' or \
                self.textSpec.GetValue() == 'None':
            if self.cmbCode.GetValue() and self.cmbPipe.GetValue():
                self.textSpec.ChangeValue(self.cmbCode.GetValue() + '-' +
                                          self.cmbPipe.GetValue())
                DataStrg.append(self.textSpec.GetValue())
            else:
                NoData = 256 + NoData
        else:
            DataStrg.append(self.textSpec.GetValue())

        # check that all the records are present, if not use
        # the sum of the binary numbers to see which are missing
        if len(DataStrg) < 9:
            # use the binary numbers as keys for the DataBxs dictionary
            DataBxs = {1: 'Commodity Code', 2: 'Commodity Description',
                       3: 'Pipe Material Spec', 4: 'Fluid Category',
                       5: 'Design Pressure', 6: 'Min. Design Temp',
                       7: 'Max. Design Temp', 8: 'End Connections',
                       9: 'Pipe Spec Code'}
            binry = str('{0:09b}'.format(NoData))
            digstr = [pos for pos, char in enumerate(binry) if char == '1']
            for dig in digstr:
                DialogStr = DataBxs[9-dig] + ',\n' + DialogStr
            wx.MessageBox('Value(s) needed for;\n' + DialogStr,
                          'Missing Data', wx.OK | wx.ICON_INFORMATION)
        else:
            self.AddRec(DataStrg)

    def AddRec(self, DataStrg):
        # set a default to cancel changes as safety net
        SQL_step = 3

        # last add the two none mandatory fields Pending and notes
        if self.chkPend.GetValue():
            DataStrg.append(1)
        else:
            DataStrg.append(0)
        DataStrg.append(self.notes.GetValue())
        choices = ['Save this as a new commodity property',
                   '''Update the existing the existing
                    commodity property with this data''']

        # use a SingleChioce dialog to determine if data is to
        # be a new record or edited record
        SQL_Dialog = wx.SingleChoiceDialog(
            self, '''The commodity property data has changed do you
             want\nto save as new record or update existing?''',
            'Commodity Property Has Changed', choices,
            style=wx.CHOICEDLG_STYLE)
        if SQL_Dialog.ShowModal() == wx.ID_OK:
            SQL_step = SQL_Dialog.GetSelection()
        SQL_Dialog.Destroy()

        if SQL_step == 1:
            wx.MessageBox(
                '''Updating the record will maintain the
                 original Commodity Code.\n'''
                "All other data will be updated as needed.", "Update Record",
                wx.OK | wx.ICON_INFORMATION)

            UpQuery = ('UPDATE CommodityProperties SET Design_Pressure = ' +
                       DataStrg[4] + ', Minimum_Design_Temperature = ' +
                       DataStrg[5] + ', Maximum_Design_Temperature = ' +
                       DataStrg[6] + ', Pending = "' + str(DataStrg[9]) +
                       '", Note = "' + DataStrg[10] + '", Pipe_Code = "' +
                       DataStrg[8] + '''", End_Connection = (SELECT ID FROM
                        CommodityEnds WHERE Commodity_End = "''' +
                       DataStrg[7] + '''"), Pipe_Material_Code = (SELECT
                         Material_Spec_ID FROM PipeMaterialSpec WHERE
                         Pipe_Material_Spec = "''' +
                       DataStrg[2] + '''"), Fluid_Category = (SELECT Fluid_ID
                        FROM FluidCategory WHERE Designation = "''' +
                       DataStrg[3] + '") WHERE CommodityPropertyID = '
                       + str(self.ComdPrtyID))

            Dbase().TblEdit(UpQuery)

        elif SQL_step == 0:
            # if it is to be a new record first check to see that
            # a duplicate Commodity Code + Pipe Material Spec does not exist
            # if it does then open dialogue to state duplicate commodity
            # property could not be saved
            ChkQuery = ('''SELECT Commodity_Code, Pipe_Material_Code FROM
                         CommodityProperties WHERE Commodity_Code = "''' +
                        DataStrg[0] + '''" AND Pipe_Material_Code = (SELECT
                         Material_Spec_ID FROM PipeMaterialSpec WHERE
                         Pipe_Material_Spec = "''' + DataStrg[2] + '")')

            if Dbase().Dsqldata(ChkQuery):
                wx.MessageBox('''The combination of Commodity Code\nand
                               Pipe Material Specification\nalready exist
                               and cannot be added!''', "Cannot Add Record",
                              wx.OK | wx.ICON_INFORMATION)
                return

            # get the next available autoincremented value for the table
            IDE = cursr.execute(
                "SELECT MAX(CommodityPropertyID) FROM "
                + self.tblname).fetchone()[0]
            if IDE is None:
                IDE = 0
            IDE = int(IDE+1)
            del DataStrg[1]

            UpQuery1 = ('''SELECT Fluid_ID FROM FluidCategory
                         WHERE Designation = "''' + DataStrg[2] + '"')
            UpQuery2 = ('''SELECT ID FROM CommodityEnds WHERE
                         Commodity_End = "''' + DataStrg[6] + '"')
            UpQuery3 = ('''SELECT Material_Spec_ID FROM PipeMaterialSpec WHERE
                         Pipe_Material_Spec = "''' + DataStrg[1] + '"')

            FluidID = Dbase().Dsqldata(UpQuery1)[0][0]

            ComdEndID = Dbase().Dsqldata(UpQuery2)[0][0]

            SpecID = Dbase().Dsqldata(UpQuery3)[0][0]

            UpQuery = ('''INSERT INTO CommodityProperties(CommodityPropertyID,
                        Commodity_Code, Design_Pressure,''' +
                       '''Minimum_Design_Temperature,
                        Maximum_Design_Temperature, Pipe_Code, Note,
                        Pending, ''' + '''Fluid_Category, End_Connection,
                        Pipe_Material_Code) VALUES (''' + str(IDE) + ',"'
                       + DataStrg[0] + '",' + DataStrg[3] + ',' + DataStrg[4]
                       + ',' + DataStrg[5] + ',"' + DataStrg[7] + '","'
                       + DataStrg[9] + '",' + str(DataStrg[8]) + ',"' +
                       FluidID + '",' + str(ComdEndID) + ',"' +
                       str(SpecID) + '")')

            Dbase().TblEdit(UpQuery)

            ValueList = []
            New_ID = cursr.execute(
                "SELECT MAX(Pipe_Spec_ID) FROM PipeSpecification").\
                fetchone()
            if New_ID[0] is None:
                Max_ID = '1'
            else:
                Max_ID = str(New_ID[0]+1)
            colinfo = Dbase().Dcolinfo('PipeSpecification')
            for n in range(0, len(colinfo)-2):
                ValueList.append(None)

            num_vals = ('?,'*len(colinfo))[:-1]
            ValueList.insert(0, Max_ID)
            ValueList.insert(1, str(IDE))

            UpQuery = ("INSERT INTO PipeSpecification VALUES ("
                       + num_vals + ")")
            Dbase().TblEdit(UpQuery, ValueList)

        # if cancel is selected do nothing
        elif SQL_step == 3:
            return

        self.data = Dbase().Dsqldata(self.Dsql)
        self.model = DataMods(self.tblname, self.data)
        self.dvc.AssociateModel(self.model)
        self.dvc.Refresh

    def OnSelectP(self, evt):
        if self.NewSpec is False:
            self.cmbPipe.ChangeValue('')
        if self.cmbCode.GetValue() and self.cmbPipe.GetValue():
            self.textSpec.ChangeValue(
                self.cmbCode.GetValue() + '-' + self.cmbPipe.GetValue())
        else:
            self.textSpec.ChangeValue('')

    def OnSelectF(self, evt):
        if self.NewSpec is False:
            self.cmbFluid.ChangeValue('')

    def OnSelectE(self, evt):
        if self.NewSpec is False:
            self.cmbEnd.ChangeValue('')

    def OnSelectC(self, evt):
        if self.NewSpec is False:
            self.cmbCode.ChangeValue('')
        else:
            tbl_name = 'CommodityCodes'
            query2 = ('SELECT * FROM ' + tbl_name +
                      ' WHERE Commodity_Code = "' +
                      self.cmbCode.GetValue() + '"')
            newvalue = Dbase().Dsqldata(query2)
            if newvalue != []:
                self.txtDscrpt.ChangeValue(newvalue[0][1])

    def OnRestoreBoxs(self, evt):
        self.RestoreBoxs()

    def OnRestoreTbl(self, evt):
        self.RestoreTbl()

    def OnAddRec(self, evt):
        self.ValData()
        self.NewSpec = False

    def OnSearch(self, evt):
        srchstrg = ''
        qry = ('''SELECT * FROM CommodityCodes WHERE Commodity_Description
                LIKE '%''' + self.txtDscrpt.GetValue() + "%' COLLATE NOCASE")
        srchdata = Dbase().Search(qry)
        if srchdata:
            n = len(srchdata)
            m = 1
            for item in srchdata:
                srchstrg = srchstrg + "'" + item[0] + "'"
                if m < n:
                    srchstrg = srchstrg + " OR a.Commodity_Code = "
                m += 1
            ShQuery = self.Dsql + '\n WHERE a.Commodity_Code = ' + srchstrg
            srchdata = Dbase().Search(ShQuery)

        else:
            srchdata = []

        self.data = srchdata
        self.model = DataMods(self.tblname, self.data)
        self.dvc.AssociateModel(self.model)
        self.dvc.Refresh
        self.Search_Restore()

    def OnDeleteRow(self, evt):
        # Remove the selected row from the model. The model will take care
        # of notifying the view (and any other observers) that the change has
        # happened.
        txt1 = 'Deletion of this commmodity property'
        txt2 = '\nwill remove it from the pipe specification'
        txt3 = '\n\nComponents related to the commodity will'
        txt4 = '\nhowever not be affected.'
        msg = txt1 + txt2 + txt3 + txt4
        dlg = wx.MessageDialog(
            self, message=msg, caption="Confirm Deletion",
            style=wx.OK | wx.CANCEL | wx.ICON_QUESTION)
        if dlg.ShowModal() == wx.ID_OK:
            item = self.dvc.GetSelection()
            rowD = self.model.GetRow(item)
            Dbase().TblDelete('PipeSpecification', self.data[rowD][0],
                              'Commodity_Property_ID')
            self.RestoreBoxs()
            self.model.DeleteRow(rowD, 'CommodityPropertyID')

    def OnDevSpec(self, evt):
        self.Bind(wx.EVT_BUTTON, self.OnAddRec, self.b2)
        self.NewSpec = True

    def OnGridSelect(self, evt):
        item = self.dvc.GetSelection()
        rowGS = self.model.GetRow(item)

        self.ComdPrtyID = self.data[rowGS][0]
        self.cmbCode.ChangeValue(self.data[rowGS][1])
        self.txtDscrpt.ChangeValue(self.data[rowGS][2])
        self.cmbPipe.ChangeValue(self.data[rowGS][3])
        self.cmbFluid.ChangeValue(self.data[rowGS][4])
        self.textDP.ChangeValue(str(self.data[rowGS][5]))
        self.textMinT.ChangeValue(str(self.data[rowGS][6]))
        self.textMaxT.ChangeValue(str(self.data[rowGS][7]))
        self.cmbEnd.ChangeValue(self.data[rowGS][8])
        self.textSpec.ChangeValue(str(self.data[rowGS][9]))
        self.chkPend.SetValue(self.data[rowGS][10])
        self.notes.ChangeValue(str(self.data[rowGS][11]))

        for i in range(24):
            self.btns[i].Enable()

        self.b6.SetLabel("<-- Search\nDescription")
        self.Bind(wx.EVT_BUTTON, self.OnSearch, self.b6)
        self.Srch = False
        self.data = Dbase().Dsqldata(self.Dsql)
        self.model = DataMods(self.tblname, self.data)
        self.dvc.AssociateModel(self.model)
        self.dvc.Refresh

    def OnAddCode(self, evt):
        boxnums = [0]
        CmbLst1(self, 'CommodityCodes')
        self.ReFillList('CommodityCodes', boxnums)

    def OnAddFld(self, evt):
        boxnums = [2]
        CmbLst1(self, 'FluidCategory')
        self.ReFillList('FluidCategory', boxnums)

    def OnAddSpec(self, evt):
        boxnums = [1]
        PipeMtrSpc(self, 'PipeMaterialSpec')
        self.ReFillList('PipeMaterialSpec', boxnums)

    def OnAddEnd(self, evt):
        boxnums = [3]
        CmbLst1(self, 'CommodityEnds')
        self.ReFillList('CommodityEnds', boxnums)

    def ReFillList(self, cmbtbl, boxnums):
        for n in boxnums:
            self.lc = self.lctrls[n]
            self.lc.DeleteAllItems()
            index = 0
            ReFillQuery = 'SELECT * FROM "' + cmbtbl + '"'
            for values in Dbase().Dsqldata(ReFillQuery):
                col = 0
                for value in values:
                    if col == 0:
                        self.lc.InsertItem(index, str(value))
                    else:
                        self.lc.SetItem(index, col, str(value))
                    col += 1
                index += 1


class PDFFrm(wx.Frame):
    def __init__(self, parent):
        super(PDFFrm, self).__init__(parent)
        from wx.lib.pdfviewer import pdfViewer, pdfButtonPanel
        self.Maximize(True)

        self.parent = parent
        self.Bind(wx.EVT_CLOSE, self.OnCloseFrm)

        hsizer = wx.BoxSizer(wx.HORIZONTAL)
        vsizer = wx.BoxSizer(wx.VERTICAL)
        self.buttonpanel = pdfButtonPanel(self, wx.ID_ANY,
                                          wx.DefaultPosition,
                                          wx.DefaultSize, 0)
        vsizer.Add(self.buttonpanel, 0,
                   wx.GROW | wx.LEFT | wx.RIGHT | wx.TOP, 5)
        self.viewer = pdfViewer(self, wx.ID_ANY, wx.DefaultPosition,
                                wx.DefaultSize, wx.HSCROLL |
                                wx.VSCROLL | wx.SUNKEN_BORDER)
        vsizer.Add(self.viewer, 1, wx.GROW | wx.LEFT | wx.RIGHT |
                   wx.BOTTOM, 5)
        loadbutton = wx.Button(self, wx.ID_ANY, "Load PDF file",
                               wx.DefaultPosition, wx.DefaultSize, 0)
        loadbutton.SetForegroundColour((255, 0, 0))
        vsizer.Add(loadbutton, 0, wx.ALIGN_CENTER | wx.ALL, 5)
        hsizer.Add(vsizer, 1, wx.GROW | wx.ALL, 5)
        self.SetSizer(hsizer)
        self.SetAutoLayout(True)

        # introduce buttonpanel and viewer to each other
        self.buttonpanel.viewer = self.viewer
        self.viewer.buttonpanel = self.buttonpanel

        self.Bind(wx.EVT_BUTTON, self.OnLoadButton, loadbutton)

        self.CenterOnParent()
        self.GetParent().Enable(False)
        self.Show(True)
        self.__eventLoop = wx.GUIEventLoop()
        self.__eventLoop.Run()

    def OnLoadButton(self, event):
        dlg = wx.FileDialog(self, wildcard="*.pdf")
        if dlg.ShowModal() == wx.ID_OK:
            wx.BeginBusyCursor()
            self.viewer.LoadFile(dlg.GetPath())
            wx.EndBusyCursor()
        dlg.Destroy()

    def OnCloseFrm(self, evt):
        self.GetParent().Enable(True)   # add for child form
        self.__eventLoop.Exit()        # add for child form
        self.Destroy()


class CalcFrm(wx.Frame):
    def __init__(self, parent):
        ttl = 'Wall Thickness and Hydro-Test calculation'
        super(CalcFrm, self).__init__(parent,
                                      title=ttl,
                                      size=(580, 720),
                                      style=wx.DEFAULT_FRAME_STYLE &
                                      ~(wx.RESIZE_BORDER | wx.MAXIMIZE_BOX |
                                        wx.MINIMIZE_BOX))
        self.lctrls = []
        self.parent = parent

        self.FrmSizer = wx.BoxSizer(wx.VERTICAL)
        self.Bind(wx.EVT_CLOSE, self.OnClose)

        self.pnl = CalcPnl(self)
        self.FrmSizer.Add(self.pnl, 1, wx.EXPAND)
        self.FrmSizer.Add((35, 10))
        self.FrmSizer.Add((10, 20))
        self.pnl.b4.Bind(wx.EVT_BUTTON, self.OnClose)
        self.SetSizer(self.FrmSizer)

        # add these 5 following lines to child parent form
        self.CenterOnParent()
        self.GetParent().Enable(False)
        self.Show(True)
        self.__eventLoop = wx.GUIEventLoop()
        self.__eventLoop.Run()

    def OnClose(self, evt):
        self.GetParent().Enable(True)   # add for child form
        self.__eventLoop.Exit()        # add for child form
        self.Destroy()


class CalcPnl(sc.ScrolledPanel):

    def __init__(self, parent):
        super(CalcPnl, self).__init__(parent, size=(560, 630))

        self.lctrls = []

        font = wx.Font(10, wx.DEFAULT, wx.NORMAL, wx.NORMAL)
        font1 = wx.Font(12, wx.DEFAULT, wx.NORMAL, wx.BOLD)

        self.Sizer = wx.BoxSizer(wx.VERTICAL)

        DTSizer = wx.BoxSizer(wx.HORIZONTAL)
        DTbox_border = wx.StaticBox(self, -1)
        DTbox = wx.StaticBoxSizer(DTbox_border, wx.VERTICAL)
        text = wx.StaticText(self, label="Allowed Stress")
        text.SetForegroundColour('red')
        DTbox.Add(text, 0, wx.ALIGN_LEFT)
        Stressnote = wx.StaticText(self, label='Sd = ', style=wx.ALIGN_LEFT)
        self.Stresstxt = wx.TextCtrl(self, size=(80, -1), value='',
                                     style=wx.TE_CENTER)
        self.Stresstxt.SetForegroundColour('red')
        self.Stresstxt.SetFont(font)
        DTempnote = wx.StaticText(self, label='psi @ Design Temperature',
                                  style=wx.ALIGN_LEFT)
        DTSizer.Add(DTbox, 0, wx.BOTTOM | wx.LEFT, border=10)
        DTSizer.Add(Stressnote, 0, wx.LEFT | wx.ALIGN_CENTER_VERTICAL,
                    border=20)
        DTSizer.Add(self.Stresstxt, 0, wx.LEFT | wx.ALIGN_CENTER, border=10)
        DTSizer.Add(DTempnote, 0, wx.LEFT | wx.ALIGN_CENTER, border=25)

        HTSizer = wx.BoxSizer(wx.HORIZONTAL)
        Hydronote = wx.StaticText(self, label='Sh = ', style=wx.ALIGN_LEFT)
        self.Hydrotxt = wx.TextCtrl(self, size=(80, -1), value='',
                                    style=wx.TE_CENTER)
        self.Hydrotxt.SetForegroundColour('red')
        self.Hydrotxt.SetFont(font)
        HTnote = wx.StaticText(self, label='psi @ Hydro Test Temperature',
                               style=wx.ALIGN_LEFT)
        HTSizer.Add(Hydronote, 0, wx.LEFT | wx.ALIGN_CENTER_VERTICAL,
                    border=135)
        HTSizer.Add(self.Hydrotxt, 0, wx.LEFT | wx.ALIGN_CENTER, border=10)
        HTSizer.Add(HTnote, 0, wx.LEFT | wx.ALIGN_CENTER, border=25)

        # draw a line between upper and lower section
        ln1 = wx.StaticLine(self, 0, size=(560, 2), style=wx.LI_VERTICAL)
        ln1.SetBackgroundColour('Black')

        DataSizer = wx.BoxSizer(wx.HORIZONTAL)
        Databox_border = wx.StaticBox(self, -1)
        Databox = wx.StaticBoxSizer(Databox_border, wx.VERTICAL)
        text = wx.StaticText(self, label="Design Data")
        text.SetForegroundColour('red')
        Databox.Add(text, 0, wx.ALIGN_LEFT)
        DPnote = wx.StaticText(self, label='Design Pressure Pd = ',
                               style=wx.ALIGN_LEFT)
        self.DPtxt = wx.TextCtrl(self, size=(80, -1), value='',
                                 style=wx.TE_CENTER)
        self.DPtxt.SetForegroundColour('red')
        self.DPtxt.SetFont(font)
        Unitsnote = wx.StaticText(self, label='psig', style=wx.ALIGN_LEFT)
        b3 = wx.Button(self, label="View Flange\nRatings")
        self.Bind(wx.EVT_BUTTON, self.OnFlange, b3)
        DataSizer.Add(Databox, 0, wx.BOTTOM | wx.LEFT, border=10)
        DataSizer.Add(DPnote, 0, wx.LEFT | wx.ALIGN_CENTER_VERTICAL, border=20)
        DataSizer.Add(self.DPtxt, 0, wx.LEFT | wx.ALIGN_CENTER, border=10)
        DataSizer.Add(Unitsnote, 0, wx.LEFT | wx.ALIGN_CENTER, border=25)
        DataSizer.Add(b3, 0, wx.LEFT | wx.ALIGN_CENTER, 15)

        # draw a line between upper and lower section
        ln2 = wx.StaticLine(self, 0, size=(560, 2), style=wx.LI_VERTICAL)
        ln2.SetBackgroundColour('Black')

        self.cmbsizer1 = wx.BoxSizer(wx.HORIZONTAL)
        noteCor = wx.StaticText(
            self, label='Corrosion Allowance c (inches) = ',
            style=wx.ALIGN_LEFT)
        noteCor.SetFont(font)
        self.cmbCor = wx.ComboCtrl(self, pos=(10, 10), size=(100, -1),
                                   style=wx.CB_READONLY)
        self.cmbCor.SetPopupControl(ListCtrlComboPopup('CorrosionAllowance',
                                                       showcol=1,
                                                       lctrls=self.lctrls))
        self.cmbsizer1.Add(noteCor, 0, wx.LEFT | wx.ALIGN_CENTER_VERTICAL, 15)
        self.cmbsizer1.Add(self.cmbCor, 0, wx.LEFT |
                           wx.ALIGN_CENTER_VERTICAL, 5)

        self.cmbsizer2 = wx.BoxSizer(wx.HORIZONTAL)
        noteOD = wx.StaticText(
            self, label='Pipe OD (inches) = ',
            style=wx.ALIGN_LEFT)
        noteOD.SetFont(font)
        noteWall = wx.StaticText(self, label='Wall Thk.',
                                 style=wx.ALIGN_LEFT)
        noteWall.SetFont(font)
        self.cmbOD = wx.ComboCtrl(self, pos=(10, 10),
                                  size=(100, -1), style=wx.CB_READONLY)
        self.cmbOD.SetPopupControl(ListCtrlComboPopup('Pipe_OD',
                                                      showcol=1,
                                                      lctrls=self.lctrls))
        self.Bind(wx.EVT_COMBOBOX_CLOSEUP, self.CutDpth, self.cmbOD)
        FillQuery = (
            "SELECT Wall,Sch FROM PipeDimensions WHERE Size = '" +
            self.cmbOD.GetValue() + "'")
        self.cmbSch = wx.ComboCtrl(self, pos=(10, 10),
                                   size=(100, -1), style=wx.CB_READONLY)
        self.cmbSch.SetPopupControl(
            ListCtrlComboPopup('PipeDimensions', FillQuery,
                               showcol=0,
                               lctrls=self.lctrls))
        self.Bind(wx.EVT_COMBOBOX_CLOSEUP, self.WallThk, self.cmbSch)
        self.cmbSch.Disable()
        self.noteSch = wx.TextCtrl(self, size=(100, -1), value='Sch',
                                   style=wx.TE_READONLY | wx.TE_LEFT)
        self.cmbsizer2.Add(noteOD, 0, wx.LEFT | wx.ALIGN_CENTER_VERTICAL, 15)
        self.cmbsizer2.Add(self.cmbOD, 0, wx.LEFT |
                           wx.ALIGN_CENTER_VERTICAL, 5)
        self.cmbsizer2.Add(noteWall, 0, wx.LEFT | wx.ALIGN_CENTER_VERTICAL, 15)
        self.cmbsizer2.Add(self.cmbSch, 0, wx.LEFT |
                           wx.ALIGN_CENTER_VERTICAL, 5)
        self.cmbsizer2.Add(self.noteSch, 0, wx.LEFT |
                           wx.ALIGN_CENTER_VERTICAL, 15)

        radioboxsizer1 = wx.BoxSizer(wx.HORIZONTAL)
        noteDpth = wx.StaticText(self, label='Specify pipe end: ',
                                 style=wx.ALIGN_LEFT)
        noteDpth.SetFont(font)
        rdBoxSzr1 = wx.BoxSizer(wx.VERTICAL)
        lblDpth = wx.StaticText(self, label='Cut Depth Cd',
                                style=wx.ALIGN_LEFT)
        lblDpth.SetFont(font)
        self.txtThrd = wx.TextCtrl(self, size=(80, -1), value='',
                                   style=wx.TE_CENTER | wx.TE_READONLY)
        self.txtThrd.SetFont(font)
        self.txtGrv = wx.TextCtrl(self, size=(80, -1), value='',
                                  style=wx.TE_CENTER | wx.TE_READONLY)
        self.txtGrv.SetFont(font)
        rdBoxSzr1.Add(lblDpth, 0, wx.TOP | wx.LEFT, 10)
        rdBoxSzr1.Add((15, 10))
        rdBoxSzr1.Add(self.txtThrd, 0, wx.TOP | wx.LEFT | wx.BOTTOM, 5)
        rdBoxSzr1.Add(self.txtGrv, 0, wx.ALL, 5)

        self.rdBox1 = wx.RadioBox(self,
                                  choices=['Plain', '\nThreaded\n', 'Grooved'],
                                  majorDimension=1, style=wx.RA_SPECIFY_COLS)
        self.rdBox1.SetSelection(0)
        self.Bind(wx.EVT_RADIOBOX, self.CutDpth, self.rdBox1)
        radioboxsizer1.Add(noteDpth, 0, wx.LEFT | wx.ALIGN_CENTER_VERTICAL, 10)
        radioboxsizer1.Add((15, 10))
        radioboxsizer1.Add(self.rdBox1, 0, wx.ALL | wx.ALIGN_TOP, 5)
        radioboxsizer1.Add((5, 10))
        radioboxsizer1.Add(rdBoxSzr1, 0, wx.TOP | wx.ALIGN_TOP, 8)

        radioboxsizer = wx.BoxSizer(wx.HORIZONTAL)
        noteEff = wx.StaticText(
            self, label='E = Pipe Quality Factor, default values: ',
            style=wx.ALIGN_LEFT)
        noteEff.SetFont(font)
        rdBoxSzr = wx.BoxSizer(wx.VERTICAL)
        self.ERWtxt = wx.TextCtrl(self, size=(80, -1), value='.085',
                                  style=wx.TE_CENTER)
        self.ERWtxt.SetFont(font)
        self.Smlstxt = wx.TextCtrl(self, size=(80, -1), value='1.00',
                                   style=wx.TE_CENTER)
        self.Smlstxt.SetFont(font)
        rdBoxSzr.Add(self.ERWtxt, 0, wx.ALL, 5)
        rdBoxSzr.Add(self.Smlstxt, 0, wx.ALL, 5)

        self.rdBox = wx.RadioBox(self,
                                 choices=['ERW', '\nSeamless\n'],
                                 majorDimension=1, style=wx.RA_SPECIFY_COLS)
        self.rdBox.SetSelection(0)
        radioboxsizer.Add(noteEff, 0, wx.LEFT | wx.ALIGN_CENTER_VERTICAL, 10)
        radioboxsizer.Add(rdBoxSzr, 0, wx.TOP | wx.ALIGN_TOP, 8)
        radioboxsizer.Add(self.rdBox, 0, wx.ALL | wx.ALIGN_TOP, 5)

        self.StrFactsizer = wx.BoxSizer(wx.HORIZONTAL)
        noteStrsFact = wx.StaticText(
            self, label=('''Y = stress - temperature compensation factor,
                          default value: '''), style=wx.ALIGN_LEFT)
        noteStrsFact.SetFont(font)
        self.StrsFact = wx.TextCtrl(self, size=(80, -1), value='0.40',
                                    style=wx.TE_CENTER)
        self.StrsFact.SetFont(font)
        self.StrFactsizer.Add(noteStrsFact, 0, wx.LEFT |
                              wx.ALIGN_CENTER_VERTICAL, 15)
        self.StrFactsizer.Add(self.StrsFact, 0, wx.ALL, 10)

        # draw a line between upper and lower section
        ln3 = wx.StaticLine(self, 0, size=(560, 2), style=wx.LI_VERTICAL)
        ln3.SetBackgroundColour('Black')

        notesizer = wx.BoxSizer(wx.VERTICAL)
        note1 = wx.StaticText(
            self, label=('''Minimum Pipe Wall Thickness tm = (Pd * Pipe OD) / 2
                           (Sd * E + Pd * Y) + c + Cd'''), style=wx.ALIGN_LEFT)
        note1.SetFont(font)
        note1.SetForegroundColour(('red'))
        notesizer.Add(note1, 0, wx.TOP | wx.LEFT | wx.ALIGN_LEFT, border=10)

        CalcWallszr = wx.BoxSizer(wx.HORIZONTAL)
        self.b1 = wx.Button(self, label="Calculate Minimum Wall Thickness")
        self.b1.SetForegroundColour('green')
        self.Bind(wx.EVT_BUTTON, self.CalcWall, self.b1)
        self.noteCalcWall = wx.StaticText(self, label=' ', style=wx.ALIGN_LEFT)
        self.noteCalcWall.SetFont(font1)
        CalcWallszr.Add(self.b1, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        CalcWallszr.Add(self.noteCalcWall, 0, wx.ALL |
                        wx.ALIGN_CENTER_VERTICAL, 5)

        # draw a line between upper and lower section
        ln4 = wx.StaticLine(self, 0, size=(560, 2), style=wx.LI_VERTICAL)
        ln4.SetBackgroundColour('Black')

        notesizer1 = wx.BoxSizer(wx.VERTICAL)
        note3 = wx.StaticText(
            self, label='Required Hydro Test Pressure = 1.5 * Pd * Sh / Sd',
            style=wx.ALIGN_LEFT)
        note3.SetFont(font)
        note3.SetForegroundColour(('red'))
        notesizer1.Add(note3, 0, wx.TOP | wx.LEFT | wx.ALIGN_LEFT, border=10)

        CalcHydroszr = wx.BoxSizer(wx.HORIZONTAL)
        self.b2 = wx.Button(self, label="Calculate Hydro Test Pressure")
        self.b2.SetForegroundColour('green')
        self.Bind(wx.EVT_BUTTON, self.CalcHydro, self.b2)
        self.noteCalcHydro = wx.StaticText(self, label=' ',
                                           style=wx.ALIGN_LEFT)
        self.noteCalcHydro.SetFont(font1)
        CalcHydroszr.Add(self.b2, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        CalcHydroszr.Add(self.noteCalcHydro, 0, wx.ALL |
                         wx.ALIGN_CENTER_VERTICAL, 5)

        btnbox = wx.BoxSizer(wx.VERTICAL)
        self.b4 = wx.Button(self, label="Exit")
        self.b4.SetForegroundColour('red')
        btnbox.Add(self.b4, 0, wx.RIGHT | wx.ALIGN_CENTER, 15)

        self.Sizer.Add(DTSizer, 0, wx.ALL, 5)
        self.Sizer.Add(HTSizer, 0, wx.ALL, 5)
        self.Sizer.Add(ln1, 0, wx.ALIGN_LEFT)
        self.Sizer.Add(DataSizer, 0, wx.ALL, 5)
        self.Sizer.Add(ln2, 0, wx.ALIGN_LEFT)
        self.Sizer.Add(self.cmbsizer1, 0, wx.TOP | wx.ALIGN_LEFT, 15)
        self.Sizer.Add(self.cmbsizer2, 0, wx.TOP | wx.ALIGN_LEFT, 15)
        self.Sizer.Add(radioboxsizer1, 0, wx.ALL, 5)
        self.Sizer.Add(radioboxsizer, 0, wx.ALL, 5)
        self.Sizer.Add(self.StrFactsizer, 0, wx.ALL | wx.ALIGN_LEFT)
        self.Sizer.Add(ln3, 0, wx.ALIGN_LEFT)
        self.Sizer.Add(notesizer, 0, wx.ALIGN_LEFT)
        self.Sizer.Add(CalcWallszr, 0, wx.ALL, 5)
        self.Sizer.Add(ln4, 0, wx.ALIGN_LEFT)
        self.Sizer.Add(notesizer1, 0, wx.ALIGN_LEFT)
        self.Sizer.Add(CalcHydroszr, 0, wx.ALL, 5)
        self.Sizer.Add(btnbox, 0, wx.ALIGN_RIGHT)
        self.SetupScrolling()

    def CutDpth(self, evt):
        if self.cmbOD.GetValue() != '':
            if self.rdBox1.GetSelection() == 1:
                qry = ("SELECT ThreadDpth FROM PipeDimensions WHERE Size = '"
                       + self.cmbOD.GetValue() + "'")
                self.txtThrd.ChangeValue(Dbase().Dsqldata(qry)[0][0])
                self.txtGrv.ChangeValue('')
                self.Cd = eval(self.txtThrd.GetValue())
            elif self.rdBox1.GetSelection() == 2:
                qry = ("SELECT GrooveDpth FROM PipeDimensions WHERE Size = '"
                       + self.cmbOD.GetValue() + "'")
                Dbase().Dsqldata(qry)
                self.txtGrv.ChangeValue(Dbase().Dsqldata(qry)[0][0])
                self.txtThrd.ChangeValue('')
                self.Cd = eval(self.txtGrv.GetValue())
            elif self.rdBox1.GetSelection() == 0:
                self.txtThrd.ChangeValue('')
                self.txtGrv.ChangeValue('')
                self.Cd = 0

            # fill the schedule drop down based on pipe OD
            ReFillQuery = (
                "SELECT Wall,Sch FROM PipeDimensions WHERE Size = '" +
                self.cmbOD.GetValue() + "'")
            self.cmbSch.Enable()

            lctr = self.lctrls[2]
            lctr.DeleteAllItems()
            index = 0

            for values in Dbase().Dsqldata(ReFillQuery):
                col = 0
                for value in values:
                    if col == 0:
                        lctr.InsertItem(index, str(value))
                    else:
                        lctr.SetItem(index, col, str(value))
                    col += 1
                index += 1

    def WallThk(self, evt):
        qry = ('SELECT Sch FROM PipeDimensions WHERE Wall = "'
               + str(self.cmbSch.GetValue())) + '"'
        self.noteSch.ChangeValue(Dbase().Dsqldata(qry)[0][0])

    def CalcWall(self, evt):
        chck = self.ValWallData()
        if chck != []:
            self.DataWarn(chck)
        else:
            if self.rdBox.GetStringSelection() == 'ERW':
                E = eval(self.ERWtxt.GetValue())
            else:
                E = eval(self.Smlstxt.GetValue())

            Pd = eval(self.DPtxt.GetValue())
            pipeOD = eval(self.cmbOD.GetValue()[:-1].replace('-', '+'))
            Sd = eval(self.Stresstxt.GetValue())
            Y = eval(self.StrsFact.GetValue())
            c = eval(self.cmbCor.GetValue())

            WT1 = Pd * pipeOD
            WT2 = 2 * (Sd * E + pipeOD * Y)
            WT = WT1 / WT2 + c + self.Cd
            self.noteCalcWall.SetLabel(str('{0:5.3f}'.format(WT)) + ' inches')

    def CalcHydro(self, evt):
        chckH = self.ValHydroData()
        if chckH != []:
            self.DataWarn(chckH)
        else:
            Pd = eval(self.DPtxt.GetValue())
            Sh = eval(self.Hydrotxt.GetValue())
            Sd = eval(self.Stresstxt.GetValue())
            HP1 = Sh / Sd
            if Sh > Sd:
                HP1 = 1
            HP = 1.5 * Pd * HP1
            self.noteCalcHydro.SetLabel(str('{0:5.3f}'.format(HP)) + ' psig')

    def ValWallData(self):
        chck = []
        if self.Stresstxt.GetValue() == '':
            chck.append(1)
        if self.DPtxt.GetValue() == '':
            chck.append(3)
        if self.cmbOD.GetValue() == '':
            chck.append(4)
        if self.cmbCor.GetValue() == '':
            chck.append(5)
        if (self.ERWtxt.GetValue() == '' and self.Smlstxt.GetValue() == ''):
            chck.append(6)
        if self.StrsFact.GetValue() == '':
            chck.append(7)
        return chck

    def DataWarn(self, chck):
        bxnames = ['Allowable Stress Sd', 'Allowable Stress Sh',
                   'Design Pressure Pd', 'Pipe OD', 'Corrosion Allowance',
                   'Pipe Quality factor E', 'Temp Compensation factor Y']
        missingstr = ''
        for num in chck:
            missingstr = missingstr + bxnames[num-1] + '\n'
        MsgBx = wx.MessageDialog(self, 'Value(s) needed for;\n' + missingstr,
                                 'Missing Data', wx.OK | wx.ICON_HAND)

        MsgBx_val = MsgBx.ShowModal()
        MsgBx.Destroy()
        if MsgBx_val == wx.ID_OK:
            return False

    def ValHydroData(self):
        chckH = []
        if self.Stresstxt.GetValue() == '':
            chckH.append(1)
        if self.Hydrotxt.GetValue() == '':
            chckH.append(2)
        if self.DPtxt.GetValue() == '':
            chckH.append(3)
        return chckH

    def OnFlange(self, evt):
        FlgRatg(self)

    def OnClose(self, evt):
        self.GetParent().Enable(True)   # add for child form
        self.__eventLoop.Exit()        # add for child form
        self.Destroy()


class FlgRatg(wx.Frame):
    def __init__(self, parent):
        self.lctrls = []

        super(FlgRatg, self).__init__(
            parent,
            title='Flange Pressure and Temperature Rating',
            size=(660, 750),
            style=wx.DEFAULT_FRAME_STYLE & ~ (wx.RESIZE_BORDER |
                                              wx.MAXIMIZE_BOX |
                                              wx.MINIMIZE_BOX))

        self.Bind(wx.EVT_CLOSE, self.OnClose)

        self.parent = parent
        self.tempbxs = []
        self.pressbxs = []
        self.datax = []
        self.datay = []
        self.InitUI()

    def InitUI(self):
        self.pnl = FlgRtgPnl(self)

        self.mainSizer = wx.BoxSizer(wx.VERTICAL)

        sizer_1 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_1.Add(self.pnl, 1, wx.EXPAND)

        self.tempszr = wx.BoxSizer(wx.HORIZONTAL)
        self.presszr = wx.BoxSizer(wx.HORIZONTAL)
        self.txttemp = wx.StaticText(self, label="Temperature Data:")
        self.txtpress = wx.StaticText(self, label="Pressure Data:")
        self.tempszr.Add(self.txttemp, 0,
                         wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 15)
        self.presszr.Add(self.txtpress, 0,
                         wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 45)

        btnszr = wx.BoxSizer(wx.HORIZONTAL)
        self.cmbPipe = wx.ComboCtrl(self, size=(90, -1), style=wx.CB_READONLY)
        self.Bind(wx.EVT_COMBOBOX_CLOSEUP, self.OnSelectP, self.cmbPipe)
        self.Bind(wx.EVT_TEXT, self.OnSelectP, self.cmbPipe)
        self.cmbPipe.SetPopupControl(ListCtrlComboPopup(
            'PipeMaterialSpec', showcol=1, PupQuery='', lctrls=self.lctrls))
        self.tmp = wx.TextCtrl(self, name='temp', style=wx.TE_PROCESS_ENTER)
        self.tmp.SetHint('Temperature')
        self.tmp.Enable(False)
        self.Bind(wx.EVT_TEXT_ENTER, self.OnClosePt, self.tmp)
        self.prs = wx.TextCtrl(self, name='press', style=wx.TE_PROCESS_ENTER)
        self.prs.SetHint('Pressure')
        self.prs.Enable(False)
        self.Bind(wx.EVT_TEXT_ENTER, self.OnClosePt, self.prs)
        lblpt = wx.StaticText(self, label='\t<--\nMaximum\n\t-->')
        self.b1 = wx.Button(self, label="Exit")
        self.Bind(wx.EVT_BUTTON, self.OnClose, self.b1)
        btnszr.Add(self.cmbPipe, 0, wx.LEFT | wx.RIGHT | wx.ALIGN_CENTER, 25)
        btnszr.Add(self.tmp, 0, wx.ALIGN_CENTER)
        btnszr.Add(lblpt, 0, wx.ALIGN_CENTER_VERTICAL)
        btnszr.Add(self.prs, 0, wx.ALIGN_CENTER)
        btnszr.Add(self.b1, 0, wx.LEFT | wx.ALIGN_CENTER, 25)

        lblszr = wx.BoxSizer(wx.HORIZONTAL)
        lbltmp = wx.StaticText(self, label='Temperature')
        lblprs = wx.StaticText(self, label='Pressure')
        lblszr.Add(lbltmp, 0, wx.ALIGN_TOP)
        lblszr.Add((65, 10))
        lblszr.Add(lblprs, 0, wx.ALIGN_TOP)

        self.mainSizer.Add(sizer_1, 0, wx.EXPAND)
        self.mainSizer.Add((50, 25))
        self.mainSizer.Add(self.tempszr, 0,
                           wx.EXPAND | wx.ALIGN_LEFT)
        self.mainSizer.Add((50, 25))
        self.mainSizer.Add(self.presszr, 0, wx.ALIGN_LEFT | wx.BOTTOM, 20)
        self.mainSizer.Add(btnszr, 0, wx.ALIGN_CENTER)
        self.mainSizer.Add(lblszr, 0, wx.ALIGN_CENTER)

        self.SetSizer(self.mainSizer)
        self.Layout()

        self.CenterOnParent()
        self.GetParent().Enable(False)
        self.Show(True)
        self.__eventLoop = wx.GUIEventLoop()
        self.__eventLoop.Run()

    def Update_pnl(self):
        self.pnl.draw(self.datax, self.datay)

    def OnSelectP(self, evt):
        code = self.cmbPipe.GetValue()
        qry = ('''SELECT Material_Spec_ID FROM PipeMaterialSpec
                WHERE Pipe_Material_Spec = "''' +
               self.cmbPipe.GetValue() + '"')
        id = Dbase().Dsqldata(qry)[0][0]
        qry = ('''SELECT Temperature, Pressure FROM PressTempTables
                WHERE Specification_Number = ''' + str(id))
        data1 = Dbase().Dsqldata(qry)
        if data1 != []:
            sortdata = sorted(data1, key=lambda tup: tup[0])
            self.datax, self.datay = map(list, zip(*sortdata))
            if len(self.datax) <= 10:
                nr = 10
            else:
                nr = len(self.datax)
            self.SetSize(66*nr, 750)
            self.Update_pnl()
            self.pnl.axes.set_title("Material Code " + code)
            self.tblfill()
            self.prs.Enable()
            self.tmp.Enable()
        else:
            self.pnl.axes.set_title("Material Code not setup for " + code)
            self.tblfill()
            n = 0
            for item in self.pressbxs:
                item.ChangeValue('')
                self.tempbxs[n].ChangeValue('')
                n += 1

    def tblfill(self):
        for n in range(len(self.tempbxs)):
            self.tempbxs[n].Destroy()
            self.pressbxs[n].Destroy()
        self.tempbxs = []
        self.pressbxs = []
        self.tempszr.Clear()
        self.presszr.Clear()
        self.txtpress.SetLabel('')
        self.txttemp.SetLabel('')
        self.txttemp = wx.StaticText(self, label="Temperature Data:")
        self.txtpress = wx.StaticText(self, label="Pressure Data:")
        self.tempszr.Add(self.txttemp, 0, wx.ALIGN_LEFT |
                         wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 15)
        self.presszr.Add(self.txtpress, 0, wx.ALIGN_LEFT |
                         wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 45)
        for n in range(len(self.datax)):
            tempbx = wx.TextCtrl(self, value=str(self.datax[n]), size=(50, 33),
                                 style=wx.TE_READONLY | wx.TE_RIGHT)
            pressbx = wx.TextCtrl(self, value=str(self.datay[n]),
                                  size=(50, 33),
                                  style=wx.TE_READONLY | wx.TE_RIGHT)
            self.tempszr.Add(tempbx, 0, wx.ALIGN_LEFT)
            self.presszr.Add(pressbx, 0, wx.ALIGN_LEFT)
            self.tempbxs.append(tempbx)
            self.pressbxs.append(pressbx)
        self.Layout()

    def OnClosePt(self, evt):
        # call closept calculate the value coressponding to the specified value
        # indicate the value of the point and the box in which it was specified
        maxvl = self.closept(evt.GetEventObject().GetValue(),
                             evt.GetEventObject().GetName())
        # place the calculated value in the proper text box
        if evt.GetEventObject().GetName() == 'temp':
            self.prs.ChangeValue(str(maxvl))
        else:
            self.tmp.ChangeValue(str(maxvl))

    def closept(self, pt, axs):
        # find closest number to pt
        pt = int(pt)
        if axs == 'temp':
            dt = self.datax
            # if the point specified is an existing data point
            # just return corresponding value
            if pt in dt:
                indx = dt.index(pt)
                maxvl = self.datay[indx]
                return maxvl
            # if the specified point is outside the range
            # of data do not extrapolate
            if pt < min(self.datax) or pt > max(self.datax):
                maxvl = 'none'
                return maxvl
        else:
            dt = self.datay
            # if the point specified is an existing data point
            # just return corresponding value
            if pt in dt:
                indx = dt.index(pt)
                maxvl = self.datax[indx]
                return maxvl
            # if the specified point is outside the range
            # of data do not extrapolate
            if pt < min(self.datay) or pt > max(self.datay):
                maxvl = 'none'
                return maxvl

        nrpt = min(dt, key=lambda x: abs(x - pt))
        # find the index value of the near point
        indx1 = dt.index(nrpt)
        # set the index value of the point to the other side of your number
        if nrpt > pt:
            indx2 = indx1 - 1
        else:
            indx2 = indx1 + 1

        # set up the coefficents needed
        a = self.datay[indx2] - self.datay[indx1]
        b = self.datax[indx2] - self.datax[indx1]
        c = (self.datax[indx2]*self.datay[indx1] -
             self.datax[indx1]*self.datay[indx2])
        # based on the line between the two points being linear
        # calculate point of intersection
        if axs == 'temp':
            maxvl = a*pt/b + c/b
        else:
            maxvl = pt*b/a - c/a
        return round(maxvl, 1)

    def OnClose(self, evt):
        self.GetParent().Enable(True)   # add this line for child
        self.__eventLoop.Exit()     # add this line for child
        self.Destroy()


class FlgRtgPnl(wx.Panel):
    def __init__(self, parent):
        super(FlgRtgPnl, self).__init__(parent)
        self.figure = Figure()
        self.axes = self.figure.add_subplot(111)
        self.axes.grid(color='b', linestyle='-', linewidth=.2)
        self.canvas = FigureCanvas(self, -1, self.figure)
        self.sizer = wx.BoxSizer(wx.VERTICAL)
        self.sizer.Add(self.canvas, 1, wx.LEFT | wx.TOP | wx.GROW)
        self.SetSizer(self.sizer)
        self.Fit()

    def draw(self, datax, datay):
        temp = datax
        press = datay
        self.axes.clear()
        self.axes.grid(color='b', linestyle='-', linewidth=.2)
        self.axes.set_xticks(temp)
        self.axes.set_yticks(press)
        self.axes.set_ylabel('Pressure (psig)')
        self.axes.set_xlabel('Temperature (F)')
        self.axes.plot(temp, press, marker='.')
        self.canvas.draw()


class MergeFrm(wx.Frame):
    def __init__(self, parent):

        super(MergeFrm, self).__init__(parent,
                                       title='Select pdf files to merge',
                                       size=(725, 550))

        self.Bind(wx.EVT_CLOSE, self.OnExit)
        self.parent = parent
        self.InitUI()

    def InitUI(self):
        from wx.lib.itemspicker import ItemsPicker, \
             EVT_IP_SELECTION_CHANGED, IP_REMOVE_FROM_CHOICES

        self.Sizer = wx.BoxSizer(wx.VERTICAL)

        pick_sizer = wx.BoxSizer(wx.VERTICAL)
        self.ip = ItemsPicker(self, -1,
                              [],
                              'All pdf files in\nselected directory:',
                              'Files to be merged\nin order shown:',
                              size=(700, 400), ipStyle=IP_REMOVE_FROM_CHOICES)
        # event occures when bAdd is pressed
        self.ip.Bind(EVT_IP_SELECTION_CHANGED, self.OnSelectionChange)
        self.ip._source.SetMinSize((-1, 150))

        self.ip.bAdd.SetLabel('Add =>')
        self.ip.bRemove.SetLabel('<= Remove')
        pick_sizer.Add(self.ip, 0, wx.ALL, 10)

        btnsizer = wx.BoxSizer(wx.HORIZONTAL)
        b1 = wx.Button(self, label="Select File\nDirectory")
        self.Bind(wx.EVT_BUTTON, self.OnAdd, b1)

        b2 = wx.Button(self, label="Exit")
        self.Bind(wx.EVT_BUTTON, self.OnExit, b2)

        b3 = wx.Button(self, label='Merge\nPDF files')
        self.Bind(wx.EVT_BUTTON, self.OnMerge, b3)

        btnsizer.Add(b1, 0, wx.ALL, 5)
        btnsizer.Add((340, 10))
        btnsizer.Add(b3, 0, wx.ALL, 5)
        btnsizer.Add(b2, 0, wx.ALL, 5)
        self.Sizer.Add(pick_sizer, 0, wx.ALL, 10)
        self.Sizer.Add(btnsizer, 0, wx.ALL | wx.ALIGN_CENTER, 10)

        self.CenterOnParent()
        self.GetParent().Enable(False)
        self.Show(True)
        self.__eventLoop = wx.GUIEventLoop()
        self.__eventLoop.Run()

    def OnSelectionChange(self, evt):
        # items in left side box
        self.pdf_list = self.ip.GetSelections()

    def OnMerge(self, evt):
        self.MergePDF()

    def MergePDF(self):
        # get a specific user file name
        msg = 'Save Report as PDF.'
        wildcard = 'PDF (*.pdf)|*.pdf'
        filename = self.FylDilog(wildcard, msg, wx.FD_SAVE |
                                 wx.FD_OVERWRITE_PROMPT)
        # confirm an extension was specifed as pdf if not
        # add pdf extension
        if filename[-4:].lower() != '.pdf':
            filename += '.pdf'

        merger = PdfFileMerger()
        for pdf in self.pdf_list:
            pdf_file = self.path + '/' + pdf
            merger.append(open(pdf_file, 'rb'))

        with open(filename, "wb") as fout:
            merger.write(fout)

    def OnAdd(self, evt):
        dlg = wx.DirDialog(self, "Choose a directory containing PDF files:",
                           style=wx.DD_DEFAULT_STYLE)

        if dlg.ShowModal() == wx.ID_OK:

            self.path = dlg.GetPath()
            files = [a for a in os.listdir(self.path) if a.endswith(".pdf")]
        dlg.Destroy()
        self.ip.SetItems(files)

    def FylDilog(self, wildcard, msg, stl):
        self.currentDirectory = os.getcwd()
        path = 'No File'
        dlg = wx.FileDialog(
            self, message=msg,
            defaultDir=self.currentDirectory,
            defaultFile="",
            wildcard=wildcard,
            style=stl
            )
        if dlg.ShowModal() == wx.ID_OK:
            path = dlg.GetPaths()[0]
        dlg.Destroy()

        return path

    def OnExit(self, evt):
        self.GetParent().Enable(True)   # add for child form
        self.__eventLoop.Exit()        # add for child form
        self.Destroy()


class RPSImport(wx.Frame):
    '''Routine to build form and populate grid'''
    def __init__(self, parent):

        self.frmtitle = 'Request for New Pipe Specification'

        super(RPSImport, self).__init__(parent, title=self.frmtitle,
                                        size=(900, 880))

        self.Bind(wx.EVT_CLOSE, self.OnClose)

        self.columnames = []
        self.rec_num = 0
        self.lctrls = []
        self.data = []
        self.MainSQL = ''
        self.tblname = 'NewPipeSpec'
        self.MainSQL = 'SELECT * FROM NewPipeSpec'
        self.data = Dbase().Dsqldata(self.MainSQL)
        # specify which listbox column to display in the combobox
        self.showcol = int

        self.InitUI()

    def InitUI(self):
        # set the Sizer property (same as SetSizer)
        self.Sizer = wx.BoxSizer(wx.VERTICAL)

        self.IDsizer = wx.BoxSizer(wx.HORIZONTAL)
        RcdIDlbl = wx.StaticText(self, label='Record ID',
                                 style=wx.ALIGN_LEFT)
        RcdIDlbl.SetForegroundColour((255, 0, 0))
        self.RcdIDtxt = wx.TextCtrl(self, size=(80, 33), value='',
                                    style=wx.TE_CENTER)
        self.RcdIDtxt.Enable(False)
        TagIDlbl = wx.StaticText(self, label='Tags',
                                 style=wx.ALIGN_LEFT)
        TagIDlbl.SetForegroundColour((255, 0, 0))
        self.TagIDtxt = wx.TextCtrl(self, size=(200, 33), value='',
                                    style=wx.TE_CENTER)
        St2 = wx.StaticText(self, label='* S',
                            style=wx.ALIGN_LEFT)
        St2.SetForegroundColour((255, 0, 0))
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.TagIDtxt)
        self.IDsizer.Add(RcdIDlbl, 0, wx.CENTER | wx.LEFT, 50)
        self.IDsizer.Add(self.RcdIDtxt, 0, wx.ALIGN_CENTRE | wx.RIGHT, 50)
        self.IDsizer.Add(TagIDlbl, 0, wx.LEFT | wx.CENTER, 50)
        self.IDsizer.Add(self.TagIDtxt, 0, wx.ALIGN_CENTER, 10)
        self.IDsizer.Add(St2, 0, wx.CENTER)

        self.textsizer = wx.BoxSizer(wx.HORIZONTAL)

        self.notesizer2 = wx.BoxSizer(wx.VERTICAL)
        self.note1 = wx.StaticText(self, label='Requestor',
                                   style=wx.ALIGN_LEFT)
        self.note1.SetForegroundColour((255, 0, 0))
        self.note2 = wx.StaticText(self, label='Issue date',
                                   style=wx.ALIGN_LEFT)
        self.note2.SetForegroundColour((255, 0, 0))
        self.notesizer2.Add(self.note1, 0, wx.ALIGN_RIGHT)
        self.notesizer2.Add(self.note2, 0, wx.ALIGN_RIGHT |
                            wx.LEFT | wx.TOP, 30)

        self.txtsizer2 = wx.BoxSizer(wx.VERTICAL)
        self.text1 = wx.TextCtrl(self, size=(150, 33), value='',
                                 style=wx.TE_CENTER)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text1)

        self.text2 = wx.TextCtrl(self, size=(100, 33), value='',
                                 style=wx.TE_CENTER)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text2)
        self.txtsizer2.Add(self.text1, 0, wx.ALIGN_CENTER)
        self.txtsizer2.Add(self.text2, 0, wx.ALIGN_CENTER | wx.TOP, 10)

        starszr1 = wx.BoxSizer(wx.VERTICAL)
        St3 = wx.StaticText(self, label='* S',
                            style=wx.ALIGN_LEFT)
        St3.SetForegroundColour((255, 0, 0))
        St4 = wx.StaticText(self, label='* S',
                            style=wx.ALIGN_LEFT)
        St4.SetForegroundColour((255, 0, 0))
        starszr1.Add(St3, 0, wx.ALIGN_LEFT | wx.BOTTOM, 25)
        starszr1.Add(St4, 0, wx.ALIGN_LEFT)

        self.notesizer3 = wx.BoxSizer(wx.VERTICAL)
        self.note3 = wx.StaticText(self, label='P&&ID Line Number',
                                   style=wx.ALIGN_LEFT)
        self.note3.SetForegroundColour((255, 0, 0))
        self.text3 = wx.TextCtrl(self, size=(130, 33), value='',
                                 style=wx.TE_CENTER)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text3)
        self.notesizer3.Add(self.note3, 0, wx.ALIGN_RIGHT)
        self.notesizer3.Add(self.text3, 0, wx.ALIGN_CENTER)

        self.Dscrptsizer = wx.BoxSizer(wx.VERTICAL)
        self.Descrptlbl = wx.StaticText(self, label='Supporting Documents',
                                        style=wx.ALIGN_LEFT)
        self.Descrptlbl.SetForegroundColour((255, 0, 0))
        self.Descrpttxt = wx.TextCtrl(self, size=(350, 90), value='',
                                      style=wx.TE_MULTILINE)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.Descrpttxt)
        self.Dscrptsizer.Add(self.Descrptlbl, 0, wx.ALIGN_LEFT | wx.LEFT, 10)
        self.Dscrptsizer.Add(self.Descrpttxt, 0, wx.ALIGN_RIGHT)

        self.textsizer.Add(self.notesizer2, 0, wx.LEFT |
                           wx.ALIGN_CENTER_VERTICAL, 10)
        self.textsizer.Add(self.txtsizer2, 0,
                           wx.ALIGN_CENTER_VERTICAL)
        self.textsizer.Add(starszr1, 0, wx.RIGHT |
                           wx.ALIGN_CENTER_VERTICAL, 10)
        self.textsizer.Add((30, 10))
        self.textsizer.Add(self.notesizer3, 0, wx.LEFT |
                           wx.ALIGN_CENTER_VERTICAL, 10)
        self.textsizer.Add((30, 10))
        self.textsizer.Add(self.Dscrptsizer, 0, wx.LEFT |
                           wx.ALIGN_CENTER_VERTICAL, 10)

        self.textsizer8 = wx.BoxSizer(wx.HORIZONTAL)
        self.OpDsizer = wx.BoxSizer(wx.VERTICAL)
        self.note13 = wx.StaticText(self, label='Similar Pipe Spec',
                                    style=wx.ALIGN_LEFT)
        self.note13.SetForegroundColour((255, 0, 0))

        self.note14 = wx.StaticText(self,
                                    label='Max. Operating\nPressure (psig)',
                                    style=wx.ALIGN_LEFT)
        self.note14.SetForegroundColour((255, 0, 0))

        self.note15 = wx.StaticText(self,
                                    label='Max. Operating\nTemperature (F)',
                                    style=wx.ALIGN_LEFT)
        self.note15.SetForegroundColour((255, 0, 0))
        self.OpDsizer.Add(self.note13, 0, wx.ALIGN_RIGHT)
        self.OpDsizer.Add(self.note14, 0, wx.ALIGN_RIGHT |
                          wx.TOP | wx.BOTTOM, 15)
        self.OpDsizer.Add(self.note15, 0, wx.ALIGN_RIGHT)

        self.txtsizer6 = wx.BoxSizer(wx.VERTICAL)
        self.text4 = wx.TextCtrl(self, size=(100, 33), value='',
                                 style=wx.TE_CENTER)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text4)

        self.text5 = wx.TextCtrl(self, size=(100, 33), value='',
                                 style=wx.TE_CENTER)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text5)
        self.text6 = wx.TextCtrl(self, size=(100, 33), value='',
                                 style=wx.TE_CENTER)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text6)
        self.txtsizer6.Add(self.text4, 0)
        self.txtsizer6.Add(self.text5, 0, wx.TOP | wx.BOTTOM, 10)
        self.txtsizer6.Add(self.text6, 0)

        starszr2 = wx.BoxSizer(wx.VERTICAL)
        St5 = wx.StaticText(self, label='*',
                            style=wx.ALIGN_LEFT)
        St5.SetForegroundColour((255, 0, 0))
        St6 = wx.StaticText(self, label='*',
                            style=wx.ALIGN_LEFT)
        St6.SetForegroundColour((255, 0, 0))
        starszr2.Add(St5, 0, wx.ALIGN_LEFT | wx.BOTTOM, 25)
        starszr2.Add(St6, 0, wx.ALIGN_LEFT)

        self.OpDsizer2 = wx.BoxSizer(wx.VERTICAL)
        self.note16 = wx.StaticText(self, label='Max. Design\nPressure (psig)',
                                    style=wx.ALIGN_LEFT)
        self.note16.SetForegroundColour((255, 0, 0))
        self.note17 = wx.StaticText(self, label='Max. Design\nTemperature (F)',
                                    style=wx.ALIGN_LEFT)
        self.note17.SetForegroundColour((255, 0, 0))
        self.note18 = wx.StaticText(self, label='Min. Design\nTemperature (F)',
                                    style=wx.ALIGN_LEFT)
        self.note18.SetForegroundColour((255, 0, 0))
        self.OpDsizer2.Add(self.note16, 0, wx.ALIGN_RIGHT)
        self.OpDsizer2.Add(self.note17, 0, wx.ALIGN_RIGHT |
                           wx.TOP | wx.BOTTOM, 15)
        self.OpDsizer2.Add(self.note18, 0, wx.ALIGN_RIGHT)

        self.txtsizer7 = wx.BoxSizer(wx.VERTICAL)
        self.text7 = wx.TextCtrl(self, size=(100, 33), value='',
                                 style=wx.TE_CENTER)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text7)

        self.text8 = wx.TextCtrl(self, size=(100, 33), value='',
                                 style=wx.TE_CENTER)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text8)
        self.text9 = wx.TextCtrl(self, size=(100, 33), value='',
                                 style=wx.TE_CENTER)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text9)
        self.txtsizer7.Add(self.text7, 0)
        self.txtsizer7.Add(self.text8, 0, wx.TOP | wx.BOTTOM, 15)
        self.txtsizer7.Add(self.text9, 0)

        starszr3 = wx.BoxSizer(wx.VERTICAL)
        St7 = wx.StaticText(self, label='*',
                            style=wx.ALIGN_LEFT)
        St7.SetForegroundColour((255, 0, 0))
        St8 = wx.StaticText(self, label='*',
                            style=wx.ALIGN_LEFT)
        St8.SetForegroundColour((255, 0, 0))
        St9 = wx.StaticText(self, label='*',
                            style=wx.ALIGN_LEFT)
        St9.SetForegroundColour((255, 0, 0))
        starszr3.Add(St7, 0, wx.ALIGN_LEFT)
        starszr3.Add(St8, 0, wx.ALIGN_LEFT | wx.TOP | wx.BOTTOM, 30)
        starszr3.Add(St9, 0, wx.ALIGN_LEFT)

        lbl1 = '* - indicates required field'
        lbl2 = '\nS - indicates searchable field'
        note = wx.StaticText(self,
                             label=lbl1 + lbl2,
                             style=wx.ALIGN_LEFT)
        note.SetForegroundColour((255, 0, 0))

        self.textsizer8.Add(self.OpDsizer, 0, wx.LEFT |
                            wx.ALIGN_CENTER_VERTICAL, 30)
        self.textsizer8.Add(self.txtsizer6, 0, wx.LEFT |
                            wx.ALIGN_CENTER_VERTICAL, 10)
        self.textsizer8.Add(starszr2, 0, wx.ALIGN_BOTTOM | wx.BOTTOM, 15)
        self.textsizer8.Add((65, 10))
        self.textsizer8.Add(self.OpDsizer2, 0, wx.LEFT |
                            wx.ALIGN_CENTER_VERTICAL, 10)
        self.textsizer8.Add(self.txtsizer7, 0,
                            wx.ALIGN_CENTER)
        self.textsizer8.Add(starszr3, 0, wx.ALIGN_CENTER_VERTICAL)
        self.textsizer8.Add(note, 0, wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 45)

        self.Finalsizer = wx.BoxSizer(wx.HORIZONTAL)

        self.Reprsizer = wx.BoxSizer(wx.VERTICAL)
        self.Reprlbl = wx.StaticText(self, label='Fluid Description  *',
                                     style=wx.ALIGN_LEFT)
        self.Reprlbl.SetForegroundColour((255, 0, 0))
        self.Reprtxt = wx.TextCtrl(self, size=(350, 100), value='',
                                   style=wx.TE_MULTILINE)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.Reprtxt)
        self.Reprsizer.Add(self.Reprlbl, 0, wx.ALIGN_LEFT | wx.LEFT, 10)
        self.Reprsizer.Add(self.Reprtxt, 0, wx.ALIGN_RIGHT)

        self.notesizer4 = wx.BoxSizer(wx.VERTICAL)
        self.note5 = wx.StaticText(self, label='Approved Pipe\nSpecification',
                                   style=wx.ALIGN_LEFT)
        self.note5.SetForegroundColour((255, 0, 0))
        self.note7 = wx.StaticText(self, label='Name of Signing\nQC Manager',
                                   style=wx.ALIGN_LEFT)
        self.note7.SetForegroundColour((255, 0, 0))
        self.notesizer4.Add(self.note5, 0, wx.ALIGN_RIGHT | wx.ALL, 10)
        self.notesizer4.Add(self.note7, 0, wx.ALIGN_RIGHT | wx.ALL, 10)

        self.txtsizer4 = wx.BoxSizer(wx.VERTICAL)
        self.text10 = wx.TextCtrl(self, size=(150, 33), value='',
                                  style=wx.TE_CENTER)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text10)
        self.text11 = wx.TextCtrl(self, size=(150, 33), value='',
                                  style=wx.TE_CENTER)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text11)
        self.txtsizer4.Add(self.text10, 0, wx.ALL | wx.ALIGN_CENTER, 10)
        self.txtsizer4.Add(self.text11, 0, wx.ALL | wx.ALIGN_CENTER, 10)

        self.notesizer5 = wx.BoxSizer(wx.VERTICAL)
        self.note4 = wx.StaticText(self, label='Date',
                                   style=wx.ALIGN_LEFT)
        self.note4.SetForegroundColour((255, 0, 0))
        self.note10 = wx.StaticText(self, label='Date',
                                    style=wx.ALIGN_LEFT)
        self.note10.SetForegroundColour((255, 0, 0))
        self.notesizer5.Add(self.note4, 0, wx.ALIGN_RIGHT | wx.ALL, 20)
        self.notesizer5.Add(self.note10, 0, wx.ALIGN_RIGHT | wx.ALL, 20)

        self.txtsizer5 = wx.BoxSizer(wx.VERTICAL)
        self.text12 = wx.TextCtrl(self, size=(80, 33), value='',
                                  style=wx.TE_CENTER)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text12)
        self.text13 = wx.TextCtrl(self, size=(80, 33), value='',
                                  style=wx.TE_CENTER)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text13)

        self.txtsizer5.Add(self.text12, 0, wx.ALL | wx.ALIGN_CENTER, 10)
        self.txtsizer5.Add(self.text13, 0, wx.ALL | wx.ALIGN_CENTER, 10)

        self.Finalsizer.Add(self.Reprsizer, 0, wx.LEFT |
                            wx.ALIGN_CENTER_VERTICAL, 10)
        self.Finalsizer.Add(self.notesizer4, 0, wx.LEFT |
                            wx.ALIGN_CENTER_VERTICAL, 10)
        self.Finalsizer.Add(self.txtsizer4, 0, wx.RIGHT |
                            wx.ALIGN_CENTER_VERTICAL, 10)
        self.Finalsizer.Add(self.notesizer5, 0, wx.LEFT |
                            wx.ALIGN_CENTER_VERTICAL, 10)
        self.Finalsizer.Add(self.txtsizer5, 0, wx.RIGHT |
                            wx.ALIGN_CENTER_VERTICAL, 10)

        self.Finalsizer1 = wx.BoxSizer(wx.HORIZONTAL)

        self.notesizer6 = wx.BoxSizer(wx.VERTICAL)
        self.note8 = wx.StaticText(self, label='Name of Signing\nOperations',
                                   style=wx.ALIGN_LEFT)
        self.note8.SetForegroundColour((255, 0, 0))

        self.note9 = wx.StaticText(self, label='Name of Signing\nRequestor',
                                   style=wx.ALIGN_LEFT)
        self.note9.SetForegroundColour((255, 0, 0))
        self.notesizer6.Add(self.note8, 0, wx.ALIGN_RIGHT | wx.ALL, 10)
        self.notesizer6.Add(self.note9, 0, wx.ALIGN_RIGHT | wx.ALL, 10)

        self.textsizer9 = wx.BoxSizer(wx.VERTICAL)
        self.text14 = wx.TextCtrl(self, size=(150, 33), value='',
                                  style=wx.TE_CENTER)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text14)

        self.text15 = wx.TextCtrl(self, size=(150, 33), value='',
                                  style=wx.TE_CENTER)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text15)
        self.textsizer9.Add(self.text14, 0, wx.ALL | wx.ALIGN_CENTER, 10)
        self.textsizer9.Add(self.text15, 0, wx.ALL | wx.ALIGN_CENTER, 10)

        self.notesizer7 = wx.BoxSizer(wx.VERTICAL)
        self.note11 = wx.StaticText(self, label='Date',
                                    style=wx.ALIGN_LEFT)
        self.note11.SetForegroundColour((255, 0, 0))

        self.note12 = wx.StaticText(self, label='Date',
                                    style=wx.ALIGN_LEFT)
        self.note12.SetForegroundColour((255, 0, 0))
        self.notesizer7.Add(self.note11, 0, wx.ALIGN_RIGHT | wx.ALL, 20)
        self.notesizer7.Add(self.note12, 0, wx.ALIGN_RIGHT | wx.ALL, 20)

        self.textsizer10 = wx.BoxSizer(wx.VERTICAL)
        self.text16 = wx.TextCtrl(self, size=(80, 33), value='',
                                  style=wx.TE_CENTER)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text16)

        self.text17 = wx.TextCtrl(self, size=(80, 33), value='',
                                  style=wx.TE_CENTER)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text17)
        self.textsizer10.Add(self.text16, 0, wx.ALL | wx.ALIGN_CENTER, 10)
        self.textsizer10.Add(self.text17, 0, wx.ALL | wx.ALIGN_CENTER, 10)

        self.Finalsizer1.Add(self.notesizer6, 0,
                             wx.ALIGN_CENTER_VERTICAL)
        self.Finalsizer1.Add(self.textsizer9, 0, wx.RIGHT |
                             wx.ALIGN_CENTER_VERTICAL, 10)
        self.Finalsizer1.Add(self.notesizer7, 0, wx.LEFT |
                             wx.ALIGN_CENTER_VERTICAL, 10)
        self.Finalsizer1.Add(self.textsizer10, 0, wx.RIGHT |
                             wx.ALIGN_CENTER_VERTICAL, 10)

        self.Sizer.Add(self.IDsizer, 0, wx.BOTTOM | wx.TOP, 25)
        self.Sizer.Add(self.textsizer, 0, wx.ALIGN_CENTER |
                       wx.TOP | wx.BOTTOM, 5)
        self.Sizer.Add(self.textsizer8, 0, wx.ALIGN_CENTER |
                       wx.TOP | wx.BOTTOM, 20)
        self.Sizer.Add(self.Finalsizer, 0, wx.ALIGN_CENTER | wx.BOTTOM, 20)
        self.Sizer.Add(self.Finalsizer1, 0, wx.ALIGN_CENTER)

        self.btnbox = wx.BoxSizer(wx.HORIZONTAL)
        # Add buttons for grid modifications
        self.b1 = wx.Button(self, label="Import HTML")
        self.Bind(wx.EVT_BUTTON, self.ImportHTML, self.b1)

        self.b2 = wx.Button(self, label="Save Report Data")
        self.b2.Disable()
        self.Bind(wx.EVT_BUTTON, self.OnAddRec, self.b2)

        self.b3 = wx.Button(self, label="Delete Spec")
        self.Bind(wx.EVT_BUTTON, self.OnDelete, self.b3)

        self.b4 = wx.Button(self, label="Exit")
        self.Bind(wx.EVT_BUTTON, self.OnClose, self.b4)

        # add a button box and place the buttons
        self.btnbox.Add(self.b1, 0, wx.ALL, 5)
        self.btnbox.Add(self.b2, 0, wx.ALL, 5)
        self.btnbox.Add(self.b3, 0, wx.ALL, 5)
        self.btnbox.Add(self.b4, 0, wx.ALL, 5)

        self.srchbox = wx.BoxSizer(wx.HORIZONTAL)

        self.b5 = wx.Button(self, label="Search")
        self.Bind(wx.EVT_BUTTON, self.OnSearch, self.b5)

        self.b6 = wx.Button(self, label="Clear Boxes")
        self.Bind(wx.EVT_BUTTON, self.OnClear, self.b6)

        self.b7 = wx.Button(self, label="Restore")
        self.Bind(wx.EVT_BUTTON, self.OnRestore, self.b7)

        self.srchbox.Add(self.b6, 0, wx.ALL, 5)
        self.srchbox.Add(self.b5, 0, wx.ALL, 5)
        self.srchbox.Add(self.b7, 0, wx.ALL, 5)

        self.navbox = wx.BoxSizer(wx.HORIZONTAL)
        self.fst = wx.Button(self, label='<<')
        self.lst = wx.Button(self, label='>>')
        self.nxt = wx.Button(self, label='>')
        self.pre = wx.Button(self, label='<')
        self.fst.Bind(wx.EVT_BUTTON, self.OnMovefst)
        self.lst.Bind(wx.EVT_BUTTON, self.OnMovelst)
        self.nxt.Bind(wx.EVT_BUTTON, self.OnMovenxt)
        self.pre.Bind(wx.EVT_BUTTON, self.OnMovepre)

        self.navbox.Add(self.fst, 0, wx.ALL, 5)
        self.navbox.Add(self.pre, 0, wx.ALL, 5)
        self.navbox.Add(self.nxt, 0, wx.ALL, 5)
        self.navbox.Add(self.lst, 0, wx.ALL, 5)

        self.numbox = wx.BoxSizer(wx.HORIZONTAL)
        self.recnum1 = wx.StaticText(self, label='Record ',
                                     style=wx.ALIGN_LEFT)
        self.recnum1.SetForegroundColour((255, 0, 0))

        self.recnum2 = wx.StaticText(self, label=str(self.rec_num+1),
                                     style=wx.ALIGN_LEFT)
        self.recnum2.SetForegroundColour((255, 0, 0))
        self.recnum3 = wx.StaticText(self, label='/ '+str(len(self.data)),
                                     style=wx.ALIGN_LEFT)
        self.recnum3.SetForegroundColour((255, 0, 0))
        self.numbox.Add(self.recnum1, 0, wx.ALL, 5)
        self.numbox.Add(self.recnum2, 0, wx.ALL, 5)
        self.numbox.Add(self.recnum3, 0, wx.ALL, 5)

        self.Sizer.Add(self.btnbox, 0, wx.ALIGN_CENTER | wx.TOP, 10)
        self.Sizer.Add(self.navbox, 0, wx.ALIGN_CENTER | wx.TOP, 10)
        self.Sizer.Add(self.srchbox, 0, wx.ALIGN_CENTER | wx.TOP, 10)
        self.Sizer.Add(self.numbox, 0, wx.ALIGN_CENTER | wx.TOP, 10)
        self.b4.SetFocus()

        self.FillScreen()

        # add these following lines to child form
        self.CenterOnParent()
        self.GetParent().Enable(False)
        self.Show(True)
        self.__eventLoop = wx.GUIEventLoop()
        self.__eventLoop.Run()

    def OnMovefst(self, evt):
        self.rec_num = 0
        self.FillScreen()

    def OnMovelst(self, evt):
        self.rec_num = len(self.data)-1
        if self.rec_num < 0:
            self.rec_num = 0
        self.FillScreen()

    def OnMovenxt(self, evt):
        if len(self.data) != 0:
            self.rec_num += 1
        if self.rec_num == len(self.data):
            self.rec_num = 0
        self.FillScreen()

    def OnMovepre(self, evt):
        if len(self.data) != 0:
            self.rec_num -= 1
        if self.rec_num < 0:
            self.rec_num = len(self.data)-1
        self.FillScreen()

    def ImportHTML(self, evt):
        self.Html_convert()

    def Html_convert(self):
        self.currentDirectory = os.getcwd()

        # show the html file in browser
        wildcard = "HTML file (*.html)|*.html"
        msg = 'Select HTML File to View'
        fylnm = ''

        dlg = wx.FileDialog(
            self, message=msg,
            defaultFile="",
            wildcard=wildcard,
            style=wx.FD_OPEN |
            wx.FD_FILE_MUST_EXIST)
        if dlg.ShowModal() == wx.ID_OK:
            fylnm = dlg.GetPath()

        dlg.Destroy()

        if fylnm != '':
            with open(fylnm) as f:
                soup = BeautifulSoup(f, "html.parser")

            id_txt = []
            vl_txt = []
            id_txtar = []
            vl_txtar = []
            chk_chkd = []

            # get the name of the form for proper selection of pdf formate
            titl = [ttl.get_text() for ttl in soup.select('h1')][0]

            # get all the tex input values
            for item in soup.find_all("input", {"type": "text"}):
                id_txt.append(item.get('id'))
                if item.get('value') is None:
                    vl_txt.append('TBD')
                else:
                    vl_txt.append(item.get('value'))
            txt_inpt = dict(zip(id_txt, vl_txt))
            txt_boxes = txt_inpt

            # get all the text area input
            for item in soup.find_all('textarea'):
                id_txtar.append(item.get('id'))
                if item.contents != []:
                    vl_txtar.append(item.contents[0])
                else:
                    vl_txtar.append('')
            txtar_inpt = dict(zip(id_txtar, vl_txtar))
            txt_area = txtar_inpt

            # collect the checkboxes which are checked only
            for item in soup.find_all('input', checked=True):
                chk_chkd.append(item.get('id'))
            chkd_boxes = chk_chkd

            # use the form title to select the proper convertion program
            if titl.find('Specification') != -1:
                rptdta = [txt_boxes, txt_area, chkd_boxes]
                self.FillScreen(rptdta)
                self.b2.Enable()
            else:
                msg = ('''This is not a valid Material Substitution Record\n
                generated using this pipe specification software''')
                cpt = 'Unable to import report'
                dlg = wx.MessageDialog(self, msg, cpt, wx.OK |
                                       wx.ICON_INFORMATION)
                dlg.ShowModal()
                dlg.Destroy()

    def FillScreen(self, htmldta=None):
        # all the IDs for the various tables making up the package
        if htmldta is None:
            # this is filling the screen with database information
            if len(self.data) == 0:
                self.recnum2.SetLabel(str(self.rec_num))
                return
            else:
                recrd = self.data[self.rec_num]
            self.RcdIDtxt.ChangeValue(str(recrd[0]))
            self.TagIDtxt.ChangeValue(str(recrd[1]))
            self.text1.ChangeValue(str(recrd[2]))
            self.text2.ChangeValue(str(recrd[3]))
            self.text3.ChangeValue(str(recrd[4]))
            self.Descrpttxt.ChangeValue(str(recrd[5]))
            self.text4.ChangeValue(str(recrd[6]))
            self.text5.ChangeValue(str(recrd[7]))
            self.text6.ChangeValue(str(recrd[8]))
            self.text7.ChangeValue(str(recrd[9]))
            self.text8.ChangeValue(str(recrd[10]))
            self.text9.ChangeValue(str(recrd[11]))
            self.Reprtxt.ChangeValue(str(recrd[12]))
            self.text10.ChangeValue(str(recrd[13]))
            self.text11.ChangeValue(str(recrd[15]))
            self.text12.ChangeValue(str(recrd[14]))
            self.text13.ChangeValue(str(recrd[16]))
            self.text14.ChangeValue(str(recrd[17]))
            self.text15.ChangeValue(str(recrd[19]))
            self.text16.ChangeValue(str(recrd[18]))
            self.text17.ChangeValue(str(recrd[20]))

            self.recnum2.SetLabel(str(self.rec_num+1))

        else:
            # this fills the form with html data
            self.Descrpttxt.ChangeValue(htmldta[1]['ta_1'])
            self.Reprtxt.ChangeValue(htmldta[1]['ta_2'])

            self.text1.ChangeValue(htmldta[0]['t_1'])
            self.text2.ChangeValue(htmldta[0]['t_3'])
            self.text3.ChangeValue(htmldta[0]['t_4'])
            self.text4.ChangeValue(htmldta[0]['t_5'])
            self.text5.ChangeValue(htmldta[0]['t_6'])
            self.text6.ChangeValue(htmldta[0]['t_7'])
            self.text7.ChangeValue(htmldta[0]['t_8'])
            self.text8.ChangeValue(htmldta[0]['t_9'])
            self.text9.ChangeValue(htmldta[0]['t_10'])
            self.text10.ChangeValue(htmldta[0]['t_11'])
            self.text11.ChangeValue(htmldta[0]['t_12'])
            self.text12.ChangeValue(htmldta[0]['t_15'])
            self.text13.ChangeValue(htmldta[0]['t_16'])
            self.text14.ChangeValue(htmldta[0]['t_13'])
            self.text15.ChangeValue(htmldta[0]['t_14'])
            self.text16.ChangeValue(htmldta[0]['t_17'])
            self.text17.ChangeValue(htmldta[0]['t_18'])

            htmldta = []

    def ValData(self):
        NoData = 0

        if self.text1.GetValue() == '':
            msg = 'Requestor'
            NoData = 1
        elif self.text5.GetValue() == '':
            msg = 'Max. Operating Pressure'
            NoData = 1
        elif self.text6.GetValue() == '':
            msg = 'Max. Operating Temperature'
            NoData = 1
        elif self.text7.GetValue() == '':
            msg = 'Max. Design Pressure'
            NoData = 1
        elif self.text8.GetValue() == '':
            msg = 'Max. Design Temperature'
            NoData = 1
        elif self.text9.GetValue() == '':
            msg = 'Min. Design Temperature'
            NoData = 1
        elif self.Reprtxt.GetValue() == '':
            msg = 'Description of Fluid'
            NoData = 1
        elif self.TagIDtxt.GetValue() == '':
            msg = 'Tag ID'
            NoData = 1

        if NoData == 1:
            wx.MessageBox('Value needed for;\n' + msg +
                          'to complete information.',
                          'Missing Data', wx.OK | wx.ICON_INFORMATION)
            return False
        else:
            return True

    def OnAddRec(self, evt):
        # check first that data is all present and clean up incomplete boxes
        check = self.ValData()
        if check:
            SQL_step = 3

            choice1 = ('1) Save this as a new ' + self.frmtitle +
                       ' Specification')
            choice2 = ('2) Update the existing ' + self.frmtitle +
                       ' Specification with this data')
            txt1 = ('NOTE: Updating this information will be\nreflected\
                     in all associated ' + self.frmtitle)
            txt2 = (' Specifications!\nRecommendation is to save as a new\
                     specification.\n\n\tHow do you want to proceed?')

            SQL_Dialog = wx.SingleChoiceDialog(self, txt1+txt2,
                                               'Information Has Changed',
                                               [choice1, choice2],
                                               style=wx.CHOICEDLG_STYLE)
            if SQL_Dialog.ShowModal() == wx.ID_OK:
                SQL_step = SQL_Dialog.GetSelection()
            SQL_Dialog.Destroy()

            self.AddRec(SQL_step)
            self.MainSQL = 'SELECT * FROM ' + self.tblname
            self.data = Dbase().Dsqldata(self.MainSQL)

            if SQL_step == 0:
                self.rec_num = len(self.data)-1
            self.FillScreen()
            self.recnum3.SetLabel('/ '+str(len(self.data)))

    def AddRec(self, SQL_step):
        realnames = []
        ValueList = []

        New_ID = cursr.execute(
            'SELECT MAX(ReportID) FROM NewPipeSpec').fetchone()
        if New_ID[0] is None:
            Max_ID = '1'
        else:
            Max_ID = str(New_ID[0]+1)
        for item in Dbase().Dcolinfo(self.tblname):
            realnames.append(item[1])

        ValueList.append(Max_ID)
        ValueList.append(self.TagIDtxt.GetValue())
        ValueList.append(self.text1.GetValue())
        ValueList.append(self.text2.GetValue())
        ValueList.append(self.text3.GetValue())
        st = self.Descrpttxt.GetValue()
        st = st.replace("'", "").replace('"', '')
        ValueList.append(st)
        ValueList.append(self.text4.GetValue())
        ValueList.append(self.text5.GetValue())
        ValueList.append(self.text6.GetValue())
        ValueList.append(self.text7.GetValue())
        ValueList.append(self.text8.GetValue())
        ValueList.append(self.text9.GetValue())
        st = self.Reprtxt.GetValue()
        st = st.replace("'", "").replace('"', '')
        ValueList.append(st)
        ValueList.append(self.text10.GetValue())
        ValueList.append(self.text12.GetValue())
        ValueList.append(self.text11.GetValue())
        ValueList.append(self.text13.GetValue())
        ValueList.append(self.text14.GetValue())
        ValueList.append(self.text16.GetValue())
        ValueList.append(self.text15.GetValue())
        ValueList.append(self.text17.GetValue())

        if SQL_step == 0:  # enter new record
            CurrentID = Max_ID
            UpQuery = ('INSERT INTO ' + self.tblname + ' VALUES (' + "'"
                       + "','".join(map(str, ValueList)) + "'" + ')')
            Dbase().TblEdit(UpQuery)

        elif SQL_step == 1:  # update edited record
            CurrentID = self.data[self.rec_num][0]
            realnames.remove('ReportID')
            del ValueList[0]

            SQL_str = dict(zip(realnames, ValueList))
            Update_str = ", ".join(["%s='%s'" % (k, v)
                                    for k, v in SQL_str.items()])
            UpQuery = ('UPDATE ' + self.tblname + ' SET ' + Update_str +
                       ' WHERE ReportID = ' + str(CurrentID))
            Dbase().TblEdit(UpQuery)

        elif SQL_step == 3:
            return

        self.b2.Disable()

        return CurrentID

    def OnSelect2(self, evt):
        self.b2.Enable()

    def OnDelete(self, evt):
        recrd = self.data[self.rec_num][0]

        try:
            Dbase().TblDelete(self.tblname, recrd, 'ReportID')
            self.MainSQL = 'SELECT * FROM ' + self.tblname
            self.data = Dbase().Dsqldata(self.MainSQL)
            self.rec_num -= 1
            if self.rec_num < 0:
                self.rec_num = len(self.data)-1
            self.FillScreen()
            self.recnum3.SetLabel('/ '+str(len(self.data)))
        except sqlite3.IntegrityError:
            wx.MessageBox("This Record is associated"
                          " with\nother tables and cannot be deleted!",
                          "Cannot Delete",
                          wx.OK | wx.ICON_INFORMATION)

    def OnSearch(self, evt):
        sqlstr = 'SELECT * FROM NewPipeSpec'
        frst = 0

        if self.TagIDtxt.GetValue() != '':
            sqlstr = (sqlstr + " WHERE TagID LIKE '%" +
                      self.TagIDtxt.GetValue() + "%' COLLATE NOCASE")
            frst = 1

        if self.text1.GetValue() != '':
            if frst:
                str_and = " OR "
            else:
                str_and = " WHERE "
            sqlstr = (sqlstr + str_and + "Requestor LIKE '%" +
                      self.text1.GetValue() + "%' COLLATE NOCASE")
            frst = 1

        if self.text2.GetValue() != '':
            if frst:
                str_and = " OR "
            else:
                str_and = " WHERE "
            sqlstr = (sqlstr + str_and + "IssueDate LIKE '%" +
                      self.text2.GetValue() + "%' COLLATE NOCASE")

        self.data = Dbase().Dsqldata(sqlstr)
        self.rec_num = 0
        self.recnum3.SetLabel('/ '+str(len(self.data)))
        self.FillScreen()

    def OnClear(self, evt):
        self.RcdIDtxt.ChangeValue('')
        self.TagIDtxt.ChangeValue('')
        self.text1.ChangeValue('')
        self.text2.ChangeValue('')
        self.text3.ChangeValue('')
        self.Descrpttxt.ChangeValue('')
        self.text4.ChangeValue('')
        self.text5.ChangeValue('')
        self.text6.ChangeValue('')
        self.text7.ChangeValue('')
        self.text8.ChangeValue('')
        self.text9.ChangeValue('')
        self.Reprtxt.ChangeValue('')
        self.text10.ChangeValue('')
        self.text11.ChangeValue('')
        self.text12.ChangeValue('')
        self.text13.ChangeValue('')
        self.text14.ChangeValue('')
        self.text15.ChangeValue('')
        self.text16.ChangeValue('')
        self.text17.ChangeValue('')

    def OnRestore(self, evt):
        self.data = Dbase().Dsqldata(self.MainSQL)
        self.recnum3.SetLabel('/ '+str(len(self.data)))
        self.FillScreen()

    def OnClose(self, evt):
        self.GetParent().Enable(True)   # add for child form
        self.__eventLoop.Exit()        # add for child form
        self.Destroy()


class NCRImport(wx.Frame):
    '''Routine to build form and populate grid'''
    def __init__(self, parent):

        self.frmtitle = 'Non-conformance Report'
        super(NCRImport, self).__init__(parent, title=self.frmtitle,
                                        size=(880, 890))

        self.Bind(wx.EVT_CLOSE, self.OnClose)

        self.parent = parent   # add for child form

        self.columnames = []
        self.rec_num = 0
        self.lctrls = []
        self.data = []
        self.MainSQL = ''
        self.tblname = 'NonConformance'

        self.MainSQL = 'SELECT * FROM NonConformance'
        self.data = Dbase().Dsqldata(self.MainSQL)
        # specify which listbox column to display in the combobox
        self.showcol = int

        self.InitUI()

    def InitUI(self):
        # set the Sizer property (same as SetSizer)
        self.Sizer = wx.BoxSizer(wx.VERTICAL)

        self.IDsizer = wx.BoxSizer(wx.HORIZONTAL)
        RcdIDlbl = wx.StaticText(self, label='Record ID',
                                 style=wx.ALIGN_LEFT)
        RcdIDlbl.SetForegroundColour((255, 0, 0))
        self.RcdIDtxt = wx.TextCtrl(self, size=(80, 33), value='',
                                    style=wx.TE_CENTER)
        self.RcdIDtxt.Enable(False)
        TagIDlbl = wx.StaticText(self, label='Tags',
                                 style=wx.ALIGN_LEFT)
        TagIDlbl.SetForegroundColour((255, 0, 0))
        self.TagIDtxt = wx.TextCtrl(self, size=(200, 33), value='',
                                    style=wx.TE_CENTER)
        St2 = wx.StaticText(self, label='* S',
                            style=wx.ALIGN_LEFT)
        St2.SetForegroundColour((255, 0, 0))
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.TagIDtxt)
        self.IDsizer.Add((150, 10))
        self.IDsizer.Add(RcdIDlbl, 0, wx.ALL | wx.CENTER, 10)
        self.IDsizer.Add(self.RcdIDtxt, 0, wx.TOP | wx.BOTTOM | wx.LEFT, 10)
        self.IDsizer.Add((150, 10))
        self.IDsizer.Add(TagIDlbl, 0, wx.ALL | wx.CENTER, 10)
        self.IDsizer.Add(self.TagIDtxt, 0, wx.TOP | wx.BOTTOM | wx.LEFT, 10)
        self.IDsizer.Add(St2, 0, wx.CENTER)

        self.textsizer = wx.BoxSizer(wx.HORIZONTAL)

        self.notesizer2 = wx.BoxSizer(wx.VERTICAL)
        note1 = wx.StaticText(self, label='Job Number',
                              style=wx.ALIGN_LEFT)
        note1.SetForegroundColour((255, 0, 0))

        note2 = wx.StaticText(self, label='Requestor',
                              style=wx.ALIGN_LEFT)
        note2.SetForegroundColour((255, 0, 0))

        note3 = wx.StaticText(self, label='Fabricator',
                              style=wx.ALIGN_LEFT)
        note3.SetForegroundColour((255, 0, 0))
        self.notesizer2.Add(note1, 0)
        self.notesizer2.Add((10, 25))
        self.notesizer2.Add(note2, 0)
        self.notesizer2.Add((10, 25))
        self.notesizer2.Add(note3, 0)

        self.notesizer3 = wx.BoxSizer(wx.VERTICAL)
        note4 = wx.StaticText(self, label='Issue date',
                              style=wx.ALIGN_RIGHT)
        note4.SetForegroundColour((255, 0, 0))

        note5 = wx.StaticText(self, label='Pipe Spec Code',
                              style=wx.ALIGN_LEFT)
        note5.SetForegroundColour((255, 0, 0))

        note6 = wx.StaticText(self, label='Line Number',
                              style=wx.ALIGN_RIGHT)
        note6.SetForegroundColour((255, 0, 0))
        self.notesizer3.Add(note4, 0, wx.ALIGN_RIGHT | wx.RIGHT, 5)
        self.notesizer3.Add((10, 25))
        self.notesizer3.Add(note5, 0)
        self.notesizer3.Add((10, 25))
        self.notesizer3.Add(note6, 0, wx.ALIGN_RIGHT | wx.RIGHT, 5)

        self.txtsizer2 = wx.BoxSizer(wx.VERTICAL)
        self.text1 = wx.TextCtrl(self, size=(100, 33), value='',
                                 style=wx.TE_CENTER)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text1)

        self.text2 = wx.TextCtrl(self, size=(100, 33), value='',
                                 style=wx.TE_CENTER)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text2)

        self.text3 = wx.TextCtrl(self, size=(100, 33), value='',
                                 style=wx.TE_CENTER)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text3)
        self.txtsizer2.Add(self.text1, 0)
        self.txtsizer2.Add(self.text2, 0, wx.TOP | wx.BOTTOM, 10)
        self.txtsizer2.Add(self.text3, 0)

        starszr1 = wx.BoxSizer(wx.VERTICAL)
        St3 = wx.StaticText(self, label='* S',
                            style=wx.ALIGN_LEFT)
        St3.SetForegroundColour((255, 0, 0))
        St4 = wx.StaticText(self, label='* S',
                            style=wx.ALIGN_LEFT)
        St4.SetForegroundColour((255, 0, 0))
        St5 = wx.StaticText(self, label='* S',
                            style=wx.ALIGN_LEFT)
        St5.SetForegroundColour((255, 0, 0))
        starszr1.Add(St3, 0, wx.ALIGN_LEFT)
        starszr1.Add(St4, 0, wx.ALIGN_LEFT | wx.TOP | wx.BOTTOM, 25)
        starszr1.Add(St5, 0, wx.ALIGN_LEFT)

        self.txtsizer3 = wx.BoxSizer(wx.VERTICAL)
        self.text4 = wx.TextCtrl(self, size=(100, 33), value='',
                                 style=wx.TE_CENTER)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text4)

        self.text5 = wx.TextCtrl(self, size=(100, 33), value='',
                                 style=wx.TE_CENTER)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text5)

        self.text6 = wx.TextCtrl(self, size=(100, 33), value='',
                                 style=wx.TE_CENTER)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text6)
        self.txtsizer3.Add(self.text4, 0)
        self.txtsizer3.Add(self.text5, 0, wx.TOP | wx.BOTTOM, 10)
        self.txtsizer3.Add(self.text6, 0)

        starszr2 = wx.BoxSizer(wx.VERTICAL)
        St6 = wx.StaticText(self, label='* S',
                            style=wx.ALIGN_LEFT)
        St6.SetForegroundColour((255, 0, 0))
        St7 = wx.StaticText(self, label='* S',
                            style=wx.ALIGN_LEFT)
        St7.SetForegroundColour((255, 0, 0))
        St8 = wx.StaticText(self, label='* S',
                            style=wx.ALIGN_LEFT)
        St8.SetForegroundColour((255, 0, 0))
        starszr2.Add(St6, 0, wx.ALIGN_LEFT)
        starszr2.Add(St7, 0, wx.ALIGN_LEFT | wx.TOP | wx.BOTTOM, 25)
        starszr2.Add(St8, 0, wx.ALIGN_LEFT)

        self.Dscrptsizer = wx.BoxSizer(wx.VERTICAL)
        self.Descrptlbl = wx.StaticText(self, label='Project Description',
                                        style=wx.ALIGN_LEFT)
        self.Descrptlbl.SetForegroundColour((255, 0, 0))
        self.Descrpttxt = wx.TextCtrl(self, size=(350, 90), value='',
                                      style=wx.TE_MULTILINE)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.Descrpttxt)
        self.Dscrptsizer.Add(self.Descrptlbl, 0, wx.ALIGN_LEFT | wx.LEFT, 10)
        self.Dscrptsizer.Add(self.Descrpttxt, 0, wx.ALIGN_RIGHT | wx.ALL, 10)

        self.textsizer.Add(self.notesizer2, 0, wx.LEFT |
                           wx.ALIGN_CENTER_VERTICAL, 10)
        self.textsizer.Add(self.txtsizer2, 0,
                           wx.ALIGN_CENTER_VERTICAL)
        self.textsizer.Add(starszr1, 0, wx.RIGHT |
                           wx.ALIGN_CENTER_VERTICAL, 10)
        self.textsizer.Add(self.notesizer3, 0, wx.LEFT |
                           wx.ALIGN_CENTER_VERTICAL, 10)
        self.textsizer.Add(self.txtsizer3, 0,
                           wx.ALIGN_CENTER_VERTICAL)
        self.textsizer.Add(starszr2, 0, wx.RIGHT |
                           wx.ALIGN_CENTER_VERTICAL, 10)
        self.textsizer.Add(self.Dscrptsizer, 0, wx.LEFT |
                           wx.ALIGN_CENTER_VERTICAL, 10)

        self.Infosizer = wx.BoxSizer(wx.HORIZONTAL)
        self.Docsizer = wx.BoxSizer(wx.VERTICAL)
        Doclbl = wx.StaticText(self, label='Related Documents',
                               style=wx.ALIGN_LEFT)
        Doclbl.SetForegroundColour((255, 0, 0))
        self.Doctxt = wx.TextCtrl(self, size=(400, 80), value='',
                                  style=wx.TE_MULTILINE)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.Doctxt)
        self.Docsizer.Add(Doclbl, 0, wx.ALIGN_LEFT)
        self.Docsizer.Add(self.Doctxt, 0, wx.ALIGN_RIGHT)

        self.NCRsizer = wx.BoxSizer(wx.VERTICAL)
        NCRlbl = wx.StaticText(self, label='Nonconformance Description  *',
                               style=wx.ALIGN_LEFT)
        NCRlbl.SetForegroundColour((255, 0, 0))
        self.NCRtxt = wx.TextCtrl(self, size=(400, 80), value='',
                                  style=wx.TE_MULTILINE)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.NCRtxt)
        self.NCRsizer.Add(NCRlbl, 0, wx.ALIGN_LEFT | wx.LEFT, 10)
        self.NCRsizer.Add(self.NCRtxt, 0, wx.ALIGN_RIGHT)
        self.Infosizer.Add(self.Docsizer, 0)
        self.Infosizer.Add(self.NCRsizer, 0, wx.LEFT, 20)

        self.Fixsizer = wx.BoxSizer(wx.HORIZONTAL)
        self.Prpsdsizer = wx.BoxSizer(wx.VERTICAL)
        self.Prpsdlbl = wx.StaticText(self, label='Proposed Repair',
                                      style=wx.ALIGN_LEFT)
        self.Prpsdlbl.SetForegroundColour((255, 0, 0))
        self.Prpsdtxt = wx.TextCtrl(self, size=(400, 80), value='',
                                    style=wx.TE_MULTILINE)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.Prpsdtxt)
        self.Prpsdsizer.Add(self.Prpsdlbl, 0, wx.ALIGN_LEFT)
        self.Prpsdsizer.Add(self.Prpsdtxt, 0, wx.ALIGN_LEFT)

        self.Chksizer = wx.BoxSizer(wx.VERTICAL)
        self.QCchk = wx.CheckBox(self, label='Was QC Manager Notified?',
                                 style=wx.ALIGN_RIGHT)
        self.QCchk.SetForegroundColour((255, 0, 0))
        self.Bind(wx.EVT_CHECKBOX, self.OnSelect2, self.QCchk)
        self.Opschk = wx.CheckBox(self, label='Was Operations Notified?',
                                  style=wx.ALIGN_RIGHT)
        self.Opschk.SetForegroundColour((255, 0, 0))
        self.Bind(wx.EVT_CHECKBOX, self.OnSelect2, self.Opschk)
        self.Chksizer.Add(self.QCchk, 0, wx.BOTTOM, 15)
        self.Chksizer.Add(self.Opschk, 0)

        self.Fixsizer.Add(self.Prpsdsizer, 0, wx.LEFT | wx.RIGHT, 20)
        self.Fixsizer.Add((60, 10))
        self.Fixsizer.Add(self.Chksizer, 0, wx.ALIGN_CENTER)

        self.Finalsizer = wx.BoxSizer(wx.HORIZONTAL)

        self.notesizer4 = wx.BoxSizer(wx.VERTICAL)
        self.note7 = wx.StaticText(self, label='Name of Signing\nQC Manager',
                                   style=wx.ALIGN_LEFT)
        self.note7.SetForegroundColour((255, 0, 0))

        self.note8 = wx.StaticText(self, label='Name of Signing\nOperations',
                                   style=wx.ALIGN_LEFT)
        self.note8.SetForegroundColour((255, 0, 0))

        self.note9 = wx.StaticText(self, label='Name of Signing\nRequestor',
                                   style=wx.ALIGN_LEFT)
        self.note9.SetForegroundColour((255, 0, 0))
        self.notesizer4.Add(self.note7, 0)
        self.notesizer4.Add(self.note8, 0, wx.TOP, 15)
        self.notesizer4.Add(self.note9, 0, wx.TOP, 15)

        self.notesizer5 = wx.BoxSizer(wx.VERTICAL)
        note10 = wx.StaticText(self, label='Date',
                               style=wx.ALIGN_LEFT)
        note10.SetForegroundColour((255, 0, 0))

        note11 = wx.StaticText(self, label='Date',
                               style=wx.ALIGN_LEFT)
        note11.SetForegroundColour((255, 0, 0))

        note12 = wx.StaticText(self, label='Date',
                               style=wx.ALIGN_LEFT)
        note12.SetForegroundColour((255, 0, 0))
        self.notesizer5.Add(note10, 0)
        self.notesizer5.Add(note11, 0, wx.TOP | wx.BOTTOM, 30)
        self.notesizer5.Add(note12, 0)

        self.txtsizer4 = wx.BoxSizer(wx.VERTICAL)
        self.text7 = wx.TextCtrl(self, size=(150, 33), value='',
                                 style=wx.TE_CENTER)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text7)

        self.text8 = wx.TextCtrl(self, size=(150, 33), value='',
                                 style=wx.TE_CENTER)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text8)

        self.text9 = wx.TextCtrl(self, size=(150, 33), value='',
                                 style=wx.TE_CENTER)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text9)
        self.txtsizer4.Add(self.text7, 0, wx.ALIGN_CENTER)
        self.txtsizer4.Add(self.text8, 0, wx.ALIGN_CENTER |
                           wx.TOP | wx.BOTTOM, 15)
        self.txtsizer4.Add(self.text9, 0, wx.ALIGN_CENTER)

        self.txtsizer5 = wx.BoxSizer(wx.VERTICAL)
        self.text10 = wx.TextCtrl(self, size=(100, 33), value='',
                                  style=wx.TE_CENTER)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text10)

        self.text11 = wx.TextCtrl(self, size=(100, 33), value='',
                                  style=wx.TE_CENTER)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text11)

        self.text12 = wx.TextCtrl(self, size=(100, 33), value='',
                                  style=wx.TE_CENTER)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text12)
        self.txtsizer5.Add(self.text10, 0, wx.ALIGN_CENTER)
        self.txtsizer5.Add(self.text11, 0, wx.ALIGN_CENTER |
                           wx.TOP | wx.BOTTOM, 15)
        self.txtsizer5.Add(self.text12, 0, wx.ALIGN_CENTER)

        self.Reprsizer = wx.BoxSizer(wx.VERTICAL)
        Reprlbl = wx.StaticText(self, label='Final Repair Requirements',
                                style=wx.ALIGN_LEFT)
        Reprlbl.SetForegroundColour((255, 0, 0))
        self.Reprtxt = wx.TextCtrl(self, size=(350, 90), value='',
                                   style=wx.TE_MULTILINE)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.Reprtxt)
        self.Reprsizer.Add(Reprlbl, 0, wx.ALIGN_LEFT | wx.LEFT, 10)
        self.Reprsizer.Add(self.Reprtxt, 0, wx.ALIGN_LEFT)

        self.Finalsizer.Add(self.Reprsizer, 0, wx.ALIGN_LEFT)
        self.Finalsizer.Add((30, 10))
        self.Finalsizer.Add(self.notesizer4, 0, wx.LEFT |
                            wx.ALIGN_CENTER_VERTICAL, 10)
        self.Finalsizer.Add(self.txtsizer4, 0, wx.RIGHT |
                            wx.ALIGN_CENTER_VERTICAL, 10)
        self.Finalsizer.Add(self.notesizer5, 0, wx.LEFT |
                            wx.ALIGN_CENTER_VERTICAL, 10)
        self.Finalsizer.Add(self.txtsizer5, 0, wx.RIGHT |
                            wx.ALIGN_CENTER_VERTICAL, 10)

        self.Sizer.Add(self.IDsizer, 0, wx.BOTTOM | wx.TOP, 25)
        self.Sizer.Add(self.textsizer, 0, wx.ALIGN_CENTER)
        self.Sizer.Add(self.Infosizer, 0, wx.CENTER | wx.BOTTOM | wx.TOP, 25)
        self.Sizer.Add(self.Fixsizer, 0, wx.LEFT | wx.BOTTOM, 25)
        self.Sizer.Add(self.Finalsizer, 0, wx.ALIGN_CENTER | wx.BOTTOM, 25)

        self.btnbox = wx.BoxSizer(wx.HORIZONTAL)
        # Add buttons for grid modifications
        self.b1 = wx.Button(self, label="Import HTML")
        self.Bind(wx.EVT_BUTTON, self.ImportHTML, self.b1)

        self.b2 = wx.Button(self, label="Save Report Data")
        self.b2.Disable()
        self.Bind(wx.EVT_BUTTON, self.OnAddRec, self.b2)

        self.b3 = wx.Button(self, label="Delete Spec")
        self.Bind(wx.EVT_BUTTON, self.OnDelete, self.b3)

        self.b4 = wx.Button(self, label="Exit")
        self.Bind(wx.EVT_BUTTON, self.OnClose, self.b4)

        # add a button box and place the buttons
        self.btnbox.Add(self.b1, 0, wx.ALL, 5)
        self.btnbox.Add(self.b2, 0, wx.ALL, 5)
        self.btnbox.Add(self.b3, 0, wx.ALL, 5)
        self.btnbox.Add(self.b4, 0, wx.ALL, 5)

        self.srchbox = wx.BoxSizer(wx.HORIZONTAL)

        self.b5 = wx.Button(self, label="Search")
        self.Bind(wx.EVT_BUTTON, self.OnSearch, self.b5)

        self.b6 = wx.Button(self, label="Clear Boxes")
        self.Bind(wx.EVT_BUTTON, self.OnClear, self.b6)

        self.b7 = wx.Button(self, label="Restore")
        self.Bind(wx.EVT_BUTTON, self.OnRestore, self.b7)

        self.srchbox.Add(self.b6, 0, wx.ALL, 5)
        self.srchbox.Add(self.b5, 0, wx.ALL, 5)
        self.srchbox.Add(self.b7, 0, wx.ALL, 5)

        self.navbox = wx.BoxSizer(wx.HORIZONTAL)
        self.fst = wx.Button(self, label='<<')
        self.lst = wx.Button(self, label='>>')
        self.nxt = wx.Button(self, label='>')
        self.pre = wx.Button(self, label='<')
        self.fst.Bind(wx.EVT_BUTTON, self.OnMovefst)
        self.lst.Bind(wx.EVT_BUTTON, self.OnMovelst)
        self.nxt.Bind(wx.EVT_BUTTON, self.OnMovenxt)
        self.pre.Bind(wx.EVT_BUTTON, self.OnMovepre)

        lbl1 = '* - indicates required field'
        lbl2 = '\nS - indicates searchable field'
        note = wx.StaticText(self,
                             label=lbl1 + lbl2,
                             style=wx.ALIGN_LEFT)
        note.SetForegroundColour((255, 0, 0))

        self.navbox.Add(self.fst, 0, wx.ALL, 5)
        self.navbox.Add(self.pre, 0, wx.ALL, 5)
        self.navbox.Add(self.nxt, 0, wx.ALL, 5)
        self.navbox.Add(self.lst, 0, wx.ALL, 5)
        self.navbox.Add(note, 0, wx.LEFT, 20)

        self.numbox = wx.BoxSizer(wx.HORIZONTAL)
        self.recnum1 = wx.StaticText(self, label='Record ',
                                     style=wx.ALIGN_LEFT)
        self.recnum1.SetForegroundColour((255, 0, 0))

        self.recnum2 = wx.StaticText(self, label=str(self.rec_num+1),
                                     style=wx.ALIGN_LEFT)
        self.recnum2.SetForegroundColour((255, 0, 0))
        self.recnum3 = wx.StaticText(self, label='/ '+str(len(self.data)),
                                     style=wx.ALIGN_LEFT)
        self.recnum3.SetForegroundColour((255, 0, 0))
        self.numbox.Add(self.recnum1, 0, wx.ALL, 5)
        self.numbox.Add(self.recnum2, 0, wx.ALL, 5)
        self.numbox.Add(self.recnum3, 0, wx.ALL, 5)

        self.Sizer.Add(self.srchbox, 0, wx.ALIGN_CENTER)
        self.Sizer.Add(self.btnbox, 0, wx.ALIGN_CENTER | wx.TOP, 10)
        self.Sizer.Add(self.navbox, 0, wx.ALIGN_CENTER | wx.TOP, 10)
        self.Sizer.Add(self.numbox, 0, wx.ALIGN_CENTER | wx.TOP, 10)
        self.b4.SetFocus()

        self.FillScreen()

        # add these following lines to child form
        self.CenterOnParent()
        self.GetParent().Enable(False)
        self.Show(True)
        self.__eventLoop = wx.GUIEventLoop()
        self.__eventLoop.Run()

    def OnMovefst(self, evt):
        self.rec_num = 0
        self.FillScreen()

    def OnMovelst(self, evt):
        self.rec_num = len(self.data)-1
        if self.rec_num < 0:
            self.rec_num = 0
        self.FillScreen()

    def OnMovenxt(self, evt):
        if len(self.data) != 0:
            self.rec_num += 1
        if self.rec_num == len(self.data):
            self.rec_num = 0
        self.FillScreen()

    def OnMovepre(self, evt):
        if len(self.data) != 0:
            self.rec_num -= 1
        if self.rec_num < 0:
            self.rec_num = len(self.data)-1
        self.FillScreen()

    def ImportHTML(self, evt):
        self.Html_convert()

    def Html_convert(self):
        self.currentDirectory = os.getcwd()

        # show the html file in browser
        wildcard = "HTML file (*.html)|*.html"
        msg = 'Select HTML File to View'
        fylnm = ''

        dlg = wx.FileDialog(
            self, message=msg,
            defaultFile="",
            wildcard=wildcard,
            style=wx.FD_OPEN |
            wx.FD_FILE_MUST_EXIST)
        if dlg.ShowModal() == wx.ID_OK:
            fylnm = dlg.GetPath()

        dlg.Destroy()

        if fylnm != '':
            with open(fylnm) as f:
                soup = BeautifulSoup(f, "html.parser")

            id_txt = []
            vl_txt = []
            id_txtar = []
            vl_txtar = []
            chk_chkd = []

            # get the name of the form for proper selection of pdf formate
            titl = [ttl.get_text() for ttl in soup.select('h1')][0]

            # get all the tex input values
            for item in soup.find_all("input", {"type": "text"}):
                id_txt.append(item.get('id'))
                if item.get('value') is None:
                    vl_txt.append('TBD')
                else:
                    vl_txt.append(item.get('value'))
            txt_inpt = dict(zip(id_txt, vl_txt))
            txt_boxes = txt_inpt

            # get all the text area input
            for item in soup.find_all('textarea'):
                id_txtar.append(item.get('id'))
                if item.contents != []:
                    vl_txtar.append(item.contents[0])
                else:
                    vl_txtar.append('')
            txtar_inpt = dict(zip(id_txtar, vl_txtar))
            txt_area = txtar_inpt

            # collect the checkboxes which are checked only
            for item in soup.find_all('input', checked=True):
                chk_chkd.append(item.get('id'))
            chkd_boxes = chk_chkd

            # use the form title to select the proper convertion program
            if titl.find('Nonconformance') != -1:
                rptdta = [txt_boxes, txt_area, chkd_boxes]
                self.FillScreen(rptdta)
                self.b2.Enable()
            else:
                msg = ('''This is not a valid Nonconformance Report\n
                generated using this pipe specification software''')
                cpt = 'Unable to import report'
                dlg = wx.MessageDialog(self, msg, cpt, wx.OK |
                                       wx.ICON_INFORMATION)
                dlg.ShowModal()
                dlg.Destroy()

    def FillScreen(self, htmldta=None):
        # all the IDs for the various tables making up the package
        if htmldta is None:
            # this is filling the screen with database information
            if len(self.data) == 0:
                self.recnum2.SetLabel(str(self.rec_num))
                return
            else:
                recrd = self.data[self.rec_num]
            self.RcdIDtxt.ChangeValue(str(recrd[0]))
            self.text1.ChangeValue(str(recrd[3]))
            self.text2.ChangeValue(str(recrd[4]))
            self.text3.ChangeValue(str(recrd[5]))
            self.text4.ChangeValue(str(recrd[6]))
            self.text5.ChangeValue(str(recrd[7]))
            self.text6.ChangeValue(str(recrd[8]))
            self.text7.ChangeValue(str(recrd[15]))
            self.text8.ChangeValue(str(recrd[17]))
            self.text9.ChangeValue(str(recrd[19]))
            self.text10.ChangeValue(str(recrd[16]))
            self.text11.ChangeValue(str(recrd[18]))
            self.text12.ChangeValue(str(recrd[20]))
            self.TagIDtxt.ChangeValue(str(recrd[1]))
            self.Descrpttxt.ChangeValue(str(recrd[2]))
            self.Doctxt.ChangeValue(str(recrd[9]))
            self.NCRtxt.ChangeValue(str(recrd[10]))
            self.Prpsdtxt.ChangeValue(str(recrd[11]))
            self.Reprtxt.ChangeValue(str(recrd[14]))

            if recrd[12] in (True, 'True', '1', 1):
                self.QCchk.SetValue(1)
            else:
                self.QCchk.SetValue(0)

            if recrd[13] in (True, 'True', '1', 1):
                self.Opschk.SetValue(1)
            else:
                self.Opschk.SetValue(0)
            self.recnum2.SetLabel(str(self.rec_num+1))

        else:
            # this fills the form with html data
            self.Descrpttxt.ChangeValue(htmldta[1]['ta_1'])
            self.Doctxt.ChangeValue(htmldta[1]['ta_2'])
            self.NCRtxt.ChangeValue(htmldta[1]['ta_3'])
            self.Prpsdtxt.ChangeValue(htmldta[1]['ta_4'])
            self.Reprtxt.ChangeValue(htmldta[1]['ta_5'])

            self.text1.ChangeValue(htmldta[0]['t_1'])
            self.text2.ChangeValue(htmldta[0]['t_2'])
            self.text3.ChangeValue(htmldta[0]['t_4'])
            self.text4.ChangeValue(htmldta[0]['t_5'])
            self.text5.ChangeValue(htmldta[0]['t_6'])
            self.text6.ChangeValue(htmldta[0]['t_7'])
            self.text7.ChangeValue(htmldta[0]['t_8'])
            self.text8.ChangeValue(htmldta[0]['t_10'])
            self.text9.ChangeValue(htmldta[0]['t_12'])
            self.text10.ChangeValue(htmldta[0]['t_9'])
            self.text11.ChangeValue(htmldta[0]['t_11'])
            self.text12.ChangeValue(htmldta[0]['t_13'])

            if htmldta[2] != []:
                for item in htmldta[2]:
                    if item == 'chk_1':
                        self.QCchk.SetValue(True)
                    if item == 'chk_2':
                        self.Opschk.SetValue(True)
            htmldta = []

    def ValData(self):
        NoData = 0

        if self.Descrpttxt.GetValue() == '':
            msg = 'Project Description'
            NoData = 1
        elif self.text1.GetValue() == '':
            msg = 'Job Number'
            NoData = 1
        elif self.text2.GetValue() == '':
            msg = 'Requestor'
            NoData = 1
        elif self.text3.GetValue() == '':
            msg = 'Fabricator'
            NoData = 1
        elif self.text5.GetValue() == '':
            msg = 'Pipe Specification Code'
            NoData = 1
        elif self.text6.GetValue() == '':
            msg = 'Line Number'
            NoData = 1
        elif self.NCRtxt.GetValue() == '':
            msg = 'Nonconformance Description'
            NoData = 1
        elif self.TagIDtxt.GetValue() == '':
            msg = 'Tag ID'
            NoData = 1

        if NoData == 1:
            wx.MessageBox('Value needed for;\n' + msg +
                          'to complete information.',
                          'Missing Data', wx.OK | wx.ICON_INFORMATION)
            return False
        else:
            return True

    def OnAddRec(self, evt):
        # check first that data is all present and clean up incomplete boxes
        check = self.ValData()
        if check:
            SQL_step = 3

            choice1 = ('1) Save this as a new ' + self.frmtitle +
                       ' Specification')
            choice2 = ('2) Update the existing ' + self.frmtitle +
                       ' Specification with this data')
            txt1 = ('NOTE: Updating this information will be\nreflected\
                     in all associated ' + self.frmtitle)
            txt2 = (' Specifications!\nRecommendation is to save as a new\
                     specification.\n\n\tHow do you want to proceed?')

            SQL_Dialog = wx.SingleChoiceDialog(self, txt1+txt2,
                                               'Information Has Changed',
                                               [choice1, choice2],
                                               style=wx.CHOICEDLG_STYLE)
            if SQL_Dialog.ShowModal() == wx.ID_OK:
                SQL_step = SQL_Dialog.GetSelection()
            SQL_Dialog.Destroy()

            self.AddRec(SQL_step)
            self.MainSQL = 'SELECT * FROM ' + self.tblname
            self.data = Dbase().Dsqldata(self.MainSQL)

            if SQL_step == 0:
                self.rec_num = len(self.data)-1
            self.FillScreen()
            self.recnum3.SetLabel('/ '+str(len(self.data)))

    def AddRec(self, SQL_step):
        realnames = []
        ValueList = []

        New_ID = cursr.execute(
            'SELECT MAX(ReportID) FROM NonConformance').fetchone()
        if New_ID[0] is None:
            Max_ID = '1'
        else:
            Max_ID = str(New_ID[0]+1)
        for item in Dbase().Dcolinfo(self.tblname):
            realnames.append(item[1])

        ValueList.append(Max_ID)
        ValueList.append(self.TagIDtxt.GetValue())
        st = self.Descrpttxt.GetValue()
        st = st.replace("'", "").replace('"', '')
        ValueList.append(st)
        ValueList.append(self.text1.GetValue())
        ValueList.append(self.text2.GetValue())
        ValueList.append(self.text3.GetValue())
        ValueList.append(self.text4.GetValue())
        ValueList.append(self.text5.GetValue())
        ValueList.append(self.text6.GetValue())
        st = self.Doctxt.GetValue()
        st = st.replace("'", "").replace('"', '')
        ValueList.append(st)
        st = self.NCRtxt.GetValue()
        st = st.replace("'", "").replace('"', '')
        ValueList.append(st)
        ValueList.append(self.Prpsdtxt.GetValue())
        ValueList.append(self.QCchk.GetValue())
        ValueList.append(self.Opschk.GetValue())
        st = self.Reprtxt.GetValue()
        st = st.replace("'", "").replace('"', '')
        ValueList.append(st)
        ValueList.append(self.text7.GetValue())
        ValueList.append(self.text8.GetValue())
        ValueList.append(self.text9.GetValue())
        ValueList.append(self.text10.GetValue())
        ValueList.append(self.text11.GetValue())
        ValueList.append(self.text12.GetValue())

        if SQL_step == 0:  # enter new record
            CurrentID = Max_ID
            UpQuery = ('INSERT INTO ' + self.tblname + ' VALUES (' + "'"
                       + "','".join(map(str, ValueList)) + "'" + ')')
            Dbase().TblEdit(UpQuery)

        elif SQL_step == 1:  # update edited record
            CurrentID = self.data[self.rec_num][0]
            realnames.remove('ReportID')
            del ValueList[0]

            SQL_str = dict(zip(realnames, ValueList))
            Update_str = ", ".join(["%s='%s'" % (k, v)
                                    for k, v in SQL_str.items()])
            UpQuery = ('UPDATE ' + self.tblname + ' SET ' + Update_str +
                       ' WHERE ReportID = ' + str(CurrentID))
            Dbase().TblEdit(UpQuery)

        elif SQL_step == 3:
            return

        self.b2.Disable()

        return CurrentID

    def OnSelect2(self, evt):
        self.b2.Enable()

    def OnDelete(self, evt):
        recrd = self.data[self.rec_num][0]

        try:
            Dbase().TblDelete(self.tblname, recrd, 'ReportID')
            self.MainSQL = 'SELECT * FROM ' + self.tblname
            self.data = Dbase().Dsqldata(self.MainSQL)
            self.rec_num -= 1
            if self.rec_num < 0:
                self.rec_num = len(self.data)-1
            self.FillScreen()
            self.recnum3.SetLabel('/ '+str(len(self.data)))
        except sqlite3.IntegrityError:
            wx.MessageBox("This Record is associated"
                          " with\nother tables and cannot be deleted!",
                          "Cannot Delete",
                          wx.OK | wx.ICON_INFORMATION)

    def OnSearch(self, evt):
        sqlstr = 'SELECT * FROM NonConformance'
        frst = 0

        if self.TagIDtxt.GetValue() != '':
            sqlstr = (sqlstr + " WHERE TagID LIKE '%" +
                      self.TagIDtxt.GetValue() + "%' COLLATE NOCASE")
            frst = 1

        if self.text1.GetValue() != '':
            if frst:
                str_and = " OR "
            else:
                str_and = " WHERE "
            sqlstr = (sqlstr + str_and + "Job_Number LIKE '%" +
                      self.text1.GetValue() + "%' COLLATE NOCASE")
            frst = 1

        if self.text2.GetValue() != '':
            if frst:
                str_and = " OR "
            else:
                str_and = " WHERE "
            sqlstr = (sqlstr + str_and + "Requestor LIKE '%" +
                      self.text2.GetValue() + "%' COLLATE NOCASE")
            frst = 1

        if self.text3.GetValue() != '':
            if frst:
                str_and = " OR "
            else:
                str_and = " WHERE "
            sqlstr = (sqlstr + str_and + "Fabricator LIKE '%" +
                      self.text3.GetValue() + "%' COLLATE NOCASE")
            frst = 1

        if self.text4.GetValue() != '':
            if frst:
                str_and = " OR "
            else:
                str_and = " WHERE "
            sqlstr = (sqlstr + str_and + "IssueDate LIKE '%" +
                      self.text4.GetValue() + "%' COLLATE NOCASE")
            frst = 1

        if self.text5.GetValue() != '':
            if frst:
                str_and = " OR "
            else:
                str_and = " WHERE "
            sqlstr = (sqlstr + str_and + "Pipe_Spec_Code LIKE '%" +
                      self.text5.GetValue() + "%' COLLATE NOCASE")
            frst = 1

        if self.text6.GetValue() != '':
            if frst:
                str_and = " OR "
            else:
                str_and = " WHERE "
            sqlstr = (sqlstr + str_and + "Line_Number LIKE '%" +
                      self.text6.GetValue() + "%' COLLATE NOCASE")

        self.data = Dbase().Dsqldata(sqlstr)
        self.rec_num = 0
        self.recnum3.SetLabel('/ '+str(len(self.data)))
        self.FillScreen()

    def OnClear(self, evt):

        self.RcdIDtxt.ChangeValue('')
        self.TagIDtxt.ChangeValue('')
        self.text1.ChangeValue('')
        self.text2.ChangeValue('')
        self.text3.ChangeValue('')
        self.Descrpttxt.ChangeValue('')
        self.text4.ChangeValue('')
        self.text5.ChangeValue('')
        self.text6.ChangeValue('')
        self.text7.ChangeValue('')
        self.text8.ChangeValue('')
        self.text9.ChangeValue('')
        self.Reprtxt.ChangeValue('')
        self.text10.ChangeValue('')
        self.text11.ChangeValue('')
        self.text12.ChangeValue('')
        self.Doctxt.ChangeValue('')
        self.NCRtxt.ChangeValue('')
        self.Prpsdtxt.ChangeValue('')

    def OnRestore(self, evt):
        self.data = Dbase().Dsqldata(self.MainSQL)
        self.recnum3.SetLabel('/ '+str(len(self.data)))
        self.FillScreen()

    def OnClose(self, evt):
        self.GetParent().Enable(True)   # add for child form
        self.__eventLoop.Exit()        # add for child form
        self.Destroy()


class MSRImport(wx.Frame):
    '''Routine to build form and populate grid'''
    def __init__(self, parent):

        self.frmtitle = 'Material Substitution Record'
        super(MSRImport, self).__init__(parent, title=self.frmtitle,
                                        size=(880, 900))

        self.Bind(wx.EVT_CLOSE, self.OnClose)

        self.parent = parent   # add for child form

        self.columnames = []
        self.rec_num = 0
        self.lctrls = []
        self.data = []
        self.MainSQL = ''
        self.tblname = 'MtrSubRcrd'

        self.MainSQL = 'SELECT * FROM MtrSubRcrd'
        self.data = Dbase().Dsqldata(self.MainSQL)
        # specify which listbox column to display in the combobox
        self.showcol = int

        self.InitUI()

    def InitUI(self):
        # set the Sizer property (same as SetSizer)
        self.Sizer = wx.BoxSizer(wx.VERTICAL)

        self.IDsizer = wx.BoxSizer(wx.HORIZONTAL)
        RcdIDlbl = wx.StaticText(self, label='Record ID',
                                 style=wx.ALIGN_LEFT)
        RcdIDlbl.SetForegroundColour((255, 0, 0))
        self.RcdIDtxt = wx.TextCtrl(self, size=(80, 33), value='',
                                    style=wx.TE_CENTER)
        self.RcdIDtxt.Enable(False)
        TagIDlbl = wx.StaticText(self, label='Tags',
                                 style=wx.ALIGN_LEFT)
        TagIDlbl.SetForegroundColour((255, 0, 0))
        self.TagIDtxt = wx.TextCtrl(self, size=(200, 33), value='',
                                    style=wx.TE_CENTER)
        St2 = wx.StaticText(self, label='* S',
                            style=wx.ALIGN_LEFT)
        St2.SetForegroundColour((255, 0, 0))
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.TagIDtxt)
        self.IDsizer.Add((150, 10))
        self.IDsizer.Add(RcdIDlbl, 0, wx.ALL | wx.CENTER, 10)
        self.IDsizer.Add(self.RcdIDtxt, 0, wx.TOP | wx.BOTTOM | wx.LEFT, 10)
        self.IDsizer.Add((150, 10))
        self.IDsizer.Add(TagIDlbl, 0, wx.ALL | wx.CENTER, 10)
        self.IDsizer.Add(self.TagIDtxt, 0, wx.TOP | wx.BOTTOM | wx.LEFT, 10)
        self.IDsizer.Add(St2, 0, wx.CENTER)

        self.textsizer = wx.BoxSizer(wx.HORIZONTAL)

        self.notesizer2 = wx.BoxSizer(wx.VERTICAL)
        note1 = wx.StaticText(self, label='Job Number',
                              style=wx.ALIGN_LEFT)
        note1.SetForegroundColour((255, 0, 0))

        note2 = wx.StaticText(self, label='Requestor',
                              style=wx.ALIGN_LEFT)
        note2.SetForegroundColour((255, 0, 0))

        note3 = wx.StaticText(self, label='Fabricator',
                              style=wx.ALIGN_LEFT)
        note3.SetForegroundColour((255, 0, 0))
        self.notesizer2.Add(note1, 0)
        self.notesizer2.Add((10, 25))
        self.notesizer2.Add(note2, 0)
        self.notesizer2.Add((10, 25))
        self.notesizer2.Add(note3, 0)

        self.notesizer3 = wx.BoxSizer(wx.VERTICAL)
        note4 = wx.StaticText(self, label='Issue date',
                              style=wx.ALIGN_RIGHT)
        note4.SetForegroundColour((255, 0, 0))

        note5 = wx.StaticText(self, label='Pipe Spec Code',
                              style=wx.ALIGN_LEFT)
        note5.SetForegroundColour((255, 0, 0))

        note6 = wx.StaticText(self, label='Line Number',
                              style=wx.ALIGN_RIGHT)
        note6.SetForegroundColour((255, 0, 0))
        self.notesizer3.Add(note4, 0, wx.ALIGN_RIGHT | wx.RIGHT, 5)
        self.notesizer3.Add((10, 25))
        self.notesizer3.Add(note5, 0)
        self.notesizer3.Add((10, 25))
        self.notesizer3.Add(note6, 0, wx.ALIGN_RIGHT | wx.RIGHT, 5)

        self.txtsizer2 = wx.BoxSizer(wx.VERTICAL)
        self.text1 = wx.TextCtrl(self, size=(100, 33), value='',
                                 style=wx.TE_CENTER)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text1)

        self.text2 = wx.TextCtrl(self, size=(100, 33), value='',
                                 style=wx.TE_CENTER)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text2)

        self.text3 = wx.TextCtrl(self, size=(100, 33), value='',
                                 style=wx.TE_CENTER)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text3)
        self.txtsizer2.Add(self.text1, 0)
        self.txtsizer2.Add(self.text2, 0, wx.TOP | wx.BOTTOM, 10)
        self.txtsizer2.Add(self.text3, 0)

        starszr1 = wx.BoxSizer(wx.VERTICAL)
        St3 = wx.StaticText(self, label='* S',
                            style=wx.ALIGN_LEFT)
        St3.SetForegroundColour((255, 0, 0))
        St4 = wx.StaticText(self, label='* S',
                            style=wx.ALIGN_LEFT)
        St4.SetForegroundColour((255, 0, 0))
        St5 = wx.StaticText(self, label='* S',
                            style=wx.ALIGN_LEFT)
        St5.SetForegroundColour((255, 0, 0))
        starszr1.Add(St3, 0, wx.ALIGN_LEFT)
        starszr1.Add(St4, 0, wx.ALIGN_LEFT | wx.TOP | wx.BOTTOM, 25)
        starszr1.Add(St5, 0, wx.ALIGN_LEFT)

        self.txtsizer3 = wx.BoxSizer(wx.VERTICAL)
        self.text4 = wx.TextCtrl(self, size=(100, 33), value='',
                                 style=wx.TE_CENTER)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text4)

        self.text5 = wx.TextCtrl(self, size=(100, 33), value='',
                                 style=wx.TE_CENTER)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text5)

        self.text6 = wx.TextCtrl(self, size=(100, 33), value='',
                                 style=wx.TE_CENTER)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text6)
        self.txtsizer3.Add(self.text4, 0)
        self.txtsizer3.Add(self.text5, 0, wx.TOP | wx.BOTTOM, 10)
        self.txtsizer3.Add(self.text6, 0)

        starszr2 = wx.BoxSizer(wx.VERTICAL)
        St6 = wx.StaticText(self, label='* S',
                            style=wx.ALIGN_LEFT)
        St6.SetForegroundColour((255, 0, 0))
        St7 = wx.StaticText(self, label='* S',
                            style=wx.ALIGN_LEFT)
        St7.SetForegroundColour((255, 0, 0))
        St8 = wx.StaticText(self, label='* S',
                            style=wx.ALIGN_LEFT)
        St8.SetForegroundColour((255, 0, 0))
        starszr2.Add(St6, 0, wx.ALIGN_LEFT)
        starszr2.Add(St7, 0, wx.ALIGN_LEFT | wx.TOP | wx.BOTTOM, 25)
        starszr2.Add(St8, 0, wx.ALIGN_LEFT)

        self.Dscrptsizer = wx.BoxSizer(wx.VERTICAL)
        self.Descrptlbl = wx.StaticText(self, label='Project Description',
                                        style=wx.ALIGN_LEFT)
        self.Descrptlbl.SetForegroundColour((255, 0, 0))
        self.Descrpttxt = wx.TextCtrl(self, size=(350, 90), value='',
                                      style=wx.TE_MULTILINE)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.Descrpttxt)
        self.Dscrptsizer.Add(self.Descrptlbl, 0, wx.ALIGN_LEFT | wx.LEFT, 10)
        self.Dscrptsizer.Add(self.Descrpttxt, 0, wx.ALIGN_RIGHT | wx.ALL, 10)

        self.textsizer.Add(self.notesizer2, 0, wx.LEFT |
                           wx.ALIGN_CENTER_VERTICAL, 20)
        self.textsizer.Add(self.txtsizer2, 0,
                           wx.ALIGN_CENTER_VERTICAL)
        self.textsizer.Add(starszr1, 0, wx.RIGHT |
                           wx.ALIGN_CENTER_VERTICAL, 10)
        self.textsizer.Add(self.notesizer3, 0, wx.LEFT |
                           wx.ALIGN_CENTER_VERTICAL, 10)
        self.textsizer.Add(self.txtsizer3, 0,
                           wx.ALIGN_CENTER_VERTICAL)
        self.textsizer.Add(starszr2, 0, wx.RIGHT |
                           wx.ALIGN_CENTER_VERTICAL, 10)
        self.textsizer.Add(self.Dscrptsizer, 0, wx.LEFT |
                           wx.ALIGN_CENTER_VERTICAL, 10)

        self.Infosizer = wx.BoxSizer(wx.HORIZONTAL)
        self.Docsizer = wx.BoxSizer(wx.VERTICAL)
        Doclbl = wx.StaticText(self, label='Support Documents',
                               style=wx.ALIGN_LEFT)
        Doclbl.SetForegroundColour((255, 0, 0))
        self.Doctxt = wx.TextCtrl(self, size=(400, 80), value='',
                                  style=wx.TE_MULTILINE)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.Doctxt)
        self.Docsizer.Add(Doclbl, 0, wx.ALIGN_LEFT | wx.LEFT, 10)
        self.Docsizer.Add(self.Doctxt, 0, wx.ALIGN_RIGHT | wx.ALL, 10)

        self.MSRsizer = wx.BoxSizer(wx.VERTICAL)
        MSRlbl = wx.StaticText(self, label='Description of Deviation  *',
                               style=wx.ALIGN_LEFT)
        MSRlbl.SetForegroundColour((255, 0, 0))
        self.MSRtxt = wx.TextCtrl(self, size=(400, 80), value='',
                                  style=wx.TE_MULTILINE)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.MSRtxt)
        self.MSRsizer.Add(MSRlbl, 0, wx.ALIGN_LEFT | wx.LEFT, 10)
        self.MSRsizer.Add(self.MSRtxt, 0, wx.ALIGN_RIGHT | wx.ALL, 10)
        self.Infosizer.Add(self.Docsizer, 0, wx.ALL, 10)
        self.Infosizer.Add(self.MSRsizer, 0, wx.ALL, 10)

        self.textsizer8 = wx.BoxSizer(wx.HORIZONTAL)
        self.OpDsizer = wx.BoxSizer(wx.VERTICAL)
        note13 = wx.StaticText(self, label='Design Pressure (psig)',
                               style=wx.ALIGN_LEFT)
        note13.SetForegroundColour((255, 0, 0))

        note14 = wx.StaticText(self, label='Hydro Test Pressure (psig)',
                               style=wx.ALIGN_LEFT)
        note14.SetForegroundColour((255, 0, 0))
        self.OpDsizer.Add(note13, 0, wx.ALIGN_RIGHT | wx.ALL, 10)
        self.OpDsizer.Add(note14, 0, wx.ALIGN_RIGHT | wx.ALL, 10)

        self.OpDsizer2 = wx.BoxSizer(wx.VERTICAL)
        note15 = wx.StaticText(self, label='Design Temp Min (F)',
                               style=wx.ALIGN_LEFT)
        note15.SetForegroundColour((255, 0, 0))
        note16 = wx.StaticText(self, label='Design Temp Max (F)',
                               style=wx.ALIGN_LEFT)
        note16.SetForegroundColour((255, 0, 0))
        self.OpDsizer2.Add(note15, 0, wx.ALIGN_RIGHT | wx.ALL, 10)
        self.OpDsizer2.Add(note16, 0, wx.ALIGN_RIGHT | wx.ALL, 10)

        self.txtsizer6 = wx.BoxSizer(wx.VERTICAL)
        self.text13 = wx.TextCtrl(self, size=(100, 33), value='',
                                  style=wx.TE_CENTER)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text13)

        self.text14 = wx.TextCtrl(self, size=(100, 33), value='',
                                  style=wx.TE_CENTER)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text14)
        self.txtsizer6.Add(self.text13, 0, wx.ALL | wx.ALIGN_CENTER, 5)
        self.txtsizer6.Add(self.text14, 0, wx.ALL | wx.ALIGN_CENTER, 5)

        self.txtsizer7 = wx.BoxSizer(wx.VERTICAL)
        self.text15 = wx.TextCtrl(self, size=(100, 33), value='',
                                  style=wx.TE_CENTER)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text15)

        self.text16 = wx.TextCtrl(self, size=(100, 33), value='',
                                  style=wx.TE_CENTER)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text16)

        self.txtsizer7.Add(self.text15, 0, wx.ALL | wx.ALIGN_CENTER, 5)
        self.txtsizer7.Add(self.text16, 0, wx.ALL | wx.ALIGN_CENTER, 5)

        lbl1 = '* - indicates required field'
        lbl2 = '\nS - indicates searchable field'
        note = wx.StaticText(self,
                             label=lbl1 + lbl2,
                             style=wx.ALIGN_LEFT)
        note.SetForegroundColour((255, 0, 0))

        self.textsizer8.Add(self.OpDsizer, 0, wx.LEFT |
                            wx.ALIGN_CENTER_VERTICAL, 10)
        self.textsizer8.Add(self.txtsizer6, 0, wx.LEFT |
                            wx.ALIGN_CENTER_VERTICAL, 10)
        self.textsizer8.Add(self.OpDsizer2, 0, wx.LEFT |
                            wx.ALIGN_CENTER_VERTICAL, 10)
        self.textsizer8.Add(self.txtsizer7, 0, wx.LEFT |
                            wx.ALIGN_CENTER_VERTICAL, 10)
        self.textsizer8.Add((25, 10))
        self.textsizer8.Add(note, 0, wx.RIGHT |
                            wx.ALIGN_CENTER_VERTICAL, 10)

        self.Finalsizer = wx.BoxSizer(wx.HORIZONTAL)

        self.notesizer4 = wx.BoxSizer(wx.VERTICAL)
        note7 = wx.StaticText(self, label='Name of Signing\nQC Manager',
                              style=wx.ALIGN_LEFT)
        note7.SetForegroundColour((255, 0, 0))

        note8 = wx.StaticText(self, label='Name of Signing\nOperations',
                              style=wx.ALIGN_LEFT)
        note8.SetForegroundColour((255, 0, 0))

        note9 = wx.StaticText(self, label='Name of Signing\nRequestor',
                              style=wx.ALIGN_LEFT)
        note9.SetForegroundColour((255, 0, 0))
        self.notesizer4.Add(note7, 0)
        self.notesizer4.Add(note8, 0, wx.TOP | wx.BOTTOM, 10)
        self.notesizer4.Add(note9, 0)

        self.notesizer5 = wx.BoxSizer(wx.VERTICAL)
        note10 = wx.StaticText(self, label='Date',
                               style=wx.ALIGN_LEFT)
        note10.SetForegroundColour((255, 0, 0))

        note11 = wx.StaticText(self, label='Date',
                               style=wx.ALIGN_LEFT)
        note11.SetForegroundColour((255, 0, 0))

        note12 = wx.StaticText(self, label='Date',
                               style=wx.ALIGN_LEFT)
        note12.SetForegroundColour((255, 0, 0))
        self.notesizer5.Add(note10, 0)
        self.notesizer5.Add(note11, 0, wx.TOP | wx.BOTTOM, 25)
        self.notesizer5.Add(note12, 0)

        self.txtsizer4 = wx.BoxSizer(wx.VERTICAL)
        self.text7 = wx.TextCtrl(self, size=(150, 33), value='',
                                 style=wx.TE_CENTER)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text7)

        self.text8 = wx.TextCtrl(self, size=(150, 33), value='',
                                 style=wx.TE_CENTER)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text8)

        self.text9 = wx.TextCtrl(self, size=(150, 33), value='',
                                 style=wx.TE_CENTER)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text9)
        self.txtsizer4.Add(self.text7, 0, wx.ALL | wx.ALIGN_CENTER, 5)
        self.txtsizer4.Add(self.text8, 0, wx.ALL | wx.ALIGN_CENTER, 5)
        self.txtsizer4.Add(self.text9, 0, wx.ALL | wx.ALIGN_CENTER, 5)

        self.txtsizer5 = wx.BoxSizer(wx.VERTICAL)
        self.text10 = wx.TextCtrl(self, size=(100, 33), value='',
                                  style=wx.TE_CENTER)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text10)

        self.text11 = wx.TextCtrl(self, size=(100, 33), value='',
                                  style=wx.TE_CENTER)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text11)

        self.text12 = wx.TextCtrl(self, size=(100, 33), value='',
                                  style=wx.TE_CENTER)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text12)
        self.txtsizer5.Add(self.text10, 0, wx.ALL | wx.ALIGN_CENTER, 5)
        self.txtsizer5.Add(self.text11, 0, wx.ALL | wx.ALIGN_CENTER, 5)
        self.txtsizer5.Add(self.text12, 0, wx.ALL | wx.ALIGN_CENTER, 5)

        self.Reprsizer = wx.BoxSizer(wx.VERTICAL)
        Reprlbl = wx.StaticText(self, label='Requested Substitution  *',
                                style=wx.ALIGN_LEFT)
        Reprlbl.SetForegroundColour((255, 0, 0))
        self.Reprtxt = wx.TextCtrl(self, size=(350, 100), value='',
                                   style=wx.TE_MULTILINE)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.Reprtxt)
        self.Reprsizer.Add(Reprlbl, 0, wx.ALIGN_LEFT | wx.LEFT, 10)
        self.Reprsizer.Add(self.Reprtxt, 0, wx.ALIGN_RIGHT | wx.ALL, 10)

        self.Finalsizer.Add(self.Reprsizer, 0, wx.LEFT |
                            wx.ALIGN_CENTER_VERTICAL, 10)
        self.Finalsizer.Add(self.notesizer4, 0, wx.LEFT |
                            wx.ALIGN_CENTER_VERTICAL, 10)
        self.Finalsizer.Add(self.txtsizer4, 0, wx.RIGHT |
                            wx.ALIGN_CENTER_VERTICAL, 10)
        self.Finalsizer.Add(self.notesizer5, 0, wx.LEFT |
                            wx.ALIGN_CENTER_VERTICAL, 10)
        self.Finalsizer.Add(self.txtsizer5, 0, wx.RIGHT |
                            wx.ALIGN_CENTER_VERTICAL, 10)

        self.Sizer.Add(self.IDsizer, 0, wx.TOP | wx.BOTTOM, 20)
        self.Sizer.Add(self.textsizer, 0, wx.ALIGN_CENTER)
        self.Sizer.Add(self.Infosizer, 0, wx.TOP | wx.BOTTOM, 20)
        self.Sizer.Add(self.textsizer8, 0)
        self.Sizer.Add(self.Finalsizer, 0, wx.ALIGN_CENTER | wx.TOP |
                       wx.BOTTOM, 20)

        self.btnbox = wx.BoxSizer(wx.HORIZONTAL)
        # Add buttons for grid modifications
        self.b1 = wx.Button(self, label="Import HTML")
        self.Bind(wx.EVT_BUTTON, self.ImportHTML, self.b1)

        self.b2 = wx.Button(self, label="Save Report Data")
        self.b2.Disable()
        self.Bind(wx.EVT_BUTTON, self.OnAddRec, self.b2)

        self.b3 = wx.Button(self, label="Delete Spec")
        self.Bind(wx.EVT_BUTTON, self.OnDelete, self.b3)

        self.b4 = wx.Button(self, label="Exit")
        self.Bind(wx.EVT_BUTTON, self.OnClose, self.b4)

        # add a button box and place the buttons
        self.btnbox.Add(self.b1, 0, wx.ALL, 5)
        self.btnbox.Add(self.b2, 0, wx.ALL, 5)
        self.btnbox.Add(self.b3, 0, wx.ALL, 5)
        self.btnbox.Add(self.b4, 0, wx.ALL, 5)

        self.navbox = wx.BoxSizer(wx.HORIZONTAL)
        self.fst = wx.Button(self, label='<<')
        self.lst = wx.Button(self, label='>>')
        self.nxt = wx.Button(self, label='>')
        self.pre = wx.Button(self, label='<')
        self.fst.Bind(wx.EVT_BUTTON, self.OnMovefst)
        self.lst.Bind(wx.EVT_BUTTON, self.OnMovelst)
        self.nxt.Bind(wx.EVT_BUTTON, self.OnMovenxt)
        self.pre.Bind(wx.EVT_BUTTON, self.OnMovepre)

        self.navbox.Add(self.fst, 0, wx.ALL, 5)
        self.navbox.Add(self.pre, 0, wx.ALL, 5)
        self.navbox.Add(self.nxt, 0, wx.ALL, 5)
        self.navbox.Add(self.lst, 0, wx.ALL, 5)

        self.srchbox = wx.BoxSizer(wx.HORIZONTAL)

        self.b5 = wx.Button(self, label="Search")
        self.Bind(wx.EVT_BUTTON, self.OnSearch, self.b5)

        self.b6 = wx.Button(self, label="Clear Boxes")
        self.Bind(wx.EVT_BUTTON, self.OnClear, self.b6)

        self.b7 = wx.Button(self, label="Restore")
        self.Bind(wx.EVT_BUTTON, self.OnRestore, self.b7)

        self.srchbox.Add(self.b6, 0, wx.ALL, 5)
        self.srchbox.Add(self.b5, 0, wx.ALL, 5)
        self.srchbox.Add(self.b7, 0, wx.ALL, 5)

        self.numbox = wx.BoxSizer(wx.HORIZONTAL)
        self.recnum1 = wx.StaticText(self, label='Record ',
                                     style=wx.ALIGN_LEFT)
        self.recnum1.SetForegroundColour((255, 0, 0))

        self.recnum2 = wx.StaticText(self, label=str(self.rec_num+1),
                                     style=wx.ALIGN_LEFT)
        self.recnum2.SetForegroundColour((255, 0, 0))
        self.recnum3 = wx.StaticText(self, label='/ '+str(len(self.data)),
                                     style=wx.ALIGN_LEFT)
        self.recnum3.SetForegroundColour((255, 0, 0))
        self.numbox.Add(self.recnum1, 0, wx.ALL, 5)
        self.numbox.Add(self.recnum2, 0, wx.ALL, 5)
        self.numbox.Add(self.recnum3, 0, wx.ALL, 5)

        self.Sizer.Add(self.btnbox, 0, wx.ALIGN_CENTER)
        self.Sizer.Add(self.navbox, 0, wx.ALIGN_CENTER | wx.TOP, 10)
        self.Sizer.Add(self.srchbox, 0, wx.ALIGN_CENTER | wx.TOP, 10)
        self.Sizer.Add(self.numbox, 0, wx.ALIGN_CENTER | wx.TOP, 10)

        self.b4.SetFocus()

        self.FillScreen()

        # add these following lines to child form
        self.CenterOnParent()
        self.GetParent().Enable(False)
        self.Show(True)
        self.__eventLoop = wx.GUIEventLoop()
        self.__eventLoop.Run()

    def OnMovefst(self, evt):
        self.rec_num = 0
        self.FillScreen()

    def OnMovelst(self, evt):
        self.rec_num = len(self.data)-1
        if self.rec_num < 0:
            self.rec_num = 0
        self.FillScreen()

    def OnMovenxt(self, evt):
        if len(self.data) != 0:
            self.rec_num += 1
        if self.rec_num == len(self.data):
            self.rec_num = 0
        self.FillScreen()

    def OnMovepre(self, evt):
        if len(self.data) != 0:
            self.rec_num -= 1
        if self.rec_num < 0:
            self.rec_num = len(self.data)-1
        self.FillScreen()

    def ImportHTML(self, evt):
        self.Html_convert()

    def Html_convert(self):
        self.currentDirectory = os.getcwd()

        # show the html file in browser
        wildcard = "HTML file (*.html)|*.html"
        msg = 'Select HTML File to View'
        fylnm = ''

        dlg = wx.FileDialog(
            self, message=msg,
            defaultFile="",
            wildcard=wildcard,
            style=wx.FD_OPEN |
            wx.FD_FILE_MUST_EXIST)
        if dlg.ShowModal() == wx.ID_OK:
            fylnm = dlg.GetPath()

        dlg.Destroy()

        if fylnm != '':
            with open(fylnm) as f:
                soup = BeautifulSoup(f, "html.parser")

            id_txt = []
            vl_txt = []
            id_txtar = []
            vl_txtar = []
            chk_chkd = []

            # get the name of the form for proper selection of pdf formate
            titl = [ttl.get_text() for ttl in soup.select('h1')][0]

            # get all the tex input values
            for item in soup.find_all("input", {"type": "text"}):
                id_txt.append(item.get('id'))
                if item.get('value') is None:
                    vl_txt.append('TBD')
                else:
                    vl_txt.append(item.get('value'))
            txt_inpt = dict(zip(id_txt, vl_txt))
            txt_boxes = txt_inpt

            # get all the text area input
            for item in soup.find_all('textarea'):
                id_txtar.append(item.get('id'))
                if item.contents != []:
                    vl_txtar.append(item.contents[0])
                else:
                    vl_txtar.append('')
            txtar_inpt = dict(zip(id_txtar, vl_txtar))
            txt_area = txtar_inpt

            # collect the checkboxes which are checked only
            for item in soup.find_all('input', checked=True):
                chk_chkd.append(item.get('id'))
            chkd_boxes = chk_chkd

            # use the form title to select the proper convertion program
            if titl.find('Substitution') != -1:
                rptdta = [txt_boxes, txt_area, chkd_boxes]
                self.FillScreen(rptdta)
                self.b2.Enable()
            else:
                msg = ('''This is not a valid Material Substitution Record\n
                generated using this pipe specification software''')
                cpt = 'Unable to import report'
                dlg = wx.MessageDialog(self, msg, cpt, wx.OK |
                                       wx.ICON_INFORMATION)
                dlg.ShowModal()
                dlg.Destroy()

    def FillScreen(self, htmldta=None):
        # all the IDs for the various tables making up the package
        if htmldta is None:
            # this is filling the screen with database information
            if len(self.data) == 0:
                self.recnum2.SetLabel(str(self.rec_num))
                return
            else:
                recrd = self.data[self.rec_num]
            self.RcdIDtxt.ChangeValue(str(recrd[0]))
            self.text1.ChangeValue(str(recrd[3]))
            self.text2.ChangeValue(str(recrd[4]))
            self.text3.ChangeValue(str(recrd[5]))
            self.text4.ChangeValue(str(recrd[6]))
            self.text5.ChangeValue(str(recrd[7]))
            self.text6.ChangeValue(str(recrd[8]))
            self.text7.ChangeValue(str(recrd[16]))
            self.text8.ChangeValue(str(recrd[18]))
            self.text9.ChangeValue(str(recrd[20]))
            self.text10.ChangeValue(str(recrd[17]))
            self.text11.ChangeValue(str(recrd[19]))
            self.text12.ChangeValue(str(recrd[21]))
            self.text13.ChangeValue(str(recrd[11]))
            self.text14.ChangeValue(str(recrd[12]))
            self.text15.ChangeValue(str(recrd[13]))
            self.text16.ChangeValue(str(recrd[14]))
            self.TagIDtxt.ChangeValue(str(recrd[1]))
            self.Descrpttxt.ChangeValue(str(recrd[2]))
            self.Doctxt.ChangeValue(str(recrd[9]))
            self.MSRtxt.ChangeValue(str(recrd[10]))
            self.Reprtxt.ChangeValue(str(recrd[15]))

            self.recnum2.SetLabel(str(self.rec_num+1))

        else:
            # this fills the form with html data
            self.Descrpttxt.ChangeValue(htmldta[1]['ta_1'])
            self.Doctxt.ChangeValue(htmldta[1]['ta_2'])
            self.MSRtxt.ChangeValue(htmldta[1]['ta_3'])
            self.Reprtxt.ChangeValue(htmldta[1]['ta_4'])

            self.text1.ChangeValue(htmldta[0]['t_1'])
            self.text2.ChangeValue(htmldta[0]['t_2'])
            self.text3.ChangeValue(htmldta[0]['t_4'])
            self.text4.ChangeValue(htmldta[0]['t_5'])
            self.text5.ChangeValue(htmldta[0]['t_6'])
            self.text6.ChangeValue(htmldta[0]['t_7'])
            self.text7.ChangeValue(htmldta[0]['t_12'])
            self.text8.ChangeValue(htmldta[0]['t_13'])
            self.text9.ChangeValue(htmldta[0]['t_14'])
            self.text10.ChangeValue(htmldta[0]['t_15'])
            self.text11.ChangeValue(htmldta[0]['t_16'])
            self.text12.ChangeValue(htmldta[0]['t_17'])
            self.text13.ChangeValue(htmldta[0]['t_8'])
            self.text14.ChangeValue(htmldta[0]['t_9'])
            self.text15.ChangeValue(htmldta[0]['t_10'])
            self.text16.ChangeValue(htmldta[0]['t_11'])

            htmldta = []

    def ValData(self):
        NoData = 0

        if self.Descrpttxt.GetValue() == '':
            msg = 'Project Description'
            NoData = 1
        elif self.text1.GetValue() == '':
            msg = 'Job Number'
            NoData = 1
        elif self.text2.GetValue() == '':
            msg = 'Requestor'
            NoData = 1
        elif self.text3.GetValue() == '':
            msg = 'Fabricator'
            NoData = 1
        elif self.text5.GetValue() == '':
            msg = 'Pipe Specification Code'
            NoData = 1
        elif self.text6.GetValue() == '':
            msg = 'Line Number'
            NoData = 1
        elif self.MSRtxt.GetValue() == '':
            msg = 'Description of Deviation'
            NoData = 1
        elif self.TagIDtxt.GetValue() == '':
            msg = 'Tag ID'
            NoData = 1

        if NoData == 1:
            wx.MessageBox('Value needed for;\n' + msg +
                          'to complete information.',
                          'Missing Data', wx.OK | wx.ICON_INFORMATION)
            return False
        else:
            return True

    def OnAddRec(self, evt):
        # check first that data is all present and clean up incomplete boxes
        check = self.ValData()
        if check:
            SQL_step = 3

            choice1 = ('1) Save this as a new ' + self.frmtitle +
                       ' Specification')
            choice2 = ('2) Update the existing ' + self.frmtitle +
                       ' Specification with this data')
            txt1 = ('NOTE: Updating this information will be\nreflected\
                     in all associated ' + self.frmtitle)
            txt2 = (' Specifications!\nRecommendation is to save as a new\
                     specification.\n\n\tHow do you want to proceed?')

            SQL_Dialog = wx.SingleChoiceDialog(self, txt1+txt2,
                                               'Information Has Changed',
                                               [choice1, choice2],
                                               style=wx.CHOICEDLG_STYLE)
            if SQL_Dialog.ShowModal() == wx.ID_OK:
                SQL_step = SQL_Dialog.GetSelection()
            SQL_Dialog.Destroy()

            self.AddRec(SQL_step)
            self.MainSQL = 'SELECT * FROM ' + self.tblname
            self.data = Dbase().Dsqldata(self.MainSQL)

            if SQL_step == 0:
                self.rec_num = len(self.data)-1
            self.FillScreen()
            self.recnum3.SetLabel('/ '+str(len(self.data)))

    def AddRec(self, SQL_step):
        realnames = []
        ValueList = []

        New_ID = cursr.execute(
            'SELECT MAX(ReportID) FROM NonConformance').fetchone()
        if New_ID[0] is None:
            Max_ID = '1'
        else:
            Max_ID = str(New_ID[0]+1)
        for item in Dbase().Dcolinfo(self.tblname):
            realnames.append(item[1])

        ValueList.append(Max_ID)
        ValueList.append(self.TagIDtxt.GetValue())
        st = self.Descrpttxt.GetValue()
        st = st.replace("'", "").replace('"', '')
        ValueList.append(st)
        ValueList.append(self.text1.GetValue())
        ValueList.append(self.text2.GetValue())
        ValueList.append(self.text3.GetValue())
        ValueList.append(self.text4.GetValue())
        ValueList.append(self.text5.GetValue())
        ValueList.append(self.text6.GetValue())
        st = self.Doctxt.GetValue()
        st = st.replace("'", "").replace('"', '')
        ValueList.append(st)
        st = self.MSRtxt.GetValue()
        st = st.replace("'", "").replace('"', '')
        ValueList.append(st)
        ValueList.append(self.text13.GetValue())
        ValueList.append(self.text14.GetValue())
        ValueList.append(self.text15.GetValue())
        ValueList.append(self.text16.GetValue())
        ValueList.append(self.Reprtxt.GetValue())
        ValueList.append(self.text7.GetValue())
        ValueList.append(self.text8.GetValue())
        ValueList.append(self.text9.GetValue())
        ValueList.append(self.text10.GetValue())
        ValueList.append(self.text11.GetValue())
        ValueList.append(self.text12.GetValue())

        if SQL_step == 0:  # enter new record
            CurrentID = Max_ID
            UpQuery = ('INSERT INTO ' + self.tblname + ' VALUES (' + "'"
                       + "','".join(map(str, ValueList)) + "'" + ')')
            Dbase().TblEdit(UpQuery)

        elif SQL_step == 1:  # update edited record
            CurrentID = self.data[self.rec_num][0]
            realnames.remove('ReportID')
            del ValueList[0]

            SQL_str = dict(zip(realnames, ValueList))
            Update_str = ", ".join(["%s='%s'" % (k, v)
                                    for k, v in SQL_str.items()])
            UpQuery = ('UPDATE ' + self.tblname + ' SET ' + Update_str +
                       ' WHERE ReportID = ' + str(CurrentID))
            Dbase().TblEdit(UpQuery)

        elif SQL_step == 3:
            return

        self.b2.Disable()

        return CurrentID

    def OnSelect2(self, evt):
        self.b2.Enable()

    def OnDelete(self, evt):
        recrd = self.data[self.rec_num][0]

        try:
            Dbase().TblDelete(self.tblname, recrd, 'ReportID')
            self.MainSQL = 'SELECT * FROM ' + self.tblname
            self.data = Dbase().Dsqldata(self.MainSQL)
            self.rec_num -= 1
            if self.rec_num < 0:
                self.rec_num = len(self.data)-1
            self.FillScreen()
            self.recnum3.SetLabel('/ '+str(len(self.data)))
        except sqlite3.IntegrityError:
            wx.MessageBox("This Record is associated"
                          " with\nother tables and cannot be deleted!",
                          "Cannot Delete",
                          wx.OK | wx.ICON_INFORMATION)

    def OnSearch(self, evt):
        sqlstr = 'SELECT * FROM MtrSubRcrd'
        frst = 0

        if self.TagIDtxt.GetValue() != '':
            sqlstr = (sqlstr + " WHERE TagID LIKE '%" +
                      self.TagIDtxt.GetValue() + "%' COLLATE NOCASE")
            frst = 1

        if self.text1.GetValue() != '':
            if frst:
                str_and = " OR "
            else:
                str_and = " WHERE "
            sqlstr = (sqlstr + str_and + "Job_Number LIKE '%" +
                      self.text1.GetValue() + "%' COLLATE NOCASE")
            frst = 1

        if self.text2.GetValue() != '':
            if frst:
                str_and = " OR "
            else:
                str_and = " WHERE "
            sqlstr = (sqlstr + str_and + "Requestor LIKE '%" +
                      self.text2.GetValue() + "%' COLLATE NOCASE")
            frst = 1

        if self.text3.GetValue() != '':
            if frst:
                str_and = " OR "
            else:
                str_and = " WHERE "
            sqlstr = (sqlstr + str_and + "Fabricator LIKE '%" +
                      self.text3.GetValue() + "%' COLLATE NOCASE")
            frst = 1

        if self.text4.GetValue() != '':
            if frst:
                str_and = " OR "
            else:
                str_and = " WHERE "
            sqlstr = (sqlstr + str_and + "IssueDate LIKE '%" +
                      self.text4.GetValue() + "%' COLLATE NOCASE")
            frst = 1

        if self.text5.GetValue() != '':
            if frst:
                str_and = " OR "
            else:
                str_and = " WHERE "
            sqlstr = (sqlstr + str_and + "Pipe_Spec_Code LIKE '%" +
                      self.text5.GetValue() + "%' COLLATE NOCASE")
            frst = 1

        if self.text6.GetValue() != '':
            if frst:
                str_and = " OR "
            else:
                str_and = " WHERE "
            sqlstr = (sqlstr + str_and + "Line_Number LIKE '%" +
                      self.text6.GetValue() + "%' COLLATE NOCASE")

        self.data = Dbase().Dsqldata(sqlstr)
        self.rec_num = 0
        self.recnum3.SetLabel('/ '+str(len(self.data)))
        self.FillScreen()

    def OnClear(self, evt):
        self.RcdIDtxt.ChangeValue('')
        self.TagIDtxt.ChangeValue('')
        self.text1.ChangeValue('')
        self.text2.ChangeValue('')
        self.text3.ChangeValue('')
        self.Descrpttxt.ChangeValue('')
        self.text4.ChangeValue('')
        self.text5.ChangeValue('')
        self.text6.ChangeValue('')
        self.text7.ChangeValue('')
        self.text8.ChangeValue('')
        self.text9.ChangeValue('')
        self.Reprtxt.ChangeValue('')
        self.text10.ChangeValue('')
        self.text11.ChangeValue('')
        self.text12.ChangeValue('')
        self.text13.ChangeValue('')
        self.text14.ChangeValue('')
        self.text15.ChangeValue('')
        self.text16.ChangeValue('')
        self.Doctxt.ChangeValue('')
        self.MSRtxt.ChangeValue('')

    def OnRestore(self, evt):
        self.data = Dbase().Dsqldata(self.MainSQL)
        self.recnum3.SetLabel('/ '+str(len(self.data)))
        self.FillScreen()

    def OnClose(self, evt):
        self.GetParent().Enable(True)   # add for child form
        self.__eventLoop.Exit()        # add for child form
        self.Destroy()


class BldTrvlSht(wx.Frame):
    '''Routine to build form and populate grid'''
    def __init__(self, parent, model=None):

        self.parent = parent
        self.model = model
        self.Lvl2tbl = 'InspectionTravelSheet'
        self.NoteStr = []
        self.ComCode = ''
        self.PipeMtrSpec = ''

        if self.Lvl2tbl.find("_") != -1:
            frmtitle = (self.Lvl2tbl.replace("_", " "))
        else:
            frmtitle = (' '.join(re.findall('([A-Z][a-z]*)', self.Lvl2tbl)))

        super(BldTrvlSht, self).__init__(parent,
                                         title=frmtitle,
                                         size=(1200, 875))

        self.Bind(wx.EVT_CLOSE, self.OnClose)

        self.InitUI()

    def InitUI(self):
        model1 = None

        font1 = wx.Font(12, wx.DEFAULT, wx.NORMAL, wx.FONTWEIGHT_BOLD)
        link_fld = 'Timing'
        self.frg_tbl = 'TrvlShtTime'
        self.frg_fld = 'TimeID'
        frgn_col = 'Timing'
        self.columnames = ['ID', 'Timing', 'Note']

        realnames = []
        datatable_str = ''

        # here we get the information needed in the report
        # and for the SQL from the Lvl2 table
        # and determine report column width based either
        # on data type or specified field size
        n = 0
        ColWdth = []
        for item in Dbase().Dcolinfo(self.Lvl2tbl):
            col_wdth = ''
        # check to see if field length is specified if
        # so use it to set grid col width
            for s in re.findall(r'\d+', item[2]):
                if s.isdigit():
                    col_wdth = int(s)
                    ColWdth.append(col_wdth)
            realnames.append(item[1])
            if item[5] == 1:
                self.pk_Name = item[1]
                self.pk_col = n
                if 'INTEGER' in item[2]:
                    self.autoincrement = True
                    if col_wdth == '':
                        ColWdth.append(6)
                # include the primary key and table name into SELECT statement
                datatable_str = (datatable_str + self.Lvl2tbl + '.'
                                 + self.pk_Name + ',')
            # need to make frgn_fld column noneditable in DVC
            elif 'INTEGER' in item[2] or 'FLOAT' in item[2]:
                if col_wdth == '' and 'FLOAT' in item[2]:
                    ColWdth.append(10)
                elif col_wdth == '':
                    ColWdth.append(6)
            elif 'BLOB' in item[2]:
                if col_wdth == '':
                    ColWdth.append(30)
            elif 'TEXT' in item[2] or 'BOOLEAN' in item[2]:
                if col_wdth == '':
                    ColWdth.append(10)
            elif 'DATE' in item[2]:
                if col_wdth == '':
                    ColWdth.append(10)

            # get first Lvl2 datatable column name in item[1]
            # check to see if name is lvl2 primary key or lvl1 linked field
            # if they are not then add tablename and
            # datafield to SELECT statement
            if item[1] != link_fld and item[1] != self.pk_Name:
                datatable_str = (datatable_str + ' ' + self.Lvl2tbl +
                                 '.' + item[1] + ',')
            elif item[1] == link_fld:
                datatable_str = (datatable_str + ' ' + self.frg_tbl +
                                 '.' + frgn_col + ',')

            n += 1

        self.realnames = realnames
        self.ColWdth = ColWdth

        datatable_str = datatable_str[:-1]

        DsqlLvl2 = ('SELECT ' + datatable_str + ' FROM ' + self.Lvl2tbl
                    + ' INNER JOIN ' + self.frg_tbl)
        DsqlLvl2 = (DsqlLvl2 + ' ON ' + self.Lvl2tbl + '.' + link_fld
                    + ' = ' + self.frg_tbl + '.' + self.frg_fld)

        # specify data for upper dvc with all the notes in databasse
        self.DsqlLvl2 = DsqlLvl2

        self.data = Dbase().Dsqldata(self.DsqlLvl2)

        self.data1 = []

        # Create the dataview for the commocity property notes (lower table)
        self.dvc1 = dv.DataViewCtrl(self, wx.ID_ANY, wx.DefaultPosition,
                                    wx.Size(500, 150),
                                    style=wx.BORDER_THEME
                                    | dv.DV_ROW_LINES
                                    | dv.DV_VERT_RULES
                                    | dv.DV_HORIZ_RULES
                                    | dv.DV_MULTIPLE
                                    )

        self.dvc1.SetMinSize = (wx.Size(100, 200))
        self.dvc1.SetMaxSize = (wx.Size(500, 400))

        if model1 is None:
            self.model1 = DataMods(self.Lvl2tbl, self.data1)
        else:
            self.model1 = model1
        self.dvc1.AssociateModel(self.model1)

        # specify which listbox column to display in the combobox
        self.showcol = int

        # Create a dataview control for all the database notes (upper table)
        self.dvc = dv.DataViewCtrl(self, wx.ID_ANY, wx.DefaultPosition,
                                   wx.Size(500, 300),
                                   style=wx.BORDER_THEME
                                   | dv.DV_ROW_LINES
                                   | dv.DV_VERT_RULES
                                   | dv.DV_HORIZ_RULES
                                   | dv.DV_MULTIPLE
                                   )

        self.dvc.SetMinSize = (wx.Size(100, 200))
        self.dvc.SetMaxSize = (wx.Size(500, 400))

    # if autoincrement is false then the data can be sorted based on ID_col
        if self.autoincrement == 0:
            self.data.sort(key=lambda tup: tup[self.pk_col])

    # use the sorted data to load the dataviewlistcontrol
        if self.model is None:
            self.model = DataMods(self.Lvl2tbl, self.data)

        self.dvc.AssociateModel(self.model)

        n = 0
        for colname in self.columnames:
            col_mode = dv.DATAVIEW_CELL_INERT
            self.dvc.AppendTextColumn(colname, n,
                                      width=wx.LIST_AUTOSIZE_USEHEADER,
                                      mode=col_mode)

            self.dvc1.AppendTextColumn(colname, n,
                                       width=wx.LIST_AUTOSIZE_USEHEADER,
                                       mode=col_mode)
            n += 1

        # make columns not sortable and but reorderable.
        n = 0
        for c in self.dvc.Columns:
            c.Sortable = False
            # make the category column sortable
            if n == 1:
                c.Sortable = True
            c.Reorderable = True
            c.Resizeable = True
            n += 1

        # change to not let the ID col be moved.
        self.dvc.Columns[(self.pk_col)].Reorderable = False
        self.dvc.Columns[(self.pk_col)].Resizeable = False

        # set the Sizer property (same as SetSizer)
        self.Sizer = wx.BoxSizer(wx.VERTICAL)

        # develope the comboctrl and attach popup list
        self.cmb1 = wx.ComboCtrl(self, pos=(10, 10), size=(200, -1))
        self.Bind(wx.EVT_TEXT, self.OnSelect, self.cmb1)
        self.cmb1.SetHint(frgn_col)
        self.showcol = 1
        self.popup = ListCtrlComboPopup(self.frg_tbl, showcol=self.showcol)

        self.cmbsizer = wx.BoxSizer(wx.HORIZONTAL)

        # add a button to call main form to search combo list data
        self.b6 = wx.Button(self, label="Restore Data")
        self.Bind(wx.EVT_BUTTON, self.OnRestore, self.b6)
        self.cmbsizer.Add(self.b6, 0)

        self.cmb1.SetPopupControl(self.popup)
        self.cmbsizer.Add(self.cmb1, 0, wx.ALIGN_TOP)

        # add a button to call main form to search combo list data
        self.b5 = wx.Button(self, label="<= Search Data")
        self.Bind(wx.EVT_BUTTON, self.OnSearch, self.b5)
        self.cmbsizer.Add(self.b5, 0, wx.BOTTOM, 15)

        self.b9 = wx.Button(self, label='Add Note(s) to\nTravel Sheet Report')
        self.Bind(wx.EVT_BUTTON, self.OnSelectDone, self.b9)
        self.cmbsizer.Add(20, 10)
        self.cmbsizer.Add(self.b9, 0, wx.BOTTOM, 15)

        self.addlbl = wx.StaticText(self, -1, style=wx.ALIGN_CENTER_HORIZONTAL)
        txt = '   Complete Listing of All Inspection Travel Sheet Notes'
        self.addlbl.SetLabel(txt)
        self.addlbl.SetForegroundColour((255, 0, 0))
        self.addlbl.SetFont(font1)
        self.Sizer.Add(self.addlbl, 0, wx.ALIGN_LEFT)
        self.Sizer.Add((10, 20))
        self.Sizer.Add(self.cmbsizer, 0, wx.ALIGN_CENTER)
        self.Sizer.Add(self.dvc, 1, wx.EXPAND)

        self.cmbsizer2 = wx.BoxSizer(wx.HORIZONTAL)
        # add a button to call main form to search combo list data
        self.b10 = wx.Button(self, label="Restore Data")
        self.Bind(wx.EVT_BUTTON, self.OnRstrDVC1, self.b10)
        self.cmbsizer2.Add(self.b10, 0, wx.ALIGN_LEFT, 5)

        # develope the comboctrl and attach popup list
        self.cmb10 = wx.ComboCtrl(self, pos=(10, 10), size=(200, -1))
        self.Bind(wx.EVT_TEXT, self.OnSelect, self.cmb10)
        self.cmb10.SetHint('Note Category')
        self.showcol = 1
        self.popup = ListCtrlComboPopup(self.frg_tbl, showcol=self.showcol)
        self.cmb10.SetPopupControl(self.popup)
        self.cmbsizer2.Add(self.cmb10, 0, wx.ALIGN_TOP, 5)

        # add a button to call main form to search combo list data
        self.b11 = wx.Button(self, label="<= Search Data")
        self.Bind(wx.EVT_BUTTON, self.OnSrchDVC1, self.b11)
        self.cmbsizer2.Add(self.b11, 0, 5)

        self.b8 = wx.Button(self, id=2, label="Print Inspection\nTravel Sheet")
        self.Bind(wx.EVT_BUTTON, self.PrintFile, self.b8)
        self.cmbsizer2.Add(self.b8, 0, 5)

        self.b12 = wx.Button(self,
                             label='Remove Item(s) From\nTravel Sheet List')
        self.Bind(wx.EVT_BUTTON, self.OnRmvNote, self.b12)
        self.cmbsizer2.Add((30, 10))
        self.cmbsizer2.Add(self.b12, 0, 5)

        self.b4 = wx.Button(self, label="Exit")
        self.b4.SetForegroundColour('red')
        self.Bind(wx.EVT_BUTTON, self.OnClose, self.b4)
        self.cmbsizer2.Add((60, 10))
        self.cmbsizer2.Add(self.b4, 0, wx.ALL | wx.ALIGN_CENTER, 5)

        self.Sizer.Add((10, 20))
        self.Sizer.Add(self.dvc1, 1, wx.EXPAND)
        self.Sizer.Add((10, 15))
        self.Sizer.Add(self.cmbsizer2, 0, wx.ALIGN_CENTER)
        self.Sizer.Add((10, 20))

        self.CenterOnParent()
        self.GetParent().Enable(False)
        self.Show(True)
        self.__eventLoop = wx.GUIEventLoop()
        self.__eventLoop.Run()

    def OnClose(self, evt):
        self.GetParent().Enable(True)
        self.__eventLoop.Exit()
        self.Destroy()

    def PrintFile(self, evt):
        import Report_Lvl2

        colwdths = [5, 20, 70, 5, 5]
        colnames = ['QC', 'Timing', 'Note', 'Hold', 'Pass']
        ttl = 'Inspection Travel Sheet'
        rptdata = []

        # confirm there is data to print
        if self.data1 == []:
            NoData = wx.MessageDialog(
                None, 'No Items Selected to Print', 'Error', wx.OK |
                wx.ICON_EXCLAMATION)
            NoData.ShowModal()
            return
        else:
            for item in self.data1:
                item = list(item)
                item[0] = ''
                item.append('')
                item.append('')
                rptdata.append(item)

        filename = self.ReportName()

        Report_Lvl2.Report(self.Lvl2tbl, rptdata, colnames,
                           colwdths, filename, ttl).create_pdf()

    def ReportName(self):

        saveDialog = wx.FileDialog(self, message='Save Report as PDF.',
                                   wildcard='PDF (*.pdf)|*.pdf',
                                   style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT)
        if saveDialog.ShowModal() == wx.ID_CANCEL:
            filename = ''
        filename = saveDialog.GetPath()
        if filename.find(".pdf") == -1:
            filename = filename + '.pdf'
        saveDialog.Destroy()

        if not filename:
            exit()

        return filename

    def OnSelectDone(self, evt):
        # add note(s) selected from upper table to lower table
        # rowID is table index value
        NoteIds = []
        items = self.dvc.GetSelections()
        for item in items:
            row = self.model.GetRow(item)
            rowID = self.model.GetValueByRow(row, self.pk_col)
            if int(rowID) in [self.data1[i][0]
                              for i in range(len(self.data1))]:
                continue
            NoteIds.append(rowID)
        if self.NoteStr == []:
            self.NoteStr = str(tuple(NoteIds))
        else:
            self.NoteStr = self.NoteStr[:-1] + ', ' + str(tuple(NoteIds))[1:]

        self.NoteStr = "".join(self.NoteStr.split())
        if len(NoteIds) <= 1:
            self.NoteStr = self.NoteStr[:-2] + ')'

        self.Dsql1 = (self.DsqlLvl2 + ' WHERE ' + self.pk_Name +
                      ' IN ' + self.NoteStr)
        self.data1 = Dbase().Dsqldata(self.Dsql1)
        self.model1 = DataMods(self.Lvl2tbl, self.data1)
        self.dvc1.AssociateModel(self.model1)
        self.dvc1.Refresh

    def OnRmvNote(self, evt):
        # remove selected note(s) from Commodity property
        # rowID is table index value
        NoteIds = []
        items = self.dvc1.GetSelections()
        if list(items) != []:
            for item in items:
                row = self.model1.GetRow(item)
                rowID = self.model1.GetValueByRow(row, self.pk_col)
                NoteIds.append(rowID)

            NoteLst = self.ConvertStr_Lst(self.NoteStr)

            for i in NoteIds:
                NoteLst.remove(i)

            self.NoteStr = str(tuple(NoteLst))
            if self.NoteStr[-2] == ',':
                self.NoteStr = self.NoteStr[:-2] + ')'
            elif self.NoteStr == '()':
                self.NoteStr = None

            if self.NoteStr is not None:
                self.Dsql1 = (self.DsqlLvl2 + ' WHERE ' + self.pk_Name +
                              ' IN ' + self.NoteStr)
                self.data1 = Dbase().Dsqldata(self.Dsql1)
            else:
                self.data1 = []
            self.model1 = DataMods(self.Lvl2tbl, self.data1)
            self.dvc1.AssociateModel(self.model1)
            self.dvc1.Refresh

    def ConvertStr_Lst(self, vals):
        lst = vals.split("'")
        newlst = []
        for i in range(len(lst)):
            if lst[i].isdigit():
                newlst.append(lst[i])
        return newlst

    def OnSrchDVC1(self, evt):
        # collect feign table info
        frgn_info = Dbase().Dtbldata(self.Lvl2tbl)
        field = frgn_info[0][4]
        frg_tbl1 = frgn_info[0][2]

        # do search of string value from combobox
        # equal to value in the self.frg_tbl
        Shcol = Dbase().Dcolinfo(frg_tbl1)[self.showcol][1]
        ShQry = ("SELECT " + field + " FROM " + frg_tbl1 + " WHERE " +
                 Shcol + " LIKE '%" + self.cmb10.GetValue() +
                 "%' COLLATE NOCASE")
        ShQryVal = str(Dbase().Dsqldata(ShQry)[0][0])
        # append the found frgn_fld to the original data grid SQL and
        # find only records equal to the combo selection
        ShQry = (self.Dsql1 + ' AND ' + frg_tbl1 + '.' +
                 field + ' = "' + ShQryVal + '"')
        OSdata = Dbase().Search(ShQry)
        # if nothing is found show blank grid
        if OSdata is False:
            OSdata = []
        self.model1 = DataMods(self.Lvl2tbl, OSdata)
        self.dvc1.AssociateModel(self.model1)
        self.dvc1.Refresh

    def OnRstrDVC1(self, evt):
        ORdata1 = Dbase().Restore(self.Dsql1)
        self.cmb10.ChangeValue('')
        self.cmb10.SetHint(self.frg_fld)
        self.model1 = DataMods(self.Lvl2tbl, ORdata1)
        self.dvc1._AssociateModel(self.model1)
        self.dvc1.Refresh

    def OnSearch(self, evt):
        # collect feign table info
        frgn_info = Dbase().Dtbldata(self.Lvl2tbl)
        field = frgn_info[0][4]
        self.frg_tbl = frgn_info[0][2]

        # do search of string value from combobox equal
        # to value in the self.frg_tbl
        Shcol = Dbase().Dcolinfo(self.frg_tbl)[self.showcol][1]
        ShQuery = ('SELECT ' + field + ' FROM ' + self.frg_tbl +
                   " WHERE " + Shcol + " LIKE '%" + self.cmb1.GetValue()
                   + "%' COLLATE NOCASE")
        ShQueryVal = str(Dbase().Dsqldata(ShQuery)[0][0])
        # append the found frgn_fld to the original data grid SQL and
        # find only records equal to the combo selection
        ShQuery = (self.DsqlLvl2 + ' WHERE ' + self.frg_tbl + '.' +
                   field + ' = "' + ShQueryVal + '"')

        OSdata = Dbase().Search(ShQuery)
        # if nothing is found show blank grid
        if OSdata is False:
            OSdata = []
        self.model = DataMods(self.Lvl2tbl, OSdata)
        self.dvc.AssociateModel(self.model)
        self.dvc.Refresh

    def OnSelect(self, evt):
        txt = ('''To complete adding a new record click "Add Row".\nThen
                edit data by double click on the cell.''')
        self.editlbl.SetLabel(txt)
        self.editlbl.SetForegroundColour((255, 0, 0))
        self.Sizer.Layout()

    def OnRestore(self, evt):
        self.ORdata = Dbase().Restore(self.DsqlLvl2)
        self.cmb1.ChangeValue('')
        self.cmb1.SetHint(self.frg_fld)
        self.model = DataMods(self.Lvl2tbl, self.ORdata)
        self.dvc._AssociateModel(self.model)
        self.dvc.Refresh


class BldConduit(wx.Frame):
    '''Routine to build form and populate grid'''
    def __init__(self, parent, tblname, ComdPrtyID=None):

        self.parent = parent   # add line for child
        self.tblname = tblname

        self.ComCode = ''
        self.PipeMtrSpec = ''
        self.ComdPrtyID = ComdPrtyID
        self.columnames = []
        self.rec_num = 0
        self.addbtns = []
        self.lctrls = []
        self.data = []
        self.MainSQL = ''

        if self.tblname.find("_") != -1:
            self.frmtitle = (self.tblname.replace("_", " "))
        else:
            self.frmtitle = (' '.join(re.findall('([A-Z][a-z]*)',
                             self.tblname)))

        if self.tblname == 'Tubing':
            frmsz = (800, 640)
        elif self.tblname == 'Piping':
            frmsz = (1000, 600)

        super(BldConduit, self).__init__(parent,
                                         title=self.frmtitle,
                                         size=frmsz)

        self.Bind(wx.EVT_CLOSE, self.OnClose)

        self.InitUI()

    def InitUI(self):
        StartQry = None

        if self.ComdPrtyID is not None:
            query = ('''SELECT Commodity_Code,Pipe_Material_Code,End_Connection,
                      Pipe_Code FROM CommodityProperties WHERE
                      CommodityPropertyID = '''
                     + str(self.ComdPrtyID))
            dataset = Dbase().Dsqldata(query)[0]
            self.PipeMtrSpec = dataset[1]
            self.ComCode = dataset[0]
            self.PipeCode = dataset[3]

        titles = []
        tles = []

        # set values for either tubing or piping table
        if self.tblname == 'Tubing':
            self.TopRows = 3
            self.BtmRows = 2
            # list order is: hint,table,box width,combo or
            # text box (0 or 1),tbldID,column name shown in cmbbox
            self.top_hints_tbls = [
                ('Tube OD', 'TubeSize', 100, 1, 'SizeID', 'OD'),
                ('Wall Thk', 'TubeWall', 100, 1, 'WallID', 'Wall'),
                ('Tube Material', 'TubeMaterial', 350, 1, 'ID', 'Material')]

            top_space = (30, 100, 140)

            self.btm_hints_tbls = [
                ('Tube Valve\nMaterial', 'TubeValveMatr', 150, 1, 'TVID',
                 'Tube_Valve_Body_Material'),
                ('Valve Type', '', 100, 0), ('Manufacturer', '', 100, 0),
                ('Model Number', '', 200, 0)]

            btm_space = (5, 30, 55, 80)

            # name of main tbl ID field
            self.field = 'TubeID'
            # name of filed linked in the PipeSpecification table
            self.field2 = 'Tube_ID'

        elif self.tblname == 'Piping':
            self.TopRows = 4
            self.BtmRows = 2
            self.top_hints_tbls = [
                ('Pipe OD', 'Pipe_OD', 100, 1, 'PipeOD_ID', 'Pipe_OD'),
                ('Pipe OD', 'Pipe_OD', 100, 1, 'PipeOD_ID', 'Pipe_OD'),
                ('Pipe Material', 'PipeMaterial', 300, 1, 'ID',
                 'Pipe_Material'),
                ('Schedule', 'PipeSchedule', 120, 1, 'ID', 'Pipe_Schedule')]

            top_space = (35, 70, 165, 0)

            self.btm_hints_tbls = [
                ('Pipe OD', 'Pipe_OD', 100, 1, 'PipeOD_ID', 'Pipe_OD'),
                ('Pipe OD', 'Pipe_OD', 100, 1, 'PipeOD_ID', 'Pipe_OD'),
                ('Nipple Material', 'PipeMaterial', 300, 1, 'ID',
                 'Pipe_Material'),
                ('Schedule', 'PipeSchedule', 120, 1, 'ID', 'Pipe_Schedule'),
                ('End Connection', 'EndConnects', 200, 1, 'EndID',
                 'Connection')]

            btm_space = (20, 25, 85, 60, 60)

            self.field = 'PipingID'
            self.field2 = 'ComponentPack_ID'

        # setup label field if commodity does not have item specified
        font2 = wx.Font(12, wx.DEFAULT, wx.NORMAL, wx.FONTWEIGHT_BOLD)
        txt_nodata = (' The ' + self.frmtitle +
                      ' have not been setup for this Commodity Property')
        self.lbl_nodata = wx.StaticText(self, -1, label=txt_nodata,
                                        size=(600, 40), style=wx.LEFT)
        self.lbl_nodata.SetForegroundColour((255, 0, 0))
        self.lbl_nodata.SetFont(font2)
        self.lbl_nodata.SetLabel('   ')

        # select data based on if the form is called from commodity
        # properties form or if you are to see all data
        if self.ComdPrtyID is not None:
            query = ('SELECT ' + self.field2 +
                     ' FROM PipeSpecification WHERE Commodity_Property_ID = '
                     + str(self.ComdPrtyID))
            chk = Dbase().Dsqldata(query)
            if chk != []:
                StartQry = chk[0][0]
                if StartQry is not None:
                    self.MainSQL = ('SELECT * FROM ' + self.tblname + ' WHERE '
                                    + self.field + ' = ' + str(StartQry))
                    self.data = Dbase().Dsqldata(self.MainSQL)
                else:   # no item setup for commodity property
                    self.lbl_nodata.SetLabel(txt_nodata)
        else:    # no commodity property specified
            self.MainSQL = 'SELECT * FROM ' + self.tblname
            self.data = Dbase().Dsqldata(self.MainSQL)
        # specify which listbox column to display in the combobox
        self.showcol = int

        # set the Sizer property (same as SetSizer)
        self.Sizer = wx.BoxSizer(wx.VERTICAL)

        self.specsizer = wx.BoxSizer(wx.HORIZONTAL)
        font = wx.Font(10, wx.DEFAULT, wx.NORMAL, wx.NORMAL)
        # set up bxw to show that there is no item for commodity property
        self.warningsizer = wx.BoxSizer(wx.HORIZONTAL)
        self.warningsizer.Add(self.lbl_nodata, 0, wx.ALIGN_CENTER)

        if self.tblname == 'Piping':
            # build the pipe material spec combobox
            # depending on if the value is preset
            self.cmbMtr = wx.ComboCtrl(self, pos=(10, 10),
                                       size=(80, -1), style=wx.CB_READONLY)
            # need to chage the pipespec index for
            # material to the material code
            if self.PipeMtrSpec != '':
                query = ('''SELECT * FROM PipeMaterialSpec WHERE
                          Material_Spec_ID = ''' + str(self.PipeMtrSpec))
                MtrSpc = Dbase().Dsqldata(query)[0][1]
                self.cmbMtr.SetPopupControl(ListCtrlComboPopup(
                    'PipeMaterialSpec', query, showcol=1))
                self.cmbMtr.SetValue(MtrSpc)
            # if there is no commodity property specified
            # show all material codes in combo
            else:
                self.cmbMtr.SetPopupControl(ListCtrlComboPopup(
                    'PipeMaterialSpec', showcol=1))

            self.Bind(wx.EVT_COMBOBOX_CLOSEUP, self.OnCmbMtrClose, self.cmbMtr)
            self.Bind(wx.EVT_COMBOBOX_DROPDOWN, self.OnCmbMtrOpen, self.cmbMtr)

            if self.ComdPrtyID is None:
                self.specsizer.Add((20, 20))
            else:      # show commodity property ticker
                if self.PipeCode is None:
                    lbl = self.ComCode + ' -'
                    sz = (60, -1)
                else:
                    lbl = self.PipeCode
                    sz = (90, -1)

            self.text1 = wx.TextCtrl(self, size=sz, value=lbl,
                                     style=wx.TE_READONLY | wx.TE_RIGHT)
            self.text1.SetForegroundColour((255, 0, 0))
            self.text1.SetFont(font)
            self.specsizer.Add(self.text1, 0, wx.LEFT | wx.BOTTOM, 10)
            self.specsizer.Add(self.cmbMtr, 0, wx.RIGHT | wx.BOTTOM, 10)

        # setup upper section for comboboxs and labels
        self.titlesizer = wx.BoxSizer(wx.HORIZONTAL)

        font1 = wx.Font(12, wx.DEFAULT, wx.NORMAL, wx.FONTWEIGHT_BOLD)
        n = 0

        for i in self.top_hints_tbls:
            title = wx.StaticText(self, label=self.top_hints_tbls[n][0],
                                  style=wx.ALIGN_LEFT)
            title.SetForegroundColour((255, 0, 0))
            titles.append(title)
            self.titlesizer.Add(titles[n], 0, wx.RIGHT |
                                wx.ALIGN_CENTER, border=5)

            if i[3] == 1:
                addbtn = wx.Button(self, label='+', size=(35, -1))
                addbtn.SetForegroundColour((255, 0, 0))
                addbtn.SetFont(font1)
                self.Bind(wx.EVT_BUTTON, self.OnAdd1, addbtn)
                self.addbtns.append(addbtn)
                self.titlesizer.Add(addbtn, 0, wx.RIGHT, border=top_space[n])

            n += 1

        self.titlesizer.Add((30, 10))

        self.cmbs = []
        self.cmbsizers = []
        self.txtbxs = []
        self.btmcmbsizers = []
        self.btmcmbs = []

        # draw a line between upper and lower section
        self.ln = wx.StaticLine(self, 0, size=(800, 5), style=wx.LI_VERTICAL)
        self.ln.SetBackgroundColour('Black')

        # setup lower section of comboboxs and labels
        self.tlesizer = wx.BoxSizer(wx.HORIZONTAL)

        n = 0
        for i in self.btm_hints_tbls:
            tle = wx.StaticText(self, label=self.btm_hints_tbls[n][0],
                                style=wx.ALIGN_LEFT)
            tle.SetForegroundColour((255, 0, 0))
            tles.append(tle)
            self.tlesizer.Add(tles[n], 0, wx.LEFT | wx.ALIGN_CENTER,
                              border=btm_space[n])

            if i[3] == 1:
                addbtn = wx.Button(self, label='+', size=(35, -1))
                addbtn.SetForegroundColour((255, 0, 0))
                addbtn.SetFont(font1)
                self.Bind(wx.EVT_BUTTON, self.OnAdd1, addbtn)
                self.addbtns.append(addbtn)
                self.tlesizer.Add(addbtn, 0, wx.LEFT, border=5)

            n += 1

        if self.tblname == 'Tubing':
            self.textsizer = wx.BoxSizer(wx.HORIZONTAL)

            self.text6 = wx.TextCtrl(self, size=(650, 80), value='',
                                     style=wx.TE_MULTILINE)

            self.Bind(wx.EVT_TEXT, self.OnSelect, self.text6)

            self.note7 = wx.StaticText(self, label='Notes',
                                       style=wx.ALIGN_LEFT)
            self.note7.SetForegroundColour((255, 0, 0))

            self.textsizer.Add(self.note7, 0, wx.ALIGN_CENTER_VERTICAL |
                               wx.RIGHT, 10)
            self.textsizer.Add(self.text6, 0, wx.RIGHT, 10)

        self.Sizer.Add(self.warningsizer, 0, wx.CENTER | wx.TOP, 5)
        self.Sizer.Add(self.specsizer, 0, wx.ALL | wx.ALIGN_CENTER)
        self.Sizer.Add(self.titlesizer, 0, wx.ALL | wx.ALIGN_CENTER)

        # build the comboboxs for the upper section
        self.BxBld(self.TopRows, self.top_hints_tbls, 'Top')
        for n in range(self.TopRows):
            self.Sizer.Add(self.cmbsizers[n], 0, wx.ALL | wx.ALIGN_CENTER)

        self.Sizer.Add((20, 15))
        self.Sizer.Add(self.ln, 0, wx.ALIGN_CENTER)
        self.Sizer.Add((20, 15))
        self.Sizer.Add(self.tlesizer, 0, wx.ALL | wx.ALIGN_CENTER)

        # build the comboboxs for the lower section
        self.BxBld(self.BtmRows, self.btm_hints_tbls, 'Btm')
        for n in range(self.BtmRows):
            self.Sizer.Add(self.btmcmbsizers[n], 0, wx.ALL | wx.ALIGN_CENTER)

        self.btnDict = {self.addbtns[i]: i for i in range(len(self.addbtns))}

        if self.tblname == 'Tubing':    # add textbox for Notes
            self.Sizer.Add(self.textsizer, 0, wx.ALIGN_CENTER | wx.TOP, 15)

        # Add buttons for grid modifications
        self.b2 = wx.Button(self, label="Add/Update\nto " + self.frmtitle)
        self.b2.Disable()
        self.Bind(wx.EVT_BUTTON, self.OnAddRec, self.b2)

        self.b3 = wx.Button(self, label="Delete\nSpecification")
        self.Bind(wx.EVT_BUTTON, self.OnDelete, self.b3)

        self.b7 = wx.Button(self, label="Print\nReport")
        self.Bind(wx.EVT_BUTTON, self.PrintFile, self.b7)

        self.b4 = wx.Button(self, label="Exit")
        self.Bind(wx.EVT_BUTTON, self.OnClose, self.b4)

        self.b5 = wx.Button(self, label="Clear\nBoxes")
        self.Bind(wx.EVT_BUTTON, self.OnRestoreCmbs, self.b5)

        # add a button box and place the buttons
        self.btnbox = wx.BoxSizer(wx.HORIZONTAL)
        self.btnbox.Add(self.b5, 0, wx.ALIGN_CENTER | wx.RIGHT, 35)
        self.btnbox.Add(self.b2, 0, wx.ALL, 5)
        if self.ComdPrtyID is not None:
            self.b6 = wx.Button(self, size=(120, 45),
                                label="Show All\n" + self.frmtitle)
            self.b3.SetLabel('Delete From\nCommodity')
            if StartQry is None:
                self.b3.Disable()
            self.Bind(wx.EVT_BUTTON, self.OnAddComd, self.b6)
            self.btnbox.Add(self.b6, 0, wx.ALL | wx.ALIGN_CENTER, 5)
        self.btnbox.Add(self.b3, 0, wx.ALL, 5)
        self.btnbox.Add((20, 10))
        self.btnbox.Add(self.b7, 0, wx.ALL, 5)
        self.btnbox.Add((20, 10))
        self.btnbox.Add(self.b4, 0, wx.ALL, 5)

        self.fst = wx.Button(self, label='<<')
        self.lst = wx.Button(self, label='>>')
        self.nxt = wx.Button(self, label='>')
        self.pre = wx.Button(self, label='<')
        self.fst.Bind(wx.EVT_BUTTON, self.OnMovefst)
        self.lst.Bind(wx.EVT_BUTTON, self.OnMovelst)
        self.nxt.Bind(wx.EVT_BUTTON, self.OnMovenxt)
        self.pre.Bind(wx.EVT_BUTTON, self.OnMovepre)

        self.navbox = wx.BoxSizer(wx.HORIZONTAL)
        self.navbox.Add(self.fst, 0, wx.ALL, 5)
        self.navbox.Add(self.pre, 0, wx.ALL, 5)
        self.navbox.Add(self.nxt, 0, wx.ALL, 5)
        self.navbox.Add(self.lst, 0, wx.ALL, 5)

        self.numbox = wx.BoxSizer(wx.HORIZONTAL)
        self.recnum1 = wx.StaticText(self, label='Record ',
                                     style=wx.ALIGN_LEFT)
        self.recnum1.SetForegroundColour((255, 0, 0))

        self.recnum2 = wx.StaticText(self, label=str(self.rec_num+1),
                                     style=wx.ALIGN_LEFT)
        self.recnum2.SetForegroundColour((255, 0, 0))
        self.recnum3 = wx.StaticText(self, label='/ '+str(len(self.data)),
                                     style=wx.ALIGN_LEFT)
        self.recnum3.SetForegroundColour((255, 0, 0))
        self.numbox.Add(self.recnum1, 0, wx.ALL, 5)
        self.numbox.Add(self.recnum2, 0, wx.ALL, 5)
        self.numbox.Add(self.recnum3, 0, wx.ALL, 5)

        self.Sizer.Add(self.btnbox, 0, wx.ALIGN_CENTER | wx.TOP, 10)
        self.Sizer.Add(self.navbox, 0, wx.ALIGN_CENTER | wx.TOP, 10)
        self.Sizer.Add(self.numbox, 0, wx.ALIGN_CENTER | wx.TOP)
        self.b4.SetFocus()

        self.FillScreen()

        # add these following lines for child parent
        self.CenterOnParent()
        self.GetParent().Enable(False)
        self.Show(True)
        self.__eventLoop = wx.GUIEventLoop()
        self.__eventLoop.Run()

    def BxBld(self, rows, hnts, lvl):
        # build one row at a time for either level
        for n in range(rows):
            # build the box sizer for the row
            self.cmbsizer = wx.BoxSizer(wx.HORIZONTAL)
            if lvl == 'Top':
                self.cmbsizers.append(self.cmbsizer)
            else:
                self.btmcmbsizers.append(self.cmbsizer)
            # build a list of combo and text boxes for each level
            for item in hnts:
                # item[3] = 1 designates a combobox is needed
                if item[3] == 1:
                    cmb = wx.ComboCtrl(self, pos=(10, 10), size=(item[2], -1),
                                       style=wx.CB_READONLY)
                    self.Bind(wx.EVT_COMBOBOX_DROPDOWN, self.OnCmbOpen, cmb)
                    self.Bind(wx.EVT_COMBOBOX_CLOSEUP, self.OnCmbClose, cmb)
                    cmb.SetHint(item[0])
                    self.showcol = 1
                    if item[0] == 'Schedule':
                        self.showcol = 2
                    # build a dictionary of the combobox list table
                    # and combobox listctrl
                    cmb.SetPopupControl(ListCtrlComboPopup(
                        item[1], showcol=self.showcol, lctrls=self.lctrls))
                    if lvl == 'Top':
                        self.cmbs.append(cmb)
                    else:
                        self.btmcmbs.append(cmb)

                    if lvl == 'Top':
                        self.cmbsizers[n].Add(cmb, 0,)
                        self.cmbsizers[n].Add((25, 10))
                    else:
                        self.btmcmbsizers[n].Add(cmb, 0)
                # if not a combo box then build a textbox
                else:
                    txtbx = wx.TextCtrl(self, size=(item[2], 33), value='',
                                        style=wx.TE_CENTER)
                    self.Bind(wx.EVT_TEXT, self.OnSelect, txtbx)
                    txtbx.SetHint(item[0])
                    self.txtbxs.append(txtbx)

                    if lvl == 'Top':
                        self.cmbsizers[n].Add(txtbx, 0)
                        self.cmbsizers[n].Add((25, 10))
                    else:
                        self.btmcmbsizers[n].Add(txtbx, 0)
                        self.btmcmbsizers[n].Add((25, 10))

    # called to update the item table and commodity table if needed
    def OnAddRec(self, evt):
        # check first that data is all present and clean up incomplete boxes
        check = self.ValData()
        if check:
            SQL_step = 3

            choice1 = ('1) Save this as a new ' + self.frmtitle +
                       ' Specification')
            choice2 = ('2) Update the existing ' + self.frmtitle +
                       ' Specification with this data')
            txt1 = ('''NOTE: Updating this information will be\n
                    reflected in all associated ''' + self.frmtitle)
            txt2 = (''' Specifications!\nRecommendation is to save as a new
                     specification.\n\n\tHow do you want to proceed?''')

            # if this is a not commodity related change
            if self.ComdPrtyID is None:
                if self.data == []:
                    SQL_step = 0
                else:
                    # Make a selection as to whether the record
                    # is to be a new or an update item
                    # use a SingleChioce dialog to determine
                    # if data is new record or edited record
                    SQL_Dialog = wx.SingleChoiceDialog(
                        self, txt1+txt2, 'Information Has Changed',
                        [choice1, choice2], style=wx.CHOICEDLG_STYLE)
                    if SQL_Dialog.ShowModal() == wx.ID_OK:
                        SQL_step = SQL_Dialog.GetSelection()
                    SQL_Dialog.Destroy()

                self.AddRec(SQL_step)
                self.MainSQL = 'SELECT * FROM ' + self.tblname
                self.data = Dbase().Dsqldata(self.MainSQL)

            else:  # this is a commodity related change
                choice1 = choice1 + ' for this commodity?'
                choice2 = choice2 + ' and save for this commodity?'
                # use a SingleChioce dialog to determine if data
                # is new record or edited record
                SQL_Dialog = wx.SingleChoiceDialog(
                    self, txt1+txt2, 'Information Has Changed',
                    [choice1, choice2], style=wx.CHOICEDLG_STYLE)
                if SQL_Dialog.ShowModal() == wx.ID_OK:
                    SQL_step = SQL_Dialog.GetSelection()

                SQL_Dialog.Destroy()

                # always save as a new spec if commodity property is specified
                cmd_addID = self.AddRec(SQL_step)
                # no matter the change over write or add the specification ID
                # to the PipeSpec table under the commodity property ID

                self.ChgSpecID(cmd_addID)
                query = (
                    'SELECT ' + self.field2 +
                    ' FROM PipeSpecification WHERE Commodity_Property_ID = '
                    + str(self.ComdPrtyID))
                StartQry = Dbase().Dsqldata(query)
                self.MainSQL = ('SELECT * FROM ' + self.tblname + ' WHERE '
                                + self.field + ' = ' + str(StartQry[0][0]))
                self.data = Dbase().Dsqldata(self.MainSQL)

            if SQL_step == 0:
                self.rec_num = len(self.data)-1
            self.FillScreen()
            self.recnum3.SetLabel('/ '+str(len(self.data)))

    def AddRec(self, SQL_step):
        realnames = []
        ValueList = []

        colinfo = Dbase().Dcolinfo(self.tblname)
        ValueList = [None for i in range(0, len(colinfo))]
        # if the table index is auto increment then assign
        # next value otherwise do nothing
        for item in colinfo:
            if item[5] == 1:
                IDname = item[1]
                if 'INTEGER' in item[2]:
                    New_ID = cursr.execute(
                        "SELECT MAX(" + IDname + ") FROM " +
                        self.tblname).fetchone()
                    if New_ID[0] is None:
                        Max_ID = '1'
                    else:
                        Max_ID = str(New_ID[0]+1)
            realnames.append(item[1])
        ValueList[0] = str(Max_ID)

        if self.tblname == 'Piping':
            query = ('''SELECT Material_Spec_ID FROM PipeMaterialSpec
                      WHERE Pipe_Material_Spec = "''' + self.cmbMtr.GetValue()
                     + '"')
            ValueList[1] = str(Dbase().Dsqldata(query)[0][0])

        bc = 0

        if self.tblname == 'Piping':
            m = 18  # start of data column for start of bottom combo boxes
            increment = 2
        if self.tblname == 'Tubing':
            m = 10
            increment = 1

        for g in range(0, self.TopRows):
            # determine if this item is a combo box
            if self.top_hints_tbls[g][3] == 1:
                h = [i for i in range(g, len(self.cmbs), self.TopRows)]
                for n in h:
                    cmbtbl = self.top_hints_tbls[g][1]
                    field = self.top_hints_tbls[g][5]
                    tblID = self.top_hints_tbls[g][4]
                    if self.cmbs[n].GetValue():
                        query = ('SELECT ' + tblID + ' FROM ' + cmbtbl +
                                 ' WHERE ' + field + " = '" +
                                 self.cmbs[n].GetValue() + "'")
                        ValueList[n+increment] = str(
                            Dbase().Dsqldata(query)[0][0])

        for g in range(0, len(self.btm_hints_tbls)):
            # determine if this item is a combo box
            if self.btm_hints_tbls[g][3] == 1:
                h = [i for i in range(g, len(self.btmcmbs),
                     len(self.btm_hints_tbls))]
                for n in h:
                    cmbtbl = self.btm_hints_tbls[g][1]
                    field = self.btm_hints_tbls[g][5]
                    tblID = self.btm_hints_tbls[g][4]
                    if self.btmcmbs[n].GetValue():
                        query = ("SELECT " + tblID + " FROM " + cmbtbl +
                                 " WHERE " + field + " = '" +
                                 self.btmcmbs[n].GetValue() + "'")
                        ValueList[m+n] = str(Dbase().Dsqldata(query)[0][0])

        if self.tblname == 'Tubing':
            # force correct data into ValueList from btmcbm box 1
            cmbtbl = self.btm_hints_tbls[0][1]
            field = self.btm_hints_tbls[0][5]
            tblID = self.btm_hints_tbls[0][4]
            if self.btmcmbs[1].GetValue():
                query = ('SELECT ' + tblID + ' FROM ' + cmbtbl + ' WHERE ' +
                         field + " = '" + self.btmcmbs[1].GetValue() + "'")
                ValueList[14] = str(Dbase().Dsqldata(query)[0][0])
            # collect the data from the text boxes and place inot ValueList
            # numbers are data column numbers coresponding to lower textboxes
            for n in (10, 11, 12, 14, 15, 16):
                if self.txtbxs[bc].GetValue():
                    ValueList[n+increment] = self.txtbxs[bc].GetValue()
                bc += 1
            # add the note box to the ValueList
            ValueList[len(colinfo)-increment] = self.text6.GetValue()

        if SQL_step == 0:  # enter new record
            CurrentID = Max_ID
            num_vals = ('?,'*len(colinfo))[:-1]
            UpQuery = ('INSERT INTO ' + self.tblname + ' VALUES (' +
                       num_vals + ')')
            Dbase().TblEdit(UpQuery, ValueList)
            self.rec_num = len(self.data)

        elif SQL_step == 1:  # update edited record
            CurrentID = self.data[self.rec_num][0]
            realnames.remove(self.field)
            del ValueList[0]

            SQL_str = ','.join(["%s=?" % (name) for name in realnames])
            UpQuery = ('UPDATE ' + self.tblname + ' SET ' + SQL_str +
                       ' WHERE ' + self.field + ' = ' + str(CurrentID))
            Dbase().TblEdit(UpQuery, ValueList)
            self.data = Dbase().Dsqldata(self.MainSQL)

        elif SQL_step == 3:
            return

        self.b2.Disable()
        return CurrentID

    def OnRestoreCmbs(self, evt):
        self.RestoreCmbs()

    def OnMovefst(self, evt):
        self.rec_num = 0
        self.FillScreen()

    def OnMovelst(self, evt):
        if len(self.data) == 0:
            return
        self.rec_num = len(self.data)-1
        self.FillScreen()

    def OnMovenxt(self, evt):
        if len(self.data) == 0:
            return
        self.rec_num += 1
        if self.rec_num == len(self.data):
            self.rec_num = 0
        self.FillScreen()

    def OnMovepre(self, evt):
        if len(self.data) == 0:
            return
        self.rec_num -= 1
        if self.rec_num < 0:
            self.rec_num = len(self.data)-1
        self.FillScreen()

    def ReportName(self):

        saveDialog = wx.FileDialog(self, message='Save Report as PDF.',
                                   wildcard='PDF (*.pdf)|*.pdf',
                                   style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT)
        if saveDialog.ShowModal() == wx.ID_CANCEL:
            filename = ''
        filename = saveDialog.GetPath()
        if filename.find(".pdf") == -1:
            filename = filename + '.pdf'
        saveDialog.Destroy()

        if not filename:
            exit()

        return filename

    def PrintFile(self, evt):
        import Report_Conduit

        if self.data == []:
            NoData = wx.MessageDialog(
                None, 'No Data to Print', 'Error', wx.OK | wx.ICON_EXCLAMATION)
            NoData.ShowModal()
            return

        ttl = self.tblname
        if self.tblname == 'Piping':
            if len(self.data) == 1:
                # specify a title for the report if table name
                # is not to be used
                if self.PipeCode is None:
                    ttl = self.ComCode + '-' + self.cmbMtr.GetValue()
                else:
                    ttl = self.PipeCode
                ttl = self.tblname + ' for ' + ttl

        filename = self.ReportName()

        if self.tblname == 'Piping':
            Colnames = [('Pipe ID', 'Pipe Material Specification'),
                        ('Minimum\nPipe OD', 'Maximum\nPipe OD',
                         'Pipe\nMaterial', 'Pipe\nSchedule'),
                        ('Minimum\nNipple OD', 'Maximum\nNipple OD',
                         'Nipple\nMaterial', 'Nipple\nSchedule',
                         'End\nConnection')]

            Colwdths = [(10, 10),
                        (10, 10, 20, 10), (10, 10, 20, 10, 10)]

            rptdata = self.ReportdataPipe()

        if self.tblname == 'Tubing':
            Colnames = [('Tube ID', 'Notes'),
                        ('Tube\nSize', 'Tube Wall\nThickness',
                         'Tube\nMaterial'),
                        ('Valve\nType', 'Valve Body\nMaterial',
                         'Valve\nManufacture', 'Model\nNumber')]

            Colwdths = [(10, 40),
                        (10, 10, 20), (15, 20, 20, 15)]

            rptdata = self.ReportdataTube()

        Report_Conduit.Report(self.tblname, rptdata, Colnames,
                              Colwdths, ttl, filename).create_pdf()

    def ReportdataTube(self):
        rptdata = []

        for segn in range(0, len(self.data)):
            data1 = [(self.data[segn][i]) for i in [0, 18]]
            data2 = [x for x in list(self.data[segn][1:9]) if x is not None]
            data3 = [x for x in list(self.data[segn][10:17]) if x is not None]

            n = 0
            for i in data2:
                # first column numbers
                if n in [m*3 for m in range(0, 3)]:
                    qry = ("SELECT OD FROM TubeSize WHERE SizeID = "
                           + str(i))
                # second column numbers
                elif n in [m*3+1 for m in range(0, 3)]:
                    qry = ("SELECT Wall FROM TubeWall WHERE WallID = "
                           + str(i))
                # third column numbers
                elif n in [m*3+2 for m in range(0, 3)]:
                    qry = ('''SELECT Material FROM TubeMaterial
                            WHERE ID = ''' + str(i))
                data2[n] = str(Dbase().Dsqldata(qry)[0][0])
                n += 1

            m = 0
            for i in data3:
                # first column numbers
                if m in [p*4 for p in range(0, 2)]:
                    qry = ('''SELECT Tube_Valve_Body_Material FROM
                            TubeValveMatr WHERE TVID = ''' + str(i))
                    data3[m] = str(Dbase().Dsqldata(qry)[0][0])
                # second column numbers
                elif m in [p*4+1 for p in range(0, 2)]:
                    data3[m] = i
                # third column numbers
                elif m in [p*4+2 for p in range(0, 2)]:
                    data3[m] = i
                # fourth column numbers
                elif m in [p*4+3 for p in range(0, 2)]:
                    data3[m] = i

                m += 1

            rptdata2 = []
            rptdata3 = []
            for n in range(0, 3):
                if data2 != []:
                    rptdata2.append(tuple(data2[n*3:(n+1)*3]))
            for m in range(0, 2):
                if data3 != []:
                    rptdata3.append(tuple(data3[m*4:(m+1)*4]))

            rptdata.append(data1)
            rptdata.append(rptdata2)
            rptdata.append(rptdata3)

        return rptdata

    def ReportdataPipe(self):
        rptdata = []

        for segn in range(0, len(self.data)):
            data1 = list(self.data[segn][0:2])
            data2 = [x for x in list(self.data[segn][2:18]) if x is not None]
            data3 = [x for x in list(self.data[segn][18:28]) if x is not None]

            # determine the matr spec designation used for all tables
            query = ('''SELECT Pipe_Material_Spec FROM PipeMaterialSpec
                    WHERE Material_Spec_ID = ''' + str(data1[1]))
            data1[1] = Dbase().Dsqldata(query)[0][0]

            n = 0
            for i in data2:
                # first column numbers
                if n in [m*4 for m in range(0, 4)]:
                    qry = ("SELECT Pipe_OD FROM Pipe_OD WHERE PipeOD_ID = "
                           + str(i))
                # second column numbers
                elif n in [m*4+1 for m in range(0, 4)]:
                    qry = ("SELECT Pipe_OD FROM Pipe_OD WHERE PipeOD_ID = "
                           + str(i))
                # third column numbers
                elif n in [m*4+2 for m in range(0, 4)]:
                    qry = ('''SELECT Pipe_Material FROM PipeMaterial
                            WHERE ID = ''' + str(i))
                # fourth column numbers
                elif n in [m*4+3 for m in range(0, 4)]:
                    qry = ('''SELECT Pipe_Schedule FROM PipeSchedule
                            WHERE ID = ''' + str(i))
                data2[n] = str(Dbase().Dsqldata(qry)[0][0])
                n += 1

            m = 0
            for i in data3:
                # first column numbers
                if m in [p*5 for p in range(0, 2)]:
                    qry = ("SELECT Pipe_OD FROM Pipe_OD WHERE PipeOD_ID = "
                           + str(i))
                # second column numbers
                elif m in [p*5+1 for p in range(0, 2)]:
                    qry = ("SELECT Pipe_OD FROM Pipe_OD WHERE PipeOD_ID = "
                           + str(i))
                # third column numbers
                elif m in [p*5+2 for p in range(0, 2)]:
                    qry = ('''SELECT Pipe_Material FROM PipeMaterial
                            WHERE ID = ''' + str(i))
                # fourth column numbers
                elif m in [p*5+3 for p in range(0, 2)]:
                    qry = ('''SELECT Pipe_Schedule FROM PipeSchedule
                            WHERE ID = ''' + str(i))
                # fifth column numbers
                elif m in [p*5+4 for p in range(0, 2)]:
                    qry = ('''SELECT Connection FROM EndConnects
                            WHERE EndID = ''' + str(i))
                data3[m] = str(Dbase().Dsqldata(qry)[0][0])
                m += 1

            rptdata2 = []
            rptdata3 = []
            for n in range(0, 4):
                if data2 != []:
                    rptdata2.append(tuple(data2[n*4:(n+1)*4]))
            for m in range(0, 2):
                if data3 != []:
                    rptdata3.append(tuple(data3[m*5:(m+1)*5]))

            rptdata.append(data1)
            rptdata.append(rptdata2)
            rptdata.append(rptdata3)

        return rptdata

    def FillScreen(self):
        # all the IDs for the various tables making up the package
        if len(self.data) == 0:
            self.recnum2.SetLabel(str(self.rec_num))
            return
        else:
            recrd = self.data[self.rec_num]

        if self.tblname == 'Tubing':
            # 3 sets of cmbbxs for TubeSize, Tubewall & TubeMaterial
            cmbvals = [(0, 3, 6), (1, 4, 7), (2, 5, 8)]
            increment = 1

        if self.tblname == 'Piping':
            cmbvals = [(0, 4, 8, 12), (1, 5, 9, 13), (2, 6, 10, 14),
                       (3, 7, 11, 15)]
            # cycle thru each combobox
            increment = 2
            query = ('SELECT * FROM PipeMaterialSpec WHERE Material_Spec_ID = '
                     + str(recrd[1]))
            self.cmbMtr.ChangeValue(Dbase().Dsqldata(query)[0][1])

        for n in cmbvals:
            # collect all the data from the tube size table that is
            # equal to the size values in the tubing table
            for i in n:
                if recrd[i+increment] is None:
                    self.cmbs[i].ChangeValue('')
                    continue
                query = ('SELECT * FROM ' + self.top_hints_tbls[n[0]][1] +
                         ' WHERE ' + self.top_hints_tbls[n[0]][4] + ' = ' +
                         str(recrd[i+increment]))
                val = Dbase().Dsqldata(query)
                if val == []:
                    self.cmbs[i].ChangeValue('')
                else:
                    self.cmbs[i].ChangeValue(val[0][1])
                    if self.top_hints_tbls[n[0]][0] == 'Schedule':
                        self.cmbs[i].ChangeValue(val[0][2])
                    else:
                        self.cmbs[i].ChangeValue(val[0][1])

        if self.tblname == 'Tubing':
            # txtvals are (database record number, txtbox number)
            m = 11
            txtvals = [(11, 0), (12, 1), (13, 2), (15, 3), (16, 4), (17, 5)]

        # the piping form has no textboxes
        if self.tblname == 'Piping':
            txtvals = []
        for v in txtvals:
            if recrd[v[0]] is not None:
                self.txtbxs[v[1]].ChangeValue(recrd[v[0]])
            else:
                self.txtbxs[v[1]].ChangeValue('')

        if self.tblname == 'Tubing':
            # 2 cmbbxs for TubeValveMaterial,first number is cmbbox
            # index second is location in datastring
            btmcmbvals = [(0, 10), (1, 14)]
            self.text6.ChangeValue(str(recrd[-1]))

        if self.tblname == 'Piping':
            m = 18   # datatble starting record for the first bottom combo box
            btmcmbvals = [(i, i+m) for i in range(0, len(self.btmcmbs))]

        # cycle thru each combobox
        m = 0
        for n in btmcmbvals:
            if recrd[n[1]] is None:
                self.btmcmbs[n[0]].ChangeValue('')
                continue
            # collect all the data from the tube size table
            # that is equal to the size values in the tubing table
            query = ('SELECT * FROM ' + self.btm_hints_tbls[m][1] +
                     ' WHERE ' + self.btm_hints_tbls[m][4] + ' = ' +
                     str(recrd[n[1]]))
            val = Dbase().Dsqldata(query)
            if val == []:
                self.btmcmbs[n[0]].ChangeValue('')
            else:
                if self.btm_hints_tbls[m][0] == 'Schedule':
                    self.btmcmbs[n[0]].ChangeValue(val[0][2])
                else:
                    self.btmcmbs[n[0]].ChangeValue(val[0][1])
            m += 1
            if m > (len(btmcmbvals)/2-1):
                m = 0

        self.recnum2.SetLabel(str(self.rec_num+1))

    def OnAdd1(self, evt):
        btn = evt.GetEventObject()
        callbtn = self.btnDict[btn]

        if callbtn < self.TopRows:
            # numbers are for the index number for each group of comboboxes
            boxnums = [n*len(self.top_hints_tbls)+callbtn
                       for n in range(0, self.TopRows)]
            cmbtbl = self.top_hints_tbls[callbtn][1]
            CmbLst1(self, cmbtbl)
        else:
            # numbers are for the index number for each group of comboboxes
            if self.tblname == 'Piping':
                boxnums = [n*len(self.btm_hints_tbls)+12+callbtn
                           for n in range(0, self.BtmRows)]
            if self.tblname == 'Tubing':
                boxnums = [9, 10]
            cmbtbl = self.btm_hints_tbls[callbtn-self.TopRows][1]
            CmbLst1(self, cmbtbl)

        self.ReFillList(cmbtbl, boxnums)

    def ReFillList(self, cmbtbl, boxnums):
        for n in boxnums:
            self.lc = self.lctrls[n]
            self.lc.DeleteAllItems()
            index = 0
            ReFillQuery = 'SELECT * FROM "' + cmbtbl + '"'
            for values in Dbase().Dsqldata(ReFillQuery):
                col = 0
                for value in values:
                    if col == 0:
                        self.lc.InsertItem(index, str(value))
                    else:
                        self.lc.SetItem(index, col, str(value))
                    col += 1
                index += 1

    # Return the widget that is to be used for the popup
    def GetControl(self):
        return self.lc

    # Return a string representation of the current item.
    def GetStringValue(self):
        if self.value == -1:
            return
        return self.lc.GetItemText(self.value, self.showcol)

    def ValData(self):
        DataStrg = []
        DialogStr = ''

        if self.tblname == 'Piping':
            if self.cmbMtr.GetValue() == '':
                wx.MessageBox('''Value required for the Pipe Material
                               Spec before saving the data.''',
                              'Missing Data', wx.OK | wx.ICON_INFORMATION)
                return False

        for n in range(len(self.cmbs)):
            if self.cmbs[n].GetValue() == '':
                DialogStr = DialogStr + '0'
            else:
                DialogStr = DialogStr + '1'

        for n in range(0, (len(self.top_hints_tbls)*self.TopRows),
                       len(self.top_hints_tbls)):
            strgsum = DialogStr[n:n+4]
            NoData = [pos for pos, char in enumerate(strgsum) if char == '0']
            if len(NoData) not in (0, len(self.top_hints_tbls)):
                DialogStr = ''
                for i in NoData:
                    DataStrg.append(self.cmbs[n+int(i)].GetHint())

        if len(DataStrg) > 0:
            DialogStr = '\n'.join(DataStrg)
            wx.MessageBox('Value needed for;\n' + DialogStr +
                          '\nto complete information.', 'Missing Data',
                          wx.OK | wx.ICON_INFORMATION)
            return False
        else:
            self.b2.Disable()
            return True

    def OnAddComd(self, evt):
        self.AddComd()

    # link this ID to the commodity property
    def AddComd(self):
        # this can happen only for tubing form
        if self.b6.GetLabel() == 'Show All\n' + self.frmtitle:
            if self.MainSQL == '':
                self.MainSQL = 'SELECT * FROM ' + self.tblname
                self.data = Dbase().Dsqldata(self.MainSQL)
            else:
                self.MainSQL = self.MainSQL[:self.MainSQL.find('WHERE')]
                self.data = Dbase().Dsqldata(self.MainSQL)
            self.b6.SetLabel("Add Item\nto Commodity")
            self.b6.Enable()
            self.b3.Disable()
            self.FillScreen()
            self.recnum3.SetLabel('/ '+str(len(self.data)))
        else:        # this will be the default for the piping form
            query = ('SELECT ' + self.field2 +
                     ' FROM PipeSpecification WHERE Commodity_Property_ID = '
                     + str(self.ComdPrtyID))
            StartQry = Dbase().Dsqldata(query)
            if StartQry == []:
                ValueList = []
                New_ID = cursr.execute(
                    "SELECT MAX(Pipe_Spec_ID) FROM PipeSpecification").\
                    fetchone()
                if New_ID[0] is None:
                    Max_ID = '1'
                else:
                    Max_ID = str(New_ID[0]+1)
                colinfo = Dbase().Dcolinfo('PipeSpecification')
                for n in range(0, len(colinfo)-2):
                    ValueList.append(None)

                num_vals = ('?,'*len(colinfo))[:-1]
                ValueList.insert(0, Max_ID)
                ValueList.insert(1, str(self.ComdPrtyID))

                UpQuery = ("INSERT INTO PipeSpecification VALUES ("
                           + num_vals + ")")
                Dbase().TblEdit(UpQuery, ValueList)
                StartQry = Max_ID
            else:
                StartQry = str(StartQry[0][0])

            cmd_addID = self.data[self.rec_num][0]
            self.ChgSpecID(cmd_addID)

            self.MainSQL = ('SELECT * FROM ' + self.tblname + ' WHERE ' +
                            self.field + ' = ' + str(cmd_addID))
            self.data = Dbase().Dsqldata(self.MainSQL)

            self.rec_num = 0
            self.FillScreen()
            self.lbl_nodata.SetLabel('   ')
            self.recnum3.SetLabel('/ '+str(len(self.data)))
            self.b6.Disable()
            self.b3.Enable()
            self.b6.SetLabel("Show All\n" + self.frmtitle)

    def ChgSpecID(self, ID=None):
        UpQuery = ('UPDATE PipeSpecification SET ' + self.field2 +
                   ' = ? WHERE Commodity_Property_ID = ' +
                   str(self.ComdPrtyID))
        Dbase().TblEdit(UpQuery, [ID])

    def OnCmbOpen(self, evt):
        self.cmbOld = evt.GetEventObject().GetValue()

    def OnCmbClose(self, evt):
        if self.cmbOld != evt.GetEventObject().GetValue():
            self.b2.Enable()
            self.cmbMtr.Enable()

    def OnCmbMtrClose(self, evt):
        self.update = 2
        self.cmbMtrNew = self.cmbMtr.GetValue()
        if self.cmbMtrOld == self.cmbMtrNew:
            self.update = 1

        self.CmbMtrClose()
        self.b2.Disable()

    def CmbMtrClose(self):
        # convert the material code selected to the material ID
        qry = ('''SELECT Material_Spec_ID FROM PipeMaterialSpec
                WHERE Pipe_Material_Spec LIKE '%''' + self.cmbMtrNew +
               "%' COLLATE NOCASE")
        SpecID = Dbase().Dsqldata(qry)[0][0]
        # use the material ID to source all records in the piping table
        qry = 'SELECT * FROM Piping WHERE Spec = ' + str(SpecID)
        self.data = Dbase().Dsqldata(qry)
        if self.data != []:
            self.recnum3.SetLabel('/ '+str(len(self.data)))
            self.rec_num = 0
            self.b3.Disable()
            self.FillScreen()
            if self.ComdPrtyID:
                self.b6.SetLabel("Add Item\nto Commodity")
                self.b6.Enable()

        # if nothing has been set up in the piping table for this material code
        else:
            self.data = []
            self.RestoreCmbs()
            self.cmbMtr.ChangeValue(self.cmbMtrNew)
            self.rec_num = 0
            self.recnum2.SetLabel(str(self.rec_num))
            self.recnum3.SetLabel('/ 0')

    def OnCmbMtrOpen(self, evt):
        self.cmbMtrOld = self.cmbMtr.GetValue()

    def OnSelect(self, evt):
        self.b2.Enable()
        self.cmbMtr.Enable()

    def RestoreCmbs(self):
        for cmb in self.cmbs:
            cmb.Clear()
            cmb.ChangeValue('')
        for cmb in self.btmcmbs:
            cmb.Clear()
            cmb.ChangeValue('')
        for txtbx in self.txtbxs:
            txtbx.ChangeValue('')

        if self.tblname == 'Tubing':
            self.text6.ChangeValue('')

        if self.tblname == 'Piping':
            self.cmbMtr.ChangeValue('')

    def OnDelete(self, evt):
        if self.data != []:
            recrd = self.data[self.rec_num][0]
            if self.ComdPrtyID is None:
                try:
                    Dbase().TblDelete(self.tblname, recrd, self.field)
                    self.MainSQL = 'SELECT * FROM ' + self.tblname
                    self.data = Dbase().Dsqldata(self.MainSQL)
                    self.rec_num -= 1
                    if self.rec_num < 0:
                        self.rec_num = len(self.data)-1
                    self.FillScreen()
                    self.recnum3.SetLabel('/ '+str(len(self.data)))

                except sqlite3.IntegrityError:
                    wx.MessageBox("This Record is associated"
                                  " with\nother tables and cannot be deleted!",
                                  "Cannot Delete",
                                  wx.OK | wx.ICON_INFORMATION)
            else:
                self.ChgSpecID()
                self.RestoreCmbs()
                self.b3.Disable()
                self.b6.Enable()
                self.data = []
                self.rec_num = 0
                self.recnum2.SetLabel(str(self.rec_num))
                self.recnum3.SetLabel('/ '+str(len(self.data)))
                self.lbl_nodata.SetLabel(
                    'The ' + self.frmtitle +
                    ' have not been setup for this Commodity Property')

    def OnClose(self, evt):
        self.GetParent().Enable(True)   # add this line for child
        self.__eventLoop.Exit()     # add this line for child
        self.Destroy()


class BldValve(wx.Frame):
    '''Routine to build form and populate grid'''
    def __init__(self, parent, tblname, ComdPrtyID=None, model=None):

        self.parent = parent
        self.model = model

        self.ComdPrtyID = ComdPrtyID
        self.mtr = None
        self.VlvIDs = []

        self.tblname = tblname
        if self.tblname.find("_") != -1:
            self.frmtitle = (self.tblname.replace("_", " "))
        else:
            self.frmtitle = (' '.join(re.findall('([A-Z][a-z]*)',
                             self.tblname)))

        super(BldValve, self).__init__(parent,
                                       title=self.frmtitle,
                                       size=(870, 680))

        self.Bind(wx.EVT_CLOSE, self.OnClose)
        self.InitUI()

    def InitUI(self):
        self.data = self.LoadData()

        # specify which listbox column to display in the combobox
        self.showcol = int

        # set up the table column names, width and if column can be
        # edited ie primary autoincrement
        tblinfo = []

        tblinfo = Dbase(self.tblname).Fld_Size_Type()
        self.ID_col = tblinfo[1]
        self.pkcol_name = tblinfo[0]
        self.autoincrement = tblinfo[2]
        columnames = Dbase(self.tblname).ColNames()

        self.txtparts = {}
        self.search = False
        self.grid_select = False

        if self.frmtitle == 'Gate Valve':
            self.txtparts[0] = 'GT'
        elif self.frmtitle == 'Ball Valve':
            self.txtparts[0] = 'BL'
        elif self.frmtitle == 'Plug Valve':
            self.txtparts[0] = 'PG'
        elif self.frmtitle == 'Globe Valve':
            self.txtparts[0] = 'GL'
        elif self.frmtitle == 'Butterfly Valve':
            self.txtparts[0] = 'BF'
        elif self.frmtitle == 'Swing Check Valve':
            self.txtparts[0] = 'SC'
        elif self.frmtitle == 'Piston Check Valve':
            self.txtparts[0] = 'PC'

        # Create a dataview control
        self.dvc = dv.DataViewCtrl(self, wx.ID_ANY, wx.DefaultPosition,
                                   wx.Size(500, 300),
                                   style=wx.BORDER_THEME
                                   | dv.DV_ROW_LINES
                                   | dv.DV_VERT_RULES
                                   | dv.DV_HORIZ_RULES
                                   | dv.DV_SINGLE
                                   )
        self.dvc.SetMinSize = (wx.Size(100, 200))
        self.dvc.SetMaxSize = (wx.Size(500, 400))

        # if autoincrement is false then the data can be sorted based on ID_col
        if self.autoincrement == 0:
            self.data.sort(key=lambda tup: tup[self.ID_col])

        # use the sorted data to load the dataviewlistcontrol
        if self.model is None:
            self.model = DataMods(self.tblname, self.data)

        self.dvc.AssociateModel(self.model)

        n = 0
        for colname in columnames:
            self.dvc.AppendTextColumn(colname, n,
                                      width=wx.LIST_AUTOSIZE_USEHEADER,
                                      mode=dv.DATAVIEW_CELL_INERT)
            n += 1

        # make columns not sortable and but reorderable.
        for c in self.dvc.Columns:
            c.Sortable = False
            c.Reorderable = True
            c.Resizeable = True

        # set the Sizer property (same as SetSizer)
        self.Sizer = wx.BoxSizer(wx.VERTICAL)

        # change to not let the ID col be moved.
        self.dvc.Columns[(self.ID_col)].Reorderable = False
        self.dvc.Columns[(self.ID_col)].Resizeable = False

        self.dvsizer = wx.BoxSizer(wx.HORIZONTAL)
        self.dvsizer.Add(10, -1, 0)
        self.dvsizer.Add(self.dvc, 1, wx.EXPAND)
        self.dvsizer.Add(10, -1, 0)

        # set up first row of combo boxesand label
        self.searchsizer = wx.BoxSizer(wx.HORIZONTAL)
        self.cmbsizer1 = wx.BoxSizer(wx.HORIZONTAL)
        self.cmbsizer2 = wx.BoxSizer(wx.HORIZONTAL)
        self.cmbsizer3 = wx.BoxSizer(wx.HORIZONTAL)
        self.cmbsizer4 = wx.BoxSizer(wx.HORIZONTAL)

        font = wx.Font(12, wx.DEFAULT, wx.NORMAL, wx.NORMAL)
        font2 = wx.Font(12, wx.DEFAULT, wx.NORMAL, wx.FONTWEIGHT_BOLD)

        txt_nodata = (
            ' The ' + self.frmtitle +
            ' have not been setup for this Commodity Property')
        self.lbl_nodata = wx.StaticText(self, -1, label='',
                                        size=(650, 30), style=wx.LEFT)
        self.lbl_nodata.SetForegroundColour((255, 0, 0))
        self.lbl_nodata.SetFont(font2)

        if self.VlvIDs == []:
            self.lbl_nodata.SetLabel(txt_nodata)
        else:
            self.lbl_nodata.SetLabel('   ')

        self.warningsizer = wx.BoxSizer(wx.HORIZONTAL)
        self.warningsizer.Add(self.lbl_nodata, 0, wx.ALIGN_CENTER)

        self.text1 = wx.TextCtrl(self, size=(350, 33), value='Valve Code',
                                 style=wx.TE_READONLY | wx.TE_CENTER)
        self.text1.SetForegroundColour((255, 0, 0))
        self.text1.SetFont(font)

        if self.ComdPrtyID is None:
            self.lbl_nodata.SetLabel('   ')
            self.searchsizer.Add((20, 20))
        else:      # show commodity property ticker
            if self.PipeCode is not None:
                txt = self.PipeCode
            else:
                query = ('''SELECT * FROM PipeMaterialSpec WHERE
                          Material_Spec_ID = ''' + str(self.PipeMtrSpec))
                MtrSpc = Dbase().Dsqldata(query)[0][1]
                txt = self.ComCode + ' - ' + MtrSpc
            self.text2 = wx.TextCtrl(self, size=(130, 33), value=txt,
                                     style=wx.TE_READONLY | wx.TE_CENTER)
            self.text2.SetForegroundColour((255, 0, 0))
            self.text2.SetFont(font)
            self.searchsizer.Add(self.text2, 0, wx.ALIGN_LEFT, 5)

        # add a button to call main form to search combo list data
        self.b1 = wx.Button(self, label="<=Search Data")
        self.Bind(wx.EVT_BUTTON, self.OnSearch, self.b1)
        self.searchsizer.Add(self.text1, 0, wx.ALIGN_LEFT, 5)
        self.searchsizer.Add((10, 10))
        self.searchsizer.Add(self.b1, 0, wx.ALIGN_LEFT, 5)

        # Start the generation of the required combo boxes
        # using a dictionary of column names and table names

        self.hints_tbls = {}
        # list of combobox objects
        self.cmbctrls = []
        # the list of grid column names excluding ID,Valve ID and notes
        self.cmb_range = columnames[2:-1]
        # dictionary using grid column name as the key (combobox hint)
        # for the value (combobox table)
        self.hints_tbls = self.TblInfoRebuild()
        # counter to track which comboboxes need activation
        # after material type is selected
        self.a = []
        n = 1

        # cmb_range is the list of the combo box
        for name in self.cmb_range:
            cmb_tbl = self.hints_tbls[name]
            cmbbox = wx.ComboCtrl(self, size=(200, -1))
            self.Bind(wx.EVT_TEXT, self.OnSelect, cmbbox)

        # list of combo boxes to disable and have
        # hint = '1st Select Material Type'
            if (cmb_tbl.find('Material') != -1
                    and cmb_tbl != 'MaterialType'
                    and cmb_tbl != 'SleeveMaterial'):
                cmbbox.SetHint('1st Select Material Type')
                cmbbox.Disable()
                self.a.append(n-1)
                self.showcol = 2
            else:  # combo boxes to be populated
                cmbbox.SetHint(name)
                self.showcol = 1

            cmbbox.SetPopupControl(ListCtrlComboPopup(
                cmb_tbl, showcol=self.showcol))
            self.cmbctrls.append(cmbbox)

        # position the combo boxes in the correct sizer
            if n <= 4:
                self.cmbsizer1.Add(cmbbox, 0, wx.ALIGN_LEFT, 5)
            elif n <= 8 and n > 4:
                self.cmbsizer2.Add(cmbbox, 0, wx.ALIGN_LEFT, 5)
            elif n <= 12 and n > 8:
                self.cmbsizer3.Add(cmbbox, 0, wx.ALIGN_LEFT, 5)
            else:
                self.cmbsizer4.Add(cmbbox, 0, wx.ALIGN_LEFT, 5)
            n += 1

        self.b5 = wx.Button(self, label="Reset Screen")
        self.Bind(wx.EVT_BUTTON, self.OnRestoreCmbs, self.b5)
        self.cmbsizer4.Add(self.b5, 0, wx.ALIGN_LEFT, 5)

        self.notebox = wx.BoxSizer(wx.HORIZONTAL)
        self.notes = wx.TextCtrl(self, size=(700, 80), value='',
                                 style=wx.TE_MULTILINE | wx.TE_LEFT)
        self.notebox.Add(self.notes, 0, wx.ALIGN_LEFT, 5)

        # Add some buttons
        self.b1 = wx.Button(self, label="Print\nReport")
        self.Bind(wx.EVT_BUTTON, self.PrintFile, self.b1)

        self.b2 = wx.Button(self, label="Add/Update\nValve")
        self.b2.Disable()
        self.Bind(wx.EVT_BUTTON, self.OnAddRow, self.b2)

        self.b3 = wx.Button(self, label="Delete\nRecord")
        self.Bind(wx.EVT_BUTTON, self.OnDeleteRow, self.b3)
        if self.ComdPrtyID is not None:
            self.b3.SetLabel('Remove\nRecord')
        self.b3.Disable()

        self.b4 = wx.Button(self, label="Exit")
        self.Bind(wx.EVT_BUTTON, self.OnClose, self.b4)

        # add a button box and place the buttons
        self.btnbox = wx.BoxSizer(wx.HORIZONTAL)
        self.btnbox.Add(self.b1, 0, wx.ALL | wx.ALIGN_CENTER, 5)
        self.btnbox.Add(self.b2, 0, wx.ALL | wx.ALIGN_CENTER, 5)
        if self.ComdPrtyID is not None:
            self.b6 = wx.Button(self, size=(120, 45), label="Show All\nValves")
            self.Bind(wx.EVT_BUTTON, self.OnAddComd, self.b6)
            self.btnbox.Add(self.b6, 0, wx.ALL | wx.ALIGN_CENTER, 5)
        self.btnbox.Add(self.b3, 0, wx.ALL | wx.ALIGN_CENTER, 5)
        self.btnbox.Add(self.b4, 0, wx.LEFT | wx.BOTTOM |
                        wx.TOP | wx.ALIGN_CENTER, 20)

        # add static label to explain how to add / edit data
        self.editlbl = wx.StaticText(self, -1, style=wx.ALIGN_LEFT)
        txt = '''\tTo edit data double click on the cell, then edit the boxes as needed.
        To add a new record complete the data in the text boxes then click
        "Search Data" to see if the spec already exits, if it does not you will
        be able to save the spec.'''
        self.editlbl.SetLabel(txt)
        self.editlbl.SetForegroundColour((255, 0, 0))

        self.Sizer.Add(self.warningsizer, 0, wx.CENTER | wx.TOP, 5)
        self.Sizer.Add((10, 10))
        self.Sizer.Add(self.searchsizer, 0, wx.ALIGN_CENTER)
        self.Sizer.Add((10, 10))
        self.Sizer.Add(self.cmbsizer1, 0, wx.ALIGN_CENTER)
        self.Sizer.Add(self.cmbsizer2, 0, wx.ALIGN_CENTER)
        self.Sizer.Add(self.cmbsizer3, 0, wx.ALIGN_CENTER)
        self.Sizer.Add(self.cmbsizer4, 0, wx.ALIGN_CENTER)
        self.Sizer.Add(self.dvsizer, 1, wx.EXPAND)
        self.Sizer.Add(self.notebox, 0, wx.ALIGN_CENTER)
        self.Sizer.Add(self.editlbl, 0, wx.ALIGN_CENTER, 5)
        self.Sizer.Add(self.btnbox, 0, wx.ALIGN_CENTER, 5)

        # Bind some events so we can see what the DVC sends us
        self.Bind(dv.EVT_DATAVIEW_ITEM_ACTIVATED, self.OnGridSelect, self.dvc)

        if self.mtr is not None:
            self.cmbctrls[4].SetValue(self.mtr)
            for i in (i for i, x in enumerate(self.cmbctrls)
                      if self.cmb_range[i].find('Connection') != -1):
                tbl = 'EndConnects'
                strg = ()
                if self.ComEnds == 1:
                    strg = (1, 2, 3, 4, 5, 6)
                elif self.ComEnds == 2:
                    strg = (1, 3, 4, 6)
                elif self.ComEnds == 3:
                    strg = (1, 3, 4)
                elif self.ComEnds == 4:
                    strg = (1, 3)
                elif self.ComEnds == 5:
                    strg = (7)
                elif self.ComEnds == 6:
                    strg = (8)
                query = 'SELECT * FROM EndConnects WHERE EndID IN ' + str(strg)
                self.cmbctrls[i].SetPopupControl(ListCtrlComboPopup(
                    tbl, PupQuery=query, cmbvalue='', showcol=1))

            # add these following lines for child parent form
        self.CenterOnParent()
        self.GetParent().Enable(False)
        self.Show(True)
        self.__eventLoop = wx.GUIEventLoop()
        self.__eventLoop.Run()

    # The function builds a dictionary of value=table name and key=column name
    def TblInfoRebuild(self):
        keys = []
        values = []
        TblInfo = Dbase().Dtbldata(self.tblname)

        for element in TblInfo:
            # modify the column names to remove underscore and seperate words
            colname = element[3]
            if colname.find("ID", -2) != -1:
                colname = "ID"
            elif colname.find("_") != -1:
                colname = colname.replace("_", " ")
            else:
                colname = (' '.join(re.findall('([A-Z][a-z]*)', colname)))

            keys.append(colname)
            values.append(element[2])

        tbl_cmb = dict(zip(keys, values))
        return (tbl_cmb)

    def PrintFile(self, evt):
        import Report_General

        if self.data == []:
            NoData = wx.MessageDialog(
                None, 'No Data to Print', 'Error', wx.OK | wx.ICON_EXCLAMATION)
            NoData.ShowModal()
            return

        saveDialog = wx.FileDialog(self, message='Save Report as PDF.',
                                   wildcard='PDF (*.pdf)|*.pdf',
                                   style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT)

        if saveDialog.ShowModal() == wx.ID_CANCEL:
            filename = ''

        filename = saveDialog.GetPath()
        if filename.find(".pdf") == -1:
            filename = filename + '.pdf'

        saveDialog.Destroy()

        columnames = Dbase(self.tblname).ColNames()
        columnames[2] = 'Rating\nDesignation'
        columnames[3] = 'Minimum\nPipe Size'
        columnames[4] = 'Maximum\nPipe Size'
        columnames[6] = 'Matr Type\nDesignation'

        if self.tblname in ['SwingCheckValve', 'PistonCheckValve',
                            'GlobeValve']:
            colwdth = [4, 25, 10, 10, 10, 15, 15, 15, 12, 12, 15,
                       15, 15, 35]
        elif self.tblname in ['PlugValve', 'BallValve']:
            colwdth = [4, 25, 10, 10, 10, 15, 15, 15, 12, 12, 15,
                       15, 15, 10, 35]
        elif self.tblname == 'GateValve':
            columnames[8] = 'Wedge\nMaterial'
            columnames[9] = 'Stem\nMaterial'
            columnames[10] = 'Stem Packing\nGasket'
            columnames[11] = 'Bonnet\nType'
            columnames[12] = 'Seat\nMaterial'
            colwdth = [4, 25, 10, 10, 10, 15, 15, 15, 10, 10, 20,
                       12, 10, 10, 10, 30]
        elif self.tblname == 'ButterflyValve':
            colwdth = [4, 25, 10, 10, 10, 15, 15, 15, 12, 12, 15,
                       15, 35]

        rows_pg = 3
        colms_pg = 8

        ttl = None
        if self.mtr is not None:
            # specify a title for the report if table name
            # is not to be used
            if self.PipeCode is None:
                ttl = self.text2.GetValue()
            else:
                ttl = self.PipeCode
            ttl = ' for ' + ttl

        Report_General.Report(self.frmtitle, self.data, colwdth, columnames,
                              rows_pg, colms_pg, filename, ttl).create_pdf()

    def LoadData(self):
        if self.ComdPrtyID is not None:
            query = ('''SELECT Commodity_Code,Pipe_Material_Code,End_Connection,
                     Pipe_Code FROM CommodityProperties WHERE
                      CommodityPropertyID = '''
                     + str(self.ComdPrtyID))
            dataset = Dbase().Dsqldata(query)[0]
            self.PipeMtrSpec = dataset[1]
            self.ComEnds = dataset[2]
            self.ComCode = dataset[0]
            self.PipeCode = dataset[3]

            # specify the material type on start up but do not
            # limit this for valve material selection
            # i.e. allow the selection of a stainless valve for
            # carbon steel piping if needed
            query = ('''SELECT Pipe_Material_Spec FROM PipeMaterialSpec
                      WHERE Material_Spec_ID = ''' + str(self.PipeMtrSpec))
            mtrspc = Dbase().Dsqldata(query)[0][0]
            query = ('''SELECT Material_Type FROM MaterialType
                      WHERE Type_Designation = "''' + mtrspc[1] + '"')
            self.mtr = Dbase().Dsqldata(query)[0][0]

        DsqlVlv = ('''SELECT v.ID, v.Valve_Code, a.ANSI_class, b.Pipe_OD,
                    c.Pipe_OD, d.Connection, e.Material_Type, f.Body_Material,
                    g.Wedge_Material, ''')

        if self.tblname == 'BallValve':
            DsqlVlv = (DsqlVlv + '''h.Stem_Material, i.Packing_Material,
                        j.Body_Type, k.Seat_Material, l.Porting, v.Notes
                       \n FROM ''')
            self.tblID = 'Ball_Valve_ID'
            self.PrtyTbl = 'BallProperty'
            self.PrtyID = 'BAID'
        elif self.tblname == 'GateValve':
            DsqlVlv = (DsqlVlv + '''h.Stem_Material, i.Packing_Material,
                        j.Bonnet_Type, k.Seat_Material, l.Wedge_Type,
                        m.Porting, v.Notes \n FROM ''')
            self.tblID = 'Gate_Valve_ID'
            self.PrtyTbl = 'GateProperty'
            self.PrtyID = 'GTID'
        elif self.tblname == 'GlobeValve':
            DsqlVlv = (DsqlVlv + '''h.Stem_Material, i.Packing_Material,
                        j.Bonnet_Type, k.Seat_Material, v.Notes \n FROM ''')
            self.tblID = 'Globe_Valve_ID'
            self.PrtyTbl = 'GlobeProperty'
            self.PrtyID = 'GBID'
        elif self.tblname == 'PlugValve':
            DsqlVlv = (DsqlVlv + '''h.Stem_Material, i.Packing_Material,
                        j.Body_Type, k.Sleeve_Material, l.Porting, v.Notes
                        \n FROM ''')
            self.tblID = 'Plug_Valve_ID'
            self.PrtyTbl = 'PlugProperty'
            self.PrtyID = 'PGID'
        elif self.tblname == 'SwingCheckValve':
            DsqlVlv = (DsqlVlv + '''h.Wedge_Material, i.Packing_Material,
                        j.Bonnet_Type, k.Seat_Material, v.Notes \n FROM ''')
            self.tblID = 'Swing_ID'
            self.PrtyTbl = 'SwingProperty'
            self.PrtyID = 'SWID'
        elif self.tblname == 'PistonCheckValve':
            DsqlVlv = (DsqlVlv + '''h.Spring_Material, i.Packing_Material,
                        j.Bonnet_Type, k.Seat_Material, v.Notes \n FROM ''')
            self.tblID = 'Piston_Check_ID'
            self.PrtyTbl = 'PistonProperty'
            self.PrtyID = 'PSID'
        elif self.tblname == 'ButterflyValve':
            DsqlVlv = ('''SELECT v.ID, v.Valve_Code, a.ANSI_class, b.Pipe_OD,
                        c.Pipe_OD, d.Body_Type, e.Material_Type,
                        f.Body_Material, g.Disc_Material, h.Shaft_Material,
                        i.Packing_Material, j.Seat_Material,
                        v.Notes \n FROM ''')
            self.tblID = 'Butterfly_Valve_ID'
            self.PrtyTbl = 'ButterflyProperty'
            self.PrtyID = 'BUID'

        DsqlVlv = DsqlVlv + self.tblname + ' v' + '\n'

        # the index fields are returned from the PRAGMA statement tbldata
        # in reverse order so the alpha characters need to count up not down
        n = 0
        tbldata = Dbase().Dtbldata(self.tblname)
        for element in tbldata:
            alpha = chr(96-n+len(tbldata))
            DsqlVlv = (DsqlVlv + ' INNER JOIN ' + element[2] + ' '
                       + alpha + ' ON v.' + element[3] + ' = ' +
                       alpha + '.' + element[4] + '\n')
            n += 1

        if self.ComdPrtyID is not None:
            Valves = ('SELECT ' + self.tblID + ' FROM ' + self.PrtyTbl +
                      ' WHERE Commodity_Property_ID = ' + str(self.ComdPrtyID))
            self.VlvIDs = Dbase().Dsqldata(Valves)
            ValveIDs = tuple([i[0] for i in self.VlvIDs])

            if len(ValveIDs) == 1:
                self.vlvstrg = '(' + str(ValveIDs[0]) + ')'
            else:
                self.vlvstrg = str(ValveIDs)
            self.DsqlVlv = DsqlVlv + ' WHERE v.ID IN ' + self.vlvstrg
            data = Dbase().Dsqldata(self.DsqlVlv)
        else:
            self.DsqlVlv = DsqlVlv
            data = Dbase().Dsqldata(self.DsqlVlv)

        return data

    def NewTextStr(self):
        s = ''
        orderly = sorted(self.txtparts)
        for n in orderly:
            s = s + str(self.txtparts[n]) + '.'
        s = s[:-1]
        self.text1.ChangeValue(s)

    def cmbSelected(self, cmbselect):
        i = int()
        if self.search is False:
            for i in (i for i, x in enumerate(self.cmbctrls)
                      if x == cmbselect):
                cmbvalue = str(self.cmbctrls[i].GetValue())
                tbl = self.hints_tbls[self.cmb_range[i]]
                if (5 in self.txtparts) and i in self.a:
                    newlbl = self.LblData(tbl, cmbvalue, self.txtparts[5])
                else:
                    newlbl = self.LblData(tbl, cmbvalue)
                self.txtparts[1+i] = newlbl
                self.NewTextStr()

                if i == 4:  # Material Type combobox
                    for n in self.a:
                        if self.grid_select:
                            tbl = self.hints_tbls[self.cmb_range[n]]
                            fldstrg = ''
                            cmbinfo = Dbase().Dcolinfo(tbl)
                            for set in cmbinfo:
                                fldstrg = fldstrg + tbl + '.' + set[1] + ','

                            fldstrg = fldstrg[:-1]
                            # what to do if there is no data to
                            # populate the lists with
                            query = ('SELECT ' + fldstrg + ' FROM ' + tbl +
                                     ' INNER JOIN MaterialType ON ')
                            query = (query + tbl + '''.Material_Type =
                                      MaterialType.Type_Designation WHERE ''')
                            query = (query + "MaterialType.Material_Type = '"
                                     + cmbvalue + "'")
                            self.cmbctrls[n].ChangeValue('')
                            self.cmbctrls[n].SetHint(self.cmb_range[n])
                            self.cmbctrls[n].SetPopupControl(
                                ListCtrlComboPopup(tbl, PupQuery=query,
                                                   cmbvalue=cmbvalue,
                                                   showcol=2))
                        else:
                            tbl = self.hints_tbls[self.cmb_range[n]]
                            fldstrg = ''
                            cmbinfo = Dbase().Dcolinfo(tbl)
                            for set in cmbinfo:
                                fldstrg = fldstrg + tbl + '.' + set[1] + ','

                            fldstrg = fldstrg[:-1]

                            query = ('SELECT ' + fldstrg + ' FROM ' +
                                     tbl + ' INNER JOIN MaterialType ON ')
                            query = (query + tbl + '''.Material_Type =
                                      MaterialType.Type_Designation WHERE ''')
                            query = (query + "MaterialType.Material_Type = '"
                                     + cmbvalue + "'")

                            self.cmbctrls[n].SetHint(self.cmb_range[n])
                            self.cmbctrls[n].Enable()
                            self.cmbctrls[n].SetPopupControl(
                                ListCtrlComboPopup(tbl, PupQuery=query,
                                                   cmbvalue=cmbvalue,
                                                   showcol=2))

        self.grid_select = False

    # this builds the ValveID text string
    def LblData(self, tbl_name, cmbvalue, condition=None):
        self.query3 = ''

        if cmbvalue == "":
            return

        if condition:
            mtrstg = "' AND Material_Type = '" + condition + "'"
        else:
            mtrstg = "'"

        if tbl_name == 'ANSI_Rating':
            # this is used if the combobox shows a numberic value
            self.query3 = ('SELECT Rating_Designation FROM ' + tbl_name
                           + ' WHERE ANSI_Class = ' + cmbvalue)
        elif tbl_name == 'Pipe_OD':
            # this query is used when the combobox shows a string value
            self.query3 = ('SELECT PipeOD_ID FROM ' + tbl_name +
                           " WHERE Pipe_OD = '" + cmbvalue + mtrstg)
        elif tbl_name == 'EndConnects':
            self.query3 = ('SELECT EndID FROM ' + tbl_name +
                           " WHERE Connection = '" + cmbvalue + mtrstg)
        elif tbl_name == 'MaterialType':
            self.query3 = ('SELECT Type_Designation FROM ' + tbl_name
                           + " WHERE Material_Type = '" + cmbvalue + mtrstg)
        elif tbl_name == 'BodyType' or tbl_name == 'ButterflyBodyType':
            self.query3 = ('SELECT BodyTypeID FROM ' + tbl_name +
                           " WHERE Body_Type = '" + cmbvalue + mtrstg)
        elif tbl_name == 'BonnetType':
            self.query3 = ('SELECT BonnetTypeID FROM ' + tbl_name +
                           " WHERE Bonnet_Type = '" + cmbvalue + mtrstg)
        elif tbl_name == 'ValveBodyMaterial' or \
                tbl_name == 'ButterflyBodyMaterial':
            self.query3 = ('SELECT BodyMaterialID FROM ' + tbl_name +
                           " WHERE Body_Material = '" + cmbvalue + mtrstg)
        elif tbl_name == 'ButterflyShaftMaterial':
            self.query3 = ('SELECT ShaftMaterialID FROM ' + tbl_name +
                           " WHERE Shaft_Material = '" + cmbvalue + mtrstg)
        elif tbl_name == 'WedgeMaterial':
            self.query3 = ('SELECT WedgeMaterialID FROM ' + tbl_name +
                           " WHERE Wedge_Material = '" + cmbvalue + mtrstg)
        elif tbl_name == 'ButterflyDiscMaterial':
            self.query3 = ('SELECT DiscMaterialID FROM ' + tbl_name +
                           " WHERE Disc_Material = '" + cmbvalue + mtrstg)
        elif tbl_name == 'StemMaterial':
            self.query3 = ('SELECT StemMaterialID FROM ' + tbl_name +
                           " WHERE Stem_Material = '" + cmbvalue + mtrstg)
        elif tbl_name == 'PackingMaterial':
            self.query3 = ('SELECT StemPackingID FROM ' + tbl_name +
                           " WHERE Packing_Material = '" + cmbvalue + mtrstg)
        elif tbl_name == 'SpringMaterial':
            self.query3 = ('SELECT Spring_MaterialID FROM ' + tbl_name +
                           " WHERE Spring_Material = '" + cmbvalue + mtrstg)
        elif tbl_name == 'SeatMaterial' or tbl_name == 'ButterflySeatMaterial':
            self.query3 = ('SELECT SeatMaterialID FROM ' + tbl_name +
                           " WHERE Seat_Material = '" + cmbvalue + mtrstg)
        elif tbl_name == 'SleeveMaterial':
            self.query3 = ('SELECT Sleeve_ID FROM ' + tbl_name +
                           " WHERE Sleeve_Material = '" + cmbvalue + mtrstg)
        elif tbl_name == 'Porting':
            self.query3 = ('SELECT ID FROM ' + tbl_name +
                           " WHERE Porting = '" + cmbvalue + mtrstg)
        elif tbl_name == 'WedgeType':
            self.query3 = ('SELECT WedgeTypeID FROM ' + tbl_name +
                           " WHERE Wedge_Type = '" + cmbvalue + mtrstg)

        lbldata = Dbase().Dsqldata(self.query3)
        lblstr = str(lbldata[0][0])
        return lblstr

    def RestoreCmbs(self):
        n = 0
        for cmbbox in self.cmbctrls:
            cmbbox.Clear()
            cmbbox.ChangeValue('')
            if n in self.a:
                cmbbox.SetHint('1st Select Material Type')
                cmbbox.Disable()
            else:
                cmbbox.SetHint(self.cmb_range[n])
            n += 1

        self.b2.SetLabel("Add/Update\nValve")
        self.b2.Disable()
        self.b3.Disable()

        temptxt = self.txtparts[0]
        self.txtparts = {}
        self.txtparts[0] = temptxt
        self.text1.ChangeValue('Valve Code')
        self.text1.SetFocus()

        if self.ComdPrtyID is not None:
            if self.b6.GetLabel() == 'Add Valve\nto Commodity':
                if self.DsqlVlv.find('WHERE') == -1:
                    self.DsqlVlv = (self.DsqlVlv + ' WHERE v.ID IN '
                                    + self.vlvstrg)
                else:
                    self.DsqlVlv = self.DsqlVlv[:self.DsqlVlv.find('WHERE')]
                    self.DsqlVlv = (self.DsqlVlv + ' WHERE v.ID IN '
                                    + self.vlvstrg)

            self.data = Dbase().Dsqldata(self.DsqlVlv)
            self.model = DataMods(self.tblname, self.data)
            self.dvc.AssociateModel(self.model)
            self.dvc.Refresh

            self.b6.Enable(True)
            self.b6.SetLabel("Show All\nValves")

    def AddRow(self, SQL_step):
        cmbvalues = []
        self.b2.Disable()
        tbl_ID = None
        code = self.text1.GetValue()

        # build the list of values in comboboxes to save or update to database
        # gets the valve code
        cmbvalues.append(code)
        # adds the other combo box values to the list
        n = 0
        if len(self.cmb_range) == code.count('.'):
            for cmbbox in self.cmbctrls:
                cmbvalues.append(self.cmbctrls[n].GetValue())
                n += 1
        # adds last item the notes box to the database list
        cmbvalues.append(self.notes.GetValue())

        # if a new valve is being added do this
        if SQL_step == 0:
            txtparts_values = code.split('.')
            # remove the valve type designation from the code string
            del txtparts_values[0]
            # remove and store the material type from the
            # string leaving only intergers

            data_str = ''
            for char in txtparts_values:
                if char.isalpha():
                    data_str = data_str + ',"'+char+'"'
                else:
                    data_str = data_str + ','+char

            data_str = "'" + code + "'" + data_str
            data_str = data_str + ",'" + self.notes.GetValue() + "'"

            # get the next available autoincremented value for the table
            IDE = cursr.execute(
                "SELECT MAX(ID) FROM " + self.tblname).fetchone()[0]
            if IDE is None:
                IDE = 0
            IDE = int(IDE+1)

            UpQuery = ('INSERT INTO ' + self.tblname + ' VALUES ('
                       + str(IDE) + ',' + data_str + ')')
            Dbase().TblEdit(UpQuery)
            return IDE
        # if an existing valve is being updated do the following
        elif SQL_step == 1:
            item = self.dvc.GetSelection()
            rowA = self.model.GetRow(item)

            tbl_ID = self.model.GetValueByRow(rowA, 0)

            txtparts_values = code.split('.')
            for char in txtparts_values:
                if char.isalpha():
                    txtparts_values.remove(char)
                    mtrType = char

            realnames = []
            for item in Dbase().Dcolinfo(self.tblname):
                realnames.append(item[1])
            realnames.remove('ID')
            realnames.remove('Notes')
            realnames.remove('Valve_Code')
            realnames.remove('Material_Type_Designation')

            SQL_str = dict(zip(realnames, txtparts_values))
            Update_str = (", ".join(["%s=%s" % (k, v)
                          for k, v in SQL_str.items()]))
            UpQuery = ('UPDATE ' + self.tblname + ' SET Valve_Code = "' +
                       code + '", Material_Type_Designation = "' + mtrType +
                       '", ' + Update_str + ', Notes = "' +
                       self.notes.GetValue() + '"' + ' WHERE ID = '
                       + str(tbl_ID))

            Dbase().TblEdit(UpQuery)

        # if cancel is selected do nothing
        elif SQL_step == 3:
            return

    # link this valve ID to the commodity property
    def AddComd(self, VlvID=None):
        item = self.dvc.GetSelection()
        if VlvID is None:
            rowGS = self.model.GetRow(item)
            VlvID = self.data[rowGS][self.ID_col]
        # check to make sure valve is not already linked to commodity property
        ChkQry = ('SELECT * FROM ' + self.PrtyTbl +
                  ' WHERE Commodity_Property_ID = ' + str(self.ComdPrtyID)
                  + ' AND ' + self.tblID + ' = ' + str(VlvID))
        if Dbase().Dsqldata(ChkQry) == []:
            # get the next available autoincremented ID
            # value for the valve property table

            if self.data == []:
                IDE = 1
            else:
                IDE = cursr.execute(
                    "SELECT MAX(" + self.PrtyID + ") FROM "
                    + self.PrtyTbl).fetchone()[0]+1

            qry = ('INSERT INTO ' + self.PrtyTbl + ' VALUES (' + str(IDE) +
                   ',' + str(VlvID) + ',' + str(self.ComdPrtyID) + ')')
            Dbase().TblEdit(qry)

            Valves = ('SELECT ' + self.tblID + ' FROM ' + self.PrtyTbl +
                      ' WHERE Commodity_Property_ID = ' + str(self.ComdPrtyID))
            VlvIDs = Dbase().Dsqldata(Valves)
            ValveIDs = tuple([i[0] for i in VlvIDs])
            if len(ValveIDs) == 1:
                self.vlvstrg = '(' + str(ValveIDs[0]) + ')'
            else:
                self.vlvstrg = str(ValveIDs)

            if self.DsqlVlv.find('WHERE') == -1:
                self.DsqlVlv = self.DsqlVlv + ' WHERE v.ID IN ' + self.vlvstrg
            else:
                self.DsqlVlv = self.DsqlVlv[:self.DsqlVlv.find('WHERE')]
                self.DsqlVlv = self.DsqlVlv + ' WHERE v.ID IN ' + self.vlvstrg

            self.data = Dbase().Dsqldata(self.DsqlVlv)
            self.lbl_nodata.SetLabel('   ')
        else:
            txtmsg = ('''Operation could not be completed,\n
            valve is already linked to this commodity property''')
            wx.MessageBox(txtmsg, 'Warning', wx.OK | wx.ICON_WARNING)
            return

    def RmvComd(self):
        item = self.dvc.GetSelection()
        rowGS = self.model.GetRow(item)
        Dbase().TblDelete(self.PrtyTbl, str(self.data[rowGS][self.ID_col]),
                          self.PrtyID)

    def ValData(self):
        datamissing = False
        missdata = 'The following data is missing:\n'

        # first check that all the combo boxes are completed
        for i, x in enumerate(self.cmbctrls):
            cmbvalue = str(self.cmbctrls[i].GetValue())
            if cmbvalue == '':
                missdata = missdata + self.cmb_range[i] + '\n'
                datamissing = True
        if datamissing is True:
            wx.MessageBox(missdata, 'Data Error',
                          wx.OK | wx.ICON_INFORMATION)
            return False

        # if both diameters are complete compare them and confirm Max > Min
        if (self.cmbctrls[1].GetValue() != '' or
                self.cmbctrls[2].GetValue() != ''):
            test = (eval(self.cmbctrls[2].GetValue().replace('"', '')
                    .replace('-', '+')) -
                    eval(self.cmbctrls[1].GetValue().replace('"', '')
                    .replace('-', '+')))
            if test <= 0:
                wx.MessageBox(
                    'Minimum Diameter must be\nless than Maximum Diameter.',
                    'Diameter Error', wx.OK | wx.ICON_INFORMATION)
                return False

        return True

    def OnSelect(self, evt):
        self.cmbSelected(evt.GetEventObject())

    def OnRestoreCmbs(self, evt):
        self.RestoreCmbs()

    def OnSearch(self, evt):

        if self.ValData() is True:
            field = Dbase().Dcolinfo(self.tblname)[1][1]
            ShQuery = ('SELECT * FROM ' + self.tblname + ' WHERE ' +
                       field + ' = "' + self.text1.GetValue() + '"')
            existing = Dbase().Search(ShQuery)

            if len(self.cmb_range) == self.text1.GetValue().count('.'):
                self.search = True

            if existing is not False:
                self.b2.SetLabel("Existing Spec")
                self.b3.Enable()
            else:
                self.b2.SetLabel("Add/Update\nValve")
                self.b2.Enable()
                self.b3.Disable()

            self.search = False

    def OnDeleteRow(self, evt):
        item = self.dvc.GetSelection()
        row = self.model.GetRow(item)
        try:
            if self.ComdPrtyID is None:
                self.model.DeleteRow(row, self.pkcol_name)
            elif self.b6.GetLabel() == "Show All\nValves":
                DeQuery = ("DELETE FROM " + self.PrtyTbl +
                           " WHERE Commodity_Property_ID = " +
                           str(self.ComdPrtyID) + " AND " +
                           self.tblID + " = " +
                           str(self.data[row][self.ID_col]))
                Dbase().PrptyDelete(DeQuery)

                Valves = ('SELECT ' + self.tblID + ' FROM ' + self.PrtyTbl
                          + ' WHERE Commodity_Property_ID = '
                          + str(self.ComdPrtyID))
                VlvIDs = Dbase().Dsqldata(Valves)
                ValveIDs = tuple([i[0] for i in VlvIDs])

                if len(ValveIDs) == 1:
                    self.vlvstrg = '(' + str(ValveIDs[0]) + ')'
                else:
                    self.vlvstrg = str(ValveIDs)

                if self.DsqlVlv.find('WHERE') == -1:
                    self.DsqlVlv = (self.DsqlVlv + ' WHERE v.ID IN '
                                    + self.vlvstrg)
                else:
                    self.DsqlVlv = self.DsqlVlv[:self.DsqlVlv.find('WHERE')]
                    self.DsqlVlv = (self.DsqlVlv + ' WHERE v.ID IN '
                                    + self.vlvstrg)

                self.data = Dbase().Dsqldata(self.DsqlVlv)
                if self.data == []:
                    self.lbl_nodata.SetLabel(
                        'The ' + self.frmtitle +
                        ' have not been setup for this Commodity Property')
                self.model = DataMods(self.tblname, self.data)
                self.dvc.AssociateModel(self.model)
                self.dvc.Refresh

        except sqlite3.IntegrityError:
            wx.MessageBox(
                "This " + self.pkcol_name +
                " is associated with other tables and cannot be deleted!",
                "Cannot Delete",
                wx.OK | wx.ICON_INFORMATION)

        self.b2.Enable()
        self.b3.Disable()
        self.RestoreCmbs()

    def OnAddComd(self, evt):
        if self.b6.GetLabel() == 'Show All\nValves':
            self.DsqlVlv = self.DsqlVlv[:self.DsqlVlv.find('WHERE')]
            self.data = Dbase().Dsqldata(self.DsqlVlv)
            self.model = DataMods(self.tblname, self.data)
            self.dvc.AssociateModel(self.model)
            self.dvc.Refresh
            self.b6.Enable(False)
            self.b6.SetLabel("Add Valve\nto Commodity")
        else:
            self.AddComd()
            self.RestoreCmbs()
            self.b3.Disable()
            self.b6.Enable(True)
            self.b6.SetLabel("Show All\nValves")

    def OnAddRow(self, evt):
        SQL_step = 3
        # if this is a not commodity related change
        if self.ComdPrtyID is None:
            # Make a selection as to whether the record is to be a new or
            # an update valve
            choices = ['Save this as a new valve code?',
                       'Update the valve code with this data?']
            # use a SingleChioce dialog to determine if data is
            # to be a new record or edited record
            SQL_Dialog = wx.SingleChoiceDialog(
                self, '''The valve record has changed do you want\n to save as
                 new record or update existing?''', 'Valve Record Has Changed',
                choices, style=wx.CHOICEDLG_STYLE)

            if SQL_Dialog.ShowModal() == wx.ID_OK:
                SQL_step = SQL_Dialog.GetSelection()
            SQL_Dialog.Destroy()
            self.AddRow(SQL_step)
            self.data = Dbase().Dsqldata(self.DsqlVlv)

        else:  # this is a commodity related valve change
            choices = ['Save this as a new valve code for this commodity?',
                       'Update this valve code for this commodity?']
            # use a SingleChioce dialog to determine if data is new
            # record or edited record
            SQL_Dialog = wx.SingleChoiceDialog(
                self, '''The valve record has changed do you want\n to save as
                new record or update existing?''', 'Valve Record Has Changed',
                choices, style=wx.CHOICEDLG_STYLE)
            if SQL_Dialog.ShowModal() == wx.ID_OK:
                SQL_step = SQL_Dialog.GetSelection()
            SQL_Dialog.Destroy()

            # save the modified valve as a replacement valve to the commodity
            if SQL_step == 1:
                # removes the existing valve ID from the valve property table
                self.RmvComd()
            # always save as a new valve code if commodity
            # property is specified
            maxID = self.AddRow(0)
            self.AddComd(maxID)

        self.model = DataMods(self.tblname, self.data)
        self.dvc.AssociateModel(self.model)
        self.dvc.Refresh

    def OnGridSelect(self, evt):
        item = self.dvc.GetSelection()
        rowGS = self.model.GetRow(item)
        self.text1.ChangeValue(self.data[rowGS][1])
        txtparts_values = self.text1.GetValue().split('.')
        txtparts_keys = list(range(len(txtparts_values)))
        self.txtparts = dict(zip(txtparts_keys, txtparts_values))

        n = 2
        for cmbbox in self.cmbctrls:
            cmbbox.Clear()
            cmbbox.ChangeValue(str(self.data[rowGS][n]))
            n += 1
        self.notes.ChangeValue(self.data[rowGS][n])
        self.grid_select = True

        cmbvalue = str(self.cmbctrls[4].GetValue())

        for m in self.a:  # selection of material related combo boxes only
            self.cmbctrls[m].Enable()
            tbl = self.hints_tbls[self.cmb_range[m]]

            fldstrg = ''
            cmbinfo = Dbase().Dcolinfo(tbl)
            for set in cmbinfo:
                fldstrg = fldstrg + tbl + '.' + set[1] + ','

            fldstrg = fldstrg[:-1]

            query = ('SELECT ' + fldstrg + ' FROM ' + tbl +
                     ' INNER JOIN MaterialType ON ')
            query = (query + tbl +
                     '.Material_Type = MaterialType.Type_Designation WHERE ')
            query = query + "MaterialType.Material_Type ='" + cmbvalue + "'"

            self.cmbctrls[m].SetPopupControl(ListCtrlComboPopup(
                tbl, query, cmbvalue, showcol=2))

        if self.ComdPrtyID is not None:
            self.b6.Enable()

        self.b3.Enable()
        self.b2.Disable()

    def OnClose(self, evt):
        # add these 2 lines as part of child parent form
        self.GetParent().Enable(True)
        self.__eventLoop.Exit()
        self.Destroy()


class BldFtgs(wx.Frame):
    '''Routine to build form and populate grid'''
    def __init__(self, parent, tblname, ComdPrtyID=None):

        ''' need to pass the;
        CommmodityProperties.Pipe_Material_Code
        (PipeMaterialSpec.Material_Spec_ID) for text2,
        .Commodity_Code (CommoditCodes.Commodity_Code) for text1 and
        .End_Connection (CommodityEnds.ID) for text3 into this class
        CommmodityProperties.CommodityPropertyID is needed
        to look up each attachement ID
        PipeSpecification.Flange_ID, .Union_ID, .OLet_ID, .OrificeFlange_ID'''

        self.parent = parent   # add for child parent form
        self.ComEnds = ''
        self.ComCode = ''
        self.PipeMtrSpec = ''

        # commodity property ID linked to pipespecification
        self.ComdPrtyID = ComdPrtyID

        if self.ComdPrtyID is not None:
            query = (
                '''SELECT Commodity_Code,Pipe_Material_Code,End_Connection,
                Pipe_Code FROM CommodityProperties
                 WHERE CommodityPropertyID = '''
                + str(self.ComdPrtyID))

            dataset = Dbase().Dsqldata(query)[0]
            self.PipeMtrSpec = dataset[1]
            self.ComCode = dataset[0]
            self.ComEnds = dataset[2]
            self.PipeCode = dataset[3]

        self.data = []
        self.lctrls = []
        self.rec_num = 0
        self.columnames = []
        self.addbtns = []
        self.tblname = tblname
        self.restore = False
        self.bxdict = {}
        self.DsqlFtg = ''
        StartQry = None

        self.txtparts = {}

        if self.tblname.find("_") != -1:
            self.frmtitle = (self.tblname.replace("_", " "))
        else:
            self.frmtitle = (
                ' '.join(re.findall('([A-Z][a-z]*)', self.tblname)))

        self.txtparts[0] = str(self.tblname[0]) + 'C'

        if self.ComdPrtyID is not None:
            query = ('SELECT ' + self.tblname[:-1] +
                     '''_ID FROM PipeSpecification WHERE
                      Commodity_Property_ID = '''
                     + str(self.ComdPrtyID))
            chk = Dbase().Dsqldata(query)
            if chk != []:
                StartQry = chk[0][0]
                if StartQry is not None:
                    self.DsqlFtg = (
                        'SELECT * FROM ' + self.tblname + ' WHERE '
                        + self.tblname[:-1] + 'ID = ' + str(StartQry)
                    )

                    self.data = Dbase().Dsqldata(self.DsqlFtg)
        else:
            self.DsqlFtg = 'SELECT * FROM ' + self.tblname
            self.data = Dbase().Dsqldata(self.DsqlFtg)

        # specify which listbox column to display in the combobox
        self.showcol = int

        # set up the table column names, width and if column can be
        # edited ie primary autoincrement
        tblinfo = []

        tblinfo = Dbase(self.tblname).Fld_Size_Type()
        self.ID_col = tblinfo[1]
        self.colms = tblinfo[3]
        self.pkcol_name = tblinfo[0]
        self.autoincrement = tblinfo[2]

        if self.tblname == 'Fittings':  # info for pipe fittings;elbows,tees
            self.lbl_txt = ['Min. Pipe\nDiameter', 'Max. Pipe\nDiameter',
                            'End\nConnections', 'Material', 'Schedule\nClass']
            self.cmb_tbls = ['Pipe_OD', 'Pipe_OD', 'EndConnects',
                             'Material', 'Rating']
            self.wrn_lbl = (
                'No Fittings have been setup for this Commodity Property'
                )
            siz = (780, 580)
        elif self.tblname == 'GrooveClamps':
            self.lbl_txt = ['Min. Pipe\nDiameter', 'Max. Pipe\nDiameter',
                            'Schedule', 'Groove Type', 'Style',
                            'Seal\nMaterial', 'Clamp\nMaterial']
            self.cmb_tbls = ['Pipe_OD', 'Pipe_OD', 'PipeSchedule',
                             'ClampGroove', 'GrooveClampVendor',
                             'GasketSealMaterial', 'ForgedMaterial']
            self.wrn_lbl = '''No Grooved Style Clamps have been setup
                            for this Commodity Property'''
            siz = (1230, 520)
        else:  # info for forgings; unions, flanges etc
            self.lbl_txt = ['Min. Pipe\nDiameter', 'Max. Pipe\nDiameter',
                            'Flange\nStyle', 'Flange\nFace', 'Material',
                            'Schedule\nClass']
            self.cmb_tbls = ['Pipe_OD', 'Pipe_OD', 'FlangeStyle', 'FlangeFace',
                             'ForgedMaterial', 'ANSI_Rating']
            self.wrn_lbl = (
                'No Flanges have been setup for this Commodity Property'
                )
            siz = (930, 580)
            if self.tblname == 'Unions':
                self.lbl_txt[2] = 'End\nConnections'
                self.lbl_txt[3] = 'Seat\nMaterial'
                self.cmb_tbls[2] = 'EndConnects'
                self.cmb_tbls[3] = 'SeatMaterial'
                self.cmb_tbls[5] = 'ForgeClass'
                self.wrn_lbl = (
                    'Not Unions have been setup for this Commodity Property'
                )
                siz = (930, 520)
            elif self.tblname == 'OLets':
                self.lbl_txt[0] = 'Min. Run\nDiameter'
                self.lbl_txt[1] = 'Max. Run\nDiameter'
                self.lbl_txt[2] = 'Minimum Delta\nBranch to Run'
                self.lbl_txt[3] = 'OLet\nStyle'
                self.cmb_tbls[2] = 'Pipe_OD'
                self.cmb_tbls[3] = 'OLetStyle'
                self.cmb_tbls[5] = 'OLetWt'
                self.wrn_lbl = (
                    'Not O-Lets have been setup for this Commodity Property'
                    )
                siz = (930, 650)

        self.columnames = Dbase(self.tblname).ColNames()[3:]

        self.Num_Cols = len(self.lbl_txt)
        self.Num_Rows = len(self.columnames)//self.Num_Cols

        super(BldFtgs, self).__init__(parent, title=self.frmtitle, size=siz)

        self.Bind(wx.EVT_CLOSE, self.OnClose)

        # set the Sizer property (same as SetSizer)
        self.Sizer = wx.BoxSizer(wx.VERTICAL)

        # set up first row of combo boxesand label
        self.searchsizer = wx.BoxSizer(wx.HORIZONTAL)
        self.codesizer = wx.BoxSizer(wx.HORIZONTAL)

        cmbsizers = []
        for row in range(0, self.Num_Rows):
            cmbsizer = wx.BoxSizer(wx.HORIZONTAL)
            cmbsizers.append(cmbsizer)

        font = wx.Font(10, wx.DEFAULT, wx.NORMAL, wx.NORMAL)
        font1 = wx.Font(12, wx.DEFAULT, wx.NORMAL, wx.FONTWEIGHT_BOLD)
        font2 = wx.Font(12, wx.DEFAULT, wx.NORMAL, wx.FONTWEIGHT_BOLD)

        txt_nodata = (' The ' + self.frmtitle +
                      ' have not been setup for this Commodity Property')
        self.lbl_nodata = wx.StaticText(self, -1, label=txt_nodata,
                                        size=(650, 30), style=wx.LEFT)
        self.lbl_nodata.SetForegroundColour((255, 0, 0))
        self.lbl_nodata.SetFont(font2)
        self.lbl_nodata.SetLabel('   ')

        self.warningsizer = wx.BoxSizer(wx.HORIZONTAL)
        self.warningsizer.Add(self.lbl_nodata, 0, wx.ALIGN_CENTER)

        if self.ComdPrtyID is not None:
            if StartQry is None:
                self.lbl_nodata.SetLabel(txt_nodata)
            query = ('SELECT * FROM PipeMaterialSpec WHERE Material_Spec_ID = '
                     + str(self.PipeMtrSpec))
            self.MtrSpc = Dbase().Dsqldata(query)[0][1]
            self.spec_ID = Dbase().Dsqldata(query)[0][0]

            if self.PipeCode is not None:
                txt = self.PipeCode
            else:
                txt = self.ComCode + ' - ' + self.MtrSpc

            self.text1 = wx.TextCtrl(self, size=(100, 33), value=txt,
                                     style=wx.TE_READONLY | wx.TE_CENTER)

            self.text1.SetForegroundColour((255, 0, 0))
            self.text1.SetFont(font)
            self.searchsizer.Add(self.text1, 0)

            query = (
                'SELECT * FROM CommodityEnds WHERE ID = ' + str(self.ComEnds))
            End = Dbase().Dsqldata(query)[0][1]
            self.text3 = wx.TextCtrl(
                self, size=(300, 33), value=End,
                style=wx.TE_READONLY | wx.TE_CENTER)
            self.text3.SetForegroundColour((255, 0, 0))
            self.text3.SetFont(font)

            self.searchsizer.Add(self.text3, 0)

        else:
            self.searchsizer.Add((20, 20))

        if self.tblname == 'Fittings':
            strg = ()
            if self.ComEnds == 1:
                strg = (1, 2, 3, 4, 5, 6)
            elif self.ComEnds == 2:
                strg = (1, 3, 4, 6)
            elif self.ComEnds == 3:
                strg = (1, 3, 4)
            elif self.ComEnds == 4:
                strg = (1, 3)
            elif self.ComEnds == 5:
                strg = (7)
            elif self.ComEnds == 6:
                strg = (8)

            queryEnds = 'SELECT * FROM EndConnects WHERE EndID IN ' + str(strg)
            txtsiz = (450, 33)
        else:
            txtsiz = (650, 33)

        self.text4 = wx.TextCtrl(
            self, size=txtsiz,
            value=str(self.frmtitle) + ' Code',
            style=wx.TE_READONLY | wx.TE_CENTER
                )
        self.text4.SetForegroundColour((255, 0, 0))
        # add a button to call main form to search combo list data
        self.text4.SetFont(font)
        self.lblCode = wx.StaticText(self, -1, label='CODE', style=wx.CENTER)

        self.codesizer.Add(self.lblCode, 0, wx.ALIGN_CENTER_VERTICAL)
        self.codesizer.Add((10, 10))
        self.codesizer.Add(self.text4, 0, wx.ALIGN_CENTER)

        self.lblsizer = wx.BoxSizer(wx.HORIZONTAL)
        xply = 1

        for txt in self.lbl_txt:
            self.lbl = wx.StaticText(self, -1, label=txt, style=wx.CENTER)

            self.addbtn = wx.Button(self, label='+', size=(35, -1))
            self.addbtn.SetForegroundColour((255, 0, 0))
            self.addbtn.SetFont(font1)
            self.Bind(wx.EVT_BUTTON, self.OnAdd1, self.addbtn)
            self.addbtns.append(self.addbtn)

            if xply < 1.18:
                xply += .28
            self.lblsizer.Add(self.lbl, 0, wx.ALL | wx.ALIGN_BOTTOM, 10)
            self.lblsizer.Add(self.addbtn, 0, wx.ALIGN_CENTER)
            self.lblsizer.Add((int(26*xply), 25))

        if self.tblname == 'GrooveClamps':
            self.lblVndr = wx.StaticText(self, -1, label='Vendor',
                                         style=wx.CENTER)
            self.lblsizer.Add(self.lblVndr, 0, wx.RIGHT | wx.ALIGN_CENTER, 85)

        # Start the generation of the required combo boxes
        # list of combobox objects
        self.cmbctrls = []
        self.txtVndrs = []

        # counter to track which comboboxes need
        # activation after material type is selected
        n = 0
        # cmb_range is the list of the combo box
        for h in range(0, self.Num_Rows):
            for cmb_tbl in self.cmb_tbls:
                cmbbox = wx.ComboCtrl(self, size=(150, -1),
                                      style=wx.CB_READONLY)
                self.Bind(wx.EVT_COMBOBOX_CLOSEUP, self.OnCmbClose, cmbbox)
                self.Bind(wx.EVT_COMBOBOX_DROPDOWN, self.OnCmbOpen, cmbbox)

                if self.tblname == 'Fittings':
                    # can only load lists for the boxes for PipeOD and Ends
                    # material type and rating type is not know yet,
                    # set in CmbSelected
                    if cmb_tbl == 'Material' or cmb_tbl == 'Rating':
                        cmbbox.Disable()
                        self.showcol = 1
                    elif cmb_tbl == 'EndConnects':
                        self.showcol = 1
                        cmbbox.SetPopupControl(
                            ListCtrlComboPopup(cmb_tbl, queryEnds,
                                               showcol=self.showcol,
                                               lctrls=self.lctrls))
                    else:
                        self.showcol = 1
                        cmbbox.SetPopupControl(
                            ListCtrlComboPopup(cmb_tbl, showcol=self.showcol,
                                               lctrls=self.lctrls))

                    # disable the boxes below row 1
                    if h >= 1:
                        cmbbox.Disable()
                    self.cmbctrls.append(cmbbox)

                    # and the cmbbox to the cmbsizer
                    cmbsizers[h].Add(cmbbox, 0, wx.ALIGN_LEFT, 5)

                    n += 1
                else:
                    if cmb_tbl == 'ForgedMaterial':
                        # get the data for the material combo box
                        query = ('SELECT * FROM ' + cmb_tbl +
                                 ' WHERE Material_Type = "' +
                                 self.MtrSpc[1] +
                                 '" AND Material_Grade = "' +
                                 self.MtrSpc[2] + '"')
                        self.showcol = 1
                        cmbbox.SetPopupControl(
                            ListCtrlComboPopup(cmb_tbl, query,
                                               showcol=self.showcol,
                                               lctrls=self.lctrls))

                    elif cmb_tbl == 'SeatMaterial':
                        query = ('SELECT * FROM ' + cmb_tbl +
                                 ' WHERE Material_Type = "' +
                                 self.MtrSpc[1]) + '"'
                        self.showcol = 2
                        cmbbox.SetPopupControl(
                            ListCtrlComboPopup(cmb_tbl, query,
                                               showcol=self.showcol,
                                               lctrls=self.lctrls))

                    elif cmb_tbl == 'GasketSealMaterial':
                        query = ('SELECT * FROM ' + cmb_tbl +
                                 ' WHERE Material_Type = "' +
                                 self.MtrSpc[1] + '"')
                        self.showcol = 1
                        cmbbox.SetPopupControl(
                            ListCtrlComboPopup(cmb_tbl, query,
                                               showcol=self.showcol,
                                               lctrls=self.lctrls))

                    elif cmb_tbl in ['GrooveClampVendor', "PipeSchedule"]:
                        query = 'SELECT * FROM ' + cmb_tbl
                        self.showcol = 2
                        cmbbox.SetPopupControl(
                            ListCtrlComboPopup(cmb_tbl, query,
                                               showcol=self.showcol,
                                               lctrls=self.lctrls))
                    # set up the fitting end connects to correspond
                    # with the ends stipulated for the commodity property
                    elif cmb_tbl in ['EndConnects', 'OLetStyle',
                                     'FlangeStyle']:
                        endQ = 'SELECT * FROM ' + cmb_tbl
                        style1 = ''
                        style2 = ''
                        if str(self.ComEnds) == '1':
                            # Unions
                            if cmb_tbl == 'EndConnects':
                                style1 = '(2,3,4,5)'
                        if str(self.ComEnds) == '2':
                            if cmb_tbl == 'EndConnects':
                                style1 = '(1,3,4,6)'
                            # OLets
                            if cmb_tbl == 'OLetStyle':
                                style2 = '(1,2,4,5,7)'
                            # Orifice Flange and Flange
                            if cmb_tbl == 'FlangeStyle':
                                style2 = '(1,2,3,4,5,7,8,9)'
                        if str(self.ComEnds) == '3':
                            if cmb_tbl == 'EndConnects':
                                style1 = '(1,3,4,6)'
                            if cmb_tbl == 'FlangeStyle':
                                style2 = '(1,3,4,5,7,8,9)'
                            if cmb_tbl == 'OLetStyle':
                                style2 = '(1,2,4,5,7)'
                        if str(self.ComEnds) == '4':
                            if cmb_tbl == 'EndConnects':
                                style1 = '(1,3,6)'
                            if cmb_tbl == 'FlangeStyle':
                                style2 = '(1,3,4,5,7,8,9)'
                            if cmb_tbl == 'OLetStyle':
                                style2 = '(1,2,3,4,5,7)'
                        if str(self.ComEnds) == '5':
                            if cmb_tbl == 'EndConnects':
                                style1 = '(6)'
                            if cmb_tbl == 'FlangeStyle':
                                style2 = '(7,8)'
                        if str(self.ComEnds) == '6':
                            if cmb_tbl == 'EndConnects':
                                style1 = '(6)'
                            if cmb_tbl == 'OLetStyle':
                                style2 = '(2,4,5,7)'
                            if cmb_tbl == 'FlangeStyle':
                                style2 = '(1,5,7,8)'
                        strg = ''

                        if style1 != '':
                            strg = strg + " WHERE EndID IN " + style1
                        elif style2 != '':
                            strg = strg + " WHERE StyleID IN " + style2
                        endQ = endQ + strg

                        self.showcol = 1
                        cmbbox.SetPopupControl(
                            ListCtrlComboPopup(cmb_tbl, endQ,
                                               showcol=self.showcol,
                                               lctrls=self.lctrls))

                    # all other combo boxes get generic set up
                    else:
                        self.showcol = 1
                        cmbbox.SetPopupControl(
                            ListCtrlComboPopup(cmb_tbl, PupQuery='',
                                               showcol=self.showcol,
                                               lctrls=self.lctrls))

                    # disable the boxes below row 1
                    if h >= 1:
                        cmbbox.Disable()
                    self.cmbctrls.append(cmbbox)
                    # and the cmbbox to the cmbsizer
                    cmbsizers[h].Add(cmbbox, 0, wx.ALIGN_LEFT, 5)

                    n += 1

            if self.tblname == 'GrooveClamps':
                txtVndr = wx.TextCtrl(self, size=(150, 33), value='',
                                      style=wx.TE_READONLY | wx.TE_CENTER)
                self.txtVndrs.append(txtVndr)
                cmbsizers[h].Add(txtVndr, 0, wx.ALIGN_LEFT, 5)

        # Add some buttons
        self.b5 = wx.Button(self, label="Clear\nBoxes")
        self.Bind(wx.EVT_BUTTON, self.OnRestoreCmbs, self.b5)

        self.b1 = wx.Button(self, label="Print\nReport")
        self.Bind(wx.EVT_BUTTON, self.PrintFile, self.b1)

        self.b2 = wx.Button(self, label="Add/Update\nto " + self.frmtitle)
        self.b2.Disable()
        self.Bind(wx.EVT_BUTTON, self.OnAddRec, self.b2)

        self.b3 = wx.Button(self, label="Delete\nSpecification")
        self.Bind(wx.EVT_BUTTON, self.OnDeleteRec, self.b3)

        self.b4 = wx.Button(self, label="Exit")
        self.Bind(wx.EVT_BUTTON, self.OnClose, self.b4)

        # add a button box and place the buttons
        self.btnbox = wx.BoxSizer(wx.HORIZONTAL)
        self.btnbox.Add(self.b5, 0, wx.ALIGN_CENTER | wx.RIGHT, 35)
        self.btnbox.Add(self.b1, 0, wx.ALIGN_CENTER | wx.LEFT, 5)
        self.btnbox.Add(self.b2, 0, wx.ALIGN_CENTER | wx.LEFT, 5)
        if self.ComdPrtyID is not None:
            self.b6 = wx.Button(self, size=(120, 45),
                                label="Show All\n" + self.frmtitle)
            self.b3.SetLabel('Delete From\nCommodity')
            if StartQry is None:
                self.b3.Disable()
            self.Bind(wx.EVT_BUTTON, self.OnAddComd, self.b6)
            self.btnbox.Add(self.b6, 0, wx.ALL | wx.ALIGN_CENTER, 5)
        self.btnbox.Add(self.b3, 0, wx.ALIGN_CENTER | wx.LEFT, 5)
        self.btnbox.Add((30, 10))
        self.btnbox.Add(self.b4, 0, wx.ALIGN_CENTER | wx.LEFT, 5)

        self.fst = wx.Button(self, label='<<')
        self.lst = wx.Button(self, label='>>')
        self.nxt = wx.Button(self, label='>')
        self.pre = wx.Button(self, label='<')
        self.fst.Bind(wx.EVT_BUTTON, self.OnMovefst)
        self.lst.Bind(wx.EVT_BUTTON, self.OnMovelst)
        self.nxt.Bind(wx.EVT_BUTTON, self.OnMovenxt)
        self.pre.Bind(wx.EVT_BUTTON, self.OnMovepre)

        self.navbox = wx.BoxSizer(wx.HORIZONTAL)
        self.navbox.Add(self.fst, 0, wx.ALL, 5)
        self.navbox.Add(self.pre, 0, wx.ALL, 5)
        self.navbox.Add(self.nxt, 0, wx.ALL, 5)
        self.navbox.Add(self.lst, 0, wx.ALL, 5)

        self.numbox = wx.BoxSizer(wx.HORIZONTAL)
        self.recnum1 = wx.StaticText(self, label='Record ',
                                     style=wx.ALIGN_LEFT)
        self.recnum1.SetForegroundColour((255, 0, 0))

        self.recnum2 = wx.StaticText(self, label=str(self.rec_num+1),
                                     style=wx.ALIGN_LEFT)
        self.recnum2.SetForegroundColour((255, 0, 0))
        self.recnum3 = wx.StaticText(self, label='/ '+str(len(self.data)),
                                     style=wx.ALIGN_LEFT)
        self.recnum3.SetForegroundColour((255, 0, 0))
        self.numbox.Add(self.recnum1, 0, wx.ALL, 5)
        self.numbox.Add(self.recnum2, 0, wx.ALL, 5)
        self.numbox.Add(self.recnum3, 0, wx.ALL, 5)

        # add static label to explain how to add / edit data
        self.editlbl = wx.StaticText(self, -1, style=wx.ALIGN_LEFT)
        txt1 = '\tTo edit data select the item from the corresponding drop '
        txt2 = 'downlist. If the final Code is a\n\tduplicate you will not be '
        txt3 = 'able to save the information. To add a new record click "Clear'
        txt4 = ' Boxes"\n\tand complete the data selection, if it does not '
        txt5 = ' already exist you will be able to save the spec.\n\tYou can '
        txt6 = ' add items to the drop down list by clicking the RED + '
        txt7 = 'beside each column label.'
        txt = txt1 + txt2 + txt3 + txt4 + txt5 + txt6 + txt7

        self.editlbl.SetLabel(txt)
        self.editlbl.SetForegroundColour((255, 0, 0))

        self.Sizer.Add(self.warningsizer, 0, wx.CENTER | wx.TOP, 5)
        self.Sizer.Add(self.searchsizer, 0, wx.ALL | wx.ALIGN_CENTER, 5)
        self.Sizer.Add(self.codesizer, 0, wx.ALL | wx.ALIGN_CENTER, 5)
        self.Sizer.Add(self.lblsizer, 0, wx.ALIGN_CENTER)
        for cmbsizer in cmbsizers:
            self.Sizer.Add(cmbsizer, 0, wx.ALIGN_CENTER)

        self.Sizer.Add((25, 10))
        self.Sizer.Add(self.btnbox, 0, wx.ALIGN_CENTER, 5)
        self.Sizer.Add((25, 10))
        self.Sizer.Add(self.editlbl, 0, wx.ALIGN_CENTER, 5)
        self.Sizer.Add((25, 10))
        self.Sizer.Add(self.navbox, 0, wx.ALIGN_CENTER, 5)
        self.Sizer.Add(self.numbox, 0, wx.ALIGN_CENTER, 5)

        self.btnDict = {self.addbtns[i]: i for i in range(len(self.addbtns))}

        self.FillScreen()

        # add these 5 following lines for the child parent form
        self.CenterOnParent()
        self.GetParent().Enable(False)
        self.Show(True)
        self.__eventLoop = wx.GUIEventLoop()
        self.__eventLoop.Run()

    def OnMovefst(self, evt):
        self.rec_num = 0
        self.FillScreen()

    def OnMovelst(self, evt):
        if len(self.data) == 0:
            return
        self.rec_num = len(self.data)-1
        self.FillScreen()

    def OnMovenxt(self, evt):
        if len(self.data) == 0:
            return
        self.rec_num += 1
        if self.rec_num == len(self.data):
            self.rec_num = 0
        self.FillScreen()

    def OnMovepre(self, evt):
        if len(self.data) == 0:
            return
        self.rec_num -= 1
        if self.rec_num < 0:
            self.rec_num = len(self.data)-1
        self.FillScreen()

    def FillScreen(self):
        # all the IDs for the various tables making up the package
        if len(self.data) == 0:
            self.recnum2.SetLabel(str(self.rec_num))
            return
        else:
            recrd = self.data[self.rec_num]
            minum = self.Num_Rows * self.Num_Cols
            if self.tblname == 'GrooveClamps':
                for i in range(len(self.txtVndrs)):
                    self.txtVndrs[i].ChangeValue('')

        if self.tblname == 'Fittings':
            # numbers for the Pipe_OD combos
            for i in [0, 1, 5, 6, 10, 11, 15, 16]:
                # select the next combo box which matches its position
                # & ID numbers for each record
                tbl_ID = str(recrd[i+3])
                if tbl_ID != 'None':
                    tbl_ID_nam = Dbase().Dcolinfo(self.cmb_tbls[0])[0][1]
                    query = ('SELECT * FROM ' + self.cmb_tbls[0] +
                             ' WHERE ' + tbl_ID_nam + ' = ' + tbl_ID)
                    self.cmbctrls[i].ChangeValue(
                        str(Dbase().Dsqldata(query)[0][1]))
                    self.cmbctrls[i].Enable()
                else:
                    self.cmbctrls[i].ChangeValue('')
                    self.cmbctrls[i].Disable()
                    if i < minum:
                        minum = i

            # determine which material table and rating table
            # to use bsed on end connection
            # numbers of the combo boxes for the end connections
            for i in [2, 7, 12, 17]:
                # select the next combo box which matches its position
                # & ID numbers for each record
                tbl_ID = str(recrd[i+3])
                if tbl_ID != 'None':
                    tbl_ID_nam = Dbase().Dcolinfo(self.cmb_tbls[2])[0][1]
                    query = ('SELECT * FROM ' + self.cmb_tbls[2] +
                             ' WHERE ' + tbl_ID_nam + ' = ' + tbl_ID)
                    self.cmbctrls[i].ChangeValue(
                        str(Dbase().Dsqldata(query)[0][1]))
                    self.cmbctrls[i].Enable()

                    if tbl_ID in ('2', '4', '5'):
                        rtg_tbl = 'ForgeClass'
                        mtr_tbl = 'ForgedMaterial'
                    elif tbl_ID == '1':
                        rtg_tbl = 'ANSI_Rating'
                        mtr_tbl = 'ForgedMaterial'
                    elif tbl_ID == '3':
                        rtg_tbl = 'PipeSchedule'
                        mtr_tbl = 'ButtWeldMaterial'
                    elif tbl_ID == '6':
                        rtg_tbl = 'PipeSchedule'
                        mtr_tbl = 'PipeMaterial'

                    self.bxdict[i] = (rtg_tbl, mtr_tbl)
                else:
                    self.cmbctrls[i].ChangeValue('')
                    self.cmbctrls[i].Disable()

            # specify the material table to
            # use based on the type of fitting end
            for i in [3, 8, 13, 18]:   # numbers for the material combo boxes
                # select the next combo box which matches its position
                # & ID numbers for each record
                tbl_ID = str(recrd[i+3])
                if tbl_ID != 'None':
                    mtr_tbl = self.bxdict[i-1][1]
                    tbl_ID_nam = Dbase().Dcolinfo(mtr_tbl)[0][1]
                    query = ('SELECT * FROM ' + mtr_tbl +
                             ' WHERE Material_Type = "' +
                             self.MtrSpc[1] +
                             '" AND Material_Grade = "' +
                             self.MtrSpc[2] + '"')
                    self.showcol = 1
                    self.cmbctrls[i].SetPopupControl(
                        ListCtrlComboPopup(mtr_tbl, query,
                                           showcol=self.showcol,
                                           lctrls=self.lctrls))
                    tbl_ID_nam = str(Dbase().Dcolinfo(mtr_tbl)[0][1])
                    query = ('SELECT * FROM ' + mtr_tbl + ' WHERE '
                             + tbl_ID_nam + ' = ' + tbl_ID)
                    self.cmbctrls[i].ChangeValue(
                        str(Dbase().Dsqldata(query)[0][1]))
                    self.cmbctrls[i].Enable()
                else:
                    self.cmbctrls[i].ChangeValue('')
                    self.cmbctrls[i].Disable()

            # specify the rating table to use based on the type of fitting end
            # numbers of the combo boxes for the fitting class
            for i in [4, 9, 14, 19]:
                # select the next combo box which matches its position
                # & ID numbers for each record
                tbl_ID = str(recrd[i+3])
                if tbl_ID != 'None':
                    rtg_tbl = self.bxdict[i-2][0]
                    query = 'SELECT * FROM ' + rtg_tbl
                    self.showcol = 1
                    if rtg_tbl == 'PipeSchedule':
                        self.showcol = 2
                    tbl_ID_nam = Dbase().Dcolinfo(rtg_tbl)[0][1]
                    self.cmbctrls[i].SetPopupControl(
                        ListCtrlComboPopup(rtg_tbl, query,
                                           showcol=self.showcol,
                                           lctrls=self.lctrls))
                    query = ('SELECT * FROM ' + rtg_tbl + ' WHERE '
                             + tbl_ID_nam + ' = ' + tbl_ID)
                    if rtg_tbl == 'PipeSchedule':
                        self.cmbctrls[i].ChangeValue(
                            Dbase().Dsqldata(query)[0][2])
                    else:
                        self.cmbctrls[i].ChangeValue(
                            str(Dbase().Dsqldata(query)[0][1]))

                    self.cmbctrls[i].Enable()
                else:
                    self.cmbctrls[i].ChangeValue('')
                    self.cmbctrls[i].Disable()

        else:
            # cycle through each column by row
            for m in range(0, self.Num_Rows * self.Num_Cols):
                # select the next combo box which matches its
                # position & ID numbers for each record
                tbl_ID = str(recrd[m+3])

                tbl_ID_nam = Dbase().Dcolinfo(
                    self.cmb_tbls[m % self.Num_Cols])[0][1]

                if tbl_ID != 'None':
                    query = ('SELECT * FROM ' +
                             self.cmb_tbls[m % self.Num_Cols]
                             + ' WHERE ' + tbl_ID_nam + ' = ' + tbl_ID)
                    if self.cmb_tbls[m % self.Num_Cols] \
                            in ['SeatMaterial', 'GrooveClampVendor',
                                'PipeSchedule']:
                        self.cmbctrls[m].ChangeValue(
                            str(Dbase().Dsqldata(query)[0][2]))
                        if self.tblname == 'GrooveClamps' and m in [4, 11]:
                            Vndr = Dbase().Dsqldata(query)[0][1]
                            self.txtVndrs[m//11].ChangeValue(Vndr)
                    else:
                        self.cmbctrls[m].ChangeValue(
                            str(Dbase().Dsqldata(query)[0][1]))

                    self.cmbctrls[m].Enable()
                else:
                    self.cmbctrls[m].ChangeValue('')
                    self.cmbctrls[m].Disable()
                    if m < minum:
                        minum = m

        if minum < self.Num_Cols * self.Num_Rows:
            for m in range(minum, minum+self.Num_Cols):
                self.cmbctrls[m].Enable()

        if recrd[1] == 'None':
            self.text4.ChangeValue(self.tblname[:-1] + ' Code')
        else:
            self.text4.ChangeValue(str(recrd[1]))

        tmp = self.text4.GetValue()[:2]
        lcd = self.text4.GetValue()[3:].split('.')
        if '0' in lcd:
            lcd = lcd[:(lcd.index('0')-lcd.index('0') % self.Num_Cols)]
        else:
            lcd = lcd[:len(lcd)-len(lcd) % self.Num_Cols]
        lcd.insert(0, tmp)
        self.txtparts = {i: lcd[i] for i in range(0, len(lcd))}

        self.recnum2.SetLabel(str(self.rec_num+1))

    def PrintFile(self, evt):
        import Report_Lvl4

        if self.data == []:
            NoData = wx.MessageDialog(
                None, 'No Data to Print', 'Error', wx.OK | wx.ICON_EXCLAMATION)
            NoData.ShowModal()
            return

        saveDialog = wx.FileDialog(self, message='Save Report as PDF.',
                                   wildcard='PDF (*.pdf)|*.pdf',
                                   style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT)

        if saveDialog.ShowModal() == wx.ID_CANCEL:
            filename = ''

        filename = saveDialog.GetPath()
        if filename.find(".pdf") == -1:
            filename = filename + '.pdf'

        saveDialog.Destroy()

        # call the function to get the data for the report,
        # data,column names,column widths
        rptdata, Colnames, Colwdths = self.ReportData()

        ttl = None
        if len(self.data) == 1:
            # specify a title for the report if table name
            # is not to be used
            if self.PipeCode is None:
                ttl = self.text1.GetValue()
            else:
                ttl = self.PipeCode
            ttl = ' for ' + ttl

        Report_Lvl4.Report(self.tblname, rptdata, Colnames,
                           Colwdths, filename, ttl).create_pdf()

    def NewTextStr(self):
        s = ''
        orderly = sorted(self.txtparts)
        for n in orderly:
            s = s + str(self.txtparts[n]) + '.'
        s = s[:-1]
        return s

    # this builds the Fitting Code text string
    def LblData(self, tbl_name, cmbvalue):
        query3 = ''

        if cmbvalue == '':
            return

        if tbl_name == 'ANSI_Rating':
            # this is used if the combobox shows a numberic value
            query3 = ('SELECT Rating_Designation FROM ' + tbl_name +
                      ' WHERE ANSI_Class = ' + cmbvalue)
        elif tbl_name == 'Pipe_OD':
            # this query is used when the combobox shows a string value
            query3 = ("SELECT PipeOD_ID FROM " + tbl_name +
                      " WHERE Pipe_OD = '" + cmbvalue + "'")
        elif tbl_name == 'EndConnects':
            query3 = ('SELECT EndID FROM ' + tbl_name +
                      " WHERE Connection = '" + cmbvalue + "'")
        elif tbl_name == 'ForgeClass':
            query3 = ('SELECT ClassID FROM ' + tbl_name +
                      ' WHERE Forged_Class = ' + cmbvalue)
        elif tbl_name == 'FlangeStyle':
            query3 = ('SELECT StyleID FROM ' + tbl_name +
                      " WHERE Style_Type = '" + cmbvalue + "'")
        elif tbl_name == 'FlangeFace':
            query3 = ('SELECT FaceID FROM ' + tbl_name +
                      " WHERE Face_Style = '" + cmbvalue + "'")
        elif tbl_name == 'ForgedMaterial':
            query3 = ('SELECT ID FROM ' + tbl_name +
                      " WHERE Forged_Material = '" + cmbvalue +
                      "' AND Material_Grade = '" + self.MtrSpc[2] +
                      "'")
        elif tbl_name == 'SeatMaterial':
            query3 = ('SELECT SeatMaterialID FROM ' + tbl_name +
                      " WHERE Seat_Material = '" + cmbvalue +
                      "' AND Material_Type = '" + self.MtrSpc[1] +
                      "'")
        elif tbl_name == 'OLetStyle':
            query3 = ('SELECT StyleID FROM ' + tbl_name +
                      " WHERE Style = '" + cmbvalue + "'")
        elif tbl_name == 'OLetWt':
            query3 = ('SELECT OLetWtID FROM ' + tbl_name +
                      " WHERE Weight = '" + cmbvalue + "'")
        elif tbl_name == 'ButtWeldMaterial':
            query3 = ('SELECT ID FROM ' + tbl_name +
                      " WHERE Butt_Weld_Material = '" + cmbvalue +
                      "' AND Material_Grade = '" + self.MtrSpc[2] +
                      "'")
        elif tbl_name == 'PipeMaterial':
            query3 = ('SELECT ID FROM ' + tbl_name + " WHERE Pipe_Material = '"
                      + cmbvalue + "' AND Material_Grade = '" +
                      self.MtrSpc[2] + "'")
        elif tbl_name == 'PipeSchedule':
            query3 = ('SELECT ID FROM ' + tbl_name + " WHERE Pipe_Schedule = '"
                      + cmbvalue + "'")
        elif tbl_name == 'ClampGroove':
            query3 = ('SELECT ClampGrooveID FROM ' + tbl_name +
                      " WHERE Groove = '" + cmbvalue + "'")
        elif tbl_name == 'GrooveClampVendor':
            query3 = ('SELECT VendorID FROM ' + tbl_name +
                      " WHERE Style = '" + cmbvalue + "'")
        elif tbl_name == 'GasketSealMaterial':
            query3 = ('SELECT SealID FROM ' + tbl_name +
                      " WHERE GasketSealmaterial = '" + cmbvalue +
                      "' AND Material_Type = '" + self.MtrSpc[1] +
                      "'")

        if query3 != '':
            lbldata = Dbase().Dsqldata(query3)
            lblstr = str(lbldata[0][0])

        return lblstr

    def RestoreCmbs(self):
        n = 0
        self.restore = True
        for cmbbox in self.cmbctrls:
            cmbbox.Clear()
            cmbbox.ChangeValue('')
            cmbbox.Disable()
            n += 1

        if self.tblname == 'Fittings':
            self.cmbctrls[0].Enable()
            self.cmbctrls[1].Enable()
            self.cmbctrls[2].Enable()
        else:
            for m in range(0, self.Num_Cols):
                self.cmbctrls[m].Enable()

        if self.tblname == 'GrooveClamps':
            for i in range(len(self.txtVndrs)):
                self.txtVndrs[i].ChangeValue('')

        temptxt = self.txtparts[0]
        self.txtparts = {}
        self.txtparts[0] = temptxt
        self.text4.ChangeValue(str(self.tblname) + ' Code')
        self.text4.SetFocus()
        self.b2.Disable()
        self.b6.Enable()
        self.restore = False

    def ValData(self):
        # validate the data and truncate incomplete comboboxes

        DataStrg = []

        data_end = (self.Num_Cols*self.Num_Rows)

        # builds tuple of the max dia combobox indexs (1,7,13,19,25,31)
        col_num = 1
        row_num = 0
        aray = [n*self.Num_Cols+col_num for n in range(row_num, self.Num_Rows)]
        aray = tuple(aray)

        for i, x in enumerate(self.cmbctrls):
            cmbvalue = str(self.cmbctrls[i].GetValue())
            # at first empty combobox note the location and exit the FOR loop
            if cmbvalue == '':
                lastcmb = i
                break
            else:
                DataStrg.append(cmbvalue)
                lastcmb = self.Num_Cols*self.Num_Rows
            # if both diamters are complete compare them and confirm Max > Min
            if i in aray:
                if self.cmbctrls[i-1].GetValue() != '' and \
                        self.cmbctrls[i].GetValue() != '':
                    test = eval(cmbvalue.replace('"', '').replace('-', '+'))
                    - eval(self.cmbctrls[i-1].GetValue().
                           replace('"', '').replace('-', '+'))
                    if test <= 0:
                        wx.MessageBox('''Minimum Diameter must be less
                                      than Maximum Diameter.''',
                                      'Diameter Error',
                                      wx.OK | wx.ICON_INFORMATION)
                        self.cmbctrls[i].ChangeValue('')
                        return False

        if i/self.Num_Cols < 1:  # any cmb in the first row
            DataStrg = []
        else:
            if(i % self.Num_Cols):   # anywhere but the start of a new row
                if i == (self.Num_Cols*self.Num_Rows) and \
                        self.cmbctrls[i] != '':
                    data_end = i
                    MsgBx = wx.MessageDialog(
                        self, 'Value needed for ' + self.columnames[i] +
                        '''\n OK will delete incomplete row of data.\n
                          CANCEL will return for data correction.''',
                        'Missing Data', wx.OK | wx.CANCEL | wx.ICON_HAND)

                    MsgBx_val = MsgBx.ShowModal()
                    MsgBx.Destroy()
                    if MsgBx_val == wx.ID_CANCEL:
                        return False
                    # reverts back to the end of the previous row
                    data_end = (i//self.Num_Cols)*self.Num_Cols
            else:  # this represents the start of a new row (5,10,15)
                data_end = i

        # truncate the Fitting Code to reflect
        # the last completed row of comboboxes
        DataStrg = DataStrg[:data_end]
        code = self.NewTextStr()
        skip = 2
        if self.tblname == 'GrooveClamps':
            skip = 3
        n = 0
        for pos, char in enumerate(code):
            if '.' == char:
                n += 1
                if n == data_end:
                    break
        # add 2 because pos is 1 behind n and there is a period
        code = code[:pos+2]
        # build array of all the row end cmbbox numbers
        aray = [n*self.Num_Cols+5 for n in range(0, self.Num_Rows)]
        # lastcmb box completed is the final box then do nothing
        if lastcmb != max(aray):
            for i in aray:
                # this is the end box of the last completed row
                if lastcmb >= i:
                    end_data = i
            for m in range(end_data + skip, (self.Num_Cols*self.Num_Rows)):
                if m in self.txtparts:
                    del self.txtparts[m]
        self.text4.ChangeValue(code)
        self.b2.Disable()
        return True

    def Search(self):
        if self.Num_Cols <= self.text4.GetValue().count('.'):
            field = Dbase().Dcolinfo(self.tblname)[1][1]
            ShQuery = ('SELECT ' + self.tblname[:-1] + 'ID FROM ' +
                       self.tblname + ' WHERE ' + field + ' = "' +
                       self.text4.GetValue() + '"')
            existing = Dbase().Search(ShQuery)

            if existing is False:
                if self.ComdPrtyID is not None:
                    # self.b2.SetLabel('Add/Update\nto Commodity')
                    self.b3.SetLabel('Delete From\nCommodity')
                # else:
                    # self.b2.SetLabel("Add/Update\nto " + self.frmtitle)

            return existing

    def AddRec(self, SQL_step):
        # this adds or updates a record to the fitting table only
        data_strg = []

        if SQL_step == 0:  # enter new record
            # get the next available autoincremented
            # value for the fitting table
            IDE = cursr.execute(
                "SELECT MAX(" + self.tblname[:-1] + "ID) FROM "
                + self.tblname).fetchone()[0]
            if IDE is None:
                IDE = 0
            IDE = int(IDE+1)

            names = []
            for record in Dbase().Dcolinfo(self.tblname):
                names.append(record[1])
            names.remove(self.tblname[:-1] + 'ID')
            names.remove(self.tblname[:-1] + '_Code')
            names.remove('Spec')
            num_col = len(names)

            data_strg.append(str(IDE),)
            data_strg.append(self.code)
            data_strg.append(str(self.spec_ID))

            end_txtparts = len(self.txtparts)

            for n in range(1, end_txtparts):
                if n in self.txtparts:
                    data_strg.append(self.txtparts[n])

            num_rec = len(self.txtparts)

            if self.tblname in ['Fittings', 'GrooveClamps']:
                strt = end_txtparts
                stp = (num_col + 1)
            else:
                strt = 0
                stp = (num_col - num_rec + 1)

            for n in range(strt, stp):
                data_strg.append(None)

            num_vals = ('?,'*len(self.colms))[:-1]
            UpQuery = ("INSERT INTO " + self.tblname +
                       " VALUES (" + num_vals + ")")
            Dbase().TblEdit(UpQuery, data_strg)
            CurrentID = IDE

        elif SQL_step == 1:  # update edited record
            realnames = []
            for item in Dbase().Dcolinfo(self.tblname):
                realnames.append(item[1])

            # remove the name of table ID from list as it is not updating
            realnames.remove(self.pkcol_name)
            realnames.remove('Spec')
            # and will use the current record ID
            CurrentID = self.data[self.rec_num][0]
            # break the code down into string of individual IDs
            self.txtparts = self.code.split('.')
            del self.txtparts[0]
            data_lst = [None if x == '0' else x for x in self.txtparts]
            data_lst.insert(0, (self.code))

            for i in range(len(data_lst), len(realnames)):
                data_lst.append(None)

            Name_str = ','.join(["%s=?" % (name) for name in realnames])
            UpQuery = ('UPDATE ' + self.tblname + ' SET ' + Name_str +
                       ' WHERE ' + self.pkcol_name + ' = ' + str(CurrentID))
            Dbase().TblEdit(UpQuery, data_lst)

        # cancel the update step
        elif SQL_step == 3:
            return

        return CurrentID

    def OnDeleteRec(self, evt):
        recrd = self.data[self.rec_num][0]
        if self.ComdPrtyID is None:
            try:
                Dbase().TblDelete(self.tblname, recrd, self.pkcol_name)
                self.Dsql = 'SELECT * FROM ' + self.tblname
                self.data = Dbase().Dsqldata(self.Dsql)
                self.rec_num -= 1
                if self.rec_num < 0:
                    self.rec_num = len(self.data)-1
                self.FillScreen()
                self.recnum3.SetLabel('/ '+str(len(self.data)))

            except sqlite3.IntegrityError:
                wx.MessageBox("This Record is associated"
                              " with\nother tables and cannot be deleted!",
                              "Cannot Delete",
                              wx.OK | wx.ICON_INFORMATION)
        else:
            self.ChgSpecID()
            self.RestoreCmbs()
            self.b3.Disable()
            self.data = []
            self.rec_num = 0
            self.recnum2.SetLabel(str(self.rec_num))
            self.recnum3.SetLabel('/ '+str(len(self.data)))
            self.lbl_nodata.SetLabel(
                'The ' + self.frmtitle +
                ' have not been setup for this Commodity Property')

    def OnAdd1(self, evt):
        btn = evt.GetEventObject()
        callbtn = self.btnDict[btn]
        cmbtbl = ''
        # numbers are for the index number for each group of comboboxes
        if self.tblname == 'Fittings':
            if callbtn in (0, 1):
                boxnums = (0, 1, 5, 6, 10, 11, 15, 16)
                cmbtbl = 'Pipe_OD'
                CmbLst1(self, cmbtbl)
            elif callbtn == 2:
                boxnums = (2, 7, 12, 17)
                cmbtbl = 'EndConnects'
                CmbLst1(self, cmbtbl)
            elif callbtn == 3:
                boxnums = []
                txt = ('''    NOTE: There are 3 potential
            specifications for material sources;''')
                choices = ['1) Forged material',
                           '2) Butt-weld material',
                           '3) Pipe material']
                # use a SingleChioce dialog to determine if
                # data is new record or edited record
                Update_Dialog = wx.SingleChoiceDialog(
                    self, txt, 'Select Material Type', choices,
                    style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)

                Update_step = 3
                if Update_Dialog.ShowModal() == wx.ID_OK:
                    Update_step = Update_Dialog.GetSelection()
                    Update_Dialog.Destroy()
                # selected Forged material table to update so any
                # flanged, threded of SW item needs to be updated for material
                if Update_step == 0:
                    m = 0
                    for n in (2, 7, 12, 17):
                        val = self.cmbctrls[n].GetValue()
                        if val == '':
                            break
                        if val in ['Flanged', 'Threaded',
                                   'Socket Welded', 'SW or Threaded']:
                            cmbtbl = 'ForgedMaterial'
                            boxnums.append(12+m)
                        m += 1
                # selected Butt weld material table to update so any
                # butt welded item needs to be updated for material
                elif Update_step == 1:
                    m = 0
                    for n in (2, 7, 12, 17):
                        val = self.cmbctrls[n].GetValue()
                        if val == '':
                            break
                        if val == 'Butt Welded':
                            cmbtbl = 'ButtWeldMaterial'
                            boxnums.append(12+m)
                        m += 1
                # selected Fabricated material table to update so any
                # piping item item needs to be updated for material
                elif Update_step == 2:
                    m = 0
                    for n in (2, 7, 12, 17):
                        val = self.cmbctrls[n].GetValue()
                        if val == '':
                            break
                        if val == 'Fabricated':
                            cmbtbl = 'PipeMaterial'
                            boxnums.append(12+m)
                        m += 1
                if cmbtbl != '':
                    CmbLst2(self, cmbtbl)
                boxnums = tuple(boxnums)

            elif callbtn == 4:
                boxnums = []
                txt = ('''    NOTE: There are 3 potential
                       specifications for fitting weight;    ''')
                choices = ['1) ANSI Class',
                           '2) Forged Class.',
                           '3) Pipe Schedule']
                # use a SingleChioce dialog to determine
                # if data is new record or edited record
                Update_Dialog = wx.SingleChoiceDialog(
                    self, txt, 'Select Class Type', choices)

                Update_step = 3
                if Update_Dialog.ShowModal() == wx.ID_OK:
                    Update_step = Update_Dialog.GetSelection()
                    Update_Dialog.Destroy()
                # selected ANSI rating so update all classes
                # where ends are flanged
                if Update_step == 0:
                    m = 0
                    for n in (2, 7, 12, 17):
                        val = self.cmbctrls[n].GetValue()
                        if val == '':
                            break
                        if val == 'Flanged':
                            cmbtbl = 'ANSI_Rating'
                            boxnums.append(16+m)
                        m += 1
                # selected forged class so any ends with forged
                # fittings need the class upated for forgings
                elif Update_step == 1:
                    m = 0
                    for n in (2, 7, 12, 17):
                        val = self.cmbctrls[n].GetValue()
                        if val == '':
                            break
                        if val in ['Threaded', 'Socket Welded',
                                   'SW or Threaded']:
                            cmbtbl = 'ForgeClass'
                            boxnums.append(16+m)
                        m += 1
                # selected pipe schedules so any ends which require
                # pipe schedule class need class box list updated
                elif Update_step == 2:
                    m = 0
                    for n in (2, 7, 12, 17):
                        val = self.cmbctrls[n].GetValue()
                        if val == '':
                            break
                        if val in ['Fabricated', 'Butt Welded']:
                            cmbtbl = 'PipeSchedule'
                            boxnums.append(16+m)
                        m += 1
                if cmbtbl != '':
                    CmbLst1(self, cmbtbl)
                boxnums = tuple(boxnums)

        else:
            boxnums = [n*self.Num_Cols+callbtn
                       for n in range(0, self.Num_Rows)]
            cmbtbl = self.cmb_tbls[callbtn]
            CmbLst1(self, cmbtbl)

        if cmbtbl != '':
            self.ReFillList(cmbtbl, boxnums)

    def OnAddComd(self, evt):
        self.AddComd()

    # link this ID to the commodity property
    def AddComd(self):
        # if this is a show all update just revise self.data
        if self.b6.GetLabel() == 'Show All\n' + self.frmtitle:
            if self.DsqlFtg == '':
                self.DsqlFtg = 'SELECT * FROM ' + self.tblname
                self.data = Dbase().Dsqldata(self.DsqlFtg)
            else:
                self.DsqlFtg = self.DsqlFtg[:self.DsqlFtg.find('WHERE')]
                self.data = Dbase().Dsqldata(self.DsqlFtg)
            self.b6.SetLabel("Add Item\nto Commodity")
            self.b6.Enable()
            self.b3.Disable()
            self.FillScreen()
            self.recnum3.SetLabel('/ '+str(len(self.data)))
        # this is now an update spec ID in pipe specification do the following
        else:
            # check if the commodity property exists in pipespec
            query = ('SELECT ' + self.tblname[:-1] + '_ID' +
                     ' FROM PipeSpecification WHERE Commodity_Property_ID = '
                     + str(self.ComdPrtyID))
            StartQry = Dbase().Dsqldata(query)
            # commodity code has not been set up in pipespec set up new ID
            if StartQry == []:
                ValueList = []
                New_ID = cursr.execute("SELECT MAX(Pipe_Spec_ID)\
                 FROM PipeSpecification").fetchone()
                if New_ID[0] is None:
                    Max_ID = '1'
                else:
                    Max_ID = str(New_ID[0]+1)
                colinfo = Dbase().Dcolinfo('PipeSpecification')
                for n in range(0, len(colinfo)-2):
                    ValueList.append(None)

                num_vals = ('?,'*len(colinfo))[:-1]
                ValueList.insert(0, Max_ID)
                ValueList.insert(1, str(self.ComdPrtyID))

                UpQuery = ("INSERT INTO PipeSpecification VALUES ("
                           + num_vals + ")")
                Dbase().TblEdit(UpQuery, ValueList)
                StartQry = Max_ID
            else:
                StartQry = str(StartQry[0][0])

            cmd_addID = self.data[self.rec_num][0]
            self.ChgSpecID(cmd_addID)

            self.DsqlFtg = ('SELECT * FROM ' + self.tblname +
                            ' WHERE ' + self.pkcol_name +
                            ' = ' + str(cmd_addID))
            self.data = Dbase().Dsqldata(self.DsqlFtg)

            self.rec_num = 0
            self.FillScreen()
            self.lbl_nodata.SetLabel('   ')
            self.recnum3.SetLabel('/ '+str(len(self.data)))
            self.b3.Enable()
            self.b6.SetLabel("Show All\n" + self.frmtitle)

    def ReFillList(self, cmbtbl, boxnums):
        for n in boxnums:
            lctr = self.lctrls[n]
            lctr.DeleteAllItems()
            index = 0

            ReFillQuery = 'SELECT * FROM "' + cmbtbl + '"'

            if cmbtbl in ['ForgedMaterial', 'ButtWeldMaterial',
                          'PipeMaterial']:
                ReFillQuery = ('SELECT * FROM ' + cmbtbl +
                               ' WHERE Material_Type = "' +
                               self.MtrSpc[1] +
                               '" AND Material_Grade = "' +
                               self.MtrSpc[2] +
                               '"')
            for values in Dbase().Dsqldata(ReFillQuery):
                col = 0
                for value in values:
                    if col == 0:
                        lctr.InsertItem(index, str(value))
                    else:
                        lctr.SetItem(index, col, str(value))
                    col += 1
                index += 1

    def OnCmbOpen(self, evt):
        self.cmbOld = evt.GetEventObject().GetValue()

    def OnCmbClose(self, evt):
        cmbNew = evt.GetEventObject().GetValue()
        if self.restore is False and self.cmbOld != cmbNew:
            self.CmbClose(evt.GetEventObject())
            self.b2.Enable()
            if self.ComdPrtyID is not None:
                self.b6.Disable()

    def CmbClose(self, cmbselect):
        i = int()

        # the required table is self.cmb_tbls[i]
        # number each combo in turn and check if it is the
        # one selected, record number of selected box
        for i in (i for i, x in enumerate(self.cmbctrls) if x == cmbselect):
            cmbvalue = str(self.cmbctrls[i].GetValue())
        # build the string for the fittings code
        # check if index 2 is has been selected the end connections box
        # if it has been
            if self.tblname == 'Fittings':
                # if the end connection has been selected
                # then the material type and rating can be specified
                # this is the number of the end connection boxes
                if i in (2, 7, 12, 17):
                    if cmbvalue in ('Threaded', 'Socket Welded',
                                    'SW or Threaded', ''):
                        # empty quotes for deleted end connection
                        rtg_tbl = 'ForgeClass'
                        mtr_tbl = 'ForgedMaterial'
                    elif cmbvalue == 'Flanged':
                        rtg_tbl = 'ANSI_Rating'
                        mtr_tbl = 'ForgedMaterial'
                    elif cmbvalue == 'Butt Welded':
                        rtg_tbl = 'PipeSchedule'
                        mtr_tbl = 'ButtWeldMaterial'
                    elif cmbvalue == 'Fabricated':
                        rtg_tbl = 'PipeSchedule'
                        mtr_tbl = 'PipeMaterial'

                    # set the combo list table for the
                    # material and rating/class combo box
                    self.cmb_tbls[3] = mtr_tbl
                    self.cmb_tbls[4] = rtg_tbl
                    self.bxdict[i] = (rtg_tbl, mtr_tbl)

                    # get the data for the material combo box
                    query = ('SELECT * FROM ' + mtr_tbl +
                             ' WHERE Material_Type = "' +
                             self.MtrSpc[1] +
                             '" AND Material_Grade = "' +
                             self.MtrSpc[2] +
                             '"')
                    self.cmbctrls[i+1].ChangeValue('')
                    self.cmbctrls[i+1].Enable()
                    self.showcol = 1
                    self.cmbctrls[i+1].SetPopupControl(
                        ListCtrlComboPopup(self.bxdict[i][1], query, cmbvalue,
                                           showcol=self.showcol,
                                           lctrls=self.lctrls))

                    # get the data for the rating or class combo box
                    query = 'SELECT * FROM ' + rtg_tbl
                    self.cmbctrls[i+2].ChangeValue('')
                    self.cmbctrls[i+2].Enable()
                    self.showcol = 1
                    if rtg_tbl == 'PipeSchedule':
                        self.showcol = 2
                    self.cmbctrls[i+2].SetPopupControl(
                        ListCtrlComboPopup(self.bxdict[i][0], query, cmbvalue,
                                           showcol=self.showcol,
                                           lctrls=self.lctrls))
                    # txtparts is a dictionary of
                    # combobox number:combobox recordID
                    if i == 2:
                        if 4 in self.txtparts:
                            self.txtparts[4] = 0
                        if 5 in self.txtparts:
                            self.txtparts[5] = 0

                    if i == 7:
                        if 9 in self.txtparts:
                            self.txtparts[9] = 0
                        if 10 in self.txtparts:
                            self.txtparts[10] = 0

                    if i == 12:
                        if 14 in self.txtparts:
                            self.txtparts[14] = 0
                        if 15 in self.txtparts:
                            self.txtparts[15] = 0

                    if i == 17:
                        if 19 in self.txtparts:
                            self.txtparts[19] = 0
                        if 20 in self.txtparts:
                            self.txtparts[20] = 0

                # build the string for the fittings code
                # check if index 2 is has been selected the end connections box
                # if it has been
                # b the remainder will be 0,1,2,3,4 representing each column
                if cmbvalue != '':
                    b = i % self.Num_Cols
                    if b <= 2:   # this is for the 1st 3 boxes in each row
                        tbl_name = self.cmb_tbls[b]
                    elif b == 3:   # this is for the material boxes
                        tbl_name = self.bxdict[i-1][1]
                    else:
                        # this is for the rating boxes
                        tbl_name = self.bxdict[i-2][0]
                    newlbl = self.LblData(tbl_name, cmbvalue)
                    self.txtparts[1+i] = newlbl
                else:    # if the cmb box is empty then set the txtpart to 0
                    if (i+1) in self.txtparts:
                        self.txtparts[i+1] = 0

                # redevelope the fitting code based on
                # the newly selected combo box
                self.text4.ChangeValue(self.NewTextStr())

                if (len(self.txtparts)-1) in [5, 10, 15]:
                    for n in range((len(self.txtparts)-1),
                                   len(self.txtparts) + 2):
                        self.cmbctrls[n].Enable()
            else:
                # this builds a tuple of all the last combobox
                # indexs in each column for forgings
                col_num = 0
                row_num = 1
                # start the aray list all the combo boxes in the
                # first combo box(0) starting with the second row(1)
                aray = [n*self.Num_Cols+col_num for n in
                        range(row_num, self.Num_Rows)]
                aray = tuple(aray)

                if cmbvalue != '':
                    if self.tblname == 'GrooveClamps' and i in [4, 11]:
                        query = ('''SELECT Manufacture FROM
                                 GrooveClampVendor WHERE Style = "'''
                                 + cmbvalue + '"')
                        Vndr = Dbase().Dsqldata(query)[0][0]
                        self.txtVndrs[i//11].ChangeValue(Vndr)

                    b = i % self.Num_Cols
                    newlbl = self.LblData(self.cmb_tbls[b], cmbvalue)

                    if newlbl:
                        self.txtparts[1+i] = newlbl
                    elif (i+1) in self.txtparts:
                        del self.txtparts[i+1]
                else:
                    self.txtparts[1+i] = '0'

                self.text4.ChangeValue(self.NewTextStr())

                # row has been completed enable the next row of boxes
                if (len(self.txtparts)-1) in aray:
                    distance = 5
                    if self.tblname == "GrooveClamps":
                        distance = 6
                    for n in range((len(self.txtparts)-1),
                                   len(self.txtparts) + distance):
                        self.cmbctrls[n].Enable()

    def OnRestoreCmbs(self, evt):
        self.RestoreCmbs()

    def OnSearch(self, evt):
        self.Search()

    def OnAddRec(self, evt):
        tmp = self.text4.GetValue()[:2]
        lcd = self.text4.GetValue()[3:].split('.')

        if '0' in lcd:
            lcd = lcd[:(lcd.index('0')-lcd.index('0') % self.Num_Cols)]
        else:
            lcd = lcd[:len(lcd)-len(lcd) % self.Num_Cols]
        lcd.insert(0, tmp)
        self.code = '.'.join(str(r) for r in lcd)
        self.txtparts = {i: lcd[i] for i in range(0, len(lcd))}
        self.text4.ChangeValue(self.code)

        # check first that data is all present and clean up incomplete boxes
        check = self.ValData()
        if check:
            # check to see if code is unique
            existing = self.Search()
            # if the fitting code already exist do nothing
            if existing:
                wx.MessageBox('This ' + self.frmtitle +
                              ' Code already exists.', "Existing Record",
                              wx.OK | wx.ICON_INFORMATION)
                self.FillScreen()
                self.b6.Enable()
                return

            SQL_step = 3

            choice1 = ('1) Save this as a new '
                       + self.frmtitle + ' Specification')
            choice2 = ('2) Update the existing '
                       + self.frmtitle + ' Specification with this data')
            txt1 = ('''NOTE: Updating this information will be
                    reflected in all associated ''' + self.frmtitle)
            txt2 = (''' Specifications!\nRecommendation is to save as a
                     new specification.\n\n\tHow do you want to proceed?''')

            # if this is a not commodity related change
            if self.ComdPrtyID is None:
                # Make a selection as to whether the record
                # is to be a new or an update valve
                # use a SingleChioce dialog to determine
                # if data is new record or edited record
                SQL_Dialog = wx.SingleChoiceDialog(
                    self, txt1 + txt2, 'Information Has Changed',
                    [choice1, choice2], style=wx.CHOICEDLG_STYLE)
                if SQL_Dialog.ShowModal() == wx.ID_OK:
                    SQL_step = SQL_Dialog.GetSelection()
                SQL_Dialog.Destroy()

                self.AddRec(SQL_step)
                self.DsqlFtg = 'SELECT * FROM ' + self.tblname
                self.data = Dbase().Dsqldata(self.DsqlFtg)

            else:  # this is a commodity related change
                if self.data == []:
                    SQL_step = 0
                else:
                    choice1 = choice1 + ' for this commodity?'
                    choice2 = choice2 + ' for this commodity?'
                    # use a SingleChioce dialog to determine
                    # if data is new record or edited record
                    SQL_Dialog = wx.SingleChoiceDialog(
                        self, txt1+txt2, 'Information Has Changed',
                        [choice1, choice2], style=wx.CHOICEDLG_STYLE)
                    if SQL_Dialog.ShowModal() == wx.ID_OK:
                        SQL_step = SQL_Dialog.GetSelection()

                    SQL_Dialog.Destroy()

                cmd_addID = self.AddRec(SQL_step)
                # always save as a new spec if commodity property is specified
                # no matter the change over write or add the specification ID
                # to the PipeSpecification table under the ComdPrtyID

                self.ChgSpecID(cmd_addID)
                # reset the self.data to reflect changes
                query = ('SELECT ' + self.tblname[:-1] + '_ID' +
                         ''' FROM PipeSpecification
                         WHERE Commodity_Property_ID = '''
                         + str(self.ComdPrtyID))
                StartQry = Dbase().Dsqldata(query)
                self.DsqlFtg = ('SELECT * FROM ' + self.tblname +
                                ' WHERE ' + self.pkcol_name +
                                ' = ' + str(StartQry[0][0]))
                self.data = Dbase().Dsqldata(self.DsqlFtg)
                self.b3.Enable()
                self.b6.SetLabel("Show All\n" + self.frmtitle)
                self.b6.Enable()

            if SQL_step == 0:
                self.rec_num = len(self.data)-1
            if SQL_step == 1 and self.ComdPrtyID is not None:
                self.rec_num = 0
            self.FillScreen()
            self.recnum3.SetLabel('/ '+str(len(self.data)))

    def ChgSpecID(self, ID=None):
        UpQuery = ('UPDATE PipeSpecification SET ' + self.tblname[:-1] +
                   '_ID=?  WHERE Commodity_Property_ID = '
                   + str(self.ComdPrtyID))
        Dbase().TblEdit(UpQuery, [ID])

    def ReportData(self):
        rptdata = []

        if self.tblname == 'Fittings':
            Colnames = [('ID', 'Fitting Code', 'Material\nSpecification'),
                        tuple(self.lbl_txt)]
            Colwdths = [(6, 40, 10), (10, 10, 20, 20, 10)]
        elif self.tblname == 'Flanges':
            Colnames = [('ID', 'Flange Code', 'Material\nSpecification'),
                        tuple(self.lbl_txt)]
            Colwdths = [(6, 40, 10), (10, 10, 20, 20, 10, 10)]
        elif self.tblname == 'OrificeFlanges':
            Colnames = [('ID', 'Orifice\nFlange Code',
                        'Material\nSpecification'),
                        tuple(self.lbl_txt)]
            Colwdths = [(6, 40, 10), (10, 10, 20, 20, 10, 10)]
        elif self.tblname == 'Unions':
            Colnames = [('ID', 'Union Code', 'Material\nSpecification'),
                        tuple(self.lbl_txt)]
            Colwdths = [(6, 40, 10), (10, 10, 20, 20, 10, 10)]
        elif self.tblname == 'OLets':
            Colnames = [('ID', 'Olet Code', 'Material\nSpecification'),
                        tuple(self.lbl_txt)]
            Colwdths = [(6, 40, 10), (10, 10, 15, 10, 20, 10)]
        elif self.tblname == 'GrooveClamps':
            Colnames = [('ID', 'Groove\nClamp Code',
                         'Material\nSpecification'),
                        tuple(self.lbl_txt)]
            Colwdths = [(6, 40, 10), (10, 10, 10, 10, 20, 20, 20)]

        numcols = len(Colwdths[1])

        for segn in range(len(self.data)):
            data1 = list(self.data[segn][0:3])
            data2 = list(self.data[segn][3:])

            numrows = len(data2)//numcols

            # determine the matr spec designation used for all tables
            query = ('''SELECT Pipe_Material_Spec FROM PipeMaterialSpec
                    WHERE Material_Spec_ID = ''' + str(data1[2]))
            data1[2] = Dbase().Dsqldata(query)[0][0]

            n = 0
            for i in data2:
                if i is None:
                    break
                # first column numbers
                if n in [m*numcols for m in range(0, numrows)]:
                    # [0,1,5,6,10,11,15,16]:
                    query3 = ("SELECT Pipe_OD FROM Pipe_OD WHERE PipeOD_ID = "
                              + str(i))
                # second column numbers
                elif n in [m*numcols+1 for m in range(0, numrows)]:
                    query3 = ("SELECT Pipe_OD FROM Pipe_OD WHERE PipeOD_ID = "
                              + str(i))
                # third column numbers
                elif n in [m*numcols+2 for m in range(0, numrows)]:
                    # [2,7,12,17]:
                    if self.tblname in ['Fittings', 'Unions']:
                        query3 = ('''SELECT Connection FROM EndConnects
                                WHERE EndID = ''' + str(i))
                        if i in [2, 4, 5]:
                            self.rtg_tbl = 'ForgeClass'
                            self.mtr_tbl = 'ForgedMaterial'
                        elif i == 1:
                            self.rtg_tbl = 'ANSI_Rating'
                            self.mtr_tbl = 'ForgedMaterial'
                        elif i == 3:
                            self.rtg_tbl = 'PipeSchedule'
                            self.mtr_tbl = 'ButtWeldMaterial'
                        elif i == 6:
                            self.rtg_tbl = 'PipeSchedule'
                            self.mtr_tbl = 'PipeMaterial'
                    elif self.tblname in ['Flanges', 'OrificeFlanges']:
                        query3 = ('''SELECT Style_Type FROM FlangeStyle WHERE
                                StyleID = ''' + str(i))
                    elif self.tblname == 'OLets':
                        query3 = (
                            'SELECT Pipe_OD FROM Pipe_OD WHERE PipeOD_ID = '
                            + str(i))
                    elif self.tblname == 'GrooveClamps':
                        query3 = ('''SELECT Pipe_Schedule FROM PipeSchedule
                                WHERE ID = ''' + str(i))
                # fourth column numbers
                elif n in [m*numcols+3 for m in range(0, numrows)]:
                    # [3,8,13,18]:
                    if self.tblname == 'Fittings':
                        tbl_name = self.mtr_tbl
                        if tbl_name == 'ForgedMaterial':
                            query3 = (
                                'SELECT Forged_Material FROM ' +
                                tbl_name + " WHERE ID = " + str(i))
                        elif tbl_name == 'ButtWeldMaterial':
                            query3 = ('SELECT Butt_Weld_Material FROM ' +
                                      tbl_name + " WHERE ID = " + str(i))
                        elif tbl_name == 'PipeMaterial':
                            query3 = ('SELECT Pipe_Material FROM ' + tbl_name +
                                      " WHERE ID = " + str(i))
                    elif self.tblname in ['Flanges', 'OrificeFlanges']:
                        query3 = ('''SELECT Face_Style FROM FlangeFace
                                WHERE FaceID = ''' + str(i))
                    elif self.tblname == 'Unions':
                        query3 = ('''SELECT Seat_Material FROM SeatMaterial
                                WHERE SeatMaterialID = ''' + str(i))
                    elif self.tblname == 'OLets':
                        query3 = (
                            "SELECT Style FROM OLetStyle WHERE StyleID = "
                            + str(i))
                    elif self.tblname == 'GrooveClamps':
                        query3 = ('''SELECT Groove FROM ClampGroove
                                WHERE ClampGrooveID = ''' + str(i))
                # fifth column numbers
                elif n in [m*numcols+4 for m in range(0, numrows)]:
                    # [4,9,14,19]:
                    if self.tblname == 'Fittings':
                        tbl_name = self.rtg_tbl
                        if tbl_name == 'ANSI_Rating':
                            query3 = ('SELECT ANSI_Class FROM ' + tbl_name +
                                      ' WHERE Rating_Designation = ' + str(i))
                        elif tbl_name == 'ForgeClass':
                            query3 = ('SELECT Forged_Class FROM ' + tbl_name +
                                      ' WHERE ClassID = ' + str(i))
                        elif tbl_name == 'PipeSchedule':
                            query3 = ('SELECT Pipe_Schedule FROM ' + tbl_name +
                                      " WHERE ID = " + str(i))
                    elif self.tblname in ['Flanges', 'OrificeFlanges',
                                          'Unions', 'OLets']:
                        query3 = ('''SELECT Forged_Material FROM
                                ForgedMaterial WHERE ID = ''' + str(i))
                    elif self.tblname == 'GrooveClamps':
                        query3 = ('''SELECT Style FROM GrooveClampVendor
                                WHERE VendorID = ''' + str(i))
                # sixth column numbers
                elif n in [m*numcols+5 for m in range(0, numrows)]:
                    if self.tblname in ['Flanges', 'OrificeFlanges']:
                        query3 = ('''SELECT ANSI_Class FROM ANSI_Rating
                                    WHERE Rating_Designation = ''' + str(i))
                    elif self.tblname == 'Unions':
                        query3 = ('''SELECT Forged_Class FROM ForgeClass
                                WHERE ClassID = ''' + str(i))
                    elif self.tblname == 'OLets':
                        query3 = ("SELECT Weight FROM OLetWt WHERE OLetWtID = "
                                  + str(i))
                    elif self.tblname == 'GrooveClamps':
                        query3 = ('''SELECT GasketSealMaterial FROM GasketSealMaterial
                                WHERE SealID = ''' + str(i))
                # seventh column numbers
                elif n in [m*numcols+6 for m in range(0, numrows)]:
                    if self.tblname == 'GrooveClamps':
                        query3 = ('''SELECT Forged_Material FROM
                                ForgedMaterial WHERE ID = ''' + str(i))

                data2[n] = str(Dbase().Dsqldata(query3)[0][0])

                n += 1

            rptdata.append(tuple(data1))
            rptdata.append(tuple(data2[:n]))

        return (rptdata, Colnames, Colwdths)

    def OnClose(self, evt):
        # add following 2 lines for child parent form
        self.GetParent().Enable(True)
        self.__eventLoop.Exit()
        self.Destroy()


class BldWeld(wx.Frame):
    def __init__(self, parent, tblname, ComdPrtyID=None):

        self.parent = parent    # add for child form

        self.tblname = tblname
        if self.tblname.find("_") != -1:
            self.frmtitle = (self.tblname.replace("_", " "))
        else:
            self.frmtitle = (' '.join(re.findall('([A-Z][a-z]*)',
                             self.tblname)))

        super(BldWeld, self).__init__(parent,
                                      title=self.frmtitle,
                                      size=(900, 720))

        self.Bind(wx.EVT_CLOSE, self.OnClose)

        self.recrd = []
        self.rec_num = 0
        self.sub_rec_num = 0
        self.lctrls = []
        self.data = []
        self.subdata = []
        self.btnDict = {}

        self.startstrg = ''
        self.field = 'WeldRequirementsID'
        self.MainSQL = ''
        StartQry = None

        self.ComCode = ''
        self.PipeMtrSpec = ''

        self.ComdPrtyID = ComdPrtyID

        if self.ComdPrtyID is not None:
            query = ('''SELECT Commodity_Code,Pipe_Material_Code,
                     End_Connection,Pipe_Code FROM CommodityProperties WHERE
                      CommodityPropertyID = ''' + str(self.ComdPrtyID))
            dataset = Dbase().Dsqldata(query)[0]
            self.PipeMtrSpec = dataset[1]
            self.ComCode = dataset[0]
            self.PipeCode = dataset[3]

        # order of information text label,table name,combo width,
        # combo box true = 1,table ID col name,column name shown in box
        self.hints_tbls = [
            ('Filler', 'WeldFiller', (325, -1), 0, 'ID', 'Metal_Spec'),
            ('Weld Thickness\nQualification', 'WeldQualifyThickness',
             (150, -1), 0, 'ID', 'Thickness'),
            ('   Welder Qualification\n   Certificate to Show', '',
             (350, 60), 0),
            ('Weld\nProcess', 'WeldProcessList', (250, -1),
             0, 'ID', 'Process'),
            ('Material\nGroup', 'WeldMaterialGroup', (115, -1), 0, 'ID',
             'Material_Group'),
            ('Filler Metal\nGroup Number', 'WeldFillerGroup',
             (175, -1), 0, 'ID', 'Filler_Group'),
            ('Position\nQualified', 'WeldQualifyPosition',
             (100, -1), 0, 'ID', 'Position'),
            ('Special Notes:', '', (500, 60), 0)]

        # select data based on if the form is called from commodity
        # properties form or if you are to see all data
        if self.ComdPrtyID is not None:
            query = ('SELECT ' + self.tblname +
                     '''_ID FROM PipeSpecification WHERE
                      Commodity_Property_ID = ''' +
                     str(self.ComdPrtyID))
            chk = Dbase().Dsqldata(query)
            if chk != []:
                StartQry = chk[0][0]
                if StartQry is not None:
                    self.MainSQL = ('SELECT * FROM ' + self.tblname + ' WHERE '
                                    + self.tblname + 'ID = ' + str(StartQry))
                    self.data = Dbase().Dsqldata(self.MainSQL)
                    # range of procedureID fields is 2 to 5
                    for n in range(2, 6):
                        SubSQL = ('''SELECT * FROM WeldProcedures
                                   WHERE ProcedureID = "''' +
                                  str(self.data[0][n]) + '"')
                        if Dbase().Dsqldata(SubSQL) != []:
                            self.subdata.append(Dbase().Dsqldata(SubSQL)[0])
        else:
            self.MainSQL = 'SELECT * FROM ' + self.tblname
            self.data = Dbase().Dsqldata(self.MainSQL)

        tblinfo = Dbase(self.tblname).Fld_Size_Type()
        self.ID_col = tblinfo[1]
        self.colms = tblinfo[3]
        self.pkcol_name = tblinfo[0]
        self.autoincrement = tblinfo[2]

        # specify which listbox column to display in the combobox
        self.showcol = int

        # set the Sizer property (same as SetSizer)
        self.Sizer = wx.BoxSizer(wx.VERTICAL)

        font1 = wx.Font(12, wx.DEFAULT, wx.NORMAL, wx.FONTWEIGHT_BOLD)

        self.nodata_txt = (' The ' + self.frmtitle +
                           ' have not been setup for this Commodity Property')
        self.lbl_nodata = wx.StaticText(self, -1, label=self.nodata_txt,
                                        size=(700, 30), style=wx.LEFT)
        self.lbl_nodata.SetForegroundColour((255, 0, 0))
        self.lbl_nodata.SetFont(font1)
        self.lbl_nodata.SetLabel('   ')

        self.warningsizer = wx.BoxSizer(wx.HORIZONTAL)
        self.warningsizer.Add(self.lbl_nodata, 0, wx.ALIGN_CENTER)

        specsizer1 = wx.BoxSizer(wx.HORIZONTAL)
        # build the welding process combobox
        # depending on if the value is preset
        prcsttl = wx.StaticText(self, label='Welding Process',
                                style=wx.ALIGN_LEFT)
        prcsttl.SetForegroundColour((255, 0, 0))
        self.cmbProces = wx.ComboCtrl(self, pos=(10, 10), size=(550, -1))
        self.cmbProces.SetPopupControl(
            ListCtrlComboPopup('WeldProcesses', showcol=1, lctrls=self.lctrls))
        self.Bind(wx.EVT_COMBOBOX_DROPDOWN, self.OnCmbOpen, self.cmbProces)
        self.Bind(wx.EVT_COMBOBOX_CLOSEUP, self.OnCmbExit, self.cmbProces)

        addbtn1 = wx.Button(self, label='+', size=(35, -1))
        addbtn1.SetForegroundColour((255, 0, 0))
        addbtn1.SetFont(font1)
        self.Bind(wx.EVT_BUTTON, self.OnAdd, addbtn1)
        self.btnDict[addbtn1] = 0

        specsizer2 = wx.BoxSizer(wx.HORIZONTAL)
        title1 = wx.StaticText(
            self, label='Procedure Notes:\n(Press Enter to\naccept changes)',
            style=wx.ALIGN_LEFT | wx.ALIGN_CENTER)
        title1.SetForegroundColour((255, 0, 0))
        self.text1 = wx.TextCtrl(self, size=(550, 60), value='',
                                 style=wx.TE_MULTILINE | wx.TE_PROCESS_ENTER)
        self.Bind(wx.EVT_TEXT_ENTER, self.OnSelect1, self.text1)

        specsizer1.Add(prcsttl, 0, wx.ALIGN_LEFT | wx.TOP, 15)
        specsizer1.Add((10, 10))
        specsizer1.Add(self.cmbProces, 0, wx.TOP, 10)
        specsizer1.Add(addbtn1, 0, wx.TOP, 10)
        specsizer2.Add(title1, 0, wx.TOP, 5)
        specsizer2.Add(self.text1, 0, wx.BOTTOM | wx.LEFT, 10)

        self.Sizer.Add(self.warningsizer, 0, wx.CENTER | wx.TOP, 5)
        self.Sizer.Add(specsizer1, 0, wx.CENTER | wx.ALL, 10)
        self.Sizer.Add(specsizer2, 0, wx.CENTER | wx.ALL, 10)

        self.BxBld()

        prcsbxszr = wx.StaticBox(self, -1, '*Procedure:',
                                 style=wx.ALIGN_CENTER)
        prcsbxszr.SetForegroundColour('blue')
        bxSizer = wx.StaticBoxSizer(prcsbxszr, wx.VERTICAL)

        tlesizer1 = wx.BoxSizer(wx.HORIZONTAL)
        txtProced = wx.StaticText(self, label='Weld Procedure',
                                  style=wx.ALIGN_LEFT)
        txtProced.SetForegroundColour((255, 0, 0))
        self.cmbProced = wx.ComboCtrl(self, size=(100, -1))
        self.Bind(wx.EVT_TEXT, self.OnSelectProced, self.cmbProced)
        self.showcol = 1
        query = 'SELECT WeldProcedures.Procedure FROM WeldProcedures'
        self.cmbProced.SetPopupControl(ListCtrlComboPopup(
            'WeldProcedures', PupQuery=query, lctrls=self.lctrls))

        addbtn2 = wx.Button(self, label='+', size=(35, -1))
        addbtn2.SetForegroundColour((255, 0, 0))
        addbtn2.SetFont(font1)
        self.Bind(wx.EVT_BUTTON, self.OnAdd, addbtn2)
        self.btnDict[addbtn2] = 1

        tlesizer1.Add(txtProced, 0, wx.LEFT, border=15)
        tlesizer1.Add(self.cmbProced, 0, wx.LEFT, border=15)
        tlesizer1.Add(addbtn2, 0)
        tlesizer1.Add(self.txtlbls[6], 0, wx.LEFT, border=25)
        tlesizer1.Add(self.txtbxs[6], 0, wx.LEFT, border=10)

        bxSizer.Add(tlesizer1, 0, wx.CENTER | wx.ALL, border=10)

        tlesizer2 = wx.BoxSizer(wx.HORIZONTAL)

        tlesizer2.Add(self.txtlbls[3], 0, wx.LEFT, border=15)
        tlesizer2.Add(self.txtbxs[3], 0, wx.LEFT, border=10)
        tlesizer2.Add(self.txtlbls[1], 0, wx.LEFT, border=25)
        tlesizer2.Add(self.txtbxs[1], 0, wx.LEFT, border=10)
        bxSizer.Add(tlesizer2, 0, wx.CENTER | wx.ALL, border=5)

        tlesizer3 = wx.BoxSizer(wx.HORIZONTAL)

        tlesizer3.Add(self.txtlbls[0], 0, wx.LEFT, border=15)
        tlesizer3.Add(self.txtbxs[0], 0, wx.LEFT, border=10)
        tlesizer3.Add(self.txtlbls[5], 0, wx.LEFT, border=15)
        tlesizer3.Add(self.txtbxs[5], 0, wx.LEFT, border=10)

        bxSizer.Add(tlesizer3, 0, wx.CENTER | wx.ALL, border=5)

        tlesizer4 = wx.BoxSizer(wx.HORIZONTAL)

        tlesizer4.Add(self.txtlbls[4], 0, wx.LEFT, border=20)
        tlesizer4.Add(self.txtbxs[4], 0, wx.LEFT, border=10)
        tlesizer4.Add(self.txtlbls[2], 0, wx.RIGHT, border=15)
        tlesizer4.Add(self.txtbxs[2], 0, wx.RIGHT, border=20)

        bxSizer.Add(tlesizer4, 0, wx.CENTER | wx.ALL, border=5)

        tlesizer5 = wx.BoxSizer(wx.HORIZONTAL)

        tlesizer5.Add(self.txtlbls[7], 0, wx.RIGHT, border=20)
        tlesizer5.Add(self.txtbxs[7], 0, wx.RIGHT, border=50)

        bxSizer.Add(tlesizer5, 0, wx.CENTER | wx.ALL, border=5)

        # form level controls
        fstsub = wx.Button(self, label='*<<')
        lstsub = wx.Button(self, label='>>*')
        nxtsub = wx.Button(self, label='>*')
        presub = wx.Button(self, label='*<')
        fstsub.SetForegroundColour('blue')
        lstsub.SetForegroundColour('blue')
        nxtsub.SetForegroundColour('blue')
        presub.SetForegroundColour('blue')
        fstsub.Bind(wx.EVT_BUTTON, self.OnMovefstsub)
        lstsub.Bind(wx.EVT_BUTTON, self.OnMovelstsub)
        nxtsub.Bind(wx.EVT_BUTTON, self.OnMovenxtsub)
        presub.Bind(wx.EVT_BUTTON, self.OnMovepresub)

        navbox = wx.BoxSizer(wx.HORIZONTAL)
        navbox.Add(fstsub, 0, wx.ALL | wx.ALIGN_LEFT, 5)
        navbox.Add(presub, 0, wx.ALL | wx.ALIGN_LEFT, 5)
        navbox.Add(nxtsub, 0, wx.ALL | wx.ALIGN_LEFT, 5)
        navbox.Add(lstsub, 0, wx.ALL | wx.ALIGN_LEFT, 5)

        bxSizer.Add(navbox, 0, wx.CENTER | wx.ALL, border=5)

        numboxsub = wx.BoxSizer(wx.HORIZONTAL)
        self.recnumsub1 = wx.StaticText(self, label='Record ',
                                        style=wx.ALIGN_LEFT)
        self.recnumsub1.SetForegroundColour('blue')

        self.recnumsub2 = wx.StaticText(self, label=str(self.sub_rec_num+1),
                                        style=wx.ALIGN_LEFT)
        self.recnumsub2.SetForegroundColour('blue')
        self.recnumsub3 = wx.StaticText(self, label='/ ' + str(len(self.data))
                                        + ' (MAX. 4)', style=wx.ALIGN_LEFT)
        self.recnumsub3.SetForegroundColour('blue')
        numboxsub.Add(self.recnumsub1, 0, wx.ALL | wx.ALIGN_LEFT, 5)
        numboxsub.Add(self.recnumsub2, 0, wx.ALL | wx.ALIGN_LEFT, 5)
        numboxsub.Add(self.recnumsub3, 0, wx.ALL | wx.ALIGN_LEFT, 5)

        bxSizer.Add(numboxsub, 0, wx.CENTER | wx.ALL, border=3)

        # Add buttons for form modifications
        self.subb2 = wx.Button(self, label="Add to Welding\nRequirement")
        self.Bind(wx.EVT_BUTTON, self.OnAddSubSpec, self.subb2)
        self.subb2.Disable()

        self.subb3 = wx.Button(self, label="Remove from\nWelding Requirement")
        self.Bind(wx.EVT_BUTTON, self.OnRmvSubSpec, self.subb3)
        navbox.Add((100, 30))
        navbox.Add(self.subb2, 1, wx.CENTER, 10)
        navbox.Add(self.subb3, 1, wx.CENTER, 10)

        self.Sizer.Add(bxSizer, 0, wx.CENTER, 10)

        # form level controls
        fst = wx.Button(self, label='<<')
        lst = wx.Button(self, label='>>')
        nxt = wx.Button(self, label='>')
        pre = wx.Button(self, label='<')
        fst.Bind(wx.EVT_BUTTON, self.OnMovefst)
        lst.Bind(wx.EVT_BUTTON, self.OnMovelst)
        nxt.Bind(wx.EVT_BUTTON, self.OnMovenxt)
        pre.Bind(wx.EVT_BUTTON, self.OnMovepre)

        navbox = wx.BoxSizer(wx.HORIZONTAL)
        navbox.Add(fst, 0, wx.ALL | wx.ALIGN_LEFT, 2)
        navbox.Add(pre, 0, wx.ALL | wx.ALIGN_LEFT, 2)
        navbox.Add(nxt, 0, wx.ALL | wx.ALIGN_LEFT, 2)
        navbox.Add(lst, 0, wx.ALL | wx.ALIGN_LEFT, 2)
        navbox.Add((15, 10))

        numbox = wx.BoxSizer(wx.HORIZONTAL)
        self.recnum1 = wx.StaticText(self, label='Record ',
                                     style=wx.ALIGN_LEFT)
        self.recnum1.SetForegroundColour((255, 0, 0))

        self.recnum2 = wx.StaticText(self, label=str(self.rec_num+1),
                                     style=wx.ALIGN_LEFT)
        self.recnum2.SetForegroundColour((255, 0, 0))
        self.recnum3 = wx.StaticText(self, label='/ '+str(len(self.data)),
                                     style=wx.ALIGN_LEFT)
        self.recnum3.SetForegroundColour((255, 0, 0))
        numbox.Add(self.recnum1, 0, wx.ALL | wx.ALIGN_CENTER, 5)
        numbox.Add(self.recnum2, 0, wx.ALL | wx.ALIGN_CENTER, 5)
        numbox.Add(self.recnum3, 0, wx.ALL | wx.ALIGN_CENTER, 5)

        self.Sizer.Add(navbox, 0, wx.ALIGN_CENTER | wx.TOP, 10)
        self.Sizer.Add(numbox, 0, wx.ALIGN_CENTER)

        self.b3 = wx.Button(self, label="Delete\nSpecification")
        if self.ComdPrtyID is not None:
            if StartQry is None:
                self.nodata_txt = (
                    'The ' + self.frmtitle +
                    ' have not been setup for this Commodity Property')
            else:
                if self.PipeCode is None:
                    query = (
                        '''SELECT * FROM PipeMaterialSpec WHERE
                         Material_Spec_ID = '''
                        + str(self.PipeMtrSpec))
                    self.MtrSpc = Dbase().Dsqldata(query)[0][1]
                    self.nodata_txt = ('Weld Requirements for ' +
                                       self.ComCode + '-' + self.MtrSpc)
                else:
                    self.nodata_txt = ('Weld Requirements for ' +
                                       self.PipeCode)
            self.lbl_nodata.SetLabel(self.nodata_txt)
            self.b6 = wx.Button(self, size=(120, -1), label="Show All\nItems")
            self.b3.SetLabel('Delete From\nCommodity')
            if StartQry is None:
                self.b3.Disable()
            self.Bind(wx.EVT_BUTTON, self.OnAddComd, self.b6)
            navbox.Add(self.b6, 0, wx.CENTER)
        self.Bind(wx.EVT_BUTTON, self.OnDelete, self.b3)
        self.b7 = wx.Button(self, label="Print\nReport")
        self.Bind(wx.EVT_BUTTON, self.PrintFile, self.b7)
        self.b4 = wx.Button(self, label="Exit")
        self.Bind(wx.EVT_BUTTON, self.OnClose, self.b4)

        navbox.Add(self.b3, 0, wx.CENTER)
        navbox.Add((12, 25))
        navbox.Add(self.b7, 0, wx.CENTER)
        navbox.Add((12, 25))
        navbox.Add(self.b4, 0, wx.CENTER)

        self.FillScreen()
        self.FillSubScrn()

        # add following 5 lines for child parent form
        self.CenterOnParent()
        self.GetParent().Enable(False)
        self.Show(True)
        self.__eventLoop = wx.GUIEventLoop()
        self.__eventLoop.Run()

    def BxBld(self):
        self.txts = []
        self.cmbs = []
        self.txtbxs = []
        self.txtlbls = []

        font1 = wx.Font(12, wx.DEFAULT, wx.NORMAL, wx.FONTWEIGHT_BOLD)

        items = self.hints_tbls
        # order of information text label,table name,combo width,
        # combo box true = 1,table ID col name,column name shown in box

        for item in items:
            if item[3] == 0:
                txtbx = wx.TextCtrl(self, size=item[2], value='',
                                    style=wx.TE_LEFT | wx.TE_READONLY)
                self.Bind(wx.EVT_TEXT, self.OnSelect2, txtbx)
                self.txtbxs.append(txtbx)
                txtlbl = wx.StaticText(self, label=item[0],
                                       style=wx.ALIGN_LEFT)
                txtlbl.SetForegroundColour((255, 0, 0))
                self.txtlbls.append(txtlbl)
            elif item[3] == 1:
                cmb = wx.ComboCtrl(self, size=item[2])
                self.Bind(wx.EVT_TEXT, self.OnSelect2, cmb)
                self.showcol = 1
                cmb.SetPopupControl(ListCtrlComboPopup(
                    item[1], showcol=self.showcol, lctrls=self.lctrls))
                self.cmbs.append(cmb)

                txt = wx.StaticText(self, label=item[0], style=wx.ALIGN_LEFT)
                txt.SetForegroundColour((255, 0, 0))
                self.txts.append(txt)

                addbtn = wx.Button(self, label='+', size=(35, -1))
                addbtn.SetForegroundColour((255, 0, 0))
                addbtn.SetFont(font1)
                self.Bind(wx.EVT_BUTTON, self.OnAdd1, addbtn)
                self.addbtns.append(addbtn)

    def PrintFile(self, evt):
        import Report_Weld

        Colnames = [('Process', 'Notes'),
                    ('Weld\nProcedure', 'Process', 'Weld Filler',
                     'Filler\nGroup', 'Approved\nThickness', 'Material',
                     'Position', 'Position\nDescription',
                     'Welder Qualification\nCertificate Notes', 'Notes')]

        Colwdths = [(20, 40),
                    (10, 15, 8, 8, 8, 8, 10, 15, 15, 15)]

        if self.data == []:
            NoData = wx.MessageDialog(
                None, 'No Data to Print', 'Error', wx.OK | wx.ICON_EXCLAMATION)
            NoData.ShowModal()
            return

        saveDialog = wx.FileDialog(self, message='Save Report as PDF.',
                                   wildcard='PDF (*.pdf)|*.pdf',
                                   style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT)

        if saveDialog.ShowModal() == wx.ID_CANCEL:
            filename = ''

        filename = saveDialog.GetPath()
        if filename.find(".pdf") == -1:
            filename = filename + '.pdf'

        saveDialog.Destroy()

        rptdata = self.Reportdata()

        ttl = self.frmtitle
        if len(self.data) == 1:
            # specify a title for the report if table name
            # is not to be used
            if self.PipeCode is None:
                ttl = self.text1.GetValue()
            else:
                ttl = self.PipeCode
            ttl = self.frmtitle + ' for ' + ttl

        Report_Weld.Report(rptdata, Colnames,
                           Colwdths, filename, ttl).create_pdf()

    def Reportdata(self):
        rptdata = []

        for seg in self.data:
            rptdata3 = []
            rptdata1 = []
            qry = ('SELECT Process FROM WeldProcesses WHERE ProcessID = ' +
                   str(seg[1]))
            rptdata1.append(Dbase().Dsqldata(qry)[0][0])
            rptdata1.append(seg[6])

            rptdata3.append(tuple(rptdata1))

            for m in seg[2:6]:
                if m is not None:
                    qry = ('SELECT * FROM WeldProcedures WHERE ProcedureID = '
                           + str(m))
                    suddt = Dbase().Dsqldata(qry)
                    rptdata2 = []
                    rptdata2.append(suddt[0][1])
                    qry = ('''SELECT Process FROM WeldProcesses
                            WHERE ProcessID = ''' + str(suddt[0][5]))
                    rptdata2.append(Dbase().Dsqldata(qry)[0][0])
                    qry = ('SELECT Metal_Spec FROM WeldFiller WHERE ID = ' +
                           str(suddt[0][2]))
                    rptdata2.append(Dbase().Dsqldata(qry)[0][0])
                    qry = ('''SELECT Filler_Group FROM WeldFillerGroup
                            WHERE ID = ''' + str(suddt[0][7]))
                    rptdata2.append(Dbase().Dsqldata(qry)[0][0])
                    qry = ('''SELECT Thickness FROM WeldQualifyThickness
                            WHERE ID = ''' + str(suddt[0][3]))
                    rptdata2.append(Dbase().Dsqldata(qry)[0][0])
                    qry = ('''SELECT Material_Group FROM WeldMaterialGroup
                            WHERE ID = ''' + str(suddt[0][6]))
                    rptdata2.append(Dbase().Dsqldata(qry)[0][0])
                    qry = ('''SELECT Position, Description FROM WeldQualifyPosition
                            WHERE ID = ''' + str(suddt[0][8]))
                    tmp = Dbase().Dsqldata(qry)
                    rptdata2.append(tmp[0][0])
                    rptdata2.append(tmp[0][1])

                    rptdata2.append(suddt[0][4])
                    rptdata2.append(suddt[0][9])
                    rptdata3.append(tuple(rptdata2))

            rptdata.append(rptdata3)

        return rptdata

    def OnClose(self, evt):
        # check to make sure there are processes saved for the weld requirement
        # if there are none then remove the weld requirement from the pipespec
        # then remove the weld requirement ID from the WeldRequirements table
        if self.recrd != []:
            query = ('SELECT Procedure1, Procedure2, Procedure3,\
                     Procedure4 FROM '
                     + self.tblname + ' WHERE ' + self.field + '=' +
                     str(self.recrd[0]))
            procedures = Dbase().Dsqldata(query)
            if procedures == [(None, None, None, None)]:
                if self.ComdPrtyID is not None:
                    self.ChgSpecID()
                Dbase().TblDelete(self.tblname, self.recrd[0], self.field)

        # add for child form
        self.GetParent().Enable(True)
        self.__eventLoop.Exit()
        self.Destroy()

    def OnMovefstsub(self, evt):
        self.sub_rec_num = 0
        self.FillSubScrn()

    def OnMovelstsub(self, evt):
        if len(self.subdata) == 0:
            return
        self.sub_rec_num = len(self.subdata)-1
        self.FillSubScrn()

    def OnMovenxtsub(self, evt):
        if len(self.subdata) == 0:
            return
        self.sub_rec_num += 1
        if self.sub_rec_num == len(self.subdata):
            self.sub_rec_num = 0
        self.FillSubScrn()

    def OnMovepresub(self, evt):
        if len(self.subdata) == 0:
            return
        self.sub_rec_num -= 1
        if self.sub_rec_num < 0:
            self.sub_rec_num = len(self.subdata)-1
        self.FillSubScrn()

    def OnMovefst(self, evt):
        self.rec_num = 0
        self.FillScreen()
        self.sub_rec_num = 0
        self.FillSubScrn()

    def OnMovelst(self, evt):
        if len(self.data) == 0:
            return
        self.rec_num = len(self.data)-1
        self.FillScreen()
        self.sub_rec_num = 0
        self.FillSubScrn()

    def OnMovenxt(self, evt):
        if len(self.data) == 0:
            return
        self.rec_num += 1
        if self.rec_num == len(self.data):
            self.rec_num = 0
        self.FillScreen()
        self.sub_rec_num = 0
        self.FillSubScrn()

    def OnMovepre(self, evt):
        if len(self.data) == 0:
            return
        self.rec_num -= 1
        if self.rec_num < 0:
            self.rec_num = len(self.data)-1
        self.FillScreen()
        self.sub_rec_num = 0
        self.FillSubScrn()

    def FillScreen(self):
        # one record string of data from the main table
        if len(self.data) == 0:
            self.recnum2.SetLabel(str(self.rec_num))
            self.text1.ChangeValue('')
            self.cmbProces.ChangeValue('')
            return
        else:
            self.recrd = self.data[self.rec_num]

        self.text1.ChangeValue(str(self.recrd[6]))

        query = ('SELECT Process FROM WeldProcesses WHERE ProcessID = '
                 + str(self.recrd[1]))
        cmbval = Dbase().Dsqldata(query)
        self.cmbProces.ChangeValue(cmbval[0][0])

        self.recnum2.SetLabel(str(self.rec_num+1))

    def FillSubScrn(self):
        self.subdata = []
        # if there is a weld requirement then look up any of the weld
        # procedures attached to it range of procedureID fields is 2 to 5
        if self.data != []:
            for n in range(2, 6):
                SubSQL = ('SELECT * FROM WeldProcedures WHERE ProcedureID IN ('
                          + str(self.recrd[n]) + ')')
                if self.recrd[n] is not None:
                    if Dbase().Dsqldata(SubSQL) != []:
                        self.subdata.append(Dbase().Dsqldata(SubSQL)[0])

        if self.subdata != []:
            subrecrd = self.subdata[self.sub_rec_num]
        else:
            for grp in [(0, 0, 2), (1, 1, 3), (2, 2, 4), (3, 3, 5), (4, 4, 6),
                        (5, 5, 7), (6, 6, 8), (7, 7, 9)]:
                # gets build info from combo andtext box
                subtbl = self.hints_tbls[grp[0]]
                if len(subtbl) > 4:
                    if subtbl[3] == 1:  # specifies this is a combo box
                        self.cmbs[grp[1]].ChangeValue('')
                    if subtbl[3] == 0:   # specifies this is a text box
                        self.txtbxs[grp[1]].ChangeValue('')
                else:
                    self.txtbxs[2].ChangeValue('')
                    self.txtbxs[7].ChangeValue('')
            self.cmbProced.ChangeValue('')
            self.recnumsub2.SetLabel('0')
            self.recnumsub3.SetLabel('/ 0 ' + ' (Maximum 4)')
            return

        # numbers represent (record in hints_tbls,
        # cmbbxs/txtbx number,ID in subrecrd data)
        for grp in [(0, 0, 2), (1, 1, 3), (2, 2, 4), (3, 3, 5), (4, 4, 6),
                    (5, 5, 7), (6, 6, 8), (7, 7, 9)]:
            # gets build info from combo and text box
            subtbl = self.hints_tbls[grp[0]]
            if len(subtbl) > 4:
                qry = ('SELECT ' + subtbl[5] + ' FROM ' + subtbl[1] +
                       ' WHERE ' + subtbl[4] + ' = ' + str(subrecrd[grp[2]]))

                recrd = Dbase().Dsqldata(qry)
                if recrd != []:
                    recrd = recrd[0][0]

                if subtbl[3] == 1:  # specifies this is a combo box
                    self.cmbs[grp[1]].ChangeValue(recrd)
                if subtbl[3] == 0:   # specifies this is a text box
                    self.txtbxs[grp[1]].ChangeValue(recrd)
            else:
                if subrecrd[4] is not None:
                    self.txtbxs[2].ChangeValue(subrecrd[4])
                else:
                    self.txtbxs[2].ChangeValue('')

                if subrecrd[9] is not None:
                    self.txtbxs[7].ChangeValue(subrecrd[9])
                else:
                    self.txtbxs[7].ChangeValue('')

        self.cmbProced.ChangeValue(subrecrd[1])
        self.subb2.Disable()
        self.recnumsub2.SetLabel(str(self.sub_rec_num+1))
        self.recnumsub3.SetLabel('/ ' + str(len(self.subdata)) +
                                 ' (Maximum 4)')

    def OnAdd(self, evt):
        # call the specific form to add to the list of the drop down box
        btn = evt.GetEventObject()
        callbtn = self.btnDict[btn]

        if callbtn == 1:
            tbl = 'WeldProcedures'
            BldProced(self, tbl)
        elif callbtn == 0:
            tbl = 'WeldProcesses'
            CmbLst1(self, tbl)

        self.ReFillList(tbl, callbtn)

    def OnAddComd(self, evt):
        if self.b6.GetLabel() == 'Show All\nItems':
            if self.MainSQL == '':
                self.MainSQL = 'SELECT * FROM ' + self.tblname
                self.data = Dbase().Dsqldata(self.MainSQL)
            else:
                self.MainSQL = self.MainSQL[:self.MainSQL.find('WHERE')+1]
                self.data = Dbase().Dsqldata(self.MainSQL)
            self.lbl_nodata.SetLabel('   ')
            self.b6.SetLabel("Add Item to\nCommodity")
            self.b6.Enable()
            self.FillScreen()
            self.recnum3.SetLabel('/ '+str(len(self.data)))
            self.FillSubScrn()
        else:
            self.AddComd()
            if self.PipeCode is None:
                query = (
                    'SELECT * FROM PipeMaterialSpec WHERE Material_Spec_ID = '
                    + str(self.PipeMtrSpec))
                self.MtrSpc = Dbase().Dsqldata(query)[0][1]
                self.nodata_txt = ('Weld Requirements for ' +
                                   self.ComCode + '-' + self.MtrSpc)
            else:
                self.nodata_txt = ('Weld Requirements for ' +
                                   self.PipeCode)
            self.lbl_nodata.SetLabel(self.nodata_txt)
            self.b6.Enable()
            self.b3.Enable()
            self.b6.SetLabel("Show All\nItems")

    # link this ID to the commodity property
    def AddComd(self):

        query = ('SELECT ' + self.tblname + '''_ID FROM PipeSpecification WHERE
                  Commodity_Property_ID = ''' + str(self.ComdPrtyID))
        StartQry = Dbase().Dsqldata(query)
        if StartQry == []:
            ValueList = []
            New_ID = cursr.execute(
                "SELECT MAX(Pipe_Spec_ID) FROM PipeSpecification").\
                fetchone()
            if New_ID[0] is None:
                Max_ID = '1'
            else:
                Max_ID = str(New_ID[0]+1)
            colinfo = Dbase().Dcolinfo('PipeSpecification')
            for n in range(0, len(colinfo)-2):
                ValueList.append(None)

            num_vals = ('?,'*len(colinfo))[:-1]
            ValueList.insert(0, Max_ID)
            ValueList.insert(1, str(self.ComdPrtyID))

            UpQuery = "INSERT INTO PipeSpecification VALUES (" + num_vals + ")"
            Dbase().TblEdit(UpQuery, ValueList)
            StartQry = Max_ID
        else:
            StartQry = str(StartQry[0][0])

        cmd_addID = self.data[self.rec_num][0]
        self.ChgSpecID(cmd_addID)

        self.MainSQL = ('SELECT * FROM ' + self.tblname + ' WHERE ' +
                        self.pkcol_name + ' = ' + str(cmd_addID))
        self.data = Dbase().Dsqldata(self.MainSQL)

        self.rec_num = 0
        self.FillScreen()
        self.recnum3.SetLabel('/ '+str(len(self.data)))

    def EditSubTbl(self):
        query = ('SELECT Procedure1, Procedure2, Procedure3, Procedure4 FROM '
                 + self.tblname + ' WHERE ' + self.field + '=' +
                 str(self.recrd[0]))
        procedures = Dbase().Dsqldata(query)

        if None not in procedures[0]:
            mesg = ('''There are presently already\nfour procedures.
                      Either modify an existing\nor delete one before saving\na
                     new procedure''')
            wx.MessageBox(mesg, "Exceeded Allowed Number of Procedures",
                          wx.OK | wx.ICON_INFORMATION)
            return

        if len(self.nodata_txt) < 35:
            txt = ('This change will be reflected in all' +
                   '\nreferences to this Welding Requirement.\n' +
                   'Do you wish to proceed')
            new_data = wx.MessageDialog(
                self, txt, 'Universal Data Change',
                wx.OK | wx.CANCEL | wx.ICON_INFORMATION)

            response = new_data.ShowModal()
            new_data.Destroy()
        else:
            response = wx.ID_OK

        if response == wx.ID_OK:
            query = ('''SELECT ProcedureID FROM WeldProcedures WHERE
                     Procedure LIKE '%''' + self.cmbProced.GetValue() +
                     "%' COLLATE NOCASE")
            # this is the ID value for the linked field in the foreign table
            procedID = Dbase().Dsqldata(query)[0][0]
            n = 0
            for i in procedures[0]:
                if i is None:
                    query = ("UPDATE " + self.tblname + " SET Procedure" +
                             str(n+1) + " = ? WHERE " + self.field + " = ?")
                    update_list = (str(procedID), str(self.recrd[0]))
                    Dbase().TblEdit(query, update_list)
                    break
                n += 1

            self.data = Dbase().Dsqldata(self.MainSQL)
            self.FillScreen()
            self.FillSubScrn()

    def ReFillList(self, cmbtbl, boxnum):
        self.lc = self.lctrls[boxnum]
        self.lc.DeleteAllItems()
        index = 0

        if cmbtbl == 'WeldProcedures':
            self.showcol = 1
            query = 'SELECT WeldProcedures.Procedure FROM WeldProcedures'
        else:
            query = 'SELECT * FROM "' + cmbtbl + '"'

        for values in Dbase().Dsqldata(query):
            col = 0
            for value in values:
                if col == 0:
                    self.lc.InsertItem(index, str(value))
                else:
                    self.lc.SetItem(index, col, str(value))
                col += 1
            index += 1

    def OnRmvSubSpec(self, evt):
        if len(self.subdata) == 1:
            txt = ('There must be at least one Process attached' +
                   ' to the Weld Requirement.\nYou may only change' +
                   ' the exiting Weld Process or select a new Weld Requiremnt')
            no_delete = wx.MessageDialog(self, txt, 'Unable to Delete',
                                         wx.OK | wx.ICON_INFORMATION)
            no_delete.ShowModal()
            return

        else:
            txt = ('This change will be reflected in all' +
                   '\nreferences to this Welding Requirement.\n' +
                   'Do you wish to proceed')
            new_data = wx.MessageDialog(
                self, txt, 'Universal Data Change',
                wx.OK | wx.CANCEL | wx.ICON_INFORMATION)

            response = new_data.ShowModal()
            new_data.Destroy()
            if response == wx.ID_OK:
                rqrdID = self.data[self.rec_num][0]

                # use the selected weld procedure value to determine
                # which Procedure column you are at
                # this determines the procedureID for the item
                query = (
                    '''SELECT ProcedureID FROM WeldProcedures WHERE
                     Procedure LIKE '%''' +
                    self.cmbProced.GetValue() + "%' COLLATE NOCASE")
                # this is the ID value for the linked field
                # in the foreign table
                procedID = Dbase().Dsqldata(query)[0][0]
                # get the list of the procedure IDs
                query = ('SELECT Procedure1, Procedure2, Procedure3,\
                         Procedure4 FROM '
                         + self.tblname + ' WHERE ' + self.field + '=' +
                         str(self.recrd[0]))
                procedures = Dbase().Dsqldata(query)
                # this determines the actual procedure column number for
                # the selected procedure ID
                n = 1
                for item in procedures[0]:
                    if procedID == item:
                        proced_num = n
                        break
                    n += 1
                # now use the correct column number to delete the specified
                # procedure from the weld requirement
                qry = ('UPDATE ' + self.tblname + ' SET Procedure' +
                       str(proced_num) + '=? WHERE ' + self.field + ' = '
                       + str(rqrdID))
                val = [None]
                Dbase().TblEdit(qry, val)
                self.data = Dbase().Dsqldata(self.MainSQL)
                self.FillScreen()
                self.sub_rec_num = 0
                self.FillSubScrn()

    def OnAddSubSpec(self, evt):
        self.EditSubTbl()
        self.subb2.Disable()

    def OnSelect1(self, evt):
        note_chg = wx.MessageDialog(
            self, '''This change will be reflected in all\nreferences\
            to this Welding Requirement.''',
            'Universal Data Change',
            wx.OK | wx.CANCEL | wx.ICON_INFORMATION)

        response = note_chg.ShowModal()
        note_chg.Destroy()

        if response == wx.ID_OK:
            rqrdID = self.data[self.rec_num][0]
            qry = ('UPDATE ' + self.tblname + ' SET WeldNote =? WHERE ' +
                   self.field + ' = ' + str(rqrdID))
            val = [self.text1.GetValue()]
            Dbase().TblEdit(qry, val)
            self.data = Dbase().Dsqldata(self.MainSQL)
            self.FillScreen()
        elif response == wx.ID_CANCEL:
            self.FillScreen()

    def OnCmbOpen(self, evt):
        self.startstrg = self.cmbProces.GetValue()

    def OnCmbExit(self, evt):
        self.TextChg()

    def TextChg(self):
        if self.cmbProces.GetValue() != self.startstrg:
            Update_step = 3

            choice1 = ('1) Save this as a new ' +
                       self.frmtitle + ' Specification')
            choice2 = ('2) Update the existing ' + self.frmtitle +
                       ' Specification with this data')
            txt1 = ('NOTE: Updating this information will be reflected' +
                    ' in all associated ' + self.frmtitle)
            txt2 = (' Specifications!\n\nRecommendation is to save as a new' +
                    ' specification.\n\n\tHow do you want to proceed?')

            if self.ComdPrtyID is None:
                # if this is a not commodity related change
                # Make a selection as to whether the record
                # is to be a new or an update valve use a
                # SingleChioce dialog to determine if data
                # is new record or edited record

                # use a SingleChioce dialog to determine if data
                # is new record or edited record
                Update_Dialog = wx.SingleChoiceDialog(
                    self, txt1+txt2, 'Information Has Changed',
                    [choice1, choice2], style=wx.CHOICEDLG_STYLE)

                if Update_Dialog.ShowModal() == wx.ID_OK:
                    Update_step = Update_Dialog.GetSelection()

                Update_Dialog.Destroy()

                if Update_step == 0:  # enter new record
                    if self.AddRequire():
                        msg = ('The new Welding Requirement has been saved.' +
                               '\nYou can now modify, delete or add new Weld' +
                               ' Procedures\nby selecting them from the ' +
                               ' Procedure drop down list\n or selecting ' +
                               'the red + next to it and\nadding new Weld' +
                               ' Procedures then selecting them\nfrom the ' +
                               'updated drop down list')
                        wx.MessageBox(msg, "Completeing Welding Requirements",
                                      wx.OK | wx.ICON_INFORMATION)

                elif Update_step == 1:  # update edited record
                    self.ModProced()

                elif Update_step == 3:  # cancel changes
                    self.FillScreen()

            else:  # this is a commodity related change
                # before preparing to save any data confirm this is not
                # an empty weld requirement
                # check to make sure there are processes saved for the
                # weld requirement if there are none then remove the weld
                # requirement from the pipespec then remove the weld
                # requirement ID from the WeldRequirements table
                if self.recrd != []:
                    query = ('SELECT Procedure1, Procedure2, Procedure3,\
                             Procedure4 FROM '
                             + self.tblname + ' WHERE ' + self.field + '=' +
                             str(self.recrd[0]))
                    procedures = Dbase().Dsqldata(query)
                    if procedures == [(None, None, None, None)]:
                        self.ChgSpecID()
                        Dbase().TblDelete(self.tblname,
                                          self.recrd[0],
                                          self.field)

                choice1 = choice1 + ' for this commodity?'
                choice2 = choice2 + ' for this commodity?'
                if len(self.subdata) == 0:
                    Update_step = 0
                else:
                    Update_Dialog = wx.SingleChoiceDialog(
                        self, txt1+txt2, 'Information Has Changed',
                        [choice1, choice2], style=wx.CHOICEDLG_STYLE)

                    if Update_Dialog.ShowModal() == wx.ID_OK:
                        Update_step = Update_Dialog.GetSelection()

                    Update_Dialog.Destroy()

                if Update_step == 0:  # enter new record
                    # the AddRequire module will add the weld Procedure to the
                    # WeldRequirement table and return the new ID number
                    # the ChgSpecID will add the new ID number to the
                    # PipeSepcification table under the corresponding
                    # Commodity Property
                    self.ChgSpecID(self.AddRequire())
                    msg = ('The new Welding Requirement has been saved.' +
                           '\nYou can now modify, delete or add new Weld' +
                           ' Procedures\nby selecting them from the ' +
                           ' Procedure drop down list\n or selecting ' +
                           'the red + next to it and\nadding new Weld' +
                           ' Procedures then selecting them\nfrom the ' +
                           'updated drop down list')
                    wx.MessageBox(msg, "Completeing Welding Requirements",
                                  wx.OK | wx.ICON_INFORMATION)

                elif Update_step == 1:  # update edited record
                    self.ModProced()

                elif Update_step == 3:  # cancel changes
                    self.FillScreen()

    def ModProced(self):
        query = ("SELECT ProcessID FROM WeldProcesses WHERE Process LIKE '%" +
                 self.cmbProces.GetValue() + "%' COLLATE NOCASE")
        val = str(Dbase().Dsqldata(query)[0][0])
        rqrdID = self.data[self.rec_num][0]
        qry = ('UPDATE ' + self.tblname + ' SET ProcessID =? WHERE ' +
               self.field + '= ' + str(rqrdID))
        Dbase().TblEdit(qry, val)
        self.data = Dbase().Dsqldata(self.MainSQL)
        self.FillScreen()

    def AddRequire(self):    # add a new welding requirement
        colinfo = Dbase().Dcolinfo(self.tblname)
        ValueList = [None for i in range(0, len(colinfo))]
        num_vals = ('?,'*len(colinfo))[:-1]

        # get the new maximum ID value from the Weld requirements table
        New_ID = cursr.execute("SELECT MAX(" + self.field + ") FROM "
                               + self.tblname).fetchone()
        if New_ID[0] is None:
            Max_ID = '1'
        else:
            Max_ID = str(New_ID[0]+1)
        ValueList[0] = Max_ID
        # get the new process ID from the updated combobox
        query = ("SELECT ProcessID FROM WeldProcesses WHERE Process LIKE '%" +
                 self.cmbProces.GetValue() + "%' COLLATE NOCASE")
        ValueList[1] = Dbase().Dsqldata(query)[0][0]
        # save the information into the Weld requirements table as new record
        UpQuery = 'INSERT INTO ' + self.tblname + ' VALUES (' + num_vals + ')'
        Dbase().TblEdit(UpQuery, ValueList)

        if self.ComdPrtyID is not None:
            # save the new Weld requirement ID to the Pipe Specification table
            self.ChgSpecID(Max_ID)

            # requery the pipe specification table to be able to
            # populate the form
            query = ('SELECT ' + self.tblname +
                     '''_ID FROM PipeSpecification WHERE
                      Commodity_Property_ID = ''' +
                     str(self.ComdPrtyID))
            StartQry = Dbase().Dsqldata(query)[0][0]
            if StartQry is not None:
                self.MainSQL = ('SELECT * FROM ' + self.tblname + ' WHERE ' +
                                self.tblname + 'ID = ' + str(StartQry))
                self.data = Dbase().Dsqldata(self.MainSQL)
            self.rec_num = len(self.data)-1
            self.recnum3.SetLabel('/ '+str(len(self.data)))
        else:
            self.MainSQL = 'SELECT * FROM ' + self.tblname
            self.data = Dbase().Dsqldata(self.MainSQL)
            self.rec_num = len(self.data)-1
            self.recnum3.SetLabel('/ '+str(len(self.data)))

        self.FillScreen()

        for grp in [(0, 0, 2), (1, 1, 3), (2, 2, 4), (3, 3, 5),
                    (4, 4, 6), (5, 5, 7), (6, 6, 8), (7, 7, 9)]:
            # gets build info from combo / text box
            subtbl = self.hints_tbls[grp[0]]
            if len(subtbl) > 4:
                if subtbl[3] == 1:  # specifies this is a combo box
                    self.cmbs[grp[1]].ChangeValue('')
                if subtbl[3] == 0:   # specifies this is a text box
                    self.txtbxs[grp[1]].ChangeValue('')
            else:
                self.txtbxs[2].ChangeValue('')
                self.txtbxs[7].ChangeValue('')

        self.sub_rec_num = 0
        self.FillSubScrn()

        return Max_ID

    def OnSelectProced(self, evt):
        query = ("SELECT * FROM WeldProcedures WHERE Procedure LIKE '%" +
                 self.cmbProced.GetValue() + "%' COLLATE NOCASE")
        # this is the ID value for the linked field in the foreign table
        subrecrd = Dbase().Dsqldata(query)[0]
        # numbers represent (record in hints_tbls,cmbbxs/txtbx
        # number,ID in subrecrd data)
        for grp in [(0, 0, 2), (1, 1, 3), (2, 2, 4), (3, 3, 5), (4, 4, 6),
                    (5, 5, 7), (6, 6, 8), (7, 7, 9)]:
            # gets build info from combo / text box
            subtbl = self.hints_tbls[grp[0]]
            if len(subtbl) > 4:
                qry = ('SELECT ' + subtbl[5] + ' FROM ' + subtbl[1] +
                       ' WHERE ' + subtbl[4] + ' = ' + str(subrecrd[grp[2]]))
                recrd = Dbase().Dsqldata(qry)
                if recrd != []:
                    recrd = recrd[0][0]
                if subtbl[3] == 1:  # specifies this is a combo box
                    self.cmbs[grp[1]].ChangeValue(recrd)
                if subtbl[3] == 0:   # specifies this is a text box
                    self.txtbxs[grp[1]].ChangeValue(recrd)
            else:
                if subrecrd[4] is not None:
                    self.txtbxs[2].ChangeValue(subrecrd[4])
                else:
                    self.txtbxs[2].ChangeValue('')

                if subrecrd[9] is not None:
                    self.txtbxs[7].ChangeValue(subrecrd[9])
                else:
                    self.txtbxs[7].ChangeValue('')
        self.subb2.Enable()

    def OnSelect2(self, evt):
        self.subb2.Enable()

    def OnDelete(self, evt):
        recrd = self.data[self.rec_num][0]
        if self.ComdPrtyID is None:
            try:
                Dbase().TblDelete(self.tblname, recrd, self.field)
                self.MainSQL = 'SELECT * FROM ' + self.tblname
                self.data = Dbase().Dsqldata(self.MainSQL)
                self.rec_num -= 1
                if self.rec_num < 0:
                    self.rec_num = len(self.data)-1
                self.FillScreen()
                self.FillSubScrn()
                self.recnum3.SetLabel('/ '+str(len(self.data)))

            except sqlite3.IntegrityError:
                wx.MessageBox("This Record is associated"
                              " with\nother tables and cannot be deleted!",
                              "Cannot Delete",
                              wx.OK | wx.ICON_INFORMATION)
        else:
            self.ChgSpecID()
            self.data = []
            self.FillScreen()
            self.subdata = []
            self.FillSubScrn()
            self.b3.Disable()
            self.recnum3.SetLabel('/ '+str(len(self.data)))
            self.nodata_txt = (
                            'The ' + self.frmtitle +
                            ' have not been setup for this Commodity Property')
            self.lbl_nodata.SetLabel(self.nodata_txt)

    def ChgSpecID(self, ID=None):
        UpQuery = ('UPDATE PipeSpecification SET ' + self.tblname +
                   '_ID =?  WHERE Commodity_Property_ID = ' +
                   str(self.ComdPrtyID))
        Dbase().TblEdit(UpQuery, [ID])


class BldProced(wx.Frame):  # Sub form
    def __init__(self, parent, tblname):
        super(BldProced, self).__init__(parent, title="Welding Procedure",
                                        size=(850, 460))

        self.Bind(wx.EVT_CLOSE, self.Close)

        self.parent = parent
        self.rec_num = 0
        self.sub_rec_num = 0
        self.lctrls = []
        self.subdata = []
        self.tblname = tblname

        SubSQL = ('SELECT * FROM WeldProcedures')
        self.subdata = Dbase().Dsqldata(SubSQL)
        self.InitUI()

    def InitUI(self):
        # order of information text label,table name,combo width,
        # combo box true = 1,table ID col name,column name shown in box
        self.hints_tbls = [
            ('See WPS\nSheet(s)', '', (200, -1), 0),
            ('Filler', 'WeldFiller', (325, -1), 1, 'ID', 'Metal_Spec'),
            ('Weld Thickness\nQualification', 'WeldQualifyThickness',
             (150, -1), 1, 'ID', 'Thickness'),
            ('Welder Qualification\nCertificate to Show', '',
             (350, 60), 0),
            ('Weld\nProcess', 'WeldProcessList', (250, -1), 1,
             'ID', 'Process'),
            ('Material\nGroup', 'WeldMaterialGroup', (115, -1), 1,
             'ID', 'Material_Group'),
            ('Filler Metal\nGroup Number', 'WeldFillerGroup', (175, -1), 1,
             'ID', 'Filler_Group'),
            ('Position\nQualified', 'WeldQualifyPosition', (100, -1), 1,
             'ID', 'Position'),
            ('Special Notes:', '', (500, 60), 0)]

        self.BxBld()

        self.btnDict = {self.addbtns[i]: i for i in range(len(self.addbtns))}

        self.Sizer = wx.BoxSizer(wx.VERTICAL)

        prcsbxszr = wx.StaticBox(self, -1, '', style=wx.ALIGN_CENTER)
        bxSizer = wx.StaticBoxSizer(prcsbxszr, wx.VERTICAL)

        tlesizer1 = wx.BoxSizer(wx.HORIZONTAL)

        tlesizer1.Add(self.txtlbls[0], 0, wx.LEFT, border=15)
        tlesizer1.Add(self.txtbxs[0], 0, wx.LEFT, border=15)
        tlesizer1.Add(self.txts[5], 0, wx.LEFT, border=25)
        tlesizer1.Add(self.cmbs[5], 0, wx.LEFT, border=10)
        tlesizer1.Add(self.addbtns[5], 0, wx.RIGHT, border=15)

        bxSizer.Add(tlesizer1, 0, wx.CENTER | wx.ALL, border=10)

        tlesizer2 = wx.BoxSizer(wx.HORIZONTAL)

        tlesizer2.Add(self.txts[2], 0, wx.LEFT, border=15)
        tlesizer2.Add(self.cmbs[2], 0, wx.LEFT, border=10)
        tlesizer2.Add(self.addbtns[2], 0, wx.RIGHT, border=15)
        tlesizer2.Add(self.txts[1], 0, wx.LEFT, border=25)
        tlesizer2.Add(self.cmbs[1], 0, wx.LEFT, border=10)
        tlesizer2.Add(self.addbtns[1], 0, wx.RIGHT, border=15)
        bxSizer.Add(tlesizer2, 0, wx.CENTER | wx.ALL, border=5)

        tlesizer3 = wx.BoxSizer(wx.HORIZONTAL)

        tlesizer3.Add(self.txts[0], 0, wx.LEFT, border=15)
        tlesizer3.Add(self.cmbs[0], 0, wx.LEFT, border=10)
        tlesizer3.Add(self.addbtns[0], 0, wx.RIGHT, border=15)
        tlesizer3.Add(self.txts[4], 0, wx.LEFT, border=15)
        tlesizer3.Add(self.cmbs[4], 0, wx.LEFT, border=10)
        tlesizer3.Add(self.addbtns[4], 0, wx.RIGHT, border=15)

        bxSizer.Add(tlesizer3, 0, wx.CENTER | wx.ALL, border=5)

        tlesizer4 = wx.BoxSizer(wx.HORIZONTAL)

        tlesizer4.Add(self.txts[3], 0, wx.LEFT, border=20)
        tlesizer4.Add(self.cmbs[3], 0, wx.LEFT, border=10)
        tlesizer4.Add(self.addbtns[3], 0, wx.RIGHT, border=35)
        tlesizer4.Add(self.txtlbls[1], 0, wx.RIGHT, border=15)
        tlesizer4.Add(self.txtbxs[1], 0, wx.RIGHT, border=20)

        bxSizer.Add(tlesizer4, 0, wx.CENTER | wx.ALL, border=5)

        tlesizer5 = wx.BoxSizer(wx.HORIZONTAL)

        tlesizer5.Add(self.txtlbls[2], 0, wx.RIGHT, border=20)
        tlesizer5.Add(self.txtbxs[2], 0, wx.RIGHT, border=50)

        bxSizer.Add(tlesizer5, 0, wx.CENTER | wx.ALL, border=5)

        # form level controls
        fstsub = wx.Button(self, label='<<')
        lstsub = wx.Button(self, label='>>')
        nxtsub = wx.Button(self, label='>')
        presub = wx.Button(self, label='<')
        fstsub.Bind(wx.EVT_BUTTON, self.OnMovefstsub)
        lstsub.Bind(wx.EVT_BUTTON, self.OnMovelstsub)
        nxtsub.Bind(wx.EVT_BUTTON, self.OnMovenxtsub)
        presub.Bind(wx.EVT_BUTTON, self.OnMovepresub)

        navbox = wx.BoxSizer(wx.HORIZONTAL)
        navbox.Add(fstsub, 0, wx.ALL | wx.ALIGN_LEFT, 5)
        navbox.Add(presub, 0, wx.ALL | wx.ALIGN_LEFT, 5)
        navbox.Add(nxtsub, 0, wx.ALL | wx.ALIGN_LEFT, 5)
        navbox.Add(lstsub, 0, wx.ALL | wx.ALIGN_LEFT, 5)

        bxSizer.Add(navbox, 0, wx.CENTER | wx.ALL, border=5)

        numboxsub = wx.BoxSizer(wx.HORIZONTAL)
        self.recnumsub1 = wx.StaticText(self, label='Record ',
                                        style=wx.ALIGN_LEFT)
        self.recnumsub1.SetForegroundColour('blue')

        self.recnumsub2 = wx.StaticText(self, label=str(self.sub_rec_num+1),
                                        style=wx.ALIGN_LEFT)
        self.recnumsub2.SetForegroundColour('blue')
        self.recnumsub3 = wx.StaticText(self, label='/ ' +
                                        str(len(self.subdata)),
                                        style=wx.ALIGN_LEFT)
        self.recnumsub3.SetForegroundColour('blue')
        numboxsub.Add(self.recnumsub1, 0, wx.ALL | wx.ALIGN_LEFT, 5)
        numboxsub.Add(self.recnumsub2, 0, wx.ALL | wx.ALIGN_LEFT, 5)
        numboxsub.Add(self.recnumsub3, 0, wx.ALL | wx.ALIGN_LEFT, 5)

        bxSizer.Add(numboxsub, 0, wx.CENTER | wx.ALL, border=3)

        # Add buttons for form modifications
        self.subb2 = wx.Button(self, label="Update Procedure")
        self.Bind(wx.EVT_BUTTON, self.OnAddSubSpec, self.subb2)
        self.subb2.Disable()

        self.subb3 = wx.Button(self, label="Delete Procedure")
        self.Bind(wx.EVT_BUTTON, self.OnDeleteSubSpec, self.subb3)

        self.subb4 = wx.Button(self, label="Exit", size=(80, -1))
        self.Bind(wx.EVT_BUTTON, self.Close, self.subb4)

        navbox.Add((30, 10))
        navbox.Add(self.subb2, 0, wx.ALL | wx.ALIGN_LEFT, 5)
        navbox.Add(self.subb3, 0, wx.ALL | wx.ALIGN_LEFT, 5)
        navbox.Add(self.subb4, 0, wx.ALL | wx.ALIGN_LEFT, 5)

        self.Sizer.Add(bxSizer, 0, wx.CENTER)

        self.FillSubScrn()

        self.CenterOnParent()
        self.GetParent().Enable(False)
        self.Show(True)

        self.__eventLoop = wx.GUIEventLoop()
        self.__eventLoop.Run()

    def BxBld(self):
        self.txts = []
        self.cmbs = []
        self.txtbxs = []
        self.txtlbls = []
        self.addbtns = []

        font1 = wx.Font(12, wx.DEFAULT, wx.NORMAL, wx.FONTWEIGHT_BOLD)

        items = self.hints_tbls
        # order of information text label,table name,combo width,
        # combo box true = 1,table ID col name,column name shown in box

        for item in items:
            if item[3] == 0:
                txtbx = wx.TextCtrl(self, size=item[2], value='',
                                    style=wx.TE_LEFT)
                self.Bind(wx.EVT_TEXT, self.OnSelect2, txtbx)
                self.txtbxs.append(txtbx)
                txtlbl = wx.StaticText(self, label=item[0],
                                       style=wx.ALIGN_LEFT)
                txtlbl.SetForegroundColour((255, 0, 0))
                self.txtlbls.append(txtlbl)
            elif item[3] == 1:
                cmb = wx.ComboCtrl(self, size=item[2])
                self.Bind(wx.EVT_TEXT, self.OnSelect2, cmb)
                self.showcol = 1
                cmb.SetPopupControl(ListCtrlComboPopup(
                    item[1], showcol=self.showcol, lctrls=self.lctrls))
                self.cmbs.append(cmb)

                txt = wx.StaticText(self, label=item[0], style=wx.ALIGN_LEFT)
                txt.SetForegroundColour((255, 0, 0))
                self.txts.append(txt)

                addbtn = wx.Button(self, label='+', size=(35, -1))
                addbtn.SetForegroundColour((255, 0, 0))
                addbtn.SetFont(font1)
                self.Bind(wx.EVT_BUTTON, self.OnAdd1, addbtn)
                self.addbtns.append(addbtn)

    def OnClose(self, evt):
        self.Destroy()

    def OnSelect2(self, evt):
        self.subb2.Enable()

    def OnMovefstsub(self, evt):
        self.sub_rec_num = 0
        self.FillSubScrn()

    def OnMovelstsub(self, evt):
        if len(self.subdata) == 0:
            return
        self.sub_rec_num = len(self.subdata)-1
        self.FillSubScrn()

    def OnMovenxtsub(self, evt):
        if len(self.subdata) == 0:
            return
        self.sub_rec_num += 1
        if self.sub_rec_num == len(self.subdata):
            self.sub_rec_num = 0
        self.FillSubScrn()

    def OnMovepresub(self, evt):
        if len(self.subdata) == 0:
            return
        self.sub_rec_num -= 1
        if self.sub_rec_num < 0:
            self.sub_rec_num = len(self.subdata)-1
        self.FillSubScrn()

    def FillSubScrn(self):
        SubSQL = ('SELECT * FROM WeldProcedures')
        self.subdata = Dbase().Dsqldata(SubSQL)
        subrecrd = self.subdata[self.sub_rec_num]

        # numbers represent (record in hints_tbls,
        # cmbbxs number,ID in subrecrd data)
        for grp in [(0, 0, 1), (1, 0, 2), (2, 1, 3), (3, 1, 4), (4, 2, 5),
                    (5, 3, 7), (6, 4, 6), (7, 5, 8), (8, 2, 9)]:
            subtbl = self.hints_tbls[grp[0]]
            if subtbl[3] == 1:
                qry = ('SELECT ' + subtbl[5] + ' FROM ' + subtbl[1] + ' WHERE '
                       + subtbl[4] + ' = ' + str(subrecrd[grp[2]]))
                recrd = Dbase().Dsqldata(qry)
                if recrd != []:
                    recrd = recrd[0][0]
                    self.cmbs[grp[1]].ChangeValue(recrd)
            if subtbl[3] == 0:
                if subrecrd[grp[2]] is None:
                    self.txtbxs[grp[1]].ChangeValue('')
                else:
                    self.txtbxs[grp[1]].ChangeValue(str(subrecrd[grp[2]]))

        self.recnumsub2.SetLabel(str(self.sub_rec_num+1))
        self.recnumsub3.SetLabel('/ ' + str(len(self.subdata)))

    def OnAdd1(self, evt):
        btn = evt.GetEventObject()
        callbtn = self.btnDict[btn]

        if callbtn == 0:
            cmbtbl = 'WeldFiller'
        elif callbtn == 1:
            cmbtbl = 'WeldQualifyThickness'
        elif callbtn == 2:
            cmbtbl = 'WeldProcessList'
        elif callbtn == 3:
            cmbtbl = 'WeldMaterialGroup'
        elif callbtn == 4:
            cmbtbl = 'WeldFillerGroup'
        elif callbtn == 5:
            cmbtbl = 'WeldQualifyPosition'

        CmbLst1(self, cmbtbl)
        self.ReFillList(cmbtbl, callbtn)

    def EditSubTbl(self):
        realnames = []
        ValueList = []

        choices = ['''1) Add this as a new procedure in the overall
                    welding procedures''', '''2) Update the existing procedure
                    and reflect the change in the welding procedures''']
        txt = ('''NOTE: Updating this information will be\nreflected in all
                associated welding procedure!\nRecommendation is to save as
                a new specification.\n\n\tHow do you want to proceed?''')
        # use a SingleChioce dialog to determine if data
        # is new record or edited record
        SQL_Dialog = wx.SingleChoiceDialog(self, txt,
                                           'Information Has Changed', choices,
                                           style=wx.CHOICEDLG_STYLE)

        SQL_step = 3
        if SQL_Dialog.ShowModal() == wx.ID_OK:
            SQL_step = SQL_Dialog.GetSelection()
        SQL_Dialog.Destroy()

        colinfo = Dbase().Dcolinfo('WeldProcedures')
        ValueList = [None for i in range(0, len(colinfo))]
        # if the table index is auto increment then assign
        # next value otherwise do nothing
        for item in colinfo:
            if item[5] == 1:
                # IDcol = item[0]
                IDname = item[1]
                if 'INTEGER' in item[2]:
                    New_ID = cursr.execute(
                        "SELECT MAX(" + IDname + ") FROM WeldProcedures").\
                        fetchone()
                    if New_ID[0] is None:
                        Max_ID = '1'
                    else:
                        Max_ID = str(New_ID[0]+1)
            realnames.append(item[1])
        ValueList[0] = Max_ID

        # numbers represent (record in hints_tbls,cmbbxs number,
        # ID in subrecrd data)
        for grp in [(0, 0, 1), (1, 0, 2), (2, 1, 3), (3, 1, 4), (4, 2, 5),
                    (5, 3, 6), (6, 4, 7), (7, 5, 8), (8, 2, 9)]:
            subtbl = self.hints_tbls[grp[0]]
            if subtbl[3] == 1:
                qry = ("SELECT " + subtbl[4] + " FROM " + subtbl[1] + " WHERE "
                       + subtbl[5] + " = '" + self.cmbs[grp[1]].GetValue()
                       + "'")
                recrd = Dbase().Dsqldata(qry)
                ValueList[grp[2]] = recrd[0][0]

            if subtbl[3] == 0:
                if self.txtbxs[grp[1]].GetValue() == '':
                    ValueList[grp[2]] = None
                else:
                    ValueList[grp[2]] = self.txtbxs[grp[1]].GetValue()

        if SQL_step == 0:  # enter new record

            num_vals = ('?,'*len(colinfo))[:-1]
            UpQuery = 'INSERT INTO WeldProcedures VALUES (' + num_vals + ')'
            Dbase().TblEdit(UpQuery, ValueList)
            self.sub_rec_num = len(self.subdata)

        elif SQL_step == 1:  # update edited record
            CurrentID = self.subdata[self.sub_rec_num][0]
            realnames.remove('ProcedureID')
            del ValueList[0]

            SQL_str = ','.join(["%s=?" % (name) for name in realnames])
            UpQuery = ('UPDATE WeldProcedures SET ' + SQL_str +
                       ' WHERE ProcedureID = ' + str(CurrentID))
            Dbase().TblEdit(UpQuery, ValueList)

        elif SQL_step == 3:
            return

        self.FillSubScrn()
        self.recnumsub3.SetLabel('/ '+str(len(self.subdata)))

    def ReFillList(self, cmbtbl, boxnum):
        self.lc = self.lctrls[boxnum]
        self.lc.DeleteAllItems()
        index = 0
        ReFillQuery = 'SELECT * FROM "' + cmbtbl + '"'
        for values in Dbase().Dsqldata(ReFillQuery):
            col = 0
            for value in values:
                if col == 0:
                    self.lc.InsertItem(index, str(value))
                else:
                    self.lc.SetItem(index, col, str(value))
                col += 1
            index += 1

    def OnDeleteSubSpec(self, evt):
        recrd = self.subdata[self.sub_rec_num][0]
        try:
            Dbase().TblDelete('WeldProcedures', recrd, 'ProcedureID')
            self.MainSQL = 'SELECT * FROM WeldProcedures'
            self.subdata = Dbase().Dsqldata(self.MainSQL)
            self.sub_rec_num -= 1
            if self.sub_rec_num < 0:
                self.sub_rec_num = len(self.subdata)-1
            self.FillSubScrn()
            self.recnumsub3.SetLabel('/ '+str(len(self.subdata)))

        except sqlite3.IntegrityError:
            wx.MessageBox(
                "This Record is associated"
                " with\nother tables and cannot be deleted!", "Cannot Delete",
                wx.OK | wx.ICON_INFORMATION)

    def OnAddSubSpec(self, evt):
        LogStr = []

        # numbers represent (record in hints_tbls,
        # cmbbxs number,ID in subrecrd data)
        for grp in [(0, 0, 1), (1, 0, 2), (2, 1, 3), (3, 1, 4), (4, 2, 5),
                    (5, 3, 7), (6, 4, 6), (7, 5, 8), (8, 2, 9)]:
            if self.cmbs[grp[1]].GetValue() == '':
                LogStr.append(self.hints_tbls[grp[0]][0].replace('\n', ' '))

        if len(LogStr) > 0:
            LogStr = '\n'.join(LogStr)
            wx.MessageBox('Value needed for;\n' + LogStr +
                          '\nto complete information.', 'Missing Data',
                          wx.OK | wx.ICON_INFORMATION)
            return False
        else:
            self.EditSubTbl()
            self.subb2.Disable()

    def Close(self, evt):
        self.GetParent().Enable(True)
        self.__eventLoop.Exit()
        self.Destroy()


class BldInsul(wx.Frame):
    '''Routine to build form and populate grid'''
    def __init__(self, parent, tblname, ComdPrtyID=None):

        self.parent = parent  # added for child form
        self.tblname = tblname

        # there are 29 data IDs making up the Insulation Code
        self.txtparts = {}

        for i in range(0, 30):
            self.txtparts[i] = 0

        if self.tblname.find("_") != -1:
            self.frmtitle = (self.tblname.replace("_", " "))
        else:
            self.frmtitle = (' '.join(re.findall('([A-Z][a-z]*)',
                             self.tblname)))
        self.txtparts[0] = str(self.tblname[0]) + 'C'

        super(BldInsul, self).__init__(parent,
                                       title=self.frmtitle,
                                       size=(830, 780))

        self.Bind(wx.EVT_CLOSE, self.OnClose)

        self.ComEnds = ''
        self.ComCode = ''
        self.PipeMtrSpec = ''
        self.ComdPrtyID = ComdPrtyID

        self.InitUI()

    def InitUI(self):

        if self.ComdPrtyID is not None:
            query = ('''SELECT Commodity_Code,Pipe_Material_Code,
                     End_Connection,Pipe_Code FROM CommodityProperties
                      WHERE CommodityPropertyID = '''
                     + str(self.ComdPrtyID))
            dataset = Dbase().Dsqldata(query)[0]
            self.PipeMtrSpec = dataset[1]
            self.ComCode = dataset[0]
            self.ComEnds = dataset[2]
            self.PipeCode = dataset[3]

        self.data = []
        self.columnames = []
        self.cmbctrls = []
        self.rec_num = 0
        self.addbtns = []
        self.lctrls = []
        self.restore = False
        self.bxdict = {}
        self.DsqlFtg = ''
        StartQry = None

        font1 = wx.Font(12, wx.DEFAULT, wx.NORMAL, wx.FONTWEIGHT_BOLD)

        txt_nodata = ('The ' + self.frmtitle +
                      ' have not been setup for this Commodity Property')
        self.lbl_nodata = wx.StaticText(self, -1, label=txt_nodata,
                                        size=(600, 30), style=wx.LEFT)
        self.lbl_nodata.SetForegroundColour((255, 0, 0))
        self.lbl_nodata.SetFont(font1)
        self.lbl_nodata.SetLabel('  ')

        if self.ComdPrtyID is not None:
            query = ('SELECT ' + self.tblname +
                     '''_ID FROM PipeSpecification WHERE
                      Commodity_Property_ID = '''
                     + str(self.ComdPrtyID))
            chk = Dbase().Dsqldata(query)
            if chk != []:
                StartQry = chk[0][0]
                if StartQry is not None:
                    self.DsqlFtg = ('SELECT * FROM ' + self.tblname + ' WHERE '
                                    + self.tblname + 'ID = ' + str(StartQry))
                    self.data = Dbase().Dsqldata(self.DsqlFtg)
                else:
                    self.lbl_nodata.SetLabel(txt_nodata)
            else:
                self.lbl_nodata.SetLabel(txt_nodata)
        else:
            self.DsqlFtg = 'SELECT * FROM ' + self.tblname
            self.data = Dbase().Dsqldata(self.DsqlFtg)

        # specify which listbox column to display in the combobox, default 1
        self.showcol = int

        # set up the table column names, width and if column can be
        # edited ie primary autoincrement
        tblinfo = []

        tblinfo = Dbase(self.tblname).Fld_Size_Type()
        self.ID_col = tblinfo[1]
        self.colms = tblinfo[3]
        self.pkcol_name = tblinfo[0]
        self.autoincrement = tblinfo[2]

        names = ['Min. Dia', 'Max. Dia', 'Material', 'Thickness']
        self.lbl_txt = ['Min. Pipe\nDiameter', 'Max. Pipe\nDiameter',
                        'Material', 'Thickness']
        self.cmb_tbls = ['Pipe_OD', 'Pipe_OD', 'InsulationMaterial',
                         'InsulationThickness']
        self.cmb_tbls2 = ['InsulationJacket', 'PaintPrep', 'InsulationClass',
                          'InsulationAdhesive', 'InsulationSealer']

        for n in range(0, (len(self.colms)-3)//len(names)):
            for col in names:
                self.columnames.append(col)

        self.Num_Cols = len(self.lbl_txt)
        self.Num_Rows = 6

        # set the Sizer property (same as SetSizer)
        self.Sizer = wx.BoxSizer(wx.VERTICAL)

        self.warningsizer = wx.BoxSizer(wx.HORIZONTAL)
        self.warningsizer.Add(self.lbl_nodata, 0, wx.ALIGN_CENTER)

        self.searchsizer = wx.BoxSizer(wx.HORIZONTAL)
        font = wx.Font(10, wx.DEFAULT, wx.NORMAL, wx.NORMAL)

        if self.ComdPrtyID is not None:
            query = ('SELECT * FROM PipeMaterialSpec WHERE Material_Spec_ID = '
                     + str(self.PipeMtrSpec))
            MtrSpc = Dbase().Dsqldata(query)[0][1]
            self.spec_ID = Dbase().Dsqldata(query)[0][0]

            if self.PipeCode is not None:
                txt = self.PipeCode
            else:
                query = ('''SELECT * FROM PipeMaterialSpec WHERE
                          Material_Spec_ID = ''' + str(self.PipeMtrSpec))
                MtrSpc = Dbase().Dsqldata(query)[0][1]
                txt = self.ComCode + ' - ' + MtrSpc

            self.text1 = wx.TextCtrl(self, size=(100, 33), value=txt,
                                     style=wx.TE_READONLY | wx.TE_CENTER)
            self.text1.SetForegroundColour((255, 0, 0))
            self.text1.SetFont(font)
            self.searchsizer.Add(self.text1, 0, wx.ALIGN_LEFT, 5)

        self.text4 = wx.TextCtrl(self, size=(480, 33), value=str(self.frmtitle)
                                 + ' Code',
                                 style=wx.TE_READONLY | wx.TE_CENTER)
        self.text4.SetForegroundColour((255, 0, 0))
        self.text4.SetFont(font)
        self.searchsizer.Add(self.text4, 0, wx.ALIGN_LEFT, 5)

        self.Upsizer = wx.BoxSizer(wx.HORIZONTAL)

        self.UpLsizer = wx.BoxSizer(wx.VERTICAL)
        self.UpLbtnsizer = wx.BoxSizer(wx.VERTICAL)

        self.lblJkt = wx.StaticText(self, -1, label='Jacket Material',
                                    style=wx.CENTER)
        self.cmbJkt = wx.ComboCtrl(self, id=25, size=(250, -1),
                                   style=wx.CB_READONLY)
        self.Bind(wx.EVT_COMBOBOX_CLOSEUP, self.OnCmbTypeClose, self.cmbJkt)
        self.Bind(wx.EVT_COMBOBOX_DROPDOWN, self.OnCmbTypeOpen, self.cmbJkt)
        self.showcol = 1
        self.cmbJkt.SetPopupControl(ListCtrlComboPopup(
            self.cmb_tbls2[0], PupQuery='',
            lctrls=self.lctrls, showcol=self.showcol))
        self.txtJkt = wx.StaticText(self, -1, style=wx.ALIGN_CENTER)
        txt = 'Jacket Thickness'
        self.txtJkt.SetLabel(txt)

        self.addJkt = wx.Button(self, label='+', size=(35, -1))
        self.addJkt.SetForegroundColour((255, 0, 0))
        self.addJkt.SetFont(font1)
        self.Bind(wx.EVT_BUTTON, self.OnAddJkt, self.addJkt)

        self.lblSuf = wx.StaticText(self, -1, label='Surface Preparation',
                                    style=wx.CENTER)
        self.cmbSuf = wx.ComboCtrl(self, id=26, size=(250, -1),
                                   style=wx.CB_READONLY)
        self.Bind(wx.EVT_COMBOBOX_CLOSEUP, self.OnCmbTypeClose, self.cmbSuf)
        self.Bind(wx.EVT_COMBOBOX_DROPDOWN, self.OnCmbTypeOpen, self.cmbSuf)
        self.showcol = 1
        self.cmbSuf.SetPopupControl(ListCtrlComboPopup(
            self.cmb_tbls2[1], PupQuery='', lctrls=self.lctrls,
            showcol=self.showcol))
        self.txtSuf = wx.StaticText(self, -1, style=wx.ALIGN_CENTER)
        txt = 'Preparation'
        self.txtSuf.SetLabel(txt)

        self.addSuf = wx.Button(self, label='+', size=(35, -1))
        self.addSuf.SetForegroundColour((255, 0, 0))
        self.addSuf.SetFont(font1)
        self.Bind(wx.EVT_BUTTON, self.OnAddSuf, self.addSuf)

        self.UpLsizer.Add((10, 5))
        self.UpLsizer.Add(self.lblJkt, 0, wx.ALIGN_CENTER)
        self.UpLsizer.Add(self.cmbJkt, 0, wx.ALIGN_CENTER)
        self.UpLsizer.Add(self.txtJkt, 0, wx.ALIGN_LEFT)
        self.UpLsizer.Add((10, 33))
        self.UpLsizer.Add(self.lblSuf, 0, wx.ALIGN_CENTER)
        self.UpLsizer.Add(self.cmbSuf, 0, wx.ALIGN_CENTER)
        self.UpLsizer.Add(self.txtSuf, 0, wx.ALIGN_LEFT)

        self.UpLbtnsizer.Add((10, 22))
        self.UpLbtnsizer.Add(self.addJkt, 0, wx.ALIGN_CENTER)
        self.UpLbtnsizer.Add((10, 60))
        self.UpLbtnsizer.Add(self.addSuf, 0, wx.ALIGN_CENTER)

        self.UpRsizer = wx.BoxSizer(wx.VERTICAL)
        self.UpRbtnsizer = wx.BoxSizer(wx.VERTICAL)

        self.lblCls = wx.StaticText(self, -1, label='Insulation Class',
                                    style=wx.CENTER)
        self.cmbCls = wx.ComboCtrl(self, id=27, size=(250, -1),
                                   style=wx.CB_READONLY)
        self.Bind(wx.EVT_COMBOBOX_CLOSEUP, self.OnCmbTypeClose, self.cmbCls)
        self.Bind(wx.EVT_COMBOBOX_DROPDOWN, self.OnCmbTypeOpen, self.cmbCls)
        self.showcol = 1
        self.cmbCls.SetPopupControl(ListCtrlComboPopup(
            'InsulationClass', PupQuery='', lctrls=self.lctrls,
            showcol=self.showcol))

        self.addCls = wx.Button(self, label='+', size=(35, -1))
        self.addCls.SetForegroundColour((255, 0, 0))
        self.addCls.SetFont(font1)
        self.Bind(wx.EVT_BUTTON, self.OnAddCls, self.addCls)

        self.lblAdh = wx.StaticText(self, -1, label='Adhesive',
                                    style=wx.CENTER)
        self.cmbAdh = wx.ComboCtrl(self, id=28, size=(250, -1),
                                   style=wx.CB_READONLY)
        self.Bind(wx.EVT_COMBOBOX_CLOSEUP, self.OnCmbTypeClose, self.cmbAdh)
        self.Bind(wx.EVT_COMBOBOX_DROPDOWN, self.OnCmbTypeOpen, self.cmbAdh)
        self.showcol = 1
        self.cmbAdh.SetPopupControl(ListCtrlComboPopup(
            'InsulationAdhesive', PupQuery='', lctrls=self.lctrls,
            showcol=self.showcol))

        self.addAdh = wx.Button(self, label='+', size=(35, -1))
        self.addAdh.SetForegroundColour((255, 0, 0))
        self.addAdh.SetFont(font1)
        self.Bind(wx.EVT_BUTTON, self.OnAddAdh, self.addAdh)

        self.lblSel = wx.StaticText(self, -1, label='Sealer', style=wx.CENTER)
        self.cmbSel = wx.ComboCtrl(self, id=29, size=(250, -1),
                                   style=wx.CB_READONLY)
        self.Bind(wx.EVT_COMBOBOX_CLOSEUP, self.OnCmbTypeClose, self.cmbSel)
        self.Bind(wx.EVT_COMBOBOX_DROPDOWN, self.OnCmbTypeOpen, self.cmbSel)
        self.showcol = 1
        self.cmbSel.SetPopupControl(ListCtrlComboPopup(
            'InsulationSealer', PupQuery='', lctrls=self.lctrls,
            showcol=self.showcol))

        self.addSel = wx.Button(self, label='+', size=(35, -1))
        self.addSel.SetForegroundColour((255, 0, 0))
        self.addSel.SetFont(font1)
        self.Bind(wx.EVT_BUTTON, self.OnAddSel, self.addSel)

        self.UpRsizer.Add((10, 5))
        self.UpRsizer.Add(self.lblCls, 0, wx.ALIGN_CENTER)
        self.UpRsizer.Add(self.cmbCls, 0, wx.ALIGN_CENTER)
        self.UpRsizer.Add((10, 5))
        self.UpRsizer.Add(self.lblAdh, 0, wx.ALIGN_CENTER)
        self.UpRsizer.Add(self.cmbAdh, 0, wx.ALIGN_CENTER)
        self.UpRsizer.Add((10, 5))
        self.UpRsizer.Add(self.lblSel, 0, wx.ALIGN_CENTER)
        self.UpRsizer.Add(self.cmbSel, 0, wx.ALIGN_CENTER)

        self.UpRbtnsizer.Add((10, 20))
        self.UpRbtnsizer.Add(self.addCls, 0, wx.ALIGN_CENTER)
        self.UpRbtnsizer.Add((10, 15))
        self.UpRbtnsizer.Add(self.addAdh, 0, wx.ALIGN_CENTER)
        self.UpRbtnsizer.Add((10, 15))
        self.UpRbtnsizer.Add(self.addSel, 0, wx.ALIGN_CENTER)

        self.Upsizer.Add(self.UpLsizer, 0, wx.LEFT)
        self.Upsizer.Add(self.UpLbtnsizer, 0, wx.LEFT)
        self.Upsizer.Add((80, 10))
        self.Upsizer.Add(self.UpRsizer, 0, wx.RIGHT)
        self.Upsizer.Add(self.UpRbtnsizer, 0, wx.RIGHT)

        self.notesizer = wx.BoxSizer(wx.HORIZONTAL)
        self.lblnote = wx.StaticText(self, -1, label='Notes:', style=wx.CENTER)
        self.notes = wx.TextCtrl(self, size=(500, 60), value='',
                                 style=wx.TE_MULTILINE | wx.TE_LEFT)
        self.notesizer.Add(self.lblnote, 0, wx.LEFT, 5)
        self.notesizer.Add(self.notes, 0, wx.ALIGN_LEFT | wx.LEFT, 10)

        # set up first row of combo boxes and label
        self.codesizer = wx.BoxSizer(wx.HORIZONTAL)

        cmbsizers = []
        for row in range(0, self.Num_Rows):
            cmbsizer = wx.BoxSizer(wx.HORIZONTAL)
            cmbsizers.append(cmbsizer)

        # set up the cmbbox labels
        self.lblsizer = wx.BoxSizer(wx.HORIZONTAL)
        xply = 1
        for txt in self.lbl_txt:
            self.lbl = wx.StaticText(self, -1, label=txt,
                                     style=wx.ALIGN_CENTER_HORIZONTAL)

            self.addbtn = wx.Button(self, label='+', size=(35, -1))
            self.addbtn.SetForegroundColour((255, 0, 0))
            self.addbtn.SetFont(font1)
            self.Bind(wx.EVT_BUTTON, self.OnAdd1, self.addbtn)
            self.addbtns.append(self.addbtn)

            if xply < 1.18:
                xply += .28
            self.lblsizer.Add(self.lbl, 0, wx.ALL | wx.ALIGN_BOTTOM, 10)
            self.lblsizer.Add(self.addbtn, 0, wx.ALIGN_CENTER)
            self.lblsizer.Add((20*xply, 25))

        # Start the generation of the required combo boxes
        # list of combobox objects

        # the list of grid column names excluding ID,Valve ID and notes
        self.cmb_range = self.columnames

        # counter to track which comboboxes need activation
        # after material type is selected
        self.a = []
        n = 0
        # cmb_range is the list of the combo box
        for h in range(0, self.Num_Rows):
            for cmb_tbl in self.cmb_tbls:
                cmbbox = wx.ComboCtrl(self, size=(150, -1),
                                      style=wx.CB_READONLY)
                self.Bind(wx.EVT_COMBOBOX_CLOSEUP, self.OnCmbClose, cmbbox)
                self.Bind(wx.EVT_COMBOBOX_DROPDOWN, self.OnCmbOpen, cmbbox)

                cmbbox.SetPopupControl(ListCtrlComboPopup(
                    cmb_tbl, PupQuery='', lctrls=self.lctrls, showcol=1))

                n += 1

                # disable the boxes below row 1
                if h >= 1:
                    cmbbox.Disable()
                self.cmbctrls.append(cmbbox)

                # and the cmbbox to the cmbsizer
                cmbsizers[h].Add(cmbbox, 0, wx.ALIGN_LEFT, 5)

        # add the upper combo boxes to the list of cmbctrls
        self.cmbctrls.append(self.cmbJkt)
        self.cmbctrls.append(self.cmbSuf)
        self.cmbctrls.append(self.cmbCls)
        self.cmbctrls.append(self.cmbAdh)
        self.cmbctrls.append(self.cmbSel)

        # Add some buttons
        self.b5 = wx.Button(self, label="Clear\nBoxes")
        self.Bind(wx.EVT_BUTTON, self.OnRestoreCmbs, self.b5)

        self.b1 = wx.Button(self, label="Print\nReport")
        self.Bind(wx.EVT_BUTTON, self.PrintFile, self.b1)

        self.b2 = wx.Button(self, label="Add/Update\nto " + self.frmtitle)
        self.b2.Disable()
        self.Bind(wx.EVT_BUTTON, self.OnAddRec, self.b2)

        self.b3 = wx.Button(self, label="Delete\nSpecification")
        self.Bind(wx.EVT_BUTTON, self.OnDeleteRec, self.b3)

        self.b4 = wx.Button(self, label="Exit")
        self.Bind(wx.EVT_BUTTON, self.OnClose, self.b4)

        # add a button box and place the buttons
        self.btnbox = wx.BoxSizer(wx.HORIZONTAL)
        self.btnbox.Add(self.b5, 0, wx.ALIGN_CENTER | wx.RIGHT, 35)
        self.btnbox.Add(self.b1, 0, wx.ALIGN_CENTER | wx.LEFT, 5)
        self.btnbox.Add(self.b2, 0, wx.ALIGN_CENTER | wx.LEFT, 5)
        if self.ComdPrtyID is not None:
            self.b6 = wx.Button(self, size=(120, 45), label="Show All\nItems")
            self.b3.SetLabel('Delete From\nCommodity')
            if StartQry is None:
                self.b3.Disable()
            self.Bind(wx.EVT_BUTTON, self.OnAddComd, self.b6)
            self.btnbox.Add(self.b6, 0, wx.ALL | wx.ALIGN_CENTER, 5)
        self.btnbox.Add(self.b3, 0, wx.ALIGN_CENTER | wx.LEFT, 5)
        self.btnbox.Add((30, 10))
        self.btnbox.Add(self.b4, 0, wx.ALIGN_CENTER | wx.LEFT, 5)

        self.fst = wx.Button(self, label='<<')
        self.lst = wx.Button(self, label='>>')
        self.nxt = wx.Button(self, label='>')
        self.pre = wx.Button(self, label='<')
        self.fst.Bind(wx.EVT_BUTTON, self.OnMovefst)
        self.lst.Bind(wx.EVT_BUTTON, self.OnMovelst)
        self.nxt.Bind(wx.EVT_BUTTON, self.OnMovenxt)
        self.pre.Bind(wx.EVT_BUTTON, self.OnMovepre)

        self.navbox = wx.BoxSizer(wx.HORIZONTAL)
        self.navbox.Add(self.fst, 0, wx.ALL | wx.ALIGN_LEFT, 5)
        self.navbox.Add(self.pre, 0, wx.ALL | wx.ALIGN_LEFT, 5)
        self.navbox.Add(self.nxt, 0, wx.ALL | wx.ALIGN_LEFT, 5)
        self.navbox.Add(self.lst, 0, wx.ALL | wx.ALIGN_LEFT, 5)

        self.numbox = wx.BoxSizer(wx.HORIZONTAL)
        self.recnum1 = wx.StaticText(self, label='Record ',
                                     style=wx.ALIGN_LEFT)
        self.recnum1.SetForegroundColour((255, 0, 0))

        self.recnum2 = wx.StaticText(self, label=str(self.rec_num+1),
                                     style=wx.ALIGN_LEFT)
        self.recnum2.SetForegroundColour((255, 0, 0))
        self.recnum3 = wx.StaticText(self, label='/ '+str(len(self.data)),
                                     style=wx.ALIGN_LEFT)
        self.recnum3.SetForegroundColour((255, 0, 0))
        self.numbox.Add(self.recnum1, 0, wx.ALL | wx.ALIGN_LEFT, 5)
        self.numbox.Add(self.recnum2, 0, wx.ALL | wx.ALIGN_LEFT, 5)
        self.numbox.Add(self.recnum3, 0, wx.ALL | wx.ALIGN_LEFT, 5)

        self.Sizer.Add(self.warningsizer, 0, wx.CENTER | wx.TOP, 5)
        self.Sizer.Add(self.searchsizer, 0, wx.ALL | wx.ALIGN_CENTER, 5)
        self.Sizer.Add(self.codesizer, 0, wx.ALL | wx.ALIGN_CENTER, 5)
        self.Sizer.Add(self.Upsizer, 0, wx.ALL | wx.ALIGN_CENTER, 5)
        self.Sizer.Add(self.notesizer, 0, wx.ALIGN_CENTER | wx.TOP |
                       wx.BOTTOM, 10)
        self.Sizer.Add(self.lblsizer, 0, wx.ALIGN_CENTER)

        for cmbsizer in cmbsizers:
            self.Sizer.Add(cmbsizer, 0, wx.ALIGN_CENTER)
        self.Sizer.Add((25, 10))
        self.Sizer.Add(self.btnbox, 0, wx.ALIGN_CENTER, 5)
        self.Sizer.Add((25, 10))
        self.Sizer.Add(self.navbox, 0, wx.ALIGN_CENTER, 5)
        self.Sizer.Add(self.numbox, 0, wx.ALIGN_CENTER, 5)

        self.btnDict = {self.addbtns[i]: i for i in range(len(self.addbtns))}

        self.FillScreen()

        # add the following 5 lines for child parent form
        self.CenterOnParent()
        self.GetParent().Enable(False)
        self.Show(True)
        self.__eventLoop = wx.GUIEventLoop()
        self.__eventLoop.Run()

    def OnMovefst(self, evt):
        self.rec_num = 0
        self.FillScreen()

    def OnMovelst(self, evt):
        if len(self.data) == 0:
            return
        self.rec_num = len(self.data)-1
        self.FillScreen()

    def OnMovenxt(self, evt):
        if len(self.data) == 0:
            return
        self.rec_num += 1
        if self.rec_num == len(self.data):
            self.rec_num = 0
        self.FillScreen()

    def OnMovepre(self, evt):
        if len(self.data) == 0:
            return
        self.rec_num -= 1
        if self.rec_num < 0:
            self.rec_num = len(self.data)-1
        self.FillScreen()

    def FillScreen(self):
        m = 0
        SufPrep = 'Not Specified'
        InslThk = 'Not Specified'
        # all the IDs for the various tables making up the insulation package
        if len(self.data) == 0:
            self.recnum2.SetLabel(str(self.rec_num))
            return
        else:
            recrd = self.data[self.rec_num]
            minum = self.Num_Rows * self.Num_Cols

        # cycle through each row by column
        for m in range(0, self.Num_Rows * self.Num_Cols):
            tbl_ID = str(recrd[m+1])

            tbl_ID_nam = Dbase().Dcolinfo(
                self.cmb_tbls[m % self.Num_Cols])[0][1]

            if tbl_ID != 'None':
                query = ('SELECT * FROM ' + self.cmb_tbls[m % self.Num_Cols] +
                         ' WHERE ' + tbl_ID_nam + ' = ' + tbl_ID)
                self.cmbctrls[m].ChangeValue(
                    str(Dbase().Dsqldata(query)[0][1]))
                self.txtparts[m+1] = tbl_ID
                self.cmbctrls[m].Enable()
            else:
                self.cmbctrls[m].ChangeValue('')
                self.txtparts[m+1] = '0'
                self.cmbctrls[m].Disable()
                if m < minum:
                    minum = m

        if minum < self.Num_Cols * self.Num_Rows:
            for m in range(minum, minum+self.Num_Cols):
                self.cmbctrls[m].Enable()

        # the above code uses the first data IDs to
        # fill the lower grid of combo boxes
        # this fills the upper 5 combo boxes which is
        # made up of the last data IDs in the Insulation code
        n = 0
        for m in range(25, 30):
            tbl_ID = str(recrd[m])

            tbl_ID_nam = Dbase().Dcolinfo(self.cmb_tbls2[n])[0][1]

            if tbl_ID != 'None':
                query = ('SELECT * FROM ' + self.cmb_tbls2[n] +
                         ' WHERE ' + tbl_ID_nam + ' = ' + tbl_ID)
                data = Dbase().Dsqldata(query)[0]
                self.cmbctrls[m-1].ChangeValue(str(data[1]))
                self.txtparts[m] = tbl_ID
                if m == 25:
                    if data[2] is not None:
                        InslThk = data[2]
                elif m == 26:
                    if data[2] is not None:
                        SufPrep = data[2]
            else:
                self.cmbctrls[m-1].ChangeValue('')
                self.txtparts[m] = '0'

            self.txtSuf.SetLabel('Preparation = ' + SufPrep)
            self.txtJkt.SetLabel('Jacket Thickness = ' + InslThk)
            n += 1

        if recrd[30] == '':
            self.notes.ChangeValue('')
        else:
            self.notes.ChangeValue(str(recrd[30]))

        if recrd[31] == 'None':
            self.text4.ChangeValue(str(self.frmtitle) + ' Code')
        else:
            self.text4.ChangeValue(str(recrd[31]))
        self.recnum2.SetLabel(str(self.rec_num+1))

    def PrintFile(self, evt):
        import Report_Insul

        if self.data == []:
            NoData = wx.MessageDialog(
                None, 'No Data to Print', 'Error', wx.OK | wx.ICON_EXCLAMATION)
            NoData.ShowModal()
            return

        saveDialog = wx.FileDialog(self, message='Save Report as PDF.',
                                   wildcard='PDF (*.pdf)|*.pdf',
                                   style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT)

        if saveDialog.ShowModal() == wx.ID_CANCEL:
            filename = ''

        filename = saveDialog.GetPath()
        if filename.find(".pdf") == -1:
            filename = filename + '.pdf'

        saveDialog.Destroy()

        Colwdths, Colnames, rptdata = self.ReportData()

        ttl = None
        if len(self.data) == 1:
            # specify a title for the report if table name
            # is not to be used
            if self.PipeCode is None:
                ttl = self.text1.GetValue()
            else:
                ttl = self.PipeCode
            ttl = ' for ' + ttl

        Report_Insul.Report(
            self.tblname, rptdata, Colnames, Colwdths,
            filename, ttl).create_pdf()

    def ReportData(self):
        rptdata = []

        Colnames = [('ID', 'Insulation Code', 'Jacketing', 'Surface Prep',
                     'Insulation Class', 'Adhesive', 'Sealer', 'Note'),
                    ('Min. Dia', 'Max. Dia', 'Material', 'Thickness')]

        Colwdths = [(5, 35, 10, 15, 15, 15, 15, 30), (10, 10, 15, 10)]

        for segn in range(0, len(self.data)):
            # specify the general data for the insulation
            data1 = []
            data1.append(str(self.data[segn][0]))  # Insulation ID
            data1.append(str(self.data[segn][31]))  # Insulation code

            n = 0
            for i in self.data[segn][25:30]:
                if i is None:
                    data1.append('N/A')
                    continue
                else:
                    if n == 0:
                        qry = ('''SELECT JacketMatr, Jacket_Thk FROM InsulationJacket
                                WHERE JacketID = ''' + str(i))
                        data1.append(Dbase().Dsqldata(qry)[0][0])
                    if n == 1:
                        qry = ('''SELECT Notes FROM PaintPrep
                                WHERE PaintPrepID = ''' + str(i))
                        data1.append(Dbase().Dsqldata(qry)[0][0])
                    if n == 2:
                        qry = ('''SELECT Insulation_Class FROM InsulationClass
                                WHERE ClassID = ''' + str(i))
                        data1.append(Dbase().Dsqldata(qry)[0][0])
                    if n == 3:
                        qry = ('''SELECT Adhesive, Vendor FROM InsulationAdhesive
                                WHERE AdhesiveID = ''' + str(i))
                        data1.append(Dbase().Dsqldata(qry)[0][0])
                    if n == 4:
                        qry = ('''SELECT Sealer, Vendor FROM InsulationSealer
                                WHERE SealerID = ''' + str(i))
                        data1.append(Dbase().Dsqldata(qry)[0][0])
                n += 1
            data1.append(str(self.data[segn][30]))  # Insulation Note

            # generate the data for the various diameter ranges
            qry = []
            qry.append('SELECT Pipe_OD FROM Pipe_OD WHERE PipeOD_ID = ')
            qry.append('SELECT Pipe_OD FROM Pipe_OD WHERE PipeOD_ID = ')
            qry.append('''SELECT Material FROM InsulationMaterial
                        WHERE MatrID = ''')
            qry.append('''SELECT Thickness FROM InsulationThickness
                        WHERE ThkID = ''')

            data2 = []
            data2.extend(range(len(self.data[segn][31].split('.'))-5))

            for s in range(1, 5):
                for n in [m*4+s for m in range(0, 6)]:
                    ID = self.data[segn][n]
                    if ID is None:
                        break
                    data2[n] = str(Dbase().Dsqldata(qry[s-1] + str(ID))[0][0])
            del data2[0]

            rptdata.append(tuple(data1))
            rptdata.append(tuple(data2))

        return (Colwdths, Colnames, rptdata)

    def NewTextStr(self):
        s = ''
        orderly = sorted(self.txtparts)
        for n in orderly:
            s = s + str(self.txtparts[n]) + '.'
        s = s[:-1]
        return s

    # this builds the Fitting Code text string
    def LblData(self, tbl_name, cmbvalue):
        query3 = ''

        if cmbvalue == '':
            return '0'

        if tbl_name == 'InsulationJacket':
            # this is used if the combobox shows a numberic value
            query3 = ('SELECT JacketID FROM ' + tbl_name +
                      ' WHERE JacketMatr = "' + cmbvalue + '"')
        elif tbl_name == 'Pipe_OD':
            # this query is used when the combobox shows a string value
            query3 = ("SELECT PipeOD_ID FROM " + tbl_name +
                      " WHERE Pipe_OD = '" + cmbvalue + "'")
        elif tbl_name == 'PaintPrep':
            query3 = ('SELECT PaintPrepID FROM ' + tbl_name +
                      " WHERE Sandblast_Spec = '" + cmbvalue + "'")
        elif tbl_name == 'InsulationClass':
            query3 = ('SELECT ClassID FROM ' + tbl_name +
                      ' WHERE Insulation_Class = "' + cmbvalue + '"')
        elif tbl_name == 'InsulationAdhesive':
            query3 = ('SELECT AdhesiveID FROM ' + tbl_name +
                      " WHERE Adhesive = '" + cmbvalue + "'")
        elif tbl_name == 'InsulationSealer':
            query3 = ('SELECT SealerID FROM ' + tbl_name +
                      " WHERE Sealer = '" + cmbvalue + "'")
        elif tbl_name == 'InsulationMaterial':
            query3 = ('SELECT MatrID FROM ' + tbl_name +
                      " WHERE Material = '" + cmbvalue + "'")
        elif tbl_name == 'InsulationThickness':
            query3 = ('SELECT ThkID FROM ' + tbl_name +
                      " WHERE Thickness = '" + cmbvalue + "'")

        if query3 != '':
            lbldata = Dbase().Dsqldata(query3)
            lblstr = str(lbldata[0][0])

        return lblstr

    def OnDeleteRec(self, evt):
        recrd = self.data[self.rec_num][0]
        if self.ComdPrtyID is None:
            try:
                Dbase().TblDelete(self.tblname, recrd, self.pkcol_name)
                self.Dsql = 'SELECT * FROM ' + self.tblname
                self.data = Dbase().Dsqldata(self.Dsql)
                self.rec_num -= 1
                if self.rec_num < 0:
                    self.rec_num = len(self.data)-1
                self.FillScreen()
                self.recnum3.SetLabel('/ '+str(len(self.data)))

            except sqlite3.IntegrityError:
                wx.MessageBox("This Record is associated"
                              " with\nother tables and cannot be deleted!",
                              "Cannot Delete",
                              wx.OK | wx.ICON_INFORMATION)
        else:
            self.ChgSpecID()
            self.RestoreCmbs()
            self.b3.Disable()
            self.data = []
            self.rec_num = 0
            self.recnum2.SetLabel(str(self.rec_num))
            self.recnum3.SetLabel('/ '+str(len(self.data)))
            self.lbl_nodata.SetLabel(
                'The ' + self.frmtitle +
                ' have not been setup for this Commodity Property')

    def OnAdd1(self, evt):
        btn = evt.GetEventObject()
        callbtn = self.btnDict[btn]
        # numbers are for the index number for each group of comboboxes
        boxnums = [n*self.Num_Cols+callbtn for n in range(0, self.Num_Rows)]
        cmbtbl = self.cmb_tbls[callbtn]
        CmbLst1(self, cmbtbl)
        self.ReFillList(cmbtbl, boxnums)

    def ReFillList(self, cmbtbl, boxnums):
        if len(boxnums) > 1:
            for n in boxnums:
                lctr = self.lctrls[n+5]
                lctr.DeleteAllItems()
                index = 0
                ReFillQuery = 'SELECT * FROM "' + cmbtbl + '"'
                for values in Dbase().Dsqldata(ReFillQuery):
                    col = 0
                    for value in values:
                        if col == 0:
                            lctr.InsertItem(index, str(value))
                        else:
                            lctr.SetItem(index, col, str(value))
                        col += 1
                    index += 1
        else:
            lctr = self.lctrls[boxnums[0]]
            index = 0
            ReFillQuery = 'SELECT * FROM "' + cmbtbl + '"'
            for values in Dbase().Dsqldata(ReFillQuery):
                col = 0
                for value in values:
                    if col == 0:
                        lctr.InsertItem(index, str(value))
                    else:
                        lctr.SetItem(index, col, str(value))
                    col += 1
                index += 1

    def RestoreCmbs(self):
        n = 0
        self.restore = True
        for cmbbox in self.cmbctrls:
            cmbbox.Clear()
            cmbbox.ChangeValue('')
            cmbbox.Disable()
            n += 1
        for m in range(0, self.Num_Cols):
            self.cmbctrls[m].Enable()

        self.cmbAdh.ChangeValue('')
        self.cmbCls.ChangeValue('')
        self.cmbJkt.ChangeValue('')
        self.cmbSel.ChangeValue('')
        self.cmbSuf.ChangeValue('')
        self.notes.SetLabel('')

        temptxt = self.txtparts[0]
        self.txtparts = {}
        for i in range(0, 30):
            self.txtparts[i] = 0
        self.txtparts[0] = temptxt
        self.text4.ChangeValue(str(self.tblname) + ' Code')
        self.b2.Disable()
        self.b6.Enable()
        self.restore = False

    def ValData(self):
        # validate the data and truncate incomplete comboboxes
        #  digstr = []
        DataStrg = []
        #  NoData = 0
        #  DialogStr = ''
        data_end = (self.Num_Cols*self.Num_Rows)

        # builds tuple of the max dia combobox indexs (1,5,9,13,17,21)
        # this builds a tuple of all the last combobox indexs in each
        # column for forgings
        col_num = 1
        row_num = 0
        # start the aray list all the combo boxes in the first
        # combo box(0) starting with the second row(1)
        # in this case all the minimum diameter cmbboxes
        aray = [n*self.Num_Cols+col_num for n in range(row_num, self.Num_Rows)]
        aray = tuple(aray)

        # all the cmb's in the lower grid
        for i in range(0, self.Num_Cols*self.Num_Cols):
            cmbvalue = str(self.cmbctrls[i].GetValue())
            # at first empty combobox note the location and exist the FOR loop
            if cmbvalue == '':
                #  NoData = 2**i
                lastcmb = i
                break
            else:
                DataStrg.append(cmbvalue)
                lastcmb = self.Num_Cols*self.Num_Rows

            # if both diameters are complete compare them and confirm Max > Min
            if i in aray:
                if (self.cmbctrls[i-1].GetValue() != '' or
                        self.cmbctrls[i].GetValue() != ''):
                    test = (eval(cmbvalue.replace('"', '').replace('-', '+')) -
                            eval(self.cmbctrls[i-1].GetValue().replace('"', '')
                            .replace('-', '+')))
                    if test <= 0:
                        wx.MessageBox('''Minimum Diameter must be less than
                                       Maximum Diameter.''', 'Diameter Error',
                                      wx.OK | wx.ICON_INFORMATION)
                        self.cmbctrls[i].ChangeValue('')
                        return False

        if lastcmb/self.Num_Cols < 1:  # any cmb in the first row
            DataStrg = []
        else:
            # if first empty cmb is anywhere but the start of a new row
            if(lastcmb % self.Num_Cols):
                if lastcmb == ((self.Num_Cols * self.Num_Rows)
                               and self.cmbctrls[lastcmb]) != '':
                    data_end = lastcmb
                    MsgBx = wx.MessageDialog(
                        self, 'Value needed for ' + self.columnames[lastcmb] +
                        '''\n OK will delete incomplete row of data.\n CANCEL
                         will return for data correction.''',
                        'Missing Data', wx.OK | wx.CANCEL | wx.ICON_HAND)

                    MsgBx_val = MsgBx.ShowModal()
                    MsgBx.Destroy()
                    if MsgBx_val == wx.ID_CANCEL:
                        return False
                    # reverts back to the end of the previous row
                    data_end = (lastcmb//self.Num_Cols)*self.Num_Cols
            else:  # this represents the start of a new row (5,10,15)
                data_end = lastcmb

        # truncate the Fitting Code to reflect
        # the last completed row of comboboxes
        DataStrg = DataStrg[:data_end]

        # this builds a tuple of all
        # the last combobox index in each row
        col_num = 0
        row_num = 0
        # start the aray list all the combo boxes in
        # the first combo box(3) starting with the second row(0)
        aray = [n*self.Num_Cols+col_num for n in range(row_num, self.Num_Rows)]
        aray = tuple(aray)
        # lastcmb box completed is the end box then do nothing
        if lastcmb != max(aray):
            for i in aray:
                # this is the end box of the last completed row
                if lastcmb >= i:
                    end_data = i
            for m in range(end_data+1, (self.Num_Cols*self.Num_Rows)+1):
                if m in self.txtparts:
                    del self.txtparts[m]

        code = self.NewTextStr()
        self.text4.ChangeValue(code)
        self.b2.Disable()
        return True

    def Search(self):
        ShQuery = ('SELECT InsulationID FROM Insulation WHERE Code = "'
                   + self.text4.GetValue() + '"')
        existing = Dbase().Search(ShQuery)

        if existing:
            self.b2.SetLabel("Existing Spec")
        else:
            if self.ComdPrtyID is not None:
                self.b2.SetLabel('Add/Update\nto ' + self.frmtitle)
                self.b3.SetLabel('Delete From\nCommodity')
            else:
                self.b2.SetLabel('Add/Update\nto ' + self.frmtitle)

        return existing

    def OnAddComd(self, evt):
        self.AddComd()

    # link this ID to the pipespec for this commodity
    def AddComd(self):
        if self.b6.GetLabel() == 'Show All\nItems':
            if self.DsqlFtg == '':
                self.DsqlFtg = 'SELECT * FROM ' + self.tblname
                self.data = Dbase().Dsqldata(self.DsqlFtg)
            else:
                self.DsqlFtg = self.DsqlFtg[:self.DsqlFtg.find('WHERE')]
                self.data = Dbase().Dsqldata(self.DsqlFtg)
            self.b6.SetLabel("Add Item\nto Commodity")
            self.b6.Enable()
            self.b3.Disable()
            self.FillScreen()
            self.recnum3.SetLabel('/ '+str(len(self.data)))
        else:
            query = ('SELECT ' + self.tblname + '_ID' +
                     ' FROM PipeSpecification WHERE Commodity_Property_ID = '
                     + str(self.ComdPrtyID))
            StartQry = Dbase().Dsqldata(query)
            if StartQry == []:
                ValueList = []
                New_ID = cursr.execute(
                    '''SELECT MAX(Pipe_Spec_ID) FROM
                     PipeSpecification''').fetchone()
                if New_ID[0] is None:
                    Max_ID = '1'
                else:
                    Max_ID = str(New_ID[0]+1)
                colinfo = Dbase().Dcolinfo('PipeSpecification')
                for n in range(0, len(colinfo)-2):
                    ValueList.append(None)

                num_vals = ('?,'*len(colinfo))[:-1]
                ValueList.insert(0, Max_ID)
                ValueList.insert(1, str(self.ComdPrtyID))

                UpQuery = ("INSERT INTO PipeSpecification VALUES (" +
                           num_vals + ")")
                Dbase().TblEdit(UpQuery, ValueList)
                StartQry = Max_ID
            else:
                StartQry = str(StartQry[0][0])

            cmd_addID = self.data[self.rec_num][0]
            self.ChgSpecID(cmd_addID)

            self.DsqlFtg = ('SELECT * FROM ' + self.tblname + ' WHERE ' +
                            self.pkcol_name + ' = ' + str(cmd_addID))
            self.data = Dbase().Dsqldata(self.DsqlFtg)

            self.rec_num = 0
            self.FillScreen()
            self.lbl_nodata.SetLabel('   ')
            self.recnum3.SetLabel('/ '+str(len(self.data)))
            self.b3.Enable()
            self.b6.SetLabel("Show All\nItems")

    def AddRec(self, SQL_step):
        code = self.text4.GetValue()
        data_lst = []

        if SQL_step == 0:  # enter new record

            # get the next available autoincremented value for the table
            IDE = cursr.execute(
                "SELECT MAX(" + self.pkcol_name +
                ") FROM " + self.tblname).fetchone()[0]
            if IDE is None:
                IDE = 0
            IDE = int(IDE+1)

            # build the data string based on the Insulation Code
            txtparts = code.split('.')
            del txtparts[0]  # remove code type indicator
            # use all but the last five items which represent
            # the upper cbmbxs on the form
            data_lst = txtparts[:-5]
            for i in range(len(data_lst), self.Num_Cols*self.Num_Rows):
                # add all the lower cmbbxs to the datalist
                data_lst.append(None)

            # now add back into the data list the upper cmbbxs
            data_lst.extend(txtparts[len(txtparts)-5:])
            # retreive the note and code values and append to the data list
            data_lst = [None if x == '0' else x for x in data_lst]
            data_lst.append(self.notes.GetValue())
            data_lst.append(code)
            # insert the newest ID number
            data_lst.insert(0, str(IDE))

            # generate a string of ?'s to represent the database field names
            num_vals = ('?,'*len(self.colms))[:-1]

            CurrentID = IDE
            UpQuery = ("INSERT INTO " + self.tblname +
                       " VALUES (" + num_vals + ")")
            Dbase().TblEdit(UpQuery, data_lst)

        elif SQL_step == 1:  # update edited record
            realnames = []
            for item in Dbase().Dcolinfo(self.tblname):
                realnames.append(item[1])

            CurrentID = self.data[self.rec_num][0]

            # remove the name InsulationID from list as it is not updating
            realnames.remove(self.pkcol_name)
            # break the insulation code down into string of individual IDs
            txtparts = code.split('.')
            del txtparts[0]
            data_lst = txtparts[:-5]
            for i in range(len(data_lst), self.Num_Cols*self.Num_Rows):
                data_lst.append(None)

            data_lst.extend(txtparts[len(txtparts)-5:])

            data_lst = [None if x == '0' else x for x in data_lst]
            data_lst.append(self.notes.GetValue())
            data_lst.append(code)

            Name_str = ','.join(["%s=?" % (name) for name in realnames])
            UpQuery = ('UPDATE ' + self.tblname + ' SET ' + Name_str +
                       ' WHERE ' + self.pkcol_name + ' = ' + str(CurrentID))
            Dbase().TblEdit(UpQuery, data_lst)

        # cancel the update step
        elif SQL_step == 3:
            return

        self.b2.Disable()
        return CurrentID

    def OnCmbTypeOpen(self, evt):
        self.cmbTypeOld = evt.GetEventObject().GetValue()

    def OnCmbTypeClose(self, evt):
        cmbvalue = evt.GetEventObject().GetValue()
        if cmbvalue != self.cmbTypeOld:
            spot = evt.GetEventObject().Id
            tbl = self.cmb_tbls2[spot-25]
            self.CmbTypeClose(cmbvalue, spot, tbl)

    def CmbTypeClose(self, cmbvalue, spot, tbl):
        if tbl == 'InsulationJacket':
            query = ('''SELECT Jacket_Thk FROM InsulationJacket
                      WHERE JacketMatr = "''' + cmbvalue + '"')
            InslThk = Dbase().Dsqldata(query)[0][0]
            if InslThk is None:
                InslThk = 'Not Specified'
            strg1 = 'Jacket Thickness = '
            self.txtJkt.SetLabel(strg1 + InslThk)
        elif tbl == 'PaintPrep':
            query = ('SELECT Notes FROM PaintPrep WHERE SandBlast_Spec = "'
                     + cmbvalue + '"')
            Prep = Dbase().Dsqldata(query)[0][0]
            if Prep is None:
                Prep = 'Not Specified'
            self.txtSuf.SetLabel(Prep)

        newlbl = self.LblData(tbl, cmbvalue)
        if newlbl:
            self.txtparts[spot] = newlbl
        self.b2.Enable()
        self.text4.ChangeValue(self.NewTextStr())

    def OnAddJkt(self, evt):
        boxnums = [0]
        CmbLst1(self, 'InsulationJacket')
        self.ReFillList('InsulationJacket', boxnums)

    def OnAddSuf(self, evt):
        boxnums = [1]
        CmbLst1(self, 'PaintPrep')
        self.ReFillList('PaintPrep', boxnums)

    def OnAddCls(self, evt):
        boxnums = [2]
        CmbLst1(self, 'InsulationClass')
        self.ReFillList('InsulationClass', boxnums)

    def OnAddAdh(self, evt):
        boxnums = [3]
        CmbLst1(self, 'InsulationAdhesive')
        self.ReFillList('InsulationAdhesive', boxnums)

    def OnAddSel(self, evt):
        boxnums = [4]
        CmbLst1(self, 'InsulationSealer')
        self.ReFillList('InsulationSealer', boxnums)

    def OnCmbOpen(self, evt):
        self.cmbOld = evt.GetEventObject().GetValue()

    def OnCmbClose(self, evt):
        cmbNew = evt.GetEventObject().GetValue()
        if self.restore is False and self.cmbOld != cmbNew:
            self.CmbClose(evt.GetEventObject())
            self.b2.Enable()
            if self.ComdPrtyID is not None:
                self.b6.Disable()

    def CmbClose(self, cmbselect):
        i = int()

        # this builds a tuple of all the last combobox
        # indexs in each column for forgings
        col_num = 0
        row_num = 1
        # start the aray list all the combo boxes in the first
        # combo box(0) starting with the second row(1)
        aray = [n*self.Num_Cols+col_num for n in range(row_num, self.Num_Rows)]
        aray = tuple(aray)

        for i in (i for i, x in enumerate(self.cmbctrls) if x == cmbselect):
            cmbvalue = str(self.cmbctrls[i].GetValue())

        # build the string for the item code
        if cmbvalue != '':
            b = i % self.Num_Cols
            newlbl = self.LblData(self.cmb_tbls[b], cmbvalue)

            if newlbl:
                self.txtparts[1+i] = newlbl
            elif (i+1) in self.txtparts:
                del self.txtparts[i+1]
        else:
            if (i+1) in self.txtparts:
                self.txtparts[1+i] = '0'

        self.text4.ChangeValue(self.NewTextStr())

        # row has been completed enable the next row of boxes
        if (len(self.txtparts)-1) in aray:
            for n in range((len(self.txtparts)-1), len(self.txtparts)+5):
                self.cmbctrls[n].Enable()

    def OnRestoreCmbs(self, evt):
        self.RestoreCmbs()

    def OnAddRec(self, evt):
        # check first that data is all present and clean up incomplete boxes
        check = self.ValData()
        if check:
            existing = self.Search()   # check to see if code is unique
            # if the fitting code already exist do nothing
            if existing:
                wx.MessageBox('This ' + self.frmtitle +
                              ' Code already exists.',
                              "Existing Record", wx.OK | wx.ICON_INFORMATION)
                return

            SQL_step = 3

            choice1 = ('1) Save this as a new ' + self.frmtitle
                       + ' Specification')
            choice2 = ('2) Update the existing ' + self.frmtitle
                       + ' Specification with this data')
            txt1 = ('''NOTE: Updating this information will be\nreflected
                     in all associated ''' + self.frmtitle)
            txt2 = (''' Specifications!\nRecommendation is to save as a new
                     specification.\n\n\tHow do you want to proceed?''')

            # if this is a not commodity related change
            if self.ComdPrtyID is None:
                if self.data == []:
                    SQL_step = 0
                else:
                    # Make a selection as to whether the record is to be a new
                    # or an update valve use a SingleChioce dialog to determine
                    # if data is new record or edited record
                    SQL_Dialog = wx.SingleChoiceDialog(
                        self, txt1+txt2, 'Information Has Changed',
                        [choice1, choice2], style=wx.CHOICEDLG_STYLE)
                    if SQL_Dialog.ShowModal() == wx.ID_OK:
                        SQL_step = SQL_Dialog.GetSelection()
                    SQL_Dialog.Destroy()

                self.AddRec(SQL_step)
                self.DsqlFtg = 'SELECT * FROM ' + self.tblname
                self.data = Dbase().Dsqldata(self.DsqlFtg)

            else:   # this is a commodity related change
                choice1 = choice1 + ' for this commodity?'
                choice2 = choice2 + ' for this commodity?'
                # use a SingleChioce dialog to determine if
                # data is new record or edited record
                SQL_Dialog = wx.SingleChoiceDialog(
                    self, txt1+txt2, 'Information Has Changed',
                    [choice1, choice2], style=wx.CHOICEDLG_STYLE)
                if SQL_Dialog.ShowModal() == wx.ID_OK:
                    SQL_step = SQL_Dialog.GetSelection()
                SQL_Dialog.Destroy()

                # always save as a new spec if commodity property is specified
                CurrentID = self.AddRec(SQL_step)
                # no matter the change over write or add the specification ID
                # to the PipeSpecification table under the
                # commodity property ID

                self.ChgSpecID(CurrentID)

                query = ('SELECT ' + self.tblname + '_ID' +
                         ''' FROM PipeSpecification WHERE
                          Commodity_Property_ID = ''' + str(self.ComdPrtyID))
                StartQry = Dbase().Dsqldata(query)
                self.DsqlFtg = ('SELECT * FROM ' + self.tblname + ' WHERE '
                                + self.pkcol_name + ' = '
                                + str(StartQry[0][0]))
                self.data = Dbase().Dsqldata(self.DsqlFtg)

            if SQL_step == 0:
                self.rec_num = len(self.data)-1
            self.FillScreen()
            self.recnum3.SetLabel('/ '+str(len(self.data)))

    def ChgSpecID(self, ID=None):
        UpQuery = ('UPDATE PipeSpecification SET ' + self.tblname +
                   '_ID=?  WHERE Commodity_Property_ID = '
                   + str(self.ComdPrtyID))
        Dbase().TblEdit(UpQuery, [ID])

    def OnClose(self, evt):
        # add 2 lines for child parent form
        self.GetParent().Enable(True)
        self.__eventLoop.Exit()
        self.Destroy()


class PipeMtrSpc(wx.Frame):
    '''Routine to build form and populate grid'''
    def __init__(self, parent, tblname):

        self.parent = parent
        self.tblname = tblname
        self.lctrls = []

        super(PipeMtrSpc, self).__init__(
                          parent, title="Pipe Specification Code",
                          style=wx.DEFAULT_FRAME_STYLE ^ wx.RESIZE_BORDER,
                          size=(680, 320))

        self.Bind(wx.EVT_CLOSE, self.OnClose)

        self.search = False
        self.txtparts = {}
        font1 = wx.Font(10, wx.DEFAULT, wx.NORMAL, wx.FONTWEIGHT_BOLD)
        # set the Sizer property (same as SetSizer)
        self.Sizer = wx.BoxSizer(wx.VERTICAL)

        self.lblsizer = wx.BoxSizer(wx.HORIZONTAL)
        font = wx.Font(20, wx.DEFAULT, wx.NORMAL, wx.BOLD)

        self.text1 = wx.ComboCtrl(self, pos=(10, 10), size=(180, 40),
                                  style=wx.CB_READONLY)
        self.text1.SetForegroundColour((255, 0, 0))
        self.text1.SetFont(font)
        self.text1.SetPopupControl(ListCtrlComboPopup('PipeMaterialSpec',
                                                      showcol=1,
                                                      lctrls=self.lctrls))
        self.Bind(wx.EVT_TEXT, self.OnTxtChg, self.text1)
        self.Bind(wx.EVT_COMBOBOX_CLOSEUP, self.PrprtyFill, self.text1)

        # add a button to call main form to search combo list data
        self.b1 = wx.Button(self, label="<= Search Data")
        self.Bind(wx.EVT_BUTTON, self.OnSearch, self.b1)

        self.b3 = wx.Button(self, label="Clear Boxes")
        self.Bind(wx.EVT_BUTTON, self.OnRestore, self.b3)

        self.lblsizer.Add(self.text1, 0, wx.ALL |
                          wx.ALIGN_CENTER_VERTICAL, border=5)

        self.lblsizer.Add(self.b1, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL,
                          border=5)
        self.lblsizer.Add(self.b3, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL,
                          border=5)

        self.Sizer.Add(self.lblsizer, 0, wx.ALL | wx.ALIGN_CENTER)

        self.dblsizer = wx.BoxSizer(wx.HORIZONTAL)

        self.notesizer = wx.BoxSizer(wx.HORIZONTAL)

        inst = 'In order to add a new pipe specification code:\n'
        inst = inst + '1) Use the ANSI Rating list box to select the'
        inst = inst + (''' pressure class.\n2) Then make a selection
                        in the Material Type\n''')
        inst = inst + '3) Now you can select the Material Grade\n'
        inst = inst + '4) Select the Corrosion Allowance from the list box\n'
        inst = inst + ('''5) The last digit is optional and specifies
                        any special case\n''')
        inst = inst + ('''6) Click the Search button to see if the spec
                        exists before saving''')

        self.note1 = wx.StaticText(self, label=inst, style=wx.ALIGN_LEFT)
        self.note1.SetForegroundColour((255, 0, 0))

        self.notesizer.Add(self.note1, 0, wx.ALIGN_CENTER | wx.RIGHT,
                           border=15)

        self.cmbsizer2 = wx.BoxSizer(wx.VERTICAL)

        # develope the comboctrls and attach popup list
        self.cmb2 = wx.ComboCtrl(self, pos=(10, 10), size=(200, -1))
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.cmb2)
        self.cmb2.SetHint('ANSI Rating')
        self.cmb2.SetPopupControl(ListCtrlComboPopup('ANSI_Rating',
                                                     showcol=1,
                                                     lctrls=self.lctrls))
        self.cmbsizer2.Add(self.cmb2, 0, wx.ALIGN_LEFT)

        self.cmb3 = wx.ComboCtrl(self, pos=(10, 10), size=(200, -1))
        self.Bind(wx.EVT_TEXT, self.OnSelect3, self.cmb3)
        self.cmb3.SetHint('Material Type')
        self.cmb3.SetPopupControl(ListCtrlComboPopup('MaterialType',
                                                     showcol=1,
                                                     lctrls=self.lctrls))
        self.cmbsizer2.Add(self.cmb3, 0, wx.ALIGN_LEFT)

        self.cmb4 = wx.ComboCtrl(self, pos=(10, 10), size=(200, -1))
        self.Bind(wx.EVT_TEXT, self.OnSelect4, self.cmb4)
        self.cmb4.SetHint('Material Grade')
        self.cmb4.SetPopupControl(ListCtrlComboPopup('MaterialGrade',
                                                     showcol=3,
                                                     lctrls=self.lctrls))
        self.cmbsizer2.Add(self.cmb4, 0, wx.ALIGN_LEFT)
        self.cmb4.Disable()

        self.cmb5 = wx.ComboCtrl(self, pos=(10, 10), size=(200, -1))
        self.Bind(wx.EVT_TEXT, self.OnSelect5, self.cmb5)
        self.cmb5.SetHint('Corrosion Allowance')
        self.cmb5.SetPopupControl(ListCtrlComboPopup('CorrosionAllowance',
                                                     showcol=1,
                                                     lctrls=self.lctrls))
        self.cmbsizer2.Add(self.cmb5, 0, wx.ALIGN_LEFT)

        self.cmb6 = wx.ComboCtrl(self, pos=(10, 10), size=(200, -1))
        self.Bind(wx.EVT_TEXT, self.OnSelect6, self.cmb6)
        self.cmb6.SetHint('Special Case')
        self.cmb6.SetPopupControl(ListCtrlComboPopup('SpecialCase',
                                                     showcol=2,
                                                     lctrls=self.lctrls))
        self.cmbsizer2.Add(self.cmb6, 0, wx.ALIGN_LEFT)

        self.addsizer = wx.BoxSizer(wx.VERTICAL)

        self.addClass = wx.Button(self, label='+', size=(35, -1))
        self.addClass.SetForegroundColour((255, 0, 0))
        self.addClass.SetFont(font1)
        self.Bind(wx.EVT_BUTTON, self.OnAddClass, self.addClass)
        self.addsizer.Add(self.addClass, 0, wx.ALIGN_LEFT)

        self.addType = wx.Button(self, label='+', size=(35, -1))
        self.addType.SetForegroundColour((255, 0, 0))
        self.addType.SetFont(font1)
        self.Bind(wx.EVT_BUTTON, self.OnAddType, self.addType)
        self.addsizer.Add(self.addType, 0, wx.ALIGN_LEFT)

        self.addGrade = wx.Button(self, label='+', size=(35, -1))
        self.addGrade.SetForegroundColour((255, 0, 0))
        self.addGrade.SetFont(font1)
        self.addGrade.Disable()
        self.Bind(wx.EVT_BUTTON, self.OnAddGrade, self.addGrade)
        self.addsizer.Add(self.addGrade, 0, wx.ALIGN_LEFT)

        self.addCor = wx.Button(self, label='+', size=(35, -1))
        self.addCor.SetForegroundColour((255, 0, 0))
        self.addCor.SetFont(font1)
        self.Bind(wx.EVT_BUTTON, self.OnAddCor, self.addCor)
        self.addsizer.Add(self.addCor, 0, wx.ALIGN_LEFT)

        self.addSpec = wx.Button(self, label='+', size=(35, -1))
        self.addSpec.SetForegroundColour((255, 0, 0))
        self.addSpec.SetFont(font1)
        self.Bind(wx.EVT_BUTTON, self.OnAddSpecial, self.addSpec)
        self.addsizer.Add(self.addSpec, 0, wx.ALIGN_LEFT)

        self.dblsizer.Add(self.notesizer, 0)
        self.dblsizer.Add(self.cmbsizer2, 0)
        self.dblsizer.Add(self.addsizer, 0)

        self.Sizer.Add(self.dblsizer, 0, wx.ALIGN_CENTER)

        # Add buttons for grid modifications
        self.b2 = wx.Button(self, label="Build\nSpecificaton")
        self.b2.Disable()
        self.Bind(wx.EVT_BUTTON, self.OnAddSpec, self.b2)
        self.b4 = wx.Button(self, label="Exit")
        self.Bind(wx.EVT_BUTTON, self.OnClose, self.b4)
        self.b5 = wx.Button(self, label="Delete")
        self.Bind(wx.EVT_BUTTON, self.OnDelete, self.b5)
        self.b6 = wx.Button(self, label="Print\nList")
        self.Bind(wx.EVT_BUTTON, self.OnPrint, self.b6)
        # add a button box and place the buttons
        self.btnbox = wx.BoxSizer(wx.HORIZONTAL)
        self.btnbox.Add(self.b5, 0, wx.ALL, 5)
        self.btnbox.Add(self.b2, 0, wx.ALL, 5)
        self.btnbox.Add(self.b6, 0, wx.ALL, 5)
        self.btnbox.Add(self.b4, 0, wx.ALL, 5)

        self.Sizer.Add(self.btnbox, 0, wx.ALIGN_CENTER, 5)
        self.b4.SetFocus()

        self.CenterOnParent()
        self.GetParent().Enable(False)
        self.Show(True)
        self.__eventLoop = wx.GUIEventLoop()
        self.__eventLoop.Run()

    def LblData(self, tbl_name, cmbvalue, condition=None):

        self.query3 = ''

        if cmbvalue == "":
            return

        if tbl_name == 'ANSI_Rating':
            self.query3 = ('SELECT Rating_Designation FROM ' + tbl_name +
                           ' WHERE ANSI_Class = ' + cmbvalue)
        elif tbl_name == 'MaterialType':
            self.query3 = ('SELECT Type_Designation FROM ' + tbl_name +
                           " WHERE Material_Type = '" + cmbvalue + "'")
        elif tbl_name == 'MaterialGrade':
            self.query3 = ('SELECT Material_Grade_Designation FROM ' + tbl_name
                           + " WHERE Material_Grade = '" + cmbvalue + "'")
        elif tbl_name == 'CorrosionAllowance':
            self.query3 = ('SELECT Corrosion_Designation_ID FROM ' + tbl_name
                           + ' WHERE Corrosion_Allowance = ' + cmbvalue)
        elif tbl_name == 'SpecialCase':
            self.query3 = ('SELECT Case_Designation FROM ' + tbl_name +
                           " WHERE Case_Description = '" + cmbvalue + "'")
        lbldata = Dbase().Dsqldata(self.query3)
        lblstr = str(lbldata[0][0])
        return lblstr

    def EditTbl(self):
        # rowID is table index value, colID is column number of table index
        # new_value is the link foreign field value
        # colChgNum is the column number which has edited cell,
        # values is the new edited value(s)

        values = self.text1.GetValue()
        if len(values) >= 4:
            #  added_row = 1
            realnames = []
            for item in Dbase().Dcolinfo(self.tblname):
                realnames.append(item[1])
            #  enable_b2 = 0

            ValueList = []
            if type(values) == str:
                ValueList.append(values)
            else:
                ValueList = values
            # if the table index is auto increment then assign next
            # value otherwise do nothing
            for item in Dbase().Dcolinfo(self.tblname):
                if item[5] == 1:
                    IDcol = item[0]
                    IDname = item[1]
                    if 'INTEGER' in item[2]:
                        New_ID = cursr.execute(
                            "SELECT MAX(" + IDname + ") FROM " +
                            self.tblname).fetchone()
                        if New_ID[0] is None:
                            Max_ID = '1'
                        else:
                            Max_ID = str(New_ID[0]+1)
                        if type(values) == str:
                            ValueList.insert(IDcol, Max_ID)

            UpQuery = ('INSERT INTO ' + self.tblname + ' VALUES (' + "'" +
                       "','".join(map(str, ValueList)) + "'" + ')')

            self.RestoreForm()
            Dbase().TblEdit(UpQuery)
        else:
            wx.MessageBox("The Pipe Specification is not complete!",
                          "Cannot Be Added", wx.OK | wx.ICON_INFORMATION)

    def OnAddClass(self, evt):
        boxnums = 1
        CmbLst1(self, 'ANSI_Rating')
        self.ReFillList('ANSI_Rating', boxnums)

    def OnAddType(self, evt):
        boxnums = 2
        CmbLst1(self, 'MaterialType')
        self.ReFillList('MaterialType', boxnums)

    def OnAddGrade(self, evt):
        boxnums = 3
        BldMtr(self, 'MaterialGrade')
        self.ReFillList('MaterialGrade', boxnums)

    def OnAddCor(self, evt):
        boxnums = 4
        CmbLst1(self, 'CorrosionAllowance')
        self.ReFillList('CorrosionAllowance', boxnums)

    def OnAddSpecial(self, evt):
        boxnums = 5
        CmbLst1(self, 'SpecialCase')
        self.ReFillList('SpecialCase', boxnums)

    def ReFillList(self, cmbtbl, boxnum):
        self.lc = self.lctrls[boxnum]
        if self.lc:
            self.lc.DeleteAllItems()
        index = 0

        ReFillQuery = 'SELECT * FROM "' + cmbtbl + '"'
        if cmbtbl == 'MaterialGrade':
            ReFillQuery = ('SELECT * FROM ' + cmbtbl +
                           ' WHERE Type_Designation = "' +
                           self.txtparts[1] + '"')
        for values in Dbase().Dsqldata(ReFillQuery):
            col = 0
            for value in values:
                if col == 0:
                    self.lc.InsertItem(index, str(value))
                else:
                    self.lc.SetItem(index, col, str(value))
                col += 1
            index += 1

    def OnAddSpec(self, evt):
        digstr = []
        DataStrg = []
        NoData = 0
        DialogStr = ''

    # before saving racord check that all the required data has been entered
        if self.cmb2.GetValue() == '':
            # each record is assigned value of binary number
            # 000001 = 1 for no data
            NoData = 1
        else:
            # develop list of each data item
            DataStrg.append(self.cmb2.GetValue())

        if self.cmb3.GetValue() == '':
            # next binary number level 000010 = 2 etc.
            NoData = 2 + NoData
        else:
            DataStrg.append(self.cmb3.GetValue())

        if self.cmb4.GetValue() == '':
            NoData = 4 + NoData
        else:
            DataStrg.append(self.cmb4.GetValue())

        if self.cmb5.GetValue() == '':
            NoData = 8 + NoData
        else:
            DataStrg.append(self.cmb5.GetValue())

        # check that all the records are present, if not use
        # the sum of the binary numbers to see which are missing
        if NoData > 0:
            # use the binary numbers as keys for the DataBxs dictionary
            DataBxs = {1: 'ANSI Rating', 2: 'Material Type',
                       3: 'Material Grade', 4: 'Corrosion Allowance'}
            binry = str('{0:09b}'.format(NoData))
            digstr = [pos for pos, char in enumerate(binry) if char == '1']
            for dig in digstr:
                DialogStr = DataBxs[9-dig] + ',\n' + DialogStr
            wx.MessageBox('Value(s) needed for;\n' + DialogStr, 'Missing Data',
                          wx.OK | wx.ICON_INFORMATION)
        else:
            self.EditTbl()
            self.ReFillList('PipeMaterialSpec', 0)

    def NewTextStr(self):
        s = ''
        orderly = sorted(self.txtparts)
        for n in orderly:
            s = s + str(self.txtparts[n])
        self.text1.ChangeValue(s)

    def RestoreForm(self):

        self.cmb2.ChangeValue('')
        self.cmb2.SetHint("ANSI Rating")
        self.cmb3.ChangeValue('')
        self.cmb3.SetHint("Material Type")
        self.cmb4.ChangeValue('')
        self.cmb4.SetHint("Material Grade")
        self.cmb4.Disable()
        self.addGrade.Disable()
        self.cmb5.ChangeValue('')
        self.cmb5.SetHint("Corrosion Allowance")
        self.cmb6.ChangeValue('')
        self.cmb6.SetHint("Special Case")
        self.b2.SetLabel("Build\nSpecificaton")
        self.b2.Disable()
        self.txtparts = {}
        self.text1.ChangeValue('')

    def OnSelect2(self, evt):
        if self.search is False:
            tbl = 'ANSI_Rating'
            cmbvalue = str(self.cmb2.GetValue())
            if cmbvalue != "":
                newlbl = self.LblData(tbl, cmbvalue)
                self.txtparts[1] = newlbl
                self.NewTextStr()
            else:
                self.txtparts[1] = ''
                self.NewTextStr()

    def OnSelect3(self, evt):
        if self.search is False:
            tbl = 'MaterialType'
            cmbvalue = self.cmb3.GetValue()
            if cmbvalue != "":
                newlbl = self.LblData(tbl, cmbvalue)
                self.txtparts[2] = newlbl
                self.NewTextStr()
                query = ('''SELECT * FROM MaterialGrade WHERE
                          Type_Designation = "''' + newlbl + '"')
                self.cmb4.Enable()
                self.addGrade.Enable()
                self.cmb4.SetValue('')
                self.cmb4.SetPopupControl(
                    ListCtrlComboPopup('MaterialGrade', query, cmbvalue,
                                       showcol=3))
                self.cmb4.SetHint("Material Grade")
            else:
                self.txtparts[2] = ''
                self.NewTextStr()

    def OnSelect4(self, evt):
        if self.search is False:
            tbl = 'MaterialGrade'
            cmbvalue = str(self.cmb4.GetValue())
            if cmbvalue != "":
                newlbl = self.LblData(tbl, cmbvalue)
                self.txtparts[3] = newlbl
                self.NewTextStr()
            else:
                self.txtparts[3] = ''
                self.NewTextStr()

    def OnSelect5(self, evt):
        if self.search is False:
            tbl = 'CorrosionAllowance'
            cmbvalue = str(self.cmb5.GetValue())
            if cmbvalue != "":
                newlbl = self.LblData(tbl, cmbvalue)
                self.txtparts[4] = newlbl
                self.NewTextStr()
            else:
                self.txtparts[4] = ''
                self.NewTextStr()

    def OnSelect6(self, evt):
        if self.search is False:
            tbl = 'SpecialCase'
            cmbvalue = str(self.cmb6.GetValue())
            if cmbvalue != "":
                newlbl = self.LblData(tbl, cmbvalue)
                self.txtparts[5] = newlbl
                self.NewTextStr()
            else:
                self.txtparts[5] = ''
                self.NewTextStr()

    def OnTxtChg(self, evt):
        org_txt = self.text1.GetValue()
        new_txt = org_txt.upper()
        if len(new_txt) > 5:
            new_txt = new_txt[:5]
        self.text1.ChangeValue(new_txt)
        self.text1.SetInsertionPointEnd()
        self.text1.Refresh()

    def PrprtyFill(self, evt):
        lbl = self.text1.GetValue()
        if lbl != '':
            self.cmb6.ChangeValue('')
            qry = ('''SELECT ANSI_Class FROM ANSI_Rating WHERE
                    Rating_Designation = ''' + lbl[0])
            self.cmb2.ChangeValue(str(Dbase().Dsqldata(qry)[0][0]))
            qry = ('''SELECT Material_Type FROM MaterialType WHERE
                    Type_Designation = "''' + lbl[1] + '"')
            self.cmb3.ChangeValue(str(Dbase().Dsqldata(qry)[0][0]))
            qry = ('''SELECT Material_Grade FROM MaterialGrade WHERE
                    Material_Grade_Designation = "''' + lbl[2] + '"' +
                   ' AND Type_Designation = "' + lbl[1] + '"')
            self.cmb4.ChangeValue(str(Dbase().Dsqldata(qry)[0][0]))
            self.cmb4.Enable()
            self.addGrade.Enable()

            qry = ('''SELECT Corrosion_Allowance FROM CorrosionAllowance WHERE
                    Corrosion_Designation_ID = "''' + lbl[3] + '"')
            self.cmb5.ChangeValue(str(Dbase().Dsqldata(qry)[0][0]))
            if len(lbl) > 4:
                qry = ('''SELECT Case_Description FROM SpecialCase
                        WHERE Case_Designation = "''' + lbl[4] + '"')
                self.cmb6.ChangeValue(str(Dbase().Dsqldata(qry)[0][0]))

    def OnSearch(self, evt):
        existing = None
        if (len(self.text1.GetValue()) >= 4 and
                (' ' in self.text1.GetValue()) is False):
            field = Dbase().Dcolinfo(self.tblname)[1][1]
            ShQuery = ('SELECT * FROM ' + self.tblname + ' WHERE ' + field +
                       ' = "' + self.text1.GetValue() + '"')
            existing = Dbase().Search(ShQuery)
            self.search = True

        if existing is not None:
            if existing:
                self.b2.SetLabel("Existing\nSpec")
            else:
                self.b2.SetLabel("Add\nSpecification")
                self.b2.Enable()

        self.search = False

    def OnDelete(self, evt):
        field = Dbase().Dcolinfo(self.tblname)[1][1]
        val = self.text1.GetValue()
        try:
            Dbase().TblDelete(self.tblname, val, field)
        except sqlite3.IntegrityError:
            wx.MessageBox('''This Record is associated with\nother tables
                           and cannot be deleted!''', "Cannot Delete",
                          wx.OK | wx.ICON_INFORMATION)
        self.RestoreForm()
        self.ReFillList('PipeMaterialSpec', 0)

    def OnPrint(self, evt):
        import Report_Lvl1

        saveDialog = wx.FileDialog(
            self, message='Save Report as PDF.',
            wildcard='PDF (*.pdf)|*.pdf',
            style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT)

        if saveDialog.ShowModal() == wx.ID_CANCEL:
            filename = ''

        filename = saveDialog.GetPath()
        if filename.find(".pdf") == -1:
            filename = filename + '.pdf'

        saveDialog.Destroy()
        data1 = []
        rptdata = []

        colnames = ['Pipe\nMaterial\nSpecification', 'ANSI Rating',
                    'Material Type', 'Material Grade',
                    'Corrosion\nAllowance\n(inches)', 'Special Case']

        colwdths = [10, 10, 20, 20, 20]

        qry = 'SELECT Pipe_Material_Spec FROM PipeMaterialSpec'
        mtrdata = Dbase().Dsqldata(qry)
        for seg in mtrdata:
            for i in range(len(seg[0])):
                if i == 0:
                    qry = ('''SELECT ANSI_Class FROM ANSI_Rating WHERE
                            Rating_Designation = "''' + str(seg[0][i]) + '"')
                    new_data = Dbase().Dsqldata(qry)[0][0]
                elif i == 1:
                    qry = (
                        '''SELECT Material_Type FROM MaterialType
                         WHERE Type_Designation = "''' + str(seg[0][i]) + '"')
                    new_data = Dbase().Dsqldata(qry)[0][0]
                elif i == 2:
                    qry = ('''SELECT Material_Grade FROM MaterialGrade
                            WHERE Material_Grade_Designation = "'''
                           + str(seg[0][i]) + '"')
                    new_data = Dbase().Dsqldata(qry)[0][0]
                elif i == 3:
                    qry = ('''SELECT Corrosion_Allowance FROM CorrosionAllowance
                            WHERE Corrosion_Designation_ID = "'''
                           + str(seg[0][i]) + '"')
                    new_data = Dbase().Dsqldata(qry)[0][0]
                elif i == 4:
                    qry = (
                        '''SELECT Case_Description FROM SpecialCase
                         WHERE Case_Designation = "''' + str(seg[0][i]) + '"')
                    new_data = Dbase().Dsqldata(qry)[0][0]
                data1.append(new_data)
            if len(seg[0]) == 4:
                data1.append(' ')
            data1.insert(0, seg[0])
            rptdata.append(tuple(data1))
            data1 = []

        Report_Lvl1.Report(self.tblname, rptdata, colnames,
                           filename, Colwdths=colwdths).create_pdf()

    def OnRestore(self, evt):
        self.RestoreForm()

    def OnClose(self, evt):
        self.GetParent().Enable(True)
        self.__eventLoop.Exit()
        self.Destroy()


class BldFst(wx.Frame):
    '''Routine to build form and populate grid'''
    def __init__(self, parent, tblname, ComdPrtyID=None):

        # add line for child parent form
        self.parent = parent
        self.tblname = tblname

        if self.tblname.find("_") != -1:
            self.frmtitle = (self.tblname.replace("_", " "))
        else:
            self.frmtitle = (' '.join(re.findall('([A-Z][a-z]*)',
                             self.tblname)))

        super(BldFst, self).__init__(parent,
                                     title=self.frmtitle,
                                     size=(730, 430))

        self.ComdPrtyID = ComdPrtyID
        self.PipeMtrSpec = ''
        self.ComCode = ''
        self.columnames = []
        self.lctrls = []
        self.data = []
        self.MainSQL = ''
        self.rec_num = 0
        StartQry = None

        self.Bind(wx.EVT_CLOSE, self.OnClose)

        font1 = wx.Font(12, wx.DEFAULT, wx.NORMAL, wx.FONTWEIGHT_BOLD)

        # setup label field if commodity does not have item specified
        font2 = wx.Font(12, wx.DEFAULT, wx.NORMAL, wx.FONTWEIGHT_BOLD)
        txt_nodata = (
            'Fasteners have not been setup for this Commodity Property')
        self.lbl_nodata = wx.StaticText(self, -1, label=txt_nodata,
                                        size=(550, -1), style=wx.CENTER)
        self.lbl_nodata.SetForegroundColour((255, 0, 0))
        self.lbl_nodata.SetFont(font2)
        self.lbl_nodata.SetLabel('   ')

        if self.ComdPrtyID is not None:
            query = ('''SELECT Commodity_Code,Pipe_Material_Code,End_Connection,
                     Pipe_Code FROM CommodityProperties WHERE
                      CommodityPropertyID = '''
                     + str(self.ComdPrtyID))
            dataset = Dbase().Dsqldata(query)[0]
            self.PipeMtrSpec = dataset[1]
            self.ComCode = dataset[0]
            self.PipeCode = dataset[3]

            query = ('''SELECT Fastener_ID FROM PipeSpecification
                      WHERE Commodity_Property_ID = ''' + str(self.ComdPrtyID))
            chk = Dbase().Dsqldata(query)
            if chk != []:
                StartQry = chk[0][0]
                if StartQry is not None:
                    self.MainSQL = ('SELECT * FROM ' + self.tblname
                                    + ' WHERE FastenerID = ' + str(StartQry))
                    self.data = Dbase().Dsqldata(self.MainSQL)
                else:   # no item setup for commodity property
                    self.lbl_nodata.SetLabel(txt_nodata)
            else:   # no item setup for commodity property
                self.lbl_nodata.SetLabel(txt_nodata)
        else:
            self.MainSQL = 'SELECT * FROM ' + self.tblname
            self.data = Dbase().Dsqldata(self.MainSQL)
        # specify which listbox column to display in the combobox
        self.showcol = int

        # set the Sizer property (same as SetSizer)
        self.Sizer = wx.BoxSizer(wx.VERTICAL)

        self.specsizer = wx.BoxSizer(wx.HORIZONTAL)
        font = wx.Font(10, wx.DEFAULT, wx.NORMAL, wx.NORMAL)
        # set up bxw to show that there is no item for commodity property
        self.warningsizer = wx.BoxSizer(wx.HORIZONTAL)
        self.warningsizer.Add(self.lbl_nodata, 0, wx.ALIGN_CENTER)

        if self.ComCode == '':
            self.specsizer.Add((20, 20))
        else:      # show commodity property ticker
            if self.PipeCode is not None:
                txt = self.PipeCode
            else:
                query = ('''SELECT * FROM PipeMaterialSpec WHERE
                          Material_Spec_ID = ''' + str(self.PipeMtrSpec))
                MtrSpc = Dbase().Dsqldata(query)[0][1]
                txt = self.ComCode + ' - ' + MtrSpc

            self.text1 = wx.TextCtrl(self, size=(100, 33), value=txt,
                                     style=wx.TE_READONLY | wx.TE_CENTER)
            self.text1.SetForegroundColour((255, 0, 0))
            self.text1.SetFont(font)
            self.specsizer.Add(self.text1, 0, wx.ALL, 10)

        self.notesizer = wx.BoxSizer(wx.HORIZONTAL)

        self.note1 = wx.StaticText(self, label='Stud / Bolt Material',
                                   style=wx.ALIGN_LEFT)
        self.note1.SetForegroundColour((255, 0, 0))

        self.addblt = wx.Button(self, label='+', size=(35, -1))
        self.addblt.SetForegroundColour((255, 0, 0))
        self.addblt.SetFont(font1)
        self.Bind(wx.EVT_BUTTON, self.OnAddBlt, self.addblt)

        self.note2 = wx.StaticText(self, label='Nut Material',
                                   style=wx.ALIGN_LEFT)
        self.note2.SetForegroundColour((255, 0, 0))

        self.addnut = wx.Button(self, label='+', size=(35, -1))
        self.addnut.SetForegroundColour((255, 0, 0))
        self.addnut.SetFont(font1)
        self.Bind(wx.EVT_BUTTON, self.OnAddNut, self.addnut)

        self.notesizer.Add(self.note1, 0, wx.LEFT | wx.ALIGN_CENTER, border=60)
        self.notesizer.Add(self.addblt, 0, wx.LEFT, 10)
        self.notesizer.Add(self.note2, 0, wx.LEFT | wx.ALIGN_CENTER, border=65)
        self.notesizer.Add(self.addnut, 0, wx.LEFT, 10)

        self.cmbsizer1 = wx.BoxSizer(wx.HORIZONTAL)

        self.note3 = wx.StaticText(self, label='Material Option 1',
                                   style=wx.ALIGN_CENTER_VERTICAL)
        self.note3.SetForegroundColour((255, 0, 0))

        self.blt1 = wx.ComboCtrl(self, id=1, pos=(10, 10), size=(200, -1),
                                 style=wx.CB_READONLY)
        self.Bind(wx.EVT_COMBOBOX_DROPDOWN, self.OnCmbBltOpen, self.blt1)
        self.Bind(wx.EVT_COMBOBOX_CLOSEUP, self.OnCmbBltClose, self.blt1)
        self.blt1.SetHint('Stud Material')
        self.blt1.SetPopupControl(ListCtrlComboPopup('BoltMaterial',
                                  showcol=1, lctrls=self.lctrls))

        self.nut1 = wx.ComboCtrl(self, id=2, pos=(10, 10),
                                 size=(200, -1), style=wx.CB_READONLY)
        self.Bind(wx.EVT_COMBOBOX_DROPDOWN, self.OnCmbNutOpen, self.nut1)
        self.Bind(wx.EVT_COMBOBOX_CLOSEUP, self.OnCmbNutClose, self.nut1)
        self.nut1.SetHint('Nut Material')
        self.nut1.SetPopupControl(ListCtrlComboPopup('NutMaterial',
                                  showcol=1, lctrls=self.lctrls))

        self.cmbsizer1.Add(self.note3, 0, wx.RIGHT | wx.TOP, border=8)
        self.cmbsizer1.Add(self.blt1, 0, wx.ALIGN_LEFT, 5)
        self.cmbsizer1.Add(self.nut1, 0, wx.ALIGN_LEFT, 5)

        self.cmbsizer2 = wx.BoxSizer(wx.HORIZONTAL)

        self.note4 = wx.StaticText(self, label='Material Option 2',
                                   style=wx.ALIGN_CENTER_VERTICAL)
        self.note4.SetForegroundColour((255, 0, 0))

        # develope the comboctrls and attach popup list
        self.blt2 = wx.ComboCtrl(self, id=3, pos=(10, 10), size=(200, -1),
                                 style=wx.CB_READONLY)
        self.Bind(wx.EVT_COMBOBOX_DROPDOWN, self.OnCmbBltOpen, self.blt2)
        self.Bind(wx.EVT_COMBOBOX_CLOSEUP, self.OnCmbBltClose, self.blt2)
        self.blt2.SetHint('Stud Material')
        self.blt2.SetPopupControl(ListCtrlComboPopup('BoltMaterial', showcol=1,
                                                     lctrls=self.lctrls))

        self.nut2 = wx.ComboCtrl(self, id=4, pos=(10, 10), size=(200, -1),
                                 style=wx.CB_READONLY)
        self.Bind(wx.EVT_COMBOBOX_DROPDOWN, self.OnCmbNutOpen, self.nut2)
        self.Bind(wx.EVT_COMBOBOX_CLOSEUP, self.OnCmbNutClose, self.nut2)
        self.nut2.SetHint('Nut Material')
        self.nut2.SetPopupControl(ListCtrlComboPopup('NutMaterial', showcol=1,
                                                     lctrls=self.lctrls))

        self.cmbsizer2.Add(self.note4, 0, wx.RIGHT | wx.TOP, border=8)
        self.cmbsizer2.Add(self.blt2, 0, wx.ALIGN_LEFT, 5)
        self.cmbsizer2.Add(self.nut2, 0, wx.ALIGN_LEFT, 5)

        self.Sizer.Add(self.warningsizer, 0, wx.CENTER | wx.TOP, 5)
        self.Sizer.Add(self.specsizer, 0, wx.ALL | wx.ALIGN_CENTER)
        self.Sizer.Add(self.notesizer, 0, wx.ALL | wx.ALIGN_CENTER)
        self.Sizer.Add(self.cmbsizer1, 0, wx.ALL | wx.ALIGN_CENTER, 10)
        self.Sizer.Add(self.cmbsizer2, 0, wx.ALL | wx.ALIGN_CENTER, 10)

        # Add buttons for grid modifications
        self.b1 = wx.Button(self, label="Print\nReport")
        self.Bind(wx.EVT_BUTTON, self.PrintFile, self.b1)
        self.b2 = wx.Button(self, label="Add/Update\nto " + self.frmtitle)
        self.b2.Disable()
        self.Bind(wx.EVT_BUTTON, self.OnAddRec, self.b2)

        self.b3 = wx.Button(self, label="Delete\nSpecification")
        self.Bind(wx.EVT_BUTTON, self.OnDelete, self.b3)

        self.b4 = wx.Button(self, label="Exit")
        self.Bind(wx.EVT_BUTTON, self.OnClose, self.b4)

        self.b5 = wx.Button(self, label="Clear\nBoxes")
        self.Bind(wx.EVT_BUTTON, self.OnRestoreCmbs, self.b5)

        # add a button box and place the buttons
        self.btnbox = wx.BoxSizer(wx.HORIZONTAL)
        self.btnbox.Add(self.b5, 0, wx.ALIGN_CENTER | wx.RIGHT, 20)
        self.btnbox.Add(self.b2, 0, wx.ALL, 5)
        if self.ComdPrtyID is not None:
            if self.tblname == 'Piping':
                lbl = 'Add/Update to\nCommodity'
            else:
                lbl = "Show All\nItems"
            self.b6 = wx.Button(self, size=(120, 45), label=lbl)
            self.b6.Enable()
            self.b3.SetLabel('Delete From\nCommodity')
            if StartQry is None:
                self.b3.Disable()
            self.Bind(wx.EVT_BUTTON, self.OnAddComd, self.b6)
            self.btnbox.Add(self.b6, 0, wx.ALL | wx.ALIGN_CENTER, 5)
        self.btnbox.Add(self.b3, 0, wx.ALL, 5)
        self.btnbox.Add(self.b1, 0, wx.ALL, 5)
        self.btnbox.Add((30, 10))
        self.btnbox.Add(self.b4, 0, wx.ALL, 5)

        self.fst = wx.Button(self, label='<<')
        self.lst = wx.Button(self, label='>>')
        self.nxt = wx.Button(self, label='>')
        self.pre = wx.Button(self, label='<')
        self.fst.Bind(wx.EVT_BUTTON, self.OnMovefst)
        self.lst.Bind(wx.EVT_BUTTON, self.OnMovelst)
        self.nxt.Bind(wx.EVT_BUTTON, self.OnMovenxt)
        self.pre.Bind(wx.EVT_BUTTON, self.OnMovepre)

        self.navbox = wx.BoxSizer(wx.HORIZONTAL)
        self.navbox.Add(self.fst, 0, wx.ALL, 5)
        self.navbox.Add(self.pre, 0, wx.ALL, 5)
        self.navbox.Add(self.nxt, 0, wx.ALL, 5)
        self.navbox.Add(self.lst, 0, wx.ALL, 5)

        self.numbox = wx.BoxSizer(wx.HORIZONTAL)
        self.recnum1 = wx.StaticText(self, label='Record ',
                                     style=wx.ALIGN_LEFT)
        self.recnum1.SetForegroundColour((255, 0, 0))

        self.recnum2 = wx.StaticText(self, label=str(self.rec_num + 1),
                                     style=wx.ALIGN_LEFT)
        self.recnum2.SetForegroundColour((255, 0, 0))
        self.recnum3 = wx.StaticText(self, label='/ ' + str(len(self.data)),
                                     style=wx.ALIGN_LEFT)
        self.recnum3.SetForegroundColour((255, 0, 0))
        self.numbox.Add(self.recnum1, 0, wx.ALL, 5)
        self.numbox.Add(self.recnum2, 0, wx.ALL, 5)
        self.numbox.Add(self.recnum3, 0, wx.ALL, 5)

        self.Sizer.Add(self.btnbox, 0, wx.ALIGN_CENTER | wx.TOP, 10)
        self.Sizer.Add(self.navbox, 0, wx.ALIGN_CENTER | wx.TOP, 10)
        self.Sizer.Add(self.numbox, 0, wx.ALIGN_CENTER | wx.TOP)
        self.SetSizer(self.Sizer)
        self.b4.SetFocus()

        self.FillScreen()

        # add these following 5 lines to child parent form
        self.CenterOnParent()
        self.GetParent().Enable(False)
        self.Show(True)
        self.__eventLoop = wx.GUIEventLoop()
        self.__eventLoop.Run()

    # called to update the item table and commodity table if needed
    def OnAddRec(self, evt):
        # check first that data is all present and clean up incomplete boxes
        check = self.ValData()
        if check:
            SQL_step = 3

            choice1 = ('1) Save this as a new ' + self.frmtitle +
                       ' Specification')
            choice2 = ('2) Update the existing ' + self.frmtitle +
                       ' Specification with this data')
            txt1 = ('''NOTE: Updating this information will be\nreflected
                     in all associated ''' + self.frmtitle)
            txt2 = (''' Specifications!\nRecommendation is to save as a
                     new specification.\n\n\tHow do you want to proceed?''')

            # if this is a not commodity related change
            if self.ComdPrtyID is None:
                if self.data == []:
                    SQL_step = 0
                else:
                    # Make a selection as to whether the record
                    # is to be a new or an update valve
                    # use a SingleChioce dialog to determine
                    # if data is new record or edited record
                    SQL_Dialog = wx.SingleChoiceDialog(
                        self, txt1+txt2, 'Information Has Changed',
                        [choice1, choice2], style=wx.CHOICEDLG_STYLE)
                    if SQL_Dialog.ShowModal() == wx.ID_OK:
                        SQL_step = SQL_Dialog.GetSelection()
                    SQL_Dialog.Destroy()

                self.AddRec(SQL_step)
                self.MainSQL = 'SELECT * FROM ' + self.tblname
                self.data = Dbase().Dsqldata(self.MainSQL)

            else:  # this is a commodity related change
                choice1 = choice1 + ' for this commodity?'
                choice2 = choice2 + ' and save for this commodity?'
                # use a SingleChioce dialog to determine if data
                # is new record or edited record
                SQL_Dialog = wx.SingleChoiceDialog(
                    self, txt1+txt2, 'Information Has Changed',
                    [choice1, choice2], style=wx.CHOICEDLG_STYLE)
                if SQL_Dialog.ShowModal() == wx.ID_OK:
                    SQL_step = SQL_Dialog.GetSelection()
                SQL_Dialog.Destroy()

                cmd_addID = self.AddRec(SQL_step)
                # always save as a new spec if commodity property is specified
                # no matter the change over write or add the specification ID
                # to the PipeSpecification table under the
                # commodity property ID

                self.ChgSpecID(cmd_addID)

                query = ('''SELECT Fastener_ID FROM PipeSpecification
                          WHERE Commodity_Property_ID = ''' +
                         str(self.ComdPrtyID))
                StartQry = Dbase().Dsqldata(query)
                self.MainSQL = ('SELECT * FROM ' + self.tblname +
                                ' WHERE FastenerID = ' + str(StartQry[0][0]))
                self.data = Dbase().Dsqldata(self.MainSQL)

            if SQL_step == 0:
                self.rec_num = len(self.data)-1
            self.FillScreen()
            self.recnum3.SetLabel('/ '+str(len(self.data)))

    def AddRec(self, SQL_step):
        realnames = []
        ValueList = []

        colinfo = Dbase().Dcolinfo(self.tblname)
        ValueList = [None for i in range(0, len(colinfo))]
        # if the table index is auto increment then assign
        # next value otherwise do nothing
        for item in colinfo:
            if item[5] == 1:
                #    IDcol = item[0]
                IDname = item[1]
                if 'INTEGER' in item[2]:
                    New_ID = cursr.execute(
                        "SELECT MAX(" + IDname + ") FROM " +
                        self.tblname).fetchone()
                    if New_ID[0] is None:
                        Max_ID = '1'
                    else:
                        Max_ID = str(New_ID[0]+1)
            realnames.append(item[1])
        ValueList[0] = str(Max_ID)

        qry = ("SELECT BoltID FROM BoltMaterial WHERE Bolt_Material = '" +
               self.blt1.GetValue() + "'")
        ValueList[1] = Dbase().Dsqldata(qry)[0][0]
        qry = ("SELECT NutID FROM NutMaterial WHERE Nut_Material = '" +
               self.nut1.GetValue() + "'")
        ValueList[2] = Dbase().Dsqldata(qry)[0][0]
        if self.blt2.GetValue() != '':
            qry = ("SELECT BoltID FROM BoltMaterial WHERE Bolt_Material = '" +
                   self.blt2.GetValue() + "'")
            ValueList[3] = Dbase().Dsqldata(qry)[0][0]
            qry = ("SELECT NutID FROM NutMaterial WHERE Nut_Material = '" +
                   self.nut2.GetValue() + "'")
            ValueList[4] = Dbase().Dsqldata(qry)[0][0]

        if SQL_step == 0:  # enter new record
            CurrentID = Max_ID
            num_vals = ('?,'*len(colinfo))[:-1]
            UpQuery = ('INSERT INTO ' + self.tblname + ' VALUES ('
                       + num_vals + ')')
            Dbase().TblEdit(UpQuery, ValueList)
            self.rec_num = len(self.data)

        elif SQL_step == 1:  # update edited record
            CurrentID = self.data[self.rec_num][0]
            realnames.remove('FastenerID')
            del ValueList[0]

            SQL_str = ','.join(["%s=?" % (name) for name in realnames])
            UpQuery = ('UPDATE ' + self.tblname + ' SET ' + SQL_str +
                       ' WHERE FastenerID = ' + str(CurrentID))
            Dbase().TblEdit(UpQuery, ValueList)
            self.data = Dbase().Dsqldata(self.MainSQL)

        elif SQL_step == 3:
            return

        self.b2.Disable()
        return CurrentID

    def OnRestoreCmbs(self, evt):
        self.RestoreCmbs()

    def PrintFile(self, evt):
        import Report_Lvl1

        data2 = []
        rptdata = []

        if self.data == []:
            NoData = wx.MessageDialog(
                None, 'No Data to Print', 'Error', wx.OK | wx.ICON_EXCLAMATION)
            NoData.ShowModal()
            return

        saveDialog = wx.FileDialog(self, message='Save Report as PDF.',
                                   wildcard='PDF (*.pdf)|*.pdf',
                                   style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT)

        if saveDialog.ShowModal() == wx.ID_CANCEL:
            filename = ''

        filename = saveDialog.GetPath()
        if filename.find(".pdf") == -1:
            filename = filename + '.pdf'

        saveDialog.Destroy()

        Colnames = ['Bolt\nMaterial', 'Nut\nMaterial']

        for segn in range(len(self.data)):
            data1 = [x for x in list(self.data[segn][1:]) if x is not None]
            rptID = (self.data[segn][0])

            for n in range(4):
                if n <= len(data1)-1:
                    if n in (0, 2):
                        qry = ('''SELECT Bolt_Material FROM BoltMaterial
                                WHERE BoltID = ''' + str(data1[n]))
                        data2.append(Dbase().Dsqldata(qry)[0][0])
                    else:
                        qry = ('''SELECT Nut_Material FROM NutMaterial
                                WHERE NutID = ''' + str(data1[n]))
                        data2.append(Dbase().Dsqldata(qry)[0][0])
                        rptdata.append(tuple(data2))
                        data2 = []

        Colwdths = [20, 20]

        ttl = None
        if len(self.data) == 1:
            # specify a title for the report if table name
            # is not to be used
            if self.PipeCode is None:
                ttl = self.text1.GetValue() + '-' + self.text2.GetValue()
            else:
                ttl = self.PipeCode
            ttl = ' for ' + ttl + '<br/>' + self.tblname + ' ID ' + str(rptID)

        Report_Lvl1.Report(self.tblname, rptdata, Colnames,
                           filename, ttl, Colwdths).create_pdf()

    def OnMovefst(self, evt):
        self.rec_num = 0
        self.FillScreen()

    def OnMovelst(self, evt):
        if len(self.data) == 0:
            return
        self.rec_num = len(self.data)-1
        self.FillScreen()

    def OnMovenxt(self, evt):
        if len(self.data) == 0:
            return
        self.rec_num += 1
        if self.rec_num == len(self.data):
            self.rec_num = 0
        self.FillScreen()

    def OnMovepre(self, evt):
        if len(self.data) == 0:
            return
        self.rec_num -= 1
        if self.rec_num < 0:
            self.rec_num = len(self.data)-1
        self.FillScreen()

    def FillScreen(self):
        # all the IDs for the various tables making up the package
        if len(self.data) == 0:
            self.recnum2.SetLabel(str(self.rec_num))
            return
        else:
            recrd = self.data[self.rec_num]
        qry = 'SELECT * FROM BoltMaterial WHERE BoltID = ' + str(recrd[1])
        txtval = Dbase().Dsqldata(qry)
        self.blt1.ChangeValue(txtval[0][1])
        qry = 'SELECT * FROM NutMaterial WHERE NutID = ' + str(recrd[2])
        txtval = Dbase().Dsqldata(qry)
        self.nut1.ChangeValue(txtval[0][1])

        if recrd[3] is not None:
            qry = 'SELECT * FROM BoltMaterial WHERE BoltID = ' + str(recrd[3])
            txtval = Dbase().Dsqldata(qry)
            self.blt2.ChangeValue(txtval[0][1])
            qry = 'SELECT * FROM NutMaterial WHERE NutID = ' + str(recrd[4])
            txtval = Dbase().Dsqldata(qry)
            self.nut2.ChangeValue(txtval[0][1])
        else:
            self.blt2.ChangeValue('')
            self.nut2.ChangeValue('')

        self.recnum2.SetLabel(str(self.rec_num + 1))

    def ValData(self):
        if self.blt1.GetValue() == '' or self.nut1.GetValue() == '':
            wx.MessageBox('''Value needed for;\nMaterial Option 1 to complete
                           information.''', 'Missing Data',
                          wx.OK | wx.ICON_INFORMATION)
            return False
        elif self.blt2.GetValue() != '' and self.nut2.GetValue() == '':
            wx.MessageBox('''Value needed for;\nMaterial Option 2 Nut Material
                           to complete information.''', 'Missing Data',
                          wx.OK | wx.ICON_INFORMATION)
            return False
        elif self.blt2.GetValue() == '' and self.nut2.GetValue() != '':
            wx.MessageBox('''Value needed for;\nMaterial Option 2 Bolt Material
                           to complete information.''', 'Missing Data',
                          wx.OK | wx.ICON_INFORMATION)
            return False
        else:
            self.b2.Disable()
            return True

    def OnAddComd(self, evt):
        self.AddComd()

    # link this ID to the commodity property
    def AddComd(self):

        if self.b6.GetLabel() == 'Show All\nItems':
            if self.MainSQL == '':
                self.MainSQL = 'SELECT * FROM ' + self.tblname
                self.data = Dbase().Dsqldata(self.MainSQL)
            else:
                self.MainSQL = self.MainSQL[:self.MainSQL.find('WHERE')]
                self.data = Dbase().Dsqldata(self.MainSQL)
            self.b6.SetLabel("Add Item\nto Commodity")
            self.b6.Enable()
            self.b3.Disable()
            self.FillScreen()
            self.recnum3.SetLabel('/ '+str(len(self.data)))
        else:
            query = ('''SELECT Fastener_ID FROM PipeSpecification
                      WHERE Commodity_Property_ID = ''' +
                     str(self.ComdPrtyID))
            StartQry = Dbase().Dsqldata(query)
            if StartQry == []:
                ValueList = []
                New_ID = cursr.execute(
                    '''SELECT MAX(Pipe_Spec_ID) FROM
                     PipeSpecification''').fetchone()
                if New_ID[0] is None:
                    Max_ID = '1'
                else:
                    Max_ID = str(New_ID[0]+1)
                colinfo = Dbase().Dcolinfo('PipeSpecification')
                for n in range(0, len(colinfo)-2):
                    ValueList.append(None)

                num_vals = ('?,'*len(colinfo))[:-1]
                ValueList.insert(0, Max_ID)
                ValueList.insert(1, str(self.ComdPrtyID))

                UpQuery = ("INSERT INTO PipeSpecification VALUES ("
                           + num_vals + ")")
                Dbase().TblEdit(UpQuery, ValueList)
                StartQry = Max_ID
            else:
                StartQry = str(StartQry[0][0])

            cmd_addID = self.data[self.rec_num][0]
            self.ChgSpecID(cmd_addID)

            self.MainSQL = ('SELECT * FROM ' + self.tblname +
                            ' WHERE FastenerID = ' + str(cmd_addID))
            self.data = Dbase().Dsqldata(self.MainSQL)

            self.rec_num = 0
            self.FillScreen()
            self.lbl_nodata.SetLabel('   ')
            self.recnum3.SetLabel('/ '+str(len(self.data)))
            self.b6.Disable()
            self.b3.Enable()
            self.b6.SetLabel("Show All\nItems")

    def ChgSpecID(self, ID=None):
        UpQuery = ('''UPDATE PipeSpecification SET Fastener_ID = ?
                    WHERE Commodity_Property_ID = ''' + str(self.ComdPrtyID))
        Dbase().TblEdit(UpQuery, [ID])

    def EditTbl(self):
        # if the table index is auto increment then assign next
        # value otherwise do nothing
        for item in Dbase().Dcolinfo(self.tblname):
            if item[5] == 1:
                IDname = item[1]
                if 'INTEGER' in item[2]:
                    New_ID = cursr.execute(
                        "SELECT MAX(" + IDname + ") FROM " +
                        self.tblname).fetchone()
                    if New_ID[0] is None:
                        Max_ID = '1'
                    else:
                        Max_ID = str(New_ID[0]+1)

        if self.blt1.GetValue() != '' and self.nut1.GetValue() != '':
            IDQuery = ('''SELECT ID FROM BoltMaterial WHERE
                        Bolt_Material = "''' + self.blt1.GetValue() + '"')
            ID1 = Dbase().Dsqldata(IDQuery)[0][0]
        if self.blt2.GetValue() != '' and self.nut2.GetValue() != '':
            IDQuery = ('SELECT ID FROM BoltMaterial WHERE Bolt_Material = "'
                       + self.blt2.GetValue() + '"')
            ID2 = Dbase().Dsqldata(IDQuery)[0][0]

        ValueList = []
        ValueList.append(Max_ID)
        ValueList.append(ID1)
        ValueList.append(ID2)

        Query = ('SELECT * FROM ' + self.tblname +
                 ' WHERE Bolting_Material_1 = "' + str(ID1) +
                 '" AND Bolting_Material_2 = "' + str(ID2) + '"')
        if Dbase().Dsqldata(Query):
            return
        else:
            UpQuery = ('INSERT INTO ' + self.tblname + ' VALUES (' + "'" +
                       "','".join(map(str, ValueList)) + "'" + ')')
            Dbase().TblEdit(UpQuery)

    def OnAddSpec(self, evt):
        # before saving racord check that all the
        # required data has been entered
        if self.blt1.GetValue() != '' and self.nut1.GetValue() != '':
            self.EditTbl()
        elif self.blt2.GetValue() != '' and self.nut2.GetValue() != '':
            self.EditTbl()
        else:
            wx.MessageBox('Value(s) needed for at least one Option',
                          'Missing Data', wx.OK | wx.ICON_INFORMATION)

    def OnCmbBltOpen(self, evt):
        self.cmbOld = evt.GetEventObject().GetValue()

    def OnCmbBltClose(self, evt):
        self.cmbNew = evt.GetEventObject().GetValue()
        if self.cmbNew != self.cmbOld:
            self.b2.Enable()

    def OnCmbNutOpen(self, evt):
        self.cmbOld = evt.GetEventObject().GetValue()

    def OnCmbNutClose(self,     evt):
        self.cmbNew = evt.GetEventObject().GetValue()
        if self.cmbNew != self.cmbOld:
            self.b2.Enable()

    def RestoreCmbs(self):
        self.blt1.ChangeValue('')
        self.nut1.ChangeValue('')
        self.blt2.ChangeValue('')
        self.nut2.ChangeValue('')

    def OnAddBlt(self, evt):
        boxnums = (0, 1)
        cmbtbl = 'BoltMaterial'
        CmbLst1(self, cmbtbl)

        self.ReFillList(cmbtbl, boxnums)

    def OnAddNut(self, evt):
        boxnums = (0, 1)
        cmbtbl = 'NutMaterial'
        CmbLst1(self, cmbtbl)

        self.ReFillList(cmbtbl, boxnums)

    def ReFillList(self, cmbtbl, boxnums):
        for n in boxnums:
            self.lc = self.lctrls[n]
            self.lc.DeleteAllItems()
            index = 0
            ReFillQuery = 'SELECT * FROM "' + cmbtbl + '"'
            for values in Dbase().Dsqldata(ReFillQuery):
                col = 0
                for value in values:
                    if col == 0:
                        self.lc.InsertItem(index, str(value))
                    else:
                        self.lc.SetItem(index, col, str(value))
                    col += 1
                index += 1

    def OnDelete(self, evt):
        if self.data != []:
            recrd = self.data[self.rec_num][0]
            if self.ComdPrtyID is None:
                try:
                    Dbase().TblDelete(self.tblname, recrd, 'FastenerID')
                    self.MainSQL = 'SELECT * FROM ' + self.tblname
                    self.data = Dbase().Dsqldata(self.MainSQL)
                    self.rec_num -= 1
                    if self.rec_num < 0:
                        self.rec_num = len(self.data)-1
                    self.FillScreen()
                    self.recnum3.SetLabel('/ '+str(len(self.data)))

                except sqlite3.IntegrityError:
                    wx.MessageBox('''This Record is associated
                                   with\nother tables and cannot
                                   be deleted!''', "Cannot Delete",
                                  wx.OK | wx.ICON_INFORMATION)
            else:
                self.ChgSpecID()
                self.RestoreCmbs()
                self.b3.Disable()
                self.data = []
                self.rec_num = 0
                self.recnum2.SetLabel(str(self.rec_num))
                self.recnum3.SetLabel('/ '+str(len(self.data)))
                self.lbl_nodata.SetLabel('The ' + self.frmtitle +
                                         ''' have not been setup for this
                                          Commodity Property''')

    def OnClose(self, evt):
        # add next 2 lines for child parent form
        self.GetParent().Enable(True)
        self.__eventLoop.Exit()
        self.Destroy()


class BldLvl3(wx.Frame):
    '''Routine to build form and populate grid'''
    def __init__(self, parent, tblname, ComdPrtyID=None):

        self.parent = parent   # add for child parent form

        self.ComCode = ''
        self.ComdPrtyID = ComdPrtyID

        if self.ComdPrtyID is not None:
            query = ('''SELECT Commodity_Code,Pipe_Material_Code,
                     End_Connection,Pipe_Code FROM CommodityProperties
                      WHERE CommodityPropertyID = '''
                     + str(self.ComdPrtyID))
            dataset = Dbase().Dsqldata(query)[0]
            self.PipeMtrSpec = dataset[1]
            self.ComCode = dataset[0]
            self.PipeCode = dataset[3]

        self.columnames = []
        self.rec_num = 0
        self.lctrls = []
        self.data = []
        self.MainSQL = ''
        self.tblname = tblname
        StartQry = None

        txtboxSize3 = txtboxSize4 = (350, 33)
        spcs = [12, 60, 60, 23]

        self.frgtbl2 = None

        if self.tblname == 'InspectionPacks':
            frmSize = (650, 620)
            txtboxSize3 = txtboxSize4 = (600, 130)
            self.IDcol = 'Inspection'
            self.cmb1hint = 'Fluid Category'
            self.frgtbl1 = 'FluidCategory'
            self.cmb1fld = 'Fluid_ID'
            self.cmb1val = 'Designation'
            self.lbl_txt = ['Fluid Category', 'Description']
            self.RptColWdth = [6, 20, 30, 30]
            self.RptColNames = ['ID', 'Fluid Category',
                                'Enhanced Inspection', 'Notes']
            spcs = [12, 120, 60, 23]
        elif self.tblname == 'GasketPacks':
            frmSize = (600, 500)
            txtboxSize3 = txtboxSize4 = (350, 66)
            self.IDcol = 'Gasket'
            self.cmb1hint = 'ANSI Rating'
            self.frgtbl1 = 'ANSI_Rating'
            self.cmb1fld = 'Rating_Designation'
            self.cmb1val = 'ANSI_Class'
            self.lbl_txt = ['ANSI Class', 'Description']
            self.RptColWdth = [6, 10, 30, 30]
            self.RptColNames = ['ID', 'ANSI Class', 'Description', 'Notes']
        elif self.tblname == 'PaintSpec':
            frmSize = (575, 550)
            self.IDcol = 'PaintSpec'
            self.cmb1hint = 'Surface Prep'
            self.frgtbl1 = 'PaintPrep'
            self.cmb1fld = 'PaintPrepID'
            self.cmb1val = 'Sandblast_Spec'
            self.cmb2hint = 'Color'
            self.frgtbl2 = 'PaintColors'
            self.cmb2fld = 'ColorCodeID'
            self.cmb2val = 'Color'
            self.lbl_txt = ['Surface Preparation', 'Paint Color',
                            'Base Coat', 'Final Coat']
            spcs = [10, 24, 25, 25]
            self.RptColWdth = [6, 15, 15, 15, 20, 15, 30]
            self.RptColNames = ['ID', 'Surface Prep', 'Base Coat',
                                'Final Coat', 'Color', 'Tagging', 'Notes']

        if self.tblname.find("_") != -1:
            self.frmtitle = (self.tblname.replace("_", " "))
        else:
            self.frmtitle = (' '.join(re.findall('([A-Z][a-z]*)',
                             self.tblname)))

        super(BldLvl3, self).__init__(parent,
                                      title=self.frmtitle,
                                      size=frmSize)

        self.Bind(wx.EVT_CLOSE, self.OnClose)

        font2 = wx.Font(11, wx.DEFAULT, wx.NORMAL, wx.FONTWEIGHT_BOLD)

        txt_nodata = (' The ' + self.frmtitle +
                      ' have not been setup for this Commodity Property')
        self.lbl_nodata = wx.StaticText(self, -1, label=txt_nodata,
                                        size=(650, 40), style=wx.ALIGN_CENTRE)
        self.lbl_nodata.SetForegroundColour((255, 0, 0))
        self.lbl_nodata.SetFont(font2)
        self.lbl_nodata.SetLabel('  ')

        if self.ComdPrtyID is not None:
            query = ('SELECT ' + self.IDcol +
                     '''_ID FROM PipeSpecification WHERE
                      Commodity_Property_ID = '''
                     + str(self.ComdPrtyID))
            chk = Dbase().Dsqldata(query)
            if chk != []:
                StartQry = chk[0][0]
                if StartQry is not None:
                    self.MainSQL = ('SELECT * FROM ' + self.tblname + ' WHERE '
                                    + self.IDcol + 'ID = ' + str(StartQry))
                    self.data = Dbase().Dsqldata(self.MainSQL)
                else:
                    self.lbl_nodata.SetLabel(txt_nodata)
            else:
                self.lbl_nodata.SetLabel(txt_nodata)
        else:
            self.MainSQL = 'SELECT * FROM ' + self.tblname
            self.data = Dbase().Dsqldata(self.MainSQL)

        self.warningsizer = wx.BoxSizer(wx.HORIZONTAL)
        self.warningsizer.Add(self.lbl_nodata, 0, wx.ALIGN_CENTER)

        # specify which listbox column to display in the combobox
        self.showcol = int

        # set the Sizer property (same as SetSizer)
        self.Sizer = wx.BoxSizer(wx.VERTICAL)

        self.specsizer = wx.BoxSizer(wx.HORIZONTAL)
        font = wx.Font(10, wx.DEFAULT, wx.NORMAL, wx.NORMAL)

        if StartQry is None:
            self.lbl_nodata.Show()

        if self.ComdPrtyID is not None:
            qry = ('''SELECT Pipe_Material_Spec FROM PipeMaterialSpec WHERE
                    Material_Spec_ID = ''' + str(self.PipeMtrSpec))
            mtrspec = Dbase().Dsqldata(qry)[0][0]
            if self.PipeCode is None:
                txt = self.ComCode + ' - ' + mtrspec
            else:
                txt = self.PipeCode
            self.text1 = wx.TextCtrl(self, size=(100, 33), value=txt,
                                     style=wx.TE_READONLY | wx.TE_CENTER)
            self.text1.SetForegroundColour((255, 0, 0))
            self.text1.SetFont(font)
            self.specsizer.Add(self.text1, 0, wx.ALL, 10)

        self.notesizer = wx.BoxSizer(wx.HORIZONTAL)
        self.note1 = wx.StaticText(self, label=self.lbl_txt[0],
                                   style=wx.ALIGN_LEFT)
        self.note1.SetForegroundColour((255, 0, 0))
        self.notesizer.Add(self.note1, 0, wx.LEFT, border=5)
        self.txthints = ['Description', 'Note', 'Vendor', 'Model']
        if self.tblname == 'PaintSpec':
            self.txthints = ['Base Coat', 'Final Coat', 'Tagging', 'Notes']
            self.note2 = wx.StaticText(self, label='Commodity Color',
                                       style=wx.ALIGN_LEFT)
            self.note2.SetForegroundColour((255, 0, 0))
            self.notesizer.Add(self.note2, 0, wx.LEFT, border=150)
            self.notesizer.Add((30, 10))

        self.cmbsizer = wx.BoxSizer(wx.HORIZONTAL)

        # develope the comboctrls and attach popup list
        self.cmb1 = wx.ComboCtrl(self, pos=(10, 10), size=(200, -1))
        self.Bind(wx.EVT_TEXT, self.OnSelect1, self.cmb1)
        # self.cmb1.SetHint(self.cmb1hint)
        self.cmb1.SetPopupControl(ListCtrlComboPopup(self.frgtbl1, showcol=1,
                                                     lctrls=self.lctrls))
        font1 = wx.Font(12, wx.DEFAULT, wx.NORMAL, wx.FONTWEIGHT_BOLD)
        self.addcmb1 = wx.Button(self, label='+', size=(35, -1))
        self.addcmb1.SetForegroundColour((255, 0, 0))
        self.addcmb1.SetFont(font1)
        self.Bind(wx.EVT_BUTTON, self.OnAddcmb1, self.addcmb1)
        self.cmbsizer.Add(self.cmb1, 0, wx.ALIGN_LEFT, 5)
        self.cmbsizer.Add(self.addcmb1, 0, wx.ALIGN_LEFT, 5)

        self.notesizer2 = wx.BoxSizer(wx.HORIZONTAL)
        if self.tblname == 'PaintSpec':
            self.cmb2 = wx.ComboCtrl(self, pos=(10, 10), size=(200, -1))
            self.Bind(wx.EVT_TEXT, self.OnSelect3, self.cmb2)
            self.cmb2.SetHint('Color')
            self.cmb2.SetPopupControl(ListCtrlComboPopup(
                'PaintColors', showcol=1, lctrls=self.lctrls))
            self.addcmb2 = wx.Button(self, label='+', size=(35, -1))
            self.addcmb2.SetForegroundColour((255, 0, 0))
            self.addcmb2.SetFont(font1)
            self.Bind(wx.EVT_BUTTON, self.OnAddcmb2, self.addcmb2)
            self.cmbsizer.Add((25, 10))
            self.cmbsizer.Add(self.cmb2, 0, wx.ALIGN_LEFT, 5)
            self.cmbsizer.Add(self.addcmb2, 0, wx.ALIGN_LEFT, 5)
            self.note3 = wx.TextCtrl(self, size=(300, 33), value='',
                                     style=wx.TE_READONLY | wx.TE_LEFT)
            self.note3.SetHint('Select a Surface Preparation')
            self.notesizer2.Add(self.note3, 0, wx.ALL | wx.RIGHT, 5)
            self.notesizer2.Add((100, 10))

        self.textsizer = wx.BoxSizer(wx.HORIZONTAL)

        self.txtsizer = wx.BoxSizer(wx.VERTICAL)
        self.text3 = wx.TextCtrl(self, size=txtboxSize3, value='',
                                 style=wx.TE_LEFT | wx.TE_WORDWRAP
                                 | wx.TE_MULTILINE)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text3)
        self.notesizer3 = wx.BoxSizer(wx.VERTICAL)
        self.text4 = wx.TextCtrl(self, size=txtboxSize4, value='',
                                 style=wx.TE_LEFT | wx.TE_WORDWRAP
                                 | wx.TE_MULTILINE)
        self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text4)
        self.txtsizer.Add(self.text3, 0, wx.ALL | wx.ALIGN_CENTER, 5)
        self.txtsizer.Add(self.text4, 0, wx.ALL | wx.ALIGN_CENTER, 5)

        self.note4 = wx.StaticText(self, label=self.txthints[0],
                                   style=wx.ALIGN_LEFT)
        self.note4.SetForegroundColour((255, 0, 0))
        self.note5 = wx.StaticText(self, label=self.txthints[1],
                                   style=wx.ALIGN_LEFT)
        self.note5.SetForegroundColour((255, 0, 0))
        self.notesizer3.Add(self.note4, 0, wx.ALIGN_RIGHT | wx.TOP, spcs[0])
        self.notesizer3.Add(self.note5, 0, wx.ALIGN_RIGHT | wx.TOP, spcs[1])

        if self.tblname == 'PaintSpec':
            self.text5 = wx.TextCtrl(self, size=(350, 33), value='',
                                     style=wx.TE_LEFT | wx.TE_WORDWRAP |
                                     wx.TE_MULTILINE)
            self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text5)
            self.text6 = wx.TextCtrl(self, size=(350, 33), value='',
                                     style=wx.TE_LEFT | wx.TE_WORDWRAP
                                     | wx.TE_MULTILINE)
            self.Bind(wx.EVT_TEXT, self.OnSelect2, self.text6)
            self.note6 = wx.StaticText(self, label=self.txthints[2],
                                       style=wx.ALIGN_LEFT)
            self.note6.SetForegroundColour((255, 0, 0))
            self.note7 = wx.StaticText(self, label=self.txthints[3],
                                       style=wx.ALIGN_LEFT)
            self.note7.SetForegroundColour((255, 0, 0))
            self.txtsizer.Add(self.text5, 0, wx.ALL | wx.ALIGN_CENTER, 5)
            self.txtsizer.Add(self.text6, 0, wx.ALL | wx.ALIGN_CENTER, 5)
            self.notesizer3.Add(self.note6, 0, wx.ALIGN_RIGHT |
                                wx.TOP, spcs[2])
            self.notesizer3.Add(self.note7, 0, wx.ALIGN_RIGHT |
                                wx.TOP, spcs[3])

        self.textsizer.Add(self.notesizer3, 0, wx.LEFT, 10)
        self.textsizer.Add(self.txtsizer, 0, wx.RIGHT, 10)

        self.Sizer.Add(self.warningsizer, 0, wx.CENTER | wx.TOP, 5)
        self.Sizer.Add(self.specsizer, 0, wx.ALL | wx.ALIGN_CENTER)
        self.Sizer.Add(self.notesizer, 0, wx.ALL | wx.ALIGN_CENTER)
        self.Sizer.Add(self.cmbsizer, 0, wx.ALL | wx.ALIGN_CENTER)
        self.Sizer.Add(self.notesizer2, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL)
        self.Sizer.Add(self.textsizer, 0, wx.ALIGN_CENTER | wx.TOP, 5)

        self.btnbox = wx.BoxSizer(wx.HORIZONTAL)
        # Add buttons for grid modifications

        self.b5 = wx.Button(self, label="Clear\nBoxes")
        self.Bind(wx.EVT_BUTTON, self.OnRestoreCmbs, self.b5)
        self.btnbox.Add(self.b5, 0, wx.ALIGN_CENTER | wx.RIGHT, 20)

        self.b3 = wx.Button(self, label="Delete\nSpec")
        self.Bind(wx.EVT_BUTTON, self.OnDelete, self.b3)

        self.b1 = wx.Button(self, label="Print\nReport")
        self.Bind(wx.EVT_BUTTON, self.PrintFile, self.b1)
        self.btnbox.Add(self.b1, 0, wx.ALL | wx.ALIGN_LEFT, 5)

        # add a button box and place the buttons
        self.b2 = wx.Button(self, label="Add/Update\n" + self.frmtitle)
        self.b2.Disable()
        self.Bind(wx.EVT_BUTTON, self.OnAddRec, self.b2)
        self.btnbox.Add(self.b2, 0, wx.ALL | wx.ALIGN_LEFT, 5)

        if self.ComdPrtyID is not None:
            self.b6 = wx.Button(self, size=(120, 45), label="Show All\nItems")
            self.b3.SetLabel('Delete From\nCommodity')
            if StartQry is None:
                self.b3.Disable()
            self.Bind(wx.EVT_BUTTON, self.OnAddComd, self.b6)
            self.btnbox.Add(self.b6, 0, wx.ALL | wx.ALIGN_CENTER, 5)
        self.btnbox.Add(self.b3, 0, wx.ALL | wx.ALIGN_LEFT, 5)

        self.navbox = wx.BoxSizer(wx.HORIZONTAL)
        self.fst = wx.Button(self, label='<<')
        self.lst = wx.Button(self, label='>>')
        self.nxt = wx.Button(self, label='>')
        self.pre = wx.Button(self, label='<')
        self.fst.Bind(wx.EVT_BUTTON, self.OnMovefst)
        self.lst.Bind(wx.EVT_BUTTON, self.OnMovelst)
        self.nxt.Bind(wx.EVT_BUTTON, self.OnMovenxt)
        self.pre.Bind(wx.EVT_BUTTON, self.OnMovepre)

        self.b4 = wx.Button(self, label="Exit")
        self.Bind(wx.EVT_BUTTON, self.OnClose, self.b4)

        self.navbox.Add(self.fst, 0, wx.ALL | wx.ALIGN_LEFT, 5)
        self.navbox.Add(self.pre, 0, wx.ALL | wx.ALIGN_LEFT, 5)
        self.navbox.Add(self.nxt, 0, wx.ALL | wx.ALIGN_LEFT, 5)
        self.navbox.Add(self.lst, 0, wx.ALL | wx.ALIGN_LEFT, 5)
        self.navbox.Add((20, 10))
        self.navbox.Add(self.b4, 0, wx.ALL | wx.ALIGN_LEFT, 5)
        self.numbox = wx.BoxSizer(wx.HORIZONTAL)
        self.recnum1 = wx.StaticText(self, label='Record ',
                                     style=wx.ALIGN_LEFT)
        self.recnum1.SetForegroundColour((255, 0, 0))

        self.recnum2 = wx.StaticText(self, label=str(self.rec_num+1),
                                     style=wx.ALIGN_LEFT)
        self.recnum2.SetForegroundColour((255, 0, 0))
        self.recnum3 = wx.StaticText(self, label='/ ' + str(len(self.data)),
                                     style=wx.ALIGN_LEFT)
        self.recnum3.SetForegroundColour((255, 0, 0))
        self.numbox.Add(self.recnum1, 0, wx.ALL | wx.ALIGN_LEFT, 5)
        self.numbox.Add(self.recnum2, 0, wx.ALL | wx.ALIGN_LEFT, 5)
        self.numbox.Add(self.recnum3, 0, wx.ALL | wx.ALIGN_LEFT, 5)

        self.Sizer.Add(self.btnbox, 0, wx.ALIGN_CENTER | wx.TOP, 10)
        self.Sizer.Add(self.navbox, 0, wx.ALIGN_CENTER | wx.TOP, 10)
        self.Sizer.Add(self.numbox, 0, wx.ALIGN_CENTER | wx.TOP)
        self.b4.SetFocus()

        self.FillScreen()

        # add these 5 following lines to child parent form
        self.CenterOnParent()
        self.GetParent().Enable(False)
        self.Show(True)
        self.__eventLoop = wx.GUIEventLoop()
        self.__eventLoop.Run()

    def AddRec(self, SQL_step):
        realnames = []
        ValueList = []

        # if the table index is auto increment then assign
        # next value otherwise do nothing
        for item in Dbase().Dcolinfo(self.tblname):
            if item[5] == 1:
                IDname = item[1]
                if 'INTEGER' in item[2]:
                    New_ID = cursr.execute(
                        "SELECT MAX(" + IDname + ") FROM " +
                        self.tblname).fetchone()
                    if New_ID[0] is None:
                        Max_ID = '1'
                    else:
                        Max_ID = str(New_ID[0]+1)
            realnames.append(item[1])
        ValueList.append(Max_ID)
        if self.cmb1.GetValue() != '':   # and self.note3.GetValue()!='':
            IDQuery = ('SELECT ' + self.cmb1fld + ' FROM ' + self.frgtbl1 +
                       ' WHERE ' + self.cmb1val + ' LIKE "%' +
                       self.cmb1.GetValue() + '%" COLLATE NOCASE')
            ID1 = Dbase().Dsqldata(IDQuery)[0][0]

        if self.tblname != 'PaintSpec':
            ValueList.append(ID1)
            ValueList.append(self.text3.GetValue())
            ValueList.append(self.text4.GetValue())

        if self.tblname == 'PaintSpec':
            if self.cmb2.GetValue() != '':
                IDQuery = ('SELECT ' + self.cmb2fld + ' FROM ' + self.frgtbl2 +
                           ' WHERE ' + self.cmb2val + ' LIKE "%' +
                           self.cmb2.GetValue() + '%" COLLATE NOCASE')
                ID2 = Dbase().Dsqldata(IDQuery)[0][0]
            ValueList.append(ID1)
            ValueList.append(self.text3.GetValue())
            ValueList.append(self.text4.GetValue())
            ValueList.append(ID2)
            ValueList.append(self.text5.GetValue())
            ValueList.append(self.text6.GetValue())

        if SQL_step == 0:  # enter new record
            CurrentID = Max_ID
            UpQuery = ('INSERT INTO ' + self.tblname +
                       ' VALUES (' + "'" + "','".
                       join(map(str, ValueList)) + "'" + ')')
            Dbase().TblEdit(UpQuery)

        elif SQL_step == 1:  # update edited record
            CurrentID = self.data[self.rec_num][0]
            realnames.remove(self.IDcol+'ID')
            del ValueList[0]

            SQL_str = dict(zip(realnames, ValueList))
            Update_str = (", ".join(["%s='%s'" % (k, v)
                          for k, v in SQL_str.items()]))
            UpQuery = ('UPDATE ' + self.tblname + ' SET ' + Update_str +
                       ' WHERE ' + self.IDcol + 'ID = ' + str(CurrentID))
            Dbase().TblEdit(UpQuery)

        elif SQL_step == 3:
            return
        self.b2.Disable()
        return CurrentID

    def OnMovefst(self, evt):
        self.rec_num = 0
        self.FillScreen()

    def OnMovelst(self, evt):
        self.rec_num = len(self.data)-1
        if self.rec_num < 0:
            self.rec_num = 0
        self.FillScreen()

    def OnMovenxt(self, evt):
        if len(self.data) != 0:
            self.rec_num += 1
        if self.rec_num == len(self.data):
            self.rec_num = 0
        self.FillScreen()

    def OnMovepre(self, evt):
        if len(self.data) != 0:
            self.rec_num -= 1
        if self.rec_num < 0:
            self.rec_num = len(self.data)-1
        self.FillScreen()

    def OnRestoreCmbs(self, evt):
        self.RestoreCmbs()

    def RestoreCmbs(self):
        self.cmb1.Clear()
        self.cmb1.ChangeValue('')
        self.text3.Clear()
        self.text3.ChangeValue('')
        self.text4.Clear()
        self.text4.ChangeValue('')
        if self.tblname == 'PaintSpec':
            self.cmb2.Clear()
            self.cmb2.ChangeValue('')
            self.note3.Clear()
            self.note3.ChangeValue('')
            self.text5.Clear()
            self.text5.ChangeValue('')
            self.text6.Clear()
            self.text6.ChangeValue('')

    def FillScreen(self):
        # all the IDs for the various tables making up the package
        if len(self.data) == 0:
            self.recnum2.SetLabel(str(self.rec_num))
            return
        else:
            recrd = self.data[self.rec_num]

        if type(recrd[1]) == str:
            query = ('SELECT * FROM ' + self.frgtbl1 +
                     ' WHERE ' + self.cmb1fld + ' = "' + str(recrd[1]) + '"')
        else:
            query = ('SELECT * FROM ' + self.frgtbl1 +
                     ' WHERE ' + self.cmb1fld + ' = ' + str(recrd[1]))
        val = Dbase().Dsqldata(query)
        if val == []:
            self.cmb1.ChangeValue('')
            if self.tblname == 'PaintSpec':
                self.note3.ChangeValue('')
        else:
            self.cmb1.ChangeValue(str(val[0][1]))
            if self.tblname == 'PaintSpec':
                self.note3.ChangeValue(val[0][2])

        self.text3.ChangeValue(recrd[2])
        self.text4.ChangeValue(str(recrd[3]))

        if self.tblname == 'PaintSpec':
            query = ('SELECT * FROM paintColors WHERE ColorCodeID = "'
                     + str(recrd[4]) + '"')
            val = Dbase().Dsqldata(query)
            if val == []:
                self.cmb2.ChangeValue('')
            else:
                self.cmb2.ChangeValue(val[0][1])

            self.text5.ChangeValue(recrd[5])
            self.text6.ChangeValue(recrd[6])

        self.recnum2.SetLabel(str(self.rec_num+1))

    def OnAddcmb1(self, evt):
        CmbLst1(self, self.frgtbl1)
        self.ReFillList(0, self.frgtbl1)

    def OnAddcmb2(self, evt):
        CmbLst1(self, self.frgtbl2)
        self.ReFillList(1, self.frgtbl2)

    def ReFillList(self, combo, tbl):
        self.lc = self.lctrls[combo]
        self.lc.DeleteAllItems()

        index = 0
        ReFillQuery = 'SELECT * FROM "' + tbl + '"'
        for values in Dbase().Dsqldata(ReFillQuery):
            col = 0
            for value in values:
                if col == 0:
                    self.lc.InsertItem(index, str(value))
                else:
                    self.lc.SetItem(index, col, str(value))
                col += 1
            index += 1

    # Return the widget that is to be used for the popup
    def GetControl(self):
        return self.lc

    # Return a string representation of the current item.
    def GetStringValue(self):
        if self.value == -1:
            return
        return self.lc.GetItemText(self.value, self.showcol)

    def OnAddRec(self, evt):
        # check first that data is all present and clean up incomplete boxes
        check = self.ValData()

        if check:
            SQL_step = 3

            choice1 = ('1) Save this as a new ' + self.frmtitle +
                       ' Specification')
            choice2 = ('2) Update the existing ' + self.frmtitle +
                       ' Specification with this data')
            txt1 = ('''NOTE: Updating this information will be\n
                    reflected in all associated ''' + self.frmtitle)
            txt2 = (''' Specifications!\nRecommendation is to save as a new
                     specification.\n\n\tHow do you want to proceed?''')

            # if this is a not commodity related change
            if self.ComdPrtyID is None:
                # Make a selection as to whether the record
                # is to be a new or an update valve
                # use a SingleChioce dialog to determine
                # if data is new record or edited record
                SQL_Dialog = wx.SingleChoiceDialog(
                    self, txt1+txt2, 'Information Has Changed',
                    [choice1, choice2], style=wx.CHOICEDLG_STYLE)
                if SQL_Dialog.ShowModal() == wx.ID_OK:
                    SQL_step = SQL_Dialog.GetSelection()
                SQL_Dialog.Destroy()

                self.AddRec(SQL_step)
                self.MainSQL = 'SELECT * FROM ' + self.tblname
                self.data = Dbase().Dsqldata(self.MainSQL)

            else:  # this is a commodity related change
                choice1 = choice1 + ' for this commodity?'
                choice2 = choice2 + ' for this commodity?'
                # use a SingleChioce dialog to determine if
                # data is new record or edited record
                SQL_Dialog = wx.SingleChoiceDialog(
                    self, txt1+txt2, 'Information Has Changed',
                    [choice1, choice2], style=wx.CHOICEDLG_STYLE)
                if SQL_Dialog.ShowModal() == wx.ID_OK:
                    SQL_step = SQL_Dialog.GetSelection()

                SQL_Dialog.Destroy()

                cmd_addID = self.AddRec(SQL_step)
                # always save as a new spec if commodity property is specified
                # no matter the change over write or add the specification ID
                # to the PipeSpec table under the commodity property ID

                self.ChgSpecID(cmd_addID)

                query = (
                    'SELECT ' + self.IDcol + '_ID' +
                    ' FROM PipeSpecification WHERE Commodity_Property_ID = '
                    + str(self.ComdPrtyID))
                StartQry = Dbase().Dsqldata(query)
                self.MainSQL = ('SELECT * FROM ' + self.tblname + ' WHERE '
                                + self.IDcol + 'ID = ' + str(StartQry[0][0]))
                self.data = Dbase().Dsqldata(self.MainSQL)

            if SQL_step == 0:
                self.rec_num = len(self.data)-1
            if SQL_step == 1 and self.ComdPrtyID is not None:
                self.rec_num = 0
            self.FillScreen()
            self.recnum3.SetLabel('/ '+str(len(self.data)))

    def ValData(self):
        digstr = []
        #  DataStrg = []
        NoData = 0
        DialogStr = ''

        NoData = 0
        if self.tblname == 'PaintSpec':
            if self.note3.GetValue() == '':
                NoData = 2**0
            if self.cmb2.GetValue() == '':
                NoData = 2**1
            if self.text4.GetValue() == '':
                NoData = 2**3

        if self.text3.GetValue() == '':
            NoData = 2**2

        numrange = []
        for num in range(1, len(self.lbl_txt)+1):
            numrange.append(num)
        DataBxs = dict(zip(numrange, self.lbl_txt))
        numCells = int(len(self.lbl_txt))
        if NoData > 0:
            # use the binary numbers as keys for the DataBxs dictionary
            binstr = '{0:' + str(numCells) + 'b}'
            binry = str(binstr.format(NoData))
            digstr = [pos for pos, char in enumerate(binry) if char == '1']
            for dig in digstr:
                DialogStr = DataBxs[len(self.lbl_txt)-dig] + ',\n' + DialogStr
            wx.MessageBox('Value needed for;\n' + DialogStr +
                          'to complete information.', 'Missing Data',
                          wx.OK | wx.ICON_INFORMATION)
            return False
        else:
            return True

    def OnSelect1(self, evt):
        # only needed to update the paint prep note field note3
        recrd = self.data[self.rec_num]
        query = ('SELECT ' + self.cmb1val + ' FROM ' + self.frgtbl1 +
                 ' WHERE ' + self.cmb1fld + ' = "' + str(recrd[1]) + '"')

        if str(self.cmb1.GetValue()) != str(Dbase().Dsqldata(query)[0][0]):
            self.b2.Enable()
            if self.tblname == 'PaintSpec':
                cmbvalue = str(self.cmb1.GetValue())
                if cmbvalue != "":
                    query = ('''SELECT Notes FROM PaintPrep WHERE
                             Sandblast_Spec = "''' + cmbvalue + '"')
                    note_3 = Dbase().Dsqldata(query)
                if note_3 == []:
                    self.note3.ChangeValue('')
                else:
                    self.note3.ChangeValue(note_3[0][0])

    def OnSelect2(self, evt):
        self.b2.Enable()

    def OnSelect3(self, evt):
        recrd = self.data[self.rec_num]
        query = ('SELECT ' + self.cmb2val + ' FROM ' + self.frgtbl2 + ' WHERE '
                 + self.cmb2fld + ' = "' + str(recrd[4]) + '"')
        if self.cmb2.GetValue() != Dbase().Dsqldata(query)[0][0]:
            self.b2.Enable()

    def OnAddComd(self, evt):
        self.AddComd()

    # link this ID to the commodity property
    def AddComd(self):
        if self.b6.GetLabel() == 'Show All\nItems':
            if self.MainSQL == '':
                self.MainSQL = 'SELECT * FROM ' + self.tblname
                self.data = Dbase().Dsqldata(self.MainSQL)
            else:
                self.MainSQL = self.MainSQL[:self.MainSQL.find('WHERE')]
                self.data = Dbase().Dsqldata(self.MainSQL)
            self.b6.SetLabel("Add Item\nto Commodity")
            self.b6.Enable()
            self.b3.Disable()
            self.FillScreen()
            self.recnum3.SetLabel('/ '+str(len(self.data)))
        else:
            query = ('SELECT ' + self.IDcol + '_ID' +
                     ' FROM PipeSpecification WHERE Commodity_Property_ID = '
                     + str(self.ComdPrtyID))
            StartQry = Dbase().Dsqldata(query)
            if StartQry == []:
                ValueList = []
                New_ID = cursr.execute(
                    '''SELECT MAX(Pipe_Spec_ID) FROM
                     PipeSpecification''').fetchone()
                if New_ID[0] is None:
                    Max_ID = '1'
                else:
                    Max_ID = str(New_ID[0]+1)
                colinfo = Dbase().Dcolinfo('PipeSpecification')
                for n in range(0, len(colinfo)-2):
                    ValueList.append(None)

                num_vals = ('?,'*len(colinfo))[:-1]
                ValueList.insert(0, Max_ID)
                ValueList.insert(1, str(self.ComdPrtyID))

                UpQuery = ("INSERT INTO PipeSpecification VALUES ("
                           + num_vals + ")")
                Dbase().TblEdit(UpQuery, ValueList)
                StartQry = Max_ID
            else:
                StartQry = str(StartQry[0][0])

            cmd_addID = self.data[self.rec_num][0]
            self.ChgSpecID(cmd_addID)

            self.MainSQL = ('SELECT * FROM ' + self.tblname +
                            ' WHERE ' + self.IDcol + 'ID = ' + str(cmd_addID))
            self.data = Dbase().Dsqldata(self.MainSQL)

            self.rec_num = 0
            self.FillScreen()
            self.lbl_nodata.SetLabel('   ')
            self.recnum3.SetLabel('/ '+str(len(self.data)))
            self.b6.Enable()
            self.b3.Enable()
            self.b6.SetLabel("Show All\nItems")

    def ChgSpecID(self, ID=None):
        UpQuery = ('UPDATE PipeSpecification SET ' + self.IDcol +
                   '_ID=?  WHERE Commodity_Property_ID = ' +
                   str(self.ComdPrtyID))
        Dbase().TblEdit(UpQuery, [ID])

    def OnDelete(self, evt):
        recrd = self.data[self.rec_num][0]
        if self.ComdPrtyID is None:
            try:
                Dbase().TblDelete(self.tblname, recrd, self.IDcol + 'ID')
                self.MainSQL = 'SELECT * FROM ' + self.tblname
                self.data = Dbase().Dsqldata(self.MainSQL)
                self.rec_num -= 1
                if self.rec_num < 0:
                    self.rec_num = len(self.data)-1
                self.FillScreen()
                self.recnum3.SetLabel('/ '+str(len(self.data)))
                self.lbl_nodata.SetLabel(
                    'The ' + self.frmtitle +
                    ' have not been setup for this Commodity Property')

            except sqlite3.IntegrityError:
                wx.MessageBox("This Record is associated"
                              " with\nother tables and cannot be deleted!",
                              "Cannot Delete", wx.OK | wx.ICON_INFORMATION)
        else:
            self.ChgSpecID()
            self.AddComd()
            self.b3.Disable()
            self.lbl_nodata.SetLabel(
                'The ' + self.frmtitle +
                ' have not been setup for this Commodity Property')

    def PrintFile(self, evt):
        import Report_Lvl2

        RptData = []

        if self.data == []:
            NoData = wx.MessageDialog(
                None, 'No Data to Print', 'Error', wx.OK | wx.ICON_EXCLAMATION)
            NoData.ShowModal()
            return

        saveDialog = wx.FileDialog(self, message='Save Report as PDF.',
                                   wildcard='PDF (*.pdf)|*.pdf',
                                   style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT)

        if saveDialog.ShowModal() == wx.ID_CANCEL:
            filename = ''

        filename = saveDialog.GetPath()
        if filename.find(".pdf") == -1:
            filename = filename + '.pdf'

        saveDialog.Destroy()
        newlst = []
        for i in self.data:
            newlst = list(i)
            frgID = i[1]

            if type(frgID) == str:
                qry = ('SELECT ' + self.cmb1val + ' FROM ' + self.frgtbl1 +
                       ' WHERE ' + self.cmb1fld + ' = "' + frgID + '"')
            else:
                qry = ('SELECT ' + self.cmb1val + ' FROM ' + self.frgtbl1 +
                       ' WHERE ' + self.cmb1fld + ' = ' + str(frgID))

            newtxt1 = Dbase().Dsqldata(qry)[0][0]
            newlst[1] = newtxt1
            if self.tblname == 'PaintSpec':
                qry = ('SELECT ' + self.cmb2val + ' FROM ' + self.frgtbl2 +
                       ' WHERE ' + self.cmb2fld + ' = ' + str(frgID))
                newtxt2 = Dbase().Dsqldata(qry)[0][0]
                newlst[4] = newtxt2
            newlst = tuple(newlst)
            RptData.append(newlst)

            if self.ComdPrtyID is not None:
                ttl = self.frmtitle + ' for ' + self.text1.GetValue()
            else:
                ttl = None

        Report_Lvl2.Report(self.tblname, RptData, self.RptColNames,
                           self.RptColWdth, filename, ttl).create_pdf()

    def OnClose(self, evt):
        # add following 2 lines for child parent
        self.GetParent().Enable(True)
        self.__eventLoop.Exit()
        self.Destroy()


class BrchFrm(wx.Frame):

    def __init__(self, parent, ComdPrtyID=None):
        self.parent = parent

        super(BrchFrm, self).__init__(
                          parent,
                          title='Branch Chart Specifications',
                          size=(790, 640))

        self.Bind(wx.EVT_CLOSE, self.OnClose)

        self.FrmSizer = wx.BoxSizer(wx.VERTICAL)

        self.pnl = BrchPnl(self, ComdPrtyID)
        self.FrmSizer.Add(self.pnl, 1, wx.EXPAND)
        self.FrmSizer.Add((10, 20))
        self.pnl.b5.Bind(wx.EVT_BUTTON, self.OnClose)
        self.SetSizer(self.FrmSizer)

        # add these 5 following lines to child parent form
        self.CenterOnParent()
        self.GetParent().Enable(False)
        self.Show(True)
        self.__eventLoop = wx.GUIEventLoop()
        self.__eventLoop.Run()

    def OnClose(self, evt):
        # add following 2 lines for child parent
        self.GetParent().Enable(True)
        self.__eventLoop.Exit()
        # Dbase().close_database()
        self.Destroy()


class BrchPnl(sc.ScrolledPanel):
    def __init__(self, parent, ComdPrtyID):
        super(BrchPnl, self).__init__(parent, size=(760, 550))

        self.tblname = 'BranchCriteria'
        self.ComdPrtyID = ComdPrtyID

        self.InitUI()

    def InitUI(self):
        font = wx.Font(10, wx.DEFAULT, wx.NORMAL, wx.NORMAL)
        font1 = wx.Font(12, wx.DEFAULT, wx.NORMAL, wx.BOLD)
        query = ('''SELECT Commodity_Code,Pipe_Material_Code, Pipe_Code,
                 End_Connection FROM CommodityProperties WHERE
                 CommodityPropertyID = ''' + str(self.ComdPrtyID))
        dataset = Dbase().Dsqldata(query)[0]
        PipeMtrSpec = dataset[1]
        self.ComCode = dataset[0]
        self.PipeCode = dataset[2]
        Ends = dataset[3]

        qry = ('''SELECT Pipe_Material_Spec FROM PipeMaterialSpec WHERE
                Material_Spec_ID = ''' + str(PipeMtrSpec))
        self.mtrcode = Dbase().Dsqldata(qry)[0][0]

        qry = 'SELECT Commodity_End FROM CommodityEnds WHERE ID = ' + str(Ends)
        self.End = Dbase().Dsqldata(qry)[0][0]

        query = ('''SELECT * FROM PipeSpecification WHERE
                 Commodity_Property_ID = ''' + str(self.ComdPrtyID))
        chk = Dbase().Dsqldata(query)
        if chk != []:    # commodity code exists in pipespec
            self.CurrentID = chk[0][18]  # get the current branch criteria ID
            if self.CurrentID is not None:  # ID of branch criteria exists
                # for commodity code
                self.MainSQL = ('''SELECT * FROM BranchCriteria WHERE
                                BranchID = ''' + str(self.CurrentID))
                self.data = Dbase().Dsqldata(self.MainSQL)
            else:   # branch ID does not exist for the commodity code
                self.MainSQL = ''
                self.data = []
            self.NewPipeSpec = False
        else:     # commodity code does not exist in pipespec
            # therefore neither does branch ID
            self.NewPipeSpec = True
            self.CurrentID = None
            self.MainSQL = ''
            self.data = []

        if self.CurrentID is None:
            New_ID = cursr.execute(
                "SELECT MAX(BranchID) FROM BranchCriteria").fetchone()
            if New_ID == []:    # the branch criteria table is empty
                self.CurrentID = '1'
            else:        # the maximum ID in the branch criteria table
                self.CurrentID = str(New_ID[0]+1)

        DirectionTxt = 'This form allows the setting of the general '\
            'guidelines for the generation of the branch connection charts.'\
            'Branch charts are based on the criteria of the type of end '\
            'connection and size differences between the branch and main run.'\
            '  In the default data it is assumed that when the branch equals '\
            ' the run an equal tee will be used. When the branch is one size '\
            'smaller than the run a reducing tee is to be used.'\
            ' The type of end connection allowed in the chart is determined'\
            ' by those specified for the correspondinig commodity property.'

        self.Sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.SizerLeft = wx.BoxSizer(wx.VERTICAL)

        self.text1 = wx.TextCtrl(self, size=(325, 220),
                                 value=DirectionTxt,
                                 style=wx.TE_READONLY | wx.TE_MULTILINE)
        self.text1.SetForegroundColour('red')
        self.text1.SetFont(font)
        self.SizerLeft.Add(self.text1, 0, wx.TOP | wx.LEFT, 15)

        self.leftszr1 = wx.BoxSizer(wx.HORIZONTAL)
        self.b1 = wx.Button(self, label="Save\nCriteria")
        self.Bind(wx.EVT_BUTTON, self.OnAddRec, self.b1)
        self.b3 = wx.Button(self, label='Delete\nCriteria')
        self.Bind(wx.EVT_BUTTON, self.OnDelete, self.b3)
        self.b4 = wx.Button(self, label='Print', size=(60, 45))
        self.Bind(wx.EVT_BUTTON, self.OnPrint, self.b4)
        self.leftszr1.Add((20, 10))
        self.leftszr1.Add(self.b1, 0, wx.ALIGN_TOP | wx.LEFT, 10)
        self.leftszr1.Add(self.b3, 0, wx.ALIGN_TOP | wx.LEFT, 10)
        self.leftszr1.Add(self.b4, 0, wx.ALIGN_TOP | wx.LEFT, 10)
        self.SizerLeft.Add((10, 15))
        self.SizerLeft.Add(self.leftszr1, 0, wx.TOP, 5)

        self.leftszr2 = wx.BoxSizer(wx.HORIZONTAL)
        if self.PipeCode is None:
            txtlbl = ('Branch Criteria for ' + self.ComCode +
                      ' - ' + self.mtrcode)
        else:
            txtlbl = 'Branch Criteria for ' + self.PipeCode
        if self.data == []:
            txtlbl = 'No Data for this commodity.'
            self.b3.Enable(False)
            self.b4.Enable(False)
        self.txtComd = wx.StaticText(self, -1,
                                     style=wx.ALIGN_CENTER_HORIZONTAL)
        self.txtComd.SetLabel(txtlbl)
        self.txtComd.SetForegroundColour('red')
        self.txtComd.SetFont(font1)
        self.leftszr2.Add(self.txtComd, 0, wx.ALIGN_CENTER | wx.LEFT, 15)
        self.SizerLeft.Add(self.leftszr2, 0, wx.ALL, 15)

        self.leftszr3 = wx.BoxSizer(wx.HORIZONTAL)
        lblSelect = wx.StaticText(
            self, label='1) Select the size limits for the chart',
            style=wx.ALIGN_RIGHT)
        lblSelect.SetForegroundColour('red')
        lblSelect.SetFont(font)
        ln0 = wx.StaticLine(self, 0, size=(350, 2), style=wx.LI_VERTICAL)
        ln0.SetBackgroundColour('Black')
        self.leftszr3.Add(lblSelect, 0, wx.LEFT, 40)
        self.SizerLeft.Add(ln0, 0, wx.ALIGN_LEFT | wx.BOTTOM, 5)
        self.SizerLeft.Add(self.leftszr3, 0, wx.TOP | wx.BOTTOM, 10)
        self.SizerLeft.Add((20, 5))

        self.leftszr4 = wx.BoxSizer(wx.HORIZONTAL)
        lblMinOD = wx.StaticText(
            self, label='Select the Minimum\nPipe OD for the Chart',
            style=wx.ALIGN_LEFT)
        self.cmbMin = wx.ComboCtrl(self, size=(80, -1), style=wx.CB_READONLY)
        self.cmbMin.SetHint('Min. Dia')
        self.cmbMin.SetPopupControl(ListCtrlComboPopup('Pipe_OD', showcol=1))
        self.leftszr4.Add(lblMinOD, 0, wx.LEFT, 30)
        self.leftszr4.Add(self.cmbMin, 0, wx.ALIGN_CENTER | wx.LEFT, 60)
        self.SizerLeft.Add(self.leftszr4, 0, wx.BOTTOM, 7)

        self.leftszr5 = wx.BoxSizer(wx.HORIZONTAL)
        lblMaxOD = wx.StaticText(
            self, label='Select the Maximum\nPipe OD for the Chart',
            style=wx.ALIGN_LEFT)
        self.cmbMax = wx.ComboCtrl(self, size=(80, -1), style=wx.CB_READONLY)
        self.cmbMax.SetHint('Max. Dia')
        self.cmbMax.SetPopupControl(ListCtrlComboPopup('Pipe_OD', showcol=1))
        self.leftszr5.Add(lblMaxOD, 0, wx.LEFT, 30)
        self.leftszr5.Add(self.cmbMax, 0, wx.ALIGN_CENTER | wx.LEFT, 60)
        self.SizerLeft.Add(self.leftszr5, 0, wx.BOTTOM, 5)

        self.leftszr6 = wx.BoxSizer(wx.HORIZONTAL)
        ln = wx.StaticLine(self, 0, size=(350, 2), style=wx.LI_VERTICAL)
        ln.SetBackgroundColour('Black')
        self.SizerLeft.Add(ln, 0, wx.TOP | wx.BOTTOM, 10)
        self.SizerLeft.Add(self.leftszr6, 0, wx.BOTTOM, 5)
        self.SizerLeft.Add((20, 5))

        self.leftszr7 = wx.BoxSizer(wx.HORIZONTAL)
        lblRdT = wx.StaticText(
            self, label='2) For Reducing Tees use the following limits:',
            style=wx.ALIGN_LEFT)
        lblRdT.SetForegroundColour('red')
        lblRdT.SetFont(font)
        self.leftszr7.Add(lblRdT, 0, wx.ALIGN_CENTER | wx.LEFT, 40)
        self.SizerLeft.Add(self.leftszr7, 0, wx.BOTTOM, 15)

        self.leftszr8 = wx.BoxSizer(wx.HORIZONTAL)
        lblRnSz = wx.StaticText(
            self, label='Run Size Must Be\nLess Than or Equal to:\n\nAND',
            style=wx.ALIGN_LEFT)
        self.cmbRTr = wx.ComboCtrl(self, size=(80, -1), style=wx.CB_READONLY)
        self.cmbRTr.SetHint('Max. OD')
        self.cmbRTr.SetPopupControl(ListCtrlComboPopup('Pipe_OD', showcol=1))
        self.leftszr8.Add(lblRnSz, 0, wx.LEFT, 30)
        self.leftszr8.Add(self.cmbRTr, 0, wx.ALIGN_TOP | wx.LEFT, 60)
        self.SizerLeft.Add(self.leftszr8, 0, wx.BOTTOM, 10)

        self.leftszr9 = wx.BoxSizer(wx.HORIZONTAL)
        lblRnSz1 = wx.StaticText(
            self, label='Branch Size Must\nBe At Least',
            style=wx.ALIGN_LEFT)
        self.txtRTDif = wx.TextCtrl(self, size=(30, 30), value='')
        self.txtRTDif.SetHint('RdDelta')
        lblRnSz2 = wx.StaticText(
            self, label=' Size(s) Smaller Than\nthe Main Run.',
            style=wx.ALIGN_LEFT)
        self.leftszr9.Add((20, 10))
        self.leftszr9.Add(lblRnSz1, 0, wx.ALIGN_CENTER | wx.LEFT, 10)
        self.leftszr9.Add(self.txtRTDif, 0, wx.ALIGN_CENTER | wx.LEFT, 10)
        self.leftszr9.Add(lblRnSz2, 0, wx.ALIGN_CENTER | wx.LEFT, 10)
        self.SizerLeft.Add(self.leftszr9, 0, wx.BOTTOM, 5)

        # Develop Widgets for right side of form
        self.SizerRight = wx.BoxSizer(wx.VERTICAL)

        self.rghtszr5 = wx.BoxSizer(wx.HORIZONTAL)
        # draw a line between upper and lower section
        lblRdT = wx.StaticText(
            self, label='3) For Forgings (O-Lets) use the following limits:',
            style=wx.ALIGN_RIGHT)
        lblRdT.SetForegroundColour('red')
        lblRdT.SetFont(font)
        self.rghtszr5.Add((30, 10))
        self.rghtszr5.Add(lblRdT, 0, wx.ALIGN_CENTER | wx.LEFT | wx.TOP, 10)
        self.SizerRight.Add((10, 15))
        self.SizerRight.Add(self.rghtszr5, 0, wx.BOTTOM, 5)

        self.rghtszr6 = wx.BoxSizer(wx.HORIZONTAL)
        lblBrSz1 = wx.StaticText(
            self, label='Branch Size Must Be\nLess Than or Equal to:',
            style=wx.ALIGN_LEFT)
        self.cmbFgb = wx.ComboCtrl(self, size=(80, -1), style=wx.CB_READONLY)
        self.cmbFgb.SetHint('Min. OD')
        self.cmbFgb.SetPopupControl(ListCtrlComboPopup('Pipe_OD', showcol=1))
        self.rghtszr6.Add((30, 10))
        self.rghtszr6.Add(lblBrSz1, 0, wx.ALIGN_CENTER | wx.LEFT | wx.TOP, 10)
        self.rghtszr6.Add((75, 10))
        self.rghtszr6.Add(self.cmbFgb, 0, wx.ALIGN_CENTER | wx.TOP, 10)
        self.SizerRight.Add(self.rghtszr6, 0, wx.BOTTOM, 5)

        self.rghtszr8 = wx.BoxSizer(wx.HORIZONTAL)
        lblRnSz4 = wx.StaticText(
            self, label='Branch Size Must\nBe At Least',
            style=wx.ALIGN_LEFT)
        self.txtFgDif = wx.TextCtrl(self, size=(30, 30), value='')
        self.txtFgDif.SetHint('RdDelta')
        lblRnSz5 = wx.StaticText(
            self, label=' Size(s) Smaller Than\nthe Main Run.',
            style=wx.ALIGN_LEFT)
        self.rghtszr8.Add((30, 10))
        self.rghtszr8.Add(lblRnSz4, 0, wx.ALIGN_CENTER | wx.LEFT | wx.TOP, 10)
        self.rghtszr8.Add(self.txtFgDif, 0, wx.ALIGN_CENTER |
                          wx.LEFT | wx.TOP, 10)
        self.rghtszr8.Add(lblRnSz5, 0, wx.ALIGN_CENTER | wx.LEFT | wx.TOP, 10)
        self.SizerRight.Add(self.rghtszr8, 0, wx.BOTTOM, 10)

        self.rghtszr9 = wx.BoxSizer(wx.HORIZONTAL)
        txt1 = '4) For Set-On fabrications use the following\nlimits.'
        txt2 = '  Set-On fabrications will require\nsupporting calculations.'
        txtlbl = txt1 + txt2
        lblBtmR = wx.StaticText(self, label=txtlbl, style=wx.ALIGN_LEFT)
        lblBtmR.SetForegroundColour('red')
        lblBtmR.SetFont(font)
        ln2 = wx.StaticLine(self, 0, size=(370, 2), style=wx.LI_VERTICAL)
        ln2.SetBackgroundColour('Black')
        self.rghtszr9.Add((30, 10))
        self.rghtszr9.Add(lblBtmR, 0, wx.ALIGN_CENTER | wx.LEFT | wx.TOP, 10)
        self.SizerRight.Add(ln2, 0, wx.ALIGN_LEFT)
        self.SizerRight.Add((10, 15))
        self.SizerRight.Add(self.rghtszr9, 0, wx.BOTTOM, 5)

        self.rghtszr12 = wx.BoxSizer(wx.HORIZONTAL)
        self.SOchk = wx.CheckBox(self, label='Allow Set-Ons ',
                                 style=wx.ALIGN_RIGHT)
        self.rghtszr12.Add((45, 10))
        self.rghtszr12.Add(self.SOchk, 0, wx.ALIGN_CENTER |
                           wx.LEFT | wx.TOP, 10)
        self.SizerRight.Add(self.rghtszr12, 0, wx.BOTTOM, 10)

        self.rghtszr10 = wx.BoxSizer(wx.HORIZONTAL)
        lblBWOD = wx.StaticText(
            self, label=('''Run Size Must Be\nGreater Than or Equal to:\n\nAND\
            \t\t  (Run size between)'''), style=wx.ALIGN_LEFT)
        self.cmbBWr2 = wx.ComboCtrl(self, size=(80, -1), style=wx.CB_READONLY)
        self.cmbBWr2.SetHint('Max. OD')
        self.cmbBWr2.SetPopupControl(ListCtrlComboPopup('Pipe_OD', showcol=1))
        self.rghtszr10.Add(lblBWOD, 0, wx.ALIGN_CENTER | wx.LEFT, 35)
        self.rghtszr10.Add(self.cmbBWr2, 0, wx.ALIGN_TOP | wx.RIGHT, 20)
        self.SizerRight.Add(self.rghtszr10, 0, wx.BOTTOM, 15)

        self.rghtszr7 = wx.BoxSizer(wx.HORIZONTAL)
        lblRnSz3 = wx.StaticText(
            self, label='Run Size Must Be\nLess Than or Equal to: \n\nAND',
            style=wx.ALIGN_LEFT)
        self.cmbBWr1 = wx.ComboCtrl(self, size=(80, -1), style=wx.CB_READONLY)
        self.cmbBWr1.SetHint('Min. OD')
        self.cmbBWr1.SetPopupControl(ListCtrlComboPopup('Pipe_OD', showcol=1))
        self.rghtszr7.Add(lblRnSz3, 0, wx.ALIGN_CENTER | wx.LEFT, 40)
        self.rghtszr7.Add(self.cmbBWr1, 0, wx.ALIGN_TOP | wx.LEFT, 75)
        self.SizerRight.Add(self.rghtszr7, 0, wx.BOTTOM, 5)

        self.rghtszr11 = wx.BoxSizer(wx.HORIZONTAL)
        lblRnSz6 = wx.StaticText(
            self, label='Branch Size Must\nBe At Least',
            style=wx.ALIGN_LEFT)
        self.txtBWDif = wx.TextCtrl(self, size=(30, 30), value='')
        self.txtBWDif.SetHint('RdDelta')
        lblRnSz7 = wx.StaticText(
            self, label=' Size(s) Smaller Than\nthe Main Run.',
            style=wx.ALIGN_LEFT)
        self.rghtszr11.Add((30, 10))
        self.rghtszr11.Add(lblRnSz6, 0, wx.ALIGN_CENTER | wx.LEFT | wx.TOP, 10)
        self.rghtszr11.Add(self.txtBWDif, 0, wx.ALIGN_CENTER | wx.LEFT, 10)
        self.rghtszr11.Add(lblRnSz7, 0, wx.ALIGN_CENTER | wx.LEFT | wx.TOP, 10)
        self.SizerRight.Add(self.rghtszr11, 0, wx.BOTTOM, 5)
        self.SizerRight.Add((10, 14))

        self.rghtszr13 = wx.BoxSizer(wx.HORIZONTAL)
        txtlbl = '5) For Sweep Outlet fabrications use the following limits.'
        lblSW = wx.StaticText(self, label=txtlbl, style=wx.ALIGN_LEFT)
        lblSW.SetForegroundColour('red')
        lblSW.SetFont(font)
        ln4 = wx.StaticLine(self, 0, size=(370, 2), style=wx.LI_VERTICAL)
        ln4.SetBackgroundColour('Black')
        self.rghtszr13.Add((30, 10))
        self.rghtszr13.Add(lblSW, 0, wx.ALIGN_CENTER | wx.LEFT | wx.TOP, 15)
        self.SizerRight.Add(ln4, 0, wx.ALIGN_LEFT)
        self.SizerRight.Add(self.rghtszr13, 0, wx.BOTTOM, 5)

        self.rghtszr14 = wx.BoxSizer(wx.HORIZONTAL)
        self.SWchk = wx.CheckBox(self, label='Allow Sweep Outlets ',
                                 style=wx.ALIGN_RIGHT)
        self.rghtszr14.Add((45, 10))
        self.rghtszr14.Add(self.SWchk, 0,
                           wx.ALIGN_CENTER | wx.LEFT | wx.TOP, 10)
        self.SizerRight.Add(self.rghtszr14, 0, wx.BOTTOM, 5)

        self.rghtszr16 = wx.BoxSizer(wx.HORIZONTAL)
        lblRnSz8 = wx.StaticText(self, label='Branch Size Must\nBe At Least',
                                 style=wx.ALIGN_LEFT)
        self.txtSWDif = wx.TextCtrl(self, size=(30, 30), value='')
        self.txtSWDif.SetHint('RdDelta')
        lblRnSz9 = wx.StaticText(
            self, label=' Size(s) Smaller Than\nthe Main Run.',
            style=wx.ALIGN_LEFT)
        self.rghtszr16.Add((30, 10))
        self.rghtszr16.Add(lblRnSz8, 0, wx.ALIGN_CENTER | wx.LEFT | wx.TOP, 10)
        self.rghtszr16.Add(self.txtSWDif, 0, wx.ALIGN_CENTER | wx.LEFT, 10)
        self.rghtszr16.Add(lblRnSz9, 0, wx.ALIGN_CENTER | wx.LEFT | wx.TOP, 10)
        self.SizerRight.Add(self.rghtszr16, 0, wx.BOTTOM, 5)

        btnbox = wx.BoxSizer(wx.VERTICAL)
        self.b5 = wx.Button(self, label="Exit")
        btnbox.Add(self.b5, 0, wx.RIGHT | wx.ALIGN_CENTER, 15)
        self.b5.SetForegroundColour('red')
        self.SizerRight.Add((25, 10))
        self.SizerRight.Add(btnbox, 1, wx.EXPAND)

        self.Sizer.Add(self.SizerLeft, 0, wx.LEFT)
        self.Sizer.Add(self.SizerRight, 0, wx.LEFT)

        self.SetupScrolling()
        self.FillScreen()

    def FillScreen(self):
        if self.data != []:
            self.cmbMin.ChangeValue(self.data[0][1])
            self.cmbMax.ChangeValue(self.data[0][2])
            self.cmbRTr.ChangeValue(self.data[0][3])
            self.txtRTDif.SetValue(self.data[0][4])
            self.cmbFgb.ChangeValue(self.data[0][5])
            self.txtFgDif.SetValue(self.data[0][6])
            self.SOchk.SetValue(self.data[0][7])
            if self.data[0][7]:
                self.cmbBWr2.ChangeValue(self.data[0][8])
                self.cmbBWr1.ChangeValue(self.data[0][9])
                self.txtBWDif.SetValue(self.data[0][10])
            else:
                self.cmbBWr2.ChangeValue('')
                self.cmbBWr1.ChangeValue('')
                self.txtBWDif.SetValue('')
            self.SWchk.SetValue(self.data[0][11])
            if self.data[0][11]:
                self.txtSWDif.SetValue(self.data[0][12])
            else:
                self.txtSWDif.SetValue('')
        else:
            self.cmbMin.ChangeValue('')
            self.cmbMax.ChangeValue('')
            self.cmbRTr.ChangeValue('')
            self.txtRTDif.SetValue('')
            self.cmbFgb.ChangeValue('')
            self.txtFgDif.SetValue('')
            self.SOchk.SetValue(False)
            self.cmbBWr2.ChangeValue('')
            self.cmbBWr1.ChangeValue('')
            self.txtBWDif.SetValue('')
            self.cmbBWr2.ChangeValue('')
            self.cmbBWr1.ChangeValue('')
            self.txtBWDif.SetValue('')
            self.SWchk.SetValue(False)
            self.txtSWDif.SetValue('')
            self.txtSWDif.SetValue('')

    def OnDelete(self, evt):
        self.ChgSpecID()
        Dbase().TblDelete('BranchCriteria', str(self.CurrentID), 'BranchID')
        self.data = []
        self.FillScreen()
        self.txtComd.SetLabel('No Data for this commodity.')
        self.b3.Enable(False)
        self.b4.Enable(False)

        # called to update the item table and commodity table if needed
    def OnAddRec(self, evt):
        # check first that data is all present and clean up incomplete boxes
        check = self.ValData()
        if check:
            SQL_step = 5101   # default to cancel action
            dlg = wx.MessageDialog(
                self, "Update/Save Criteria to\nCommodity Property",
                "Data Change", wx.OK | wx.CANCEL | wx.ICON_QUESTION)
            SQL_step = dlg.ShowModal()
            dlg.Destroy()
            if SQL_step == 5100:
                if self.data == []:
                    SQL_step = 0
                else:
                    SQL_step = 1

            self.AddRec(SQL_step)    # always save as a new spec if commodity
            # property is specified no matter the change over write or add the
            # specification ID to the PipeSpecification table under the
            # commodity property ID
            self.ChgSpecID(self.CurrentID)

            query = ('''SELECT Branch_ID FROM PipeSpecification WHERE
                     Commodity_Property_ID = ''' + str(self.ComdPrtyID))
            StartQry = Dbase().Dsqldata(query)

            self.MainSQL = ('SELECT * FROM BranchCriteria WHERE BranchID = ' +
                            str(StartQry[0][0]))
            self.data = Dbase().Dsqldata(self.MainSQL)

            self.FillScreen()
            self.b3.Enable()
            self.b4.Enable()

    def ValData(self):
        self.ValueList = []
        name = []

        colinfo = Dbase().Dcolinfo(self.tblname)
        self.num_vals = ('?,'*len(colinfo))[:-1]
        self.ValueList = [None for i in range(0, len(colinfo))]
        name = ['' for i in range(len(colinfo))]

        self.ValueList[0] = self.CurrentID
        name[0] = 'New BranchID'
        self.ValueList[1] = self.cmbMin.GetValue()
        name[1] = 'Minimum Diameter'
        self.ValueList[2] = self.cmbMax.GetValue()
        name[2] = 'Maximum Diameter'
        self.ValueList[3] = self.cmbRTr.GetValue()
        name[3] = 'Reducing Tee maximum run size'
        self.ValueList[4] = self.txtRTDif.GetValue()
        name[4] = 'Reducing Tee size delta between branch and run'
        self.ValueList[5] = self.cmbFgb.GetValue()
        name[5] = 'O-Let maximum branch size'
        self.ValueList[6] = self.txtFgDif.GetValue()
        name[6] = 'O-Let size delta between branch and run'
        self.ValueList[7] = self.SOchk.GetValue()
        if self.ValueList[7] is True:
            self.ValueList[8] = self.cmbBWr2.GetValue()
            name[8] = 'Set-On minimum run size'
            self.ValueList[9] = self.cmbBWr1.GetValue()
            name[9] = 'Set-On maximum run size'
            self.ValueList[10] = self.txtBWDif.GetValue()
            name[10] = 'Set-On size delta between branch and run'
        self.ValueList[11] = self.SWchk.GetValue()
        if self.ValueList[11] is True:
            self.ValueList[12] = self.txtSWDif.GetValue()
            name[12] = 'Sweept Outlet size delta between branch and run'

        for n in range(1, len(self.ValueList)):
            if n in range(7, 13):
                if self.ValueList[7] is False:
                    continue
                elif n == 7:
                    continue
                if self.ValueList[11] is False:
                    continue
                elif n == 11:
                    continue
            if self.ValueList[n] == '':
                txt = 'Data needed for:\n' + name[n]
                wx.MessageBox(txt, 'Info', wx.OK | wx.ICON_INFORMATION)
                return False

        maxOD = eval(self.cmbMax.GetValue()[:-1])
        minOD = eval(self.cmbMin.GetValue()[:-1])
        if maxOD-minOD <= 0:
            wx.MessageBox('''Maximum diameter must be greater then
                          Minimum Diameter''', 'Info',
                          wx.OK | wx.ICON_INFORMATION)
            return False

        return True

    def AddRec(self, SQL_step):
        realnames = ['BranchID', 'Min', 'Max', 'RTMax', 'RTDelta', 'FgMin',
                     'FgDelta', 'SOchk', 'SOMin', 'SOMax', 'SODelta', 'SWchk',
                     'SWDelta']
        ValList = []

        if SQL_step == 0:    # enter new record
            if self.NewPipeSpec is True:
                New_ID = cursr.execute(
                    ''''SELECT MAX(Pipe_Spec_ID) FROM
                     PipeSpecification''').fetchone()
                if New_ID[0] is None:
                    Max_ID = '1'
                else:
                    Max_ID = str(New_ID[0]+1)
                colinfo = Dbase().Dcolinfo('PipeSpecification')
                for n in range(0, len(colinfo)-2):
                    ValList.append(None)
                num_val = ('?,'*len(colinfo))[:-1]
                ValList.insert(0, Max_ID)
                ValList.insert(1, str(self.ComdPrtyID))
                UpQry = ('''INSERT INTO PipeSpecification VALUES ('''
                         + num_val + ')')
                Dbase().TblEdit(UpQry, ValList)

            UpQuery = ('''INSERT INTO BranchCriteria VALUES ('''
                       + self.num_vals + ')')
            Dbase().TblEdit(UpQuery, self.ValueList)

        elif SQL_step == 1:    # update edited record
            realnames.remove('BranchID')
            del self.ValueList[0]
            SQL_str = ','.join(["%s=?" % (name) for name in realnames])
            UpQuery = ('UPDATE BranchCriteria SET ' + SQL_str +
                       ' WHERE BranchID = ' + str(self.CurrentID))
            Dbase().TblEdit(UpQuery, self.ValueList)
            self.data = Dbase().Dsqldata(self.MainSQL)

    def ChgSpecID(self, ID=None):
        UpQuery = ('''UPDATE PipeSpecification SET Branch_ID=?
                   WHERE Commodity_Property_ID = ''' + str(self.ComdPrtyID))
        Dbase().TblEdit(UpQuery, [ID])

    def OnPrint(self, evt):

        saveDialog = wx.FileDialog(self, message='Save Report as PDF.',
                                   wildcard='PDF (*.pdf)|*.pdf',
                                   style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT)

        if saveDialog.ShowModal() == wx.ID_CANCEL:
            filename = ''

        filename = saveDialog.GetPath()
        if filename.find(".pdf") == -1:
            filename = filename + '.pdf'

        saveDialog.Destroy()

        self.PrintPdf(filename)

    def PrintPdf(self, filename):
        import Report_Branch
        filename = filename

        if self.PipeCode is None:
            ttl = self.ComCode + '-' + self.mtrcode
        else:
            ttl = self.PipeCode

        self.Dia = []
        self.Axis = []
        self.rMin = self.cmbMin.GetValue()
        qry = ("SELECT PipeOD_ID FROM Pipe_OD WHERE Pipe_OD = '" +
               self.rMin + "'")
        self.rMinID = Dbase().Dsqldata(qry)[0][0]

        self.rMax = self.cmbMax.GetValue()
        qry = ("SELECT PipeOD_ID FROM Pipe_OD WHERE Pipe_OD = '" +
               self.rMax + "'")
        self.rMaxID = Dbase().Dsqldata(qry)[0][0]

        qry = ("SELECT Pipe_OD FROM Pipe_OD WHERE PipeOD_ID >= "
               + str(self.rMinID) + ' AND PipeOD_ID <= ' + str(self.rMaxID))
        Axis = Dbase().Dsqldata(qry)

        for item in Axis:
            self.Axis.append(item[0])
        for D in self.Axis:
            self.Dia.append(eval(D[:-1].replace('-', '+')))

        nRTr = eval(self.cmbRTr.GetValue()[:-1].replace('-', '+'))
        nFgb = eval(self.cmbFgb.GetValue()[:-1].replace('-', '+'))
        nBWr1 = eval(self.cmbBWr1.GetValue()[:-1].replace('-', '+'))
        nBWr2 = eval(self.cmbBWr2.GetValue()[:-1].replace('-', '+'))
        nRTDif = eval(self.txtRTDif.GetValue())
        nFgDif = eval(self.txtFgDif.GetValue())
        nBWDif = eval(self.txtBWDif.GetValue())

        row = []
        rows = []
        Rindex = 0
        for r in self.Dia:
            Bindex = 0
            row = []
            for b in self.Dia:
                if b <= r:
                    # equal tee
                    if b == r:
                        row.append('ET')
                    # Reducing Tee determined by run size
                    elif (r <= nRTr) or (Rindex-Bindex) <= nRTDif:
                        row.append('RT')
                    # O-Let fittings
                    elif (b <= nFgb and (Rindex-Bindex) >= nFgDif):
                        row.append('OL')
                    # Set-On w/wo Repad
                    elif (self.SOchk.GetValue() is True and b >= nFgb and
                          (Rindex - Bindex) >= nBWDif and nBWr1 >= r >= nBWr2):
                        row.append('BO')
                    # Sweep Outlet
                    elif (self.SWchk.GetValue() is True and
                          r > nBWr1 and b > nFgb):
                        row.append('SW')
                    else:
                        row.append('Eng')
                    Bindex += 1
                else:
                    continue
            row.insert(0, self.Axis[Rindex])
            Rindex += 1
            rows.append(row)

        self.Axis.insert(0, 'Run^')

        rows.append(self.Axis)
        colsz = 560//len(self.Axis)

        Report_Branch.Report(rows, colsz, self.End, filename, ttl)


class BldSpc_Nts(wx.Frame):
    '''Routine to build form and populate grid'''
    def __init__(self, parent, tblname, ComdPrtyID=None, model=None):

        self.parent = parent
        self.ComdPrtyID = ComdPrtyID
        # set up the table column names, width and if column can be
        # edited ie primary autoincrement
        self.Lvl2tbl = tblname
        self.model = model

        if self.Lvl2tbl.find("_") != -1:
            frmtitle = (self.Lvl2tbl.replace("_", " "))
        else:
            frmtitle = (' '.join(re.findall('([A-Z][a-z]*)', self.Lvl2tbl)))

        super(BldSpc_Nts, self).__init__(parent,
                                         title=frmtitle,
                                         size=(1200, 900))

        self.Bind(wx.EVT_CLOSE, self.OnClose)

        self.InitUI()

    def InitUI(self):
        self.lctrls = []
        self.ComCode = ''
        self.PipeMtrSpec = ''
        font1 = wx.Font(12, wx.DEFAULT, wx.NORMAL, wx.FONTWEIGHT_BOLD)
        model1 = None

        if self.ComdPrtyID is not None:
            query = ('''SELECT Commodity_Code,Pipe_Material_Code,Pipe_Code FROM
                      CommodityProperties WHERE CommodityPropertyID = '''
                     + str(self.ComdPrtyID))
            dataset = Dbase().Dsqldata(query)[0]
            self.PipeMtrSpec = dataset[1]
            self.ComCode = dataset[0]
            self.PipeCode = dataset[2]

        if self.Lvl2tbl == 'Notes':
            link_fld = 'Category_ID'
            self.frg_tbl = 'NoteCategory'
            self.frg_fld = 'CategoryID'
            frgn_col = 'Category'
            self.columnames = ['ID', 'Category', 'Note']
            spec_col = 19
        elif self.Lvl2tbl == 'Specials':
            link_fld = 'Item_Type_ID'
            self.frg_tbl = 'SpecialItems'
            self.frg_fld = 'SpecialTypeID'
            frgn_col = 'Item_Type'
            self.columnames = ['ID', 'Item Type', 'Description', 'Notes',
                               'Vendor', 'Part Number']
            spec_col = 17

        realnames = []
        datatable_str = ''

        # here we get the information needed in the report
        # and for the SQL from the Lvl2 table
        # and determine report column width based either
        # on data type or specified field size
        n = 0
        ColWdth = []
        for item in Dbase().Dcolinfo(self.Lvl2tbl):
            col_wdth = ''
        # check to see if field length is specified if
        # so use it to set grid col width
            for s in re.findall(r'\d+', item[2]):
                if s.isdigit():
                    col_wdth = int(s)
                    ColWdth.append(col_wdth)
            realnames.append(item[1])
            if item[5] == 1:
                self.pk_Name = item[1]
                self.pk_col = n
                if 'INTEGER' in item[2]:
                    self.autoincrement = True
                    if col_wdth == '':
                        ColWdth.append(6)
                # include the primary key and table name into SELECT statement
                datatable_str = (datatable_str + self.Lvl2tbl + '.'
                                 + self.pk_Name + ',')
            # need to make frgn_fld column noneditable in DVC
            elif 'INTEGER' in item[2] or 'FLOAT' in item[2]:
                if col_wdth == '' and 'FLOAT' in item[2]:
                    ColWdth.append(10)
                elif col_wdth == '':
                    ColWdth.append(6)
            elif 'BLOB' in item[2]:
                if col_wdth == '':
                    ColWdth.append(30)
            elif 'TEXT' in item[2] or 'BOOLEAN' in item[2]:
                if col_wdth == '':
                    ColWdth.append(10)
            elif 'DATE' in item[2]:
                if col_wdth == '':
                    ColWdth.append(10)

            # get first Lvl2 datatable column name in item[1]
            # check to see if name is lvl2 primary key or lvl1 linked field
            # if they are not then add tablename and
            # datafield to SELECT statement
            if item[1] != link_fld and item[1] != self.pk_Name:
                datatable_str = (datatable_str + ' ' + self.Lvl2tbl +
                                 '.' + item[1] + ',')
            elif item[1] == link_fld:
                datatable_str = (datatable_str + ' ' + self.frg_tbl +
                                 '.' + frgn_col + ',')

            n += 1
        self.realnames = realnames
        self.ColWdth = ColWdth

        datatable_str = datatable_str[:-1]

        DsqlLvl2 = ('SELECT ' + datatable_str + ' FROM ' + self.Lvl2tbl
                    + ' INNER JOIN ' + self.frg_tbl)
        DsqlLvl2 = (DsqlLvl2 + ' ON ' + self.Lvl2tbl + '.' + link_fld
                    + ' = ' + self.frg_tbl + '.' + self.frg_fld)

        # specify data for upper dvc with all the notes in databasse
        self.DsqlLvl2 = DsqlLvl2
        self.data = Dbase().Dsqldata(self.DsqlLvl2)
        # if there is a specified ComdPrtyID then show dvc1
        if self.ComdPrtyID is not None:
            query = ('''SELECT * FROM PipeSpecification WHERE
                      Commodity_Property_ID = ''' + str(self.ComdPrtyID))
            chk = Dbase().Dsqldata(query)
            # the commodity property exists in pipe secification
            # and has Notes assigned
            if chk != []:
                # this is the list of the IDs for the Notes
                self.NoteStr = chk[0][spec_col]
                # if there are notes assinged the data to dvc1
                if self.NoteStr is not None:
                    # populate the lower dvc1 with just the notes
                    # related to the commodity property
                    self.Dsql1 = (self.DsqlLvl2 + ' WHERE ' + self.pk_Name +
                                  ' IN ' + self.NoteStr)
                    self.data1 = Dbase().Dsqldata(self.Dsql1)
                    txtlbl1 = ('   ' + self.Lvl2tbl +
                               ' Related to Commodity Property ')
                # if there are no notes specifed then assign empty data to dvc1
                else:
                    txtlbl1 = ('  No ' + self.Lvl2tbl +
                               ' have been related to this Commodity Property')
                    self.data1 = []

            else:
                ValList = []
                New_ID = cursr.execute(
                    '''SELECT MAX(Pipe_Spec_ID) FROM
                     PipeSpecification''').fetchone()
                if New_ID[0] is None:
                    Max_ID = '1'
                else:
                    Max_ID = str(New_ID[0]+1)

                colinfo = Dbase().Dcolinfo('PipeSpecification')
                for n in range(0, len(colinfo)-2):
                    ValList.append(None)

                num_val = ('?,'*len(colinfo))[:-1]
                ValList.insert(0, Max_ID)
                ValList.insert(1, str(self.ComdPrtyID))
                UpQry = ("INSERT INTO PipeSpecification VALUES ("
                         + num_val + ")")
                Dbase().TblEdit(UpQry, ValList)

                txtlbl1 = ('No ' + self.Lvl2tbl +
                           ' have been related to this Commodity Property   ')
                self.data1 = []
                self.NoteStr = None

            # Create the dataview for the commocity property notes
            self.dvc1 = dv.DataViewCtrl(self, wx.ID_ANY, wx.DefaultPosition,
                                        wx.Size(500, 150),
                                        style=wx.BORDER_THEME
                                        | dv.DV_ROW_LINES
                                        | dv.DV_VERT_RULES
                                        | dv.DV_HORIZ_RULES
                                        | dv.DV_MULTIPLE
                                        | dv.DV_VARIABLE_LINE_HEIGHT
                                        )

            if model1 is None:
                self.model1 = DataMods(self.Lvl2tbl, self.data1)
            else:
                self.model1 = model1
            self.dvc1.AssociateModel(self.model1)

        # specify which listbox column to display in the combobox
        self.showcol = int

        # Create a dataview control for all the database notes
        self.dvc = dv.DataViewCtrl(self, wx.ID_ANY, wx.DefaultPosition,
                                   wx.Size(500, 300),
                                   style=wx.BORDER_THEME
                                   | dv.DV_ROW_LINES
                                   | dv.DV_VERT_RULES
                                   | dv.DV_HORIZ_RULES
                                   | dv.DV_MULTIPLE
                                   | dv.DV_VARIABLE_LINE_HEIGHT
                                   )

        # pull out the foreign field name for the lable and combo box text
        self.frg_tbl = Dbase().Dtbldata(self.Lvl2tbl)[0][2]
        frg_fld = Dbase().Dtbldata(self.Lvl2tbl)[0][4]
        self.frg_fld = frg_fld.replace("_", " ")

        # if autoincrement is false then the data can be sorted based on ID_col
        if self.autoincrement == 0:
            self.data.sort(key=lambda tup: tup[self.pk_col])

        # use the sorted data to load the dataviewlistcontrol
        if self.model is None:
            self.model = DataMods(self.Lvl2tbl, self.data)

        self.dvc.AssociateModel(self.model)

        n = 0
        for colname in self.columnames:
            if n == self.pk_col and self.autoincrement or n == self.pk_col:
                col_mode = dv.DATAVIEW_CELL_INERT
            else:
                if n == 1:
                    col_mode = dv.DATAVIEW_CELL_INERT
                else:
                    col_mode = dv.DATAVIEW_CELL_EDITABLE

            self.dvc.AppendTextColumn(colname, n,
                                      width=wx.LIST_AUTOSIZE_USEHEADER,
                                      mode=col_mode)
            if self.ComdPrtyID is not None:
                self.dvc1.AppendTextColumn(colname, n,
                                           width=wx.LIST_AUTOSIZE_USEHEADER,
                                           mode=dv.DATAVIEW_CELL_INERT)
            n += 1

        # make columns not sortable and but reorderable.
        n = 0
        for c in self.dvc.Columns:
            c.Sortable = False
            # make the category column sortable
            if n == 1:
                c.Sortable = True
            c.Reorderable = True
            c.Resizeable = True
            n += 1

        # change to not let the ID col be moved.
        self.dvc.Columns[(self.pk_col)].Reorderable = False
        self.dvc.Columns[(self.pk_col)].Resizeable = False

        # Bind some events so we can see what the DVC sends us
        self.Bind(dv.EVT_DATAVIEW_ITEM_VALUE_CHANGED,
                  self.OnValueChanged, self.dvc)

        # set the Sizer property (same as SetSizer)
        self.Sizer = wx.BoxSizer(wx.VERTICAL)

        # develope the comboctrl and attach popup list
        self.cmb1 = wx.ComboCtrl(self, pos=(10, 10), size=(200, -1))
        self.Bind(wx.EVT_TEXT, self.OnSelect, self.cmb1)
        self.cmb1.SetHint(frgn_col)
        self.showcol = 1
        self.popup = ListCtrlComboPopup(self.frg_tbl, showcol=self.showcol,
                                        lctrls=self.lctrls)

        self.cmbsizer = wx.BoxSizer(wx.HORIZONTAL)

        self.lblSrch = wx.StaticText(self, -1,
                                     style=wx.ALIGN_CENTER_HORIZONTAL)
        txt = '   To add a new record\nfirst select a Note Category    '
        self.lblSrch.SetLabel(txt)
        self.lblSrch.SetForegroundColour((255, 0, 0))
        self.cmbsizer.Add(self.lblSrch, 0, wx.ALIGN_TOP)

        # add a button to call main form to search combo list data
        self.b6 = wx.Button(self, label="Restore Data")
        self.Bind(wx.EVT_BUTTON, self.OnRestore, self.b6)
        self.cmbsizer.Add(self.b6, 0)

        self.cmb1.SetPopupControl(self.popup)
        self.cmbsizer.Add(self.cmb1, 0, wx.ALIGN_TOP)

        # add a button to call main form to search combo list data
        self.b5 = wx.Button(self, label="<= Search Data")
        self.Bind(wx.EVT_BUTTON, self.OnSearch, self.b5)
        self.cmbsizer.Add(self.b5, 0, wx.BOTTOM, 15)

        font1 = wx.Font(10, wx.DEFAULT, wx.NORMAL, wx.FONTWEIGHT_BOLD)
        self.addcmb1 = wx.Button(self, label='+', size=(35, -1))
        self.addcmb1.SetForegroundColour((255, 0, 0))
        self.addcmb1.SetFont(font1)
        self.Bind(wx.EVT_BUTTON, self.OnAddcmb1, self.addcmb1)
        self.cmbsizer.Add(self.addcmb1, 0)

        self.addlbl = wx.StaticText(self, -1, style=wx.ALIGN_CENTER_HORIZONTAL)
        txt = '   Complete Listing of All ' + self.Lvl2tbl
        self.addlbl.SetLabel(txt)
        self.addlbl.SetForegroundColour((255, 0, 0))
        self.addlbl.SetFont(font1)
        self.Sizer.Add(self.addlbl, 0, wx.ALIGN_LEFT)
        self.Sizer.Add(self.cmbsizer, 0, wx.ALIGN_CENTER)
        self.Sizer.Add(self.dvc, 1, wx.EXPAND)

        # Add buttons for grid modifications
        self.b1 = wx.Button(self, id=1, label="Print Report")
        self.Bind(wx.EVT_BUTTON, self.PrintFile, self.b1)
        self.b2 = wx.Button(self, label="Add Row")
        self.Bind(wx.EVT_BUTTON, self.OnAddRow, self.b2)
        self.b3 = wx.Button(self, label="Delete Row")
        self.Bind(wx.EVT_BUTTON, self.OnDeleteRow, self.b3)

        self.b9 = wx.Button(self, label='Add Selected to\nCommodity Property')
        self.Bind(wx.EVT_BUTTON, self.OnSelectDone, self.b9)

        # add a button box and place the buttons
        self.btnbox = wx.BoxSizer(wx.HORIZONTAL)
        self.btnbox.Add(self.b1, 0, wx.ALL | wx.ALIGN_CENTER, 5)
        self.btnbox.Add(self.b2, 0, wx.ALL | wx.ALIGN_CENTER, 5)
        self.btnbox.Add(self.b3, 0, wx.ALL | wx.ALIGN_CENTER, 5)
        self.btnbox.Add(self.b9, 0, wx.ALL | wx.ALIGN_CENTER, 5)

        if self.ComdPrtyID is None:
            self.b4 = wx.Button(self, label="Exit")
            self.b4.SetForegroundColour('red')
            self.Bind(wx.EVT_BUTTON, self.OnClose, self.b4)
            self.btnbox.Add((30, 10))
            self.btnbox.Add(self.b4, 0, wx.ALL | wx.ALIGN_CENTER, 35)

        # need to disable the add row button until an
        # item is selected in the combo box
        self.b2.Disable()

        # add static label to explain how to add / edit data
        self.editlbl = wx.StaticText(self, -1,
                                     style=wx.ALIGN_CENTER)
        txt = '    \nTo edit data double click on the cell.'
        self.editlbl.SetLabel(txt)
        self.editlbl.SetForegroundColour((255, 0, 0))

        self.Sizer.Add(self.editlbl, 0, wx.ALIGN_CENTER)
        self.Sizer.Add(self.btnbox, 0, wx.ALIGN_CENTER, 5)

        if self.ComdPrtyID is not None:
            qry = ('''SELECT Pipe_Material_Spec FROM PipeMaterialSpec
                    WHERE Material_Spec_ID = ''' + str(self.PipeMtrSpec))
            self.mtrcode = Dbase().Dsqldata(qry)[0][0]
            ln = wx.StaticLine(self, 0, size=(350, 2), style=wx.LI_VERTICAL)
            ln.SetBackgroundColour('Black')
            self.Sizer.Add(ln, 0, wx.EXPAND | wx.TOP | wx.BOTTOM, 10)

            self.title1 = wx.BoxSizer(wx.HORIZONTAL)
            self.addlbl1 = wx.StaticText(self, -1,
                                         style=wx.ALIGN_CENTER_HORIZONTAL)
            txtlbl1 = txtlbl1
            self.addlbl1.SetLabel(txtlbl1)
            self.addlbl1.SetFont(font1)
            self.addlbl1.SetForegroundColour((255, 0, 0))

            self.addlbl2 = wx.StaticText(self, -1,
                                         style=wx.ALIGN_CENTER_HORIZONTAL)
            txtlbl2 = self.ComCode + ' - ' + self.mtrcode
            self.addlbl2.SetLabel(txtlbl2)
            self.addlbl2.SetFont(font1)
            self.addlbl2.SetForegroundColour((255, 0, 0))

            self.title1.Add(self.addlbl1, 0)
            self.title1.Add(self.addlbl2, 0)
            self.Sizer.Add(self.title1, 0)

            self.cmbsizer2 = wx.BoxSizer(wx.HORIZONTAL)
            # add a button to call main form to search combo list data
            self.b10 = wx.Button(self, label="Restore Data")
            self.Bind(wx.EVT_BUTTON, self.OnRstrDVC1, self.b10)
            self.cmbsizer2.Add(self.b10, 0)

            # develope the comboctrl and attach popup list
            self.cmb10 = wx.ComboCtrl(self, pos=(10, 10), size=(200, -1))
            self.Bind(wx.EVT_TEXT, self.OnSelect, self.cmb10)
            self.cmb10.SetHint('Note Category')
            self.showcol = 1
            self.popup = ListCtrlComboPopup(self.frg_tbl, showcol=self.showcol,
                                            lctrls=self.lctrls)
            self.cmb10.SetPopupControl(self.popup)
            self.cmbsizer2.Add(self.cmb10, 0, wx.ALIGN_TOP, 5)

            # add a button to call main form to search combo list data
            self.b11 = wx.Button(self, label="<= Search Data")
            self.Bind(wx.EVT_BUTTON, self.OnSrchDVC1, self.b11)
            self.cmbsizer2.Add(self.b11, 0)

            self.b8 = wx.Button(self, id=2, label="Print Commodity\nReport")
            self.Bind(wx.EVT_BUTTON, self.PrintFile, self.b8)
            self.cmbsizer2.Add(self.b8, 0)

            self.b12 = wx.Button(
                self, label='Remove ' + self.Lvl2tbl +
                ' from\nCommodity Property')
            self.Bind(wx.EVT_BUTTON, self.OnRmvNote, self.b12)
            self.cmbsizer2.Add((30, 10))
            self.cmbsizer2.Add(self.b12, 0)

            self.b4 = wx.Button(self, label="Exit")
            self.b4.SetForegroundColour('red')
            self.Bind(wx.EVT_BUTTON, self.OnClose, self.b4)
            self.cmbsizer2.Add((60, 10))
            self.cmbsizer2.Add(self.b4, 0, wx.ALL | wx.ALIGN_CENTER, 5)

            self.Sizer.Add((10, 20))
            self.Sizer.Add(self.dvc1, 1, wx.EXPAND)
            self.Sizer.Add((10, 15))
            self.Sizer.Add(self.cmbsizer2, 0, wx.ALIGN_CENTER)
            self.Sizer.Add((10, 20))

        self.CenterOnParent()
        self.GetParent().Enable(False)
        self.Show(True)

        self.__eventLoop = wx.GUIEventLoop()
        self.__eventLoop.Run()

    def PrintFile(self, evt):
        import Report_Lvl2

        if self.Lvl2tbl == 'Notes':
            colwdths = [5, 20, 70]
            ttlname = 'Commodity Notes'
        elif self.Lvl2tbl == 'Specials':
            colwdths = [5, 15, 40, 40, 15, 15]
            ttlname = 'Specialty Items'
        rptdata = []
        # if commodity related report is selected
        if evt.GetId() == 2:
            # confirm there is data to print
            if self.data1 == []:
                NoData = wx.MessageDialog(
                    None, 'No Data to Print', 'Error', wx.OK |
                    wx.ICON_EXCLAMATION)
                NoData.ShowModal()
                return
            # specify a title for the report if table name is not to be used
            if self.PipeCode is None:
                ttl = self.ComCode + '-' + self.mtrcode
            else:
                ttl = self.PipeCode
            ttl = ttlname + ' for ' + ttl

            # collect the raw data for the report and seperate it
            # from the forms data
            for ln in self.data1:
                rptdata.append(tuple(ln))
        # if the full list of notes is to be printed
        else:
            if self.data == []:
                NoData = wx.MessageDialog(
                    None, 'No Data to Print', 'Error', wx.OK |
                    wx.ICON_EXCLAMATION)
                NoData.ShowModal()
                return

            ttl = ttlname

            for ln in self.data:
                rptdata.append(tuple(ln))

        filename = self.ReportName()
        Report_Lvl2.Report(self.Lvl2tbl, rptdata, self.columnames,
                           colwdths, filename, ttl).create_pdf()

    def ReportName(self):

        saveDialog = wx.FileDialog(self, message='Save Report as PDF.',
                                   wildcard='PDF (*.pdf)|*.pdf',
                                   style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT)
        if saveDialog.ShowModal() == wx.ID_CANCEL:
            filename = ''
        filename = saveDialog.GetPath()
        if filename.find(".pdf") == -1:
            filename = filename + '.pdf'
        saveDialog.Destroy()

        if not filename:
            exit()

        return filename

    def TblCellEdit(self):
        # rowID is table index value, colID is column number of table index
        # new_value is the link foreign field value
        # colChgNum is the column number which has edited cell,
        # values is the new edited value(s)
        self.edit_data = self.model.edit_data
        rowID = self.model.GetValueByRow(self.edit_data[0], self.pk_col)
        self.edit_data.append(rowID)

        # if just a cell is being edited then do the following
        colID = self.pk_col
        colChgNum = self.edit_data[1]
        values = self.edit_data[2]
        rowID = int(self.edit_data[3])
        colIDName = self.realnames[colID]
        colChgName = self.realnames[colChgNum]

        if values.find('"') != -1:
            values = values.replace('"', '""')
        if type(rowID) != str:
            UpQuery = ('UPDATE ' + self.Lvl2tbl + ' SET ' + colChgName
                       + ' = "' + str(values) + '" WHERE ' + str(colIDName)
                       + ' = ' + str(rowID))
        else:
            UpQuery = ('UPDATE ' + self.Lvl2tbl + ' SET ' + colChgName + ' = "'
                       + str(values) + '" WHERE ' + str(colIDName)
                       + ' = "' + rowID + '"')

        Dbase().TblEdit(UpQuery)

        # check to see if added row has been properly
        # edited if so enable the addrow button
        self.data1 = Dbase().Restore(self.Dsql1)
        self.model1 = DataMods(self.Lvl2tbl, self.data1)
        self.dvc1._AssociateModel(self.model1)
        self.dvc1.Refresh

    def AddTblRow(self):
        # this returns the text value in the combo box not the linked ID value
        new_value = self.cmb1.GetValue()

        lnk_fld_name = Dbase().Dtbldata(self.Lvl2tbl)

        # this is the name of the linked field in the foreign table
        field = lnk_fld_name[0][4]
        # do a search in the foriegn table for the value of the
        # link ID which corresponds to the name in the combo box
        Shcol = Dbase().Dcolinfo(self.frg_tbl)[self.showcol][1]
        ShQuery = ('SELECT ' + field + ' FROM ' + self.frg_tbl +
                   " WHERE " + Shcol + " LIKE '%" + self.cmb1.GetValue()
                   + "%' COLLATE NOCASE")
        # this is the ID value for the linked field in the foreign table
        ShQueryVal = str(Dbase().Dsqldata(ShQuery)[0][0])

        if new_value:
            self.b2.Disable()
        # update data structure
            FldInfo = Dbase(self.Lvl2tbl).Fld_Size_Type()
            values = FldInfo[4]

        # if the table index is auto increment then
        # assign next value otherwise do nothing
            if FldInfo[2]:
                New_ID = cursr.execute("SELECT MAX(" + self.pk_Name +
                                       ") FROM " + self.Lvl2tbl).fetchone()
                if New_ID[0] is None:
                    Max_ID = '1'
                else:
                    Max_ID = str(New_ID[0]+1)
                values[self.pk_col] = Max_ID

        # if a full row is being added then do the following
            ValueList = []
            if type(values) == str:
                ValueList.append(values)
            else:
                ValueList = values

        # need to change the linked foriegn field to
        # the value selected in the combobox
            if lnk_fld_name != []:
                ValueList[self.realnames.index(
                    lnk_fld_name[0][3])] = ShQueryVal
            UpQuery = ('INSERT INTO ' + self.Lvl2tbl + ' VALUES (' + "'" +
                       "','".join(map(str, ValueList)) + "'" + ')')

            Dbase(self.Lvl2tbl).TblEdit(UpQuery)

            self.data.append(values)

            self.data = Dbase().Dsqldata(self.DsqlLvl2)
            self.model = DataMods(self.Lvl2tbl, self.data)
            self.dvc.AssociateModel(self.model)
            self.dvc.Refresh

        else:
            wx.MessageBox('''A Category needs to be selected,\nbefore adding a
                           new row''', "Missing Information",
                          wx.OK | wx.ICON_INFORMATION)

    def OnAddcmb1(self, evt):
        CmbLst1(self, self.frg_tbl)
        self.ReFillList(0, self.frg_tbl)

    def ReFillList(self, combo, tbl):
        self.lc = self.lctrls[combo]
        self.lc.DeleteAllItems()

        index = 0
        ReFillQuery = 'SELECT * FROM "' + tbl + '"'
        for values in Dbase().Dsqldata(ReFillQuery):
            col = 0
            for value in values:
                if col == 0:
                    self.lc.InsertItem(index, str(value))
                else:
                    self.lc.SetItem(index, col, str(value))
                col += 1
            index += 1

    def OnSelectDone(self, evt):
        # add note(s) selected from upper table to lower table
        # rowID is table index value
        NoteIds = []
        items = self.dvc.GetSelections()
        for item in items:
            row = self.model.GetRow(item)
            rowID = self.model.GetValueByRow(row, self.pk_col)
            if int(rowID) in [self.data1[i][0]
                              for i in range(len(self.data1))]:
                continue
            NoteIds.append(rowID)

        if self.NoteStr is None:
            self.NoteStr = str(tuple(NoteIds))
        else:
            self.NoteStr = self.NoteStr[:-1] + ', ' + str(tuple(NoteIds))[1:]

        self.NoteStr = "".join(self.NoteStr.split())
        if len(NoteIds) <= 1:
            self.NoteStr = self.NoteStr[:-2] + ')'

        self.ChgSpecID(self.NoteStr)

        self.Dsql1 = (self.DsqlLvl2 + ' WHERE ' + self.pk_Name +
                      ' IN ' + self.NoteStr)
        self.data1 = Dbase().Dsqldata(self.Dsql1)
        self.model1 = DataMods(self.Lvl2tbl, self.data1)
        self.dvc1.AssociateModel(self.model1)
        self.dvc1.Refresh

    def ChgSpecID(self, ID=None):
        UpQuery = ('UPDATE PipeSpecification SET ' + self.Lvl2tbl +
                   ' =? WHERE Commodity_Property_ID = ' + str(self.ComdPrtyID))
        Dbase().TblEdit(UpQuery, [ID])

    def OnRmvNote(self, evt):
        # remove selected note(s) from Commodity property
        # rowID is table index value
        NoteIds = []
        items = self.dvc1.GetSelections()
        for item in items:
            row = self.model1.GetRow(item)
            rowID = self.model1.GetValueByRow(row, self.pk_col)
            NoteIds.append(rowID)

        NoteLst = self.ConvertStr_Lst(self.NoteStr)

        for i in NoteIds:
            NoteLst.remove(i)

        self.NoteStr = str(tuple(NoteLst))
        if self.NoteStr[-2] == ',':
            self.NoteStr = self.NoteStr[:-2] + ')'
        elif self.NoteStr == '()':
            self.NoteStr = None

        self.ChgSpecID(self.NoteStr)

        if self.NoteStr is not None:
            self.Dsql1 = (self.DsqlLvl2 + ' WHERE ' + self.pk_Name +
                          ' IN ' + self.NoteStr)
            self.data1 = Dbase().Dsqldata(self.Dsql1)
        else:
            self.data1 = []
        self.model1 = DataMods(self.Lvl2tbl, self.data1)
        self.dvc1.AssociateModel(self.model1)
        self.dvc1.Refresh

    def ConvertStr_Lst(self, vals):
        lst = vals.split("'")
        newlst = []
        for i in range(len(lst)):
            if lst[i].isdigit():
                newlst.append(lst[i])
        return newlst

    def OnSrchDVC1(self, evt):
        # collect feign table info
        frgn_info = Dbase().Dtbldata(self.Lvl2tbl)
        field = frgn_info[0][4]
        frg_tbl1 = frgn_info[0][2]

        # do search of string value from combobox
        # equal to value in the self.frg_tbl
        Shcol = Dbase().Dcolinfo(frg_tbl1)[self.showcol][1]
        ShQry = ("SELECT " + field + " FROM " + frg_tbl1 + " WHERE " +
                 Shcol + " LIKE '%" + self.cmb10.GetValue() +
                 "%' COLLATE NOCASE")
        ShQryVal = str(Dbase().Dsqldata(ShQry)[0][0])
        # append the found frgn_fld to the original data grid SQL and
        # find only records equal to the combo selection
        ShQry = (self.Dsql1 + ' AND ' + frg_tbl1 + '.' +
                 field + ' = "' + ShQryVal + '"')
        OSdata = Dbase().Search(ShQry)
        # if nothing is found show blank grid
        if OSdata is False:
            OSdata = []
        self.model1 = DataMods(self.Lvl2tbl, OSdata)
        self.dvc1.AssociateModel(self.model1)
        self.dvc1.Refresh

    def OnRstrDVC1(self, evt):
        ORdata1 = Dbase().Restore(self.Dsql1)
        self.cmb10.ChangeValue('')
        self.cmb10.SetHint(self.frg_fld)
        self.model1 = DataMods(self.Lvl2tbl, ORdata1)
        self.dvc1._AssociateModel(self.model1)
        self.dvc1.Refresh

    def OnDeleteRow(self, evt):
        item = self.dvc.GetSelection()
        row = self.model.GetRow(item)
        self.b2.Disable()
        self.model.DeleteRow(row, self.pk_Name)
        self.data1 = Dbase().Restore(self.Dsql1)
        self.model1 = DataMods(self.Lvl2tbl, self.data1)
        self.dvc1._AssociateModel(self.model1)
        self.dvc1.Refresh

    def OnAddRow(self, evt):
        self.AddTblRow()

    def OnValueChanged(self, evt):
        self.TblCellEdit()
        evt.Skip()

    def OnSearch(self, evt):
        # collect feign table info
        frgn_info = Dbase().Dtbldata(self.Lvl2tbl)
        field = frgn_info[0][4]
        self.frg_tbl = frgn_info[0][2]

        # do search of string value from combobox equal
        # to value in the self.frg_tbl
        Shcol = Dbase().Dcolinfo(self.frg_tbl)[self.showcol][1]
        ShQuery = ('SELECT ' + field + ' FROM ' + self.frg_tbl +
                   " WHERE " + Shcol + " LIKE '%" + self.cmb1.GetValue()
                   + "%' COLLATE NOCASE")
        ShQueryVal = str(Dbase().Dsqldata(ShQuery)[0][0])
        # append the found frgn_fld to the original data grid SQL and
        # find only records equal to the combo selection
        ShQuery = (self.DsqlLvl2 + ' WHERE ' + self.frg_tbl + '.' +
                   field + ' = "' + ShQueryVal + '"')

        OSdata = Dbase().Search(ShQuery)
        # if nothing is found show blank grid
        if OSdata is False:
            OSdata = []
        self.model = DataMods(self.Lvl2tbl, OSdata)
        self.dvc.AssociateModel(self.model)
        self.dvc.Refresh
        self.b2.Enable()

    def OnSelect(self, evt):
        self.b2.Enable()
        txt = ('''To complete adding a new record click "Add Row".\nThen
                edit data by double click on the cell.''')
        self.editlbl.SetLabel(txt)
        self.editlbl.SetForegroundColour((255, 0, 0))
        self.Sizer.Layout()

    def OnRestore(self, evt):
        self.ORdata = Dbase().Restore(self.DsqlLvl2)
        self.cmb1.ChangeValue('')
        self.cmb1.SetHint(self.frg_fld)
        self.model = DataMods(self.Lvl2tbl, self.ORdata)
        self.dvc._AssociateModel(self.model)
        self.dvc.Refresh

    def OnClose(self, evt):
        self.GetParent().Enable(True)
        self.__eventLoop.Exit()
        self.Destroy()


class SupportFrms(wx.Frame):
    def __init__(self, parent):

        super(SupportFrms, self).__init__(parent,
                                          title='Direct Table Modification',
                                          size=(870, 700))

        self.parent = parent   # add for child parent form

        self.Bind(wx.EVT_CLOSE, self.OnClose)
        self.InitUI()

    def InitUI(self):
        # the table names are modified to become
        # the button name
        TableList1 = [
            'None', 'ANSI_Rating', 'BodyType', 'BoltMaterial', 'BonnetType',
            'ButterflyBodyType', 'CommodityCodes', 'CommodityEnds',
            'CorrosionAllowance', 'EndConnects', 'FlangeFace',
            'FlangeStyle', 'FluidCategory', 'ForgeClass', 'InsulationAdhesive',
            'InsulationClass', 'InsulationJacket', 'InsulationMaterial',
            'InsulationSealer', 'InsulationThickness', 'MaterialType',
            'NoteCategory', 'NutMaterial', 'OLetStyle', 'OLetWt',
            'PaintColors', 'PaintPrep', 'PipeDimensions',
            'PipeMaterialSpec', 'Pipe_OD', 'PipeSchedule',
            'Porting', 'SleeveMaterial', 'SpecialCase', 'SpecialItems',
            'TrvlShtTime', 'TubeMaterial', 'TubeSize', 'TubeValveMatr',
            'TubeWall', 'UnionSeats', 'ValveEnds', 'WedgeType',
            'WeldFiller', 'WeldFillerGroup', 'WeldMaterialGroup',
            'WeldProcessList', 'WeldProcesses',
            'WeldQualifyPosition', 'WeldQualifyThickness']

        rdBox1Choices = []

        for tblName1 in TableList1:
            if tblName1.find("_") != -1:
                btnName = (tblName1.replace("_", " "))
            else:
                btnName = (' '.join(re.findall('([A-Z][a-z]*)', tblName1)))
            rdBox1Choices.append(btnName)

        # the new button name is linked in a dictionary to the tablename
        self.tbl_bt1 = dict(zip(rdBox1Choices, TableList1))

        TableList2 = [
            'None', 'ButterflyBodyMaterial', 'ButterflyDiscMaterial',
            'ButterflySeatMaterial', 'ButterflyShaftMaterial',
            'ButtWeldMaterial', 'ChemMatr', 'ForgedMaterial',
            'InspectionTravelSheet', 'MaterialGrade', 'PackingMaterial',
            'PipeMaterial', 'PressTempTables', 'SeatMaterial',
            'SpringMaterial', 'StemMaterial', 'ValveBodyMaterial',
            'WedgeMaterial', 'WeldProcedures', 'WeldRequirements']

        rdBox2Choices = []

        for tblName2 in TableList2:
            if tblName2.find("_") != -1:
                btnName = (tblName2.replace("_", " "))
            else:
                btnName = (' '.join(re.findall('([A-Z][a-z]*)', tblName2)))
            rdBox2Choices.append(btnName)

        self.tbl_bt2 = dict(zip(rdBox2Choices, TableList2))

        TableList3 = ['None', 'BallValve', 'ButterflyValve', 'GateValve',
                      'GlobeValve', 'PlugValve', 'PistonCheckValve',
                      'SwingCheckValve']

        rdBox3Choices = []

        for tblName3 in TableList3:
            if tblName3.find("_") != -1:
                btnName = (tblName3.replace("_", " "))
            else:
                btnName = (' '.join(re.findall('([A-Z][a-z]*)', tblName3)))
            rdBox3Choices.append(btnName)

        self.tbl_bt3 = dict(zip(rdBox3Choices, TableList3))

        TableList4 = ['None', 'NewPipeSpec', 'NonConformance', 'MtrSubRcrd']

        rdBox4Choices = []

        for tblName4 in TableList4:
            if tblName4.find("_") != -1:
                btnName = (tblName4.replace("_", " "))
            else:
                btnName = (' '.join(re.findall('([A-Z][a-z]*)', tblName4)))
            rdBox4Choices.append(btnName)

        self.tbl_bt4 = dict(zip(rdBox4Choices, TableList4))

        self.Sizer = wx.BoxSizer(wx.VERTICAL)

        self.rb1 = wx.CheckBox(self, 11, label='Pipe Material Spec',
                               style=wx.RB_GROUP)
        self.Bind(wx.EVT_CHECKBOX, self.OnMtrSpec)

        bxSzr1 = wx.BoxSizer(wx.VERTICAL)
        self.rdBox1 = wx.RadioBox(self, wx.ID_ANY, "", (5, 5),
                                  wx.DefaultSize, rdBox1Choices,
                                  5, wx.RA_SPECIFY_COLS)
        self.rdBox1.SetSelection(0)
        self.rdBox1.Bind(wx.EVT_RADIOBOX, self.OnRadioBox1)
        bxSzr1.Add(self.rdBox1, 0, wx.ALL, 5)

        bxSzr2 = wx.BoxSizer(wx.VERTICAL)
        self.rdBox2 = wx.RadioBox(self, wx.ID_ANY, "", (5, 5),
                                  wx.DefaultSize, rdBox2Choices,
                                  4, wx.RA_SPECIFY_COLS)
        self.rdBox2.SetSelection(0)
        self.rdBox2.Bind(wx.EVT_RADIOBOX, self.OnRadioBox2)
        bxSzr2.Add(self.rdBox2, 0, wx.ALL, 5)

        bxSzr3 = wx.BoxSizer(wx.HORIZONTAL)
        self.rdBox3 = wx.RadioBox(self, wx.ID_ANY, "", (5, 5),
                                  wx.DefaultSize, rdBox3Choices, 4,
                                  wx.RA_SPECIFY_COLS)
        self.rdBox3.SetSelection(0)
        self.rdBox3.Bind(wx.EVT_RADIOBOX, self.OnRadioBox3)
        bxSzr3.Add(self.rdBox3, 0, wx.ALL, 5)

        bxSzr4 = wx.BoxSizer(wx.HORIZONTAL)
        self.rdBox4 = wx.RadioBox(self, wx.ID_ANY, "", (5, 5),
                                  wx.DefaultSize, rdBox4Choices, 4,
                                  wx.RA_SPECIFY_COLS)
        self.rdBox4.SetSelection(0)
        self.rdBox4.Bind(wx.EVT_RADIOBOX, self.OnRadioBox4)
        bxSzr4.Add(self.rdBox4, 0, wx.ALL, 5)

        bxSzrBt = wx.BoxSizer(wx.HORIZONTAL)
        self.b4 = wx.Button(self, label="Exit")
        self.b4.SetForegroundColour('red')
        self.Bind(wx.EVT_BUTTON, self.OnClose, self.b4)

        self.btnbox = wx.BoxSizer(wx.VERTICAL)
        self.btnbox.Add(self.b4, 0, wx.ALL | wx.ALIGN_CENTER, 5)

        bxSzrBt.Add(self.btnbox, 1, wx.ALIGN_BOTTOM, 25)

        self.Sizer.Add(self.rb1, 1, wx.ALL, 5)
        self.Sizer.Add(bxSzr1, 1, wx.ALL, 5)
        self.Sizer.Add(bxSzr2, 1, wx.ALL, 5)
        self.Sizer.Add(bxSzr3, 1, wx.ALL, 5)
        self.Sizer.Add(bxSzr4, 1, wx.ALL, 5)

        self.Sizer.Add(bxSzrBt, 1, wx.ALIGN_RIGHT | wx.ALL, 5)

        # add these 5 following lines to child parent form
        self.CenterOnParent()
        self.GetParent().Enable(False)
        self.Show(True)
        self.__eventLoop = wx.GUIEventLoop()
        self.__eventLoop.Run()

    def OnRadioBox1(self, evt):
        # use the selected raio button to specify the corresponding
        # table/form name to open
        if self.rdBox1.GetStringSelection() != 'None':
            tablename = (self.tbl_bt1.get(self.rdBox1.GetStringSelection()))
            CmbLst1(self, tablename)

    def OnRadioBox2(self, evt):
        if self.rdBox2.GetStringSelection() != 'None':
            tablename = (self.tbl_bt2.get(self.rdBox2.GetStringSelection()))
            if tablename in ['PipeMaterial', 'ButtWeldMaterial',
                             'ForgedMaterial']:
                BldMtr(self, tablename)
            elif tablename == 'WeldRequirements':
                BldWeld(self, tablename)
            elif tablename == 'WeldProcedures':
                BldProced(self, tablename)
            else:
                CmbLst2(self, tablename)

    def OnRadioBox3(self, evt):
        if self.rdBox3.GetStringSelection() != 'None':
            tablename = (self.tbl_bt3.get(self.rdBox3.GetStringSelection()))
            BldValve(self, tablename)

    def OnRadioBox4(self, evt):
        # use the selected raio button to specify the corresponding
        # table/form name to open
        if self.rdBox4.GetStringSelection() != 'None':
            tablename = (self.tbl_bt4.get(self.rdBox4.GetStringSelection()))
            if tablename == 'NewPipeSpec':
                RPSImport(self)
            elif tablename == 'NonConformance':
                NCRImport(self)
            elif tablename == 'MtrSubRcrd':
                MSRImport(self)

    def OnMtrSpec(self, evt):
        PipeMtrSpc(self, 'PipeMaterialSpec')
        self.rb1.SetValue(0)

    def OnClose(self, evt):
        # add following 2 lines for child parent
        self.GetParent().Enable(True)
        self.__eventLoop.Exit()
        self.Destroy()


class BldMtr(wx.Frame):
    '''Routine to build form and populate grid'''
    def __init__(self, parent, tblname, ComdPrtyID=None, model=None):

        self.parent = parent
        self.model = model
        # set up the table column names, width and if column can be
        # edited ie primary autoincrement
        self.tblname = tblname
        # commodity property ID linked to pipespecification
        self.ComdPrtyID = ComdPrtyID

        if self.tblname.find("_") != -1:
            frmtitle = (self.tblname.replace("_", " "))
        else:
            frmtitle = (' '.join(re.findall('([A-Z][a-z]*)', self.tblname)))

        super(BldMtr, self).__init__(parent, title=frmtitle, size=(680, 500))

        self.Bind(wx.EVT_CLOSE, self.OnClose)

        self.ComCode = ''
        self.PipeMtrSpec = ''
        self.frg_tbl = 'PipeMaterialSpec'
        self.edit_data = []

        if self.ComdPrtyID is not None:
            query = ('''SELECT Commodity_Code, Pipe_Material_Code,
                      End_Connection FROM CommodityProperties WHERE
                      CommodityPropertyID = ''' + str(self.ComdPrtyID))
            dataset = Dbase().Dsqldata(query)[0]
            self.PipeMtrSpec = dataset[1]
            self.ComCode = dataset[0]

        grd_tbl_data = Dbase().Dtbldata(self.tblname)

        link_fld = []
        frgn_tbl = []
        frgn_key = []
        frgn_col = []
        link_fld_col = []

        n = 0
        for item in grd_tbl_data:
            link_fld.append(grd_tbl_data[n][3])
            frgn_tbl.append(grd_tbl_data[n][2])
            frgn_key.append(grd_tbl_data[n][4])
            frgn_col.append(Dbase().Dcolinfo(frgn_tbl[n]))
            n += 1

        info = Dbase().Dcolinfo(self.tblname)

        # here we get the information needed for the SQL from the Lvl2 table
        # and determine report column width based either
        # on data type or specified field size
        n = 0
        for item in info:
            for fld in link_fld:
                if item[1] == fld:
                    # need to make frgn_fld column noneditable in DVC
                    link_fld_col.append(n)
            n += 1

        tblinfo = []
        # specify which listbox column to display in the combobox
        self.showcol = int

        tblinfo = Dbase(self.tblname).Fld_Size_Type()
        self.ID_col = tblinfo[1]
        #  colms = tblinfo[3]
        self.pkcol_name = tblinfo[0]
        self.autoincrement = tblinfo[2]
        columnames = Dbase(self.tblname).ColNames()

        if self.PipeMtrSpec:
            query = ('''SELECT Pipe_Material_Spec FROM PipeMaterialSpec
                      WHERE Material_Spec_ID = "''' +
                     str(self.PipeMtrSpec) + '"')
            self.spec = Dbase().Dsqldata(query)[0][0]
            mtrtyp = self.spec[1]
            mtrgrd = self.spec[2]
            self.GrdSql = ('SELECT * FROM ' + self.tblname +
                           ' WHERE Material_Type = "' + mtrtyp +
                           '" AND Material_Grade = "' + mtrgrd + '"')
            self.data = Dbase().Dsqldata(self.GrdSql)

            query = ('''SELECT Material_Type FROM MaterialType WHERE
                      Type_Designation = "''' + mtrtyp + '"')
            MtrTyp = Dbase().Dsqldata(query)[0][0]
            strg2 = 'Material Type = '
            self.text1 = (wx.StaticText(
                self, size=(500, 30), label=strg2 + MtrTyp,
                style=wx.ST_NO_AUTORESIZE))

            query = ('''SELECT Material_Grade FROM MaterialGrade
                      WHERE Type_Designation = "''' +
                     mtrtyp + '" AND Material_Grade_Designation = "'
                     + mtrgrd + '"')
            MtrGrd = Dbase().Dsqldata(query)[0][0]
            strg3 = 'Material Grade = '
            self.text2 = (wx.StaticText(self, size=(500, 30),
                          label=strg3 + MtrGrd, style=wx.ST_NO_AUTORESIZE))

        else:
            self.GrdSql = 'SELECT * FROM ' + self.tblname
            self.data = Dbase().Dsqldata(self.GrdSql)
            self.text1 = wx.StaticText(
                self, size=(500, 30), label='Material Type',
                style=wx.ST_NO_AUTORESIZE)
            self.text2 = wx.StaticText(
                self, size=(500, 30), label='Material Grade',
                style=wx.ST_NO_AUTORESIZE)

        # Create a dataview control
        self.dvc = dv.DataViewCtrl(self, wx.ID_ANY, wx.DefaultPosition,
                                   wx.Size(660, 200),
                                   style=wx.BORDER_THEME
                                   | dv.DV_ROW_LINES
                                   | dv.DV_VERT_RULES
                                   | dv.DV_HORIZ_RULES
                                   | dv.DV_SINGLE
                                   )

        # Bind some events so we can see what the DVC sends us
        self.Bind(dv.EVT_DATAVIEW_ITEM_VALUE_CHANGED,
                  self.OnValueChanged, self.dvc)
        # use the sorted data to load the dataviewlistcontrol
        if self.model is None:
            self.model = DataMods(self.tblname, self.data, self.edit_data)
        self.dvc.AssociateModel(self.model)

        n = 0
        # formwidth = 0
        for colname in columnames:
            if n == self.ID_col and self.autoincrement or (n in link_fld_col):
                col_mode = dv.DATAVIEW_CELL_INERT
            else:
                col_mode = dv.DATAVIEW_CELL_EDITABLE

            self.dvc.AppendTextColumn(colname, n,
                                      width=wx.LIST_AUTOSIZE_USEHEADER,
                                      mode=col_mode)
            n += 1

        # make columns not sortable and but reorderable.
        n = 0
        for c in self.dvc.Columns:
            c.Sortable = True
            if n == 0 and self.autoincrement != 0:
                c.Sortable = False
            c.Reorderable = True
            c.Resizeable = True
            n += 1

        # change to not let the ID col be moved.
        self.dvc.Columns[(self.ID_col)].Reorderable = False
        self.dvc.Columns[(self.ID_col)].Resizeable = False
        # if autoincrement is false then the data can be sorted based on ID_col
        if self.autoincrement == 0:
            self.data.sort(key=lambda tup: tup[self.ID_col])

        # set the Sizer property (same as SetSizer)
        self.Sizer = wx.BoxSizer(wx.VERTICAL)

        self.dvcbox = wx.BoxSizer(wx.HORIZONTAL)
        self.dvcbox.Add(self.dvc, 1, wx.ALL | wx.ALIGN_CENTER, 5)

        self.addlbl = wx.StaticText(self, -1, style=wx.ALIGN_CENTER_HORIZONTAL)
        txt = ('''To add a new record first select
    a Pipe Material Specification''')
        self.addlbl.SetLabel(txt)
        self.addlbl.SetForegroundColour((255, 0, 0))

        self.cmbsizer = wx.BoxSizer(wx.HORIZONTAL)
        # develope the comboctrl and attach popup list
        self.cmb1 = wx.ComboCtrl(self, size=(100, -1))
        self.Bind(wx.EVT_TEXT, self.OnSelect, self.cmb1)
        self.showcol = 1
        self.cmb1.SetPopupControl(ListCtrlComboPopup(
            self.frg_tbl, showcol=self.showcol))
        if self.PipeMtrSpec:
            self.cmb1.ChangeValue(self.spec)

        # add a button to call main form to search combo list data
        self.b6 = wx.Button(self, label="Restore Data")
        self.Bind(wx.EVT_BUTTON, self.OnRestore, self.b6)

        self.cmbsizer.Add(self.b6, 0, wx.ALL, 10)
        self.cmbsizer.Add(self.cmb1, 0, wx.ALIGN_CENTER, 5)

        self.lblsizer = wx.BoxSizer(wx.VERTICAL)
        self.collbl = wx.StaticText(self, -1, style=wx.ALIGN_CENTER_HORIZONTAL)
        txt = 'Columns can be sorted by clicking title!'
        self.collbl.SetLabel(txt)
        self.collbl.SetForegroundColour((255, 0, 0))
        self.lblsizer.Add(self.text1, 0, wx.ALIGN_RIGHT, 5)
        self.lblsizer.Add(self.text2, 0)  # , wx.ALIGN_BOTTOM, 5)
        self.lblsizer.Add(self.collbl, 0)  # , wx.ALIGN_BOTTOM, 5)

        # Add buttons for grid modifications
        self.b1 = wx.Button(self, label="Print Report")
        self.Bind(wx.EVT_BUTTON, self.PrintFile, self.b1)
        self.b2 = wx.Button(self, label="Add Row")
        self.Bind(wx.EVT_BUTTON, self.OnAddRow, self.b2)
        self.b3 = wx.Button(self, label="Delete Row")
        self.Bind(wx.EVT_BUTTON, self.OnDeleteRow, self.b3)
        self.b4 = wx.Button(self, label="Exit")
        self.Bind(wx.EVT_BUTTON, self.OnClose, self.b4)

        # add a button box and place the buttons
        self.btnbox = wx.BoxSizer(wx.HORIZONTAL)
        self.btnbox.Add(self.b1, 0, wx.ALL | wx.ALIGN_CENTER, 5)
        self.btnbox.Add(self.b2, 0, wx.ALL | wx.ALIGN_CENTER, 5)
        self.btnbox.Add(self.b3, 0, wx.ALL | wx.ALIGN_CENTER, 5)
        self.btnbox.Add(self.b4, 0, wx.ALL | wx.ALIGN_CENTER, 5)

        # need to disable the add row button
        # until an item is selected in the combo box
        self.b2.Disable()

        # add static label to explain how to add / edit data
        self.editlbl = wx.StaticText(self, -1, style=wx.ALIGN_CENTER)
        txt = (
            '''To edit the forge material description double click on the cell.
        NOTE: changes will be reflected in all related records''')
        self.editlbl.SetLabel(txt)
        self.editlbl.SetForegroundColour((255, 0, 0))

        self.Sizer.Add((5, 10))
        self.Sizer.Add(self.addlbl, 0, wx.ALIGN_CENTER_HORIZONTAL)
        self.Sizer.Add((5, 10))
        self.Sizer.Add(self.cmbsizer, 0, wx.ALIGN_CENTER)
        self.Sizer.Add((5, 10))
        self.Sizer.Add(self.lblsizer, 0, wx.ALIGN_CENTER)
        self.Sizer.Add(self.dvcbox, 1, wx.ALIGN_CENTER)
        self.Sizer.Add(self.editlbl, 0, wx.ALIGN_CENTER)
        self.Sizer.Add((5, 10))
        self.Sizer.Add(self.btnbox, 0, wx.ALL | wx.ALIGN_CENTER, 10)

        # add these following lines to child form
        self.CenterOnParent()
        self.GetParent().Enable(False)
        self.Show(True)
        self.__eventLoop = wx.GUIEventLoop()
        self.__eventLoop.Run()

    def cmbSelect(self):
        mtrtyp = self.cmb1.GetValue()[1]
        mtrgrd = self.cmb1.GetValue()[2]
        query = ('SELECT * FROM ' + self.tblname + ' WHERE Material_Type = "'
                 + mtrtyp + '" AND Material_Grade = "' + mtrgrd + '"')
        self.data = Dbase().Search(query)

        # if nothing is found show blank grid
        if self.data is False:
            self.data = []
        else:
            query = ('''SELECT Material_Type FROM MaterialType WHERE
                      Type_Designation = "''' + mtrtyp + '"')
            MtrTyp = Dbase().Dsqldata(query)[0][0]
            strg1 = 'Material Type = '
            self.text1.SetLabel(strg1 + MtrTyp)

            query = ('''SELECT Material_Grade FROM MaterialGrade WHERE
                      Type_Designation = "''' + mtrtyp +
                     '" AND Material_Grade_Designation = "' + mtrgrd + '"')
            MtrGrd = Dbase().Dsqldata(query)[0][0]
            strg2 = 'Material Grade = '
            self.text2.SetLabel(strg2 + MtrGrd)

        self.model = DataMods(self.tblname, self.data)
        self.dvc.AssociateModel(self.model)
        self.dvc.Refresh

    def TblCellEdit(self):
        # rowID is table index value, colID is column number of table index
        # new_value is the link foreign field value
        # colChgNum is the column number which has edited cell,
        # values is the new edited value(s)
        self.edit_data = self.model.edit_data
        rowID = self.model.GetValueByRow(self.edit_data[0], self.ID_col)
        self.edit_data.append(rowID)

        realnames = []
        for item in Dbase().Dcolinfo(self.tblname):
            realnames.append(item[1])

        enable_b2 = 0

        # if just a cell is being edited then do the following
        colID = self.ID_col
        colChgNum = self.edit_data[1]
        values = self.edit_data[2]
        rowID = self.edit_data[3]
        colIDName = realnames[colID]
        colChgName = realnames[colChgNum]

        if values.find('"') != -1:
            values = values.replace('"', '""')
        if type(rowID) != str:
            UpQuery = ('UPDATE ' + self.tblname + ' SET ' + colChgName + ' = "'
                       + str(values) + '" WHERE ' + str(colIDName) + ' = '
                       + str(rowID))
            check_query = ('SELECT * FROM ' + self.tblname + ' WHERE '
                           + str(colIDName) + ' = ' + str(rowID))
        else:
            UpQuery = ('UPDATE ' + self.tblname + ' SET ' + colChgName + ' = "'
                       + str(values) + '" WHERE ' + str(colIDName) + ' = "'
                       + rowID + '"')

        # if the database ID col has been changed then base the
        # query on the new ID value
        # if anyother column was changed then base the query
        # on the original colID
            if colID == colChgNum:
                check_query = ('SELECT * FROM ' + self.tblname + ' WHERE ' +
                               str(colIDName) + ' = "' + values + '"')
            else:
                check_query = ('SELECT * FROM ' + self.tblname + ' WHERE ' +
                               str(colIDName) + ' = "' + rowID + '"')

        Dbase().TblEdit(UpQuery)

        # check to see if added row has been properly edited if
        # so enable the addrow button
        data_string = Dbase().Dsqldata(check_query)
        if 'REQUIRED' not in data_string[0]:
            enable_b2 = 1

        data = Dbase().Dsqldata(self.GrdSql)

        if enable_b2:
            self.b2.Disable()
        # if autoincrement is false then the data can be sorted based on ID_col
            if self.autoincrement == 0:
                data.sort(key=lambda tup: tup[self.ID_col])

        txt = ('''To edit the forge material description double click on the cell.
        NOTE: changes will be reflected in all related records''')
        self.editlbl.SetLabel(txt)

    def AddTblRow(self):
        # this returns the text value in the combo box not the linked ID value
        values = [0]*4

        new_value = self.cmb1.GetValue()

        if new_value:
            realnames = []
            for item in Dbase().Dcolinfo(self.tblname):
                realnames.append(item[1])

            self.b2.Disable()

            # if the table index is auto increment then assign next
            # value otherwise do nothing
            New_ID = cursr.execute(
                "SELECT MAX(" + self.pkcol_name + ") FROM "
                + self.tblname).fetchone()
            if New_ID[0] is None:
                Max_ID = '1'
            else:
                Max_ID = str(New_ID[0]+1)
            values[self.ID_col] = Max_ID

            # update data structure
            values[1] = 'REQUIRED'
            values[2] = new_value[1]
            values[3] = new_value[2]

            # if a full row is being added then do the following
            ValueList = []
            if type(values) == str:
                ValueList.append(values)
            else:
                ValueList = values

            UpQuery = ('INSERT INTO ' + self.tblname + ' VALUES (' +
                       "'" + "','".join(map(str, ValueList)) + "'" + ')')
            Dbase(self.tblname).TblEdit(UpQuery)

            self.data.append(values)

            self.data = Dbase().Dsqldata(self.GrdSql)
            self.model = DataMods(self.tblname, self.data)
            self.dvc.AssociateModel(self.model)
            self.dvc.Refresh

            txt = ('''To complete the addition of the new record
                double the click cell to modify description''')
            self.editlbl.SetLabel(txt)

        else:
            wx.MessageBox('''A material type needs to be selected
                          ,\nbefore adding a new row"''',
                          "Missing Information",
                          wx.OK | wx.ICON_INFORMATION)

    def PrintFile(self, evt):
        import Report_Lvl2

        saveDialog = wx.FileDialog(self, message='Save Report as PDF.',
                                   wildcard='PDF (*.pdf)|*.pdf',
                                   style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT)

        if saveDialog.ShowModal() == wx.ID_CANCEL:
            filename = ''

        filename = saveDialog.GetPath()
        if filename.find(".pdf") == -1:
            filename = filename + '.pdf'

        saveDialog.Destroy()

        rptdata = self.data

        ColWdth = [6, 30, 10, 10]

        if self.tblname == 'ForgedMaterial':
            colnames = ['ID', 'Forged\nMaterial', 'Material\nType',
                        'Material\nGrade']

        elif self.tblname == 'ButtWeldMaterial':
            colnames = ['ID', 'Butt Weld\nMaterial', 'Material\nType',
                        'Material\nGrade']

        elif self.tblname == 'PipeMaterial':
            colnames = ['ID', 'Pipe\nMaterial', 'Material\nType',
                        'Material\nGrade']

        Report_Lvl2.Report(self.tblname, rptdata, colnames,
                           ColWdth, filename).create_pdf()

    def OnDeleteRow(self, evt):
        item = self.dvc.GetSelection()
        row = self.model.GetRow(item)
        self.b2.Disable()
        self.model.DeleteRow(row, self.pkcol_name)

    def OnAddRow(self, evt):
        self.AddTblRow()

    def OnValueChanged(self, evt):
        self.TblCellEdit()

    def OnSelect(self, evt):
        self.b2.Enable()
        txt = ('''To complete adding a new record click "Add Row".
            Then edit data by double click on the cell.''')
        self.editlbl.SetLabel(txt)
        self.editlbl.SetForegroundColour((255, 0, 0))
        self.Sizer.Layout()
        self.cmbSelect()

    def OnRestore(self, evt):
        self.cmb1.ChangeValue('')
        self.GrdSql = 'SELECT * FROM ' + self.tblname
        self.data = Dbase().Dsqldata(self.GrdSql)
        self.model = DataMods(self.tblname, self.data)
        self.dvc.AssociateModel(self.model)
        self.dvc.Refresh

    def OnClose(self, evt):
        self.GetParent().Enable(True)
        self.__eventLoop.Exit()
        self.Destroy()


class CmbLst1(wx.Frame):
    def __init__(self, parent, Lvl1tbl, model=None):
        '''Class to build the popup forms for the tables selected
        by the addbtn next to each combo box
        Routine to build form and populate grid
        set up the table column names, width and if column can be
        edited ie primary autoincrement'''

        self.parent = parent
        self.Lvl1tbl = Lvl1tbl
        self.model = model

        self.tblinfo = []
        self.tblinfo = Dbase(self.Lvl1tbl).Fld_Size_Type()
        self.ID_col = self.tblinfo[1]
        self.colms = self.tblinfo[3]
        self.pkcol_name = self.tblinfo[0]
        self.autoincrement = self.tblinfo[2]
        self.columnames = Dbase(self.Lvl1tbl).ColNames()

        if self.Lvl1tbl.find("_") != -1:
            frmtitle = (self.Lvl1tbl.replace("_", " "))
        else:
            frmtitle = (' '.join(re.findall('([A-Z][a-z]*)', self.Lvl1tbl)))

        frmwdth = 0
        for item in self.colms:
            frmwdth = item + frmwdth

        if frmwdth < 50:
            frmwdth = 100
        elif 50 <= frmwdth < 75:
            frmwdth = 150
        else:
            frmwdth = 180
        frmSize = (frmwdth*5, 350)

        super(CmbLst1, self).__init__(parent, title=frmtitle, size=frmSize)

        self.Bind(wx.EVT_CLOSE, self.Close)

        self.InitUI()

    def InitUI(self):

        self.MainSQL = 'SELECT * FROM ' + self.Lvl1tbl
        self.data = Dbase().Dsqldata(self.MainSQL)

        # Create a dataview control
        self.dvc = dv.DataViewCtrl(self, wx.ID_ANY, wx.DefaultPosition,
                                   wx.Size(500, 300),
                                   style=wx.BORDER_THEME
                                   | dv.DV_ROW_LINES
                                   | dv.DV_VERT_RULES
                                   | dv.DV_HORIZ_RULES
                                   | dv.DV_SINGLE
                                   )
        self.dvc.SetMinSize = (wx.Size(100, 200))
        self.dvc.SetMaxSize = (wx.Size(500, 400))

        # Bind some events so we can see what the DVC sends us
        self.Bind(dv.EVT_DATAVIEW_ITEM_VALUE_CHANGED, self.OnValueChanged,
                  self.dvc)

        # if autoincrement is false then the data can be sorted based on ID_col
        if self.autoincrement == 0:
            self.data.sort(key=lambda tup: tup[self.ID_col])

        # use the sorted data to load the dataviewlistcontrol
        if self.model is None:
            self.model = DataMods(self.Lvl1tbl, self.data)

        self.dvc.AssociateModel(self.model)

        n = 0
        for colname in self.columnames:
            colname = str(colname)

            if n == self.ID_col and self.autoincrement:
                col_mode = dv.DATAVIEW_CELL_INERT
            else:
                col_mode = dv.DATAVIEW_CELL_EDITABLE

            self.dvc.AppendTextColumn(colname, n,
                                      width=wx.LIST_AUTOSIZE_USEHEADER,
                                      mode=col_mode)
            n += 1

        # make columns not sortable and but reorderable.
        for c in self.dvc.Columns:
            c.Sortable = False
            c.Reorderable = True
            c.Resizeable = True

        # change to not let the ID col be moved.
        self.dvc.Columns[(self.ID_col)].Reorderable = False
        self.dvc.Columns[(self.ID_col)].Resizeable = False

        # set the Sizer property (same as SetSizer)
        self.Sizer = wx.BoxSizer(wx.VERTICAL)
        self.Sizer.Add(self.dvc, 1, wx.EXPAND)

        # Add some buttons
        self.b1 = wx.Button(self, label="Print Report")
        self.Bind(wx.EVT_BUTTON, self.PrintFile, self.b1)
        self.b2 = wx.Button(self, label="Add Row")
        self.Bind(wx.EVT_BUTTON, self.OnAddRow, self.b2)
        self.b3 = wx.Button(self, label="Delete Row")
        self.Bind(wx.EVT_BUTTON, self.OnDeleteRow, self.b3)
        self.b4 = wx.Button(self, label="Exit")
        self.Bind(wx.EVT_BUTTON, self.Close, self.b4)

        # add a button box and place the buttons
        self.btnbox = wx.BoxSizer(wx.HORIZONTAL)
        self.btnbox.Add(self.b1, 0, wx.LEFT | wx.ALIGN_CENTER, 5)
        self.btnbox.Add(self.b2, 0, wx.ALIGN_CENTER, 5)
        self.btnbox.Add(self.b3, 0, wx.ALIGN_CENTER, 5)
        self.btnbox.Add(self.b4, 0, wx.ALIGN_CENTER | wx.RIGHT, 5)

        # add static label to explain how to add / edit data
        self.editlbl = wx.StaticText(self, -1, style=wx.ALIGN_CENTER)
        txt = ('''To edit data double click on the cell.  To add a new
         record click "Add Row", then edit data.''')
        self.editlbl.SetLabel(txt)
        self.editlbl.SetForegroundColour((255, 0, 0))

        self.Sizer.Add(self.editlbl, 0, wx.ALIGN_CENTER, 5)

        self.Sizer.Add(self.btnbox, 0, wx.ALIGN_CENTER, 5)

        # add the next 5 lines for the child parent relation
        self.CenterOnParent()
        self.GetParent().Enable(False)
        self.Show(True)

        self.__eventLoop = wx.GUIEventLoop()
        self.__eventLoop.Run()

    def PrintFile(self, evt):
        import Report_Lvl1

        saveDialog = wx.FileDialog(
            self, message='Save Report as PDF.',
            wildcard='PDF (*.pdf)|*.pdf',
            style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT)

        if saveDialog.ShowModal() == wx.ID_CANCEL:
            filename = ''

        filename = saveDialog.GetPath()
        if filename.find(".pdf") == -1:
            filename = filename + '.pdf'

        saveDialog.Destroy()

        Report_Lvl1.Report(self.Lvl1tbl, self.data, self.columnames,
                           filename).create_pdf()

    def TblCellEdit(self):
        # rowID is table index value, colID is column number of table index
        # new_value is the link foreign field value
        # colChgNum is the column number which has edited cell,
        # values is the new edited value(s)

        self.edit_data = self.model.edit_data
        rowID = self.model.GetValueByRow(self.edit_data[0], self.ID_col)
        self.edit_data.append(rowID)

        realnames = []
        for item in Dbase().Dcolinfo(self.Lvl1tbl):
            realnames.append(item[1])

        enable_b2 = 0

        # if just a cell is being edited then do the following
        colID = self.ID_col
        colChgNum = self.edit_data[1]
        values = self.edit_data[2]
        rowID = int(self.edit_data[3])
        colIDName = realnames[colID]
        colChgName = realnames[colChgNum]

        if values.find('"') != -1:
            values = values.replace('"', '""')
        if type(rowID) != str:
            UpQuery = ('UPDATE ' + self.Lvl1tbl + ' SET ' + colChgName + ' = "'
                       + str(values) + '" WHERE ' + str(colIDName) +
                       ' = ' + str(rowID))
            check_query = ('SELECT * FROM ' + self.Lvl1tbl +
                           ' WHERE ' + str(colIDName) + ' = ' + str(rowID))
        else:
            UpQuery = ('UPDATE ' + self.Lvl1tbl + ' SET ' + colChgName + ' = "'
                       + str(values) + '" WHERE ' + str(colIDName) +
                       ' = "' + rowID + '"')

            # if the database ID col has been changed
            # then base the query on the new ID value
            # if anyother column was changed
            # then base the query on the original colID
            if colID == colChgNum:
                check_query = ('SELECT * FROM ' + self.Lvl1tbl +
                               ' WHERE ' + str(colIDName) +
                               ' = "' + values + '"')
            else:
                check_query = ('SELECT * FROM ' + self.Lvl1tbl + ' WHERE '
                               + str(colIDName) + ' = "' + rowID + '"')

        Dbase().TblEdit(UpQuery)

        # check to see if added row has been properly edited
        # if so enable the addrow button
        data_string = Dbase().Dsqldata(check_query)
        if 'Required' not in data_string[0]:
            enable_b2 = 1

        data = Dbase().Dsqldata(self.MainSQL)

        if enable_b2:
            # if autoincrement is false then the data
            # can be sorted based on ID_col
            if self.autoincrement == 0:
                data.sort(key=lambda tup: tup[self.ID_col])

    def AddTblRow(self):
        FldInfo = Dbase(self.Lvl1tbl).Fld_Size_Type()
        values = FldInfo[4]
        # if the table index is auto increment then assign
        # next value otherwise do nothing
        if FldInfo[2]:
            New_ID = cursr.execute(
                "SELECT MAX(" + self.pkcol_name +
                ") FROM " + self.Lvl1tbl).fetchone()
            if New_ID[0] is None:
                Max_ID = '1'
            else:
                Max_ID = str(New_ID[0]+1)
            values[self.ID_col] = Max_ID

        # if a full row is being added then do the following
        ValueList = []
        if type(values) == str:
            ValueList.append(values)
        else:
            ValueList = values
        # if the table index is auto increment then
        # assign next value otherwise do nothing
        for item in Dbase().Dcolinfo(self.Lvl1tbl):
            if item[5] == 1:
                IDcol = item[0]
                IDname = item[1]
                if 'INTEGER' in item[2]:
                    New_ID = cursr.execute(
                        "SELECT MAX(" + IDname +
                        ") FROM " + self.Lvl1tbl).fetchone()
                    if New_ID[0] is None:
                        Max_ID = '1'
                    else:
                        Max_ID = str(New_ID[0]+1)
                    if type(values) == str:
                        ValueList.insert(IDcol, Max_ID)

        UpQuery = ('INSERT INTO ' + self.Lvl1tbl +
                   ' VALUES (' + "'" + "','".join(
                       map(str, ValueList)) + "'" + ')')
        Dbase().TblEdit(UpQuery)

        self.data.append(values)
        self.model = DataMods(self.Lvl1tbl, self.data)
        self.dvc.AssociateModel(self.model)
        self.dvc.Refresh

    def OnDeleteRow(self, evt):
        item = self.dvc.GetSelection()
        row = self.model.GetRow(item)
        self.model.DeleteRow(row, self.pkcol_name)

    def OnAddRow(self, evt):
        self.AddTblRow()

    def OnValueChanged(self, evt):
        self.TblCellEdit()
        self.b2.Enable()

    def Close(self, evt):
        # add next 2 lines for child parent relation
        self.GetParent().Enable(True)
        self.__eventLoop.Exit()
        self.Destroy()


class CmbLst2(wx.Frame):
    '''Routine to build form and populate grid'''
    def __init__(self, parent, LVl2tbl, model=None):

        self.parent = parent
        self.model = model
        # set up the table column names, width and if column can be
        # edited ie primary autoincrement
        self.Lvl2tbl = LVl2tbl
        self.edit_data = []

        grd_tbl_data = Dbase().Dtbldata(self.Lvl2tbl)
        if grd_tbl_data != []:
            link_fld = grd_tbl_data[0][3]
            self.frg_tbl = grd_tbl_data[0][2]
            self.frg_fld = grd_tbl_data[0][4]
            frgn_col = Dbase().Dcolinfo(self.frg_tbl)[1][1]
        realnames = []
        datatable_str = ''
        self.RptColWdth = []

        # here we get the information needed in the report and
        # for the SQL from the Lvl2 table and determine report
        # column width based either on data type or specified field size
        n = 0
        ColWdth = []
        for item in Dbase().Dcolinfo(self.Lvl2tbl):
            col_wdth = ''
            # check to see if field length is specified if so
            # use it to set grid col width
            for s in re.findall(r'\d+', item[2]):
                if s.isdigit():
                    col_wdth = int(s)
                    ColWdth.append(col_wdth)
            realnames.append(item[1])
            if item[5] == 1:
                self.pk_Name = item[1]
                self.pk_col = n
                if 'INTEGER' in item[2]:
                    self.autoincrement = True
                    if col_wdth == '':
                        ColWdth.append(6)
                # include the primary key and table name into SELECT statement
                datatable_str = (datatable_str + self.Lvl2tbl + '.' +
                                 self.pk_Name + ',')
            # need to make frgn_fld column noneditable in DVC
            elif 'INTEGER' in item[2] or 'FLOAT' in item[2]:
                if col_wdth == '' and 'FLOAT' in item[2]:
                    ColWdth.append(10)
                elif col_wdth == '':
                    ColWdth.append(6)
            elif 'BLOB' in item[2]:
                if col_wdth == '':
                    ColWdth.append(30)
            elif 'TEXT' in item[2] or 'BOOLEAN' in item[2]:
                if col_wdth == '':
                    ColWdth.append(10)
            elif 'DATE' in item[2]:
                if col_wdth == '':
                    ColWdth.append(10)

            # get first Lvl2 datatable column name in item[1]
            # check to see if name is lvl2 primary key or lvl1 linked field
            # if they are not then add tablename and
            # datafield to SELECT statement
            if item[1] != link_fld and item[1] != self.pk_Name:
                datatable_str = (
                    datatable_str + ' ' + self.Lvl2tbl + '.' + item[1] + ',')
            elif item[1] == link_fld:
                datatable_str = (
                    datatable_str + ' ' + self.frg_tbl + '.' + frgn_col + ',')

            n += 1
        self.realnames = realnames
        self.ColWdth = ColWdth

        datatable_str = datatable_str[:-1]

        DsqlLvl2 = ('SELECT ' + datatable_str + ' FROM ' +
                    self.Lvl2tbl + ' INNER JOIN ' + self.frg_tbl)
        DsqlLvl2 = (DsqlLvl2 + ' ON ' + self.Lvl2tbl + '.' +
                    link_fld + ' = ' + self.frg_tbl + '.' + self.frg_fld)

        self.DsqlLvl2 = DsqlLvl2
        self.data = Dbase().Dsqldata(self.DsqlLvl2)
        # specify which listbox column to display in the combobox
        self.showcol = int

        self.columnames = Dbase(self.Lvl2tbl).ColNames()

        if self.Lvl2tbl.find("_") != -1:
            frmtitle = (self.Lvl2tbl.replace("_", " "))
        else:
            frmtitle = (' '.join(re.findall('([A-Z][a-z]*)', self.Lvl2tbl)))

        frmwdth = 0
        for item in self.ColWdth:
            frmwdth = item + frmwdth
        if frmwdth <= 65:
            frmSize = (480, 450)
        elif 75 >= frmwdth > 65:
            frmSize = (650, 450)
        else:
            frmSize = (1200, 450)

        super(CmbLst2, self).__init__(parent, title=frmtitle, size=frmSize)

        self.Bind(wx.EVT_CLOSE, self.Close)

        # Create a dataview control
        self.dvc = dv.DataViewCtrl(self, wx.ID_ANY, wx.DefaultPosition,
                                   wx.Size(500, 300),
                                   style=wx.BORDER_THEME
                                   | dv.DV_ROW_LINES
                                   | dv.DV_VERT_RULES
                                   | dv.DV_HORIZ_RULES
                                   | dv.DV_SINGLE
                                   )

        self.dvc.SetMinSize = (wx.Size(100, 200))
        self.dvc.SetMaxSize = (wx.Size(500, 400))

        # Bind some events so we can see what the DVC sends us
        self.Bind(dv.EVT_DATAVIEW_ITEM_VALUE_CHANGED, self.OnValueChanged,
                  self.dvc)

        # pull out the foreign field name for the lable and combo box text
        self.frg_tbl = Dbase().Dtbldata(self.Lvl2tbl)[0][2]
        frg_fld = Dbase().Dtbldata(self.Lvl2tbl)[0][4]
        self.frg_fld = frg_fld.replace("_", " ")

        # if autoincrement is false then the data can be sorted based on ID_col
        if self.autoincrement == 0:
            self.data.sort(key=lambda tup: tup[self.pk_col])

        # use the sorted data to load the dataviewlistcontrol
        if self.model is None:
            self.model = DataMods(self.Lvl2tbl, self.data, self.edit_data)

        self.dvc.AssociateModel(self.model)

        if link_fld.find("_") != -1:
            frgn_flg_ttl = (link_fld.replace("_", " "))
        else:
            frgn_flg_ttl = (' '.join(re.findall('([A-Z][a-z]*)', link_fld)))
        frgn_col_num = self.columnames.index(frgn_flg_ttl)

        n = 0
        for colname in self.columnames:
            if len(colname) > self.ColWdth[n]:
                colwdth = len(colname)
            else:
                colwdth = self.ColWdth[n]

            self.RptColWdth.append(colwdth)

            if n == self.pk_col and self.autoincrement or n == self.pk_col:
                col_mode = dv.DATAVIEW_CELL_INERT
            else:
                if n == frgn_col_num:
                    col_mode = dv.DATAVIEW_CELL_INERT
                else:
                    col_mode = dv.DATAVIEW_CELL_EDITABLE

            self.dvc.AppendTextColumn(colname, n,
                                      width=wx.LIST_AUTOSIZE_USEHEADER,
                                      mode=col_mode)
            n += 1

        # make columns not sortable and but reorderable.
        for c in self.dvc.Columns:
            c.Sortable = False
            c.Reorderable = True
            c.Resizeable = True

        # change to not let the ID col be moved.
        self.dvc.Columns[(self.pk_col)].Reorderable = False
        self.dvc.Columns[(self.pk_col)].Resizeable = False

        # set the Sizer property (same as SetSizer)
        self.Sizer = wx.BoxSizer(wx.VERTICAL)

        # develope the comboctrl and attach popup list
        self.cmb1 = wx.ComboCtrl(self, pos=(10, 10), size=(200, -1))
        self.Bind(wx.EVT_TEXT, self.OnSelect, self.cmb1)
        self.cmb1.SetHint(self.frg_fld)
        self.showcol = 1
        self.popup = ListCtrlComboPopup(self.frg_tbl, showcol=self.showcol)

        self.cmbsizer = wx.BoxSizer(wx.HORIZONTAL)

        # add a button to call main form to search combo list data
        self.b6 = wx.Button(self, label="Restore\nData")
        self.Bind(wx.EVT_BUTTON, self.OnRestore, self.b6)
        self.cmbsizer.Add(self.b6, 0, wx.ALIGN_LEFT, 5)

        self.cmb1.SetPopupControl(self.popup)
        self.cmbsizer.Add(self.cmb1, 0, wx.ALIGN_CENTER, 5)

        # add a button to call main form to search combo list data
        self.b5 = wx.Button(self, label="<= Search Data")
        self.Bind(wx.EVT_BUTTON, self.OnSearch, self.b5)
        self.cmbsizer.Add(self.b5, 0)  # , wx.ALIGN_RIGHT)

        self.addlbl = wx.StaticText(self, -1, style=wx.ALIGN_CENTER_HORIZONTAL)

        txt = '   To add a new record first select a ' + self.frg_fld
        self.addlbl.SetLabel(txt)
        self.addlbl.SetForegroundColour((255, 0, 0))
        self.Sizer.Add(self.addlbl, 0, wx.ALIGN_LEFT)
        self.Sizer.Add(self.cmbsizer, 0, wx.ALIGN_CENTER)
        self.Sizer.Add(self.dvc, 1, wx.EXPAND)

        # Add buttons for grid modifications
        self.b1 = wx.Button(self, label="Print Report")
        self.Bind(wx.EVT_BUTTON, self.PrintFile, self.b1)
        self.b2 = wx.Button(self, label="Add Row")
        self.Bind(wx.EVT_BUTTON, self.OnAddRow, self.b2)
        self.b3 = wx.Button(self, label="Delete Row")
        self.Bind(wx.EVT_BUTTON, self.OnDeleteRow, self.b3)
        self.b4 = wx.Button(self, label="Exit")
        self.Bind(wx.EVT_BUTTON, self.Close, self.b4)

        # add a button box and place the buttons
        self.btnbox = wx.BoxSizer(wx.HORIZONTAL)
        self.btnbox.Add(self.b1, 0, wx.ALL | wx.ALIGN_CENTER, 5)
        self.btnbox.Add(self.b2, 0, wx.ALL | wx.ALIGN_CENTER, 5)
        self.btnbox.Add(self.b3, 0, wx.ALL | wx.ALIGN_CENTER, 5)
        self.btnbox.Add(self.b4, 0, wx.ALL | wx.ALIGN_CENTER, 5)

        # need to disable the add row button until an
        # item is selected in the combo box
        self.b2.Disable()

        # add static label to explain how to add / edit data
        self.editlbl = wx.StaticText(self, -1, style=wx.ALIGN_CENTER)
        txt = '    \nTo edit data double click on the cell.'
        self.editlbl.SetLabel(txt)
        self.editlbl.SetForegroundColour((255, 0, 0))

        self.Sizer.Add(self.editlbl, 0, wx.ALIGN_CENTER)

        self.Sizer.Add(self.btnbox, 0, wx.ALIGN_CENTER, 5)

        # add the next 5 lines for the child parent relation
        self.CenterOnParent()
        self.GetParent().Enable(False)
        self.Show(True)

        self.__eventLoop = wx.GUIEventLoop()
        self.__eventLoop.Run()

    def PrintFile(self, evt):
        import Report_Lvl2

        saveDialog = wx.FileDialog(self, message='Save Report as PDF.',
                                   wildcard='PDF (*.pdf)|*.pdf',
                                   style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT)

        if saveDialog.ShowModal() == wx.ID_CANCEL:
            filename = ''

        filename = saveDialog.GetPath()
        if filename.find(".pdf") == -1:
            filename = filename + '.pdf'

        saveDialog.Destroy()

        Report_Lvl2.Report(self.Lvl2tbl, self.data, self.columnames,
                           self.RptColWdth, filename).create_pdf()

    def TblCellEdit(self):
        # rowID is table index value, colID is column number of table index
        # new_value is the link foreign field value
        # colChgNum is the column number which has edited cell,
        # values is the new edited value(s)

        self.edit_data = self.model.edit_data
        rowID = self.model.GetValueByRow(self.edit_data[0], self.pk_col)
        self.edit_data.append(rowID)

        realnames = []
        for item in Dbase().Dcolinfo(self.Lvl2tbl):
            realnames.append(item[1])

        enable_b2 = 0
        # if just a cell is being edited then do the following
        colID = self.pk_col
        colChgNum = self.edit_data[1]
        values = self.edit_data[2]
        rowID = int(self.edit_data[3])
        colIDName = realnames[colID]
        colChgName = realnames[colChgNum]

        if values.find('"') != -1:
            values = values.replace('"', '""')
        if type(rowID) != str:
            UpQuery = ('UPDATE ' + self.Lvl2tbl + ' SET ' + colChgName + ' = "'
                       + str(values) + '" WHERE ' + str(colIDName) + ' = '
                       + str(rowID))
            check_query = ('SELECT * FROM ' + self.Lvl2tbl + ' WHERE ' +
                           str(colIDName) + ' = ' + str(rowID))
        else:
            UpQuery = ('UPDATE ' + self.Lvl2tbl + ' SET ' + colChgName + ' = "'
                       + str(values) + '" WHERE ' + str(colIDName) + ' = "'
                       + rowID + '"')

            # if the database ID col has been changed then
            # base the query on the new ID value if any other column
            # was changed then base the query on the original colID
            if colID == colChgNum:
                check_query = ('SELECT * FROM ' + self.Lvl2tbl + ' WHERE '
                               + str(colIDName) + ' = "' + values + '"')
            else:
                check_query = ('SELECT * FROM ' + self.Lvl2tbl + ' WHERE ' +
                               str(colIDName) + ' = "' + rowID + '"')

        Dbase().TblEdit(UpQuery)

        # check to see if added row has been properly edited
        # if so enable the addrow button
        data_string = Dbase().Dsqldata(check_query)
        if 'Required' not in data_string[0]:
            enable_b2 = 1

        self.data = Dbase().Dsqldata(self.DsqlLvl2)

        if enable_b2:
            self.b2.Disable()
            # if autoincrement is false then the data can
            # be sorted based on ID_col
            if self.autoincrement == 0:
                self.data.sort(key=lambda tup: tup[self.pk_col])

    def AddTblRow(self):
        # this returns the text value in the combo box not the linked ID value
        new_value = self.cmb1.GetValue()

        lnk_fld_name = Dbase().Dtbldata(self.Lvl2tbl)

        # this is the name of the linked field in the foreign table
        field = lnk_fld_name[0][4]
        # do a search in the foriegn table for the value of the
        # link ID which corresponds to the name in the combo box
        Shcol = Dbase().Dcolinfo(self.frg_tbl)[self.showcol][1]
        ShQuery = ('SELECT ' + field + ' FROM ' + self.frg_tbl +
                   " WHERE " + Shcol + " LIKE '%" + self.cmb1.GetValue()
                   + "%' COLLATE NOCASE")
        # this is the ID value for the linked field in the foreign table
        ShQueryVal = str(Dbase().Dsqldata(ShQuery)[0][0])

        if new_value:
            realnames = []
            for item in Dbase().Dcolinfo(self.Lvl2tbl):
                realnames.append(item[1])

            self.b2.Disable()

            # update data structure
            FldInfo = Dbase(self.Lvl2tbl).Fld_Size_Type()
            values = FldInfo[4]

            # if the table index is auto increment then assign
            # next value otherwise do nothing
            if FldInfo[2]:
                New_ID = cursr.execute(
                    "SELECT MAX(" + self.pk_Name + ") FROM "
                    + self.Lvl2tbl).fetchone()
                if New_ID[0] is None:
                    Max_ID = '1'
                else:
                    Max_ID = str(New_ID[0]+1)
                values[self.pk_col] = Max_ID

            # if a full row is being added then do the following
            ValueList = []
            if type(values) == str:
                ValueList.append(values)
            else:
                ValueList = values

            # need to change the linked foriegn
            # field to the value selected in the combobox

            if lnk_fld_name != []:
                ValueList[realnames.index(lnk_fld_name[0][3])] = ShQueryVal
            UpQuery = ('INSERT INTO ' + self.Lvl2tbl +
                       ' VALUES (' + "'" + "','".join(map(str, ValueList)) +
                       "'" + ')')

            Dbase(self.Lvl2tbl).TblEdit(UpQuery)

            self.data.append(values)

            self.data = Dbase().Dsqldata(self.DsqlLvl2)
            self.model = DataMods(self.Lvl2tbl, self.data)
            self.dvc.AssociateModel(self.model)
            self.dvc.Refresh

        else:
            wx.MessageBox(
                '''A material type needs to be selected,\n
                before adding a new row''',
                "Missing Information", wx.OK | wx.ICON_INFORMATION)

    def OnDeleteRow(self, evt):
        item = self.dvc.GetSelection()
        row = self.model.GetRow(item)
        self.b2.Disable()
        self.model.DeleteRow(row, self.pk_Name)

    def OnAddRow(self, evt):
        self.AddTblRow()

    def OnValueChanged(self, evt):
        self.TblCellEdit()

    def OnSearch(self, evt):
        # collect feign table info
        frgn_info = Dbase().Dtbldata(self.Lvl2tbl)
        field = frgn_info[0][4]
        self.frg_tbl = frgn_info[0][2]

        # do search of string value from combobox
        # equal to value in the self.frg_tbl
        Shcol = Dbase().Dcolinfo(self.frg_tbl)[self.showcol][1]
        ShQuery = ('SELECT ' + field + ' FROM ' + self.frg_tbl +
                   " WHERE " + Shcol + " LIKE '%" + self.cmb1.GetValue()
                   + "%' COLLATE NOCASE")
        ShQueryVal = str(Dbase().Dsqldata(ShQuery)[0][0])
        # append the found frgn_fld to the original data grid SQL and
        # find only records equal to the combo selection
        ShQuery = (self.DsqlLvl2 + ' WHERE ' + self.frg_tbl + '.' +
                   field + ' = "' + ShQueryVal + '"')

        OSdata = Dbase().Search(ShQuery)
        # if nothing is found show blank grid
        if OSdata is False:
            OSdata = []
        self.model = DataMods(self.Lvl2tbl, OSdata)
        self.dvc.AssociateModel(self.model)
        self.dvc.Refresh
        self.b2.Enable()

    def OnSelect(self, evt):
        self.b2.Enable()
        txt = ('''To complete adding a new record click "Add Row".\n
            Then edit data by double click on the cell.''')
        self.editlbl.SetLabel(txt)
        self.editlbl.SetForegroundColour((255, 0, 0))
        self.Sizer.Layout()

    def OnRestore(self, evt):
        self.ORdata = Dbase().Restore(self.DsqlLvl2)
        self.cmb1.ChangeValue('')
        self.cmb1.SetHint(self.frg_fld)
        self.model = DataMods(self.Lvl2tbl, self.ORdata)
        self.dvc._AssociateModel(self.model)
        self.dvc.Refresh

    def Close(self, evt):
        # add the next 2 lines for the child parent relation
        self.GetParent().Enable(True)
        self.__eventLoop.Exit()
        # Dbase().close_database()
        self.Destroy()


class ListCtrlComboPopup(wx.ComboPopup):

    '''CLASS TO HANDLE THE CHANGES TO THE COMBO POPUP LIST'''
    def __init__(self, tbl_name, PupQuery='',
                 cmbvalue='', showcol=0, lctrls=None):

        wx.ComboPopup.__init__(self)
        self.tbl_name = tbl_name
        self.PupQuery = PupQuery
        self.cmbvalue = cmbvalue
        self.showcol = showcol
        self.lctrls = lctrls
        self.name_list = []

        if self.PupQuery == '':
            self.PupQuery = 'SELECT * FROM ' + self.tbl_name
            for item in Dbase().Dcolinfo(self.tbl_name):
                self.name_list.append(item[1])
        else:    # get last of required column names for the combo box list
            spot = self.PupQuery.index('FROM', 0, len(self.PupQuery))
            tblnms = self.PupQuery[6:spot]
            tblnms = tblnms.replace(self.tbl_name+'.', '')
            tblnms = tblnms.strip(' ')
            if tblnms == '*':
                for item in Dbase().Dcolinfo(self.tbl_name):
                    self.name_list.append(item[1])
            else:
                self.name_list = tblnms.split(',')

    def AddItem(self, txt):
        self.lc.InsertItem(self.lc.GetItemCount(), txt)

    def OnLeftDown(self, evt):
        item, flags = self.lc.HitTest(evt.GetPosition())
        if item == -1:
            return
        if item >= 0:
            self.lc.Select(item, on=0)
            self.curitem = item
        self.value = self.curitem
        self.Dismiss()

    # This is called immediately after construction finishes.  You can
    # use self.GetCombo if needed to get to the ComboCtrl instance.
    def Init(self):
        self.value = -1
        self.curitem = -1

    # Create the popup child control.  Return true for success.
    def Create(self, parent):
        self.index = 0

        self.lc = wx.ListCtrl(parent, size=wx.DefaultSize,
                              style=wx.LC_REPORT | wx.BORDER_SUNKEN
                              | wx.LB_SINGLE)

        # this looks up the data for
        # the individual tables representing each combo box
        self.lc.Bind(wx.EVT_LEFT_DOWN, self.OnLeftDown)
        if self.lctrls is not None:
            self.lctrls.append(self.lc)

        LstColWdth = []
        for item in Dbase().Dcolinfo(self.tbl_name):
            if item[1] in self.name_list:
                lstcol_wdth = ''
                # check to see if field length is specified if so
                # use it to set grid col width
                for s in re.findall(r'\d+', item[2]):
                    if s.isdigit():
                        lstcol_wdth = int(s)
                        LstColWdth.append(lstcol_wdth)

                if 'INTEGER' in item[2] or 'FLOAT' in item[2]:
                    if lstcol_wdth == '' and 'FLOAT' in item[2]:
                        LstColWdth.append(10)
                    elif lstcol_wdth == '' and 'INTEGER' in item[2]:
                        LstColWdth.append(6)
                elif 'BLOB' in item[2]:
                    if lstcol_wdth == '':
                        LstColWdth.append(30)
                elif 'TEXT' in item[2]:
                    if lstcol_wdth == '':
                        LstColWdth.append(10)
                elif 'DATE' in item[2]:
                    if lstcol_wdth == '':
                        LstColWdth.append(10)

        n = 0
        for info in Dbase().Dcolinfo(self.tbl_name):
            info = list(info)
            colname = info[1]
            if colname in self.name_list:
                if colname.find("ID", -2) != -1:
                    colname = "ID"
                elif colname.find("_") != -1:
                    colname = colname.replace("_", " ")
                else:
                    colname = (' '.join(re.findall('([A-Z][a-z]*)', colname)))
                self.lc.InsertColumn(n, colname)

                lstcolwdth = LstColWdth[n]*9
                if (len(colname)*9) > lstcolwdth:
                    lstcolwdth = len(colname)*9

                self.lc.SetColumnWidth(n, lstcolwdth)
                n += 1

        index = 0
        for values in Dbase().Dsqldata(self.PupQuery):
            col = 0
            for value in values:
                if col == 0:
                    self.lc.InsertItem(index, str(value))
                else:
                    self.lc.SetItem(index, col, str(value))
                col += 1
            index += 1
        return True

    # Return the widget that is to be used for the popup
    def GetControl(self):
        return self.lc

    # Return a string representation of the current item.
    def GetStringValue(self):
        if self.value == -1:
            return
        return self.lc.GetItemText(self.value, self.showcol)

    # Called immediately after the popup is shown
    def OnPopup(self):
        # this provides the original combox value
        # if self.value >= 0:
        #    print('OnPopUp',self.lc.GetItemText(self.value))
        wx.ComboPopup.OnPopup(self)

    # Called when popup is dismissed
    def OnDismiss(self):
        # this provides the new combo value
        # print('OnDismiss',self.lc.GetItemText(self.value))
        wx.ComboPopup.OnDismiss(self)

    def PaintComboControl(self, dc, rect):
        # This is called to custom tube in the combo control itself
        # (ie. not the popup).  Default implementation draws value as
        # string.
        wx.ComboPopup.PaintComboControl(self, dc, rect)

    # Return final size of popup. Called on every popup, just prior to OnPopup.
    # minWidth = preferred minimum width for window
    # prefHeight = preferred height. Only applies if > 0,
    # maxHeight = max height for window, as limited by screen size
    #   and should only be rounded down, if necessary.
    def GetAdjustedSize(self, minWidth, prefHeight, maxHeight):
        minWidthNew = 0
        for cl in range(0, self.lc.GetColumnCount()):
            minWidthNew = self.lc.GetColumnWidth(cl) + minWidthNew
        minWidth = minWidthNew
        return wx.ComboPopup.GetAdjustedSize(self, minWidth, 150, 800)

    # Return true if you want delay the call to Create until the popup
    # is shown for the first time. It is more efficient, but note that
    # it is often more convenient to have the contrminWidth
    # immediately.
    # Default returns false.
    def LazyCreate(self):
        return wx.ComboPopup.LazyCreate(self)


class DataMods(dv.DataViewIndexListModel):
    '''CLASS TO HANDLE THE CHANGES TO THE TABLE DATA'''
    def __init__(self, frmtbl, data, edit_data=None):
        dv.DataViewIndexListModel.__init__(self, len(data))
        self.data = data
        self.frmtbl = frmtbl
        self.edit_data = edit_data

    # This method is called when the user edits a data item in the view.
    def SetValueByRow(self, value, row, col):
        self.edit_data = [row, col, value]
        tmp = list(self.data[row])
        tmp[col] = value
        self.data[row] = tuple(tmp)
        return True

    # This method is called to provide the data object for a particular row,col
    def GetValueByRow(self, row, col):
        if (len(str(self.data[row][col]))) > 70:
            new_strg = self.wrap_paragraphs(str(self.data[row][col]), 70)
            new_strg = " ".join(new_strg)
            return new_strg
        else:
            return str(self.data[row][col])

    # Report how many columns this model provides data for.
    def GetColumnCount(self):
        if self.data == []:
            return 0
        return len(self.data[0])

    # Specify the data type for a column
    def GetColumnType(self, col):
        return "string"

    # Report the number of rows in the model
    def GetCount(self):
        return len(self.data)

    # Called to check if non-standard attributes should be used
    # in the cell at (row, col)
    def GetAttrByRow(self, row, col, attr):
        if col == 0:
            attr.SetColour('blue')
            attr.SetBold(True)
            return True
        return False

    def dedent(self, text):

        if text.startswith('\n'):
            # text starts with blank line, don't ignore the first line
            return textwrap.dedent(text)

        # split first line
        splits = text.split('\n', 1)
        if len(splits) == 1:
            # only one line
            return textwrap.dedent(text)

        first, rest = splits
        # dedent everything but the first line
        rest = textwrap.dedent(rest)
        return '\n'.join([first, rest])

    def wrap_paragraphs(self, text, ncols):

        paragraph_re = re.compile(r'\n(\s*\n)+', re.MULTILINE)
        text = self.dedent(text).strip()
        # every other entry is space
        paragraphs = paragraph_re.split(text)[::2]
        out_ps = []
        indent_re = re.compile(r'\n\s+', re.MULTILINE)

        for p in paragraphs:
            # presume indentation that survives dedent
            # is meaningful formatting,
            # so don't fill unless text is flush.
            if indent_re.search(p) is None:
                # wrap paragraph
                p = textwrap.fill(p, ncols)
            out_ps.append(p)
        return out_ps

    # copy then delete the row of data
    def DeleteRow(self, row, delPK_name):
        try:
            DtQueryVal = self.data[row][0]
            Dbase().TblDelete(self.frmtbl, DtQueryVal, delPK_name)
            del self.data[row]
            # notify the view(s) using this model that it has been removed
            self.RowDeleted(row)
        except sqlite3.IntegrityError:
            wx.MessageBox('This ' + delPK_name +
                          ''' is associated with other tables and cannot
                           be deleted!''', "Cannot Delete",
                          wx.OK | wx.ICON_INFORMATION)

    # insert a blank row
    def AddRow(self):
        # notify views
        self.RowAppended()
        return


class Dbase(object):
    '''DATABASE CLASS HANDLER'''
    # this initializes the database and opens the specified table
    def __init__(self, frmtbl=None):
        # this sets the path to the database and needs
        # to be changed accordingly
        self.cursr = cursr
        self.db = db
        self.frmtbl = frmtbl

    def Dcolinfo(self, table):
        # sequence for items in colinfo is column number, column name,
        # data type(size), not null, default value, primary key
        self.cursr.execute("PRAGMA table_info(" + table + ");")
        colinfo = self.cursr.fetchall()
        return colinfo

    def Dtbldata(self, table):
        # this will provide the foreign keys and their related tables
        # unknown,unknown,Frgn Tbl,Parent Tbl Link fld,
        # Frgn Tbl Link fld,action,action,default
        self.cursr.execute('PRAGMA foreign_key_list(' + table + ')')
        tbldata = list(self.cursr.fetchall())
        return tbldata

    def Dsqldata(self, DataQuery):
        # provides the actual data from the table
        self.cursr.execute(DataQuery)
        sqldata = self.cursr.fetchall()
        return sqldata

    def TblDelete(self, table, val, field):
        '''Call the function to delete the values in
        the grid into the database error trapping will occure
        in the call def delete_data'''

        if type(val) != str:
            DeQuery = ("DELETE FROM " + table + " WHERE "
                       + field + " = " + str(val))
        else:
            DeQuery = ("DELETE FROM " + table + " WHERE "
                       + field + " = '" + val + "'")
        self.cursr.execute(DeQuery)
        self.db.commit()

    def TblEdit(self, UpQuery, data_strg=None):
        if data_strg is None:
            self.cursr.execute(UpQuery)
        else:
            self.cursr.execute(UpQuery, data_strg)
        self.db.commit()

    def PrptyDelete(self, query):
        self.cursr.execute(query)
        self.db.commit()

        # determine the required column width, name and primary status
    def Fld_Size_Type(self):
        # specified field type or size
        values = []
        auto_incr = True
        ColWdth = []
        n = 0

        # collect all the table information needed to build the grid
        # colinfo includes schema for each column: column number, name,
        # field type & length, is null , default value, is primary key
        for item in self.Dcolinfo(self.frmtbl):
            col_wdth = ''
            # check to see if field length is specified if so
            # use it to set grid col width
            for s in re.findall(r'\d+', item[2]):
                if s.isdigit():
                    col_wdth = int(s)
                    ColWdth.append(col_wdth)
            # check if it is a text string and primary key if it is then
            # it is not auto incremented develope a string of data based
            # on record field type for adding new row
            if item[5] == 1:
                pk = item[1]
                pk_col = n
                # if the primary key is not an interger then assign a text
                # value and indicate it is not auto incremented
                if 'INTEGER' not in item[2]:
                    values.append('Required')
                    auto_incr = False
                # it must be an integer and will be auto incremeted,
                # New_ID is assigned in AddRow routine
                else:
                    values.append('New_ID')
                    if col_wdth == '':
                        ColWdth.append(6)
            # remaining steps assing values to not
            # null fields otherwise leave empty
            elif 'INTEGER' in item[2] or 'FLOAT' in item[2]:
                if item[3]:
                    values.append(0)
                else:
                    values.append('')
                if col_wdth == '' and 'FLOAT' in item[2]:
                    ColWdth.append(10)
                elif col_wdth == '':
                    ColWdth.append(6)
            elif 'BLOB' in item[2]:
                if item[3]:
                    values.append('Required')
                else:
                    values.append('')
                if col_wdth == '':
                    ColWdth.append(30)
            elif 'TEXT' in item[2] or 'BOOLEAN' in item[2]:
                if item[3]:
                    values.append('Required')
                else:
                    values.append('')
                if col_wdth == '':
                    ColWdth.append(10)
            elif 'DATE' in item[2]:
                i = datetime.datetime.now()
                today = ("%s-%s-%s" % (i.month, i.day, i.year))
                if item[3]:
                    values.append(today)
                if col_wdth == '':
                    ColWdth.append(10)
            n += 1

        # the variables in FldInfo are;database column name for ID,
        # database number of ID column, if ID is autoincremented
        # (imples interger or stg), list of database specified column
        # width, a list of database column names,
        # a list of values to insert into none null fields
        FldInfo = [pk, pk_col, auto_incr, ColWdth, values]
        return FldInfo

    def ColNames(self):
        colnames = []
        for item in self.Dcolinfo(self.frmtbl):
            # modify the column names to remove
            # underscore and seperate words
            colname = item[1]
            if colname.find("ID", -2) != -1:
                colname = "ID"
            elif colname.find("_") != -1:
                colname = colname.replace("_", " ")
            else:
                colname = (' '.join(re.findall('([A-Z][a-z]*)', colname)))
            colnames.append(colname)
        return colnames

    def Search(self, ShQuery):
        self.cursr.execute(ShQuery)
        data = self.cursr.fetchall()
        if data == []:
            return False
        else:
            return data

    def Restore(self, RsQuery):
        self.cursr.execute(RsQuery)
        data = self.cursr.fetchall()
        # self.cursr.close()
        return data

    # close out the database
    def close_database(self):
        self.cursr.close()
        del self.cursr
        self.db.close()


if __name__ == "__main__":
    app = wx.App(False)
    frame = StrUpFrm()
    frame.Show()
    frame.CenterOnScreen()
    app.MainLoop()
