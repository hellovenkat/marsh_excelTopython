import os
import sys
from os.path import join, dirname, abspath
import wx
import wx.animate
import math
import wx.lib.agw.multidirdialog as MDD
import wx.lib.scrolledpanel as scrolled
import wx.lib.mixins.gridlabelrenderer as glr
import xlrd
import wx.grid as grid
import xlsgrid as XG
import wx.grid as gridlib
import matplotlib
matplotlib.use('WXAgg')
import matplotlib.pyplot as plt
from matplotlib.backends.backend_wxagg import FigureCanvasWxAgg as FigureCanvas
from matplotlib.backends.backend_wx import NavigationToolbar2Wx
from matplotlib.figure import Figure
wildcard = "Excel Workbook (*.xls)|*.xls|" \
            "All files (*.*)|*.*"
class MyGrid(grid.Grid, glr.GridWithLabelRenderersMixin):
    def __init__(self, *args, **kw):
        grid.Grid.__init__(self, *args, **kw)
        glr.GridWithLabelRenderersMixin.__init__(self)

class TextLabelRenderer(glr.GridLabelRenderer):
    def __init__(self, text, colspan,bgcolour=None):
        self.text = text
        self.colspan = colspan
        if bgcolour is not None:
            self.bgcolour = bgcolour
        else:
            self.bgcolour = "white"

    def Draw(self, grid, dc, rect, col):
        if self.colspan == 0:
            rect.SetSize((0,0))
        if self.colspan > 1:
            add_cols = self.colspan - 1
            l = rect.left
            r = rect.right + ((rect.Size.x -1) * add_cols)
            rect.left = l
            rect.right = r
        dc.SetBrush(wx.Brush(self.bgcolour))
        dc.SetPen(wx.TRANSPARENT_PEN)
        dc.DrawRectangleRect(rect)
        hAlign, vAlign = grid.GetColLabelAlignment()
        text = self.text
        if self.colspan != 0:
            self.DrawBorder(grid, dc, rect)
        self.DrawText(grid, dc, rect, text, hAlign, vAlign)

# Define the tab content as classes:
class TabOne(wx.Panel):

    def __init__(self, parent):
        wx.Panel.__init__(self, parent)
        self.InitUI()
    #def populate(self):
        #self.rupPan.SetBackgroundColour('red')
    def InitUI(self):
        #panel = wx.Panel(self)
        #panel.SetBackgroundColour('#4f5049')
        '''t = wx.StaticText(self, -1, "This is the first tab", (20, 20))
        hbox1 = wx.BoxSizer(wx.HORIZONTAL)
        left = wx.Panel(self)
        left.SetBackgroundColour('cyan')
        #hbox1.Add(left, border=8)
        hbox1.Add(left, 3, wx.EXPAND | wx.ALL, 5)
        right = wx.Panel(self)
        right.SetBackgroundColour('red')
        #hbox1.Add(right,  border=8)
        hbox1.Add(right, 3, wx.EXPAND | wx.ALL, 5)
        self.SetSizer(hbox1)'''
        #panel = wx.Panel(self)
        #self.SetBackgroundColour('#4f5049')
        self.SetBackgroundColour('#edeeff')
        vbox = wx.BoxSizer(wx.VERTICAL)
        hbox = wx.BoxSizer(wx.HORIZONTAL)

        #leftPan = wx.Panel(self)
        leftPan = wx.lib.scrolledpanel.ScrolledPanel(self)
        leftPan.SetupScrolling()
        #rightPan = wx.lib.scrolledpanel.ScrolledPanel(self)
        #rightPan.SetupScrolling()
        # leftPan.SetBackgroundColour('cyan')
        vbox_leftPan = wx.BoxSizer(wx.VERTICAL)
        global cb1, cb2, cb3, cb4
        cb1 = wx.CheckBox(leftPan, label='Use a generic biom profile')
        cb2 = wx.CheckBox(leftPan, label='Add thin layer')
        cb3 = wx.CheckBox(leftPan, label='Calibrate to accretion rate')
        cb4 = wx.CheckBox(leftPan, label='for future development')
        '''vbox_leftPan.Add(cb1,0,wx.ALIGN_CENTER)
        vbox_leftPan.Add(cb2,0,wx.ALIGN_CENTER)
        vbox_leftPan.Add(cb3,0,wx.ALIGN_CENTER)
        vbox_leftPan.Add(cb4,0,wx.ALIGN_CENTER)'''
        leftPan.SetSizer(vbox_leftPan)
        vbox_leftPan.Add((-1, 3))
        vbox_leftPan.Add(cb1)
        vbox_leftPan.Add(cb2)
        vbox_leftPan.Add(cb3)
        vbox_leftPan.Add(cb4)
        vbox_leftPan.Add((-1, 10))
        runSim = wx.Button(leftPan, -1, 'Run Simulation')
        runSim.Bind(wx.EVT_BUTTON, self.onCalculate)
        vbox_leftPan.Add(runSim)
        vbox_leftPan.Add((-1, 25))


        #############################
        lbl = wx.StaticText(leftPan, -1, style=wx.ALIGN_CENTER)
        txt = "                   Physical Inputs"
        font = wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD)
        lbl.SetFont(font)
        lbl.SetLabel(txt)
        vbox_leftPan.Add(lbl)
        fgs_phy = wx.FlexGridSizer(8, 3, 0, 0)
        global phy_sea_level_forecast, phy_sea_level_start, phy_20th, phy_MTA, phy_Marsh_ele, phy_sus_minSed, phy_sus_org, phy_lt
        phy_sea_level_forecast = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="",size=(40, -1))
        phy_sea_level_start = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="-1",size=(40, -1))
        phy_20th = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="0.2",size=(40, -1))
        phy_MTA = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="80",size=(40, -1))
        phy_Marsh_ele = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="80",size=(40, -1))
        phy_sus_minSed = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="80",size=(40, -1))
        phy_sus_org = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="80",size=(40, -1))
        phy_lt = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="80",size=(40, -1))
        fgs_phy.AddMany([(wx.StaticText(leftPan, style=wx.TE_RIGHT, label="Sea Level Forecast")),
                    (phy_sea_level_forecast),
                    (wx.StaticText(leftPan, style=wx.TE_RIGHT, label=" cm/100y")),
                    (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="Sea Level at Start")),
                    phy_sea_level_start,
                    (wx.StaticText(leftPan, style=wx.TE_LEFT, label=" cm (NAVD)")),
                    (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="20th Cent Sea Level Rate")),
                    phy_20th,
                    (wx.StaticText(leftPan, style=wx.TE_RIGHT, label=" cm/yr")),
                    (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="Mean Tidal Amplitude")),
                    phy_MTA,
                    (wx.StaticText(leftPan, style=wx.TE_RIGHT, label=" cm/ (MSL)")),
                    (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="Marsh Elevation @ t0")),
                    phy_Marsh_ele,
                    (wx.StaticText(leftPan, style=wx.TE_RIGHT, label=" cm/ (MSL)")),
                    (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="Suspended Min. Sed. Conc.")),
                    phy_sus_minSed,
                    (wx.StaticText(leftPan, style=wx.TE_RIGHT, label=" mg/l")),
                    (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="Suspended Org. Conc.")),
                    phy_sus_org,
                    (wx.StaticText(leftPan, style=wx.TE_RIGHT, label=" mg/l")),
                    (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="LT Accretion Rate")),
                    phy_lt,
                    (wx.StaticText(leftPan, style=wx.TE_RIGHT, label=" cm/yr"))
                    ])

        vbox_leftPan.Add(fgs_phy, proportion=0, flag=wx.EXPAND)
        vbox_leftPan.Add((-1, 25))
        ############################
        lbl1 = wx.StaticText(leftPan, -1, style=wx.ALIGN_CENTER)
        txt1 = "                   Biological Inputs"
        font1 = wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD)
        lbl1.SetFont(font1)
        lbl1.SetLabel(txt1)
        vbox_leftPan.Add(lbl1)
        #gs1 = wx.GridSizer(10, 3, 0, 0)
        fgs_bio = wx.FlexGridSizer(10, 3, 0, 0)
        global bio_max_growth, bio_min_growth, bio_opt_growth, bio_max_peak, bio_OM_below_root, bio_OM_decay, bio_BGBio, bio_BG_turnover, bio_max_root_depth, bio_reserved
        bio_max_growth = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="80",size=(40, -1))
        bio_min_growth = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="-1",size=(40, -1))
        bio_opt_growth = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="0.2",size=(40, -1))
        bio_max_peak = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="80",size=(40, -1))
        bio_OM_below_root = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="80",size=(40, -1))
        bio_OM_decay = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="80",size=(40, -1))
        bio_BGBio = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="80",size=(40, -1))
        bio_BG_turnover = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="80",size=(40, -1))
        bio_max_root_depth = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="80",size=(40, -1))
        bio_reserved = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="80",size=(40, -1))
        fgs_bio.AddMany([((wx.StaticText(leftPan, style=wx.TE_RIGHT, label="max growth limit (rel MSL)"))),
                     #(wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="80")),
                     (bio_max_growth),
                     ((wx.StaticText(leftPan, style=wx.TE_RIGHT, label=" cm"))),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="min growth limit (rel MSL)")),
                     (bio_min_growth),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label=" cm")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="opt growth elev (rel MSL)")),
                     (bio_opt_growth),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label=" cm")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="max peak biomass")),
                     (bio_max_peak),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label=" g/m2")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="%OM below root zone")),
                     (bio_OM_below_root),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="OM decay rate")),
                     (bio_OM_decay),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label=" 1/year")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="BGBio to Shoot Ratio")),
                     (bio_BGBio),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label=" g/g")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="BG turnover rate")),
                     (bio_BG_turnover),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label=" 1/year")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="Max Root Depth")),
                     (bio_max_root_depth),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label=" cm")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="Reserved")),
                     (bio_reserved),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label=" cm"))
                     ])

        vbox_leftPan.Add(fgs_bio, proportion=0, flag=wx.EXPAND)
        vbox_leftPan.Add((-1, 25))
        ###########################
        lbl2 = wx.StaticText(leftPan, -1, style=wx.ALIGN_CENTER)
        txt2 = "            Model Coefficients"
        font2 = wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD)
        lbl2.SetFont(font2)
        lbl2.SetLabel(txt2)
        vbox_leftPan.Add(lbl2)
        global model_max_capture,model_refrac
        model_max_capture = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="80",size=(40, -1))
        model_refrac = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="-1",size=(40, -1))
        fgs_model = wx.FlexGridSizer(2, 3, 0, 0)
        fgs_model.AddMany([(wx.StaticText(leftPan, style=wx.TE_RIGHT, label="Max Capture Eff (q)              ")),
                     (model_max_capture),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label=" tide")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="Refrac. Fraction (kr)")),
                     (model_refrac),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label=" g/g"))
                     ])

        vbox_leftPan.Add(fgs_model, proportion=0, flag=wx.EXPAND)
        vbox_leftPan.Add((-1, 25))
        ###############################
        lbl3 = wx.StaticText(leftPan, -1, style=wx.ALIGN_CENTER)
        txt3 =''' Episodic Storm Inputs or 
         Thin Layer Placement'''
        #txt3 = "      Episodic"
        font3 = wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD)
        lbl3.SetFont(font3)
        lbl3.SetLabel(txt3)
        vbox_leftPan.Add(lbl3)
        global epi_years, epi_repeat, epi_recoveryTime, epi_addElevation
        epi_years = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="20",size=(40, -1))
        epi_repeat = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="20",size=(40, -1))
        epi_recoveryTime = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="10",size=(40, -1))
        epi_addElevation = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="10",size=(40, -1))
        fgs_epi = wx.FlexGridSizer(4, 3, 0, 0)
        fgs_epi.AddMany([(wx.StaticText(leftPan, style=wx.TE_RIGHT, label="years from start                    ")),
                     (epi_years),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label=" years")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="repeat interval")),
                     (epi_repeat),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label=" years")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="recovery time")),
                     (epi_recoveryTime),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label=" years")),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label="add elevation")),
                     (epi_addElevation),
                     (wx.StaticText(leftPan, style=wx.TE_RIGHT, label=" cm"))
                      ])

        vbox_leftPan.Add(fgs_epi, proportion=0, flag=wx.EXPAND)
        vbox_leftPan.Add((-1, 25))
        # panel2 = wx.lib.scrolledpanel.ScrolledPanel(self, -1, size=(screenWidth, 400), pos=(0, 28),
        # style=wx.SIMPLE_BORDER)
        # leftPan.SetupScrolling()
        ###############################
        hbox.Add(leftPan, 1, wx.EXPAND | wx.ALL, 5)

        self.rupPan = wx.Panel(self)
        self.rupPan.SetBackgroundColour('#edeeff')
        #self.drawImages()
        #self.vbox_rupPan = wx.BoxSizer(wx.VERTICAL)

        #self.hbox_rupPan = wx.BoxSizer(wx.HORIZONTAL)
        wx.StaticText(self.rupPan, -1,"Copyright University of South Carolina 2010. All Rights Reserved, JT Morris 6-9-10", (325, 390))
        #qw = wx.StaticText(self.rupPan, style=wx.TE_CENTER,label="                                                                             Copyright University of South Carolina 2010. All Rights Reserved, JT Morris 6-9-10")
        # vbox.Add(qw, 2, wx.EXPAND | wx.ALL, 0)
        # hbox_rdownPan.Add(text_rdownPan, 2, wx.EXPAND | wx.ALL, 0)

        # st1 = wx.StaticText(rupPan, label='North Inlet, SC')
        # st2 = wx.StaticText(rupPan, label='MEM-TLP 6.0')
        vbox.Add(self.rupPan, 2, wx.EXPAND | wx.ALL, 0)

        # rdownPan = wx.Panel(panel)
        # rdownPan.SetBackgroundColour('#eeeeee')
        hbox_rdownPan = wx.BoxSizer(wx.HORIZONTAL)
        # rupPan.SetSizer(vbox_rupPan)
        rbut_rdownPan = wx.Panel(self)
        rtext_rdownPan = wx.Panel(self)
        rtext_vbox = wx.BoxSizer(wx.VERTICAL)
        # rbut_rdownPan.SetBackgroundColour('cyan')
        hbox_rdownPan.Add(rbut_rdownPan, 0.90, wx.EXPAND | wx.ALL, 0)
        hbox_rdownPan.Add(rtext_rdownPan, 2, wx.EXPAND | wx.ALL, 0)
        #rtext_rdownPan.SetBackgroundColour('red')
        rdown_text1='''Metrics computed over the final 50 years of simulation
   1.16 avg vert accretion (cm/yr)  last 50 yr of the simulation (yrs 51-100 average)
   36.7 refractory c seq (g C/m2/yr) at the end of the simulation from top 50 cohorts
   52.1 total C/m2 in belowground biomass in top 50 cohorts (50 years) (g C/m2/yr)'''

        lbl_1 = wx.StaticText(rtext_rdownPan, -1, style=wx.ALIGN_CENTER)

        rtextfont = wx.Font(11, wx.ROMAN, wx.FONTSTYLE_NORMAL, wx.NORMAL)
        lbl_1.SetFont(rtextfont)
        lbl_1.SetLabel(rdown_text1)
        rtext_vbox.Add((-1, 10))
        rtext_vbox.Add((-1, 10))
        rtext_vbox.Add(lbl_1, 0, wx.ALIGN_CENTER)
        rtext_vbox.Add((-1, 25))
        rdown_text2='''Metrics computed over the first 50 years of simulation
   0.87 avg vert accretion (cm/yr)  first 50 yr of the simulation (yrs 1-50 average)
41.4 refractory c seq (g C/m2/yr) at the mid point of the simulation from top 50 cohorts
   58.1 total C/m2 in belowground biomass from top 50 cohorts (50 years) (g C/m2/yr)
'''
        lbl_2 = wx.StaticText(rtext_rdownPan, -1, style=wx.ALIGN_CENTER)
        lbl_2.SetFont(rtextfont)
        lbl_2.SetLabel(rdown_text2)
        rtext_vbox.Add(lbl_2, 0, wx.ALIGN_CENTER)
        #rtext_rdownPan.Add(rtext_vbox)
        rtext_rdownPan.SetSizer(rtext_vbox)
        vbox.Add(hbox_rdownPan, 1, wx.EXPAND | wx.ALL, 0)
        rbut_vbox = wx.BoxSizer(wx.VERTICAL)
        r1 = wx.RadioButton(rbut_rdownPan, label='Plum Island, MA')
        r2 = wx.RadioButton(rbut_rdownPan, label='North Inlet, SC')
        r3 = wx.RadioButton(rbut_rdownPan, label='Apalachicola, FL')
        r4 = wx.RadioButton(rbut_rdownPan, label='Grand Bay, MS')
        r5 = wx.RadioButton(rbut_rdownPan, label='Coon Isl, SFB')
        r6 = wx.RadioButton(rbut_rdownPan, label='Other Estuary')
        #r2.SetValue(True)
        r1.Bind(wx.EVT_RADIOBUTTON, self.onRadioButton)
        # rupPan.SetSizer(vbox_rupPan)
        r2.Bind(wx.EVT_RADIOBUTTON, self.onRadioButton)
        r3.Bind(wx.EVT_RADIOBUTTON, self.onRadioButton)
        r4.Bind(wx.EVT_RADIOBUTTON, self.onRadioButton)
        r5.Bind(wx.EVT_RADIOBUTTON, self.onRadioButton)
        r6.Bind(wx.EVT_RADIOBUTTON, self.onRadioButton)
        r2.SetValue(True)
        self.onRadioButton(None)

        #wx.PostEvent(r2, wx.CommandEvent(wx.wxEVT_RADIOBUTTON_CLICKED))
        rbut_vbox.Add((-1, 10))
        rbut_vbox.Add((-1, 10))
        rbut_vbox.Add(r1,0,wx.LEFT,20)
        rbut_vbox.Add((-1, 15))
        rbut_vbox.Add(r2,0,wx.LEFT,20)
        rbut_vbox.Add((-1, 15))
        rbut_vbox.Add(r3,0,wx.LEFT,20)
        rbut_vbox.Add((-1, 15))
        rbut_vbox.Add(r4,0,wx.LEFT,20)
        rbut_vbox.Add((-1, 15))
        rbut_vbox.Add(r5,0,wx.LEFT,20)
        rbut_vbox.Add((-1, 15))
        rbut_vbox.Add(r6,0,wx.LEFT,20)
        rbut_rdownPan.SetSizer(rbut_vbox)
        #text_rdownPan = wx.Panel(self)
        # text_rdownPan.SetBackgroundColour('red')
        #text_rdownPan.Add((-1, 10))
       # //wx.StaticText(text_rdownPan, style=wx.TE_LEFT, label=" Metrics computed over the final 50 years of simulation")
        #hbox_rdownPan.Add(text_rdownPan, 1, wx.EXPAND | wx.ALL, 0)

        '''browsePanel = wx.Panel(self)
        hbox_rdownPan.Add(browsePanel, 1, wx.EXPAND | wx.ALL, 0)
        browsePanel_vbox = wx.BoxSizer(wx.VERTICAL)
        AnotherFile = wx.StaticText(browsePanel, style=wx.TE_RIGHT, label="Choose another excel file")
        self.currentDirectory = os.getcwd()
        openFileDlgBtn = wx.Button(browsePanel, label="Browse")
        openFileDlgBtn.Bind(wx.EVT_BUTTON, self.onOpenFile)
        closeBtn = wx.Button(browsePanel, label="Change")
        closeBtn.Bind(wx.EVT_BUTTON, self.onClose)
        #button = wx.Button(browsePanel, id=wx.ID_ANY, label="Change")
        #button.Bind(wx.EVT_BUTTON, self.onButton, what)
        browsePanel_vbox.Add((-1, 10))
        browsePanel_vbox.Add(AnotherFile)
        browsePanel_vbox.Add((-1, 10))
        #browsePanel_vbox.Add(textBox)
        #browsePanel_vbox.Add((-1, 10))
        browsePanel_vbox.Add(openFileDlgBtn)
        browsePanel_vbox.Add((-1, 10))
        browsePanel_vbox.Add(closeBtn)
        browsePanel_vbox.Add((-1, 10))
        browsePanel.SetSizer(browsePanel_vbox)'''

        hbox.Add(vbox, 3, wx.EXPAND | wx.ALL, 5)
        self.SetSizer(hbox)

        #panel.EnableScrolling(True,True)
    def drawImages(self):
        #vbox_rupPan = wx.BoxSizer(wx.VERTICAL)

        #hbox_rupPan = wx.BoxSizer(wx.HORIZONTAL)
        # rupPan.SetSizer(hbox_rupPan)
        #self.vbox_rupPan.Add((-1, 10))

        # rup_title = wx.StaticText(rupPan, style=wx.ALIGN_CENTRE, label = "North Inlet, SC")
        # vbox_rupPan.Add(rup_title, wx.LEFT|wx.RIGHT, 30)
        # ruptxt = "North Inlet, SC"+"\n"+"MEM-TLP 6.0"
        # rupfont = wx.Font(16, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD)
        # ruplbl.SetFont(rupfont)
        # ruplbl.SetLabel(ruptxt)
        #self.vbox_rupPan.Add((-1, 30))
        #print "in panel"
        #print self
        folder = 'asset'
        #print join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))),folder,'first.png')
        #image1 = join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))), folder, 'first.png')
        image1 = join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))), folder, 'first.png')
        #print image1
        try:
            #print "try"
            image = wx.Image(image1, wx.BITMAP_TYPE_ANY)
        except:
            print "catch"
            image = wx.Image('asset/first.png', wx.BITMAP_TYPE_ANY)
        #print image
        image = image.Scale(320, 170, wx.IMAGE_QUALITY_HIGH)
        imageBitmap = wx.StaticBitmap(self.rupPan, wx.ID_ANY, wx.BitmapFromImage(image))
        imageBitmap.SetPosition((30, 30))
        #self.hbox_rupPan.Add(imageBitmap, 0, wx.LEFT | wx.RIGHT, 15)
        #image2 = join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))), folder, 'second.png')
        image2 = join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))), folder, 'second.png')
        try:
            image = wx.Image(image2, wx.BITMAP_TYPE_ANY)
        except:
            image = wx.Image('asset/second.png', wx.BITMAP_TYPE_ANY)
        image = image.Scale(320, 170, wx.IMAGE_QUALITY_HIGH)
        imageBitmap = wx.StaticBitmap(self.rupPan, wx.ID_ANY, wx.BitmapFromImage(image))
        imageBitmap.SetPosition((360, 30))
        #self.hbox_rupPan.Add(imageBitmap, 0, wx.LEFT | wx.RIGHT, 15)
        #image3 = join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))), folder, 'third.png')
        image3 = join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))), folder, 'third.png')
        try:
            image = wx.Image(image3, wx.BITMAP_TYPE_ANY)
        except:
            image = wx.Image('asset/third.png', wx.BITMAP_TYPE_ANY)
        image = image.Scale(320, 170, wx.IMAGE_QUALITY_HIGH)
        imageBitmap = wx.StaticBitmap(self.rupPan, wx.ID_ANY, wx.BitmapFromImage(image))
        imageBitmap.SetPosition((690, 30))
        #self.hbox_rupPan.Add(imageBitmap, 0, wx.LEFT | wx.RIGHT, 15)
        #self.vbox_rupPan.Add(self.hbox_rupPan)
        #vbox_rupPan.Add(hbox_rupPan)
        #vbox_rupPan.Add((-1, 20))
        #hbox_rupPan1 = wx.BoxSizer(wx.HORIZONTAL)
        #image4 = join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))), folder, 'fourth.png')
        image4 = join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))), folder, 'fourth.png')
        try:
            image = wx.Image(image4, wx.BITMAP_TYPE_ANY)
        except:
            image = wx.Image('asset/fourth.png', wx.BITMAP_TYPE_ANY)
        image = image.Scale(320, 170, wx.IMAGE_QUALITY_HIGH)
        imageBitmap = wx.StaticBitmap(self.rupPan, wx.ID_ANY, wx.BitmapFromImage(image))
        imageBitmap.SetPosition((30, 210))
        #hbox_rupPan1.Add(imageBitmap, 0, wx.LEFT | wx.RIGHT, 15)
        #image5 = join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))), folder, 'fifth.png')
        image5 = join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))), folder, 'fifth.png')
        try:
            image = wx.Image(image5, wx.BITMAP_TYPE_ANY)
        except:
            image = wx.Image('asset/fifth.png', wx.BITMAP_TYPE_ANY)
        image = image.Scale(320, 170, wx.IMAGE_QUALITY_HIGH)
        imageBitmap = wx.StaticBitmap(self.rupPan, wx.ID_ANY, wx.BitmapFromImage(image))
        imageBitmap.SetPosition((360, 210))
        #hbox_rupPan1.Add(imageBitmap, 0, wx.LEFT | wx.RIGHT, 15)
        #image6 = join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))), folder, 'first.png')
        image6 = join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))), folder, 'first.png')
        try:
            image = wx.Image(image6, wx.BITMAP_TYPE_ANY)
        except:
            image = wx.Image('asset/first.png', wx.BITMAP_TYPE_ANY)
        image = image.Scale(320, 170, wx.IMAGE_QUALITY_HIGH)
        imageBitmap = wx.StaticBitmap(self.rupPan, wx.ID_ANY, wx.BitmapFromImage(image))
        imageBitmap.SetPosition((690, 210))
        #hbox_rupPan1.Add(imageBitmap, 0, wx.LEFT | wx.RIGHT, 15)
        #vbox_rupPan.Add(hbox_rupPan1)
        #hbox_rupPan.Add(canvas, 0, wx.LEFT | wx.RIGHT, 30)
        '''figs = [Figure(figsize=(3, 1.5)) for _ in range(3)]
        axes = [fig.add_subplot(111) for fig in figs]
        canvases = [FigureCanvas(rupPan, -1, fig) for fig in figs]
        for canvas in canvases:
            hbox_rupPan.Add(canvas, 0, wx.LEFT|wx.RIGHT, 30)
        vbox_rupPan.Add(hbox_rupPan)
        vbox_rupPan.Add((-1, 50))

        hbox_rupPan1 = wx.BoxSizer(wx.HORIZONTAL)
        figs = [Figure(figsize=(3, 1.5)) for _ in range(3)]
        axes = [fig.add_subplot(111) for fig in figs]
        canvases = [FigureCanvas(rupPan, -1, fig) for fig in figs]
        for canvas in canvases:
                hbox_rupPan1.Add(canvas, 0, wx.LEFT|wx.RIGHT|wx.TOP, 30)
        vbox_rupPan.Add(hbox_rupPan1)'''
        #self.vbox_rupPan.Add((-1, 20))

        #self.rupPan.SetSizer(self.vbox_rupPan)

        #qw = wx.StaticText(self.rupPan, style=wx.TE_CENTER,size=(600,-1),label="                                                                             Copyright University of South Carolina 2010. All Rights Reserved, JT Morris 6-9-10")
        # vbox_rupPan.Add(qw, 2, wx.EXPAND | wx.ALL, 0)
        # hbox_rdownPan.Add(text_rdownPan, 2, wx.EXPAND | wx.ALL, 0)

        # st1 = wx.StaticText(rupPan, label='North Inlet, SC')
        # st2 = wx.StaticText(rupPan, label='MEM-TLP 6.0')
    def onOpenFile(self, event):
        """
        Create and show the Open FileDialog
        """
        dlg = wx.FileDialog(
            self, message="Choose a file",
            defaultDir=self.currentDirectory,
            defaultFile="",
            wildcard=wildcard,
            style=wx.FD_OPEN | wx.FD_MULTIPLE | wx.FD_CHANGE_DIR
            )
        if dlg.ShowModal() == wx.ID_OK:
            paths = dlg.GetPaths()
            #print "You chose the following file(s):"
            for path in paths:
                #print path
                MainFrame.filePath = path
        dlg.Destroy()
        #wx.Window.Destroy()
        #self.Destroy()
        #self.Close()
        #MainFrame().Show(False)

        #self.Update()
    def onClose(self, event):
        """"""
        #self.Close()
        frame = self.GetParent()
        ##print "hello"
        frame.Destroy()
        #wx.GetApp().Exit()
        #app = wx.App()
        MainFrame().Show()
        #app.MainLoop()
    def onCalculate(self,e):
        MySplash = MySplashScreen()
        MySplash.Show()
        '''ani = wx.animate.Animation("asset/loader.gif")
        ctrl = wx.animate.AnimationCtrl(self, -1, ani)
        ctrl.SetUseWindowBackgroundColour()
        ctrl.Play()
        wx.Yield()'''
        '''hbox_rupPan = wx.BoxSizer(wx.HORIZONTAL)
        vbox_rupPan = wx.BoxSizer(wx.VERTICAL)
        rupPan = wx.Panel(self)
        vbox_rupPan.Add(hbox_rupPan)
        vbox_rupPan.Add((-1, 20))
        hbox_rupPan1 = wx.BoxSizer(wx.HORIZONTAL)
        image = wx.Image('asset/first.png', wx.BITMAP_TYPE_ANY)
        image = image.Scale(350, 200, wx.IMAGE_QUALITY_HIGH)
        imageBitmap = wx.StaticBitmap(rupPan, wx.ID_ANY, wx.BitmapFromImage(image))
        # imageBitmap.SetPosition((1000, 100))
        hbox_rupPan1.Add(imageBitmap, 0, wx.LEFT | wx.RIGHT, 15)
        rupPan.SetSizer(vbox_rupPan)'''
        #self.draw()
        #print "Cal"
        k1 = 0.085
        k2 = 1.99
        troubleshoot = 1
        BGB = [0] * 1800
        SedD = [0] * 1800
        lbgb = [0] * 1800
        dzdt = [0] * 1800
        MSL = [0] * 1800
        inorg = [0] * 600
        sorg = [0] * 600
        soildepth = [0] * 700
        MHW= [0] * 1800
        MHWs = [0] * 1800
        bio= [0] * 1800
        SOM = [0] * 1800
        cquest = [0] * 600
        cmat = [0] * 600
        OMmat= [0] * 1800
        decay = [0] * 1800
        MHWs= [0] * 1800
        marshelev= [0] * 1800
        D= [0] * 1800
        T = [0] * 1800
        IT = [0] * 1800
        ybio = [0] * 1800
        a= [0] * 10
        b= [0] * 10
        c= [0] * 10
        aleft= [0] * 10
        bleft= [0] * 10
        cleft= [0] * 10
        aright= [0] * 10
        bright= [0] * 10
        cright = [0] * 10
        dzdd = [0] * 1800
        dref = [0] * 1800
        dlab = [0] * 1800
        ddcay = [0] * 1800
        bulkd = [0] * 1800
        sedi = [0] * 1800
        droot = [0] * 1800
        totdepth = [0] * 1800
        totBGB = [0] * 1800
        dreftot = [0] * 600
        sorg = [0] * 600
        corelbg = [0] * 141
        coretbg = [0] * 141
        coretin = [0] * 141
        coresom = [0] * 141
        bins = [["" for x in range(405)] for y in range(405)]
        bincounts = [0] * 405
        cohortbins = [0] * 405
        #marshelev=[100]

        ## declaring the needed 2d values in python
        w, h = 26, 1000
        data_list = [["" for x in range(w)] for y in range(h)]
        comp_list = [["" for x in range(w)] for y in range(h)]
        rootdist_list = [["" for x in range(20)] for y in range(200)]
        sheet12_list = [["" for x in range(50)] for y in range(550)]
        sheet10_list = [["" for x in range(1000)] for y in range(1000)]
        num_output_list = [["" for x in range(2000)] for y in range(2000)]
        comp_elev_list = []
        comp_biomass_list = []
        comp_year_list = []
        comp_fourth_biomass_list = []
        comp_msl_list = []
        comp_marshele_list = []
        comp_ind_time = []
        comp_cquest_list = []
        clslr = float(phy_sea_level_forecast.GetValue())
        MSL0 = float(phy_sea_level_start.GetValue())
        SLR100 = float(phy_20th.GetValue())
        RUNL = 100
        j_msl = MSL0
        bsea = (clslr / RUNL - SLR100) / (RUNL - 1)
        asea = SLR100 - bsea

        Tamp = float(phy_MTA.GetValue())
        marshelev[1] = float(phy_Marsh_ele.GetValue()) - (MSL0 - j_msl)
        MHW0 = MSL0 + Tamp
        Trange = 2 * Tamp
        mtsi0 = float(phy_sus_minSed.GetValue())
        tss0 = mtsi0
        mts0 = float(phy_sus_org.GetValue())
        ymax = float(bio_max_peak.GetValue())
        sedload = mtsi0 * 0.000001 * (Tamp + MSL0 - float(phy_Marsh_ele.GetValue())) * 704/2
        orgload = mts0 * 0.000001 * (Tamp + MSL0 - float(phy_Marsh_ele.GetValue())) * 704/2

        if cb1.GetValue() is True:
            res = Tamp + 30
            bio_max_growth.SetLabel(str(res))
            bio_min_growth.SetLabel("-10")
        j2 = 1

        MHW[1] = MHW0
        D[1] = MHW[1] - marshelev[1]

        if cb1.GetValue() is False:
            maxE = float(bio_max_growth.GetValue())
            minE = float(bio_min_growth.GetValue())
            Eopt = float(bio_opt_growth.GetValue())
            if Eopt > maxE :
                Eopt = maxE
                bio_max_growth.SetLabel(str(maxE))
            if Eopt < minE :
                Eopt = minE
                bio_min_growth.SetLabel(str(minE))
            Dopt = Tamp - Eopt
            minD = Tamp - maxE
            maxDleft = Dopt + (Dopt - minD)
            maxD = Tamp - minE
            minDright = Dopt - (maxD - Dopt)
        else:
            minD = -20
            maxD = MHW0 + 20
            maxE = Tamp - minD
            minE = MSL0 + minD
            bio_max_growth.SetLabel(str(round(maxE,0)))
            bio_min_growth.SetLabel(str(round(minE,0)))
            Dopt = (maxD + minD) /2
            bio_opt_growth.SetLabel(str(round((minE+maxE)/2, 0)))
            Eopt = float(bio_opt_growth.GetValue())
            Dopt = Tamp - Eopt
            minD = Tamp - maxE
            maxDleft = Dopt + (Dopt - minD)
            maxD = Tamp - minE
            minDright = Dopt - (maxD - Dopt)
        if cb2.GetValue() is True:
            recovtime = int(float(epi_recoveryTime.GetValue()))
        else:
            recovtime = 1
        for j in reversed(range(1,recovtime+1)):
            i=j
            a[i] = -((-minD * ymax / i - maxD * ymax / i) / ((minD - Dopt) * (-maxD + Dopt)))
            b[i] = -((ymax / i) / ((minD - Dopt) * (-maxD + Dopt)))
            c[i] = (minD * maxD * ymax / i) / ((minD - Dopt) * (maxD - Dopt))
            aleft[i] = -((-minD * ymax / i - maxDleft * ymax / i) / ((minD - Dopt) * (-maxDleft + Dopt)))
            bleft[i] = -((ymax / i) / ((minD - Dopt) * (-maxDleft + Dopt)))
            cleft[i] = (minD * maxDleft * ymax / i) / ((minD - Dopt) * (maxDleft - Dopt))
            aright[i] = -((-minDright * ymax / i - maxD * ymax / i) / ((minDright - Dopt) * (-maxD + Dopt)))
            bright[i] = -(((ymax / i) / i) / ((minDright - Dopt) * (-maxD + Dopt)))
            cright[i] = (minDright * maxD * ymax / i) / ((minDright - Dopt) * (maxD - Dopt))

        omdr = float(bio_OM_decay.GetValue())
        bgmult = float(bio_BGBio.GetValue())
        if D[1] <= Dopt and cb1.GetValue() is False:
            a[1] = aleft[1]
            b[1] = bleft[1]
            c[1] = cleft[1]
        if D[1] > Dopt and cb1.GetValue() is False:
            a[1] = aright[1]
            b[1] = bright[1]
            c[1] = cright[1]
        Drmax = float(bio_max_root_depth.GetValue())
        lna = 1.1
        p = -0.512
        bscale = 0.0001

        '''#print a[1]
        #print b[1]
        #print c[1]
        #print D[1]'''
        bio[1] = (a[1] * D[1] + b[1] * (D[1]*D[1]) + c[1]) * bscale
        #print "-------"
        ##print bio[1]
        if bio[1] < 0:
            bio[1] = 0
        Indtime = min(1, D[1] / Trange)
        Indtime = max(0, D[1] / Trange)
        IT[1] = Indtime
        #w, h = 26, 1000

        for k in range(1,101):
            comp_list[k+1][4] = bio[1] /bscale
            comp_list[k+1][5] = IT[1]
            comp_list[k+1][15] = marshelev[1]
            comp_list[k+1][12] = MSL[1]
        for k in range(2,602):
            comp_list[k][4] = ""
            comp_list[k][8] = ""
            comp_list[k][5] = ""
            comp_list[k][15] = ""
            comp_list[k][12] = ""

        rangex = maxE - minE
        dele = rangex / 80
        for j in range(0,81):
            elev = minE + j * dele
            x = Tamp - elev
            if x<=Dopt and not cb1.GetValue:
                a[1] = aleft[1]
                b[1] = bleft[1]
                c[1] = cleft[1]
            if x>Dopt and not cb1.GetValue:
                a[1] = aright[1]
                b[1] = bright[1]
                c[1] = cright[1]
            ybio[j+1] = (a[1] * x + b[1] * (x*x) + c[1]) * bscale
            if ybio[j+1] < 0:
                ybio[j+1]=0
            comp_list[j][16] = elev
            comp_list[j][17] = ybio[j+1] * 10000
        #print comp_list
        RT = bio[1] * bgmult
        if Drmax > 0:
            Rmax = 2 * RT / Drmax
            kd = Rmax / Drmax
            Rtest = 0.5 * Rmax * Drmax
        LRV = RT / k1
        q = float(model_max_capture.GetValue())
        kr = float(model_refrac.GetValue())
        BGTR = float(bio_BG_turnover.GetValue())
        setvel = 2.8 * Indtime
        if setvel > q:
                inorg[1] = q * tss0 * D[1] * 0.000001 * 704 / 2
                sorg[1] = q * mts0 * D[1] * 0.000001 * 704 / 2
        else:
                inorg[1] = max(setvel * tss0 * D[1] * 0.000001 * 704 / 2, 0)
                sorg[1] = max(setvel * mts0 * D[1] * 0.000001 * 704 / 2, 0)
        minwt = inorg[1]
        cohort_size = inorg[1] / k2 + sorg[1] / k1 + kr * BGTR * RT / k1
        dzm = cohort_size
        krBGTR = kr * BGTR
        #BRZOM = 90 * (sorg[1] / k1 + krBGTR * RT / k1) / (inorg[1] / k2 + sorg[1] / k1 + krBGTR * RT / k1)
        if cb3.GetValue() is True:
                maxvertacc = float(phy_lt.GetValue()) #if phy_lt.GetValue()=='' else float(phy_lt.GetValue())
                krBGTR = (k1 / RT) * (maxvertacc - inorg[1] / k2 - sorg[1] / k1)
                kr = krBGTR / BGTR
                BGTR = min(3, BGTR)
                model_refrac.SetLabel(str(kr))
                #BRZOM = 90 * (sorg[1] / k1 + krBGTR * RT / k1) /   (inorg[1] / k2 + sorg[1] / k1 + krBGTR * RT / k1)
                cohort_size = inorg[1] / k2 + sorg[1] / k1 + krBGTR * RT / k1
                phy_lt.SetLabel(str(round(cohort_size, 2)))
                maxvertacc = cohort_size
                bio_BG_turnover.SetLabel(str(BGTR))

        # changed here
        #Bden = (sorg[1] + krBGTR * RT + inorg[1]) / ((sorg[1] + krBGTR * RT) / k1 + inorg[1] / k2)
        Bden = 1
        # check the root distribution **********

        rootdist_list[3][2] = RT
        rootdist_list[4][2] = kd
        rootdist_list[5][2] = Rmax
        rootdist_list[6][2] = Drmax
        rootdist_list[7][2] = kr
        rootdist_list[8][2] = omdr
        rootdist_list[9][2] = BGTR

        #'************************************************
        if minwt == 0 and D[1] < 0:
            dzm = (sorg[1] + krBGTR * RT) / k1
        if minwt == 0 and cb3.GetValue() is True:
            dzm = maxvertacc
        nocohort = 500

        for ico in (1,nocohort+1):
            dzdd[ico] = dzm
            sedi[ico] = minwt
            dref[ico] = sorg[1] + krBGTR * RT
            bulkd[ico] = Bden
            #BD = dzdd[ico] / dref[ico]
        lstroot = droot[500]
        lstref = dref[500]
        lstlab = 0
        if bio[1] == 0 and D[1] < 0:
            Scenario = 1
        if bio[1] > 0 and D[1] < 0:
            Scenario = 2
        if bio[1] > 0 and D[1] > 0:
            Scenario = 3
        if sedload > 0 and D[1] > 0 and bio[1] <= 0:
            Scenario = 4


        if Scenario == 1: #this is the case where the surface is out of the water and above the growth zone
            for ico in (1, nocohort+1):
                inorg[1] = 0
                dzdd[ico] = 0.2
                soildepth[ico] = dzdd[ico]
                dlab[ico] = 0
                droot[ico] = 0
                OMmat[ico] = BRZON
                if OMmat(ico) > 0:
                    bulkd[ico] = 1 / (0.01 * OMmat[ico] / k1 + (1 - 0.01 * OMmat[ico]) / k2)
                    sedi[ico] = (1 - 0.01 * OMmat[ico]) * k2
                    dref[ico] = 0.01 * OMmat[ico] * k1 # g/g x g/cm3
                elif bulkd(ico) > k2:
                    bulkd[ico] = k2#bulk density below the root zone
                    OMmat[ico] = 0
                    sedi[ico] = dzdd[ico] * bulkd[ico]
                    dref[ico] = 0
        elif Scenario == 2:
            Bot = Drmax - dzm
            droot[500] = 0.5 * (Rmax / Drmax) * ((Drmax*Drmax) - (Bot * Bot))
            dlab[500] = (1 - kr) * (BGTR * droot[500])
            decay[500] = -dlab[500] * omdr
            dlab[500] = dlab[500] - decay[500]
            dref[500] = (krBGTR * droot[500]) + sorg[1] + lstref
            tdz = (droot[500] + dlab[500] + dref[500]) / k1
            dzdd[500] = tdz
            soildepth[500] = dzdd[500]
            BGB[500] = dref[500] + dlab[500] + droot[500]
            OMmat[500] = 90 * BGB[500] / (sedi[500] + BGB[500])
            bulkd[500] = 1 / (0.01 * OMmat[500] / k1 + (1 - 0.01 * OMmat[500]) / k2)
            Top = Drmax - tdz
            for ico in reversed(range(1, nocohort)):
                lstref = dref[500]
            if troubleshoot==1:
                sheet12_list[2][9] = soildepth[500]
                sheet12_list[2][10] = OMmat[500]
                sheet12_list[2][11] = BGB[500]
                sheet12_list[2][12] = sedi[500]
                sheet12_list[2][13] = dref[500]
                sheet12_list[2][14] = droot[500]
                sheet12_list[2][15] = dlab[500]
                sheet12_list[2][16] = dzdd[500]
            for ico in reversed(range(1, nocohort)):
                Bot = Top - dzm
                rsection = 0.5 * (Rmax / Drmax) * ((Top * Top) - (Bot * Bot))
                if Top < 0:
                    rsection = 0
                if Bot < 0 and Top > 0:
                    rsection = 0.5 * (Rmax / Drmax) * (Top * Top)
                droot[ico] = max(0, rsection)
                dlab[ico] = lstlab + (BGTR * droot[ico]) * (1 - kr)
                decay[ico] = max(0, -1 * dlab[ico] * omdr)
                dlab[ico] = dlab[ico] - decay[ico]
                dref[ico] = lstref + krBGTR * droot[ico]
                tdz = (droot[ico] + dlab[ico] + dref[ico]) / k1
                soildepth[ico] = dzdd[ico]
                BGB[ico] = dref[ico] + dlab[ico] + droot[ico]
                OMmat[ico] = 90 * BGB[ico] / [sedi[ico] + BGB[ico]]
                bulkd[ico] = 1 / (0.01 * OMmat[ico] / k1 + (1 - 0.01 * OMmat[ico]) / k2)
                dzdd[ico] = tdz
                soildepth[ico] = dzdd[ico]
                Top = Top - dzdd[ico]
                Bot = Top - dzm
                Start = nocohort - 2
                lstref = dref[ico]
                lstlab = dlab[ico]
            ico = 1
            if troubleshoot==1:
                ik = 2
                for ico in reversed(range(1, nocohort+1)):
                    sheet12_list[ik][9] = soildepth[ico]
                    sheet12_list[ik][10] = OMmat[ico]
                    sheet12_list[ik][11] = BGB[ico]
                    sheet12_list[ik][12] = sedi[ico]
                    sheet12_list[ik][13] = dref[ico]
                    sheet12_list[ik][14] = droot[ico]
                    sheet12_list[ik][15] = dlab[ico]
                    sheet12_list[ik][16] = dzdd[ico]
                    ik = ik + 1
        elif Scenario == 3:
            if bio[1] == 0 and D[1] < 0:
                Bot = Drmax - (dref[500] + sorg[1]) / k1
            else:
                Bot = Drmax - dzm
            tdzl = dzm
            delroot = 0
            displacedroot = 0
            for kk in range(1,101):
                droot[500] = 0.5 * (Rmax / Drmax) * ((Drmax * Drmax) - (Bot * Bot))  # area under the tdz segment
                dlab[500] = (1 - kr) * (BGTR * droot[500])
                decay[500] = -dlab[500] * omdr  # total decay in cohort 1 in 1 year
                dlab[500] = dlab[500] - decay[500]
                dref[500] = (krBGTR * droot[500]) + sorg[1]
                tdz = max((dref[ico] + sorg[1]) / k1, (droot[500] + dlab[500] + dref[500]) / k1 + minwt / k2)  # dzm is the mineral fraction and does not change
                dzdd[500] = tdz
                Bot = Drmax - tdz
                if abs(tdz - tdzl) < 0.000001:
                    break
                tdzl = tdz
            lstdzdd = dzdd[500]
            soildepth[500] = dzdd[500]
            BGB[500] = dref[500] + dlab[500] + droot[500]
            OMmat[500] = 90 * BGB[500] / (sedi[500] + BGB[500])
            bulkd[500] = 1 / (0.01 * OMmat[500] / k1 + (1 - 0.01 * OMmat[500]) / k2)
            Top = Drmax - tdz
            Bot = Top - tdz
            if troubleshoot==1:
                sheet12_list[2][ 9] = soildepth[500]
                sheet12_list[2][10] = OMmat[500]
                sheet12_list[2][ 11] = BGB[500]
                sheet12_list[2][ 12] = sedi[500]
                sheet12_list[2][ 13] = dref[500]
                sheet12_list[2][ 14] = droot[500]
                sheet12_list[2][ 15] = dlab[500]
                sheet12_list[2][ 16] = dzdd[500]
            for ico in reversed(range(1, nocohort)):
                lstroot = droot[ico + 1]
                lstref = 0.5 * (Rmax / Drmax) * ((Drmax * Drmax) - (Top * Top)) * krBGTR + sorg[1]
                if Top < 0:
                    lstref = 0.5 * (Rmax / Drmax) * (Drmax * Drmax) * krBGTR + sorg[1]
                lstlab = dlab[ico + 1]
                for kk in range(1,51):
                    rsection = 0.5 * (Rmax / Drmax) * ((Top * Top) - (Bot * Bot))
                    if Top < 0 or Bot < 0:
                        rsection = 0
                    droot[ico] = max(0, rsection)
                    dlab[ico] = lstlab + (BGTR * droot[ico]) * (1 - kr)#+ (1 - kr) * delroot
                    decay[ico] = max(0, -1 * dlab[ico] * omdr)
                    dlab[ico] = dlab[ico] - decay[ico]
                    dref[ico] = lstref + krBGTR * droot[ico]
                    tdz = (droot[ico] + dlab[ico] + dref[ico]) / k1 + minwt / k2
                    tdz = max(lstdzdd, tdz)
                    Bot = Top - tdz
                    if abs(tdz - tdzl) < 0.00001:
                        break
                    tdzl = tdz
                soildepth[ico] = dzdd[ico]
                BGB[ico] = dref[ico] + dlab[ico] + droot[ico]
                ##print sedi[ico]
                ##print BGB[ico]
                OMmat[ico] = 90 * BGB[ico] / (sedi[ico] + BGB[ico])
                bulkd[ico] = 1 / (0.01 * OMmat[ico] / k1 + (1 - 0.01 * OMmat[ico]) / k2)
                dzdd[ico] = tdz
                soildepth[ico] = dzdd[ico]
                Top = Top - dzdd[ico]
                Bot = Top - tdz
            ico = 1
            if troubleshoot==1:
                ik = 2
            for ico in reversed(range(1, nocohort+1)):
                    sheet12_list[ik][9] = soildepth[ico]
                    sheet12_list[ik][10] = OMmat[ico]
                    sheet12_list[ik][11] = BGB[ico]
                    sheet12_list[ik][12] = sedi[ico]
                    sheet12_list[ik][13] = dref[ico]
                    sheet12_list[ik][14] = droot[ico]
                    sheet12_list[ik][15] = dlab[ico]
                    sheet12_list[ik][16] = dzdd[ico]
                    ik = ik + 1
        elif Scenario == 4:
            D1 = 0
            d2 = dzdd[nocohort]
            for k in range(1,16):
                droot[nocohort] = 0  # annual root production in cohort 1
                dref[nocohort] = orgload * q  # refractory input in cohort 1
                dlab[nocohort] = 0  # labile remaining in cohort 1
                decay[nocohort] = 0  # total decay in cohort 1 in 1 year
                BGB[nocohort] = dref[nocohort]
                if q * sedload + BGB[nocohort] > 0 :
                    OMmat[nocohort] = 90 * BGB[nocohort] / (q * sedload + BGB[nocohort]) # percent OM in cohort 1
                if OMmat[nocohort] > 0 :
                    bulkd[nocohort] = 1 / (0.01 * OMmat[nocohort] / k1 + (1 - 0.01 * OMmat[nocohort]) / k2)
                if bulkd[nocohort] > k2 :
                    bulkd[nocohort] = k2
                dzdd[nocohort] = (sedi[nocohort] + BGB[nocohort]) / bulkd[nocohort]
                if abs(d2 - dzdd[nocohort]) < 0.0001:
                    break
                d2 = D1 + dzdd[nocohort]
            D1 = d2
            for ico in reversed(range(1, nocohort)):
                sedi[ico] = q * sedload
                d2 = D1 + dzdd[ico]
                for k in range(0,16):
                    droot[ico] = 0  # annual root production in cohort 1
                    dref[ico] = dref[ico + 1]  # refractory input in cohort 1
                    dlab[ico] = 0  # labile remaining in cohort 1
                    decay[ico] = 0  # total decay in cohort 1 in 1 year
                    BGB[ico] = dref[ico]
                    if q * sedload + BGB[ico] > 0:
                        OMmat[ico] = 90 * BGB[ico] / (q * sedload + BGB[ico])  # percent OM in cohort 1
                    if OMmat[ico] > 0:
                        bulkd[ico] = 1 / (0.01 * OMmat[ico] / k1 + (1 - 0.01 * OMmat[ico]) / k2)
                    if bulkd[ico] > k2:
                        bulkd[ico] = k2
                    dzdd[ico] = (sedi[ico] + BGB[ico]) / bulkd[ico]
                    if abs(d2 - dzdd[ico]) < 0.001:
                        break
                    d2 = D1 + dzdd[ico]
                D1 = D1+dzdd[ico]
        inorg[1] = sedload  # annual sediment load [g cm-2 yr-1]
        sheet10_list[1][ 1] = "cohort"
        sheet10_list[1][ 2] = "depth"
        sheet10_list[1][ 3] = "sedload"
        sheet10_list[1][ 4] = "droot"
        sheet10_list[1][ 5] = "dref"
        sheet10_list[1][ 6] = "dlab"
        sheet10_list[1][ 7] = "bulkd"
        sheet10_list[1][ 8] = "%OMmat"
        sheet10_list[1][ 9] = "decay"

        k = 0
        dztot = 0
        for ico in reversed(range(1, nocohort+1)):
            k = 502 - ico
            dztot = dztot + dzdd[ico]
            ##print k
            sheet10_list[k][ 1] = ico
            sheet10_list[k][ 2] = dztot  # this is depth
            sheet10_list[k][ 3] = sedload
            sheet10_list[k][ 4] = droot[ico] * 1000  # output mg/cm2
            sheet10_list[k][ 5] = dref[ico] * 1000
            sheet10_list[k][ 6] = dlab[ico] * 1000
            sheet10_list[k][ 7] = bulkd[ico]
            sheet10_list[k][ 8] = OMmat[ico]
            sheet10_list[k][ 9] = decay[ico] * 1000
        tlbgbio = 0
        totsom = 0
        totdepth[1] = 0
        tlabbio0 = 0
        for k in range(3, 1801):
            num_output_list[k][ 2] = ""
            num_output_list[k][ 3] = ""
            num_output_list[k][ 4] = ""
            num_output_list[k][ 5] = ""
            num_output_list[k][ 6] = ""
            num_output_list[k][ 7] = ""
            num_output_list[k][ 8] = ""
            num_output_list[k][ 9] = ""
            num_output_list[k][ 10] = ""
            num_output_list[k][ 11] = ""
            num_output_list[k][ 12] = ""
            num_output_list[k][ 13] = ""
            num_output_list[k][ 14] = ""
            num_output_list[k][ 15] = ""
            num_output_list[k][ 16] = ""
            num_output_list[k][ 17] = ""
            num_output_list[k][ 18] = ""
            num_output_list[k][ 19] = ""
            num_output_list[k][ 20] = ""
        kk = 2
        for ico in reversed(range(1, 501)):
            tlbgbio = tlbgbio + droot[ico]
            totsom = totsom + BGB[ico]
            totdepth[1] = totdepth[1] + dzdd[ico]
            kk = kk + 1
            num_output_list[kk][ 5] = round(dzdd[ico], 4)
            num_output_list[kk][ 6] = totdepth[1]
            num_output_list[kk][ 7] = droot[ico] * 10000  # g/m2
            num_output_list[kk][ 8] = dlab[ico] * 10000
            num_output_list[kk][ 9] = dref[ico] * 10000
            num_output_list[kk][ 10] = BGB[ico] * 10000
            num_output_list[kk][ 11] = OMmat[ico]
            num_output_list[kk][ 12] = sedi[ico]
            num_output_list[kk][ 13] = bulkd[ico]
            if dzdd[ico] > 0:
                num_output_list[kk][14] = droot[ico] * 1000 / dzdd[ico]
            num_output_list[kk][ 15] = decay[ico] * 10000
        for k in range(3, 501):
            num_output_list[k][ 16] = ""
            num_output_list[k][ 17] = ""
            num_output_list[k][ 18] = ""
            num_output_list[k][ 19] = ""
            num_output_list[k][ 20] = ""

        bin = 0
        for i in range(1,41):
            bincounts[i] = 0
            for j in range(1, 401):
                bins[i][ j] = 0
        for i in range(1,141):
            corelbg[i] = 0
            coretbg[i] = 0
            coretin[i] = 0
        counts = 1
        cohortBot = 0
        Top = 2.5
        jlast = 1
        for i in range(1,41):
            tprop = 0
            for j in range(1, 401):
                jlast = jlast + 1
                cohortTop = num_output_list[j + 2][ 6]
                if cohortTop > Top:
                    break
                prop = (cohortTop - cohortBot) / 2.5
                tprop = tprop + prop
                cohortBot = cohortTop
                #If bins(i, j) = 1 Then
                corelbg[i] = corelbg[i] + num_output_list[j + 2][7] * prop# add lbg
                coretbg[i] = coretbg[i] + num_output_list[j + 2][ 10] * prop
                coretin[i] = coretin[i] + num_output_list[j + 2][ 12] * prop
            jlast = jlast - 1
            cohortBot = Top
            Top = Top + 2.5
            if Top > 100:
                break
        k = 3
        for i in range(1, 41):
            if (i * 2.5) > 100:
                break
            num_output_list[k][ 16] = i * 2.5
            num_output_list[k][ 17] = corelbg[i]
            num_output_list[k][ 18] = coretbg[i]
            num_output_list[k][ 19] = 10000 * coretin[i]
            if coretbg[i] + coretin[i] > 0:
                num_output_list[k][ 20] = 90 * coretbg[i] / (coretbg[i] + 10000 * coretin[i])
            k = k + 1
        for k in range(2,501):
            comp_list[k][1] = " "
            comp_list[k][2] = " "
            comp_list[k][3] = " "
            comp_list[k][4] = " "
            comp_list[k][5] = " "
            comp_list[k][6] = " "
            comp_list[k][7] = " "
            comp_list[k][8] = " "
            comp_list[k][9] = " "
            comp_list[k][10] = " "
            comp_list[k][11] = " "
            comp_list[k][12] = " "
            comp_list[k][13] = " "
            comp_list[k][14] = " "
            comp_list[k][15] = " "
            #comp_list[k][16] = " "
            comp_list[k][19] = " "
            comp_list[k][20] = " "
            comp_list[k][21] = " "
            comp_list[k][22] = " "
            comp_list[k][23] = " "
        irecov = 1
        mtsi = mtsi0
        jt = 5
        thintime = float(epi_years.GetValue())
        for jtime in range(1,101):
            deadbymove = 0
            MSL[jtime] = j_msl + asea * jtime + bsea * (jtime * jtime)
            MHW[jtime] = Tamp + MSL[jtime] + lna * math.sin(2 * 3.14159265 * jtime / 18.6 + p)
            D[jtime] = MHW[jtime] - marshelev[jtime]
            if cb2.GetValue() is True and jtime == thintime:
                D[jtime] = D[jtime] - float(epi_addElevation)
                marshelev[jtime] = marshelev[jtime] + float(epi_addElevation)
            Trange = (MHW[jtime] - MSL[jtime]) * 2
            if cb2.GetValue()==True:
                if jtime == thintime:
                    irecov = recovtime
                if jtime - thintime > 0:
                    irecov = irecov - 1
                if irecov < 1:
                    irecov = 1
            if D[jtime] <= Dopt:
                a[irecov] = aleft[irecov]
                b[irecov] = bleft[irecov]
                c[irecov] = cleft[irecov]
            else:
                a[irecov] = aright[irecov]
                b[irecov] = bright[irecov]
                c[irecov] = cright[irecov]
            bio[jtime] = (a[irecov] * D[jtime] + b[irecov] * (D[jtime] * D[jtime]) + c[irecov]) * bscale
            if bio[jtime] < 0:
                bio[jtime] = 0
            Indtime = D[jtime] / Trange
            if Indtime <= 0:
                Indtime = 0
            if Indtime > 1:
                Indtime = 1
            IT[jtime] = Indtime
            sedload = mtsi * D[jtime] * 704 * 0.000001 / 2
            orgload = mts0 * D[jtime] * 704 * 0.000001 / 2
            if D[jtime] < 0:
                sedload = 0
                orgload = 0
            setvel = q * Indtime
            if setvel > 1:
                inorg[jtime] = sedload
                sorg[jtime] = orgload
            else:
                inorg[jtime] = max(setvel * sedload, 0)
                sorg[jtime] = max(setvel * orgload, 0)
            newco = 0
            if inorg[jtime] + sorg[jtime] > 0:
                newco = 1
                nocohort = nocohort + 1  # add a new cohort if there is some sedimentation
                sedi[nocohort] = inorg[jtime]

            if cb2.GetValue()==True and jtime == thintime:
                sedi[nocohort] = sedi[nocohort] + float(epi_recoveryTime.GetValue()) * k2  # read in thin layer
                inorg[jtime] = inorg[jtime] + float(epi_recoveryTime.GetValue()) * k2  # add the thin layer amount to the inorganic load
                #    D[jtime] = D[jtime] - Cells[38, 2] this is done later
                newco = 1

            RT = bio[jtime] * bgmult  # RT is the total live belowground biomass
            Rmax = 2 * RT / Drmax
            # Rmax = RT / [Drmax + kd * Log[kd / [Drmax + kd]]] # asymptotic root biomass [units are per cm2]
            # RT = bio[jtime] * bgmult  # comput current root distribution
            # kd = Log[0.05] / Drmax
            # Rof = -0.95 * RT * kd / [1 - Exp[kd * Drmax]]
            # mult root biomass by the turnover rate and refractory fraction to get the C sequestration
            # but also add the roots that turnover by vertical displacement
            cquest[jtime] = krBGTR * RT + sorg[jtime]
            if newco == 1:
                # tdz is the initial dimension of the new cohort
                tdz = sedi[nocohort] / k2 + sorg[jtime] / k1  # inorg is the natural input of mineral sediment at jtime
                # if adding sediment layer, increase tdz
                lstroot = 0
                lstref = 0
                lstlab = 0
                dzm = tdz  # here tdz is just from the mineral input
            # droot[nocohort] = 0
            if newco == 0:
                tdz = dzdd[nocohort]  # use the previous dimension if there is not a new cohort
                lstroot = droot[nocohort]
                lstref = dref[nocohort]
                lstlab = dlab[nocohort]
            Top = Drmax
            Bot = Drmax - tdz
            tdzl = tdz
            delroot = 0
            for kk in range(1,101):
                droot[nocohort] = 0.5 * (Rmax / Drmax) * ((Drmax*Drmax) - (Bot*Bot))  # area under the tdz segment
                dlab[nocohort] = lstlab + (1 - kr) * (BGTR * droot[nocohort])
                decay[nocohort] = -dlab[nocohort] * omdr  # total decay in cohort 1 in 1 year
                dlab[nocohort] = dlab[nocohort] - decay[nocohort]
                dref[nocohort] = lstref + (krBGTR * droot[nocohort]) + sorg[jtime]
                tdz = (droot[nocohort] + dlab[nocohort] + dref[nocohort]) / k1 + inorg[jtime] / k2
                # account for displaced roots, positive delroot is root lost from the cohort
                # dzm should be the cohort dimension created by sediment input
                dzdd[nocohort] = tdz  # here tdz is from mineral and organic inputs
                Bot = Drmax - tdz
                if abs(tdz - tdzl) < 0.0001:
                    break
                tdzl = tdz
            tdztest = tdz
            soildepth[nocohort] = dzdd[nocohort]
            BGB[nocohort] = dref[nocohort] + dlab[nocohort] + droot[nocohort]
            OMmat[nocohort] = 90 * BGB[nocohort] / (sedi[nocohort] + BGB[nocohort])
            bulkd[nocohort] = 1 / (0.01 * OMmat[nocohort] / k1 + (1 - 0.01 * OMmat[nocohort]) / k2)
            # this is the death of roots by displacement
            if cb4.GetValue()==True and thintime == jtime:
                deadbymove = 0.5 * (Rmax / Drmax) * ((Drmax*Drmax) - (Bot*Bot))
                deadbymove = max(0, deadbymove)
            # cquest[jtime] = cquest[jtime] + dref[nocohort]
            # deadbymove = delroot
            # ico is the cohort number starting with nocohort at the top of the stack
            if troubleshoot==1:
                sheet12_list[2][1] = soildepth[nocohort]
                sheet12_list[2][2] = OMmat[nocohort]
                sheet12_list[2][3] = BGB[nocohort]
                sheet12_list[2][4] = sedi[nocohort]
                sheet12_list[2][5] = dref[nocohort]
                sheet12_list[2][6] = droot[nocohort]
                sheet12_list[2][7] = dlab[nocohort]
                sheet12_list[2][8] = dzdd[nocohort]
            Top = Drmax - tdz  # the top of the next cohort
            for ico in reversed(range(1, nocohort )):
                Bot = Top - dzdd[ico]  # bottom of the next cohort
                lstroot = droot[ico]  # this is the root biomass in this cohort in the previous year
                lstlab = dlab[ico]
                lstref = dref[ico]
                # tdz = (dlab[ico] + dref[ico]) / k1 + sedi[ico] / k2
                # Bot = Top - tdz # subtract the mineral dimension
                for kk in range(1, 51):
                    # WR to the cohorts, the top cohort is at elevation Drmax relative to root dist
                    rsection = 0.5 * (Rmax / Drmax) * ((Top*Top) - (Bot *Bot))  # area under the tdz segment
                    if Top < 0:
                        rsection = 0.5 * (Rmax / Drmax) * (Top*Top)  # area under the tdz segment
                    if Bot < 0:
                        rsection = 0

                    droot[ico] = max(0, rsection)  # rsection is root biomass in the top to bottom section
                    # delroot = max(0, lstroot - droot[ico])
                    # If Not CheckBox4 Then delroot = max(0, delroot - BGTR * lstroot) # zero delroot if it is less than turnover
                    # If CheckBox4 Then delroot = lstroot - droot[ico]
                    # droot[ico] = Rmax * (bot - (0.5 * (bot ^ 2#) / Drmax)) - Rmax * (Top - (0.5 * (Top ^ 2#) / Drmax))
                    dlab[ico] = lstlab + (BGTR * droot[ico]) * (1 - kr)  # + (1 - kr) * delroot
                    decay[ico] = max(0, -1 * dlab[ico] * omdr)  # total decay in cohort 1 in 1 year (omdr is negative)
                    dlab[ico] = dlab[ico] - decay[ico]
                    # add the refractory roots that turnover by vertical displacement (delroot)
                    dref[ico] = lstref + krBGTR * droot[ico]  # + kr * delroot
                    tdz = (droot[ico] + dlab[ico] + dref[ico]) / k1 + sedi[ico] / k2
                    Bot = Top - tdz
                    if abs(tdz - tdzl) < 0.00001:
                        break
                    tdzl = tdz  # tdz is the size of the top cohort
                soildepth[ico] = soildepth[ico + 1] + dzdd[ico]
                BGB[ico] = dref[ico] + dlab[ico] + droot[ico]
                OMmat[ico] = 90 * BGB[ico] / (sedi[ico] + BGB[ico])
                bulkd[ico] = 1 / (0.01 * OMmat[ico] / k1 + (1 - 0.01 * OMmat[ico]) / k2)
                dzdd[ico] = tdz
                Top = Bot
            if troubleshoot==1:
                ik=2
                for ico in reversed(range(400, nocohort+1)):
                    sheet12_list[ik][1] = soildepth[ico]
                    sheet12_list[ik][2] = OMmat[ico]
                    sheet12_list[ik][3] = BGB[ico]
                    sheet12_list[ik][4] = sedi[ico]
                    sheet12_list[ik][5] = dref[ico]
                    sheet12_list[ik][6] = droot[ico]
                    sheet12_list[ik][7] = dlab[ico]
                    sheet12_list[ik][8] = dzdd[ico]
                    ik = ik + 1
            if cb4.GetValue()==True:
                cquest[jtime] = cquest[jtime] + kr * deadbymove
            jt = jt + 5
            comp_list[jtime + 1][ 11] = OMmat[nocohort - 50]
            # Sheet12.Cells[jtime + 1, 7] = deadbymove * 10000
            BDL = bulkd[ico]

            tlbgbio = 0
            totsom = 0
            totdepth[jtime + 1] = 0

            for ico in reversed(range(1, nocohort+1)):
                totdepth[jtime + 1] = totdepth[jtime + 1] + dzdd[ico]
                tlbgbio = tlbgbio + droot[ico]
                totsom = totsom + BGB[ico]


            lbgb[jtime] = tlbgbio
            totBGB[jtime] = totsom
            dzdt[jtime] = (totdepth[jtime + 1] - totdepth[jtime])
            marshelev[jtime + 1] = marshelev[jtime] + dzdt[jtime]
            if cb2.GetValue() == True and jtime == thintime + recovtime:
                thintime = thintime + epi_repeat# next thin layer appl
            k = jtime + 1
            comp_list[k][1] = jtime
            comp_list[k][2] = MHW[jtime]
            comp_list[k][3] = dzdt[jtime]
            comp_list[k][4] = bio[jtime] * 10000
            comp_list[k][5] = round(IT[jtime],2)
            comp_list[k][6] = sorg[jtime]
            comp_list[k][8] = round(cquest[jtime] * 0.42 * 10000,2)
            comp_list[k][9] = sorg[jtime]
            comp_list[k][10] = totBGB[jtime]
            comp_list[k][12] = MSL[jtime]
            comp_list[k][13] = inorg[jtime]
            if jtime == 1:
                comp_list[k][ 14] = MSL[jtime] - MSL0
            else:
                comp_list[k][ 14] = MSL[jtime] - MSL[jtime - 1]
            comp_list[k][15] = round(marshelev[jtime],2)
            #comp_list[k][16] = round(totdepth[jtime],2)
            comp_list[k][19] = round(lbgb[jtime],2)
            comp_list[k][20] = bulkd[nocohort - 50]
            comp_list[k][21] = 10000 * droot[nocohort]
            comp_list[k][22] = dzdd[nocohort]
            comp_list[k][23] = deadbymove
            if jtime == 50:
                totC50 = 0
                for j in reversed(range(nocohort-50, nocohort+1)):
                    totC50 = totC50 + 0.42 * BGB[j] * 10000
            if jtime == 100:
                totC100 = 0
                for j in reversed(range(nocohort - 50, nocohort + 1)):
                    totC100 = totC100 + 0.42 * BGB[j] * 10000
        jtime = jtime - 1
        k = 2
        totd = 0
        totsom = 0
        tlbgbio = 0
        acr50 = 0
        acr100 = 0
        refracC100 = 0
        refracC50 = 0
        for j in range(1,51):
            acr50 = acr50 + dzdt[j]
            refracC50 = refracC50 + cquest[j] * 0.42 * 10000
        for j in range(50,101):
            acr100 = acr100 + dzdt[j]
            refracC100 = refracC100 + cquest[j] * 0.42 * 10000
        totC50 = totC50 / 50
        totC100 = totC100 / 50
        acr50 = acr50 / 50
        acr100 = acr100 / 50
        data_list[26][ 9] = round((marshelev[100] - marshelev[50]) / 50, 2)
        data_list[27][ 9] = round(refracC100 / 50, 1)
        data_list[28][ 9] = round(totC100, 1)
        data_list[30][ 9] = round((marshelev[51] - marshelev[1]) / 50, 2)
        data_list[31][ 9] = round(refracC50 / 50, 1)
        data_list[32][ 9] = round(totC50, 1)
        for ico in reversed(range(1, nocohort+1)):
            k = k + 1
            totd = totd + dzdd[ico]
            tlbgbio = tlbgbio + droot[ico]
            totsom = totsom + BGB[ico]
            num_output_list[k][5] = round(dzdd[ico],4)
            num_output_list[k][6] = totd
            num_output_list[k][7] = droot[ico] * 10000
            num_output_list[k][8] = dlab[ico] * 10000
            num_output_list[k][9] = dref[ico] * 10000
            num_output_list[k][10] = BGB[ico] * 10000
            if ico > 550:
                totC50 = totC50 + 0.42 * BGB[ico] * 10000
            num_output_list[k][ 11] = OMmat[ico]
            num_output_list[k][ 12] = sedi[ico]
            num_output_list[k][ 13] = bulkd[ico]
            if dzdd[ico] > 0:
                num_output_list[k][ 1] = droot[ico] * 1000 / dzdd[ico]
            else:
                num_output_list[k][ 14] = 0
            num_output_list[k][ 15] = decay[ico] * 10000
        for j in range(1,101):
            #'marshelev(j + 1) = marshelev(j) + totdepth(j + 1) - totdepth(j)
            num_output_list[j + 2][ 1] = j
            #' years before present
            num_output_list[2 + j][ 2] = MSL[j]
            num_output_list[2 + j][ 3] = round(marshelev[j], 1)
            num_output_list[2 + j][ 4] = bio[j] * 10000
        for i in range(1,141):
            corelbg[i] = 0
            #' fill the 2.5 cm bins with zeros
            coretbg[i] = 0
            coretin[i] = 0
        counts = 1

        cohortBot = 0
        Top = 2.5
        jlast = 1
        for i in range(1, 141):
            #' for each 2.5 cm section
            tprop = 0
            for j in range(jlast, 401):
                jlast = jlast + 1
                cohortTop = num_output_list[j + 2][ 6]
                if cohortTop > Top:
                    break
                #'slice = WorksheetFunction.Max(2.5, Top - Bot)
                prop = (cohortTop - cohortBot) / 2.5
                tprop = tprop + prop
                cohortBot = cohortTop
                #'If bins(i, j) = 1 Then
                corelbg[i] = corelbg[i] + num_output_list[j + 2][ 7] * prop
                #' add lbg
                coretbg[i] = coretbg[i] + num_output_list[j + 2][ 10] * prop
                coretin[i] = coretin[i] + num_output_list[j + 2][ 12] * prop

            jlast = jlast - 1
            cohortBot = Top
            Top = Top + 2.5
            if Top > 100:
                break
        k=45
        for i in range(1, 141):
            if i * 2.5 > 100:
                break
            num_output_list[k][ 16] = i * 2.5
            num_output_list[k][ 17] = corelbg[i]
            num_output_list[k][ 18] = coretbg[i]
            num_output_list[k][ 19] = 10000 * coretin[i]
            if coretin[i] > 0:
                num_output_list[k][ 20] = 90 * coretbg[i] / (coretbg[i] + 10000 * coretin[i])
            k = k + 1
        tlbgbio = 0
        totsom = 0
        SedD[1] = dzdd[1]

        for i in range(0, len(comp_list)):
            for j in range(0, len(comp_list[i])):
                compGrid.SetCellValue(i, j, str(comp_list[i][j]))
        #for i in range(0, len(num_output_list)):
            #for j in range(0, len(num_output_list[i])):
                #numGrid.SetCellValue(i, j, str(num_output_list[i][j]))
        for i in range(0, len(sheet10_list)):
            for j in range(0, len(sheet10_list[i])):
                sheet10Grid.SetCellValue(i, j, str(sheet10_list[i][j]))
        for i in range(0, len(sheet12_list)):
            for j in range(0, len(sheet12_list[i])):
                sheet12Grid.SetCellValue(i, j, str(sheet12_list[i][j]))
        for i in range(0, len(rootdist_list)):
            for j in range(0, len(rootdist_list[i])):
                rootdistGrid.SetCellValue(i, j, str(rootdist_list[i][j]))
        #print comp_list
        #comp_year_list.append(0)
        #comp_ind_time.append(0.06)
        for i in range(1, 78):
            comp_elev_list.append(comp_list[i][16])
        for i in range(3, 80):
            comp_biomass_list.append(comp_list[i][17])
        for i in range(2, 102):
            #for j in range(0, len(comp_list[i])):
                #comp_elev_list.append(comp_list[i][16])
                #comp_biomass_list.append(comp_list[i][17])

                comp_year_list.append(comp_list[i][1])
                comp_ind_time.append(comp_list[i][5])
                comp_cquest_list.append(comp_list[i][8])
                comp_fourth_biomass_list.append(comp_list[i][4])
                comp_msl_list.append(comp_list[i][12])
                comp_marshele_list.append(comp_list[i][15])
       # for i in range(2, 102):
               # comp_fourth_biomass_list.append(comp_list[][])
        #print comp_year_list
        #print comp_biomass_list
        matplotlib.rc('xtick',labelsize=15)
        matplotlib.rc('ytick', labelsize=15)
        plt.rc('axes', labelsize=15)
        plt.axis([0, 100, 0, 200])
        plt.plot(comp_year_list, comp_msl_list, comp_year_list, comp_marshele_list)
        plt.xlabel("time (yrs)")
        plt.ylabel("cm NAVD")
        ax = plt.gca()

        ax.set_facecolor('#edeeff')
        #ax.set_facecolor('red')
        #plt.xticks([1, 20, 40, 60, 80, 100], ['0', '20', '40', '60', '80', '100'])
        #image1 = join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))), folder, 'first.png')
        '''try:

            image = wx.Image(image1, wx.BITMAP_TYPE_ANY)
        except IOError:
            image = wx.Image('asset/first.png', wx.BITMAP_TYPE_ANY)
        '''
        #print join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))), 'asset', 'fifth.png')
        try:
            plt.savefig(join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))), 'asset', 'fifth.png'),facecolor = '#edeeff', dpi=96)
        except:
            plt.savefig('asset/fifth.png', facecolor='#edeeff', dpi=96)
        plt.close()
        matplotlib.rc('ytick', labelsize=12)
        plt.axis([0, 100, 0, 1600])
        plt.plot(comp_year_list, comp_fourth_biomass_list)
        plt.xlabel("time (yrs)")
        plt.ylabel("Standing Biomass (g/m2)")
        #plt.rc('axes', labelsize=20)
        ax = plt.gca()

        ax.set_facecolor('#edeeff')
        #plt.xticks([1, 20, 40, 60, 80, 100], ['0', '20', '40', '60', '80', '100'])
        try:
            plt.savefig(join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))), 'asset', 'fourth.png'),facecolor = '#edeeff')
        except:
            plt.savefig('asset/fourth.png',facecolor = '#edeeff')
        plt.close()
        matplotlib.rc('ytick', labelsize=12)
        plt.axis([0, 300, 0, 1600])

        plt.plot(comp_elev_list, comp_biomass_list)
        plt.xlabel("Elevation (cm) Rel to MSL")
        plt.ylabel("Standing Biomass (g/m2)")

        ax = plt.gca()
        ax.set_facecolor('#edeeff')
        #plt.xticks([1, 20, 40, 60, 80, 100], ['0', '20', '40', '60', '80', '100'])
        try:
            plt.savefig(join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))), 'asset', 'first.png'),facecolor = '#edeeff', dpi=96)
        except:
            plt.savefig('asset/first.png',facecolor = '#edeeff', dpi=96)
        #plt.savefig('asset/first.png',facecolor = '#edeeff')
        plt.close()
        matplotlib.rc('xtick',labelsize=15)
        matplotlib.rc('ytick', labelsize=15)
        plt.axis([1, 100, 0, 1])
        plt.plot(comp_year_list, comp_ind_time)
        plt.xlabel("time (yrs)")
        plt.ylabel("Inundation Time (0-1)")
        ax = plt.gca()
        ax.set_facecolor('#edeeff')
        # set the locations and labels of the xticks
        plt.xticks([1,20,40,60,80,100], ['0', '20', '40', '60', '80', '100'])
        try:
            plt.savefig(join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))), 'asset', 'second.png'),facecolor = '#edeeff')
        except:
            plt.savefig('asset/second.png',facecolor = '#edeeff')
        #plt.savefig('asset/second.png',facecolor = '#edeeff')
        plt.close()

        plt.axis([0, 100, 0, 70])
        plt.plot(comp_year_list, comp_cquest_list)
        plt.xlabel("time (yrs)")
        plt.ylabel("CSequestraion (g Cm2 y2)")
        ax = plt.gca()
        ax.set_facecolor('#edeeff')
        #plt.xticks([1, 20, 40, 60, 80, 100], ['0', '20', '40', '60', '80', '100'])
        try:
            plt.savefig(join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))), 'asset', 'third.png'),facecolor = '#edeeff')
        except:
            plt.savefig('asset/third.png',facecolor = '#edeeff')
        #plt.savefig('asset/third.png',facecolor = '#edeeff')
        plt.close()
        #self.afterFill()
        #parent = self.GetParent()
        #print self
        #print "-----"
        #print parent
        #self.Update()
        #self.Refresh()
        self.drawImages()
    def onRadioButton(self, e):
        #print "here"
        if e is None:
            lab = 'North Inlet, SC'
        else:
            cb_r = e.GetEventObject()
            lab = cb_r.GetLabel()
        #self.rupPan.SetBackgroundColour("red")
        #print cb_r.GetLabel()
        #cb_r = e.GetEventObject()
        #self.SetTitle(cb_r.GetLabel())
        # self.rupPan.ruplbl.SetLabel(cb_r.GetLabel())
        ##print TabOne.filePath
        ##print cb_r.GetLabel()
        title_lbl = wx.StaticText(self.rupPan, -1, pos=(370, 10),size=(500,10))
        title_font = wx.Font(14, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD)
        title_lbl.SetFont(title_font)
        title_lbl.SetLabel(lab)
        if lab == 'North Inlet, SC':
            #print 'North'
            #object1 = TabTwo("abcd")
            #sum = object1.rows
            ##print texts[81]
            cb1.SetValue(False)
            cb2.SetValue(False)
            cb3.SetValue(False)
            cb4.SetValue(True)

            #k_4=k_2= k_3= k_5= k_6= k_7=k_8=k_9=[]
            ##print data_texts
            w, h = 8, 14
            data_list = []
            #del data_list[:]
            myGrid.ClearGrid()
            ##print data_list
            data_list = [["" for x in range(w)] for y in range(h)]
            ##print data_list
            ind=ind_1=ind_2=ind_3=0
            ##print data_texts[62]
            #for i in range(62, 75):
            for i in range(62, 76):
                #MSL - let 1996 be t0
                #data_list

                data_list[ind][0] = data_texts[i][5]
                data_list[ind][1] = data_texts[i][6]
                ind=ind+1
                #k_2.append(float(data_texts[i][5]))
                #k_3.append(float(data_texts[i][6]))
            for i in range(80, 89):
                #MSL - let 1996 be t0
                ##print data_texts[i]
                data_list[ind_1][2] = str(float(data_texts[i][0]) - 1996)
                data_list[ind_1][3] = str(float(data_texts[i][1]) * 100)
                ind_1=ind_1+1
                #k_4.append(float(data_texts[i][0]) - 1996)
                #k_5.append(float(data_texts[i][1]) * 100)
                ##print k_5 # convert to cm
            for i in range(17,31):
                data_list[ind_2][4] = data_texts[i][9]
                data_list[ind_2][5] = data_texts[i][10]
                ind_2 = ind_2 + 1
                #k_6.append(float(data_texts[i][9]))
                #k_7.append(float(data_texts[i][10]))
            for i in range(18,32):
                data_list[ind_3][6] = str(float(data_texts[i][14]) - 1996)
                data_list[ind_3][7] = str(float(data_texts[i][15]))
                ind_3=ind_3+1
                #k_8.append(float(data_texts[i][14])-1996)
                #k_9.append(float(data_texts[i][15]))
            ##print data_list
            myGrid.SetCellValue(0, 0, "Hello")
            for i in range(0,len(data_list)):
                for j in range(0,len(data_list[i])):
                    myGrid.SetCellValue(i, j, data_list[i][j])
            phy_sea_level_forecast.SetLabel("30")
            phy_sea_level_start.SetLabel("-1")
            phy_20th.SetLabel("0.2")
            phy_MTA.SetLabel("70")
            phy_Marsh_ele.SetLabel("43")
            phy_sus_minSed.SetLabel("20")
            phy_sus_org.SetLabel("0")
            phy_lt.SetLabel("")
            bio_max_growth.SetLabel("110")
            bio_min_growth.SetLabel("-25")
            bio_opt_growth.SetLabel("35")
            bio_max_peak.SetLabel("1200")
            bio_OM_below_root.SetLabel("5")
            bio_OM_decay.SetLabel("-0.3")
            bio_BGBio.SetLabel("2")
            bio_BG_turnover.SetLabel("1")
            bio_max_root_depth.SetLabel("25")
            bio_reserved.SetLabel("")
            model_max_capture.SetLabel("1")
            model_refrac.SetLabel("0.1")
        if lab == 'Grand Bay, MS':
            #print 'Grand'
            cb1.SetValue(False)
            cb2.SetValue(False)
            cb3.SetValue(False)
            cb4.SetValue(False)
            #object1 = TabTwo("abcd")
            #sum = object1.rows
            ##print texts[81]

            w, h = 8, 59
            data_list = []
            #del data_list[:]
            myGrid.ClearGrid()

            data_list = [["" for x in range(w)] for y in range(h)]
            ##print data_list
            ind = ind_1 = ind_2 = ind_3 = 0
            ##print data_texts[61]
            ##print data_texts[61][5]
            ##print data_texts[61][6]
            # for i in range(62, 75):
            for i in range(2, 61):
                # MSL - let 1996 be t0
                # data_list

                data_list[ind][0] = str(data_texts[i][5])
                data_list[ind][1] = str(data_texts[i][6])
                ind = ind + 1
                # k_2.append(float(data_texts[i][5]))
                # k_3.append(float(data_texts[i][6]))
            for i in range(41, 79):
                # MSL - let 1996 be t0
                ##print data_texts[i]
                data_list[ind_1][2] = str(float(data_texts[i][0]) - 2013)
                data_list[ind_1][3] = str(float(data_texts[i][1]))
                ind_1 = ind_1 + 1
                # k_4.append(float(data_texts[i][0]) - 1996)
                # k_5.append(float(data_texts[i][1]) * 100)
                # #print k_5 # convert to cm
            for i in range(2, 14):
                data_list[ind_2][4] = data_texts[i][9]
                data_list[ind_2][5] = data_texts[i][10]
                ind_2 = ind_2 + 1
                # k_6.append(float(data_texts[i][9]))
                # k_7.append(float(data_texts[i][10]))
            for i in range(2, 15):
                if data_texts[i][15] > -30 :
                    data_list[ind_3][6] = data_texts[i][14]
                #data_list[ind_3][6] = str(float(data_texts[i][14]) - 1996)
                data_list[ind_3][7] = data_texts[i][15]
                ind_3 = ind_3 + 1
                # k_8.append(float(data_texts[i][14])-1996)
                # k_9.append(float(data_texts[i][15]))'''
            ##print data_list
            for i in range(0, len(data_list)):
                for j in range(0, len(data_list[i])):
                    myGrid.SetCellValue(i, j, data_list[i][j])
            phy_sea_level_forecast.SetLabel("100")
            phy_sea_level_start.SetLabel("9")
            phy_20th.SetLabel("0.25")
            phy_MTA.SetLabel("30")
            phy_Marsh_ele.SetLabel("14")
            phy_sus_minSed.SetLabel("15")
            phy_sus_org.SetLabel("0")
            phy_lt.SetLabel("")
            bio_max_growth.SetLabel("50")
            bio_min_growth.SetLabel("-30")
            bio_opt_growth.SetLabel("25")
            bio_max_peak.SetLabel("2400")
            bio_OM_below_root.SetLabel("8")
            bio_OM_decay.SetLabel("-0.4")
            bio_BGBio.SetLabel("2")
            bio_BG_turnover.SetLabel("0.8")
            bio_max_root_depth.SetLabel("30")
            bio_reserved.SetLabel("")
            model_max_capture.SetLabel("1")
            model_refrac.SetLabel("0.05")
            '''items=[['a','b','c']]
            file_name = "C:\Users\VKOTHA\Downloads\Temp.xls"
            #filename = MainFrame.filePath
            book = xlrd.open_workbook(file_name, formatting_info=1)
            sheetname = "Numerical_Output"
            # sheetname = "Data"
            sheet = book.sheet_by_name(sheetname)
            rows, cols = sheet.nrows, sheet.ncols
            #print rows
            #print cols
            comments, texts = XG.ReadExcelCOM(file_name, sheetname, rows, cols)
            xlsGrid = XG.XLSGrid(self)
            #print book
            #print sheet
            #print texts
            #print comments
            xlsGrid.PopulateGrid(book, sheet, items, comments)
            ##print k_2'''

        if lab == 'Plum Island, MA':
            #print 'Plum'
            #object1 = TabTwo("abcd")
            #sum = object1.rows
            ##print texts[81]
            cb1.SetValue(False)
            cb2.SetValue(False)
            cb3.SetValue(False)
            cb4.SetValue(False)

            w, h = 8, 51
            data_list = []
            #del data_list[:]
            myGrid.ClearGrid()
            data_list = [["" for x in range(w)] for y in range(h)]
            ##print data_list
            ind = ind_1 = ind_2 = ind_3 = 0
            ##print data_texts[61]
            ##print data_texts[61][5]
            ##print data_texts[61][6]
            # for i in range(62, 75):
            for i in range(78, 99):
                # MSL - let 1996 be t0
                # data_list

                data_list[ind][0] = str(data_texts[i][5])
                data_list[ind][1] = str(data_texts[i][6])
                ind = ind + 1
                # k_2.append(float(data_texts[i][5]))
                # k_3.append(float(data_texts[i][6]))
            for i in range(90, 141):
                # MSL - let 1996 be t0
                ##print data_texts[i]
                data_list[ind_1][2] = str(float(data_texts[i][0]) - 2013)
                data_list[ind_1][3] = str(float(data_texts[i][1]))
                ind_1 = ind_1 + 1
                # k_4.append(float(data_texts[i][0]) - 1996)
                # k_5.append(float(data_texts[i][1]) * 100)
                # #print k_5 # convert to cm
            '''for i in range(2, 14):
                data_list[ind_2][4] = data_texts[i][9]
                data_list[ind_2][5] = data_texts[i][10]
                ind_2 = ind_2 + 1'''
                # k_6.append(float(data_texts[i][9]))
                # k_7.append(float(data_texts[i][10]))
            for i in range(33, 45):
                data_list[ind_2][6] = data_texts[i][14]
                data_list[ind_2][7] = data_texts[i][15]
                ind_2 = ind_2 + 1
            '''for i in range(2, 15):
                if data_texts[i][15] > -30 :
                    data_list[ind_3][6] = data_texts[i][14]
                #data_list[ind_3][6] = str(float(data_texts[i][14]) - 1996)
                data_list[ind_3][7] = data_texts[i][15]
                ind_3 = ind_3 + 1'''
                # k_8.append(float(data_texts[i][14])-1996)
                # k_9.append(float(data_texts[i][15]))'''
            #print data_list
            for i in range(0, len(data_list)):
                for j in range(0, len(data_list[i])):
                    myGrid.SetCellValue(i, j, data_list[i][j])
            phy_sea_level_forecast.SetLabel("40")
            phy_sea_level_start.SetLabel("1.8")
            phy_20th.SetLabel("0.2")
            phy_MTA.SetLabel("160")
            phy_Marsh_ele.SetLabel("142.7")
            phy_sus_minSed.SetLabel("15")
            phy_sus_org.SetLabel("1")
            phy_lt.SetLabel("")
            bio_max_growth.SetLabel("195")
            bio_min_growth.SetLabel("0")
            bio_opt_growth.SetLabel("100")
            bio_max_peak.SetLabel("1400")
            bio_OM_below_root.SetLabel("18")
            bio_OM_decay.SetLabel("-0.2")
            bio_BGBio.SetLabel("2")
            bio_BG_turnover.SetLabel("1")
            bio_max_root_depth.SetLabel("25")
            bio_reserved.SetLabel("")
            model_max_capture.SetLabel("1")
            model_refrac.SetLabel("0.05")
            cb1.SetValue(False)
            cb2.SetValue(False)
            cb3.SetValue(False)
            cb4.SetValue(False)
        if lab == 'Apalachicola, FL':
            #print 'Apcola'
            # object1 = TabTwo("abcd")
            # sum = object1.rows
            # #print texts[81]
            cb1.SetValue(False)
            cb2.SetValue(False)
            cb3.SetValue(False)
            cb4.SetValue(False)
            w, h = 8, 75
            data_list = []
            #del data_list[:]
            myGrid.ClearGrid()
            data_list = [["" for x in range(w)] for y in range(h)]
            ##print data_list
            ind = ind_1 = ind_2 = ind_3 = 0
            ##print data_texts[61]
            ##print data_texts[61][5]
            ##print data_texts[61][6]
            # for i in range(62, 75):
            for i in range(2, 77):
                # MSL - let 1996 be t0
                # data_list
                ##print data_texts[i][7]
                ##print data_texts[i][8]
                data_list[ind][0] = str(data_texts[i][7])
                data_list[ind][1] = str(data_texts[i][8])
                ind = ind + 1
                # k_2.append(float(data_texts[i][5]))
                # k_3.append(float(data_texts[i][6]))
            for i in range(2, 40):
                # MSL - let 1996 be t0
                ##print data_texts[i]
                data_list[ind_1][2] = str(float(data_texts[i][0]) - 2013)
                data_list[ind_1][3] = str(float(data_texts[i][1]))
                ind_1 = ind_1 + 1
                # k_4.append(float(data_texts[i][0]) - 1996)
                # k_5.append(float(data_texts[i][1]) * 100)
                # #print k_5 # convert to cm
            for i in range(2, 14):
                data_list[ind_2][4] = data_texts[i][11]
                data_list[ind_2][5] = data_texts[i][12]
                ind_2 = ind_2 + 1
            #for i in range(2, 15):
            if data_texts[2][16] > -30 :
                data_list[1][6] = data_texts[2][14]
                #data_list[ind_3][6] = str(float(data_texts[i][14]) - 1996)
            data_list[1][7] = data_texts[2][16]
            #print data_list
            for i in range(0, len(data_list)):
                for j in range(0, len(data_list[i])):
                    myGrid.SetCellValue(i, j, data_list[i][j])
            phy_sea_level_forecast.SetLabel("100")
            phy_sea_level_start.SetLabel("11")
            phy_20th.SetLabel("0.2")
            phy_MTA.SetLabel("22")
            phy_Marsh_ele.SetLabel("24.2")
            phy_sus_minSed.SetLabel("20")
            phy_sus_org.SetLabel("0")
            phy_lt.SetLabel("")
            bio_max_growth.SetLabel("70")
            bio_min_growth.SetLabel("-10")
            bio_opt_growth.SetLabel("25")
            bio_max_peak.SetLabel("2400")
            bio_OM_below_root.SetLabel("25")
            bio_OM_decay.SetLabel("-0.3")
            bio_BGBio.SetLabel("2")
            bio_BG_turnover.SetLabel("0.8")
            bio_max_root_depth.SetLabel("30")
            bio_reserved.SetLabel("")
            #model_max_capture.SetLabel("1")
            model_refrac.SetLabel("0.05")
        if lab == 'Coon Isl, SFB':
            #print 'Coon'
            # object1 = TabTwo("abcd")
            # sum = object1.rows
            # #print texts[81]
            cb1.SetValue(False)
            cb2.SetValue(False)
            cb3.SetValue(False)
            cb4.SetValue(False)
            w, h = 8, 100
            data_list = []
            # del data_list[:]
            myGrid.ClearGrid()
            data_list = [["" for x in range(w)] for y in range(h)]
            # #print data_list
            ind = ind_1 = ind_2 = ind_3 = 0
            # #print data_texts[61]
            # #print data_texts[61][5]
            # #print data_texts[61][6]
            # for i in range(62, 75):
            for i in range(78, 178):
                # MSL - let 1996 be t0
                # data_list
                # #print data_texts[i][7]
                # #print data_texts[i][8]
                data_list[ind][0] = str(data_texts[i][7])
                data_list[ind][1] = str(data_texts[i][8])
                ind = ind + 1
            #print data_list
            for i in range(0, len(data_list)):
                for j in range(0, len(data_list[i])):
                    myGrid.SetCellValue(i, j, data_list[i][j])
            phy_sea_level_forecast.SetLabel("100")
            phy_sea_level_start.SetLabel("106")
            phy_20th.SetLabel("0.24")
            phy_MTA.SetLabel("85")
            phy_Marsh_ele.SetLabel("179")
            phy_sus_minSed.SetLabel("100")
            phy_sus_org.SetLabel("0")
            phy_lt.SetLabel("")
            bio_max_growth.SetLabel("89")
            bio_min_growth.SetLabel("-36")
            bio_opt_growth.SetLabel("64")
            bio_max_peak.SetLabel("1200")
            bio_OM_below_root.SetLabel("10")
            bio_OM_decay.SetLabel("-0.3")
            bio_BGBio.SetLabel("4")
            bio_BG_turnover.SetLabel("0.5")
            bio_max_root_depth.SetLabel("20")
            bio_reserved.SetLabel("")
            #model_max_capture.SetLabel("1")
            model_refrac.SetLabel("0.1")
        if lab == 'Other Estuary':
            #print 'Other'
            # object1 = TabTwo("abcd")
            # sum = object1.rows
            # #print texts[81]
            myGrid.ClearGrid()
            phy_sea_level_forecast.SetLabel("100")
            phy_sea_level_start.SetLabel("0")
            phy_20th.SetLabel("0.2")
            phy_MTA.SetLabel("70")
            phy_Marsh_ele.SetLabel("45")
            phy_sus_minSed.SetLabel("20")
            phy_sus_org.SetLabel("0")
            phy_lt.SetLabel("")
            if cb1.GetValue() == True:
                bio_max_growth.SetLabel("120")
                bio_min_growth.SetLabel("-30")
                bio_opt_growth.SetLabel("35")
            bio_max_peak.SetLabel("1200")
            bio_OM_below_root.SetLabel("5")
            bio_OM_decay.SetLabel("-0.4")
            bio_BGBio.SetLabel("4")
            bio_BG_turnover.SetLabel("0.5")
            bio_max_root_depth.SetLabel("30")
            bio_reserved.SetLabel("0.2")
            #model_max_capture.SetLabel("1")
            model_refrac.SetLabel("0.1")
            cb1.SetValue(True)
            cb2.SetValue(False)
            cb3.SetValue(False)
            cb4.SetValue(False)

        self.onCalculate(e)


    '''def onButton(self, event, qqq):
        """
        This method is fired when its corresponding button is pressed
        """
        #what = self.textBox.GetValue()
        #print qqq
        #print "Button pressed!"'''

class TabThree(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)
        #panel = wx.Panel(self)
        #t = wx.StaticText(self, -1, "This is the third tab", (20, 20))
        '''try:
            imageFile = "C:\Users\VKOTHA\Downloads\Instructions.jpg"
            #imageFile = join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))), 'asset', 'Instructions.jpg')


        except:'''
            #print "exceptinstr"
        imageFile = join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))), 'asset','Instructions.jpg')
            #bitmap = wx.Bitmap("asset/logo.jpg")
        #print imageFile
        #imageFile = "C:\Users\VKOTHA\Downloads\Instructions.jpg"
        #leftPan = wx.lib.scrolledpanel.ScrolledPanel(self)
        #leftPan.SetupScrolling()
        png = wx.Image(imageFile, wx.BITMAP_TYPE_ANY).ConvertToBitmap()
        wx.StaticBitmap(self, -1, png, (10, 5), (png.GetWidth(), png.GetHeight()))



class TabTwo(wx.Panel):
    '''def __init__(self, parent):
        wx.Panel.__init__(self, parent)
        #t = wx.StaticText(self, -1, "This is the third tab", (20, 20))
        filename = "C:\Users\VKOTHA\Downloads\Temp.xls"
        #filename = MainFrame.filePath
        book = xlrd.open_workbook(filename, formatting_info=1)
        sheetname = "Numerical_Output"
        #sheetname = "Data"
        sheet = book.sheet_by_name(sheetname)
        rows, cols = sheet.nrows, sheet.ncols
        #global comments, texts
        comments, texts= XG.ReadExcelCOM(filename, sheetname, rows, cols)
        #global xlsGrid
        xlsGrid = XG.XLSGrid(self)
        xlsGrid.PopulateGrid(book, sheet, texts, comments)
        ##print texts
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(xlsGrid, 1, wx.EXPAND, 5)
        self.SetSizer(sizer)'''
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)
        global numGrid
        numGrid = gridlib.Grid(self)
        numGrid.CreateGrid(1000, 1000)
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(numGrid, 1, wx.EXPAND)
        self.SetSizer(sizer)


class TabFour(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)
        #t = wx.StaticText(self, -1, "This is the third tab", (20, 20))
        #filename = "C:\Users\VKOTHA\Downloads\Temp.xls"
        try:
            fname =  join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))),'asset','Temp.xls')
            book = xlrd.open_workbook(fname, formatting_info=1)
        except IOError:
            fname = join(os.path.dirname(os.path.abspath(sys.argv[0])), 'asset', 'Temp.xls')
            book = xlrd.open_workbook(fname, formatting_info=1)
        #print os.path.abspath(sys.argv[0])

        #print os.path.abspath("__file__")
        #print fname
        #print abspath(__file__)
        #filename = MainFrame.filePath
        #book = xlrd.open_workbook(fname, formatting_info=1)
        sheetname = "Data"
        sheet = book.sheet_by_name(sheetname)
        rows, cols = sheet.nrows, sheet.ncols
        global data_texts
        comments, data_texts = XG.ReadExcelCOM(fname, sheetname, rows, cols)
        xlsGrid = XG.XLSGrid(self)
        xlsGrid.PopulateGrid(book, sheet, data_texts, comments)
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(xlsGrid, 1, wx.EXPAND, 5)
        self.SetSizer(sizer)

class TabFive(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)
        # t = wx.StaticText(self, -1, "This is the third tab", (20, 20))
        '''filename = "C:\Users\VKOTHA\Downloads\Temp.xls"
        #filename = MainFrame.filePath
        book = xlrd.open_workbook(filename, formatting_info=1)
        sheetname = "IO_data"
        # sheetname = "Data"
        sheet = book.sheet_by_name(sheetname)
        rows, cols = sheet.nrows, sheet.ncols
        ##print rows
        ##print cols
        comments, texts = XG.ReadExcelCOM(filename, sheetname, rows, cols)


        aaGrid = XG.XLSGrid(self)
        ##print book
        ##print sheet
        ##print texts
        ##print comments
        aaGrid.PopulateGrid(book, sheet, texts, comments)
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(aaGrid, 1, wx.EXPAND, 5)
        self.SetSizer(sizer)'''
        global myGrid
        myGrid = MyGrid(self, size=(100, 100))
        myGrid.CreateGrid(100, 8)
        myGrid.SetColLabelRenderer(0, TextLabelRenderer('SOM', 2))
        myGrid.SetColLabelRenderer(1, TextLabelRenderer('', 0))
        myGrid.SetColLabelRenderer(2, TextLabelRenderer('MSL', 2))
        myGrid.SetColLabelRenderer(3, TextLabelRenderer('', 0))
        myGrid.SetColLabelRenderer(4, TextLabelRenderer('Growth', 2))
        myGrid.SetColLabelRenderer(5, TextLabelRenderer('', 0))
        myGrid.SetColLabelRenderer(6, TextLabelRenderer('Marsh Elevation', 2))
        myGrid.SetColLabelRenderer(7, TextLabelRenderer('', 0))
        '''global myGrid
        myGrid = gridlib.Grid(self)
        myGrid.CreateGrid(100, 10)
        #myGrid.SetColLabelSize(2)
        myGrid.SetColLabelSize(0)
        myGrid.SetCellSize(0, 0, 1, 3)
        myGrid.SetCellValue(0, 0, "Yesterday")
        myGrid.SetColLabelValue(0, "abcd")'''
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(myGrid, 1, wx.EXPAND)

        self.SetSizer(sizer)


class TabSix(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)
        global rootdistGrid
        rootdistGrid = gridlib.Grid(self)
        rootdistGrid.CreateGrid(1000, 1000)
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(rootdistGrid, 1, wx.EXPAND)
        self.SetSizer(sizer)
class TabSeven(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)
        global compGrid
        compGrid = gridlib.Grid(self)
        compGrid.CreateGrid(1000, 1000)
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(compGrid, 1, wx.EXPAND)
        self.SetSizer(sizer)
class TabEight(wx.Panel):
    '''def __init__(self, parent):
        wx.Panel.__init__(self, parent)
        t = wx.StaticText(self, -1, "This is the third tab", (20, 20))'''
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)
        global inunGrid
        inunGrid = gridlib.Grid(self)
        inunGrid.CreateGrid(1000, 1000)
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(inunGrid, 1, wx.EXPAND)
        self.SetSizer(sizer)
class TabNine(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)
        global sheet12Grid
        sheet12Grid = gridlib.Grid(self)
        sheet12Grid.CreateGrid(1000, 1000)
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(sheet12Grid, 1, wx.EXPAND)
        self.SetSizer(sizer)
class TabTen(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)
        global sheet10Grid
        sheet10Grid = gridlib.Grid(self)
        sheet10Grid.CreateGrid(1000, 1000)
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(sheet10Grid, 1, wx.EXPAND)
        self.SetSizer(sizer)
class MySplashScreen(wx.SplashScreen):
    """
Create a splash screen widget.
    """
    def __init__(self, parent=None):
        # This is a recipe to a the screen.
        # Modify the following variables as necessary.
        try:
            #print "trybitmap"
            #print join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))), 'asset', 'logo.jpg')
            #plt.savefig(join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))), 'asset', 'fifth.png'),facecolor = '#edeeff', dpi=96)
            aBitmap = wx.Image(name=join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))), 'asset', 'logo.jpg')).ConvertToBitmap()

        except:
            #print "exceptbitmap"
            aBitmap = wx.Image(name="asset/logo.jpg").ConvertToBitmap()



        #aBitmap = wx.Image(name="asset/logo.jpg").ConvertToBitmap()
        splashStyle = wx.SPLASH_CENTRE_ON_SCREEN | wx.SPLASH_TIMEOUT
        splashDuration = 1000 # milliseconds
        # Call the constructor with the above arguments in exactly the
        # following order.
        wx.SplashScreen.__init__(self, aBitmap, splashStyle,
                                 splashDuration, parent)
        self.Bind(wx.EVT_CLOSE, self.OnExit)

        wx.Yield()
#----------------------------------------------------------------------#

    def OnExit(self, evt):
        self.Hide()
        # MyFrame is the main frame.
        #MyFrame = MyGUI(None, -1, "Hello from wxPython")
        #app.SetTopWindow(MyFrame)
        #MyFrame.Show(True)
        # The program will freeze without this line.
        evt.Skip()  # Make sure the default handler runs too...
class GifPanel(wx.Panel):
    """ class MyPanel creates a panel, inherits wx.Panel """
    def __init__(self, parent, id):
        # default pos and size creates a panel that fills the frame
        wx.Panel.__init__(self, parent, id)
        self.SetBackgroundColour("white")
        # pick the filename of an animated GIF file you have ...
        # give it the full path and file name!
        ag_fname = "asset/loading.gif"
        ag = wx.animate.GIFAnimationCtrl(self, id, ag_fname, pos=(10, 10))
        # clears the background
        ag.GetPlayer().UseBackgroundColour(True)
        # continuously loop through the frames of the gif file (default)
        ag.Play()

class MainFrame(wx.Frame):
    #filePath = "MEM_file.xls"
    #filePath = "test.xls"
    #filePath = "C://Users/VKOTHA/Downloads/Temp.xls"
    def __init__(self):
        wx.Frame.__init__(self, None, size = (1200,650), title="MEM v6.0")
        # Create a panel and notebook (tabs holder)
        #self.Centre()
        p = wx.Panel(self)
        #self.SetDimensions(0, 0, 1200, 480)
        #self.Show()
        #print wx.GetDisplaySize()
        #print wx.DefaultSize
        #self.ShowFullScreen(True)
        nb = wx.Notebook(p)

        # Create the tab windows
        tab5 = TabFive(nb)
        tab4 = TabFour(nb)

        #tab2 = TabTwo(nb)
        tab3 = TabThree(nb)


        tab6 = TabSix(nb)
        tab7 = TabSeven(nb)
        tab8 = TabEight(nb)
        tab9 = TabNine(nb)
        tab10 = TabTen(nb)
        tab1 = TabOne(nb)
        # Add the windows to tabs and name them.
        nb.AddPage(tab1, "IO Page")
        #nb.AddPage(tab2, "Numerical Output")
        nb.AddPage(tab3, "Instructions")
        nb.AddPage(tab7, "Computations")
        nb.AddPage(tab6, "rootdist")
        #nb.AddPage(tab8, "Inundation Time")
        nb.AddPage(tab4, "Data")
        nb.AddPage(tab9, "Sheet12")
        nb.AddPage(tab10, "Sheet15")
        nb.AddPage(tab5, "IO_data")
        #nb.AddPage(tab6, "rootdist")
        #nb.AddPage(tab7, "Computations")
        #nb.AddPage(tab8, "Sheet12")


        # Set noteboook in a sizer to create the layout
        sizer = wx.BoxSizer()
        sizer.Add(nb, 1, wx.EXPAND)
        p.SetSizer(sizer)


def show_splash():
    # create, show and return the splash screen
    try:
        #print "tryshowsplash"
        #print join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))), 'asset', 'logo.jpg')
        # plt.savefig(join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))), 'asset', 'fifth.png'),facecolor = '#edeeff', dpi=96)
        #aBitmap = wx.Image(name=join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))), 'asset','logo.jpg')).ConvertToBitmap()
        bitmap = wx.Bitmap(join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))), 'asset','logo.jpg'))


    except:
        #print "exceptshowsplash"
        bitmap = wx.Bitmap("asset/logo.jpg")

    #bitmap = wx.Bitmap("asset/logo.jpg")
    splash = wx.SplashScreen(bitmap, wx.SPLASH_CENTRE_ON_SCREEN|wx.SPLASH_NO_TIMEOUT, 0, None, -1)
    splash.Show()
    return splash
if __name__ == "__main__":
    app = wx.App()
    splash = show_splash()
    MainFrame().Show()
    splash.Destroy()
    app.MainLoop()
