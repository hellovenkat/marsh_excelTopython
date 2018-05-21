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
import matplotlib.patches as mpatches
import matplotlib.pyplot as plt
from matplotlib.backends.backend_wxagg import FigureCanvasWxAgg as FigureCanvas
from matplotlib.backends.backend_wx import NavigationToolbar2Wx
from matplotlib.figure import Figure
wildcard = "Excel Workbook (*.xls)|*.xls|" \
            "All files (*.*)|*.*"
GIFNames = ['C:/Users/VKOTHA/Desktop/marsh_excelToPython/asset/loader.gif']
global data_1,data_2,data_3,data_4,data_5,data_6
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

    def InitUI(self):

        self.SetBackgroundColour('#edeeff')
        vbox = wx.BoxSizer(wx.VERTICAL)
        hbox = wx.BoxSizer(wx.HORIZONTAL)
        #leftPan = wx.Panel(self)
        leftPan = wx.lib.scrolledpanel.ScrolledPanel(self)
        leftPan.SetupScrolling()

        vbox_leftPan = wx.BoxSizer(wx.VERTICAL)
        global cb1, cb2, cb3, cb4
        cb1 = wx.CheckBox(leftPan, label='Use a generic biom profile')
        cb2 = wx.CheckBox(leftPan, label='Add thin layer')
        cb3 = wx.CheckBox(leftPan, label='Calibrate to accretion rate')
        cb4 = wx.CheckBox(leftPan, label='for future development')
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
        global data1_label,data2_label,data3_label,data4_label,data5_label,data6_label,user_label,phy_sea_level_forecast, phy_sea_level_start, phy_20th, phy_MTA, phy_Marsh_ele, phy_sus_minSed, phy_sus_org, phy_lt
        phy_sea_level_forecast = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="",size=(40, -1))
        phy_sea_level_start = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="",size=(40, -1))
        phy_20th = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="",size=(40, -1))
        phy_MTA = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="",size=(40, -1))
        phy_Marsh_ele = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="",size=(40, -1))
        phy_sus_minSed = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="",size=(40, -1))
        phy_sus_org = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="",size=(40, -1))
        phy_lt = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="",size=(40, -1))
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
        fgs_bio = wx.FlexGridSizer(10, 3, 0, 0)
        global bio_max_growth, bio_min_growth, bio_opt_growth, bio_max_peak, bio_OM_below_root, bio_OM_decay, bio_BGBio, bio_BG_turnover, bio_max_root_depth, bio_reserved
        bio_max_growth = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="",size=(40, -1))
        bio_min_growth = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="",size=(40, -1))
        bio_opt_growth = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="",size=(40, -1))
        bio_max_peak = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="",size=(40, -1))
        bio_OM_below_root = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="",size=(40, -1))
        bio_OM_decay = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="",size=(40, -1))
        bio_BGBio = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="",size=(40, -1))
        bio_BG_turnover = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="",size=(40, -1))
        bio_max_root_depth = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="",size=(40, -1))
        bio_reserved = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="",size=(40, -1))
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
        model_max_capture = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="",size=(40, -1))
        model_refrac = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="",size=(40, -1))
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
        epi_years = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="",size=(40, -1))
        epi_repeat = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="",size=(40, -1))
        epi_recoveryTime = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="",size=(40, -1))
        epi_addElevation = wx.TextCtrl(leftPan, style=wx.TE_RIGHT, value="",size=(40, -1))
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
        ###############################
        hbox.Add(leftPan, 1, wx.EXPAND | wx.ALL, 5)

        self.rupPan = wx.Panel(self)
        self.rupPan.SetBackgroundColour('#edeeff')

        wx.StaticText(self.rupPan, -1, "Copyright University of South Carolina 2010. All Rights Reserved, JT Morris 6-9-10", (325, 390))

        vbox.Add(self.rupPan, 2, wx.EXPAND | wx.ALL, 0)


        hbox_rdownPan = wx.BoxSizer(wx.HORIZONTAL)

        rbut_rdownPan = wx.Panel(self)
        rtext_rdownPan = wx.Panel(self)
        rtext_vbox = wx.BoxSizer(wx.VERTICAL)

        hbox_rdownPan.Add(rbut_rdownPan, 0.90, wx.EXPAND | wx.ALL, 0)
        hbox_rdownPan.Add(rtext_rdownPan, 2, wx.EXPAND | wx.ALL, 0)
        rtextfont_underline = wx.Font(11, wx.ROMAN, wx.FONTSTYLE_NORMAL, wx.NORMAL, underline=True)
        rtextfont = wx.Font(11, wx.ROMAN, wx.FONTSTYLE_NORMAL, wx.NORMAL, underline=False)
        rdown_text1 = '''Metrics computed over the final 50 years of simulation'''
        lbl_1 = wx.StaticText(rtext_rdownPan, -1, style=wx.ALIGN_CENTER)
        lbl_1.SetFont(rtextfont_underline)
        lbl_1.SetLabel(rdown_text1)
        rtext_vbox.Add((-1, 10))
        rtext_vbox.Add((-1, 10))
        rtext_vbox.Add(lbl_1, 0, wx.ALIGN_CENTER)
        rtext_vbox.Add((-1, 5))
        data1_label = wx.StaticText(rtext_rdownPan, label="null", style=wx.ALIGN_CENTER)
        data2_label = wx.StaticText(rtext_rdownPan, label="null", style=wx.ALIGN_CENTER)
        data3_label = wx.StaticText(rtext_rdownPan, label="null", style=wx.ALIGN_CENTER)
        data4_label = wx.StaticText(rtext_rdownPan, label="null", style=wx.ALIGN_CENTER)
        data5_label = wx.StaticText(rtext_rdownPan, label="null", style=wx.ALIGN_CENTER)
        data6_label = wx.StaticText(rtext_rdownPan, label="null", style=wx.ALIGN_CENTER)
        data1_label.SetFont(rtextfont)
        data2_label.SetFont(rtextfont)
        data3_label.SetFont(rtextfont)
        data4_label.SetFont(rtextfont)
        data5_label.SetFont(rtextfont)
        data6_label.SetFont(rtextfont)

        text1_label = wx.StaticText(rtext_rdownPan, label=" avg vert accretion (cm/yr)  last 50 yr of the simulation (yrs 51-100 average)", style=wx.ALIGN_CENTER)
        text2_label = wx.StaticText(rtext_rdownPan, label=" refractory c seq (g C/m2/yr) at the end of the simulation from top 50 cohorts", style=wx.ALIGN_CENTER)
        text3_label = wx.StaticText(rtext_rdownPan, label=" total C/m2 in belowground biomass in top 50 cohorts (50 years) (g C/m2/yr)", style=wx.ALIGN_CENTER)
        text4_label = wx.StaticText(rtext_rdownPan, label=" avg vert accretion (cm/yr)  first 50 yr of the simulation (yrs 1-50 average)", style=wx.ALIGN_CENTER)
        text5_label = wx.StaticText(rtext_rdownPan, label=" refractory c seq (g C/m2/yr) at the mid point of the simulation from top 50 cohorts", style=wx.ALIGN_CENTER)
        text6_label = wx.StaticText(rtext_rdownPan, label=" total C/m2 in belowground biomass from top 50 cohorts (50 years) (g C/m2/yr)", style=wx.ALIGN_CENTER)
        text1_label.SetFont(rtextfont)
        text2_label.SetFont(rtextfont)
        text3_label.SetFont(rtextfont)
        text4_label.SetFont(rtextfont)
        text5_label.SetFont(rtextfont)
        text6_label.SetFont(rtextfont)

        txt_Pan1 = wx.BoxSizer(wx.HORIZONTAL)
        txt_Pan1.Add(data1_label)
        txt_Pan1.Add(text1_label)
        rtext_vbox.Add(txt_Pan1, 0,wx.ALIGN_CENTER, wx.ALL, 5)
        txt_Pan2 = wx.BoxSizer(wx.HORIZONTAL)
        txt_Pan2.Add(data2_label)
        txt_Pan2.Add(text2_label)
        rtext_vbox.Add(txt_Pan2, 0,wx.ALIGN_CENTER, wx.ALL, 5)

        txt_Pan3 = wx.BoxSizer(wx.HORIZONTAL)
        txt_Pan3.Add(data3_label)
        txt_Pan3.Add(text3_label)
        rtext_vbox.Add(txt_Pan3, 0,wx.ALIGN_CENTER, wx.ALL, 5)


        rdown_text2='''Metrics computed over the first 50 years of simulation'''
        lbl_2 = wx.StaticText(rtext_rdownPan, -1, style=wx.ALIGN_CENTER)
        lbl_2.SetFont(rtextfont_underline)
        lbl_2.SetLabel(rdown_text2)
        rtext_vbox.Add((-1, 20))
        rtext_vbox.Add(lbl_2, 0,wx.ALIGN_CENTER, wx.ALIGN_CENTER)
        rtext_vbox.Add((-1, 5))
        txt_Pan4 = wx.BoxSizer(wx.HORIZONTAL)
        txt_Pan4.Add(data4_label)
        txt_Pan4.Add(text4_label)
        rtext_vbox.Add(txt_Pan4, 0,wx.ALIGN_CENTER, wx.ALL, 5)

        txt_Pan5 = wx.BoxSizer(wx.HORIZONTAL)
        txt_Pan5.Add(data5_label)
        txt_Pan5.Add(text5_label)
        rtext_vbox.Add(txt_Pan5, 0,wx.ALIGN_CENTER, wx.ALL, 5)

        txt_Pan6 = wx.BoxSizer(wx.HORIZONTAL)
        txt_Pan6.Add(data6_label)
        txt_Pan6.Add(text6_label)
        rtext_vbox.Add(txt_Pan6, 0,wx.ALIGN_CENTER, wx.ALL, 5)
        rtext_rdownPan.SetSizer(rtext_vbox)
        vbox.Add(hbox_rdownPan, 1, wx.EXPAND | wx.ALL, 0)
        rbut_vbox = wx.BoxSizer(wx.VERTICAL)
        r1 = wx.RadioButton(rbut_rdownPan, label='Plum Island, MA')
        r2 = wx.RadioButton(rbut_rdownPan, label='North Inlet, SC')
        r3 = wx.RadioButton(rbut_rdownPan, label='Apalachicola, FL')
        r4 = wx.RadioButton(rbut_rdownPan, label='Grand Bay, MS')
        r5 = wx.RadioButton(rbut_rdownPan, label='Coon Isl, SFB')
        r6 = wx.RadioButton(rbut_rdownPan, label='Other Estuary')

        r1.Bind(wx.EVT_RADIOBUTTON, self.onRadioButton)

        r2.Bind(wx.EVT_RADIOBUTTON, self.onRadioButton)
        r3.Bind(wx.EVT_RADIOBUTTON, self.onRadioButton)
        r4.Bind(wx.EVT_RADIOBUTTON, self.onRadioButton)
        r5.Bind(wx.EVT_RADIOBUTTON, self.onRadioButton)
        r6.Bind(wx.EVT_RADIOBUTTON, self.onRadioButton)
        r2.SetValue(True)
        self.onRadioButton(None)


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




        hbox.Add(vbox, 3, wx.EXPAND | wx.ALL, 5)
        self.SetSizer(hbox)

        #panel.EnableScrolling(True,True)
    def drawImages(self):

        folder = 'asset'

        image1 = join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))), folder, 'first.png')

        image = wx.Image(image1, wx.BITMAP_TYPE_ANY)
        image = image.Scale(320, 170, wx.IMAGE_QUALITY_HIGH)
        imageBitmap = wx.StaticBitmap(self.rupPan, wx.ID_ANY, wx.BitmapFromImage(image))
        imageBitmap.SetPosition((30, 30))
        image2 = join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))), folder, 'second.png')
        image = wx.Image(image2, wx.BITMAP_TYPE_ANY)
        image = image.Scale(320, 170, wx.IMAGE_QUALITY_HIGH)
        imageBitmap = wx.StaticBitmap(self.rupPan, wx.ID_ANY, wx.BitmapFromImage(image))
        imageBitmap.SetPosition((360, 30))
        image3 = join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))), folder, 'third.png')

        image = wx.Image(image3, wx.BITMAP_TYPE_ANY)
        image = image.Scale(320, 170, wx.IMAGE_QUALITY_HIGH)
        imageBitmap = wx.StaticBitmap(self.rupPan, wx.ID_ANY, wx.BitmapFromImage(image))
        imageBitmap.SetPosition((690, 30))
        image4 = join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))), folder, 'fourth.png')

        image = wx.Image(image4, wx.BITMAP_TYPE_ANY)
        image = image.Scale(320, 170, wx.IMAGE_QUALITY_HIGH)
        imageBitmap = wx.StaticBitmap(self.rupPan, wx.ID_ANY, wx.BitmapFromImage(image))
        imageBitmap.SetPosition((30, 210))
        image5 = join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))), folder, 'fifth.png')
        image = wx.Image(image5, wx.BITMAP_TYPE_ANY)
        image = image.Scale(320, 170, wx.IMAGE_QUALITY_HIGH)
        imageBitmap = wx.StaticBitmap(self.rupPan, wx.ID_ANY, wx.BitmapFromImage(image))
        imageBitmap.SetPosition((360, 210))
        image6 = join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))), folder, 'sixth.png')
        image = wx.Image(image6, wx.BITMAP_TYPE_ANY)
        image = image.Scale(320, 170, wx.IMAGE_QUALITY_HIGH)
        imageBitmap = wx.StaticBitmap(self.rupPan, wx.ID_ANY, wx.BitmapFromImage(image))
        imageBitmap.SetPosition((690, 210))
    def onOpenFile(self, event):

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
    def changeVal(self,data_1, data_2 , data_3, data_4, data_5, data_6):
        data1_label.SetLabel(str(data_1))
        data2_label.SetLabel(str(data_2))
        data3_label.SetLabel(str(data_3))
        data4_label.SetLabel(str(data_4))
        data5_label.SetLabel(str(data_5))
        data6_label.SetLabel(str(data_6))
    def onCalculate(self,e,lab):
        MySplash = MySplashScreen()
        MySplash.Show()


        BGB = [0] * 1801
        SedD = [0] * 1801
        lbgb = [0] * 1801
        dzdt = [0] * 1801
        MSL = [0] * 1801
        inorg = [0] * 601

        MHW = [0] * 1801
        bio = [0] * 1801
        OMmat = [0] * 1801
        MHWs = [0] * 1801
        marshelev = [0] * 1801
        D = [0] * 1801
        T = [0] * 1801

        decay = [0] * 1801
        SOM = [0] * 1801
        cquest = [0] * 601
        cmat = [0] * 601
        soildepth = [0] * 601

        dzdd = [0] * 1801
        dref = [0] * 1801
        dlab = [0] * 1801
        ddcay = [0] * 1801
        bulkd = [0] * 1801
        sedi = [0] * 1801
        droot = [0] * 1801

        totdepth = [0] * 1801
        totBGB = [0] * 1801
        IT = [0] * 1801
        ybio = [0] * 101
        dreftot = [0] * 601
        sorg = [0] * 601

        bins = [["" for x in range(401)] for y in range(41)]
        bincounts = [0] * 41
        cohortbins = [0] * 401

        corelbg = [0] * 141
        coretbg = [0] * 141
        coretin = [0] * 141
        coresom = [0] * 141

        a= [0] * 11
        b= [0] * 11
        c= [0] * 11
        aleft= [0] * 11
        bleft= [0] * 11
        cleft= [0] * 11
        aright= [0] * 11
        bright= [0] * 11
        cright= [0] * 11


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
        num_sed_depth_1 = []
        num_sed_org_1 = []
        num_sed_depth_2 = []
        num_sed_org_2 = []
        IO_data_1 = []
        IO_data_2 = []



        k1 = 0.085
        k2 = 1.99
        troubleshoot = 1
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
            #print "cb1"
            res = Tamp + 30
            bio_max_growth.SetLabel(str(res))
            bio_min_growth.SetLabel("-10")
        j2 = 1

        MHW[1] = MHW0
        D[1] = MHW[1] - marshelev[1]

        if cb1.GetValue() is False:
            #print "cb1 false"
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
            #print "cb1 else"
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
            a[i] = aleft[i]
            b[i] = bleft[i]
            c[i] = cleft[i]
        if D[1] > Dopt and cb1.GetValue() is False:
            a[i] = aright[i]
            b[i] = bright[i]
            c[i] = cright[i]
        Drmax = float(bio_max_root_depth.GetValue())
        lna = 1.1
        p = -0.512
        bscale = 0.0001

        bio[1] = (a[1] * D[1] + b[1] * (D[1]*D[1]) + c[1]) * bscale

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
            if x<=Dopt and cb1.GetValue() is False:
                a[1] = aleft[1]
                b[1] = bleft[1]
                c[1] = cleft[1]
            if x>Dopt and cb1.GetValue() is False:

                a[1] = aright[1]
                b[1] = bright[1]
                c[1] = cright[1]

            ybio[j+1] = (a[1] * x + b[1] * (x*x) + c[1]) * bscale
            if ybio[j+1] < 0:
                ybio[j+1]=0
            comp_list[j+2][17] = elev
            comp_list[j+2][18] = ybio[j+1] * 10000

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
        BRZOM = 90 * (sorg[1] / k1 + krBGTR * RT / k1) / (inorg[1] / k2 + sorg[1] / k1 + krBGTR * RT / k1)

        if cb3.GetValue() is True:

                maxvertacc = float(phy_lt.GetValue())
                krBGTR = (k1 / RT) * (maxvertacc - inorg[1] / k2 - sorg[1] / k1)
                kr = krBGTR / BGTR
                BGTR = min(3, BGTR)
                model_refrac.SetLabel(str(kr))
                BRZOM = 90 * (sorg[1] / k1 + krBGTR * RT / k1) /   (inorg[1] / k2 + sorg[1] / k1 + krBGTR * RT / k1)
                cohort_size = inorg[1] / k2 + sorg[1] / k1 + krBGTR * RT / k1
                phy_lt.SetLabel(str(round(cohort_size, 2)))
                maxvertacc = cohort_size
                bio_BG_turnover.SetLabel(str(BGTR))

        # changed here
        Bden = (sorg[1] + krBGTR * RT + inorg[1]) / ((sorg[1] + krBGTR * RT) / k1 + inorg[1] / k2)

        rootdist_list[0][0] = RT
        rootdist_list[1][0] = kd
        rootdist_list[2][0] = Rmax
        rootdist_list[3][0] = Drmax
        rootdist_list[4][0] = kr
        rootdist_list[5][0] = omdr
        rootdist_list[6][0] = BGTR

        #'************************************************
        if minwt == 0 and D[1] < 0:
            dzm = (sorg[1] + krBGTR * RT) / k1
        if minwt == 0 and cb3.GetValue() is True:
            dzm = maxvertacc

        nocohort = 500

        for ico in range(1,nocohort+1):
            dzdd[ico] = dzm
            sedi[ico] = minwt
            dref[ico] = sorg[1] + krBGTR * RT
            bulkd[ico] = Bden
            BD = dzdd[ico] / dref[ico]
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
            for ico in range(1, nocohort+1):
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
            if troubleshoot == 1:
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
            #print Bot
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
            for j in range(jlast, 401):
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
            comp_list[k][16] = " "
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
            bio[jtime] = (a[irecov] * D[jtime] + b[irecov] * (D[jtime]*D[jtime]) + c[irecov]) * bscale
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
                thintime = thintime + int(epi_repeat)# next thin layer appl
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
            comp_list[k][16] = round(totdepth[jtime],2)
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
        # wrong code----should be done
        data_1 = round((marshelev[100] - marshelev[50]) / 50, 2)
        data_2 = round(refracC100 / 50, 1)
        data_3 = round(totC100, 1)
        data_4 = round((marshelev[51] - marshelev[1]) / 50, 2)
        data_5 = round(refracC50 / 50, 1)
        data_6 = round(totC50, 1)
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

        for i in range(0, len(num_output_list)):
            for j in range(0, 30):
                numGrid.SetCellValue(i, j, str(num_output_list[i][j]))
        for i in range(0, len(sheet10_list)):
            for j in range(0, len(sheet10_list[i])):
                sheet10Grid.SetCellValue(i, j, str(sheet10_list[i][j]))
        for i in range(0, len(sheet12_list)):
            for j in range(0, len(sheet12_list[i])):
                sheet12Grid.SetCellValue(i, j, str(sheet12_list[i][j]))
        for i in range(0, len(rootdist_list)):
            for j in range(0, len(rootdist_list[i])):
                rootdistGrid.SetCellValue(i, j, str(rootdist_list[i][j]))
        for i in range(2, 83):
            comp_elev_list.append(comp_list[i][17])
        for i in range(2, 83):
            comp_biomass_list.append(comp_list[i][18])
        for i in range(2, 102):

                comp_year_list.append(comp_list[i][1])
                comp_ind_time.append(comp_list[i][5])
                comp_cquest_list.append(comp_list[i][8])
                comp_fourth_biomass_list.append(comp_list[i][4])
                comp_msl_list.append(comp_list[i][12])
                comp_marshele_list.append(comp_list[i][15])

        for i in range(2, 200):
            if data_list[i][0] != '':
               IO_data_1.append(data_list[i][0])
            if data_list[i][1] != '':
               IO_data_2.append(data_list[i][1])

        for i in range(3, 43):
            if num_output_list[i][20] != '':
               num_sed_depth_1.append(num_output_list[i][16])
               num_sed_org_1.append(num_output_list[i][20])
        for i in range(45, 85):
            if num_output_list[i][20] != '':
               num_sed_depth_2.append(num_output_list[i][16])
               num_sed_org_2.append(num_output_list[i][20])

        matplotlib.rc('xtick',labelsize=15)
        matplotlib.rc('ytick', labelsize=15)
        plt.rc('axes', labelsize=15)
        plt.axis([0, 100, 0, 200])


        plot_msl, = plt.plot(comp_year_list, comp_msl_list, 'black')
        plot_marshele, = plt.plot(comp_year_list, comp_marshele_list, 'green')

        l1 = plt.legend([plot_marshele], ["Marsh Elevation"], loc=1, fontsize="xx-large", framealpha=0)
        l2 = plt.legend([plot_msl], ["MSL"], loc=4, fontsize="xx-large", framealpha=0)  # this removes l1 from the axes.
        plt.gca().add_artist(l1)
        '''fifth_plot_lines = []
        fifth_plot_lines.append([plot_msl,plot_marshele])
        fifth_legend = plt.legend(fifth_plot_lines[0], ["MSL","Marsh Elevation" ], loc=0, fontsize="xx-large")
        plt.gca().add_artist(fifth_legend)'''

        plt.xlabel("time (yrs)")
        plt.ylabel("cm NAVD")


        IO_data_marsh_elev_year=[]
        IO_data_marsh_elev_cm=[]
        for row in range(1,myGrid.GetNumberRows()):
            temp_marsh_elev_year = myGrid.GetCellValue(row, 6)
            temp_marsh_elev_cm = myGrid.GetCellValue(row, 7)
            if temp_marsh_elev_year!='' and temp_marsh_elev_cm!='':
                IO_data_marsh_elev_year.append(float(temp_marsh_elev_year))
            #if temp_marsh_elev_cm!='':
                IO_data_marsh_elev_cm.append(float(temp_marsh_elev_cm))
        plt.scatter(IO_data_marsh_elev_year, IO_data_marsh_elev_cm, color='green')


        ax = plt.gca()
        ax.set_facecolor('#edeeff')
        ax.autoscale(enable=True)
        ax.yaxis.grid(True)



        plt.savefig(join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))), 'asset', 'fifth.png'),facecolor = '#edeeff', dpi=96)

        plt.close()
        matplotlib.rc('ytick', labelsize=12)
        plt.axis([0, 100, 0, 1600])

        plt.plot(comp_year_list, comp_fourth_biomass_list)

        plt.xlabel("time (yrs)")
        plt.ylabel("Standing Biomass (g/m2)")

        ax = plt.gca()
        ax.set_facecolor('#edeeff')
        ax.autoscale(enable=True)
        ax.yaxis.grid(True)

        plt.savefig(join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))), 'asset', 'fourth.png'),facecolor = '#edeeff')

        plt.close()
        matplotlib.rc('ytick', labelsize=12)
        plt.axis([0, 300, 0, 1600])

        plt.plot(comp_elev_list, comp_biomass_list)
        plt.xlabel("Elevation (cm) Rel to MSL")
        plt.ylabel("Standing Biomass (g/m2)")
        IO_data_elev=[]
        IO_data_biom=[]
        for row in range(1,myGrid.GetNumberRows()):
            temp_elev = myGrid.GetCellValue(row, 4)
            temp_biom = myGrid.GetCellValue(row, 5)
            if temp_elev!='':
                IO_data_elev.append(float(temp_elev))
            if temp_biom!='':
                IO_data_biom.append(float(temp_biom))
        plt.axvline(x=0)
        plt.scatter(IO_data_elev, IO_data_biom, color='red')
        ax = plt.gca()
        ax.set_facecolor('#edeeff')
        ax.autoscale(enable=True)
        ax.yaxis.grid(True)

        plt.savefig(join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))), 'asset', 'first.png'),facecolor = '#edeeff')

        plt.close()
        matplotlib.rc('xtick',labelsize=15)
        matplotlib.rc('ytick', labelsize=15)
        plt.axis([1, 100, 0, 1])
        plt.plot(comp_year_list, comp_ind_time)
        plt.xlabel("time (yrs)")
        plt.ylabel("Inundation Time (0-1)")
        ax = plt.gca()
        ax.set_facecolor('#edeeff')
        ax.autoscale(enable=True)
        ax.yaxis.grid(True)

        plt.savefig(join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))), 'asset', 'second.png'),facecolor = '#edeeff')

        plt.close()

        plt.axis([0, 100, 0, 70])
        plt.plot(comp_year_list, comp_cquest_list)
        plt.xlabel("time (yrs)")
        plt.ylabel("CSequestraion (g Cm2 y2)")
        ax = plt.gca()
        ax.set_facecolor('#edeeff')
        ax.autoscale(enable=True)
        ax.yaxis.grid(True)

        plt.savefig(join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))), 'asset', 'third.png'),facecolor = '#edeeff')

        plt.close()
        matplotlib.rc('ytick', labelsize=12)
        plt.axis([0, 100, 0, 50])


        pres, = plt.plot(num_sed_depth_1, num_sed_org_1, 'b', label="Present")
        fut, = plt.plot(num_sed_depth_2, num_sed_org_2, 'r', label = "Future")

        plt.xlabel("Sediment Depth (cm)")
        plt.ylabel("Sediment Org. Matter (%)")
        sixth_plot_lines = []
        sixth_plot_lines.append([pres,fut])
        sixth_legend = plt.legend(sixth_plot_lines[0], ["Present", "Future"], loc=1, fontsize="xx-large", framealpha=0)
        plt.gca().add_artist(sixth_legend)
        IO_data_Dsom = []
        IO_data_LOI = []
        for row in range(1, myGrid.GetNumberRows()):
            temp_dsom = myGrid.GetCellValue(row, 0)
            temp_Loi = myGrid.GetCellValue(row, 1)
            if temp_dsom != '':
                IO_data_Dsom.append(float(temp_dsom))
            if temp_Loi != '':
                IO_data_LOI.append(float(temp_Loi))
        plt.scatter(IO_data_Dsom, IO_data_LOI, color='blue')
        ax = plt.gca()
        ax.set_facecolor('#edeeff')
        ax.autoscale(enable=True)
        ax.yaxis.grid(True)

        plt.savefig(join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))), 'asset', 'sixth.png'),facecolor = '#edeeff')

        plt.close()
        title_lbl = wx.StaticText(self.rupPan, -1, pos=(370, 10),size=(500,10))
        title_font = wx.Font(14, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD)
        title_lbl.SetFont(title_font)
        title_lbl.SetLabel(lab)
        self.changeVal(data_1, data_2 , data_3, data_4, data_5, data_6)
        self.drawImages()
    def onRadioButton(self, e):

        if e is None:
            lab = 'North Inlet, SC'
        else:
            cb_r = e.GetEventObject()
            lab = cb_r.GetLabel()


        if lab == 'North Inlet, SC':

            cb1.SetValue(False)
            cb2.SetValue(False)
            cb3.SetValue(False)
            cb4.SetValue(True)


            ##print data_texts
            w, h = 8, 14
            data_list = []

            myGrid.ClearGrid()
            abcd(self)
            data_list = [["" for x in range(w)] for y in range(h)]

            ind=ind_1=ind_2=ind_3=0

            for i in range(62, 76):
                #MSL - let 1996 be t0
                #data_list

                data_list[ind][0] = data_texts[i][5]
                data_list[ind][1] = data_texts[i][6]
                ind=ind+1



            for i in range(80, 89):

                data_list[ind_1][2] = str(float(data_texts[i][0]) - 1996)
                data_list[ind_1][3] = str(float(data_texts[i][1]) * 100)
                ind_1=ind_1+1

            for i in range(17,31):
                data_list[ind_2][4] = data_texts[i][9]
                data_list[ind_2][5] = data_texts[i][10]
                ind_2 = ind_2 + 1

            for i in range(18,32):
                data_list[ind_3][6] = str(float(data_texts[i][14]) - 1996)
                data_list[ind_3][7] = str(float(data_texts[i][15]))
                ind_3=ind_3+1


            for i in range(0,len(data_list)):
                for j in range(0,len(data_list[i])):
                    myGrid.SetCellValue(i+1, j, data_list[i][j])
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
            epi_years.SetLabel("20")
            epi_repeat.SetLabel("20")
            epi_recoveryTime.SetLabel("10")
            epi_addElevation.SetLabel("10")
        if lab == 'Grand Bay, MS':

            cb1.SetValue(False)
            cb2.SetValue(False)
            cb3.SetValue(False)
            cb4.SetValue(False)


            w, h = 8, 59
            data_list = []

            myGrid.ClearGrid()
            abcd(self)
            data_list = [["" for x in range(w)] for y in range(h)]

            ind = 0
            ind_1 = 0
            ind_2 = 0
            ind_3 = 0

            for i in range(2, 61):
                # MSL - let 1996 be t0
                # data_list

                data_list[ind][0] = str(data_texts[i][5])
                data_list[ind][1] = str(data_texts[i][6])
                ind = ind + 1

            for i in range(41, 79):
                # MSL - let 1996 be t0

                data_list[ind_1][2] = str(float(data_texts[i][0]) - 2013)
                data_list[ind_1][3] = str(float(data_texts[i][1]))
                ind_1 = ind_1 + 1

            for i in range(2, 14):
                data_list[ind_2][4] = data_texts[i][9]
                data_list[ind_2][5] = data_texts[i][10]
                ind_2 = ind_2 + 1

            for i in range(2, 15):
                if data_texts[i][15] > -30 :
                    data_list[ind_3][6] = data_texts[i][14]

                data_list[ind_3][7] = data_texts[i][15]
                ind_3 = ind_3 + 1


            for i in range(0, len(data_list)):
                for j in range(0, len(data_list[i])):
                    myGrid.SetCellValue(i+1, j, data_list[i][j])
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
            epi_years.SetLabel("20")
            epi_repeat.SetLabel("20")
            epi_recoveryTime.SetLabel("10")
            epi_addElevation.SetLabel("10")


        if lab == 'Plum Island, MA':

            cb1.SetValue(False)
            cb2.SetValue(False)
            cb3.SetValue(False)
            cb4.SetValue(False)

            w, h = 8, 51

            myGrid.ClearGrid()
            abcd(self)
            data_list = [["" for x in range(w)] for y in range(h)]

            ind = ind_1 = ind_2 = ind_3 = 0

            for i in range(78, 99):
                # MSL - let 1996 be t0
                # data_list

                data_list[ind][0] = str(data_texts[i][5])
                data_list[ind][1] = str(data_texts[i][6])
                ind = ind + 1

            for i in range(90, 141):
                # MSL - let 1996 be t0

                data_list[ind_1][2] = str(float(data_texts[i][0]) - 2013)
                data_list[ind_1][3] = str(float(data_texts[i][1]))
                ind_1 = ind_1 + 1


            for i in range(33, 45):
                data_list[ind_2][6] = data_texts[i][14]
                data_list[ind_2][7] = data_texts[i][15]
                ind_2 = ind_2 + 1

            for i in range(0, len(data_list)):
                for j in range(0, len(data_list[i])):
                    myGrid.SetCellValue(i+1, j, data_list[i][j])
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
            epi_years.SetLabel("20")
            epi_repeat.SetLabel("20")
            epi_recoveryTime.SetLabel("10")
            epi_addElevation.SetLabel("10")
            cb1.SetValue(False)
            cb2.SetValue(False)
            cb3.SetValue(False)
            cb4.SetValue(False)
        if lab == 'Apalachicola, FL':

            cb1.SetValue(False)
            cb2.SetValue(False)
            cb3.SetValue(False)
            cb4.SetValue(False)
            w, h = 8, 75
            data_list = []

            myGrid.ClearGrid()
            abcd(self)
            data_list = [["" for x in range(w)] for y in range(h)]

            ind = ind_1 = ind_2 = ind_3 = 0

            for i in range(2, 77):

                data_list[ind][0] = str(data_texts[i][7])
                data_list[ind][1] = str(data_texts[i][8])
                ind = ind + 1

            for i in range(2, 40):
                # MSL - let 1996 be t0

                data_list[ind_1][2] = str(float(data_texts[i][0]) - 2013)
                data_list[ind_1][3] = str(float(data_texts[i][1]))
                ind_1 = ind_1 + 1

            for i in range(2, 14):
                data_list[ind_2][4] = data_texts[i][11]
                data_list[ind_2][5] = data_texts[i][12]
                ind_2 = ind_2 + 1

            if data_texts[2][16] > -30 :
                data_list[1][6] = data_texts[2][14]

            data_list[1][7] = data_texts[2][16]

            for i in range(0, len(data_list)):
                for j in range(0, len(data_list[i])):
                    myGrid.SetCellValue(i+1, j, data_list[i][j])
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
            model_max_capture.SetLabel("1")
            model_refrac.SetLabel("0.05")
            epi_years.SetLabel("20")
            epi_repeat.SetLabel("20")
            epi_recoveryTime.SetLabel("10")
            epi_addElevation.SetLabel("10")
        if lab == 'Coon Isl, SFB':

            cb1.SetValue(False)
            cb2.SetValue(False)
            cb3.SetValue(False)
            cb4.SetValue(False)
            w, h = 8, 100
            data_list = []

            myGrid.ClearGrid()
            abcd(self)
            data_list = [["" for x in range(w)] for y in range(h)]

            ind = ind_1 = ind_2 = ind_3 = 0

            for i in range(78, 178):
                # MSL - let 1996 be t0

                data_list[ind][0] = str(data_texts[i][7])
                data_list[ind][1] = str(data_texts[i][8])
                ind = ind + 1

            for i in range(0, len(data_list)):
                for j in range(0, len(data_list[i])):
                    myGrid.SetCellValue(i+1, j, data_list[i][j])
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
            model_max_capture.SetLabel("1")
            model_refrac.SetLabel("0.1")
            epi_years.SetLabel("20")
            epi_repeat.SetLabel("20")
            epi_recoveryTime.SetLabel("10")
            epi_addElevation.SetLabel("10")
        if lab == 'Other Estuary':

            myGrid.ClearGrid()
            abcd(self)
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
            model_max_capture.SetLabel("1")
            model_refrac.SetLabel("0.1")
            cb1.SetValue(True)
            cb2.SetValue(False)
            cb3.SetValue(False)
            cb4.SetValue(False)
            epi_years.SetLabel("20")
            epi_repeat.SetLabel("20")
            epi_recoveryTime.SetLabel("10")
            epi_addElevation.SetLabel("10")

        self.onCalculate(e,lab)

class TabThree(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)
        #Pan = wx.lib.scrolledpanel.ScrolledPanel(self)
        #Pan.SetupScrolling()
        #siz = wx.BoxSizer(wx.VERTICAL)

        imageFile = join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))), 'asset','Instructions.jpg')

        png = wx.Image(imageFile, wx.BITMAP_TYPE_ANY).ConvertToBitmap()
        wx.StaticBitmap(self, -1, png, (5, 5), (png.GetWidth(), png.GetHeight()))
        #siz.Add(k, 1, wx.EXPAND)

        #self.SetSizer(siz)


class TabTwo(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)
        global numGrid
        numGrid = gridlib.Grid(self)
        numGrid.CreateGrid(2000, 30)

        numGrid.SetColLabelValue(0, "")

        numGrid.SetColLabelValue(1, "year")
        numGrid.SetColLabelValue(2, "MSL (cm NAVD)")
        numGrid.SetColLabelValue(3, "Marsh Elevtion (cm NAVD)")
        numGrid.SetColLabelValue(4, "Standing Biomass (g/m2)")
        numGrid.SetColLabelValue(5, "dzdd")
        numGrid.SetColLabelValue(6, "Sediment Depth (cm)")
        numGrid.SetColLabelValue(7, "Live BG Biomass (g/m2)")
        numGrid.SetColLabelValue(8, "Labile OM (g/m2)")
        numGrid.SetColLabelValue(9, "Refractory OM (g/m2)")
        numGrid.SetColLabelValue(10, "Tot BG Biomass (g/m2)")
        numGrid.SetColLabelValue(11, "SOM (%)")
        numGrid.SetColLabelValue(12, "Inorganic Sed. (g/cm2)")
        numGrid.SetColLabelValue(13, "Bulk Density (g/cm3)")
        numGrid.SetColLabelValue(14, "Live Root (mg/cm3)")
        numGrid.SetColLabelValue(15, "Decay Rate (g dry wt/m2/yr)")
        numGrid.SetColLabelValue(16, "Sediment Depth (cm)")
        numGrid.SetColLabelValue(17, "Live BG Biomass (g/m2)")
        numGrid.SetColLabelValue(18, "Tot BG Biomass (g/m2)")
        numGrid.SetColLabelValue(19, "Inorganic Sed. (g/m2)")
        numGrid.SetColLabelValue(20, "SOM (%)")
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(numGrid, 1, wx.EXPAND)
        self.SetSizer(sizer)


class TabFour(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)

        fname =  join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))),'asset','Temp.xls')
        book = xlrd.open_workbook(fname, formatting_info=1)


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

        global myGrid
        myGrid = MyGrid(self, size=(100, 100))
        myGrid.CreateGrid(105, 8)
        myGrid.SetColLabelRenderer(0, TextLabelRenderer('SOM', 2))
        myGrid.SetColLabelRenderer(1, TextLabelRenderer('', 0))
        myGrid.SetColLabelRenderer(2, TextLabelRenderer('MSL', 2))
        myGrid.SetColLabelRenderer(3, TextLabelRenderer('', 0))
        myGrid.SetColLabelRenderer(4, TextLabelRenderer('Growth', 2))
        myGrid.SetColLabelRenderer(5, TextLabelRenderer('', 0))
        myGrid.SetColLabelRenderer(6, TextLabelRenderer('Marsh Elevation', 2))
        myGrid.SetColLabelRenderer(7, TextLabelRenderer('', 0))
        attr = wx.grid.GridCellAttr()
        attr.SetTextColour("navyblue")
        #attr.SetBackgroundColour("pink")
        attr.SetFont(wx.Font(10, wx.SWISS, wx.NORMAL, wx.BOLD))
        myGrid.SetRowAttr(0, attr)
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(myGrid, 1, wx.EXPAND)

        self.SetSizer(sizer)


class TabSix(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)
        global rootdistGrid
        rootdistGrid = gridlib.Grid(self)
        rootdistGrid.CreateGrid(1000, 1000)
        rootdistGrid.SetRowLabelValue(0, "RT")
        rootdistGrid.SetRowLabelValue(1, "kd")
        rootdistGrid.SetRowLabelValue(2, "Rmax")
        rootdistGrid.SetRowLabelValue(3, "Drmax")
        rootdistGrid.SetRowLabelValue(4, "kr")
        rootdistGrid.SetRowLabelValue(5, "OM decay")
        rootdistGrid.SetRowLabelValue(6, "BGTurn")
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(rootdistGrid, 1, wx.EXPAND)
        self.SetSizer(sizer)
class TabSeven(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)

        global compGrid
        compGrid = gridlib.Grid(self)
        compGrid.CreateGrid(1000, 30)
        compGrid.SetColLabelValue(0, "")
        compGrid.SetColLabelValue(1, "Year")
        compGrid.SetColLabelValue(2, "MHW")
        compGrid.SetColLabelValue(3, "dzdt")
        compGrid.SetColLabelValue(4, "biom")
        compGrid.SetColLabelValue(5, "Ind Time")
        compGrid.SetColLabelValue(6, "Sorg")
        compGrid.SetColLabelValue(7, "")
        compGrid.SetColLabelValue(8, "Cquest")
        compGrid.SetColLabelValue(9, "sorg")
        compGrid.SetColLabelValue(10, "totBGB")
        compGrid.SetColLabelValue(11, "Omat")
        compGrid.SetColLabelValue(12, "MSL")
        compGrid.SetColLabelValue(13, "Inorg")
        compGrid.SetColLabelValue(14, "slr")
        compGrid.SetColLabelValue(15, "marshE")
        compGrid.SetColLabelValue(16, "pedonD")
        compGrid.SetColLabelValue(17, "Elev")
        compGrid.SetColLabelValue(18, "Biomass")
        compGrid.SetColLabelValue(19, "LBGB")
        compGrid.SetColLabelValue(20, "BD")
        compGrid.SetColLabelValue(21, "droot")
        compGrid.SetColLabelValue(22, "dzdd")
        compGrid.SetColLabelValue(23, "doabymo")

        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(compGrid, 1, wx.EXPAND)
        self.SetSizer(sizer)
class TabEight(wx.Panel):

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


        aBitmap = wx.Image(name=join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))), 'asset','logo.jpg')).ConvertToBitmap()

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

    def __init__(self):
        wx.Frame.__init__(self, None, size = (1200,650), title="MEM v6.0")
        # Create a panel and notebook (tabs holder)

        p = wx.Panel(self)

        nb = wx.Notebook(p)

        # Create the tab windows
        tab5 = TabFive(nb)
        tab4 = TabFour(nb)
        tab2 = TabTwo(nb)
        tab3 = TabThree(nb)
        tab6 = TabSix(nb)
        tab7 = TabSeven(nb)
        tab8 = TabEight(nb)
        tab9 = TabNine(nb)
        tab10 = TabTen(nb)
        tab1 = TabOne(nb)
        # Add the windows to tabs and name them.
        nb.AddPage(tab1, "IO Page")
        nb.AddPage(tab2, "Numerical Output")
        nb.AddPage(tab3, "Instructions")
        nb.AddPage(tab7, "Computations")
        nb.AddPage(tab6, "rootdist")
        nb.AddPage(tab8, "Inundation Time")
        nb.AddPage(tab4, "Data")
        nb.AddPage(tab9, "Sheet12")
        nb.AddPage(tab10, "Sheet15")
        nb.AddPage(tab5, "IO_data")



        # Set noteboook in a sizer to create the layout
        sizer = wx.BoxSizer()
        sizer.Add(nb, 1, wx.EXPAND)

#        wx.Panel.__init__(self, parent, -1)

        #sizer = wx.FlexGridSizer(2,3,5,5)

        '''for name in GIFNames:
            ani = wx.animate.Animation(name)
            ctrl = wx.animate.AnimationCtrl(self, -1, ani)
            ctrl.SetUseWindowBackgroundColour()
            ctrl.Play()

            sizer.AddF(ctrl, wx.SizerFlags().Border(wx.ALL, 10))

        border = wx.BoxSizer()
        border.AddF(sizer, wx.SizerFlags(1).Expand().Border(wx.ALL, 20))
        self.SetSizer(border)'''

        p.SetSizer(sizer)
        _icon = wx.Icon(join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))), 'asset', 'icon.ico'), wx.BITMAP_TYPE_ICO)
        self.SetIcon(_icon)

def show_splash():
    # create, show and return the splash screen

    bitmap = wx.Bitmap(join(os.path.dirname(os.path.dirname(os.path.abspath(sys.argv[0]))), 'asset', 'logo.jpg'))

    splash = wx.SplashScreen(bitmap, wx.SPLASH_CENTRE_ON_SCREEN|wx.SPLASH_NO_TIMEOUT, 0, None, -1)
    splash.Show()
    return splash
def abcd(self):
        myGrid.SetCellValue(0, 0, "D (cm)")
        myGrid.SetCellValue(0, 1, "LOI (%)")
        myGrid.SetCellValue(0, 2, "year")
        myGrid.SetCellValue(0, 3, "cm")
        myGrid.SetCellValue(0, 4, "elev")
        myGrid.SetCellValue(0, 5, "biom")
        myGrid.SetCellValue(0, 6, "year")
        myGrid.SetCellValue(0, 7, "cm")
if __name__ == "__main__":
    app = wx.App()
    splash = show_splash()
    MainFrame().Show()
    splash.Destroy()
    app.MainLoop()
