VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{5936A75C-3F42-11D6-AF6B-AA0004005F12}#1.3#0"; "MeansCtrl.ocx"
Begin VB.Form frmCCIIdxDtlExcGrid 
   Caption         =   "CCI Index Detail Exception Report Grid"
   ClientHeight    =   6750
   ClientLeft      =   2265
   ClientTop       =   2835
   ClientWidth     =   12000
   Icon            =   "frmCCIIdxDtlExcRptGrid.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6750
   ScaleWidth      =   12000
   Begin VB.Frame fraSelectDates 
      Caption         =   "Select Dates"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4320
      TabIndex        =   28
      Top             =   2160
      Width           =   1905
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   1695
         TabIndex        =   52
         Top             =   180
         Width           =   1695
         Begin VB.OptionButton optQuarter 
            Caption         =   "Quarter"
            Height          =   225
            Left            =   0
            TabIndex        =   29
            Top             =   0
            Value           =   -1  'True
            Width           =   870
         End
         Begin VB.OptionButton optPeriod 
            Caption         =   "Period"
            Height          =   225
            Left            =   880
            TabIndex        =   30
            Top             =   0
            Width           =   765
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Variance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7320
      TabIndex        =   4
      Top             =   360
      Width           =   4575
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   260
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   4380
         TabIndex        =   49
         Top             =   180
         Width           =   4380
         Begin VB.OptionButton optVar0Pct 
            Caption         =   "None"
            Height          =   225
            Left            =   0
            TabIndex        =   5
            Top             =   30
            Value           =   -1  'True
            Width           =   780
         End
         Begin VB.TextBox txtVariancePct 
            Height          =   285
            Left            =   4020
            MaxLength       =   2
            TabIndex        =   11
            Top             =   1
            Visible         =   0   'False
            Width           =   350
         End
         Begin VB.OptionButton optVar5Pct 
            Caption         =   "5%"
            Height          =   225
            Left            =   840
            TabIndex        =   6
            Top             =   30
            Width           =   540
         End
         Begin VB.OptionButton optVar10Pct 
            Caption         =   "10%"
            Height          =   225
            Left            =   1455
            TabIndex        =   7
            Top             =   30
            Width           =   660
         End
         Begin VB.OptionButton optVar15Pct 
            Caption         =   "15%"
            Height          =   225
            Left            =   2205
            TabIndex        =   8
            Top             =   30
            Width           =   630
         End
         Begin VB.OptionButton optVarOtherPct 
            Caption         =   "Other"
            Height          =   225
            Left            =   2925
            TabIndex        =   9
            Top             =   30
            Width           =   705
         End
         Begin VB.Label lblVariance 
            AutoSize        =   -1  'True
            Caption         =   "Var:"
            Height          =   195
            Left            =   3655
            TabIndex        =   10
            Top             =   30
            Visible         =   0   'False
            Width           =   285
         End
      End
   End
   Begin VB.Frame fraReport 
      Caption         =   "Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4320
      TabIndex        =   0
      Top             =   360
      Width           =   2940
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   2775
         TabIndex        =   48
         Top             =   200
         Width           =   2775
         Begin VB.OptionButton optMaterialRpt 
            Caption         =   "Material"
            Height          =   255
            Left            =   0
            TabIndex        =   1
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton optInstallRpt 
            Caption         =   "Installation"
            Height          =   255
            Left            =   940
            TabIndex        =   2
            Top             =   0
            Width           =   1095
         End
         Begin VB.OptionButton optTotalRpt 
            Caption         =   "Total"
            Height          =   255
            Left            =   2040
            TabIndex        =   3
            Top             =   0
            Value           =   -1  'True
            Width           =   735
         End
      End
   End
   Begin VB.Frame fraSelType 
      Caption         =   "Geographic Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   7320
      TabIndex        =   15
      Top             =   840
      Width           =   4575
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   4215
         TabIndex        =   51
         Top             =   240
         Width           =   4215
         Begin VB.OptionButton optAllCities 
            Caption         =   "All CCI Cities (731-Cities)"
            Height          =   255
            Left            =   2085
            TabIndex        =   19
            Top             =   280
            Width           =   2070
         End
         Begin VB.OptionButton optPriCity 
            Caption         =   "Primary Cities (316-Cities)"
            Height          =   255
            Left            =   2085
            TabIndex        =   17
            Top             =   0
            Width           =   2055
         End
         Begin VB.OptionButton optCCICities 
            Caption         =   "CCI Cities (727-Cities)"
            Height          =   255
            Left            =   0
            TabIndex        =   18
            Top             =   280
            Width           =   1875
         End
         Begin VB.OptionButton optNatlAvg 
            Caption         =   "Nat'l Avg (30-City)"
            Height          =   255
            Left            =   0
            TabIndex        =   16
            Top             =   0
            Value           =   -1  'True
            Width           =   1695
         End
      End
   End
   Begin VB.Frame fraClassSystemID 
      Caption         =   "Classification System"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   4320
      TabIndex        =   46
      Top             =   840
      Width           =   2940
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   2535
         TabIndex        =   50
         Top             =   240
         Width           =   2535
         Begin VB.OptionButton optClassSysRes 
            Caption         =   "Residential"
            Height          =   210
            Left            =   1320
            TabIndex        =   14
            Top             =   270
            Width           =   1095
         End
         Begin VB.OptionButton optClassSysUF 
            Caption         =   "Uniformat"
            Height          =   210
            Left            =   0
            TabIndex        =   13
            Top             =   270
            Width           =   1140
         End
         Begin VB.OptionButton optClassSysMF 
            Caption         =   "MasterFormat"
            Height          =   210
            Left            =   0
            TabIndex        =   12
            Top             =   0
            Value           =   -1  'True
            Width           =   1395
         End
      End
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview &Report"
      Height          =   495
      Left            =   120
      TabIndex        =   42
      Top             =   6120
      Width           =   1695
   End
   Begin VB.ComboBox cmbQuarterID 
      Height          =   315
      Left            =   7200
      TabIndex        =   36
      Top             =   2280
      Width           =   1005
   End
   Begin VB.Frame fraPeriod 
      Caption         =   "Index Period"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   6360
      TabIndex        =   31
      Top             =   2160
      Visible         =   0   'False
      Width           =   2880
      Begin VB.PictureBox Picture6 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   2655
         TabIndex        =   53
         Top             =   180
         Width           =   2655
         Begin VB.OptionButton optPeriodCurrent 
            Caption         =   "Current"
            Height          =   195
            Left            =   0
            TabIndex        =   32
            Top             =   0
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton optPeriodJan1 
            Caption         =   "Jan 1st"
            Height          =   195
            Left            =   840
            TabIndex        =   33
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton optPeriodHist 
            Caption         =   "Historical"
            Height          =   195
            Left            =   1680
            TabIndex        =   34
            Top             =   0
            Width           =   990
         End
      End
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   435
      Left            =   10440
      TabIndex        =   39
      Top             =   2280
      Width           =   1150
   End
   Begin VB.ComboBox cmbCity 
      Enabled         =   0   'False
      Height          =   315
      Left            =   7890
      TabIndex        =   25
      Top             =   1800
      Width           =   2055
   End
   Begin VB.ComboBox cmbState 
      Height          =   315
      Left            =   6555
      TabIndex        =   23
      Top             =   1800
      Width           =   765
   End
   Begin VB.TextBox Zip 
      Height          =   285
      Left            =   10440
      TabIndex        =   27
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox ClassificationID 
      Height          =   285
      Left            =   9690
      TabIndex        =   38
      Top             =   2280
      Width           =   585
   End
   Begin VB.ComboBox cmbCountry 
      Height          =   315
      Left            =   5040
      TabIndex        =   21
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Create Index Dtl Exception Report"
      Height          =   495
      Left            =   8400
      TabIndex        =   43
      Top             =   6120
      Width           =   2715
   End
   Begin ConstructionCostDatabase.DynaTree FormatTree 
      Height          =   2715
      Left            =   120
      TabIndex        =   44
      Top             =   0
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   4789
   End
   Begin VB.CheckBox ckbRowWrap 
      Caption         =   "Row Wrap"
      Height          =   315
      Left            =   120
      TabIndex        =   40
      Top             =   2880
      Width           =   1215
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid 
      Height          =   2715
      Left            =   120
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   3240
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   4789
      _LayoutType     =   0
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(0)._MinWidth=15"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowDelete     =   -1  'True
      DataMode        =   2
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   12632256
      RowDividerColor =   12632256
      RowSubDividerColor=   12632256
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H0&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=29,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=33"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=30,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=31,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=32"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=34"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=35"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=36"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=37,.parent=2,.namedParent=39"
      _StyleDefs(23)  =   "FilterBarStyle:id=40,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=11,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=20,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=12,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=13,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=14,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=16,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=15,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=17,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=18,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=19,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=38,.parent=37"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=41,.parent=40"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=24,.parent=11"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=11"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=12"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=15"
      _StyleDefs(44)  =   "Named:id=29:Normal"
      _StyleDefs(45)  =   ":id=29,.parent=0"
      _StyleDefs(46)  =   "Named:id=30:Heading"
      _StyleDefs(47)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(48)  =   ":id=30,.wraptext=-1"
      _StyleDefs(49)  =   "Named:id=31:Footing"
      _StyleDefs(50)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(51)  =   "Named:id=32:Selected"
      _StyleDefs(52)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=33:Caption"
      _StyleDefs(54)  =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(55)  =   "Named:id=34:HighlightRow"
      _StyleDefs(56)  =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(57)  =   "Named:id=35:EvenRow"
      _StyleDefs(58)  =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
      _StyleDefs(59)  =   "Named:id=36:OddRow"
      _StyleDefs(60)  =   ":id=36,.parent=29"
      _StyleDefs(61)  =   "Named:id=39:RecordSelector"
      _StyleDefs(62)  =   ":id=39,.parent=30"
      _StyleDefs(63)  =   "Named:id=42:FilterBar"
      _StyleDefs(64)  =   ":id=42,.parent=29"
   End
   Begin VB.Label lblFromQtr 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Quarter:"
      Height          =   225
      Left            =   6480
      TabIndex        =   35
      Top             =   2310
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   47
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ID:"
      Height          =   255
      Left            =   9360
      TabIndex        =   37
      Top             =   2310
      Width           =   240
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "State:"
      Height          =   255
      Left            =   5970
      TabIndex        =   22
      Top             =   1830
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "City:"
      Height          =   255
      Left            =   7410
      TabIndex        =   24
      Top             =   1830
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ZIP:"
      Height          =   255
      Left            =   10020
      TabIndex        =   26
      Top             =   1830
      Width           =   375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Country:"
      Height          =   255
      Left            =   4320
      TabIndex        =   20
      Top             =   1830
      Width           =   615
   End
   Begin VB.Label lblRowCount 
      Caption         =   "0 rows returned"
      Height          =   255
      Left            =   5340
      TabIndex        =   45
      Top             =   2880
      Width           =   3255
   End
   Begin VB.Line Line1 
      X1              =   4080
      X2              =   4080
      Y1              =   2730
      Y2              =   90
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   11040
      Y1              =   2820
      Y2              =   2820
   End
End
Attribute VB_Name = "frmCCIIdxDtlExcGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'<modulename> frmCCIIdxDtlExcGrid.frm</modulename>
'<functionname>General (Main) </functionname>
'
'<summary>
' (CCI) INDEX DETAIL EXCEPTION REPORT GRID:
'
'This window/form tracks the (%) progression/regression of MATERIALS, INSTALLATION COSTS AND TOTAL COSTS across:
'
'"   Current Qtr   (Material, Installation, Total) - Current Quarter of current yr.
'"   Current Year (Material, Installation, Total) - goes back to July of previous yr.
'"   Historically    (Material, Installation, Total) - goes back to 01/01/93!
'
'These percentages (%) can be displayed on a per MasterFormat Division basis with use of the menu tree to the upper left.
'Follow on selections can be performed as follows:
'
'"Variance":
'"   None
'"   5%
'"   10%
'"   15%
'"   Other
'
'"Classification System":
'"   MasterFormat (default)
'"   Res (residential)
'"   Uni (uniformat)
'
'"Geographic Selection" :
'"   NATL AVG (30 city)
'"   PRIMARY CITIES (316 cities)
'"   CCI CITIES (727 cities)
'"   ALL CITIES (731 cities)
'
'Select Dates:
'"   Quarter
'"   Period  (NOT SUPPORTED!!!)
'
'Index Period:
'"   Current
'"   Jan 1st
'"   Historical
'
'(By)
'
'1.  Quarter Id              (YYYYQN)
'2.  Country
'3.  City
'4.  State
'5.  Zip
'6.  ID
'
'
'(BUTTONS)
'"   SEARCH              (CmdSearch_Click() )
'
'Search for "index detail" data based upon Selections and filled in boxes:
'
'sp_select_published_cci_index_dtl_rpt_rlh
'
'"   Preview Report          (PreviewReport()  )
'rptCCIIndexDetail.XML
'"   Create Index Dtl Exception Report   (cmdCreate_Click())
'
'ExecStoredProcSelectedQuarter (" SP_REPORT_PUB_CCI_INDEX_MASTERFORMAT_RPT_WITH_FUEL_RLH ")
'
'NOTE: "Anytown" = 30 city average  and is displayed by selecting:
'"   Country = USA
'"   State = US
'"   City = Anytown
'
'COMPUTATIONAL NOTES:
'"   Jan 1 is actually July of previous year  (headers/captions need changing)
'"   % columns are:  tot col / 30 city value = %
'"   Historical reference date is:  01/01/93
'"   0 in the "Mat Total" means that there are no material values for that location (city, state or location)
'
'Key Subs / Functions:
'"   CmdSearch_Click()
'Prepares parameters to be passed with the stored procedure to retrieve needed "All Cities" or "Anytown" data
'
'HELPER Class: CCCIIdxExcMap.Cls
' </summary>
'
' <seealso> CCCIIdxExcMap.cls</seealso>
'<seealso> </seealso>
'
' <datastruct>m_rec</datastruct>
'<datastruct>m_objGridMap</datastruct>
'
' <storedprocedurename> sp_select_published_cci_index_dtl_rpt_rlh   </storedprocedurename>
'<storedprocedurename> sp_report_pub_cci_index_masterformat_rpt_with_fuel_rlh </storedprocedurename>
'
'
' <returns>N/A</returns>
' <exception>Always trap with an accompanying message box</exception>
' <example>
' <code>
'* * *
'SELECTED "MF", QUARTER_ID AND select_type (Nat'l Avg (30 city))
'VARIANCE=None (0), class_id = Masterformat division 023
'* * *
'exec sp_select_published_cci_index_dtl_rpt_rlh   @class_id = '023%', @quarter_id = '2006Q3', @variance_type = 'T', @variance = 0, @class_system_id = 'MF', @zip_3 = '%', @loc_id = 0, @country_code = '%', @state_code = '%', @select_type = 2
'</code>
' <code>
'    * * *
'            SELECTED "MF", QUARTER_ID, state_code='CA',
'select_type=2 (Nat'l Avg (30 city))
'* * *
'exec sp_select_published_cci_index_dtl_rpt_rlh   @class_id = '023%', @quarter_id = '2006Q3', @variance_type = 'T', @variance = 0, @class_system_id = 'MF', @zip_3 = '%', @loc_id = 400, @country_code = 'USA%', @state_code = 'CA%', @select_type = 2
'</code>
'<code>
'* * *  (ANYTOWN - loc_id=23)
'SELECTED "MF", QUARTER_ID, STATE_CODE='US' "
' * * *
'exec sp_select_published_cci_index_dtl_rpt_rlh   @class_id = '023%', @quarter_id = '2006Q3', @variance_type = 'T', @variance = 5, @class_system_id = 'MF', @zip_3 = '%', @loc_id = 23, @country_code = 'USA%', @state_code = 'US%', @select_type = 2
'</code>
'<code>* * * VARIANCE selected was 5%
'
'exec sp_select_published_cci_index_dtl_rpt_rlh   @class_id = '023%', @quarter_id = '2006Q3', @variance_type = 'T', @variance = 5, @class_system_id = 'MF', @zip_3 = '%', @loc_id = 0, @country_code = 'USA%', @state_code = '%', @select_type = 2
'
'NOTE: You can 't be sure that anything will be returned for different variances
'</code>
'<code> * * * HISTORICAL selected  (NOTE: I don't think "periods" are supported yet…
'
'exec sp_select_published_cci_index_dtl_rpt_rlh   @class_id = '023%', @quarter_id = '1993Q1', @variance_type = 'T', @variance = 5, @class_system_id = 'MF', @zip_3 = '%', @loc_id = 0, @country_code = 'USA%', @state_code = 'CA%', @select_type = 2
'</code>
'</example>
'<permission>Public</Permission>
'<dependson>This component depends on the following
'1.  CCCIIdxExcMap.cls
'2.  CGridMap.cls
'3.  CCDdal.CRSMDataAccess (
'4.  Access to the DAL (data access layer dll) opened in MainModule_Main() )
'</dependson>




Dim m_objGridMap As New CCCIIdxExcMap ' Class to handle grid
Dim m_blnFirstSearch As Boolean
Dim m_rec As New ADODB.RecordSet ' Recordset to hold query results
Dim m_blnDoubleClick As Boolean
Dim m_blnWereErrors As Boolean ' True if the Update had errors, used in QueryUnload
Dim m_strCurrentFormControl As String
Dim m_CurrentQtr As String
Dim m_FirstQtr As String
Dim m_State As String

Dim strRFQText As String
Dim strPrintContact As String
Dim blnSuppressPrices As Boolean
Dim blnSuppressAddressee As Boolean
Dim blnUseRecipientPrice As Boolean
Dim StartMatID As String

Private Function FillRptParm(sValue As String) As Variant
    If Len(sValue) = 0 Then
        FillRptParm = """"""
    Else
        FillRptParm = """" + SQLChangeWildcard(sValue) + """"
    End If
End Function

Private Function FindRcdType() As String
End Function

Private Function ReportType() As String
    If optMaterialRpt Then
        ReportType = "M"
    ElseIf optInstallRpt Then
        ReportType = "I"
    ElseIf optTotalRpt Then
        ReportType = "T"
    End If

End Function

Private Function SelectedVariance() As String
    ' 10/04/2005 RTD - ADDED SUPPORT FOR 'NONE' SELECTION (== 0%)
    If optVar0Pct Then
        SelectedVariance = 0
    ElseIf optVar5Pct Then
        SelectedVariance = 5
    ElseIf optVar10Pct Then
        SelectedVariance = 10
    ElseIf optVar15Pct Then
        SelectedVariance = 15
    ElseIf optVarOtherPct Then
        If Len(txtVariancePct) > 0 Then
            SelectedVariance = txtVariancePct
        Else
            SelectedVariance = 0
        End If
    End If
End Function

Public Sub Sort(intDir As Integer)
    m_objGridMap.Sort intDir
End Sub

Public Sub SelectAllRows()
    ' Pass recordset to handler class
    m_objGridMap.RecordSet = m_rec
    
    If m_rec.RecordCount > 0 Then
        m_objGridMap.SelectAllRows
    End If
End Sub

' Handles Row Wrap feature
Private Sub ckbRowWrap_Click()
    m_objGridMap.RowWrap (ckbRowWrap)
End Sub

Private Sub cmbState_Change()
    Dim iSelStart  As Integer
    Dim iSelLen As Integer
    
    iSelStart = cmbState.SelStart
    iSelLen = cmbState.SelLength
    cmbState = UCase(cmbState)
    cmbState.SelStart = iSelStart
    cmbState.SelLength = iSelLen
End Sub

Private Sub cmbState_Click()
    If cmbState.ListIndex = -1 Then
        cmbCity.ListIndex = -1
        cmbCity.Enabled = False
    Else
        cmbCity.Enabled = True
        If m_State <> cmbState.Text Then
            LoadCities cmbCity, cmbState.Text
        End If
    End If

End Sub

Private Sub cmbState_GotFocus()
    m_State = cmbState.Text
End Sub

Private Sub cmbState_LostFocus()
    If cmbState.ListIndex = -1 Then
        cmbCity.ListIndex = -1
        cmbCity.Enabled = False
    Else
        cmbCity.Enabled = True
        If m_State <> cmbState.Text Then
            LoadCities cmbCity, cmbState.Text
        End If
    End If

End Sub

Private Sub cmbState_Validate(Cancel As Boolean)
Dim i As Integer
    Dim bFound As Boolean
    If cmbState.Text <> "" Then
        For i = 0 To cmbState.listcount - 1
            If cmbState.Text = cmbState.List(i) Then
                cmbState.ListIndex = i
                bFound = True
                Exit For
            End If
        Next i
        If Not bFound Then
            MsgBox "Please enter valid state"
            Cancel = True
        End If
    End If
End Sub

Private Sub cmdCreate_Click()
    
    If ExecStoredProcSelectedQuarter("SP_REPORT_PUB_CCI_INDEX_MASTERFORMAT_RPT_WITH_FUEL_RLH") Then
        MsgBox "The Index Detail Exception table has been created.", vbInformation + vbOKOnly
    End If
    
End Sub

Private Function GetRFQTextID(ListRFQTextID As cdlgLstSel) As Long
Dim sql As String
Dim rec As ADODB.RecordSet
Dim varCurSelectedRow  As Variant
Dim blnResult As Boolean

'A list of available body texts for the user will
' be constructed, and the list selections populated from it.

ListRFQTextID.Caption = "Letter Text Selection"
ListRFQTextID.ComboCaption = "Select Text:"
ListRFQTextID.Check1Caption = "Suppress Prices"
ListRFQTextID.Check2Caption = "Suppress Addressee"

    sql = "exec sp_rfq_rpt_options_current_user"
    g_objDAL.GetRecordset CONNECT, sql, rec
    If rec.EOF And rec.BOF Then
        MsgBox "No texts have been set up for this report.  Please contact the IS department for help."
        GoTo Exit_Sub
    Else
        If rec.RecordCount = 0 Then     'invalid
            MsgBox "No contacts found."
        Else
            Do Until rec.EOF
                ListRFQTextID.AddUniqueItem rec.Fields("value_description"), 0, rec.Fields("report_option_value_id")
                rec.MoveNext
            Loop
        End If
        rec.Close
    End If

If ListRFQTextID.itemCount > 0 Then
    If ListRFQTextID.SetList = True Then
        Screen.MousePointer = vbNormal
        blnResult = ListRFQTextID.ShowList()
        Screen.MousePointer = vbHourglass
    End If
End If

If blnResult = True And ListRFQTextID.itemCount > 0 Then  'Contact selected or only 1 found - if none, ignore
    GetRFQTextID = ListRFQTextID.SingleItemData
Else
    GetRFQTextID = -1
End If

Exit_Sub:

End Function

Private Sub cmdPreview_Click()
    PreviewReport
End Sub

Private Sub Form_Activate()
    Dim ctl As Control
    If Me.WindowState <> vbMinimized Then
       If Len(m_strCurrentFormControl) > 0 Then
           For Each ctl In Me.Controls
               If ctl.Name = m_strCurrentFormControl Then
                   ctl.SetFocus
                   Exit For
               End If
           Next ctl
       End If
    '    TDBGrid.ReBind
       OutputView False
       ShowGridSort
       ShowToolbarIcons True
       m_objGridMap.SetMenuBar
    End If
End Sub

Private Sub Form_Deactivate()
    m_strCurrentFormControl = Me.ActiveControl.Name
    ShowToolbarIcons False
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim strSELECT As String
    Dim blnReturn As Boolean
    Move START_LEFT, START_TOP, START_WIDTH, START_HEIGHT
    LoadCombos Me, True, True, True
    m_CurrentQtr = g_sQuarterID
    m_FirstQtr = Left(g_sQuarterID, 5) & "1"
    ' This will never return any rows, just used to create recordset
    ClassificationID.Text = "~"
    cmdSearch_Click
    ClassificationID.Text = ""
    Status ("")
End Sub

Private Sub Form_Initialize()
    Status ("Loading CCI Index Detail Exception...")
    Screen.MousePointer = vbHourglass
    m_blnFirstSearch = True
    FormatTree.InitData g_cnShared, "CCI_INDEX"
    DoEvents    'Paint screen
   ' Initialize grid only once
    m_objGridMap.SetGrid TDBGrid
    m_objGridMap.InitGrid
    Screen.MousePointer = vbNormal
    m_blnFirstSearch = False
End Sub

Private Sub Form_LostFocus()
    TDBGrid.Update
    HideGridSort
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbNormal Or Me.WindowState = vbMaximized Then
        If Me.Width >= 11250 Then
            TDBGrid.Width = Me.Width - 255
            Line2.X2 = Me.Width - 210
        Else
            Me.Width = 11250
        End If
        If Me.Height >= 7260 Then
            TDBGrid.Height = Me.Height - 4545
            cmdCreate.Top = Me.Height - 1020
            cmdPreview.Top = Me.Height - 1020
        Else
            Me.Height = 7260
        End If
    Else
        ShowMinimizedForms
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    HideGridSort
    ShowToolbarIcons False
End Sub

' Leaf in MasterFormat tree selected.
Private Sub FormatTree_NodeSelected(ByVal strID As String)
    Dim rs As New ADODB.RecordSet
    Dim strSELECT As String
    Dim blnReturn As Boolean
    
    On Error Resume Next
    If m_blnFirstSearch = True Then
        m_blnFirstSearch = False
    Else
        ' 10/03/2005 RTD - REWROTE TO SUPPORT 'R1' NODE FOR RESI CLASS SYSTEM
        Select Case strID
        Case "U2"
            optClassSysUF.Value = True
            ClassificationID.Text = ""
        Case "MF"
            optClassSysMF.Value = True
            ClassificationID.Text = ""
        Case "R1"
            optClassSysRes.Value = True
            ClassificationID.Text = ""
        Case Else
            ClassificationID.Text = strID & "*"
        End Select
        
        ' Kick-off search
        cmdSearch_Click
    End If
End Sub

Private Sub cmdSearch_Click()
    On Error Resume Next
    Dim blnRet As Boolean
    Dim strSELECT As String
    Dim dtmToday As Date
    Dim dtmStart As Date
    Dim strStartMatSrch As String
    Dim iSelectType As Integer
    Dim sQtr As String
    Dim sMktCode As String
    Dim sClassSystemID  As String
    Dim sType As String
    Dim sReport As String
    Dim iPct As Integer
    
    TDBGrid.Update

    Screen.MousePointer = vbHourglass
    dtmToday = Date
    
    ' Synch tree with text box
    If Not ClassificationID.Text = "" Then
        FormatTree.FocusItem (ClassificationID.Text)
    End If
    If optQuarter Then
        If cmbQuarterID.ListIndex = -1 Then
            MsgBox "Please select a quarter."
            GoTo Exit_Sub
        Else
            sQtr = cmbQuarterID.Text
        End If
    Else
        If optPeriodHist Then
            sQtr = "1993Q1"
        ElseIf optPeriodJan1 Then
            sQtr = m_FirstQtr
        ElseIf optPeriodCurrent Then
            sQtr = m_CurrentQtr
        End If
    End If
    
     If Len(ClassificationID.Text) = 0 And Len(cmbCity.Text) = 0 And Len(cmbState.Text) = 0 And Len(Zip.Text) = 0 Then
        Screen.MousePointer = vbNormal
        MsgBox "You must enter search criteria before searching."
        GoTo Exit_Sub
    End If
    If optClassSysMF Then
        sClassSystemID = "MF"
    ElseIf optClassSysRes Then
        sClassSystemID = "R1"   ' 10/03/2005 RTD - SUPPORT RESIDENTIAL CLASS SYSTEM
    Else
        sClassSystemID = "U2"
    End If
    
    strSELECT = "exec sp_select_published_cci_index_dtl_rpt_rlh "
    strSELECT = strSELECT + "  @class_id = '" + FillWildCard(SQLChangeWildcard(ClassificationID)) + "'"
    strSELECT = strSELECT + ", @quarter_id = '" + sQtr + "'"
    strSELECT = strSELECT + ", @variance_type = '" + ReportType() + "'"
    strSELECT = strSELECT + ", @variance = " + SelectedVariance()
    strSELECT = strSELECT + ", @class_system_id = '" + sClassSystemID + "'"
    strSELECT = strSELECT + ", @zip_3 = '" + FillWildCard(SQLChangeWildcard(Zip.Text)) + "'"
    If cmbCity.ListIndex = -1 Then
        strSELECT = strSELECT + ", @loc_id = 0"
    Else
        strSELECT = strSELECT + ", @loc_id = " + CStr(cmbCity.ItemData(cmbCity.ListIndex))
    End If
    strSELECT = strSELECT + ", @country_code = '" + FillWildCard(cmbCountry.Text) + "'"
    strSELECT = strSELECT + ", @state_code = '" + FillWildCard(cmbState.Text) + "'"
    strSELECT = strSELECT + ", @select_type = " + GeographicType(Me)
    
    m_rec.Close ' Make sure it is closed
    ' 10/03/2005 RTD - MODIFIED MAX RECORDS RETURNED PER CR#1380
    m_rec.MaxRecords = MAX_RECORDS * 5 ' Set the maximum number to bring back
    dtmStart = Now
    ' Use g_objDAL to perform select
    
    If DEBUGON Then Stop
    
    blnRet = g_objDAL.GetRecordset(vbNullString, strSELECT, m_rec)
    If blnRet = False Then
        MsgBox "An error occurred while searching."
        lblRowCount.Caption = "0 rows returned."
        Screen.MousePointer = vbNormal
        Exit Sub
    End If
    
    ' Pass recordset to handler class
    m_objGridMap.RecordSet = m_rec
        
    If m_rec.RecordCount > 0 Then
        lblRowCount.Caption = str(m_rec.RecordCount) + " rows returned in " + str(DateDiff("s", dtmStart, Now)) + " seconds"
    Else
        lblRowCount.Caption = "0 rows returned."
    End If
    
    ' If the upper bound was hit, inform user
    If m_rec.RecordCount = MAX_RECORDS And m_rec.State = adStateOpen Then
        MsgBox "The search returned the maximum number of records allowed. More records may be available."
    End If
    
    ' Reset the grid contents
    TDBGrid.Bookmark = Null
    TDBGrid.ReBind
    TDBGrid.ApproxCount = m_rec.RecordCount
Exit_Sub:
    Screen.MousePointer = vbNormal

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    ' Check if there are pending changes
    If m_objGridMap.IsPendingChange = True Then
        Dim Button
        Button = MsgBox("Do you want to save your changes?", vbYesNoCancel)
        If Button = vbYes Then
            ' If there were errors, cancel the close
            If m_blnWereErrors Then
                Cancel = True
            End If
        ElseIf Button = vbCancel Then
            Cancel = True
            Exit Sub
        End If
    End If
End Sub

Private Sub optInstallRpt_Click()
    'cmdPreview.Caption = "Preview &Install"
End Sub

Private Sub optMaterialRpt_Click()
    'cmdPreview.Caption = "Preview &Material"
End Sub

Private Sub optPeriod_Click()
    fraPeriod.Visible = optPeriod.Value
    lblFromQtr.Visible = Not optPeriod.Value
    cmbQuarterID.Visible = Not optPeriod.Value
End Sub

Private Sub optQuarter_Click()
    fraPeriod.Visible = Not optQuarter.Value
    lblFromQtr.Visible = optQuarter.Value
    cmbQuarterID.Visible = optQuarter.Value
End Sub

Private Sub optTotalRpt_Click()
    'cmdPreview.Caption = "Preview &Totals"
End Sub

Private Sub optVar10Pct_Click()
    ShowOtherVariance
End Sub

Private Sub optVar15Pct_Click()
    ShowOtherVariance
End Sub

Private Sub optVar5Pct_Click()
    ShowOtherVariance
End Sub
Private Sub ShowOtherVariance()
    lblVariance.Visible = optVarOtherPct.Value
    txtVariancePct.Visible = optVarOtherPct.Value
    On Error Resume Next
    If txtVariancePct.Visible Then
        txtVariancePct.SetFocus
    End If
End Sub

Private Sub optVarOtherPct_Click()
    ShowOtherVariance
End Sub

Private Sub TDBGrid_DblClick()
    ' Signal that double-click has occurred
    m_blnDoubleClick = True
End Sub

Private Sub TDBGrid_Error(ByVal DataError As Integer, Response As Integer)
    Response = 0
    TDBGrid.DataChanged = False
End Sub

Private Sub TDBGrid_GotFocus()
    TDBGrid.TabStop = True
End Sub

Private Sub TDBGrid_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyBack Then
        If TDBGrid.Col = 2 Or TDBGrid.Col = 3 Then
            If Len(TDBGrid.Text) + 1 > 75 Then
                KeyAscii = 0
            End If
        End If
    End If
End Sub

Private Sub TDBGrid_LostFocus()
    TDBGrid.TabStop = False
End Sub

Private Sub TDBGrid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If this is the mouse-up form a double click
    If m_blnDoubleClick Then
    Else
        If Button = vbRightButton And IsNumeric(TDBGrid.Bookmark) Then
            Dim strErrorMsg As String
            strErrorMsg = m_objGridMap.GetError(TDBGrid.Bookmark)
            If Len(strErrorMsg) > 0 Then
                MsgBox strErrorMsg
            End If
        End If
    End If
End Sub

Private Sub txtVariancePct_Validate(Cancel As Boolean)
    Dim bError As Boolean
    If Len(txtVariancePct.Text) = 0 Then Exit Sub     'Valid
    If txtVariancePct.Text = 0 Then Exit Sub
    If IsNumeric(txtVariancePct.Text) Then
        If CInt(txtVariancePct.Text) > 99 Or CInt(txtVariancePct.Text) < 1 Then
            bError = True
        End If
    Else
        bError = True
    End If
    If bError Then
        Cancel = True
        MsgBox "Please enter a valid percentage between 1 and 99."
    End If
End Sub

Public Sub PrintReport()
    PreviewReport
End Sub

Public Sub PreviewReport()
    Dim fPreviewWindow As New frmReportPreview
    Dim sReport As String
    
    If optMaterialRpt.Value Then
        sReport = "Material Exceptions"
    ElseIf optInstallRpt.Value Then
        sReport = "Installation Exceptions"
    Else
        sReport = "Totals Exceptions"
    End If
    
    If cmbCity.Text <> "" Then
        sReport = sReport & " by City"
    End If
    
    If m_rec.RecordCount >= 1 Then
        fPreviewWindow.ReportName = sReport
        fPreviewWindow.ReportFile = "rptCCIIndexDetail.xml"
        fPreviewWindow.OpenEvent = "select_type_sel = """ & GeographicType(Me) & """" & vbCrLf & _
                                    "variance_sel = """ & SelectedVariance & """"
        fPreviewWindow.RecordSet = m_rec
        fPreviewWindow.RenderReport
        fPreviewWindow.Show
    Else
        MsgBox "Please choose or search for a CCI index.", vbInformation + vbOKOnly, "Warning"
    End If
End Sub

Private Sub ShowToolbarIcons(bShowIcons As Boolean)

    fMainForm.tbToolBar.Buttons.Item(tbrPRINT).Enabled = bShowIcons
    fMainForm.tbToolBar.Buttons.Item(tbrPRINT).Visible = bShowIcons
    fMainForm.tbToolBar.Buttons.Item(tbrPREVIEW).Enabled = bShowIcons
    fMainForm.tbToolBar.Buttons.Item(tbrPREVIEW).Visible = bShowIcons
    fMainForm.tbToolBar.Buttons.Item(tbrEXPORTDATA).Enabled = bShowIcons
    fMainForm.tbToolBar.Buttons.Item(tbrEXPORTDATA).Visible = bShowIcons
    fMainForm.tbToolBar.Buttons.Item(tbrEXPORTDATA + 1).Visible = bShowIcons
    fMainForm.mnuFilePageSetup.Enabled = bShowIcons
    fMainForm.mnuFilePrint.Enabled = bShowIcons
    fMainForm.mnuFilePrintPreview.Enabled = bShowIcons

End Sub

Public Sub ExportData()
    
    If m_rec.RecordCount > 0 Then
        Dim fExport As New frmExport
        fExport.SetRow TDBGrid, m_rec
        fExport.Title = "CCI Index Detail Exception"
        fExport.Show
    Else
        MsgBox "Please choose or search for a CCI index.", vbInformation + vbOKOnly
    End If
    
End Sub
