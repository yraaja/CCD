VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{5936A75C-3F42-11D6-AF6B-AA0004005F12}#1.3#0"; "MeansCtrl.ocx"
Begin VB.Form frmCCICompUsageGrid 
   Caption         =   "CCI Component Usage Maintenance Grid"
   ClientHeight    =   6855
   ClientLeft      =   2265
   ClientTop       =   2835
   ClientWidth     =   12015
   Icon            =   "frmCCICompUsageGrid.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6855
   ScaleWidth      =   12015
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Create Quarterly Report Table"
      Height          =   495
      Left            =   4680
      TabIndex        =   41
      Top             =   6240
      Width           =   3075
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export"
      Height          =   495
      Left            =   360
      TabIndex        =   40
      Top             =   6240
      Width           =   1150
   End
   Begin VB.Frame fraSelType 
      Caption         =   "Geographic Selection"
      Height          =   945
      Left            =   6120
      TabIndex        =   4
      Top             =   360
      Width           =   4335
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   240
         ScaleHeight     =   615
         ScaleWidth      =   3975
         TabIndex        =   37
         Top             =   240
         Width           =   3975
         Begin VB.OptionButton optAllCities 
            Caption         =   "All CCI Cities (731-Cities)"
            Height          =   210
            Left            =   1845
            TabIndex        =   8
            Top             =   300
            Width           =   2070
         End
         Begin VB.OptionButton optPriCity 
            Caption         =   "Primary Cities (316-Cities)"
            Height          =   210
            Left            =   1845
            TabIndex        =   6
            Top             =   0
            Width           =   2055
         End
         Begin VB.OptionButton optCCICities 
            Caption         =   "CCI Cities (727-Cities)"
            Height          =   210
            Left            =   0
            TabIndex        =   7
            Top             =   300
            Width           =   1875
         End
         Begin VB.OptionButton optNatlAvg 
            Caption         =   "Nat'l Avg (30-City)"
            Height          =   210
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Value           =   -1  'True
            Width           =   1695
         End
      End
   End
   Begin VB.Frame fraClassSystemID 
      Caption         =   "Classification System"
      Height          =   945
      Left            =   4320
      TabIndex        =   0
      Top             =   360
      Width           =   1740
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   1575
         TabIndex        =   36
         Top             =   240
         Width           =   1575
         Begin VB.OptionButton optResidential 
            Caption         =   "Res"
            Height          =   210
            Left            =   0
            TabIndex        =   3
            Top             =   300
            Width           =   615
         End
         Begin VB.OptionButton optClassSysUF 
            Caption         =   "Uni"
            Height          =   210
            Left            =   840
            TabIndex        =   2
            Top             =   300
            Width           =   780
         End
         Begin VB.OptionButton optClassSysMF 
            Caption         =   "Master Format"
            Height          =   210
            Left            =   0
            TabIndex        =   1
            Top             =   0
            Value           =   -1  'True
            Width           =   1395
         End
      End
   End
   Begin VB.Frame fraConstType 
      Caption         =   "Records"
      ForeColor       =   &H00404040&
      Height          =   1455
      Left            =   10560
      TabIndex        =   9
      Top             =   360
      Width           =   1335
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   120
         ScaleHeight     =   1095
         ScaleWidth      =   1095
         TabIndex        =   38
         Top             =   240
         Width           =   1095
         Begin VB.OptionButton optRcdsLabor 
            Caption         =   "Labor"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   13
            Top             =   840
            Width           =   735
         End
         Begin VB.OptionButton optRcdsEquip 
            Caption         =   "Equipment"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   12
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton optRcdsAll 
            Caption         =   "All"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   10
            Top             =   0
            Value           =   -1  'True
            Width           =   495
         End
         Begin VB.OptionButton optRcdsMatl 
            Caption         =   "Material"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   11
            Top             =   360
            Width           =   855
         End
      End
   End
   Begin VB.ComboBox cmbCity 
      Enabled         =   0   'False
      Height          =   315
      Left            =   7770
      TabIndex        =   19
      Top             =   1440
      Width           =   1815
   End
   Begin VB.ComboBox cmbState 
      Height          =   315
      Left            =   6555
      TabIndex        =   17
      Top             =   1440
      Width           =   645
   End
   Begin VB.TextBox Zip 
      Height          =   285
      Left            =   9930
      TabIndex        =   21
      Top             =   1470
      Width           =   495
   End
   Begin VB.ComboBox cmbCountry 
      Height          =   315
      Left            =   5040
      TabIndex        =   15
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox UseID 
      Height          =   285
      Left            =   6480
      TabIndex        =   25
      Top             =   2010
      Width           =   975
   End
   Begin VB.ComboBox cmbQuarterID 
      Height          =   315
      Left            =   5040
      TabIndex        =   23
      Top             =   1980
      Width           =   1005
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Enabled         =   0   'False
      Height          =   495
      Left            =   9960
      TabIndex        =   32
      Top             =   6240
      Visible         =   0   'False
      Width           =   1150
   End
   Begin ConstructionCostDatabase.DynaTree FormatTree 
      Height          =   2430
      Left            =   60
      TabIndex        =   33
      Top             =   60
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   4286
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   435
      Left            =   10440
      TabIndex        =   29
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CheckBox ckbRowWrap 
      Caption         =   "Row Wrap"
      Height          =   315
      Left            =   120
      TabIndex        =   30
      Top             =   2610
      Width           =   1215
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid 
      Height          =   3180
      Left            =   120
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   2970
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   5609
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=29,.bgcolor=&H80000005&,.bold=0,.fontsize=825"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
   Begin VB.Frame Frame1 
      Height          =   495
      Left            =   7800
      TabIndex        =   26
      Top             =   1920
      Width           =   2415
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   2055
         TabIndex        =   39
         Top             =   120
         Width           =   2055
         Begin VB.OptionButton optSummary 
            Caption         =   "Summary"
            Height          =   315
            Left            =   0
            TabIndex        =   27
            Top             =   0
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optDetail 
            Caption         =   "Detail"
            Height          =   315
            Left            =   1200
            TabIndex        =   28
            Top             =   0
            Width           =   975
         End
      End
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "State:"
      Height          =   255
      Left            =   5970
      TabIndex        =   16
      Top             =   1500
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "City:"
      Height          =   255
      Left            =   7290
      TabIndex        =   18
      Top             =   1500
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Zip:"
      Height          =   255
      Left            =   9540
      TabIndex        =   20
      Top             =   1500
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "Country:"
      Height          =   255
      Left            =   4320
      TabIndex        =   14
      Top             =   1500
      Width           =   615
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   10680
      Y1              =   2550
      Y2              =   2550
   End
   Begin VB.Label lblFromQtr 
      Alignment       =   2  'Center
      Caption         =   "Quarter:"
      Height          =   255
      Left            =   4320
      TabIndex        =   22
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label lblRowCount 
      Caption         =   "0 rows returned"
      Height          =   255
      Left            =   5400
      TabIndex        =   35
      Top             =   2610
      Width           =   3255
   End
   Begin VB.Line Line1 
      X1              =   4200
      X2              =   4200
      Y1              =   2400
      Y2              =   120
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "ID:"
      Height          =   255
      Left            =   6135
      TabIndex        =   24
      Top             =   2040
      Width           =   255
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
      TabIndex        =   34
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmCCICompUsageGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'<modulename> frmCCICompUsageGrid.frm</modulename>
'<functionname>General (Main) </functionname>
'
'<summary>
' (CCI) Component Usage Maintenance GRID:
'
'"   This window/form tracks:  Material, Equipment, and Labor costs in real dollars
'"   By MF division, Class ID (MF division ##), and CCI ID
'"   It breaks report down by "detail" and "summary" lines  (see "Typ" column in grid)
'
'
'"Geographic Selection":
'"   NATL AVG (30 city)
'"   PRIMARY CITIES (316 cities)
'"   CCI CITIES (727 cities)
'"   ALL CITIES (731 cities)
'
'Records:
'"   All
'"   Material
'"   Equipment
'"   Labor
'
'(By)
'
'1.  Country
'2.  City
'3.  State
'4.  Zip
'5.  Quarter Id              (YYYYQN)
'6.  ID
'7.  Summary/Detail  (Report Format)
'Blue lines = summary lines  (Typ=(T)otal)
'White lines = detail lines      (Typ=(E)lectric, (M)aterial, (L)abor
'
'
'(BUTTONS)
'"   SEARCH              (CmdSearch_Click() )
'
'Search for "index detail" data based upon Selections and filled in boxes:
'
'sp_select_published_cci_index_dtl_rpt_rlh
'
'"   Export                  (Export_Click()  )
'frmExport
'rptCCIIndexDetail.XML
'"   Create Quarterly Report TAble   (cmdCreate_Click())
'
'ExecStoredProcSelectedQuarter (" SP_REPORT_CCI_DETAIL_COMPONENT_USAGE_REPORT")
'
'"   Delete                  (cmdDelete_Click() )
'
'NOTE: "Anytown" = 30 city average  and is displayed by selecting:
'"   Country = USA
'"   State = US
'"   City = Anytown
'
'Key Subs / Functions:
'"   CmdSearch_Click()
'Prepares parameters to be passed with the stored procedure to retrieve needed "All Cities" or "Anytown" data
'
'HELPER Class: CCCICompUseMap.Cls
' </summary>
'
' <seealso> CCCICompUseMap.cls</seealso>
'<seealso> </seealso>
'
' <datastruct>m_rec</datastruct>
'<datastruct>m_objGridMap</datastruct>
'
' <storedprocedurename> sp_select_cci_comp_usage_ksr </storedprocedurename>
'<storedprocedurename>sp_report_cci_detail_component_usage_report
'</storedprocedurename>
'
'
' <returns>N/A</returns>
' <exception>Always trap with an accompanying message box</exception>
' <example>
' <code>
'* * *
'"   SPECIFIED  QUARTER_ID ONLY!
'AND
'"   (default: Master Format classification
'"   (default: Nat'l Avg (30-City) )
'"   (default: Record (All) )
'* * *
'exec sp_select_cci_comp_usage_ksr   @class_id = '', @class_system_id = 'MF', @quarter_id = '2006Q4', @summary = '1', @loc_id = 0, @state_code = '%', @country_code = '%', @select_rcd_types = 'A', @select_type = 2
'</code>
' <code>
'    * * *
'            SELECTED Master Format "MF", QUARTER_ID, state_code='CA',
'select_type=2 (Nat'l Avg (30 city)) , All Records
'* * *
'exec sp_select_cci_comp_usage_ksr   @class_id = '', @class_system_id = 'MF', @quarter_id = '2006Q4', @summary = '1', @loc_id = 0, @state_code = 'CA%', @country_code = 'USA%', @select_rcd_types = 'A', @select_type = 2
'</code>
'<code>
'* * *  SELECTED "MF", QUARTER_ID, STATE_CODE='CA'
'RECORDS= (M)aterial, Summary format (summary='1')
' * * *
'exec sp_select_cci_comp_usage_ksr   @class_id = '', @class_system_id = 'MF', @quarter_id = '2006Q4', @summary = '1', @loc_id = 0, @state_code = 'CA%', @country_code = 'USA%', @select_rcd_types = 'M', @select_type = 2
'</code>
'<code>
'* * * SELECTED "MF", QUARTER_ID, STATE_CODE='CA'
'RECORDS= (E)quipment, Detail Format (summary='0')
'* * *
'exec sp_select_cci_comp_usage_ksr   @class_id = '', @class_system_id = 'MF', @quarter_id = '2006Q4', @summary = '0', @loc_id = 0, @state_code = 'CA%', @country_code = 'USA%', @select_rcd_types = 'E', @select_type = 2
'</code>
'<code>
'* * * SELECTED "MF", QUARTER_ID, STATE_CODE='CA'
'Master Format Division = 23,
'RECORDS= (E)quipment, Detail Format (summary='0')
'* * *
'exec sp_select_cci_comp_usage_ksr   @class_id = '023%', @class_system_id = 'MF', @quarter_id = '2006Q4', @summary = '0', @loc_id = 0, @state_code = 'CA%', @country_code = 'USA%', @select_rcd_types = 'E', @select_type = 2
'</code>
'<code>
'* * * SELECTED "MF", QUARTER_ID, STATE_CODE='CA'
'Master Format Division = 23,
'RECORDS= ALL, Detail Format (summary='0')
'* * *
'exec sp_select_cci_comp_usage_ksr   @class_id = ' ', @class_system_id = 'MF', @quarter_id = '2006Q4', @summary = '0', @loc_id = 0, @state_code = 'CA%', @country_code = 'USA%', @select_rcd_types = 'A', @select_type = 2
'</code>
'</example>
'<permission>Public</Permission>
'<dependson>This component depends on the following
'1.  CCCICompUseMap.cls
'2.  CGridMap.cls
'3.  CCDdal.CRSMDataAccess (
'4.  Access to the DAL (data access layer dll) opened in MainModule_Main() )
'</dependson>



Dim m_objGridMap As New CCCICompUseMap ' Class to handle grid
Dim m_blnFirstSearch As Boolean
Dim m_rec As New ADODB.RecordSet ' Recordset to hold query results
Dim m_blnDoubleClick As Boolean
Dim m_blnWereErrors As Boolean ' True if the Update had errors, used in QueryUnload
Dim m_strCurrentFormControl As String
Dim m_CurrentQtr As String
Dim m_State As String
Dim m_rsUsageClone As New ADODB.RecordSet

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

Private Sub cmdClone_Click()
    ExecStoredProcSelectedQuarter "sp_clone_pub_cci_material_price"
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
If DEBUGON Then Stop
        
    ExecStoredProcSelectedQuarter "SP_REPORT_CCI_DETAIL_COMPONENT_USAGE_REPORT" '03/03/2010

End Sub

Private Sub cmdDelete_Click()
    Dim varButton
    varButton = MsgBox("Are you sure you want to delete?", vbYesNo + vbCritical)
    If varButton = vbYes Then
        TDBGrid.Delete
    End If

End Sub

Private Sub cmdNew_Click()
Dim blnUnitCost As Boolean
Dim blnAssembly As Boolean
Dim bln_Continue As Boolean
Dim varCurrentM_recBookmark As Variant

'If IsNull(TDBGrid.Bookmark) Then
        bln_Continue = True
'Else
''    If ValidGridRow() = True Then
''        bln_Continue = True
''    End If
'End If
'
On Error GoTo Exit_Sub
    
    If TDBGrid.DataChanged = True Then
'    If TDBGrid.AddNewMode = dbgAddNewCurrent Then
        TDBGrid.Update
    End If

On Error Resume Next

If bln_Continue = True Then
    TDBGrid.SetFocus
    TDBGrid.MoveLast
    TDBGrid.AllowAddNew = True
    TDBGrid.Row = TDBGrid.Row + 1
    
    TDBGrid.Split = 0
    m_rec.AddNew
    m_rec.MoveLast
    varCurrentM_recBookmark = m_rec.Bookmark
    ' Defaults for new added row
    
    m_rec.Fields("cci_skey") = 0
    m_rec.Fields("quarter_id") = cmbQuarterID.Text
    m_rec.Fields("rec_type") = "M"
    'Select Class System
        If optClassSysMF.Value = True Then
            m_rec.Fields("class_system_id") = "MF"
        ElseIf optClassSysUF.Value = True Then
            m_rec.Fields("class_system_id") = "U2"
        ElseIf optResidential.Value = True Then
            m_rec.Fields("class_system_id") = "R1"
        End If
        m_rec.Fields("quarter_id") = cmbQuarterID.Text

    m_objGridMap.SetRowState m_rec.Bookmark, STATE_NEW
        
    TDBGrid.SetFocus
    TDBGrid.AllowAddNew = False
    TDBGrid.ReOpen m_rec.Bookmark
 '   m_rec.MoveLast
    DoEvents
    
'Select Class System
    If optClassSysMF.Value = True Then
        TDBGrid.Columns("Sys").Text = "MF"
    ElseIf optClassSysUF.Value = True Then
        TDBGrid.Columns("Sys").Text = "U2"
    ElseIf optResidential.Value = True Then
        TDBGrid.Columns("Sys").Text = "R1"
    End If
    TDBGrid.Columns("Type").Text = "M" 'Dft to material
    TDBGrid.Columns("Quarter").Text = cmbQuarterID.Text
    TDBGrid.Columns("cci_skey").Text = 0
    
    TDBGrid.AllowAddNew = False
    TDBGrid.Col = TDBGrid.Columns("Cls ID").ColIndex
'    m_objGridMap.SetRowState TDBGrid.Bookmark, STATE_NEW
End If
TDBGrid.SetFocus

Exit_Sub:

End Sub

Private Sub cmdUpdate_Click()
    On Error GoTo Out
    Dim blnRet As Boolean
    Dim vntBookmark As Variant
    m_blnWereErrors = False
    
    vntBookmark = TDBGrid.Bookmark
    TDBGrid.Update
    blnRet = m_objGridMap.Update
    If blnRet = False Then
        m_blnWereErrors = True
    End If
    TDBGrid.Bookmark = vntBookmark
Out:
End Sub



Private Sub cmdExport_Click()

    Dim fExport As New frmExport
    
    If m_rec.RecordCount >= 1 Then
        fExport.SetRow TDBGrid, m_rec
        fExport.Title = "CCI Component Usage"
        fExport.Show
    Else
        MsgBox "Please choose or search for a CCI Component usage row.", vbInformation + vbOKOnly
    End If
    

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
       m_objGridMap.SetMenuBar
    End If
End Sub

Private Sub Form_Deactivate()
m_strCurrentFormControl = Me.ActiveControl.Name
End Sub


Private Sub Form_Load()
    On Error Resume Next
    Dim strSELECT As String
    Dim blnReturn As Boolean
    Move START_LEFT, START_TOP, START_WIDTH, START_HEIGHT
    LoadCombos Me, True, True, True

    ' This will never return any rows, just used to create recordset
    UseID.Text = "~"
    cmdSearch_Click
    UseID.Text = ""
    Status ("")
End Sub

Private Sub Form_Initialize()
    ' 10/03/2005 RTD - CORRECTED INCORRECT STATUS MESSAGE
    Status ("Loading CCI Component Usage Maintenance...")
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
'            TDBGrid.Height = Me.Height - 3745
            TDBGrid.Height = Me.Height - 4545
'            cmdNew.Top = Me.Height - 1020
'            cmdUpdate.Top = Me.Height - 1020
'            cmdDelete.Top = Me.Height - 1020
            Me.cmdExport.Top = Me.Height - 1020   'rlh 03/03/2010
            Me.cmdCreate.Top = Me.Height - 1020
        Else
            Me.Height = 7260
        End If
    Else
        ShowMinimizedForms
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
HideGridSort
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
        If strID = "U2" Then
            optClassSysUF = True
            UseID.Text = ""
        ElseIf strID = "MF" Then
            optClassSysMF = True
            UseID.Text = ""
        Else
            UseID.Text = strID & "*"
        End If
        ' Kick-off search
        cmdSearch_Click
    End If
End Sub

Private Sub cmdSearch_Click()
    On Error Resume Next
    Dim blnRet As Boolean
    Dim strSELECT As String
    Dim sSelectType As String
    Dim sClassSys As String
    Dim iSummaryFlag As Integer
    Dim dtmToday As Date
    Dim dtmStart As Date
    Dim strStartMatSrch As String
    Dim iSelectType As Integer
    TDBGrid.Update

    If m_objGridMap.IsPendingChange = True Then
        Dim Button
        Button = MsgBox("Do you want to save your changes?", vbYesNoCancel)
        If Button = vbYes Then
            cmdUpdate_Click
            ' If there were errors, cancel the search
            If m_blnWereErrors Then
                Exit Sub
            End If
        ElseIf Button = vbNo Then
            TDBGrid.DataChanged = False
        ElseIf Button = vbCancel Then
            ' Cancel the search
            Exit Sub
        End If
    End If
    Screen.MousePointer = vbHourglass
    dtmToday = Date
    
    ' Synch tree with text box
    If Not UseID.Text = "" Then
        FormatTree.FocusItem (UseID.Text)
    End If
    
    If cmbQuarterID.ListIndex = -1 Then
        MsgBox "Please select a quarter."
        Exit Sub
    End If
    If Len(UseID.Text) = 0 _
           And Len(cmbQuarterID.Text) = 0 _
           And Len(cmbCity.Text) = 0 _
           And Len(cmbState.Text) = 0 Then
        Screen.MousePointer = vbNormal
        MsgBox "You must enter search criteria before searching."
        Exit Sub
    End If
    
'Select Class System
    If optClassSysMF.Value = True Then
        sClassSys = "MF"
    ElseIf optClassSysUF.Value = True Then
        sClassSys = "U2"
    ElseIf optResidential.Value = True Then
        sClassSys = "R1"
    End If

'Select Record types for Result Recordset
    If optRcdsAll.Value = True Then
        sSelectType = "A"
    ElseIf optRcdsMatl.Value = True Then
        sSelectType = "M"
    ElseIf optRcdsEquip.Value = True Then
        sSelectType = "E"
    ElseIf optRcdsLabor.Value = True Then
        sSelectType = "L"
    End If
    
    If optSummary.Value = True Then
        iSummaryFlag = 1
    Else
        iSummaryFlag = 0
    End If
    
'    strSELECT = "exec sp_select_cci_comp_usage_rlh "
    strSELECT = "exec sp_select_cci_comp_usage_ksr "
    strSELECT = strSELECT + "  @class_id = '" + SQLChangeWildcard(UseID) + "'"
    strSELECT = strSELECT + ", @class_system_id = '" + sClassSys + "'"
    strSELECT = strSELECT + ", @quarter_id = '" + cmbQuarterID.Text + "'"
    strSELECT = strSELECT + ", @summary = '" + CStr(iSummaryFlag) + "'"
    If cmbCity.ListIndex = -1 Then
        strSELECT = strSELECT + ", @loc_id = 0"
    Else
        strSELECT = strSELECT + ", @loc_id = " + CStr(cmbCity.ItemData(cmbCity.ListIndex))
    End If
    strSELECT = strSELECT + ", @state_code = '" + FillWildCard(cmbState.Text) + "'"
    strSELECT = strSELECT + ", @country_code = '" + FillWildCard(cmbCountry.Text) + "'"
    strSELECT = strSELECT + ", @select_rcd_types = '" + sSelectType + "'"
    strSELECT = strSELECT + ", @select_type = " + GeographicType(Me)
    m_rec.Close ' Make sure it is closed
    m_rec.MaxRecords = MAX_RECORDS ' Set the maximum number to bring back
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
    Set m_rsUsageClone = m_rec.Clone

    ' Pass recordset to handler class
    m_objGridMap.RecordSet = m_rec
    Debug.Print "Record Count: " & m_rec.RecordCount  'rlh
        
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
    Screen.MousePointer = vbNormal

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    ' Check if there are pending changes
    If m_objGridMap.IsPendingChange = True Then
        Dim Button
        Button = MsgBox("Do you want to save your changes?", vbYesNoCancel)
        If Button = vbYes Then
            cmdUpdate_Click
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

Private Sub optDetail_Click()
'MsgBox ("(optDetail_Click)Record count: " & m_rec.RecordCount)   'rlh
Debug.Print "(optDetail_Click)Record count: " & m_rec.RecordCount
End Sub

Private Sub TDBGrid_DblClick()
    ' Signal that double-click has occurred
    m_blnDoubleClick = True
End Sub

Private Sub TDBGrid_Error(ByVal DataError As Integer, Response As Integer)
    Response = 0
    TDBGrid.DataChanged = False
End Sub

'*** APEX Migration Utility Code Change ***
'Private Sub TDBGrid_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueOleDBGrid70.StyleDisp)
Private Sub TDBGrid_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueOleDBGrid80.StyleDisp)
    On Error Resume Next

    Dim bLocked As Boolean
    Dim recClone As ADODB.RecordSet
    If IsNumeric(Bookmark) Then
        Set recClone = m_rec.Clone
        recClone.Bookmark = Bookmark
        With TDBGrid.Columns(Col)
            Select Case .Caption
                Case "Cls ID", "CCI ID", "Type"
                Debug.Print "(FetchCellStyle)Record Count:" + m_rec.RecordCount
                If recClone.Fields("last_update_person") = "" And recClone.Fields("rec_type") <> "T" Then   'New record
                    bLocked = False
                Else
                    bLocked = True
                End If
            End Select
            If bLocked = True Then
                bLocked = True
                CellStyle.Locked = True
                CellStyle.ForeColor = vbGrayText
            Else
                CellStyle.ForeColor = vbBlack
                CellStyle.Locked = False
            End If
        End With
        recClone.Close
        Set recClone = Nothing
    End If
    
    ' If the row is highlighted, then let it be
    If (Condition And dbgSelectedRow) = 8 Then
            CellStyle.ForeColor = vbWhite
        Exit Sub
    End If

End Sub

'*** APEX Migration Utility Code Change ***
'Private Sub TDBGrid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
Private Sub TDBGrid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid80.StyleDisp)
On Error Resume Next
m_rsUsageClone.Bookmark = Bookmark
Select Case m_rsUsageClone.Fields("rec_type").Value
        Case "E", "L", "M"  'Regular lines
            RowStyle.Locked = False
            RowStyle.ForeColor = vbBlack
        Case "T"
            RowStyle.Locked = True
            RowStyle.Font.Bold = True
            RowStyle.ForeColor = vbGrayText
            RowStyle.BackColor = "15658689"
End Select

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

