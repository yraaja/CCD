VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{5936A75C-3F42-11D6-AF6B-AA0004005F12}#1.3#0"; "MeansCtrl.ocx"
Begin VB.Form frmCCICSIFmtSumRptGrid 
   Caption         =   "CCI Dollar Listing"
   ClientHeight    =   6750
   ClientLeft      =   2265
   ClientTop       =   2835
   ClientWidth     =   11340
   Icon            =   "frmCCICSIFmtSumRptGrid.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6750
   ScaleWidth      =   11340
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
      Left            =   6240
      TabIndex        =   15
      Top             =   600
      Width           =   4335
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   4095
         TabIndex        =   25
         Top             =   240
         Width           =   4095
         Begin VB.OptionButton optAllCities 
            Caption         =   "All CCI Cities (731-Cities)"
            Height          =   195
            Left            =   1920
            TabIndex        =   32
            Top             =   300
            Width           =   2070
         End
         Begin VB.OptionButton optPriCity 
            Caption         =   "Primary Cities (316-Cities)"
            Height          =   255
            Left            =   1920
            TabIndex        =   31
            Top             =   0
            Width           =   2055
         End
         Begin VB.OptionButton optCCICities 
            Caption         =   "CCI Cities (727-Cities)"
            Height          =   255
            Left            =   0
            TabIndex        =   30
            Top             =   285
            Width           =   1875
         End
         Begin VB.OptionButton optNatlAvg 
            Caption         =   "Nat'l Avg (30-City)"
            Height          =   255
            Left            =   0
            TabIndex        =   29
            Top             =   0
            Value           =   -1  'True
            Width           =   1695
         End
      End
   End
   Begin VB.Frame fraClassSystemID 
      Caption         =   "Classification System"
      Height          =   825
      Left            =   4440
      TabIndex        =   16
      Top             =   600
      Width           =   1740
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   1575
         TabIndex        =   24
         Top             =   240
         Width           =   1575
         Begin VB.OptionButton optResidential 
            Caption         =   "Res"
            Height          =   255
            Left            =   660
            TabIndex        =   28
            Top             =   270
            Width           =   615
         End
         Begin VB.OptionButton optClassSysUF 
            Caption         =   "Uni"
            Height          =   210
            Left            =   0
            TabIndex        =   27
            Top             =   300
            Width           =   660
         End
         Begin VB.OptionButton optClassSysMF 
            Caption         =   "Master Format"
            Height          =   210
            Left            =   0
            TabIndex        =   26
            Top             =   0
            Value           =   -1  'True
            Width           =   1395
         End
      End
   End
   Begin VB.ComboBox cmbCity 
      Enabled         =   0   'False
      Height          =   315
      Left            =   7890
      TabIndex        =   19
      Top             =   1560
      Width           =   1815
   End
   Begin VB.ComboBox cmbState 
      Height          =   315
      Left            =   6675
      TabIndex        =   18
      Top             =   1560
      Width           =   765
   End
   Begin VB.TextBox Zip 
      Height          =   285
      Left            =   10050
      TabIndex        =   17
      Top             =   1590
      Width           =   495
   End
   Begin VB.ComboBox cmbCountry 
      Height          =   315
      Left            =   5160
      TabIndex        =   14
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton cmdHistRpt 
      Caption         =   "&Historical"
      Height          =   495
      Left            =   9960
      TabIndex        =   13
      Top             =   6120
      Width           =   1150
   End
   Begin VB.CommandButton cmdCurYrRpt 
      Caption         =   "Current &Year"
      Height          =   495
      Left            =   8760
      TabIndex        =   12
      Top             =   6120
      Width           =   1150
   End
   Begin VB.CommandButton cmdCurPerRpt 
      Caption         =   "&Current Period"
      Height          =   495
      Left            =   7560
      TabIndex        =   11
      Top             =   6120
      Width           =   1150
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Create CSI Format Summary Map Report"
      Height          =   495
      Left            =   4200
      TabIndex        =   10
      Top             =   6120
      Width           =   3315
   End
   Begin VB.ComboBox cmbQuarterID 
      Height          =   315
      Left            =   5160
      TabIndex        =   8
      Top             =   2100
      Width           =   1005
   End
   Begin VB.TextBox ClassificationID 
      Height          =   285
      Left            =   7080
      TabIndex        =   7
      Top             =   2160
      Width           =   855
   End
   Begin ConstructionCostDatabase.DynaTree FormatTree 
      Height          =   2715
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   4789
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   435
      Left            =   9480
      TabIndex        =   0
      Top             =   2040
      Width           =   1150
   End
   Begin VB.CheckBox ckbRowWrap 
      Caption         =   "Row Wrap"
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   1215
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid 
      Height          =   2715
      Left            =   60
      TabIndex        =   6
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
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=30"
      _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=31"
      _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=32"
      _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=34"
      _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=35"
      _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "RecordSelectorStyle:id=37,.parent=2,.namedParent=39"
      _StyleDefs(19)  =   "FilterBarStyle:id=40,.parent=1,.namedParent=42"
      _StyleDefs(20)  =   "Splits(0).Style:id=11,.parent=1"
      _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=20,.parent=4"
      _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=12,.parent=2"
      _StyleDefs(23)  =   "Splits(0).FooterStyle:id=13,.parent=3"
      _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=14,.parent=5"
      _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=16,.parent=6"
      _StyleDefs(26)  =   "Splits(0).EditorStyle:id=15,.parent=7"
      _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=17,.parent=8"
      _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=18,.parent=9"
      _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=19,.parent=10"
      _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=38,.parent=37"
      _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=41,.parent=40"
      _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=24,.parent=11"
      _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12"
      _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15"
      _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=28,.parent=11"
      _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=15"
      _StyleDefs(40)  =   "Named:id=29:Normal"
      _StyleDefs(41)  =   ":id=29,.parent=0"
      _StyleDefs(42)  =   "Named:id=30:Heading"
      _StyleDefs(43)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(44)  =   ":id=30,.wraptext=-1"
      _StyleDefs(45)  =   "Named:id=31:Footing"
      _StyleDefs(46)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(47)  =   "Named:id=32:Selected"
      _StyleDefs(48)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(49)  =   "Named:id=33:Caption"
      _StyleDefs(50)  =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(51)  =   "Named:id=34:HighlightRow"
      _StyleDefs(52)  =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(53)  =   "Named:id=35:EvenRow"
      _StyleDefs(54)  =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
      _StyleDefs(55)  =   "Named:id=36:OddRow"
      _StyleDefs(56)  =   ":id=36,.parent=29"
      _StyleDefs(57)  =   "Named:id=39:RecordSelector"
      _StyleDefs(58)  =   ":id=39,.parent=30"
      _StyleDefs(59)  =   "Named:id=42:FilterBar"
      _StyleDefs(60)  =   ":id=42,.parent=29"
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "State:"
      Height          =   255
      Left            =   6090
      TabIndex        =   23
      Top             =   1620
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "City:"
      Height          =   255
      Left            =   7410
      TabIndex        =   22
      Top             =   1620
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Zip:"
      Height          =   255
      Left            =   9660
      TabIndex        =   21
      Top             =   1620
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "Country:"
      Height          =   255
      Left            =   4440
      TabIndex        =   20
      Top             =   1620
      Width           =   735
   End
   Begin VB.Label lblFromQtr 
      Caption         =   "Quarter:"
      Height          =   255
      Left            =   4440
      TabIndex        =   9
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label lblRowCount 
      Caption         =   "0 rows returned"
      Height          =   255
      Left            =   5340
      TabIndex        =   5
      Top             =   2880
      Width           =   3255
   End
   Begin VB.Line Line1 
      X1              =   4320
      X2              =   4320
      Y1              =   2760
      Y2              =   120
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   11040
      Y1              =   2820
      Y2              =   2820
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Class ID:"
      Height          =   255
      Left            =   6195
      TabIndex        =   4
      Top             =   2160
      Width           =   720
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
      Left            =   4440
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmCCICSIFmtSumRptGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'<modulename> frmCCICsiFmtSumRptGrid.frm</modulename>
'<functionname>General (Main) </functionname>
'
'<summary>
' (CCI) DOLLAR LISTING:
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
'"Geographic Selection" :
'"   NATL AVG (30 city)
'"   PRIMARY CITIES (316 cities)
'"   CCI CITIES (727 cities)
'"   ALL CITIES (731 cities)
'
'(Classification System)
'"   MasterFormat (default)
'"   Res (residential)
'"   Uni (uniformat)
'
'Select Dates:
'"   Quarter
'"   Period  (NOT SUPPORTED!!!)
'
'Index Period:   (NOT SUPPORTED!!!)
'"   Current
'"   Jan 1st
'"   Historical
'
'(By)
'
'1.  Quarter Id              (YYYYQn)
'2.  Country
'3.  City
'4.  State
'5.  Zip
'6.  Equip Id
'
'
'(BUTTONS)
'"   SEARCH              (CmdSearch_Click() )
'
'Search for "index detail" data based upon Selections and filled in boxes:
'
'sp_select_cci_csi_format_sum_map_rpt_RLH
'
'"   Current Period          (BuildReport CURRENT_PERIOD)
'"   Current Year                (BuildReport JAN_1_PERIOD)
'"   Historical              (BuildReport HIST_PERIOD)
'
'"   Create CSI Format Summary Map Report    (ExportData() )
'
'ExecStoredProcSelectedQuarter ("SP_REPORT_PUB_CCI_CSIFORMAT_SUM_MAP_RPT_WITH_FUEL_RLH")
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
'HELPER Class: CCCICSIFmtMap.Cls
' </summary>
'
' <seealso> CCCICSIFmtMap.cls</seealso>
'<seealso> </seealso>
'
' <datastruct>m_rec</datastruct>
'<datastruct>m_objGridMap</datastruct>
'
' <storedprocedurename> sp_select_cci_csi_format_sum_map_rpt_RLH </storedprocedurename>
'<storedprocedurename> sp_report_pub_cci_csiformat_sum_map_rpt_with_fuel_rlh </storedprocedurename>
'
'
' <returns>N/A</returns>
' <exception>Always trap with an accompanying message box</exception>
' <example>
' <code>
'* * *
'SELECTED "MF", QUARTER_ID AND select_type (Nat'l Avg (30 city))  ONLY !!!
'* * *
'
'exec sp_select_cci_csi_format_sum_map_rpt_RLH  @class_id = '', @class_system_id = 'MF', @quarter_id = '2006Q3', @select_type = 2, @zip_3 = '', @loc_id = 0, @state_code = ''
'</code>
' <code>
'    * * *
'            SELECTED "MF", QUARTER_ID, state_code, select_type (Nat'l Avg (30 city))
'* * *
'exec sp_select_cci_csi_format_sum_map_rpt_RLH  @class_id = '', @class_system_id = 'MF', @quarter_id = '2006Q3', @select_type = 2, @zip_3 = '', @loc_id = 0, @state_code = 'CA'
'</code>
'<code>
'* * *  (ANYTOWN)
'SELECTED "MF", QUARTER_ID, LOC_ID=23, STATE_CODE='US' "
' * * *
'exec sp_select_cci_csi_format_sum_map_rpt_RLH  @class_id = '', @class_system_id = 'MF', @quarter_id = '2006Q3', @select_type = 2, @zip_3 = '', @loc_id = 23, @state_code = 'US'
'
'</code>
'<code>* * * SELECTED (UNIFORMAT/UNI),
'quarter_id, state='CA'
'exec sp_select_cci_csi_format_sum_map_rpt_RLH  @class_id = '', @class_system_id = 'U2', @quarter_id = '2006Q3', @select_type = 2, @zip_3 = '', @loc_id = 0, @state_code = 'CA'
'
'NOTE: Not sure if [Classification System]=Uni or Res works…or even applies!?
'</code>
'</example>
'<permission>Public</Permission>
'<dependson>This component depends on the following
'1.  CCCICSIFmtMap.cls
'2.  CGridMap.cls
'3.  CCDdal.CRSMDataAccess (
'4.  Access to the DAL (data access layer dll) opened in MainModule_Main() )
'</dependson>



Dim m_objGridMap As New CCCICSIFmtMap ' Class to handle grid
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

Const CURRENT_PERIOD = 0
Const JAN_1_PERIOD = 1
Const HIST_PERIOD = 2

Private Function FillRptParm(sValue As String) As Variant
    If Len(sValue) = 0 Then
        FillRptParm = """"""
    Else
        FillRptParm = """" + SQLChangeWildcard(sValue) + """"
    End If
End Function

Private Sub RetrievePct(sType As String, dblMatPct As Double, dblInstPct As Double, dblTotPct As Double)
    Dim iSelectType As Integer
    Dim sQtr As String
    Dim sClassSystemID  As String
    Dim rec As ADODB.RecordSet
    Dim strSELECT As String
    Dim blnRet As Boolean
    
    On Error GoTo Err_Handler
    sQtr = cmbQuarterID.Text
    If cmbQuarterID.Text = "" Then
        MsgBox "The start quarter is required."
    End If
    If optClassSysMF Then
        sClassSystemID = "MF"
    ElseIf optClassSysUF Then
        sClassSystemID = "U2"
    ElseIf optResidential Then
        sClassSystemID = "RS"
    End If
    strSELECT = "exec sp_select_cci_dollar_total_pct "
    strSELECT = strSELECT + " @class_id = '" + SQLChangeWildcard(ClassificationID) + "'"
    strSELECT = strSELECT + ", @class_system_id = '" + sClassSystemID + "'"
    strSELECT = strSELECT + ", @quarter_id = '" + sQtr + "'"
    strSELECT = strSELECT + ", @select_type = " + GeographicType(Me)
    strSELECT = strSELECT + ", @zip_3 = '" + SQLChangeWildcard(Zip.Text) + "'"
    If cmbCity.ListIndex = -1 Then
        strSELECT = strSELECT + ", @loc_id = 0"
    Else
        strSELECT = strSELECT + ", @loc_id = " + CStr(cmbCity.ItemData(cmbCity.ListIndex))
    End If
    strSELECT = strSELECT + ", @type = '" + CStr(iSelectType) + "'"
    strSELECT = strSELECT + ", @state_code = '" + cmbState.Text + "'"
    ' Use g_objDAL to perform select
    blnRet = g_objDAL.GetRecordset(vbNullString, strSELECT, rec)
    If blnRet = False Then
        MsgBox "An error occurred while searching."
        lblRowCount.Caption = "0 rows returned."
        Screen.MousePointer = vbNormal
        Exit Sub
    Else
        If IsNull(rec.Fields(0)) Then   'rlh 03/03/2010
        Else
            dblMatPct = rec.Fields(0)
        End If
        If IsNull(rec.Fields(1)) Then   'rlh 03/03/2010
        Else
            dblInstPct = rec.Fields(1)
        End If
        If IsNull(rec.Fields(2)) Then   'rlh 03/03/2010
        Else
            dblTotPct = rec.Fields(2)
        End If
    End If
    rec.Close
    Set rec = Nothing
    Exit Sub
    
Err_Handler:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number
    dblMatPct = 0
    dblInstPct = 0
    dblTotPct = 0
    Exit Sub
    
End Sub

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



Private Sub cmdUpdate_Click()
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
    'RLH 03/03/2010  Can't find stored procedure  (we'll use mine!)
    If DEBUGON Then Stop
    'ExecStoredProcSelectedQuarter "sp_report_pub_cci_masterformat_sum_map_WITH_FUEL" '03/03/2010
    
    ExecStoredProcSelectedQuarter "SP_REPORT_PUB_CCI_CSIFORMAT_SUM_MAP_RPT_WITH_FUEL_RLH" '03/03/2010

End Sub

Private Sub cmdCurPerRpt_Click()
    BuildReport CURRENT_PERIOD
End Sub

Private Sub cmdCurYrRpt_Click()
    BuildReport JAN_1_PERIOD
End Sub

Private Sub cmdHistRpt_Click()
    BuildReport HIST_PERIOD
End Sub

Private Sub Form_Activate()
    Dim ctl As Control
    ShowToolbarIcons True
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
    ShowToolbarIcons False
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim strSELECT As String
    Dim blnReturn As Boolean
    Move START_LEFT, START_TOP, START_WIDTH, START_HEIGHT
    LoadCombos Me, True, True, True

    ' This will never return any rows, just used to create recordset
    ClassificationID.Text = "~"
    cmdSearch_Click
    ClassificationID.Text = ""
    Status ("")
End Sub
Private Sub Form_Initialize()
    Status ("Loading CCI Material Price Maintenance...")
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
            cmdCurPerRpt.Top = Me.Height - 1020
            cmdCurYrRpt.Top = Me.Height - 1020
            cmdHistRpt.Top = Me.Height - 1020
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
        If strID = "U2" Then
            optClassSysUF = True
            ClassificationID.Text = ""
        ElseIf strID = "MF" Then
            optClassSysMF = True
            ClassificationID.Text = ""
        Else
            ClassificationID.Text = strID & "*"
        End If
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
    Dim iSelectType As Integer
    Dim sQtr As String
    Dim sClassSystemID  As String
    
    TDBGrid.Update

    Screen.MousePointer = vbHourglass
    dtmToday = Date
    
    ' Synch tree with text box
    If Not ClassificationID.Text = "" Then
        FormatTree.FocusItem (ClassificationID.Text)
    End If
    
    If Len(ClassificationID.Text) = 0 And Len(cmbCity.Text) = 0 And Len(cmbState.Text) = 0 And Len(Zip.Text) = 0 And Len(cmbQuarterID.Text) = 0 Then
        Screen.MousePointer = vbNormal
        MsgBox "You must enter search criteria before searching."
'        GoTo Exit_Sub
    End If
    
    sQtr = cmbQuarterID.Text
    
    If cmbQuarterID.Text = "" Then
        MsgBox "The start quarter is required."
    End If
    If optClassSysMF Then
        sClassSystemID = "MF"
    ElseIf optClassSysUF Then
        sClassSystemID = "U2"
    ElseIf optResidential Then
        sClassSystemID = "RS"
    End If
    strSELECT = "exec sp_select_cci_csi_format_sum_map_rpt_RLH "  'rlh 03/03/2010
    strSELECT = strSELECT + " @class_id = '" + SQLChangeWildcard(ClassificationID) + "'"
    strSELECT = strSELECT + ", @class_system_id = '" + sClassSystemID + "'"
    strSELECT = strSELECT + ", @quarter_id = '" + sQtr + "'"
    strSELECT = strSELECT + ", @select_type = " + GeographicType(Me)
    strSELECT = strSELECT + ", @zip_3 = '" + SQLChangeWildcard(Zip.Text) + "'"
    If cmbCity.ListIndex = -1 Then
        strSELECT = strSELECT + ", @loc_id = 0"
    Else
        strSELECT = strSELECT + ", @loc_id = " + CStr(cmbCity.ItemData(cmbCity.ListIndex))
    End If
    strSELECT = strSELECT + ", @state_code = '" + cmbState.Text + "'"
    m_rec.Close ' Make sure it is closed
    m_rec.MaxRecords = 5000 ' Set the maximum number to bring back
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
    If m_rec.RecordCount = 5000 And m_rec.State = adStateOpen Then
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

Public Sub PrintReport()
    PreviewReport
End Sub

Public Sub PreviewReport()
    Dim iReportType As Long
    
    iReportType = CURRENT_PERIOD
    BuildReport iReportType
    
End Sub

Public Sub BuildReport(ReportType As Long)
    Dim fPreviewWindow As New frmReportPreview
    Dim sReport As String
    Dim dblMatPct As Double
    Dim dblInstPct As Double
    Dim dblTotPct As Double
    
    Select Case ReportType
    Case CURRENT_PERIOD
        sReport = "Current Period Dollar Report"
        RetrievePct CURRENT_PERIOD, dblMatPct, dblInstPct, dblTotPct
    Case JAN_1_PERIOD
        sReport = "Current Year Dollar Report"
        RetrievePct JAN_1_PERIOD, dblMatPct, dblInstPct, dblTotPct
    Case HIST_PERIOD
        sReport = "History Dollar Report"
        RetrievePct HIST_PERIOD, dblMatPct, dblInstPct, dblTotPct
    End Select
    
    If m_rec.RecordCount >= 1 Then
        fPreviewWindow.ReportName = sReport
        fPreviewWindow.ReportFile = "rptCCICSIFormat.xml"
        fPreviewWindow.OpenEvent = "select_type_sel = """ & GeographicType(Me) & """" & vbCrLf & _
                                    "mat_pct = " & dblMatPct & vbCrLf & _
                                    "inst_pct = " & dblInstPct & vbCrLf & _
                                    "tot_pct = " & dblTotPct
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
    fMainForm.mnuFilePageSetup.Enabled = bShowIcons
    fMainForm.mnuFilePrint.Enabled = bShowIcons
    fMainForm.mnuFilePrintPreview.Enabled = bShowIcons

End Sub

