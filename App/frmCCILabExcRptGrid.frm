VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{5936A75C-3F42-11D6-AF6B-AA0004005F12}#1.3#0"; "MeansCtrl.ocx"
Begin VB.Form frmCCILabExcGrid 
   Caption         =   "CCI Labor Rate Exception Report Display Grid"
   ClientHeight    =   6855
   ClientLeft      =   2265
   ClientTop       =   2835
   ClientWidth     =   11355
   Icon            =   "frmCCILabExcRptGrid.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6855
   ScaleWidth      =   11355
   Begin VB.Frame fraRateCalc 
      Caption         =   "Labor Rates"
      Height          =   975
      Left            =   9360
      TabIndex        =   5
      Top             =   360
      Width           =   1815
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   120
         ScaleHeight     =   615
         ScaleWidth      =   1455
         TabIndex        =   29
         Top             =   240
         Width           =   1455
         Begin VB.OptionButton optNoWorkComp 
            Caption         =   "without W/C"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   300
            Width           =   1195
         End
         Begin VB.OptionButton optWorkComp 
            Caption         =   "with W/C"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   0
            Value           =   -1  'True
            Width           =   1335
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
      Height          =   975
      Left            =   4440
      TabIndex        =   0
      Top             =   360
      Width           =   4695
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   240
         ScaleHeight     =   495
         ScaleWidth      =   4215
         TabIndex        =   28
         Top             =   260
         Width           =   4215
         Begin VB.OptionButton optAllCities 
            Caption         =   "All CCI Cities (731-Cities)"
            Height          =   195
            Left            =   2085
            TabIndex        =   4
            Top             =   300
            Width           =   2070
         End
         Begin VB.OptionButton optPriCity 
            Caption         =   "Primary Cities (316-Cities)"
            Height          =   255
            Left            =   2085
            TabIndex        =   2
            Top             =   0
            Width           =   2055
         End
         Begin VB.OptionButton optCCICities 
            Caption         =   "CCI Cities (727-Cities)"
            Height          =   255
            Left            =   0
            TabIndex        =   3
            Top             =   280
            Width           =   1875
         End
         Begin VB.OptionButton optNatlAvg 
            Caption         =   "Nat'l Avg (30-City)"
            Height          =   255
            Left            =   0
            TabIndex        =   1
            Top             =   0
            Value           =   -1  'True
            Width           =   1695
         End
      End
   End
   Begin VB.ComboBox cmbCountry 
      Height          =   315
      Left            =   5220
      TabIndex        =   9
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox Zip 
      Height          =   285
      Left            =   10740
      TabIndex        =   15
      Top             =   1470
      Width           =   405
   End
   Begin VB.ComboBox cmbState 
      Height          =   315
      Left            =   6795
      TabIndex        =   11
      Top             =   1440
      Width           =   750
   End
   Begin VB.ComboBox cmbCity 
      Enabled         =   0   'False
      Height          =   315
      Left            =   8250
      TabIndex        =   13
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Create Labor Exception Report"
      Height          =   495
      Left            =   6960
      TabIndex        =   24
      Top             =   6240
      Width           =   2715
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "&Report"
      Height          =   495
      Left            =   120
      TabIndex        =   23
      Top             =   6240
      Width           =   1150
   End
   Begin VB.ComboBox cmbTradeID 
      Height          =   315
      Left            =   7335
      TabIndex        =   19
      Top             =   2010
      Width           =   1320
   End
   Begin VB.ComboBox cmbQuarterID 
      Height          =   315
      Left            =   5220
      TabIndex        =   17
      Top             =   2010
      Width           =   1005
   End
   Begin ConstructionCostDatabase.DynaTree FormatTree 
      Height          =   2400
      Left            =   60
      TabIndex        =   25
      Top             =   45
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   4233
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   435
      Left            =   9120
      TabIndex        =   20
      Top             =   1920
      Width           =   1150
   End
   Begin VB.CheckBox ckbRowWrap 
      Caption         =   "Row Wrap"
      Height          =   315
      Left            =   120
      TabIndex        =   21
      Top             =   2580
      Width           =   1215
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid 
      Height          =   3120
      Left            =   75
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2940
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   5503
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
   Begin VB.Label Label6 
      Caption         =   "Country:"
      Height          =   255
      Left            =   4560
      TabIndex        =   8
      Top             =   1500
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Zip:"
      Height          =   255
      Left            =   10320
      TabIndex        =   14
      Top             =   1500
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "City:"
      Height          =   255
      Left            =   7770
      TabIndex        =   12
      Top             =   1500
      Width           =   375
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "State:"
      Height          =   255
      Left            =   6210
      TabIndex        =   10
      Top             =   1500
      Width           =   495
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Quarter:"
      Height          =   255
      Left            =   4545
      TabIndex        =   16
      Top             =   2070
      Width           =   585
   End
   Begin VB.Label lblRowCount 
      Caption         =   "0 rows returned"
      Height          =   255
      Left            =   5355
      TabIndex        =   27
      Top             =   2580
      Width           =   3255
   End
   Begin VB.Line Line1 
      X1              =   4200
      X2              =   4200
      Y1              =   2400
      Y2              =   60
   End
   Begin VB.Line Line2 
      X1              =   135
      X2              =   11070
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Trade ID:"
      Height          =   255
      Left            =   6435
      TabIndex        =   18
      Top             =   2070
      Width           =   855
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
      TabIndex        =   26
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmCCILabExcGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'<modulename> frmCCILabExcGrid.frm</modulename>
'<functionname>General (Main) </functionname>
'
'<summary>
' (CCI) LABOR RATE EXCEPTION REPORT GRID:
'
'This window/form tracks the progression/regression of different selections of labor rates (single or groups of trade_ids), year over year for 4 quarters.
'Labor rate data is displayed by trade_id  based upon a selected "Geographic Selection" as follows:
'"   NATL AVG (30 city)
'"   PRIMARY CITIES (316 cities)
'"   CCI CITIES (727 cities)
'"   ALL CITIES (731 cities)
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
'(Labor Rates)
'"   With W/C        (worker's compensation
'"   Without W/C
'
'
'(BUTTONS)
'"   SEARCH              (CmdSearch_Click() )
'
'Search for labor rate "exception" data based upon selected "Geographic Selection", "W/C" options,  and filled in boxes:
'
'sp_select_published_cci_labor_exc_rpt_rlh
'
'"   Report                  (PreviewReport() )
'
'rptCCIExceptionReport.xml                 XML TEMPLATE/ComponentOne
'
'"   Create Labor Exception Table
'
'ExecStoredProcSelectedQuarter " SP_REPORT_PUB_CCI_LABOR_RATE_GRID " (Quarter Id)
'
'NOTE: "Anytown" = 30 city average  and is displayed by selecting:
'"   Country = USA
'"   State = US
'"   City = Anytown
'
'Key Subs / Functions:
'"   CmdSearch_Click()
'Prepares parameters to be passed with the stored procedure to retrieve needed "All Cities" or "Anytown" data
'"   ExecStoredProcSelectedQuarter()
'
'HELPER Class: CCCILabExcMap.Cls
' </summary>
'
' <seealso> CCCILabExcMap.cls </seealso>
'<seealso> </seealso>
'
' <datastruct>m_rec</datastruct>
'<datastruct>m_objGridMap</datastruct>
'
' <storedprocedurename> sp_select_published_cci_labor_exc_rpt_rlh </storedprocedurename>
'<storedprocedurename> SP_REPORT_PUB_CCI_LABOR_RATE_GRID </storedprocedurename>
'
'
' <returns>N/A</returns>
' <exception>Always trap with an accompanying message box</exception>
' <example>
' <code>
'* * * NOTE: You MUST "Create Labor Exception Table" before any data reporting can happen !!!
'
'* * * SELECT trade id from TREE: CEFI   (no other  selections /specifications)
'
'exec sp_select_published_cci_labor_exc_rpt_rlh   @trade_skey = 31, @quarter_id = '2006Q3', @zip_3 = '', @loc_id = 0, @state_code = '', @select_type = 2, @country_code = '%', @workers_comp = '0'
'</code>
' <code>
'* * * NOTE: You MUST "Create Labor Exception Report" before any data reporting can happen !!!
'
'* * * SPECIFY TRADE ID (trade_skey), CITY & STATE (i.e. loc_id) for trade id, CEFI, only
'
'exec sp_select_published_cci_labor_exc_rpt_rlh   @trade_skey = 31, @quarter_id = '2006Q3', @zip_3 = '', @loc_id = 0, @state_code = 'CA', @select_type = 2, @country_code = '%', @workers_comp = '0' </code>
'<code>
'* * * NOTE: You MUST "Create Labor Exception Table" before any data reporting can happen  !!!
'
'* * * ANYTOWN (loc_id=23)  (across ALL cci equipment)
'
'exec sp_select_published_cci_labor_exc_rpt_rlh  @trade_skey = 0, @quarter_id = '2006Q3', @zip_3 = '', @loc_id = 23, @state_code = 'US', @select_type = 2, @country_code = '%', @workers_comp = '0'
'</code>
'<code>
'* * * ANYTOWN (loc_id = 23), Workers Comp, for all CCI trade_ids…
'
'exec sp_select_published_cci_labor_exc_rpt_rlh  @trade_skey = 0, @quarter_id = '2006Q3', @zip_3 = '', @loc_id = 23, @state_code = 'US', @select_type = 2, @country_code = '%', @workers_comp = '1'
'</code>
'</example>
'<permission>Public</Permission>
'<dependson>This component depends on the following
'1.  CCCILabExcMap.cls
'2.  CGridMap.cls
'3.  CCDdal.CRSMDataAccess (
'4.  Access to the DAL (data access layer dll) opened in MainModule_Main() )
'</dependson>



Dim m_objGridMap As New CCCILabExcMap ' Class to handle grid
Dim m_blnFirstSearch As Boolean
Dim m_rec As New ADODB.RecordSet ' Recordset to hold query results
Dim m_blnDoubleClick As Boolean
Dim m_blnWereErrors As Boolean ' True if the Update had errors, used in QueryUnload
Dim m_strCurrentFormControl As String
Dim m_CurrentQtr As String
Dim m_State As String
Dim m_blnWorkersComp As Boolean ' True if Data includes the Worker's Comp calculation

Private Sub ShowToolbarIcons(bShowIcons As Boolean)

    fMainForm.tbToolBar.Buttons.Item(tbrPRINT).Enabled = bShowIcons
    fMainForm.tbToolBar.Buttons.Item(tbrPRINT).Visible = bShowIcons
    fMainForm.tbToolBar.Buttons.Item(tbrPREVIEW).Enabled = bShowIcons
    fMainForm.tbToolBar.Buttons.Item(tbrPREVIEW).Visible = bShowIcons
    fMainForm.mnuFilePageSetup.Enabled = bShowIcons
    fMainForm.mnuFilePrint.Enabled = bShowIcons
    fMainForm.mnuFilePrintPreview.Enabled = bShowIcons

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

Private Sub cmdLaborRates_Click()
    If IsNumeric(TDBGrid.Bookmark) = False Then
        MsgBox "You must select a row."
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    ' Navigate to single-record view
    Dim frm As frmLaborRateGrid
    Dim rec As ADODB.RecordSet
    Set frm = New frmLaborRateGrid
    frm.JumpIn Trim(TDBGrid.Columns("CCI Trade ID").CellText(TDBGrid.Bookmark)), TDBGrid.Columns("State").CellText(TDBGrid.Bookmark), TDBGrid.Columns("City").CellText(TDBGrid.Bookmark), ""
    Screen.MousePointer = vbNormal

End Sub

Private Sub cmdPublishLaborRates_Click()
    ExecStoredProcSelectedQuarter "SP_UPDATE_PUBLISHED_CCI_LABOR_RATE"
End Sub

Private Sub cmdCreate_Click()

    ExecStoredProcSelectedQuarter "SP_REPORT_PUB_CCI_LABOR_RATE_GRID"
    MsgBox "The report tables have been updated.", vbInformation + vbOKOnly

End Sub

Private Sub cmdPreview_Click()
    
    PreviewReport
    
End Sub

Private Function FillRptParm(sValue As String) As Variant
    If Len(sValue) = 0 Then
        FillRptParm = """"""
    Else
        FillRptParm = """" + SQLChangeWildcard(sValue) + """"
    End If
End Function

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
    ShowToolbarIcons False
    m_strCurrentFormControl = Me.ActiveControl.Name
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim strSELECT As String
    Dim blnReturn As Boolean
    
    Move START_LEFT, START_TOP, START_WIDTH, START_HEIGHT
    LoadCombos Me, True, True, True, True

    m_blnWorkersComp = True
    optWorkComp = m_blnWorkersComp
       
    ' This will never return any rows, just used to create recordset
    Zip = "~"
    cmdSearch_Click
    Zip = ""
    Status ("")
    
End Sub

Private Sub Form_Initialize()
    Status ("Loading CCI Labor Exception Report Maintenance...")
    Screen.MousePointer = vbHourglass
    m_blnFirstSearch = True
    FormatTree.InitData g_cnShared, "CCI_LABOR"
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
        If Me.Width >= 11475 Then
            TDBGrid.Width = Me.Width - 255
            Line2.X2 = Me.Width - 210
        Else
            Me.Width = 11475
        End If
        If Me.Height >= 7260 Then
            cmdCreate.Top = Me.Height - 1020
            cmdPreview.Top = cmdCreate.Top
            TDBGrid.Height = cmdCreate.Top - TDBGrid.Top - 240
        Else
            Me.Height = 7260
        End If
    Else
        ShowMinimizedForms
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ShowToolbarIcons False
    HideGridSort
End Sub

' Leaf in MasterFormat tree selected.
Private Sub FormatTree_NodeSelected(ByVal strID As String)
Dim rs As New ADODB.RecordSet
Dim strSELECT As String
Dim blnReturn As Boolean
Dim i As Integer

On Error Resume Next
    If m_blnFirstSearch = True Then
        m_blnFirstSearch = False
    Else
        If Len(strID) = 0 Or strID = "op" Then
            cmbTradeID.ListIndex = -1
        Else
            For i = 0 To cmbTradeID.listcount - 1
                If strID = cmbTradeID.List(i) Then
                    cmbTradeID.Text = strID
                    cmbTradeID.ListIndex = i
                    Exit For
                End If
            Next i
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
    Dim strStartMatSrch As String
    Dim iSelectType As Integer
    Dim iWorkersCompType As Integer
    
    TDBGrid.Update
    
   
    If m_objGridMap.IsPendingChange = True Then
        Dim Button
        Button = MsgBox("Do you want to save your changes?", vbYesNoCancel)
        If Button = vbYes Then
'            cmdUpdate_Click
            ' If there were errors, cancel the search
            If m_blnWereErrors Then
                Exit Sub
            End If
        Else
            If Button = vbNo Then
                TDBGrid.DataChanged = False
            ElseIf Button = vbCancel Then
                ' Cancel the search
                Exit Sub
            End If
        End If
    End If
    Screen.MousePointer = vbHourglass
    dtmToday = Date
    
    ' Synch tree with text box
    If Not cmbTradeID.Text = "" Then
        FormatTree.FocusItem (cmbTradeID.Text)
    End If
    
     If Len(cmbTradeID.Text) = 0 And Len(cmbCity.Text) = 0 And Len(cmbState.Text) = 0 And Len(Zip.Text) = 0 And Len(cmbQuarterID.Text) = 0 Then
        Screen.MousePointer = vbNormal
        MsgBox "You must enter search criteria before searching."
'        GoTo Exit_Sub
    End If
    
    If cmbQuarterID.Text = "" Then
        MsgBox "The quarter is required."
    End If
    
    ' 9/12/2005 RTD - SET THE WORKER'S COMP CALCULATION FLAG
    m_blnWorkersComp = optWorkComp.Value
    If optWorkComp.Value Then
        iWorkersCompType = 1
    End If
    
    strSELECT = "exec sp_select_published_cci_labor_exc_rpt_rlh " 'rlh 02/27/2010
    If cmbTradeID.ListIndex = -1 Then
        strSELECT = strSELECT + " @trade_skey = 0"
    Else
        strSELECT = strSELECT + "  @trade_skey = " + CStr(cmbTradeID.ItemData(cmbTradeID.ListIndex))
    End If
    strSELECT = strSELECT + ", @quarter_id = '" + cmbQuarterID.Text + "'"
    strSELECT = strSELECT + ", @zip_3 = '" + SQLChangeWildcard(Zip.Text) + "'"
    If cmbCity.ListIndex = -1 Then
        strSELECT = strSELECT + ", @loc_id = 0"
    Else
        strSELECT = strSELECT + ", @loc_id = " + CStr(cmbCity.ItemData(cmbCity.ListIndex))
    End If
    strSELECT = strSELECT + ", @state_code = '" + cmbState.Text + "'"
    strSELECT = strSELECT + ", @select_type = " + GeographicType(Me)
    strSELECT = strSELECT + ", @country_code = '" + FillWildCard(cmbCountry.Text) + "'"
    ' 9/12/2005 RTD - PASS @workers_comp VARIABLE TO STORED PROC
    '                 (VALUE IS OPTIONAL, DEFAULT = 1 == INCLUDE WORKER'S COMP CALCULATION)
    strSELECT = strSELECT + ", @workers_comp = '" & iWorkersCompType & "'"
    m_rec.Close ' Make sure it is closed
    'm_rec.MaxRecords = MAX_RECORDS ' Set the maximum number to bring back
    m_rec.MaxRecords = 0 ' 7/18/2005 RTD ALLOW ALL RECORDS FOR REPORT PREVIEW
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
    Screen.MousePointer = vbNormal

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    ' Check if there are pending changes
    If m_objGridMap.IsPendingChange = True Then
        Dim Button
        Button = MsgBox("Do you want to save your changes?", vbYesNoCancel)
        If Button = vbYes Then
            'cmdUpdate_Click
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

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

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
    Dim fPreviewWindow As New frmReportPreview
    
    If m_rec.RecordCount >= 1 Then
        fPreviewWindow.ReportName = "CCI Labor Rate Exception Report"
        fPreviewWindow.ReportFile = "rptCCIExceptionReport.xml"
        fPreviewWindow.OpenEvent = "select_type_sel = """ & CInt(GeographicType(Me)) & """" & _
                                    vbCrLf & "workers_comp = " & m_blnWorkersComp
        fPreviewWindow.RecordSet = m_rec
        fPreviewWindow.RenderReport
        fPreviewWindow.Show
    Else
        MsgBox "Please choose or search for a CCI trade.", vbInformation + vbOKOnly, "Warning"
    End If
    
End Sub

