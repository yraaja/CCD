VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{5936A75C-3F42-11D6-AF6B-AA0004005F12}#1.3#0"; "MeansCtrl.ocx"
Begin VB.Form frmCCIMatEquRptGrid 
   Caption         =   "CCI Material/Equipment Exceptions Report Grid"
   ClientHeight    =   6765
   ClientLeft      =   2265
   ClientTop       =   2835
   ClientWidth     =   11475
   Icon            =   "frmCCIMatEquRptGrid.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6765
   ScaleWidth      =   11475
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
      Left            =   5040
      TabIndex        =   17
      Top             =   720
      Width           =   5175
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   520
         Left            =   240
         ScaleHeight     =   525
         ScaleWidth      =   4695
         TabIndex        =   25
         Top             =   220
         Width           =   4695
         Begin VB.OptionButton optAllCities 
            Caption         =   "All CCI Cities (731-Cities)"
            Height          =   195
            Left            =   2085
            TabIndex        =   29
            Top             =   300
            Width           =   2070
         End
         Begin VB.OptionButton optPriCity 
            Caption         =   "Primary Cities (316-Cities)"
            Height          =   255
            Left            =   2085
            TabIndex        =   28
            Top             =   0
            Width           =   2055
         End
         Begin VB.OptionButton optCCICities 
            Caption         =   "CCI Cities (727-Cities)"
            Height          =   255
            Left            =   0
            TabIndex        =   27
            Top             =   300
            Width           =   1875
         End
         Begin VB.OptionButton optNatlAvg 
            Caption         =   "Nat'l Avg (30-City)"
            Height          =   255
            Left            =   0
            TabIndex        =   26
            Top             =   0
            Value           =   -1  'True
            Width           =   1695
         End
      End
   End
   Begin VB.ComboBox cmbCity 
      Enabled         =   0   'False
      Height          =   315
      Left            =   8370
      TabIndex        =   20
      Top             =   1710
      Width           =   1815
   End
   Begin VB.ComboBox cmbState 
      Height          =   315
      Left            =   7155
      TabIndex        =   19
      Top             =   1710
      Width           =   765
   End
   Begin VB.TextBox Zip 
      Height          =   285
      Left            =   10560
      TabIndex        =   18
      Top             =   1710
      Width           =   495
   End
   Begin VB.ComboBox cmbCountry 
      Height          =   315
      Left            =   5640
      TabIndex        =   16
      Top             =   1710
      Width           =   855
   End
   Begin VB.TextBox EquClassificationID 
      Height          =   285
      Left            =   9000
      TabIndex        =   15
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CheckBox optRcdsEquip 
      Caption         =   "Equipment"
      Height          =   255
      Left            =   6720
      TabIndex        =   13
      Top             =   390
      Width           =   1215
   End
   Begin VB.CheckBox optRcdsMatl 
      Caption         =   "Material"
      Height          =   255
      Left            =   5280
      TabIndex        =   12
      Top             =   390
      Width           =   975
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "&Report"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   6120
      Width           =   1150
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Create Mat/Equ Exception Table"
      Height          =   495
      Left            =   8400
      TabIndex        =   10
      Top             =   6120
      Width           =   2715
   End
   Begin VB.ComboBox cmbQuarterID 
      Height          =   315
      Left            =   5640
      TabIndex        =   8
      Top             =   2160
      Width           =   1005
   End
   Begin VB.TextBox MatClassificationID 
      Height          =   285
      Left            =   7320
      TabIndex        =   7
      Top             =   2160
      Width           =   855
   End
   Begin ConstructionCostDatabase.DynaTree FormatTree 
      Height          =   2715
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   4789
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   435
      Left            =   10200
      TabIndex        =   0
      Top             =   2160
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
      Left            =   120
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
      ColumnFooters   =   -1  'True
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
      _StyleDefs(51)  =   ":id=31,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(52)  =   ":id=31,.fontname=MS Sans Serif"
      _StyleDefs(53)  =   "Named:id=32:Selected"
      _StyleDefs(54)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(55)  =   "Named:id=33:Caption"
      _StyleDefs(56)  =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(57)  =   "Named:id=34:HighlightRow"
      _StyleDefs(58)  =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(59)  =   "Named:id=35:EvenRow"
      _StyleDefs(60)  =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
      _StyleDefs(61)  =   "Named:id=36:OddRow"
      _StyleDefs(62)  =   ":id=36,.parent=29"
      _StyleDefs(63)  =   "Named:id=39:RecordSelector"
      _StyleDefs(64)  =   ":id=39,.parent=30"
      _StyleDefs(65)  =   "Named:id=42:FilterBar"
      _StyleDefs(66)  =   ":id=42,.parent=29"
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "State:"
      Height          =   255
      Left            =   6570
      TabIndex        =   24
      Top             =   1710
      Width           =   495
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "City:"
      Height          =   255
      Left            =   7890
      TabIndex        =   23
      Top             =   1710
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Zip:"
      Height          =   255
      Left            =   10140
      TabIndex        =   22
      Top             =   1710
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Country:"
      Height          =   255
      Left            =   4920
      TabIndex        =   21
      Top             =   1710
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "Equ ID:"
      Height          =   255
      Left            =   8400
      TabIndex        =   14
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblFromQtr 
      Caption         =   "Quarter:"
      Height          =   255
      Left            =   4920
      TabIndex        =   9
      Top             =   2160
      Width           =   615
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
      X1              =   4800
      X2              =   4800
      Y1              =   2730
      Y2              =   90
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   11040
      Y1              =   2820
      Y2              =   2820
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Mat ID:"
      Height          =   255
      Left            =   6675
      TabIndex        =   4
      Top             =   2160
      Width           =   600
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
      Left            =   4920
      TabIndex        =   3
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmCCIMatEquRptGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' <modulename> frmCCIMatEquRptGrid.frm</modulename>
' <functionname>General (Main) </functionname>
'
' <summary>
' (CCI) MATERIAL/EQUIPMENT EXCEPTIONS REPORT GRID:
'
'Display Equipment rental rates based upon a selected "Geographic Selection". As follows:
'"   NATL AVG (30 city)
'"   PRIMARY CITIES (316 cities)
'"   CCI CITIES (727 cities)
'"   ALL CITIES (731 cities)
'
'(By)
'
'1.  Quarter Id                  (YYYYQn)
'2.  Country
'3.  City
'4.  State
'5.  Zip
'6.  Equip Id
'
'(BUTTONS)
'"   SEARCH              (CmdSearch_Click() )
'       Search for equipment rental rate data based upon selected "Geographic Selection" and filled in boxes
'"   Report              (PreviewReport)
'    rptCCIExceptionReport.xml   (XML TEMPLATE)
'"   Create Mat/Equ Exception Table
'       ExecStoredProcSelectedQuarter "SP_REPORT_PUB_CCI_MATERIAL_EQUIPMENT_WITH_FUEL_RLH" (Quarter Id)
'
'NOTE: "Anytown" = 30 city average  and is displayed by selecting:
'Country = USA
'State = US
'City = Anytown
'
'NOTE: "Rolling Quarters" on the data grid columns
'Given the requested/selected "Quarter Id", the previous (4) quarters of data will be displayed accordingly
'
'Key Subs / Functions:
'"   CmdSearch_Click()
'Prepares parameters to be passed with the stored procedure to retrieve needed "All Cities" or "Anytown" data
'"   ExecStoredProcSelectedQuarter()
'
'HELPER Class: CCCIMatEqMap.Cls
' </summary>
'
' <seealso>CCCIEqpRtMap.cls</seealso>
'<seealso> </seealso>
'
' <datastruct>m_rec</datastruct>
'<datastruct>m_objGridMap</datastruct>
'
' <storedprocedurenamesp_select_cci_mat_equ_rpt_RLH</storedprocedurename>
'<storedprocedurename> SP_REPORT_PUB_CCI_MATERIAL_EQUIPMENT_WITH_FUEL_RLH</storedprocedurename>
'
'
' <returns>N/A</returns>
' <exception>Always trap with an accompanying message box</exception>
' <example>
' <code> * * * NOTE: You MUST "Create Mat/Equ Exception Table" before any data reporting can happen !!!
'* * * SELECT material id from TREE: M10MA   (no other selections/specifications)
'
'exec sp_select_cci_mat_equ_rpt_RLH  @cci_mat_id = 'M10MA%', @cci_equip_id = '%', @quarter_id = '2010Q3', @zip_3 = '', @loc_id = 0, @state_code = '', @select_type = 2, @type_code = 'M', @country_code = '%'
'</code>
' <code>
'* * * NOTE: You MUST "Create Mat/Equ Exception Table" before any data reporting can happen !!!
'* * * SPECIFY MATERIAL ID, CITY & STATE (i.e. loc_id) for equipment id, EAC60, only
'exec sp_select_cci_mat_equ_rpt_RLH  @cci_mat_id = 'M10MA%', @cci_equip_id = '%', @quarter_id = '2010Q3', @zip_3 = '', @loc_id = 59, @state_code = 'CA', @select_type = 2, @type_code = 'M', @country_code = '%'
'</code>
'<code>
'* * * NOTE: You MUST "Create Mat/Equ Exception Table" before any data reporting can happen  !!!
'* * * ANYTOWN (loc_id=23)  (across ALL cci equipment)
'exec sp_select_cci_equipment_rate_rlh   @cci_equip_id = '', @qtr_id = '2006Q4', @zip = '', @loc_id = 23, @state_code = 'US', @select_type = 2
'</code>
'<code>
'* * * CLONE  (copy a specified quarter's worth of data for equipment rates into
'the previous quarter in PUBLISHED_CCI_EQUIPMENT_RATE
'* * *
'exec sp_clone_pub_cci_equipment_rate '2006Q4'
'</code>
'</example>
'<permission>Public</Permission>
'<dependson>This component depends on the following
'1.  CCCIMatEqMap.cls
'2.  CGridMap.cls
'3.  CCDdal.CRSMDataAccess (
'4.  Access to the DAL (data access layer dll) opened in MainModule_Main() )
'</dependson>


Dim m_objGridMap As New CCCIMatEqMap ' Class to handle grid
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

Private Function FindRcdType() As String
    
    If optRcdsMatl.Value = 1 And optRcdsEquip.Value = 1 Then
        FindRcdType = "A"
    ElseIf optRcdsMatl.Value = 1 Then
        FindRcdType = "M"
    ElseIf optRcdsEquip.Value = 1 Then
        FindRcdType = "E"
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
    'rlh 07/15/2009  changed to "with_fuel" as per ksr
    If DEBUGON Then Stop
    ExecStoredProcSelectedQuarter "SP_REPORT_PUB_CCI_MATERIAL_EQUIPMENT_WITH_FUEL_RLH"
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
        'TDBGrid.ReBind
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
    Dim strSelect As String
    Dim blnReturn As Boolean
    Move START_LEFT, START_TOP, START_WIDTH, START_HEIGHT
    LoadCombos Me, True, True, True
    
    TDBGrid.FooterFont.Bold = True
    
    ' This will never return any rows, just used to create recordset
    MatClassificationID.Text = "~"
    cmdSearch_Click
    MatClassificationID.Text = ""
    Status ("")
End Sub

Private Sub Form_Initialize()
    ' 10/04/2005 RTD - CORRECTED INCORRECT STATUS MESSAGE
    Status ("Loading CCI Material/Equipment Report...")
    Screen.MousePointer = vbHourglass
    m_blnFirstSearch = True
    FormatTree.InitData g_cnShared, "CCI_MAT_EQUIP"
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
Dim strSelect As String
Dim blnReturn As Boolean

On Error Resume Next
    If m_blnFirstSearch = True Then
        m_blnFirstSearch = False
    Else
        If strID = "MA" Then
            optRcdsMatl.Value = 1
            optRcdsEquip.Value = 0
            EquClassificationID = ""
            MatClassificationID = ""
        ElseIf strID = "EQ" Then
            optRcdsMatl.Value = 0
            optRcdsEquip.Value = 1
            EquClassificationID = ""
            MatClassificationID = ""
        ElseIf Mid(strID, 1, 1) = "E" Then
            MatClassificationID = ""
            EquClassificationID = strID
            optRcdsEquip.Value = 1
            optRcdsMatl.Value = 0
        ElseIf Mid(strID, 1, 1) = "M" Then
            optRcdsMatl.Value = 1
            optRcdsEquip.Value = 0
            EquClassificationID = ""
            MatClassificationID = strID
        ElseIf strID = "op" Then    'All
            optRcdsMatl.Value = 1
            optRcdsEquip.Value = 1
            EquClassificationID = ""
            MatClassificationID = ""
        End If
        
        ' Kick-off search
        cmdSearch_Click
    End If
End Sub

Private Sub cmdSearch_Click()
    On Error Resume Next
    Dim blnRet As Boolean
    Dim strSelect As String
    Dim dtmToday As Date
    Dim dtmStart As Date
    Dim strStartMatSrch As String
    Dim sQtr As String
    Dim sMktCode As String
    Dim sClassSystemID  As String
    Dim sType As String
    
    TDBGrid.Update

    Screen.MousePointer = vbHourglass
    dtmToday = Date
    
    ' Synch tree with text box
    If Not MatClassificationID.Text = "" Then
        FormatTree.FocusItem (MatClassificationID.Text)
    End If
    
     If Len(MatClassificationID.Text) = 0 And Len(EquClassificationID.Text) = 0 And Len(cmbCity.Text) = 0 And Len(cmbState.Text) = 0 And Len(Zip.Text) = 0 And Len(cmbQuarterID.Text) = 0 Then
        Screen.MousePointer = vbNormal
        MsgBox "You must enter search criteria before searching."
'        GoTo Exit_Sub
    End If
    
    sQtr = cmbQuarterID.Text
    
    If cmbQuarterID.Text = "" Then
        MsgBox "The start quarter is required."
    End If
    
    strSelect = "exec sp_select_cci_mat_equ_rpt_RLH "
    strSelect = strSelect + " @cci_mat_id = '" + FillWildCard(SQLChangeWildcard(MatClassificationID.Text)) + "'"
    strSelect = strSelect + ", @cci_equip_id = '" + FillWildCard(SQLChangeWildcard(EquClassificationID.Text)) + "'"
    strSelect = strSelect + ", @quarter_id = '" + sQtr + "'"
    strSelect = strSelect + ", @zip_3 = '" + SQLChangeWildcard(Zip.Text) + "'"
    If cmbCity.ListIndex = -1 Then
        strSelect = strSelect + ", @loc_id = 0"
    Else
        strSelect = strSelect + ", @loc_id = " + CStr(cmbCity.ItemData(cmbCity.ListIndex))
    End If
    strSelect = strSelect + ", @state_code = '" + cmbState.Text + "'"
    strSelect = strSelect + ", @select_type = " + GeographicType(Me)
    strSelect = strSelect + ", @type_code = '" + FindRcdType() + "'"
    strSelect = strSelect + ", @country_code = '" + FillWildCard(cmbCountry.Text) + "'"
    
    m_rec.Close ' Make sure it is closed
    ' CHANGED 6/16/2005 RTD FOR VERSION 7.4.0 CR#1526
    'm_rec.MaxRecords = MAX_RECORDS ' Set the maximum number to bring back
    m_rec.MaxRecords = 0 ' Return ALL records
    dtmStart = Now
    ' Use g_objDAL to perform select
    If DEBUGON Then Stop
    
    blnRet = g_objDAL.GetRecordset(vbNullString, strSelect, m_rec)
    If blnRet = False Then
        MsgBox "An error occurred while searching." & vbCrLf & g_objDAL.LastErrorDescription
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
    
    UpdateFooter
    
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

Private Sub cmdPreview_Click()
    PreviewReport
End Sub

Public Sub PrintReport()
    PreviewReport
End Sub

Public Sub PreviewReport()
    Dim fPreviewWindow As New frmReportPreview
    
    If m_rec.RecordCount >= 1 Then
        fPreviewWindow.ReportName = "CCI Exception Report"
        fPreviewWindow.ReportFile = "rptCCIExceptionReport.xml"
        fPreviewWindow.OpenEvent = "select_type_sel = """ & CInt(GeographicType(Me)) & """"
        fPreviewWindow.RecordSet = m_rec
        fPreviewWindow.RenderReport
        fPreviewWindow.Show
    Else
        MsgBox "Please choose or search for a CCI.", vbInformation + vbOKOnly, "Warning"
    End If
    
End Sub

Private Sub UpdateFooter()
    Dim dAvg(8) As Double
    Dim i As Long
    
    TDBGrid.Columns("City").FooterText = "AVERAGE"
    If DEBUGON Then Stop
    For i = 1 To 8
        dAvg(i) = 0
    Next
    m_rec.MoveFirst
    Do While Not m_rec.EOF
        dAvg(1) = dAvg(1) + m_rec("Q1")
        dAvg(2) = dAvg(2) + m_rec("Q2")
        dAvg(3) = dAvg(3) + m_rec("Q3")
        dAvg(4) = dAvg(4) + m_rec("Q4")
        dAvg(5) = dAvg(5) + m_rec("Q1_pct")
        dAvg(6) = dAvg(6) + m_rec("Q2_pct")
        dAvg(7) = dAvg(7) + m_rec("Q3_pct")
        dAvg(8) = dAvg(8) + m_rec("Q4_pct")
        m_rec.MoveNext
    Loop
    For i = 1 To 8
        dAvg(i) = dAvg(i) / m_rec.RecordCount
    Next
    
    TDBGrid.Columns("Q1").FooterText = Format(dAvg(1), "Standard")
    TDBGrid.Columns("Q2").FooterText = Format(dAvg(2), "Standard")
    TDBGrid.Columns("Q3").FooterText = Format(dAvg(3), "Standard")
    TDBGrid.Columns("Q4").FooterText = Format(dAvg(4), "Standard")
    TDBGrid.Columns("Q1_pct").FooterText = Format(dAvg(5), "0.000")
    TDBGrid.Columns("Q2_pct").FooterText = Format(dAvg(6), "0.000")
    TDBGrid.Columns("Q3_pct").FooterText = Format(dAvg(7), "0.000")
    TDBGrid.Columns("Q4_pct").FooterText = Format(dAvg(8), "0.000")

End Sub

Public Sub ExportData()
    
    If m_rec.RecordCount > 0 Then
        Dim fExport As New frmExport
    
        fExport.SetRow TDBGrid, m_rec
        fExport.Title = "CCI Material-Equipment Report"
        fExport.Show
    Else
        MsgBox "Please choose or search for a CCI material or equipment.", vbInformation + vbOKOnly
    End If
    
End Sub
Private Sub SetEstimateIndicators(blnNewValue As Boolean)
' 10/10/2005 RTD - UPDATED TO CORRECT BUG REPORTED BY G. MEDEIROS
    If m_rec.RecordCount > 0 Then
        If TDBGrid.SelBookmarks.Count > 0 Then
            ' If rows selected, update only the selected rows
            Dim vntBookmark As Variant
            For Each vntBookmark In TDBGrid.SelBookmarks
                TDBGrid.Bookmark = vntBookmark
                TDBGrid.Columns("estimated_ind") = blnNewValue
            Next
            TDBGrid.Refresh
        Else
            ' Update All Rows
            TDBGrid.MoveFirst
            Do While Not TDBGrid.EOF
                TDBGrid.Columns("estimated_ind") = blnNewValue
                TDBGrid.MoveNext
            Loop
            TDBGrid.Refresh
        End If
    End If
    
End Sub
