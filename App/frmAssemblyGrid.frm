VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{5936A75C-3F42-11D6-AF6B-AA0004005F12}#1.3#0"; "MeansCtrl.ocx"
Begin VB.Form frmAssemblyGrid 
   Caption         =   "Assembly Maintenance Grid"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11130
   Icon            =   "frmAssemblyGrid.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6855
   ScaleWidth      =   11130
   Begin VB.ComboBox cboMasterFormat 
      Height          =   315
      Left            =   7800
      TabIndex        =   30
      Text            =   " "
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Frame fraAssemblyType 
      Caption         =   "Assembly Type"
      Height          =   615
      Left            =   8040
      TabIndex        =   11
      Top             =   120
      Width           =   2895
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   2535
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   240
         Width           =   2535
         Begin VB.OptionButton optAssemblyType 
            Caption         =   "Co&mmercial"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   0
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optAssemblyType 
            Caption         =   "R&esidential"
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   0
            Width           =   1215
         End
      End
   End
   Begin VB.TextBox EndAssemblyID 
      Height          =   315
      Left            =   9600
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox AltAssemblyId 
      Height          =   315
      Left            =   8640
      TabIndex        =   5
      Top             =   1515
      Width           =   1335
   End
   Begin VB.CommandButton cmdClone 
      Caption         =   "C&lone"
      Height          =   495
      Left            =   10260
      TabIndex        =   26
      Top             =   6240
      Width           =   795
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   495
      Left            =   9360
      TabIndex        =   25
      Top             =   6240
      Width           =   795
   End
   Begin VB.Frame Frame1 
      Caption         =   "Go To"
      Height          =   855
      Left            =   120
      TabIndex        =   16
      Top             =   6000
      Width           =   7095
      Begin VB.CommandButton cmdHistory 
         Caption         =   "&History"
         Height          =   495
         Left            =   6000
         TabIndex        =   22
         Top             =   240
         Width           =   915
      End
      Begin VB.CommandButton cmdCostworks 
         Caption         =   "Cost&Works"
         Enabled         =   0   'False
         Height          =   495
         Left            =   3750
         TabIndex        =   20
         Top             =   240
         Width           =   1035
      End
      Begin VB.CommandButton cmdAssemblySqFtUsage 
         Caption         =   "Assembly U&sage"
         Height          =   495
         Left            =   4920
         TabIndex        =   21
         Top             =   240
         Width           =   915
      End
      Begin VB.CommandButton cmdAssembly 
         Caption         =   "&Assembly"
         Height          =   495
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   1035
      End
      Begin VB.CommandButton cmdAssemblyUCUsage 
         Caption         =   "Unit &Cost Usage"
         Height          =   495
         Left            =   2580
         TabIndex        =   19
         Top             =   240
         Width           =   1035
      End
      Begin VB.CommandButton cmdAssemblyBkDtl 
         Caption         =   "&Book Detail"
         Height          =   495
         Left            =   1440
         TabIndex        =   18
         Top             =   240
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   495
      Left            =   8460
      TabIndex        =   24
      Top             =   6240
      Width           =   795
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   495
      Left            =   7440
      TabIndex        =   23
      Top             =   6240
      Width           =   915
   End
   Begin VB.CheckBox ckbRowWrap 
      Caption         =   "Row Wrap"
      Height          =   315
      Left            =   120
      TabIndex        =   13
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   375
      Left            =   9360
      TabIndex        =   8
      Top             =   2355
      Width           =   1150
   End
   Begin VB.TextBox TechDesc 
      Height          =   315
      Left            =   8640
      TabIndex        =   7
      Top             =   1920
      Width           =   2115
   End
   Begin VB.TextBox StartAssemblyID 
      Height          =   315
      Left            =   8040
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
   Begin ConstructionCostDatabase.DynaTree FormatTree 
      Height          =   2775
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   4895
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid 
      Height          =   2715
      Left            =   60
      TabIndex        =   15
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
      Splits(0).AllowRowSelect=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
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
   Begin VB.Label lblMstrFmt 
      Alignment       =   1  'Right Justify
      Caption         =   "Mstr Fmt:"
      Height          =   255
      Left            =   6840
      TabIndex        =   31
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "To"
      Height          =   255
      Left            =   9600
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "From"
      Height          =   255
      Left            =   8040
      TabIndex        =   12
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Alt Assembly ID:"
      Height          =   255
      Left            =   6840
      TabIndex        =   4
      Top             =   1515
      Width           =   1695
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
      Left            =   6780
      TabIndex        =   28
      Top             =   60
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Technical Description:"
      Height          =   255
      Left            =   6840
      TabIndex        =   6
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Assembly ID:"
      Height          =   255
      Left            =   6960
      TabIndex        =   0
      Top             =   1080
      Width           =   975
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   11040
      Y1              =   2820
      Y2              =   2820
   End
   Begin VB.Line Line1 
      X1              =   6660
      X2              =   6660
      Y1              =   2760
      Y2              =   0
   End
   Begin VB.Label lblRowCount 
      Caption         =   "0 rows returned"
      Height          =   255
      Left            =   5160
      TabIndex        =   14
      Top             =   2880
      Width           =   3255
   End
End
Attribute VB_Name = "frmAssemblyGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''' <modulename> frmAssemblyGrid</modulename>
''' <functionname>General (Main) </functionname>
'''
''' <summary>
''' Provides u/i permitting user to do the following:
'''
'''(Major function buttons)
'''
'''1.  Display "Assembly" form         (frmAssembly.frm)
'''2.  Display "Book Detail." form         (frmAssemblyBookGrid.frm)
'''3.  Display "Unit Cost Usage            (frmUCostUsageGrid.frm)
'''4.  Display "CostWorks"
'''5.  Display "Assembly Usage"            (frmAsmUsageGrid.frm)
'''6.  Display "History"               (frmAssemblyHistoryGrid.frm)
'''7.  Save "Update" new/changed data       No form.
''' (m_objGridMap.Update())
'''8.  Create a NEW assembly line          (frmAssembly.frm)
'''9.  Delete an existing assembly line        No form
'''10. Clone a selected material price line        (frmAssembly.frm

'''HELPER Class: CAssemblyMap.Cls
'''</summary>
'''
'''<seealso>CAssemblyMap.cls</seealso>
'''<seealso>frmAssembly.frm</seealso>
'''<seealso>frmAssemblyBookGrid.frm</seealso>
'''<seealso>frmAssemblyBookDetail.frm</seealso>
'''<seealso>frmAssemblyBookGrid.frm</seealso>
'''<seealso>frmAsmUsageGrid.frm</seealso>
'''
'''<datastruct>m_objGridMap</datastruct>
'''<datastruct>m_rec</datastruct>
'''
'''<storedprocedurename> sp_select_assembly</storedprocedurename>
'''<storedprocedurename> sp_update_assembly_driver</storedprocedurename>
'''
'''<returns>N/A</returns>
''' <exception>Always trap with an accompanying message box</exception>
''' <example>
''' <code>
'''exec sp_select_assembly @start_assembly_id='D10100000000', @end_assembly_id='D10109999999', @alt_assembly_id='%', @tech_desc='%', @assembly_type = 0
'''</code>
''' <code>
'''exec sp_update_assembly_driver @type_code='M', @assembly_skey= 20776, @assembly_id='E10101100100', @alt_assembly_id='1112001100', @rev_uni2_L3='     ',
'''@rev_uni2_L5='    ', @rev_uni2_L6='    ', @book_desc='Bank equipment, drive up window, drawer & mike, no glazing, economy', @metric_book_desc='',
'''@tech_desc='Architectural equipment, bank equipment drive up window, drawer & mike, no glazing, economy', @metric_tech_desc='Architectural equipment, bank equipment drive up window, drawer and mike, no glazing, economy',
'''@coml_ind= 1, @resi_ind= 0, @labor_equip_ind= 0, @comment='', @unit='Day', @metric_unit='Ea.', @ad_change_ind= 1, @std_mat_cost= 4950, @std_inst_cost= 660,
'''@std_equip_cost= 0, @std_labor_cost= 660, @std_total_cost= 5610, @std_mat_cost_op= 5450, @std_inst_cost_op= 1200, @std_equip_cost_op= 0, @std_labor_cost_op= 1200,
'''@std_total_cost_op= 6650, @std_labor_hour= 16, @std_change_ind= 0, @opn_mat_cost= 4950, @opn_inst_cost= 495, @opn_equip_cost= 0, @opn_labor_cost= 495, @opn_total_cost= 5445,
'''@opn_mat_cst_op= 5450, @opn_inst_cost_op= 970, @opn_equip_cost_op= 0, @opn_labor_cost_op= 970, @opn_total_cost_op= 6420, @opn_labor_hour= 16, @opn_change_ind= 0,
'''@rr_mat_cost= 4950, @rr_inst_cost= 660, @rr_equip_cost= 0, @rr_labor_cost= 660, @rr_total_cost= 5610,
'''@rr_mat_cost_op= 5450, @rr_inst_cost_op= 1250, @rr_equip_cost_op= 0, @rr_labor_cost_op= 1250, @rr_total_cost_op= 6700,
'''@rr_labor_hour= 16, @rr_change_ind= 0, @metric_mat_cost= 4950, @metric_inst_cost= 660, @metric_equip_cost= 0, @metric_labor_cost= 660,
'''@metric_total_cost= 5610, @metric_mat_cost_op= 5450, @metric_inst_cost_op= 1200, @metric_equip_cost_op= 0, @metric_labor_cost_op= 1200,
'''@metric_total_cost_op= 6650, @metric_labor_hour= 16, @pct_ind= 0, @ad_last_update_id= 3, @std_last_update_id= 1, @opn_last_update_id= 1,
'''@rr_last_update_id= 1, @last_update_person='Hancockrl',  @update_unitcost_usage_ind=0, @cost_change_ind=0</code>
'''</example>
'''<permission>Public</Permission>
'''<dependson>This component depends on the following
'''1.  CAssemblyMap.cls
'''2.  CGridMap.cls
'''3.  CCDdal.CRSMDataAccess (
'''Access to the DAL (data access layer dll) opened in MainModule_Main() )
'''</dependson>


Dim m_objGridMap As New CAssemblyMap ' Class to handle grid
Dim m_blnFirstSearch As Boolean ' Is this the first search we have made on this screen
Dim m_blnJumpIn As Boolean ' Are we jumping here from another screen
Dim m_rec As New ADODB.RecordSet ' Recordset to hold query results
Dim m_blnDoubleClick As Boolean ' Did a double click just occurr
Dim m_blnWereErrors As Boolean ' True if the Update had errors, used in QueryUnload

Dim m_iAssemblyType As Integer
Const USEBOOKMARK = 1
Const USECOORD = 0

Const COMMERCIAL_ASSEMBLIES = 0
Const RESIDENTIAL_ASSEMBLIES = 1

Dim rsAssemblyClone As RecordSet
Dim m_strCurrentFormControl As String

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

Private Sub cboMasterFormat_Change()
MasterFormatChanged
End Sub

Private Sub cboMasterFormat_Click()

MasterFormatChanged
End Sub

Private Sub cmdAssemblySqFtUsage_Click()
If IsNumeric(TDBGrid.Bookmark) = False Then
        MsgBox "You must select a row."
        Exit Sub
    End If
    ' Navigate to single-record view
    Dim frm As frmAsmUsageGrid
    Dim rec As ADODB.RecordSet
    Screen.MousePointer = vbHourglass
    Set frm = New frmAsmUsageGrid
    ' Make copy of recordset
    Set rec = m_rec.Clone
    ' Get the selected row from grid
    rec.Bookmark = TDBGrid.Bookmark
    frm.SetRow rec ' Pass the current record into the form
    Set frm.frmCallingForm = Me
    Set frm.tdbCols = Me.TDBGrid.Columns
    Set frm.myTDBGrid = Me.TDBGrid
    'frm.Show
    frm.JumpIn3 Compress_String(TDBGrid.Columns("Assembly ID").CellText(TDBGrid.Bookmark))
    frm.Show
    Screen.MousePointer = vbDefault
    
    
'     Use this (below) later to get assembly id into assembly usage grid "assembly id" box
    
'     If IsNumeric(TDBGrid.Bookmark) = True Then
'        ' Open spreadsheet view with data from row selected
'        Dim frm As frmUCostUsageGrid
'        Set frm = New frmUCostUsageGrid
'        frm.MasterFormat = MasterFormat
'        frm.JumpIn Compress_String(TDBGrid.Columns("Unit Cost ID").CellText(TDBGrid.Bookmark)) + "*"
'    Else
'        MsgBox "You must select a row."
'    End If
End Sub

Private Sub startassemblyid_LostFocus()
    StartAssemblyID.Text = Trim(StartAssemblyID.Text)
End Sub

' Handles Row Wrap feature
Private Sub ckbRowWrap_Click()
    m_objGridMap.RowWrap (ckbRowWrap)
End Sub

Private Sub cmdAssembly_Click()
    If IsNumeric(TDBGrid.Bookmark) = False Then
        MsgBox "You must select a row."
        Exit Sub
    End If
    ' Navigate to single-record view
    Dim frm As frmAssembly
    Dim rec As ADODB.RecordSet
    Screen.MousePointer = vbHourglass
    Set frm = New frmAssembly
    ' Make copy of recordset
    Set rec = m_rec.Clone
    ' Get the selected row from grid
    rec.Bookmark = TDBGrid.Bookmark
    frm.SetRow rec ' Pass the current record into the form
    Set frm.frmCallingForm = Me
    Set frm.tdbCols = Me.TDBGrid.Columns
    Set frm.myTDBGrid = Me.TDBGrid
    frm.Show
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmdAssemblyBkDtl_Click()
    If IsNumeric(TDBGrid.Bookmark) = False Then
        MsgBox "You must select a row."
        Exit Sub
    End If
'    ' Navigate to Book Detail form
    Dim frm As New frmAssemblyBookGrid
    frm.JumpIn TDBGrid.Columns("Assembly ID").Value
'    frm.Show

End Sub

Private Sub cmdAssemblyUCUsage_Click()
    If IsNumeric(TDBGrid.Bookmark) = True Then
        ' Open spreadsheet view with data from row selected
        Dim frm As frmUCostUsageGrid
        Set frm = New frmUCostUsageGrid
        frm.JumpIn2 TDBGrid.Columns("Assembly ID").CellText(TDBGrid.Bookmark)
    Else
        MsgBox "You must select a row."
    End If

End Sub

Private Sub cmdClone_Click()
    On Error GoTo Out
    If IsNumeric(TDBGrid.Bookmark) = False Then
        MsgBox "You must select a row."
        Exit Sub
    End If
    Dim rec As ADODB.RecordSet
    Dim bln_Continue As Boolean
    If IsNull(TDBGrid.Bookmark) Then
        bln_Continue = True
    ElseIf ValidGridRow = True Then
        bln_Continue = True
    End If
    If bln_Continue = True Then
        Set rec = m_objGridMap.CloneRowRecordset
        ' Navigate to single-record view
        Dim frm As frmAssembly
        Set frm = New frmAssembly
        frm.SetRow rec, True ' Pass the current record into the form
        frm.Show
    End If
Out:
End Sub

Private Sub cmdDelete_Click()
    
    m_objGridMap.Delete
    Screen.MousePointer = vbNormal

End Sub

Private Sub cmdUnitCost_Click()
    ' Navigate to grid view
    Dim frm As frmUnitCostGrid
    Set frm = New frmUnitCostGrid
    frm.JumpIn TDBGrid.Columns("Unit Cost ID").CellText(TDBGrid.Bookmark)
End Sub

Private Sub cmdHistory_Click()
    If IsNumeric(TDBGrid.Bookmark) = False Then
        MsgBox "You must select a row."
        Exit Sub
    End If
    If TDBGrid.Columns("Type").CellText(TDBGrid.Bookmark) = "H" Then
        MsgBox "No history is available for header (H) rows."
        Exit Sub
    End If
    ' Open single record view with data from row selected
    Dim frm As frmAssemblyHistoryGrid
    Set frm = New frmAssemblyHistoryGrid
    frm.JumpIn TDBGrid.Columns("Assembly ID").CellText(TDBGrid.Bookmark)
End Sub

Private Sub cmdNew_Click()
'TDBGrid.SetFocus
'TDBGrid.MoveLast
'TDBGrid.Row = TDBGrid.Row + 1
    On Error GoTo Out
    Dim rec As New ADODB.RecordSet
    Dim bln_Continue As Boolean
    If IsNull(TDBGrid.Bookmark) Then
        bln_Continue = True
    ElseIf ValidGridRow = True Then
        bln_Continue = True
    End If
    If bln_Continue = True Then
        CopyRSFields rec, m_rec
        ' Open empty single record view
        Dim frm As frmAssembly
        Set frm = New frmAssembly
        ' Force any changes into recordset from grid
        TDBGrid.Update
        frm.SetRow rec, True
        frm.Show
    End If
Out:

End Sub

Private Sub cmdUpdate_Click()
    Dim blnRet As Boolean
    Dim vntBookmark As Variant
    Dim bln_Continue As Boolean
    
    m_blnWereErrors = False
    If IsNull(TDBGrid.Bookmark) Then
        bln_Continue = True
    ElseIf ValidGridRow = True Then
        bln_Continue = True
    End If
    If bln_Continue = True Then
        vntBookmark = TDBGrid.Bookmark
        TDBGrid.Update
        blnRet = m_objGridMap.Update
        If blnRet = False Then
            m_blnWereErrors = True
        End If
        TDBGrid.Bookmark = vntBookmark
    End If
End Sub

Private Function ValidGridRow() As Boolean

    If TDBGrid.Columns("Comm'l Ind").Value = 0 And TDBGrid.Columns("Resi Ind").Value = 0 Then
        TDBGrid.SetFocus
        MsgBox "Commercial or Residential Use indicator must be selected."
        ValidGridRow = False
    ElseIf Len(Trim(TDBGrid.Columns("Assembly ID").Value)) = 0 Then
        ValidGridRow = False
        MsgBox "Please enter a valid Assembly Id."
    Else
        ValidGridRow = True
    End If

End Function

Private Sub Form_Deactivate()
    m_strCurrentFormControl = Me.ActiveControl.Name
End Sub

Private Sub Form_Initialize()
    ' Initialize grid
    Screen.MousePointer = vbHourglass
    m_blnFirstSearch = True
    FormatTree.InitData g_cnShared, "ASSEMBLY_COMMERCIAL"
    m_objGridMap.SetGrid TDBGrid
    m_objGridMap.InitGrid
    m_blnJumpIn = False
    Screen.MousePointer = vbNormal
    m_blnFirstSearch = False
    
    MASTER_FORMAT_ASSEMBLIES = 2004 'by default   (rlh)  05/07/2008
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 And Screen.ActiveControl.Name = "TDBGrid" Then       'Escape pressed
        If TDBGrid.AddNewMode > 0 Then
            TDBGrid.DataChanged = False
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim blnReturn As Boolean
    Dim strSelect As String
    
    'rlh 05/06/2008
    LoadMasterFormatCombo Me.cboMasterFormat       'rlh
    
    
    Move START_LEFT, START_TOP, START_WIDTH, START_HEIGHT
    
    ' This will never return any rows, just used to create recordset
    StartAssemblyID.Text = "~"
    cmdSearch_Click
    StartAssemblyID.Text = ""
    
    'StartAssemblyID.SetFocus
    
End Sub

' Called when coming here from another screen
Public Sub JumpIn(strAssemblyId As String)
    StartAssemblyID.Text = strAssemblyId
    cmdSearch_Click
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
            Frame1.Top = Me.Height - 1260
            cmdUpdate.Top = Me.Height - 1020
            cmdNew.Top = Me.Height - 1020
            cmdClone.Top = Me.Height - 1020
            cmdDelete.Top = Me.Height - 1020
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
    Dim rs As ADODB.RecordSet
    Dim strSelect As String
    Dim blnReturn As Boolean
    
    On Error Resume Next
    ' Synch text box with tree
    rs.Close ' Make sure it is closed
    strSelect = "select assembly_id_start, assembly_id_end from UNIFORMAT2_ID_HIERARCHY where uni2_category_id='" + strID + "'"
    blnReturn = g_objDAL.GetRecordset(vbNullString, strSelect, rs)
    StartAssemblyID.Text = rs.Fields("assembly_id_start")
    EndAssemblyID.Text = rs.Fields("assembly_id_end")
    ' Clear other boxes
    rs.Close
    TechDesc.Text = ""
    ' Kick-off search
    cmdSearch_Click
End Sub

Private Sub cmdSearch_Click()
    On Error Resume Next
    Dim blnReturn As Boolean
    Dim strSelect As String
    Dim dtmToday As Date
    Dim dtmStart As Date
    Dim strError As String
    Dim strSrchStartAssemblyID As String
    Dim strSrchEndAssemblyID As String
    Dim strSrchTechDesc As String
    Dim strSrchAltAssemblyId As String
    
    
    dtmToday = Date
    rsAssemblyClone.Close
    Set rsAssemblyClone = Nothing
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
    lblRowCount.Caption = "Working..."
    lblRowCount.Refresh
    
    ' Synch tree with text box
    If Not StartAssemblyID.Text = "" Then
        FormatTree.FocusItem (Compress_String(StartAssemblyID.Text))
    End If
    
    m_rec.Close ' Make sure it is closed
    m_rec.MaxRecords = MAX_RECORDS ' Set the maximum number to bring back
    dtmStart = Now
    
    If Len(StartAssemblyID.Text) = 0 And Len(TechDesc.Text) = 0 And Len(AltAssemblyId.Text) = 0 Then
        Screen.MousePointer = vbNormal
        MsgBox "You must enter Search Criteria"
        Exit Sub
    End If

    If StartAssemblyID.Text = "" Then
        strSrchStartAssemblyID = "*"
    Else
        strSrchStartAssemblyID = Compress_String(StartAssemblyID.Text)
    End If
    If EndAssemblyID.Text = "" Then
        strSrchEndAssemblyID = StartAssemblyID.Text
    Else
        strSrchEndAssemblyID = Compress_String(EndAssemblyID.Text)
    End If
    If TechDesc.Text = "" Then
        strSrchTechDesc = "*"
    Else
        strSrchTechDesc = TechDesc.Text
    End If
    If AltAssemblyId.Text = "" Then
        strSrchAltAssemblyId = "*"
    Else
        strSrchAltAssemblyId = AltAssemblyId.Text
    End If
    
    strSelect = "exec sp_select_assembly @start_assembly_id='" + _
    SQLChangeWildcard(strSrchStartAssemblyID) + "', @end_assembly_id='" + _
    SQLChangeWildcard(strSrchEndAssemblyID) + "', @alt_assembly_id='" + _
    SQLChangeWildcard(strSrchAltAssemblyId) + "', @tech_desc='" + _
    SQLFixString(SQLChangeWildcard(strSrchTechDesc)) + "'" + _
    ", @assembly_type = " + CStr(m_iAssemblyType)
    ' Use DAL to perform select
    blnReturn = g_objDAL.GetRecordset(vbNullString, strSelect, m_rec)
    If blnReturn = False Then
        MsgBox "An error occurred while searching. Error:" + Error$
        lblRowCount.Caption = "0 rows returned."
        Screen.MousePointer = vbNormal
        Exit Sub
    End If
    
    ' Pass recordset to handler class
    m_objGridMap.RecordSet = m_rec
    
    If m_rec.RecordCount > 0 Then
        lblRowCount.Caption = str(m_rec.RecordCount) + " rows returned in " + str(DateDiff("s", dtmStart, Now)) + " seconds"
        cmdClone.Enabled = True
        cmdDelete.Enabled = True
        cmdHistory.Enabled = True
        cmdUpdate.Enabled = True
        cmdAssembly.Enabled = True
        cmdAssemblyBkDtl.Enabled = True
        cmdAssemblyUCUsage.Enabled = True
        Set rsAssemblyClone = m_rec.Clone
    Else
        lblRowCount.Caption = "0 rows returned."
        cmdAssemblyBkDtl.Enabled = False
        cmdAssemblyUCUsage.Enabled = False
        cmdClone.Enabled = False
        cmdHistory.Enabled = False
        cmdDelete.Enabled = False
        cmdUpdate.Enabled = False
        cmdAssembly.Enabled = False
    End If
    
    ' If the upper bound was hit, inform user
    If m_rec.RecordCount = MAX_RECORDS And m_rec.State = adStateOpen Then
        MsgBox "The search returned the maximum number of records allowed. More records may be available."
    End If

    ' Reset the grid contents
    TDBGrid.Bookmark = Null
    TDBGrid.ReBind
    TDBGrid.ApproxCount = m_rec.RecordCount
    SetButtons USEBOOKMARK
    Screen.MousePointer = vbNormal
End Sub

Private Sub SetButtons(Mode As Single, Optional Coord As Variant)
    On Error GoTo Exit_Sub
    Select Case Mode
        Case USEBOOKMARK
            rsAssemblyClone.Bookmark = TDBGrid.Bookmark
        Case USECOORD
            rsAssemblyClone.Bookmark = TDBGrid.RowBookmark(TDBGrid.RowContaining(Coord))
    End Select
    
    'No Unit Costs should be associated with E records
    If rsAssemblyClone.Fields("type_code") = "E" Or rsAssemblyClone.Fields("type_code") = "M" Then
    '    cmdAssemblySqFtUsage.Enabled = False
        cmdAssemblyUCUsage.Enabled = True
    Else
    '    cmdAssemblySqFtUsage.Enabled = True
        cmdAssemblyUCUsage.Enabled = False
    End If
    
    cmdAssemblyUCUsage.Enabled = True  'rlh - temporary
    
Exit_Sub:
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

Private Sub optAssemblyType_Click(Index As Integer)

    FormatTree.ClearTree
    Select Case Index
        ' Fill the MasterFormat tree
        Case 0 'Commercial
            FormatTree.InitData g_cnShared, "ASSEMBLY_COMMERCIAL"
            m_iAssemblyType = COMMERCIAL_ASSEMBLIES
        Case 1 'Residential
            FormatTree.InitData g_cnShared, "ASSEMBLY_RESI"
            m_iAssemblyType = RESIDENTIAL_ASSEMBLIES
    End Select
    If StartAssemblyID.Text = "" And TechDesc.Text = "" Then
        StartAssemblyID.Text = "~"
    End If
    cmdSearch_Click
    If StartAssemblyID.Text = "~" And TechDesc.Text = "" Then
        StartAssemblyID.Text = ""
    End If

End Sub

Private Sub TDBGrid_DblClick()
    ' Signal that double-click has occurred
    m_blnDoubleClick = True
End Sub

Private Sub TDBGrid_GotFocus()
    TDBGrid.TabStop = True
End Sub

Private Sub TDBGrid_KeyUp(KeyCode As Integer, Shift As Integer)
    SetButtons USEBOOKMARK

End Sub

Private Sub TDBGrid_LostFocus()
    TDBGrid.TabStop = False
End Sub

Private Sub TDBGrid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If this is the mouse-up form a double click
    If m_blnDoubleClick Then
        ' Make sure it is the left button
        If Button = vbLeftButton Then
            m_blnDoubleClick = False
            ' Same function as clicking Assembly button, open single record view
            cmdAssembly_Click
        End If
    Else
        TDBGrid.Splits(0).ClearSelCols
        If Button = vbRightButton Then
            Dim strErrorMsg As String
            If Not IsNull(TDBGrid.Bookmark) Then
                strErrorMsg = m_objGridMap.GetError(TDBGrid.Bookmark)
                If Len(strErrorMsg) > 0 Then
                    MsgBox strErrorMsg
                End If
            End If
        Else
            SetButtons USECOORD, Y
        End If
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
        TDBGrid.Refresh
        OutputView False
        ShowGridSort
        m_objGridMap.SetMenuBar
    End If
End Sub

Private Sub techdesc_LostFocus()
    TechDesc.Text = Trim(TechDesc.Text)
End Sub
Private Sub MasterFormatChanged()
'A NEW MASTERFORMAT WAS SELECTED FROM THE DROP-DOWN BOX
'ADDED 6/20/2005 RTD FOR VERSION 7.4.0+
    Dim sTreeType As String
    
    If cboMasterFormat.ListIndex < 0 Then
        Exit Sub
    End If
    
    If MF95_ENABLED Then
        Select Case cboMasterFormat.ItemData(cboMasterFormat.ListIndex)
        Case EXT_MASTERFORMAT_VERSION
    '        UnLockField Me, "EndUnitCostID"
    '        lblUnitCostId.Caption = "Unit Cost ID " & Right(EXT_MASTERFORMAT_VERSION, 2) & ":"
            'sTreeType = "UNITCOST" & Right(EXT_MASTERFORMAT_VERSION, 2)
            MASTER_FORMAT_ASSEMBLIES = EXT_MASTERFORMAT_VERSION
        Case UCD_MASTERFORMAT_VERSION
'            UnLockField Me, "EndUnitCostID"
'            lblUnitCostId.Caption = "Unit Cost ID " & Right(UCD_MASTERFORMAT_VERSION, 2) & ":"
'            sTreeType = "UNITCOST"
            MASTER_FORMAT_ASSEMBLIES = UCD_MASTERFORMAT_VERSION
        Case ALT_MASTERFORMAT_VERSION
'            LockField Me, "EndUnitCostID"
'            'EndUnitCostID.Text = ""
'            lblUnitCostId.Caption = "Alt Unit Cost ID:"
'            sTreeType = "UNITCOST"
        Case Else
'            UnLockField Me, "EndUnitCostID"
'            lblUnitCostId.Caption = "Unit Cost ID " & Right(UCD_MASTERFORMAT_VERSION, 2) & ":"
'            sTreeType = "UNITCOST"
            MASTER_FORMAT_ASSEMBLIES = UCD_MASTERFORMAT_VERSION
    
        End Select
    Else
        Select Case cboMasterFormat.ItemData(cboMasterFormat.ListIndex)
        Case EXT_MASTERFORMAT_VERSION
    '        UnLockField Me, "EndUnitCostID"
    '        lblUnitCostId.Caption = "Unit Cost ID " & Right(EXT_MASTERFORMAT_VERSION, 2) & ":"
            'sTreeType = "UNITCOST" & Right(EXT_MASTERFORMAT_VERSION, 2)
            MASTER_FORMAT_ASSEMBLIES = EXT_MASTERFORMAT_VERSION
    

    End Select
    End If
    
    
    'CHECK IF WE NEED TO RE-INITIALIZE TREE
'    If FormatTree.TreeType <> sTreeType Then
'        Screen.MousePointer = vbHourglass
'        FormatTree.DisableRedraw = True
'        FormatTree.ClearTree
'        FormatTree.InitData g_cnShared, sTreeType
'        FormatTree.DisableRedraw = False
'        Screen.MousePointer = vbDefault
'    End If

    On Error Resume Next
'    StartUnitCostID.SetFocus
    Screen.MousePointer = vbDefault

End Sub
Public Function SelectMasterFormat(iMasterFormat As Long) As Boolean
'SET THE MASTERFORMAT COMBO BOX TO THE NEW SELECTION
'ADDED 8/2/2005 RTD
    Dim i As Long
    
    cboMasterFormat.ListIndex = -1
    For i = 0 To cboMasterFormat.listcount - 1
        If cboMasterFormat.ItemData(i) = iMasterFormat Then
            cboMasterFormat.ListIndex = i
            SelectMasterFormat = True
            Exit For
        End If
    Next
    
End Function
