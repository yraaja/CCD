VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{5936A75C-3F42-11D6-AF6B-AA0004005F12}#1.3#0"; "MeansCtrl.ocx"
Begin VB.Form frmUnitCostGrid 
   Caption         =   "Unit Cost Grid"
   ClientHeight    =   6870
   ClientLeft      =   1500
   ClientTop       =   2790
   ClientWidth     =   11595
   Icon            =   "frmUnitCostGrid.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6870
   ScaleWidth      =   11595
   Begin TrueOleDBGrid80.TDBGrid TDBGrid 
      Height          =   2655
      Left            =   120
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   3240
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   4683
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
      Splits(0)._ColumnProps(5)=   "Column(0)._MinWidth=29552"
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
      RowDividerStyle =   7
      Caption         =   "MasterFormat"
      MultipleLines   =   0
      CellTips        =   1
      CellTipsWidth   =   0
      DeadAreaBackColor=   12632256
      RowDividerColor =   -2147483632
      RowSubDividerColor=   -2147483632
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
      _StyleDefs(58)  =   ":id=35,.parent=29,.bgcolor=&HC0C0C0&"
      _StyleDefs(59)  =   "Named:id=36:OddRow"
      _StyleDefs(60)  =   ":id=36,.parent=29"
      _StyleDefs(61)  =   "Named:id=39:RecordSelector"
      _StyleDefs(62)  =   ":id=39,.parent=30"
      _StyleDefs(63)  =   "Named:id=42:FilterBar"
      _StyleDefs(64)  =   ":id=42,.parent=29"
   End
   Begin VB.ComboBox cboMasterFormat 
      Height          =   315
      ItemData        =   "frmUnitCostGrid.frx":0442
      Left            =   8400
      List            =   "frmUnitCostGrid.frx":0444
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox EndUnitCostID 
      Height          =   315
      Left            =   9960
      TabIndex        =   1
      Top             =   1260
      Width           =   1515
   End
   Begin VB.TextBox altunitcostid 
      Height          =   315
      Left            =   10200
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.TextBox Description 
      Height          =   315
      Left            =   8400
      TabIndex        =   3
      Top             =   1680
      Width           =   3080
   End
   Begin VB.CommandButton cmdClone 
      Caption         =   "Clone"
      Height          =   495
      Left            =   11280
      TabIndex        =   10
      Top             =   6195
      Width           =   795
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   10360
      TabIndex        =   9
      Top             =   6195
      Width           =   795
   End
   Begin VB.Frame Frame1 
      Caption         =   "Go To"
      Height          =   855
      Left            =   120
      TabIndex        =   12
      Top             =   5955
      Width           =   8220
      Begin VB.CommandButton cmdOutput 
         Caption         =   "&Output"
         Height          =   495
         Left            =   4875
         TabIndex        =   31
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdCrews 
         Caption         =   "Crews"
         Enabled         =   0   'False
         Height          =   495
         Left            =   5688
         TabIndex        =   30
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdHistory 
         Caption         =   "History"
         Height          =   495
         Left            =   6501
         TabIndex        =   29
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdUnitCost 
         Caption         =   "Unit Cost"
         Height          =   495
         Left            =   993
         TabIndex        =   28
         Top             =   240
         Width           =   915
      End
      Begin VB.CommandButton cmdUnitCostUsage 
         Caption         =   "Unit Cost    Usage"
         Height          =   495
         Left            =   3957
         TabIndex        =   27
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton cmdMaterialUsage 
         Caption         =   "Material Usage"
         Height          =   495
         Left            =   3039
         TabIndex        =   26
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton cmdLongDesc 
         Caption         =   "     &Long Description"
         Height          =   495
         Left            =   1986
         TabIndex        =   25
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdSearchExtID 
         Caption         =   "MF 04 Record"
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   795
      End
      Begin VB.CommandButton cmdBookPreview 
         Caption         =   "  &Book Preview"
         Height          =   495
         Left            =   7320
         TabIndex        =   23
         Top             =   240
         Width           =   795
      End
      Begin VB.CommandButton cmdCostworks 
         Caption         =   "CostWorks"
         Enabled         =   0   'False
         Height          =   495
         Left            =   7200
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   495
      Left            =   9440
      TabIndex        =   8
      Top             =   6195
      Width           =   795
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   495
      Left            =   8520
      TabIndex        =   7
      Top             =   6195
      Width           =   795
   End
   Begin VB.CheckBox ckbRowWrap 
      Caption         =   "Row Wrap"
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   495
      Left            =   8400
      TabIndex        =   4
      Top             =   2160
      Width           =   1150
   End
   Begin VB.TextBox StartUnitCostID 
      Height          =   315
      Left            =   8400
      TabIndex        =   0
      Top             =   1260
      Width           =   1515
   End
   Begin ConstructionCostDatabase.DynaTree FormatTree 
      Height          =   2775
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   4895
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "MasterFormat:"
      Height          =   255
      Left            =   6900
      TabIndex        =   21
      Top             =   650
      Width           =   1335
   End
   Begin VB.Label lblUnitCostId 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Cost ID:"
      Height          =   255
      Left            =   6900
      TabIndex        =   19
      Top             =   1310
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      Height          =   255
      Left            =   9960
      TabIndex        =   18
      Top             =   1020
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      Height          =   255
      Left            =   8400
      TabIndex        =   17
      Top             =   1020
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Alt Unit Cost ID:"
      Height          =   255
      Left            =   10200
      TabIndex        =   16
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tech Desc:"
      Height          =   255
      Left            =   7020
      TabIndex        =   15
      Top             =   1730
      Width           =   1215
   End
   Begin VB.Label lblRowCount 
      Alignment       =   2  'Center
      Caption         =   "0 rows returned"
      Height          =   195
      Left            =   4560
      TabIndex        =   14
      Top             =   2880
      Width           =   6510
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
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
      TabIndex        =   13
      Top             =   60
      Width           =   1215
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   11220
      Y1              =   2820
      Y2              =   2820
   End
   Begin VB.Line Line1 
      X1              =   6660
      X2              =   6660
      Y1              =   2700
      Y2              =   60
   End
End
Attribute VB_Name = "frmUnitCostGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_objGridMap As New CUnitCostMap ' Class to handle grid
Dim m_blnFirstSearch As Boolean ' Is this the first search we have made on this screen
Dim m_blnJumpIn As Boolean ' Are we jumping here from another screen
Dim m_rec As New ADODB.RecordSet ' Recordset to hold query results
Dim m_blnDoubleClick As Boolean ' Did a double click just occurr
Dim m_blnWereErrors As Boolean ' True if the Update had errors, used in QueryUnload
Dim rsUnitCostClone As RecordSet
Dim m_sngYCoord As Single
Dim m_strCurrentFormControl As String
Dim m_intMasterFormat As Long   ' Stores MasterFormat version to use by Search et al
Dim m_blnMasterFormatNotSpecified As Boolean    ' True if MF was never explicitly set

Const USEBOOKMARK = 1
Const USECOORD = 0

' <modulename> frmUnitCostGrid</modulename>
' <functionname>General (Main) </functionname>
' <summary>
' Provides u/i permitting user to do the following:
'    1.  Flip flop from  MF95 to MF04
'    2.  Display Unit Cost Tabs form
'    3.  Display Long Descriptions for selected unit cost line (ucl)
'    4.  Display Material Usage for selected ucl
'    5.  Set "book type assignments" / Output usage
'    6.  Open "Crews" dialog
'    7.  Open "History" dialog reflecting significant data change per selected ucl
'    8.  Open "Book Preview" rendering of a selected book level/line
'    9.  Update / Save any changes to unit cost related data
'    10. Delete a selected unit cost line
'    11. Clone a selected ucl
' </summary>
' <seealso>N/A</seealso>
' <datastruct>m_rec</datastruct>
' <storedprocedurename> usp_select_unit_cost_ext_rlh2</storedprocedurename>
' <storedprocedurename> sp_temp_output_init</storedprocedurename>
' <param name="data">
'       ???a dataset containing all the data for updating ?
' </param>
' <param name="someParameter">
'       ???? Description of someParameter goes here  updating
' </param>
' <returns>N/A</returns>
' <exception>Always trap with an accompanying message box</exception>
' <example>
' <code>
'       exec usp_select_unit_cost_ext_rlh2 @start_unit_cost_id = '030100000000', @end_unit_cost_id = '030499999999', @alt_unit_cost_id = '', @tech_desc = '', @master_format=2004
' </code>
' <code>
'       exec sp_temp_output_init
'</code>
'</example>
'<permission>Public</Permission>
'<dependson>This component depends on the following
'    CUnitCostMap.Cls
'    CGridMap.Cls
'    CCDdal.CRSMDataAccess (
'       Access to the DAL (data access layer dll) opened in MainModule_Main() )
'</dependson>


' MasterFormat property
' Returns/sets the CSI MasterFormat version of the Unit Cost IDs
Public Property Get MasterFormat() As Long
    MasterFormat = m_intMasterFormat
End Property
Public Property Let MasterFormat(NewValue As Long)
    m_intMasterFormat = NewValue
    SelectMasterFormat m_intMasterFormat
    m_blnMasterFormatNotSpecified = False
End Property

Public Sub Sort(intDir As Integer)
    m_objGridMap.Sort intDir
End Sub

Private Sub position_output(Optional Y As Single = 0)
    Dim sKey As String
    ' Only send data to the Output dialog if it is open
    Dim frm As Form
    Dim blnVisible As Boolean
    If FormOpen("dlgOutput", frm, blnVisible) = True Then
        If blnVisible = True Then
            If m_rec.BOF Or m_rec.EOF Then
                DoOutput
            Else
                DoOutput
                Me.SetFocus
            End If
        End If
    End If
End Sub

Private Sub SetButtons(Mode As Single, Optional Coord As Variant)
    
    On Error GoTo Exit_Sub
    Select Case Mode
        Case USEBOOKMARK
            rsUnitCostClone.Bookmark = TDBGrid.Bookmark
        Case USECOORD
            rsUnitCostClone.Bookmark = TDBGrid.RowBookmark(TDBGrid.RowContaining(Coord))
    End Select
    
    'No material should be associated with H or E records
    If rsUnitCostClone.Fields("type_code") = "H" Or rsUnitCostClone.Fields("type_code") = "E" _
             Then
        cmdMaterialUsage.Enabled = False
        cmdCrews.Enabled = False
        cmdClone.Enabled = True
        If rsUnitCostClone.Fields("type_code") = "E" Then
            cmdUnitCostUsage.Enabled = True
            cmdClone.Enabled = True
        Else
            cmdUnitCostUsage.Enabled = False
            
        End If
    Else
        If rsUnitCostClone.Fields("type_code") = "B" Then
            cmdMaterialUsage.Enabled = False
            cmdClone.Enabled = False
        Else
            cmdMaterialUsage.Enabled = True
            cmdClone.Enabled = True
        End If
        cmdUnitCostUsage.Enabled = True
        
        If Len(rsUnitCostClone.Fields("crew_id")) > 0 Then
            cmdCrews.Enabled = True
        Else
            cmdCrews.Enabled = False
        End If
    End If
    
Exit_Sub:
    
End Sub

Public Sub SelectAllRows()
    ' Pass recordset to handler class
    m_objGridMap.RecordSet = m_rec
    
    If m_rec.RecordCount > 0 Then
        m_objGridMap.SelectAllRows
    End If
End Sub

Private Sub altunitcostid_LostFocus()
    altunitcostid = Trim(altunitcostid)
End Sub

Private Sub cboMasterFormat_Click()
    MasterFormatChanged
End Sub

' Handles Row Wrap feature
Private Sub ckbRowWrap_Click()
    m_objGridMap.RowWrap (ckbRowWrap)
End Sub

Private Sub cmdAssembly_Click()
'    If IsNumeric(TDBGrid.Bookmark) = False Then
'        MsgBox "You must select a row."
'        Exit Sub
'    End If
'    ' Navigate to single-record view
'    Dim frm As frmEquipment
'    Dim rec As ADODB.RecordSet
'    Set frm = New frmEquipment
'    ' Make copy of recordset
'    Set rec = m_rec.Clone
'    ' Get the selected row from grid
'    rec.Bookmark = TDBGrid.Bookmark
'    frm.SetRow rec ' Pass the current record into the form
'    frm.Show
End Sub

Private Sub cmdBookPreview_Click()
    PreviewReport
End Sub

Private Sub cmdClone_Click()
    On Error GoTo Out
    If IsNumeric(TDBGrid.Bookmark) = False Then
        MsgBox "You must select a row."
        Exit Sub
    End If
    Dim rec As ADODB.RecordSet
       
    m_blnClone_pub = True     'rlh ccd 8.4  04/15/2009
    
    Set rec = m_objGridMap.CloneRowRecordset
    ' Navigate to single-record view
    Dim frm As frmUnitCost
    Set frm = New frmUnitCost
    frm.SetRow rec, True ' Pass the current record into the form
    ' Set MasterFormat so cloned record is saved appropriately
    frm.MasterFormat = m_intMasterFormat
'''    frm.ext_unit_cost_id = ""           'rlh 04/2008
    frm.Show
Out:
End Sub

Private Sub cmdCostworks_Click()
    Dim clsUnitCostOut As New CUnitCostOut
    Screen.MousePointer = vbHourglass
    ' If there are selected rows, just show them
    If TDBGrid.SelBookmarks.Count > 0 Then
        Dim vntBookmark As Variant
        For Each vntBookmark In TDBGrid.SelBookmarks
            m_rec.Bookmark = vntBookmark
            clsUnitCostOut.Add m_rec.Fields(0)
        Next
    ' Otherwise do all rows in grid
    Else
        If Not (m_rec.BOF And m_rec.EOF) Then
            m_rec.MoveFirst
            While Not m_rec.EOF
                clsUnitCostOut.Add m_rec.Fields(0)
                m_rec.MoveNext
            Wend
        End If
    End If
    clsUnitCostOut.Done
    Screen.MousePointer = vbNormal

End Sub

Private Sub cmdDelete_Click()
    m_objGridMap.Delete
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
    Dim frm As frmUCostHistoryGrid
    Set frm = New frmUCostHistoryGrid
    frm.MasterFormat = MasterFormat
    frm.JumpIn TDBGrid.Columns("Unit Cost ID").CellText(TDBGrid.Bookmark)
End Sub

Private Sub cmdLongDesc_Click()
    Dim sUnitCostId As String
    
    If IsNumeric(TDBGrid.Bookmark) = True Then
        ' Open long description grid view with data from row selected
        ' 8/24/2005 RTD - Set long description MasterFormat for unit cost ID
        sUnitCostId = TDBGrid.Columns("Unit Cost ID").CellText(TDBGrid.Bookmark)
        Dim frm As frmLongDescriptionGrid
        Set frm = New frmLongDescriptionGrid
        frm.MasterFormat = MasterFormat
        frm.JumpIn Compress_String(sUnitCostId)
    Else
        MsgBox "You must select a row.", vbInformation
    End If

End Sub

Private Sub cmdMaterialUsage_Click()

    If IsNumeric(TDBGrid.Bookmark) = False Then
        MsgBox "You must select a row.", vbInformation
        Exit Sub
    End If
    ' Open spreadsheet view with data from row selected
    Dim frm As frmMatUsageGrid
    Set frm = New frmMatUsageGrid
    frm.strSource = "Unit Cost"
    frm.MasterFormat = MasterFormat
    frm.JumpIn2 Compress_String(TDBGrid.Columns("Unit Cost ID").CellText(TDBGrid.Bookmark))
    
End Sub

Private Sub cmdOutput_Click()
    Dim frm As Form
    Dim blnVisible As Boolean
    
    m_strKeyType2 = "U"      'flag as unit cost line processing  (rlh) 07/14/2008
    
    If FormOpen("dlgOutput", frm, blnVisible) = True Then
        If blnVisible = False Then
            frm.Visible = True
            DoOutput
        Else
            frm.Visible = False
        End If
    Else
        DoOutput
    End If
End Sub

Private Sub cmdSearchExtID_Click()
'GET THE OPPOSITE MASTERFORMAT ID AND EXECUTE A SEARCH TO RETURN ITS DATA
    Dim sUnitCostId As String
    
    If MF95_ENABLED = False Then    'rlh 2/19/2009
        Exit Sub                    'rlh 02/10/2009  MF95 Disable
    End If
    
    If IsNumeric(TDBGrid.Bookmark) = True Then
        sUnitCostId = TDBGrid.Columns(m_objGridMap.ExtUnitCostIDColumnName).CellText(TDBGrid.Bookmark)
        If m_intMasterFormat = EXT_MASTERFORMAT_VERSION Then
            If sUnitCostId = "" Then
                MsgBox "This Unit Cost row does not have a corresponding MasterFormat " & UCD_MASTERFORMAT_VERSION & " Unit Cost ID.", vbInformation
                Exit Sub
            End If
            SelectMasterFormat UCD_MASTERFORMAT_VERSION
        Else
            If sUnitCostId = "" Then
                MsgBox "This Unit Cost row does not have a corresponding MasterFormat " & EXT_MASTERFORMAT_VERSION & " Unit Cost ID.", vbInformation
                Exit Sub
            End If
            'SelectMasterFormat EXT_MASTERFORMAT_VERSION  'rlh 02/10/2009  MF95 Disable
        End If
        ' Kick off Search
        StartUnitCostID.Text = sUnitCostId
        EndUnitCostID.Text = ""
        Description.Text = ""
        cmdSearch_Click
    Else
        MsgBox "You must select a row.", vbInformation
    End If
    
End Sub

Private Sub cmdUnitCost_Click()
    If IsNumeric(TDBGrid.Bookmark) = False Then
        MsgBox "You must select a row.", vbInformation
        Exit Sub
    End If
    ' Navigate to single-record view
    Dim frm As frmUnitCost
    Dim rec As ADODB.RecordSet
    Set frm = New frmUnitCost
    ' Make copy of recordset
    Set rec = m_rec.Clone
    ' Get the selected row from grid
    rec.Bookmark = TDBGrid.Bookmark
    frm.SetRow rec                          ' Pass the current record into the form
    frm.MasterFormat = m_intMasterFormat    ' Pass the current MasterFormat to the form
    Set frm.frmCallingForm = Me
    Set frm.tdbCols = Me.TDBGrid.Columns
    Set frm.myTDBGrid = Me.TDBGrid
    frm.Show

End Sub

Private Sub cmdCrews_Click()
    ' Open single record view with data from row selected
    Dim frm As frmCrewGrid
    Set frm = New frmCrewGrid
    frm.JumpIn2 Trim(TDBGrid.Columns("Crew ID").CellText(TDBGrid.Bookmark))
End Sub

Private Sub cmdNew_Click()
    On Error GoTo Out
    Dim rec As New ADODB.RecordSet
    
    m_blnNew_pub = True     'rlh ccd 8.4  04/15/2009
    
    CopyRSFields rec, m_rec
    ' Open empty single record view
    Dim frm As frmUnitCost
    Set frm = New frmUnitCost
    ' Force any changes into recordset from grid
    TDBGrid.Update
    frm.SetRow rec, True
    ' Set MasterFormat so new record is saved appropriately
    frm.MasterFormat = m_intMasterFormat
    frm.Show
    Exit Sub
    
Out:

End Sub

Private Sub cmdUnitCostUsage_Click()
    If IsNumeric(TDBGrid.Bookmark) = True Then
        ' Open spreadsheet view with data from row selected
        Dim frm As frmUCostUsageGrid
        Set frm = New frmUCostUsageGrid
        frm.MasterFormat = MasterFormat
        frm.JumpIn Compress_String(TDBGrid.Columns("Unit Cost ID").CellText(TDBGrid.Bookmark)) + "*"
    Else
        MsgBox "You must select a row."
    End If

End Sub

Private Sub cmdUpdate_Click()
    On Error GoTo Out
    Dim blnRet As Boolean
    Dim vntBookmark As Variant
    
    m_blnWereErrors = False
    vntBookmark = TDBGrid.Bookmark
    If TDBGrid.DataChanged = True Then
        If TDBGrid.Columns("Crew Type").Value = "L" Then
            If Not IsNumeric(TDBGrid.Columns("Crew Qty")) Then
                MsgBox "Please enter a valid Crew quantity."
                Exit Sub
            ElseIf Val(CInt(TDBGrid.Columns("Crew Qty"))) <> Val(TDBGrid.Columns("Crew Qty")) Then
                MsgBox "Please enter a valid Crew quantity."
                Exit Sub
                Screen.MousePointer = vbNormal
            ElseIf Val(TDBGrid.Columns("Crew Qty")) < 1 Then    'Quantity is required
                MsgBox "Please enter a valid Crew quantity."
                Exit Sub
                Screen.MousePointer = vbNormal
            End If
        Else
            If TDBGrid.Columns("Type") <> "H" And Trim(TDBGrid.Columns("Crew ID")) <> "" And Trim(TDBGrid.Columns("Crew Qty")) = "" Then
                MsgBox "Please enter a valid Crew quantity."
                Exit Sub
            End If
        End If
    End If
    If TDBGrid.Columns("Type") = "M" Then
        If Len(Trim(TDBGrid.Columns("Unit"))) = 0 Then    'unit required
            MsgBox "Please enter a valid Unit of Measure."
            Exit Sub
        End If
    End If
    TDBGrid.Update
    blnRet = m_objGridMap.Update
    If blnRet = False Then
        m_blnWereErrors = True
        Screen.MousePointer = vbNormal
    End If
    TDBGrid.Bookmark = vntBookmark
    Screen.MousePointer = vbNormal

Out:

End Sub

Private Sub Description_LostFocus()
    Description = Trim(Description)
End Sub

Private Sub EndUnitCostID_LostFocus()
    EndUnitCostID = Trim(EndUnitCostID)
End Sub

Private Sub Form_Deactivate()
    m_strCurrentFormControl = Me.ActiveControl.Name
    ShowToolbarIcons False
End Sub

Private Sub Form_Initialize()

    ' Get the Default MasterFormat
    ' Fill the MasterFormat tree
    m_blnMasterFormatNotSpecified = True
    m_intMasterFormat = g_intMasterFormat
    FormatTree.ShowMasterFormatRoot = True
    FormatTree.ClearTree
    If m_intMasterFormat = EXT_MASTERFORMAT_VERSION Then
        FormatTree.InitData g_cnShared, "UNITCOST" & Right(EXT_MASTERFORMAT_VERSION, 2)
    Else
        FormatTree.InitData g_cnShared, "UNITCOST"
    End If

    ' Initialize grid
    m_objGridMap.SetGrid TDBGrid
    m_objGridMap.InitGrid
    m_blnFirstSearch = True
    m_blnJumpIn = False
    
    
    
End Sub

Private Sub Form_Load()
    Dim blnReturn As Boolean
    Dim strSelect As String
    Dim rec As ADODB.RecordSet
    
    Move START_LEFT, START_TOP, START_WIDTH, START_HEIGHT
'(rlh) 02/10/2009
'    g_objDAL.GetRecordset CONNECT, "select country_code from country order by country_code", rec
'    While Not rec.EOF
'        CountryList.AddItem (rec.Fields("country_code").Value)
'        If rec.Fields("country_code").Value = "USA" Then
'            CountryList.Selected(CountryList.ListCount - 1) = True
'        End If
'        rec.MoveNext
'    Wend
'    rec.Close
'    g_objDAL.GetRecordset CONNECT, "select region_code from region order by region_code", rec
'    While Not rec.EOF
'        RegionList.AddItem (rec.Fields("region_code").Value)
'        If rec.Fields("region_code").Value = "NAT" Then
'            RegionList.Selected(RegionList.ListCount - 1) = True
'        End If
'        rec.MoveNext
'    Wend
'    rec.Close
'
    
    LoadMasterFormatCombo Me.cboMasterFormat
    
    ' This will never return any rows, just used to create recordset
    StartUnitCostID.Text = "~"
    cmdSearch_Click
    StartUnitCostID.Text = ""
    
    
    cmdSearchExtID.Enabled = False      '(rlh) 02/10/2009
    
    If MF95_ENABLED = False Then    'rlh 3/24/2009
        Me.cmdSearchExtID.Enabled = False  'rlh 03/24/2009  MF95 Disable
    End If
    
End Sub

' Called when coming here from another screen
Public Sub JumpIn(strUnitCostId As String)
'    StartUnitCostID.Text = Compress_String(strUnitCostId)
    StartUnitCostID.Text = Trim(strUnitCostId) 'CR 919 4-30-01 ep: Removed line above and added Trim function to strip out leading spaces.
    If m_blnMasterFormatNotSpecified Then
        
        'CCD 8.4 rlh - commented out masterformat (re)setting.  it was causing problems!
        
        ' MF was never explicitly set, so default to 1995 for compatibility purposes
        'MasterFormat = UCD_MASTERFORMAT_VERSION
       
    End If
    cmdSearch_Click
End Sub

Private Sub Form_LostFocus()
    TDBGrid.Update
    HideGridSort
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbNormal Or Me.WindowState = vbMaximized Then
        If Me.Width >= 11715 Then
            TDBGrid.Width = Me.Width - (TDBGrid.Left * 3)
            Line2.X2 = Me.Width - 210
        Else
            Me.Width = 11715
        End If
        
        If Me.Height >= 7130 Then
            TDBGrid.Height = Me.Height - 4545
            Frame1.Top = Me.Height - 1260
            cmdUpdate.Top = Me.Height - 1020
            cmdNew.Top = Me.Height - 1020
            cmdClone.Top = Me.Height - 1020
            cmdDelete.Top = Me.Height - 1020
        Else
            Me.Height = 7130
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
    Dim sTableName As String
    
    On Error Resume Next
    ' Synch text box with tree
    If Len(strID) = 12 Then
        StartUnitCostID.Text = strID + "*"
        EndUnitCostID.Text = ""
        altunitcostid.Text = ""
    Else
        rs.Close ' Make sure it is closed
        sTableName = FormatTree.TableName
        strSelect = "select unit_cost_id_start, unit_cost_id_end from " & sTableName & " where hier_id='" + strID + "'"
        blnReturn = g_objDAL.GetRecordset(vbNullString, strSelect, rs)
        If Not rs.EOF Then
            StartUnitCostID.Text = rs.Fields("unit_cost_id_start")
            EndUnitCostID.Text = rs.Fields("unit_cost_id_end")
        Else
            StartUnitCostID.Text = ""
            EndUnitCostID.Text = ""
        End If
        rs.Close
    End If
    ' Clear other boxes
    Description.Text = ""
    ' Kick-off search
    cmdSearch_Click
    
End Sub

Private Sub cmdSearch_Click()
    On Error Resume Next
    Dim blnReturn As Boolean
    Dim strSelect As String
    Dim dtmToday As Date
    Dim dtmStart As Date
    Dim strStartUnitCostSrch As String
    Dim iMasterFormat As Long
    
    If m_objGridMap.IsPendingChange = True Then
        Dim Button
        Button = MsgBox("Do you want to save your changes?", vbYesNoCancel)
        If Button = vbYes Then
            cmdUpdate_Click
            ' If there were errors, cancel the search
            If m_blnWereErrors Then
                Exit Sub
            End If
        ElseIf Button = vbCancel Then
            ' Cancel the search
            TDBGrid.ReBind
            Exit Sub
        Else
            TDBGrid.DataChanged = False
        End If
    End If
    Screen.MousePointer = vbHourglass
    dtmToday = Date
    DoEvents
    
    ' Sync tree with text box
'    If Not StartUnitCostID.Text = "" Then
'        FormatTree.FocusItem (UnitCostID.Text)
'    End If
    
    If Len(StartUnitCostID.Text) = 0 And Len(altunitcostid.Text) = 0 And Len(Description.Text) = 0 Then
        Screen.MousePointer = vbNormal
        MsgBox "You must enter a Unit Cost ID or Description.", vbInformation
        StartUnitCostID.SetFocus
        Exit Sub
    End If
    If Len(StartUnitCostID) = 12 And InStr(1, StartUnitCostID, "*") = 0 And Len(EndUnitCostID) = 0 Then
        strStartUnitCostSrch = Compress_String(StartUnitCostID) + "*"
    Else
        strStartUnitCostSrch = Compress_String(StartUnitCostID)
    End If
    
    lblRowCount.Caption = "Working..."
    lblRowCount.Refresh

    m_rec.Close ' Make sure it is closed
    m_rec.MaxRecords = MAX_RECORDS ' Set the maximum number to bring back
    dtmStart = Now

    Dim strError As String
    iMasterFormat = cboMasterFormat.ItemData(cboMasterFormat.ListIndex)
    If iMasterFormat = ALT_MASTERFORMAT_VERSION Then
        strSelect = "exec usp_select_unit_cost_ext_rlh2 @start_unit_cost_id = '', @end_unit_cost_id = '', @alt_unit_cost_id = '" + SQLChangeWildcard(strStartUnitCostSrch) + "', @tech_desc = '" + SQLChangeWildcard(Description.Text) + "'"
        'strSELECT = "exec usp_select_unit_cost_ext_rlhDaveDrain @start_unit_cost_id = '', @end_unit_cost_id = '', @alt_unit_cost_id = '" + SQLChangeWildcard(strStartUnitCostSrch) + "', @tech_desc = '" + SQLChangeWildcard(Description.Text) + "'"
    Else
        strSelect = "exec usp_select_unit_cost_ext_rlh2 @start_unit_cost_id = '" + SQLChangeWildcard(strStartUnitCostSrch) + "', @end_unit_cost_id = '" + Compress_String(EndUnitCostID.Text) + "', @alt_unit_cost_id = '', @tech_desc = '" + SQLChangeWildcard(Description.Text) + "', @master_format=" & iMasterFormat
        'strSELECT = "exec usp_select_unit_cost_ext_rlhDaveDrain @start_unit_cost_id = '" + SQLChangeWildcard(strStartUnitCostSrch) + "', @end_unit_cost_id = '" + Compress_String(EndUnitCostID.Text) + "', @alt_unit_cost_id = '', @tech_desc = '" + SQLChangeWildcard(Description.Text) + "', @master_format=" & iMasterFormat
    End If
    ' Use DAL to perform select
    blnReturn = g_objDAL.GetRecordset(vbNullString, strSelect, m_rec)
    If blnReturn = False Then
        '8/16/2005 RTD - Added new DAL property LastErrorDescription
        MsgBox "An error occurred while searching:" & vbCrLf & g_objDAL.LastErrorDescription, vbCritical
        lblRowCount.Caption = "0 rows returned."
        Screen.MousePointer = vbNormal
        Exit Sub
    End If
    Set rsUnitCostClone = m_rec.Clone

    ' Pass recordset to handler class
    m_objGridMap.RecordSet = m_rec
    
    If m_rec.RecordCount > 0 Then
        lblRowCount.Caption = str(m_rec.RecordCount) + " rows returned in " + str(DateDiff("s", dtmStart, Now)) + " seconds"
        cmdOutput.Enabled = True
        cmdSearchExtID.Enabled = True
        cmdBookPreview.Enabled = True
    Else
        lblRowCount.Caption = "0 rows returned."
        cmdOutput.Enabled = False
        cmdSearchExtID.Enabled = False
        cmdBookPreview.Enabled = False
    End If
    
    ' Set MasterFormat - indicating the resulting search's format
    m_objGridMap.MasterFormat = iMasterFormat
    m_intMasterFormat = iMasterFormat
    TDBGrid.Caption = "MasterFormat " & iMasterFormat
    If iMasterFormat = EXT_MASTERFORMAT_VERSION Then
        cmdSearchExtID.Caption = "MF-" & Right(UCD_MASTERFORMAT_VERSION, 2) & " Record"
    Else
        cmdSearchExtID.Caption = "MF-" & Right(EXT_MASTERFORMAT_VERSION, 2) & " Record"
    End If
    
    ' If the upper bound was hit, inform user
    If m_rec.RecordCount = MAX_RECORDS And m_rec.State = adStateOpen Then
        MsgBox "The search returned the maximum number of records allowed. More records may be available."
    End If
    m_objGridMap.SetMenuBar
    
    ' Reset the grid contents
    TDBGrid.Bookmark = Null
    TDBGrid.ReBind
    TDBGrid.ApproxCount = m_rec.RecordCount
    SetButtons USEBOOKMARK
    Screen.MousePointer = vbNormal
    
    If MF95_ENABLED = False Then                'rlh 3/24/2009
        Me.cmdSearchExtID.Enabled = False       'rlh 03/24/2009  MF95 Disable
    End If
 
    Me.SetFocus                                 'rlh 4/14/2009  CCD 8.4 ksr

End Sub

Public Sub DoOutput()
    Dim sKey As String
    Dim frm As Form
    Dim blnVisible As Boolean
    Dim blnRefresh As Boolean
    Dim strUpdate As String
    Dim rec As ADODB.RecordSet
    Dim strError As String
    Dim blnReturn As Boolean
    Dim varBookmark As Variant
    Dim strUpdate1 As String
    Dim strSelect As String
    Dim oOUFormat As OUTPUT_USAGE_FORMAT

    If FormOpen("dlgOutput", frm, blnVisible) = True Then
        If blnVisible = False Then
            frm.Visible = True
            blnRefresh = True
        Else
            frm.Visible = False
            blnRefresh = False
        End If
    Else
        If Not IsNull(TDBGrid.Bookmark) Then
            blnRefresh = True
            Set frm = New dlgOutput
        End If
    End If

    If Not (TDBGrid.BOF = True Or TDBGrid.EOF = True) Then
        strUpdate = "exec sp_temp_output_init"
        blnReturn = frm.m_objOutput.ExecQuery(vbNullString, strUpdate, strError)

        strUpdate1 = "exec sp_temp_add_output_keys @skey_type = 'U', @skey = "
        If TDBGrid.SelBookmarks.Count = 0 Then  'No rows selected
            If Not IsNull(TDBGrid.Bookmark) Then    'Use current row
                m_rec.Bookmark = TDBGrid.Bookmark
                strUpdate = strUpdate1 + CStr(m_rec.Fields("unit_cost_skey"))
                blnReturn = frm.m_objOutput.ExecQuery(vbNullString, strUpdate, strError)
            End If
        Else
            For Each varBookmark In TDBGrid.SelBookmarks
                m_rec.Bookmark = varBookmark
                strUpdate = strUpdate1 + CStr(m_rec.Fields("unit_cost_skey"))
                blnReturn = frm.m_objOutput.ExecQuery(vbNullString, strUpdate, strError)
            Next varBookmark
        End If
        
        '8/15/2005 RTD
        'SET ALLOWABLE OUTPUT USAGE MASTERFORMATS
        oOUFormat = OUTPUT_BOTH
        If Trim(m_rec.Fields("ext_unit_cost_id") & "") = "" Then
            'IF XUCID IS EMPTY, THEN ONLY 1 FORMAT IS ALLOWED
            If MasterFormat = EXT_MASTERFORMAT_VERSION Then
                oOUFormat = OUTPUT_MF2004_ONLY
            Else
                oOUFormat = OUTPUT_MF1995_ONLY
            End If
        End If
        frm.OutputUsageFormat = oOUFormat
        
        frm.FillData
        frm.Show vbModeless, fMainForm
        frm.Caption = "Output Usage"
    End If

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

Private Sub StartUnitCostID_Change()
    Dim blnReturn As Boolean
    
    If InStr(1, StartUnitCostID.Text, "*") > 0 Then
        blnReturn = LockField(Me, "EndUnitCostID")
    Else
        If Me.cboMasterFormat.Text <> "MF-1988" Then
            blnReturn = UnLockField(Me, "EndUnitCostID")
        End If
    End If

End Sub

Private Sub StartUnitCostID_LostFocus()
    StartUnitCostID = Trim(StartUnitCostID)
End Sub

Private Sub TDBGrid_DblClick()
    ' Signal that double-click has occurred
    m_blnDoubleClick = True
End Sub

Private Sub TDBGrid_FetchCellTips(ByVal SplitIndex As Integer, ByVal ColIndex As Integer, ByVal RowIndex As Long, CellTip As String, ByVal FullyDisplayed As Boolean, ByVal TipStyle As TrueOleDBGrid80.StyleDisp)
'SHOW CELLTIP FOR DESCRIPTION COLUMNS
    
    If ColIndex >= 0 And ColIndex <= TDBGrid.Columns.Count Then
        If Right(TDBGrid.Columns(ColIndex).DataField, 4) <> "desc" Then
            CellTip = ""
        End If
    Else
        CellTip = ""
    End If
    
End Sub

Private Sub TDBGrid_GotFocus()
    TDBGrid.TabStop = True
End Sub

Private Sub TDBGrid_KeyPress(KeyAscii As Integer)
    m_sngYCoord = 0
End Sub

Private Sub TDBGrid_KeyUp(KeyCode As Integer, Shift As Integer)
    m_sngYCoord = 0
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
            ' Same function as clicking Unit Cost button, open single record view
            cmdUnitCost_Click
        End If
    Else
        If Button = vbRightButton And IsNumeric(TDBGrid.Bookmark) Then
            Dim strErrorMsg As String
            strErrorMsg = m_objGridMap.GetError(TDBGrid.Bookmark)
            If Len(strErrorMsg) > 0 Then
                MsgBox strErrorMsg
            End If
        Else
            If TDBGrid.RowContaining(Y) <> TDBGrid.Row Then
                m_sngYCoord = Y
            End If
        End If
    End If

End Sub

Private Sub Form_Activate()
    'TDBGrid.ReBind
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
        OutputView True
        ShowGridSort
        m_objGridMap.SetMenuBar
    End If
    ShowToolbarIcons True
    
End Sub

Private Sub TDBGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim lngCurrentRow As Long
    If m_sngYCoord > 0 Then
        lngCurrentRow = TDBGrid.RowContaining(m_sngYCoord)
        If lngCurrentRow <> -1 Then
            If IsNumeric(LastRow) - 1 <> lngCurrentRow Then
            TDBGrid.Row = lngCurrentRow
                m_rec.Bookmark = TDBGrid.Bookmark
                position_output
                SetButtons USECOORD, m_sngYCoord
            End If
        End If
    Else
        If IsNumeric(LastRow) Then
            If CLng(TDBGrid.Row) <> CLng(LastRow) - 1 Then
                position_output
                SetButtons USEBOOKMARK
            End If
        Else    'no last row, must have changed
                position_output
                SetButtons USEBOOKMARK
        End If
    End If
    m_sngYCoord = 0
End Sub

Private Sub TDBGrid_SelChange(Cancel As Integer)
    TDBGrid.SetFocus
End Sub

Private Sub ShowToolbarIcons(bShowIcons As Boolean)
    
    On Error GoTo Err_Handler
    With fMainForm
        .tbToolBar.Buttons.Item(tbrPRINT).Enabled = bShowIcons
        .tbToolBar.Buttons.Item(tbrPRINT).Visible = bShowIcons
        .tbToolBar.Buttons.Item(tbrPREVIEW).Enabled = bShowIcons
        .tbToolBar.Buttons.Item(tbrPREVIEW).Visible = bShowIcons
        .tbToolBar.Buttons.Item(tbrEXPORT).Enabled = False
        .tbToolBar.Buttons.Item(tbrEXPORT).Visible = False
        .mnuFilePageSetup.Enabled = bShowIcons
        .mnuFilePrint.Enabled = bShowIcons
        .mnuFileSaveAs.Enabled = False
        .mnuFilePrintPreview.Enabled = bShowIcons
    End With
    Exit Sub

Err_Handler:
    Exit Sub
    
End Sub

Public Sub PreviewReport()
    Dim sFieldName As String
    
    If m_rec.RecordCount > 0 Then
        BookFormatPrintPreviewRS m_rec, MasterFormat
    End If
    
End Sub

Public Sub PrintReport()
    PreviewReport
End Sub

Private Sub MasterFormatChanged()
'A NEW MASTERFORMAT WAS SELECTED FROM THE DROP-DOWN BOX
'ADDED 6/20/2005 RTD FOR VERSION 7.4.0+
    Dim sTreeType As String
    
    If cboMasterFormat.ListIndex < 0 Then
        Exit Sub
    End If
    
    Select Case cboMasterFormat.ItemData(cboMasterFormat.ListIndex)
    Case EXT_MASTERFORMAT_VERSION
        UnLockField Me, "EndUnitCostID"
        lblUnitCostId.Caption = "Unit Cost ID " & Right(EXT_MASTERFORMAT_VERSION, 2) & ":"
        sTreeType = "UNITCOST" & Right(EXT_MASTERFORMAT_VERSION, 2)
    Case UCD_MASTERFORMAT_VERSION
        UnLockField Me, "EndUnitCostID"
        lblUnitCostId.Caption = "Unit Cost ID " & Right(UCD_MASTERFORMAT_VERSION, 2) & ":"
        sTreeType = "UNITCOST"
    Case ALT_MASTERFORMAT_VERSION
        LockField Me, "EndUnitCostID"
        'EndUnitCostID.Text = ""
        lblUnitCostId.Caption = "Alt Unit Cost ID:"
        sTreeType = "UNITCOST"
    Case Else
        UnLockField Me, "EndUnitCostID"
        lblUnitCostId.Caption = "Unit Cost ID " & Right(UCD_MASTERFORMAT_VERSION, 2) & ":"
        sTreeType = "UNITCOST"
    End Select
    
    'CHECK IF WE NEED TO RE-INITIALIZE TREE
    If FormatTree.TreeType <> sTreeType Then
        Screen.MousePointer = vbHourglass
        FormatTree.DisableRedraw = True
        FormatTree.ClearTree
        FormatTree.InitData g_cnShared, sTreeType
        FormatTree.DisableRedraw = False
        Screen.MousePointer = vbDefault
    End If

    On Error Resume Next
    StartUnitCostID.SetFocus
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
