VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmTradeGroupGrid 
   Caption         =   "Trade Group Grid"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11190
   Icon            =   "frmTradeGroupGrid.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6855
   ScaleWidth      =   11190
   Begin VB.CommandButton cmdRemap 
      Caption         =   "Remap Groups"
      Height          =   495
      Left            =   10080
      TabIndex        =   23
      Top             =   6240
      Width           =   1035
   End
   Begin VB.CheckBox chkIncludeHistory 
      Caption         =   "Include History"
      Height          =   375
      Left            =   9360
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdRmvGroup 
      Caption         =   "Remo&ve Group Mbr"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7365
      TabIndex        =   12
      Top             =   6240
      Width           =   1150
   End
   Begin VB.CommandButton cmdAddMbr 
      Caption         =   " &Add Group          Mbr"
      Height          =   495
      Left            =   4590
      TabIndex        =   10
      Top             =   6240
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   "Go To"
      Height          =   855
      Left            =   240
      TabIndex        =   22
      Top             =   6000
      Width           =   1455
      Begin VB.CommandButton cmdLaborRate 
         Caption         =   "Labor &Rate"
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1150
      End
   End
   Begin VB.ComboBox cmbCity 
      Height          =   315
      Left            =   7815
      TabIndex        =   4
      Top             =   1680
      Width           =   1935
   End
   Begin VB.ComboBox cmbTradeGrpCode 
      Height          =   315
      Left            =   7800
      TabIndex        =   0
      Top             =   525
      Width           =   1485
   End
   Begin VB.ComboBox cmbTradeID 
      Height          =   315
      Left            =   7815
      TabIndex        =   2
      Top             =   945
      Width           =   1470
   End
   Begin VB.ComboBox cmbState 
      Height          =   315
      Left            =   7815
      TabIndex        =   3
      Top             =   1320
      Width           =   765
   End
   Begin VB.CommandButton cmdRenameGrp 
      Caption         =   "Rena&me Group"
      Height          =   495
      Left            =   8880
      TabIndex        =   13
      Top             =   6240
      Width           =   1035
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete Group"
      Height          =   495
      Left            =   5985
      TabIndex        =   11
      Top             =   6240
      Width           =   1035
   End
   Begin ConstructionCostDatabase.DynaTree FormatTree 
      Height          =   2775
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   4895
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New Group"
      Height          =   495
      Left            =   3195
      TabIndex        =   9
      Top             =   6240
      Width           =   1035
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   495
      Left            =   1800
      TabIndex        =   8
      Top             =   6240
      Width           =   1035
   End
   Begin VB.CheckBox ckbRowWrap 
      Caption         =   "Row Wrap"
      Height          =   315
      Left            =   60
      TabIndex        =   6
      Top             =   2880
      Width           =   1215
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
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(0)._MinWidth=1216"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(1)._MinWidth=-3"
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
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   480
      Left            =   8520
      TabIndex        =   5
      Top             =   2160
      Width           =   1150
   End
   Begin VB.Line Line1 
      X1              =   6660
      X2              =   6660
      Y1              =   2700
      Y2              =   60
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "State:"
      Height          =   255
      Left            =   7110
      TabIndex        =   21
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label lblRowCount 
      Caption         =   "0 rows returned"
      Height          =   255
      Left            =   5160
      TabIndex        =   20
      Top             =   2880
      Width           =   3255
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
      Height          =   315
      Left            =   6780
      TabIndex        =   19
      Top             =   60
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "City:"
      Height          =   255
      Left            =   7350
      TabIndex        =   18
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Trade Group:"
      Height          =   255
      Left            =   6720
      TabIndex        =   17
      Top             =   600
      Width           =   1035
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Trade ID:"
      Height          =   255
      Left            =   6990
      TabIndex        =   16
      Top             =   990
      Width           =   735
   End
   Begin VB.Line Line2 
      X1              =   60
      X2              =   11040
      Y1              =   2820
      Y2              =   2820
   End
End
Attribute VB_Name = "frmTradeGroupGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''' <modulename> frmTradeGroupGrid</modulename>
''' <functionname>General (Main) </functionname>
'''
''' <summary>
''' Provides u/i permitting user to do the following:
'''
'''Research Labor Rate "Groups" with the context of:
'''"   Trade Group (selected group of trade_ids by city and state)
'''"   Trade ID    (e.g. carpenters (CARP), plumbers (PLUMB) etc.)
'''"   City
'''"   State
'''
'''(Major function buttons)
'''
'''1.  "Labor Rate" - Display form         (frmLaborRate.frm)
'''2.  "Update" (save) Labor Rate change(s)    (CLaborRateMap.Update() )
'''3.  "New Group" - Build a new trade group   (frmLaborRate.frm)
'''4.  "Add Group Mbr"             (LoadTradeGroups() )
'''5.  "Delete  Group"             (frmLaborRate.frm)
'''6.  "Remove Group Mbr"                   (m_objGridMap.RemoveTradeGroupMbr() )
'''7.  "Rename Group"              (frmTradeGroup)
'''8.  "Remap Groups"              (frmTradeGroupRemap)
'''
'''
'''HELPER Class: CTradeGroupMap.Cls
''' </summary>
'''
''' <seealso>frmLaborRate</seealso>
'''<seealso>frmTradeGroup</seealso>
'''
''' <datastruct>m_rec</datastruct>
'''<datastruct>m_objGridMap</datastruct>
'''
''' <storedprocedurename> sp_replace_trade_groups</storedprocedurename>
''' <storedprocedurename> sp_LaborRatesMaxStart</storedprocedurename>
'''
''' <returns>N/A</returns>
''' <exception>Always trap with an accompanying message box</exception>
''' <example>
''' <code>
'''exec sp_LaborRatesMaxStart @trade_id='BOIL', @trade_group_code='', @city='', @state='AK', @start_date=' ', @term_date=' ', @includehistory = 1, @maxrowcount = 5000
'''</code>
''' <code>
'''exec sp_replace_trade_groups  @trade_group_code = 'BOIL001', @new_trade_group_code = '', @last_update_person='Hancockrl'
'''</code>
'''<code>
'''NOTE:  This is how the menu tree gets initialized for "Labor" menu activity/nodes
'''
'''Private Sub Form_Initialize()
'''    Status ("Loading Trade Groups...")
'''    Screen.MousePointer = vbHourglass
'''    ' Fill the MasterFormat tree
'''    FormatTree.InitData g_cnShared, "LABOR"    '<<<<<<<<<<<<<<<<<<<<< Here!
'''    ' Initialize grid
'''    m_objGridMap.SetGrid TDBGrid
'''    m_objGridMap.InitGrid
'''    m_objGridMap.Warnings = True
'''    DoEvents    'Paint screen
'''    LoadCombos
'''    m_blnFirstSearch = True
'''    m_blnJumpIn = False
'''    Screen.MousePointer = vbNormal
'''End Sub
'''
'''</code>
'''</example>
'''<permission>Public</Permission>
'''<dependson>This component depends on the following
'''1.  CTradeGroupMap.cls
'''2.  CGridMap.cls
'''3.  CCDdal.CRSMDataAccess (
'''4.  Access to the DAL (data access layer dll) opened in MainModule_Main() )
'''</dependson>


Dim m_objGridMap As New CTradeGroupMap ' Class to handle grid
Dim m_blnFirstSearch As Boolean ' Is this the first search we have made on this screen
Dim m_blnJumpIn As Boolean ' Are we jumping here from another screen
Dim m_rec As New ADODB.RecordSet ' Recordset to hold query results
Dim m_blnDoubleClick As Boolean ' Did a double click just occurr
Dim m_blnWereErrors As Boolean ' True if the Update had errors, used in QueryUnload
Dim rsLaborClone As RecordSet
Dim varBookmark As Variant


Const TRADE_Group_TABLE = "Trade_Group"
Const LABOR_TRADE_TABLE = "Labor_Trade"
Const LOCATION_TABLE = "Location"
Const LABORTRADE_TABLE = "Labor_Trade"
Const USEBOOKMARK = 1
Const USECOORD = 0

Dim m_State As String
Dim m_strCurrentFormControl As String
Private Sub BldUnionSQL(iRecords As Integer, m_rec As RecordSet)
' This procedure will create a recordset from the Labor Rate and Location tables
'   All grid fields will be included in the recordset
'   The key of Surrogate Trade Key (trade_skey, Location ID (loc_id), and termination date (term_date) will determine
'   the group of records, with the last lRecord (number of records) based on Termination Date (descending).
'   Each iteration of iRecords will be a separate subquery to be joined in the union query
'   Each iteration will contain a subquery with the previous date used for comparison for the current max date



End Sub

Private Sub LoadCities(Optional strCity As String)
Dim strSelect As String
Dim rsTemp As RecordSet
Dim blnReturn As Boolean
'Load Cities
    
cmbCity.Clear
If cmbState.Text > "" Then
    strSelect = "select distinct city from location where location.state_code = '" + cmbState.Text + "'  order by city"
Else
    strSelect = "select distinct city from location order by city"
End If
    ' Use DAL to perform select
    blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
    If blnReturn = False Then
        MsgBox "An error occurred loading Cities."
    Else
        If Not (rsTemp.EOF And rsTemp.BOF) Then
            Do Until rsTemp.EOF
                If strCity > "" Then
                    If strCity = rsTemp![City] Then
                        cmbCity.Text = ConvertCase(rsTemp![City])
                    End If
                End If
                cmbCity.AddItem ConvertCase(rsTemp![City])
                rsTemp.MoveNext
            Loop
        End If
    End If
    rsTemp.Close

End Sub
Private Function ConvertCase(strText As String) As String
Dim strTemp As String
Dim strTemp2 As String
Dim iStarta As Integer
Dim iStartb As Integer
strTemp = Left(strText, 1) + LCase(Right(strText, Len(strText) - 1))
iStarta = InStr(1, strText, " ")
If iStarta = 0 Then
    iStarta = InStr(1, strText, ",")
End If
If iStarta <> 0 Then
    While iStarta <> 0
        strTemp = Left(strTemp, Len(strTemp) - (Len(strTemp) - iStarta)) + UCase(Mid(strTemp, iStarta + 1, 1)) + Right(strTemp, Len(strTemp) - iStarta - 1)
        iStartb = InStr(iStarta + 1, strText, " ")
        If iStartb = 0 Then
            iStartb = InStr(iStarta, strText, ",")
        End If
        iStarta = iStartb
    Wend
    ConvertCase = strTemp
Else
    ConvertCase = strTemp
End If
End Function

Private Sub LoadCombos()
    Dim blnReturn As Boolean
    Dim strSelect As String
    Dim rsTemp As RecordSet

'Load All Selection Combos

'Load Trade IDs
    strSelect = "SELECT LABOR_TRADE.trade_id FROM LABOR_TRADE ORDER BY LABOR_TRADE.trade_id"
    
    ' Use DAL to perform select
    blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
    If blnReturn = False Then
        MsgBox "An error occurred loading Trade IDs."
        lblRowCount.Caption = "0 rows returned."
    Else
        If Not (rsTemp.EOF And rsTemp.BOF) Then
            Do Until rsTemp.EOF
                cmbTradeID.AddItem rsTemp![Trade_ID]
                rsTemp.MoveNext
            Loop
        End If
    End If
    rsTemp.Close

'Load Trade Groups
    strSelect = "SELECT TRADE_Group.trade_group_code FROM Trade_Group ORDER BY TRADE_Group.trade_group_code"
    strSelect = "select distinct trade_group_code from labor_rate order by trade_group_code"
    ' Use DAL to perform select
    blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
    If blnReturn = False Then
        MsgBox "An error occurred loading Trade Groups."
    Else
        If Not (rsTemp.EOF And rsTemp.BOF) Then
            Do Until rsTemp.EOF
                cmbTradeGrpCode.AddItem rsTemp![Trade_Group_Code]
                rsTemp.MoveNext
            Loop
        End If
    End If
    rsTemp.Close

    LoadCities

'Load States
    strSelect = "select distinct state_code from location order by state_code;"

    ' Use DAL to perform select
    blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
    If blnReturn = False Then
        MsgBox "An error occurred loading States."
    Else
        If Not (rsTemp.EOF And rsTemp.BOF) Then
            Do Until rsTemp.EOF
                cmbState.AddItem rsTemp![State_Code]
                rsTemp.MoveNext
            Loop
        End If
    End If
    rsTemp.Close

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

Private Sub cmbState_GotFocus()
m_State = cmbState.Text
End Sub

Private Sub cmbState_LostFocus()
If m_State <> cmbState.Text Then
    LoadCities
End If
End Sub

Private Sub cmdAddMbr_Click()
    Dim strUpdate As String
    Dim blnRet As Boolean
    Dim strError As String
    Dim rsTemp As ADODB.RecordSet
    Dim varButton
'Rename the trade group code in all labor rate records that it is in.
    Dim frm As frmTradeGroup
    Dim rec As ADODB.RecordSet
    
    On Error Resume Next

    
    Screen.MousePointer = vbHourglass
    
    Set frm = New frmTradeGroup
    frm.Caption = "Add Trade Group Member"
    frm.blnAddMbr = True
    frm.show_detail
    If Len(Trim(TDBGrid.Columns("Trade Group Code").CellText(TDBGrid.Bookmark))) > 0 Then
        frm.Trade_Group_Code = TDBGrid.Columns("Trade Group Code").CellText(TDBGrid.Bookmark)
        frm.start_date = TDBGrid.Columns("MaxStartDate").CellText(TDBGrid.Bookmark)
    End If
    frm.LoadTradeGroups
    'frm.Load_Trade_IDs         'AK- 6/7/2006 - Removed as it is called twice
    frm.get_counts
    frm.Show
    
    Screen.MousePointer = vbNormal
    Status ("")
End Sub
Private Sub cmdDelete_Click()
    Dim strSelect As String
    Dim strUpdate As String
    Dim blnRet As Boolean
    Dim strError As String
    Dim rsTemp As ADODB.RecordSet
    Dim varButton
'Remove the trade group code from the Trade_Group table and all labor rate records.

    On Error Resume Next
    If TDBGrid.Columns("trade_group_code").Value > " " Then
        strSelect = "select count(*) as RcdsToDelete from labor_rate where trade_group_code='" + TDBGrid.Columns("trade_group_code").Value + "'"
        blnRet = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
        If Not blnRet Then
            MsgBox "An error occurred retrieving data."
        Else
            If rsTemp![RcdsToDelete] > 0 Then
                Dim strMsg As String
                strMsg = CStr(rsTemp![RcdsToDelete]) + " Labor rate records will be removed/detached from the trade group " + TDBGrid.Columns("trade_group_code").Value + ".  Are you sure you want to remove them?"
                varButton = MsgBox(strMsg, vbYesNo + vbCritical)
            End If
        End If
        rsTemp.Close
        Set rsTemp = Nothing

     If varButton = vbYes Then
        Screen.MousePointer = vbHourglass
         strUpdate = "exec sp_replace_trade_groups "
         strUpdate = strUpdate + " @trade_group_code = '" + TDBGrid.Columns("trade_group_code").Value
         strUpdate = strUpdate + "', @new_trade_group_code = ''"
         strUpdate = strUpdate + ", @last_update_person='" + strUserName + "'"
         blnRet = g_objDAL.ExecQuery(CONNECT, strUpdate, strError)
        If Len(strError) > 0 Then
        'If Not blnRet Then
             MsgBox strError
             m_blnWereErrors = True
         Else
         
        'Load Trade Groups
            cmbTradeGrpCode.Clear
            strSelect = "select distinct trade_group_code from labor_rate order by trade_group_code"
            blnRet = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
            If blnRet = False Then
                MsgBox "An error occurred loading Trade Groups."
            Else
                If Not (rsTemp.EOF And rsTemp.BOF) Then
                    Do Until rsTemp.EOF
                        cmbTradeGrpCode.AddItem rsTemp![Trade_Group_Code]
                        rsTemp.MoveNext
                    Loop
                End If
            End If
            rsTemp.Close
            m_rec.Requery
            cmbTradeGrpCode.Text = "~"
            cmdSearch_Click
            cmbTradeGrpCode.Text = ""
            Screen.MousePointer = vbNormal
         End If
    
     End If
    End If

End Sub
Private Sub cmdLaborRate_Click()
    If IsNumeric(TDBGrid.Bookmark) = False Then
        MsgBox "You must select a row."
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    ' Navigate to single-record view
    Dim frm As frmLaborRateGrid
    Dim rec As ADODB.RecordSet
    Set frm = New frmLaborRateGrid

    frm.JumpIn TDBGrid.Columns("Trade ID").CellText(TDBGrid.Bookmark), TDBGrid.Columns("State").CellText(TDBGrid.Bookmark), TDBGrid.Columns("City").CellText(TDBGrid.Bookmark), TDBGrid.Columns("Trade Group Code").CellText(TDBGrid.Bookmark)
    Screen.MousePointer = vbNormal
End Sub

Private Sub cmdNew_Click()
    Dim strUpdate As String
    Dim blnRet As Boolean
    Dim strError As String
    Dim rsTemp As ADODB.RecordSet
    Dim varButton
    Dim frm As frmTradeGroup
    Dim rec As ADODB.RecordSet
    
    On Error Resume Next
    
    Screen.MousePointer = vbHourglass
    DoEvents

    Set frm = New frmTradeGroup
    frm.Visible = False
    DoEvents
    frm.Caption = "New Trade Group"
    frm.blnNewGroup = True
    frm.show_detail
    frm.LoadTradeGroups
    frm.Load_Trade_IDs
    frm.Show
    
  Screen.MousePointer = vbNormal
Status ("")

End Sub

Public Sub Sort(intDir As Integer)
    m_objGridMap.Sort intDir
End Sub

Private Sub cmdRefresh_Click()
cmdSearch_Click
End Sub

Private Sub cmdRemap_Click()
 Dim strUpdate As String
    Dim blnRet As Boolean
    Dim strError As String
    Dim rsTemp As ADODB.RecordSet
    Dim varButton
'Rename the trade group code in all labor rate records that it is in.
    Dim frm As frmTradeGroupRemap
    Dim rec As ADODB.RecordSet
    
    On Error Resume Next
    
    Screen.MousePointer = vbHourglass
    
    Set frm = New frmTradeGroupRemap   'rlh CCD 8.4 +
    frm.Trade_Group_Code = TDBGrid.Columns("Trade Group Code").Text
    frm.LoadTradeGroups
    frm.show_detail
    
    frm.Caption = "Remap Trade Groups"
    frm.Show
    
    Screen.MousePointer = vbNormal
End Sub

Private Sub cmdRenameGrp_Click()
    Dim strUpdate As String
    Dim blnRet As Boolean
    Dim strError As String
    Dim rsTemp As ADODB.RecordSet
    Dim varButton
'Rename the trade group code in all labor rate records that it is in.
    Dim frm As frmTradeGroup
    Dim rec As ADODB.RecordSet
    
    On Error Resume Next
    
    Screen.MousePointer = vbHourglass
    
    Set frm = New frmTradeGroup
    'frm.Trade_Group_Code = TDBGrid.Columns("Trade Group Code").Text
    frm.Caption = "Rename Trade Group"
    frm.Show
    
    Screen.MousePointer = vbNormal

End Sub

Private Sub cmdRmvGroup_Click()

'Remove the group code for one trade id/location/start date
    On Error GoTo Error_Processing
    
If m_objGridMap.RemoveTradeGroupMbr Then
    Screen.MousePointer = vbHourglass
    cmdSearch_Click
End If
Exit_Sub:
    Screen.MousePointer = vbNormal
    Exit Sub

Error_Processing:
    MsgBox Error$
    Resume Exit_Sub
    Resume 0

End Sub

Private Sub cmdTradeGroup_Click()
    Dim strUpdate As String
    Dim blnRet As Boolean
    Dim strError As String
    Dim rsTemp As ADODB.RecordSet
    Dim varButton
'Rename the trade group code in all labor rate records that it is in.
    Dim frm As frmTradeGroup
    Dim rec As ADODB.RecordSet
    
    On Error Resume Next
    
    Screen.MousePointer = vbHourglass
    
    Set frm = New frmTradeGroup
        ' Make copy of recordset
    Set rec = m_rec.Clone
    ' Get the selected row from grid
    rec.Bookmark = TDBGrid.Bookmark
    frm.SetRow rec ' Pass the current record into the form
    frm.LoadTradeGroups
    frm.Caption = "Trade Group Update"
    frm.Show
    
    Screen.MousePointer = vbNormal

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
    cmdSearch_Click
Out:
End Sub

Private Sub Form_Deactivate()
m_strCurrentFormControl = Me.ActiveControl.Name
End Sub

Private Sub Form_GotFocus()
 frmTradeGroupGrid.TDBGrid.Refresh
End Sub

Private Sub Form_Initialize()
    Status ("Loading Trade Groups...")
    Screen.MousePointer = vbHourglass
    ' Fill the MasterFormat tree
    FormatTree.InitData g_cnShared, "LABOR"
    ' Initialize grid
    m_objGridMap.SetGrid TDBGrid
    m_objGridMap.InitGrid
    m_objGridMap.Warnings = True
    DoEvents    'Paint screen
    LoadCombos
    m_blnFirstSearch = True
    m_blnJumpIn = False
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Load()
    Dim blnReturn As Boolean
    Dim strSelect As String
    Move START_LEFT, START_TOP, START_WIDTH, START_HEIGHT

    cmbTradeGrpCode.Text = "0"
'    cmdSearch_Click
    cmbTradeGrpCode.Text = ""
    
    Status ("")
    SetButtons USEBOOKMARK
 
End Sub
' Called when coming here from another screen
Public Sub JumpIn(strTradeGrpCode As String, strTradeID As String, strState As String, strCity As String)
    Screen.MousePointer = vbHourglass
    cmbTradeGrpCode.Text = strTradeGrpCode
    cmbTradeID.Text = strTradeID
    cmbState.Text = strState
    cmbCity.Text = strCity
    cmdSearch_Click
End Sub

Private Sub Form_LostFocus()
TDBGrid.Update
HideGridSort
End Sub

Private Sub Form_Resize()
Dim i As Integer
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
            cmdRenameGrp.Top = Me.Height - 1020
            cmdRemap.Top = Me.Height - 1020     'rlh 03/25/2009  CCD 8.4+
            cmdDelete.Top = Me.Height - 1020
            cmdAddMbr.Top = Me.Height - 1020
            cmdRmvGroup.Top = Me.Height - 1020
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
'    ' Synch text box with tree
'     ' Clear other boxes
Dim strCity As String

cmbTradeGrpCode = ""

cmbTradeID = Left(strID, 4)

If Len(strID) > 6 Then
    strCity = Right(strID, Len(strID) - 6)
Else
    strCity = ""
End If
If Len(strID) >= 6 Then
    cmbState = Mid(strID, 5, 2)
Else
    cmbState = ""
End If
LoadCities strCity

'    ' Kick-off search
    cmdSearch_Click
End Sub

Public Sub cmdSearch_Click()
    On Error Resume Next
    Dim blnReturn As Boolean
    Dim strSelect As String
    Dim dtmToday As Date
    Dim dtmStart As Date
    Dim bSelectAnd As Boolean
    Dim sTmp As String
    Dim strID As String
    Dim rsTemp As ADODB.RecordSet
    Dim rsTemp2 As RecordSet
    Dim iIncludeHistory As Integer
    
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
        GoTo Exit_Sub
        Else
            TDBGrid.DataChanged = False
        End If
    End If
    Screen.MousePointer = vbHourglass
    dtmToday = Date

' Synch tree with text box - all 3, or first 2, or first
    strID = ""
    If cmbTradeID.Text <> "" And cmbState.Text <> "" And cmbCity.Text <> "" Then
        strID = cmbTradeID.Text + UCase(cmbState.Text) + UCase(cmbCity.Text)
    ElseIf cmbTradeID.Text <> "" And cmbState.Text <> "" Then
        strID = cmbTradeID.Text + cmbState.Text
    ElseIf cmbTradeID.Text <> "" Then
        strID = cmbTradeID.Text
    End If
    If strID > "" Then
        FormatTree.FocusItem (strID)
    End If

    If Len(cmbTradeID.Text) = 0 And Len(cmbTradeGrpCode.Text) = 0 And Len(cmbCity.Text) = 0 And Len(cmbState.Text) = 0 Then
        Screen.MousePointer = vbNormal
        MsgBox "You must enter search criteria before searching."
        GoTo Exit_Sub
    End If
'    If Len(cmbTradeGrpCode.Text) > 0 Then
'       iIncludeHistory = chkIncludeHistory
'    Else
'        iIncludeHistory = 1
'    End If

    lblRowCount.Caption = "Working..."
    lblRowCount.Refresh
  
    strSelect = "exec sp_LaborRatesMaxStart @trade_id='" + SQLChangeWildcard(cmbTradeID.Text) + "', @trade_group_code='"
    strSelect = strSelect + SQLChangeWildcard(cmbTradeGrpCode.Text) + "', @city='"
    strSelect = strSelect + SQLChangeWildcard(Trim(cmbCity.Text)) + "', @state='"
    strSelect = strSelect + SQLChangeWildcard(Trim(cmbState.Text)) + "', @start_date=''"
    strSelect = strSelect + ", @term_date=''"
    strSelect = strSelect + ", @includehistory = " + CStr(iIncludeHistory)
    strSelect = strSelect + ", @maxrowcount = " + CStr(MAX_RECORDS)
    
    m_rec.Close ' Make sure it is closed
    m_rec.MaxRecords = MAX_RECORDS ' Set the maximum number to bring back
    dtmStart = Now
    ' Use DAL to perform select
    blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, m_rec)
    If blnReturn = False Then
        MsgBox "An error occurred while searching."
        lblRowCount.Caption = "0 rows returned."
        GoTo Exit_Sub
    End If
    Set rsLaborClone = m_rec.Clone
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
    TDBGrid.FetchRowStyle = True
    Set rsLaborClone = m_rec.Clone
    SetButtons USEBOOKMARK
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

'*** APEX Migration Utility Code Change ***
'Private Sub TDBGrid_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal col As Integer, ByVal CellStyle As TrueOleDBGrid60.StyleDisp)
'*** APEX Migration Utility Code Change ***
'Private Sub TDBGrid_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal col As Integer, ByVal CellStyle As TrueOleDBGrid70.StyleDisp)
Private Sub TDBGrid_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueOleDBGrid80.StyleDisp)
'    rsLaborClone.Bookmark = Bookmark
'    If rsLaborClone.Fields("trade_group_code") > " " Then
'        CellStyle.Locked = True
'        CellStyle.Font.Bold = False
'    Else
'        CellStyle.Locked = False
'        CellStyle.Font.Bold = True
'    End If

End Sub

'*** APEX Migration Utility Code Change ***
'Private Sub TDBGrid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
'*** APEX Migration Utility Code Change ***
'Private Sub TDBGrid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
Private Sub TDBGrid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid80.StyleDisp)
'If the row is not the last row for the trade group/start date, lock it.
On Error Resume Next
rsLaborClone.Bookmark = Bookmark
'If m_GroupDate(Bookmark) <> rsLaborClone.Fields("start_date") And rsLaborClone.Fields("trade_group_code") > " " Then
If rsLaborClone.Fields("MaxStartDate") <> rsLaborClone.Fields("start_date") Then
    RowStyle.Locked = True
    RowStyle.ForeColor = vbGrayText
Else
    RowStyle = "ActiveRow"
End If
End Sub

Private Sub TDBGrid_GotFocus()
TDBGrid.TabStop = True
End Sub

Private Sub TDBGrid_KeyUp(KeyCode As Integer, Shift As Integer)
        SetButtons USEBOOKMARK

End Sub

Private Sub SetButtons(Mode As Single, Optional Coord As Variant)
On Error GoTo Exit_Sub

Select Case Mode
    Case USEBOOKMARK
        If Not IsNull(TDBGrid.Bookmark) Then
            rsLaborClone.Bookmark = TDBGrid.Bookmark
        End If
    Case USECOORD
        rsLaborClone.Bookmark = TDBGrid.RowBookmark(TDBGrid.RowContaining(Coord))
End Select

If Not IsNull(TDBGrid.Bookmark) Then
    If TDBGrid.Bookmark > 0 Then 'valid bookmark
        cmdLaborRate.Enabled = True
    Else
        cmdLaborRate.Enabled = False
    End If
    
    If rsLaborClone.Fields("MaxStartDate") <> rsLaborClone.Fields("start_date") Then
        cmdRmvGroup.Enabled = False
    Else
        If Trim(rsLaborClone.Fields("trade_group_code")) > "" Then
            cmdRmvGroup.Enabled = True
        Else
            cmdRmvGroup.Enabled = False
        End If
    End If
    If Trim(rsLaborClone.Fields("trade_group_code")) > "" Then
        cmdRenameGrp.Enabled = True
        cmdDelete.Enabled = True
    Else
        cmdRenameGrp.Enabled = False
        cmdDelete.Enabled = False
    End If
Else
    cmdRenameGrp.Enabled = False
    cmdDelete.Enabled = False
    cmdLaborRate.Enabled = False
    cmdRmvGroup.Enabled = False
End If

Exit_Sub:

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
        End If
    Else
        If Button = vbRightButton And IsNumeric(TDBGrid.Bookmark) Then
            Dim strErrorMsg As String
            strErrorMsg = m_objGridMap.GetError(TDBGrid.Bookmark)
            If Len(strErrorMsg) > 0 Then
            puzzling
                MsgBox strErrorMsg
            End If
        End If
        SetButtons USECOORD, Y
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
     Dim i As Integer
    '    TDBGrid.ReBind
        OutputView False
        For i = 0 To Forms.Count - 1
            If Forms(i).Name = "frmTradeGroup" Then
                If Forms(i).Visible = True Then
                    Forms(i).ZOrder
                    If Me.WindowState = vbNormal Then
                        Forms(i).WindowState = vbNormal
                    End If
                End If
                Exit For
            End If
        Next i
        ShowGridSort
        m_objGridMap.SetMenuBar
    End If
    
End Sub

