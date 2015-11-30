VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{5936A75C-3F42-11D6-AF6B-AA0004005F12}#1.3#0"; "MeansCtrl.ocx"
Begin VB.Form frmEquipRateGrid 
   Caption         =   "Equipment Rate Grid"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11130
   Icon            =   "frmEquipRateGrid.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6855
   ScaleWidth      =   11130
   Begin VB.ListBox RegionList 
      Height          =   645
      Left            =   10080
      MultiSelect     =   2  'Extended
      TabIndex        =   26
      Top             =   1320
      Width           =   735
   End
   Begin VB.ListBox CountryList 
      Height          =   645
      Left            =   8340
      MultiSelect     =   2  'Extended
      TabIndex        =   25
      Top             =   1320
      Width           =   735
   End
   Begin VB.ComboBox CountryCode 
      Height          =   315
      Left            =   6780
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Tag             =   "3S"
      Top             =   2040
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.ComboBox RegionCode 
      Height          =   315
      Left            =   6780
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Tag             =   "3S"
      Top             =   2400
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.TextBox EquipmentID 
      Height          =   315
      Left            =   8340
      TabIndex        =   0
      Top             =   480
      Width           =   1515
   End
   Begin VB.TextBox ContactID 
      Height          =   315
      Left            =   8340
      TabIndex        =   1
      Top             =   900
      Width           =   1515
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   495
      Left            =   8340
      TabIndex        =   4
      Top             =   2160
      Width           =   1150
   End
   Begin VB.CheckBox ckbRowWrap 
      Caption         =   "Row Wrap"
      Height          =   315
      Left            =   60
      TabIndex        =   5
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   495
      Left            =   5940
      TabIndex        =   12
      Top             =   6240
      Width           =   1150
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   495
      Left            =   7260
      TabIndex        =   13
      Top             =   6240
      Width           =   1150
   End
   Begin VB.Frame Frame1 
      Caption         =   "Go To"
      Height          =   855
      Left            =   120
      TabIndex        =   17
      Top             =   6000
      Width           =   4575
      Begin VB.CommandButton cmdEquipment 
         Caption         =   "Equip Maint."
         Height          =   495
         Left            =   960
         TabIndex        =   8
         Top             =   240
         Width           =   795
      End
      Begin VB.CommandButton cmdEquipmentRate 
         Caption         =   "Equip Rate"
         Height          =   495
         Left            =   60
         TabIndex        =   7
         Top             =   240
         Width           =   795
      End
      Begin VB.CommandButton cmdHistory 
         Caption         =   "History"
         Height          =   495
         Left            =   1860
         TabIndex        =   9
         Top             =   240
         Width           =   795
      End
      Begin VB.CommandButton cmdCrews 
         Caption         =   "Crews"
         Height          =   495
         Left            =   2760
         TabIndex        =   10
         Top             =   240
         Width           =   795
      End
      Begin VB.CommandButton cmdInfoSources 
         Caption         =   "Info Sources"
         Height          =   495
         Left            =   3660
         TabIndex        =   11
         Top             =   240
         Width           =   795
      End
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   8580
      TabIndex        =   14
      Top             =   6240
      Width           =   1150
   End
   Begin VB.CommandButton cmdClone 
      Caption         =   "Clone"
      Height          =   495
      Left            =   9900
      TabIndex        =   15
      Top             =   6240
      Width           =   1150
   End
   Begin VB.CommandButton cmdFactor 
      Caption         =   "Factor"
      Height          =   315
      Left            =   1920
      TabIndex        =   6
      Top             =   2880
      Width           =   1150
   End
   Begin ConstructionCostDatabase.DynaTree FormatTree 
      Height          =   2775
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   4895
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid 
      Height          =   2715
      Left            =   60
      TabIndex        =   18
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
      Splits(0)._ColumnProps(5)=   "Column(0)._MinWidth=49"
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=29"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=33"
      _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=30"
      _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=31"
      _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=32"
      _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=34"
      _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=35"
      _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=36"
      _StyleDefs(16)  =   "RecordSelectorStyle:id=37,.parent=2,.namedParent=39"
      _StyleDefs(17)  =   "FilterBarStyle:id=40,.parent=1,.namedParent=42"
      _StyleDefs(18)  =   "Splits(0).Style:id=11,.parent=1"
      _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=20,.parent=4"
      _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=12,.parent=2"
      _StyleDefs(21)  =   "Splits(0).FooterStyle:id=13,.parent=3"
      _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=14,.parent=5"
      _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=16,.parent=6"
      _StyleDefs(24)  =   "Splits(0).EditorStyle:id=15,.parent=7"
      _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=17,.parent=8"
      _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=18,.parent=9"
      _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=19,.parent=10"
      _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=38,.parent=37"
      _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=41,.parent=40"
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=24,.parent=11"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=28,.parent=11"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=15"
      _StyleDefs(38)  =   "Named:id=29:Normal"
      _StyleDefs(39)  =   ":id=29,.parent=0"
      _StyleDefs(40)  =   "Named:id=30:Heading"
      _StyleDefs(41)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(42)  =   ":id=30,.wraptext=-1"
      _StyleDefs(43)  =   "Named:id=31:Footing"
      _StyleDefs(44)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(45)  =   "Named:id=32:Selected"
      _StyleDefs(46)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(47)  =   "Named:id=33:Caption"
      _StyleDefs(48)  =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(49)  =   "Named:id=34:HighlightRow"
      _StyleDefs(50)  =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(51)  =   "Named:id=35:EvenRow"
      _StyleDefs(52)  =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
      _StyleDefs(53)  =   "Named:id=36:OddRow"
      _StyleDefs(54)  =   ":id=36,.parent=29"
      _StyleDefs(55)  =   "Named:id=39:RecordSelector"
      _StyleDefs(56)  =   ":id=39,.parent=30"
      _StyleDefs(57)  =   "Named:id=42:FilterBar"
      _StyleDefs(58)  =   ":id=42,.parent=29"
   End
   Begin VB.Label Label62 
      Alignment       =   1  'Right Justify
      Caption         =   "Country:"
      Height          =   255
      Left            =   7500
      TabIndex        =   24
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label63 
      Alignment       =   1  'Right Justify
      Caption         =   "Region:"
      Height          =   255
      Left            =   9360
      TabIndex        =   23
      Top             =   1320
      Width           =   615
   End
   Begin VB.Line Line1 
      X1              =   6660
      X2              =   6660
      Y1              =   2700
      Y2              =   60
   End
   Begin VB.Line Line2 
      X1              =   60
      X2              =   11040
      Y1              =   2820
      Y2              =   2820
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Equipment ID:"
      Height          =   255
      Left            =   7020
      TabIndex        =   22
      Top             =   540
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Contact ID:"
      Height          =   255
      Left            =   7020
      TabIndex        =   21
      Top             =   960
      Width           =   1215
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
      TabIndex        =   20
      Top             =   60
      Width           =   1215
   End
   Begin VB.Label lblRowCount 
      Caption         =   "0 rows returned"
      Height          =   255
      Left            =   5160
      TabIndex        =   19
      Top             =   2880
      Width           =   3255
   End
End
Attribute VB_Name = "frmEquipRateGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' <modulename> frmEquipRateGrid.frm</modulename>
' <functionname>General (Main) </functionname>
'
' <summary>
' (CCI) EQUIPMENT RATE GRID
'
'* * * WARNING * * *  WARNING * * *  WARNING * * *  * * * * * * * * * * * * * * * * * *
'
'A significant amount of functionality is not working with this app
'Please approach Steve Plotner regarding functionality issues.
'It is my understanding (per K. R.) that this is routinely managed by way of a spreadsheet
'
'
'Display Equipment rental rates based upon "Equipment ID":
'"   Equipment ID
'"   Contact ID:
'"   Country
'"   Region
'
'
'(BUTTONS)
'"   SEARCH              (CmdSearch_Click() )
'Search for equipment rental rate data based upon "Equipment ID" or "Tech Desc"
'"   Factor                  (CEquipRateMap.Factor()
'
'"   Equip Rate              (frmEquipRate)
'"   Equip Maint             (frmEquipment)
'"   History
'"   Crews                   - NOT SUPPORTED -
'"   Info Sources                (frmInfoSourceGrid)
'- NOT WORKING -
'"   Update              (CEquipmentRateMap.Update() )
'"   New                 (frmEquipRate)
'"   Delete                  (TDBGrid.Delete)
'"   Clone                   (frmEquipRate)
'
'
'Key Subs / Functions:
'"   CmdSearch_Click()
'select Equipment.equip_skey, Equipment.equip_id, Equipment.alt_equip_id, Equipment.type_code, Equipment.book_desc, Equipment.tech_desc, Equipment.unit, Equipment.metric_book_desc, Equipment.metric_tech_desc, Equipment.crew_equip_desc, Equipment.crew_equip_desc_plural, Equipment.index_code, Equipment.index_desc, Equipment.metric_crew_equip_desc, Equipment.metric_crew_equip_desc_plural, Equipment.metric_unit, Equipment.model_name, Equipment.traces_ind, Equipment.indent_code, Equipment.format_characters, Equipment.format_code, Equipment.last_update_id as 'equip_last_update_id', Equipment.last_update_date as 'equip_last_update_date', Equipment.last_update_person as 'equip_last_update_person', #er.contact_id, #er.start_date, #er.term_date, #er.rent_per_week, #er.operating_cost_hrly, #er.factor_ind, #er.estimated_ind, #er.comment, #er.info_source_ref, #er.last_update_date as 'equiprate_last_update_date', #er.last_update_person as 'equiprate_last_update_person', #er.last_update_id as 'equiprate_last_update_id'into
'
'#eer from equipment LEFT OUTER JOIN #er ON (#er.equip_skey = equipment.equip_skey) where Equipment.equip_id like '01%'
'
'HELPER Class: CEquipRateMap.Cls
' </summary>
'
' <seealso> CEquipRateMap.cls</seealso>
'<seealso> </seealso>
'
' <datastruct>m_rec</datastruct>
'<datastruct>m_objGridMap</datastruct>
'
' <storedprocedurename> n/a </storedprocedurename>
'<storedprocedurename> n/a </storedprocedurename>
'
'
' <returns>N/A</returns>
' <exception>Always trap with an accompanying message box</exception>
' <example>
'<code>* * * DROP temp tables ?
'drop table #er
'</code>
'
'<code>* * * DROP temp tables ?
'drop table #eer
'</code>
'
'<code>* * * DROP temp tables ?
'drop table #pere
'</code>
'<code>* * * SEARCH/SELECT * * *
'select Equipment.equip_skey, Equipment.equip_id, Equipment.alt_equip_id, Equipment.type_code, Equipment.book_desc, Equipment.tech_desc, Equipment.unit, Equipment.metric_book_desc, Equipment.metric_tech_desc, Equipment.crew_equip_desc, Equipment.crew_equip_desc_plural, Equipment.index_code, Equipment.index_desc, Equipment.metric_crew_equip_desc, Equipment.metric_crew_equip_desc_plural, Equipment.metric_unit, Equipment.model_name, Equipment.traces_ind, Equipment.indent_code, Equipment.format_characters, Equipment.format_code, Equipment.last_update_id as 'equip_last_update_id', Equipment.last_update_date as 'equip_last_update_date', Equipment.last_update_person as 'equip_last_update_person', #er.contact_id, #er.start_date, #er.term_date, #er.rent_per_week, #er.operating_cost_hrly, #er.factor_ind, #er.estimated_ind, #er.comment, #er.info_source_ref, #er.last_update_date as 'equiprate_last_update_date', #er.last_update_person as 'equiprate_last_update_person', #er.last_update_id as 'equiprate_last_update_id'into
'#eer from equipment LEFT OUTER JOIN #er ON (#er.equip_skey = equipment.equip_skey) where Equipment.equip_id like '01%'
'</code>
'<code> * * *  UPON CLICKING THE "FACTOR" BUTTON * * *
'All selected rows are "factored" by values indicated in code below
'
'Public Sub Factor(dblFactor As Double, intApply As Integer)
'    Dim vntBookmark As Variant
'
'    For Each vntBookmark In TDBGrid.SelBookmarks
'        m_rec.Bookmark = vntBookmark
'        If intApply = EQUIP_FACTOR_RENT Or intApply = EQUIP_FACTOR_BOTH Then
'            m_rec.Fields("Rent_per_week") = m_rec.Fields("Rent_per_week") + m_rec.Fields("Rent_per_week") * dblFactor / 100
'        End If
'        If intApply = EQUIP_FACTOR_OPERATING Or intApply = EQUIP_FACTOR_BOTH Then
'            m_rec.Fields("Operating_cost_hrly") = m_rec.Fields("Operating_cost_hrly") + m_rec.Fields("Operating_cost_hrly") * dblFactor / 100
'        End If
'        m_rec.Fields("Factor_ind") = -1
'        m_objGridMap.SetRowState Int(vntBookmark), STATE_MODIFIED
'    Next
'    vntBookmark = TDBGrid.SelBookmarks(0)
'    TDBGrid.ReBind ' Reset grid contents
'    TDBGrid.Bookmark = vntBookmark ' Set bookmark back again
'End Sub
'</code>
'</example>
'<permission>Public</Permission>
'<dependson>This component depends on the following:
'1.  CEquipmentMap.cls
'2.  CGridMap.cls
'3.  CCDdal.CRSMDataAccess (
'4.  Access to the DAL (data access layer dll) opened in MainModule_Main() )
'</dependson>




Dim m_objGridMap As New CEquipRateMap ' Class to handle grid
Public m_blnFirstSearch As Boolean ' Is this the first search we have made on this screen
Dim m_blnJumpIn As Boolean ' Are we jumping here from another screen
Dim m_rec As New ADODB.RecordSet ' Recordset to hold query results
Dim m_blnDoubleClick As Boolean ' Did a double click just occurr
Dim m_blnWereErrors As Boolean ' True if the Update had errors, used in QueryUnload
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

' Handles Row Wrap feature
Private Sub ckbRowWrap_Click()
    m_objGridMap.RowWrap (ckbRowWrap)
End Sub

Private Sub cmdClone_Click()
    On Error GoTo Out
    If IsNumeric(TDBGrid.Bookmark) = False Then
        MsgBox "You must select a row."
        Exit Sub
    End If
    Dim rec As ADODB.RecordSet
    
    Set rec = m_objGridMap.CloneRow
    ' Force any changes into recordset from grid
    TDBGrid.Update
    ' Navigate to single-record view
    Dim frm As frmEquipRate
'    Dim rec As ADODB.RecordSet
    Set frm = New frmEquipRate
    ' Make copy of recordset
'    Set rec = m_rec.Clone
    ' Get the selected row from grid
'    rec.Bookmark = TDBGrid.Bookmark
    frm.SetRow rec, True ' Pass the current record into the form
    frm.Show
Out:
End Sub

Private Sub cmdDelete_Click()
    Dim varButton
    varButton = MsgBox("Are you sure you want to delete?", vbYesNo + vbCritical)
    If varButton = vbYes Then
        TDBGrid.Delete
    End If
End Sub

Private Sub cmdFactor_Click()
    Dim dblFactor As Double
    Dim intApply As Integer
    
    If TDBGrid.Columns("Type").CellText(TDBGrid.Bookmark) = "H" Then
        MsgBox "You cannot apply a factor to header (H) rows."
        Exit Sub
    End If
    If TDBGrid.SelBookmarks.Count > 0 Then
        dblFactor = -1
        dlgEquipFactor.GetFactor dblFactor, intApply
        m_objGridMap.Factor dblFactor, intApply
    Else
        MsgBox "You must select a row first"
    End If
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
    Dim frm As frmEquipHistoryGrid
    Set frm = New frmEquipHistoryGrid
    frm.JumpIn TDBGrid.Columns("Equip ID").CellText(TDBGrid.Bookmark)
End Sub

Private Sub cmdInfoSources_Click()
    If IsNumeric(TDBGrid.Bookmark) = False Then
        MsgBox "You must select a row."
        Exit Sub
    End If
    If TDBGrid.Columns("Type").CellText(TDBGrid.Bookmark) = "H" Then
        MsgBox "No Information Source is available for header (H) rows."
        Exit Sub
    End If
    ' Open spreadsheet view with data from row selected
    Dim frm As frmInfoSourceGrid
    Set frm = New frmInfoSourceGrid
    frm.JumpIn TDBGrid.Columns("Contact").CellText(TDBGrid.Bookmark)
End Sub

Private Sub cmdEquipment_Click()
    If IsNumeric(TDBGrid.Bookmark) = False Then
        MsgBox "You must select a row."
        Exit Sub
    End If
    ' Navigate to single-record view
    Dim frm As frmEquipment
    Dim rec As ADODB.RecordSet
    Set frm = New frmEquipment
    ' Make copy of recordset
    Set rec = m_rec.Clone
    ' Get the selected row from grid
    rec.Bookmark = TDBGrid.Bookmark
    frm.SetRow rec ' Pass the current record into the form
    frm.Show
End Sub

Private Sub cmdEquipmentRate_Click()
    If IsNumeric(TDBGrid.Bookmark) = False Then
        MsgBox "You must select a row."
        Exit Sub
    End If
    ' Navigate to single-record view
    Dim frm As frmEquipRate
    Dim rec As ADODB.RecordSet
    Set frm = New frmEquipRate
    ' Make copy of recordset
    Set rec = m_rec.Clone
    ' Get the selected row from grid
    rec.Bookmark = TDBGrid.Bookmark
    frm.SetRow rec ' Pass the current record into the form
    frm.Show
End Sub

Private Sub cmdCrews_Click()
    ' Open single record view with data from row selected
'    Dim frm As frmMatUsageGrid
'    Set frm = New frmMatUsageGrid
'    frm.JumpIn TDBGrid.Columns("Material ID").CellText(TDBGrid.Bookmark)
End Sub

Private Sub cmdNew_Click()
    On Error GoTo Out
    Dim rec As New ADODB.RecordSet
    
    CopyRSFields rec, m_rec
    ' Open empty single record view
    Dim frm As frmEquipRate
    Set frm = New frmEquipRate
    ' Force any changes into recordset from grid
    TDBGrid.Update
    frm.SetRow rec, True
    frm.Show
Out:
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

Private Sub ContactID_Change()
ContactID.Text = Trim(ContactID.Text)
End Sub

Private Sub EquipmentID_LostFocus()
EquipmentID.Text = Trim(EquipmentID.Text)
End Sub


Private Sub Form_Deactivate()
m_strCurrentFormControl = Me.ActiveControl.Name
End Sub

Private Sub Form_Initialize()
    Screen.MousePointer = vbHourglass
    m_blnFirstSearch = False
    ' Fill the MasterFormat tree
    FormatTree.InitData g_cnShared, "EQUIPMENT"
    ' Initialize grid
    m_objGridMap.SetGrid TDBGrid
    m_objGridMap.InitGrid
    m_blnJumpIn = False
    Screen.MousePointer = vbNormal
    m_blnFirstSearch = False
End Sub
Private Sub Form_Load()
    Dim blnReturn As Boolean
    Dim strSelect As String
    Dim rec As ADODB.RecordSet
    Move START_LEFT, START_TOP, START_WIDTH, START_HEIGHT
    
    g_objDAL.GetRecordset CONNECT, "select country_code from country order by country_code", rec
    While Not rec.EOF
        CountryList.AddItem (rec.Fields("country_code").Value)
        If rec.Fields("country_code").Value = "USA" Then
            CountryList.Selected(CountryList.listcount - 1) = True
        End If
        rec.MoveNext
    Wend
    rec.Close
    g_objDAL.GetRecordset CONNECT, "select region_code from region order by region_code", rec
    While Not rec.EOF
        RegionList.AddItem (rec.Fields("region_code").Value)
        If rec.Fields("region_code").Value = "NAT" Then
            RegionList.Selected(RegionList.listcount - 1) = True
        End If
        rec.MoveNext
    Wend
    rec.Close
    
    ' This will never return any rows, just used to create recordset
'    strSelect = "select * from Equipment, Equipment_rate where Equipment.equip_skey = Equipment_rate.equip_skey "
'    strSelect = strSelect + " and Equipment.equip_id = '0'"
    EquipmentID.Text = "~"
    cmdSearch_Click
    EquipmentID.Text = ""
    
'    blnReturn = g_objDAL.GetRecordset(vbNullString, strSelect, m_rec)
'    m_objGridMap.RecordSet = m_rec
End Sub

' Called when coming here from another screen
Public Sub JumpIn(strMatID As String)
    EquipmentID.Text = Compress_String(strMatID)
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
    ' Synch text box with tree
    EquipmentID.Text = strID + "*"
    ' Clear other boxes
    ContactID.Text = ""
    ' Kick-off search
    cmdSearch_Click
End Sub

Private Sub cmdSearch_Click()
    On Error Resume Next
    Dim blnReturn As Boolean
    Dim strSelect As String
    Dim dtmToday As Date
    Dim dtmStart As Date
    
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
            Exit Sub
        Else
            TDBGrid.DataChanged = False
        End If
    End If
    
    Screen.MousePointer = vbHourglass
    dtmToday = Date
    
    ' Synch tree with text box
    If Not EquipmentID.Text = "" Then
        FormatTree.FocusItem (EquipmentID.Text)
    End If
    
    If Len(EquipmentID.Text) = 0 And Len(ContactID.Text) = 0 Then
        Screen.MousePointer = vbNormal
        MsgBox "You must enter Equipment ID or Contact ID"
        Exit Sub
    End If
    
    lblRowCount.Caption = "Working..."
    lblRowCount.Refresh
    
    m_rec.Close ' Make sure it is closed
    m_rec.MaxRecords = MAX_RECORDS ' Set the maximum number to bring back
    dtmStart = Now
    
    Dim strError As String
    blnReturn = g_objDAL.ExecQuery(vbNullString, "drop table #er", strError)
    
    strSelect = "select er.* into #er from equipment_rate as er, equipment " + _
                "where er.equip_skey = equipment.equip_skey" + _
                " and start_date <= '" + Format(dtmToday, "mm/dd/yyyy") + "' and term_date >= '" + Format(dtmToday, "mm/dd/yyyy") + "'"
    
    If Not EquipmentID.Text = "" Then
        strSelect = strSelect + " and Equipment.equip_id like '"
        strSelect = strSelect + SQLChangeWildcard(EquipmentID.Text) + "'"
    End If
    If Not ContactID.Text = "" Then
        strSelect = strSelect + " and er.contact_id like '" + SQLChangeWildcard(ContactID.Text) + "'"
    End If

    ' Use DAL to perform select
    blnReturn = g_objDAL.ExecQuery(vbNullString, strSelect, strError)
    
    blnReturn = g_objDAL.ExecQuery(vbNullString, "drop table #eer", strError)
    
    strSelect = "select Equipment.equip_skey, Equipment.equip_id, Equipment.alt_equip_id, Equipment.type_code, Equipment.book_desc, Equipment.tech_desc, Equipment.unit, Equipment.metric_book_desc, Equipment.metric_tech_desc, Equipment.crew_equip_desc, Equipment.crew_equip_desc_plural, Equipment.index_code, Equipment.index_desc, Equipment.metric_crew_equip_desc, Equipment.metric_crew_equip_desc_plural, Equipment.metric_unit, Equipment.model_name, Equipment.traces_ind, Equipment.indent_code, Equipment.format_characters, " + _
                "Equipment.format_code, Equipment.last_update_id as 'equip_last_update_id', Equipment.last_update_date as 'equip_last_update_date', Equipment.last_update_person as 'equip_last_update_person', " + _
                "#er.contact_id, #er.start_date, #er.term_date, #er.rent_per_week, #er.operating_cost_hrly, #er.factor_ind, #er.estimated_ind, #er.comment, #er.info_source_ref, #er.last_update_date as 'equiprate_last_update_date', #er.last_update_person as 'equiprate_last_update_person', #er.last_update_id as 'equiprate_last_update_id'" + _
                "into #eer from equipment LEFT OUTER JOIN #er ON (#er.equip_skey = equipment.equip_skey) where "
    
    If Not Len(EquipmentID.Text) = 0 Then
        strSelect = strSelect + "Equipment.equip_id like '"
        strSelect = strSelect + SQLChangeWildcard(EquipmentID.Text) + "'"
    End If
    If Not Len(ContactID.Text) = 0 Then
        If Not Len(EquipmentID.Text) = 0 Then
            strSelect = strSelect + " and "
        End If
        strSelect = strSelect + "#er.contact_id like '" + SQLChangeWildcard(ContactID.Text) + "'"
    End If
    
    ' Use DAL to perform select
    blnReturn = g_objDAL.ExecQuery(vbNullString, strSelect, strError)
    
    blnReturn = g_objDAL.ExecQuery(vbNullString, "drop table #pere", strError)
    
    strSelect = "select pere.equip_skey, pere.country_code, pere.region_code, pere.start_date as start_date_x, pere.term_date as term_date_x, pere.rent_per_day as rent_per_day_x, " + _
                "pere.rent_per_week as rent_per_week_x, pere.rent_per_month as rent_per_month_x, pere.operating_cost_hrly as operating_cost_hrly_x, pere.crew_equip_cost as crew_equip_cost_x, " + _
                "pere.metric_rent_per_day as metric_rent_per_day_x, pere.metric_rent_per_week as metric_rent_per_week_x, pere.metric_rent_per_month as metric_rent_per_month_x, pere.metric_operating_cost_hrly as metric_operating_cost_hrly_x, pere.metric_crew_equip_cost as metric_crew_equip_cost_x, " + _
                "pere.pct_ind, pere.last_update_date as equiprate_last_update_date_x, pere.last_update_person as equiprate_last_update_person_x, pere.last_update_id as equiprate_last_update_id_x into #pere from published_equipment_rate_excep as pere " + _
                "where start_date <= '" + Format(dtmToday, "mm/dd/yyyy") + "' and term_date >= '" + Format(dtmToday, "mm/dd/yyyy") + "' " + _
                "and region_code IN " + BuildINFromListbox(RegionList) + " and country_code IN " + BuildINFromListbox(CountryList) + ""

    ' Use DAL to perform select
    blnReturn = g_objDAL.ExecQuery(vbNullString, strSelect, strError)
    
    strSelect = "select #eer.contact_id, #eer.start_date, #eer.term_date, #eer.rent_per_week, #eer.operating_cost_hrly, #eer.factor_ind, #eer.estimated_ind, #eer.comment, #eer.info_source_ref, #eer.equiprate_last_update_person, #eer.equiprate_last_update_date, " + _
        "#pere.country_code, #pere.region_code, #pere.start_date_x, #pere.term_date_x, #pere.rent_per_day_x, #pere.rent_per_week_x, #pere.rent_per_month_x, #pere.operating_cost_hrly_x, #pere.crew_equip_cost_x, #pere.metric_rent_per_day_x," + _
        "#pere.metric_rent_per_week_x, #pere.metric_rent_per_month_x, #pere.metric_operating_cost_hrly_x, #pere.metric_crew_equip_cost_x, #pere.pct_ind, #pere.equiprate_last_update_date_x, #pere.equiprate_last_update_person_x, #pere.equiprate_last_update_id_x, " + _
        "#eer.equip_skey, #eer.equip_id, #eer.alt_equip_id, #eer.type_code, #eer.book_desc, #eer.tech_desc, #eer.unit, #eer.metric_book_desc, #eer.metric_tech_desc, #eer.crew_equip_desc, #eer.crew_equip_desc_plural, #eer.index_code," + _
        "#eer.index_desc, #eer.metric_crew_equip_desc, #eer.metric_crew_equip_desc_plural, #eer.metric_unit, #eer.model_name, #eer.traces_ind, #eer.indent_code, #eer.format_characters, " + _
        "#eer.format_code, #eer.equip_last_update_id, #eer.equip_last_update_date, #eer.equip_last_update_person, " + _
        "#eer.equiprate_last_update_id " + _
        "from #eer LEFT OUTER JOIN #pere ON (#eer.equip_skey = #pere.equip_skey) " ' + _

    strSelect = strSelect + " ORDER BY #eer.equip_id, #eer.contact_id"
    
    ' Use DAL to perform select
    blnReturn = g_objDAL.GetRecordset(vbNullString, strSelect, m_rec)
    If blnReturn = False Then
        MsgBox "An error occurred while searching."
        lblRowCount.Caption = "0 rows returned."
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

Private Sub TDBGrid_LostFocus()
TDBGrid.TabStop = False
End Sub

Private Sub TDBGrid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If this is the mouse-up form a double click
    If m_blnDoubleClick Then
        ' Make sure it is the left button
        If Button = vbLeftButton Then
            m_blnDoubleClick = False
            ' Same function as clicking Material Price button, open single record view
            cmdEquipmentRate_Click
        End If
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
        TDBGrid.ReBind
        OutputView False
        ShowGridSort
        m_objGridMap.SetMenuBar
    End If
End Sub



