VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{5936A75C-3F42-11D6-AF6B-AA0004005F12}#1.3#0"; "MeansCtrl.ocx"
Begin VB.Form frmBuildingGrid 
   Caption         =   "Building Grid"
   ClientHeight    =   6900
   ClientLeft      =   1500
   ClientTop       =   2790
   ClientWidth     =   11130
   Icon            =   "frmFacilityGrid.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6900
   ScaleWidth      =   11130
   Begin VB.Frame fraBldgType 
      Caption         =   "Building Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6960
      TabIndex        =   19
      Top             =   480
      Width           =   4095
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   3735
         TabIndex        =   20
         Top             =   240
         Width           =   3735
         Begin VB.OptionButton opttype_codeB 
            Caption         =   "Both"
            Height          =   255
            Left            =   2880
            TabIndex        =   23
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton opttype_codeC 
            Caption         =   "Commercial"
            Height          =   255
            Left            =   0
            TabIndex        =   22
            Top             =   0
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton opttype_codeR 
            Caption         =   "Residential"
            Height          =   255
            Left            =   1440
            TabIndex        =   21
            Top             =   0
            Width           =   1095
         End
      End
   End
   Begin VB.ComboBox bldg_category 
      Height          =   315
      ItemData        =   "frmFacilityGrid.frx":0442
      Left            =   7785
      List            =   "frmFacilityGrid.frx":0452
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1200
      Width           =   2750
   End
   Begin ConstructionCostDatabase.DynaTree FormatTree 
      Height          =   2775
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   4895
   End
   Begin VB.TextBox txtbldg_desc 
      Height          =   285
      Left            =   7800
      TabIndex        =   2
      Top             =   1920
      Width           =   3495
   End
   Begin VB.TextBox txtbldg_id 
      Height          =   285
      Left            =   7800
      TabIndex        =   1
      Top             =   1560
      Width           =   1065
   End
   Begin VB.CommandButton cmdClone 
      Caption         =   "&Clone"
      Height          =   495
      Left            =   10080
      TabIndex        =   10
      Top             =   6240
      Width           =   915
   End
   Begin VB.Frame fraGoTo 
      Caption         =   "Go To"
      Height          =   855
      Left            =   120
      TabIndex        =   13
      Top             =   6000
      Width           =   3435
      Begin VB.CommandButton cmdBuilding 
         Caption         =   "&Building"
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   915
      End
      Begin VB.CommandButton cmdModels 
         Caption         =   "&Models"
         Height          =   495
         Left            =   1260
         TabIndex        =   6
         Top             =   240
         Width           =   915
      End
      Begin VB.CommandButton cmdOutput 
         Caption         =   "&Output"
         Height          =   495
         Left            =   2280
         TabIndex        =   7
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   495
      Left            =   9000
      TabIndex        =   9
      Top             =   6240
      Width           =   915
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7920
      TabIndex        =   8
      Top             =   6240
      Width           =   915
   End
   Begin VB.CheckBox ckbRowWrap 
      Caption         =   "Row Wrap"
      Height          =   315
      Left            =   60
      TabIndex        =   4
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   375
      Left            =   7800
      TabIndex        =   3
      Top             =   2280
      Width           =   1150
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGridBuilding 
      Height          =   2715
      Left            =   60
      TabIndex        =   11
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
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(1)._MinWidth=1241540"
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
   Begin VB.Label lblbldg_category 
      Alignment       =   1  'Right Justify
      Caption         =   "Category:"
      Height          =   255
      Left            =   6840
      TabIndex        =   18
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblBuildingDesc 
      Alignment       =   1  'Right Justify
      Caption         =   "Description:"
      Height          =   255
      Left            =   6840
      TabIndex        =   17
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label lblBuildingID 
      Alignment       =   1  'Right Justify
      Caption         =   "Building ID:"
      Height          =   255
      Left            =   6840
      TabIndex        =   16
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label lblRowCount 
      Alignment       =   2  'Center
      Caption         =   "0 rows returned"
      Height          =   195
      Left            =   2865
      TabIndex        =   15
      Top             =   2880
      Width           =   6510
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
      Left            =   6825
      TabIndex        =   14
      Top             =   120
      Width           =   1215
   End
   Begin VB.Line Line2 
      X1              =   60
      X2              =   11040
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
Attribute VB_Name = "frmBuildingGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''' <modulename> frmBuildingGrid</modulename>
''' <functionname>General (Main) </functionname>
'''
''' <summary>
''' Provides u/i permitting user to do the following:
'''
'''Add or Change a "building prototype" that comes with a building id (bldg_id)
'''E.g.  "Apartment, 1-3 Story, Community Center, Factory 1 Story etc. etc.)
'''Buildings fall into (2) categories:
'''"   Commercial
'''"   Residential
'''Buildings are built from "models".
'''
'''(Major function buttons)
'''
'''1.  Display "Building" form         (frmBuilding.frm)
'''2.  Display "Models." form          (frmModelGrid.frm)
'''3.  Display "Output"                (dlgOutput.frm)
'''4.  "Update" new/changed data           No form.
''' (m_objGridMap.Update())
'''5.  Create a NEW building line          (frmBuilding.frm)
'''6.  Clone a new building line           (frmBuilding.frm
'''
'''* * * WARNING  * * * *
'''DO NOT CLICK CLONE BUTTON UNLESS YOU REALLY MEAN IT!
'''IT ALWAYS CREATES THE NEW BUILDING WHETHER YOU WANT TO KEEP IT OR NOT!!!
'''* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *  * * * * * * * * * * * *  * * * *  *
'''HELPER Class: CBuildingMap.Cls
'''
'''NOTE:  file names do not match Project component names (e.g. frmBuildingGrid is really frmFacilityGrid.frm)
'''
'''</summary>
'''
'''<seealso> CBuildingMap.cls </seealso>
'''<seealso>frmBuilding.frm</seealso>
'''
''' <datastruct>m_objGridMap</datastruct>
'''<datastruct>m_rec</datastruct>
'''
'''<storedprocedurename> sp_select_building </storedprocedurename>
'''<storedprocedurename> sp_clone_building </storedprocedurename>
'''<storedprocedurename> sp_temp_output_init</storedprocedurename>
'''<storedprocedurename> sp_temp_add_output_keys</storedprocedurename>
'''
'''
'''
''' <returns>N/A</returns>
''' <exception>Always trap with an accompanying message box</exception>
''' <example>
''' <code>
'''exec sp_select_building @type_code = '%', @bldg_category = '%', @bldg_id = '001', @bldg_desc = '%' </code>
''' <code>
'''</code>
'''</example>
'''<permission>Public</Permission>
'''<dependson>This component depends on the following
'''1.  CBuildingMap.cls
'''2.  CGridMap.cls
'''3.  CCDdal.CRSMDataAccess (
'''Access to the DAL (data access layer dll) opened in MainModule_Main() )
'''</dependson>



'
'   Class to handle grid
Dim m_objGridMap As New CBuildingMap
'
'   Is this the first search we have made
'   on this screen.
Dim m_blnFirstSearch As Boolean
'
'   Recordset to hold query results
Dim m_rec As New ADODB.RecordSet
'
'   True if the Update had errors, used in QueryUnload
Dim m_blnWereErrors As Boolean
'
'   Keeps up with the field that last had focus when form
'   is deactivate, so when activated can set focus.
Dim m_strCurrentFormControl As String
'
'   Notifies that it wants to see changes.
Dim sEventSubscriberID              As String

Const USEBOOKMARK = 1
Const USECOORD = 0

Private Sub Form_Load()

    On Error Resume Next
    Move START_LEFT, START_TOP, START_WIDTH, START_HEIGHT
    
    ' ADDED 6/16/2005 RTD FOR VERSION 7.4.0
    TDBGridBuilding.CellTips = dbgAnchored
    
    '
    '   This will never return any rows, just used to create recordset????
    cmdSearch_Click
    
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim Button  As String
    
    On Error Resume Next
    '
    '   Check if there are pending changes
    If m_objGridMap.IsPendingChange = True Then
        Button = MsgBox("Do you want to save your changes to " & Me.Caption & "?", vbYesNoCancel, "Close Building Grid Form")
        If Button = vbYes Then
            cmdUpdate_Click
            '
            '   If there were errors, cancel the close.
            If m_blnWereErrors Then
                Cancel = True
            End If
        ElseIf Button = vbCancel Then
            Cancel = True
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '
    '   Disables & hides the sort buttons on the main form.
    HideGridSort
    ShowToolbarIcons False
    EventSubscriberRemove sEventSubscriberID
    
End Sub

Private Sub Form_Initialize()

    On Error Resume Next
    Screen.MousePointer = vbHourglass
    Status ("Loading Building Maintenance Grid ...")
    sEventSubscriberID = EventSubscriberAdd(Me)
    m_blnFirstSearch = True
    '
    '   Fill the MasterFormat tree.
    FormatTree.InitData g_cnShared, "BUILDING"
    '
    '   Initialize grid.
    With m_objGridMap
        .SetGrid TDBGridBuilding
        .InitGrid
    End With
    '
    '   When records are loaded the Bldg class has
    '   to do manipulations to the Type & Category
    '   columns causing it to look like data was changed.
    TDBGridBuilding.DataChanged = False
    PopulateBldgCategories
    m_blnFirstSearch = False
    Status ("")
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Activate()
    Dim ctl As Control
    
    On Error Resume Next
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

Private Sub Form_Deactivate()
    ShowToolbarIcons False
    m_strCurrentFormControl = Me.ActiveControl.Name
End Sub

Private Sub Form_LostFocus()
    TDBGridBuilding.Update
    HideGridSort
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    '
    '   Need to place in common routine for all forms.
    '   Possibly place all buttons in a frame like frame1 with
    '   common name and can just place it.
    If Me.WindowState = vbNormal Or Me.WindowState = vbMaximized Then
        If Me.Width >= 11250 Then
            TDBGridBuilding.Width = Me.Width - 255
            Line2.X2 = Me.Width - 210
        Else
            Me.Width = 11250
        End If
        
        If Me.Height >= 7260 Then
            TDBGridBuilding.Height = Me.Height - 4545
            fraGoTo.Top = Me.Height - 1260
            cmdUpdate.Top = Me.Height - 1020
            cmdNew.Top = Me.Height - 1020
            cmdClone.Top = Me.Height - 1020
        Else
            Me.Height = 7260
        End If
    Else
        ShowMinimizedForms
    End If
End Sub

Private Sub SetButtons(Mode As Single, Optional Coord As Variant)
    On Error Resume Next
    '
    '   No current record - disable buttons
    If m_rec.RecordCount > 0 And TDBGridBuilding.Bookmark >= 1 Then
        Select Case Mode
            Case USEBOOKMARK
                m_rec.Bookmark = TDBGridBuilding.Bookmark
            Case USECOORD
                m_rec.Bookmark = TDBGridBuilding.RowBookmark(TDBGridBuilding.RowContaining(Coord))
        End Select
        
        If IsNumeric(m_rec.Bookmark) Then
            cmdBuilding.Enabled = True
            cmdModels.Enabled = True
            cmdOutput.Enabled = True
            cmdClone.Enabled = True
            '
            '   Don't set update unless there has been a change in the grid.
            'cmdUpdate.Enabled = True
        Else
            cmdBuilding.Enabled = False
            cmdModels.Enabled = False
            cmdOutput.Enabled = False
            cmdClone.Enabled = False
            cmdUpdate.Enabled = False
        End If
    Else
        cmdBuilding.Enabled = False
        cmdModels.Enabled = False
        cmdOutput.Enabled = False
        cmdClone.Enabled = False
        cmdUpdate.Enabled = False
    End If
End Sub

Private Sub PopulateBldgCategories()
    Dim rec         As New ADODB.RecordSet
    Dim strSelect   As String
    
    Screen.MousePointer = vbHourglass
    '
    '   Fill the available categories based on the type code.
    bldg_category.Clear
        
    strSelect = "select bldg_category from bldg_category order by bldg_category"
    If Not g_objDAL.GetRecordset(vbNullString, strSelect, rec) Then
        Screen.MousePointer = vbNormal
        MsgBox "An error occurred while searching for available building categories."
    Else
        With rec
            If .RecordCount = 0 Then
                bldg_category.AddItem "(unknown)"
            Else
                bldg_category.AddItem "ALL"
                While Not .EOF
                    bldg_category.AddItem Trim(.Fields("bldg_category").Value)
                    .MoveNext
                Wend
            End If
            .Close
        End With
    End If
    Screen.MousePointer = vbNormal
End Sub

Public Sub EventNotify(eNotifyType As EEventSubscriberNotifyType, sAffectedRecordIdentifier As String)
    Dim varBookmark
    
    On Error Resume Next
    '
    '   If the record that was updated is in our grid results
    '   we need to refresh.
    If eNotifyType = esnBuildingRecordUpdated And _
        txtbldg_id.Text = "" Or Trim(txtbldg_id.Text) = Trim(sAffectedRecordIdentifier) Then
        '
        '   Need to clear fields that could have been updated if the bldg_id matches
        '   the bldg_id we updated.
        If Trim(txtbldg_id.Text) = Trim(sAffectedRecordIdentifier) Then
            bldg_category.ListIndex = -1
            txtbldg_desc.Text = ""
        End If
        
        varBookmark = TDBGridBuilding.Bookmark
        cmdSearch_Click
        TDBGridBuilding.Bookmark = varBookmark
        '
        '   Fill the MasterFormat tree.
        With FormatTree
            .ClearTree
            .InitData g_cnShared, "BUILDING"
            .FocusItem "0"
        End With
    End If
End Sub
'
'   Called from frmMain when the user clicks on the
'   toolbar buttons for sorting.
Public Sub Sort(intDir As Integer)
    m_objGridMap.Sort intDir
End Sub
'
'   Called when coming here from another screen
Public Sub JumpIn(strBuildingID As String)
    txtbldg_id.Text = Trim(strBuildingID)
    opttype_codeB.Value = True
    cmdSearch_Click
End Sub

Public Sub DoOutput()
    Dim frm         As Form
    Dim blnVisible  As Boolean
    Dim strUpdate   As String
    Dim strError    As String
    Dim blnReturn   As Boolean
    Dim varBookmark As Variant

    If FormOpen("dlgOutput", frm, blnVisible) = False Then
        If Not IsNull(TDBGridBuilding.Bookmark) Then
            Set frm = New dlgOutput
        Else
            Exit Sub
        End If
    End If
    frm.Visible = True
    
    If Not (TDBGridBuilding.BOF = True Or TDBGridBuilding.EOF = True) Then
        strUpdate = "exec sp_temp_output_init"
        '
        '   m_objOutput is a global object for CCDal.CRSMDataAccess
        blnReturn = frm.m_objOutput.ExecQuery(vbNullString, strUpdate, strError)
        '
        '   No rows selected.
        If TDBGridBuilding.SelBookmarks.Count = 0 Then
            '
            '   Use current row
            If Not IsNull(TDBGridBuilding.Bookmark) Then
                m_rec.Bookmark = TDBGridBuilding.Bookmark
                '
                '   Valid values are A = assembly, SF = SquareFoot, E = equipment, U = unit
                strUpdate = "exec sp_temp_add_output_keys @skey_type = 'SF', @skey = " _
                            & CStr(m_rec.Fields("bldg_skey"))
                blnReturn = frm.m_objOutput.ExecQuery(vbNullString, strUpdate, strError)
            End If
        Else
            For Each varBookmark In TDBGridBuilding.SelBookmarks
                m_rec.Bookmark = varBookmark
                '
                '   Valid values are A = assembly, SF = SquareFoot, E = equipment, U = unit
                strUpdate = "exec sp_temp_add_output_keys @skey_type = 'SF', @skey = " _
                            & CStr(m_rec.Fields("bldg_skey"))
                blnReturn = frm.m_objOutput.ExecQuery(vbNullString, strUpdate, strError)
            Next varBookmark
        End If

        With frm
            .FillData
            .Show vbModeless, fMainForm
            .Caption = "Output Usage"
        End With
    End If
End Sub

Private Sub position_output(Optional Y As Single = 0)
    Dim frm             As Form
    Dim blnVisible      As Boolean
    Dim lngCurrentRow   As Long
    '
    '   Only send data to the Output dialog if it is open.
    Screen.MousePointer = vbHourglass
    With TDBGridBuilding
        If FormOpen("dlgOutput", frm, blnVisible) = True Then
            If Y <> 0 Then
                lngCurrentRow = .RowContaining(Y)
                If lngCurrentRow <> -1 Then
                    .Row = lngCurrentRow
                    m_rec.Bookmark = .Bookmark
                End If
            End If
            DoOutput
            Me.SetFocus
        End If
    End With
    Screen.MousePointer = vbNormal
End Sub

Private Sub TDBGridBuilding_FetchCellTips(ByVal SplitIndex As Integer, ByVal ColIndex As Integer, ByVal RowIndex As Long, CellTip As String, ByVal FullyDisplayed As Boolean, ByVal TipStyle As TrueOleDBGrid80.StyleDisp)
' Display Cell Tip for the "Building Description" column
' ADDED 6/16/2005 RTD FOR VERSION 7.4.0
    
    If ColIndex >= 0 And ColIndex < TDBGridBuilding.Columns.Count Then
        If TDBGridBuilding.Columns(ColIndex).DataField <> "bldg_desc" Then
            CellTip = ""
        End If
    Else
        CellTip = ""
    End If
    
End Sub

Private Sub txtbldg_desc_GotFocus()
    HiliteTextBox txtbldg_desc
End Sub

Private Sub txtbldg_id_GotFocus()
    HiliteTextBox txtbldg_id
End Sub
'
'   Handles Row Wrap feature.
Private Sub ckbRowWrap_Click()
    m_objGridMap.RowWrap (ckbRowWrap)
End Sub

Private Sub cmdOutput_Click()
    DoOutput
End Sub

Private Sub cmdBuilding_Click()
    Dim frm     As frmBuilding
    Dim rec     As ADODB.RecordSet

    On Error Resume Next
    If IsNumeric(TDBGridBuilding.Bookmark) = False Then
        MsgBox "Please select a row.", vbCritical
    Else
        '
        '   Make copy of recordset, using the gridmap NOT 'm_rec.Clone'
        '   so that if they have changed values and not updated the recordset
        '   we pass to the form will contain the original values.
        '
        Set rec = m_objGridMap.CloneRowRecordset

        If Not rec.EOF Then
            Set frm = New frmBuilding
            '
            '   Pass the current record into the form,
            '   Navigating to single-record view.
            With frm
                .SetRow Trim(m_rec.Fields("bldg_id").Value)
                .Show
            End With
        End If
    End If
End Sub

Private Sub cmdUpdate_Click()
    Dim varBookmark As Variant
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    Status ("Updating Building Details ...")
    m_blnWereErrors = False
    varBookmark = TDBGridBuilding.Bookmark
    TDBGridBuilding.Update
    
    With m_objGridMap
        .Update
        Screen.MousePointer = vbNormal
        If .UpdateErrors = 0 Then
            Status ("Building Details Updated Successfully ...")
            MsgBox .SuccessfulUpdates & " rows were updated successfully."
            cmdSearch_Click
        Else
            Status ("")
            MsgBox .SuccessfulUpdates & " rows were updated successfully." _
                    & vbCrLf & .UpdateErrors & " errors were received."
        End If
    End With
    TDBGridBuilding.Bookmark = varBookmark
    cmdUpdate.Enabled = False
    Status ("")
End Sub

Private Sub cmdClone_Click()
    Dim frm         As New frmBuilding
    Dim strSelect   As String
    Dim sError      As String
    Dim recTemp     As New ADODB.RecordSet
    Dim cnTemp      As New ADODB.Connection

    With TDBGridBuilding
        If IsNumeric(TDBGridBuilding.Bookmark) = False Then
            MsgBox "You must select a row.", vbCritical
        ElseIf .Columns("Bldg ID").Value = "100" Or .Columns("Bldg ID").Value = "200" _
        Or .Columns("Bldg ID").Value = "300" Or .Columns("Bldg ID").Value = "400" Then
            MsgBox "Residential Quality Series buildings cannot be cloned.", vbCritical
        Else
            Status ("Cloning Building: " & Trim(.Columns("Bldg ID").Value) & " ...")
            Screen.MousePointer = vbHourglass
            With cnTemp
                .ConnectionTimeout = 50000
                .CommandTimeout = 50000
                '.Open "UID=" + strUserName + ";PWD=;DATABASE=" + strConnectDatabase + ";SERVER=" + strConnectServer + ";DRIVER={SQL SERVER};DSN='';"
                .Open strConnect
                .BeginTrans
            End With

            strSelect = "exec sp_clone_building @NewBldg_skey = '',"
            strSelect = strSelect & "@bldg_id = '" & .Columns("Bldg ID").Value & "',"
            strSelect = strSelect & "@last_update_person = '" & strUserName & "'"
                      
            recTemp.CursorLocation = adUseClient
            recTemp.Open _
                Source:=strSelect, _
                ActiveConnection:=cnTemp, _
                CursorType:=adOpenStatic, _
                LockType:=adLockBatchOptimistic
            
            If cnTemp.Errors.Count <> 0 Then
                Screen.MousePointer = vbNormal
                MsgBox "Error cloning building model: " _
                    & .Columns("Bldg ID").Value _
                    & vbCrLf & cnTemp.Errors(0).Description, vbCritical
                Status ("")
            
            ElseIf recTemp.RecordCount = 0 Then
                Screen.MousePointer = vbNormal
                MsgBox "Error cloning building model: " & .Columns("Bldg ID").Value, vbCritical
                Status ("")
            Else
                cnTemp.CommitTrans
                '
                '   Pass the current record into the form,
                '   Navigating to single-record view.
                With frm
                    If .SetRow(Trim(recTemp.Fields("bldg_id").Value), True) Then
                        .Show
                    Else
                        frm.Visible = False
                        Set frm = Nothing
                    End If
                End With
            End If
        End If
    End With
    recTemp.Close
    cnTemp.Close
    Set cnTemp = Nothing
    Screen.MousePointer = vbNormal
End Sub

Private Sub cmdNew_Click()
    Dim rec     As New ADODB.RecordSet
    Dim frm     As frmBuilding

    On Error Resume Next
    CopyRSFields rec, m_rec
    '
    '   Open empty single record view
    Set frm = New frmBuilding
    '
    '   Force any changes into recordset from grid
    TDBGridBuilding.Update
    With frm
        '
        '   Pass X as the bldg_id so that query won't return results
        '   just empty fields for new building.
        .SetRow "", True, rec
        .Show
    End With
End Sub

Private Sub cmdSearch_Click()
    Dim strSelect               As String
    Dim dtmToday                As Date
    Dim dtmStart                As Date
    Dim strError                As String
    Dim Button                  As String
    Dim i                       As Integer
    
    On Error Resume Next
    If m_objGridMap.IsPendingChange = True Then
        Button = MsgBox("Do you want to save your changes to " & Me.Caption & "?", vbYesNoCancel, "Search For New Building")
        If Button = vbYes Then
            cmdUpdate_Click
            '
            '   If there were errors, cancel the search
            If m_blnWereErrors Then
                Exit Sub
            End If
        ElseIf Button = vbCancel Then
            '
            ' Cancel the search
            Exit Sub
        Else
            TDBGridBuilding.DataChanged = False
            cmdUpdate.Enabled = False
        End If
    End If
    Screen.MousePointer = vbHourglass
    dtmToday = Date
    With lblRowCount
        .Caption = "Working..."
        .Refresh
    End With
    '
    '   Make sure it is closed.
    With m_rec
        .Close
        '
        '   Set the maximum number to bring back.
        .MaxRecords = MAX_RECORDS
    End With
    dtmStart = Now

    strSelect = "exec sp_select_building @type_code = '"
    If opttype_codeB.Value = True Then
        strSelect = strSelect & "%"
    ElseIf opttype_codeC.Value = True Then
        strSelect = strSelect & "C"
    Else
        strSelect = strSelect & "R"
    End If
    
    strSelect = strSelect & "', @bldg_category = '"
    If Len(Trim(bldg_category.Text)) = 0 Or Trim(bldg_category.Text) = "ALL" Then
        strSelect = strSelect & "%"
    Else
        strSelect = strSelect & SQLChangeWildcard(bldg_category.Text)
    End If
    
    strSelect = strSelect & "', @bldg_id = '"
    If Len(Trim(txtbldg_id.Text)) > 0 Then
        strSelect = strSelect & SQLChangeWildcard(txtbldg_id.Text)
    Else
        strSelect = strSelect & "%"
    End If
    strSelect = strSelect & "', @bldg_desc = '"
    
    If Len(Trim(txtbldg_desc.Text)) > 0 Then
        '
        '   Never know if we'll have an apos ' in our desc.
        strSelect = strSelect & SQLChangeWildcard(Replace(txtbldg_desc.Text, "'", "''")) & "'"
    Else
        strSelect = strSelect & "%'"
    End If
    '
    '   Use DAL to perform select.
    If Not g_objDAL.GetRecordset(vbNullString, strSelect, m_rec) Then
        Screen.MousePointer = vbNormal
        MsgBox "An error occurred while searching for building(s)."
        lblRowCount.Caption = "0 rows returned."
    Else
        '
        '   Pass recordset to handler class.
        m_objGridMap.RecordSet = m_rec
        '
        '   Need to make sure that the user cannot set
        '   max_records = 0
        With m_rec
            If m_rec.RecordCount > 0 Then
                lblRowCount.Caption = str(.RecordCount) & " rows returned in " & _
                                        str(DateDiff("s", dtmStart, Now)) + " seconds"
                '
                ' If the upper bound was hit, inform user.
                If .RecordCount = MAX_RECORDS And .State = adStateOpen Then
                    MsgBox "The search returned the maximum number of records allowed. More records may be available."
                End If
                '
                '   If we got only 1 record set the description and bldg_id
                '   equal to that record.
                If .RecordCount = 1 Then
                    txtbldg_id.Text = Trim(.Fields("bldg_id").Value)
                    txtbldg_desc.Text = Trim(.Fields("bldg_desc").Value)
                    For i = 0 To bldg_category.listcount - 1
                        If bldg_category.List(i) = Trim(.Fields("bldg_category").Value) Then
                            bldg_category.ListIndex = i
                            Exit For
                        End If
                    Next i
                End If
            Else
                lblRowCount.Caption = "0 rows returned."
            End If
        End With
        DoEvents
        '
        '   Reset the grid contents
        With TDBGridBuilding
            .Bookmark = Null
            .ReBind
            .ApproxCount = m_rec.RecordCount
        End With
        
        SetButtons USEBOOKMARK
        Screen.MousePointer = vbNormal
    End If
    
End Sub
'
'   Leaf in MasterFormat tree selected.  So populate the grid
'   based upon the bldg_id selected.
Private Sub FormatTree_NodeSelected(ByVal strID As String)
    Dim rs As New ADODB.RecordSet
    Dim strSelect As String
    Dim blnReturn As Boolean
    Dim i As Integer
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    '
    '   Clear other fields so won't search on.
    txtbldg_desc.Text = ""
    txtbldg_id.Text = ""
    bldg_category.ListIndex = -1
    opttype_codeB.Value = True
    Select Case strID
        Case "C"
            opttype_codeC.Value = True
            
        Case "R"
            opttype_codeR.Value = True
            
        Case "Commercial", "Institutional", "Industrial"
            For i = 0 To bldg_category.listcount - 1
                If bldg_category.List(i) = strID Then
                    bldg_category.ListIndex = i
                    Exit For
                End If
            Next i
            opttype_codeC.Value = True
            
        Case "Luxury", "Economy", "Custom", "Average"
            For i = 0 To bldg_category.listcount - 1
                If bldg_category.List(i) = strID Then
                    bldg_category.ListIndex = i
                    Exit For
                End If
            Next i
            opttype_codeR.Value = True
        
        Case "ALL", "op"
             For i = 0 To bldg_category.listcount - 1
                If bldg_category.List(i) = strID Then
                    bldg_category.ListIndex = i
                    Exit For
                End If
            Next i
            opttype_codeB.Value = True
       
        Case Else
            '
            '   Synch text box with tree.
            If Len(Trim(strID)) = 1 Then
                txtbldg_id.Text = strID + "*"
            Else
                txtbldg_id.Text = strID
            End If
    End Select

    Screen.MousePointer = vbNormal
    '
    '   Kick-off search.
    cmdSearch_Click
End Sub

Private Sub cmdModels_Click()
    '
    '   Open single record view with data from row selected.
    Dim frm As frmModelGrid
    
    Set frm = New frmModelGrid
    With TDBGridBuilding
        frm.JumpIn Trim(.Columns("Bldg ID").CellText(.Bookmark))
    End With
End Sub

Private Sub TDBGridBuilding_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    '
    '   Populate the dlgOutput form based upon the currently
    '   selected row, only if the user moved to a new row.
    position_output
End Sub

Private Sub TDBGridBuilding_GotFocus()
    TDBGridBuilding.TabStop = True
End Sub

Private Sub TDBGridBuilding_LostFocus()
    TDBGridBuilding.TabStop = False
End Sub

Private Sub TDBGridBuilding_KeyUp(KeyCode As Integer, Shift As Integer)
    SetButtons USEBOOKMARK
End Sub

Private Sub TDBGridBuilding_DblClick()
    '
    ' Same function as clicking Building button, open single record view
    cmdBuilding_Click
End Sub

Private Sub TDBGridBuilding_Error(ByVal DataError As Integer, Response As Integer)
    Response = 0
    TDBGridBuilding.DataChanged = False
End Sub

Private Sub TDBGridBuilding_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With TDBGridBuilding
        If Button = vbRightButton And IsNumeric(.Bookmark) Then
            If Len(m_objGridMap.GetError(.Bookmark)) > 0 Then
                MsgBox m_objGridMap.GetError(.Bookmark)
            End If
        End If
    End With
    SetButtons USEBOOKMARK
End Sub
'
'   Can't use AfterUpdate since it never fires if you can't move to another row!
Private Sub TDBGridBuilding_AfterColUpdate(ByVal ColIndex As Integer)
    cmdUpdate.Enabled = True
End Sub

Public Function PrintReport()
    PreviewReport
End Function

Public Function PreviewReport()
    Dim fPreviewWindow As New frmReportPreview
    
    If m_rec.RecordCount > 0 Then
        fPreviewWindow.ReportName = "Buildings"
        fPreviewWindow.ReportFile = "rptSummaryEstimate.xml"
        fPreviewWindow.RecordSet = m_rec
        fPreviewWindow.RenderReport
        fPreviewWindow.Show
    Else
        MsgBox "You must display the records you want to report using the Search feature.", vbInformation
    End If
End Function

Private Sub ShowToolbarIcons(bShowIcons As Boolean)

    fMainForm.tbToolBar.Buttons.Item(tbrPRINT).Enabled = bShowIcons
    fMainForm.tbToolBar.Buttons.Item(tbrPRINT).Visible = bShowIcons
    fMainForm.tbToolBar.Buttons.Item(tbrPREVIEW).Enabled = bShowIcons
    fMainForm.tbToolBar.Buttons.Item(tbrPREVIEW).Visible = bShowIcons
    fMainForm.mnuFilePageSetup.Enabled = bShowIcons
    fMainForm.mnuFilePrint.Enabled = bShowIcons
    fMainForm.mnuFilePrintPreview.Enabled = bShowIcons

End Sub

