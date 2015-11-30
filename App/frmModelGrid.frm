VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{5936A75C-3F42-11D6-AF6B-AA0004005F12}#1.3#0"; "MeansCtrl.ocx"
Begin VB.Form frmModelGrid 
   Caption         =   "Model Grid"
   ClientHeight    =   7275
   ClientLeft      =   1500
   ClientTop       =   2790
   ClientWidth     =   11685
   Icon            =   "frmModelGrid.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7275
   ScaleWidth      =   11685
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7200
      TabIndex        =   24
      Top             =   6600
      Width           =   915
   End
   Begin VB.ComboBox cbobldg_category 
      Height          =   315
      ItemData        =   "frmModelGrid.frx":0442
      Left            =   9495
      List            =   "frmModelGrid.frx":0452
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   1200
      Width           =   2040
   End
   Begin VB.TextBox txtwall_type 
      Height          =   285
      Left            =   7860
      TabIndex        =   3
      Top             =   2280
      Width           =   3675
   End
   Begin VB.TextBox txtframe_type 
      Height          =   285
      Left            =   7860
      TabIndex        =   2
      Top             =   1920
      Width           =   3675
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   495
      Left            =   8280
      TabIndex        =   9
      Top             =   6600
      Width           =   915
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   495
      Left            =   9360
      TabIndex        =   10
      Top             =   6600
      Width           =   915
   End
   Begin VB.CommandButton cmdClone 
      Caption         =   "&Clone"
      Height          =   495
      Left            =   10440
      TabIndex        =   11
      Top             =   6600
      Width           =   915
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   375
      Left            =   7860
      TabIndex        =   4
      Top             =   2640
      Width           =   1150
   End
   Begin VB.TextBox txtbldg_id 
      Height          =   285
      Left            =   7860
      MaxLength       =   4
      TabIndex        =   0
      Top             =   1200
      Width           =   630
   End
   Begin VB.TextBox txtbldg_desc 
      Height          =   285
      Left            =   7860
      TabIndex        =   1
      Top             =   1560
      Width           =   3675
   End
   Begin VB.Frame fraModelType 
      Caption         =   "Model Type"
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
      Left            =   7080
      TabIndex        =   14
      Top             =   480
      Width           =   4455
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   4095
         TabIndex        =   25
         Top             =   240
         Width           =   4095
         Begin VB.OptionButton opttype_codeR 
            Caption         =   "Residential"
            Height          =   255
            Left            =   1560
            TabIndex        =   28
            Top             =   0
            Width           =   1095
         End
         Begin VB.OptionButton opttype_codeC 
            Caption         =   "Commercial"
            Height          =   255
            Left            =   0
            TabIndex        =   27
            Top             =   0
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton opttype_codeB 
            Caption         =   "Both"
            Height          =   255
            Left            =   3000
            TabIndex        =   26
            Top             =   0
            Width           =   855
         End
      End
   End
   Begin ConstructionCostDatabase.DynaTree FormatTree 
      Height          =   3135
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   5530
   End
   Begin VB.Frame fraGoTo 
      Caption         =   "Go To"
      ForeColor       =   &H00404040&
      Height          =   855
      Left            =   30
      TabIndex        =   15
      Top             =   6360
      Width           =   3255
      Begin VB.CommandButton cmdBuildingGrid 
         Caption         =   "Buildings"
         Height          =   495
         Left            =   1200
         TabIndex        =   7
         Top             =   240
         Width           =   915
      End
      Begin VB.CommandButton cmdModel 
         BackColor       =   &H80000001&
         Caption         =   "&Model"
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   915
      End
      Begin VB.CommandButton cmdBuilding 
         Caption         =   "&Building Maint."
         Height          =   495
         Left            =   2160
         TabIndex        =   8
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.CheckBox ckbRowWrap 
      Caption         =   "Row Wrap"
      Height          =   315
      Left            =   90
      TabIndex        =   5
      Top             =   3240
      Width           =   1215
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGridModels 
      Height          =   2715
      Left            =   30
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3600
      Width           =   11325
      _ExtentX        =   19976
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
      _StyleDefs(45)  =   ":id=29,.parent=0,.bgcolor=&HFFFFFF&"
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
      Left            =   8760
      TabIndex        =   23
      Top             =   1245
      Width           =   675
   End
   Begin VB.Label lblwall_type 
      Alignment       =   1  'Right Justify
      Caption         =   "Wall Type:"
      Height          =   255
      Left            =   6915
      TabIndex        =   21
      Top             =   2325
      Width           =   900
   End
   Begin VB.Label lblframe_type 
      Alignment       =   1  'Right Justify
      Caption         =   "Frame Type:"
      Height          =   255
      Left            =   6915
      TabIndex        =   20
      Top             =   1980
      Width           =   900
   End
   Begin VB.Label Label3 
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
      Left            =   6840
      TabIndex        =   19
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label lblBuildingId 
      Alignment       =   1  'Right Justify
      Caption         =   "Building ID:"
      Height          =   255
      Left            =   6915
      TabIndex        =   18
      Top             =   1245
      Width           =   900
   End
   Begin VB.Label lblBuildingDesc 
      Alignment       =   1  'Right Justify
      Caption         =   "Description:"
      Height          =   255
      Left            =   6915
      TabIndex        =   17
      Top             =   1605
      Width           =   900
   End
   Begin VB.Label lblRowCount 
      Alignment       =   2  'Center
      Caption         =   "0 rows returned"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   4080
      TabIndex        =   16
      Top             =   3240
      Width           =   6510
   End
   Begin VB.Line Line2 
      X1              =   60
      X2              =   11340
      Y1              =   3180
      Y2              =   3180
   End
   Begin VB.Line Line1 
      X1              =   6660
      X2              =   6660
      Y1              =   3120
      Y2              =   60
   End
End
Attribute VB_Name = "frmModelGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'   Class to handle grid
Dim m_objGridMap As New CModelMap
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
Dim sEventSubscriberID As String
'
'   Used to detect form load so search doesn't slow down initial load
Dim bIsInitialLoad As Boolean

Const USEBOOKMARK = 1
Const USECOORD = 0

Private Sub cmdBuildingGrid_Click()
    '
    '   Open grid view with data from row selected.
    Dim frm As frmBuildingGrid
    
    Set frm = New frmBuildingGrid
    With TDBGridModels
        frm.JumpIn Trim(.Columns("Bldg ID").CellText(.Bookmark))
    End With
    
End Sub

Private Sub Form_Load()

    On Error Resume Next
    Move START_LEFT, START_TOP, START_WIDTH, START_HEIGHT
    '
    '   This will never return any rows, just used to create recordset????
    cmdSearch_Click
    bIsInitialLoad = False
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim Button  As String
    
    On Error Resume Next
    '
    '   Check if there are pending changes
    If m_objGridMap.IsPendingChange = True Then
        Button = MsgBox("Do you want to save your changes to " & Me.Caption & "?", vbYesNoCancel, "Close ModelGrid Form")
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
    bIsInitialLoad = True
    Status ("Loading Model Maintenance Grid ...")
    sEventSubscriberID = EventSubscriberAdd(Me)
    '
    '   Fill the MasterFormat tree.
    FormatTree.InitData g_cnShared, "BUILDING"
    '
    '   Initialize grid.
    With m_objGridMap
        .SetGrid TDBGridModels
        .InitGrid
    End With
    PopulateBldgCategories
    TDBGridModels.DataChanged = False
    Status ("")
    Screen.MousePointer = vbNormal
    
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
        OutputView True
        ShowGridSort
        m_objGridMap.SetMenuBar
        ShowToolbarIcons True
    End If

End Sub

Private Sub Form_Deactivate()
    m_strCurrentFormControl = Me.ActiveControl.Name
    ShowToolbarIcons False
End Sub

Private Sub Form_LostFocus()
    TDBGridModels.Update
    HideGridSort
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    '
    '   Need to place in common routine for all forms.
    '   Possibly place all buttons in a frame like frame1 with
    '   common name and can just place it.
    If Me.WindowState = vbNormal Or Me.WindowState = vbMaximized Then
        If Me.Width >= 11500 Then
            TDBGridModels.Width = Me.Width - 210
            Line2.X2 = Me.Width - 210
        Else
            Me.Width = 11500
        End If
        
        If Me.Height >= 7135 Then
            TDBGridModels.Height = Me.Height - 4920
            fraGoTo.Top = Me.Height - 1275
            cmdUpdate.Top = Me.Height - 1035
            cmdNew.Top = Me.Height - 1035
            cmdClone.Top = Me.Height - 1035
            cmdDelete.Top = Me.Height - 1035
        Else
            Me.Height = 7135
        End If
    Else
        ShowMinimizedForms
    End If
End Sub

Public Sub EventNotify(eNotifyType As EEventSubscriberNotifyType, sAffectedRecordIdentifier As String)
    Dim varBookmark
    
    On Error Resume Next
    '
    '   If the record that was updated is in our grid results
    '   we need to refresh.
    If eNotifyType = esnBuildingRecordUpdated Or eNotifyType = esnModelRecordUpdated Then
        If txtbldg_id.Text = "" Or Trim(txtbldg_id.Text) = Trim(sAffectedRecordIdentifier) Then
            '
            '   Need to clear fields that could have been updated if the bldg_id matches
            '   the bldg_id we updated.
            If Trim(txtbldg_id.Text) = Trim(sAffectedRecordIdentifier) Then
                opttype_codeB.Value = True
                txtbldg_desc.Text = ""
                txtframe_type.Text = ""
                txtwall_type.Text = ""
            End If
            varBookmark = TDBGridModels.Bookmark
            cmdSearch_Click
            TDBGridModels.Bookmark = varBookmark
            '
            '   Fill the MasterFormat tree.
            With FormatTree
                .ClearTree
                .InitData g_cnShared, "BUILDING"
                .FocusItem "0"
            End With
        End If
    End If
End Sub

Private Sub SetButtons(Mode As Single, Optional Coord As Variant)
    On Error Resume Next
    '
    '   No current record - disable buttons
    If m_rec.RecordCount > 0 And TDBGridModels.Bookmark >= 1 Then
        Select Case Mode
            Case USEBOOKMARK
                m_rec.Bookmark = TDBGridModels.Bookmark
            Case USECOORD
                m_rec.Bookmark = TDBGridModels.RowBookmark(TDBGridModels.RowContaining(Coord))
        End Select
        
        If IsNumeric(m_rec.Bookmark) Then
            cmdModel.Enabled = True
            cmdBuilding.Enabled = True
            cmdClone.Enabled = True
            cmdDelete.Enabled = True
            cmdNew.Enabled = True
            '
            '   Don't set update unless there has been a change in the grid.
            'cmdUpdate.Enabled = True
        Else
            cmdModel.Enabled = False
            cmdBuilding.Enabled = False
            cmdClone.Enabled = False
            cmdDelete.Enabled = False
            cmdNew.Enabled = False
            cmdUpdate.Enabled = False
        End If
    Else
            cmdModel.Enabled = False
            cmdBuilding.Enabled = False
            cmdClone.Enabled = False
            cmdDelete.Enabled = False
            cmdNew.Enabled = False
            cmdUpdate.Enabled = False
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
Public Sub JumpIn(strBldgID As String)
    txtbldg_id.Text = Trim(strBldgID)
    txtbldg_desc.Text = ""
    opttype_codeB.Value = True
    cmdSearch_Click
End Sub

Private Sub txtbldg_desc_GotFocus()
    HiliteTextBox txtbldg_desc
End Sub

Private Sub txtbldg_id_GotFocus()
    HiliteTextBox txtbldg_id
End Sub

Private Sub txtframe_type_GotFocus()
    HiliteTextBox txtframe_type
End Sub

Private Sub txtwall_type_GotFocus()
    HiliteTextBox txtwall_type
End Sub
'
'   Handles Row Wrap feature.
Private Sub ckbRowWrap_Click()
    m_objGridMap.RowWrap (ckbRowWrap)
End Sub

Private Sub cmdDelete_Click()
    Dim Button          As String
    Dim strUpdate       As String
    Dim sModelCode      As String
    Dim sBldgID         As String
    Dim cnTemp          As New ADODB.Connection
    Dim cmdTemp         As New ADODB.Command
    
    On Error GoTo errorHandler:
    With TDBGridModels
        sModelCode = Trim(.Columns("model_code").Value)
        sBldgID = Trim(.Columns("Bldg ID").Value)
        
        If IsNumeric(.Bookmark) = False Then
            MsgBox "You must select a row.", vbCritical
        ElseIf sModelCode = "7" Or sModelCode = "8" Then
            MsgBox "Model codes 7 & 8 cannot be deleted.", vbCritical
        ElseIf .SelBookmarks.Count > 1 Then
            MsgBox "Only 1 model can be deleted at a time.  Please select only 1 row.", vbCritical
        Else
          Button = MsgBox("Are you sure you want to delete Model: " & sModelCode & " for building ID: " & sBldgID & "?", vbYesNo + vbCritical)
          If Button = vbYes Then
                Screen.MousePointer = vbHourglass
                Status ("Deleting Model: " & sModelCode & " For Building ID: " & sBldgID & " ...")
                    
                With cnTemp
                    .ConnectionTimeout = 50000
                    .CommandTimeout = 50000
                    '.Open "UID=" + strUserName + ";PWD=;DATABASE=" + strConnectDatabase + ";SERVER=" + strConnectServer + ";DRIVER={SQL SERVER};DSN='';"
                    .Open strConnect
                    Set cmdTemp = New ADODB.Command
                    Set cmdTemp.ActiveConnection = cnTemp
                End With
                .Delete
                strUpdate = "exec sp_delete_bldg_model @bldg_model_skey = '" & Trim(.Columns("Mdl Skey").Value) & "'"
              
                With cmdTemp
                    .CommandTimeout = 50000
                    .CommandType = adCmdText
                    .CommandText = strUpdate
                    .Execute adExecuteNoRecords
                End With
                
                If cnTemp.Errors.Count <> 0 Then
                    Screen.MousePointer = vbNormal
                    MsgBox "Error deleting building model: " _
                        & sModelCode & " for building ID: " & sBldgID _
                        & vbCrLf & cnTemp.Errors(0).Description, vbCritical
                    Status ("")
                Else
                    Status ("")
                    '
                    '   Always refresh forms that are listening for changes in case part of the update succeeded.
                    '   ie -the bldg was updated and the grid has an old last_update_id.
                    EventSubscriberNotify esnModelRecordUpdated, sBldgID
                    cmdSearch_Click
                End If
          End If
        End If
    End With
    Screen.MousePointer = vbNormal
    Exit Sub
    
errorHandler:
    Screen.MousePointer = vbNormal
    MsgBox "Errors deleting building model: " & sModelCode _
        & " for building ID: " & sBldgID & Err.Description, vbCritical
    Status ("")
End Sub

Private Sub cmdBuilding_Click()
    '
    '   Open single record view with data from row selected.
    Dim frm As frmBuilding
    
    Set frm = New frmBuilding
    With TDBGridModels
        frm.SetRow Trim(.Columns("Bldg ID").CellText(.Bookmark))
        frm.Show
    End With

End Sub

Private Sub cmdClone_Click()
    Dim frm                 As New frmModel
    Dim strSelect           As String
    Dim sError              As String
    Dim sbldg_model_skey    As String
    Dim bOK                 As Boolean
    Dim recTemp             As New ADODB.RecordSet
    Dim cnTemp              As New ADODB.Connection
    Dim cmdTemp             As New ADODB.Command

    On Error Resume Next
    With TDBGridModels
        If IsNumeric(.Bookmark) = False Then
            MsgBox "You must select a row.", vbCritical
        ElseIf .Columns("Model Code").Value = "7" Or .Columns("Model Code").Value = "8" Then
            MsgBox "Model codes 7 & 8 cannot be cloned.", vbCritical
        Else
            Status ("Cloning Model: " & Trim(.Columns("Model Code").Value) & " For Building ID: " & Trim(.Columns("Bldg ID").Value) & " ...")
            Screen.MousePointer = vbHourglass
            strSelect = "SELECT bldg_model_skey FROM bldg_model WHERE bldg_skey = '" & Trim(TDBGridModels.Columns("bldg skey").Value) & "' AND model_code != '7' AND model_code != '8'"
    
            If Not g_objDAL.GetRecordset(vbNullString, strSelect, recTemp) Then
                Screen.MousePointer = vbNormal
                MsgBox "An error occurred while searching for total number of models prior to cloning."
            ElseIf Trim(.Columns("Type").Value) = "C" And recTemp.RecordCount >= 6 Then
                MsgBox "Maximum number of models per Commercial building is 6.  Currently this building has " & recTemp.RecordCount & " models.  " & vbCrLf & _
                        "Inserting 1 additional model will exceed the maximum. ", vbCritical
            ElseIf Trim(.Columns("Type").Value) = "R" And recTemp.RecordCount >= 4 Then
                MsgBox "Maximum number of models per Residential building is 4.  Currently this building has " & recTemp.RecordCount & " models.  " & vbCrLf & _
                        "Inserting 1 additional model will exceed the maximum. ", vbCritical
            Else
                With cnTemp
                    .ConnectionTimeout = 50000
                    .CommandTimeout = 0
                    '.Open "UID=" + strUserName + ";PWD=;DATABASE=" + strConnectDatabase + ";SERVER=" + strConnectServer + ";DRIVER={SQL SERVER};DSN='';"
                    .Open strConnect
                End With
                Set cmdTemp = New ADODB.Command
                Set cmdTemp.ActiveConnection = cnTemp

                strSelect = "exec sp_clone_model @OldBldg_model_skey = '"
                strSelect = strSelect & Trim(.Columns("Mdl Skey").Value) & "',"
                strSelect = strSelect & "@last_update_person = '" & strUserName & "'"
                
                With cmdTemp
                    .CommandTimeout = 80000
                    .CommandType = adCmdText
                    .CommandText = strSelect
                    .Execute adExecuteNoRecords
                End With
    
                If cnTemp.Errors.Count <> 0 Then
                    Screen.MousePointer = vbNormal
                    MsgBox "Error cloning model: " _
                        & Trim(.Columns("Mdl Skey").Value) _
                        & vbCrLf & cnTemp.Errors(0).Description, vbCritical
                    Status ("")
                Else
                    recTemp.Close
                    cnTemp.Errors.Clear
                    strSelect = "SELECT bldg_model_skey FROM bldg_model WHERE bldg_skey = '" _
                        & Trim(.Columns("Bldg Skey").Value) & "' " _
                        & "AND model_code = (SELECT max(model_code) FROM bldg_model WHERE " _
                        & "bldg_skey = '" & Trim(.Columns("Bldg Skey").Value) & "' " _
                        & "AND model_code != '7' AND model_code != '8')" _
                        & "AND bldg_model_skey != '" & Trim(.Columns("Mdl Skey").Value) & "' " _
                        & "AND last_update_person = '" & strUserName & "'"
                    
                    recTemp.CursorLocation = adUseClient
                    recTemp.Open _
                        Source:=strSelect, _
                        ActiveConnection:=cnTemp, _
                        CursorType:=adOpenStatic, _
                        LockType:=adLockBatchOptimistic
        
                    If cnTemp.Errors.Count <> 0 Then
                        If cnTemp.Errors(0).Number <> "0" Then
                            Screen.MousePointer = vbNormal
                            bOK = False
                            MsgBox "Error locating new model" _
                                & vbCrLf & cnTemp.Errors(0).Description, vbCritical
                            Status ("")
                        Else
                            bOK = True
                        End If
                    Else
                        bOK = True
                    End If
                    
                    If bOK Then
                        If Not recTemp.EOF Then
                            sbldg_model_skey = recTemp.Fields("bldg_model_skey").Value
                            recTemp.Close
                            strSelect = "exec sp_select_model @type_code = '%',@bldg_category = '%'," _
                                & "@bldg_id = '%',@bldg_desc = '%',@frame_type = '%',@wall_type = '%'," _
                                & "@bldg_model_skey = '" & sbldg_model_skey & "'"
                            
                            recTemp.CursorLocation = adUseClient
                            recTemp.Open _
                                Source:=strSelect, _
                                ActiveConnection:=cnTemp, _
                                CursorType:=adOpenStatic, _
                                LockType:=adLockBatchOptimistic

                            If cnTemp.Errors.Count <> 0 Then
                                If cnTemp.Errors(0).Number <> "0" Then
                                    Screen.MousePointer = vbNormal
                                    bOK = False
                                    MsgBox "Error searching for new model."
                                Else
                                    bOK = True
                                End If
                            Else
                                bOK = True
                            End If
                            If bOK Then
                                '
                                '   Pass the current record into the form,
                                '   Navigating to single-record view.
                                '   Default to Open op_code if we're on Residential.
                                With frm
                                    .SetRow recTemp, True, , IIf(Trim(TDBGridModels.Columns("type").Value) = "R", "Open", "Union")
                                    .Show
                                End With
                                '
                                '   Always refresh forms that are listening for changes in case part of the update succeeded.
                                EventSubscriberNotify esnModelRecordUpdated, Trim(.Columns("Bldg ID").Value)
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Screen.MousePointer = vbNormal
    End With
End Sub

Private Sub cmdUpdate_Click()
    Dim varBookmark As Variant
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    Status ("Updating Model Details ...")
    m_blnWereErrors = False
    varBookmark = TDBGridModels.Bookmark
    TDBGridModels.Update
    
    With m_objGridMap
        .Update
        Screen.MousePointer = vbNormal
        If .UpdateErrors = 0 Then
            Status ("Model Details Updated Successfully ...")
            MsgBox .SuccessfulUpdates & " rows were updated successfully."
            cmdSearch_Click
        Else
            Status ("")
            MsgBox .SuccessfulUpdates & " rows were updated successfully." _
                    & vbCrLf & .UpdateErrors & " errors were received."
        End If
    End With
    TDBGridModels.Bookmark = varBookmark
    cmdUpdate.Enabled = False
    Status ("")
End Sub

Private Sub cmdNew_Click()
    Dim recTemp     As New ADODB.RecordSet
    Dim rec         As New ADODB.RecordSet
    Dim strSelect   As String
    Dim frm         As frmModel

    On Error Resume Next
    With TDBGridModels
        '
        '   If this is a Residential Quality Series building then use a different sp to get
        '   results since those buildings do not have records in published_bldg_matrix_cost.
        If txtbldg_id.Text = "100" Or txtbldg_id.Text = "200" _
        Or txtbldg_id.Text = "300" Or txtbldg_id.Text = "400" Then
            MsgBox "Models cannot be added to Residential Quality Series buildings.", vbCritical
        Else
            Screen.MousePointer = vbHourglass
            strSelect = "SELECT bldg_model_skey FROM bldg_model WHERE bldg_skey = '" _
                & Trim(TDBGridModels.Columns("bldg skey").Value) _
                & "' AND model_code != '7' AND model_code != '8'"
                       
            If Not g_objDAL.GetRecordset(vbNullString, strSelect, recTemp) Then
                Screen.MousePointer = vbNormal
                MsgBox "An error occurred while searching for total number of models prior to inserting.", vbCritical
            ElseIf Trim(.Columns("Type").Value) = "C" And recTemp.RecordCount >= 6 Then
                Screen.MousePointer = vbNormal
                MsgBox "Maximum number of models per Commercial building is 6.  Currently this building has " & recTemp.RecordCount & " models.  " & vbCrLf & _
                        "Inserting 1 additional model will exceed the maximum. ", vbCritical
            ElseIf Trim(.Columns("Type").Value) = "R" And recTemp.RecordCount >= 4 Then
                Screen.MousePointer = vbNormal
                MsgBox "Maximum number of models per Residential building is 4.  Currently this building has " & recTemp.RecordCount & " models.  " & vbCrLf & _
                        "Inserting 1 additional model will exceed the maximum. ", vbCritical
            Else
                
                CopyRSFields rec, m_rec
                '
                '   Open empty single record view
                Set frm = New frmModel
                '
                '   Force any changes into recordset from grid
                .Update
                With frm
                    rec.Fields("bldg_id").Value = Trim(TDBGridModels.Columns("bldg id").Value)
                    If Trim(rec.Fields("bldg_id").Value) <> "" Then
                        .SetRow rec, True
                        .Show
                    End If
                End With
            End If
            recTemp.Close
        End If
    End With
    Screen.MousePointer = vbNormal
End Sub

Private Sub cmdSearch_Click()
    Dim strSelect               As String
    Dim dtmToday                As Date
    Dim dtmStart                As Date
    Dim strError                As String
    Dim Button                  As String
    
    On Error Resume Next
    If m_objGridMap.IsPendingChange = True Then
        Button = MsgBox("Do you want to save your changes to " & Me.Caption & "?", vbYesNoCancel, "Search For New Model")
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
            TDBGridModels.DataChanged = False
            cmdUpdate.Enabled = False
        End If
    End If
    Screen.MousePointer = vbHourglass
    If bIsInitialLoad Then
        txtbldg_id.Text = "000"
    End If
    dtmToday = Date
    '
    '   If it's the first search obviously the text boxes will not
    '   be populated.
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
    '
    '   If this is a Residential Quality Series building then use a different sp to get
    '   results since those buildings do not have records in published_bldg_matrix_cost.
    If txtbldg_id.Text = "100" Or txtbldg_id.Text = "200" _
    Or txtbldg_id.Text = "300" Or txtbldg_id.Text = "400" Then
        strSelect = "exec sp_select_model_basements @bldg_id = '"
        strSelect = strSelect & Trim(txtbldg_id.Text)
        strSelect = strSelect & "', @bldg_model_skey = '%'"
    Else
        strSelect = "exec sp_select_model @type_code = "
        If opttype_codeC.Value = True Then
            strSelect = strSelect & "'C"
        ElseIf opttype_codeR.Value = True Then
            strSelect = strSelect & "'R"
        Else
           strSelect = strSelect & "'%"
        End If
        
        strSelect = strSelect & "', @bldg_category = '"
        If Len(Trim(cbobldg_category.Text)) = 0 Or Trim(cbobldg_category.Text) = "ALL" Then
           strSelect = strSelect & "%"
        Else
           strSelect = strSelect & SQLChangeWildcard(cbobldg_category.Text)
        End If
        
        strSelect = strSelect & "', @bldg_id = '"
        If Len(Trim(txtbldg_id.Text)) > 0 Then
            strSelect = strSelect & SQLChangeWildcard(txtbldg_id.Text)
        Else
            strSelect = strSelect & "%"
        End If
        
        strSelect = strSelect & "', @bldg_desc = '"
        If Len(Trim(txtbldg_desc.Text)) > 0 Then
            strSelect = strSelect & SQLChangeWildcard(Replace(Trim(txtbldg_desc.Text), "'", "''"))
        Else
            strSelect = strSelect & "%"
        End If
        
        strSelect = strSelect & "', @frame_type = '"
        If Len(Trim(txtframe_type.Text)) > 0 Then
            strSelect = strSelect & SQLChangeWildcard(txtframe_type.Text)
        Else
            strSelect = strSelect & "%"
        End If
        
        strSelect = strSelect & "', @wall_type = '"
        If Len(Trim(txtwall_type.Text)) > 0 Then
            strSelect = strSelect & SQLChangeWildcard(txtwall_type.Text)
        Else
            strSelect = strSelect & "%"
        End If
        strSelect = strSelect & "', @bldg_model_skey = '%'"
    End If
    '
    '   Use DAL to perform select.
    If Not g_objDAL.GetRecordset(vbNullString, strSelect, m_rec) Then
        Screen.MousePointer = vbNormal
        MsgBox "An error occurred while searching for model(s)."
        lblRowCount.Caption = "0 rows returned."
    Else
        '
        '   Pass recordset to handler class.
        m_objGridMap.RecordSet = m_rec
        '
        '   Need to make sure that the user cannot set
        '   max_records = 0
       With m_rec
           If .RecordCount > 0 Then
               lblRowCount.Caption = str(.RecordCount) & " rows returned in " & _
                                       str(DateDiff("s", dtmStart, Now)) + " seconds"
               '
               ' If the upper bound was hit, inform user.
               If .RecordCount = MAX_RECORDS And .State = adStateOpen Then
                   MsgBox "The search returned the maximum number of records allowed. More records may be available."
               End If
               '
               '   If we have searched for only 1 bldg then set the description
               '   and bldg_id equal to that record.
               If txtbldg_id.Text = Trim(.Fields("bldg_id").Value) Then
                   txtbldg_id.Text = Trim(.Fields("bldg_id").Value)
                   txtbldg_desc.Text = Trim(.Fields("bldg_desc").Value)
                   If .Fields("type_code") = "C" Then
                       opttype_codeC.Value = True
                   ElseIf .Fields("type_code") = "R" Then
                       opttype_codeR.Value = True
                   End If
               End If
           Else
               lblRowCount.Caption = "0 rows returned."
           End If
        End With
        DoEvents
        '
        '   Reset the grid contents
        With TDBGridModels
            .Bookmark = Null
            .ReBind
            .ApproxCount = m_rec.RecordCount
        End With
        SetButtons USEBOOKMARK
        Screen.MousePointer = vbNormal
    End If
    If bIsInitialLoad Then
        txtbldg_id.Text = ""
    End If
End Sub

Private Sub cmdModel_Click()
    Dim frm     As frmModel
    Dim rec     As ADODB.RecordSet

    On Error Resume Next
    If IsNumeric(TDBGridModels.Bookmark) = False Then
        MsgBox "Please select a row.", vbCritical
    Else
        '
        '   Make copy of recordset, using the gridmap NOT 'm_rec.Clone'
        '   so that if they have changed values and not updated the recordset
        '   we pass to the form will contain the original values.
        '
        Set rec = m_objGridMap.CloneRowRecordset
        
        If Not rec.EOF Then
        
            Set frm = New frmModel
            '
            '   Pass the current record into the form,
            '   Navigating to single-record view.
            With frm
                .SetRow rec, False, , IIf(Trim(TDBGridModels.Columns("type").Value) = "R", "Open", "Union")
                .Show
            End With
        End If
    End If
End Sub

Private Sub PopulateBldgCategories()
    Dim rec         As New ADODB.RecordSet
    Dim strSelect   As String
    
    Screen.MousePointer = vbHourglass
    '
    '   Fill the available categories based on the type code.
    cbobldg_category.Clear
        
    strSelect = "select bldg_category from bldg_category order by bldg_category"
    If Not g_objDAL.GetRecordset(vbNullString, strSelect, rec) Then
        Screen.MousePointer = vbNormal
        MsgBox "An error occurred while searching for available building categories."
    Else
        With rec
            If .RecordCount = 0 Then
                cbobldg_category.AddItem "(unknown)"
            Else
                cbobldg_category.AddItem "ALL"
                While Not .EOF
                    cbobldg_category.AddItem Trim(.Fields("bldg_category").Value)
                    .MoveNext
                Wend
            End If
            .Close
        End With
    End If
    Screen.MousePointer = vbNormal
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
    cbobldg_category.ListIndex = -1
    opttype_codeB.Value = True
    Select Case strID
        Case "C"
            opttype_codeC.Value = True
            
        Case "R"
            opttype_codeR.Value = True
            
        Case "Commercial", "Institutional", "Industrial"
            For i = 0 To cbobldg_category.listcount - 1
                If cbobldg_category.List(i) = strID Then
                    cbobldg_category.ListIndex = i
                    Exit For
                End If
            Next i
            opttype_codeC.Value = True
            
        Case "Luxury", "Economy", "Custom", "Average"
            For i = 0 To cbobldg_category.listcount - 1
                If cbobldg_category.List(i) = strID Then
                    cbobldg_category.ListIndex = i
                    Exit For
                End If
            Next i
            opttype_codeR.Value = True
        
        Case "ALL", "op"
             For i = 0 To cbobldg_category.listcount - 1
                If cbobldg_category.List(i) = strID Then
                    cbobldg_category.ListIndex = i
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

Private Sub TDBGridModels_GotFocus()
    TDBGridModels.TabStop = True
End Sub

Private Sub TDBGridModels_LostFocus()
    TDBGridModels.TabStop = False
End Sub

Private Sub TDBGridModels_KeyUp(KeyCode As Integer, Shift As Integer)
    SetButtons USEBOOKMARK
End Sub

Private Sub TDBGridModels_DblClick()
    '
    ' Same function as clicking Building button, open single record view
    cmdModel_Click
End Sub

Private Sub TDBGridModels_Error(ByVal DataError As Integer, Response As Integer)
    Response = 0
    TDBGridModels.DataChanged = False
End Sub

Private Sub TDBGridModels_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With TDBGridModels
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
Private Sub TDBGridModels_AfterColUpdate(ByVal ColIndex As Integer)
    cmdUpdate.Enabled = True
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

Public Function PrintReport()
    PreviewReport
End Function

Public Function PreviewReport()
    Dim fPreviewWindow As New frmReportPreview
    
    If m_rec.RecordCount > 0 Then
        fPreviewWindow.ReportName = "Models"
        fPreviewWindow.ReportFile = "rptSummaryEstimate.xml"
        fPreviewWindow.RecordSet = m_rec
        fPreviewWindow.RenderReport
        fPreviewWindow.Show
    Else
        MsgBox "You must display the records you want to report using the Search feature.", vbInformation
    End If
End Function
