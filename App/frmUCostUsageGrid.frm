VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{5936A75C-3F42-11D6-AF6B-AA0004005F12}#1.3#0"; "MeansCtrl.ocx"
Begin VB.Form frmUCostUsageGrid 
   Caption         =   "Unit Cost Usage Grid"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11130
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUCostUsageGrid.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6855
   ScaleWidth      =   11130
   Begin VB.ComboBox cboMasterFormat 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8100
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   480
      Width           =   1455
   End
   Begin VB.ListBox lstValidate 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3120
      Sorted          =   -1  'True
      TabIndex        =   20
      Top             =   6120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox StartUnitCostID 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8100
      TabIndex        =   0
      Top             =   1140
      Width           =   1515
   End
   Begin VB.TextBox EndUnitCostID 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9720
      TabIndex        =   1
      Top             =   1140
      Width           =   1515
   End
   Begin VB.CommandButton cmdClone 
      Caption         =   "Clone"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9900
      TabIndex        =   10
      Top             =   6240
      Width           =   1150
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8580
      TabIndex        =   9
      Top             =   6240
      Width           =   1150
   End
   Begin VB.Frame Frame1 
      Caption         =   "Go To"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   12
      Top             =   6000
      Width           =   2715
      Begin VB.CommandButton cmdAssemblyMaintenance 
         Caption         =   "Assembly Maintenance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   6
         Top             =   240
         Width           =   1155
      End
      Begin VB.CommandButton cmdUnitCost 
         Caption         =   "Unit Cost"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7260
      TabIndex        =   8
      Top             =   6240
      Width           =   1150
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5940
      TabIndex        =   7
      Top             =   6240
      Width           =   1150
   End
   Begin VB.CheckBox ckbRowWrap 
      Caption         =   "Row Wrap"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      TabIndex        =   4
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8100
      TabIndex        =   3
      Top             =   2160
      Width           =   1150
   End
   Begin VB.TextBox AssemblyID 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8100
      TabIndex        =   2
      Top             =   1680
      Width           =   1515
   End
   Begin ConstructionCostDatabase.DynaTree FormatTree 
      Height          =   2775
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   4895
      ShowMasterFormatRoot=   -1  'True
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid 
      Height          =   2715
      Left            =   60
      TabIndex        =   19
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
      Splits(0)._ColumnProps(5)=   "Column(0)._MinWidth=33"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(1)._MinWidth=149300324"
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
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "MasterFormat:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6780
      TabIndex        =   22
      Top             =   525
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8160
      TabIndex        =   18
      Top             =   880
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9720
      TabIndex        =   17
      Top             =   880
      Width           =   1575
   End
   Begin VB.Label lblUnitCostId 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Cost ID:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6780
      TabIndex        =   16
      Top             =   1200
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
      Height          =   375
      Left            =   6780
      TabIndex        =   15
      Top             =   60
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Assembly ID:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6780
      TabIndex        =   14
      Top             =   1740
      Width           =   1215
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
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   13
      Top             =   2880
      Width           =   3255
   End
End
Attribute VB_Name = "frmUCostUsageGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_objGridMap As New CUCostUsageMap ' Class to handle grid
Dim m_blnFirstSearch As Boolean ' Is this the first search we have made on this screen
Dim m_blnJumpIn As Boolean ' Are we jumping here from another screen
Dim m_rec As New ADODB.RecordSet ' Recordset to hold query results
Dim m_blnDoubleClick As Boolean ' Did a double click just occur
Dim m_blnWereErrors As Boolean ' True if the Update had errors, used in QueryUnload
Dim m_strCurrentFormControl As String
Dim m_intMasterFormat As Long   ' Stores MasterFormat version to use by Search et al
Dim m_blnMasterFormatNotSpecified As Boolean    ' True if MF was never explicitly set

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

Private Function CheckEntryErrors() As Boolean
Dim i As Integer
Dim strItem As String
Dim varBookmarks() As Variant
Dim strSaveBookMark As Variant
On Error GoTo Error_Processing
ReDim varBookmarks(0 To m_rec.RecordCount)
strSaveBookMark = TDBGrid.Bookmark

    'Validate a unique unit_cost_id/sort_order
    If m_rec.RecordCount > 0 Then
        lstValidate.Clear
        TDBGrid.MoveFirst
        i = 0
        Do Until TDBGrid.EOF
            strItem = ""
            If Not IsNull(TDBGrid.Columns("Assembly ID")) Then
                strItem = TDBGrid.Columns("Assembly ID")
            End If
            If Not IsNull(TDBGrid.Columns("Unit Cost ID")) Then
                strItem = strItem + TDBGrid.Columns("Unit Cost ID")
            End If
            If Not IsNull(TDBGrid.Columns("Sort")) Then
                strItem = strItem + Trim(TDBGrid.Columns("Sort"))
            End If
            
            lstValidate.AddItem (strItem)
            lstValidate.ItemData(lstValidate.NewIndex) = i
            varBookmarks(i) = TDBGrid.Bookmark
            i = i + 1
            TDBGrid.MoveNext
        Loop
        'Start at the second item, compare each to the prior (list is sorted)
        For i = 1 To lstValidate.listcount - 1
            If lstValidate.List(i) = lstValidate.List(i - 1) Then
                CheckEntryErrors = True
                MsgBox "The Unit Cost ID/Sort ID must be unique."
                TDBGrid.Bookmark = varBookmarks(lstValidate.ItemData(i))
                m_objGridMap.SetError TDBGrid.Bookmark, "The Unit Cost ID/Sort ID must be unique."
                TDBGrid.RefetchRow (varBookmarks(i))
                Exit For
            End If
        Next i
    End If
    TDBGrid.Bookmark = strSaveBookMark

Exit_Function:
Exit Function
Error_Processing:
'MsgBox Error$
Resume Exit_Function
End Function

Public Sub SelectAllRows()
    ' Pass recordset to handler class
    m_objGridMap.RecordSet = m_rec
    
    If m_rec.RecordCount > 0 Then
        m_objGridMap.SelectAllRows
    End If
End Sub

Private Sub AssemblyID_LostFocus()
AssemblyID.Text = Trim(AssemblyID.Text)
End Sub

Private Sub cboMasterFormat_Click()
    MasterFormatChanged
End Sub

' Handles Row Wrap feature
Private Sub ckbRowWrap_Click()
    m_objGridMap.RowWrap (ckbRowWrap)
End Sub

Private Sub cmdAssemblyMaintenance_Click()
    If IsNumeric(TDBGrid.Bookmark) = True Then
        ' Navigate to grid view
        Dim frm As frmAssemblyGrid
        Set frm = New frmAssemblyGrid
        frm.JumpIn TDBGrid.Columns("Assembly ID").CellText(TDBGrid.Bookmark)
    Else
        MsgBox "You must select a row."
    End If

End Sub
Private Sub cmdClone_Click()
    Dim rec As ADODB.RecordSet
    If IsNull(TDBGrid.Bookmark) Then
        MsgBox "Please select a row to clone."
    ElseIf ValidGridRow() = True Then
            Set rec = m_objGridMap.CloneRow
    End If
End Sub

Private Sub cmdDelete_Click()
    Dim varButton
    varButton = MsgBox("Are you sure you want to delete?", vbYesNo + vbCritical)
    If varButton = vbYes Then
        TDBGrid.Delete
    End If
End Sub

Private Sub cmdUnitCost_Click()
    ' Navigate to grid view
    If IsNull(TDBGrid.Bookmark) Then
        MsgBox "Please select a row."
    Else
        Dim frm As frmUnitCostGrid
        Set frm = New frmUnitCostGrid
        frm.MasterFormat = MasterFormat
        frm.JumpIn Compress_String(TDBGrid.Columns("Unit Cost ID").CellText(TDBGrid.Bookmark)) + "*"
    End If
End Sub

Private Sub cmdNew_Click()
    Dim blnUnitCost As Boolean
    Dim blnAssembly As Boolean
    'New
    ' Open empty single record view
    '    Dim frm As frmMatPrice
    '    Set frm = New frmMatPrice
    '    frm.Show
    Dim bln_Continue As Boolean
    Dim varCurrentM_recBookmark As Variant
    
    If IsNull(TDBGrid.Bookmark) Then
        bln_Continue = True
    Else
        If ValidGridRow() = True Then
            bln_Continue = True
        End If
    End If

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
    If (Len(StartUnitCostID.Text) = 0 And Len(AssemblyID) = 10) And Right(AssemblyID, 1) <> "*" Then
        m_rec.Fields("parent_id").value = AssemblyID.Text
        m_rec.Fields("parent_skey").value = GetAssemblySkey(AssemblyID.Text)
            'If the start/end match then one uc ID was selected; use as default for new row.
    ElseIf (Len(AssemblyID) = 0 _
                And Len(StartUnitCostID.Text) = 12) _
                And Right(StartUnitCostID.Text, 1) <> "*" _
                And StartUnitCostID.Text = EndUnitCostID.Text Then
        m_rec.Fields("unit_cost_id").value = StartUnitCostID.Text
        m_rec.Fields("unit_cost_skey").value = GetUCSkey(StartUnitCostID.Text, MasterFormat)
    Else
        m_rec.Fields("unit_cost_skey").value = 0
        m_rec.Fields("parent_skey").value = 0
    End If
    m_rec.Fields("adj_factor") = 1
    m_rec.Fields("usage_unit_qty") = 1
    m_rec.Fields("last_update_id") = 0
    m_rec.Fields("last_update_person") = strUserName
    m_rec.Fields("metric_adj_factor") = 1
    m_rec.Fields("usage_metric_unit_qty") = 1
    m_objGridMap.SetRowState m_rec.Bookmark, STATE_NEW
        
    TDBGrid.SetFocus
    TDBGrid.AllowAddNew = False
    TDBGrid.ReOpen m_rec.Bookmark
    DoEvents
    m_rec.MoveLast
        
    If (Len(StartUnitCostID.Text) = 0 And Len(AssemblyID) = 10) And Right(AssemblyID, 1) <> "*" Then
        TDBGrid.Columns("Assembly ID").value = AssemblyID.Text
        TDBGrid.Columns("Parent Skey").value = GetAssemblySkey(AssemblyID.Text)
        blnAssembly = True
        m_rec.MoveLast
        m_objGridMap.FillAssembly AssemblyID
            'If the start/end match then one uc ID was selected; use as default for new row.
    ElseIf (Len(AssemblyID) = 0 _
                And Len(StartUnitCostID.Text) = 12) _
                And Right(StartUnitCostID.Text, 1) <> "*" _
                And StartUnitCostID.Text = EndUnitCostID.Text Then
        TDBGrid.Columns("Unit Cost ID").value = StartUnitCostID.Text
        TDBGrid.Columns("Unit Cost Skey").value = GetUCSkey(StartUnitCostID.Text, MasterFormat)
        blnUnitCost = True
        m_rec.MoveLast
        m_objGridMap.FillUnitCost StartUnitCostID.Text
    Else
        TDBGrid.Columns("Unit Cost Skey").value = 0
        TDBGrid.Columns("Parent Skey").value = 0
        
        'rlh 02/18/2009  John Chiang/Kathy R. issue
        '
        'Going into the the Assembly Maintenance throught the Unit Cost Usage Form
        'and then trying to add a NEW line using the NEW button was failing?
        'However, going into the Assembly Maintenance form, directly, and then adding lines
        'on the interior grid works...
        
        'FIX: John Chiang and Kathy R.
        m_objGridMap.FillAssembly AssemblyID   'rlh 02/18/2009  Wasn't being filled in?
    End If
    TDBGrid.Columns("Adj Factor") = 1
    TDBGrid.Columns("Qty") = 1
    TDBGrid.Columns("last_update_id") = 0
    'TDBGrid.Columns("Skey Type").Value = "A"
    TDBGrid.Columns("Met Qty") = 1
    TDBGrid.Columns("Met Adj") = 1

    TDBGrid.AllowAddNew = False
    
    If blnUnitCost = True Then
        TDBGrid.Col = TDBGrid.Columns("Assembly ID").ColIndex
    Else
        If blnAssembly = True Then
            TDBGrid.Col = TDBGrid.Columns("Unit Cost ID").ColIndex
        End If
    End If
    m_objGridMap.SetRowState TDBGrid.Bookmark, STATE_NEW
End If
TDBGrid.SetFocus

Exit_Sub:

End Sub
Private Function ValidGridRow() As Boolean
    If MASTER_FORMAT_ASSEMBLIES = 1995 Then
        If Len(Trim(TDBGrid.Columns("unit cost id"))) = 0 Or Len(Trim(TDBGrid.Columns("assembly id"))) = 0 Then
            MsgBox "Both the Assembly and Unit Cost IDs must be entered."
            TDBGrid.SetFocus
            ValidGridRow = False
        Else
            ValidGridRow = True
        End If
    End If
    If MASTER_FORMAT_ASSEMBLIES = 2004 Then
        If Len(Trim(TDBGrid.Columns("unit cost id"))) = 0 Or Len(Trim(TDBGrid.Columns("assembly id"))) = 0 Then
            MsgBox "Both the Assembly and Unit Cost IDs must be entered."
            TDBGrid.SetFocus
            ValidGridRow = False
        Else
            ValidGridRow = True
        End If
    End If
    
    If AssemblyUCSortRequired(TDBGrid.Columns("parent_skey")) = True And Len(Trim(TDBGrid.Columns("Sort"))) = 0 Then
        MsgBox "This assembly has an assembly book system line:  The sort order is required."
        TDBGrid.SetFocus
        ValidGridRow = False
    Else
        ValidGridRow = True
    End If

End Function

Private Sub cmdUpdate_Click()
    Dim blnRet As Boolean
    Dim vntBookmark As Variant
    On Error GoTo Error_Processing

    Screen.MousePointer = vbHourglass
    m_blnWereErrors = False
    
    vntBookmark = TDBGrid.Bookmark

If ValidGridRow() = True Then
    TDBGrid.Update
    If CheckEntryErrors() = False Then
        blnRet = m_objGridMap.Update
        If blnRet = False Then
            m_blnWereErrors = True
        End If
        TDBGrid.Bookmark = vntBookmark
 Else
    m_blnWereErrors = True
   End If
Else
    m_blnWereErrors = True
End If
Exit_Sub:
Screen.MousePointer = vbNormal
Exit Sub

Error_Processing:
'MsgBox Error$
Resume Exit_Sub
Resume 0
End Sub

Private Sub EndUnitCostID_LostFocus()
    EndUnitCostID.Text = Trim(EndUnitCostID.Text)
End Sub

Private Sub Form_Deactivate()
    m_strCurrentFormControl = Me.ActiveControl.Name
End Sub

Private Sub Form_Initialize()
    
    ' Fill the MasterFormat tree
    m_intMasterFormat = g_intMasterFormat
    FormatTree.ClearTree
    If m_intMasterFormat = EXT_MASTERFORMAT_VERSION Then
        FormatTree.InitData g_cnShared, "UNITCOST04"
    Else
        FormatTree.InitData g_cnShared, "UNITCOST"
    End If
    m_blnMasterFormatNotSpecified = True
    
    ' Initialize grid
    m_objGridMap.SetGrid TDBGrid
    m_objGridMap.InitGrid
    m_blnFirstSearch = True
    m_blnJumpIn = False

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
    
    Move START_LEFT, START_TOP, START_WIDTH, START_HEIGHT
    
    LoadMasterFormatCombo Me.cboMasterFormat, True
    
    'rlh 05/15/2008
    Select Case MASTER_FORMAT_ASSEMBLIES        'gets set on frmAssemblyGrid/cboMasterFormat
    Case 2004
        Me.cboMasterFormat.ListIndex = 0        'rlh 05/15/2008
    Case 1995
        Me.cboMasterFormat.ListIndex = 1        'rlh 05/15/2008
    Case Else
        If MF95_ENABLED Then                    'rlh CCD 8.4  04/7/2009
            MsgBox ("MASTER FORMAT NOT YET SELECTED ON ASSEMBLIES GRID. WE'LL DEFAULT TO MF-2004.  PROCESSING WILL CONTINUE...")
        End If
        
        Me.cboMasterFormat.ListIndex = 0        'rlh 05/15/2008
    End Select
    
    ' This will never return any rows, just used to create recordset
    StartUnitCostID.Text = "~"
    cmdSearch_Click
    StartUnitCostID.Text = ""
    
End Sub

' Called when coming here from another screen
Public Sub JumpIn(strUnitCostId As String)
    StartUnitCostID.Text = strUnitCostId
    EndUnitCostID.Text = strUnitCostId
    If m_blnMasterFormatNotSpecified Then
        ' MF was never explicitly set, so default to 1995 for compatibility purposes
        MasterFormat = 1995
    End If
    cmdSearch_Click
End Sub

' Called when coming here from another screen
Public Sub JumpIn2(strAssemblyId As String)
    AssemblyID.Text = strAssemblyId
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
Dim rs As New ADODB.RecordSet
Dim strSelect As String
Dim blnReturn As Boolean
On Error Resume Next
    ' Synch text box with tree
    If Len(strID) = 12 Then
        StartUnitCostID.Text = strID + "*"
        EndUnitCostID.Text = ""
        AssemblyID.Text = ""
    Else
        rs.Close ' Make sure it is closed
        
        'Comment added by Mohan: Jan 05, 2012, Master Format is always going to be using the 2004 MASTERFORMAT, leaving the If condition just in case something comes up in the future
        If MasterFormat = EXT_MASTERFORMAT_VERSION Then
            strSelect = "select unit_cost_id_start, unit_cost_id_end from MASTERFORMAT04_ID_HIERARCHY where hier_id='" + strID + "'"
        Else
            'Line of code was changed by Mohan on Jan 05,2012, MASTERFORMAT95_ID_HIERARCHY was changed to MASTERFORMAT04_ID_HIERARCHY
            strSelect = "select unit_cost_id_start, unit_cost_id_end from MASTERFORMAT04_ID_HIERARCHY where hier_id='" + strID + "'"
        End If
        blnReturn = g_objDAL.GetRecordset(vbNullString, strSelect, rs)
        StartUnitCostID.Text = rs.Fields("unit_cost_id_start")
        EndUnitCostID.Text = rs.Fields("unit_cost_id_end")
        ' Clear other boxes
        AssemblyID.Text = ""
        rs.Close
    End If
    ' Kick-off search
    cmdSearch_Click

End Sub

Private Sub cmdSearch_Click()
    On Error Resume Next
    Dim blnReturn As Boolean
    Const ASSEMBLY = 1
    Const BOTH = 2
    Const UNITCOST = 3
    Dim iSelMode As Integer
    Dim strSelect As String
    Dim dtmToday As Date
    Dim dtmStart As Date
    Dim strError As String
    Dim strStartUnitCostSrch As String
    Dim iMasterFormat As Long
    
    dtmToday = Date
    
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
    lblRowCount.Caption = "Working..."
    lblRowCount.Refresh
    
    iMasterFormat = cboMasterFormat.ItemData(cboMasterFormat.ListIndex)
    
    ' Synch tree with text box
'    If Not StartUnitCostID.Text = "" Then
'        FormatTree.FocusItem (StartUnitCostID.Text)
'    End If
    
    m_rec.Close ' Make sure it is closed
    m_rec.MaxRecords = MAX_RECORDS ' Set the maximum number to bring back
    dtmStart = Now

'    If Len(AssemblyID.Text) = 0 And Len(UnitCostID.Text) = 0 Then
'        Screen.MousePointer = vbNormal
'        MsgBox "You must enter either Unit Cost ID or Assembly ID"
'        Exit Sub
'    End If
    If AssemblyID.Text = "" Then
        iSelMode = UNITCOST
        AssemblyID.Text = "*"
    Else
        iSelMode = ASSEMBLY
    End If
    If Len(StartUnitCostID) = 12 And InStr(1, StartUnitCostID, "*") = 0 And Len(EndUnitCostID) = 0 Then
        If iSelMode = ASSEMBLY Then
            iSelMode = BOTH
        Else
            iSelMode = UNITCOST
        End If
        strStartUnitCostSrch = StartUnitCostID + "*"
    Else
        If StartUnitCostID = "" Then
            strStartUnitCostSrch = "*"
        Else
            If iSelMode = ASSEMBLY Then
                iSelMode = BOTH
            Else
                iSelMode = UNITCOST
            End If
            strStartUnitCostSrch = StartUnitCostID
        End If
    End If

    

        strSelect = "exec usp_select_unit_cost_usage_ext @parent_id='" + _
        SQLChangeWildcard(AssemblyID.Text) + "'" + _
        ", @skey_type = 'A', @start_unit_cost_id = '" + SQLChangeWildcard(strStartUnitCostSrch) + "'" + _
        ", @end_unit_cost_id = '" + EndUnitCostID.Text + "'" + _
        ", @master_format = '" & CStr(iMasterFormat) & "'" & _
        ", @selmode = " & iSelMode

    ' Use DAL to perform select
    blnReturn = g_objDAL.GetRecordset(vbNullString, strSelect, m_rec)
    If blnReturn = False Then
        lblRowCount.Caption = "0 rows returned."
        Screen.MousePointer = vbNormal
        MsgBox "An error occurred while searching. Error:" + Error$, vbInformation
        Exit Sub
    End If
    
    ' Set MasterFormat to match new results
    MasterFormat = iMasterFormat
    m_objGridMap.MasterFormat = iMasterFormat

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
    If StartUnitCostID.Text = "*" Then
        StartUnitCostID.Text = ""
    End If
    If AssemblyID.Text = "*" Then
        AssemblyID.Text = ""
    End If

    ' Reset the grid contents
    TDBGrid.ReBind
    TDBGrid.ApproxCount = m_rec.RecordCount
    m_objGridMap.SetRowStateNone
    Screen.MousePointer = vbNormal
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    'Bypass confirm update until update is enabled
    ' Check if there are pending changes
    If m_objGridMap.IsPendingChange = True And cmdUpdate.Enabled = True Then
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
        blnReturn = UnLockField(Me, "EndUnitCostID")
    End If
End Sub

Private Sub StartUnitCostID_LostFocus()
    StartUnitCostID.Text = Trim(StartUnitCostID.Text)
End Sub

Private Sub TDBGrid_Click()
'Debug.Print TDBGrid.Columns(TDBGrid.Col).ValueItems(2)
'Debug.Print TDBGrid.Columns(20).CellText(1)
End Sub

Private Sub TDBGrid_ComboSelect(ByVal ColIndex As Integer)
    With TDBGrid
    If UCase(.Columns(ColIndex).Caption) = "UNIT" And Trim(.Columns(ColIndex).value) <> "" Then
        Dim i As Integer
        Dim blnReturn As Boolean
        Dim rstMetUnit As ADODB.RecordSet
        Dim strSQL As String
'*** APEX Migration Utility Code Change ***
'        Dim objItem As New TrueOleDBGrid70.ValueItem
        Dim objItem As New TrueOleDBGrid80.ValueItem
        
        strSQL = "SELECT DISTINCT metric_unit " & _
            "FROM metric_conversion where UPPER(RTRIM(LTRIM(unit))) = '" & UCase(Trim(.Columns(ColIndex).value)) & "'"
        blnReturn = g_objDAL.GetRecordset(vbNullString, strSQL, rstMetUnit)
        If blnReturn And rstMetUnit.RecordCount = 1 Then
            rstMetUnit.MoveFirst
            If Not IsNull(rstMetUnit!metric_unit) Then
                .Columns("Met Unit").value = CStr(rstMetUnit!metric_unit)
            End If
        End If
    Else
        .Columns("Met Unit").value = ""
    End If
    End With
End Sub

Private Sub TDBGrid_DblClick()
    ' Signal that double-click has occurred
    m_blnDoubleClick = True
End Sub

Private Sub TDBGrid_GotFocus()
    TDBGrid.TabStop = True
End Sub

Private Sub TDBGrid_LostFocus()
    TDBGrid.TabStop = False
End Sub

Private Sub TDBGrid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If this is the mouse-up form a double click
'    If m_blnDoubleClick Then
        ' Make sure it is the left button
'        If Button = vbLeftButton Then
'            m_blnDoubleClick = False
            ' Same function as clicking Material Price button, open single record view
'            cmdMaterialPrice_Click
'        End If
'    Else
        If Button = vbRightButton And Not IsNull(TDBGrid.Bookmark) Then
            Dim strErrorMsg As String
            strErrorMsg = m_objGridMap.GetError(TDBGrid.Bookmark)
            If Len(strErrorMsg) > 0 Then
                MsgBox strErrorMsg
            End If
        End If
'    End If
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

Private Sub MasterFormatChanged()
'A NEW MASTERFORMAT WAS SELECTED FROM THE DROP-DOWN BOX
'ADDED 6/20/2005 RTD FOR VERSION 7.4.0+
    Dim sTreeType As String
    
    
    Exit Sub
    Select Case cboMasterFormat.ItemData(cboMasterFormat.ListIndex)
    Case 2004
        UnLockField Me, "EndUnitCostID"
        lblUnitCostId.Caption = "Unit Cost ID 04:"
        sTreeType = "UNITCOST04"
    Case 1995
        UnLockField Me, "EndUnitCostID"
        lblUnitCostId.Caption = "Unit Cost ID 95:"
        sTreeType = "UNITCOST"
    Case 1988
        LockField Me, "EndUnitCostID"
        'EndUnitCostID.Text = ""
        lblUnitCostId.Caption = "Alt Unit Cost ID:"
        sTreeType = "UNITCOST"
    Case Else
        UnLockField Me, "EndUnitCostID"
        lblUnitCostId.Caption = "Unit Cost ID 95:"
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

