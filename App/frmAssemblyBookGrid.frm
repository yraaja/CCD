VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{5936A75C-3F42-11D6-AF6B-AA0004005F12}#1.3#0"; "MeansCtrl.ocx"
Begin VB.Form frmAssemblyBookGrid 
   Caption         =   "Assembly Book Detail Grid"
   ClientHeight    =   6900
   ClientLeft      =   1500
   ClientTop       =   2790
   ClientWidth     =   11310
   Icon            =   "frmAssemblyBookGrid.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6900
   ScaleWidth      =   11310
   Begin VB.Frame fraAssemblyType 
      Caption         =   "Assembly Type"
      Height          =   615
      Left            =   8280
      TabIndex        =   19
      Top             =   120
      Width           =   2895
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   2535
         TabIndex        =   26
         Top             =   240
         Width           =   2535
         Begin VB.OptionButton optAssemblyType 
            Caption         =   "R&esidential"
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   28
            Top             =   0
            Width           =   1215
         End
         Begin VB.OptionButton optAssemblyType 
            Caption         =   "Co&mmercial"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   27
            Top             =   0
            Value           =   -1  'True
            Width           =   1215
         End
      End
   End
   Begin VB.TextBox StartAssemblyID 
      Height          =   315
      Left            =   7800
      TabIndex        =   23
      Top             =   2160
      Width           =   1515
   End
   Begin VB.TextBox EndAssemblyID 
      Height          =   315
      Left            =   9540
      TabIndex        =   22
      Top             =   2160
      Width           =   1515
   End
   Begin VB.TextBox AltBookId 
      Height          =   315
      Left            =   8340
      TabIndex        =   20
      Top             =   1155
      Width           =   1515
   End
   Begin VB.TextBox bookdesc 
      Height          =   345
      Left            =   8340
      TabIndex        =   1
      Top             =   1530
      Width           =   2355
   End
   Begin VB.CommandButton cmdClone 
      Caption         =   "&Clone"
      Height          =   495
      Left            =   9540
      TabIndex        =   10
      Top             =   6240
      Width           =   915
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   495
      Left            =   8355
      TabIndex        =   9
      Top             =   6240
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Caption         =   "Go To"
      Height          =   855
      Left            =   120
      TabIndex        =   13
      Top             =   6000
      Width           =   3735
      Begin VB.CommandButton cmdBookDetail 
         Caption         =   "&Book Detail"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   915
      End
      Begin VB.CommandButton cmdAssembly 
         Caption         =   "&Assembly"
         Height          =   495
         Left            =   1320
         TabIndex        =   5
         Top             =   240
         Width           =   915
      End
      Begin VB.CommandButton cmdOutput 
         Caption         =   "&Output"
         Height          =   495
         Left            =   2400
         TabIndex        =   6
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   495
      Left            =   7200
      TabIndex        =   8
      Top             =   6240
      Width           =   915
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   495
      Left            =   6000
      TabIndex        =   7
      Top             =   6240
      Width           =   915
   End
   Begin VB.CheckBox ckbRowWrap 
      Caption         =   "Row Wrap"
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   255
      Left            =   8760
      TabIndex        =   2
      Top             =   2520
      Width           =   1150
   End
   Begin VB.TextBox bookid 
      Height          =   315
      Left            =   8340
      TabIndex        =   0
      Top             =   780
      Width           =   1515
   End
   Begin ConstructionCostDatabase.DynaTree FormatTree 
      Height          =   2775
      Left            =   0
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   4895
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid 
      Height          =   2715
      Left            =   60
      TabIndex        =   12
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
   Begin VB.Label Label7 
      Caption         =   "From"
      Height          =   255
      Left            =   8340
      TabIndex        =   25
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "To"
      Height          =   255
      Left            =   10080
      TabIndex        =   24
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Alt Book ID:"
      Height          =   255
      Left            =   6840
      TabIndex        =   21
      Top             =   1215
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Assembly ID:"
      Height          =   255
      Left            =   6720
      TabIndex        =   18
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Book ID:"
      Height          =   255
      Left            =   6840
      TabIndex        =   17
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Book Description:"
      Height          =   255
      Left            =   6720
      TabIndex        =   16
      Top             =   1590
      Width           =   1455
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
      Left            =   6780
      TabIndex        =   14
      Top             =   60
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
Attribute VB_Name = "frmAssemblyBookGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_objGridMap As New CAssemblyBkMap ' Class to handle grid
Dim m_blnFirstSearch As Boolean ' Is this the first search we have made on this screen
Dim m_blnJumpIn As Boolean ' Are we jumping here from another screen
Dim m_rec As New ADODB.RecordSet ' Recordset to hold query results
Dim m_blnDoubleClick As Boolean ' Did a double click just occurr
Dim m_blnWereErrors As Boolean ' True if the Update had errors, used in QueryUnload
Const USEBOOKMARK = 1
Const USECOORD = 0
Dim rsAssemblyBookClone As RecordSet
Dim m_iBookType As Integer

Const COMMERCIAL_ASSEMBLIES = 0
Const RESIDENTIAL_ASSEMBLIES = 1
Dim m_sngYCoord As Single
Dim m_strCurrentFormControl As String

Public Sub SetMenuBar()
    m_objGridMap.SetMenuBar
End Sub
Public Sub Sort(intDir As Integer)
    m_objGridMap.Sort intDir
End Sub
Public Sub PrintReport()

End Sub

Public Sub PreviewReport()

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
Dim bln_RecordsFound As Boolean
On Error GoTo Exit_Sub
If rsAssemblyBookClone.RecordCount = 0 Then
    bln_RecordsFound = False
Else
    bln_RecordsFound = True
End If
If bln_RecordsFound Then
    Select Case Mode
        Case USEBOOKMARK
            rsAssemblyBookClone.Bookmark = TDBGrid.Bookmark
        Case USECOORD
            rsAssemblyBookClone.Bookmark = TDBGrid.RowBookmark(TDBGrid.RowContaining(Coord))
    End Select
    
    If IsNumeric(rsAssemblyBookClone.Bookmark) Then
        cmdClone.Enabled = True
        cmdDelete.Enabled = True
        cmdUpdate.Enabled = True
        cmdBookDetail.Enabled = True
        cmdOutput.Enabled = True
        If IsNull(rsAssemblyBookClone.Fields("assembly_id")) Then
            cmdAssembly.Enabled = False
        Else
            cmdAssembly.Enabled = True
        End If
    Else
        cmdClone.Enabled = False
        cmdDelete.Enabled = False
        cmdUpdate.Enabled = False
        cmdAssembly.Enabled = False
        cmdBookDetail.Enabled = False
        cmdOutput.Enabled = False
    End If
Else
    cmdClone.Enabled = False
    cmdDelete.Enabled = False
    cmdUpdate.Enabled = False
    cmdAssembly.Enabled = False
    cmdBookDetail.Enabled = False
    cmdOutput.Enabled = False
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

Private Sub startassemblyid_LostFocus()
StartAssemblyID.Text = Trim(StartAssemblyID.Text)
End Sub


Private Sub bookdesc_LostFocus()
bookdesc.Text = Trim(bookdesc.Text)
End Sub


Private Sub bookid_Change()
bookid.Text = Trim(bookid.Text)
End Sub


' Handles Row Wrap feature
Private Sub ckbRowWrap_Click()
    m_objGridMap.RowWrap (ckbRowWrap)
End Sub


Private Sub cmdAssembly_Click()
If IsNumeric(TDBGrid.Bookmark) Then
   ' Navigate to grid view
   Dim frm As frmAssemblyGrid
   Set frm = New frmAssemblyGrid
   frm.JumpIn TDBGrid.Columns("Assembly ID").CellText(TDBGrid.Bookmark)
Else
   MsgBox "Please select a row first."
End If
End Sub

Private Sub cmdBookDetail_Click()
    If IsNumeric(TDBGrid.Bookmark) = False Then
        MsgBox "You must select a row."
        Exit Sub
    End If
'    ' Navigate to single-record view
    Dim frm As frmAssemblyBookDetail
    Dim rec As ADODB.RecordSet
    Set frm = New frmAssemblyBookDetail
    ' Make copy of recordset
    Set rec = m_rec.Clone
    ' Get the selected row from grid
    rec.Bookmark = TDBGrid.Bookmark
    frm.SetRow rec ' Pass the current record into the form
'    Set frm.frmCallingForm = Me
    Set frm.tdbCols = Me.TDBGrid.Columns
    Set frm.myTDBGrid = Me.TDBGrid
    frm.Show
End Sub

Private Sub cmdClone_Click()
    On Error GoTo Out
    If IsNumeric(TDBGrid.Bookmark) = False Then
        MsgBox "You must select a row."
        Exit Sub
    End If
    Dim rec As ADODB.RecordSet

    Set rec = m_objGridMap.CloneRowRecordset
    ' Navigate to single-record view
    Dim frm As frmAssemblyBookDetail
    Set frm = New frmAssemblyBookDetail
    frm.SetRow rec, True  ' Pass the current record into the form
    frm.Show
    frm.metric_calculation_factor = TDBGrid.Columns("metric_calculation_factor").Value
Out:
End Sub

Private Sub cmdDelete_Click()
    m_objGridMap.Delete
End Sub

Private Sub cmdOutput_Click()
Dim frm As Form
Dim blnVisible As Boolean

m_strKeyType2 = "A"      'flag as unit cost line processing  (rlh) 07/14/2008

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

Private Sub cmdNew_Click()
    On Error GoTo Out
    Dim rec As New ADODB.RecordSet
    
    CopyRSFields rec, m_rec
    ' Open empty single record view
    Dim frm As frmAssemblyBookDetail
    Set frm = New frmAssemblyBookDetail
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
    Dim bln_Continue As Boolean
    
    m_blnWereErrors = False
    
    If TDBGrid.Columns("metric_calculation_factor").Value = 0 Then
        MsgBox "Metric Calculation Factor cannot be 0!"

        TDBGrid.ReBind

    End If
    
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
Out:
End Sub

Private Sub Form_Deactivate()
m_strCurrentFormControl = Me.ActiveControl.Name
End Sub

Private Sub Form_Initialize()
    ' Fill the MasterFormat tree
    Screen.MousePointer = vbHourglass
    m_blnFirstSearch = True
    FormatTree.InitData g_cnShared, "ASBLY_BK_DTL_COMMERCIAL"
    ' Initialize grid
    m_objGridMap.SetGrid TDBGrid
    m_objGridMap.InitGrid
    m_blnFirstSearch = False
    Screen.MousePointer = vbNormal
    m_blnJumpIn = False
End Sub

Private Sub Form_Load()
    Dim blnReturn As Boolean
    Dim strSELECT As String
    Dim rec As ADODB.RecordSet
    Move START_LEFT, START_TOP, START_WIDTH, START_HEIGHT

'    ' This will never return any rows, just used to create recordset
    bookid.Text = "~"
    cmdSearch_Click
    bookid.Text = ""
End Sub
' Called when coming here from another screen
Public Sub JumpIn(strAssemblyId As String)
    StartAssemblyID.Text = strAssemblyId & "*"
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
Dim Rs As ADODB.RecordSet
Dim strSELECT As String
Dim blnReturn As Boolean
On Error Resume Next
    ' Synch text box with tree
    Rs.Close ' Make sure it is closed
    strSELECT = "select assembly_id_start, assembly_id_end from UNIFORMAT2_ID_HIERARCHY where uni2_category_id='" + strID + "'"
    blnReturn = g_objDAL.GetRecordset(vbNullString, strSELECT, Rs)
    StartAssemblyID.Text = Rs.Fields("assembly_id_start")
    EndAssemblyID.Text = Rs.Fields("assembly_id_end")
    
    bookid.Text = ""
    bookdesc.Text = ""
    ' Kick-off search
    cmdSearch_Click
End Sub
Private Sub cmdSearch_Click()
    On Error Resume Next
    Dim blnReturn As Boolean
    Dim strSELECT As String
    Dim dtmToday As Date
    Dim dtmStart As Date
    Dim strSrchStartAssemblyID As String
    Dim strSrchEndAssemblyID As String
    Dim strAltBookIDSrch As String
    Dim strBookIDSrch As String
    Dim strBookDescSrch As String
    Dim frm As Form
    Dim blnVisible As Boolean
    
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
    'If Not assemblyid.Text = "" Then
    '    FormatTree.FocusItem (UnitCostID.Text)
    'End If
    If Len(StartAssemblyID.Text) = 0 And Len(AltBookId.Text) = 0 And Len(bookid.Text) = 0 And Len(bookdesc.Text) = 0 Then
        Screen.MousePointer = vbNormal
        MsgBox "You must enter Search Criteria."
        Exit Sub
    End If
    If Len(StartAssemblyID) = 12 And InStr(1, StartAssemblyID, "*") = 0 Then
        strSrchStartAssemblyID = StartAssemblyID + "*"
    End If
    If Len(AltBookId) = 12 And InStr(1, AltBookId, "*") = 0 Then
        strAltBookIDSrch = AltBookId + "*"
    End If
    
    If StartAssemblyID = "" Then
        strSrchStartAssemblyID = "*"
    Else
        strSrchStartAssemblyID = StartAssemblyID
    End If
    strSrchEndAssemblyID = EndAssemblyID
    
    If AltBookId = "" Then
        strAltBookIDSrch = "*"
    Else
        strAltBookIDSrch = AltBookId
    End If
    If bookid.Text = "" Then
        strBookIDSrch = "*"
    Else
        strBookIDSrch = bookid.Text
    End If
    If bookdesc.Text = "" Then
        strBookDescSrch = "*"
    Else
        strBookDescSrch = bookdesc.Text
    End If
    
    lblRowCount.Caption = "Working..."
    lblRowCount.Refresh

    m_rec.Close ' Make sure it is closed
    m_rec.MaxRecords = MAX_RECORDS ' Set the maximum number to bring back
    dtmStart = Now

    Dim strError As String

    strSELECT = "exec sp_select_book_detail @start_assembly_id = '" + SQLChangeWildcard(strSrchStartAssemblyID) + "', @end_assembly_id = '" + SQLChangeWildcard(strSrchEndAssemblyID) + "',@alt_assembly_book_id = '" + SQLChangeWildcard(strAltBookIDSrch) + "', @assembly_book_id = '" + SQLChangeWildcard(strBookIDSrch) + "', @book_desc = '" + SQLChangeWildcard(strBookDescSrch) + _
    "', @book_type = " + CStr(m_iBookType)
    ' Use DAL to perform select
    blnReturn = g_objDAL.GetRecordset(vbNullString, strSELECT, m_rec)
    If blnReturn = False Then
        MsgBox "An error occurred while searching."
        lblRowCount.Caption = "0 rows returned."
        Screen.MousePointer = vbNormal
        Exit Sub
    End If
    Set rsAssemblyBookClone = m_rec.Clone
    If FormOpen("dlgOutput", frm, blnVisible) = True Then
        If blnVisible = True Then       'Hide output before loading recordset
            DoOutput
        End If
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
    DoEvents
    ' Reset the grid contents
    TDBGrid.Bookmark = Null
    TDBGrid.ReBind
    TDBGrid.ApproxCount = m_rec.RecordCount
    SetButtons USEBOOKMARK
    Screen.MousePointer = vbNormal
End Sub
Private Function ValidGridRow() As Boolean

If TDBGrid.Columns("Comm'l Ind").Value = 0 And TDBGrid.Columns("Resi Ind").Value = 0 Then
    TDBGrid.SetFocus
    MsgBox "Commercial or Residential Use indicator must be selected."
    ValidGridRow = False
Else
    ValidGridRow = True
End If

End Function
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
    Dim strSELECT As String
    On Error GoTo Error_Processing
    
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
        strUpdate1 = "exec sp_temp_add_output_keys @skey_type = 'A', @skey = "
        If TDBGrid.SelBookmarks.Count = 0 Then  'No rows selected
            If Not IsNull(TDBGrid.Bookmark) Then    'Use current row
                m_rec.Bookmark = TDBGrid.Bookmark
                strUpdate = strUpdate1 + CStr(m_rec.Fields("assembly_book_skey"))
                blnReturn = frm.m_objOutput.ExecQuery(vbNullString, strUpdate, strError)
            End If
        Else
            For Each varBookmark In TDBGrid.SelBookmarks
                m_rec.Bookmark = varBookmark
                strUpdate = strUpdate1 + CStr(m_rec.Fields("assembly_book_skey"))
                blnReturn = frm.m_objOutput.ExecQuery(vbNullString, strUpdate, strError)
            Next varBookmark
        End If
        frm.m_strKeyType = "A"
        frm.FillData
        frm.Show vbModeless, fMainForm
        frm.Caption = "Output Usage"
    End If
Exit_Sub:
Exit Sub

Error_Processing:
MsgBox Error$
Resume Exit_Sub
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    ' Check if there are pending changes
    If m_objGridMap.IsPendingChange = True Then
        Dim Button
        Button = MsgBox("Do you want to save your changes?", vbYesNoCancel)
        If Button = vbYes Then
            cmdUpdate_Click
        End If
            ' If there were errors, cancel the close
        If Button = vbCancel Or m_blnWereErrors = True Then
            Cancel = True
            Exit Sub
        End If
    End If
End Sub

Private Sub optAssemblyType_Click(index As Integer)
FormatTree.ClearTree

Select Case index
    ' Fill the MasterFormat tree
    Case 0 'Commercial
        FormatTree.InitData g_cnShared, "ASBLY_BK_DTL_COMMERCIAL"
        m_iBookType = COMMERCIAL_ASSEMBLIES
    Case 1 'Residential
        FormatTree.InitData g_cnShared, "ASBLY_BK_DTL_RESI"
        m_iBookType = RESIDENTIAL_ASSEMBLIES
End Select
If StartAssemblyID.Text = "" And bookid.Text = "" And bookdesc.Text = "~" Then
    StartAssemblyID.Text = "~"
End If
cmdSearch_Click
If StartAssemblyID.Text = "" And bookid.Text = "" And bookdesc.Text = "~" Then
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
            ' Same function as clicking Assembly button, open single record view
            cmdBookDetail_Click
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
'    TDBGrid.ReBind
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
End Sub





Private Sub TDBGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim lngCurrentRow As Long
    If m_sngYCoord > 0 Then
        lngCurrentRow = TDBGrid.RowContaining(m_sngYCoord)
        If lngCurrentRow <> -1 Then
            If CLng(LastRow) - 1 <> lngCurrentRow Then
            TDBGrid.Row = lngCurrentRow
                m_rec.Bookmark = TDBGrid.Bookmark
                position_output
                SetButtons USECOORD, m_sngYCoord
                m_sngYCoord = 0
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
    
End Sub

Private Sub TDBGrid_SelChange(Cancel As Integer)
    Dim lngCurrentRow As Long
    If m_sngYCoord > 0 Then
        lngCurrentRow = TDBGrid.RowContaining(m_sngYCoord)
        If lngCurrentRow <> -1 Then
            TDBGrid.Row = lngCurrentRow
                m_rec.Bookmark = TDBGrid.Bookmark
                position_output
                SetButtons USECOORD, m_sngYCoord
                m_sngYCoord = 0
        End If
    Else
        position_output
        SetButtons USEBOOKMARK
    End If
End Sub


