VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmCrewGrid 
   Caption         =   "Crew Grid"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11130
   Icon            =   "frmCrewGrid.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6855
   ScaleWidth      =   11130
   Visible         =   0   'False
   Begin VB.ComboBox TracesCrewID 
      Height          =   315
      Left            =   9840
      TabIndex        =   31
      Top             =   1650
      Width           =   1155
   End
   Begin VB.CommandButton cmdCloneCrew 
      Caption         =   "&Clone Crew"
      Height          =   495
      Left            =   3600
      TabIndex        =   29
      Top             =   6240
      Width           =   1150
   End
   Begin VB.CommandButton cmdEditCrew 
      Caption         =   "&Edit Crew"
      Height          =   495
      Left            =   4860
      TabIndex        =   30
      Top             =   6255
      Width           =   1150
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   495
      Left            =   6120
      TabIndex        =   28
      Top             =   6240
      Width           =   1150
   End
   Begin VB.ComboBox CountryCode 
      Height          =   315
      ItemData        =   "frmCrewGrid.frx":0442
      Left            =   10200
      List            =   "frmCrewGrid.frx":044F
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox RegionCode 
      Height          =   315
      ItemData        =   "frmCrewGrid.frx":0461
      Left            =   10200
      List            =   "frmCrewGrid.frx":046B
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox CrewID 
      Height          =   315
      Left            =   7740
      TabIndex        =   2
      Top             =   1650
      Width           =   795
   End
   Begin VB.ComboBox EquipmentID 
      Height          =   315
      Left            =   7740
      TabIndex        =   4
      Top             =   2400
      Width           =   1515
   End
   Begin VB.TextBox StartUnitCostID 
      Height          =   315
      Left            =   7680
      TabIndex        =   0
      Top             =   660
      Width           =   1515
   End
   Begin VB.TextBox EndUnitCostID 
      Height          =   315
      Left            =   7680
      TabIndex        =   1
      Top             =   1080
      Width           =   1515
   End
   Begin VB.ComboBox TradeID 
      Height          =   315
      Left            =   7740
      TabIndex        =   3
      Top             =   2040
      Width           =   1515
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   375
      Left            =   9600
      TabIndex        =   7
      Top             =   2400
      Width           =   1395
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   495
      Left            =   7380
      TabIndex        =   11
      Top             =   6240
      Width           =   1150
   End
   Begin VB.Frame Frame1 
      Caption         =   "Go To"
      Height          =   855
      Left            =   60
      TabIndex        =   16
      Top             =   6000
      Width           =   2115
      Begin VB.CommandButton cmdEquipPrice 
         Caption         =   "&Equipment Price"
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   915
      End
      Begin VB.CommandButton cmdLaborRate 
         Caption         =   "Labor &Rate"
         Height          =   495
         Left            =   1080
         TabIndex        =   9
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   495
      Left            =   8640
      TabIndex        =   12
      Top             =   6240
      Width           =   1150
   End
   Begin VB.CommandButton cmdClone 
      Caption         =   "&Clone"
      Height          =   495
      Left            =   9900
      TabIndex        =   13
      Top             =   6240
      Width           =   1150
   End
   Begin ConstructionCostDatabase.DynaTree FormatTree 
      Height          =   2775
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   4895
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid 
      Height          =   2715
      Left            =   60
      TabIndex        =   10
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
      Splits(0)._ColumnProps(5)=   "Column(0)._MinWidth=29800"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(1)._MinWidth=149414304"
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
   Begin VB.CheckBox ckbRowWrap 
      Caption         =   "Row Wrap"
      Height          =   315
      Left            =   120
      TabIndex        =   14
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Traces Crew ID:"
      Height          =   255
      Left            =   8640
      TabIndex        =   32
      Top             =   1710
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Country:"
      Height          =   255
      Left            =   9600
      TabIndex        =   27
      Top             =   540
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "Region:"
      Height          =   255
      Left            =   9600
      TabIndex        =   26
      Top             =   180
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Line Line4 
      X1              =   7320
      X2              =   8040
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line3 
      X1              =   8520
      X2              =   9240
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label8 
      Caption         =   "OR"
      Height          =   255
      Left            =   8160
      TabIndex        =   25
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "From:"
      Height          =   255
      Left            =   6840
      TabIndex        =   24
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "To:"
      Height          =   255
      Left            =   6720
      TabIndex        =   23
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Unit Cost ID:"
      Height          =   255
      Left            =   8040
      TabIndex        =   22
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Equipment ID:"
      Height          =   255
      Left            =   6600
      TabIndex        =   21
      Top             =   2460
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Trade ID:"
      Height          =   255
      Left            =   6840
      TabIndex        =   20
      Top             =   2085
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   6660
      X2              =   6660
      Y1              =   2700
      Y2              =   0
   End
   Begin VB.Line Line2 
      X1              =   210
      X2              =   10770
      Y1              =   2820
      Y2              =   2820
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Crew ID:"
      Height          =   255
      Left            =   6840
      TabIndex        =   18
      Top             =   1710
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
      Left            =   6780
      TabIndex        =   17
      Top             =   60
      Width           =   1215
   End
   Begin VB.Label lblRowCount 
      Caption         =   "0 rows returned"
      Height          =   255
      Left            =   5220
      TabIndex        =   19
      Top             =   2880
      Width           =   3255
   End
End
Attribute VB_Name = "frmCrewGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_objGridMap As New CCrewMap ' Class to handle grid
Dim m_blnFirstSearch As Boolean ' Is this the first search we have made on this screen
Dim m_blnJumpIn As Boolean ' Are we jumping here from another screen
Dim m_rec As New ADODB.RecordSet ' Recordset to hold query results
Dim m_blnDoubleClick As Boolean ' Did a double click just occurr
Dim m_blnWereErrors As Boolean ' True if the Update had errors, used in QueryUnload
Dim m_blnMat_ID_Error As Boolean
Const USEBOOKMARK = 1
Const USECOORD = 0

Public strSource As String  'Source initiating this form
Dim m_strCurrentFormControl As String
Private Function CheckEntryErrors() As Boolean
Dim rec As ADODB.RecordSet
Dim strSelect As String
Dim blnReturn As Boolean

If Trim(TDBGrid.Columns("Crew Id").Text) = "" Or IsNull(TDBGrid.Columns("Crew Id").Text) Then
    MsgBox "Please enter a valid crew."
    CheckEntryErrors = True
Else
    strSelect = "select * from crew where crew_id = '" + Trim(TDBGrid.Columns("Crew Id").Value) + "'"
    blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rec)
    If blnReturn = False Then
        MsgBox "Please enter a valid crew."
        CheckEntryErrors = False
    Else
        If rec.RecordCount = 0 Then  'Not found
            MsgBox "Please enter a valid crew."
            CheckEntryErrors = True
        End If
    End If
    rec.Close
    Set rec = Nothing
End If

If CheckEntryErrors = False Then
    If Trim(UCase(TDBGrid.Columns("Type").Value)) <> "2" And Trim(UCase(TDBGrid.Columns("Type").Value)) <> "1" Then
        MsgBox "Please enter a valid Type."
        CheckEntryErrors = True
    End If
End If
'Validate the ID - Labor Trade or Equipment
If CheckEntryErrors = False Then
    If Trim(TDBGrid.Columns("Trade/Equip ID").Value) = "" Or IsNull(TDBGrid.Columns("Trade/Equip ID").Value) Then
        MsgBox "Please enter a valid Trade/Equip ID."
        CheckEntryErrors = True
    Else
        If UCase(TDBGrid.Columns("Type").Value) = "E" Then  'Equipment line
            strSelect = "select * from equipment where equip_id = '" + Trim(TDBGrid.Columns("Trade/Equip ID").Value) + "'"
        ElseIf UCase(TDBGrid.Columns("Type").Value) = "L" Then  'Equipment line
            strSelect = "select * from labor_trade where trade_id = '" + Trim(TDBGrid.Columns("Trade/Equip ID").Value) + "'"
        End If
        blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rec)
        If blnReturn = False Then
            MsgBox "Please enter a valid Trade/Equip ID."
            CheckEntryErrors = True
        Else
            If rec.RecordCount = 0 Then  'Not found
                MsgBox "Please enter a valid Trade/Equip ID."
                CheckEntryErrors = True
            End If
        End If
        rec.Close
        Set rec = Nothing
    End If
End If

End Function

Private Sub Clear_UC()
StartUnitCostID.Text = ""
EndUnitCostID.Text = ""

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

Private Sub chkIncludeTraces_Click()
load_crews
End Sub

' Handles Row Wrap feature
Private Sub ckbRowWrap_Click()
    m_objGridMap.RowWrap (ckbRowWrap)
End Sub

Private Sub cmdClone_Click()
    Dim rec As ADODB.RecordSet
    Set rec = m_objGridMap.CloneRow
End Sub

Private Sub cmdCloneCrew_Click()
    If IsNumeric(TDBGrid.Bookmark) = False Then
        MsgBox "You must select a row."
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    ' Navigate to single-record view
    Dim frm As frmCrew
    Set frm = New frmCrew
    frm.m_strCrewID = TDBGrid.Columns("Crew ID").CellText(TDBGrid.Bookmark)
    frm.JumpIn
    frm.m_blnCloneCrew = True
    frm.Show
    Screen.MousePointer = vbNormal


End Sub

Private Sub cmdDelete_Click()
    Dim varButton
    varButton = MsgBox("Are you sure you want to delete?", vbYesNo + vbCritical)
    If varButton = vbYes Then
        TDBGrid.Delete
        cmdSearch_Click
    End If
End Sub

Private Sub cmdEditCrew_Click()
    If IsNumeric(TDBGrid.Bookmark) = False Then
        MsgBox "You must select a row."
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    ' Navigate to single-record view
    Dim frm As frmCrew
    Set frm = New frmCrew
    frm.m_strCrewID = TDBGrid.Columns("Crew ID").CellText(TDBGrid.Bookmark)
    frm.JumpIn
    frm.Show
    Screen.MousePointer = vbNormal

End Sub

Private Sub cmdNew_Click()

    m_rec.AddNew
    m_rec.Fields("skey") = 0
    TDBGrid.Col = 0
    TDBGrid.ReBind
    TDBGrid.SetFocus
    TDBGrid.MoveLast
    TDBGrid.RefetchCol (0)

End Sub

Private Sub SetButtons(Mode As Single, Optional Coord As Variant)
Dim rsClone As ADODB.RecordSet
On Error GoTo Exit_Sub

Set rsClone = m_rec.Clone
Select Case Mode
    Case USEBOOKMARK
        If Not IsNull(TDBGrid.Bookmark) Then
            rsClone.Bookmark = TDBGrid.Bookmark
        End If
    Case USECOORD
        rsClone.Bookmark = TDBGrid.RowBookmark(TDBGrid.RowContaining(Coord))
End Select

If Not IsNull(TDBGrid.Bookmark) Then
    If TDBGrid.Bookmark > 0 Then 'valid bookmark
        If rsClone.Fields("skey") > 0 Then
            If rsClone.Fields("skey_type").Value = "2" Then 'Equipment row
                cmdEquipPrice.Enabled = True
            Else
                cmdEquipPrice.Enabled = False
            End If
            If rsClone.Fields("skey_type").Value = "1" Then 'Labor row
                cmdLaborRate.Enabled = True
            Else
                cmdLaborRate.Enabled = False
            End If
        Else
                cmdLaborRate.Enabled = False
                cmdEquipPrice.Enabled = False
        End If
    End If
'enable/disable Editing buttons
    If cmdLaborRate.Enabled = True Or cmdEquipPrice.Enabled = True Then
        cmdDelete.Enabled = True
        cmdClone.Enabled = True
    Else
        cmdDelete.Enabled = False
        cmdClone.Enabled = False
    End If
    cmdEditCrew.Enabled = True
    cmdCloneCrew.Enabled = True
Else
    cmdDelete.Enabled = False
    cmdClone.Enabled = False
    cmdCloneCrew.Enabled = False
    cmdEditCrew.Enabled = False
End If

Exit_Sub:
End Sub

Private Sub cmdEquipPrice_Click()
    If IsNumeric(TDBGrid.Bookmark) = False Then
        MsgBox "You must select a row."
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    ' Navigate to single-record view
    Dim frm As frmEquipRateGrid
    Set frm = New frmEquipRateGrid

    frm.JumpIn TDBGrid.Columns("Trade/Equip ID").CellText(TDBGrid.Bookmark)
    Screen.MousePointer = vbNormal

End Sub

Private Sub cmdLaborRate_Click()
    If IsNumeric(TDBGrid.Bookmark) = False Then
        MsgBox "You must select a row."
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    ' Navigate to single-record view
    Dim frm As frmLaborRateGrid
    Set frm = New frmLaborRateGrid

    frm.JumpIn TDBGrid.Columns("Trade/Equip ID").CellText(TDBGrid.Bookmark)
    Screen.MousePointer = vbNormal

End Sub


Private Sub cmdUpdate_Click()
    Dim blnRet As Boolean
    Dim vntBookmark As Variant
On Error GoTo Error_Processing

    m_blnWereErrors = False
    
    If TDBGrid.DataChanged = True Then
        If CheckEntryErrors = True Then
            Exit Sub
        End If
    End If
    Status ("Updating Crew Usage Information....")
    vntBookmark = TDBGrid.Bookmark
    TDBGrid.Update
    blnRet = m_objGridMap.Update
    If blnRet = False Then
        m_blnWereErrors = True
    End If
    cmdSearch_Click
Exit_Sub:
Exit Sub

Error_Processing:
'MsgBox Error$
Resume Exit_Sub

End Sub
Private Sub CrewID_Change()
If Len(CrewID) > 0 Then
    Clear_UC
End If

End Sub

Private Sub CrewID_Click()
Clear_UC
End Sub

Private Sub CrewID_LostFocus()
CrewID.Text = Trim(CrewID.Text)
End Sub


Private Sub EndUnitCostID_LostFocus()
EndUnitCostID.Text = Trim(EndUnitCostID.Text)
End Sub


Private Sub EquipmentID_Change()
If Len(EquipmentID) > 0 Then
    Clear_UC
End If
End Sub

Private Sub EquipmentID_Click()
Clear_UC
End Sub

Private Sub EquipmentID_LostFocus()
EquipmentID.Text = Trim(EquipmentID.Text)
End Sub

Private Sub Form_Deactivate()
m_strCurrentFormControl = Me.ActiveControl.Name
End Sub

Private Sub Form_Initialize()
    ' Fill the MasterFormat tree
    Status ("Loading Crew Maintenance...")
    m_blnFirstSearch = True
    Screen.MousePointer = vbHourglass
    If Forms(0).ActiveForm.Name = "frmCrewGrid" Then
        strSource = "CREW"
    ElseIf Forms(0).ActiveForm.Name = "frmMatPriceGrid" Then
        strSource = "Material"
    End If
    FormatTree.InitData g_cnShared, "UNITCOST"

        ' Initialize grid
    m_objGridMap.SetGrid TDBGrid
    m_objGridMap.strSource = strSource  '4/19/01
    m_objGridMap.InitGrid
    m_blnJumpIn = False
    Screen.MousePointer = vbNormal
    m_blnFirstSearch = False
    Status ("Ready")

End Sub

Private Sub Form_Load()

    Dim blnReturn As Boolean
    Dim strSelect As String
    Move START_LEFT, START_TOP, START_WIDTH, START_HEIGHT
    SetButtons USEBOOKMARK
    LoadCombos
    StartUnitCostID.Text = "~"
    cmdSearch_Click
    StartUnitCostID.Text = ""
    
End Sub

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
                TradeID.AddItem rsTemp![Trade_ID]
                rsTemp.MoveNext
            Loop
        End If
    End If
    rsTemp.Close

'Load Crews
load_crews
'Load Equipment
    strSelect = "select distinct equip_id from equipment order by equip_id;"

    ' Use DAL to perform select
    blnReturn = g_objDAL.GetRecordset(vbNullString, strSelect, rsTemp)
    If blnReturn = False Then
        MsgBox "An error occurred loading Equipment."
    Else
        If Not (rsTemp.EOF And rsTemp.BOF) Then
            Do Until rsTemp.EOF
                EquipmentID.AddItem rsTemp!equip_id
                rsTemp.MoveNext
            Loop
        End If
    End If
    rsTemp.Close

    CountryCode.ListIndex = 1   ' Select USA
    RegionCode.ListIndex = 0   ' Select Nat
    
    

End Sub
Private Sub load_crews()
Dim rsTemp As ADODB.RecordSet
Dim strSelect As String
Dim blnReturn As Boolean
    CrewID.Clear
    strSelect = "select distinct crew_id from crew where type_code = 'C' order by crew_id"
    ' Use DAL to perform select
    blnReturn = g_objDAL.GetRecordset(vbNullString, strSelect, rsTemp)
    If blnReturn = False Then
        MsgBox "An error occurred loading Crews."
    Else
        If Not (rsTemp.EOF And rsTemp.BOF) Then
            Do Until rsTemp.EOF
                CrewID.AddItem rsTemp![crew_id]
                rsTemp.MoveNext
            Loop
        End If
    End If
    rsTemp.Close

    TracesCrewID.Clear
    strSelect = "select distinct traces_crew_id from crew where type_code = 'C' order by traces_crew_id"
    ' Use DAL to perform select
    blnReturn = g_objDAL.GetRecordset(vbNullString, strSelect, rsTemp)
    If blnReturn = False Then
        MsgBox "An error occurred loading Traces Crews."
    Else
        If Not (rsTemp.EOF And rsTemp.BOF) Then
            Do Until rsTemp.EOF
                TracesCrewID.AddItem rsTemp![traces_crew_id]
                rsTemp.MoveNext
            Loop
        End If
    End If
    rsTemp.Close
End Sub

' Called when coming here from another screen
Public Sub JumpIn(strTradeID As String)
    TradeID.Text = strTradeID
    cmdSearch_Click
End Sub

' Called when coming here from another screen
Public Sub JumpIn2(strCrewID As String)
    CrewID.Text = strCrewID
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
            cmdEditCrew.Top = Me.Height - 1020
            cmdNew.Top = Me.Height - 1020
            cmdCloneCrew.Top = Me.Height - 1020
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
    Else
        rs.Close ' Make sure it is closed
        'Line of code was changed by Mohan on Jan 05,2012, MASTERFORMAT95_ID_HIERARCHY was changed to MASTERFORMAT04_ID_HIERARCHY
        strSelect = "select unit_cost_id_start, unit_cost_id_end from MASTERFORMAT04_ID_HIERARCHY where hier_id='" + strID + "'"
        
        
        blnReturn = g_objDAL.GetRecordset(vbNullString, strSelect, rs)
        StartUnitCostID.Text = rs.Fields("unit_cost_id_start")
        EndUnitCostID.Text = rs.Fields("unit_cost_id_end")
        ' Clear other boxes
        TradeID.Text = ""
        EquipmentID.Text = ""
        CrewID.Text = ""
        rs.Close
        StartUnitCostID.Refresh
        DoEvents
    End If
    ' Kick-off search
    cmdSearch_Click '4/19/01
End Sub

Private Sub cmdSearch_Click()
    On Error Resume Next
    Dim blnReturn As Boolean
    Dim strSelect As String
    Dim dtmToday As Date
    Dim dtmStart As Date
    Dim strError As String
    
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
    
    If Len(StartUnitCostID.Text) = 0 _
        And Len(EndUnitCostID.Text) = 0 _
        And Len(EquipmentID.Text) = 0 _
        And Len(TradeID.Text) = 0 _
        And Len(CrewID.Text) = 0 _
        And Len(TracesCrewID.Text) = 0 Then
        MsgBox "Please enter selection criteria."
        Exit Sub
    End If
    If Len(RegionCode.Text) = 0 _
        And Len(CountryCode.Text) = 0 Then
        MsgBox "Country and Region codes are required."
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    lblRowCount.Caption = "Working..."
    lblRowCount.Refresh
    Status ("Searching")
    
    ' Synch tree with text box
    If Not StartUnitCostID.Text = "" Then
        FormatTree.FocusItem (StartUnitCostID.Text)
    End If
    
    m_rec.Close ' Make sure it is closed
    m_rec.MaxRecords = MAX_RECORDS ' Set the maximum number to bring back
    dtmStart = Now
    
    strSelect = "exec sp_select_crew @crew_id='"
    If Len(CrewID.Text) > 0 Then
        strSelect = strSelect + SQLChangeWildcard(CrewID.Text) + "',  @traces_crew_id='"
    Else
        strSelect = strSelect + "%', @traces_crew_id='"
    End If
    
    If Len(TracesCrewID.Text) > 0 Then
        strSelect = strSelect + SQLChangeWildcard(TracesCrewID.Text) + "',  @trade_id='"
    Else
        strSelect = strSelect + "%', @trade_id='"
    End If
   
    If Len(TradeID.Text) > 0 Then
        strSelect = strSelect + SQLChangeWildcard(TradeID.Text) + "',  @equip_id='"
    Else
        strSelect = strSelect + "%', @equip_id='"
    End If
    
    If Len(EquipmentID.Text) > 0 Then
        strSelect = strSelect + SQLChangeWildcard(EquipmentID.Text) + "',  @start_unit_cost_id='"
    Else
        strSelect = strSelect + "%', @start_unit_cost_id='"
    End If
    
    If Len(StartUnitCostID.Text) > 0 Then
        If Right(StartUnitCostID.Text, 1) <> "*" And Len(Trim(EndUnitCostID)) = 0 Then
            StartUnitCostID = StartUnitCostID + "*"
        End If
        strSelect = strSelect + SQLChangeWildcard(StartUnitCostID.Text) + "',  @end_unit_cost_id='"
    Else
        strSelect = strSelect + "%', @end_unit_cost_id='"
    End If
    
    If Len(EndUnitCostID.Text) > 0 Then
        strSelect = strSelect + SQLChangeWildcard(EndUnitCostID.Text) + "',  @country_code='"
    Else
        strSelect = strSelect + "%', @country_code='"
    End If
    
    strSelect = strSelect + CountryCode + "',  @region_code='"
    strSelect = strSelect + RegionCode + "',  @mode='"

'Fill the mode based on the selection criteria entered:
'   1 = Crew only
'   2 = Crew/Trade
'   3 = Crew/Equipment
'   4 = Crew/Trade/Equipment
'   5 = Unit Cost (wildcard)
'   6 = Unit Cost Range

    If Len(StartUnitCostID.Text) > 0 And Right(StartUnitCostID.Text, 1) = "*" Then
        strSelect = strSelect + "5'"
    ElseIf Len(StartUnitCostID.Text) > 0 Then
        strSelect = strSelect + "6'"
    ElseIf Len(TradeID.Text) > 0 And Len(EquipmentID) > 0 Then
        strSelect = strSelect + "4'"
    ElseIf Len(TradeID.Text) > 0 And Len(EquipmentID) = 0 Then
        strSelect = strSelect + "2'"
    ElseIf Len(TradeID.Text) = 0 And Len(EquipmentID) > 0 Then
        strSelect = strSelect + "3'"
    ElseIf Len(TradeID.Text) = 0 And Len(EquipmentID) = 0 Then
        strSelect = strSelect + "1'"
    End If
    
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
    SetButtons USEBOOKMARK
    Status ("Ready")
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

Private Sub StartUnitCostID_LostFocus()
StartUnitCostID.Text = Trim(StartUnitCostID.Text)
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
Dim rsClone As ADODB.RecordSet
Set rsClone = m_rec.Clone
If Not IsNull(TDBGrid.Bookmark) Then
rsClone.Bookmark = TDBGrid.Bookmark
    ' If this is the mouse-up form a double click
    If m_blnDoubleClick Then
        ' Make sure it is the left button
        If Button = vbLeftButton Then
            m_blnDoubleClick = False
            ' Same function as clicking Material Price button, open single record view
            If rsClone.Fields("skey_type") = "2" Then
                cmdEquipPrice_Click
            ElseIf rsClone.Fields("skey_type") = "1" Then
                cmdLaborRate_Click
            End If
        End If
    Else
        If Button = vbRightButton And IsNumeric(TDBGrid.Bookmark) Then
            Dim strErrorMsg As String
            strErrorMsg = m_objGridMap.GetError(TDBGrid.Bookmark)
            If Len(strErrorMsg) > 0 Then
                MsgBox strErrorMsg
            End If
        End If
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

Private Sub TDBGrid_OnAddNew()
'' Defaults for new added row
'If left(MaterialID, 1) = "M" Then MaterialID = right(MaterialID, Len(MaterialID) - 1)
'If (Len(UnitCostID.Text) = 0 And Len(MaterialID) = 10) And right(MaterialID, 1) <> "*" Then
''    TDBGrid.Columns("Material ID").Value = "M" + MaterialID
''    TDBGrid.Columns("mat_skey").Value = GetMatSkey(TDBGrid.Columns("Material ID").Value)
'    TDBGrid.col = TDBGrid.Columns("Unit Cost ID").ColIndex
'ElseIf (Len(MaterialID) = 0 And Len(UnitCostID.Text) = 12) And right(UnitCostID.Text, 1) <> "*" Then
''    TDBGrid.Columns("Unit Cost ID").Value = UnitCostID.Text
''    TDBGrid.Columns("unit_cost_skey").Value = GetUCSkey(UnitCostID.Text)
'    TDBGrid.col = TDBGrid.Columns("Material ID").ColIndex
'End If
'If Len(MaterialID) > 0 Then
'    MaterialID = "M" + MaterialID
'End If
TDBGrid.Split = 0
TDBGrid.AllowAddNew = False
'TDBGrid.Columns("Input Factor").DefaultValue = 1
'TDBGrid.Columns("Output Factor").Value = 1
'TDBGrid.Columns("Adj Factor").Value = 1
'TDBGrid.Columns("Unit Qty").Value = 1
'TDBGrid.Columns("last_update_id") = 0
End Sub


Private Sub TradeID_Change()
If Len(Trim(TradeID)) > 0 Then
    Clear_UC
End If
End Sub


Private Sub TradeID_Click()
Clear_UC
End Sub

Private Sub TradeID_LostFocus()
TradeID.Text = Trim(TradeID.Text)
End Sub


