VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmInfoSourceGrid 
   Caption         =   "Information Source Grid"
   ClientHeight    =   6945
   ClientLeft      =   1845
   ClientTop       =   3210
   ClientWidth     =   11160
   Icon            =   "frmInfoSourceGrid.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6945
   ScaleWidth      =   11160
   Begin VB.ComboBox cboStateCode 
      Height          =   315
      Left            =   3360
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Tag             =   "1"
      Top             =   1800
      Width           =   675
   End
   Begin VB.ComboBox cboCountryCode 
      Height          =   315
      Left            =   6240
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Tag             =   "1"
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox txtKeyword 
      Height          =   315
      Left            =   1140
      TabIndex        =   1
      Top             =   960
      Width           =   2415
   End
   Begin VB.CommandButton cmdClear 
      Appearance      =   0  'Flat
      Caption         =   "&Clear"
      Height          =   435
      Left            =   2580
      TabIndex        =   9
      Top             =   2280
      Width           =   1150
   End
   Begin VB.CheckBox ckbTicklerDate 
      Caption         =   "Tickler Date "
      Height          =   255
      Left            =   4860
      TabIndex        =   17
      Top             =   160
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "        "
      Height          =   1095
      Left            =   4680
      TabIndex        =   31
      Top             =   180
      Width           =   2415
      Begin MSComCtl2.DTPicker DTPickerFrom 
         Height          =   315
         Left            =   720
         TabIndex        =   18
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   22937601
         CurrentDate     =   36201
      End
      Begin MSComCtl2.DTPicker DTPickerTo 
         Height          =   315
         Left            =   720
         TabIndex        =   19
         Top             =   660
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   22937601
         CurrentDate     =   36201
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "from"
         Height          =   255
         Left            =   220
         TabIndex        =   34
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "to"
         Height          =   255
         Left            =   220
         TabIndex        =   33
         Top             =   660
         Width           =   435
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Go To"
      Height          =   855
      Left            =   120
      TabIndex        =   30
      Top             =   6000
      Width           =   2595
      Begin VB.CommandButton cmdReport 
         Caption         =   "Report"
         Height          =   495
         Left            =   1320
         TabIndex        =   12
         Top             =   240
         Width           =   915
      End
      Begin VB.CommandButton cmdInfoSource 
         Caption         =   "Info Source"
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdClone 
      Caption         =   "Clone"
      Height          =   495
      Left            =   9900
      TabIndex        =   16
      Top             =   6240
      Width           =   1150
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   8580
      TabIndex        =   15
      Top             =   6240
      Width           =   1150
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   495
      Left            =   7260
      TabIndex        =   14
      Top             =   6240
      Width           =   1150
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   495
      Left            =   5940
      TabIndex        =   13
      Top             =   6240
      Width           =   1150
   End
   Begin VB.TextBox txtZipCode 
      Height          =   315
      Left            =   4680
      TabIndex        =   6
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox txtCity 
      Height          =   315
      Left            =   1140
      TabIndex        =   4
      Top             =   1800
      Width           =   1515
   End
   Begin VB.TextBox txtLastName 
      Height          =   315
      Left            =   4680
      TabIndex        =   3
      Top             =   1380
      Width           =   2415
   End
   Begin VB.TextBox txtContactId 
      Height          =   315
      Left            =   1140
      TabIndex        =   0
      Top             =   540
      Width           =   855
   End
   Begin VB.TextBox txtCompanyName 
      Height          =   315
      Left            =   1140
      TabIndex        =   2
      Top             =   1380
      Width           =   2415
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid 
      Height          =   2715
      Left            =   60
      TabIndex        =   20
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
   Begin VB.CheckBox ckbRowWrap 
      Caption         =   "Row Wrap"
      Height          =   315
      Left            =   60
      TabIndex        =   10
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearch 
      Appearance      =   0  'Flat
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   435
      Left            =   1140
      TabIndex        =   8
      Top             =   2280
      Width           =   1150
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Keyword:"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   32
      Top             =   1020
      Width           =   675
   End
   Begin VB.Label lblRowCount 
      Caption         =   "0 rows returned"
      Height          =   255
      Left            =   5340
      TabIndex        =   29
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
      Left            =   120
      TabIndex        =   28
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblCountryCode 
      Alignment       =   1  'Right Justify
      Caption         =   "Country:"
      Height          =   255
      Left            =   5520
      TabIndex        =   27
      Top             =   1860
      Width           =   615
   End
   Begin VB.Label lblZipCode 
      Alignment       =   1  'Right Justify
      Caption         =   "Zip:"
      Height          =   255
      Left            =   4200
      TabIndex        =   26
      Top             =   1860
      Width           =   375
   End
   Begin VB.Label lblStateCode 
      Alignment       =   1  'Right Justify
      Caption         =   "State:"
      Height          =   255
      Left            =   2820
      TabIndex        =   25
      Top             =   1860
      Width           =   435
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "City:"
      Height          =   255
      Left            =   720
      TabIndex        =   24
      Top             =   1860
      Width           =   315
   End
   Begin VB.Label lblName 
      Alignment       =   1  'Right Justify
      Caption         =   "Last Name:"
      Height          =   255
      Left            =   3780
      TabIndex        =   23
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Contact Id:"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   22
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Company:"
      Height          =   255
      Index           =   0
      Left            =   300
      TabIndex        =   21
      Top             =   1440
      Width           =   735
   End
   Begin VB.Line Line2 
      X1              =   60
      X2              =   11040
      Y1              =   2820
      Y2              =   2820
   End
End
Attribute VB_Name = "frmInfoSourceGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_objGridMap As New CInfoSourceMap ' Class to handle grid
Dim m_blnFirstSearch As Boolean
Dim m_rec As New ADODB.RecordSet ' Recordset to hold query results
Dim m_blnDoubleClick As Boolean
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
    Dim frm As frmInfoSource
    Set frm = New frmInfoSource
    frm.SetRow rec, True ' Pass the current record into the form
    frm.Show
Out:
End Sub

' Called when coming here from another screen
Public Sub JumpIn(strInfoID As String)
    txtContactId.Text = strInfoID
    cmdSearch_Click
    txtContactId.Text = Left(strInfoID, 6)
End Sub

Private Sub cmdDelete_Click()
    Dim varButton
    varButton = MsgBox("Are you sure you want to delete?", vbYesNo + vbCritical)
    If varButton = vbYes Then
        TDBGrid.Delete
    End If
End Sub

Private Sub cmdInfoSource_Click()
    If IsNumeric(TDBGrid.Bookmark) = False Then
        MsgBox "You must select a row."
        Exit Sub
    End If
    ' Navigate to single-record view
    Dim frm As frmInfoSource
    Dim rec As ADODB.RecordSet
    Set frm = New frmInfoSource
    ' Make copy of recordset
    Set rec = m_rec.Clone
    ' Get the selected row from grid
    rec.Bookmark = TDBGrid.Bookmark
    frm.SetRow rec ' Pass the current record into the form
    frm.Show
End Sub

Private Sub cmdNew_Click()
    On Error GoTo Out
    Dim rec As New ADODB.RecordSet
    
    CopyRSFields rec, m_rec
    ' Open empty single record view
    Dim frm As frmInfoSource
    Set frm = New frmInfoSource
    ' Force any changes into recordset from grid
    TDBGrid.Update
    frm.SetRow rec, True
    frm.Show
Out:
End Sub

Private Sub cmdReport_Click()
    PreviewReport
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

Private Sub cmdClear_Click()
    txtContactId = ""
    txtLastName = ""
    txtCompanyName = ""
    cboStateCode.ListIndex = -1
    cboCountryCode.ListIndex = -1
    txtZipCode = ""
    txtCity = ""
    ckbTicklerDate.Value = 0
End Sub

Private Sub Form_Deactivate()
    m_strCurrentFormControl = Me.ActiveControl.Name
    ShowToolbarIcons False
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim strSELECT As String
    Dim blnReturn As Boolean
    Dim rec As ADODB.RecordSet
    Move START_LEFT, START_TOP, START_WIDTH, START_HEIGHT
    
    g_objDAL.GetRecordset vbNullString, "select state_code from state_country", rec
    While Not rec.EOF
        cboStateCode.AddItem (rec.Fields("state_code").Value)
        rec.MoveNext
    Wend
    rec.Close
    g_objDAL.GetRecordset vbNullString, "select country_code from country", rec
    While Not rec.EOF
        cboCountryCode.AddItem (rec.Fields("country_code").Value)
        rec.MoveNext
    Wend
    
    txtContactId.Text = "~"
    cmdSearch_Click
    txtContactId.Text = ""
    
    DTPickerFrom.Value = Date
    DTPickerTo.Value = Date
    
'    blnReturn = g_objDAL.GetRecordset(vbNullString, strSelect, m_rec)
'    m_objGridMap.RecordSet = m_rec
    Status ("")
    
End Sub

Private Sub Form_Initialize()
    Status ("Loading Information Source...")
    ' Initialize grid only once
    m_objGridMap.SetGrid TDBGrid
    m_objGridMap.InitGrid
    m_blnFirstSearch = True
End Sub

Private Sub cmdSearch_Click()
    On Error Resume Next
    Dim blnRet As Boolean
    Dim strSELECT As String
    Dim dtmToday As Date
    Dim blnAnd As Boolean
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
    blnAnd = False
    
    lblRowCount.Caption = "Working..."
    lblRowCount.Refresh
    
    dtmToday = Date
    
    strSELECT = "select * from Information_source where " ' + _
        ' "start_date <= '" + Format(dtmToday, "mm/dd/yyyy") + "' and term_date >= '" + Format(dtmToday, "mm/dd/yyyy") + "'"
    
    If Not txtContactId.Text = "" Then
        ' ADDED 6/16/2005 RTD FOR VERSION 7.4.0 (CR#292)
        ' IF CONTACT CONTAINS "," SEPARATED LIST, THEN USE THE 'IN' KEYWORD
        If InStr(txtContactId.Text, ",") > 0 Then
            strSELECT = strSELECT + " contact_id IN (" + SQLChangeWildcard(DelimitList(txtContactId.Text)) + ")"
        Else
            strSELECT = strSELECT + " contact_id LIKE '" + SQLChangeWildcard(txtContactId.Text) + "'"
        End If
        blnAnd = True
    End If
    If Not txtKeyword.Text = "" Then
        If blnAnd Then
            strSELECT = strSELECT + " and"
        End If
        strSELECT = strSELECT + " keyword LIKE '" + SQLFixString(SQLChangeWildcard(txtKeyword.Text)) + "'"
        blnAnd = True
    End If
    If Not txtLastName.Text = "" Then
        If blnAnd Then
            strSELECT = strSELECT + " and"
        End If
        strSELECT = strSELECT + " last_name LIKE '" + SQLChangeWildcard(txtLastName.Text) + "'"
        blnAnd = True
    End If
    If Not txtCompanyName.Text = "" Then
        If blnAnd Then
            strSELECT = strSELECT + " and"
        End If
        strSELECT = strSELECT + " company_name LIKE '" + SQLFixString(SQLChangeWildcard(txtCompanyName.Text)) + "'"
        blnAnd = True
    End If
    If Not txtCity.Text = "" Then
        If blnAnd Then
            strSELECT = strSELECT + " and"
        End If
        strSELECT = strSELECT + " city LIKE '" + SQLChangeWildcard(txtCity.Text) + "'"
        blnAnd = True
    End If
    If Not txtZipCode.Text = "" Then
        If blnAnd Then
            strSELECT = strSELECT + " and"
        End If
        strSELECT = strSELECT + " zip_code LIKE '" + SQLChangeWildcard(txtZipCode.Text) + "'"
        blnAnd = True
    End If
    If Not cboStateCode.Text = "" Then
        If blnAnd Then
            strSELECT = strSELECT + " and"
        End If
        strSELECT = strSELECT + " state_code LIKE '" + SQLChangeWildcard(cboStateCode.Text) + "'"
        blnAnd = True
    End If
    If Not cboCountryCode.Text = "" Then
        If blnAnd Then
            strSELECT = strSELECT + " and"
        End If
        strSELECT = strSELECT + " country_code LIKE '" + SQLChangeWildcard(cboCountryCode.Text) + "'"
        blnAnd = True
    End If
    If ckbTicklerDate Then
        If blnAnd Then
            strSELECT = strSELECT + " and"
        End If
        strSELECT = strSELECT + " tickler_date>='" + Format(DTPickerFrom.Value, "mm/dd/yyyy") + "'"
        strSELECT = strSELECT + " and"
        strSELECT = strSELECT + " tickler_date<='" + Format(DTPickerTo.Value, "mm/dd/yyyy") + "'"
        blnAnd = True
    End If
    
    strSELECT = strSELECT + " order by contact_id"
    
    If Not blnAnd = True Then
        MsgBox "You must enter search criteria before searching."
        lblRowCount.Caption = "0 rows returned."
        GoTo Exit_Sub
    End If
    
    m_rec.Close ' Make sure it is closed
    m_rec.MaxRecords = MAX_RECORDS ' Set the maximum number to bring back
    dtmStart = Now
    ' Use g_objDAL to perform select
    blnRet = g_objDAL.GetRecordset(vbNullString, strSELECT, m_rec)
    If blnRet = False Then
        MsgBox "An error occurred while searching."
        lblRowCount.Caption = "0 rows returned."
        GoTo Exit_Sub
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
    TDBGrid.Columns("Company").AutoSize
    
Exit_Sub:
    Screen.MousePointer = vbNormal

End Sub

Private Sub Form_LostFocus()
    TDBGrid.Update
    HideGridSort
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

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbNormal Or Me.WindowState = vbMaximized Then
        If Me.Width > 11250 Then
            TDBGrid.Width = Me.Width - 255
            Line2.X2 = Me.Width - 210
        Else
            Me.Width = 11250
        End If
        
        If Me.Height > 7260 Then
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
    ShowToolbarIcons False
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
    If m_blnDoubleClick Then
        ' Make sure it is the left button
        If Button = vbLeftButton Then
            m_blnDoubleClick = False
            ' Same function as clicking Material Price button, open single record view
            cmdInfoSource_Click
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
    ShowToolbarIcons True
End Sub

Private Sub txtCity_LostFocus()
    txtCity = Trim(txtCity)
End Sub

Private Sub txtCompanyName_LostFocus()
    txtCompanyName = Trim(txtCompanyName)
End Sub

Private Sub txtContactId_LostFocus()
    txtContactId = Trim(txtContactId)
End Sub

Private Sub txtKeyword_LostFocus()
    txtKeyword = Trim(txtKeyword)
End Sub

Private Sub txtLastName_LostFocus()
    txtLastName = Trim(txtLastName)
End Sub

Private Sub txtZipCode_LostFocus()
    txtZipCode = Trim(txtZipCode)
End Sub

Private Function DelimitList(sList As String) As String
' ADDED 6/16/2005 FOR VERSION 7.4.0 CR#292
' SURROUNDS A COMMA-SEPARATED LIST WITH SINGLE-QUOTES
    Dim sTemp As String
    Dim aList As Variant
    Dim I As Long
    
    If InStr(sList, ",") = 0 Then
        sTemp = "'" & sList & "'"
    Else
        aList = Split(sList, ",")
        For I = LBound(aList) To UBound(aList)
            sTemp = sTemp & ",'" & aList(I) & "'"
        Next
        sTemp = Mid(sTemp, 2)
    End If
    DelimitList = sTemp
    
End Function

Public Function PrintReport()
    PreviewReport
End Function

Public Function PreviewReport()
    Dim fPreviewWindow As New frmReportPreview
    
    If m_rec.RecordCount > 0 Then
        fPreviewWindow.ReportName = "Information Sources"
        fPreviewWindow.ReportFile = "rptInfoSources.xml"
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
