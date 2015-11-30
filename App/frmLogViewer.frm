VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmLogViewer 
   Caption         =   "Log Viewer"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8895
   Icon            =   "frmLogViewer.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5790
   ScaleWidth      =   8895
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   7680
      TabIndex        =   5
      Top             =   240
      Width           =   1095
   End
   Begin VB.ComboBox cboLogEvent 
      Height          =   315
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
   Begin VB.ComboBox cboLogDate 
      Height          =   315
      Left            =   3000
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   2775
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   8493
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
      Splits(0).DividerColor=   13160660
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
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   13160660
      RowDividerColor =   13160660
      RowSubDividerColor=   13160660
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(44)  =   "Named:id=33:Normal"
      _StyleDefs(45)  =   ":id=33,.parent=0"
      _StyleDefs(46)  =   "Named:id=34:Heading"
      _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(48)  =   ":id=34,.wraptext=-1"
      _StyleDefs(49)  =   "Named:id=35:Footing"
      _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(51)  =   "Named:id=36:Selected"
      _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=37:Caption"
      _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(55)  =   "Named:id=38:HighlightRow"
      _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(57)  =   "Named:id=39:EvenRow"
      _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(59)  =   "Named:id=40:OddRow"
      _StyleDefs(60)  =   ":id=40,.parent=33"
      _StyleDefs(61)  =   "Named:id=41:RecordSelector"
      _StyleDefs(62)  =   ":id=41,.parent=34"
      _StyleDefs(63)  =   "Named:id=42:FilterBar"
      _StyleDefs(64)  =   ":id=42,.parent=33"
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Event"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   195
      Left            =   2520
      TabIndex        =   2
      Top             =   240
      Width           =   345
   End
End
Attribute VB_Name = "frmLogViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_rec As New ADODB.RecordSet
Dim sEvent As String
Dim iWeekday As Integer

Public Property Get EventType() As String
    EventType = sEvent
End Property
Public Property Let EventType(NewValue As String)
    sEvent = NewValue
    InitGrid
End Property

Private Function GetWeekday(ByVal sDateString As String) As Integer
    Dim dtDate As Date
    
    If IsDate(sDateString) Then
        dtDate = CDate(sDateString)
        GetWeekday = Weekday(dtDate)
    Else
        If InStr(sDateString, ",") > 0 Then
            sDateString = Mid(sDateString, InStr(sDateString, ",") + 1)
            If IsDate(sDateString) Then
                dtDate = CDate(sDateString)
                GetWeekday = Weekday(dtDate)
            End If
        End If
    End If

End Function

Public Sub RefreshGrid()
    Dim rs As ADODB.RecordSet
    Dim sFilter As String
    Static bRefreshing As Boolean
    
    On Error Resume Next
    If bRefreshing Then Exit Sub
    bRefreshing = True
    Screen.MousePointer = vbHourglass
    m_rec.Close
    If sEvent = "ALL" Then
        sFilter = ""
    Else
        sFilter = "[Event] = '" & sEvent & "'"
    End If
    ReadLogIntoRecordset iWeekday, sFilter, rs
    Set m_rec = rs
    TDBGrid.FilterActive = False
    TDBGrid.DataSource = m_rec
    TDBGrid.ReBind
    TDBGrid.ApproxCount = m_rec.RecordCount
    TDBGrid.Refresh
    SetColHeaders
    bRefreshing = False
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub SetColHeaders()
    Dim Col As TrueOleDBGrid80.Column
    Dim I As Long
    Dim aHeader As Variant

    On Error Resume Next
    Screen.MousePointer = vbHourglass
    Set Col = TDBGrid.Columns(0)
    Col.Caption = "Timestamp"
    Col.Visible = True
    Col.Locked = True
    Col.DataField = "Date"
    Col.Width = 1800
    Col.NumberFormat = "FormatText Event"
    
    Set Col = TDBGrid.Columns(1)
    Col.Caption = "Event"
    Col.Visible = True
    Col.Locked = True
    Col.DataField = "Event"
    Col.Width = 800

    Select Case sEvent
    Case "FAX"
        aHeader = Array("InfoSource", "Company", "Contact Name", "Last Name", "First Name", "Fax Number")
        For I = LBound(aHeader) To UBound(aHeader)
            Set Col = TDBGrid.Columns(I + 2)
            Col.Caption = aHeader(I)
            Col.Visible = True
            Col.Locked = True
            Col.DataField = "COL_" & Format(I + 1, "000")
            Col.Width = 1500
        Next
        For I = UBound(aHeader) + 1 To 10
            TDBGrid.Columns(I + 2).Visible = False
        Next
    Case "EMAIL"
        aHeader = Array("InfoSource", "Company", "Contact Name", "Last Name", "First Name", "E-Mail Address")
        For I = LBound(aHeader) To UBound(aHeader)
            Set Col = TDBGrid.Columns(I + 2)
            Col.Caption = aHeader(I)
            Col.Visible = True
            Col.Locked = True
            Col.DataField = "COL_" & Format(I + 1, "000")
            Col.Width = 1500
        Next
        For I = UBound(aHeader) + 1 To 10
            TDBGrid.Columns(I + 2).Visible = False
        Next
    End Select
        
    TDBGrid.Refresh
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub InitGrid()
   
    TDBGrid.ScrollBars = dbgAutomatic
    TDBGrid.AlternatingRowStyle = True
    TDBGrid.DeadAreaBackColor = vbApplicationWorkspace
    TDBGrid.OddRowStyle.BackColor = vbWindowBackground
    TDBGrid.EvenRowStyle.BackColor = g_intAlternateRowColor
    'TDBGrid.DataSource = m_rec
    RefreshGrid
    
End Sub

Private Sub cboLogDate_Change()
    iWeekday = GetWeekday(cboLogDate.Text)
    'RefreshGrid
End Sub

Private Sub cboLogDate_Click()
    iWeekday = GetWeekday(cboLogDate.Text)
    RefreshGrid
End Sub

Private Sub cboLogEvent_Change()
    sEvent = cboLogEvent.Text
    'RefreshGrid
End Sub

Private Sub cboLogEvent_Click()
    sEvent = cboLogEvent.Text
    RefreshGrid
End Sub

Private Sub cmdRefresh_Click()
    
    RefreshGrid
    
End Sub

Private Sub Form_Activate()
    ShowToolbarIcons True
End Sub

Private Sub Form_Deactivate()
    ShowToolbarIcons False
End Sub

Private Sub Form_Initialize()
    sEvent = "FAX"
    iWeekday = Weekday(Date)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ShowToolbarIcons False
    On Error Resume Next
    If m_rec.State = adStateOpen Then
        m_rec.Close
    End If
    Set m_rec = Nothing
    Set frmLogViewer = Nothing
End Sub

Private Sub Form_Load()
    Dim nRet As Long
    Dim I As Long
    Dim dtDate As Date
    
    cboLogEvent.AddItem "FAX"
    cboLogEvent.AddItem "EMAIL"
    nRet = SendMessage(cboLogEvent.hWnd, CB_FINDSTRING, 0, sEvent)
    If nRet >= 0 Then
        cboLogEvent.ListIndex = nRet
    End If
    
    dtDate = Date
    For I = 1 To 7
        cboLogDate.AddItem Format(dtDate, "Long Date")
        dtDate = DateAdd("d", -1, dtDate)
    Next
    cboLogDate.ListIndex = 0
    
    InitGrid
    
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    TDBGrid.Width = Me.Width - (TDBGrid.Left * 3)
    TDBGrid.Height = Me.Height - TDBGrid.Top - (TDBGrid.Left * 4)
    cmdRefresh.Left = Me.Width - cmdRefresh.Width - (TDBGrid.Left * 2)
    
End Sub

Private Sub ShowToolbarIcons(bShowIcons As Boolean)
    
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
    
End Sub

Public Function PreviewReport()
    
    TDBGrid.PrintInfo.PreviewCaption = "CCD Event Log Print Preview"
    TDBGrid.PrintInfo.PreviewInitHeight = 0
    TDBGrid.PrintInfo.PreviewInitWidth = 0
    TDBGrid.PrintInfo.PreviewInitScreenFill = 100
    TDBGrid.PrintInfo.PageHeader = "\tCCD EVENT LOG"
    TDBGrid.PrintInfo.PageHeaderFont.Bold = True
    TDBGrid.PrintInfo.PageHeaderFont.Size = 12
    TDBGrid.PrintInfo.PageFooter = CStr(Now) & "\t\tPage \p"
    TDBGrid.PrintInfo.SettingsOrientation = 2
    TDBGrid.PrintInfo.PrintPreview

End Function

Public Function PrintReport()

    TDBGrid.PrintInfo.PageHeader = "\tCCD EVENT LOG"
    TDBGrid.PrintInfo.PageHeaderFont.Bold = True
    TDBGrid.PrintInfo.PageHeaderFont.Size = 12
    TDBGrid.PrintInfo.PageFooter = CStr(Now) & "\t\tPage \p"
    TDBGrid.PrintInfo.SettingsOrientation = 2
    TDBGrid.PrintInfo.PrintData

End Function

Private Sub TDBGrid_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    If ColIndex = 0 Then
        'TIMESTAMP COLUMN - REMOVE "#" DELIMTERS
        If Left(Value, 1) = "#" Then
            Value = Mid(Value, 2)
        End If
        If Right(Value, 1) = "#" Then
            Value = Left(Value, Len(Value) - 1)
        End If
    End If
End Sub
