VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmUCostHistoryGrid 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unit Cost History"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10830
   Icon            =   "frmUCostHistoryGrid.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   10830
   Begin VB.CheckBox ckbRowWrap 
      Caption         =   "Row Wrap"
      Height          =   315
      Left            =   180
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtUnitCostID 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1380
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid 
      Height          =   2955
      Left            =   120
      TabIndex        =   2
      Top             =   1020
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   5212
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
      Splits(0)._ColumnProps(9)=   "Column(1)._MinWidth=149299300"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
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
   Begin VB.Label lblRowCount 
      Caption         =   "0 rows returned"
      Height          =   255
      Left            =   5340
      TabIndex        =   4
      Top             =   660
      Width           =   2115
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Unit Cost ID:"
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   180
      Width           =   1215
   End
End
Attribute VB_Name = "frmUCostHistoryGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_objGridMap As New CUCostHistMap ' Class to handle grid
Dim m_blnFirstSearch As Boolean
Dim m_rec As New ADODB.RecordSet ' Recordset to hold query results
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
    m_blnMasterFormatNotSpecified = False
End Property

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

Private Sub Form_Deactivate()
    m_strCurrentFormControl = Me.ActiveControl.Name
    ShowToolbarIcons False
End Sub

Private Sub Form_Initialize()
    
    m_intMasterFormat = 1995
    m_blnMasterFormatNotSpecified = True
    
    m_blnFirstSearch = True
    
    ' Initialize grid only once
    m_objGridMap.SetGrid TDBGrid
    m_objGridMap.InitGrid
    
End Sub

' Called when coming here from another screen
Public Sub JumpIn(strUnitCostId As String)
    txtUnitCostID.Text = strUnitCostId
    Search
End Sub

Private Sub Search()
    Dim blnRet As Boolean
    Dim strSELECT As String
    Dim dtmToday As Date
    Dim strType As String
    Dim strSkey As String
    
    'MODIFIED 8/5/2005 RTD - TO SUPPORT MASTERFORMAT VERSION
    Select Case MasterFormat
    Case ALT_MASTERFORMAT_VERSION
        strSELECT = "select type_code, unit_cost_skey from unit_cost_detail where alt_unit_cost_id = '" + Compress_String(txtUnitCostID.Text) + "'"
    Case UCD_MASTERFORMAT_VERSION
        strSELECT = "select type_code, unit_cost_skey from unit_cost_detail where unit_cost_id = '" + Compress_String(txtUnitCostID.Text) + "'"
    Case EXT_MASTERFORMAT_VERSION
        strSELECT = "SELECT ucd.type_code, ucd.unit_cost_skey " & _
                    "FROM unit_cost_detail ucd, unit_cost_detail_ext ucdx " & _
                    "WHERE ucd.unit_cost_skey = ucdx.unit_cost_skey " & _
                    " AND ucdx.unit_cost_id = '" + Compress_String(txtUnitCostID.Text) + "'"
    Case Else
        strSELECT = "select type_code, unit_cost_skey from unit_cost_detail where unit_cost_id = '" + Compress_String(txtUnitCostID.Text) + "'"
    End Select
    ' Use g_objDAL to perform select
    blnRet = g_objDAL.GetRecordset(vbNullString, strSELECT, m_rec)
    If Not m_rec.EOF Then
        strType = m_rec.Fields(0)
        strSkey = m_rec.Fields(1)
    End If
    m_rec.Close

    ' Different selects based on type
    If strType = "M" Then
        strSELECT = "select * from published_unit_cost where "
        
        strSELECT = strSELECT + "unit_cost_skey = "
        strSELECT = strSELECT + strSkey
        strSELECT = strSELECT + " order by term_date desc"
        ' Remove the other split
        TDBGrid.Splits.Remove (2)
    ElseIf strType = "E" Then
        strSELECT = "select op_code, country_code, region_code, start_date, term_date, pct_ind, unit as unit_x, metric_unit as metric_unit_x, " + _
        "mat_cost as mat_cost_x, labor_cost as labor_cost_x, equip_cost as equip_cost_x, total_cost as total_cost_x, total_cost_op as total_cost_op_x, " + _
        "metric_mat_cost as metric_mat_cost_x, metric_labor_cost as metric_labor_cost_x, metric_equip_cost as metric_equip_cost_x, metric_total_cost as metric_total_cost_x, metric_total_cost_op as metric_total_cost_op_x, " + _
        "last_update_date as last_update_date_x, last_update_person as last_update_person_x " + _
        "from published_unit_cost_exception where "
        
        strSELECT = strSELECT + "unit_cost_skey = "
        strSELECT = strSELECT + strSkey
        strSELECT = strSELECT + " order by term_date desc"
        ' Remove the other split
        TDBGrid.Splits.Remove (1)
    Else
        MsgBox "No history available."
        Exit Sub
    End If
    
    ' Use g_objDAL to perform select
    blnRet = g_objDAL.GetRecordset(vbNullString, strSELECT, m_rec)
    
    ' Pass recordset to handler class
    m_objGridMap.RecordSet = m_rec
        
    lblRowCount.Caption = str(m_rec.RecordCount) + " rows returned"
    
    ' Reset the grid contents
    TDBGrid.Bookmark = Null
    TDBGrid.ReBind
    TDBGrid.ApproxCount = m_rec.RecordCount
End Sub

Private Sub Form_Load()
    Move START_LEFT, START_TOP
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
        ShowToolbarIcons True
    End If
End Sub

Private Sub Form_LostFocus()
    TDBGrid.Update
    HideGridSort
End Sub

Private Sub Form_Unload(Cancel As Integer)
    HideGridSort
    ShowToolbarIcons False
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

Public Function SetupGridPrint()
' SETUP PROPERTIES FOR TRUE DBGRID PRINTING

    TDBGrid.PrintInfo.PreviewCaption = Me.Caption & " Preview"
    TDBGrid.PrintInfo.PageHeader = "\t" & Me.Caption & " – " & Me.txtUnitCostID.Text
    
    TDBGrid.PrintInfo.PreviewInitHeight = START_HEIGHT / Screen.TwipsPerPixelX
    TDBGrid.PrintInfo.PreviewInitWidth = START_WIDTH / Screen.TwipsPerPixelY
    TDBGrid.PrintInfo.PreviewInitPosX = 5 + (fMainForm.Left / Screen.TwipsPerPixelX)
    TDBGrid.PrintInfo.PreviewInitPosY = 4 + ((fMainForm.Top + fMainForm.sbStatusBar.Height + fMainForm.tbToolBar.Height * 2) / Screen.TwipsPerPixelY)
    TDBGrid.PrintInfo.RepeatColumnHeaders = True
    TDBGrid.PrintInfo.PageHeaderFont.Bold = True
    TDBGrid.PrintInfo.PageHeaderFont.Size = 12
    TDBGrid.PrintInfo.PageFooter = CStr(Now) & "\t\tPage \p"
    ' ORIENTATION 1=PORTRAIT | 2=LANDSCAPE
    TDBGrid.PrintInfo.SettingsOrientation = 2
    TDBGrid.PrintInfo.SettingsMarginBottom = 720
    TDBGrid.PrintInfo.SettingsMarginTop = 720
    
End Function

Public Function PreviewReport()
'PREVIEW THE GRID TO THE SCREEN

    SetupGridPrint
    TDBGrid.PrintInfo.PrintPreview

End Function

Public Function PrintReport()
'SEND THE GRID TO THE PRINTER

    SetupGridPrint
    TDBGrid.PrintInfo.PrintData

End Function


