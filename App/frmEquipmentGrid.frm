VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{5936A75C-3F42-11D6-AF6B-AA0004005F12}#1.3#0"; "MeansCtrl.ocx"
Begin VB.Form frmEquipmentGrid 
   Caption         =   "Equipment Maintenance Grid"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11130
   Icon            =   "frmEquipmentGrid.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6855
   ScaleWidth      =   11130
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   495
      Left            =   7260
      TabIndex        =   7
      Top             =   6240
      Width           =   1150
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   495
      Left            =   5940
      TabIndex        =   6
      Top             =   6240
      Width           =   1150
   End
   Begin VB.CheckBox ckbRowWrap 
      Caption         =   "Row Wrap"
      Height          =   315
      Left            =   60
      TabIndex        =   3
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   495
      Left            =   8340
      TabIndex        =   2
      Top             =   1800
      Width           =   1150
   End
   Begin VB.TextBox Description 
      Height          =   315
      Left            =   8340
      TabIndex        =   1
      Top             =   1320
      Width           =   2475
   End
   Begin VB.TextBox EquipmentID 
      Height          =   315
      Left            =   8340
      TabIndex        =   0
      Top             =   840
      Width           =   1515
   End
   Begin VB.Frame Frame1 
      Caption         =   "Go To"
      Height          =   855
      Left            =   120
      TabIndex        =   10
      Top             =   6000
      Width           =   3255
      Begin VB.CommandButton cmdOutput 
         Caption         =   "&Output"
         Height          =   495
         Left            =   2160
         TabIndex        =   17
         Top             =   240
         Width           =   915
      End
      Begin VB.CommandButton cmdEquipmentRate 
         Caption         =   "Equip. Rate"
         Height          =   495
         Left            =   1200
         TabIndex        =   5
         Top             =   240
         Width           =   915
      End
      Begin VB.CommandButton cmdEquipment 
         Caption         =   "Equip. Maint."
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   8580
      TabIndex        =   8
      Top             =   6240
      Width           =   1150
   End
   Begin VB.CommandButton cmdClone 
      Caption         =   "Clone"
      Height          =   495
      Left            =   9900
      TabIndex        =   9
      Top             =   6240
      Width           =   1150
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid 
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
      Splits(0)._ColumnProps(5)=   "Column(0)._MinWidth=6"
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
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=30"
      _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=31"
      _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=32"
      _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=34"
      _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=35"
      _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "RecordSelectorStyle:id=37,.parent=2,.namedParent=39"
      _StyleDefs(19)  =   "FilterBarStyle:id=40,.parent=1,.namedParent=42"
      _StyleDefs(20)  =   "Splits(0).Style:id=11,.parent=1"
      _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=20,.parent=4"
      _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=12,.parent=2"
      _StyleDefs(23)  =   "Splits(0).FooterStyle:id=13,.parent=3"
      _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=14,.parent=5"
      _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=16,.parent=6"
      _StyleDefs(26)  =   "Splits(0).EditorStyle:id=15,.parent=7"
      _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=17,.parent=8"
      _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=18,.parent=9"
      _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=19,.parent=10"
      _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=38,.parent=37"
      _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=41,.parent=40"
      _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=24,.parent=11"
      _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12"
      _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15"
      _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=28,.parent=11"
      _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=15"
      _StyleDefs(40)  =   "Named:id=29:Normal"
      _StyleDefs(41)  =   ":id=29,.parent=0"
      _StyleDefs(42)  =   "Named:id=30:Heading"
      _StyleDefs(43)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(44)  =   ":id=30,.wraptext=-1"
      _StyleDefs(45)  =   "Named:id=31:Footing"
      _StyleDefs(46)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(47)  =   "Named:id=32:Selected"
      _StyleDefs(48)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(49)  =   "Named:id=33:Caption"
      _StyleDefs(50)  =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(51)  =   "Named:id=34:HighlightRow"
      _StyleDefs(52)  =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(53)  =   "Named:id=35:EvenRow"
      _StyleDefs(54)  =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
      _StyleDefs(55)  =   "Named:id=36:OddRow"
      _StyleDefs(56)  =   ":id=36,.parent=29"
      _StyleDefs(57)  =   "Named:id=39:RecordSelector"
      _StyleDefs(58)  =   ":id=39,.parent=30"
      _StyleDefs(59)  =   "Named:id=42:FilterBar"
      _StyleDefs(60)  =   ":id=42,.parent=29"
   End
   Begin ConstructionCostDatabase.DynaTree FormatTree 
      Height          =   2775
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   4895
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
      Caption         =   "Tech Desc:"
      Height          =   255
      Left            =   7020
      TabIndex        =   14
      Top             =   1380
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Equipment ID:"
      Height          =   255
      Left            =   7020
      TabIndex        =   13
      Top             =   900
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
   Begin VB.Label lblRowCount 
      Caption         =   "0 rows returned"
      Height          =   255
      Left            =   5160
      TabIndex        =   12
      Top             =   2880
      Width           =   3255
   End
End
Attribute VB_Name = "frmEquipmentGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' <modulename> frmEquipmentGrid.frm</modulename>
' <functionname>General (Main) </functionname>
'
' <summary>
' (CCI) EQUIPMENT MAINTENANCE GRID
'
'* * * WARNING * * *  WARNING * * *  WARNING * * *  * * * * * * * * * * * * * * * * * *
'i DON 'T SEE WHERE THIS IS BEING USED AT ALL IN "EQUIPMENT" FUNCTIONALITY !!!
'(try Main Window   Functions/Equipment/Maintenance
'
'Anyway…
'A significant amount of functionality is not working with this app and why it has been pushed into the background!
'Please approach Steve Plotner regarding functionality issues.
'It is my understanding (per K. R.) that this is routinely managed by way of a spreadsheet
'
'
'Display Equipment rental rates based upon "Equipment ID":
'"   Equipment ID
'or
'"   Tech Desc
'
'
'(BUTTONS)
'"   SEARCH              (CmdSearch_Click() )
'"   Equip Maint             (frmEquipment)
'"   Equip Rate              (frmEquipRateGrid)
'"   Output                  (dlgOutput)
'"   Update              (CEquipmentMap.Update() )
'"   New                 (frmEquipRate)
'"   Delete                  (TDBGrid.Delete)
'"   Clone                   (frmEquipment)
'
'
'Key Subs / Functions:
'"   CmdSearch_Click()
'???  (non-standard processing?!)
'
'
'HELPER Class: CEquipmentMap.Cls
' </summary>
'
' <seealso> CEquipmentMap.cls</seealso>
'<seealso> </seealso>
'
' <datastruct>m_rec</datastruct>
'<datastruct>m_objGridMap</datastruct>
'
' <storedprocedurename> sp_temp_output_init </storedprocedurename>
'<storedprocedurename> sp_temp_add_output_keys </storedprocedurename>
'
'
' <returns>N/A</returns>
' <exception>Always trap with an accompanying message box</exception>
' <example>
'<code> * * * SEARCH
'
'select equip_skey, type_code, equip_id, alt_equip_id, tech_desc, book_desc, crew_equip_desc, crew_equip_desc_plural, metric_tech_desc, metric_book_desc, metric_crew_equip_desc, metric_crew_equip_desc_plural, index_desc, index_code, unit, metric_unit, model_name, traces_ind, indent_code, format_characters, format_code, last_update_id as 'equip_last_update_id', last_update_date as 'equip_last_update_date', last_update_person as 'equip_last_update_person' from Equipment where  equip_id like '01%' ORDER BY equip_id
'
'</code>
'</example>
'<permission>Public</Permission>
'<dependson>This component depends on the following:
'1.  CEquipmentMap.cls
'2.  CGridMap.cls
'3.  CCDdal.CRSMDataAccess (
'4.  Access to the DAL (data access layer dll) opened in MainModule_Main() )
'</dependson>





Dim m_objGridMap As New CEquipmentMap ' Class to handle grid
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
    If TDBGrid.Columns(0).Value = "M" Or TDBGrid.Columns(0).Value = "E" Then
        MsgBox "To Clone an equipment that requires rate information, go to Equipment Rate."
        Exit Sub
    End If
    Dim rec As ADODB.RecordSet
    
    m_objGridMap.CloneRow
    ' Force any changes into recordset from grid
    TDBGrid.Update
    ' Navigate to single-record view
    Dim frm As frmEquipment
'    Dim rec As ADODB.RecordSet
    Set frm = New frmEquipment
    ' Make copy of recordset
'    Set rec = m_rec.Clone
    ' Get the selected row from grid
'    rec.Bookmark = TDBGrid.Bookmark
    frm.SetRow rec, True ' Pass the current record into the form
    frm.Show
Out:
End Sub

' Called when coming here from another screen
Public Sub JumpIn(strEquipID As String)
    EquipmentID.Text = strEquipID
    cmdSearch_Click
End Sub

Private Sub cmdDelete_Click()
    Dim varButton
    varButton = MsgBox("Are you sure you want to delete?", vbYesNo + vbCritical)
    If varButton = vbYes Then
        TDBGrid.Delete
    End If
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
    ' Open single record view with data from row selected
    Dim frm As frmEquipRateGrid
    Set frm = New frmEquipRateGrid
    frm.JumpIn TDBGrid.Columns("Equipment ID").CellText(TDBGrid.Bookmark)
End Sub

Private Sub cmdOutput_Click()
    Dim frm As Form
    Dim blnVisible As Boolean
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
        ' 7/11/2005 RTD - CHANGED SKEY_TYPE FROM 'A' TO 'E' PER KATHY RODRIGUEZ
        ' Valid values are A = assembly, SF = SquareFoot, E = equipment, U = unit
        strUpdate1 = "exec sp_temp_add_output_keys @skey_type = 'E', @skey = "
        If TDBGrid.SelBookmarks.Count = 0 Then  'No rows selected
            If Not IsNull(TDBGrid.Bookmark) Then    'Use current row
                m_rec.Bookmark = TDBGrid.Bookmark
                strUpdate = strUpdate1 + CStr(m_rec.Fields("equip_skey"))
                blnReturn = frm.m_objOutput.ExecQuery(vbNullString, strUpdate, strError)
            End If
        Else
            For Each varBookmark In TDBGrid.SelBookmarks
                m_rec.Bookmark = varBookmark
                strUpdate = strUpdate1 + CStr(m_rec.Fields("equip_skey"))
                blnReturn = frm.m_objOutput.ExecQuery(vbNullString, strUpdate, strError)
            Next varBookmark
        End If
        ' 7/11/2005 RTD - CHANGED SKEY_TYPE FROM 'A' TO 'E' PER KATHY RODRIGUEZ
        ' Valid values are A = assembly, SF = SquareFoot, E = equipment, U = unit
        frm.m_strKeyType = "E"
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

Private Sub Description_LostFocus()
    Description.Text = Trim(Description.Text)
End Sub


Private Sub EquipmentID_LostFocus()
    EquipmentID.Text = Trim(EquipmentID.Text)
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

Private Sub Form_Deactivate()
    m_strCurrentFormControl = Me.ActiveControl.Name
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim strSELECT As String
    Dim blnReturn As Boolean
    Move START_LEFT, START_TOP, START_WIDTH, START_HEIGHT
    
    ' This will never return any rows, just used to create recordset
'    strSelect = "select Equipment.*, Equipment.last_update_id as 'equip_last_update_id' from Equipment where "
'    strSelect = strSelect + "equip_id = '0'"
    EquipmentID.Text = "~"
    cmdSearch_Click
    EquipmentID.Text = ""
    
    
'    blnReturn = g_objDAL.GetRecordset(vbNullString, strSelect, m_rec)
'    m_objGridMap.RecordSet = m_rec
End Sub

Private Sub Form_Initialize()
    Screen.MousePointer = vbHourglass
    m_blnFirstSearch = True
    FormatTree.InitData g_cnShared, "EQUIPMENT"
    ' Initialize grid only once
    m_objGridMap.SetGrid TDBGrid
    m_objGridMap.InitGrid
    Screen.MousePointer = vbNormal
    m_blnFirstSearch = False
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
    Description.Text = ""
    ' Kick-off search
    cmdSearch_Click
End Sub

Private Sub cmdSearch_Click()
    On Error Resume Next
    Dim blnRet As Boolean
    Dim strSELECT As String
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
    
    If Len(EquipmentID.Text) = 0 And Len(Description.Text) = 0 Then
        Screen.MousePointer = vbNormal
        MsgBox "You must enter search criteria before searching."
        Exit Sub
    End If
    
    lblRowCount.Caption = "Working..."
    lblRowCount.Refresh
    
    strSELECT = "select equip_skey, type_code, equip_id, alt_equip_id, tech_desc, book_desc, crew_equip_desc, crew_equip_desc_plural, metric_tech_desc, metric_book_desc, " + _
        "metric_crew_equip_desc, metric_crew_equip_desc_plural, index_desc, index_code, unit, metric_unit, model_name, traces_ind, indent_code, " + _
        "format_characters, format_code, last_update_id as 'equip_last_update_id', last_update_date as 'equip_last_update_date', last_update_person as 'equip_last_update_person' from Equipment where " ' + _
        ' "start_date <= '" + Format(dtmToday, "mm/dd/yyyy") + "' and term_date >= '" + Format(dtmToday, "mm/dd/yyyy") + "'"
    
    If Not Len(EquipmentID.Text) = 0 Then
        strSELECT = strSELECT + " equip_id like '" + SQLChangeWildcard(EquipmentID.Text) + "'"
    End If
    If Not Len(Description.Text) = 0 Then
        If Not Len(EquipmentID.Text) = 0 Then
            strSELECT = strSELECT + " and"
        End If
        strSELECT = strSELECT + " tech_desc LIKE '" + SQLFixString(SQLChangeWildcard(Description.Text)) + "'"
    End If
    strSELECT = strSELECT + " ORDER BY equip_id"
    
    m_rec.Close ' Make sure it is closed
    m_rec.MaxRecords = MAX_RECORDS ' Set the maximum number to bring back
    dtmStart = Now
    ' Use g_objDAL to perform select
    blnRet = g_objDAL.GetRecordset(vbNullString, strSELECT, m_rec)
    If blnRet = False Then
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

Private Sub TDBGrid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    If ColIndex = 0 Then
        Dim strSELECT As String
        Dim rec As New ADODB.RecordSet ' Recordset to hold query results
        Dim blnRet As Boolean
        Dim I As Integer

        strSELECT = "Select * from Equipment where equip_id='" + TDBGrid.Text + "'"
        ' Use g_objDAL to perform select
        blnRet = g_objDAL.GetRecordset(vbNullString, strSELECT, rec)
        If rec.RecordCount > 0 Then
            m_rec.AddNew
            For I = 0 To rec.Fields.Count - 1
                m_rec.Fields(rec.Fields(I).Name) = rec.Fields(I).Value
            Next I
            Dim MyBookmark As Variant
            MyBookmark = m_rec.Bookmark
            TDBGrid.ReBind
            TDBGrid.Bookmark = MyBookmark
            Cancel = True
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

Private Sub TDBGrid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If this is the mouse-up form a double click
    If m_blnDoubleClick Then
        ' Make sure it is the left button
        If Button = vbLeftButton Then
            m_blnDoubleClick = False
            ' Same function as clicking Material Price button, open single record view
            cmdEquipment_Click
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



