VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{5936A75C-3F42-11D6-AF6B-AA0004005F12}#1.3#0"; "MeansCtrl.ocx"
Begin VB.Form frmMatUsageGrid 
   Caption         =   "Material Usage Grid"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11130
   Icon            =   "frmMatUsageGrid.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   11130
   Visible         =   0   'False
   Begin VB.ComboBox cboMasterFormat 
      Height          =   315
      Left            =   8340
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1560
      Width           =   1515
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   495
      Left            =   5760
      TabIndex        =   17
      Top             =   6240
      Width           =   1150
   End
   Begin VB.TextBox MaterialID 
      Height          =   315
      Left            =   8340
      TabIndex        =   0
      Top             =   540
      Width           =   1515
   End
   Begin VB.TextBox UnitCostID 
      Height          =   315
      Left            =   8340
      TabIndex        =   1
      Top             =   1140
      Width           =   1515
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   495
      Left            =   8340
      TabIndex        =   3
      Top             =   2160
      Width           =   1275
   End
   Begin VB.CheckBox ckbRowWrap 
      Caption         =   "Row Wrap"
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   495
      Left            =   7220
      TabIndex        =   7
      Top             =   6240
      Width           =   1150
   End
   Begin VB.Frame Frame1 
      Caption         =   "Go To"
      Height          =   855
      Left            =   120
      TabIndex        =   11
      Top             =   6000
      Width           =   2835
      Begin VB.CommandButton cmdUnitCost 
         Caption         =   "Uni&t Cost"
         Height          =   495
         Left            =   1440
         TabIndex        =   6
         Top             =   240
         Width           =   1035
      End
      Begin VB.CommandButton cmdMaterialPrice 
         Caption         =   "&Material Price"
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   495
      Left            =   8560
      TabIndex        =   8
      Top             =   6240
      Width           =   1150
   End
   Begin VB.CommandButton cmdClone 
      Caption         =   "&Clone"
      Height          =   495
      Left            =   9900
      TabIndex        =   9
      Top             =   6240
      Width           =   1150
   End
   Begin ConstructionCostDatabase.DynaTree FormatTree 
      Height          =   2775
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   4895
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid 
      Height          =   2715
      Left            =   60
      TabIndex        =   16
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
      Splits(0)._ColumnProps(5)=   "Column(0)._MinWidth=6488064"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(1)._MinWidth=-2147483633"
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
      Alignment       =   1  'Right Justify
      Caption         =   "MasterFormat:"
      Height          =   255
      Left            =   6900
      TabIndex        =   18
      Top             =   1620
      Width           =   1335
   End
   Begin VB.Label lblRowCount 
      Caption         =   "0 rows returned"
      Height          =   255
      Left            =   5220
      TabIndex        =   15
      Top             =   2880
      Width           =   3255
   End
   Begin VB.Line Line1 
      X1              =   6660
      X2              =   6660
      Y1              =   2700
      Y2              =   0
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   10680
      Y1              =   2820
      Y2              =   2820
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Material ID:"
      Height          =   255
      Left            =   7020
      TabIndex        =   14
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblUnitCostId 
      Alignment       =   1  'Right Justify
      Caption         =   "Unit Cost ID:"
      Height          =   255
      Left            =   7020
      TabIndex        =   13
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
      Left            =   6900
      TabIndex        =   12
      Top             =   60
      Width           =   1215
   End
End
Attribute VB_Name = "frmMatUsageGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_objGridMap As New CMatUsageMap ' Class to handle grid
Dim m_blnFirstSearch As Boolean ' Is this the first search we have made on this screen
Dim m_blnJumpIn As Boolean ' Are we jumping here from another screen
Dim m_rec As New ADODB.RecordSet ' Recordset to hold query results
Dim m_blnDoubleClick As Boolean ' Did a double click just occurr
Dim m_blnWereErrors As Boolean ' True if the Update had errors, used in QueryUnload
Dim m_blnMat_ID_Error As Boolean
Dim m_intMasterFormat As Long   ' Stores MasterFormat version to use by Search et al
Dim m_blnMasterFormatNotSpecified As Boolean    ' True if MF was never explicitly set

Public strSource As String  'Source initiating this form
Dim m_strCurrentFormControl As String

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

Public Sub SelectAllRows()
    ' Pass recordset to handler class
    m_objGridMap.RecordSet = m_rec
    
    If m_rec.RecordCount > 0 Then
        m_objGridMap.SelectAllRows
    End If
End Sub

Private Sub cboMasterFormat_Click()
    MasterFormatChanged
End Sub

' Handles Row Wrap feature
Private Sub ckbRowWrap_Click()
    m_objGridMap.RowWrap (ckbRowWrap)
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
    On Error Resume Next
    TDBGrid.Delete
End Sub

Private Sub cmdMaterialPrice_Click()
    ' Navigate to grid view
    If IsNumeric(TDBGrid.Bookmark) Then
        Dim frm As frmMatPriceGrid
        Set frm = New frmMatPriceGrid
        frm.JumpIn Compress_String(TDBGrid.Columns("Material ID").CellText(TDBGrid.Bookmark)) + "*"
    Else
        MsgBox "Please select a row first."
    End If
End Sub

Private Sub cmdNew_Click()
' Create a variable to hold number of Visual Basic forms loaded
' and visible.

' Open empty single record view
'    Dim frm As frmMatPrice
'    Set frm = New frmMatPrice
'    frm.Show
Dim bln_Continue As Boolean
Dim varCurrentM_recBookmark As Variant
Dim MatID As String

'MODIFIED 8/25/2005 RTD - COMPRESS MATERIAL STRING
MaterialID.Text = Compress_String(MaterialID.Text)
If Left(MaterialID.Text, 1) = "M" Then
    MatID = Right(MaterialID.Text, 12)
Else
    MatID = MaterialID.Text
End If

If IsNull(TDBGrid.Bookmark) Then
        bln_Continue = True
Else
    'If ValidGridRow() = True Then
        bln_Continue = True
    'End If
End If

If bln_Continue = True Then
    TDBGrid.SetFocus
    TDBGrid.MoveLast
    TDBGrid.AllowAddNew = True
    TDBGrid.Row = TDBGrid.Row + 1
    'MODIFIED 8/25/2005 RTD - COMPRESS MATERIAL STRING
    If Len(MaterialID.Text) > 0 And Left(MaterialID.Text, 1) <> "M" Then
        MaterialID.Text = "M" + Compress_String(MaterialID.Text)
    End If
    TDBGrid.Split = 0
    m_rec.AddNew
    m_rec.MoveLast
    varCurrentM_recBookmark = m_rec.Bookmark
    m_rec.Fields("input_factor").Value = 1
    m_rec.Fields("output_factor").Value = 1
    m_rec.Fields("adj_factor").Value = 1
    m_rec.Fields("unit_qty").Value = 1
    m_rec.Fields("last_update_id").Value = 0
    m_rec.Fields("last_update_person").Value = strUserName
    If Left(MaterialID, 1) = "M" Then MaterialID = Right(MaterialID, Len(MaterialID) - 1)
    If (Len(UnitCostID.Text) = 0 And Len(MaterialID) = 12) And Right(MaterialID, 1) <> "*" Then
        m_rec.Fields("mat_id").Value = "M" + MaterialID
        m_rec.Fields("mat_skey").Value = GetMatSkey(m_rec.Fields("mat_id").Value)
    ElseIf (Len(MaterialID) = 0 And Len(UnitCostID.Text) = 12) And Right(UnitCostID.Text, 1) <> "*" Then
        'UPDATED 8/25/2004 RTD TO SUPPORT MASTERFORMAT 2004
        m_rec.Fields("unit_cost_skey").Value = GetUCSkey(UnitCostID.Text, MasterFormat)
        If MasterFormat = EXT_MASTERFORMAT_VERSION Then
            m_rec.Fields("ext_unit_cost_id").Value = Compress_String(UnitCostID.Text)
        Else
            m_rec.Fields("unit_cost_id").Value = Compress_String(UnitCostID.Text)
        End If
    End If
    
    m_objGridMap.SetRowState m_rec.Bookmark, STATE_NEW
    
    TDBGrid.SetFocus
    TDBGrid.AllowAddNew = False
    TDBGrid.ReOpen m_rec.Bookmark
    m_rec.MoveLast
    DoEvents
    ' Defaults for new added row
    If Left(MaterialID, 1) = "M" Then MaterialID = Mid(Compress_String(MaterialID), 2)
   
        If (Len(UnitCostID.Text) = 0 And Len(MaterialID) = 12) And Right(MaterialID, 1) <> "*" Then
        
            TDBGrid.Columns("Material ID").Value = "M" + MaterialID
            
            TDBGrid.Columns("mat_skey").Value = GetMatSkey(TDBGrid.Columns("Material ID").Value)
            
            TDBGrid.Col = TDBGrid.Columns("Unit Cost ID").ColIndex
            
            m_objGridMap.FillMaterial TDBGrid.Columns("Material ID").Value, varCurrentM_recBookmark
            
        ElseIf (Len(MaterialID) = 0 And Len(UnitCostID.Text) = 12) And Right(UnitCostID.Text, 1) <> "*" Then
            'UPDATED 8/25/2004 RTD TO SUPPORT MASTERFORMAT 2004
            TDBGrid.Columns("unit_cost_skey").Value = GetUCSkey(UnitCostID.Text, MasterFormat)
            If MasterFormat = EXT_MASTERFORMAT_VERSION Then
                TDBGrid.Columns("Unit Cost ID " & Right(EXT_MASTERFORMAT_VERSION, 2)).Value = UnitCostID.Text
            Else
                TDBGrid.Columns("Unit Cost ID").Value = UnitCostID.Text
            End If
            TDBGrid.Col = TDBGrid.Columns("Material ID").ColIndex
            m_objGridMap.FillUnitCost UnitCostID.Text, MasterFormat
        End If
        
        TDBGrid.EditActive = True
        TDBGrid.Columns("Input Factor").Value = 1
        TDBGrid.Columns("Output Factor").Value = 1
        TDBGrid.Columns("Adj Factor").Value = 1
        TDBGrid.Columns("Unit Qty").Value = 1
        TDBGrid.Columns("last_update_id") = 0
        TDBGrid.Columns("last_update_person") = strUserName
    End If
End Sub

Private Function ValidGridRow() As Boolean
    
    ' MODIFIED 8/25/2005 RTD - UNIT COST ID AND EXT UNIT COST ID CAN BE EMPTY, BUT NOT BOTH
    If Len(Trim(TDBGrid.Columns("unit cost id"))) = 0 And Len(Trim(TDBGrid.Columns("unit cost id " & Right(EXT_MASTERFORMAT_VERSION, 2)))) = 0 _
        Or Len(Trim(TDBGrid.Columns("material id"))) = 0 Then
        MsgBox "Both the Material and Unit Cost ID(s) must be entered.", vbExclamation
        TDBGrid.SetFocus
        ValidGridRow = False
    Else
        ValidGridRow = True
    End If

End Function

Private Sub cmdUnitCost_Click()
    Dim sUnitCostId As String
    
    If IsNumeric(TDBGrid.Bookmark) Then
        ' Navigate to grid view
        Dim frm As frmUnitCostGrid
        Set frm = New frmUnitCostGrid
        
        If Me.cboMasterFormat.Text = "MF-2004" Then
            MasterFormat = "2004"
        End If
        If Me.cboMasterFormat.Text = "MF-1995" Then
            MasterFormat = "1995"
        End If
        
        'rlh CCD 8.4 Fix to KSR issue
        If m_intMasterFormat = EXT_MASTERFORMAT_VERSION Then
            sUnitCostId = TDBGrid.Columns("Unit Cost ID 04").CellText(TDBGrid.Bookmark)
        End If
        
        If m_intMasterFormat = UCD_MASTERFORMAT_VERSION Then
            sUnitCostId = TDBGrid.Columns("Unit Cost ID").CellText(TDBGrid.Bookmark)
        End If
        
        If sUnitCostId = "" Then
            If m_intMasterFormat = UCD_MASTERFORMAT_VERSION Then  'rlh ccd 8.4
                'sUnitCostId = TDBGrid.Columns("Unit Cost ID " & Right(EXT_MASTERFORMAT_VERSION, 2)).CellText(TDBGrid.Bookmark)
                sUnitCostId = TDBGrid.Columns("Unit Cost ID").CellText(TDBGrid.Bookmark)
                frm.MasterFormat = MasterFormat
                frm.JumpIn Compress_String(sUnitCostId) + "*"
            End If
            If m_intMasterFormat = EXT_MASTERFORMAT_VERSION Then  'rlh ccd 8.4
                 sUnitCostId = TDBGrid.Columns("Unit Cost ID 04").CellText(TDBGrid.Bookmark)
                frm.MasterFormat = MasterFormat
                frm.JumpIn Compress_String(sUnitCostId) + "*"
            End If
           
        Else
            frm.JumpIn Compress_String(sUnitCostId) + "*"
        End If
    Else
        MsgBox "Please select a row first.", vbInformation
    End If
    
End Sub

Private Sub cmdUpdate_Click()
    Dim blnRet As Boolean
    Dim vntBookmark As Variant
    
    On Error GoTo Error_Processing
    Screen.MousePointer = vbHourglass
    m_blnWereErrors = False
    vntBookmark = TDBGrid.Bookmark
    If ValidGridRow() = True Then
        TDBGrid.Update
        blnRet = m_objGridMap.Update
        If blnRet = False Then
            m_blnWereErrors = True
        End If
        TDBGrid.Bookmark = vntBookmark
    End If
Exit_Sub:
    Screen.MousePointer = vbNormal
Exit Sub

Error_Processing:
    'MsgBox Error$
    Resume Exit_Sub

End Sub

Private Sub Form_Deactivate()
    m_strCurrentFormControl = Me.ActiveControl.Name
    ShowToolbarIcons False
End Sub

Private Sub Form_Initialize()
    
    On Error Resume Next
    ' Fill the MasterFormat tree
    If Forms(0).ActiveForm.Name = "frmUnitCostGrid" Then
        strSource = "Unit Cost"
    ElseIf Forms(0).ActiveForm.Name = "frmMatPriceGrid" Then
        strSource = "Material"
    End If
    'Line of code was changed by Mohan on Jan 05,2012, added "MATERIAL04" to make sure it uses MASTERFORMAT04
    FormatTree.InitData g_cnShared, "MATERIAL04"
        ' Initialize grid
    m_objGridMap.SetGrid TDBGrid
    m_objGridMap.strSource = strSource
    m_objGridMap.InitGrid
    m_blnFirstSearch = True
    m_blnJumpIn = False

End Sub

Private Sub Form_Load()

    Dim blnReturn As Boolean
    Dim strSelect As String
    
    Move START_LEFT, START_TOP, START_WIDTH, START_HEIGHT
    
    LoadMasterFormatCombo Me.cboMasterFormat, True
    
    ' This will never return any rows, just used to create recordset
'    strSelect = "select mu.mat_skey, m.mat_id, m.tech_desc as material_tech_desc, m.metric_tech_desc as material_metric_tech_desc, " + _
'        "ucd.unit_cost_skey, ucd.unit_cost_id, ucd.tech_desc as unit_tech_desc, ucd.metric_tech_desc as unit_metric_tech_desc, " + _
'        "m.usage_unit, mu.unit_qty, mu.input_factor, mu.output_factor, mu.adj_factor, ucd.unit, " + _
'        "mu.comment, mu.last_update_date , mu.last_update_person " + _
'        "From material as m, material_usage as mu, unit_cost_detail as ucd " + _
'        "where m.mat_skey = mu.mat_skey and mu.unit_cost_skey = ucd.unit_cost_skey and "
'        strSelect = strSelect + "m.mat_id = '0'"
    
    MaterialID.Text = "~"
    cmdSearch_Click
    MaterialID.Text = ""
    
'    blnReturn = g_objDAL.GetRecordset(vbNullString, strSelect, m_rec)
'    m_objGridMap.RecordSet = m_rec

End Sub

' Called when coming here from another screen
Public Sub JumpIn(strMatID As String)
    MaterialID.Text = strMatID
    m_objGridMap.Material_id = strMatID
    cmdSearch_Click
End Sub

' Called when coming here from another screen
Public Sub JumpIn2(strUCostID As String)
    UnitCostID.Text = strUCostID
    m_objGridMap.UnitCost_ID = strUCostID
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
    ShowToolbarIcons False
End Sub

' Leaf in MasterFormat tree selected.
Private Sub FormatTree_NodeSelected(ByVal strID As String)
    ' Synch text box with tree
    MaterialID.Text = strID + "*"
    ' Clear other boxes
    UnitCostID.Text = ""
    ' Kick-off search
    cmdSearch_Click
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
    
    If Len(MaterialID.Text) = 0 And Len(UnitCostID.Text) = 0 Then
        MsgBox "You must enter either Material ID or Unit Cost ID or both.", vbInformation
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    lblRowCount.Caption = "Working..."
    lblRowCount.Refresh
    
    ' Synch tree with text box
    If Not MaterialID.Text = "" Then
        FormatTree.FocusItem (MaterialID.Text)
    End If
    
    m_rec.Close ' Make sure it is closed
    m_rec.MaxRecords = MAX_RECORDS ' Set the maximum number to bring back
    dtmStart = Now
    
    strSelect = "exec usp_select_material_usage_ext @mat_id='"
    If Len(MaterialID.Text) > 0 Then
        If Left(MaterialID.Text, 1) <> "M" Then
            strSelect = strSelect + "M"
        End If
        strSelect = strSelect + SQLChangeWildcard(Compress_String(MaterialID.Text)) + "',  @unit_cost_ID='"
    Else
        strSelect = strSelect + "', @unit_cost_ID='"
    End If
    
    If Len(UnitCostID.Text) > 0 Then
        strSelect = strSelect + SQLChangeWildcard(Compress_String(UnitCostID.Text)) + "', "
    Else
        strSelect = strSelect + "', "
    End If
    strSelect = strSelect + "@op_code='STD', @country_code='USA', @region_code='NAT', @selmode="
    If Len(UnitCostID.Text) > 0 And Len(MaterialID.Text) > 0 Then
        strSelect = strSelect + "2"
    ElseIf Len(UnitCostID.Text) > 0 Then
        strSelect = strSelect + "3"
    Else
        strSelect = strSelect + "1"
    End If
    strSelect = strSelect + ", @master_format=" & MasterFormat
    
    ' Use DAL to perform select
    blnReturn = g_objDAL.GetRecordset(vbNullString, strSelect, m_rec)
    
    If blnReturn = False Then
        Screen.MousePointer = vbNormal
        MsgBox "An error occurred while searching.", vbCritical
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

Private Sub MaterialID_Change()
    Dim intStart As Integer
    intStart = MaterialID.SelStart
    MaterialID.Text = UCase(MaterialID.Text)
    MaterialID.SelStart = intStart
End Sub

Private Sub MaterialID_LostFocus()
    MaterialID = Trim(MaterialID)
    m_objGridMap.Material_id = MaterialID.Text
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
        If Button = vbRightButton And IsNumeric(TDBGrid.Bookmark) Then
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
        ShowToolbarIcons True
    End If
End Sub

Private Sub UnitCostID_LostFocus()
    m_objGridMap.UnitCost_ID = UnitCostID.Text
End Sub

Public Function SelectMasterFormat(iMasterFormat As Long) As Boolean
'SET THE MASTERFORMAT COMBO BOX TO THE NEW SELECTION
'ADDED 8/4/2005 RTD
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

Private Sub MasterFormatChanged()
'A NEW MASTERFORMAT WAS SELECTED FROM THE DROP-DOWN BOX
'ADDED 6/20/2005 RTD FOR VERSION 7.4.0+
    Dim sTreeType As String
    
    If cboMasterFormat.ListIndex < 0 Then
        Exit Sub
    End If
    
    Select Case cboMasterFormat.ItemData(cboMasterFormat.ListIndex)
    Case EXT_MASTERFORMAT_VERSION
        lblUnitCostId.Caption = "Unit Cost ID " & Right(EXT_MASTERFORMAT_VERSION, 2) & ":"
        'code change by Mohan on Jan 06, 2012: Changed value from "MATERIAL" to "MATERIAL04"
        sTreeType = "MATERIAL04"
    Case UCD_MASTERFORMAT_VERSION
        lblUnitCostId.Caption = "Unit Cost ID " & Right(UCD_MASTERFORMAT_VERSION, 2) & ":"
        sTreeType = "MATERIAL"
    Case ALT_MASTERFORMAT_VERSION
        'EndUnitCostID.Text = ""
        lblUnitCostId.Caption = "Alt Unit Cost ID:"
        sTreeType = "MATERIAL"
    Case Else
        lblUnitCostId.Caption = "Unit Cost ID:"
        sTreeType = "MATERIAL"
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
    UnitCostID.SetFocus
    Screen.MousePointer = vbDefault

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
    TDBGrid.PrintInfo.PageHeader = "\t" & Me.Caption
    
    TDBGrid.PrintInfo.PreviewInitHeight = START_HEIGHT / Screen.TwipsPerPixelX
    TDBGrid.PrintInfo.PreviewInitWidth = START_WIDTH / Screen.TwipsPerPixelY
    TDBGrid.PrintInfo.PreviewInitPosX = 5 + (fMainForm.Left / Screen.TwipsPerPixelX)
    TDBGrid.PrintInfo.PreviewInitPosY = 4 + ((fMainForm.Top + fMainForm.sbStatusBar.Height + fMainForm.tbToolBar.Height * 2) / Screen.TwipsPerPixelY)
    TDBGrid.PrintInfo.PageHeaderFont.Bold = True
    TDBGrid.PrintInfo.PageHeaderFont.Size = 12
    TDBGrid.PrintInfo.PageFooter = CStr(Now) & "\t\tPage \p"
    ' ORIENTATION 1=PORTRAIT | 2=LANDSCAPE
    TDBGrid.PrintInfo.SettingsOrientation = 2
    TDBGrid.PrintInfo.SettingsMarginBottom = 720
    TDBGrid.PrintInfo.SettingsMarginTop = 720
    TDBGrid.PrintInfo.SettingsMarginLeft = 720
    TDBGrid.PrintInfo.SettingsMarginRight = 720
    
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

