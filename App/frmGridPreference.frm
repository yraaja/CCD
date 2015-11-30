VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmGridPreference 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grid Preferences"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10845
   Icon            =   "frmGridPreference.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2460
   ScaleWidth      =   10845
   Begin VB.CheckBox chkAlternatingBackground 
      Caption         =   "&White Grid Background"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   2175
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset Defaults"
      Height          =   495
      Left            =   6000
      TabIndex        =   3
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   495
      Left            =   4080
      TabIndex        =   1
      Top             =   1800
      Width           =   1455
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   540
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   1720
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
      Splits(0)._ColumnProps(10)=   "Column(1)._MinWidth=162211412"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
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
   Begin VB.Label lblScreen 
      AutoSize        =   -1  'True
      Caption         =   "lblScreen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   2
      Top             =   60
      Width           =   1350
   End
End
Attribute VB_Name = "frmGridPreference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_strGridType As String
Dim m_strScreen As String
Private GridArray() As Variant
Private MaxCol As Integer
Private MaxRow As Long
Private m_blnChanges As Boolean ' Track if user made changes

Private Sub cmdReset_Click()
    Dim strKey As String
    Dim hKey As Long
    Dim lRet As Long
    On Error Resume Next
    TDBGrid.Split = TDBGrid.Splits.Count - 1
    ' delete values in modified columns
    For intCol = 0 To TDBGrid.Columns.Count - 1
        strKey = CCD_KEY + "\" + m_strGridType + "\" + TDBGrid.Columns(intCol).DataField ' .Caption
        lRet = RegDeleteKey(HKEY_CURRENT_USER, strKey)
    Next intCol
    RestorePreferences
End Sub

Private Sub cmdSave_Click()
    Dim strKey As String
    Dim hKey As Long
    Dim lRet As Long
    
    TDBGrid.Split = TDBGrid.Splits.Count - 1
    
    ' Save values in modified columns
    For intCol = 0 To TDBGrid.Columns.Count - 1
        strKey = CCD_KEY + "\" + m_strGridType + "\" + TDBGrid.Columns(intCol).DataField ' .Caption
        lRet = RegOpenKeyEx(HKEY_CURRENT_USER, strKey, 0&, KEY_ALL_ACCESS, hKey)
        If lRet <> 0 Then
        Else
            lRet = RegSetValueExLong(hKey, "Order", 0&, REG_DWORD, TDBGrid.Columns(intCol).Order, 4)
            lRet = RegSetValueExLong(hKey, "Visible", 0&, REG_DWORD, TDBGrid.Columns(intCol).Value, 4)
            lRet = RegSetValueExLong(hKey, "Width", 0&, REG_DWORD, TDBGrid.Columns(intCol).Width, 4)
        End If
    Next intCol
    
    strKey = CCD_KEY + "\" + m_strGridType
    lRet = RegOpenKeyEx(HKEY_CURRENT_USER, strKey, 0&, KEY_ALL_ACCESS, hKey)
    If lRet <> 0 Then
    Else
        lRet = RegSetValueExLong(hKey, "Background", 0&, REG_DWORD, chkAlternatingBackground.Value, 4)
    End If
    m_blnChanges = False
    MsgBox "Preferences saved.  Re-open the form to apply new formatting."
    Unload Me
End Sub

Private Sub Form_Activate()
    Dim i As Integer
    For i = 1 To TDBGrid.Splits.Count - 1
        TDBGrid.Splits(i).AllowColMove = True ' Allow columns in right split to move
    Next i
    TDBGrid.ReBind
    TDBGrid.ReBind
    TDBGrid.ApproxCount = 1
    OutputView False

End Sub

Public Sub SetType(strType As String)
    m_strScreen = strType
End Sub

Private Sub Form_Load()
    Dim m_objGridMap
    Move START_LEFT, START_TOP

    Select Case m_strScreen
    Case "Material"
        Set m_objGridMap = New CMaterialMap
    Case "Material Price"
        Set m_objGridMap = New CMatPriceMap
    Case "Material History"
        Set m_objGridMap = New CMatHistoryMap
    Case "Material Usage"
        Set m_objGridMap = New CMatUsageMap
    Case "Material Manufacturer"
        Set m_objGridMap = New CMatManufacMap
    Case "Information Source"
        Set m_objGridMap = New CInfoSourceMap
    Case "Equipment Rate"
        Set m_objGridMap = New CEquipRateMap
    Case "Equipment History"
        Set m_objGridMap = New CEquipHistMap
    Case "Equipment"
        Set m_objGridMap = New CEquipmentMap
    Case "Crews"
        Set m_objGridMap = New CCrewMap
    Case "Unit Cost"
        Set m_objGridMap = New CUnitCostMap
        '8/3/2005 RTD - GRID COLUMNS SHOULD DISPLAY ACCORDING TO DEFAULT MASTERFORMAT
        m_objGridMap.MasterFormat = g_intMasterFormat
    Case "Unit Cost Usage"
        Set m_objGridMap = New CUCostUsageMap
    Case "Unit Cost History"
        Set m_objGridMap = New CUCostHistMap
    Case "Labor Rate"
        Set m_objGridMap = New CLaborRateMap
    Case "Trade Group"
        Set m_objGridMap = New CTradeGroupMap
    Case "Assembly Book Detail"
        Set m_objGridMap = New CAssemblyBkMap
    Case "Assembly Maintenance"
        Set m_objGridMap = New CAssemblyMap
    Case "Assembly History"
        Set m_objGridMap = New CAsblyHistMap
    Case "Assembly Unit Cost Usage"
        Set m_objGridMap = New CAsUCUsageMap
    Case "Building"
        Set m_objGridMap = New CBuildingMap
    Case "Model"
        Set m_objGridMap = New CModelMap
    Case "Common Additives"
        Set m_objGridMap = New CBldgComAdds
    Case "Model Assemblies"
        Set m_objGridMap = New CMdlAssembly
    Case "Summary Estimate"
        Set m_objGridMap = New CMdlComponent
    Case "CCI Equipment Rate"
        Set m_objGridMap = New CCCIEqpRtMap
    Case "CCI Index Detail"
        Set m_objGridMap = New CCCIIdxDtlMap
    Case "CCI Detail"
        Set m_objGridMap = New CCCIDetailMap
    Case "CCI Labor Rate"
        Set m_objGridMap = New CCCILabRtMap
    Case "CCI Material Price"
        Set m_objGridMap = New CCCIMatPrMap
    Case "CCI Dollar Report"
        Set m_objGridMap = New CCCICSIFmtMap
    Case "CCI Index Detail Exception Report"
        Set m_objGridMap = New CCCIIdxExcMap
    Case "CCI Labor Exception"
        Set m_objGridMap = New CCCILabExcMap
    Case "CCI Component Usage"
        Set m_objGridMap = New CCCICompUseMap
    Case "CCI Material/Equipment Exception"
        Set m_objGridMap = New CCCIMatEqMap
    Case "Project Analysis"
        Set m_objGridMap = New CAnalysisMap
    Case "Project Grid"
        Set m_objGridMap = New CProjectMap
    Case "Output Usage"
        '8/8/2005 RTD - ADDED OUTPUT GRID
        Set m_objGridMap = New COutputMap
    Case Else
        lblScreen.Caption = ""
        cmdSave.Enabled = False
        cmdReset.Enabled = False
        MsgBox "Screen Type Not Found", vbCritical
        Exit Sub
    End Select
    
    lblScreen.Caption = m_strScreen
    
    m_strGridType = m_objGridMap.GRIDTYPE
    m_objGridMap.Preferences = True
    ' Initialize grid only once
    m_objGridMap.SetGrid TDBGrid
    
    If m_strGridType = "PROJECT" Then
        '
        '   List of Project components
        Dim aryClassList() As String
        Dim aryExteriorMaterial() As String
        Dim aryState() As String
        Dim m_rec As ADODB.RecordSet

        Dim i As Integer
        i = 0
        If Not g_objDAL.GetRecordset(vbNullString, "SELECT distinct class_id, sort_order FROM CLASSIFICATION WHERE class_system_id = 'P1' and class_id not like 'T%' ORDER BY sort_order", m_rec) Then
            MsgBox "An error occurred while searching for classification code(s)."
        Else
            Do Until m_rec.EOF
                ReDim Preserve aryClassList(i)
                aryClassList(i) = m_rec.Fields("class_id")
                i = i + 1
                m_rec.MoveNext
            Loop
        End If
        m_rec.Close
        
        i = 0
        If Not g_objDAL.GetRecordset(vbNullString, "SELECT distinct exterior_material_desc FROM EXTERIOR_MATERIAL ORDER BY exterior_material_desc", m_rec) Then
            MsgBox "An error occurred while searching for exterior material code(s)."
        Else
            Do Until m_rec.EOF
                ReDim Preserve aryExteriorMaterial(i)
                aryExteriorMaterial(i) = m_rec.Fields("exterior_material_desc")
                i = i + 1
                m_rec.MoveNext
            Loop
        End If
        m_rec.Close
        
        i = 0
        If Not g_objDAL.GetRecordset(vbNullString, "SELECT distinct state_code FROM state_country ORDER BY state_code", m_rec) Then
            MsgBox "An error occurred while searching for exterior material code(s)."
        Else
            Do Until m_rec.EOF
                ReDim Preserve aryState(i)
                aryState(i) = m_rec.Fields("state_code")
                i = i + 1
                m_rec.MoveNext
            Loop
        End If
        m_rec.Close
        m_objGridMap.InitGrid aryClassList, aryExteriorMaterial, aryState, True
    ElseIf m_strGridType = "Output" Then
        m_objGridMap.InitPreferenceGrid
    Else
        m_objGridMap.InitGrid
    End If
    
    Dim strKey As String
    Dim lRet As Long
    Dim hKey As Long
    Dim intCols As Integer
    intCols = TDBGrid.Columns.Count
    
    strKey = CCD_KEY + "\" + m_strGridType
    lRet = RegOpenKeyEx(HKEY_CURRENT_USER, strKey, 0&, KEY_ALL_ACCESS, hKey)
    If lRet <> 0 Then
    Else
        lRet = RegQueryValueExLong(hKey, "Background", 0&, REG_DWORD, lValue, 4)
        chkAlternatingBackground.Value = lValue
    End If

    SetDims 1, intCols
    
    For i = 0 To intCols - 1
        'code change by Mohan on Jan 09,2012: changing string to constant "RSMeans\CCD" to CCD_KEY
        strKey = CCD_KEY + "\" + m_strGridType + "\" + TDBGrid.Columns(i).DataField ' .Caption
        lRet = RegOpenKeyEx(HKEY_CURRENT_USER, strKey, 0&, KEY_ALL_ACCESS, hKey)
        If lRet <> 0 Then
            GridArray(i, 0) = TDBGrid.Columns(i).Visible
        Else
            lSize = 4
            lRet = RegQueryValueExLong(hKey, "Visible", 0&, REG_DWORD, lValue, lSize)
'            GridArray(i, 0) = 0
            GridArray(i, 0) = lValue
        End If
    Next
    
    m_blnChanges = False
    
    ' Make the grid row fit just right
    Dim intScrollWidth As Integer
    intScrollWidth = GetSystemMetrics(3) ' SM_CYHSCROLL
    If TDBGrid.Splits(0).Caption <> "" Then
        TDBGrid.Height = 750 + intScrollWidth * Screen.TwipsPerPixelY
    Else
        TDBGrid.Height = 525 + intScrollWidth * Screen.TwipsPerPixelY
    End If

End Sub

Private Sub RestorePreferences()
    Dim m_objGridMap
    On Error Resume Next
'    Move START_LEFT, START_TOP
'    Me.Visible = False
    Select Case m_strScreen
    Case "Material"
        Set m_objGridMap = New CMaterialMap
    Case "Material Price"
        Set m_objGridMap = New CMatPriceMap
    Case "Material History"
        Set m_objGridMap = New CMatHistoryMap
    Case "Material Usage"
        Set m_objGridMap = New CMatUsageMap
    Case "Material Manufacturer"
        Set m_objGridMap = New CMatManufacMap
    Case "Information Source"
        Set m_objGridMap = New CInfoSourceMap
    Case "Equipment Rate"
        Set m_objGridMap = New CEquipRateMap
    Case "Equipment History"
        Set m_objGridMap = New CEquipHistMap
    Case "Equipment"
        Set m_objGridMap = New CEquipmentMap
    Case "Crews"
        Set m_objGridMap = New CCrewMap
    Case "Unit Cost"
        Set m_objGridMap = New CUnitCostMap
    Case "Unit Cost Usage"
        Set m_objGridMap = New CUCostUsageMap
    Case "Unit Cost History"
        Set m_objGridMap = New CUCostHistMap
    Case "Labor Rate"
        Set m_objGridMap = New CLaborRateMap
    Case "Assembly Maintenance"
        Set m_objGridMap = New CAssemblyMap
    Case "Assembly Book Detail"
        Set m_objGridMap = New CAssemblyBkMap
    Case "Assembly Unit Cost Usage"
        Set m_objGridMap = New CAsUCUsageMap
    Case "Assembly History"
        Set m_objGridMap = New CAsblyHistMap
    Case "Building"
        Set m_objGridMap = New CBuildingMap
    Case "Model"
        Set m_objGridMap = New CModelMap
    Case "Common Additives"
        Set m_objGridMap = New CBldgComAdds
    Case "Model Assemblies"
        Set m_objGridMap = New CMdlAssembly
    Case "Summary Estimate"
        Set m_objGridMap = New CMdlComponent
    Case "CCI Equipment Rate"
        Set m_objGridMap = New CCCIEqpRtMap
    Case "CCI Index Detail"
        Set m_objGridMap = New CCCIIdxDtlMap
    Case "CCI Detail"
        Set m_objGridMap = New CCCIDetailMap
    Case "CCI Labor Rate"
        Set m_objGridMap = New CCCILabRtMap
    Case "CCI Material Price"
        Set m_objGridMap = New CCCIMatPrMap
    Case "CCI Dollar Report"
        Set m_objGridMap = New CCCICSIFmtMap
    Case "CCI Index Detail Exception Report"
        Set m_objGridMap = New CCCIIdxExcMap
    Case "CCI Component Usage"
        Set m_objGridMap = New CCCICompUseMap
    Case "CCI Material/Equipment Exception"
        Set m_objGridMap = New CCCIMatEqMap
    Case "CCI Labor Exception"
        Set m_objGridMap = New CCCILabExcMap
    Case "Project Analysis"
        Set m_objGridMap = New CAnalysisMap
    Case "Project Grid"
        Set m_objGridMap = New CProjectMap
    Case "Output Usage"
        Set m_objGridMap = New COutputMap
    Case Else
        MsgBox "Screen Type Not Found"
        Exit Sub
    End Select
    
    lblScreen.Caption = m_strScreen
'    m_strGridType = m_objGridMap.GRIDTYPE
    m_objGridMap.Preferences = True
    ' Initialize grid only once
    m_objGridMap.SetGrid TDBGrid
    m_objGridMap.InitGrid
    'remove all splits
'    For i = 0 To TDBGrid.Splits.Count - 1
'        TDBGrid.Splits.Remove (0)    'Remove/rebuild last split
'    Next
    m_blnChanges = False
    MsgBox "Preferences Restored to Default"
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim intCounter As Integer
    If m_blnChanges = True Then
        Dim Button
        Button = MsgBox("Do you want to save your changes?", vbYesNoCancel)
        If Button = vbYes Then
            cmdSave_Click
            Exit Sub
        ElseIf Button = vbCancel Then
            Cancel = True
            Exit Sub
        Else
            Exit Sub
        End If
    End If
End Sub

Private Sub TDBGrid_AfterColUpdate(ByVal ColIndex As Integer)
    GridArray(ColIndex, TDBGrid.Bookmark) = TDBGrid.Columns(ColIndex).Value
    m_blnChanges = True
End Sub

'*** APEX Migration Utility Code Change ***
'Private Sub TDBGrid_UnboundReadDataEx(ByVal RowBuf As TrueOleDBGrid60.RowBuffer, StartLocation As Variant, ByVal offset As Long, ApproximatePosition As Long)
'*** APEX Migration Utility Code Change ***
'Private Sub TDBGrid_UnboundReadDataEx(ByVal RowBuf As TrueOleDBGrid70.RowBuffer, StartLocation As Variant, ByVal offset As Long, ApproximatePosition As Long)
Private Sub TDBGrid_UnboundReadDataEx(ByVal RowBuf As TrueOleDBGrid80.RowBuffer, StartLocation As Variant, ByVal offset As Long, ApproximatePosition As Long)
    Dim ColIndex As Integer, j As Integer
    Dim RowsFetched As Integer, i As Long
    Dim NewPosition As Long, Bookmark As Long
    Dim StartRow As Long
    
    Dim Cols As Long, Rows As Long
    Cols = RowBuf.COLUMNCOUNT - 1
    Rows = RowBuf.RowCount - 1
    
    RowsFetched = 0
    
    If IsNull(StartLocation) Then
        ' StartLocation reffers to either BOF (-1) or EOF (MaxRow)
        StartRow = IIf(offset < 0, MaxRow + offset, -1 + offset)
    Else
        ' StartLocation is an actual bookmark
        StartRow = StartLocation + offset
    End If
    
    For i = 0 To Rows
        Bookmark = StartRow + i
        ' If we are out of bounds quit this loop
        If Bookmark < 0 Or Bookmark >= MaxRow Then Exit For
               
        ' Fill the RowBuffer with data
        For j = 0 To Cols
            ColIndex = RowBuf.ColumnIndex(i, j)
            RowBuf.Value(i, j) = GridArray(ColIndex, Bookmark)
        Next j
        
        ' Assign a bookmark for this row
        RowBuf.Bookmark(i) = Bookmark
        ' Increment number of rows fetched
        RowsFetched = RowsFetched + 1
    Next i
    
    RowBuf.RowCount = RowsFetched
    
    ' Callibrate the VScroll bar
    If StartRow >= 0 Then ApproximatePosition = StartRow
End Sub

' Initialises array
Public Sub SetDims(ByVal Rows As Long, ByVal Cols As Integer)
    If Rows <= 0 And Cols > 0 Then
        ReDim GridArray(0 To Cols - 1, 0)
    ElseIf Rows <= 0 And Cols <= 0 Then
        ReDim GridArray(0, 0)
    ElseIf Rows > 0 And Cols <= 0 Then
        ReDim GridArray(0, 0 To Rows - 1)
    Else
        ReDim GridArray(0 To Cols - 1, 0 To Rows - 1)
    End If
    
    MaxRow = Rows
    MaxCol = Cols
End Sub

