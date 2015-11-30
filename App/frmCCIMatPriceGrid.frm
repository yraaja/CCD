VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{5936A75C-3F42-11D6-AF6B-AA0004005F12}#1.3#0"; "MeansCtrl.ocx"
Begin VB.Form frmCCIMatPriceGrid 
   Caption         =   "CCI Material Price Maintenance Grid"
   ClientHeight    =   7065
   ClientLeft      =   2265
   ClientTop       =   2835
   ClientWidth     =   11355
   Icon            =   "frmCCIMatPriceGrid.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7065
   ScaleWidth      =   11355
   Begin VB.Frame fraSelType 
      Caption         =   "Geographic Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   5160
      TabIndex        =   13
      Top             =   360
      Width           =   6855
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   540
         Left            =   120
         ScaleHeight     =   540
         ScaleWidth      =   6495
         TabIndex        =   25
         Top             =   240
         Width           =   6495
         Begin VB.OptionButton optNatlAvg 
            Caption         =   "Nat'l Avg (30-City)"
            Height          =   255
            Left            =   0
            TabIndex        =   30
            Top             =   0
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton optCCICities 
            Caption         =   "CCI Cities (727-Cities)"
            Height          =   255
            Left            =   0
            TabIndex        =   29
            Top             =   300
            Width           =   1875
         End
         Begin VB.OptionButton optPriCity 
            Caption         =   "Primary Cities (316-Cities)"
            Height          =   255
            Left            =   1920
            TabIndex        =   28
            Top             =   0
            Width           =   2055
         End
         Begin VB.OptionButton optAllCities 
            Caption         =   "All CCI Cities (731-Cities)"
            Height          =   195
            Left            =   1920
            TabIndex        =   27
            Top             =   300
            Width           =   2070
         End
         Begin VB.OptionButton opt66Cities 
            Caption         =   "Sixty Six CCI  Cities (66-Cities)"
            Height          =   255
            Left            =   4080
            TabIndex        =   26
            Top             =   0
            Width           =   2655
         End
      End
   End
   Begin VB.Frame fraEdit 
      Caption         =   "Multiple Edits"
      Height          =   735
      Left            =   120
      TabIndex        =   21
      Top             =   6240
      Width           =   4575
      Begin VB.CommandButton cmdUpdateMaterialPrice 
         Caption         =   "Material Price"
         Height          =   375
         Left            =   2880
         TabIndex        =   24
         ToolTipText     =   "Update Material Price"
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdCheckNone 
         Caption         =   "Est Ind Off"
         Height          =   375
         Left            =   1440
         TabIndex        =   23
         ToolTipText     =   "Uncheck all Estimated Ind boxes"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdCheckAll 
         Caption         =   "Est Ind On"
         Height          =   375
         Left            =   240
         TabIndex        =   22
         ToolTipText     =   "Check all Estimated Ind boxes"
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.ComboBox cmbCity 
      Enabled         =   0   'False
      Height          =   315
      Left            =   8730
      TabIndex        =   16
      Top             =   1320
      Width           =   1575
   End
   Begin VB.ComboBox cmbState 
      Height          =   315
      Left            =   7560
      TabIndex        =   15
      Top             =   1320
      Width           =   765
   End
   Begin VB.TextBox Zip 
      Height          =   285
      Left            =   10650
      TabIndex        =   14
      Top             =   1320
      Width           =   495
   End
   Begin VB.ComboBox cmbCountry 
      Height          =   315
      Left            =   5760
      TabIndex        =   12
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdClone 
      Caption         =   "Clone"
      Height          =   495
      Left            =   9840
      TabIndex        =   11
      Top             =   6360
      Width           =   1155
   End
   Begin VB.ComboBox cmbQuarterID 
      Height          =   315
      Left            =   5760
      TabIndex        =   10
      Top             =   1710
      Width           =   1125
   End
   Begin ConstructionCostDatabase.DynaTree FormatTree 
      Height          =   2070
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   4725
      _ExtentX        =   8334
      _ExtentY        =   3651
   End
   Begin VB.TextBox MatID 
      Height          =   315
      Left            =   7575
      TabIndex        =   0
      Top             =   1710
      Width           =   1515
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   435
      Left            =   9960
      TabIndex        =   1
      Top             =   1680
      Width           =   1150
   End
   Begin VB.CheckBox ckbRowWrap 
      Caption         =   "Row Wrap"
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   2250
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   495
      Left            =   8520
      TabIndex        =   3
      Top             =   6360
      Width           =   1155
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid 
      Height          =   3540
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2610
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   6244
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
      Splits(0)._ColumnProps(5)=   "Column(0)._MinWidth=15"
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
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "State:"
      Height          =   255
      Left            =   6960
      TabIndex        =   20
      Top             =   1380
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "City:"
      Height          =   255
      Left            =   8280
      TabIndex        =   19
      Top             =   1380
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Zip:"
      Height          =   255
      Left            =   10260
      TabIndex        =   18
      Top             =   1380
      Width           =   375
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Country:"
      Height          =   255
      Left            =   4965
      TabIndex        =   17
      Top             =   1380
      Width           =   735
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Quarter:"
      Height          =   255
      Left            =   4920
      TabIndex        =   9
      Top             =   1770
      Width           =   780
   End
   Begin VB.Label lblRowCount 
      Caption         =   "0 rows returned"
      Height          =   255
      Left            =   5400
      TabIndex        =   7
      Top             =   2250
      Width           =   3255
   End
   Begin VB.Line Line1 
      X1              =   4845
      X2              =   4845
      Y1              =   2085
      Y2              =   15
   End
   Begin VB.Line Line2 
      X1              =   180
      X2              =   11040
      Y1              =   2190
      Y2              =   2190
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Mat ID:"
      Height          =   255
      Left            =   6855
      TabIndex        =   6
      Top             =   1770
      Width           =   615
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
      Left            =   4965
      TabIndex        =   5
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmCCIMatPriceGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''' <modulename> frmCCIMatPriceGrid</modulename>
''' <functionname>General (Main) </functionname>
'''
''' <summary>
''' (CCI)MATERIAL PRICE display grid:
'''
'''Display Material Prices based upon a defined "Geographic Selection".  The following indexes are selectable by:
'''"   NATL AVG (30 city)
'''"   PRIMARY CITIES (316 cities)
'''"   CCI CITIES (727 cities)
'''"   ALL CITIES (731 cities)
'''
'''(By)
'''
'''1.  Quarter Id                  (YYYYQn)
'''2.  Country
'''3.  City
'''4.  State
'''5.  Zip
'''6.  Mat Id
'''
'''(BUTTONS)
'''"   SEARCH                  (CmdSearch_Click() )
'''Search for material price data based upon selected "Geographic Selection" and filled in boxes
'''"   Est Ind On                  (SetEstimateIndicators(True/False) )
'''Turn on all "Est Ind" checkboxes in datagrid display
'''"   Est Ind Off                 (SetEstimateIndicators(True/False) )
'''Turn off all "Est Ind" checkboxes in datagrid display
'''"   Material Price
'''Prompt for the new/changed material price to be applied to the "checked" material rows
'''"   Update                  (CCIMatPrMap.Update() )
'''"   Clone                       (CloneCCIMatPrice() )
'''
'''NOTE: "Anytown" = 30 city average  and is displayed by selecting:
'''Country = USA
'''State = US
'''City = Anytown
'''
'''NOTE: "Rolling Quarters" on the data grid columns
'''Given the requested/selected "Quarter Id", the previous (4) quarters of data will be displayed accordingly
'''
'''Key Subs / Functions:
'''"   ExecStoredProcSelectedQuarter()
'''This gets passed a stored procedure name where ultimately the user must make a selection for a Quarter id from a dropdown
'''
'''HELPER Class: CCCIMatPrMap.Cls
''' </summary>
'''
''' <seealso>CCIMatPrMap.cls</seealso>
'''<seealso>frmMatPriceGrid</seealso>
'''
''' <datastruct>m_rec</datastruct>
'''<datastruct>m_objGridMap</datastruct>
'''
''' <storedprocedurename> sp_select_cci_material_price_rlh</storedprocedurename>
'''
''' <returns>N/A</returns>
''' <exception>Always trap with an accompanying message box</exception>
''' <example>
''' <code>
'''* * * SELECT from TREE: M10MA  (no other selections/specifications)
'''exec sp_select_cci_material_price_rlh   @cci_mat_id = 'M10MA', @qtr_id = '2006Q4', @zip = '', @loc_id = 0, @state_code = '%', @country_code = '%', @select_type = 2 </code>
''' <code>
'''* * * SPECIFY STATE (code)
'''exec sp_select_cci_material_price_rlh   @cci_mat_id = 'M10MA', @qtr_id = '2006Q4', @zip = '', @loc_id = 0, @state_code = 'CA%', @country_code = 'USA%', @select_type = 2
'''</code>
'''<code>
'''* * * ANYTOWN
'''exec sp_select_cci_material_price_rlh   @cci_mat_id = 'M10MA', @qtr_id = '2006Q4', @zip = '', @loc_id = 23, @state_code = 'US%', @country_code = 'USA%', @select_type = 2
'''</code>
'''</example>
'''<permission>Public</Permission>
'''<dependson>This component depends on the following
'''1.  CCCIMatPrMap.cls
'''2.  CGridMap.cls
'''3.  CCDdal.CRSMDataAccess (
'''4.  Access to the DAL (data access layer dll) opened in MainModule_Main() )
'''</dependson>



Dim m_objGridMap As New CCCIMatPrMap ' Class to handle grid
Dim m_blnFirstSearch As Boolean
Dim m_rec As New ADODB.RecordSet ' Recordset to hold query results
Dim m_blnDoubleClick As Boolean
Dim m_blnWereErrors As Boolean ' True if the Update had errors, used in QueryUnload
Dim m_strCurrentFormControl As String
Dim m_CurrentQtr As String
Dim m_State As String

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

Private Sub SetEstimateIndicators(blnNewValue As Boolean)
' 10/10/2005 RTD - UPDATED TO CORRECT BUG REPORTED BY G. MEDEIROS
    If m_rec.RecordCount > 0 Then
        If TDBGrid.SelBookmarks.Count > 0 Then
            ' If rows selected, update only the selected rows
            Dim vntBookmark As Variant
            For Each vntBookmark In TDBGrid.SelBookmarks
                TDBGrid.Bookmark = vntBookmark
                TDBGrid.Columns("estimated_ind") = blnNewValue
            Next
            TDBGrid.Refresh
        Else
            ' Update All Rows
            TDBGrid.MoveFirst
            Do While Not TDBGrid.EOF
                TDBGrid.Columns("estimated_ind") = blnNewValue
                TDBGrid.MoveNext
            Loop
            TDBGrid.Refresh
        End If
    End If
    
End Sub

Private Sub UpdateMaterialPrice(dblNewPrice As Double)
    
    If m_rec.RecordCount > 0 Then
        If TDBGrid.SelBookmarks.Count > 0 Then
            ' If rows selected, update only the selected rows
            Dim vntBookmark As Variant
            For Each vntBookmark In TDBGrid.SelBookmarks
                TDBGrid.Bookmark = vntBookmark
                TDBGrid.Columns("cci_mat_price") = dblNewPrice
            Next
            TDBGrid.Update
            TDBGrid.Refresh
        Else
            ' Update All Rows
            TDBGrid.MoveFirst
            Do While Not TDBGrid.EOF
                TDBGrid.Columns("cci_mat_price") = dblNewPrice
                TDBGrid.MoveNext
            Loop
            TDBGrid.Update
            TDBGrid.Refresh
        End If
    End If

End Sub

Private Sub cmbState_Change()
    Dim iSelStart  As Integer
    Dim iSelLen As Integer
    
    iSelStart = cmbState.SelStart
    iSelLen = cmbState.SelLength
    cmbState = UCase(cmbState)
    cmbState.SelStart = iSelStart
    cmbState.SelLength = iSelLen
End Sub

Private Sub cmbState_Click()
    If cmbState.ListIndex = -1 Then
        cmbCity.ListIndex = -1
        cmbCity.Enabled = False
    Else
        cmbCity.Enabled = True
        If m_State <> cmbState.Text Then
            LoadCities cmbCity, cmbState.Text
        End If
    End If

End Sub

Private Sub cmbState_GotFocus()
    m_State = cmbState.Text
End Sub

Private Sub cmbState_LostFocus()
    If cmbState.ListIndex = -1 Then
        cmbCity.ListIndex = -1
        cmbCity.Enabled = False
    Else
        cmbCity.Enabled = True
        If m_State <> cmbState.Text Then
            LoadCities cmbCity, cmbState.Text
        End If
    End If

End Sub

Private Sub cmbState_Validate(Cancel As Boolean)
    Dim i As Integer
    Dim bFound As Boolean
    If cmbState.Text <> "" Then
        For i = 0 To cmbState.listcount - 1
            If cmbState.Text = cmbState.List(i) Then
                cmbState.ListIndex = i
                bFound = True
                Exit For
            End If
        Next i
        If Not bFound Then
            MsgBox "Please enter valid state"
            Cancel = True
        End If
    End If
End Sub

Private Sub cmdCheckAll_Click()
    SetEstimateIndicators True
End Sub

Private Sub cmdCheckNone_Click()
    SetEstimateIndicators False
End Sub

Private Sub cmdClone_Click()
    CloneCCIMatPrice
End Sub

Private Sub cmdUpdate_Click()
    On Error GoTo Out
    Dim blnRet As Boolean
    Dim vntBookmark As Variant
    
    Screen.MousePointer = vbHourglass
    m_blnWereErrors = False
    vntBookmark = TDBGrid.Bookmark
    TDBGrid.Update
    blnRet = m_objGridMap.Update
    If blnRet = False Then
        m_blnWereErrors = True
    End If
    TDBGrid.Bookmark = vntBookmark

Out:
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmdUpdateMaterialPrice_Click()
    Dim sTemp As String
    Dim dNewPrice As Double
    Dim sCaption As String
    Dim sDefault As String
    Dim iCount As Long
    Dim vBookmark As Variant
    
    If m_rec.RecordCount <= 0 Then
        Exit Sub
    End If
    If TDBGrid.SelBookmarks.Count > 0 Then
        ' If rows selected, update only the selected rows
        iCount = TDBGrid.SelBookmarks.Count
        vBookmark = TDBGrid.SelBookmarks(0)
        m_rec.Bookmark = vBookmark
        sDefault = m_rec.Fields("cci_mat_price")
    Else
        ' Update All Rows
        m_rec.MoveFirst
        sDefault = m_rec.Fields("cci_mat_price")
        iCount = m_rec.RecordCount
    End If
    
    If iCount > 0 Then
        sCaption = "Enter the new Material Price to apply to " & iCount & " records:"
        sTemp = InputBox(sCaption, "New Material Price", sDefault)
        If sTemp <> "" Then
            If IsNumeric(sTemp) Then
                dNewPrice = CDbl(sTemp)
                sCaption = "Are you sure you want to change the Material Price for all " & iCount & " records to " & Format(dNewPrice, "Standard") & "?"
                If MsgBox(sCaption, vbYesNoCancel + vbQuestion) = vbYes Then
                    UpdateMaterialPrice dNewPrice
                End If
            Else
                MsgBox "'" & sTemp & "' does not appear to be a valid numeric value. Check your data and try again.", vbCritical
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
        'TDBGrid.ReBind
        OutputView False
        ShowGridSort
        m_objGridMap.SetMenuBar
        ShowToolbarIcons True
    End If
    
End Sub

Private Sub Form_Deactivate()
    m_strCurrentFormControl = Me.ActiveControl.Name
    ShowToolbarIcons False
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim strSelect As String
    Dim blnReturn As Boolean
    Move START_LEFT, START_TOP, START_WIDTH, START_HEIGHT
    LoadCombos Me, True, True, True

    ' This will never return any rows, just used to create recordset
    MatID.Text = "~"
    cmdSearch_Click
    MatID.Text = ""
    
    'Me.cmdUpdate.Enabled = False    'rlh 03/16/2010 - Update functionality is not even close!!!
    
    Status ("")
End Sub

Private Sub Form_Initialize()
    Status ("Loading CCI Material Price Maintenance...")
    Screen.MousePointer = vbHourglass
    m_blnFirstSearch = True
    FormatTree.InitData g_cnShared, "CCI_MATERIAL"
    DoEvents    'Paint screen
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
            TDBGrid.Height = Me.Height - 3885
            cmdUpdate.Top = Me.Height - 1020
            cmdClone.Top = Me.Height - 1020
            fraEdit.Top = Me.Height - 1240
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
    Dim rs As New ADODB.RecordSet
    Dim strSelect As String
    Dim blnReturn As Boolean
    
    On Error Resume Next
    If m_blnFirstSearch = True Then
        m_blnFirstSearch = False
    Else
        MatID.Text = strID
        ' Kick-off search
        cmdSearch_Click
    End If
End Sub

Private Sub cmdSearch_Click()
    On Error Resume Next
    Dim blnRet As Boolean
    Dim strSelect As String
    Dim dtmToday As Date
    Dim dtmStart As Date
    Dim strStartMatSrch As String
    Dim iSelectType As Integer
    Dim i As Integer
    Dim yyyy As String
    
    
    If Me.cmbCity = "Anytown" Then      'rlh 03/16/2010 disallow updates in Anytown Mode
        Me.cmdUpdate.Enabled = False
    Else
        Me.cmdUpdate.Enabled = True
    End If
    
    Me.cmdUpdate.Enabled = True        'rlh 03/16/2010 - Update functionality is not even close!!!
    
    
    TDBGrid.Update
    
    
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    '
    '  APPLY LABELS FOR "ROLLING" QUARTER HEADERS FOR APPLICABLE SEARCH QUARTER ID
    '  LABELS AT FORM_LOAD ARE:  QTR-1, QTR-2, QTR-3, QTR-4
    '
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
       
    
    yyyy = Mid(Me.cmbQuarterID, 1, 4)
    Dim lastQtrId As String
    lastQtrId = Me.cmbQuarterID.Text
    
    For i = 0 To TDBGrid.Columns.Count - 1
        Select Case TDBGrid.Columns(i).Caption
            Case "QTR-1"
                'TDBGrid.Columns(i).Caption = yyyy & "Q4"
                TDBGrid.Columns(i).Caption = lastQtrId
            Case "QTR-2"
                'TDBGrid.Columns(i).Caption = yyyy & "Q3"
                TDBGrid.Columns(i).Caption = Me.GetPrevQuarter(lastQtrId)
                lastQtrId = Me.GetPrevQuarter(lastQtrId)
            Case "QTR-3"
                'TDBGrid.Columns(i).Caption = yyyy & "Q2"
                TDBGrid.Columns(i).Caption = Me.GetPrevQuarter(lastQtrId)
                lastQtrId = Me.GetPrevQuarter(lastQtrId)
            Case "QTR-4"
                'TDBGrid.Columns(i).Caption = yyyy & "Q1"
                TDBGrid.Columns(i).Caption = Me.GetPrevQuarter(lastQtrId)
                lastQtrId = Me.GetPrevQuarter(lastQtrId)
            Case "QTR-5"
                TDBGrid.Columns(i).Caption = Me.GetPrevQuarter(lastQtrId)
                lastQtrId = Me.GetPrevQuarter(lastQtrId)
        End Select
    
    Next
    
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    '
    '  AFTER CONVERTING QTR-N labels to YYYYQN labels, you have to change
    '  the "compare and change" updates  (pretty lame, huh)
    '
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    lastQtrId = Me.cmbQuarterID.Text
    
    For i = 0 To TDBGrid.Columns.Count - 1
        Debug.Print "column name: " & TDBGrid.Columns(i).Caption
        If (Mid(TDBGrid.Columns(i).Caption, 5, 1) = "Q") And (Len(Trim(TDBGrid.Columns(i).Caption)) = 6) Then

            TDBGrid.Columns(i).Caption = Me.GetPrevQuarter(lastQtrId)
            lastQtrId = Me.GetPrevQuarter(lastQtrId)
        End If
    
    Next
    
    If m_objGridMap.IsPendingChange = True Then
        Dim Button
        Button = MsgBox("Do you want to save your changes?", vbYesNoCancel)
        If Button = vbYes Then
            cmdUpdate_Click
            ' If there were errors, cancel the search
            If m_blnWereErrors Then
                Exit Sub
            End If
        Else
            If Button = vbNo Then
            TDBGrid.DataChanged = False

        ElseIf Button = vbCancel Then
            ' Cancel the search
            Exit Sub
        End If
    End If
    End If
    Screen.MousePointer = vbHourglass
    dtmToday = Date
    
    ' Synch tree with text box
    If Not MatID.Text = "" Then
        FormatTree.FocusItem (MatID.Text)
    End If
    
     If Len(MatID.Text) = 0 And Len(cmbCity.Text) = 0 And Len(cmbState.Text) = 0 And Len(Zip.Text) = 0 And Len(cmbQuarterID.Text) = 0 Then
        Screen.MousePointer = vbNormal
        MsgBox "You must enter search criteria before searching."
        Exit Sub
    End If
    
    If cmbQuarterID.Text = "" Then
        Screen.MousePointer = vbNormal
        MsgBox "The Quarter is required."
        Exit Sub
    End If
    
    strSelect = "exec sp_select_cci_material_price_rlh "  'rlh 03/02/2010
    strSelect = strSelect + "  @cci_mat_id = '" + SQLChangeWildcard(MatID) + "'"
    strSelect = strSelect + ", @qtr_id = '" + cmbQuarterID.Text + "'"
    strSelect = strSelect + ", @zip = '" + SQLChangeWildcard(Zip.Text) + "'"
    If cmbCity.ListIndex = -1 Then
        strSelect = strSelect + ", @loc_id = 0"
    Else
        strSelect = strSelect + ", @loc_id = " + CStr(cmbCity.ItemData(cmbCity.ListIndex))
    End If
    strSelect = strSelect + ", @state_code = '" + FillWildCard(cmbState.Text) + "'"
    strSelect = strSelect + ", @country_code = '" + FillWildCard(cmbCountry.Text) + "'"
    strSelect = strSelect + ", @select_type = " + GeographicType(Me)
    m_rec.Close ' Make sure it is closed
    m_rec.MaxRecords = MAX_RECORDS ' Set the maximum number to bring back
    dtmStart = Now
    ' Use g_objDAL to perform select
    If DEBUGON Then Stop
    
    blnRet = g_objDAL.GetRecordset(vbNullString, strSelect, m_rec)
    If blnRet = False Then
        Screen.MousePointer = vbNormal
        MsgBox "An error occurred while searching." & vbCrLf & g_objDAL.LastErrorDescription
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

Private Sub TDBGrid_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyBack Then
        If TDBGrid.Col = 2 Or TDBGrid.Col = 3 Then
            If Len(TDBGrid.Text) + 1 > 75 Then
                KeyAscii = 0
            End If
        End If
    End If
End Sub

Private Sub TDBGrid_LostFocus()
    TDBGrid.TabStop = False
End Sub

Private Sub TDBGrid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If this is the mouse-up form a double click
    If m_blnDoubleClick Then
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

Private Sub ShowToolbarIcons(bShowIcons As Boolean)
    
    On Error GoTo Err_Handler
    With fMainForm
        .tbToolBar.Buttons.Item(tbrPRINT).Enabled = bShowIcons
        .tbToolBar.Buttons.Item(tbrPRINT).Visible = bShowIcons
        .tbToolBar.Buttons.Item(tbrPREVIEW).Enabled = bShowIcons
        .tbToolBar.Buttons.Item(tbrPREVIEW).Visible = bShowIcons
        .tbToolBar.Buttons.Item(tbrEXPORT).Enabled = False
        .tbToolBar.Buttons.Item(tbrEXPORT).Visible = False
        .tbToolBar.Buttons.Item(tbrEXPORTDATA).Enabled = bShowIcons
        .tbToolBar.Buttons.Item(tbrEXPORTDATA).Visible = bShowIcons
        .tbToolBar.Buttons.Item(tbrEXPORTDATA + 1).Visible = bShowIcons
        .mnuFilePageSetup.Enabled = bShowIcons
        .mnuFilePrint.Enabled = bShowIcons
        .mnuFileSaveAs.Enabled = False
        .mnuFilePrintPreview.Enabled = bShowIcons
    End With
    Exit Sub

Err_Handler:
    Exit Sub
    
End Sub

Public Function PreviewReport()

    TDBGrid.PrintInfo.PreviewCaption = "CCI Material Price Maintenance"
    TDBGrid.PrintInfo.PreviewInitHeight = START_HEIGHT / Screen.TwipsPerPixelX
    TDBGrid.PrintInfo.PreviewInitWidth = START_WIDTH / Screen.TwipsPerPixelY
    TDBGrid.PrintInfo.PreviewInitPosX = 5 + (fMainForm.Left / Screen.TwipsPerPixelX)
    TDBGrid.PrintInfo.PreviewInitPosY = 4 + ((fMainForm.Top + fMainForm.sbStatusBar.Height + fMainForm.tbToolBar.Height * 2) / Screen.TwipsPerPixelY)
    TDBGrid.PrintInfo.PageHeader = "\tCCI Material Price Maintenance"
    TDBGrid.PrintInfo.PageHeaderFont.Bold = True
    TDBGrid.PrintInfo.PageHeaderFont.Size = 12
    TDBGrid.PrintInfo.PageFooter = CStr(Now) & "\t\tPage \p"
    TDBGrid.PrintInfo.SettingsOrientation = 2
    TDBGrid.PrintInfo.SettingsMarginBottom = 720
    TDBGrid.PrintInfo.SettingsMarginTop = 720
    TDBGrid.PrintInfo.PrintPreview

End Function

Public Function PrintReport()

    TDBGrid.PrintInfo.PageHeader = "\tCCI Material Price Maintenance"
    TDBGrid.PrintInfo.PageHeaderFont.Bold = True
    TDBGrid.PrintInfo.PageHeaderFont.Size = 12
    TDBGrid.PrintInfo.PageFooter = CStr(Now) & "\t\tPage \p"
    TDBGrid.PrintInfo.SettingsOrientation = 2
    TDBGrid.PrintInfo.SettingsMarginBottom = 720
    TDBGrid.PrintInfo.SettingsMarginTop = 720
    TDBGrid.PrintInfo.PrintData

End Function

Public Sub ExportData()
    
    If m_rec.RecordCount > 0 Then
        Dim fExport As New frmExport
    
        fExport.SetRow TDBGrid, m_rec
        fExport.title = "CCI Material Price"
        fExport.Show
    Else
        MsgBox "Please choose or search for a CCI material price.", vbInformation + vbOKOnly
    End If
    
End Sub
Public Function GetPrevQuarter(yyyyQnn As String) As String
'RLH 03/03/2010

Dim yyyy As String
Dim nn As String
Dim prevQtr As String

yyyy = Mid(yyyyQnn, 1, 4)
nn = Mid(yyyyQnn, 6, 1)
'nn = nn - 1

If (nn = 0) Then
    prevQtr = yyyy & "Q" & CStr(4)
Else
    If (nn = 1) Then
        yyyy = CStr(CInt(yyyy) - 1)
        prevQtr = yyyy & "Q" & CStr(4)
    Else
        prevQtr = yyyy & "Q" & CStr(nn - 1)
    End If
End If

GetPrevQuarter = prevQtr

End Function
