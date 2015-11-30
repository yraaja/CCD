VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{5936A75C-3F42-11D6-AF6B-AA0004005F12}#1.3#0"; "MeansCtrl.ocx"
Begin VB.Form frmLongDescriptionGrid 
   Caption         =   "Long Description Grid"
   ClientHeight    =   6990
   ClientLeft      =   2205
   ClientTop       =   2625
   ClientWidth     =   11235
   Icon            =   "frmLongDescriptionGrid.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6990
   ScaleWidth      =   11235
   Begin VB.Frame fraUnitCostId 
      Caption         =   "Unit Cost ID"
      Height          =   615
      Left            =   1200
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.ComboBox cboMasterFormat 
         Height          =   315
         Left            =   3600
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   240
         Width           =   1000
      End
      Begin VB.TextBox EndUnitCostID 
         Height          =   315
         Left            =   2240
         TabIndex        =   2
         Top             =   240
         Width           =   1275
      End
      Begin VB.TextBox StartUnitCostID 
         Height          =   315
         Left            =   570
         TabIndex        =   1
         Top             =   240
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "From:"
         Height          =   255
         Left            =   0
         TabIndex        =   17
         Top             =   270
         Width           =   495
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "To:"
         Height          =   255
         Left            =   1920
         TabIndex        =   16
         Top             =   270
         Width           =   255
      End
   End
   Begin VB.CommandButton cmdLongDescrReport 
      Caption         =   "Long Descr Report"
      Height          =   495
      Left            =   2760
      TabIndex        =   11
      Top             =   6240
      Width           =   1275
   End
   Begin VB.Frame fraLongDescFilter 
      Caption         =   "Description Filter"
      Height          =   615
      Left            =   6000
      TabIndex        =   3
      Top             =   0
      Width           =   2055
      Begin VB.TextBox txtLongDescFilter 
         Height          =   315
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdTextSearchReplace 
      Caption         =   "Search/Replace Text"
      Height          =   495
      Left            =   4200
      TabIndex        =   12
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Frame fraMeasSys 
      Caption         =   "Measurement System"
      Height          =   615
      Left            =   8160
      TabIndex        =   5
      Top             =   0
      Width           =   2535
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   2310
         TabIndex        =   20
         Top             =   240
         Width           =   2315
         Begin VB.OptionButton optImperial 
            Caption         =   "Imperial"
            Height          =   255
            Left            =   0
            TabIndex        =   23
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton optMetric 
            Caption         =   "Metric"
            Height          =   255
            Left            =   975
            TabIndex        =   22
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton optMeasSysAll 
            Caption         =   "All"
            Height          =   255
            Left            =   1815
            TabIndex        =   21
            Top             =   0
            Value           =   -1  'True
            Width           =   510
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Go To"
      Height          =   855
      Left            =   60
      TabIndex        =   8
      Top             =   6000
      Width           =   2460
      Begin VB.CommandButton cmdUnitCostUsage 
         Caption         =   "Unit Cost Usage"
         Height          =   495
         Left            =   1320
         TabIndex        =   10
         Top             =   240
         Width           =   915
      End
      Begin VB.CommandButton cmdUnitCost 
         Caption         =   "Unit Cost"
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   495
      Left            =   9930
      TabIndex        =   13
      Top             =   6240
      Width           =   1155
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   495
      Left            =   10800
      TabIndex        =   6
      Top             =   100
      Width           =   1035
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid 
      Height          =   4740
      Left            =   45
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1200
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   8361
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
      CellTips        =   1
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
      _StyleDefs(58)  =   ":id=35,.parent=29,.bgcolor=&HFFFF00&,.ellipsis=0"
      _StyleDefs(59)  =   "Named:id=36:OddRow"
      _StyleDefs(60)  =   ":id=36,.parent=29,.ellipsis=0"
      _StyleDefs(61)  =   "Named:id=39:RecordSelector"
      _StyleDefs(62)  =   ":id=39,.parent=30"
      _StyleDefs(63)  =   "Named:id=42:FilterBar"
      _StyleDefs(64)  =   ":id=42,.parent=29"
   End
   Begin ConstructionCostDatabase.DynaTree FormatTree 
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   450
   End
   Begin VB.Label lblCharacters 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "255 characters available"
      Height          =   195
      Left            =   2060
      TabIndex        =   18
      Top             =   960
      Width           =   1740
   End
   Begin VB.Label lblRowCount 
      Caption         =   "0 rows returned"
      Height          =   255
      Left            =   6840
      TabIndex        =   15
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   360
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   1005
   End
   Begin VB.Line Line2 
      X1              =   75
      X2              =   11055
      Y1              =   720
      Y2              =   720
   End
End
Attribute VB_Name = "frmLongDescriptionGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const LONG_DESC_MAX = 255               ' Max Length of the "Object Description" field

Dim m_objGridMap As New CLongDescMap    ' Class to handle grid
Public m_blnFirstSearch As Boolean      ' Is this the first search we have made on this screen
Dim m_blnJumpIn As Boolean              ' Are we jumping here from another screen
Dim m_rec As New ADODB.RecordSet        ' Recordset to hold query results
Dim m_blnDoubleClick As Boolean         ' Did a double-click just occurr
Dim m_blnWereErrors As Boolean          ' True if the Update had errors, used in QueryUnload
Dim m_intMasterFormat As Long                   ' Stores MasterFormat version to use by Search et al
Dim m_blnMasterFormatNotSpecified As Boolean    ' True if MasterFormat was never explicitly set

Dim m_strCurrentFormControl As String
Dim m_aryLockedMetricColumns() As String

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

Private Function UpdateCharactersAvailable() As Long
' UPDATE lblCharacters CONTROL WITH THE NUMBER OF CHARACTERS
' STILL AVAILABLE FOR ENTRY IN THE LONG/OBJECT DESCRIPTION FIELD
' ADDED 5/23/2005 BY RTD, FOR VERSION 7.3
    Dim i As Long
    Dim iUsed As Long
    Dim iAvailable As Long
    Dim sDesc As String
    
    If TDBGrid.Row >= 0 Then
        ' UPDATED 8/24/2005 RTD - OBJECT DESC COLUMN NOW STARTS AT 5
        For i = 5 To TDBGrid.Columns.Count - 1 Step 2
            If TDBGrid.Columns(i).Text <> "" Then
                If sDesc = "" Then
                    sDesc = UCase(Mid(TDBGrid.Columns(i).Text, 1, 1)) + Right(TDBGrid.Columns(i).Text, (Len(TDBGrid.Columns(i).Text) - 1))
                Else
                    sDesc = sDesc + ", " + TDBGrid.Columns(i).Text
                End If
            End If
        Next i
        If Len(sDesc) > LONG_DESC_MAX Then sDesc = Left(sDesc, LONG_DESC_MAX)
        iUsed = Len(sDesc)
        iAvailable = LONG_DESC_MAX - iUsed
        lblCharacters.Caption = iAvailable & " characters available"
        lblCharacters.Refresh
        UpdateCharactersAvailable = iAvailable
    Else
        lblCharacters.Caption = ""
        lblCharacters.Refresh
        UpdateCharactersAvailable = -1
    End If
    
End Function

Private Sub LoadLockedCols()
'Update the long description:
'  1st description has 1st char uppercase,
'  all are comma separated
    Dim i As Integer
    Dim sDesc As String

    ' UPDATED 8/24/2005 RTD - OBJECT DESC COLUMN NOW STARTS AT 5
    For i = 5 To TDBGrid.Columns.Count - 1 Step 2
        If TDBGrid.Columns(i).Text <> "" Then
            If sDesc = "" Then
                sDesc = UCase(Mid(TDBGrid.Columns(i).Text, 1, 1)) + Right(TDBGrid.Columns(i).Text, (Len(TDBGrid.Columns(i).Text) - 1))
            Else
                sDesc = sDesc + ", " + TDBGrid.Columns(i).Text
            End If
        End If
    Next i
    
    If Len(sDesc) > 255 Then sDesc = Left(sDesc, 255)
    TDBGrid.Columns("Object Desc").Text = sDesc
    UpdateCharactersAvailable
    
End Sub

Public Sub SelectAllRows()
    ' Pass recordset to handler class
    m_objGridMap.RecordSet = m_rec
    
    If m_rec.RecordCount > 0 Then
        m_objGridMap.SelectAllRows
    End If
End Sub

Private Sub cmdLongDescrReport_Click()
    TDBGrid.Update
    PreviewReport
End Sub

' Called when coming here from another screen
Public Sub JumpIn(strMatID As String)
    ' CALLED FROM OTHER FORM;
    ' DO NOT SHOW THE MASTERFORMAT TREE TO SIMULATE FORM'S PREVIOUS BEHAVIOR
    ShowMasterFormatTree False
    StartUnitCostID.Text = Compress_String(strMatID)
    If m_blnMasterFormatNotSpecified Then
        ' MF was never explicitly set, so default to 1995 for compatibility purposes
        MasterFormat = 1995
    End If
    cmdSearch_Click
End Sub

Public Sub PrintReport()
    PreviewReport
End Sub

Public Sub PreviewReport()
    Dim rtn As Integer
    Dim strSelect As String
    Dim frmLongDescr As New frmLongDescRpt
    Dim blnReturn As Boolean

    On Error Resume Next
    Screen.MousePointer = vbNormal  '<-- Reset in case some problem leaves as hourglass.
    If m_rec.RecordCount > 0 Then
        With frmLongDescr
            .LoadReport m_rec
            .RenderReport
            .Show
        End With
    Else
        MsgBox "You must display the records you want to report using the Search button.", vbInformation, "Information"
        GoTo Exit_Sub
    End If
    Exit Sub

Exit_Sub:
    Screen.MousePointer = vbNormal
    
End Sub

Private Sub cmdUnitCostUsage_Click()
    Dim sUnitCostId As String
    
    TDBGrid.Update
    If IsNumeric(TDBGrid.Bookmark) = False Then
        MsgBox "You must select a row."
        Exit Sub
    End If
    ' 8/24/2005 RTD
    ' Get the Unit Cost ID and set MasterFormat version
    If MasterFormat = EXT_MASTERFORMAT_VERSION Then
        sUnitCostId = TDBGrid.Columns("Unit Cost ID " & Right(EXT_MASTERFORMAT_VERSION, 2)).Text
    Else
        sUnitCostId = TDBGrid.Columns("Unit Cost ID").Text
    End If
    ' Navigate to single-record view
    Dim frm As frmUCostUsageGrid
    Dim rec As ADODB.RecordSet
    Set frm = New frmUCostUsageGrid
    frm.MasterFormat = MasterFormat
    frm.JumpIn Compress_String(sUnitCostId) ' Pass the current record into the form
    frm.Show
End Sub

Private Sub cmdUnitCost_Click()
    Dim sUnitCostId As String
    
    TDBGrid.Update
    If IsNumeric(TDBGrid.Bookmark) = False Then
        MsgBox "You must select a row."
        Exit Sub
    End If
    
    ' Navigate to unit cost grid
    Dim frm As frmUnitCostGrid
    Set frm = New frmUnitCostGrid
    ' 8/24/2005 RTD
    ' Get the Unit Cost ID and set MasterFormat version
    If MasterFormat = EXT_MASTERFORMAT_VERSION Then
        sUnitCostId = TDBGrid.Columns("Unit Cost ID " & Right(EXT_MASTERFORMAT_VERSION, 2)).Text
    Else
        sUnitCostId = TDBGrid.Columns("Unit Cost ID").Text
    End If
    ' Get the selected row from grid - send the MasterFormat
    frm.MasterFormat = MasterFormat
    frm.JumpIn Compress_String(sUnitCostId)   ' Pass the current unit cost id into the form
    frm.Show
    
End Sub

Public Sub Sort(intDir As Integer)
    m_objGridMap.Sort intDir
End Sub

Public Sub cmdUpdate_Click()
    On Error GoTo Out
    Dim blnRet As Boolean
    Dim vntBookmark As Variant
    
    'rlh 05/22/2007
    If MasterFormat = EXT_MASTERFORMAT_VERSION Then
     Else
        MsgBox "Descriptions Update for MF-1995 has been disabled from 05/2007 forward." & vbCrLf & " Please see IT" & _
        " Liaison for advice.  Thank you"
        Exit Sub
    End If
    
    vntBookmark = TDBGrid.Bookmark
    TDBGrid.Update
    blnRet = m_objGridMap.Update
    If blnRet = False Then
        m_blnWereErrors = True  'Update failed
    Else
        m_blnWereErrors = False
    End If
    TDBGrid.Bookmark = vntBookmark
Out:
End Sub

Private Sub cmdTextSearchReplace_Click()
    Dim intCol As Integer, intColCount As Integer, intColStart As Integer, intColEnd As Integer
    Dim intVBCompare As Integer
    Dim lngFirstMatchBookmark As Long, lngCurrentBookmark As Long
    Dim strSearchText As String, strReplacementText As String, strGridText As String
    Dim blnFindAll As Boolean, blnReplace  As Boolean, blnReplaceAll As Boolean

    '-- Open Search window
    frmSearchLongDescription.Show vbModal
    
    '-- Retrieve public properties of the Search window
    With frmSearchLongDescription
        If .Cancel Or Trim(.SearchText) = "" Then Exit Sub
        strSearchText = .SearchText
        If .MatchCase Then
            intVBCompare = vbBinaryCompare
        Else
            intVBCompare = vbTextCompare
        End If
            
        strReplacementText = .ReplacementText
        blnFindAll = .FindAll
        blnReplace = .Replace
        blnReplaceAll = .ReplaceAll
    End With
    
    ' UPDATED 8/24/2005 RTD - OBJECT DESC COLUMN NOW STARTS AT 5
    '-- Set search column to Column 4--Obj Desc
    intColStart = 5
    intColEnd = TDBGrid.Columns.Count - 1

    TDBGrid.ClearSelCols   '<-- Deselect any selected column(s) as selection will mark any matches.
    TDBGrid.SelBookmarks.Clear  '<-- Clear any row highlighting.

    lngFirstMatchBookmark = -1
    TDBGrid.MoveFirst
    '-- Traverse the grid and search/replace if matched.
    Do While Not TDBGrid.EOF
    
        For intCol = intColStart To intColEnd
            TDBGrid.Col = intCol
            If InStr(1, TDBGrid.Text, strSearchText, intVBCompare) > 0 Then
                TDBGrid.SelBookmarks.Add TDBGrid.Bookmark '<-- select the row
                If lngFirstMatchBookmark = -1 Then lngFirstMatchBookmark = TDBGrid.Bookmark
                If blnReplace Then
                    TDBGrid.Text = Replace(TDBGrid.Text, strSearchText, strReplacementText, , , intVBCompare)
                    If Not blnReplaceAll Then GoTo Xit
                Else
                    If Not blnFindAll Then GoTo Xit
                End If
            End If
        Next intCol
        TDBGrid.MoveNext
    Loop
Xit:
    '-- Position to the first hit, if there is one, otherwise position to first row
    If lngFirstMatchBookmark <> -1 Then
        TDBGrid.Bookmark = lngFirstMatchBookmark
    Else
        TDBGrid.MoveFirst
    End If
    TDBGrid.Col = 0 '-- Position top left.

End Sub

Private Sub Form_Deactivate()
    
    HideGridSort
    fMainForm.tbToolBar.Buttons.Item(tbrPRINT).Enabled = False
    fMainForm.tbToolBar.Buttons.Item(tbrPRINT).Visible = False
    fMainForm.tbToolBar.Buttons.Item(tbrPREVIEW).Enabled = False
    fMainForm.tbToolBar.Buttons.Item(tbrPREVIEW).Visible = False
    fMainForm.mnuFilePageSetup.Enabled = False
    fMainForm.mnuFilePrint.Enabled = False
    fMainForm.mnuFilePrintPreview.Enabled = False
    m_strCurrentFormControl = Me.ActiveControl.Name

End Sub

Private Sub Form_Initialize()
    
    Status ("Loading Long Description Grid...")
    
    m_intMasterFormat = g_intMasterFormat
    
    ' Fill the MasterFormat tree
    FormatTree.InitData g_cnShared, "UNITCOST"
    
    Screen.MousePointer = vbHourglass
    m_blnFirstSearch = True
    ' Initialize grid
    m_objGridMap.SetGrid TDBGrid
    m_objGridMap.InitGrid
    m_blnJumpIn = False
    Screen.MousePointer = vbNormal
    m_blnFirstSearch = False
    
    m_blnMasterFormatNotSpecified = True
    
End Sub

Private Sub Form_Load()
    Dim strSelect As String
    
    Move START_LEFT, START_TOP, START_WIDTH, START_HEIGHT
    
    LoadMasterFormatCombo Me.cboMasterFormat, True
    
    StartUnitCostID.Text = "~"
    cmdSearch_Click
    StartUnitCostID.Text = ""
    Status ("")

End Sub

Private Sub Form_LostFocus()
    TDBGrid.Update
End Sub

Private Sub Form_Resize()
    Dim i As Integer
    On Error Resume Next
    If Me.WindowState = vbNormal Or Me.WindowState = vbMaximized Then
        If Me.Width >= 11250 Then
            TDBGrid.Width = Me.Width - 255
            Line2.X2 = Me.Width - 210
        Else
            Me.Width = 11250
        End If
        
        If Me.Height >= 7260 Then
            'TDBGrid.Height = Me.Height - 2500
            TDBGrid.Height = Me.Height - TDBGrid.Top - 1420
            Frame1.Top = Me.Height - 1260
            cmdUpdate.Top = Me.Height - 1020
            cmdTextSearchReplace.Top = Me.Height - 1020
            cmdLongDescrReport.Top = Me.Height - 1020
        Else
            Me.Height = 7260
        End If
    Else
        ShowMinimizedForms
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    HideGridSort
    fMainForm.tbToolBar.Buttons.Item(tbrPRINT).Enabled = False
    fMainForm.tbToolBar.Buttons.Item(tbrPRINT).Visible = False
    fMainForm.tbToolBar.Buttons.Item(tbrPREVIEW).Enabled = False
    fMainForm.tbToolBar.Buttons.Item(tbrPREVIEW).Visible = False
    fMainForm.mnuFilePageSetup.Enabled = False
    fMainForm.mnuFilePrint.Enabled = False
    fMainForm.mnuFilePrintPreview.Enabled = False
    Set frmLongDescriptionGrid = Nothing
    
End Sub

Private Sub cmdSearch_Click()
    On Error Resume Next
    Dim blnReturn As Boolean
    Dim strSelect As String
    Dim dtmToday As Date
    Dim dtmStart As Date
    Dim strStartUnitCostSrch As String
    Dim meas_sys_cd As String
    Dim rsLongDescClone As ADODB.RecordSet
    Dim iMasterFormat As Long
    
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
            TDBGrid.ReBind
            Exit Sub
        Else
            TDBGrid.DataChanged = False
        End If
    End If
    Screen.MousePointer = vbHourglass
    dtmToday = Date
    
    '8/23/2005 RTD - Get the user MasterFormat choice
    If m_blnFirstSearch Then
        iMasterFormat = 1995
    Else
        iMasterFormat = cboMasterFormat.ItemData(cboMasterFormat.ListIndex)
    End If
    
    If Len(StartUnitCostID.Text) = 0 Then
        Screen.MousePointer = vbNormal
        MsgBox "You must enter Unit Cost ID or Description.", vbExclamation
        Exit Sub
    End If
    If Len(StartUnitCostID) = 12 And InStr(1, StartUnitCostID, "*") = 0 And Len(EndUnitCostID) = 0 Then
        strStartUnitCostSrch = Compress_String(StartUnitCostID) + "*"
    Else
        strStartUnitCostSrch = Compress_String(StartUnitCostID)
    End If
    
    lblRowCount.Caption = "Working..."
    lblRowCount.Refresh

    m_rec.Close ' Make sure it is closed
    m_rec.MaxRecords = MAX_RECORDS ' Set the maximum number to bring back
    dtmStart = Now

    Dim strError As String
    
    If optMeasSysAll = True Then
        meas_sys_cd = "A"
    ElseIf optImperial = True Then
        meas_sys_cd = "I"
    ElseIf optMetric = True Then
        meas_sys_cd = "M"
    End If
    
    ' MODIFIED 8/24/2005 - RTD
    ' CHANGE TO RETRIEVE LONG DESCRIPTION GRID DATA USING
    ' MASTERFORMAT 2004 AWARE STORED PROC: usp_select_attribute_value_ext
    strSelect = "exec usp_select_attribute_value_ext @min_object_id = '" & SQLChangeWildcard(strStartUnitCostSrch) & _
        "', @max_object_id = '" & Compress_String(EndUnitCostID.Text) & _
        "', @skey_type = 'U', @meas_sys_cd = '" & meas_sys_cd & _
        "', @obj_desc_filter = '" & SQLChangeWildcard(txtLongDescFilter.Text) & _
        "', @master_format = " & iMasterFormat
    'Use DAL to perform select
    'Debug.Print strSELECT
    blnReturn = g_objDAL.GetRecordset(vbNullString, strSelect, m_rec)
    'Debug.Print strSELECT
    'm_rec.MoveFirst
    'Dim fld As ADODB.Field
    'For Each fld In m_rec.Fields
    '    Debug.Print fld.Name & " / " & fld.Value
    'Next fld
    
    If blnReturn = False Then
        MsgBox "An error occurred while searching."
        lblRowCount.Caption = "0 rows returned."
        Screen.MousePointer = vbNormal
        Exit Sub
    End If
    
    ' Set MasterFormat to match new results
    MasterFormat = iMasterFormat
    
    If m_rec.RecordCount > 0 Then
        lblRowCount.Caption = str(m_rec.RecordCount) + " rows returned in " + str(DateDiff("s", dtmStart, Now)) + " seconds"
        cmdTextSearchReplace.Enabled = True
    Else
        lblRowCount.Caption = "0 rows returned."
        cmdTextSearchReplace.Enabled = False
    End If
    
    ' Pass recordset to handler class
    m_objGridMap.RecordSet = m_rec
    
    ' If the upper bound was hit, inform user
    If m_rec.RecordCount = MAX_RECORDS And m_rec.State = adStateOpen Then
        MsgBox "The search returned the maximum number of records allowed. More records may be available."
    End If
    m_objGridMap.SetMenuBar
    ' Reset the grid contents
    TDBGrid.Bookmark = Null
    TDBGrid.ReBind
    TDBGrid.ApproxCount = m_rec.RecordCount
'*** APEX Migration Utility Code Change ***
'    Dim Col As TrueOleDBGrid70.Column
    Dim Col As TrueOleDBGrid80.Column
    For Each Col In TDBGrid.Columns 'for dynamically built columns that are not drop down boxes, resize
        If Col.ValueItems.Presentation <> dbgComboBox And Col.Caption <> "Unit Cost ID" And Col.Caption <> "MSys" And Col.Caption <> "Object Desc" And Col.Caption <> "obj_skey" Then
            Col.AutoSize
            TDBGrid.Refresh
        End If
    Next Col
    Set Col = Nothing
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
        rs.Close
    End If
    txtLongDescFilter.Text = ""
    optMeasSysAll.value = True
    ' Kick-off search
    cmdSearch_Click
    
End Sub

Private Sub TDBGrid_AfterColUpdate(ByVal ColIndex As Integer)
    'Update the long description:
    '  1st description has 1st char uppercase,
    '  all are comma separated
    Dim i As Integer
    Dim sDesc As String

    sDesc = ""
    ' UPDATED 8/24/2005 RTD - OBJECT DESC COLUMN NOW STARTS AT 5
    For i = 5 To TDBGrid.Columns.Count - 1 Step 2
        If TDBGrid.Columns(i).Text <> "" Then
            If sDesc = "" Then
                sDesc = UCase(Mid(TDBGrid.Columns(i).Text, 1, 1)) + Right(TDBGrid.Columns(i).Text, (Len(TDBGrid.Columns(i).Text) - 1))
            Else
                sDesc = sDesc + ", " + TDBGrid.Columns(i).Text
            End If
        End If
    Next i
    If Len(sDesc) > 255 Then
        sDesc = Left(sDesc, 255)
        MsgBox "The 'object description' field is greater than 255 characters. Any characters over 255 have been removed.", vbInformation + vbOKOnly, "Warning"
    End If
    TDBGrid.Columns("Object Desc").Text = sDesc
    UpdateCharactersAvailable
    
End Sub

Private Sub TDBGrid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    If TDBGrid.Columns(ColIndex).ValueItems.Presentation = dbgComboBox Then 'meas sys change
        Select Case TDBGrid.Columns("MSys")
        Case "M"
            If TDBGrid.Columns(ColIndex).value = "I" Then     'Invalid
                MsgBox "Only M or A is valid for a Metric row.", vbExclamation
                Cancel = True
            End If
        Case "I"
            If TDBGrid.Columns(ColIndex).value = "M" Then     'Invalid
                MsgBox "Only I or A is valid for an Imperial row.", vbExclamation
                Cancel = True
            End If
        End Select
    End If
End Sub

Private Sub TDBGrid_Change()
    With TDBGrid
        If Len(.Text) > 255 Then .Text = Left(.Text, 255)
    End With
End Sub

Private Sub TDBGrid_DblClick()
    ' Signal that double-click has occurred
    m_blnDoubleClick = True
End Sub

Private Sub TDBGrid_Error(ByVal DataError As Integer, Response As Integer)
    Response = 0
    TDBGrid.DataChanged = False
End Sub

'*** APEX Migration Utility Code Change ***
'Private Sub TDBGrid_FetchCellTips(ByVal SplitIndex As Integer, ByVal ColIndex As Integer, ByVal RowIndex As Long, CellTip As String, ByVal FullyDisplayed As Boolean, ByVal TipStyle As TrueOleDBGrid70.StyleDisp)
Private Sub TDBGrid_FetchCellTips(ByVal SplitIndex As Integer, ByVal ColIndex As Integer, ByVal RowIndex As Long, CellTip As String, ByVal FullyDisplayed As Boolean, ByVal TipStyle As TrueOleDBGrid80.StyleDisp)
    ' Display Cell Tip for the "Object Description" column
    ' ADDED 5/31/2005 RTD
    
    If ColIndex <> TDBGrid.Columns("Object Desc").ColIndex Then
        CellTip = ""
    End If
    
End Sub

Private Sub TDBGrid_GotFocus()
    TDBGrid.TabStop = True
End Sub

Private Sub TDBGrid_KeyPress(KeyAscii As Integer)
   'Dim I As Integer
   With TDBGrid
        If .Columns(.Col).Caption = "Tech Desc" Or .Columns(.Col).Caption = "Metric Tech Desc" Then
            If KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn Or KeyAscii = vbKeyDelete Then
                ' do nothing. allow the key
            Else
                If .SelLength <= 0 And Len(.Text) + 1 > 75 Then KeyAscii = 0
            End If
        End If
    End With
    
End Sub

Private Sub TDBGrid_KeyUp(KeyCode As Integer, Shift As Integer)
    
    UpdateCharactersAvailable
    
End Sub

Private Sub TDBGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    UpdateCharactersAvailable
    
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
'            cmdMaterialPrice_Click
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
On Error Resume Next
Dim ctl As Control

    fMainForm.tbToolBar.Buttons.Item(tbrPRINT).Enabled = True
    fMainForm.tbToolBar.Buttons.Item(tbrPRINT).Visible = True
    fMainForm.tbToolBar.Buttons.Item(tbrPREVIEW).Enabled = True
    fMainForm.tbToolBar.Buttons.Item(tbrPREVIEW).Visible = True
    fMainForm.mnuFilePageSetup.Enabled = True
    fMainForm.mnuFilePrint.Enabled = True
    fMainForm.mnuFilePrintPreview.Enabled = True

    If Me.WindowState <> vbMinimized Then
        If Len(m_strCurrentFormControl) > 0 Then
            For Each ctl In Me.Controls
                If ctl.Name = m_strCurrentFormControl Then
                    ctl.SetFocus
                    Exit For
                End If
            Next ctl
        End If
        ShowGridSort
        OutputView False
        m_objGridMap.SetMenuBar
    End If
    
End Sub

Public Function ShowMasterFormatTree(bTreeIsVisible As Boolean) As Boolean
' IF bTreeIsVisible, SHIFT THE CONTROLS AND GRID
' TO MAKE ROOM FOR THE NEW MASTER FORMAT TREE
' ADDED 5/27/2005 RTD

    If bTreeIsVisible Then
        Me.Label4.Top = 60
        Me.Label4.Left = 6780
        Me.fraUnitCostId.Top = 480
        Me.fraUnitCostId.Left = 7240
        Me.fraLongDescFilter.Top = 1200
        Me.fraLongDescFilter.Left = 7240
        Me.fraMeasSys.Top = 1920
        Me.fraMeasSys.Left = 7240
        Me.cmdSearch.Top = 2040
        Me.cmdSearch.Left = 10120
        TDBGrid.Top = 3240
        FormatTree.Visible = True
        FormatTree.Height = 2640
        FormatTree.Width = 6620
    Else
        Me.Label4.Top = 0
        Me.Label4.Left = 120
        Me.fraUnitCostId.Top = 0
        Me.fraUnitCostId.Left = 1200
        Me.fraLongDescFilter.Top = 0
        Me.fraLongDescFilter.Left = 6000
        Me.fraMeasSys.Top = 0
        Me.fraMeasSys.Left = 8160
        Me.cmdSearch.Top = 100
        Me.cmdSearch.Left = 10800
        TDBGrid.Top = 1200
        FormatTree.Visible = False
    End If
    Line2.Y1 = TDBGrid.Top - 480
    Line2.Y2 = Line2.Y1
    lblRowCount.Top = Line2.Y1 + 120
    lblCharacters.Top = Line2.Y1 + 240
    
    TDBGrid.Height = Me.Height - TDBGrid.Top - 1420
    
    ShowMasterFormatTree = bTreeIsVisible
    
End Function

Public Function SelectMasterFormat(iMasterFormat As Long) As Boolean
'SET THE MASTERFORMAT COMBO BOX TO THE NEW SELECTION
'ADDED 8/23/2005 RTD
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


