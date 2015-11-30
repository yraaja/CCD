VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmTradeGroup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trade Group Maintenance"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8355
   Icon            =   "frmTradeGroup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5430
   ScaleWidth      =   8355
   Visible         =   0   'False
   Begin VB.TextBox loc_id 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   315
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   28
      Tag             =   "1N"
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox chkDisplayHistory 
      Caption         =   "Display Trade/Location History"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5520
      TabIndex        =   7
      Top             =   1800
      Visible         =   0   'False
      Width           =   2655
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid 
      Height          =   2055
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Visible         =   0   'False
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   3625
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
   Begin VB.ComboBox Trade_Group_Code 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   1485
   End
   Begin VB.ComboBox State_Code 
      Height          =   315
      Left            =   1320
      TabIndex        =   3
      Top             =   1290
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.ComboBox Trade_ID 
      Height          =   315
      Left            =   1320
      TabIndex        =   2
      Tag             =   "0"
      Top             =   870
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.TextBox trade_desc 
      Height          =   285
      Left            =   2925
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   900
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.TextBox NewTradeGroupCode 
      Height          =   285
      Left            =   4560
      TabIndex        =   1
      Tag             =   "1S"
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox start_date 
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Tag             =   "1D"
      Top             =   1710
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.TextBox term_date 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4200
      TabIndex        =   6
      Tag             =   "1D"
      Top             =   1710
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.ComboBox City 
      Height          =   315
      ItemData        =   "frmTradeGroup.frx":0442
      Left            =   2925
      List            =   "frmTradeGroup.frx":0444
      TabIndex        =   4
      Text            =   "City"
      Top             =   1290
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox last_update_id 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   315
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   15
      Tag             =   "1N"
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox trade_skey 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   315
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Tag             =   "1G"
      Top             =   4380
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox last_update_person 
      BackColor       =   &H00C0C0C0&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   315
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4380
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox last_update_date 
      BackColor       =   &H00C0C0C0&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   315
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4380
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Default         =   -1  'True
      Height          =   495
      Left            =   3360
      TabIndex        =   10
      Top             =   4860
      Width           =   1150
   End
   Begin VB.Label lblTrades 
      Alignment       =   2  'Center
      Caption         =   "Trades"
      Height          =   255
      Left            =   960
      TabIndex        =   27
      Top             =   480
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblLaborRecords 
      Alignment       =   2  'Center
      Caption         =   "Labor Rate Records"
      Height          =   255
      Left            =   5760
      TabIndex        =   26
      Top             =   480
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblLocations 
      Alignment       =   2  'Center
      Caption         =   "Locations"
      Height          =   255
      Left            =   3240
      TabIndex        =   25
      Top             =   480
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "City:"
      Height          =   255
      Left            =   2430
      TabIndex        =   24
      Top             =   1350
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "State:"
      Height          =   255
      Left            =   480
      TabIndex        =   23
      Top             =   1350
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Trade ID:"
      Height          =   255
      Left            =   300
      TabIndex        =   22
      Top             =   930
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblNewTradeGroup 
      Alignment       =   1  'Right Justify
      Caption         =   "New Trade Group:"
      Height          =   255
      Left            =   2760
      TabIndex        =   20
      Top             =   150
      Width           =   1635
   End
   Begin VB.Label Label20 
      Caption         =   "Trade_Skey"
      Height          =   255
      Left            =   6000
      TabIndex        =   19
      Top             =   4440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Trade Group:"
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   150
      Width           =   1155
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Start Date:"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   1740
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Term Date:"
      Height          =   255
      Left            =   3270
      TabIndex        =   16
      Top             =   1740
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Last Updated By:"
      Height          =   255
      Left            =   3180
      TabIndex        =   13
      Top             =   4440
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Last Updated:"
      Height          =   255
      Left            =   300
      TabIndex        =   12
      Top             =   4440
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "frmTradeGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_objGridMap As New CTradeHistMap ' Class to handle grid
Dim m_rec As ADODB.RecordSet
Dim m_rec2 As New ADODB.RecordSet
Dim m_blnRecFlag As Boolean ' True if the screen has a data source, False if we are adding new record
Dim m_blnDeleted As Boolean ' Indicates if the data has been deleted, used in QueryUnload
Dim m_blnWereErrors As Boolean ' True if the Update had errors, used in QueryUnload
Dim m_blnInsert As Boolean
Dim m_State As String
Dim m_trade_group_code As String
Dim m_trade_id As String
Dim m_city As String
Public blnAddMbr As Boolean
Public blnNewGroup As Boolean
Public frmCallingForm As Form

Public Sub get_counts()
    Dim strSelect As String
    Dim blnRet As Boolean
    Dim strError As String
    Dim rsTemp As New ADODB.RecordSet
    Dim lOrigPointer As Long
If Trade_Group_Code.Text <> "" Then
'Count trade IDs for group
    lOrigPointer = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    strSelect = "exec usp_count_trades @trade_group_code = '" + Trade_Group_Code.Text + "'"
    blnRet = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
    If Not blnRet Then
        MsgBox "An error occurred retrieving data."
    Else
        If rsTemp.RecordCount > 0 Then
            lblTrades = CStr(rsTemp.Fields("Trades")) + " Trade(s)"
            lblTrades.Visible = True
            lblLocations = CStr(rsTemp.Fields("locations")) + " Location(s)"
            lblLocations.Visible = True
            lblLaborRecords = CStr(rsTemp.Fields("labor_records")) + " Labor Rate Record(s)"
            lblLaborRecords.Visible = True
            start_date = rsTemp.Fields("maxstartdate")
        End If
    End If
    rsTemp.Close
    Set rsTemp = Nothing

    If blnAddMbr = True Then
        Load_Trade_IDs
    End If
    Screen.MousePointer = vbNormal
End If
End Sub

Public Sub Load_Trade_IDs()
'Add Member setup:
'   Trade IDs will be entered that have a blank trade group
'   The start date must match that of the group
    Dim blnReturn As Boolean
    Dim strSelect As String
    Dim rsTemp As RecordSet
    Trade_ID.Clear
'Load Trade IDs
    DoEvents
    If blnAddMbr = True Then
        If IsDate(start_date) Then
'            strSelect = "SELECT labor_trade.trade_id, labor_trade.trade_skey FROM labor_rate as lr" + _
'            " inner join labor_trade on lr.trade_skey = labor_trade.trade_skey" + _
'            " where lr.trade_group_code = ''" + _
'            " and (convert(varchar(2),DATEPART(m, lr.term_date)) + '/' + convert(varchar(2),DATEPART(d, lr.term_date)) + '/' + convert(varchar(4),DATEPART(yyyy, lr.term_date)))  = '" + PriorDay(start_date) + _
'            " ' and lr.start_date = (select max(start_date) from labor_rate where labor_rate.trade_skey = lr.trade_skey and labor_rate.loc_id " + _
'            " = lr.loc_id) GROUP BY labor_trade.trade_id, labor_trade.trade_skey "
'
'AK- 6/7/2006 Updated the Query to be consistent and efficient

            strSelect = "SELECT distinct labor_trade.trade_id, labor_trade.trade_skey FROM labor_rate as lr" + _
            " inner join labor_trade on lr.trade_skey = labor_trade.trade_skey" + _
            " where lr.trade_group_code = ''" + _
            " and lr.start_date = (select max(start_date) from labor_rate where labor_rate.trade_skey = lr.trade_skey and labor_rate.loc_id " + _
            " = lr.loc_id) ORDER BY labor_trade.trade_id, labor_trade.trade_skey "
        Else
            strSelect = "SELECT distinct labor_trade.trade_id, labor_trade.trade_skey FROM labor_rate as lr" + _
            " inner join labor_trade on lr.trade_skey = labor_trade.trade_skey" + _
            " where lr.trade_group_code = ''" + _
            " and lr.start_date = (select max(start_date) from labor_rate where labor_rate.trade_skey = lr.trade_skey and labor_rate.loc_id " + _
            " = lr.loc_id) ORDER BY labor_trade.trade_id "
        End If
    Else
        strSelect = "SELECT distinct labor_trade.trade_id, labor_trade.trade_skey FROM labor_rate as lr" + _
        " inner join labor_trade on lr.trade_skey = labor_trade.trade_skey" + _
        " where lr.trade_group_code = ''" + _
        " and lr.start_date = (select max(start_date) from labor_rate where labor_rate.trade_skey = lr.trade_skey and labor_rate.loc_id " + _
        " = lr.loc_id) ORDER BY labor_trade.trade_id "
        
    End If

    ' Use DAL to perform select
    blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
    If blnReturn = False Then
        MsgBox "An error occurred loading Trade IDs."
    Else
        If Not (rsTemp.EOF And rsTemp.BOF) Then
            Do Until rsTemp.EOF
                Trade_ID.AddItem rsTemp![Trade_ID]
                Trade_ID.ItemData(Trade_ID.NewIndex) = rsTemp![trade_skey]
                rsTemp.MoveNext
            Loop
        End If
    End If
    rsTemp.Close
    LoadStates
End Sub

Public Sub LoadTradeGroups()
    Dim blnReturn As Boolean
    Dim strSelect As String
    Dim rsTemp As RecordSet

'Load Trade Groups
    If blnAddMbr = True Then
        strSelect = "select distinct trade_group_code from labor_rate order by trade_group_code"
    Else
        strSelect = "select distinct trade_group_code from labor_rate as lr " + _
        "where start_date = '" + start_date + "' and start_date = (select max(start_date) from labor_rate where labor_rate.trade_skey = lr.trade_skey and " + _
        "labor_rate.loc_id = lr.loc_id) order by trade_group_code"
    End If
    ' Use DAL to perform select
    blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
    If blnReturn = False Then
        MsgBox "An error occurred loading Trade Groups."
    Else
        If Not (rsTemp.EOF And rsTemp.BOF) Then
            Do Until rsTemp.EOF
                Trade_Group_Code.AddItem rsTemp![Trade_Group_Code]
                rsTemp.MoveNext
            Loop
        End If
    End If
    If rsTemp.RecordCount = 1 Then
        rsTemp.MoveFirst
        If Trim(rsTemp![Trade_Group_Code]) = "" Then
            'No groups found - add group
            Trade_Group_Code.Visible = False
            NewTradeGroupCode.Visible = True
            NewTradeGroupCode.Left = Trade_Group_Code.Left
        End If
    End If
    rsTemp.Close
End Sub
Private Sub LoadStates()
'Load States
    Dim blnReturn As Boolean
    Dim strSelect As String
    Dim rsTemp As RecordSet
    State_Code.Clear
    City.Clear
    chkDisplayHistory = 0
    chkDisplayHistory.Enabled = False
    If Trade_ID.ListIndex <> -1 Then
        If blnAddMbr Then
            If IsDate(start_date) Then
'                strSelect = "SELECT distinct location.state_code FROM labor_rate as lr " + _
'                "inner join location on lr.loc_id = location.loc_id " + _
'                 "where lr.trade_skey = " + CStr(Trade_ID.ItemData(Trade_ID.ListIndex)) + _
'                 " and (convert(varchar(2),DATEPART(m, lr.term_date)) + '/' + convert(varchar(2),DATEPART(d, lr.term_date)) + '/' + convert(varchar(4),DATEPART(yyyy, lr.term_date)))  = '" + PriorDay(start_date) + _
'                "' and lr.trade_group_code = '' " + _
'                " and (convert(varchar(2),DATEPART(m, (select max(term_date) from labor_rate" + _
'                " where labor_rate.trade_skey = lr.trade_skey and labor_rate.loc_id  = lr.loc_id)))) + '/'" + _
'                " + (convert(varchar(2),DATEPART(d, (select max(term_date) from labor_rate" + _
'                " where labor_rate.trade_skey = lr.trade_skey and labor_rate.loc_id  = lr.loc_id)))) + '/'" + _
'                " + (convert(varchar(4),DATEPART(yyyy, (select max(term_date) from labor_rate " + _
'                " where labor_rate.trade_skey = lr.trade_skey and labor_rate.loc_id  = lr.loc_id))))  = '" + PriorDay(start_date) + _
'                "' ORDER BY  location.state_code"
'AK- 6/7/2006 - update to be consistent in State retrieval
                
            strSelect = "SELECT distinct location.state_code FROM labor_rate as lr " + _
            "inner join location on lr.loc_id = location.loc_id " + _
             "where lr.trade_skey = " + CStr(Trade_ID.ItemData(Trade_ID.ListIndex)) + _
            " and lr.trade_group_code = '' and lr.start_date = (select max(start_date) from labor_rate where labor_rate.trade_skey = lr.trade_skey and labor_rate.loc_id " + _
            " = lr.loc_id) ORDER BY  location.state_code"
            Else
                strSelect = "SELECT distinct location.state_code FROM labor_rate as lr " + _
                "inner join location on lr.loc_id = location.loc_id " + _
                 "where lr.trade_skey = " + CStr(Trade_ID.ItemData(Trade_ID.ListIndex)) + _
                 " and lr.trade_group_code = '' and lr.start_date = (select max(start_date) from labor_rate where labor_rate.trade_skey = lr.trade_skey and labor_rate.loc_id " + _
                " = lr.loc_id) ORDER BY  location.state_code"
            End If
        Else
            strSelect = "SELECT distinct location.state_code FROM labor_rate as lr " + _
            "inner join location on lr.loc_id = location.loc_id " + _
             "where lr.trade_skey = " + CStr(Trade_ID.ItemData(Trade_ID.ListIndex)) + _
            " and lr.trade_group_code = '' and lr.start_date = (select max(start_date) from labor_rate where labor_rate.trade_skey = lr.trade_skey and labor_rate.loc_id " + _
            " = lr.loc_id) ORDER BY  location.state_code"
        End If
        
        ' Use DAL to perform select
        blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
        If blnReturn = False Then
            MsgBox "An error occurred loading States."
        Else
            If Not (rsTemp.EOF And rsTemp.BOF) Then
                Do Until rsTemp.EOF
                    State_Code.AddItem rsTemp![State_Code]
                    rsTemp.MoveNext
                Loop
            End If
        End If
        rsTemp.Close
        If State_Code.listcount > 0 Then
            State_Code.ListIndex = 0
        End If
        LoadCities
    End If
End Sub

Private Sub LockField(sFieldName As String)
    Me.Controls(sFieldName).Locked = True
    Me.Controls(sFieldName).BackColor = LTGREY
    Me.Controls(sFieldName).TabStop = False
End Sub
Private Function PriorDay(datFullDate As Date) As String
            Dim datPriorDate As Date
            If IsDate(datFullDate) Then
                datPriorDate = DateAdd("d", -1, start_date)
                PriorDay = Format(datPriorDate, "m") + "/" + Format(datPriorDate, "d") + "/" + Format(datPriorDate, "yyyy")
            Else
                PriorDay = ""
            End If
End Function

Public Sub SetRow(rec As ADODB.RecordSet, Optional blnInsert As Boolean = False)


' Fills all fields with data
    Set m_rec = rec
    m_blnInsert = blnInsert
    ' If we are inserting/cloning
    If Not m_rec.Fields("trade_skey") = "" Then
        m_blnRecFlag = True
    End If
    If m_blnInsert Then
        ' Do this so OriginalValue will be set to the values copied into the row
        m_rec.UpdateBatch
    End If
   show_detail
End Sub
Public Sub show_detail()
Dim ctl As Control

    For Each ctl In Me.Controls
        ctl.Visible = True
    Next ctl
    lblNewTradeGroup.Visible = False
    last_update_id.Visible = False
    loc_id.Visible = False
    TDBGrid.Visible = False
    lblTrades.Visible = False
    lblLaborRecords.Visible = False
    lblLocations.Visible = False

    ' Lock fields that can't be changed
    If Not blnAddMbr And Not blnNewGroup Then
        LockField "trade_id"
        LockField "city"
        LockField "state_code"
        chkDisplayHistory.Enabled = True
    End If
    If blnNewGroup Then
        Trade_Group_Code.Visible = False
        NewTradeGroupCode.Visible = True
        NewTradeGroupCode.Left = Trade_Group_Code.Left
    Else
        NewTradeGroupCode.Visible = False
    End If
End Sub

Private Function NewTradeLoc() As Boolean
Dim rsTemp As ADODB.RecordSet
Dim strUpdate As String
Dim strSelect As String
Dim blnReturn As Boolean

'Verify the existance of a labor rate record for the trade ID and location code.
On Error Resume Next
    strSelect = "select distinct labor_rate.trade_skey from labor_rate inner join labor_trade on labor_rate.trade_skey = labor_trade.trade_skey Where labor_rate.loc_id = " & City.ItemData(City.ListIndex) & " And labor_trade.trade_id = '" & Trade_ID.Text & "'"
    Set rsTemp = New ADODB.RecordSet
    rsTemp.Close
        ' Use DAL to perform select
        blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
        If blnReturn = False Then
            MsgBox "An error occurred while searching."
'            lblRowCount.Caption = "0 rows returned."
            GoTo Exit_Sub
        End If
    If rsTemp.RecordCount = 0 Then
        NewTradeLoc = True
    Else
        NewTradeLoc = False
    End If
    rsTemp.Close
    Set rsTemp = Nothing
Exit_Sub:

End Function

Private Sub chkDisplayHistory_Click()
Dim strSelect As String
Dim blnReturn As Boolean
'On Error Resume Next
If chkDisplayHistory = 1 Then
    TDBGrid.Visible = True
    strSelect = "exec sp_LaborRatesMaxStart @trade_id='" + SQLChangeWildcard(Trade_ID.Text) + "', @trade_group_code='"
    strSelect = strSelect + "', @city='"
    strSelect = strSelect + SQLChangeWildcard(Trim(City.Text)) + "', @state='"
    strSelect = strSelect + SQLChangeWildcard(Trim(State_Code.Text)) + "', @start_date=''"
    strSelect = strSelect + ", @term_date=''"
    strSelect = strSelect + ", @includehistory = 1"
    strSelect = strSelect + ", @maxrowcount = " + CStr(MAX_RECORDS)
    If m_rec2.State = adStateOpen Then
        m_rec2.Close ' Make sure it is closed
    End If
    m_rec2.MaxRecords = MAX_RECORDS ' Set the maximum number to bring back
    ' Use DAL to perform select
    blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, m_rec2)
    If blnReturn = False Then
        MsgBox "An error occurred while searching."
        GoTo Exit_Sub
    End If
    ' Pass recordset to handler class
    m_objGridMap.RecordSet = m_rec2
    
    ' If the upper bound was hit, inform user
    If m_rec2.RecordCount = MAX_RECORDS And m_rec2.State = adStateOpen Then
        MsgBox "The search returned the maximum number of records allowed. More records may be available."
    End If
    TDBGrid.ReBind
    TDBGrid.ApproxCount = m_rec2.RecordCount
    TDBGrid.FetchRowStyle = True
Else
    TDBGrid.Visible = False
End If

Exit_Sub:
Exit Sub
End Sub

Private Sub City_Click()
Dim rsTemp As ADODB.RecordSet
Dim strSelect As String
Dim blnReturn As Boolean
Dim strSelDate As String

'Verify the existance of a labor rate record for the trade ID and location code.
On Error Resume Next
    If City.ListIndex <> -1 Then
        If blnAddMbr = True Then
            If IsDate(start_date) Then
'                strSELECT = "select distinct lr.trade_skey, lr.last_update_id, lr.last_update_person,lr.last_update_date, lr.start_date, lr.term_date" + _
'                " from labor_rate as lr inner join labor_trade on lr.trade_skey = labor_trade.trade_skey " + _
'                " Where lr.loc_id = " & City.ItemData(City.ListIndex) & " And labor_trade.trade_id = '" + _
'                        trade_id.Text & "' and " + _
'" (convert(varchar(2),DATEPART(m, lr.term_date)) + '/' + convert(varchar(2),DATEPART(d, lr.term_date)) + '/' + convert(varchar(4),DATEPART(yy, lr.term_date)))  = '" + PriorDay(start_date) + _
'"' and lr.trade_group_code = ''"

'AK- 6/7/2006 - update to be consistent in Labor Rate retrieval

                strSelect = "select distinct lr.trade_skey, lr.last_update_id, lr.last_update_person,lr.last_update_date, lr.start_date, lr.term_date" + _
                " from labor_rate as lr inner join labor_trade on lr.trade_skey = labor_trade.trade_skey " + _
                " Where lr.loc_id = " & City.ItemData(City.ListIndex) & " And labor_trade.trade_id = '" + _
                        Trade_ID.Text + "'" + _
                " and lr.trade_group_code = '' and lr.start_date = (select max(start_date)" + _
                                            "From labor_rate " + _
                                            " where labor_rate.loc_id = lr.loc_id And labor_rate.trade_skey = lr.trade_skey)"
            Else
                strSelect = "select distinct lr.trade_skey, lr.last_update_id, lr.last_update_person,lr.last_update_date, lr.start_date, lr.term_date" + _
                " from labor_rate as lr inner join labor_trade on lr.trade_skey = labor_trade.trade_skey " + _
                " Where lr.loc_id = " & City.ItemData(City.ListIndex) & " And labor_trade.trade_id = '" + _
                        Trade_ID.Text + "'" + _
                " and lr.trade_group_code = '' and lr.start_date = (select max(start_date)" + _
                                            "From labor_rate " + _
                                            " where labor_rate.loc_id = lr.loc_id And labor_rate.trade_skey = lr.trade_skey)"
            End If
        Else
            strSelect = "select distinct lr.trade_skey, lr.last_update_id, lr.last_update_person,lr.last_update_date, lr.start_date, lr.start_date, lr.term_date" + _
            " from labor_rate as lr inner join labor_trade on lr.trade_skey = labor_trade.trade_skey " + _
            " Where lr.loc_id = " & City.ItemData(City.ListIndex) & " And labor_trade.trade_id = '" + _
                    Trade_ID.Text + _
            "' and lr.trade_group_code = '' and lr.start_date = (select max(start_date)" + _
                                        "From labor_rate " + _
                                        " where labor_rate.loc_id = lr.loc_id And labor_rate.trade_skey = lr.trade_skey)"
        End If
        Set rsTemp = New ADODB.RecordSet
        rsTemp.Close
        ' Use DAL to perform select
        blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
        If blnReturn = False Then
            MsgBox "An error occurred while searching."
            GoTo Exit_Sub
        End If
        If rsTemp.RecordCount = 0 Then
            'MsgBox "Error - no labor rate record found."
            chkDisplayHistory.Enabled = False
        Else
            If blnNewGroup Then
                start_date = rsTemp.Fields("start_date")
            End If
            trade_skey = rsTemp.Fields("trade_skey")
            last_update_id = rsTemp.Fields("last_update_id")
            last_update_date = rsTemp.Fields("last_update_date")
            last_update_person = rsTemp.Fields("last_update_person")
            term_date = rsTemp.Fields("term_date")
            chkDisplayHistory.Enabled = True
        End If
        rsTemp.Close
        Set rsTemp = Nothing
    Else
        chkDisplayHistory.Enabled = False
    End If
    chkDisplayHistory_Click
Exit_Sub:
End Sub

Private Sub City_GotFocus()
m_city = City.Text
End Sub



Private Sub cmdUpdate_Click()
    On Error Resume Next
    Dim strUpdate As String
    Dim strError As String
    Dim blnRet As Boolean
    Dim i As Integer
If Me.Caption = "Rename Trade Group" Then
    Dim strSelect As String
    Dim rsTemp As ADODB.RecordSet
    Dim varButton
'Remove the trade group code from the Trade_Group table and all labor rate records.

    On Error Resume Next
    If Trade_Group_Code.Text > " " Then
        strSelect = "select count(*) as RcdsToDelete from labor_rate where trade_group_code='" + Trade_Group_Code.Text + "'"
        blnRet = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
        If Not blnRet Then
            MsgBox "An error occurred retrieving data."
        Else
            If rsTemp![RcdsToDelete] > 0 Then
                Dim strMsg As String
                strMsg = CStr(rsTemp![RcdsToDelete]) + " Labor rate records will be removed/detached from the trade group " + Trade_Group_Code.Text + ".  Are you sure you want to remove them?"
                varButton = MsgBox(strMsg, vbYesNo + vbCritical)
            End If
        End If
        rsTemp.Close
        Set rsTemp = Nothing

     If varButton = vbYes Then
        Screen.MousePointer = vbHourglass
         strUpdate = "exec sp_replace_trade_groups "
         strUpdate = strUpdate + " @trade_group_code = '" + Trade_Group_Code.Text
         strUpdate = strUpdate + "', @new_trade_group_code = '" + NewTradeGroupCode.Text + "'"
         strUpdate = strUpdate + ", @last_update_person='" + strUserName + "'"
         blnRet = g_objDAL.ExecQuery(CONNECT, strUpdate, strError)
        If Len(strError) > 0 Then
             MsgBox strError
             m_blnWereErrors = True
         End If
     End If
    End If
Else
        m_blnWereErrors = False
            ' If we are updating
        'strUpdate = "exec sp_update_trade_grp_mbr "
        If blnAddMbr = True Then    'Adding a group member
            If City.ListIndex = -1 Then
                MsgBox "Please enter a valid state/city."
                m_blnWereErrors = True
            Else
                trade_skey = Trade_ID.ItemData(Trade_ID.ListIndex)
                loc_id = City.ItemData(City.ListIndex)
                strUpdate = "exec sp_update_trade_grp_mbr "
                strUpdate = strUpdate + " @trade_skey=" + CStr(trade_skey)
                strUpdate = strUpdate + ", @loc_id=" + CStr(loc_id)
                strUpdate = strUpdate + ", @prior_term_date='" + PriorDay(start_date) + "'"
            End If
        Else
            If blnNewGroup = True Then
                loc_id = City.ItemData(City.ListIndex)
            End If
        End If
        If m_blnWereErrors = False Then
            If Len(NewTradeGroupCode.Text) > 0 Then
               strUpdate = "exec sp_insert_trade_groups  "
               If Len(loc_id) = 0 Then
                    strUpdate = strUpdate + " @loc_id=0,"
                Else
                    strUpdate = strUpdate + " @loc_id=" + CStr(loc_id) + ","
                End If
               If Len(trade_skey) = 0 Then
                    strUpdate = strUpdate + " @trade_skey=0"
                Else
                    strUpdate = strUpdate + " @trade_skey=" + CStr(trade_skey)
                End If
                strUpdate = strUpdate + ", @start_date='" + start_date + "'"
                strUpdate = strUpdate + ", @new_trade_group_code='" + NewTradeGroupCode.Text + "'"
                strUpdate = strUpdate + ", @trade_group_code='" + Trade_Group_Code.Text + "'"
                strUpdate = strUpdate + ", @last_update_person='" + strUserName + "'"
                strUpdate = strUpdate + ", @last_update_id=" + CStr(last_update_id)
            Else
                strUpdate = strUpdate + ", @trade_group_code='" + Trade_Group_Code.Text + "'"
                If blnAddMbr = False Then
                    strUpdate = strUpdate + ", @start_date='" + start_date + "'"
                End If
                strUpdate = strUpdate + ", @last_update_person='" + strUserName + "'"
                strUpdate = strUpdate + ", @last_update_id=" + CStr(last_update_id)
             End If
            blnRet = g_objDAL.ExecQuery(CONNECT, strUpdate, strError)
            If Not blnRet Then
                MsgBox strError
                m_blnWereErrors = True
            Else
                ' Put latest data into source recordset
                If m_blnRecFlag = True Then
                    UpdateRecordsetFromForm Me, m_rec
                    m_rec.Fields("last_update_id").Value = m_rec.Fields("last_update_id").Value + 1
                    UpdateFormFromRecordset Me, m_rec
                End If
                ' final_answer
                MsgBox "Update successful."
            End If
        End If
    End If
    Screen.MousePointer = vbNormal
 
End Sub
Private Sub Form_Activate()
    OutputView False

End Sub

Private Sub Form_Initialize()
    m_blnRecFlag = False
    m_blnDeleted = False
    Status ("Loading Trade Groups...")
    Screen.MousePointer = vbHourglass
    DoEvents    'Paint screen
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim ctr As Control
    Dim rec As ADODB.RecordSet
    Move START_LEFT, START_TOP
    DoEvents    'Paint screen
    
   ' If we are showing data
    If m_blnRecFlag = True Then
        If Not m_rec.State = adStateClosed Then
            UpdateFormFromRecordset Me, m_rec
        End If
    End If
    
    LockField "start_date"
    LockField "term_date"
    LockField "trade_desc"
    LockField "trade_skey"
    LockField "last_update_id"
        ' Initialize grid
    m_objGridMap.SetGrid TDBGrid
    m_objGridMap.InitGrid
    

End Sub
Private Sub LoadCities(Optional strCity As String)
Dim strSelect As String
Dim rsTemp As RecordSet
Dim blnReturn As Boolean
Dim strSelDate As String
'Load Cities
    City.Clear
    If blnAddMbr = True Then
        If IsDate(start_date) Then
'            strSelect = "select distinct city, location.loc_id from labor_rate as lr inner join location on lr.loc_id = location.loc_id " + _
'            " where lr.trade_skey = " + CStr(trade_skey) + _
'" and (convert(varchar(2),DATEPART(m, lr.term_date)) + '/' + convert(varchar(2),DATEPART(d, lr.term_date)) + '/' + convert(varchar(4),DATEPART(yyyy, lr.term_date)))  = '" + PriorDay(start_date) + _
'"' and location.state_code = '" + State_Code.Text + _
'            "' and lr.trade_group_code = '' and lr.start_date = (select max(start_date) from labor_rate where labor_rate.trade_skey = lr.trade_skey and labor_rate.loc_id " + _
'            " = lr.loc_id) ORDER BY city"

'AK- 6/7/2006 - update to be consistent in City retrieval

            strSelect = "select distinct city, location.loc_id from labor_rate as lr inner join location on lr.loc_id = location.loc_id " + _
            " where lr.trade_skey = " + CStr(trade_skey) + _
            " and location.state_code = '" + State_Code.Text + _
            "' and lr.trade_group_code = '' and lr.start_date = (select max(start_date) from labor_rate where labor_rate.trade_skey = lr.trade_skey and labor_rate.loc_id " + _
            " = lr.loc_id) ORDER BY city"
        Else
            strSelect = "select distinct city, location.loc_id from labor_rate as lr inner join location on lr.loc_id = location.loc_id " + _
            " where lr.trade_skey = " + CStr(trade_skey) + _
            " and location.state_code = '" + State_Code.Text + _
            "' and lr.trade_group_code = '' and lr.start_date = (select max(start_date) from labor_rate where labor_rate.trade_skey = lr.trade_skey and labor_rate.loc_id " + _
            " = lr.loc_id) ORDER BY city"
        End If
    Else
            strSelect = "select distinct city, location.loc_id from labor_rate as lr inner join location on lr.loc_id = location.loc_id " + _
            " where lr.trade_skey = " + CStr(trade_skey) + _
            " and location.state_code = '" + State_Code.Text + _
            "' and lr.trade_group_code = '' and lr.start_date = (select max(start_date) from labor_rate where labor_rate.trade_skey = lr.trade_skey and labor_rate.loc_id " + _
            " = lr.loc_id) ORDER BY city"
        
'        If State_Code.Text > "" Then
'            strSelect = "select distinct city, loc_id from location where location.state_code = '" + State_Code.Text + "'  order by city"
'        Else
'            strSelect = "select distinct city, loc_id from location order by city"
'        End If
    End If
    ' Use DAL to perform select
    blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
    If blnReturn = False Then
        MsgBox "An error occurred loading Cities."
    Else
        If Not (rsTemp.EOF And rsTemp.BOF) Then
            Do Until rsTemp.EOF
                City.AddItem ConvertCase(rsTemp![City])
                City.ItemData(City.NewIndex) = rsTemp![loc_id]
                If City.Text > "" Then
                    If UCase(City.Text) = UCase(rsTemp![City]) Then
                        City.ListIndex = City.NewIndex
                    End If
                End If
                rsTemp.MoveNext
            Loop
            City.ListIndex = 0
        End If
    End If
    rsTemp.Close
End Sub
Private Function ConvertCase(strText As String) As String
Dim strTemp As String
Dim strTemp2 As String
Dim iStarta As Integer
Dim iStartb As Integer
If strText > " " Then
    strTemp = Left(strText, 1) + LCase(Right(strText, Len(strText) - 1))
    iStarta = InStr(1, strText, " ")
    If iStarta = 0 Then
        iStarta = InStr(1, strText, ",")
    End If
    If iStarta <> 0 Then
        While iStarta <> 0
            strTemp = Left(strTemp, Len(strTemp) - (Len(strTemp) - iStarta)) + UCase(Mid(strTemp, iStarta + 1, 1)) + Right(strTemp, Len(strTemp) - iStarta - 1)
            iStartb = InStr(iStarta + 1, strText, " ")
            If iStartb = 0 Then
                iStartb = InStr(iStarta, strText, ",")
            End If
            iStarta = iStartb
        Wend
        ConvertCase = strTemp
    Else
        ConvertCase = strTemp
    End If
Else
    ConvertCase = ""
End If
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim Button
    Dim blnPendingChange As Boolean
    
    ' Only go through this if the close wasn't invoked from code
    If Not UnloadMode = vbFormCode Then
        blnPendingChange = IsControlChanged(Me, m_rec)
    
        If blnPendingChange = True Then
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
    End If
    If Cancel = False Then
        Me.Visible = False
    End If
End Sub

Private Sub Form_Resize()
ResizeForm Me
End Sub

Private Sub NewTradeGroupCode_Change()
Dim iSelStart As Integer
Dim iSelLen As Integer

    If Len(NewTradeGroupCode) > 7 Then
        NewTradeGroupCode = Left(NewTradeGroupCode, 7)
        NewTradeGroupCode.SelStart = 7
    End If

If Len(NewTradeGroupCode) = 4 Then
    Dim strSelect As String
    Dim blnRet As Boolean
    Dim strError As String
    Dim rsTemp As New ADODB.RecordSet
    strSelect = "exec sp_NextTradeGroup @prefix ='" + UCase(NewTradeGroupCode.Text) + "%'"
    blnRet = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
    If Not blnRet Then
        MsgBox "An error occurred retrieving data."
    Else
        If rsTemp.RecordCount > 0 Then
            NewTradeGroupCode = rsTemp.Fields("newcode")
            NewTradeGroupCode.SelStart = 4
            NewTradeGroupCode.SelLength = 3
        End If
    End If
    rsTemp.Close
    Set rsTemp = Nothing
Else
    iSelStart = NewTradeGroupCode.SelStart
    iSelLen = NewTradeGroupCode.SelLength
    NewTradeGroupCode = UCase(NewTradeGroupCode)
    NewTradeGroupCode.SelStart = iSelStart
    NewTradeGroupCode.SelLength = iSelLen
End If
    
End Sub

Private Sub NewTradeGroupCode_Validate(Cancel As Boolean)
    Dim strSelect As String
    Dim blnRet As Boolean
    Dim strError As String
    Dim rsTemp As New ADODB.RecordSet
    If Len(NewTradeGroupCode.Text) > 0 Then
        strSelect = "select trade_skey from labor_rate where trade_group_code = '" + NewTradeGroupCode.Text + "'"
        blnRet = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
        If Not blnRet Then
            MsgBox "An error occurred retrieving data."
        Else
            If rsTemp.RecordCount > 0 Then
                MsgBox "This trade group code is already in use."
                Cancel = True
            End If
        End If
        rsTemp.Close
        Set rsTemp = Nothing
    End If
End Sub

Private Sub State_Code_Click()
If m_State <> State_Code.Text Then
    LoadCities
    m_State = State_Code.Text
End If
End Sub

Private Sub State_Code_GotFocus()
m_State = State_Code.Text
End Sub

Private Sub Trade_Group_Code_GotFocus()
    m_trade_group_code = Trade_Group_Code.Text

End Sub


Private Sub Trade_Group_Code_LostFocus()
If Trade_Group_Code.Text <> m_trade_group_code Then
    get_counts
    m_trade_group_code = Trade_Group_Code.Text
End If

End Sub

Private Sub Trade_ID_Click()
Dim strSelect As String
Dim rsTemp As RecordSet
Dim blnReturn As Boolean
Dim i As Integer
If Trade_ID.Text <> m_trade_id Then
    strSelect = "select trade_desc from LABOR_TRADE where LABOR_TRADE.trade_id = '" + Trade_ID.Text + "'"
    ' Use DAL to perform select
    blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
    If blnReturn = False Then
        MsgBox "An error occurred loading trade information."
    Else
        If Not (rsTemp.EOF And rsTemp.BOF) Then
            trade_desc.Text = ConvertCase(rsTemp![trade_desc])
        End If
        For i = 0 To Trade_ID.listcount - 1
            If UCase(Trim(Trade_ID.Text)) = UCase(Trim(Trade_ID.List(i))) Then Exit For
        Next i
        trade_skey.Text = CStr(Trade_ID.ItemData(i))
    End If
    rsTemp.Close
    Set rsTemp = Nothing
    LoadStates
    m_trade_id = Trade_ID.Text
End If

End Sub

Private Sub Trade_ID_GotFocus()
m_trade_id = Trade_ID.Text
End Sub

