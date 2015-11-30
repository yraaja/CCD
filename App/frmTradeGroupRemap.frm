VERSION 5.00
Begin VB.Form frmTradeGroupRemap 
   Caption         =   "Trade Group Remapping"
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11085
   LinkTopic       =   "Form1"
   ScaleHeight     =   8610
   ScaleWidth      =   11085
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option2 
      Caption         =   "Off"
      Height          =   255
      Left            =   5760
      TabIndex        =   43
      Top             =   2880
      Width           =   855
   End
   Begin VB.Frame frmWarnings 
      Caption         =   "Warnings Controls"
      Height          =   615
      Left            =   4680
      TabIndex        =   41
      Top             =   2640
      Width           =   2055
      Begin VB.OptionButton Option1 
         Caption         =   "On"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   495
      Left            =   9120
      TabIndex        =   40
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   495
      Left            =   7560
      TabIndex        =   39
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save to Labor Rates"
      Height          =   495
      Left            =   4425
      TabIndex        =   35
      Top             =   7920
      Width           =   2235
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add -->"
      Height          =   375
      Left            =   4080
      TabIndex        =   34
      Top             =   120
      Width           =   735
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   5055
      TabIndex        =   33
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdShiftLeft 
      Caption         =   "UP"
      Height          =   495
      Left            =   6195
      TabIndex        =   32
      ToolTipText     =   "Shift selected records from the bottom listbox to the top listbox"
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton cmdShiftAllLeft 
      Caption         =   "ALL UP"
      Height          =   495
      Left            =   5595
      TabIndex        =   31
      ToolTipText     =   "Shift all records from the bottom listbox to the top listbox"
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton cmdShiftAllRight 
      Caption         =   "ALL DN"
      Height          =   495
      Left            =   4995
      TabIndex        =   30
      ToolTipText     =   "Shift ALL records from top listbox to the bottom listbox"
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton cmdShiftRight 
      Caption         =   "DN"
      Height          =   495
      Left            =   4395
      TabIndex        =   29
      ToolTipText     =   "Shift Selected Records from the top listbox to the bottom listbox"
      Top             =   5280
      Width           =   495
   End
   Begin VB.ListBox ListBox2 
      Height          =   1815
      Left            =   360
      TabIndex        =   28
      Top             =   5880
      Width           =   10215
   End
   Begin VB.ListBox ListBox1 
      Height          =   1815
      Left            =   360
      MultiSelect     =   2  'Extended
      TabIndex        =   27
      Top             =   3360
      Width           =   10215
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
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   6540
      Visible         =   0   'False
      Width           =   1635
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
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   6540
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
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Tag             =   "1G"
      Top             =   6540
      Visible         =   0   'False
      Width           =   1215
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
      Left            =   5676
      Locked          =   -1  'True
      TabIndex        =   10
      Tag             =   "1N"
      Top             =   7920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox City 
      Height          =   315
      ItemData        =   "frmTradeGroupRemap.frx":0000
      Left            =   4260
      List            =   "frmTradeGroupRemap.frx":0002
      TabIndex        =   9
      Text            =   "City"
      Top             =   1530
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox term_date 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5535
      TabIndex        =   8
      Tag             =   "1D"
      Top             =   1950
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.TextBox start_date 
      Height          =   285
      Left            =   2655
      TabIndex        =   7
      Tag             =   "1D"
      Top             =   1950
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.TextBox NewTradeGroupCode 
      Height          =   285
      Left            =   8175
      TabIndex        =   6
      Tag             =   "1S"
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox trade_desc 
      Height          =   285
      Left            =   4260
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1140
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.ComboBox Trade_ID 
      Height          =   315
      Left            =   2655
      TabIndex        =   4
      Tag             =   "0"
      Top             =   1110
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.ComboBox State_Code 
      Height          =   315
      Left            =   2655
      TabIndex        =   3
      Top             =   1530
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.ComboBox Trade_Group_Code 
      Height          =   315
      Left            =   2400
      TabIndex        =   2
      Top             =   120
      Width           =   1485
   End
   Begin VB.CheckBox chkDisplayHistory 
      Caption         =   "Display Trade/Location History"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6855
      TabIndex        =   1
      Top             =   2040
      Visible         =   0   'False
      Width           =   2655
   End
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
      Left            =   6996
      Locked          =   -1  'True
      TabIndex        =   0
      Tag             =   "1N"
      Top             =   7920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblLaborRecords 
      Alignment       =   2  'Center
      Caption         =   "Labor Rate Records"
      Height          =   255
      Left            =   7095
      TabIndex        =   15
      Top             =   720
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblLocations 
      Alignment       =   2  'Center
      Caption         =   "Locations"
      Height          =   255
      Left            =   4575
      TabIndex        =   16
      Top             =   720
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblTrades 
      Alignment       =   2  'Center
      Caption         =   "Trades"
      Height          =   255
      Left            =   2280
      TabIndex        =   14
      Top             =   720
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblRowCount 
      Height          =   375
      Left            =   2640
      TabIndex        =   38
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lblTargetRows 
      Caption         =   "TARGET LABOR ROWS"
      Height          =   255
      Left            =   360
      TabIndex        =   37
      Top             =   5520
      Width           =   2175
   End
   Begin VB.Label lblSource 
      Caption         =   "SOURCE LABOR ROWS"
      Height          =   375
      Left            =   360
      TabIndex        =   36
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Last Updated:"
      Height          =   255
      Left            =   660
      TabIndex        =   26
      Top             =   6600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Last Updated By:"
      Height          =   255
      Left            =   3420
      TabIndex        =   25
      Top             =   6600
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Term Date:"
      Height          =   255
      Left            =   4605
      TabIndex        =   24
      Top             =   1980
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Start Date:"
      Height          =   255
      Left            =   1575
      TabIndex        =   23
      Top             =   1980
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Trade Group:"
      Height          =   255
      Left            =   1080
      TabIndex        =   22
      Top             =   120
      Width           =   1155
   End
   Begin VB.Label Label20 
      Caption         =   "Trade_Skey"
      Height          =   255
      Left            =   6240
      TabIndex        =   21
      Top             =   6600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblNewTradeGroup 
      Alignment       =   1  'Right Justify
      Caption         =   "New Trade Group:"
      Height          =   255
      Left            =   6375
      TabIndex        =   20
      Top             =   150
      Width           =   1635
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Trade ID:"
      Height          =   255
      Left            =   1635
      TabIndex        =   19
      Top             =   1170
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "State:"
      Height          =   255
      Left            =   1815
      TabIndex        =   18
      Top             =   1590
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "City:"
      Height          =   255
      Left            =   3765
      TabIndex        =   17
      Top             =   1590
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "frmTradeGroupRemap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim m_objGridMap As New CTradeHistMap ' Class to handle grid
Dim m_objGridMap As New CLaborRateMap ' Class to handle grid
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
Dim SelectCtrl As clsSelectCtrl     'rlh 03/25/2009  CCD 8.4+
Dim strOut As String                'rlh 04/01/2009  CCD 8.4+
Dim rsTemp As ADODB.RecordSet
Dim rsClone As ADODB.RecordSet
Dim TradeID As String
Dim StateCode As String
Dim WARNINGS_ON As Boolean          'rlh 04/07/2010
'Dim City As String
Dim updateCnt As Integer             'rlh 04/27/2010
Dim successCnt As Integer            'rlh 04/27/2010

Dim cmdTemp As ADODB.Command





Public Function get_counts() As Integer
    Dim strSelect As String
    Dim blnRet As Boolean
    Dim strError As String
    Dim rsTemp As New ADODB.RecordSet
    Dim lOrigPointer As Long
    
    On Error GoTo ERRLBL
    
If Trade_Group_Code.Text <> "" Then
'Count trade IDs for group
    lOrigPointer = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    strSelect = "exec usp_count_trades @trade_group_code = '" + Trade_Group_Code.Text + "'"
    blnRet = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
    If Not blnRet Then
        MsgBox "An error occurred retrieving data."
    Else
      If rsTemp(0) > 0 Then
        If rsTemp.RecordCount > 0 Then
            lblTrades = CStr(rsTemp.Fields("Trades")) + " Trade(s)"
            lblTrades.Visible = True
            lblLocations = CStr(rsTemp.Fields("locations")) + " Location(s)"
            lblLocations.Visible = True
            lblLaborRecords = CStr(rsTemp.Fields("labor_records")) + " Labor Rate Record(s)"
            lblLaborRecords.Visible = True
            start_date = rsTemp.Fields("maxstartdate")
            
            get_counts = rsTemp(0)
        End If
        Else
            get_counts = rsTemp(0)
            Screen.MousePointer = vbNormal
            Exit Function
      End If
    End If
    rsTemp.Close
    Set rsTemp = Nothing

    If blnAddMbr = True Then
        Load_Trade_IDs
    End If
    Screen.MousePointer = vbNormal
    Exit Function
ERRLBL:
    MsgBox ("(Error)get_counts: " & Err.Description)
    
End If
get_counts = 1
End Function

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
    blnAddMbr = True
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
'''            Dim INString As String
'''            Dim i As Integer
'''            INString = ""
'''            For i = 0 To List1.listcount
'''                If i = List1.listcount Then
'''                    INString = INString & "'" & List1.List(i) & "'"
'''                Else
'''                    INString = INString & "'" & List1.List(i) & "',"
'''                End If
'''            Next
'''            If List1.listcount > 0 Then
'''                'rlh 03/25/2009   CCD 8.4+  ----------------------------
'''                strSELECT = "SELECT DISTINCT lr.trade_group_code, lt.trade_id, loc.city, loc.state_code   FROM LABOR_RATE lr, LOCATION loc, LABOR_TRADE lt"
'''                strSELECT = strSELECT & " WHERE "
'''                strSELECT = strSELECT & " Loc.loc_id = lr.loc_id"
'''                strSELECT = strSELECT & " AND"
'''                strSELECT = strSELECT & " lt.trade_skey = lr.trade_skey"
'''                strSELECT = strSELECT & " AND"
'''                strSELECT = strSELECT & " trade_group_code IN(" & INString & ")"
'''                strSELECT = strSELECT & " AND trade_group_code <> '       '"
'''                strSELECT = strSELECT & " ORDER BY trade_group_code"
'''            Else
'''                strSELECT = "SELECT DISTINCT lr.trade_group_code, lt.trade_id, loc.city, loc.state_code   FROM LABOR_RATE lr, LOCATION loc, LABOR_TRADE lt"
'''                strSELECT = strSELECT & " WHERE "
'''                strSELECT = strSELECT & " Loc.loc_id = lr.loc_id"
'''                strSELECT = strSELECT & " AND"
'''                strSELECT = strSELECT & " lt.trade_skey = lr.trade_skey"
'''                strSELECT = strSELECT & " AND trade_group_code <> '       '"
''''                strSELECT = strSELECT & " AND"
''''                strSELECT = strSELECT & " trade_group_code IN(" & INString & ")"
'''                strSELECT = strSELECT & " ORDER BY trade_group_code"
'''
'''            End If
'''
'''            blnReturn = g_objDAL.GetRecordset(CONNECT, strSELECT, rsTemp)
'''            If blnReturn = False Then
'''                MsgBox "An error occurred loading Trade Groups."
'''            Else
'''                If Not (rsTemp.EOF And rsTemp.BOF) Then
'''                    Call SelectCtrl.AddAllTradeGroupsToOneListBox(1, Me.ListBox1, rsTemp)
'''                End If
'''            End If


'            If (NewTradeGroupCode <> "") Then
'                Call Me.PopMasterTradeGroupList
'            End If
            
            'rlh 03/25/2009 -------------------------------------------
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
'''    If Trade_ID.ListIndex <> -1 Then
'''        If blnAddMbr Then
'''            If IsDate(start_date) Then
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
                
            strSelect = "SELECT distinct location.state_code FROM labor_rate as lr "
            strSelect = strSelect & "inner join location on lr.loc_id = location.loc_id "
            strSelect = strSelect & " ORDER BY  location.state_code"
'''            Else
'''                strSELECT = "SELECT distinct location.state_code FROM labor_rate as lr " + _
'''                "inner join location on lr.loc_id = location.loc_id " + _
'''                 "where lr.trade_skey = " + CStr(Trade_ID.ItemData(Trade_ID.ListIndex)) + _
'''                 " and lr.trade_group_code = '' and lr.start_date = (select max(start_date) from labor_rate where labor_rate.trade_skey = lr.trade_skey and labor_rate.loc_id " + _
'''                " = lr.loc_id) ORDER BY  location.state_code"
'''            End If
'''        Else
'''            strSELECT = "SELECT distinct location.state_code FROM labor_rate as lr " + _
'''            "inner join location on lr.loc_id = location.loc_id " + _
'''             "where lr.trade_skey = " + CStr(Trade_ID.ItemData(Trade_ID.ListIndex)) + _
'''            " and lr.trade_group_code = '' and lr.start_date = (select max(start_date) from labor_rate where labor_rate.trade_skey = lr.trade_skey and labor_rate.loc_id " + _
'''            " = lr.loc_id) ORDER BY  location.state_code"
'''        End If
        
        ' Use DAL to perform select
        blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
        If blnReturn = False Then
            MsgBox "An error occurred loading States."
        Else
            State_Code.AddItem "  "   'rlh top spot should be a blank
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
'''    End If
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
'    TDBGrid.Visible = False
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
    blnNewGroup = False
    If blnNewGroup Then
        Trade_Group_Code.Visible = False
        NewTradeGroupCode.Visible = True
        NewTradeGroupCode.Left = Trade_Group_Code.Left
    Else
        NewTradeGroupCode.Visible = False
    End If
    
    lblNewTradeGroup.Visible = True         'rlh 03/25/2009 CCD 8.4+
    NewTradeGroupCode.Visible = True        'rlh 03/25/2009 CCD 8.4+
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
'    TDBGrid.Visible = True
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
'    TDBGrid.ReBind
'    TDBGrid.ApproxCount = m_rec2.RecordCount
'    TDBGrid.FetchRowStyle = True
Else
'    TDBGrid.Visible = False
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
    If (Trim(City.Text) <> "") Then
        Call PopMasterTradeGroupList               'rlh - Apply filter
    End If
Exit_Sub:
End Sub

Private Sub City_GotFocus()
m_city = City.Text
City = Trim(m_city)       'rlh - for filtering
End Sub



Private Sub cmdAdd_Click()
Me.List1.AddItem Me.Trade_Group_Code
'Me.ListBox1.Clear
Call Me.PopMasterTradeGroupList
'If (rsTemp.EOF) Then MsgBox ("No ACTIVE labor rates found for trade group: " & Me.Trade_Group_Code)
End Sub

Private Sub cmdApply_Click()
'''    Dim blnReturn As Boolean
'''    Dim i As Integer
'''
'''    If Me.ListBox2.listcount = 0 Then
'''        If Me.List1.listcount = 0 Then
'''            MsgBox ("(ERROR)Update: Please select tradegroups to be mapped.  Thank you")
'''            Me.Trade_Group_Code.SetFocus
'''            Exit Sub
'''        End If
'''    End If
'''
'''    If Me.NewTradeGroupCode.Text <> "" Then
'''
'''    Else
'''        MsgBox ("(Error)Update: Please specify a *NEW* trade group.  Thank you.")
'''        Me.NewTradeGroupCode.SetFocus
'''        Exit Sub
'''    End If
'''
'''    'BUILD THE TRADE GROUP STRING delimited by commas
'''
'''    For i = 0 To Me.List1.listcount - 1
'''        If i = Me.List1.listcount - 1 Then
'''            strOut = strOut & Me.List1.List(i)
'''        Else
'''            strOut = strOut & Me.List1.List(i) & ","
'''        End If
'''    Next
'''
'''    'Call routine to run "remap" stored procedure
'''    blnReturn = Apply(strOut)
End Sub

Private Sub cmdRestore_Click()
''Call Me.Restore
End Sub

Private Sub cmdRefresh_Click()
Me.Trade_Group_Code = ""
Me.Trade_ID = ""
Me.City = ""
Me.State_Code = ""

'Me.List1.Clear
Me.ListBox1.Clear

Screen.MousePointer = vbHourglass
'Call Me.PopMasterTradeGroupList
Call Me.RefreshMasterTradeGroupList
Screen.MousePointer = vbNormal
End Sub

Private Sub cmdSave_Click()
Dim ary() As String
Dim i As Integer
Dim j As Integer
Dim tmpstr As String
Dim ans As Variant
Dim strHdr As String
'Arrays to store key fields for checking before saving!!!
Dim chkAryStdt() As String
Dim chkAryTdt() As String
Dim chkAryUFring() As String
Dim chkAryUBase() As String
Dim chkAryTotU() As String

Dim compareAry() As structTradeGroup
Dim deleteSql As String
Dim updateSql As String

Dim loc_id As Integer
Dim strSelect As String
Dim blnReturn As Boolean

Dim trade_skey As Integer
Dim start_date As Date
Dim term_date As Date


On Error GoTo ERRLBL

If WARNINGS_ON Then

    ans = MsgBox("Warnings are on.  Do you wish to keep them on?", vbYesNo, "Warnings Prompt")
    Select Case ans
    Case vbYes
        WARNINGS_ON = True
        Me.Option1 = True
        Me.Option2 = False
    Case vbNo
        WARNINGS_ON = False
        Me.Option1 = False
        Me.Option2 = True
    End Select
Else
     ans = MsgBox("Warnings are off.  Do you wish to keep them off?", vbYesNo, "Warnings Prompt")
    Select Case ans
    Case vbYes
        WARNINGS_ON = False
        Me.Option1 = False
        Me.Option2 = True
    Case vbNo
        WARNINGS_ON = True
        Me.Option1 = True
        Me.Option2 = False
    End Select

End If

If (Len(Trim(Me.NewTradeGroupCode)) = 0) Then

    MsgBox ("Please specify a new Trade Group Code Name.  Thank you")
    Me.NewTradeGroupCode.SetFocus
    Exit Sub
End If

With g_cnShared
           .BeginTrans
           Set cmdTemp = New ADODB.Command
           Set cmdTemp.ActiveConnection = g_cnShared

           With cmdTemp
               .CommandTimeout = 0
               .CommandType = adCmdText
               .CommandText = updateSql
'''            .Execute 'adExecuteNoRecords
           End With

End With



':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'DO CHECKS ACROSS ALL "SELECTED" LR RECORDS
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
ReDim Preserve compareAry(1)
Dim MAX_START_DATE As Date
For i = 0 To ListBox2.listcount - 1
    ary = Split(ListBox2.List(i), vbTab)
    
    'ReDim Preserve compareAry(UBound(compareAry))
    
    compareAry(i).Trade_Group_Code = ary(0)
    compareAry(i).Trade_ID = ary(1)
    compareAry(i).City = ary(2)
    compareAry(i).State_Code = ary(4)
    compareAry(i).start_date = ary(5)
    compareAry(i).term_date = ary(6)
    compareAry(i).union_base = ary(7)
    compareAry(i).union_fring = ary(8)
    compareAry(i).tot_union = ary(9)
    
    If CDate(ary(5)) > CDate(MAX_START_DATE) Then MAX_START_DATE = CDate(ary(5))
    
    ReDim Preserve compareAry(UBound(compareAry) + 1)
Next

'CHECK FOR CONSISTENCY (ALL start dates = MAX START DATE, term_dates, base, fringe, total_union
' are the same!!!

'MAX START DATES
For i = 0 To UBound(compareAry) - 2

 If (compareAry(i).start_date <> MAX_START_DATE) Then
    ListBox2.Selected(i) = True
    
    If WARNINGS_ON Then
        ans = MsgBox("Found a start date at row #," & i & " that does not match the MAX start date for the group: " & " MAX START DATE: " _
        & MAX_START_DATE & vbCrLf & "Non-matching start date: " & compareAry(i).start_date, vbOKCancel, "Continue?")
    Else
        ans = vbOK
    End If
    
    ListBox2.Selected(i) = True
    
    Select Case ans
    Case vbOK
    Case vbCancel
        Exit Sub   'comment out until testing is complete
    End Select
 End If
Next

'TERM DATE
Dim last_term_date As Date
last_term_date = compareAry(0).term_date    'seed the last term date
For i = 0 To UBound(compareAry) - 2

 If (compareAry(i).term_date <> last_term_date) Then
    
    ListBox2.Selected(i) = True
    If WARNINGS_ON Then
    ans = MsgBox("Found a term date at row #," & i & " that does not match the other term dates for the group: " & _
     vbCrLf & "Non-matching term date: " & compareAry(i).term_date, vbOKCancel, "Continue?")
    Else
     ans = vbOK
    End If
    Select Case ans
    Case vbOK
    Case vbCancel
        Exit Sub   'comment out until testing is complete
    End Select
 End If
 Next
 
 'UNION BASE
Dim last_union_base As Date
last_union_base = compareAry(0).union_base    'seed the last UNION BASE
For i = 0 To UBound(compareAry) - 2

 If (compareAry(i).union_base <> last_union_base) Then
    
    ListBox2.Selected(i) = True
    If WARNINGS_ON Then
    ans = MsgBox("Found a Union Base Hourly  at row #," & i & " that does not match the other union base hrly rate for the group: " & _
     vbCrLf & "Non-matching Union base hrly: " & compareAry(i).union_base, vbOKCancel, "Continue?")
    Else
        ans = vbOK
    End If
    Select Case ans
    Case vbOK
    Case vbCancel
        Exit Sub   'comment out until testing is complete
    End Select
 End If
Next


'UNION FRINGE
Dim last_union_fringe As Date
last_union_fringe = compareAry(0).union_fring    'seed the last UNION BASE
For i = 0 To UBound(compareAry) - 2

 If (compareAry(i).union_fring <> last_union_fringe) Then
 
    ListBox2.Selected(i) = True
    If WARNINGS_ON Then
    ans = MsgBox("Found a Union Fringe Hourly  at row #," & i & " that does not match the other union fringe hrly rate for the group: " & _
     vbCrLf & "Non-matching Union fringe hrly: " & compareAry(i).union_fring, vbOKCancel, "Continue?")
    Else
        ans = vbOK
    End If
    
    Select Case ans
    Case vbOK
    Case vbCancel
        Exit Sub   'comment out until testing is complete
    End Select
 End If
Next

'TOTAL UNION
Dim last_tot_union As Date
last_tot_union = compareAry(0).tot_union    'seed the last TOTAL UNION
For i = 0 To UBound(compareAry) - 2

 If (compareAry(i).tot_union <> last_tot_union) Then
 
    ListBox2.Selected(i) = True
    If WARNINGS_ON Then
    ans = MsgBox("Found a Total Union  at row #," & i & " that does not match the other total union rates for the group: " & _
     vbCrLf & "Non-matching Total Union: " & compareAry(i).tot_union, vbOKCancel, "Continue?")
    Else
        ans = vbOK
    End If
    Select Case ans
    Case vbOK
    Case vbCancel
        Exit Sub   'comment out until testing is complete
    End Select
 End If
Next


strHdr = "Trade Group values to be saved: "
For i = 0 To ListBox2.listcount - 1
    ary = Split(ListBox2.List(i), vbTab)
   
    tmpstr = tmpstr & vbCrLf
    For j = 0 To UBound(ary)
      tmpstr = tmpstr & ary(j) & " " & vbTab
    Next

        '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        'BUILD THE INSERT statement to add to labor_rate_super TABLE
        '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Call BuildAndInsertToDb(ary)


Next


':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'   GET RID OF ALL TRACES OF trade groups that were mapped into the SUPER GROUP
'   * * * AFTER ALL UPDATES HAVE BEEN DONE !!!  * * *
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
updateCnt = 0

For i = 0 To ListBox2.listcount - 1
    ary = Split(ListBox2.List(i), vbTab)

    start_date = ary(5)
    term_date = ary(6)


    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    ' GET LOC_ID  (from City and State Code combination)
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

    strSelect = "select loc_id from LOCATION where City='" & Trim(ary(2)) & "' AND state_code='" & Trim(ary(4)) & "'"

    blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
    If Not rsTemp.EOF Then
        loc_id = rsTemp(0)
    Else
        MsgBox ("(Error)No LOC_ID found for city," & ary(2) & ", state, " & ary(4) & " Terminating error")
        Exit Sub
    End If

    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::
    'GET trade_skey (from City and State Code combination)
    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::

    strSelect = "select trade_skey from LABOR_TRADE where Trade_id='" & Trim(ary(1)) & "'"

    blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
    If Not rsTemp.EOF Then
        trade_skey = rsTemp(0)
    Else
        MsgBox ("(Error)No trade_skey found for trade_id," & ary(1) & ", state, " & ary(3) & " Terminating error")
        Exit Sub
    End If


    'CAN'T WORK!!!  THE TRADE_GROUP_CODE HAS ALREADY BEEN CHANGED TO "NEW----" (FOR EXAMPLE) SO THESE
    '               UPDATE STATEMENTS CAN NEVER WORK!!!
    
    updateSql = "UPDATE LABOR_RATE SET trade_group_code='" & Trim(Me.NewTradeGroupCode.Text) & "', "
    updateSql = updateSql & " last_update_date='" & Now & "',"
    updateSql = updateSql & " last_update_person='" & strUserName & "'"
    updateSql = updateSql & " WHERE trade_group_code='" & Trim(ary(0)) & "'"
    updateSql = updateSql & " AND trade_skey=" & trade_skey
    updateSql = updateSql & " AND loc_id=" & loc_id
    
    Dim tmpMssg As String
    tmpMssg = "HERE IS THE SQL UPDATE THAT WILL BLANK THE FOLLOWING TRADE GROUP: " & Trim(ary(0)) & vbCrLf & vbCrLf & "SQL:  (FOR DEBUG PURPOSES ONLY!) " & vbCrLf & vbCrLf & updateSql & vbCrLf & vbCrLf & " ALL TRACES OF " & ary(0) & "for trade_skey: " & trade_skey & " WILL BE REMOVED FROM SIGHT FOREVER"
    
'    ans = MsgBox("TEMPORARY: PLEASE SCRUTINIZE RESULTS BEFORE ANSWERING YES/NO " & vbCrLf & vbCrLf & tmpMssg, vbYesNo, "TEMPORARY CHECK ON UPDATE")
'    Select Case (ans)
'        Case vbYes
'        Case vbNo
'            g_cnShared.RollbackTrans
'            MsgBox ("Processing has been cancelled")
'            Exit Sub
'        Case Else
'
'
'    End Select

        
'''    updateSql = updateSql & " AND month(start_date)=" & Month(start_date)
'''    updateSql = updateSql & " AND day(start_date)=" & Day(start_date)
'''    updateSql = updateSql & " AND year(start_date)=" & Year(start_date)
'''    updateSql = updateSql & " AND month(term_date)=" & Month(term_date)
'''    updateSql = updateSql & " AND day(term_date)=" & Day(term_date)
'''    updateSql = updateSql & " AND year(term_date)=" & Year(term_date)


        '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        'BLANK OUT ALL TRADE GROUP CODES BELONGING TO THE "SUPER" TRADE GROUP AND TRADE ID
        '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

         With g_cnShared
                    
                    Set cmdTemp = New ADODB.Command
                    Set cmdTemp.ActiveConnection = g_cnShared

                    With cmdTemp
                        .CommandTimeout = 0
                        .CommandType = adCmdText
                        .CommandText = updateSql
                        .Execute 'adExecuteNoRecords
                    End With

                    If .Errors.Count <> 0 Then
                        .RollbackTrans
                        MsgBox "Errors in the Remap Update routine. " _
                            & vbCrLf & g_cnShared.Errors(0).Description, vbCritical
                        'errCnt = errCnt + 1
                        Exit For
                        
                    Else
                       
                        updateCnt = updateCnt + 1
'                        MsgBox ("The following groups were remapped to group: " & Me.NewTradeGroupCode.Text & _
'                               vbCrLf & strOut)
                    End If
    End With


Next


If updateCnt = 0 Then
        '#############################################
        'PROBLEMS... - ROLLBACK TRANSACTION
        '#############################################
        
        g_cnShared.RollbackTrans
        MsgBox ("Errors were encountered during Mapping cleanup - after 1st stage of remapping")
Else
        '#############################################
        'EVERYTHING WAS OK - COMMIT TRANSACTION
        '#############################################
        
        g_cnShared.CommitTrans
        MsgBox ("* * * PROCESSING IS COMPLETE * * * ")
             
End If
    


Exit Sub
ERRLBL:
    MsgBox ("(Error)cmdSave_Click: " & Err.Description)
    Resume
End Sub

Private Function BuildAndInsertToDb(ary() As String) As Integer
Dim i As Integer
Dim tmpstr As String
Dim insertSql As String
Dim strSelect As String
Dim blnReturn As Boolean
Dim rsTemp As ADODB.RecordSet
'Dim loc_id As String
Dim loc_id As Integer
Dim super_trade_code As String
Dim trade_code As String
Dim Trade_ID As String
Dim City As String
Dim State_Code As String
Dim cnt As Integer
Dim savary(9) As String
Dim strUpdate As String



On Error GoTo ERRLBL

insertSql = "INSERT INTO LABOR_TRADE_SUPER("
insertSql = insertSql & "SUPER_TRADE_GROUP_CODE,"
insertSql = insertSql & "TRADE_GROUP_CODE,"
insertSql = insertSql & "TRADE_ID,"
insertSql = insertSql & "CITY,"
insertSql = insertSql & "STATE_CODE,"

For i = 0 To UBound(ary)
    If ary(i) <> "" Then
        savary(cnt) = ary(i)
        cnt = cnt + 1
    End If
Next


'GET LOC_ID  (from City and State Code combination)

strSelect = "select loc_id from LOCATION where City='" & Trim(savary(2)) & "' AND state_code='" & Trim(savary(3)) & "'"
blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
If Not rsTemp.EOF Then
    loc_id = rsTemp(0)
Else
    MsgBox ("(Error)No LOC_ID found for city," & savary(2) & ", state, " & savary(3) & " Terminating error")
    Exit Function
End If


insertSql = insertSql & "LOC_ID"
insertSql = insertSql & ") "
insertSql = insertSql & "VALUES('"
insertSql = insertSql & Me.NewTradeGroupCode.Text
For i = 0 To UBound(savary) - 1
    insertSql = insertSql & "','" & savary(i)
Next
insertSql = insertSql & "'," & loc_id
insertSql = insertSql & ")"


'::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::
'::     WRITE DIRECTLY TO "LABOR_RATES" TABLES
'::
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::

        Dim trade_skey As Integer
        'Dim loc_id As Integer
        Dim start_date As Date
        Dim term_date As Date
        Dim Trade_Group_Code As String
        
        ':::::::::::::::::::::::::::::::::::::::::::::::::::::::
        'GET trade_skey (from City and State Code combination)
        ':::::::::::::::::::::::::::::::::::::::::::::::::::::::

        strSelect = "select trade_skey from LABOR_TRADE where Trade_id='" & Trim(savary(1)) & "'"
        
        blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
        If Not rsTemp.EOF Then
            trade_skey = rsTemp(0)
        Else
            MsgBox ("(Error)No trade_skey found for trade_id," & savary(1) & ", state, " & savary(3) & " Terminating error")
            Exit Function
        End If
        
        'GET LOC_ID  (from City and State Code combination)

        strSelect = "select loc_id from LOCATION where City='" & Trim(savary(2)) & "' AND state_code='" & Trim(savary(3)) & "'"
        blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
        If Not rsTemp.EOF Then
            loc_id = rsTemp(0)
        Else
            MsgBox ("(Error)No LOC_ID found for city," & savary(2) & ", state, " & savary(3) & " Terminating error")
            Exit Function
        End If
        
        start_date = ary(5)
        term_date = ary(6)
       
        ':::::::::::::::::::::::::::::::::::::::::::::::::::::::
        '::  UPDATE SELECTED LABOR RATES TO "LABOR_RATE" TABLE
        ':::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Call UpdateToLaborRates(savary(0), _
                                trade_skey, _
                               loc_id, _
                               start_date, _
                               term_date, _
                               Trim(Me.NewTradeGroupCode.Text))


Exit Function
ERRLBL:
    MsgBox ("(Error)BuildAndInsertToDb: " & Err.Description)
'    Stop
'    Resume
    
End Function
Private Sub WriteToLaborRates(Trade_Group_Code, _
                              trade_skey As Integer, _
                              loc_id As Integer, _
                              start_date As Date, _
                              term_date As Date, _
                              New_Trade_Group_Code)
                              
Dim i As Integer
Dim tmpstr As String
Dim insertSql As String
Dim strSelect As String
Dim blnReturn As Boolean
Dim rsTemp As ADODB.RecordSet
'Dim loc_id As String
Dim City As String
Dim State_Code As String
Dim strUpdate As String
Dim cmdTemp As ADODB.Command



On Error GoTo ERRLBL

insertSql = "INSERT INTO LABOR_RATE("
insertSql = insertSql & "trade_skey,"
insertSql = insertSql & "loc_id,"
insertSql = insertSql & "start_date,"
insertSql = insertSql & "term_date,"
insertSql = insertSql & "trade_group_code"
'SET ALL NON NULLABLE columns...
insertSql = insertSql & "contact_id"

insertSql = insertSql & ")"
insertSql = insertSql & "VALUES("
insertSql = insertSql & trade_skey & ","
insertSql = insertSql & loc_id & ",'"
insertSql = insertSql & start_date & "','"
insertSql = insertSql & term_date & "','"
insertSql = insertSql & Trade_Group_Code & "'"
insertSql = insertSql & ")"

MsgBox ("WRITE TO LABOR RATES: " & insertSql)

 strUpdate = insertSql
             'Screen.MousePointer = vbHourglass   'rlh 03/04/2010
g_cnShared.Execute strUpdate    'Allow long-running procedures
Exit Sub
ERRLBL:
    MsgBox ("(Error)WriteToLaborRates: " & Err.Description)
End Sub

Private Sub UpdateToLaborRates(Trade_Group_Code, _
                              trade_skey As Integer, _
                              loc_id As Integer, _
                              start_date As Date, _
                              term_date As Date, _
                              New_Trade_Group_Code)
                              
Dim i As Integer
Dim tmpstr As String
Dim updateSql As String
Dim strSelect As String
Dim deleteSql As String
Dim blnReturn As Boolean
Dim rsTemp As ADODB.RecordSet
Dim cmdTemp As ADODB.Command
'Dim loc_id As String
Dim City As String
Dim State_Code As String
Dim strUpdate As String
Dim ans As Variant




On Error GoTo ERRLBL

'GET City and State from Loc_id
strSelect = "SELECT city, state_code from location WHERE loc_id=" & loc_id
blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
If (Not rsTemp.EOF) Then
    City = rsTemp(0)
    State_Code = rsTemp(1)
End If



''1st   Check if Labor Rate item exists (it better!!!)

strSelect = "SELECT * FROM Labor_Rate Where "
strSelect = strSelect & "trade_group_code='" & Trade_Group_Code & "'"
strSelect = strSelect & " AND trade_skey=" & trade_skey
strSelect = strSelect & " AND loc_id=" & loc_id
'strSelect = strSelect & " AND start_date='" & start_date & "'"
'strSelect = strSelect & " AND term_date='" & term_date & "'"
strSelect = strSelect & " AND month(start_date)=" & Month(start_date)
strSelect = strSelect & " AND day(start_date)=" & Day(start_date)
strSelect = strSelect & " AND year(start_date)=" & Year(start_date)
strSelect = strSelect & " AND month(term_date)=" & Month(term_date)
'strSelect = strSelect & " AND day(term_date)=" & Day(DateAdd("d", 1, term_date))
strSelect = strSelect & " AND day(term_date)=" & Day(term_date)
strSelect = strSelect & " AND year(term_date)=" & Year(term_date)



'MsgBox ("UPDATE TO LABOR RATES: " & strSelect)

blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
    If blnReturn = False Then
        MsgBox "An error occurred loading Trade Groups."
    Else
        If Not (rsTemp.EOF And rsTemp.BOF) Then
            ans = vbOK      'by default
            If WARNINGS_ON Then
                ans = MsgBox("Replacing trade_group_code of " & Trade_Group_Code & _
                vbCrLf & vbCrLf & "Start Date: " & start_date & _
                vbCrLf & "Term Date: " & term_date & _
                vbCrLf & "Loc_id: " & loc_id & _
                vbCrLf & "City: " & City & _
                vbCrLf & "State: " & State_Code & vbCrLf & " on the Labor Rates table as Trade Group Code = " & New_Trade_Group_Code, vbOKCancel, "Continue Remap of Trade Group to: " & New_Trade_Group_Code)
            End If
            Select Case ans
            Case vbOK
                updateSql = "UPDATE LABOR_RATE SET trade_group_code='" & Trim(New_Trade_Group_Code) & "',"
                updateSql = updateSql & " last_update_date='" & Now & "',"
                updateSql = updateSql & " last_update_person='" & strUserName & "'"
                updateSql = updateSql & " WHERE "
                updateSql = updateSql & "trade_group_code='" & Trim(Trade_Group_Code) & "'"
                updateSql = updateSql & " AND trade_skey=" & trade_skey
                updateSql = updateSql & " AND loc_id=" & loc_id
                updateSql = updateSql & " AND month(start_date)=" & Month(start_date)
                updateSql = updateSql & " AND day(start_date)=" & Day(start_date)
                updateSql = updateSql & " AND year(start_date)=" & Year(start_date)
                updateSql = updateSql & " AND month(term_date)=" & Month(term_date)
                'updateSql = updateSql & " AND day(term_date)=" & Day(DateAdd("d", 1, term_date))
                updateSql = updateSql & " AND day(term_date)=" & Day(term_date)
                updateSql = updateSql & " AND year(term_date)=" & Year(term_date)
             
   
                
                
                With g_cnShared
'                    .BeginTrans
                    Set cmdTemp = New ADODB.Command
                    Set cmdTemp.ActiveConnection = g_cnShared
                
                    With cmdTemp
                        .CommandTimeout = 0
                        .CommandType = adCmdText
                        .CommandText = updateSql
                        .Execute 'adExecuteNoRecords
                       
                    End With
                    
                    If .Errors.Count <> 0 Then
                        MsgBox "Errors in the Remap Update routine. " _
                            & vbCrLf & g_cnShared.Errors(0).Description, vbCritical
                        
                        .RollbackTrans
                    Else
'                        .CommitTrans
                         successCnt = successCnt + 1
'                        MsgBox ("The following groups were remapped to group: " & Me.NewTradeGroupCode.Text & _
'                               vbCrLf & strOut)
                    End If
    End With
              Case vbCancel
            End Select
        End If
    End If
            
If successCnt = 0 Then
    MsgBox ("(ERROR(S)): Encountered during the intial remap (stage 1)")
End If

Exit Sub
ERRLBL:
    MsgBox ("(Error)UpdateToLaborRates: " & Err.Description)
End Sub


Private Sub cmdSearch_Click()
'Me.Trade_Group_Code = ""
'Me.Trade_ID = ""
'Me.City = ""
'Me.State_Code = ""

Screen.MousePointer = vbHourglass
Call Me.PopMasterTradeGroupList
Screen.MousePointer = vbNormal
End Sub

Private Sub cmdShiftAllLeft_Click()
Call SelectCtrl.ShiftLeftAll(Me.ListBox1, Me.ListBox2)
End Sub

Private Sub cmdShiftAllRight_Click()
Call SelectCtrl.ShiftRightAll(Me.ListBox1, Me.ListBox2)
End Sub

Private Sub cmdShiftLeft_Click()
Call SelectCtrl.ShiftLeft(Me.ListBox1, Me.ListBox2)
End Sub

Private Sub cmdShiftRight_Click()
Call SelectCtrl.ShiftRight(Me.ListBox1, Me.ListBox2)
End Sub

Private Sub cmdUpdate_Click()
    Dim blnReturn As Boolean
    Dim i As Integer
    
    
    'BUILD THE TRADE GROUP STRING delimited by commas
    
    
    For i = 0 To Me.List1.listcount - 1
        If i = Me.List1.listcount - 1 Then
            strOut = strOut & Me.List1.List(i)
        Else
            strOut = strOut & Me.List1.List(i) & ","
        End If
    Next
    
    blnReturn = Apply(strOut)

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
    '    m_objGridMap.SetGrid TDBGrid
    '    m_objGridMap.InitGrid
    
    Call LoadStates      'rlh 03/25/2009  CCD 8.4+
    
    Call Me.Load_Trade_IDs
    
    '::::::::::::::::::::::::::::::::::::::::::::::
    'CLASS MODULE W/METHODS to handle selection and shifting
    'of Trade Groups
    '::::::::::::::::::::::::::::::::::::::::::::::
    Set SelectCtrl = New clsSelectCtrl  'rlh 03/25/2009  CCD 8.4+
    
'    Me.NewTradeGroupCode.Text = "NEW----"   'rlh temporary
    
    Me.Option1.Value = True

End Sub
Private Sub LoadCities(Optional strCity As String)
Dim strSelect As String
Dim rsTemp As RecordSet
Dim blnReturn As Boolean
Dim strSelDate As String
'Load Cities
    City.Clear
'''    If blnAddMbr = True Then
'''        If IsDate(start_date) Then
'            strSelect = "select distinct city, location.loc_id from labor_rate as lr inner join location on lr.loc_id = location.loc_id " + _
'            " where lr.trade_skey = " + CStr(trade_skey) + _
'" and (convert(varchar(2),DATEPART(m, lr.term_date)) + '/' + convert(varchar(2),DATEPART(d, lr.term_date)) + '/' + convert(varchar(4),DATEPART(yyyy, lr.term_date)))  = '" + PriorDay(start_date) + _
'"' and location.state_code = '" + State_Code.Text + _
'            "' and lr.trade_group_code = '' and lr.start_date = (select max(start_date) from labor_rate where labor_rate.trade_skey = lr.trade_skey and labor_rate.loc_id " + _
'            " = lr.loc_id) ORDER BY city"

'AK- 6/7/2006 - update to be consistent in City retrieval
            If Trim(Me.State_Code.Text) = "" Then
                'rlh 03/25/2009  CCD 8.4+
                strSelect = "select distinct city from labor_rate as lr inner join location on lr.loc_id = location.loc_id "
                strSelect = strSelect & " ORDER BY city"
            Else
                'rlh 03/25/2009  CCD 8.4+
                strSelect = "select distinct city from labor_rate as lr inner join location on lr.loc_id = location.loc_id "
                strSelect = strSelect & " and location.state_code = '" + State_Code.Text & "'"
                strSelect = strSelect & " ORDER BY city"
            End If
            
'''        Else
'''            strSELECT = "select distinct city, location.loc_id from labor_rate as lr inner join location on lr.loc_id = location.loc_id " + _
'''            " where lr.trade_skey = " + CStr(trade_skey) + _
'''            " and location.state_code = '" + State_Code.Text + _
'''            "' and lr.trade_group_code = '' and lr.start_date = (select max(start_date) from labor_rate where labor_rate.trade_skey = lr.trade_skey and labor_rate.loc_id " + _
'''            " = lr.loc_id) ORDER BY city"
'''        End If
'''    Else
'''            strSELECT = "select distinct city, location.loc_id from labor_rate as lr inner join location on lr.loc_id = location.loc_id " + _
'''            " where lr.trade_skey = " + CStr(trade_skey) + _
'''            " and location.state_code = '" + State_Code.Text + _
'''            "' and lr.trade_group_code = '' and lr.start_date = (select max(start_date) from labor_rate where labor_rate.trade_skey = lr.trade_skey and labor_rate.loc_id " + _
'''            " = lr.loc_id) ORDER BY city"
'''
''''        If State_Code.Text > "" Then
''''            strSelect = "select distinct city, loc_id from location where location.state_code = '" + State_Code.Text + "'  order by city"
''''        Else
''''            strSelect = "select distinct city, loc_id from location order by city"
''''        End If
'''    End If
    ' Use DAL to perform select
    blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
    If blnReturn = False Then
        MsgBox "An error occurred loading Cities."
    Else
        If Not (rsTemp.EOF And rsTemp.BOF) Then
            City.AddItem " "     'rlh - Top spot should be a blank
            Do Until rsTemp.EOF
                City.AddItem ConvertCase(rsTemp![City])
                'City.ItemData(City.NewIndex) = rsTemp![loc_id]
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

Public Sub NewTradeGroupCode_Change()
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

Private Sub NewTradeGroupCode_LostFocus()
'Dim rsTemp As ADODB.RecordSet
'Dim strSelect As String
'Dim blnRet As Boolean
'
'On Error GoTo ERRLBL
'
'     strSelect = "select super_trade_group_code from labor_trade_super where super_trade_group_code = '" + NewTradeGroupCode.Text + "'"
'     blnRet = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
'     If Not rsTemp.EOF Then
'        MsgBox ("Warning: this trade group, " & Me.NewTradeGroupCode & " already exists")
'     End If
'     Exit Sub
'ERRLBL:
'    MsgBox ("(Error)NewTradeGroupCode_LostFocus: " & Err.Description)
'
End Sub

Private Sub NewTradeGroupCode_Validate(Cancel As Boolean)
'    Dim strSelect As String
'    Dim blnRet As Boolean
'    Dim strError As String
'    Dim rsTemp As New ADODB.RecordSet
'    If Len(NewTradeGroupCode.Text) > 0 Then
'        strSelect = "select trade_skey from labor_rate where trade_group_code = '" + NewTradeGroupCode.Text + "'"
'        blnRet = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
'        If Not blnRet Then
'            MsgBox "An error occurred retrieving data."
'        Else
'            If rsTemp.RecordCount > 0 Then
'                MsgBox "This trade group code is already in use."
'                Cancel = True
'            End If
'        End If
'        rsTemp.Close
'        Set rsTemp = Nothing
'    End If
End Sub

Private Sub Option1_Click()
WARNINGS_ON = True
Option2 = False
End Sub

Private Sub Option2_Click()
WARNINGS_ON = False
Option1 = False
End Sub

Private Sub State_Code_Click()
If m_State <> State_Code.Text Then
    LoadCities
    m_State = State_Code.Text
    State_Code = Trim(m_State)        'rlh for filtering
End If

If Trim(State_Code) <> "" Then
    Call PopMasterTradeGroupList               'rlh - Apply filter
End If
End Sub

Private Sub State_Code_GotFocus()
m_State = State_Code.Text
End Sub

Private Sub TDBGrid_Click()

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
    'LoadStates
    m_trade_id = Trade_ID.Text
    Trade_ID = Trim(m_trade_id)       'rlh for filtering
End If

Call PopMasterTradeGroupList               'rlh - Apply filter
End Sub

Private Sub Trade_ID_GotFocus()
m_trade_id = Trade_ID.Text
End Sub

Public Sub PopMasterTradeGroupList()
    Dim INString As String
    Dim i As Integer
    Dim strSelect As String
    
    Dim blnReturn As Boolean
    
    On Error GoTo ERRLBL
    
    Set rsTemp = New ADODB.RecordSet
    
    INString = ""
    For i = 0 To List1.listcount
        If i = List1.listcount Then
            INString = INString & "'" & List1.List(i) & "'"
        Else
            INString = INString & "'" & List1.List(i) & "',"
        End If
    Next
    If List1.listcount > 0 Then
        'rlh 03/25/2009   CCD 8.4+  ----------------------------
'        strSelect = "SELECT DISTINCT lr.trade_group_code, lt.trade_id, loc.city, loc.state_code, lr.start_date, lr.term_date, lr.union_base_hrly, union_fringe_hrly, lr.union_base_hrly + union_fringe_hrly as tot_union  FROM LABOR_RATE lr, LOCATION loc, LABOR_TRADE lt"
'        strSelect = strSelect & " WHERE "
''        strSelect = strSelect & "getdate() between lr.start_date and lr.term_date "
''        strSelect = strSelect & " AND"
'        strSelect = strSelect & " Loc.loc_id = lr.loc_id"
'        strSelect = strSelect & " AND"
'        strSelect = strSelect & " lt.trade_skey = lr.trade_skey"
'        strSelect = strSelect & " AND"
'        strSelect = strSelect & " trade_group_code IN(" & INString & ")"
'        strSelect = strSelect & " AND trade_group_code <> '       '"
'        strSelect = strSelect & " ORDER BY trade_group_code"

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'   PRAYING THIS WILL SATISFY GENI M. (it should produce output exactly as the labor trade grid does when
'   entering a trade_group_code ONLY!!! Copied from stored procedure: sp_LaborRatesMaxStart
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    strSelect = "exec sp_LaborRemapMaxStart @trade_id='', @trade_group_code='" & Me.Trade_Group_Code.Text & "', @city='', @state='', @start_date='', @term_date='', @includehistory = 0, @maxrowcount = 5000"

'    blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
'    If blnReturn = False Then
'        MsgBox "An error occurred while searching."
'        lblRowCount.Caption = "0 rows returned."
'        GoTo ERRLBL
'    End If
    Else
        strSelect = "SELECT DISTINCT lr.trade_group_code, lt.trade_id, loc.city, loc.state_code, lr.start_date, lr.term_date, lr.union_base_hrly, union_fringe_hrly, lr.union_base_hrly + union_fringe_hrly as tot_union    FROM LABOR_RATE lr, LOCATION loc, LABOR_TRADE lt"
        strSelect = strSelect & " WHERE "
'        strSelect = strSelect & "getdate() between lr.start_date and lr.term_date "
'        strSelect = strSelect & " AND"
        strSelect = strSelect & " Loc.loc_id = lr.loc_id"
        strSelect = strSelect & " AND"
        strSelect = strSelect & " lt.trade_skey = lr.trade_skey"
        strSelect = strSelect & " AND trade_group_code <> '       '"
'                strSELECT = strSELECT & " AND"
'                strSELECT = strSELECT & " trade_group_code IN(" & INString & ")"
        strSelect = strSelect & " ORDER BY trade_group_code"
    
    End If
    
    Dim strFilter As String
    strFilter = ""
    ':::::::::::::::::::::
    'FILTERS
    ':::::::::::::::::::::
    If Trade_ID <> "" Then
    
        strFilter = " trade_id='" & Trade_ID & "'"
    
    End If
    
    If Trim(State_Code) <> "" Then
    
        If Len(Trim(strFilter)) = 0 Then
            strFilter = " state_code='" & State_Code & "'"
        Else
            strFilter = strFilter & " AND state_code='" & State_Code & "'"
        End If
        
    End If
        
    If get_counts = 0 Then
        MsgBox ("(ERROR) You have selected a non-existent trade group: " & List1.Text)
        List1.RemoveItem List1.listcount - 1
        Exit Sub
    End If
        
    If Trim(Me.City) <> "" Then
        If (Len(Trim(strFilter)) = 0) Then
            strFilter = " City='" & City & "'"
        Else
            strFilter = strFilter & " AND City='" & City & "'"
        End If
    End If
    
    blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
    If blnReturn = False Then
        MsgBox "An error occurred loading Trade Groups."
    Else
        If Not (rsTemp.EOF And rsTemp.BOF) Then
            If DEBUGON Then Stop
            'Set rsClone = rsTemp.Clone()
            'Apply Filter
            If Trim(strFilter) <> "" Then
                rsTemp.Filter = strFilter
            End If
            
            'DISPLAY ROW COUNT
            Me.lblRowCount.Caption = "Row Count: " & rsTemp.RecordCount
            
            '::::::::::::::::::::::::::::::::::::::::::::::
            'CLASS MODULE W/METHODS to handle selection and shifting
            'of Trade Groups
            '::::::::::::::::::::::::::::::::::::::::::::::
            Set SelectCtrl = New clsSelectCtrl  'rlh 03/25/2009  CCD 8.4+
            'Call SelectCtrl.AddAllTradeGroupsToOneListView(1, Me.lv, rsTemp)
            Call SelectCtrl.AddAllTradeGroupsToOneListBox(1, Me.ListBox1, rsTemp)
        End If
    End If
    
    'CLEAR the selection/target list box
    Me.ListBox2.Clear
    
    'rlh 03/25/2009 -------------------------------------------
    Exit Sub
ERRLBL:
    MsgBox ("(Error)PopMasterTradeGroupList: " & Err.Description)
'    Stop
'    Resume
    
End Sub
Public Sub RefreshMasterTradeGroupList()
 Dim INString As String
    Dim i As Integer
    Dim strSelect As String

    
    Dim blnReturn As Boolean
    
    On Error GoTo ERRLBL
    
    Set rsTemp = New ADODB.RecordSet
    
    INString = ""
    For i = 0 To List1.listcount - 1
        If i = List1.listcount Then
            INString = INString & "'" & List1.List(i) & "'"
        Else
            INString = INString & "'" & List1.List(i) & "',"
        End If
   
    
        If List1.listcount > 0 Then
     
            ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
            '   PRAYING THIS WILL SATISFY GENI M. (it should produce output exactly as the labor trade grid does when
            '   entering a trade_group_code ONLY!!! Copied from stored procedure: sp_LaborRatesMaxStart
            ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
            strSelect = "exec sp_LaborRemapMaxStart @trade_id='', @trade_group_code='" & List1.List(i) & "', @city='', @state='', @start_date='', @term_date='', @includehistory = 0, @maxrowcount = 5000"
        End If
        
        Dim strFilter As String
        strFilter = ""
        ':::::::::::::::::::::
        'FILTERS
        ':::::::::::::::::::::
        If Trade_ID <> "" Then
        
            strFilter = " trade_id='" & Trade_ID & "'"
        
        End If
        
        If Trim(State_Code) <> "" Then
        
            If Len(Trim(strFilter)) = 0 Then
                strFilter = " state_code='" & State_Code & "'"
            Else
                strFilter = strFilter & " AND state_code='" & State_Code & "'"
            End If
            
        End If
            
        If get_counts = 0 Then
            MsgBox ("(ERROR) You have selected a non-existent trade group: " & List1.Text)
            List1.RemoveItem List1.listcount - 1
            Exit Sub
        End If
            
        If Trim(Me.City) <> "" Then
            If (Len(Trim(strFilter)) = 0) Then
                strFilter = " City='" & City & "'"
            Else
                strFilter = strFilter & " AND City='" & City & "'"
            End If
        End If
        
        blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
        If blnReturn = False Then
            MsgBox "An error occurred loading Trade Groups."
        Else
            If Not (rsTemp.EOF And rsTemp.BOF) Then
                If DEBUGON Then Stop
                'Set rsClone = rsTemp.Clone()
                'Apply Filter
                If Trim(strFilter) <> "" Then
                    rsTemp.Filter = strFilter
                End If
                
                'DISPLAY ROW COUNT
                Me.lblRowCount.Caption = "Row Count: " & rsTemp.RecordCount
                
                '::::::::::::::::::::::::::::::::::::::::::::::
                'CLASS MODULE W/METHODS to handle selection and shifting
                'of Trade Groups
                '::::::::::::::::::::::::::::::::::::::::::::::
                Set SelectCtrl = New clsSelectCtrl  'rlh 03/25/2009  CCD 8.4+
                'Call SelectCtrl.AddAllTradeGroupsToOneListView(1, Me.lv, rsTemp)
                Call SelectCtrl.AddAllTradeGroupsToOneListBox(1, Me.ListBox1, rsTemp)
            End If
        End If
    Next
    'CLEAR the selection/target list box
    Me.ListBox2.Clear
    
    'rlh 03/25/2009 -------------------------------------------
    Exit Sub
ERRLBL:
    MsgBox ("(Error)PopMasterTradeGroupList: " & Err.Description)
'    Stop
'    Resume
End Sub
Public Function Apply(strOut As String)
Dim strUpdate As String
Dim cmdTemp As ADODB.Command

On Error GoTo ERRLBL
    
    'strUpdate = "exec sp_remap_labor_trade_groups '" & strOut & "','" & Me.NewTradeGroupCode.Text & "'"
    strUpdate = "exec sp_new_super_trade_group '" & strOut & "','" & Me.NewTradeGroupCode.Text & "'"
    With g_cnShared
        .BeginTrans
        Set cmdTemp = New ADODB.Command
        Set cmdTemp.ActiveConnection = g_cnShared
    
        With cmdTemp
            .CommandTimeout = 0
            .CommandType = adCmdText
            .CommandText = strUpdate
            '.Execute 'adExecuteNoRecords
        End With
    
        If .Errors.Count <> 0 Then
            MsgBox "Errors in the Remap Update routine. " _
                & vbCrLf & g_cnShared.Errors(0).Description, vbCritical
            
            .RollbackTrans
        Else
            .CommitTrans
            Apply = True
            MsgBox ("The following groups were remapped to group: " & Me.NewTradeGroupCode.Text & _
                   vbCrLf & strOut)
        End If
    End With
    Exit Function
ERRLBL:
    MsgBox ("(Error)Apply: " & Err.Description)
    
    
End Function

Public Function Restore() As Boolean

Dim cmdTemp As ADODB.Command
Dim strUpdate As String

strUpdate = "exec sp_Restore_Labor_Rate_Table"

On Error GoTo ERRLBL

Restore = False

With g_cnShared
        .BeginTrans
        Set cmdTemp = New ADODB.Command
        Set cmdTemp.ActiveConnection = g_cnShared
    
        With cmdTemp
            .CommandTimeout = 0
            .CommandType = adCmdText
            .CommandText = strUpdate
            '.Execute 'adExecuteNoRecords
        End With
    
        If .Errors.Count <> 0 Then
            MsgBox "Errors in the Remap Update routine. " _
                & vbCrLf & g_cnShared.Errors(0).Description, vbCritical
            
            .RollbackTrans
        Else
            .CommitTrans
            Restore = True
            MsgBox ("Restore was successful")
        End If
    End With
    Exit Function
ERRLBL:
    MsgBox ("(Error)Restore: " & Err.Description)
'    Stop
'    Resume
End Function


