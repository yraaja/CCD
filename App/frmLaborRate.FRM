VERSION 5.00
Begin VB.Form frmLaborRate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Labor Rate Maintenance"
   ClientHeight    =   5880
   ClientLeft      =   3870
   ClientTop       =   3180
   ClientWidth     =   9225
   Icon            =   "frmLaborRate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5880
   ScaleWidth      =   9225
   Begin VB.ComboBox trade_id 
      Height          =   315
      Left            =   1320
      TabIndex        =   51
      Tag             =   "0"
      Top             =   60
      Width           =   1470
   End
   Begin VB.TextBox trade_desc 
      Height          =   315
      Left            =   2925
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   60
      Width           =   4380
   End
   Begin VB.TextBox davis_bacon_overtime_hrly 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7320
      TabIndex        =   12
      Tag             =   "1N"
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox davis_bacon_travel_hrly 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7320
      TabIndex        =   17
      Tag             =   "1N"
      Top             =   3360
      Width           =   1335
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
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   48
      Tag             =   "1N"
      Top             =   5040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox worker_type_code 
      Height          =   315
      ItemData        =   "frmLaborRate.frx":0442
      Left            =   6480
      List            =   "frmLaborRate.frx":0444
      TabIndex        =   47
      Top             =   450
      Width           =   855
   End
   Begin VB.TextBox trade_group_code 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Tag             =   "1S"
      Top             =   450
      Width           =   1215
   End
   Begin VB.TextBox TotalUnion 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1320
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox TotalOpen 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3480
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox TotalDavisBacon 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5640
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox comment 
      Height          =   285
      Left            =   1320
      TabIndex        =   8
      Tag             =   "1S"
      Top             =   2040
      Width           =   6375
   End
   Begin VB.TextBox local_union_num 
      Height          =   285
      Left            =   4200
      MaxLength       =   6
      TabIndex        =   6
      Tag             =   "1G"
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox davis_bacon_fringe_hrly 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5640
      TabIndex        =   16
      Tag             =   "1N"
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox davis_bacon_base_hrly 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5640
      TabIndex        =   11
      Tag             =   "1N"
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CheckBox estimated_ind 
      Caption         =   "&Estimated"
      Height          =   255
      Left            =   5760
      TabIndex        =   7
      Tag             =   "1"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox opn_fringe_hrly 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3480
      TabIndex        =   15
      Tag             =   "1N"
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox opn_base_hrly 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3480
      TabIndex        =   10
      Tag             =   "1N"
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox union_fringe_hrly 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1320
      TabIndex        =   13
      Tag             =   "1N"
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox union_base_hrly 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1320
      TabIndex        =   9
      Tag             =   "1N"
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox worker_comp_id 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4200
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   450
      Width           =   1095
   End
   Begin VB.TextBox contact_id 
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Tag             =   "1S"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox start_date 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Tag             =   "1D"
      Top             =   1200
      Width           =   1080
   End
   Begin VB.TextBox term_date 
      Height          =   285
      Left            =   4200
      TabIndex        =   4
      Tag             =   "1D"
      Top             =   1200
      Width           =   1080
   End
   Begin VB.ComboBox State_Code 
      Height          =   315
      ItemData        =   "frmLaborRate.frx":0446
      Left            =   1320
      List            =   "frmLaborRate.frx":0448
      TabIndex        =   1
      Top             =   810
      Width           =   765
   End
   Begin VB.ComboBox City 
      Height          =   315
      ItemData        =   "frmLaborRate.frx":044A
      Left            =   4200
      List            =   "frmLaborRate.frx":044C
      TabIndex        =   2
      Text            =   "City"
      Top             =   810
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
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   25
      Tag             =   "1N"
      Top             =   5400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   495
      Left            =   4920
      TabIndex        =   20
      Top             =   5220
      Width           =   1275
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
      TabIndex        =   24
      TabStop         =   0   'False
      Tag             =   "1G"
      Top             =   4380
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
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   4380
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
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4380
      Width           =   1755
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Default         =   -1  'True
      Height          =   495
      Left            =   3180
      TabIndex        =   18
      Top             =   5220
      Width           =   1275
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   120
      X2              =   9120
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      Caption         =   "Travel/Overtime"
      Height          =   255
      Left            =   7320
      TabIndex        =   49
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label20 
      Caption         =   "Trade_Skey"
      Height          =   255
      Left            =   6240
      TabIndex        =   46
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      Caption         =   "Total:"
      Height          =   255
      Left            =   180
      TabIndex        =   45
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      Caption         =   "Comment:"
      Height          =   255
      Left            =   420
      TabIndex        =   41
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      Caption         =   "Local Union:"
      Height          =   255
      Left            =   3120
      TabIndex        =   40
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label18 
      Caption         =   "------- Davis-Bacon  -------"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      TabIndex        =   39
      Top             =   2400
      Width           =   3015
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "Open"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   38
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Caption         =   "Fringe Hourly:"
      Height          =   255
      Left            =   60
      TabIndex        =   37
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Base Hourly:"
      Height          =   255
      Left            =   180
      TabIndex        =   36
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "Union"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   35
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Worker Comp ID:"
      Height          =   255
      Left            =   2760
      TabIndex        =   34
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "Contact ID:"
      Height          =   255
      Left            =   300
      TabIndex        =   33
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "Worker Type"
      Height          =   255
      Left            =   5400
      TabIndex        =   32
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Trade ID:"
      Height          =   255
      Left            =   300
      TabIndex        =   31
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Trade Group:"
      Height          =   255
      Left            =   0
      TabIndex        =   30
      Top             =   480
      Width           =   1155
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "City:"
      Height          =   255
      Left            =   3720
      TabIndex        =   29
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "State:"
      Height          =   255
      Left            =   540
      TabIndex        =   28
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Start Date:"
      Height          =   255
      Left            =   300
      TabIndex        =   27
      Top             =   1215
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Term Date:"
      Height          =   255
      Left            =   3240
      TabIndex        =   26
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Last Updated By:"
      Height          =   255
      Left            =   3180
      TabIndex        =   23
      Top             =   4440
      Width           =   1395
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Last Updated:"
      Height          =   255
      Left            =   60
      TabIndex        =   22
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      Index           =   0
      X1              =   135
      X2              =   9120
      Y1              =   4935
      Y2              =   4935
   End
End
Attribute VB_Name = "frmLaborRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_rec As ADODB.RecordSet
Dim m_blnRecFlag As Boolean ' True if the screen has a data source, False if we are adding new record
Dim m_blnDeleted As Boolean ' Indicates if the data has been deleted, used in QueryUnload
Dim m_blnWereErrors As Boolean ' True if the Update had errors, used in QueryUnload
Dim m_blnInsert As Boolean
Dim m_State As String
Dim m_trade_id As String

Public frmCallingForm As Form

Private Sub CalcRateTotals()
    TotalUnion = Format(Val(union_base_hrly) + Val(union_fringe_hrly), "#,###,##0.00")
    TotalOpen = Format(Val(opn_base_hrly) + Val(opn_fringe_hrly), "#,###,##0.00")
    TotalDavisBacon = Format(Val(davis_bacon_base_hrly) + Val(davis_bacon_fringe_hrly), "#,###,##0.00")
End Sub

Private Function EntryErrorsFound() As Boolean
'Validate entry fields
    EntryErrorsFound = TermDateError()
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
        If m_blnRecFlag Then    'Cloning record
            m_rec.Fields("start_date") = DateAdd("d", 1, m_rec.Fields("term_date"))
            m_rec.Fields("term_date") = DateAdd("d", -1, DateAdd("yyyy", 1, m_rec.Fields("start_date")))
        End If
        ' Do this so OriginalValue will be set to the values copied into the row
        m_rec.UpdateBatch
    End If
End Sub

Private Function NewTradeLoc() As Date
Dim rsTemp As ADODB.RecordSet
Dim strUpdate As String
Dim strSelect As String
Dim blnReturn As Boolean

'Verify the existance of a labor rate record for the trade ID and location code.
    On Error Resume Next
    strSelect = "select labor_rate.trade_skey, convert(varchar(2),DATEPART(month,max(term_date))) + '/' + convert(varchar(2),DATEPART(day, max(term_date)))  + '/' + convert(varchar(4),DATEPART(year, max(term_date)))  as existing_term_date " + _
    "From labor_rate " + _
    "inner join labor_trade on labor_rate.trade_skey = labor_trade.trade_skey " + _
              "and labor_trade.trade_id = '" + Trade_ID.Text + "' " + _
    "group by labor_rate.trade_skey, labor_rate.loc_id " + _
    "Having labor_rate.loc_id = " & City.ItemData(City.ListIndex)

        ' Use DAL to perform select
        blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
        If blnReturn = False Then
            MsgBox "An error occurred while searching."
            'lblRowCount.Caption = "0 rows returned."
            GoTo Exit_Sub
        End If
    If rsTemp.RecordCount = 0 Then
        NewTradeLoc = 0
    Else
        NewTradeLoc = Format(rsTemp.Fields("existing_term_date"), "short date")
    End If
    rsTemp.Close
    Set rsTemp = Nothing
Exit_Sub:

End Function

Private Function TermDateError() As Boolean
    Dim intDateDiff As Integer
    If IsDate(term_date) Then
        intDateDiff = DateDiff("d", start_date, term_date)
        If intDateDiff <= 0 Then
            MsgBox "The Term Date must be after the Start Date."
            TermDateError = True
        End If
    Else
        MsgBox "Please enter a valid Term Date."
        TermDateError = True
    End If

End Function

Private Sub City_Click()
    Me.Refresh
    comment.Refresh
End Sub

Private Sub City_LostFocus()
    If start_date.Locked = True Then
        term_date.SetFocus
    End If
End Sub

Private Sub City_Validate(Cancel As Boolean)
' Validate the trade/loc_id combination for a new record
    Dim dtStartDate As Date

    dtStartDate = NewTradeLoc()
    If m_blnRecFlag = False And City.Text > "" Then      'new record
        If dtStartDate = 0 Then
            start_date.Locked = False
            start_date = Now()
        Else
            LockField Me, "start_date"
            start_date = DateAdd("d", 1, dtStartDate)
        End If
    End If

End Sub

Private Sub cmdDelete_Click()
    On Error Resume Next
    Dim strUpdate As String
    Dim blnRet As Boolean
    Dim strError As String
    Dim rsTemp As ADODB.RecordSet
    Dim varButton
    If m_rec.Fields("trade_group_code") > " " Then
        strUpdate = "select count(*) as RcdsToDelete from labor_rate where trade_group_code='" + m_rec.Fields("trade_group_code") + "' and start_date='" + Format(m_rec.Fields("start_date"), "mm/dd/yyyy") + "'"
        blnRet = g_objDAL.GetRecordset(CONNECT, strUpdate, rsTemp)
        If Not blnRet Then
            MsgBox "An error occurred retrieving data."
        Else
            If rsTemp![RcdsToDelete] > 0 Then
                Dim strMsg As String
                strMsg = CStr(rsTemp![RcdsToDelete]) + " Labor rate records will be removed.  Are you sure you want to delete?"
                varButton = MsgBox(strMsg, vbYesNo + vbCritical)
            End If
        End If
        rsTemp.Close
        Set rsTemp = Nothing
    Else
        varButton = MsgBox("Are you sure you want to delete?", vbYesNo + vbCritical)
    End If
    
    If varButton = vbYes Then
        ' Build SQL statement
        strUpdate = "exec sp_delete_labor_rate "
        strUpdate = strUpdate + "@trade_skey=" + str(m_rec.Fields("trade_skey")) + ","
        strUpdate = strUpdate + " @loc_id=" + str(m_rec.Fields("loc_id")) + ","
        strUpdate = strUpdate + " @last_update_id=" + str(m_rec.Fields("last_update_id")) + ","
        strUpdate = strUpdate + " @start_date='" + Format(m_rec.Fields("start_date"), "mm/dd/yyyy") + "',"
        strUpdate = strUpdate + " @trade_group_code='" + m_rec.Fields("trade_group_code") + "'"
        
        blnRet = g_objDAL.ExecQuery(CONNECT, strUpdate, strError)
        If Not blnRet Then
            MsgBox strError
        Else
            MsgBox "Delete successful.  The Labor Rate Grid will now be refreshed."
            m_rec.Delete
            m_blnDeleted = True
            frmCallingForm.cmdSearch_Click
            Unload Me
        End If
    End If

End Sub

Private Sub cmdUpdate_Click()
    On Error Resume Next
    Dim strUpdate As String
    Dim strError As String
    Dim blnRet As Boolean
    Dim i As Integer
    
    m_blnWereErrors = False
    If City.ListIndex = -1 Then
        MsgBox "Please enter a valid state/city."
    Else
        If m_blnInsert = False Then        ' Update
            strUpdate = "exec sp_update_labor_rate "
        Else                                            ' Insert
            strUpdate = "exec sp_insert_labor_rate "
            For i = 0 To City.listcount - 1
                If UCase(Trim(City.Text)) = UCase(Trim(City.List(i))) Then Exit For
            Next i
            loc_id = City.ItemData(i)
        End If
        BuildStoredProcSQL Me, strUpdate, "1", m_rec
        strUpdate = strUpdate + " @last_update_person='" + strUserName + "'"
        blnRet = g_objDAL.ExecQuery(CONNECT, strUpdate, strError)
        If Not blnRet Then
            MsgBox strError
            m_blnWereErrors = True
        Else
            ' Put latest data into source recordset
            UpdateRecordsetFromForm Me, m_rec
            m_rec.Fields("mat_last_update_id").Value = m_rec.Fields("mat_last_update_id").Value + 1
            'mat_last_update_id.Text = m_rec.Fields("mat_last_update_id").Value
            UpdateFormFromRecordset Me, m_rec
            MsgBox "Update successful."
        End If
    End If
End Sub

Private Sub davis_bacon_base_hrly_LostFocus()
    CalcRateTotals
End Sub

Private Sub davis_bacon_fringe_hrly_LostFocus()
    CalcRateTotals
End Sub

Private Sub davis_bacon_overtime_hrly_LostFocus()
    CalcRateTotals
End Sub

Private Sub davis_bacon_travel_hrly_LostFocus()
    CalcRateTotals
End Sub

Private Sub Form_Activate()
    OutputView False
End Sub

Private Sub Form_Initialize()
    m_blnRecFlag = False
    m_blnDeleted = False
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim ctr As Control
    Dim rec As ADODB.RecordSet
    Move START_LEFT, START_TOP

    If m_blnInsert = True And m_blnRecFlag = True Then
        Me.Caption = Me.Caption + " [Clone of " + m_rec.Fields("trade_id").Value + "/" + m_rec.Fields("City").Value + ", " + m_rec.Fields("State_Code").Value + "]"
    Else
        Me.Caption = Me.Caption + " [" + m_rec.Fields("trade_id").Value + "/" + m_rec.Fields("City").Value + ", " + m_rec.Fields("State_Code").Value + "]"
    End If

    ' If we are showing data
    If m_blnRecFlag = True Then
        If Not m_rec.State = adStateClosed Then
            UpdateFormFromRecordset Me, m_rec
'            term_date = Format(term_date, "mm/dd/yyyy")    'pkl 9/21/01 ????? causes prompt to save1
            CalcRateTotals
        End If
    End If
    
    ' Lock fields that can't be changed
    If (m_blnInsert = True And m_blnRecFlag = True) Or m_blnInsert = False Then   'New/Not Clone allows entry
        ' 9/30/2005 RTD - CHANGED TO MainModule.LockField() FUNCTION - REQUIRES FORM, "CONTROL"
        LockField Me, "start_date"
        start_date.Enabled = False
        LockField Me, "trade_id"
        LockField Me, "city"
        LockField Me, "state_code"
    End If
        
    If (m_blnInsert = True And m_blnRecFlag = True) Or m_blnInsert = True Then   'New/Not Clone allows entry
        cmdDelete.Visible = False
    End If

    ' 9/30/2005 RTD - CHANGED TO MainModule.LockField() FUNCTION - REQUIRES FORM, "CONTROL"
    LockField Me, "worker_type_code"
    LockField Me, "trade_desc"
    LockField Me, "TotalUnion"
    LockField Me, "TotalDavisBacon"
    LockField Me, "TotalOpen"
    LockField Me, "trade_skey"
    LockField Me, "trade_group_code"
    LockField Me, "worker_comp_id"
    LoadCombos
    
    ColorLockedFields Me
    
End Sub

Private Sub LoadCities(Optional strCity As String)
    Dim strSelect As String
    Dim rsTemp As RecordSet
    Dim blnReturn As Boolean

    'Load Cities
    If State_Code.Text > "" Then
        strSelect = "select distinct city, loc_id from location where location.state_code = '" + State_Code.Text + "'  order by city"
    Else
        strSelect = "select distinct city, loc_id from location order by city"
    End If
    ' Use DAL to perform select
    blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
    If blnReturn = False Then
        MsgBox "An error occurred loading Cities."
    Else
        City.AddItem " "
        City.ItemData(City.NewIndex) = 0
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

Private Sub LoadCombos()
    Dim blnReturn As Boolean
    Dim strSelect As String
    Dim rsTemp As RecordSet

'Load All Selection Combos

'Load Trade IDs
    strSelect = "SELECT trade_id, trade_skey  FROM LABOR_TRADE ORDER BY LABOR_TRADE.trade_id"
    
    ' Use DAL to perform select
    blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
    If blnReturn = False Then
        MsgBox "An error occurred loading Trade IDs."
        'lblRowCount.Caption = "0 rows returned."
    Else
        If Not (rsTemp.EOF And rsTemp.BOF) Then
            Do Until rsTemp.EOF
                Trade_ID.AddItem rsTemp!Trade_ID
                Trade_ID.ItemData(Trade_ID.NewIndex) = rsTemp!trade_skey
                rsTemp.MoveNext
            Loop
        End If
    End If
    rsTemp.Close


'Load States
    strSelect = "select distinct state_code from location order by state_code;"

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
    
    'Load Worker Type Codes
    strSelect = "SELECT DISTINCT worker_type_code FROM LABOR_TRADE Where LABOR_TRADE.worker_type_code <> '' ORDER BY LABOR_TRADE.worker_type_code"
    ' Use DAL to perform select
    blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
    If blnReturn = False Then
        MsgBox "An error occurred loading Worker Type Codes."
    Else
        If Not (rsTemp.EOF And rsTemp.BOF) Then
            Do Until rsTemp.EOF
                worker_type_code.AddItem rsTemp![worker_type_code]
                rsTemp.MoveNext
            Loop
        End If
    End If
    rsTemp.Close
    
    LoadCities

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim Button
    Dim blnPendingChange As Boolean
    
'    If blnPendingChange = False Then
'        Unload Me
'        Else
    
    ' Only go through this if the close wasn't invoked from code
    If Not UnloadMode = vbFormCode Then
        blnPendingChange = IsControlChanged(Me, m_rec)
        
        
    
        If blnPendingChange = True Then
            Button = MsgBox("Do you want to save your changes?", vbYesNoCancel)
            If Button = vbYes Then
                m_blnWereErrors = EntryErrorsFound()
                If m_blnWereErrors = False Then
                    cmdUpdate_Click
                End If
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
'End If
End Sub

Private Sub Form_Resize()
    ResizeForm Me
End Sub

Private Sub opn_base_hrly_LostFocus()
    CalcRateTotals
End Sub

Private Sub opn_fringe_hrly_LostFocus()
    CalcRateTotals
End Sub

Private Sub State_Code_Change()
    Dim iSelStart  As Integer
    Dim iSelLen As Integer
    
    iSelStart = State_Code.SelStart
    iSelLen = State_Code.SelLength
    State_Code = UCase(State_Code)
    State_Code.SelStart = iSelStart
    State_Code.SelLength = iSelLen

End Sub

Private Sub State_Code_GotFocus()
    m_State = State_Code.Text
End Sub

Private Sub State_Code_LostFocus()
    If m_State <> State_Code.Text Then
        City.Clear
        LoadCities
    End If
End Sub

Private Sub term_date_Validate(Cancel As Boolean)
    Cancel = TermDateError()
End Sub

Private Sub Trade_ID_GotFocus()
    m_trade_id = Trade_ID.Text
End Sub

Private Sub Trade_ID_LostFocus()
    Dim strSelect As String
    Dim rsTemp As RecordSet
    Dim blnReturn As Boolean
    Dim i As Integer
    If Trade_ID.Text <> m_trade_id Then
        strSelect = "select trade_desc, worker_type_code from LABOR_TRADE where LABOR_TRADE.trade_id = '" + Trade_ID.Text + "'"
        ' Use DAL to perform select
        blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
        If blnReturn = False Then
            MsgBox "An error occurred loading trade information."
        Else
            If Not (rsTemp.EOF And rsTemp.BOF) Then
                trade_desc.Text = ConvertCase(rsTemp![trade_desc])
                worker_type_code.Text = ConvertCase(rsTemp![worker_type_code])
            End If
            For i = 0 To Trade_ID.listcount - 1
                If UCase(Trim(Trade_ID.Text)) = UCase(Trim(Trade_ID.List(i))) Then Exit For
            Next i
            trade_skey.Text = CStr(Trade_ID.ItemData(i))
        End If
        rsTemp.Close
        Set rsTemp = Nothing
    End If
End Sub

Private Sub union_base_hrly_LostFocus()
    CalcRateTotals
End Sub

Private Sub union_fringe_hrly_LostFocus()
    CalcRateTotals
End Sub

Private Sub union_fringe_hrly_Validate(Cancel As Boolean)
    If Not IsNumeric(union_fringe_hrly) Then
        MsgBox "Please enter a valid Union Fringe amount.", vbExclamation
        Cancel = True
    End If
End Sub

