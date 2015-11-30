VERSION 5.00
Begin VB.Form frmCrew 
   Caption         =   "Edit Crew"
   ClientHeight    =   3360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6810
   Icon            =   "frmCrew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3360
   ScaleWidth      =   6810
   Visible         =   0   'False
   Begin VB.TextBox traces_crew_type_code 
      Height          =   285
      Left            =   5040
      TabIndex        =   23
      Tag             =   "1"
      Text            =   "Text1"
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox type_code 
      Height          =   285
      Left            =   5040
      TabIndex        =   22
      Tag             =   "1"
      Text            =   "Text1"
      Top             =   90
      Width           =   375
   End
   Begin VB.TextBox traces_crew_desc 
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      Tag             =   "1"
      Top             =   1560
      Width           =   3615
   End
   Begin VB.TextBox traces_crew_id 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Tag             =   "1"
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox crew_skey 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   18
      Tag             =   "1N"
      Text            =   "Text1"
      Top             =   2760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox crew_id 
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Top             =   60
      Width           =   1335
   End
   Begin VB.TextBox last_update_id 
      Height          =   375
      Left            =   5280
      TabIndex        =   17
      Tag             =   "1"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   495
      Left            =   2280
      TabIndex        =   6
      Top             =   2760
      Width           =   1150
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   3720
      TabIndex        =   7
      Top             =   2760
      Width           =   1150
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
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2280
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
      Left            =   5340
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox comment 
      Height          =   285
      Left            =   2040
      TabIndex        =   5
      Tag             =   "1"
      Top             =   1920
      Width           =   4695
   End
   Begin VB.TextBox metric_crew_desc 
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Tag             =   "1"
      Top             =   1200
      Width           =   3615
   End
   Begin VB.TextBox crew_desc 
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Tag             =   "1"
      Top             =   840
      Width           =   3615
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "Traces Crew Description:"
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Traces Crew ID:"
      Height          =   255
      Left            =   480
      TabIndex        =   20
      Top             =   540
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Traces Crew Type:"
      Height          =   255
      Left            =   3480
      TabIndex        =   19
      Top             =   540
      Width           =   1455
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Last Updated:"
      Height          =   255
      Left            =   840
      TabIndex        =   16
      Top             =   2340
      Width           =   1095
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Last Updated By:"
      Height          =   255
      Left            =   3840
      TabIndex        =   15
      Top             =   2340
      Width           =   1395
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Comment:"
      Height          =   255
      Left            =   960
      TabIndex        =   12
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Crew Type:"
      Height          =   255
      Left            =   3960
      TabIndex        =   11
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Crew Metric Description:"
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Crew Description:"
      Height          =   255
      Left            =   600
      TabIndex        =   9
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Crew ID:"
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmCrew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public m_strCrewID As String
Public m_blnCloneCrew As Boolean

Dim m_rec As New ADODB.RecordSet
Dim m_blnWereErrors As Boolean
Private Sub FillCrews()
Dim rec As ADODB.RecordSet
Dim strSelect As String
Dim blnReturn As Boolean

'Populate the crew selection list
    strSelect = "SELECT crew_id FROM crew order by crew_id"

    ' Use DAL to perform select
    blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rec)
    If blnReturn = False Then
        MsgBox "An error occurred loading crews."
    Else
        If Not (rec.EOF And rec.BOF) Then
            Do Until rec.EOF
                crew_id.AddItem rec!crew_id
                If m_strCrewID > "" Then
                    If m_strCrewID = rec!crew_id Then
                        crew_id.ListIndex = crew_id.NewIndex
                    End If
                End If
                rec.MoveNext
            Loop
        End If
    End If
    If crew_id.ListIndex = -1 Then  'New crew
        crew_id.Text = m_strCrewID
    End If
End Sub

Public Sub JumpIn()
    FillCrews   'Load combo box
    LoadCrew    'Populate form
End Sub

Private Sub LoadCrew()
Dim strSelect As String
Dim blnReturn As Boolean
On Error Resume Next
'Populate the form - new or existing
    strSelect = "SELECT  crew_skey, crew_id, traces_crew_id, traces_crew_type_code, crew_desc," + _
        " traces_crew_desc, metric_crew_desc, type_code, comment, last_update_date, last_update_person, last_update_id " + _
        " from CREW WHERE crew_id = '" + crew_id.Text + "'"
    m_rec.Close
    ' Use DAL to perform select
    blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, m_rec)
    If blnReturn = False Then
        MsgBox "An error occurred verifying the crew id."
    Else
        If m_rec.RecordCount = 0 Then  'New crew
            Me.Caption = "New Crew - " + crew_id.Text
            type_code = "C"
            m_rec.AddNew
            m_rec.Fields("crew_id") = crew_id.Text
            m_rec.Fields("type_code") = "C"
'            UpdateRecordsetFromForm Me, m_Rec
        Else
            Me.Caption = "Edit Crew - " + crew_id.Text
        End If
    End If
    UpdateFormFromRecordset Me, m_rec
End Sub
Private Sub PromptForSave(Cancel As Boolean)
    Dim iResult As Integer
    Dim blnPendingChange As Boolean
    Dim bln_New As Boolean
    Dim m_blnWereErrors As Boolean
    Dim strSaveCrewID As String
        Cancel = False
'        strSaveCrewID = crew_id.Text
'        crew_id.Text = m_strCrewID
        blnPendingChange = IsControlChanged(Me, m_rec)
'        crew_id.Text = strSaveCrewID
'        crew_id.Refresh
        If blnPendingChange = True Then
            iResult = MsgBox("Do you want to save your changes?", vbYesNoCancel)
            If iResult = vbYes Then
                m_blnWereErrors = False
                If m_blnWereErrors Then
                    Cancel = True
                Else
                    cmdUpdate_Click
                    ' If there were errors, cancel the close
                    If m_blnWereErrors Then
                        Cancel = True
                    End If
                End If
            ElseIf iResult = vbCancel Then
                Cancel = True
                Exit Sub
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
    If m_rec.Fields("crew_id") > "" Then
        varButton = MsgBox("Are you sure you want to delete this crew?" + Chr(13) + "All usage records will be deleted, as well.", vbYesNo + vbCritical)
    End If
    
    If varButton = vbYes Then
        ' Build SQL statement
        strUpdate = "exec sp_delete_crew "
        strUpdate = strUpdate + "@crew_skey ='" + CStr(crew_skey) + "'"
        
        blnRet = g_objDAL.ExecQuery(CONNECT, strUpdate, strError)
        If Not blnRet Then
            MsgBox strError
        Else
            MsgBox "Delete successful."
            Unload Me
        End If
    End If

End Sub

Private Sub cmdUpdate_Click()
Dim strUpdate As String
Dim strError As String
Dim blnReturn As Boolean

On Error Resume Next

If m_blnCloneCrew = True Then
    strUpdate = "exec sp_clone_crew @Crew_id = '" + crew_id.Text + "', "
Else
    strUpdate = "exec sp_update_crew @Crew_id = '" + m_strCrewID + "', "
End If
BuildStoredProcSQL Me, strUpdate, "1", m_rec
If m_blnCloneCrew = True Then
    strUpdate = ReplaceStr(strUpdate, "@crew_skey", "@from_crew_skey")
End If
strUpdate = strUpdate + " @last_update_person='" + strUserName + "'"

blnReturn = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
If strError <> "" Then
    MsgBox strError
Else
    MsgBox "Update successful."
    UpdateRecordsetFromForm Me, m_rec
End If

End Sub

Private Sub Refresh_Crew()
If crew_id.Text <> m_strCrewID Then
    PromptForSave m_blnWereErrors
    If m_blnWereErrors = False Then
        m_strCrewID = crew_id.Text
        LoadCrew
    End If
End If
End Sub


Private Sub crew_id_Click()
Dim intSaveListIndex As Integer

If crew_id.ListIndex <> -1 Then
    intSaveListIndex = crew_id.ListIndex
    If crew_id.List(intSaveListIndex) <> m_strCrewID And m_blnCloneCrew = False Then
        Refresh_Crew
        crew_id.ListIndex = intSaveListIndex
    End If
End If
End Sub

Private Sub crew_id_LostFocus()
If m_blnCloneCrew = False Then
    Refresh_Crew
End If
End Sub

Private Sub Form_Initialize()
Me.Height = 3735
Me.Width = 6975
Me.Top = 30
Me.Left = 2430
End Sub

Private Sub Form_Load()
Dim bResult As Boolean
bResult = LockField(Me, "type_code")
bResult = LockField(Me, "traces_crew_type_code")
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If Not UnloadMode = vbFormCode Then
        PromptForSave m_blnWereErrors
        Cancel = m_blnWereErrors
    End If

End Sub


