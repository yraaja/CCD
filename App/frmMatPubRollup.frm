VERSION 5.00
Begin VB.Form frmMatPubRollup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Published Material Rollup"
   ClientHeight    =   1800
   ClientLeft      =   3900
   ClientTop       =   5070
   ClientWidth     =   4095
   Icon            =   "frmMatPubRollup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   4095
   Begin VB.TextBox processed_date 
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Tag             =   "0D"
      Top             =   2520
      Width           =   3000
   End
   Begin VB.TextBox last_update_id 
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   1680
      TabIndex        =   9
      TabStop         =   0   'False
      Tag             =   "0N"
      Top             =   2880
      Width           =   1320
   End
   Begin VB.TextBox last_update_date 
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Tag             =   "0D"
      Top             =   3360
      Width           =   3000
   End
   Begin VB.CheckBox update_ind 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   1500
      TabIndex        =   1
      Tag             =   "0"
      Top             =   600
      Width           =   215
   End
   Begin VB.TextBox mat_skey 
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Tag             =   "0N"
      Top             =   2040
      Width           =   3000
   End
   Begin VB.TextBox mat_id 
      Height          =   315
      Left            =   1500
      MaxLength       =   13
      TabIndex        =   0
      Tag             =   "0S"
      Top             =   120
      Width           =   2160
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   495
      Left            =   540
      TabIndex        =   2
      Top             =   1140
      Width           =   1150
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   1980
      TabIndex        =   3
      Top             =   1140
      Width           =   1150
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Process Date"
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   13
      Top             =   2520
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Update ID"
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   12
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Last Update Date"
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   11
      Top             =   3360
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Update Indicator:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   1230
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Material Skey:"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Material ID:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   810
   End
End
Attribute VB_Name = "frmMatPubRollup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_rec As ADODB.RecordSet
Dim m_blnRecFlag As Boolean ' True if a populated RecordSet was passed, then we show data
Dim m_blnWereErrors As Boolean ' True if the Update had errors, used in QueryUnload
Dim m_blnInsert As Boolean ' Tells if we are doing an insert or update
Dim strLast_mat_id As String ' Holds last mat_id so we know if it changed

' Fills all fields with data
Public Sub SetRow(rec As ADODB.RecordSet, Optional blnInsert As Boolean = False)
    Set m_rec = rec
    m_blnInsert = blnInsert
    ' If we are inserting/cloning
    If m_blnInsert Then
        ' Do this so OriginalValue will be set to the values copied into the row
        m_rec.UpdateBatch
    End If
    If Not m_rec.Fields("mat_skey") = 0 Then
        m_blnRecFlag = True
    End If
End Sub

Private Sub cmdDelete_Click()
    On Error Resume Next
    Dim strUpdate As String
    Dim blnRet As Boolean
    Dim strError As String
    Dim oldRow As Long
    Dim varButton
    varButton = MsgBox("Are you sure you want to delete?", vbYesNo + vbCritical)
    If varButton = vbNo Then
        Exit Sub
    End If

    strUpdate = "exec sp_delete_mat_pub_rollup "
    strUpdate = strUpdate + " @mat_skey=" + str(Me.Controls("mat_skey"))
    
    blnRet = g_objDAL.ExecQuery(CONNECT, strUpdate, strError)
    If Not blnRet Then
        MsgBox strError
    Else
        m_rec.Delete
'        If Len(frmMatPubRollupGrid.MaterialID) > 0 Then
'            OldRow = frmMatPubRollupGrid.TDBGrid.Row
'            frmMatPubRollupGrid.cmdSearch.Value = True
'            frmMatPubRollupGrid.Refresh
'            frmMatPubRollupGrid.TDBGrid.Row = OldRow
'        End If
        MsgBox "Delete successful."
        Unload Me
    End If
End Sub

Private Sub cmdUpdate_Click()
    On Error Resume Next
    Dim strUpdate As String
    Dim strError As String
    Dim blnRet As Boolean
    Dim oldRow As Long
    
    m_blnWereErrors = False
    
    ' If we are updating
    If m_blnRecFlag = True Then
        strUpdate = "exec sp_update_mat_pub_rollup "
        BuildStoredProcSQL Me, strUpdate, "0"
        If Len(processed_date) <= 0 Then strUpdate = strUpdate + " @processed_date= Null, "
        strUpdate = strUpdate + " @last_update_person='" + strUserName + "'"
    ' If we are inserting
    Else
        If MaterialIdValidate = False Then Exit Sub
        last_update_date.Text = Now()
        strUpdate = "exec sp_insert_mat_pub_rollup "
        BuildStoredProcSQL Me, strUpdate, "0"
        strUpdate = strUpdate + " @last_update_person='" + strUserName + "'"
    End If
    
    blnRet = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
    If Not blnRet Then
        MsgBox strError
        m_blnWereErrors = True
    Else
        ' Put latest data into source recordset
        UpdateRecordsetFromForm Me, m_rec
        m_rec.Fields("last_update_id").Value = m_rec.Fields("last_update_id").Value + 1
        last_update_id.Text = m_rec.Fields("last_update_id").Value
        UpdateFormFromRecordset Me, m_rec
'        If Len(frmMatPubRollupGrid.MaterialID) > 0 Then
'            OldRow = frmMatPubRollupGrid.TDBGrid.Row
'            frmMatPubRollupGrid.cmdSearch.Value = True
'            frmMatPubRollupGrid.Refresh
'            frmMatPubRollupGrid.TDBGrid.Row = OldRow
'        End If
        MsgBox "Update successful.", vbInformation
        mat_id.SetFocus
        If m_blnRecFlag = True Then Unload Me
    End If
End Sub

Private Sub Form_Activate()
    If mat_id.Enabled = False Then update_ind.SetFocus
    OutputView False
End Sub

Private Sub Form_Initialize()
    m_blnInsert = False
    Set m_rec = Nothing
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim ctr As Control
    Dim rec As ADODB.RecordSet
    'Move START_LEFT, START_TOP
    'changed here so that the form is position in a different location
    Me.top = 5190
    Me.left = 5000
    strLast_mat_id = ""
    
    ' Load data into form
    If Not m_rec.State = adStateClosed Then
        UpdateFormFromRecordset Me, m_rec
    End If
    
    strLast_mat_id = m_rec.Fields("mat_id").Value
    
   ' If we are NOT inserting
    If m_blnInsert = False Then
        ' Lock fields that can't be changed
        mat_id.Locked = True
        mat_id.BackColor = LTGREY
        ' Set caption
        Me.Caption = Me.Caption + " [" + m_rec.Fields("mat_id").Value + "]"
        cmdDelete.Enabled = True
    Else
        ' If we are inserting and not showing data
        ' Set some defaults
        If Not m_blnRecFlag Then
            'active_status_ind.Value = 1
            m_rec.Fields("active_status_ind").Value = True
            'use_ind.Value = 1
            m_rec.Fields("use_ind").Value = True
            'pct_multiplier.Text = 100
            m_rec.Fields("pct_multiplier").Value = 100
            'purchase_usage_conv_factor.Text = 1
            m_rec.Fields("purchase_usage_conv_factor").Value = 1
            'wst_use_ind.Value = 0
            m_rec.Fields("wst_use_ind").Value = 0
            'estimated_ind.Value = 0
            m_rec.Fields("estimated_ind").Value = 0
            'traces_ind.Value = 0
            m_rec.Fields("traces_ind").Value = 0
            'update_ind.Value = 0
            m_rec.Fields("update_ind").Value = 0
            'factor_ind.Value = 0
            m_rec.Fields("factor_ind").Value = 0
            ' This is a new record
            Me.Caption = Me.Caption + " [New]"
            cmdDelete.Enabled = False
        Else
            ' This means we are cloning
            Me.Caption = Me.Caption + " [Clone of " + m_rec.Fields("mat_id").Value + "]"
        End If
    End If
End Sub

Private Sub list_price_Validate(Cancel As Boolean)
    'CheckValueForNumber list_price.Text, Cancel
End Sub

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
End Sub

Private Function MaterialIdValidate() As Boolean
    Dim rec As ADODB.RecordSet
    MaterialIdValidate = False
    If Not Len(mat_id.Text) = 0 Then
        If UCase(left(mat_id, 1)) <> "M" Then mat_id = "M" + mat_id
        mat_id = UCase(mat_id)
        g_objDAL.GetRecordset vbNullString, "Select mat_id from published_material_rollup where mat_id = '" + mat_id.Text + "'", rec
        If rec.RecordCount > 0 Then
            MsgBox "You can enter Material ID only once in Rollup", vbInformation
            mat_id.SetFocus
        Else
            rec.Close
            g_objDAL.GetRecordset vbNullString, "Select mat_skey from material where mat_id = '" + mat_id.Text + "'", rec
            If rec.RecordCount > 0 Then
                If rec.RecordCount > 0 Then
                    mat_skey.Text = rec.Fields("mat_skey")
                    MaterialIdValidate = True
                End If
            Else
                MsgBox "You must enter a Valid Material ID."
                mat_id.SetFocus
            End If
        End If
        rec.Close
    End If
End Function

Private Sub Form_Resize()
ResizeForm Me
End Sub

Private Sub mat_id_GotFocus()
    mat_id.SelStart = 1
    mat_id.SelLength = Len(mat_id)
End Sub
