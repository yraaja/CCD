VERSION 5.00
Begin VB.Form frmMaterial 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Material Maintenance"
   ClientHeight    =   4080
   ClientLeft      =   2025
   ClientTop       =   3435
   ClientWidth     =   8235
   Icon            =   "frmMaterial.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4080
   ScaleWidth      =   8235
   Begin VB.TextBox alt_mat_id 
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
      Left            =   4020
      TabIndex        =   1
      Tag             =   "1S"
      Top             =   180
      Width           =   1515
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
      Left            =   6900
      Locked          =   -1  'True
      TabIndex        =   21
      Tag             =   "N"
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox purchase_usage_conv_factor 
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
      Left            =   4560
      TabIndex        =   6
      Tag             =   "1N"
      Top             =   2220
      Width           =   675
   End
   Begin VB.ComboBox purchase_unit 
      Height          =   315
      ItemData        =   "frmMaterial.frx":0442
      Left            =   1260
      List            =   "frmMaterial.frx":0444
      TabIndex        =   5
      Tag             =   "1"
      Top             =   2220
      Width           =   1215
   End
   Begin VB.ComboBox usage_unit 
      Height          =   315
      Left            =   6900
      TabIndex        =   7
      Tag             =   "1"
      Top             =   2220
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   4260
      TabIndex        =   9
      Top             =   3420
      Width           =   1150
   End
   Begin VB.TextBox mat_skey 
      BackColor       =   &H8000000F&
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
      Left            =   6900
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Tag             =   "1N"
      Top             =   2820
      Width           =   1215
   End
   Begin VB.TextBox last_update_person 
      BackColor       =   &H8000000F&
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
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2820
      Width           =   1190
   End
   Begin VB.TextBox last_update_date 
      BackColor       =   &H8000000F&
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
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2820
      Width           =   1635
   End
   Begin VB.TextBox mat_id 
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
      Left            =   1260
      TabIndex        =   0
      Tag             =   "1S"
      Top             =   180
      Width           =   1515
   End
   Begin VB.CheckBox active_status_ind 
      Caption         =   "Active"
      Height          =   315
      Left            =   6120
      TabIndex        =   2
      Tag             =   "1"
      Top             =   180
      Width           =   975
   End
   Begin VB.TextBox tech_desc 
      Height          =   315
      Left            =   1260
      MaxLength       =   75
      TabIndex        =   3
      Tag             =   "1S"
      Top             =   780
      Width           =   6855
   End
   Begin VB.TextBox metric_tech_desc 
      Height          =   315
      Left            =   1260
      MaxLength       =   75
      TabIndex        =   4
      Tag             =   "1S"
      Top             =   1560
      Width           =   6855
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   495
      Left            =   2820
      TabIndex        =   8
      Top             =   3420
      Width           =   1150
   End
   Begin VB.Label Label31 
      Alignment       =   1  'Right Justify
      Caption         =   "Material Skey:"
      Height          =   255
      Left            =   5800
      TabIndex        =   23
      Top             =   2880
      Width           =   1035
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Alt Mat ID:"
      Height          =   255
      Left            =   3000
      TabIndex        =   22
      Top             =   240
      Width           =   915
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Purchase Unit:"
      Height          =   255
      Left            =   60
      TabIndex        =   20
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Conversion Factor:"
      Height          =   255
      Left            =   3060
      TabIndex        =   19
      Top             =   2280
      Width           =   1395
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Usage Unit:"
      Height          =   255
      Left            =   5880
      TabIndex        =   18
      Top             =   2280
      Width           =   915
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Last Updated By:"
      Height          =   255
      Left            =   3000
      TabIndex        =   16
      Top             =   2880
      Width           =   1395
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Last Updated:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Material ID:"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   240
      Width           =   915
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Tech Desc:"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   840
      Width           =   915
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Metric Tech Desc:"
      Height          =   435
      Left            =   120
      TabIndex        =   10
      Top             =   1500
      Width           =   1035
   End
End
Attribute VB_Name = "frmMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_rec As adodb.RecordSet
Dim m_blnRecFlag As Boolean ' True if the screen has a data source, False if we are adding new record
Dim m_blnDeleted As Boolean ' Indicates if the data has been deleted, used in QueryUnload
Dim m_blnWereErrors As Boolean ' True if the Update had errors, used in QueryUnload
Dim m_blnInsert As Boolean
' Fills all fields with data
Public Sub SetRow(rec As adodb.RecordSet, Optional blnInsert As Boolean = False)
Dim sMatID As String
    Set m_rec = rec
    m_blnInsert = blnInsert
    ' If we are inserting/cloning
    If m_blnInsert Then
        ' Do this so OriginalValue will be set to the values copied into the row
        sMatID = m_rec.Fields("mat_id")     'Force changed value for mat id
        m_rec.Fields("mat_id") = ""
        m_rec.UpdateBatch
        m_rec.Fields("mat_id") = sMatID
    End If
    If Not m_rec.Fields("mat_skey") = "" Then
        m_blnRecFlag = True
    End If
End Sub

Private Sub alt_mat_id_Change()
    ' CORRECTED 6/22/2005 RTD
    ' PROBLEM REPORTED BY TOM DION/BARBARA BALBONI ON 6/21/2005
    Dim intSelStart As Integer
    Dim intSelLength As Integer

    If alt_mat_id.Text <> "" Then
        intSelStart = alt_mat_id.SelStart
        intSelLength = alt_mat_id.SelLength
        If UCase(Left(alt_mat_id, 1)) <> "M" Then
            alt_mat_id.Text = "M" + alt_mat_id.Text
            intSelStart = intSelStart + 1
        Else
            alt_mat_id.Text = UCase(alt_mat_id.Text)
        End If
        alt_mat_id.SelStart = intSelStart
        alt_mat_id.SelLength = intSelLength
    End If

End Sub

Private Sub alt_mat_id_GotFocus()
    SendKeys ("{HOME}+{END}")
End Sub

Private Sub alt_mat_id_Validate(Cancel As Boolean)
If Invalid_mat_id_Format(Compress_String(alt_mat_id), "alt_mat_id", m_rec) = True Then
    Cancel = True
End If
End Sub
Private Sub cmdDelete_Click()
    On Error Resume Next
    Dim strUpdate As String
    Dim blnRet As Boolean
    Dim strError As String
   
    Dim varButton
    varButton = MsgBox("Are you sure you want to delete?", vbYesNo + vbCritical)
    If varButton = vbNo Then
        Exit Sub
    End If
    
    strUpdate = "exec sp_delete_material "
    strUpdate = strUpdate + "@mat_skey=" + str(Me.Controls("mat_skey")) + ", "
    strUpdate = strUpdate + "@last_update_person='" + strUserName + "'"
    
    blnRet = g_objDAL.ExecQuery(CONNECT, strUpdate, strError)
    If Not blnRet Then
        MsgBox strError
    Else
        MsgBox "Delete successful."
        m_rec.Delete
        m_blnDeleted = True
        Unload Me
    End If
End Sub

Private Sub cmdUpdate_Click()
    On Error Resume Next
    Dim strUpdate As String
    Dim strError As String
    Dim blnRet As Boolean
    Dim strSelect As String
    Dim rec As adodb.RecordSet
    m_blnWereErrors = False
    mat_id = Compress_String(mat_id)
    
    ' If we are updating
    If m_blnInsert = False Then
        strUpdate = "exec sp_update_material "
        BuildStoredProcSQL Me, strUpdate, "1"
        strUpdate = strUpdate + "@last_update_person='" + strUserName + "'"
        strUpdate = strUpdate + ", @last_update_id=" + CStr(last_update_id)
    ' If we are inserting
    Else
        strUpdate = "exec sp_insert_material "
        BuildStoredProcSQL Me, strUpdate, "1"
        strUpdate = strUpdate + "@last_update_person='" + strUserName + "'"
    End If
    
    blnRet = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
    If Not blnRet Then
        MsgBox strError
        m_blnWereErrors = True
    Else
        If m_blnInsert = True Then
            strSelect = "select mat_skey from material where mat_id = '" & mat_id.Text & "'"
            g_objDAL.GetRecordset vbNullString, strSelect, rec
            If (rec.EOF And rec.BOF) Then
                MsgBox "Record not added."
                m_blnWereErrors = True
                Exit Sub
            Else
                mat_skey.Text = rec.Fields("mat_skey").Value
            End If
            m_blnInsert = False
        End If
        ' Put latest data into source recordset
        UpdateRecordsetFromForm Me, m_rec
        m_rec.Fields("last_update_id").Value = m_rec.Fields("last_update_id").Value + 1
        UpdateFormFromRecordset Me, m_rec
        'MsgBox "Update successful."
        
        'code added by mohan Jan 18, 2012: update the Hierarchy tree
        Dim retBlnVal As Boolean
        retBlnVal = MainModule.Update_Tree_With_Unit_Cost_Id(mat_id.Text, alt_mat_id.Text)
        
        Screen.MousePointer = vbNormal
        
        If retBlnVal = False Then
            MsgBox "Update successful for Material, but there was an error while updating the Tree.", vbExclamation + vbOKOnly
        Else
            MsgBox "Update successful.", vbInformation + vbOKOnly
        End If
        
    End If
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
    Dim rec As adodb.RecordSet
    Dim strSelect As String
    Dim blnReturn As Boolean
    Move START_LEFT, START_TOP

    g_objDAL.GetRecordset CONNECT, "select unit from unit_of_measure order by unit", rec
    While Not rec.EOF
        purchase_unit.AddItem (rec.Fields("unit").Value)
        usage_unit.AddItem (rec.Fields("unit").Value)
        rec.MoveNext
    Wend
    rec.Close
    strSelect = "select count(*) as NbrMatsUsed from material_usage as mu where mat_skey = " + CStr(m_rec.Fields("mat_skey").Value)
    g_objDAL.GetRecordset CONNECT, strSelect, rec
    If Not rec.EOF Then
        If rec.Fields("NbrMatsUsed") > 0 Then
            blnReturn = LockField(Me, "active_status_ind")
        End If
    End If
    rec.Close

    If Not m_rec.State = adStateClosed Then
        UpdateFormFromRecordset Me, m_rec
    End If
    
    ' Lock fields that can't be changed
    'Line of code was changed by Mohan on Jan 05,2012: changed FORMAT_MATERIAL_SRV to FORMAT_MATERIAL_04_SRV
    mat_id = Format(Compress_String(mat_id), FORMAT_MATERIAL_04_SRV)
    If m_blnInsert = False Then
        Me.Caption = Me.Caption + " [" + mat_id.Text + "]"
        mat_id.Locked = True
        mat_id.BackColor = LTGREY
    Else
        active_status_ind.Value = 1
        m_rec.Fields("active_status_ind").Value = True
        blnReturn = LockField(Me, "active_status_ind")
        blnReturn = LockField(Me, "update_status_code")
        If m_blnRecFlag Then
            ' This means we are cloning
            Me.Caption = Me.Caption + " [Clone of " + mat_id.Text + "]"
        End If
    End If
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
                If Invalid_mat_id_Format(alt_mat_id, "alt_mat_id", m_rec) = True Then
                    m_blnWereErrors = True
                End If
                If Invalid_mat_id_Format(Compress_String(mat_id), "mat_id", m_rec) = True Then
                    m_blnWereErrors = True
                End If
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
End Sub

Private Sub Form_Resize()
    ResizeForm Me
End Sub

Private Sub mat_id_Change()
    ' CORRECTED 6/21/2005 RTD
    ' PROBLEM REPORTED BY TOM DION/BARBARA BALBONI ON 6/21/2005
    Dim intSelStart As Integer
    Dim intSelLength As Integer
    
    intSelStart = mat_id.SelStart
    intSelLength = mat_id.SelLength
    If UCase(Left(mat_id.Text, 1)) <> "M" Then
        mat_id.Text = "M" + mat_id.Text
        intSelStart = intSelStart + 1
    Else
        mat_id.Text = UCase(mat_id.Text)
    End If
    mat_id.SelStart = intSelStart
    mat_id.SelLength = intSelLength
    
End Sub

Private Sub mat_id_GotFocus()
    SendKeys ("{HOME}+{END}")
End Sub

Private Sub mat_id_Validate(Cancel As Boolean)
'Dim blnError As Boolean
'Dim strErrorDesc As String
'Dim strSelect As String
'Dim blnReturn As Boolean
'Dim rec As New ADODB.RecordSet
'
'If UCase(left(mat_id, 1)) <> "M" Then
'    If Not IsNumeric(mat_id) Then
'        strErrorDesc = "Please enter a valid Material - (M + 10 numbers)"
'        blnError = True
'    End If
'Else
'    If Not IsNumeric(right(mat_id, Len(mat_id) - 1)) Then
'        strErrorDesc = "Please enter a valid Material - (M + 10 numbers)"
'        blnError = True
'    End If
'End If
'
'
'
'If blnError = False And Len(mat_id) = 11 Then    'Check for duplicate
'    strSelect = "Select mat_id, mat_skey from Material where mat_id='" + mat_id.Text + "'"
'    ' Use DAL to perform select
'    blnReturn = g_objDAL.GetRecordset(vbNullString, strSelect, rec)
'    If rec.RecordCount > 0 Then
'        strErrorDesc = "The material already exists and may not be added."
'        blnError = True
'    End If
'    rec.Close
'    Set rec = Nothing
'Else
'    strErrorDesc = "Please enter a valid Material - (M + 10 numbers)"
'    blnError = True
'End If
'
'If blnError = True Then
'    Beep
'    MsgBox strErrorDesc
'    Cancel = True
'End If
If Invalid_mat_id_Format(Compress_String(mat_id), "mat_id", m_rec) = True Then
    Cancel = True
End If

End Sub



Private Sub metric_tech_desc_GotFocus()
    SendKeys ("{HOME}+{END}")
End Sub

Private Sub purchase_usage_conv_factor_GotFocus()
    SendKeys ("{HOME}+{END}")
End Sub

Private Sub purchase_usage_conv_factor_KeyPress(KeyAscii As Integer)
    If CheckNumericField(purchase_usage_conv_factor, KeyAscii, purchase_usage_conv_factor.SelStart, purchase_usage_conv_factor.SelLength, 5) = False Then
        KeyAscii = 0
    End If
End Sub
Private Sub purchase_usage_conv_factor_Validate(Cancel As Boolean)
    CheckValueForNumber purchase_usage_conv_factor.Text, Cancel
End Sub

Private Sub tech_desc_GotFocus()
    SendKeys ("{HOME}+{END}")
End Sub

Private Sub usage_unit_GotFocus()
    SendKeys ("{HOME}+{END}")
End Sub
