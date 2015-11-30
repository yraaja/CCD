VERSION 5.00
Begin VB.Form frmHierarchyTree 
   Caption         =   "Hierarchy Tree"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9225
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   9225
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   5160
      TabIndex        =   21
      Top             =   4080
      Width           =   1575
   End
   Begin VB.ComboBox level_id 
      Height          =   315
      ItemData        =   "frmHierarchyTree.frx":0000
      Left            =   1560
      List            =   "frmHierarchyTree.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Tag             =   "1N"
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert"
      Height          =   495
      Left            =   3360
      TabIndex        =   19
      Top             =   4080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Default         =   -1  'True
      Height          =   495
      Left            =   1560
      TabIndex        =   18
      Top             =   4080
      Width           =   1575
   End
   Begin VB.TextBox last_update_id 
      Height          =   285
      Left            =   1560
      MaxLength       =   8
      TabIndex        =   17
      Top             =   3600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox mf95_id 
      Height          =   285
      Left            =   1560
      MaxLength       =   8
      TabIndex        =   15
      Tag             =   "1S"
      Top             =   3240
      Width           =   1575
   End
   Begin VB.TextBox mat_id_end 
      Height          =   285
      Left            =   1560
      MaxLength       =   13
      TabIndex        =   13
      Tag             =   "1S"
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox mat_id_start 
      Height          =   285
      Left            =   1560
      MaxLength       =   13
      TabIndex        =   11
      Tag             =   "1S"
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox unit_cost_id_end 
      Height          =   285
      Left            =   1560
      MaxLength       =   12
      TabIndex        =   9
      Tag             =   "1S"
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox unit_cost_id_start 
      Height          =   285
      Left            =   1560
      MaxLength       =   12
      TabIndex        =   7
      Tag             =   "1S"
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox hier_desc 
      Height          =   285
      Left            =   1560
      MaxLength       =   75
      TabIndex        =   4
      Tag             =   "1S"
      Top             =   1440
      Width           =   7335
   End
   Begin VB.TextBox level_id1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7320
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   2
      Top             =   3360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox hier_id 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   1
      Tag             =   "1S"
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lblMode 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   9015
   End
   Begin VB.Label Label9 
      Caption         =   "Last Update Id"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "MF95 ID"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "Material Id End"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Material Id Start"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Unit Cost End"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Unit Cost Start"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Hier Desc"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Level Id"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Hier Id"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   615
   End
End
Attribute VB_Name = "frmHierarchyTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strDisplay As String
Public deleteMode As Integer

Dim m_rec As ADODB.RecordSet
Dim m_blnRecFlag As Boolean ' True if the screen has a data source, False if we are adding new record
Dim m_blnDeleted As Boolean ' Indicates if the data has been deleted, used in QueryUnload
Dim m_blnWereErrors As Boolean ' True if the Update had errors, used in QueryUnload
Dim m_blnInsert As Boolean
Dim m_blnChangesMade As Boolean
Dim m_blnLoading As Boolean

Public Sub InsertMode(ByVal blnIsert As Boolean)
    m_blnInsert = True
End Sub

Public Sub SetRow(rec As ADODB.RecordSet, Optional blnInsert As Boolean = False)
    Set m_rec = rec
End Sub

Private Sub cmdDelete_Click()
    Dim blnErrors As Boolean
    Dim errString As String
    errString = ""
    blnErrors = True

        If level_id = 1 Then
            MsgBox "Can't delete Level 1", vbExclamation
        Else
            
            Dim retYesNo As Integer
            retYesNo = MsgBox("Are you sure you want to delete the element hier_id='" + hier_id.Text + "' and level_id='" + level_id + "'?", vbQuestion + vbYesNo)
            If retYesNo = vbNo Then
                Exit Sub
            End If
            
            
            Dim strDelete As String
            
            strDelete = "exec sp_delete_masterformat04_id_hierarchy @hier_id='" + hier_id.Text + "',@level_id=" + level_id + ","
            strDelete = strDelete + "@last_update_person='" + strUserName + "'"
                

            
            Dim cmd As New ADODB.Command
            cmd.CommandText = strDelete
            cmd.CommandType = adCmdText
            
        
            ' Assuming a connection has been established and a recordset has
            '  created
            Set cmd.ActiveConnection = g_cnShared
            Dim RecordSet As ADODB.RecordSet
            Set RecordSet = cmd.Execute()
            Dim retIntStatus As Integer
            
            If Not (RecordSet.EOF And RecordSet.BOF) Then
                retIntStatus = RecordSet.Fields("retStatus")
            Else
                RecordSet.Close
                MsgBox "Could not Delete from Table!!", vbExclamation
                Exit Sub
            End If
            RecordSet.Close
            If retIntStatus = 4 Then
                MsgBox "Successfully Deleted from MASTERFORMAT04_ID_HIERARCHY and MASTERFORMAT04_LEVEL_LOOKUP", vbInformation
                strDisplay = "1" 'this will be passed back to the tree and will allow the code in the tree control to make the node invisible
            ElseIf retIntStatus = 3 Then 'levelId = 4
                MsgBox "Successfully Deleted from MASTERFORMAT04_ID_HIERARCHY", vbExclamation
                strDisplay = "1" 'this will be passed back to the tree and will allow the code in the tree control to make the node invisible
            ElseIf retIntStatus = 2 Then 'levelId <> 4
                MsgBox "Could not delete from MASTERFORMAT04_ID_HIERARCHY", vbExclamation
                strDisplay = "0" 'this will be passed back to the tree and will allow the code in the tree control to keep the node visible
            ElseIf retIntStatus = 1 Then
                MsgBox "Can't delete Level 1 Hierarchy Element", vbExclamation
                strDisplay = "0" 'this will be passed back to the tree and will allow the code in the tree control to keep the node visible
            ElseIf retIntStatus = 0 Then
                MsgBox "Combination of Hier Id ='" & hier_id.Text & "' and Level Id ='" & level_id & "' does not exists in MASTERFORMAT04_ID_HIERARCHY.", vbExclamation
                strDisplay = "0" 'this will be passed back to the tree and will allow the code in the tree control to keep the node visible
            End If
            Unload Me
        End If

End Sub

Private Sub cmdInsert_Click()

    Dim blnErrors As Boolean
    Dim errString As String
    errString = ""
    blnErrors = True

    
    If m_blnChangesMade = True Then
    
        
        If Not IsNumeric(hier_id.Text) Then
            errString = errString + "Hier Id has to be at numeric and at least 2 digits." + vbCrLf
        End If
        If InStr(1, hier_id.Text, ".") > 0 Then
            errString = errString + "Hier Id cannot have a demical point in it." + vbCrLf
        End If
        If Len(hier_id.Text) < 2 Then
            errString = errString + "Hier Id has to be at least 2 digits." + vbCrLf
        End If
            
        If level_id.Text = "" Then
            errString = errString + "Please choose a Level Id." + vbCrLf
        End If
        
        If level_id.Text = "1" Then
            'the length of hier_id should be 2
            If Len(hier_id) <> 2 Then
                errString = errString + "Hier Id has to be exactly 2 digits for Level 1." + vbCrLf
            End If
            If Len(mf95_id.Text) <> 6 Then
                errString = errString + "MF95 Id has be exactly 6 characters" + vbCrLf
            End If
            
        ElseIf level_id.Text = "2" Then
            'the length of hier_id should be 6
            If Len(hier_id) <> 6 Then
                errString = errString + "Hier Id has to be exactly 6 digits for Level 2." + vbCrLf
            End If
            If Len(mf95_id.Text) <> 5 Then
                errString = errString + "MF95 Id has be exactly 5 characters" + vbCrLf
            End If
        
        ElseIf level_id.Text = "3" Then
            'the length of hier_id should be 6
            If Len(hier_id) <> 6 Then
                errString = errString + "Hier Id has to be exactly 6 digits for Level 3." + vbCrLf
            End If
            If Len(mf95_id.Text) <> 5 Then
                errString = errString + "MF95 Id has be exactly 5 characters" + vbCrLf
            End If
        
        
        ElseIf level_id.Text = "4" Then
            'the length of hier_id should be 8
            If Len(hier_id) <> 8 Then
                errString = errString + "Hier Id has to be exactly 6 digits for Level 4." + vbCrLf
            End If
            
            If Len(mf95_id.Text) <> 8 Then
                errString = errString + "MF95 Id has be exactly 8 characters" + vbCrLf
            End If
            
        End If
        
        If Trim(hier_desc.Text) = "" Then
            errString = errString + "Hier Desc cannot be blank." + vbCrLf
        End If
        
        
        If errString <> "" Then
            MsgBox errString, vbExclamation
            Exit Sub
        End If
        
        blnErrors = False
        
        If blnErrors = False Then
        
            Dim strInsert As String
            
            strInsert = "exec sp_insert_masterformat04_id_hierarchy "
            BuildStoredProcSQL Me, strInsert, "1"
            strInsert = strInsert + "@last_update_person='" + strUserName + "'"
    
            
            Dim cmd As New ADODB.Command
            cmd.CommandText = strInsert
            cmd.CommandType = adCmdText
            
        
            ' Assuming a connection has been established and a recordset has
            '  created
            Set cmd.ActiveConnection = g_cnShared
            Dim RecordSet As ADODB.RecordSet
            Set RecordSet = cmd.Execute()
            Dim retIntStatus As Integer
            
            If Not (RecordSet.EOF And RecordSet.BOF) Then
                retIntStatus = RecordSet.Fields("retStatus")
            Else
                RecordSet.Close
                MsgBox "Could not Insert into Table!!", vbExclamation
                Exit Sub
            End If
            RecordSet.Close
            If retIntStatus = 4 Then
                MsgBox "Successfully Inserted into MASTERFORMAT04_ID_HIERARCHY and MASTERFORMAT04_LEVEL_LOOKUP", vbInformation
                strDisplay = hier_desc 'this will be passed back to the tree
            
            ElseIf retIntStatus = 3 Then 'levelId = 4
                MsgBox "Successfully Inserted into MASTERFORMAT04_ID_HIERARCHY only", vbExclamation
                strDisplay = hier_desc 'this will be passed back to the tree
            ElseIf retIntStatus = 2 Then 'levelId <> 4
                MsgBox "Successfully Inserted into MASTERFORMAT04_ID_HIERARCHY", vbInformation
                strDisplay = hier_desc 'this will be passed back to the tree
            ElseIf retIntStatus = 1 Then
                MsgBox "Could not insert into MASTERFORMAT04_ID_HIERARCHY", vbExclamation
            ElseIf retIntStatus = 0 Then
                MsgBox "Combination of Hier Id ='" & hier_id.Text & "' and Level Id ='" & level_id & "' already exists in MASTERFORMAT04_ID_HIERARCHY.", vbExclamation
            End If
            
        End If
    
    Else
        MsgBox "No changes were made.", vbInformation
    
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    m_blnLoading = True
    
    Dim ctr As Control
    Dim rec As ADODB.RecordSet
    Dim strSelect As String
    Dim blnReturn As Boolean
    Move (Screen.Width / 2) - (Me.Width / 2), (Screen.Height / 2) - (Me.Height / 2)
    
    If m_blnInsert = True Then
        hier_id.Enabled = True
        hier_id.Locked = False
        level_id.Enabled = True
        level_id.Locked = False
        cmdInsert.Visible = True
        cmdInsert.Default = True
        cmdUpdate.Visible = False
        cmdDelete.Visible = False
        lblMode.Caption = "Insert Mode"
    Else
        If deleteMode = 1 Then
            level_id.Enabled = False
            level_id.Locked = True
            hier_id.Enabled = False
            hier_id.Locked = True
            cmdUpdate.Visible = False
            cmdInsert.Visible = False
            cmdDelete.Visible = True
            cmdDelete.Default = True
            lblMode.Caption = "Delete Mode"
        Else
            level_id.Enabled = False
            level_id.Locked = True
            hier_id.Enabled = False
            hier_id.Locked = True
            cmdUpdate.Visible = True
            cmdUpdate.Default = True
            cmdInsert.Visible = False
            cmdDelete.Visible = False
            lblMode.Caption = "Update Mode"
        End If
        If Not m_rec.State = adStateClosed Then
            UpdateFormFromRecordset Me, m_rec
        End If
    End If
    
    m_blnLoading = False
    
End Sub


Private Sub cmdUpdate_Click()
    On Error Resume Next
    Dim strUpdate As String
    Dim strError As String
    Dim blnRet As Boolean
    Dim strSelect As String
    Dim rec As ADODB.RecordSet
    m_blnWereErrors = False
    hier_id = Compress_String(hier_id)
    
    Dim errString As String
    errString = ""
        
    If m_blnChangesMade = True Then
    
    
            
        If level_id.Text = "" Then
            errString = errString + "Please choose a Level Id"
        End If
        
    
        If level_id.Text = "1" Then
            'the length of hier_id should be 2
            If Len(hier_id) <> 2 Then
                errString = errString + "Hier Id has to be exactly 2 digits for Level 1." + vbCrLf
            End If
            If Len(mf95_id.Text) <> 6 Then
                errString = errString + "MF95 Id has be exactly 6 characters" + vbCrLf
            End If
            
        ElseIf level_id.Text = "2" Then
            'the length of hier_id should be 6
            If Len(hier_id) <> 6 Then
                errString = errString + "Hier Id has to be exactly 6 digits for Level 2." + vbCrLf
            End If
            If Len(mf95_id.Text) <> 5 Then
                errString = errString + "MF95 Id has be exactly 5 characters" + vbCrLf
            End If
        
        ElseIf level_id.Text = "3" Then
            'the length of hier_id should be 6
            If Len(hier_id) <> 6 Then
                errString = errString + "Hier Id has to be exactly 6 digits for Level 3." + vbCrLf
            End If
            If Len(mf95_id.Text) <> 5 Then
                errString = errString + "MF95 Id has be exactly 5 characters" + vbCrLf
            End If
        
        
        ElseIf level_id.Text = "4" Then
            'the length of hier_id should be 8
            If Len(hier_id) <> 8 Then
                errString = errString + "Hier Id has to be exactly 6 digits for Level 4." + vbCrLf
            End If
            
            If Len(mf95_id.Text) <> 8 Then
                errString = errString + "MF95 Id has be exactly 8 characters" + vbCrLf
            End If
            
        End If
        
        If Trim(hier_desc.Text) = "" Then
            errString = errString + "Hier Desc cannot be blank." + vbCrLf
        End If
        
        
        If errString <> "" Then
            MsgBox errString, vbExclamation
            Exit Sub
        End If
        
        
        ' If we are updating
        'If m_blnInsert = False Then
            strUpdate = "exec sp_update_masterformat04_id_hierarchy "
            BuildStoredProcSQL Me, strUpdate, "1"
            strUpdate = strUpdate + "@last_update_person='" + strUserName + "'"
            strUpdate = strUpdate + ", @last_update_id=" + CStr(last_update_id)
        ' If we are inserting
    '    Else
    '        strUpdate = "exec sp_insert_masterformat04_id_hierarchy "
    '        BuildStoredProcSQL Me, strUpdate, "1"
    '        strUpdate = strUpdate + "@last_update_person='" + strUserName + "'"
        'End If
        
        Dim cmd As New ADODB.Command
        cmd.CommandText = strUpdate
        cmd.CommandType = adCmdText
        
    
        ' Assuming a connection has been established and a recordset has
        '  created
        Set cmd.ActiveConnection = g_cnShared
        Dim RecordSet As ADODB.RecordSet
        Set RecordSet = cmd.Execute()
        Dim retIntStatus As Integer
        
        If Not (RecordSet.EOF And RecordSet.BOF) Then
            retIntStatus = RecordSet.Fields("retStatus")
        Else
            RecordSet.Close
            MsgBox "Could not update Table!!", vbExclamation
            Exit Sub
        End If
        RecordSet.Close
        If retIntStatus = 3 Then
            MsgBox "Successfully Updated MASTERFORMAT04_ID_HIERARCHY and MASTERFORMAT04_LEVEL_LOOKUP", vbInformation
            strDisplay = hier_desc 'this will be passed back to the tree
        
        ElseIf retIntStatus = 2 Then
            MsgBox "Successfully Updated MASTERFORMAT04_ID_HIERARCHY only", vbExclamation
            strDisplay = hier_desc 'this will be passed back to the tree
        
        ElseIf retIntStatus = 1 Then
            MsgBox "Could not updated MASTERFORMAT04_ID_HIERARCHY", vbExclamation
        ElseIf retIntStatus = 0 Then
            MsgBox "Last_Update_Id in MASTERFORMAT04_ID_HIERARCHY was changed since last update. Please try to update again", vbExclamation
        End If
    
    
        Unload Me
        'End If
    Else
        MsgBox "No changes were made", vbExclamation
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

'frmCallingForm.TreeView1.
End Sub



Private Sub appendString()
    If m_blnInsert = True Then
        unit_cost_id_start.Text = hier_id.Text + Mid("000000000000", Len(hier_id.Text) + 1)
        mat_id_start.Text = "M" + hier_id.Text + Mid("000000000000", Len(hier_id.Text) + 1)
        unit_cost_id_end.Text = hier_id.Text + Mid("999999999999", Len(hier_id.Text) + 1)
        mat_id_end.Text = "M" + hier_id.Text + Mid("999999999999", Len(hier_id.Text) + 1)
    
    End If
    

End Sub

Private Sub hier_desc_Change()
    If (m_blnLoading = False) Then
        m_blnChangesMade = True
    End If

End Sub

Private Sub hier_id_Change()
    appendString
    
    If (m_blnLoading = False) Then
        m_blnChangesMade = True
    End If
    
End Sub


Private Sub level_id_Change()
    If (m_blnLoading = False) Then
        m_blnChangesMade = True
    End If

End Sub

Private Sub mat_id_end_Change()
    If (m_blnLoading = False) Then
        m_blnChangesMade = True
    End If

End Sub

Private Sub mat_id_start_Change()
    If (m_blnLoading = False) Then
        m_blnChangesMade = True
    End If

End Sub

Private Sub mf95_id_Change()
    If (m_blnLoading = False) Then
        m_blnChangesMade = True
    End If

End Sub

Private Sub unit_cost_id_end_Change()
    If (m_blnLoading = False) Then
        m_blnChangesMade = True
    End If

End Sub

Private Sub unit_cost_id_start_Change()
    If (m_blnLoading = False) Then
        m_blnChangesMade = True
    End If

End Sub
