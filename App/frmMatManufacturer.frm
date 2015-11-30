VERSION 5.00
Begin VB.Form frmMatManufacturer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Material Manufacturer"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   Icon            =   "frmMatManufacturer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1905
   ScaleWidth      =   6360
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
      Left            =   5100
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   1260
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox manufacturer_desc 
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
      Left            =   3780
      MaxLength       =   35
      TabIndex        =   1
      Tag             =   "1S"
      Top             =   180
      Width           =   2415
   End
   Begin VB.TextBox manufacturer_id 
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
      Left            =   1380
      TabIndex        =   0
      Tag             =   "1S"
      Top             =   180
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   1260
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
      Enabled         =   0   'False
      Height          =   315
      Left            =   1380
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   660
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
      Enabled         =   0   'False
      Height          =   315
      Left            =   4980
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   660
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   3360
      TabIndex        =   5
      Top             =   1260
      Width           =   1150
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Name:"
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Top             =   240
      Width           =   915
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Manufacturer ID:"
      Height          =   255
      Left            =   60
      TabIndex        =   8
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Last Updated:"
      Height          =   255
      Left            =   180
      TabIndex        =   7
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Last Updated By:"
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   720
      Width           =   1395
   End
End
Attribute VB_Name = "frmMatManufacturer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_rec As ADODB.RecordSet
Dim m_blnRecFlag As Boolean ' True if the screen has a data source, False if we are adding new record
Dim m_blnWereErrors As Boolean ' True if the Update had errors, used in QueryUnload
Dim m_blnInsert As Boolean ' Tells if we are doing an insert or update
Dim m_blnDeleted As Boolean ' Indicates if the data has been deleted, used in QueryUnload

' Fills all fields with data
Public Sub SetRow(ByRef rec As ADODB.RecordSet, Optional blnInsert As Boolean = False)
    Set m_rec = rec
    m_blnInsert = blnInsert
    ' If we are inserting/cloning
    If m_blnInsert Then
        ' Do this so OriginalValue will be set to the values copied into the row
        m_rec.UpdateBatch
    End If
    If Not m_rec.Fields("manufacturer_id") = "" Then
        m_blnRecFlag = True
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
    
    strUpdate = "exec sp_delete_material_manufacture "
    strUpdate = strUpdate + "@manufacturer_id='" + Me.Controls("manufacturer_id") + "'"
    
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
    Dim strTempMat As String
    Dim strUpdate As String
    Dim strError As String
    Dim blnRet As Boolean
    
    m_blnWereErrors = False
    
    ' If we are updating
    If m_blnInsert = False Then
        strUpdate = "exec sp_update_material_manufacture @last_update_id=" + last_update_id.Text + ", "
        BuildStoredProcSQL Me, strUpdate, "1"
        strUpdate = strUpdate + " @last_update_person='" + strUserName + "'"
    ' If we are inserting
    Else
        strUpdate = "exec sp_insert_material_manufacture "
        BuildStoredProcSQL Me, strUpdate, "1"
        strUpdate = strUpdate + " @last_update_person='" + strUserName + "'"
    End If
    
    blnRet = g_objDAL.ExecQuery(CONNECT, strUpdate, strError)
    If Not blnRet Then
        MsgBox strError
        m_blnWereErrors = True
    Else
        ' Put latest data into source recordset
        UpdateRecordsetFromForm Me, m_rec
        m_rec.Fields("last_update_id").Value = m_rec.Fields("last_update_id").Value + 1
        last_update_id.Text = m_rec.Fields("last_update_id").Value
        m_rec.Fields("last_update_person").Value = strUserName
        last_update_person.Text = strUserName
        m_rec.Fields("last_update_date").Value = Now
        last_update_date.Text = Now
        UpdateFormFromRecordset Me, m_rec
        MsgBox "Update successful."
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
    Dim ctr As Control
    
    On Error Resume Next
    Move START_LEFT, START_TOP
    
    If m_blnInsert = False Then
        Me.Caption = Me.Caption + " [" + m_rec.Fields("manufacturer_id").Value + "]"
    ElseIf m_blnRecFlag = True Then
        Me.Caption = Me.Caption + " [Clone of " + m_rec.Fields("manufacturer_id").Value + "]"
    Else
        Me.Caption = Me.Caption + " [New]"
    End If
    
    ' If we are showing data
'    If m_blnRecFlag = True Then '
        If Not m_rec.State = adStateClosed Then
            UpdateFormFromRecordset Me, m_rec
        End If
        ' Lock fields that can't be changed
        If m_blnInsert = False Then
            manufacturer_id.Locked = True
            manufacturer_id.BackColor = LTGREY
        End If
'    Else
'        ' Set all controls to blanks
'        ' Loop through all controls on form
'        For Each ctr In Me.Controls
'            ' Check type of control
'            If TypeOf ctr Is TextBox Then
'                ctr = ""
'            ElseIf TypeOf ctr Is CheckBox Then
'                ctr = 0
'            End If
'        Next ctr
'    End If
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
            ElseIf m_blnInsert = True Then
                m_rec.Delete
            End If
        End If
    End If
End Sub

Private Sub Form_Resize()
ResizeForm Me
End Sub


