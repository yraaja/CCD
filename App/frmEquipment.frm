VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmEquipment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Equipment Maintenance"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8610
   Icon            =   "frmEquipment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   8610
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   495
      Left            =   3000
      TabIndex        =   23
      Top             =   4140
      Width           =   1150
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   4440
      TabIndex        =   24
      Top             =   4140
      Width           =   1150
   End
   Begin VB.TextBox equip_skey 
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
      TabIndex        =   31
      TabStop         =   0   'False
      Tag             =   "N"
      Top             =   3720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox equip_last_update_id 
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
      TabIndex        =   30
      TabStop         =   0   'False
      Tag             =   "N"
      Top             =   4080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox equip_last_update_person 
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
      Left            =   4500
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox equip_last_update_date 
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
      Left            =   1380
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   3720
      Width           =   1635
   End
   Begin VB.ComboBox type_code 
      BackColor       =   &H8000000D&
      ForeColor       =   &H8000000E&
      Height          =   315
      ItemData        =   "frmEquipment.frx":0442
      Left            =   6960
      List            =   "frmEquipment.frx":0452
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Tag             =   "1S"
      Top             =   60
      Width           =   1215
   End
   Begin VB.TextBox equip_id 
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
      Left            =   1200
      TabIndex        =   0
      Tag             =   "1S"
      Top             =   60
      Width           =   1215
   End
   Begin VB.TextBox alt_equip_id 
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
      Left            =   4080
      TabIndex        =   1
      Tag             =   "1S"
      Top             =   60
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   3075
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   5424
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Equipment"
      TabPicture(0)   =   "frmEquipment.frx":0462
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label33"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label32"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label31"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label30"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label28"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "traces_ind"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "model_name"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "index_code"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "index_desc"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "unit"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "metric_unit"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Frame2"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Descriptions"
      TabPicture(1)   =   "frmEquipment.frx":047E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label13"
      Tab(1).Control(1)=   "Label12"
      Tab(1).Control(2)=   "Label11"
      Tab(1).Control(3)=   "Label10"
      Tab(1).Control(4)=   "crew_equip_desc_plural"
      Tab(1).Control(5)=   "crew_equip_desc"
      Tab(1).Control(6)=   "book_desc"
      Tab(1).Control(7)=   "tech_desc"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Metric Descriptions"
      TabPicture(2)   =   "frmEquipment.frx":049A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label26"
      Tab(2).Control(1)=   "Label25"
      Tab(2).Control(2)=   "Label2"
      Tab(2).Control(3)=   "Label14"
      Tab(2).Control(4)=   "metric_tech_desc"
      Tab(2).Control(5)=   "metric_book_desc"
      Tab(2).Control(6)=   "metric_crew_equip_desc"
      Tab(2).Control(7)=   "metric_crew_equip_desc_plural"
      Tab(2).ControlCount=   8
      Begin VB.TextBox tech_desc 
         Height          =   495
         Left            =   -73740
         MaxLength       =   75
         MultiLine       =   -1  'True
         TabIndex        =   13
         Tag             =   "1S"
         Top             =   540
         Width           =   6855
      End
      Begin VB.TextBox book_desc 
         Height          =   495
         Left            =   -73740
         MaxLength       =   75
         MultiLine       =   -1  'True
         TabIndex        =   14
         Tag             =   "1S"
         Top             =   1140
         Width           =   6855
      End
      Begin VB.TextBox crew_equip_desc 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -73740
         MaxLength       =   35
         TabIndex        =   15
         Tag             =   "1S"
         Top             =   1740
         Width           =   2535
      End
      Begin VB.TextBox crew_equip_desc_plural 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -69480
         MaxLength       =   35
         TabIndex        =   16
         Tag             =   "1S"
         Top             =   1740
         Width           =   2595
      End
      Begin VB.TextBox metric_crew_equip_desc_plural 
         Height          =   315
         Left            =   -69480
         MaxLength       =   35
         TabIndex        =   20
         Tag             =   "1S"
         Top             =   1740
         Width           =   2595
      End
      Begin VB.TextBox metric_crew_equip_desc 
         Height          =   315
         Left            =   -73740
         MaxLength       =   35
         TabIndex        =   19
         Tag             =   "1S"
         Top             =   1740
         Width           =   2535
      End
      Begin VB.TextBox metric_book_desc 
         Height          =   495
         Left            =   -73740
         MaxLength       =   75
         MultiLine       =   -1  'True
         TabIndex        =   18
         Tag             =   "1S"
         Top             =   1140
         Width           =   6855
      End
      Begin VB.TextBox metric_tech_desc 
         Height          =   495
         Left            =   -73740
         MaxLength       =   75
         MultiLine       =   -1  'True
         TabIndex        =   17
         Tag             =   "1S"
         Top             =   540
         Width           =   6855
      End
      Begin VB.Frame Frame2 
         Caption         =   "Formatting"
         Height          =   675
         Left            =   360
         TabIndex        =   44
         Top             =   1560
         Width           =   7755
         Begin VB.TextBox format_characters 
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
            Left            =   3660
            TabIndex        =   11
            Tag             =   "1N"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox format_code 
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
            Left            =   6060
            TabIndex        =   12
            Tag             =   "1S"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox indent_code 
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
            TabIndex        =   10
            Tag             =   "1N"
            Top             =   240
            Width           =   435
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            Caption         =   "Indent Code:"
            Height          =   255
            Left            =   360
            TabIndex        =   47
            Top             =   300
            Width           =   915
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            Caption         =   "Format Code:"
            Height          =   255
            Left            =   4980
            TabIndex        =   46
            Top             =   300
            Width           =   1035
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Caption         =   "Format Chars:"
            Height          =   255
            Left            =   2580
            TabIndex        =   45
            Top             =   300
            Width           =   1035
         End
      End
      Begin VB.ComboBox metric_unit 
         Height          =   315
         Left            =   7020
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Tag             =   "1S"
         Top             =   540
         Width           =   1215
      End
      Begin VB.ComboBox unit 
         Height          =   315
         Left            =   4620
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Tag             =   "1S"
         Top             =   540
         Width           =   1215
      End
      Begin VB.TextBox index_desc 
         Height          =   315
         Left            =   1260
         TabIndex        =   7
         Tag             =   "1S"
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox index_code 
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
         Left            =   4620
         TabIndex        =   8
         Tag             =   "1S"
         Top             =   960
         Width           =   435
      End
      Begin VB.TextBox model_name 
         Height          =   315
         Left            =   1260
         TabIndex        =   4
         Tag             =   "1S"
         Top             =   540
         Width           =   2535
      End
      Begin VB.CheckBox traces_ind 
         Caption         =   "TRACES"
         Height          =   255
         Left            =   6120
         TabIndex        =   9
         Tag             =   "1"
         Top             =   1020
         Width           =   975
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Tech Desc:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   55
         Top             =   600
         Width           =   915
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Book Desc:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   54
         Top             =   1200
         Width           =   915
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Crew Equip:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   53
         Top             =   1800
         Width           =   915
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Crew Equip Plural:"
         Height          =   255
         Left            =   -70920
         TabIndex        =   52
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Crew Equip Plural:"
         Height          =   255
         Left            =   -70920
         TabIndex        =   51
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Crew Equip:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   50
         Top             =   1800
         Width           =   915
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         Caption         =   "Book Desc:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   49
         Top             =   1200
         Width           =   915
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Caption         =   "Tech Desc:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   48
         Top             =   600
         Width           =   915
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         Caption         =   "Metric Unit:"
         Height          =   255
         Left            =   6060
         TabIndex        =   43
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         Caption         =   "Unit:"
         Height          =   255
         Left            =   4080
         TabIndex        =   42
         Top             =   600
         Width           =   435
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         Caption         =   "Index Desc:"
         Height          =   255
         Left            =   300
         TabIndex        =   41
         Top             =   1020
         Width           =   855
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         Caption         =   "Index Code:"
         Height          =   255
         Left            =   3600
         TabIndex        =   40
         Top             =   1020
         Width           =   915
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         Caption         =   "Model Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Chng Notice Cd:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   38
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "Table Ref ID:"
         Height          =   255
         Left            =   -72300
         TabIndex        =   37
         Top             =   1020
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Graphic Ref ID:"
         Height          =   255
         Left            =   -74820
         TabIndex        =   36
         Top             =   1020
         Width           =   1155
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "Format Chars:"
         Height          =   255
         Left            =   -72360
         TabIndex        =   35
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Format Code:"
         Height          =   255
         Left            =   -69960
         TabIndex        =   34
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Table Ref Col:"
         Height          =   255
         Left            =   -69960
         TabIndex        =   33
         Top             =   1020
         Width           =   1035
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Indent Code:"
         Height          =   255
         Left            =   -74580
         TabIndex        =   32
         Top             =   600
         Width           =   915
      End
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      Caption         =   "Last Updated By:"
      Height          =   255
      Left            =   3120
      TabIndex        =   29
      Top             =   3780
      Width           =   1275
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      Caption         =   "Last Updated:"
      Height          =   255
      Left            =   180
      TabIndex        =   28
      Top             =   3780
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Type Code:"
      Height          =   255
      Left            =   5760
      TabIndex        =   27
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Equip ID:"
      Height          =   255
      Left            =   180
      TabIndex        =   26
      Top             =   120
      Width           =   915
   End
   Begin VB.Label Label29 
      Alignment       =   1  'Right Justify
      Caption         =   "Alt Equip ID:"
      Height          =   255
      Left            =   3000
      TabIndex        =   25
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmEquipment"
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
Public Sub SetRow(rec As ADODB.RecordSet, Optional blnInsert As Boolean = False)
    Set m_rec = rec
    m_blnInsert = blnInsert
    ' If we are inserting/cloning
    If m_blnInsert Then
        ' Do this so OriginalValue will be set to the values copied into the row
        m_rec.UpdateBatch
    End If
    If Not m_rec.Fields("equip_id") = 0 Then
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

    strUpdate = "exec sp_delete_equipment "
    strUpdate = strUpdate + "@equip_skey=" + str(Me.Controls("equip_skey")) + ", "
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
    
    m_blnWereErrors = False
    
    ' If we are updating
    If m_blnInsert = False Then
        strUpdate = "exec sp_update_equipment @equip_skey=" + equip_skey.Text + ", @equip_last_update_id=" + equip_last_update_id.Text + ", "
        BuildStoredProcSQL Me, strUpdate, "1"
        strUpdate = strUpdate + " @last_update_person='" + strUserName + "'"
    ' If we are inserting
    Else
        strUpdate = "exec sp_insert_equipment "
        BuildStoredProcSQL Me, strUpdate, "1"
        strUpdate = strUpdate + " @last_update_person='" + strUserName + "'"
    End If
    
    blnRet = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
    If Not blnRet Then
        MsgBox strError
        m_blnWereErrors = True
    Else
        ' Put latest data into source recordset
        UpdateRecordsetFromForm Me, m_rec
        m_rec.Fields("equip_last_update_id").Value = m_rec.Fields("equip_last_update_id").Value + 1
        equip_last_update_id.Text = m_rec.Fields("equip_last_update_id").Value
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
    On Error Resume Next
    Dim ctr As Control
    Dim rec As ADODB.RecordSet
    Move START_LEFT, START_TOP

    g_objDAL.GetRecordset CONNECT, "select unit from unit_of_measure order by unit", rec
    While Not rec.EOF
        unit.AddItem (rec.Fields("unit").Value)
        metric_unit.AddItem (rec.Fields("unit").Value)
        rec.MoveNext
    Wend
    
    ' If we are showing data
'    If m_blnRecFlag = True Then
        If Not m_rec.State = adStateClosed Then
            UpdateFormFromRecordset Me, m_rec
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
    ' If we are NOT inserting
    If m_blnInsert = False Then
        ' Lock fields that can't be changed
        equip_id.Locked = True
        equip_id.BackColor = LTGREY
        
        Me.Caption = Me.Caption + " [" + m_rec.Fields("equip_id").Value + "]"
    ElseIf m_blnInsert = True And m_blnRecFlag = True Then
        Me.Caption = Me.Caption + " [Clone of " + m_rec.Fields("equip_id").Value + "]"
    Else
        Me.Caption = Me.Caption + " [New]"
    End If
    
    SSTab.Tab = 0
    ColorLockedFields Me
    
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

Private Sub Form_Resize()
ResizeForm Me
End Sub

Private Sub indent_code_Validate(Cancel As Boolean)
    CheckValueForNumber indent_code.Text, Cancel
End Sub

Private Sub format_characters_Validate(Cancel As Boolean)
    CheckValueForNumber format_characters.Text, Cancel
End Sub

