VERSION 5.00
Begin VB.Form frmDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Report Selection"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4725
      TabIndex        =   5
      Top             =   2160
      Width           =   960
   End
   Begin VB.TextBox txtVariance 
      Height          =   315
      Left            =   2625
      TabIndex        =   3
      Top             =   1200
      Width           =   3060
   End
   Begin VB.Frame frmMeasure 
      Caption         =   "Imperial/Metric"
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   2895
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   2535
         TabIndex        =   11
         Top             =   240
         Width           =   2535
         Begin VB.OptionButton opMetric 
            Caption         =   "Metric"
            Height          =   255
            Left            =   1080
            TabIndex        =   13
            Top             =   0
            Width           =   885
         End
         Begin VB.OptionButton opImperial 
            Caption         =   "Imperial"
            Height          =   255
            Left            =   0
            TabIndex        =   12
            Top             =   0
            Width           =   990
         End
      End
   End
   Begin VB.ComboBox comboClassification 
      Height          =   315
      Left            =   1260
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3675
      TabIndex        =   4
      Top             =   2160
      Width           =   960
   End
   Begin VB.TextBox txtComparisionDate 
      Height          =   315
      Left            =   2625
      TabIndex        =   2
      Top             =   840
      Width           =   3060
   End
   Begin VB.ComboBox cboReports 
      Height          =   315
      Left            =   2625
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   3090
   End
   Begin VB.Label lblVariance 
      Caption         =   "Please provide a variance:"
      Height          =   255
      Left            =   105
      TabIndex        =   10
      Top             =   1245
      Width           =   2430
   End
   Begin VB.Label lbClassification 
      Caption         =   "Classification"
      Height          =   255
      Left            =   105
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblComparisionDate 
      Caption         =   "Please provide a comparision date:"
      Height          =   255
      Left            =   105
      TabIndex        =   7
      Top             =   890
      Width           =   2565
   End
   Begin VB.Label lblReport 
      Caption         =   "Please select a report:"
      Height          =   255
      Left            =   105
      TabIndex        =   6
      Top             =   525
      Width           =   2430
   End
End
Attribute VB_Name = "frmDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bCancel As Boolean

Public Property Get Cancelled() As Boolean
    Cancelled = bCancel
End Property

Private Sub cmdCancel_Click()
    bCancel = True
    Me.Hide
End Sub

Private Sub Form_Load()
    Dim objRS As ADODB.RecordSet
    Dim strText As String
    Dim strSQL As String
    Dim i       As Integer

    With cboReports
        .AddItem "Project Types & Components"
        .AddItem "PCIS Variance"
    End With
    '
    '   Fill the classification combo box
    comboClassification.Clear
        
    strSQL = "SELECT DISTINCT class_id, class_desc FROM CLASSIFICATION WHERE class_system_id = 'F' ORDER BY class_id"
    If Not g_objDAL.GetRecordset(vbNullString, strSQL, objRS) Then
        Screen.MousePointer = vbNormal
        MsgBox "An error occurred while searching for available classification categories."
    Else
        With objRS
            If .RecordCount = 0 Then
                comboClassification.AddItem "(unknown)"
            Else
                comboClassification.AddItem "(ALL)"
                While Not .EOF
                    strText = "(" & Trim(.Fields("class_id")) & ") "
                    For i = Len(strText) To 6
                        strText = strText & " "
                    Next
                    strText = strText & Trim(.Fields("class_desc").Value)
                    comboClassification.AddItem strText
                    .MoveNext
                Wend
            End If
            .Close
        End With
    End If
    '
    '   Now get the PCIS default variance date
    strSQL = "SELECT domain_value FROM DOMAIN_TBL WHERE domain_name = 'PCIS_VAR_DT'"
    If Not g_objDAL.GetRecordset(vbNullString, strSQL, objRS) Then
        Screen.MousePointer = vbNormal
        MsgBox "An error occurred while searching for default variance date."
    Else
        With objRS
            If .RecordCount <> 0 Then
                txtComparisionDate.Text = Trim(.Fields("domain_value").Value)
            End If
            .Close
        End With
    End If
    SetDefaults
    
    Set objRS = Nothing
End Sub

Private Sub SetDefaults()
    opImperial.Value = True
    
End Sub

Private Sub cboReports_Click()
    If cboReports.Text = "PCIS Variance" Then
        txtComparisionDate.Enabled = True
        txtVariance.Enabled = True
    Else
        txtComparisionDate.Enabled = False
        txtVariance.Enabled = False
    End If
End Sub

Private Sub cmdOK_Click()
    If ValidateScreen Then
        txtComparisionDate.Text = Format(Trim(txtComparisionDate.Text), "mm/dd/yyyy")
        Me.Hide
    End If
End Sub

Private Function ValidateScreen() As Boolean
    Dim strMessage  As String
    
    On Error Resume Next
    ValidateScreen = True
    If cboReports.Text = "PCIS Variance" Then
        If Trim(txtComparisionDate.Text) = "" Then
            strMessage = "Please provide a comparision date."
            txtComparisionDate.SetFocus
        ElseIf Trim(txtVariance.Text) = "" Then
            strMessage = "Please provide a variance."
            txtVariance.SetFocus
        End If
    End If
    
    If strMessage <> "" Then
        Screen.MousePointer = vbNormal
        ValidateScreen = False
        MsgBox strMessage, vbCritical
    End If
End Function
