VERSION 5.00
Begin VB.Form dlgEquipFactor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Update By Factor"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkOperatingCostHrly 
      Caption         =   "Operating Cost Hrly"
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   660
      Width           =   1695
   End
   Begin VB.CheckBox chkRentPerWeek 
      Caption         =   "Rent per week"
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   660
      Width           =   1395
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   1020
      TabIndex        =   2
      Top             =   1260
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2460
      TabIndex        =   1
      Top             =   1260
      Width           =   1215
   End
   Begin VB.TextBox txtFactor 
      DataField       =   "contact_id"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Tag             =   "1S"
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblFactor 
      Alignment       =   1  'Right Justify
      Caption         =   "Factor:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   180
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "%"
      Height          =   255
      Left            =   1620
      TabIndex        =   3
      Top             =   180
      Width           =   195
   End
End
Attribute VB_Name = "dlgEquipFactor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const RENT = 1
Const OPERATING = 2
Const BOTH = 3

Public Sub GetFactor(ByRef dblFactor As Double, ByRef intApply As Integer)
    Show (vbModal)
    dblFactor = Val(txtFactor)
    If chkRentPerWeek.Value = 1 Then
        If chkOperatingCostHrly.Value = 1 Then
            intApply = EQUIP_FACTOR_BOTH
        Else
            intApply = EQUIP_FACTOR_RENT
        End If
    ElseIf chkOperatingCostHrly.Value = 1 Then
        intApply = EQUIP_FACTOR_OPERATING
    End If
    Clear
End Sub

Private Sub CancelButton_Click()
    Clear
    Me.Hide
End Sub

Private Sub Form_Activate()
    OutputView False

End Sub

Private Sub OKButton_Click()
    Me.Hide
End Sub

Private Sub Clear()
    txtFactor = ""
    chkRentPerWeek = 0
    chkOperatingCostHrly = 0
End Sub

