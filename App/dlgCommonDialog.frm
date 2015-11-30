VERSION 5.00
Begin VB.Form dlgCommonDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   1890
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3630
   Icon            =   "dlgCommonDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   3630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   3630
      TabIndex        =   1
      Top             =   1320
      Width           =   3630
      Begin VB.CommandButton OKButton 
         Cancel          =   -1  'True
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   435
         Left            =   1200
         TabIndex        =   2
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Label lblMessage 
      Caption         =   "message caption"
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "dlgCommonDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub OKButton_Click()
    Me.Visible = False
End Sub
