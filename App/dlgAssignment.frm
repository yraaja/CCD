VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form dlgAssignment 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Usage Assignment"
   ClientHeight    =   4350
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   7785
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdRemoveAll 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   14
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton cmdRemoveSelected 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   13
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton cmdAddAll 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   12
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton cmdAddSelected 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   11
      Top             =   1440
      Width           =   495
   End
   Begin MSComctlLib.TreeView TreeView2 
      Height          =   3375
      Left            =   4320
      TabIndex        =   10
      Top             =   840
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   5953
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   3375
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   5953
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Frame fraChild 
      Caption         =   "Child Type"
      Height          =   615
      Left            =   3120
      TabIndex        =   6
      Top             =   120
      Width           =   2175
      Begin VB.OptionButton Option5 
         Caption         =   "Unit Cost"
         Height          =   255
         Left            =   1080
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Assembly"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame fraParent 
      Caption         =   "Parent Type"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   2775
      Begin VB.OptionButton Option3 
         Caption         =   "Facility"
         Height          =   255
         Left            =   1800
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Model"
         Height          =   195
         Left            =   1080
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Assembly"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6720
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   5520
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "dlgAssignment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()

End Sub


