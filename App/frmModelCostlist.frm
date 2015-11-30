VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmModelCostList 
   Caption         =   "Model Cost List"
   ClientHeight    =   5520
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   ScaleHeight     =   5520
   ScaleWidth      =   8490
   StartUpPosition =   3  'Windows Default
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   4695
      Left            =   120
      OleObjectBlob   =   "frmModelCostlist.frx":0000
      TabIndex        =   4
      Top             =   720
      Width           =   8295
   End
   Begin VB.Label Label4 
      Caption         =   "Basement:"
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Story Height:"
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Stories:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Model Type:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Menu mnuTop 
      Caption         =   "&View"
      Index           =   0
      Begin VB.Menu mnuView 
         Caption         =   "&Detail"
         Index           =   0
      End
      Begin VB.Menu mnuView 
         Caption         =   "&Summary"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmModelCostList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
