VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCommonDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private msngX As Single 'local copy
'local variable(s) to hold property value(s)
Private msngY As Single 'local copy
Private mstrMessage As String 'local copy
Private mstrButtons As Long 'local copy
Private m_frmDialog As Form
Public Function ShowMessage() As Boolean
m_frmDialog.lblMessage = mstrMessage
m_frmDialog.Left = msngX - m_frmDialog.Width / 2
m_frmDialog.Top = msngY - m_frmDialog.Height / 2
m_frmDialog.Show 1
'm_frmDialog.ZOrder
End Function

Public Property Let Buttons(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Buttons = 5
    mstrButtons = vData
End Property


Public Property Get Buttons() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Buttons
    Buttons = mstrButtons
End Property



Public Property Let Message(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Message = 5
    mstrMessage = vData
    
End Property


Public Property Get Message() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Message
    Message = mstrMessage
End Property



Public Property Let Y(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Y = 5
    msngY = vData
End Property


Public Property Get Y() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Y
    Y = msngY
End Property



Public Property Let X(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.X = 5
    msngX = vData
End Property


Public Property Get X() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.X
    X = msngX
End Property



Private Sub Class_Initialize()
Set m_frmDialog = New dlgCommonDialog
m_frmDialog.Caption = "Construction Cost Database Message"
End Sub


Private Sub Class_Terminate()
Unload m_frmDialog
End Sub


