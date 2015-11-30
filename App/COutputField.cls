VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COutputField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' This class stores information about a field that will be output to CostWorks
Option Explicit

Private strName As String ' The recordset field name or some other identifier
Private intSize As Integer ' The length the field must occupy in the output file
Private blnSpecial As Boolean ' Whether the field requires special processing before output
Private blnNumeric As Boolean ' Whether the field is numeric or char

Public Property Get Name() As Variant
    Name = strName
End Property

Public Property Let Name(ByVal vNewValue As Variant)
    strName = vNewValue
End Property

Public Property Get Size() As Variant
    Size = intSize
End Property

Public Property Let Size(ByVal vNewValue As Variant)
    intSize = vNewValue
End Property

Public Property Get Special() As Variant
    Special = blnSpecial
End Property

Public Property Let Special(ByVal vNewValue As Variant)
    blnSpecial = vNewValue
End Property

Public Property Get Numeric() As Variant
    Numeric = blnNumeric
End Property

Public Property Let Numeric(ByVal vNewValue As Variant)
    blnNumeric = vNewValue
End Property

