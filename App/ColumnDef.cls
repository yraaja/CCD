VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CColumnDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' This class stores information about a column that will exist in a grid
Option Explicit

Private strName As String ' Name of column shown in grid
Private strTableCol As String ' Name of column in table
Private strTable As String ' Table column comes from
Private blnLocked As Boolean ' Locked or editable
Private intColWidth As Integer ' Width of column in grid
Private intOrder As Integer ' Order of column in grid
Private blnVisible As Boolean ' Visible or hidden
Private blnClone As Boolean ' Whether to copy this column when cloning a row
Private strDropDown As String ' Name of drop-down that grid column uses
Private intFormat As Integer ' Used to define the format to apply to the data on display
Private blnFetchStyle As Boolean ' Whether to fire FetchStyle event for column

Public Property Get Name() As Variant
    Name = strName
End Property

Public Property Let Name(ByVal vNewValue As Variant)
    strName = vNewValue
End Property

Public Property Get TableCol() As Variant
    TableCol = strTableCol
End Property

Public Property Let TableCol(ByVal vNewValue As Variant)
    strTableCol = vNewValue
End Property

Public Property Get Table() As Variant
    Table = strTable
End Property

Public Property Let Table(ByVal vNewValue As Variant)
    strTable = vNewValue
End Property

Public Property Get Locked() As Variant
    Locked = blnLocked
End Property

Public Property Let Locked(ByVal vNewValue As Variant)
    blnLocked = vNewValue
End Property

Public Property Get ColWidth() As Variant
    ColWidth = intColWidth
End Property

Public Property Let ColWidth(ByVal vNewValue As Variant)
    intColWidth = vNewValue
End Property

Public Property Get Order() As Variant
    Order = intOrder
End Property

Public Property Let Order(ByVal vNewValue As Variant)
    intOrder = vNewValue
End Property

Public Property Get Visible() As Variant
    Visible = blnVisible
End Property

Public Property Let Visible(ByVal vNewValue As Variant)
    blnVisible = vNewValue
End Property

Public Property Get Clone() As Variant
    Clone = blnClone
End Property

Public Property Get FetchStyle() As Variant
    FetchStyle = blnFetchStyle
End Property

Public Property Let Clone(ByVal vNewValue As Variant)
    blnClone = vNewValue
End Property

Public Property Let FetchStyle(ByVal vNewValue As Variant)
    blnFetchStyle = vNewValue
End Property

Public Property Get Format() As Variant
    Format = intFormat
End Property

Public Property Let Format(ByVal vNewValue As Variant)
    intFormat = vNewValue
End Property

Public Property Get DropDown() As Variant
    DropDown = strDropDown
End Property

Public Property Let DropDown(ByVal vNewValue As Variant)
    strDropDown = vNewValue
End Property

