VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cdlgLstSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Wrapper for dlgListSelection
'Contains available and selected collections, and methods to display the selection screen
Dim m_colAvailable As Collection
Dim m_colSelected As Collection
Dim m_blnListChanged As Boolean
Dim m_dlgSelectionList As Form
Dim m_blnCaption As String
Dim m_intSelectType As Integer
Dim m_strSingleValue As String
Dim m_strComboSelectionCaption As String
Dim m_strCheck1Caption As String
Dim m_blnUseCheck1 As Boolean
Dim m_blnCheck1Value As Boolean
Dim m_blnCancel As Boolean

Dim m_strCheck2Caption As String
Dim m_blnUseCheck2 As Boolean
Dim m_blnCheck2Value As Boolean

Const AVAILABLE_ID = 1
Const SELECTED_ID = 2
Const NO_LIST = 4

Public Property Get SingleItemData() As Long
Dim CurrentItem As cListItem

'If m_colAvailable.Count = 1 Then
'    Set CurrentItem = m_colAvailable(1)
'    SingleValue = CurrentItem.Description
'Else
    SingleItemData = m_dlgSelectionList.SingleItemData
'End If
End Property

Public Function AddUniqueItem(Desc1 As String, _
                                        Sort As Integer, _
                                        lngItemData As Long, _
                                        Optional ListID As Integer = AVAILABLE_ID) _
                                            As Boolean
    If ItemExists(Desc1) = False Then
        AddItem Desc1, Sort, lngItemData, ListID
    End If
End Function

Public Function AddItem(Desc1 As String, _
                                        Sort As Integer, _
                                        lngItemData As Long, _
                                        Optional ListID As Integer = AVAILABLE_ID) _
                                            As Boolean
    Dim CurrentItem As New cListItem
    
    'Add one item to either the avail or selected collection
    m_blnListChanged = True
    If m_colAvailable Is Nothing Then
      Set m_colAvailable = New Collection
    End If
    CurrentItem.Description = Desc1
    
    If IsNumeric(Sort) Then
        CurrentItem.SortOrder = Sort
    Else
        CurrentItem.SortOrder = 0
    End If
    If IsNumeric(lngItemData) Then
        CurrentItem.ItemData = lngItemData
    Else
        CurrentItem.ItemData = 0
    End If
    CurrentItem.ListID = ListID

    Select Case ListID
    Case AVAILABLE_ID
        m_colAvailable.Add CurrentItem
    Case SELECTED_ID
        m_colSelected.Add CurrentItem, CStr(Sort)
    End Select

    AddItem = True

End Function

Public Property Let Caption(strCaption As String)

    m_blnCaption = strCaption

End Property

Public Property Let Check1Caption(strCaption As String)
    m_strCheck1Caption = strCaption
    If Trim(strCaption) = "" Then
        m_blnUseCheck1 = False
    Else
        m_blnUseCheck1 = True
    End If
End Property

Public Property Let Check2Caption(strCaption As String)
    m_strCheck2Caption = strCaption
    If Trim(strCaption) = "" Then
        m_blnUseCheck2 = False
    Else
        m_blnUseCheck2 = True
    End If
End Property

Property Get Check1Value() As Boolean
    Check1Value = m_blnCheck1Value
End Property

Property Get Check2Value() As Boolean
    Check2Value = m_blnCheck2Value
End Property

Public Property Let ComboCaption(strComboCaption As String)
    m_strComboSelectionCaption = strComboCaption
End Property

Public Property Get Caption() As String
    Caption = m_blnCaption
End Property

Public Property Get ComboCaption() As String
    ComboCaption = m_strComboSelectionCaption
End Property

Public Property Get itemCount() As Long
    
    If m_colAvailable Is Nothing Then
        itemCount = 0
    Else
        itemCount = m_colAvailable.Count
    End If

End Property

Public Function ItemExists(Desc1 As String, Optional ListID As Integer = AVAILABLE_ID) As Boolean
'See if the item is already in the specified list
    Dim CurrentItem As New cListItem

    Select Case ListID
    Case AVAILABLE_ID
        If m_colAvailable Is Nothing Then
            ItemExists = False
        Else
            For Each CurrentItem In m_colAvailable
                If CurrentItem.Description = Desc1 Then
                    ItemExists = True
                    Exit For
                End If
            Next CurrentItem
        End If
    Case SELECTED_ID
        If m_colSelected Is Empty Then
            ItemExists = False
        Else
            For Each CurrentItem In m_colSelected
                If CurrentItem.Description = Desc1 Then
                    ItemExists = True
                    Exit For
                End If
            Next CurrentItem
        End If
    End Select

End Function

Public Property Get SingleValue() As String
Dim CurrentItem As cListItem

'If m_colAvailable.Count = 1 Then
'    Set CurrentItem = m_colAvailable(1)
'    SingleValue = CurrentItem.Description
'Else
    SingleValue = m_dlgSelectionList.SingleItemValue
'End If
End Property

Public Property Let SingleValue(sValue As String)
    m_dlgSelectionList.SingleItemValue = sValue
End Property

Public Property Get SelectType() As Integer
    SelectType = m_intSelectType
End Property

Public Property Let SelectType(intSelectType As Integer)
    m_intSelectType = intSelectType
End Property

Public Function SetList() As Boolean
Dim I As Integer
Dim CurrentItem As New cListItem

If m_colAvailable Is Nothing And m_intSelectType <> 4 Then
    MsgBox "No items have been added for selection!"
    SetList = False
Else
    If m_dlgSelectionList Is Nothing Then
        Set m_dlgSelectionList = New dlgListSelection
    End If
    With m_dlgSelectionList
    .Caption = m_blnCaption
    .SelectType = m_intSelectType
    
    Select Case m_intSelectType
    
    Case SINGLE_LIST
        .lstSingleSelection.Clear
        .picSingle.Visible = True
        .lstSingleSelection.Visible = True
        .cmbSelection.Visible = False
        .lblComboSelection.Visible = False
        .picMulti.Visible = False
        .Width = m_dlgSelectionList.lstSingleSelection.Width + 350
        .cmdFinished.Left = m_dlgSelectionList.Width / 2 - m_dlgSelectionList.cmdFinished.Width
        .cmdCancel.Left = m_dlgSelectionList.Width / 2
        For Each CurrentItem In m_colAvailable
            .lstSingleSelection.AddItem ((CurrentItem.Description))
        Next CurrentItem
    
    Case AVAILABLE_AND_SELECTED_LISTS
        .lstAvailable.Clear
        .lstSelected.Clear
        .picSingle.Visible = False
        .picMulti.Visible = True
        .Width = 5775
        .cmdFinished.Left = 1560
        .cmdCancel.Left = 3120
        For Each CurrentItem In m_colAvailable
            .lstAvailable.AddItem ((CurrentItem.Description))
            .lstAvailable.ItemData(.lstAvailable.NewIndex) = CurrentItem.ItemData
        Next CurrentItem
    
    Case COMBO_BOX
        .Height = 3000
        .ComboSelectionCaption = m_strComboSelectionCaption
        .cmbSelection.Clear
        .Picture1.Height = 480
        .picReportOptions.Height = 975
        .picSingle.Visible = True
        .picSingle.Height = 1500
        .picMulti.Visible = False
        .lstSingleSelection.Visible = False
        .cmbSelection.Visible = True
        .lblComboSelection.Visible = True
        .Width = .lstSingleSelection.Width + 720
        .cmdFinished.Left = .Width / 2 - m_dlgSelectionList.cmdFinished.Width - 120
        .cmdCancel.Left = m_dlgSelectionList.Width / 2
        For Each CurrentItem In m_colAvailable
            .cmbSelection.AddItem ((CurrentItem.Description))
            .cmbSelection.ItemData(.cmbSelection.NewIndex) = CurrentItem.ItemData
        Next CurrentItem
        If .cmbSelection.ListCount = 1 Then
            .cmbSelection.ListIndex = 0
        End If
        .cmbSelection.Left = .lstSingleSelection.Left
        .lblComboSelection.Left = .lstSingleSelection.Left
    Case 4
        .Height = 2000
        .picSingle.Visible = False
        .picMulti.Visible = False
        .lstSingleSelection.Visible = False
        .cmbSelection.Visible = False
        .lblComboSelection.Visible = False
        .Width = .lstSingleSelection.Width + 850
        .cmdFinished.Left = .Width / 2 - m_dlgSelectionList.cmdFinished.Width - 120
        .cmdCancel.Left = (m_dlgSelectionList.Width / 2) + 120
    End Select

    If m_blnUseCheck1 = True Then
        .picReportOptions.Visible = True
        .Check1.Visible = True
        .Check1.Caption = m_strCheck1Caption
    Else
        .picReportOptions.Visible = False
        .Check1.Visible = False
    End If
    If m_blnUseCheck2 = True Then
        .Check2.Visible = True
        .Check2.Caption = m_strCheck2Caption
    Else
        .Check2.Visible = False
    End If
    
    End With
   SetList = True
End If
End Function

Public Function ShowList() As Boolean

    m_dlgSelectionList.Show vbModal
    If m_dlgSelectionList.Cancel = True Then
        ShowList = False
    Else
        If m_blnUseCheck1 = True Then
            m_blnCheck1Value = m_dlgSelectionList.Check1
        Else
            m_blnCheck1Value = False
        End If
        If m_blnUseCheck2 = True Then
            m_blnCheck2Value = m_dlgSelectionList.Check2
        Else
            m_blnCheck2Value = False
        End If
        ShowList = True
    End If
    
End Function

Private Sub Class_Initialize()
    m_blnCaption = "Selection List"
    m_intSelectType = COMBO_BOX
End Sub

Private Sub Class_Terminate()
    Dim blnVisible As Boolean
    If Not m_dlgSelectionList Is Nothing Then
        If FormOpen(m_dlgSelectionList.Name, m_dlgSelectionList, blnVisible) = True Then
            Unload m_dlgSelectionList
        End If
    End If
    Set m_dlgSelectionList = Nothing
End Sub

