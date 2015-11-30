VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCWOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' This is the base class for classes providing output to CostWorks
Option Explicit

Private m_fsOutput
Private m_fOutput
Private m_colOutputFields As New Collection

Private Sub Class_Initialize()
    Set m_fsOutput = CreateObject("Scripting.FileSystemObject")
    Set m_fOutput = m_fsOutput.CreateTextFile("cwoutput.txt", True) 'Create a file
End Sub

Private Sub AddData(varData As Variant, clsOutputField As COutputField)
    On Error Resume Next
    Dim str As String
    str = Trim(CStr(varData))
    ' For numbers that are not empty and not = 0, lead with zeros
    If clsOutputField.Numeric And Len(str) <> 0 And str <> "0" Then
        Dim temp As String
        While Len(temp) + Len(str) < clsOutputField.Size
            temp = temp + "0"
        Wend
        str = temp + str
    ' For strings and empty numbers, pad with spaces; for numbers = 0, fill with spaces
    Else
        ' Value of zero represented by spaces only
        If str = "0" Then str = " "
        ' Pad with spaces
        While Len(str) < clsOutputField.Size
            str = str + " "
        Wend
    End If
    m_fOutput.Write str
End Sub

Public Sub Flush()
    m_fOutput.Write Chr(26)
    m_fOutput.Close
End Sub

Public Sub CreateField(strName As String, intSize As Integer, Optional blnNumeric As Boolean = False, Optional blnSpecial As Boolean = False)
    Dim clsOutputField As New COutputField
    clsOutputField.Name = strName
    clsOutputField.Size = intSize
    clsOutputField.Numeric = blnNumeric
    clsOutputField.Special = blnSpecial
    m_colOutputFields.Add clsOutputField, strName
End Sub

Public Sub AddRecord(rec As ADODB.RecordSet)
    Dim clsOutputField As COutputField
    
    For Each clsOutputField In m_colOutputFields
        AddData rec.Fields(clsOutputField.Name), clsOutputField
    Next
    m_fOutput.WriteLine ("")
End Sub
