VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInfoSourceMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' This class is "derived" from CGridMap
' It serves the Information Source Grid
Option Explicit

'*** APEX Migration Utility Code Change ***
'Private WithEvents TDBGrid As TrueOleDBGrid60.TDBGrid ' Access to grid events
'*** APEX Migration Utility Code Change ***
'Private WithEvents TDBGrid As TrueOleDBGrid70.TDBGrid ' Access to grid events
Private WithEvents TDBGrid As TrueOleDBGrid80.TDBGrid ' Access to grid events
Attribute TDBGrid.VB_VarHelpID = -1

Const m_GRIDTYPE = "Information Sources"
' Names of table used by the grid
Const INFO_TABLE = "Information_source"
Const INFO_UPDATE = "Information_source update"

Private m_objGridMap As New CGridMap
Private m_rec As ADODB.RecordSet ' recordset that populates grid
Private m_blnPreferences As Boolean ' Set to True if on Grid Preferences screen

Public Sub SetMenuBar()
    m_objGridMap.SetMenuBar
End Sub
Public Sub Sort(intDir As Integer)
    m_objGridMap.Sort intDir
End Sub
Public Sub SelectAllRows()
    m_objGridMap.SelectAllRows
End Sub

Public Property Let RecordSet(RecSet As ADODB.RecordSet)
    ' Pass-through to base class
    m_objGridMap.RecordSet = RecSet
    Set m_rec = RecSet
End Property

'*** APEX Migration Utility Code Change ***
'Public Sub SetGrid(grd As TrueOleDBGrid60.TDBGrid)
'*** APEX Migration Utility Code Change ***
'Public Sub SetGrid(grd As TrueOleDBGrid70.TDBGrid)
Public Sub SetGrid(grd As TrueOleDBGrid80.TDBGrid)
    ' Pass-through to base class
    m_objGridMap.SetGrid grd
    Set TDBGrid = grd
End Sub

Public Property Let Preferences(bln As Boolean)
    ' Pass-through to base class
    m_blnPreferences = bln
End Property

Public Sub InitGrid()
    m_objGridMap.GRIDTYPE = m_GRIDTYPE
'    m_objGridMap.SetFrozenCols (1)
    
    ' Add the correct columns to the grid
    m_objGridMap.AddCol "Contact ID", "contact_id", INFO_TABLE, 1200, True
    m_objGridMap.AddSplit
    m_objGridMap.AddCol "First Name", "first_name", INFO_TABLE, 1200
    m_objGridMap.AddCol "MI", "middle_initial", INFO_TABLE, 400
    m_objGridMap.AddCol "Last Name", "last_name", INFO_TABLE, 1200
    m_objGridMap.AddCol "Suffix", "name_suffix", INFO_TABLE, 800
    m_objGridMap.AddCol "Nickname", "nickname", INFO_TABLE, 1000
    m_objGridMap.AddCol "Company", "company_name", INFO_TABLE, 1200
    m_objGridMap.AddCol "Addr", "address1", INFO_TABLE, 1600
    m_objGridMap.AddCol "Addr 2", "address2", INFO_TABLE, 1600
    m_objGridMap.AddCol "Addr 3", "address3", INFO_TABLE, 1600
    m_objGridMap.AddCol "City", "city", INFO_TABLE, 1200
    m_objGridMap.AddCol "State", "state_code", INFO_TABLE, 800, False, True, True
    m_objGridMap.AddCol "Zip Code", "zip_code", INFO_TABLE, 800
    m_objGridMap.AddCol "Sal", "salutation", INFO_TABLE, 800
    m_objGridMap.AddCol "Title", "title", INFO_TABLE, 1200
    m_objGridMap.AddCol "Phone 1", "phone1", INFO_TABLE, 1100
    m_objGridMap.AddCol "Phone 2", "phone2", INFO_TABLE, 1100
    m_objGridMap.AddCol "Email", "email", INFO_TABLE, 1200
    m_objGridMap.AddCol "Fax", "fax", INFO_TABLE, 1100
    m_objGridMap.AddCol "Web Site", "url", INFO_TABLE, 1200, , , , FORMAT_STRING_URL
    m_objGridMap.AddCol "Src Cd", "source_code", INFO_TABLE, 800
    m_objGridMap.AddCol "CCI Ind", "cci_use_ind", INFO_TABLE, 800
    m_objGridMap.AddCol "CCI Contact Nm", "cci_contact_nm", INFO_TABLE, 2000
    m_objGridMap.AddCol "Ltr Cd", "cci_letter_cd", INFO_TABLE, 800
    m_objGridMap.AddCol "Update Cd", "cci_update_cd", INFO_TABLE, 800
    m_objGridMap.AddCol "Metro Cd", "cci_metro_cd", INFO_TABLE, 800
    m_objGridMap.AddCol "Tickler Dt", "tickler_date", INFO_TABLE, 1200, False, True, False, FORMAT_DATE
    m_objGridMap.AddCol "Council", "council_assoc_ind", INFO_TABLE, 1000
    m_objGridMap.AddCol "Country", "country_code", INFO_TABLE, 1000, False, True, True
    m_objGridMap.AddCol "Create Dt", "create_date", vbNullString, 1200, True, True, False, FORMAT_DATE
    m_objGridMap.AddCol "Create By", "create_person", vbNullString, 1200, True, True, False
    m_objGridMap.AddCol "Comment", "comment", INFO_TABLE, 2000
    m_objGridMap.AddCol "Keywords", "keyword", INFO_TABLE, 2000
    m_objGridMap.AddCol "Update Date", "last_update_date", vbNullString, 1650, True, True, False, FORMAT_DATETIME
    m_objGridMap.AddCol "Update By", "last_update_person", vbNullString, 1200, True, True, False
    m_objGridMap.AddCol "last_update_id", "last_update_id", INFO_UPDATE, 800, True, False, False
    
    ' Call base class
    If m_blnPreferences Then
        m_objGridMap.BuildPreferenceGrid
    Else
        m_objGridMap.BuildGrid
    End If
               
    If Not m_blnPreferences Then
        Dim bRet As Boolean
        Dim rec As ADODB.RecordSet
'*** APEX Migration Utility Code Change ***
'        Dim Item As New TrueOleDBGrid60.ValueItem
'*** APEX Migration Utility Code Change ***
'        Dim Item As New TrueOleDBGrid70.ValueItem
        Dim Item As New TrueOleDBGrid80.ValueItem
        
        bRet = g_objDAL.GetRecordset(vbNullString, "select state_code from state_country", rec)
        If Not bRet Then
            Exit Sub
        End If
        
        With TDBGrid.Columns("State").ValueItems
            While Not rec.EOF
                Item.Value = rec.Fields(0)
                TDBGrid.Columns("State").ValueItems.Add Item
                rec.MoveNext
            Wend
            .Presentation = dbgComboBox
            .Validate = True
        End With
        rec.Close
        Set rec = Nothing
        
        bRet = g_objDAL.GetRecordset(vbNullString, "select country_code from country", rec)
        If Not bRet Then
            Exit Sub
        End If
        
        With TDBGrid.Columns("Country").ValueItems
            While Not rec.EOF
                Item.Value = rec.Fields(0)
                TDBGrid.Columns("Country").ValueItems.Add Item
                rec.MoveNext
            Wend
            .Presentation = dbgComboBox
            .Validate = True
        End With
        rec.Close
        Set rec = Nothing
    End If
End Sub

Public Function Update() As Boolean
    Dim strUpdate As String ' SQL string
    Dim bRet As Boolean
    Dim strError As String ' Error string returned from DAL
    Dim intErrors As Integer ' Tracks if any errors have occurred
    Dim intSuccess As Integer ' Tracks successful updates
    
    On Error Resume Next
    Update = True
    intErrors = 0
    intSuccess = 0
    
    m_rec.MoveFirst
    While Not m_rec.EOF
        strUpdate = ""
        ' Skip the record if it didn't change
        Dim clsRowInfo As CRowInfo
        Set clsRowInfo = m_objGridMap.rowinfo(m_rec.Bookmark)
        If Not clsRowInfo Is Nothing Then
            If Not clsRowInfo.State = STATE_NONE Then
                ' If the record was added
                If clsRowInfo.State = STATE_NEW Then
                    ' Could be new to recordset, but really an existing record
                    ' Happens when user types contact_id into new grid row
                    If m_rec.Fields("last_update_date") = "" Then
                        strUpdate = "exec sp_insert_information_source "
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, INFO_TABLE
                    Else
                        strUpdate = "exec sp_update_information_source "
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, INFO_TABLE
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, INFO_UPDATE
                    End If
                ' If the record was changed
                ElseIf clsRowInfo.State = STATE_MODIFIED Then
                    strUpdate = "exec sp_update_information_source "
                    m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, INFO_TABLE
                    m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, INFO_UPDATE
                End If
                
                ' Make sure there is something to do
                If Not strUpdate = "" Then
                    strUpdate = strUpdate + "@last_update_person='" + strUserName + "'"
                    
                    bRet = False
                    strError = ""
                    bRet = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                    ' Reset on success
                    If bRet = True Then
                        intSuccess = intSuccess + 1
                        clsRowInfo.State = STATE_NONE
                        m_rec.Fields("last_update_id").Value = m_rec.Fields("last_update_id").Value + 1
                        m_rec.Fields("last_update_date").Value = Now
                        m_rec.Fields("last_update_person").Value = strUserName
                        Reset_Orig_Values m_rec
                    Else
                        intErrors = intErrors + 1
                        clsRowInfo.Error = True
                        clsRowInfo.ErrorText = strError
                    End If
                End If
            End If
        End If
        m_rec.MoveNext
    Wend
    
 '   m_rec.UpdateBatch
    Dim strMsg As String
    strMsg = ""
    If intSuccess > 0 Then
        strMsg = str(intSuccess) + " rows updated successfully." + Chr(10)
    End If
    If intErrors > 0 Then
        strMsg = strMsg + str(intErrors) + " errors occurred."
        ' Return value will be False
        Update = False
    End If
    If Len(strMsg) > 0 Then
        MsgBox strMsg
    End If
    
    TDBGrid.ReBind
    Exit Function

CleanUp:
    MsgBox "Error in cmdUpdate_Click"
End Function

Private Sub Class_Initialize()
    m_objGridMap.m_colCloneCols.Add 1, "contact_id"
    m_objGridMap.m_colCloneCols.Add 1, "first_name"
    m_objGridMap.m_colCloneCols.Add 1, "middle_initial"
    m_objGridMap.m_colCloneCols.Add 1, "last_name"
    m_objGridMap.m_colCloneCols.Add 1, "name_suffix"
    m_objGridMap.m_colCloneCols.Add 1, "nickname"
    m_objGridMap.m_colCloneCols.Add 1, "company_name"
    m_objGridMap.m_colCloneCols.Add 1, "address1"
    m_objGridMap.m_colCloneCols.Add 1, "address2"
    m_objGridMap.m_colCloneCols.Add 1, "address3"
    m_objGridMap.m_colCloneCols.Add 1, "city"
    m_objGridMap.m_colCloneCols.Add 1, "state_code"
    m_objGridMap.m_colCloneCols.Add 1, "country_code"
    m_objGridMap.m_colCloneCols.Add 1, "zip_code"
    m_objGridMap.m_colCloneCols.Add 1, "salutation"
    m_objGridMap.m_colCloneCols.Add 1, "title"
    m_objGridMap.m_colCloneCols.Add 1, "phone1"
    m_objGridMap.m_colCloneCols.Add 1, "phone2"
    m_objGridMap.m_colCloneCols.Add 1, "tickler_date"
    m_objGridMap.m_colCloneCols.Add 1, "email"
    m_objGridMap.m_colCloneCols.Add 1, "fax"
    m_objGridMap.m_colCloneCols.Add 1, "url"
    m_objGridMap.m_colCloneCols.Add 1, "source_code"
    m_objGridMap.m_colCloneCols.Add 1, "council_assoc_ind"
    m_objGridMap.m_colCloneCols.Add 1, "comment"
    m_objGridMap.m_colCloneCols.Add 1, "keyword"
End Sub

Private Sub TDBGrid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    On Error Resume Next
    ' Only worry about first column here
    If ColIndex = 0 Then
        ' Can only edit new rows and only once
        If OldValue = Empty Then
            Dim strSELECT As String
            Dim rec As New ADODB.RecordSet ' Recordset to hold query results
            Dim blnReturn As Boolean
            Dim vntMyBookmark As Variant
            Dim I As Integer
        
            ' Check to see if the mat_id entered exists already
            strSELECT = "Select * from Information_source where contact_id='" + TDBGrid.Text + "'"
            ' Use DAL to perform select
            blnReturn = g_objDAL.GetRecordset(CONNECT, strSELECT, rec)
            ' If it does, copy that data into grid
            If rec.RecordCount > 0 Then
                m_rec.AddNew ' Add new record to recordset that populates grid
                ' Make a new RowInfo
                m_objGridMap.SetRowState m_rec.Bookmark, STATE_NEW
                ' Copy all values
                For I = 0 To rec.Fields.Count - 1
                    m_rec.Fields(rec.Fields(I).Name) = rec.Fields(I).Value
                Next I
                vntMyBookmark = m_rec.Bookmark ' Save current bookmark
                TDBGrid.ReBind ' Reset grid data
                TDBGrid.ReBind ' Do it twice, don't know why
                TDBGrid.Bookmark = vntMyBookmark ' Set bookmark back
                Cancel = True ' Cancel this event so grid doesn't do any more processing
            End If
        Else
            MsgBox "You cannot edit this column for an existing row."
            Cancel = True
        End If
    End If
End Sub

Private Sub TDBGrid_Change()
Dim iLen As Integer
    With TDBGrid
        If .Columns(.Col).Caption = "MI" _
            Or .Columns(.Col).Caption = "Update Cd" Then
            iLen = 1
        ElseIf .Columns(.Col).Caption = "Src Cd" Then
            iLen = 2
        ElseIf .Columns(.Col).Caption = "Country" _
            Or .Columns(.Col).Caption = "Ltr Cd" Then
            iLen = 3
        ElseIf .Columns(.Col).Caption = "State" _
            Or .Columns(.Col).Caption = "Metro Cd" Then
            iLen = 4
        ElseIf .Columns(.Col).Caption = "Suffix" Then
            iLen = 5
        ElseIf .Columns(.Col).Caption = "Contact ID" Then
            iLen = 6
        ElseIf .Columns(.Col).Caption = "Sal" _
            Or .Columns(.Col).Caption = "Zip Code" Then
            iLen = 10
        ElseIf .Columns(.Col).Caption = "Phone 1" _
            Or .Columns(.Col).Caption = "Phone 2" _
            Or .Columns(.Col).Caption = "Fax" Then
            iLen = 12
        ElseIf .Columns(.Col).Caption = "First Name" _
            Or .Columns(.Col).Caption = "Create By" Then
            iLen = 15
        ElseIf .Columns(.Col).Caption = "Last Name" _
            Or .Columns(.Col).Caption = "Nickname" _
            Or .Columns(.Col).Caption = "Title" Then
            iLen = 20
        ElseIf .Columns(.Col).Caption = "City" Then
            iLen = 23
        ElseIf .Columns(.Col).Caption = "Company" _
            Or .Columns(.Col).Caption = "Addr" _
            Or .Columns(.Col).Caption = "Addr 2" _
            Or .Columns(.Col).Caption = "Addr 3" _
            Or .Columns(.Col).Caption = "Email" _
            Or .Columns(.Col).Caption = "URL" Then
            iLen = 30
        ElseIf .Columns(.Col).Caption = "Keywords" Then
            iLen = 120
        ElseIf .Columns(.Col).Caption = "Comment" Then
            iLen = 255
        End If
        If Len(.Text) > iLen Then .Text = left(.Text, iLen)
    End With

End Sub


' Fired when a row is deleted
Private Sub TDBGrid_UnboundDeleteRow(Bookmark As Variant)
    On Error Resume Next
    Dim strUpdate As String
    Dim blnReturn As Boolean
    Dim strError As String
    Dim clsRowInfo As CRowInfo
   
    ' Position to the desired row and execute delete
    m_rec.Bookmark = Bookmark
    
    Set clsRowInfo = m_objGridMap.rowinfo(m_rec.Bookmark)
    If Not clsRowInfo Is Nothing Then
        If clsRowInfo.State = STATE_NEW Then
            TDBGrid.ApproxCount = TDBGrid.ApproxCount - 1
            m_objGridMap.SetRowState m_rec.Bookmark, STATE_DELETED
            m_rec.Delete ' Remove it from recordset so it doesn't show.
            MsgBox "Delete successful."
            Exit Sub
        End If
    End If
    
    strUpdate = "exec sp_delete_information_source "
    strUpdate = strUpdate + "@contact_id='" + m_rec.Fields("contact_id") + "', "
    strUpdate = strUpdate + "@last_update_person='" + strUserName + "'"
    
    blnReturn = g_objDAL.ExecQuery(CONNECT, strUpdate, strError)
    If Not blnReturn Then
        m_objGridMap.SetError m_rec.Bookmark, strError
        Bookmark = Null
    Else
        TDBGrid.ApproxCount = TDBGrid.ApproxCount - 1
        m_objGridMap.SetRowState m_rec.Bookmark, STATE_DELETED
        m_rec.Delete ' Remove it from recordset so it doesn't show.
        MsgBox "Delete successful."
    End If
    Exit Sub
End Sub

Private Sub TDBGrid_OnAddNew()
    ' Defaults for new added row
    TDBGrid.Columns("last_update_id").Value = 0
End Sub

Public Function CloneRow() As ADODB.RecordSet
    ' Pass to base class
    Set CloneRow = m_objGridMap.CloneRowRecordset
End Function

Public Property Get GRIDTYPE() As String
    GRIDTYPE = m_GRIDTYPE
End Property

Public Function GetError(intBookmark As Integer) As String
    On Error Resume Next
    Dim objRowInfo As CRowInfo
    Set objRowInfo = Nothing
    Set objRowInfo = m_objGridMap.rowinfo(intBookmark)
    If Not objRowInfo Is Nothing Then
        GetError = objRowInfo.ErrorText
    Else
        GetError = ""
    End If
End Function

Public Function IsPendingChange() As Boolean
    IsPendingChange = m_objGridMap.IsPendingChange
End Function

Public Sub RowWrap(intRowWrap As Integer)
    m_objGridMap.RowWrap (intRowWrap)
End Sub


