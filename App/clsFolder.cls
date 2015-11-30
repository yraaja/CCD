VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSelectCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mActiveFolder As String
Private mFolderPath As String
Private mFolderAlias As String
Private mPhotoCount As Integer

Public Property Get folderPath()
    folderPath = mFolderPath
End Property

Public Property Let folderPath(ByVal vNewValue As Variant)
    mFolderPath = vNewValue
End Property


Public Property Get FolderAlias() As Variant
 FolderAlias = mFolderAlias
End Property

Public Property Let FolderAlias(ByVal vNewValue As Variant)
  mFolderAlias = vNewValue
End Property

Private Sub Class_Initialize()
    mFolderAlias = ""
End Sub

Public Sub ShiftRight(ListBox1 As Control, ListBox2 As Control)
'   If ListBox1.ListIndex <> -1 Then
'     ListBox2.AddItem (ListBox1.List(ListBox1.ListIndex))
'     ListBox1.RemoveItem (ListBox1.ListIndex)
'   End If

On Error GoTo ERRLBL

   Dim i As Integer
   Dim listcount As Integer
   Dim orig_listcount As Integer
   Dim cnt As Integer
   
   If ListBox1.ListIndex <> -1 Then
    orig_listcount = ListBox1.listcount
    listcount = ListBox1.listcount
    
     For i = 0 To listcount - 1
        If ListBox1.Selected(i) = True Then
           ListBox2.AddItem (ListBox1.List(i))
           cnt = cnt + 1
        End If
     Next
     
     If cnt > 0 Then
        For i = listcount To cnt Step -1
           If ListBox1.Selected(i - 1) = True Then
              ListBox1.RemoveItem (i - 1)
           End If
        Next
     End If
     
   End If
   Exit Sub
ERRLBL:
   MsgBox ("(Error)ShiftRight: " & Err.Description)
   
End Sub

Public Sub ShiftRightAll(ListBox1 As Control, ListBox2 As Control)
Dim i As Integer

On Error GoTo ERRLBL

For i = 0 To ListBox1.listcount - 1
     ListBox2.AddItem (ListBox1.List(i))
Next

ListBox1.Clear
Exit Sub
ERRLBL:
    MsgBox ("(Error)ShiftRightAll: " & Err.Description)
    Stop
    Resume
    
End Sub
Public Sub ShiftLeft(ListBox1 As Control, ListBox2 As Control)
   If ListBox2.ListIndex <> -1 Then
     ListBox1.AddItem (ListBox2.List(ListBox2.ListIndex))
      ListBox2.RemoveItem (ListBox2.ListIndex)
   End If
End Sub
Public Sub ShiftLeftAll(ListBox1 As Control, ListBox2 As Control)
Dim i As Integer

On Error GoTo ERRLBL

'For i = 1 To ListBox2.listcount
For i = 0 To ListBox2.listcount - 1
     ListBox1.AddItem (ListBox2.List(i))
     'ListBox2.RemoveItem (i)
Next

ListBox2.Clear
Exit Sub
ERRLBL:
    MsgBox ("(Error)ShiftLeftAll: " & Err.Description)
    Stop
    Resume
    
End Sub

Public Sub AddAllTradeGroupsToOneListBox(listboxIndex As Integer, ctl As ListBox, rsin As ADODB.RecordSet)

Dim i As Integer
Dim j As Long

On Error GoTo ERRLBL

'ctl.Clear      'we don't want to clear the listbox anymore...

If rsin.EOF Then Exit Sub

rsin.MoveFirst


For j = 0 To rsin.RecordCount - 1
    If Len(Trim(rsin![Trade_Group_Code]) > 0) Then
       
        If Len(Trim(rsin![City])) <= 6 Then
        ':::::::::::::::::::::::  WARNING ::::::::::::::::::::::::::::::::::::::::
        ':: NOTE - for all term dates I had to decrement the TERM date by 1
        '::        using dateAdd
        '::        as the term date with HH:MM:SS etc. rounds UP in vb to the
        '::        next day!!!!
        ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
            'ctl.AddItem rsin![Trade_Group_Code] & "   " & vbTab & rsin![Trade_ID] & vbTab & vtab & rsin![City] & "     " & vbTab & vbTab & rsin![State_Code] & vbTab & rsin![start_date] & vbTab & DateAdd("d", -1, rsin![term_date]) & vbTab & rsin![union_base_hrly] & vbTab & rsin![union_fringe_hrly] & vbTab & rsin![tot_union]
             ctl.AddItem rsin![Trade_Group_Code] & "   " & vbTab & rsin![Trade_ID] & vbTab & vtab & rsin![City] & "     " & vbTab & vbTab & rsin![State_Code] & vbTab & rsin![start_date] & vbTab & rsin![term_date] & vbTab & rsin![union_base_hrly] & vbTab & rsin![union_fringe_hrly] & vbTab & rsin![tot_union]
        Else
            'ctl.AddItem rsin![Trade_Group_Code] & "   " & vbTab & rsin![Trade_ID] & vbTab & rsin![City] & vbTab & vbTab & rsin![State_Code] & vbTab & rsin![start_date] & vbTab & DateAdd("d", -1, rsin![term_date]) & vbTab & rsin![union_base_hrly] & vbTab & rsin![union_fringe_hrly] & vbTab & rsin![tot_union]
            ctl.AddItem rsin![Trade_Group_Code] & "   " & vbTab & rsin![Trade_ID] & vbTab & rsin![City] & vbTab & vbTab & rsin![State_Code] & vbTab & rsin![start_date] & vbTab & rsin![term_date] & vbTab & rsin![union_base_hrly] & vbTab & rsin![union_fringe_hrly] & vbTab & rsin![tot_union]
        End If
    End If
    rsin.MoveNext
Next

Exit Sub
ERRLBL:
    MsgBox ("(Error)AddAllTradeGroupsToOneListBox: " & Err.Description)
    Stop
    Resume
    

End Sub

Public Sub AddAllTradeGroupsToOneListView(listboxIndex As Integer, ctl As ListView, rsin As ADODB.RecordSet)

Dim i As Integer
Dim j As Long

On Error GoTo ERRLBL

ctl.ListItems.Clear
ctl.ColumnHeaders.Clear

If rsin.EOF Then Exit Sub

rsin.MoveFirst

ctl.ColumnHeaders.Add , , "Trade Group", ctl.Width, lvwColumnLeft
ctl.ColumnHeaders.Add , , "Trade Id", ctl.Width / 9, lvwColumnCenter
ctl.ColumnHeaders.Add , , "City", ctl.Width / 9, lvwColumnCenter
ctl.ColumnHeaders.Add , , "State", ctl.Width / 9, lvwColumnCenter
ctl.ColumnHeaders.Add , , "Start Date", ctl.Width / 9, lvwColumnCenter
ctl.ColumnHeaders.Add , , "Term Date", ctl.Width / 9, lvwColumnCenter
ctl.ColumnHeaders.Add , , "Union Base", ctl.Width / 9, lvwColumnCenter
ctl.ColumnHeaders.Add , , "Union Fringe", ctl.Width / 9, lvwColumnCenter
ctl.ColumnHeaders.Add , , "Tot Union", ctl.Width / 9, lvwColumnCenter


Dim LID As ListItem
For i = 0 To rsin.RecordCount - 1
    Set LID = ctl.ListItems.Add(, , rsin(0))
    For j = 1 To rsin.Fields.Count - 1
'    If Len(Trim(rsin![Trade_Group_Code]) > 0) Then
       
'        If Len(Trim(rsin![City])) <= 6 Then
        ':::::::::::::::::::::::  WARNING ::::::::::::::::::::::::::::::::::::::::
        ':: NOTE - for all term dates I had to decrement the TERM date by 1
        '::        using dateAdd
        '::        as the term date with HH:MM:SS etc. rounds UP in vb to the
        '::        next day!!!!
        ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
            
           LID.SubItems(j) = rsin(j)
        
'            LID.SubItems(j + 1) = rsin(j + 1)
        
'        Else
'            ctl.ListItems.Add rsin![Trade_Group_Code] ' & "   " & vbTab & rsin![Trade_ID] & vbTab & rsin![City] & vbTab & vbTab & rsin![State_Code] & vbTab & rsin![start_date] & vbTab & DateAdd("d", -1, rsin![term_date]) & vbTab & rsin![union_base_hrly] & vbTab & rsin![union_fringe_hrly] & vbTab & rsin![tot_union]
'        End If
'    End If
    Next
    'Exit Sub
    rsin.MoveNext
Next

Exit Sub
ERRLBL:
    MsgBox ("(Error)AddAllPhotoFilesToOneListBox: " & Err.Description)
    Stop
    Resume
    

End Sub

Public Property Get PhotoCount() As Variant
    PhotoCount = mPhotoCount
End Property

Public Property Let PhotoCount(ByVal vNewValue As Variant)
    mPhotoCouont = vNewValue
End Property
