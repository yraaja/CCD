Attribute VB_Name = "modLog"
Option Explicit

Private Const MAX_PATH As Integer = 255
Private Declare Function apiGetTempDir Lib "kernel32" Alias "GetTempPathA" _
        (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" _
        (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long

Private Function FileExists(sFileName As String) As Boolean

    If sFileName = "" Or Right(sFileName, 1) = "\" Then
        FileExists = False
        Exit Function
    End If
    FileExists = (Dir(sFileName) <> "")

End Function

Private Function GetTempFolder() As String
'RETURN THE CURRENT USER'S TEMP FILE FOLDER
'06/30/2005 RTD

    Dim strTempDir As String
    Dim lngx As Long
    
    strTempDir = String$(MAX_PATH, 0)
    lngx = apiGetTempDir(MAX_PATH, strTempDir)
    If lngx <> 0 Then
        strTempDir = Left$(strTempDir, lngx)
        If Right(strTempDir, 1) <> "\" Then
            strTempDir = strTempDir & "\"
        End If
    Else
        strTempDir = ""
    End If
    GetTempFolder = strTempDir
    
End Function

Private Function AppendLogOK() As Boolean
' RETURN TRUE IF IT IS OK TO APPEND TO THE CURRENT DAY'S LOG
' RETURN FALSE IF THE LOG SHOULD BE OVERWRITTEN
    Dim sFileName As String
    Dim f As Long
    Dim sTemp As String
    Dim dtLogDate As Date
    
    On Error GoTo Err_Handler
    sFileName = GetTempFolder & "CCD Log " & Weekday(Date) & ".txt"
    f = FreeFile
    Open sFileName For Input Access Read As #f
    If Not EOF(f) Then
        Line Input #f, sTemp
        Input #f, dtLogDate, sTemp
    End If
    Close #f
    AppendLogOK = Format(dtLogDate, "Medium Date") = Format(Date, "Medium Date")
    Exit Function
    
Err_Handler:
    AppendLogOK = False
    
End Function

Public Function WriteToLog(ByVal sEvent As String, ParamArray sColumns() As Variant) As Boolean
    Dim bAppend As Boolean
    Dim sFileName As String
    Dim f As Long
    Dim i As Long
    Dim frm As Form
    Dim bVisible As Boolean

    On Error GoTo Err_Handler
    bAppend = AppendLogOK
    sFileName = GetTempFolder & "CCD Log " & Weekday(Date) & ".txt"
    f = FreeFile
    If bAppend Then
        Open sFileName For Append As #f
    Else
        Open sFileName For Output As #f
        Write #f, "Date", "Event";
        For i = 1 To 9
            Write #f, "COL_" & Format(i, "000");
        Next
        Write #f, "COL_010"
    End If
    Write #f, Now, sEvent;
    For i = LBound(sColumns) To UBound(sColumns)
        If InStr(sColumns(i), Chr(34)) > 0 Then
            Write #f, Replace(sColumns(i), Chr(34), "''");
        Else
            Write #f, sColumns(i);
        End If
    Next
    Write #f, ""
    Close #f
    
    If FormOpen("frmLogViewer", frm, bVisible) Then
        frm.RefreshGrid
    End If
    
    WriteToLog = True
    Exit Function

Err_Handler:
    Debug.Print "WriteToLog Error " & Err.Number & ": " & Err.Description
    WriteToLog = False
    Exit Function
    
End Function

Public Function WriteToLogRlh(ByVal sEvent As String, strText As String) As Boolean
    Dim bAppend As Boolean
    Dim sFileName As String
    Dim f As Long
    Dim i As Long
    Dim frm As Form
    Dim bVisible As Boolean
    Dim fs As Scripting.FileSystemObject
    
    

    On Error GoTo Err_Handler
    
    Dim tmpFileStr As String
    
    tmpFileStr = "CCD Log " & Weekday(Date) & ".txt"
    tmpFileStr = GetSpecialFolderLocation(CSIDL_PERSONAL) & tmpFileStr
    'tmpFileStr = App.Path & tmpFileStr
    Set fs = New FileSystemObject
    If fs.FileExists(tmpFileStr) Then
    Else
        fs.CreateTextFile (tmpFileStr)
    End If
    
    bAppend = True
    sFileName = tmpFileStr
    'sFileName = "C:\testwritetolog.txt"
    f = FreeFile
    If bAppend Then
        Open sFileName For Append As #f
    Else
        Open sFileName For Output As #f
'        Write #f, "Date", "Event";
'        For i = 1 To 9
'            Write #f, "COL_" & Format(i, "000");
'        Next
'        Write #f, "COL_010"
    End If
    Write #f, Now, sEvent;
    Write #f, sEvent, vbTab & strText
'    For i = LBound(sColumns) To UBound(sColumns)
'        If InStr(sColumns(i), Chr(34)) > 0 Then
'            Write #f, Replace(sColumns(i), Chr(34), "''");
'        Else
'            Write #f, sColumns(i);
'        End If
'    Next
'    Write #f, ""
    Close #f
    
'    If FormOpen("frmLogViewer", frm, bVisible) Then
'        frm.RefreshGrid
'    End If
    
    WriteToLogRlh = True
    Exit Function

Err_Handler:
    Debug.Print "WriteToLog Error " & Err.Number & ": " & Err.Description
    WriteToLogRlh = False
    Exit Function
    
End Function

Public Function ReadLogIntoRecordset(ByVal DayOfWeek As Integer, ByVal sFilter As String, ByRef rsData As ADODB.RecordSet) As Long
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.RecordSet
    Dim sFileName As String
    Dim sCon As String
    Dim sSQL As String
    
    On Error GoTo Err_Handler
    Screen.MousePointer = vbHourglass
    sFileName = "CCD Log " & DayOfWeek & ".txt"
    If FileExists(GetTempFolder & sFileName) Then
        CopyFile GetTempFolder & sFileName, GetTempFolder & "~LOG.TXT", False
        sFileName = "~LOG.TXT"
        sCon = "Driver={Microsoft Text Driver (*.txt; *.csv)}; "
        sCon = sCon & "DBQ=" & GetTempFolder & "; "
        sCon = sCon & "Extensions=txt,csv; "
        sCon = sCon & "Persist Security Info=False"
        sSQL = "SELECT * FROM [" & sFileName & "]"
        If sFilter <> "" Then
            sSQL = sSQL & " WHERE (" & sFilter & ")"
        End If
        cn.Open sCon
        rs.Open sSQL, cn, adOpenStatic, adLockReadOnly, adCmdText
        ReadLogIntoRecordset = rs.RecordCount
        Set rsData = rs
        'rs.Close
        'cn.Close
        Set rs = Nothing
        Set cn = Nothing
    Else
        Set rs = Nothing
        ReadLogIntoRecordset = 0
    End If
    Screen.MousePointer = vbDefault
    Exit Function

Err_Handler:
    Screen.MousePointer = vbDefault
    Debug.Print "ReadLogIntoRecordset ERROR #" & Err.Number & ": " & Err.Description
    Set rs = Nothing
    ReadLogIntoRecordset = 0
    Exit Function
    
End Function

Public Function outlooktest()
Dim objOutlook As Outlook.Application
Dim objNamespace As Outlook.NameSpace
Dim olMail As Outlook.MailItem

    Dim SafeEmail As Redemption.SafeMailItem
    Dim objEMail As Outlook.MailItem
    Dim objInbox As Outlook.MAPIFolder
    Dim objItems As Outlook.Items
    Dim returnValue As Integer
    Dim Item As Object
    
    'Initialization
    returnValue = True
    'strErrorMessage = ""
     
    'On Error GoTo errorHandler
   
    If objOutlook Is Nothing Then
        Set objOutlook = New Outlook.Application
        Set objNamespace = objOutlook.GetNamespace("MAPI")
        Call objNamespace.Logon
    End If

    'Folder "\\Mailbox - RSMeans (RBI-US RCD)\Inbox"
    Set objInbox = objNamespace.GetFolderFromID("00000000B63F2DE0D5AABF45AD275BC1420FD66601001E64012385B01343830D9629163DBA5A000004F09C300000")
    Set objItems = objInbox.Items.Restrict("[ReceivedTime] > '" & Format("7/15/05 1:30pm", "ddddd h:nn AMPM") & "'")
    For Each Item In objItems
        Set SafeEmail = New Redemption.SafeMailItem
        SafeEmail.Item = Item
        If InStr(SafeEmail.Subject, "Gerald Keller") > 0 Then
            MsgBox SafeEmail.Body, , SafeEmail.Subject
        End If
    Next
        
    objNamespace.Logoff
    Set Item = Nothing
    Set SafeEmail = Nothing
    Set objNamespace = Nothing
    Set objOutlook = Nothing
    
    Exit Function
    
End Function
