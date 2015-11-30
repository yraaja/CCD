Attribute VB_Name = "basVersionCheck"
Option Explicit

Public Const UPGRADE_WEB_SITE = "\\bincmdgkngfap02\means share\ccd\help\upgrade.htm"
Public Const CCD_EXECUTABLE = "ConstructionCostDatabase.exe"
Public Const CCD_KEY = "RSMeans\CCD"

' Server and Database Settings
Public g_objDAL As New CCDdal.CRSMDataAccess ' Global DAL object
Public strConnectServer As String
Public strConnectDatabase As String
Public strUserName As String
Public strConnect As String

' WinAPI Declarations
Private Type VS_FIXEDFILEINFO
   dwSignature As Long
   dwStrucVersion As Long     'e.g. 0x00000042 = "0.42"
   dwFileVersionMS As Long    'e.g. 0x00030075 = "3.75"
   dwFileVersionLS As Long    'e.g. 0x00000031 = "0.31"
   dwProductVersionMS As Long 'e.g. 0x00030010 = "3.10"
   dwProductVersionLS As Long 'e.g. 0x00000031 = "0.31"
   dwFileFlagsMask As Long    'e.g. 0x3F for version "0.42"
   dwFileFlags As Long        'e.g. VFF_DEBUG Or VFF_PRERELEASE
   dwFileOS As Long           'e.g. VOS_DOS_WINDOWS16
   dwFileType As Long         'e.g. VFT_DRIVER
   dwFileSubtype As Long      'e.g. VFT2_DRV_KEYBOARD
   dwFileDateMS As Long       'e.g. 0
   dwFileDateLS As Long       'e.g. 0
End Type

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, nVerSize As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags_ As Long) As Long
Private Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long

Private Function HiWord(dw As Long) As Long
  
   If dw And &H80000000 Then
      HiWord = (dw \ 65535) - 1
   Else
      HiWord = dw \ 65535
   End If
    
End Function
  
Private Function LoWord(dw As Long) As Long
  
   If dw And &H8000& Then
      LoWord = &H8000& Or (dw And &H7FFF&)
   Else
      LoWord = dw And &HFFFF&
   End If
    
End Function

Public Function StripControlCharacters(ByVal sText As String) As String
    Dim sTemp As String
    Dim I As Integer
    
    sTemp = Replace(sText, vbTab, " ")
    For I = 1 To 31
        Do While InStr(sTemp, Chr(I)) > 0
            sTemp = Replace(sTemp, Chr(I), "")
        Loop
    Next
    StripControlCharacters = sTemp

End Function

Private Function GetFileVersion(sDriverFile As String) As String
'RETRIEVES THE VERSION INFORMATION FROM A FILE
'RETURNS A STRING-FORMATTED VERSION
    Dim FI As VS_FIXEDFILEINFO
    Dim sBuffer() As Byte
    Dim nBufferSize As Long
    Dim lpBuffer As Long
    Dim nVerSize As Long
    Dim nUnused As Long
    Dim tmpVer As String
    
    nBufferSize = GetFileVersionInfoSize(sDriverFile, nUnused)
    If nBufferSize > 0 Then
        ReDim sBuffer(nBufferSize)
        Call GetFileVersionInfo(sDriverFile, 0&, nBufferSize, sBuffer(0))
        Call VerQueryValue(sBuffer(0), "\", lpBuffer, nVerSize)
        Call CopyMemory(FI, ByVal lpBuffer, Len(FI))
        tmpVer = Format$(HiWord(FI.dwFileVersionMS)) & "." & _
                 Format$(LoWord(FI.dwFileVersionMS), "0") & "." & _
                 Format$(LoWord(FI.dwFileVersionLS), "0")
    End If
    GetFileVersion = tmpVer
   
End Function

Private Function LaunchBrowser(ByVal URL As String) As Boolean
'LAUNCH THE URL IN THE SYSTEM'S DEFAULT WEB BROWSER
    Dim res As Long
    
    If URL = "" Then URL = "http://www.reedconstructiondata.com/"
    'If (InStr(1, URL, "http", vbTextCompare) <> 1) Then
    '    URL = "http://" & URL
    'End If
    Screen.MousePointer = vbHourglass
    res = ShellExecute(0&, "open", URL, vbNullString, vbNullString, vbMaximizedFocus)
    If res > 32 Then
        Call BringWindowToTop(res)
        LaunchBrowser = True
    Else
        LaunchBrowser = False
    End If
    Screen.MousePointer = vbDefault
    
End Function

Private Function FindParm(strParm As String)
    Dim iStart As Integer
    Dim iLen As Integer
    Dim iEnd As Integer
    
    iStart = InStr(1, UCase(Command), UCase(strParm), vbTextCompare)
    If iStart = 1 Then ' server
        iStart = InStr(1, Command, "=", vbTextCompare)
        iEnd = InStr(iStart + 1, Command, " ", vbTextCompare)
        FindParm = Mid(Command, iStart + 1, iEnd - (iStart + 1))
    Else 'database
        iStart = InStr(iStart + 1, Command, "=", vbTextCompare)
        iEnd = Len(Command)
        FindParm = Mid(Command, iStart + 1, (iEnd - iStart))
    End If
End Function

Public Sub LoadRegKeys()
    Dim strKey As String
    Dim hKey As Long
    Dim lSize As Long
    Dim lRet As Long
    Dim vValue As Variant
    Dim cch As Long
    Dim lrc As Long
    Dim lType As Long
    Dim lValue As Long
    Dim sValue As String
    Dim sBuffer As String
    
    ' Check the Registry for the DBServer Value
    lSize = 1000
    strKey = CCD_KEY + "\Defaults\DBServer"
    lRet = RegOpenKeyEx(HKEY_CURRENT_USER, strKey, 0&, KEY_ALL_ACCESS, hKey)
    ' Test to see if the DBServer Value is there
    If lRet <> ERROR_NONE Then
        ' Populate the values
        'vValue = "means_deveng1"
        vValue = "bincmdgkngeng01"
        lSize = Len(vValue)
        lRet = RegCreateKeyEx(HKEY_CURRENT_USER, strKey, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRet)
        lRet = RegSetValueExString(hKey, "Value", 0&, REG_SZ, vValue, lSize)
    Else
        lrc = RegQueryValueExNULL(hKey, "Value", 0&, REG_SZ, 0&, cch)
        sValue = String(cch, 0)
        lrc = RegQueryValueExString(hKey, "Value", 0&, REG_SZ, sValue, cch)
        If lrc = ERROR_NONE Then
            vValue = Left$(sValue, cch - 1)
        End If
    End If
    strConnectServer = vValue
    RegCloseKey (hKey)

    ' Get the Server setting from the command line - If it's present;
    ' This will overwrite the registry setting.
    sBuffer = Space$(255)
    lSize = Len(sBuffer)
    If InStr(1, UCase(Command), "/CONNECT_SRVR=", vbTextCompare) <> 0 Then
        strConnectServer = FindParm("/CONNECT_SRVR=")
    End If

    ' Check the Registry for the DBase Value
    lSize = 1000
    strKey = CCD_KEY + "\Defaults\DBase"
    lRet = RegOpenKeyEx(HKEY_CURRENT_USER, strKey, 0&, KEY_ALL_ACCESS, hKey)
    ' Test to see if the DBase Value is there
    If lRet <> ERROR_NONE Then
        ' Populate the values
        vValue = "CCDprod"
        lSize = Len(vValue)
        lRet = RegCreateKeyEx(HKEY_CURRENT_USER, strKey, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRet)
        lRet = RegSetValueExString(hKey, "Value", 0&, REG_SZ, vValue, lSize)
    Else
        lrc = RegQueryValueExNULL(hKey, "Value", 0&, REG_SZ, 0&, cch)
        sValue = String(cch, 0)
        lrc = RegQueryValueExString(hKey, "Value", 0&, REG_SZ, sValue, cch)
        If lrc = ERROR_NONE Then
            vValue = Left$(sValue, cch - 1)
        End If
    End If
    strConnectDatabase = vValue
    RegCloseKey (hKey)

    ' Get the Database setting from the command line - If it's present;
    ' This will overwrite the registry setting.
    sBuffer = Space$(255)
    lSize = Len(sBuffer)
    If InStr(1, UCase(Command), "/CONNECT_DB=", vbTextCompare) <> 0 Then
        strConnectDatabase = FindParm("/CONNECT_DB=")
    End If
    
    ' Get the current username
    sBuffer = Space$(255)
    lSize = Len(sBuffer)
    Call GetUserName(sBuffer, lSize)
    strUserName = Left(sBuffer, lSize - 1)
    strUserName = strUserName

    ' Set the Database Connect String
    'strConnect = "UID=" + strUserName + ";PWD=;DATABASE=" + strConnectDatabase + ";SERVER=" + strConnectServer + ";DRIVER={SQL SERVER};DSN='';"
    strConnect = "UID=ccduser;PWD=rsmeans;DATABASE=" + strConnectDatabase + ";SERVER=" + strConnectServer + ";DRIVER={SQL SERVER};DSN='';"  'cje

End Sub


Public Sub Main()
    Dim rec As ADODB.Recordset
    Dim bResult As Boolean
    Dim sAppPath As String
    Dim sSQL As String
    Dim sInstalledVersion As String
    Dim sCurrentVersion As String
    Dim dtCurrentReleaseDate As Date
    Dim sMessage As String
    Dim sDescription As String
    Dim bOkForLaunch As Boolean
    Dim res As Long
    Dim bCheckOnly As Boolean
   
    On Error GoTo Err_Handler
    
    LoadRegKeys
    
    bCheckOnly = (InStr(Command, "/CHECKONLY") > 0)
    
    sAppPath = App.Path
    If Right(sAppPath, 1) <> "\" Then sAppPath = sAppPath & "\"
    sInstalledVersion = GetFileVersion(sAppPath & CCD_EXECUTABLE)
    
    If (sInstalledVersion = "") Then
        ' If the version string is empty, we could not determine the version info, most likely
        ' because the .exe file is not in the directory.
        MsgBox "No CCD application file found in the launcher directory.  Please contact the CCD Administrator.", vbOKOnly, _
            "CCD Launcher"
        Exit Sub
    End If
            
    sSQL = "EXEC usp_version_check @version='" & sInstalledVersion & "'"
    g_objDAL.CacheConnection (strConnect)
    bResult = g_objDAL.GetRecordset(strConnect, sSQL, rec)
    If bResult Then
        If rec.EOF Then
            ' The row in the versions table is missing for this exe version.  A row must be
            ' added that matches the version of the ccd app.  In CCD, go to System Admin | Version Admin.
            ' Deprecate the previous version and add the new version.
            MsgBox "No version information found in the database for application version " & _
                sInstalledVersion & ".  Please contact the CCD Administrator.", vbOKOnly, _
                "CCD Launcher"
            Exit Sub
        Else
            sCurrentVersion = rec.Fields("current_release")
            If sCurrentVersion = sInstalledVersion And Not bCheckOnly Then
                ' VERSION IS CURRENT, LAUNCH APP
                bOkForLaunch = True
            Else
                ' DISPLAY VERSION INFORMATION DIALOG
                dtCurrentReleaseDate = rec.Fields("current_release_date")
                sDescription = rec.Fields("status_header")
                sMessage = rec.Fields("status_text")
                sMessage = sMessage & vbCrLf & vbCrLf
                sMessage = sMessage & "Installed version:  " & sInstalledVersion & vbCrLf
                sMessage = sMessage & "Current release:   " & sCurrentVersion & vbCrLf
                Dim frm As New frmVersionCheck
                ' CURRENT STATUS=1
                If rec.Fields("status") = 1 Then
                    frm.cmdContinue.Visible = False
                    frm.cmdUpgrade.Visible = False
                    frm.cmdCancel.Caption = "Close"
                Else
                    ' UNSUPPORTED STATUS=4, DO NOT ALLOW CONTINUE
                    frm.cmdContinue.Enabled = (rec.Fields("status") <> 4)
                    ' BETA STATUS=2, DO NOT ALLOW UPGRADE
                    frm.cmdUpgrade.Enabled = (rec.Fields("status") <> 2)
                End If
                frm.Header = sDescription
                frm.Content = sMessage
                frm.Show vbModal
                If frm.Result = vbAbort Then
                    ' ABORT LAUNCH
                    Dim sUpgradePath As String
                    If Not IsNull(rec.Fields("current_upgrade_path")) Then
                        sUpgradePath = rec.Fields("current_upgrade_path")
                        If CopyFolderToAppPath(sUpgradePath) Then
                            ' FILE COPY WAS SUCCESSFUL, SO LAUNCH...
                            bOkForLaunch = True
                        Else
                            ' FILE COYP FAILED, OPEN UPGRADE SITE...
                            bOkForLaunch = False
                            LaunchBrowser UPGRADE_WEB_SITE
                        End If
                    Else
                        ' OPEN UPGRADE SITE...
                        bOkForLaunch = False
                        LaunchBrowser UPGRADE_WEB_SITE
                    End If
                ElseIf frm.Result = vbIgnore Then
                    ' IGNORE MESSAGE, CONTINUE LAUNCH...
                    bOkForLaunch = True
                End If
                Set frm = Nothing
            End If
        End If
        rec.Close
    Else
        MsgBox g_objDAL.LastErrorDescription, vbCritical
    End If
    Set rec = Nothing
    Set g_objDAL = Nothing
    If bOkForLaunch And Not bCheckOnly Then
        ' LAUNCH CCD APPLICATION
        ' PASS COMMAND LINE OPTIONS TO CCD
        res = ShellExecute(0&, "open", sAppPath & CCD_EXECUTABLE, Command$, sAppPath, vbNormalFocus)
        If res > 32 Then
            Call BringWindowToTop(res)
        End If
    End If
    Exit Sub
    
Err_Handler:
    ' AN ERROR OCCURRED; ATTEMPT TO LAUNCH CCD IF NOT IN VERSION CHECK MODE
    MsgBox Err.Description, vbCritical, "CCD"
    If Not bCheckOnly Then
        res = ShellExecute(0&, "open", sAppPath & CCD_EXECUTABLE, Command$, sAppPath, vbNormalFocus)
        If res > 32 Then
            Call BringWindowToTop(res)
        End If
    End If
    Exit Sub
    
End Sub

