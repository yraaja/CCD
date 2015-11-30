VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cUserInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================
'Description:
'This class serves as a CCD Database User object
'gets/sets user information from the application
'---------------------------------------------
'Dependencies:
'None
'---------------------------------------------
'Author: Rob Durfee
'Date: 06/22/2005
'Change Description: Initial Version
'==============================================

Option Explicit

Private m_rec As ADODB.RecordSet
Private bNewUser As Boolean
Private sUserId As String
Private sUserName As String
Private sUserExt As String
Private sUserFax As String
Private iUserRole As Integer
Private sVersion As String
Private dtLastLogin As Date
Private sLastError As String

Private Enum UserRoleBits
    ROLE_BIT_USER = 0
    ROLE_BIT_ADMIN = 128
End Enum

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'****************************************************************
'** PRIVATE FUNCTIONS
'****************************************************************

Private Function SQLDelimitString(ByVal sString As String) As String
    Dim sTemp As String
    If sString = "" Then
        sTemp = "NULL"
    Else
        sTemp = "'" & SQLFixString(sString) & "'"
    End If
    SQLDelimitString = sTemp
End Function


'****************************************************************
'** PUBLIC PROPERTIES
'****************************************************************

Public Property Get LastError() As String
    LastError = sLastError
End Property

Public Property Get isAdmin() As Boolean
    ' *** TBD
    isAdmin = (iUserRole And UserRoleBits.ROLE_BIT_ADMIN = UserRoleBits.ROLE_BIT_ADMIN)
End Property

Public Property Get isNewUser() As Boolean
    isNewUser = bNewUser
End Property

Public Property Get UserID() As String
    UserID = sUserId
End Property
Public Property Let UserID(NewValue As String)
    sUserId = NewValue
End Property

Public Property Get UserName() As String
    UserName = sUserName
End Property
Public Property Let UserName(NewValue As String)
    sUserName = NewValue
End Property

Public Property Get UserExtension() As String
    UserExtension = sUserExt
End Property
Public Property Let UserExtension(NewValue As String)
    sUserExt = NewValue
End Property

Public Property Get UserFaxNumber() As String
    UserFaxNumber = sUserFax
End Property
Public Property Let UserFaxNumber(NewValue As String)
    sUserFax = NewValue
End Property

Public Property Get UserRole() As Long
    UserRole = iUserRole
End Property
Public Property Let UserRole(NewValue As Long)
    iUserRole = NewValue
End Property

Public Property Get UserRoleDescription() As String
    Dim sRole As String
    
    ' *** TBD
    If iUserRole = 0 Then
        sRole = "User"
    Else
        If (iUserRole And UserRoleBits.ROLE_BIT_USER) = UserRoleBits.ROLE_BIT_USER Then
            sRole = sRole & "User;"
        End If
        If (iUserRole And UserRoleBits.ROLE_BIT_ADMIN) = UserRoleBits.ROLE_BIT_ADMIN Then
            sRole = sRole & "Admin;"
        End If
    End If
    UserRoleDescription = sRole
    
End Property

Public Property Get ApplicationVersion() As String
    ApplicationVersion = sVersion
End Property
Public Property Let ApplicationVersion(NewValue As String)
    sVersion = NewValue
End Property

Public Property Get LastLoginTimestamp() As Date
    LastLoginTimestamp = dtLastLogin
End Property
Public Property Let LastLoginTimestamp(NewValue As Date)
    dtLastLogin = NewValue
End Property

Public Property Get LastLogin() As String
    If dtLastLogin = 0 Then
        LastLogin = ""
    Else
        LastLogin = dtLastLogin
    End If
End Property
Public Property Let LastLogin(NewValue As String)
    If IsDate(NewValue) Then
        dtLastLogin = NewValue
    End If
End Property

Public Property Get UserWorkstation() As String
    Dim sBuffer As String
    Dim iSize As Long

    sBuffer = Space$(256)
    iSize = Len(sBuffer)
    GetComputerName sBuffer, iSize
    sBuffer = Left(sBuffer, iSize)
    UserWorkstation = sBuffer

End Property

Public Property Get UserDomain() As String
    Dim sBuffer As String
    
    sBuffer = Environ("USERDOMAIN")
    UserDomain = sBuffer
    
End Property

'****************************************************************
'** PUBLIC METHODS
'****************************************************************

Public Function GetData() As Boolean
    Dim sSelect As String
    
    On Error GoTo Err_Handler
    'If sUserId = "" Then
    '    Err.Raise 1001, , "Invalid or Null user name specified."
    'End If
    Screen.MousePointer = vbHourglass
    sSelect = "SELECT * FROM USER_NAMES WHERE USER_ID='" & sUserId & "'"
    g_objDAL.GetRecordset vbNullString, sSelect, m_rec
    If m_rec.RecordCount > 0 Then
        sUserId = sUserId
        If Not IsNull(m_rec.Fields("user_name")) Then
            sUserName = m_rec.Fields("user_name")
        Else
            sUserName = ""
        End If
        If Not IsNull(m_rec.Fields("user_extension")) Then
            sUserExt = m_rec.Fields("user_extension")
        Else
            sUserExt = ""
        End If
        If Not IsNull(m_rec.Fields("user_fax")) Then
            sUserFax = m_rec.Fields("user_fax")
        Else
            sUserFax = ""
        End If
        If Not IsNull(m_rec.Fields("user_role")) Then
            iUserRole = m_rec.Fields("user_role")
        Else
            iUserRole = 0
        End If
        If Not IsNull(m_rec.Fields("ccd_version")) Then
            sVersion = m_rec.Fields("ccd_version")
        Else
            sVersion = ""
        End If
        If Not IsNull(m_rec.Fields("last_logon")) Then
            dtLastLogin = m_rec.Fields("last_logon")
        Else
            dtLastLogin = Now
        End If
    Else
        sUserId = sUserId
         sUserName = ""
        sUserExt = ""
        sUserFax = ""
        sVersion = ""
        dtLastLogin = 0
        bNewUser = True
    End If
    sLastError = Err.Description
    GetData = True
    Screen.MousePointer = vbDefault
    Exit Function

Err_Handler:
    Screen.MousePointer = vbDefault
    MsgBox "CUserInfo.GetData() Error #" & Err.Number & vbCrLf & Err.Description, vbCritical
    sLastError = Err.Description
    GetData = False
    Exit Function
End Function

Public Function UpdateData() As Boolean

    Dim sSelect As String
    Dim iRecs As Long
    
    On Error GoTo Err_Handler
    If sUserId = "" Then
        Err.Raise 1001, , "Invalid or Null user name specified."
    End If
    Screen.MousePointer = vbHourglass
    If bNewUser Then
        sSelect = "INSERT INTO USER_NAMES " & _
                " (USER_ID, USER_NAME, USER_EXTENSION, USER_FAX, LAST_UPDATE_DATE, LAST_UPDATE_PERSON)" & _
                " VALUES ('" & sUserId & "'," & _
                " " & SQLDelimitString(sUserName) & "," & _
                " " & SQLDelimitString(sUserExt) & "," & _
                " " & SQLDelimitString(sUserFax) & "," & _
                " '" & Format(Now, "General Date") & "'," & _
                " " & SQLDelimitString(strUserName) & ")"
        
    Else
        sSelect = "UPDATE USER_NAMES " & _
                " SET USER_NAME=" & SQLDelimitString(sUserName) & "," & _
                " USER_EXTENSION=" & SQLDelimitString(sUserExt) & "," & _
                " USER_FAX=" & SQLDelimitString(sUserFax) & "," & _
                " LAST_UPDATE_DATE='" & Format(Now, "General Date") & "'," & _
                " LAST_UPDATE_PERSON='" & strUserName & "'" & _
                " WHERE USER_ID='" & sUserId & "'"
    End If
    g_cnShared.Execute sSelect, iRecs
    sLastError = Err.Description
    Screen.MousePointer = vbDefault
    bNewUser = False
    UpdateData = True
    Exit Function
    
Err_Handler:
    Screen.MousePointer = vbDefault
    MsgBox "CUserInfo.UpdateData() Error #" & Err.Number & vbCrLf & Err.Description, vbCritical
    sLastError = Err.Description
    UpdateData = False
    Exit Function
End Function

Public Function Login() As Boolean
    Dim sSelect As String
    Dim iRecs As Long
    Dim msg As String
    
    On Error GoTo Err_Handler
    If sUserId = "" Then
        Err.Raise 1001, , "Invalid or Null user name specified."
    End If
    ' UPDATE VERSION AND LAST LOGIN FIELDS
    Screen.MousePointer = vbHourglass
    GetData
    sVersion = App.Major & "." & App.Minor & "." & App.Revision
    dtLastLogin = Now
    If bNewUser Then
        'new/unauthorized user - cje
        msg = "You are currently unauthorized to access this application." & vbCrLf
        msg = msg & "Please contact the CCD Administrator for logon permissions."
        MsgBox msg, vbOKOnly, "Construction Cost Database"
        
        Login = False
        Exit Function
    
'        sSelect = "INSERT INTO USER_NAMES " & _
'                " (USER_ID, USER_NAME, USER_EXTENSION, USER_FAX, CCD_VERSION, LAST_LOGON)" & _
'                " VALUES ('" & sUserId & "'," & _
'                " " & SQLDelimitString(sUserName) & "," & _
'                " " & SQLDelimitString(sUserExt) & "," & _
'                " " & SQLDelimitString(sUserFax) & "," & _
'                " " & SQLDelimitString(sVersion) & ", '" & Format(dtLastLogin, "General Date") & "')"
                
    Else
        sSelect = "UPDATE USER_NAMES " & _
                " SET CCD_VERSION=" & SQLDelimitString(sVersion) & "," & _
                " LAST_LOGON=" & SQLDelimitString(dtLastLogin) & "" & _
                " WHERE USER_ID='" & sUserId & "'"
    End If
    g_cnShared.Execute sSelect, iRecs
    sLastError = Err.Description
    Screen.MousePointer = vbDefault
    bNewUser = False
    Login = True
    Exit Function
    
Err_Handler:
    Screen.MousePointer = vbDefault
    MsgBox "CUserInfo.Login() Error #" & Err.Number & vbCrLf & Err.Description, vbCritical
    sLastError = Err.Description
    Login = False
    Exit Function

End Function

Private Sub Class_Initialize()
    
    sUserId = ""
    sUserName = ""
    sUserExt = ""
    sUserFax = ""
    sVersion = ""
    dtLastLogin = 0
    sLastError = ""
    bNewUser = False
    iUserRole = 0
    
End Sub
