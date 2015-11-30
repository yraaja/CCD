Attribute VB_Name = "RegistryModule"
Option Explicit

Global Const REG_SZ As Long = 1
Global Const REG_DWORD As Long = 4

Global Const HKEY_CLASSES_ROOT = &H80000000
Global Const HKEY_CURRENT_USER = &H80000001
Global Const HKEY_LOCAL_MACHINE = &H80000002
Global Const HKEY_USERS = &H80000003

Global Const ERROR_NONE = 0
Global Const ERROR_BADDB = 1
Global Const ERROR_BADKEY = 2
Global Const ERROR_CANTOPEN = 3
Global Const ERROR_CANTREAD = 4
Global Const ERROR_CANTWRITE = 5
Global Const ERROR_OUTOFMEMORY = 6
Global Const ERROR_INVALID_PARAMETER = 7
Global Const ERROR_ACCESS_DENIED = 8
Global Const ERROR_INVALID_PARAMETERS = 87
Global Const ERROR_NO_MORE_ITEMS = 259

Public Const KEY_QUERY_VALUE = &H1
Global Const KEY_ALL_ACCESS = &H3F
Global Const KEY_READ = &H20019
Global Const KEY_WRITE = &H20006
Global Const REG_OPTION_NON_VOLATILE = 0

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Declare Function RegDeleteKey& Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String)
Declare Function RegDeleteValue& Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String)

Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As _
String, vValue As Variant) As Long
   Dim cch As Long
   Dim lrc As Long
   Dim lType As Long
   Dim lValue As Long
   Dim sValue As String

   On Error GoTo QueryValueExError

   ' Determine the size and type of data to be read
   lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
   If lrc <> ERROR_NONE Then Error 5

   Select Case lType
       ' For strings
       Case REG_SZ:
           sValue = String(cch, 0)

lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, _
sValue, cch)
           If lrc = ERROR_NONE Then
               vValue = left$(sValue, cch - 1)
           Else
               vValue = Empty
           End If
       ' For DWORDS
       Case REG_DWORD:
lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, _
lValue, cch)
           If lrc = ERROR_NONE Then vValue = lValue
       Case Else
           'all other data types not supported
           lrc = -1
   End Select

QueryValueExExit:
   QueryValueEx = lrc
   Exit Function

QueryValueExError:
   Resume QueryValueExExit
End Function

Public Function QueryRegistryKey(ROOTKEY As Long, sKeyName As String, sValueName As String, Optional vDefaultValue As Variant)
    Dim lRetVal As Long
    Dim hKey As Long
    Dim vValue As Variant
    
    lRetVal = RegOpenKeyEx(ROOTKEY, sKeyName, 0, KEY_QUERY_VALUE, hKey)
    lRetVal = QueryValueEx(hKey, sValueName, vValue)
    If (lRetVal <> ERROR_NONE) And (IsEmpty(vValue)) Then
        vValue = vDefaultValue
    End If
    RegCloseKey (hKey)
    QueryRegistryKey = vValue
   
End Function


