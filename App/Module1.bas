Attribute VB_Name = "MainModule"
Option Explicit

''' <modulename> MainModule</modulename>
''' <functionname>General (Main) </functionname>
'''
''' <summary>
'''This is a source code component containing subs and functions available anywhere in CCD.  There are many "key" subs/functions that are routinely used/accessed throughout CCD.  I will attempt to identify some of the most widely used ones.  First, the list.:
'''
'''GLOBAL SUB/FUNCTION SETUP
'''
'''"   WINDOWS APIs    Declare them
'''"   Main()   -  Sets up "global" connections for database accessibility to be used throughout CCD
'''"   Login()  -  Sets "global"
'''(e.g. g_objDAL or g_cnShared or  g_cnSharedLong)
'''"   LoadMasterFormatCombo()
'''LoadCombos      Builds/populates comboboxes routinely displayed on CCD forms (e.g. "quarter dates", "Countries", "States", "Equipment", "Labor trades ids",
'''"   LoadRegKeys()   Check REGISTRY for session values (MaxRecords, User Name, DBServer, Dbase etc.   Each can be reset on Toolbar "Settings" menu(s)
'''"   LoadRegKeysQtr()    Same stuff saved/retrieved on REGISTRY
'''"   Status()        Update of main form "status bar"
'''"   UpdateFormFromRecordset         Populate window grid/maintenance forms from the "master" recordset
'''"   UpdateRecordsetFromForm     Populate a recordset from the form
'''"   Note- the aforementioned (2) subs take advantage of the fact that the form "control" names are the same as the recordset/db names being updated or retrieved
'''
'''
'''HELPER Class: cUserInfo.Cls
'''
'''</summary>
'''
'''<seealso> cUserInfo</seealso>
'''
'''
''' <datastruct>m_objGridMap</datastruct>
'''<datastruct>m_rec</datastruct>
'''
''' <storedprocedurename> usp_select_unit_cost_ext</storedprocedurename>
'''<storedprocedurename> usp_update_unit_cost_driver_ext</storedprocedurename>
'''<storedprocedurename> sp_select_assembly</storedprocedurename>
'''<storedprocedurename> sp_update_assembly_driver</storedprocedurename>
'''
'''
'''<returns>N/A</returns>
''' <exception>Always trap with an accompanying message box</exception>
''' <example>
'''<code>
'''Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
'''Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
'''Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long
'''Declare Function GetWindowPlacement Lib "user" (ByVal hWnd As Integer, lpwndpl As WINDOWPLACEMENT) As Integer
'''Declare Function SetWindowPlacement Lib "user" (ByVal hWnd As Integer, lpwndpl As WINDOWPLACEMENT) As Integer
'''Private Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" _
'''  (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
'''Private Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" _
'''  (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
'''Private Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" _
'''  (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, nVerSize As Long) As Long
'''Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'''Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
'''Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long
'''Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
'''</code>
'''<code>
'''Sub Main()
'''    Dim blnServerChanged As Boolean
'''    Dim intResult As Integer
'''
'''    On Error GoTo Error_Processing
'''    InitCommonControlsVB
'''    Screen.MousePointer = vbHourglass
'''    frmSplash.Show
'''    frmSplash.Refresh
'''
'''    frmSplash.SetStatus "Initializing..."
'''
'''    Call LoadRegKeys
'''
'''Continue_Login:
'''    'rlh 05/22/2007  USE FOR PRODUCTION/DISABLED FOR TEST/DEBUG
'''    'strConnect = "UID=" + strUserName + ";PWD=;DATABASE=" + strConnectDatabase + ";SERVER=" + strConnectServer + ";DRIVER={SQL SERVER};DSN='';"
'''
''''rlh 05/22/2007  There's a problem trying to login using windows authentication uid/password
'''    '#################################################################
''''##
''''##  you MUST user ccduser to "get in the door".  Any changes made
''''##  to the data will be noted with the windows authentication
''''##  userid (eg. hancockrl)
''''##  NOTE: the CCDLAUNCH application also uses CCDUSER !!!
''''##    '#################################################################
'''
'''strConnect = "UID=ccdUser;PWD=rsmeans;DATABASE=" + strConnectDatabase + ";SERVER=" + strConnectServer + ";DRIVER={SQL SERVER};DSN='';"
'''
'''    frmSplash.SetStatus "Connecting to database(1)..."
'''
'''On Error GoTo Connect_Error_Processing
'''    'g_cnShared.ConnectionTimeout = 5000        'rlh  05/22/2007
'''    'g_cnShared.CommandTimeout = 5000           'rlh
'''    'g_cnShared.Open CONNECT                    'rlh
'''    'g_cnSharedLong.ConnectionTimeout = 15000   'rlh
'''    'g_cnSharedLong.CommandTimeout = 15000      'rlh
'''
'''     g_cnShared.ConnectionTimeout = 0           'rlh  05/22/2007 unlimited timeout (Mel Mossman)
'''    g_cnShared.CommandTimeout = 0
'''    g_cnShared.Open CONNECT
'''    g_cnSharedLong.ConnectionTimeout = 0
'''    g_cnSharedLong.CommandTimeout = 0
'''    g_cnSharedLong.Open CONNECT
'''
'''    Call LoadRegKeysQtr
'''    If blnServerChanged = True Then
'''        intResult = MsgBox("Do you want to retain these settings?", vbYesNo, "Save Server Settings")
'''        If intResult = vbYes Then       'Save registry settings
'''            Dim strKey As String
'''            Dim hKey As Long
'''            Dim lSize As Long
'''            Dim lRet As Long
'''            lSize = 1000
'''                'Save Server settings
'''            strKey = CCD_KEY + "\Defaults\DBServer"
'''            lSize = Len(strConnectServer)
'''            lRet = RegOpenKeyEx(HKEY_CURRENT_USER, strKey, 0&, KEY_ALL_ACCESS, hKey)
'''            lRet = RegSetValueExString(hKey, "Value", 0&, REG_SZ, strConnectServer, lSize)
'''            RegCloseKey (hKey)
'''                'Save Database settings
'''            strKey = CCD_KEY + "\Defaults\DBase"
'''            lSize = Len(strConnectDatabase)
'''            lRet = RegOpenKeyEx(HKEY_CURRENT_USER, strKey, 0&, KEY_ALL_ACCESS, hKey)
'''            lRet = RegSetValueExString(hKey, "Value", 0&, REG_SZ, strConnectDatabase, lSize)
'''            RegCloseKey (hKey)
'''        End If
'''    End If
'''
'''    On Error GoTo Error_Processing
'''
'''    frmSplash.SetStatus "Connecting to database(2)..."
'''
'''    ' Cache a global connection
'''    g_objDAL.CacheConnection (CONNECT)
'''
'''    'validate user in CCD user_name table - cje
'''</code>
'''<code>
'''Public Function LoadMasterFormatCombo(ByRef Combo1 As ComboBox, Optional bNoAltIDSelection As Boolean = False) As Long
'''' LOAD GIVEN COMBOBOX WITH MASTERFORMAT INFORMATION
'''' SELECT THE ITEM THAT CORRESPONDS WITH USER'S DEFAULT MASTERFORMAT SETTING
'''    Dim iIndex As Long
'''    Dim sDefaultMF As String
'''
'''    Combo1.Clear
'''
'''    ' Get Default MasterFormat and select in Combo
'''    sDefaultMF = QueryRegistryKey(HKEY_CURRENT_USER, CCD_KEY & "\Defaults\MasterFormat", "Value", CStr(UCD_MASTERFORMAT_VERSION))
'''
'''    Combo1.AddItem "MF-" & EXT_MASTERFORMAT_VERSION
'''    Combo1.ItemData(Combo1.NewIndex) = EXT_MASTERFORMAT_VERSION
'''    If sDefaultMF = EXT_MASTERFORMAT_VERSION Then
'''        Combo1.ListIndex = Combo1.NewIndex
'''    End If
'''
'''    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'''    '
'''    ' FLIP THIS SWITCH (MF95_ENABLED) AND YOU CAN TOGGLE BACK AND FORTH
'''    ' FROM SUPPORT/NO SUPPORT FOR MF95 PROCESSING!!!
'''    '
'''    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'''    If SUPER_USER_SUPPORT Then
'''    Else
'''        MF95_ENABLED = False     'rlh  02/19/2009
'''    End If
'''    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'''
'''
'''    If MF95_ENABLED Then    'rlh  02/19/2009
'''        Combo1.AddItem "MF-" & UCD_MASTERFORMAT_VERSION
'''        Combo1.ItemData(Combo1.NewIndex) = UCD_MASTERFORMAT_VERSION
'''        If sDefaultMF = UCD_MASTERFORMAT_VERSION Then
'''            Combo1.ListIndex = Combo1.NewIndex
'''        End If
'''
'''        If Not bNoAltIDSelection Then
'''            Combo1.AddItem "MF-" & ALT_MASTERFORMAT_VERSION
'''            Combo1.ItemData(Combo1.NewIndex) = ALT_MASTERFORMAT_VERSION
'''            If sDefaultMF = ALT_MASTERFORMAT_VERSION Then
'''                Combo1.ListIndex = Combo1.NewIndex
'''            End If
'''        End If
'''    End If 'rlh
'''    LoadMasterFormatCombo = iIndex
'''
'''End Function
'''</code>
'''</example>
'''<permission>Public</Permission>
'''<dependson>This component depends on the following
'''1.  cUserInfo.cls
'''2.  WINDOWS APIs
'''3.  CCDdal.CRSMDataAccess (
'''Access to the DAL (data access layer dll) opened in MainModule_Main() )
'''4.  REGISTRY
'''</dependson>





'8/16/2005 RTD
'UCD_MASTERFORMAT_VERSION defines the MasterFormat version of the
'UNIT_COST_ID field in the UNIT_COST_DETAIL table
Public Const UCD_MASTERFORMAT_VERSION As Integer = 1995
'ALT_MASTERFORMAT_VERSION defines the MasterFormat version of the
'ALT_UNIT_COST_ID field in the UNIT_COST_DETAIL table
Public Const ALT_MASTERFORMAT_VERSION As Integer = 1988
'EXT_MASTERFORMAT_VERSION defines the MasterFormat version of the
'UNIT_COST_ID field in the UNIT_COST_DETAIL_EXT table
Public Const EXT_MASTERFORMAT_VERSION As Integer = 2004
Public MF95_ENABLED As Boolean          'rlh 2/19/2009
Public SUPER_USER_SUPPORT As Boolean    'rlh 4/09/2009

Public MASTER_FORMAT_ASSEMBLIES As Integer

Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Private Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Integer) As Integer
Public Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags_ As Long) As Long

Private Const ICC_USEREX_CLASSES = &H200
Private Const COLOR_WINDOW = 5
Private Const COLOR_WINDOWTEXT = 8
Private Const COLOR_APPWORKSPACE = 12
Private Const COLOR_BTNFACE = 15
Private Const COLOR_GRAYTEXT = 17
Private Const COLOR_BTNTEXT = 18

Public Const cSndASYNC = &H1
Public Const cSndNODEFAULT = &H2
Private Const cSndSYNC = &H0
Private Const cSndASYNCH = &H1
Private Const cSndLoop = &H8
Private Const cSndNOSTOP = &H10

' Server and name
'Global Const CONN_SERVER = "nt_hanscomb"
' Server and Database Settings
Public strConnectServer As String
Public CONN_USERPW As String
Public strConnectDatabase As String
Public strUserName As String
Public strConnect As String

Global Const LTGREY = &H8000000F  ' &HC0C0C0
'Global Const CONNECT = "DSN=CCD;UID=sa;PWD=;"
Global Const DELETED = "1"

' Starting position of windows
Global Const START_TOP = 0
Global Const NAV_TREE_WIDTH = 2385 ' Width of the NavTree form

' Set some constant values (from WIN32API.TXT).
Global Const conHwndTopmost = -1
Global Const conHwndNoTopmost = -2
Global Const conSwpNoActivate = &H10
Global Const conSwpShowWindow = &H40
Global Const conSwpNoSize = &H1
Global Const conSwpNoMove = &H2

' These are used by CRowInfo
Global Const STATE_NOT_SET = 0
Global Const STATE_NONE = 1
Global Const STATE_NEW = 2
Global Const STATE_MODIFIED = 3
Global Const STATE_DELETED = 4
Global Const STATE_PUBLISHED = 5

' These identify the formatting options for columns in CColumnDef
Global Const FORMAT_PRICE = 1
Global Const FORMAT_DATE = 2
Global Const FORMAT_DATETIME = 3
Global Const FORMAT_DECIMAL = 4
Global Const FORMAT_DECIMAL3 = 5
Global Const FORMAT_MATERIAL = 6        'M12345 123 1234 for readability
Global Const FORMAT_UNIT_COST = 7       '12345 123 1234 for readability
Global Const FORMAT_UNIT_COST_04 = 8    '12 34 56.78 1234 for readability
Global Const FORMAT_STRING_TRIM = 9     'Format the string using Trim() to strip spaces
Global Const FORMAT_STRING_URL = 10     'Format the string as a URL hyperlink
Global Const FORMAT_CHECK_BOX = 11      'Format the field as a checkbox

' These are the grid format strings for 6-8 above
Global Const FORMAT_UNIT_COST_SRV = "@@@@@ @@@ @@@@" ' Format for Single Record View
Global Const FORMAT_UNIT_COST_04_SRV = "@@ @@ @@.@@ @@@@" ' Format for Single Record View
Global Const FORMAT_MATERIAL_SRV = "@@@@@@ @@@ @@@@" ' Format for Single Record View
Global Const FORMAT_MATERIAL_04_SRV = "@@@ @@ @@.@@ @@@@" ' Format for Single Record View

' These are used by EquipmentRate to determine how to apply a factor
Global Const EQUIP_FACTOR_RENT = 1
Global Const EQUIP_FACTOR_OPERATING = 2
Global Const EQUIP_FACTOR_BOTH = 3

' These are sort directions
Global Const SORT_ASCENDING = 1
Global Const SORT_DESCENDING = 2

' Max records to return on search
Public MAX_RECORDS As Long

' Material Price Rollup constants
Global Const ALWAYS_ROLLUP_MATERIAL = 1
Global Const USER_ROLLUP_MATERIAL = 2
Global Const NEVER_ROLLUP_MATERIAL = 3

'Types of selection available in the List Selection class/dialog
Global Const SINGLE_LIST = 1
Global Const AVAILABLE_AND_SELECTED_LISTS = 2
Global Const COMBO_BOX = 3

'Main Menu Toolbar Button Indexes - referenced by grid forms to enable/disable buttons
Global Const tbrNEW = 1
Global Const tbrOPEN = 2
Global Const tbrSAVE = 3
Global Const tbrPRINT = 4
Global Const tbrPREVIEW = 5
Global Const tbrEXPORT = 6
Global Const tbrFAX = 7
Global Const tbrEMAIL = 8
Global Const tbrCUT = 10
Global Const tbrCOPY = 11
Global Const tbrPASTE = 12
Global Const tbrDELETE = 13
Global Const tbrUNDO = 14
Global Const tbrFIND = 16
Global Const tbrEXPORTDATA = 17
Global Const tbrSORT_ASCENDING = 19
Global Const tbrSORT_DESCENDING = 20
Global Const tbrPRINTSCREEN = 22
Global Const tbrHELP = 24

'True DB Grid Postevent Message IDs
Global Const tdbgENABLE_SORT = 100

Public g_intRollupOption As Integer
Public g_blnMaximize As Boolean
Public g_blnWhiteBackground As Boolean
Public g_blnFlatToolbar As Boolean
Public g_intMasterFormat As Long
Public g_blnAlternateRow As Boolean
Public g_intAlternateRowColor As Long
Public g_blnUseAlternateDisabledColor As Boolean
' 10/18/2005 RTD - Determine if user is CCD Admin
Public g_blnIsUserAdmin As Boolean

Global Const CCD_KEY = "RSMeans\CCD"

Public fMainForm As frmMain

Public g_objDAL As New CCDdal.CRSMDataAccess ' Global DAL object
Public g_cnShared As New ADODB.Connection
Public g_cnSharedLong As New ADODB.Connection
Public g_sQuarterID As String
Public m_strKeyType2 As String               'rlh -07/14/2008
Private intStartLeft As Integer

'8/15/2005 RTD - enum used by OutputUsageFormat property (dlgOutput/COutputMap)
Public Enum OUTPUT_USAGE_FORMAT
    OUTPUT_MF2004_ONLY = 1
    OUTPUT_MF1995_ONLY = 2
    OUTPUT_BOTH = 3
End Enum

Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type

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

Type PointAPI
   X As Integer
   Y As Integer
End Type

Type RECT
   Left As Integer
   Top As Integer
   Right As Integer
   Bottom As Integer
End Type

Type WINDOWPLACEMENT
   Length As Integer
   flags As Integer
   showCmd As Integer
   ptMinPosition As PointAPI
   ptMaxPosition As PointAPI
   rcNormalPosition As RECT
End Type

Global lpwndplOld As WINDOWPLACEMENT
Global lpwndplNew As WINDOWPLACEMENT

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long
Declare Function GetWindowPlacement Lib "user" (ByVal hWnd As Integer, lpwndpl As WINDOWPLACEMENT) As Integer
Declare Function SetWindowPlacement Lib "user" (ByVal hWnd As Integer, lpwndpl As WINDOWPLACEMENT) As Integer
Private Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" _
  (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" _
  (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" _
  (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, nVerSize As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)

Public Const CSIDL_PERSONAL = &H5                'user My Documents
Public Const CSIDL_DESKTOPDIRECTORY = &H10       'user Desktop
Public Const CSIDL_LOCAL_APPDATA = &H1C          'user Application Data

Private Const MAX_PATH = 260
Global Const CB_FINDSTRING = &H14C
Global Const CB_FINDSTRINGEXACT = &H158
Global Const WM_ACTIVATE As Long = &H6
Global Const SW_SHOWNORMAL = 1
Global hWnds() As Long
Public m_blnNew_pub As Boolean
Public m_blnClone_pub As Boolean
Global Const DEBUGON As Boolean = False         'rlh
Public strLaborSelectedQtr As String            'rlh
Public Mode As Boolean                          'rlh GO or CANCEL
Public old_mat_skey_text As String              'rlh 07/23/2009  (Dave Drain issue)

Type structTradeGroup
    Trade_Group_Code As String
    Trade_ID As String
    City As String
    State_Code As String
    start_date As Date
    term_date As Date
    union_base As Double
    union_fring As Double
    tot_union As Double
End Type

''''' Error numbers
Public Const ERROR_CLONING_ASSEMBLIES = 100
'''''

Public Function InitCommonControlsVB() As Boolean
'INITIALIZE THE WINDOWS COMMON CONTROLS COMCTL32.DLL
'TO ENABLE WINDOWS XP THEME AND COLOR SCHEME SUPPORT
    Dim iccex As tagInitCommonControlsEx
    
    On Error Resume Next
    ' Ensure CC available:
    With iccex
        .lngSize = LenB(iccex)
        .lngICC = ICC_USEREX_CLASSES
    End With
    InitCommonControlsEx iccex
    InitCommonControlsVB = (Err.Number = 0)

End Function

Public Function FillRptParm(sValue As String, Optional bWildcard As Boolean = False) As Variant
    If bWildcard Then
        FillRptParm = """" + FillWildCard(SQLChangeWildcard(sValue)) + """"
    Else
        FillRptParm = """" + SQLChangeWildcard(sValue) + """"
    End If
End Function

Public Sub LoadCombos(frm As Form, _
            Optional bQtr As Boolean = False, _
            Optional bCountry As Boolean = False, _
            Optional bState As Boolean = False, _
            Optional bTrades As Boolean = False, _
            Optional bEquipment As Boolean = False)
    Dim blnReturn As Boolean
    Dim strSelect As String
    Dim rsTemp As RecordSet
    On Error Resume Next
    'Load All Selection Combos
    If bQtr Then
    'Load Quarter IDs
        strSelect = "SELECT quarter_id FROM QUARTER_DATE ORDER BY quarter_id"
        ' Use DAL to perform select
        blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
        If blnReturn = False Then
            MsgBox "An error occurred loading Trade IDs."
            frm.lblRowCount.Caption = "0 rows returned."
        Else
            If Not (rsTemp.EOF And rsTemp.BOF) Then
                Do Until rsTemp.EOF
                    With frm
                        .cmbQuarterID.AddItem rsTemp![quarter_id]
                        If rsTemp![quarter_id] = g_sQuarterID Then
                            .cmbQuarterID.Text = .cmbQuarterID.List(.cmbQuarterID.NewIndex)
                            .cmbQuarterID.ListIndex = .cmbQuarterID.NewIndex
                        End If
                    End With
                    rsTemp.MoveNext
                Loop
            End If
        End If
        rsTemp.Close
        Set rsTemp = Nothing
    End If
    'Load Countries
    If bCountry Then
        strSelect = "select distinct country_code from location order by country_code;"
        ' Use DAL to perform select
        blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
        If blnReturn = False Then
            MsgBox "An error occurred loading States."
        Else
            If Not (rsTemp.EOF And rsTemp.BOF) Then
                Do Until rsTemp.EOF
                    With frm
                        .cmbCountry.AddItem rsTemp![country_code]
                    End With
                    rsTemp.MoveNext
                Loop
            End If
        End If
        rsTemp.Close
        Set rsTemp = Nothing
    End If
    'Load States
    If bState Then
        strSelect = "select distinct state_code from location order by state_code;"
    
        ' Use DAL to perform select
        blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
        If blnReturn = False Then
            MsgBox "An error occurred loading States."
        Else
            If Not (rsTemp.EOF And rsTemp.BOF) Then
                Do Until rsTemp.EOF
                    With frm
                        .cmbState.AddItem rsTemp![State_Code]
                    End With
                    rsTemp.MoveNext
                Loop
            End If
        End If
        rsTemp.Close
        Set rsTemp = Nothing
    End If
    If bEquipment Then
    'Load Equipment
        strSelect = "select distinct cci_equip_id from CCI_EQUIPMENT order by cci_equip_id;"
        ' Use DAL to perform select
        blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
        If blnReturn = False Then
            MsgBox "An error occurred loading Equipment."
        Else
            If Not (rsTemp.EOF And rsTemp.BOF) Then
                Do Until rsTemp.EOF
                    frm.cmbEquipment.AddItem rsTemp![cci_equip_id]
                    rsTemp.MoveNext
                Loop
            End If
        End If
        rsTemp.Close
        Set rsTemp = Nothing
    End If
    If bTrades Then
        strSelect = "select distinct trade_id, CCI_LABOR.trade_skey from LABOR_TRADE inner join CCI_LABOR on CCI_LABOR.trade_skey = labor_trade.trade_skey order by trade_id"
        ' Use DAL to perform select
        blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
        If blnReturn = False Then
            MsgBox "An error occurred loading Trades."
        Else
            If Not (rsTemp.EOF And rsTemp.BOF) Then
                Do Until rsTemp.EOF
                    With frm
                        .cmbTradeID.AddItem rsTemp![Trade_ID]
                        .cmbTradeID.ItemData(.cmbTradeID.NewIndex) = rsTemp![trade_skey]
                    End With
                    rsTemp.MoveNext
                Loop
            End If
        End If
        rsTemp.Close
    End If
End Sub

Public Function FillWildCard(str As String) As String
    If InStr(1, str, "%") = 0 And str <> "~" Then
        FillWildCard = str & "%"
    Else
        FillWildCard = str
    End If
End Function

Public Function GeographicType(frm As Form) As String
    'Retrieve the Geographic type from the form
    With frm
        If .optPriCity.Value = True Then
            GeographicType = "1"
        ElseIf .optNatlAvg.Value = True Then
            GeographicType = "2"
        ElseIf .optCCICities.Value = True Then
            GeographicType = "3"
        ElseIf .optAllCities.Value = True Then
            GeographicType = "4"
        ElseIf .opt66Cities.Value = True Then   'rlh 11/2006
            GeographicType = "5"
        End If
    End With
End Function

Private Sub LoadRegKeys()
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
    Dim sSQL As String
    Dim rec As ADODB.RecordSet
    
    ' Check the Registry for the AlternateRow Value
    lSize = 4
    strKey = CCD_KEY + "\Defaults\AlternateRow"
    lRet = RegOpenKeyEx(HKEY_CURRENT_USER, strKey, 0&, KEY_ALL_ACCESS, hKey)
    ' Test to see if the Alternate Row Value is there
    If lRet <> ERROR_NONE Then
        ' Populate the values
        vValue = 0
        lRet = RegCreateKeyEx(HKEY_CURRENT_USER, strKey, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRet)
        lRet = RegSetValueExLong(hKey, "Value", 0&, REG_DWORD, vValue, lSize)
    Else
        lRet = RegQueryValueExLong(hKey, "Value", 0&, REG_DWORD, vValue, lSize)
    End If
    g_blnAlternateRow = IIf(vValue = 1, True, False)
    RegCloseKey (hKey)
    
    ' Check the Registry for the Maximized Grid Value
    lSize = 4
    strKey = CCD_KEY + "\Defaults\MaximizeGridForms"
    lRet = RegOpenKeyEx(HKEY_CURRENT_USER, strKey, 0&, KEY_ALL_ACCESS, hKey)
    ' Test to see if the Maximized Grid  Value is there
    If lRet <> ERROR_NONE Then
        ' Populate the values
        vValue = 1
        lRet = RegCreateKeyEx(HKEY_CURRENT_USER, strKey, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRet)
        lRet = RegSetValueExLong(hKey, "Value", 0&, REG_DWORD, vValue, lSize)
    Else
        lRet = RegQueryValueExLong(hKey, "Value", 0&, REG_DWORD, vValue, lSize)
    End If
    g_blnMaximize = IIf(vValue = 1, True, False)
    RegCloseKey (hKey)
    
    ' Check the Registry for the Flat Toolbar Value
    lSize = 4
    strKey = CCD_KEY + "\Defaults\FlatToolbar"
    lRet = RegOpenKeyEx(HKEY_CURRENT_USER, strKey, 0&, KEY_ALL_ACCESS, hKey)
    ' Test to see if the Flat Toolbar Value is there
    If lRet <> ERROR_NONE Then
        ' Populate the values
        vValue = 1
        lRet = RegCreateKeyEx(HKEY_CURRENT_USER, strKey, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRet)
        lRet = RegSetValueExLong(hKey, "Value", 0&, REG_DWORD, vValue, lSize)
    Else
        lRet = RegQueryValueExLong(hKey, "Value", 0&, REG_DWORD, vValue, lSize)
    End If
    g_blnFlatToolbar = IIf(vValue = 1, True, False)
    RegCloseKey (hKey)

    ' Check the Registry for the Alternate Disabled Color Value
    lSize = 4
    strKey = CCD_KEY + "\Defaults\UseAlternateDisabledColor"
    lRet = RegOpenKeyEx(HKEY_CURRENT_USER, strKey, 0&, KEY_ALL_ACCESS, hKey)
    ' Test to see if the Flat Toolbar Value is there
    If lRet <> ERROR_NONE Then
        ' Populate the values
        vValue = 0
        lRet = RegCreateKeyEx(HKEY_CURRENT_USER, strKey, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRet)
        lRet = RegSetValueExLong(hKey, "Value", 0&, REG_DWORD, vValue, lSize)
    Else
        lRet = RegQueryValueExLong(hKey, "Value", 0&, REG_DWORD, vValue, lSize)
    End If
    g_blnUseAlternateDisabledColor = IIf(vValue = 1, True, False)
    RegCloseKey (hKey)

    ' Check the Registry for the Default MasterFormat Value
    lSize = 4
    strKey = CCD_KEY + "\Defaults\MasterFormat"
    lRet = RegOpenKeyEx(HKEY_CURRENT_USER, strKey, 0&, KEY_ALL_ACCESS, hKey)
    ' Test to see if the MasterFormat Value is there
    If lRet <> ERROR_NONE Then
        ' Populate the values
        vValue = UCD_MASTERFORMAT_VERSION
        lRet = RegCreateKeyEx(HKEY_CURRENT_USER, strKey, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRet)
        lRet = RegSetValueExLong(hKey, "Value", 0&, REG_DWORD, vValue, lSize)
    Else
        lRet = RegQueryValueExLong(hKey, "Value", 0&, REG_DWORD, vValue, lSize)
    End If
    g_intMasterFormat = vValue
    RegCloseKey (hKey)

    ' Check the Registry for the MaxRecords Value
    lSize = 4
    strKey = CCD_KEY + "\Defaults\MaxRecords"
    lRet = RegOpenKeyEx(HKEY_CURRENT_USER, strKey, 0&, KEY_ALL_ACCESS, hKey)
    ' Test to see if the MaxRecords value is there
    If lRet <> ERROR_NONE Then
        ' Populate the values
        lValue = "1000"
        lRet = RegCreateKeyEx(HKEY_CURRENT_USER, strKey, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRet)
        lRet = RegSetValueExLong(hKey, "Value", 0&, REG_DWORD, lValue, lSize)
    Else
        lRet = RegQueryValueExLong(hKey, "Value", 0&, REG_DWORD, lValue, lSize)
    End If
    MAX_RECORDS = lValue
    RegCloseKey (hKey)

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
        ' 10/25/2005 RTD - CORRECTED THE SERVER VARIABLE
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
    
    CONN_USERPW = ""
    ' Get the current username
    sBuffer = Space$(255)
    lSize = Len(sBuffer)
    Call GetUserName(sBuffer, lSize)
    strUserName = Left(sBuffer, lSize - 1)
    strUserName = strUserName
    
End Sub

Private Sub LoadRegKeysQtr()

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
    Dim sSQL As String
    Dim rec As ADODB.RecordSet
    

' Check the Registry for the Current Quarter ID Value
    lSize = 1000
    
    strKey = CCD_KEY + "\Defaults\CurrentQuarter"

    lRet = RegOpenKeyEx(HKEY_CURRENT_USER, strKey, 0&, KEY_ALL_ACCESS, hKey)

    ' Test to see if the MaxRecords value is there
    If lRet <> ERROR_NONE Then
        ' Populate the values
        sSQL = "    SELECT qd.quarter_id FROM QUARTER_DATE QD WHERE GETDATE() BETWEEN qd.start_date AND qd.term_date"
        g_objDAL.GetRecordset CONNECT, sSQL, rec
        If rec.EOF And rec.BOF Then
            vValue = ""
        Else
            If rec.RecordCount = 0 Then     'invalid
                vValue = ""
            Else
                vValue = rec.Fields("quarter_id")
            End If
        End If
        rec.Close
        Set rec = Nothing
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

    g_sQuarterID = vValue
    
    RegCloseKey (hKey)

End Sub

Public Sub HideGridSort()
    
    fMainForm.tbToolBar.Buttons.Item(tbrSORT_ASCENDING).Enabled = False
    fMainForm.tbToolBar.Buttons.Item(tbrSORT_ASCENDING).Visible = False
    fMainForm.tbToolBar.Buttons.Item(tbrSORT_DESCENDING).Enabled = False
    fMainForm.tbToolBar.Buttons.Item(tbrSORT_DESCENDING).Visible = False
    ' ALSO HIDE THE SEPARATOR AFTER THE DESCENDING BUTTON
    ' ADDED 5/31/2005 RTD
    fMainForm.tbToolBar.Buttons.Item(tbrSORT_DESCENDING + 1).Visible = False
    
End Sub

Public Sub ShowGridSort()
    
    fMainForm.tbToolBar.Buttons.Item(tbrSORT_ASCENDING).Visible = True
    fMainForm.tbToolBar.Buttons.Item(tbrSORT_DESCENDING).Visible = True
    fMainForm.tbToolBar.Buttons.Item(tbrSORT_ASCENDING).Enabled = False
    fMainForm.tbToolBar.Buttons.Item(tbrSORT_DESCENDING).Enabled = False
    ' ALSO SHOW THE SEPARATOR AFTER THE DESCENDING BUTTON
    ' ADDED 5/31/2005 RTD
    fMainForm.tbToolBar.Buttons.Item(tbrSORT_DESCENDING + 1).Visible = True

End Sub

Public Function GetUCSkey(strUnit_Cost_ID As String, Optional intMasterFormatVersion As Long = UCD_MASTERFORMAT_VERSION) As Long
'MODIFIED 8/5/2005 RTD - ADDED SUPPORT FOR MASTERFORMAT VERSION
    Dim rec As New ADODB.RecordSet
    Dim strQuery As String
    
    Select Case intMasterFormatVersion
    Case EXT_MASTERFORMAT_VERSION
        strQuery = "select unit_cost_skey from unit_cost_detail_ext where unit_cost_id = '" + strUnit_Cost_ID + "'"
    Case UCD_MASTERFORMAT_VERSION
        strQuery = "select unit_cost_skey from unit_cost_detail where unit_cost_id = '" + strUnit_Cost_ID + "'"
    Case ALT_MASTERFORMAT_VERSION
        strQuery = "select unit_cost_skey from unit_cost_detail where alt_unit_cost_id = '" + strUnit_Cost_ID + "'"
    Case Else
        strQuery = "select unit_cost_skey from unit_cost_detail where unit_cost_id = '" + strUnit_Cost_ID + "'"
    End Select
    
    g_objDAL.GetRecordset vbNullString, strQuery, rec
    If rec.EOF Then
        GetUCSkey = 0
    Else
        GetUCSkey = rec.Fields("unit_cost_skey")
    End If
    rec.Close
    Set rec = Nothing

End Function

Public Function GetAssemblySkey(strAssembly_ID As String) As Long
    Dim rec As New ADODB.RecordSet
    
    g_objDAL.GetRecordset vbNullString, "select assembly_skey from assembly_detail where assembly_id = '" + strAssembly_ID + "'", rec
    If rec.EOF Then
        GetAssemblySkey = 0
    Else
        GetAssemblySkey = rec.Fields("assembly_skey")
    End If
    rec.Close
    Set rec = Nothing

End Function

Public Function GetMatSkey(strMat_ID As String) As Long
    Dim rec As New ADODB.RecordSet
    
    g_objDAL.GetRecordset vbNullString, "select mat_skey from material where mat_id = '" + strMat_ID + "'", rec
    If rec.EOF Then
        GetMatSkey = 0
    Else
        GetMatSkey = rec.Fields("mat_skey")
    End If
    rec.Close
    Set rec = Nothing
    
End Function

Public Function Invalid_mat_id_Format(ID As String, fldname As String, m_rec As ADODB.RecordSet) As Boolean
    Dim blnError As Boolean
    Dim strErrorDesc As String
    Dim strSelect As String
    Dim blnReturn As Boolean
    Dim rec As New ADODB.RecordSet

    'Validate the material ID
    If UCase(Left(ID, 1)) <> "M" Then
        If Not IsNumeric(ID) Then
            blnError = True
        End If
    Else
        If Not IsNumeric(Right(ID, Len(ID) - 1)) Then
            blnError = True
        Else
            If Not ((Len(ID) = 13) Or (Len(ID) = 11)) Then    'Check for duplicate
                blnError = True
            End If
        End If
    End If
    If blnError = True Then
        strErrorDesc = "The material id " + ID + " is not valid. Please enter a valid Material - (M + 10 or 12 numbers)"
    Else
        'if the ID entered has changed (not original) see if it is in use
        strSelect = "Select " + fldname + ", mat_skey from Material where mat_id='" + ID + "' or  alt_mat_id='" + ID + "'"
        ' Use DAL to perform select
        blnReturn = g_objDAL.GetRecordset(vbNullString, strSelect, rec)
        If rec.RecordCount > 0 Then
            If IsNull(m_rec.Fields(fldname)) Or m_rec.Fields(fldname).OriginalValue <> ID Then
                strErrorDesc = "The material already exists and may not be added."
                blnError = True
            End If
        End If
        rec.Close
        Set rec = Nothing
    End If

If blnError = True Then
    Beep
    MsgBox strErrorDesc
    Invalid_mat_id_Format = True
Else
    ID = UCase(ID)
End If

End Function

Public Function Invalid_ID_Format(ID As String, _
                                                    fldname As String, _
                                                    m_rec As ADODB.RecordSet, _
                                                    blnNew As Boolean, _
                                    Optional sTableName As String, _
                                    Optional UseAlt As Boolean = True _
                                                    ) As Boolean
    Dim blnError As Boolean
    Dim strErrorDesc As String
    Dim strSelect As String
    Dim blnReturn As Boolean
    Dim rec As New ADODB.RecordSet
    
    'If IsEmpty(sTableName) Then  3-21-01 EP CR#910--IsEmpty only accepts a variant datatype
    If sTableName = "" Then
        sTableName = "unit_cost_detail"
    End If
    'Validate the  ID
    If Not IsNumeric(ID) Then
        blnError = True
    Else
        If Not ((Len(ID) = 12) Or (Len(ID) = 10)) Then    'Check for valid length
            blnError = True
        End If
    End If
    If blnError = True Then
        strErrorDesc = "Please enter a valid ID - (10 or 12 numbers)"
    Else
        'if the ID entered has changed (not original) see if it is in use
        If UseAlt = True Then
            strSelect = "Select " + fldname + " from " + sTableName + " where " + fldname + "='" + ID + "' or " + fldname + "='" + ID + "'"
        Else
            strSelect = "Select " + fldname + " from " + sTableName + " where " + fldname + "='" + ID + "'"
        End If
        ' Use DAL to perform select
        blnReturn = g_objDAL.GetRecordset(vbNullString, strSelect, rec)
        If rec.RecordCount > 0 Then
            If blnNew = True Then
                'rlh 04/17/2008  Added the failing ID to the error message for clarification!!!
                strErrorDesc = "The UNIT COST ID, " & ID & ", " & " already exists and may not be added."
                blnError = True
                
                'rlh 04/17/2008 searching for the next (MAX +1) "68" unit_cost_id
                blnReturn = g_objDAL.GetRecordset(vbNullString, "SELECT MAX(unit_cost_id) as UnitCostId_68 FROM UNIT_COST_DETAIL WHERE unit_cost_id LIKE '68%'", rec)
                If rec.RecordCount > 0 Then
                    Dim newMax68UnitCostId As Double
                    newMax68UnitCostId = CDbl(rec("UnitCostId_68"))
                    newMax68UnitCostId = newMax68UnitCostId + 1
                End If
            End If
            
        End If
        rec.Close
        Set rec = Nothing
    End If
    
    If blnError = True Then
        'Beep
        'MsgBox strErrorDesc, vbCritical + vbOKOnly
        Dim ans As Variant
        ans = MsgBox(strErrorDesc & vbCrLf, vbOKOnly, "Invalid Unit Cost Id")
'        Select Case ans
'        Case vbYes
'            frmUnitCost.ext_unit_cost_id.Text = CStr(newMax68UnitCostId)
'
'        End Select
        Invalid_ID_Format = True
       
         
    End If

End Function

Public Function Invalid_Assembly_id_Format(ID As String, fldname As String, m_rec As ADODB.RecordSet, blnNew As Boolean, lSkey As Long) As Boolean
    Dim blnError As Boolean
    Dim strErrorDesc As String
    Dim strSelect As String
    Dim blnReturn As Boolean
    Dim rec As New ADODB.RecordSet

    'Validate the Assembly ID
    If Not ((Len(ID) = 12) Or (Len(ID) = 10)) Then    'Check for valid length
        blnError = True
    End If
    If blnError = True Then
        strErrorDesc = "Please enter a valid Assembly ID - (10 or 12 numbers)"
    Else
        'if the ID entered has changed (not original) see if it is in use
        strSelect = "Select " + fldname + ", assembly_skey from assembly_detail where " + fldname + "='" + ID + "'"
        ' Use DAL to perform select
        blnReturn = g_objDAL.GetRecordset(vbNullString, strSelect, rec)
        If rec.RecordCount > 0 Then
            If lSkey <> rec.Fields("assembly_skey") Then
                strErrorDesc = "The Assembly ID already exists and may not be added."
                blnError = True
            End If
        End If
        rec.Close
        Set rec = Nothing
    End If

    If blnError = True Then
        Beep
        MsgBox strErrorDesc
        Invalid_Assembly_id_Format = True
    End If
    
End Function

Public Function ConvertGeneralString(strText As Variant) As String
' Convert the string to a general number format if it contains a valid number

    If IsNull(strText) = True Then
        ConvertGeneralString = "0"
    ElseIf IsNumeric(Trim(strText)) = True Then
        ConvertGeneralString = CStr(Format(strText, "General Number"))
    Else
        ConvertGeneralString = "0"
    End If

End Function

Private Function FindParm(strParm As String)
    Dim iStart As Integer
    Dim iLen As Integer
    Dim iEnd As Integer
    
'    iStart = InStr(1, UCase(Command), UCase(strParm), vbTextCompare) + 12
'    If iStart = 12 Then
'        FindParm = Empty
'    Else
'        iEnd = InStr(iStart, Command, " ", vbTextCompare)
'        If iEnd = 0 Then
'            iEnd = Len(Command)
'        End If
'        iLen = iEnd - iStart + 1
'        FindParm = Mid(Command, iStart, iLen)
'    End If


'this routine works for both server and database cje

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

Public Function LockField(frm As Form, fld As String) As Boolean
'MODIFIED 7/25/2005 RTD - TO SUPPORT ALTERNATE ROW COLOR
'FOR INCREASED CONTRAST ON CERTAIN COLOR SCHEMES/THEMES
    frm.Controls(fld).Enabled = False
    frm.Controls(fld).Locked = True
    If Not g_blnUseAlternateDisabledColor Then
        frm.Controls(fld).BackColor = vbButtonFace
    Else
        frm.Controls(fld).BackColor = g_intAlternateRowColor
    End If
    
End Function

Public Function ColorLockedFields(ByRef frm As Form) As Boolean
'ADDED 7/25/2005 RTD
'CHANGE THE BACKGROUND COLOR OF ALL OF THE FORM'S LOCKED TEXT BOXES
'TO WINDOWS STANDARD COLOR OR TO USER-SELECTED ALTERNATE COLOR
    Dim ctl As Control
    Dim iLockColor As Long
    
    If Not g_blnUseAlternateDisabledColor Then
        iLockColor = vbButtonFace
    Else
        iLockColor = g_intAlternateRowColor
    End If
    For Each ctl In frm.Controls
        If TypeOf ctl Is TextBox Then
            If ctl.Locked Then
                ctl.BackColor = iLockColor
            End If
        End If
    Next

End Function

Public Sub OutputView(blnUseOutput As Boolean)
    Dim frm As Form
    Dim blnVisible As Boolean

    'If the output form is open and visible, hide it if output does not apply to the form
    'being activated (blnUseOutput = false)
    'If it is open and not visible, show it if the form uses output
    
    If FormOpen("dlgOutput", frm, blnVisible) = True Then
        If blnVisible = True Then
            If blnUseOutput = False Then
                frm.Visible = False
            End If
        Else
            If blnUseOutput = True Then
                frm.Visible = True
            End If
        End If
    
    End If
    On Error Resume Next
    If FormOpen("frmMain", frm, blnVisible) = True Then 'Must be open, retrieve form
    End If
    If blnUseOutput = False Then
         frm.mnuToolsOutput.Enabled = False
    Else
        frm.mnuToolsOutput.Enabled = True
    End If

End Sub
' This routine is supposed to copy over the new value to the original value in the record set.
' But when you actually try it, you get an "Error 424 Object Required" message.
' But we don't ever see it because so much of the code uses "On Error Resume Next", and it hits
' the error and jumps back to the line after this function was called and continues.
' So effectively, this routine is a NOP, and we never actually copy the new value to the
' original value.
Public Sub Reset_Orig_Values(m_rec As ADODB.RecordSet)
'Reset record original field values to current values
    Dim fld As ADODB.Field

    For Each fld In m_rec.Fields
        If Not fld.OriginalValue = fld.Value Or (IsNull(fld.OriginalValue) Xor IsNull(fld.Value)) Then
            m_rec.Fields(fld.Name).OriginalValue = fld.Value
        End If
    Next

End Sub

Public Sub ResizeForm(Form As Form)
    
    On Error Resume Next
    If Form.WindowState = vbMinimized Then
        ShowMinimizedForms
    End If

End Sub

Public Function UnLockField(frm As Form, fld As String) As Boolean

    frm.Controls(fld).Enabled = True
    frm.Controls(fld).Locked = False
    frm.Controls(fld).BackColor = &H80000005  ' was vbwhite, is window background
    frm.Controls(fld).ForeColor = &H80000008  ' was vbBlack, is window text

End Function

Sub Main()
    Dim blnServerChanged As Boolean
    Dim intResult As Integer
    
    On Error GoTo Error_Processing
    InitCommonControlsVB
    Screen.MousePointer = vbHourglass
    frmSplash.Show
    frmSplash.Refresh
    
    frmSplash.SetStatus "Initializing..."
    
    Call LoadRegKeys
        
Continue_Login:
    'rlh 05/22/2007  USE FOR PRODUCTION/DISABLED FOR TEST/DEBUG
    'strConnect = "UID=" + strUserName + ";PWD=;DATABASE=" + strConnectDatabase + ";SERVER=" + strConnectServer + ";DRIVER={SQL SERVER};DSN='';"
    
    'rlh 05/22/2007  There's a problem trying to login using windows authentication uid/password
    '###################################################################
    '##
    '##  you MUST user ccduser to "get in the door".  Any changes made
    '##  to the data will be noted with the windows authentication
    '##  userid (eg. hancockrl)
    '##  NOTE: the CCDLAUNCH application also uses CCDUSER !!!
    '##
    '###################################################################
    strConnect = "UID=ccdUser;PWD=rsmeans;DATABASE=" + strConnectDatabase + ";SERVER=" + strConnectServer + ";DRIVER={SQL SERVER};DSN='';"
    
    frmSplash.SetStatus "Connecting to database(1)..."

On Error GoTo Connect_Error_Processing
    'g_cnShared.ConnectionTimeout = 5000        'rlh  05/22/2007
    'g_cnShared.CommandTimeout = 5000           'rlh
    'g_cnShared.Open CONNECT                    'rlh
    'g_cnSharedLong.ConnectionTimeout = 15000   'rlh
    'g_cnSharedLong.CommandTimeout = 15000      'rlh
    
     g_cnShared.ConnectionTimeout = 0           'rlh  05/22/2007 unlimited timeout (Mel Mossman)
    g_cnShared.CommandTimeout = 0
    g_cnShared.Open CONNECT
    g_cnSharedLong.ConnectionTimeout = 0
    g_cnSharedLong.CommandTimeout = 0
    g_cnSharedLong.Open CONNECT
    
    Call LoadRegKeysQtr
    If blnServerChanged = True Then
        intResult = MsgBox("Do you want to retain these settings?", vbYesNo, "Save Server Settings")
        If intResult = vbYes Then       'Save registry settings
            Dim strKey As String
            Dim hKey As Long
            Dim lSize As Long
            Dim lRet As Long
            lSize = 1000
                'Save Server settings
            strKey = CCD_KEY + "\Defaults\DBServer"
            lSize = Len(strConnectServer)
            lRet = RegOpenKeyEx(HKEY_CURRENT_USER, strKey, 0&, KEY_ALL_ACCESS, hKey)
            lRet = RegSetValueExString(hKey, "Value", 0&, REG_SZ, strConnectServer, lSize)
            RegCloseKey (hKey)
                'Save Database settings
            strKey = CCD_KEY + "\Defaults\DBase"
            lSize = Len(strConnectDatabase)
            lRet = RegOpenKeyEx(HKEY_CURRENT_USER, strKey, 0&, KEY_ALL_ACCESS, hKey)
            lRet = RegSetValueExString(hKey, "Value", 0&, REG_SZ, strConnectDatabase, lSize)
            RegCloseKey (hKey)
        End If
    End If
    
    On Error GoTo Error_Processing

    frmSplash.SetStatus "Connecting to database(2)..."

    ' Cache a global connection
    g_objDAL.CacheConnection (CONNECT)
    
    'validate user in CCD user_name table - cje
    If Login = True Then
        frmSplash.SetStatus "Preparing form..."
        
        Dim intScreenWidth As Integer
        intScreenWidth = GetSystemMetrics(0) ' SM_CXSCREEN
    '    If intScreenWidth < 1024 Then
        intStartLeft = 0
    '    Else
    '        intStartLeft = NAV_TREE_WIDTH
    '    End If
        
        frmSplash.SetStatus "Loading form..."
    
        Set fMainForm = New frmMain
        Load fMainForm
        Unload frmSplash
        Set frmSplash = Nothing
        
        SetColors
        fMainForm.mnuToolsSettingsMaximizeGrids.Checked = g_blnMaximize
        fMainForm.mnuToolsSettingsWhiteBackground.Checked = g_blnAlternateRow
        fMainForm.mnuToolsSettingsFlatToolbar.Checked = g_blnFlatToolbar
        fMainForm.mnuToolsSettingsAltDisabledColor.Checked = g_blnUseAlternateDisabledColor
        If g_blnFlatToolbar Then
            fMainForm.tbToolBar.Style = tbrFlat
        Else
            fMainForm.tbToolBar.Style = tbrStandard
        End If
        fMainForm.mnuToolsSettingsMF1995.Checked = Not (g_intMasterFormat = 2004)
        fMainForm.mnuToolsSettingsMF2004.Checked = (g_intMasterFormat = 2004)
        fMainForm.Show
        
        Screen.MousePointer = vbNormal
    Else
        'end application
        Unload frmSplash
        Set frmSplash = Nothing
        End
    End If
    
    
    
Exit_Sub:
    Screen.MousePointer = vbNormal
    Exit Sub

Connect_Error_Processing:
    strConnectServer = InputBox("Unable to connect to the default server/database.  Enter a new server name to retry:", "Server Name Verification", strConnectServer)
    
    If Len(strConnectServer) = 0 Then
        End
    Else
        blnServerChanged = True
        strConnectDatabase = InputBox("Database name:", "Database Name Verification", strConnectDatabase)
        If Len(strConnectDatabase) = 0 Then
            End
        Else
            Resume Continue_Login
        End If
    End If
    
Error_Processing:
    MsgBox Error$
    Resume Exit_Sub
    Resume 0
    
End Sub

Public Sub ShowMinimizedForms()
    Dim i As Integer
    Dim iLastForm As Integer

        For i = 0 To Forms.Count - 1
            If Forms(i).WindowState = vbMinimized Then
                Forms(i).ZOrder
                iLastForm = i
            End If
        Next
        If iLastForm > 0 Then
            Forms(iLastForm).SetFocus
        End If

End Sub

Public Function START_LEFT() As Integer
    START_LEFT = intStartLeft
End Function

Public Function START_WIDTH() As Integer
    If fMainForm.ScaleWidth - START_LEFT > 11250 Then
        START_WIDTH = fMainForm.ScaleWidth - START_LEFT
    Else
        START_WIDTH = 11250
    End If
End Function

Public Function START_HEIGHT() As Integer
    If fMainForm.ScaleHeight > 7260 Then
        START_HEIGHT = fMainForm.ScaleHeight
    Else
        START_HEIGHT = 7260
    End If
End Function

Public Function CONNECT() As String
    CONNECT = strConnect
End Function

Public Sub SaveUnitCostID(colUnitCostID As Collection, ID As String)
'Add the specified ID to the collection if it does not exist
    Dim blnfound As Boolean
    Dim varUnitCostID As Variant

    'Add the skey to the collection if it does not exist in it.
    blnfound = False
    'Add to the collection of unit_cost_skey entries
    For Each varUnitCostID In colUnitCostID
        If varUnitCostID = ID Then
            blnfound = True
            Exit For
        End If
    Next varUnitCostID
    If blnfound = False Then
        colUnitCostID.Add ID
    End If

End Sub

Public Sub SaveAssemblyID(colAssemblyID As Collection, ID As String)
'Add the specified ID to the collection if it does not exist
    Dim blnfound As Boolean
    Dim varAssemblyID As Variant

    'Add the skey to the collection if it does not exist in it.
    blnfound = False
    'Add to the collection of unit_cost_skey entries
    For Each varAssemblyID In colAssemblyID
        If varAssemblyID = ID Then
            blnfound = True
            Exit For
        End If
    Next varAssemblyID
    If blnfound = False Then
        colAssemblyID.Add ID
    End If

End Sub

' This routine is not currently called because the value domain_tbl.domain_name for pub_map_option has to be
' 1 to allow the unit cost roll up to occur when updating a material price.
' But they really don't ever do this because they don't want
' updates during book season to roll up and change book data.  But if they ever do have a reason to set the
' field to 1, this will be called.
' Note that this is also attempted to be called when a material price is deleted,
' but it never is because of a bug in the code that never actually calls this routine.  This is because in
' CMatPriceMap, in the unbound delete routine, it tries to call SaveMatID by passing in a string to a var
' that expects a long.  Since on error resume next is in effect, it silently fails, never adding the
' item to the collection, which means the roll up never occurs.
' Also, it appears that this code is called from the Assembly usage map, but it never really is because the
' new/update buttons are always disabled.  So effectively, this code will never be called unless they ever
' decide to set the pub_map_option to 1 in the db.
Public Sub UpdateUnitCost(colUnitCostID As Collection)
    Dim recTemp As New ADODB.RecordSet
    Dim rec As New ADODB.RecordSet
    Dim varUnitCostID
    Dim strSelect As String
    Dim strUpdate As String
    Dim blnReturn As Boolean
    Dim strError As String
    Dim blnRet As Boolean
    Dim colAssemblyID As New Collection
    Dim iMasterFormat As Long

    On Error Resume Next
    'If g_intRollupOption = ALWAYS_ROLLUP_MATERIAL Then
    
    For Each varUnitCostID In colUnitCostID
        '8/25/2005 RTD - CHECK FOR EMPTY UNIT COST IDs - CAUSES SP TO RETURN *ALL* RECORDS
        If varUnitCostID = "" Then
            Debug.Print "MainModule.UpdateUnitCost() error: Empty UnitCostID was added into collection!"
            varUnitCostID = "000000000000"
        End If
        strSelect = "exec usp_select_unit_cost_ext_rlh2 @start_unit_cost_id = '" + CStr(varUnitCostID) + "%'," & _
                    " @end_unit_cost_id = '%', @alt_unit_cost_id = '%', @tech_desc='%'"
        recTemp.Close
        ' Use DAL to perform select
        blnReturn = g_objDAL.GetRecordset(vbNullString, strSelect, recTemp)
        If recTemp.EOF Then
            recTemp.Close
            strSelect = strSelect & ", @master_format = '" & EXT_MASTERFORMAT_VERSION & "'"
            blnReturn = g_objDAL.GetRecordset(vbNullString, strSelect, recTemp)
        End If
        '8/25/2005 RTD - DO NOT ATTEMPT UPDATE IF NO RECORDS RETURNED IN recTemp,
        '                OTHERWISE A FALSE ERROR IS GENERATED
        If blnReturn = True And Not recTemp.EOF Then
            '8/25/2005 RTD - CALL THE CORRECT STORED PROC FOR THE MASTERFORMAT VERSION
            iMasterFormat = recTemp.Fields("master_format")
            If iMasterFormat = UCD_MASTERFORMAT_VERSION Then
                strUpdate = "exec sp_update_unit_cost_driver_res"
            Else
                strUpdate = "exec usp_update_unit_cost_driver_ext_rlh"
            End If
            strUpdate = strUpdate + " @unit_cost_skey=" + CStr(recTemp.Fields("unit_cost_skey")) + ", "
            strUpdate = strUpdate + " @unit_cost_id='" + CStr(varUnitCostID) + "', "
            strUpdate = strUpdate + " @alt_unit_cost_id='" + recTemp.Fields("alt_unit_cost_id") + "', "
            strUpdate = strUpdate + " @type_code='" + Trim(CStr(recTemp.Fields("type_code"))) + "', "
            strUpdate = strUpdate + " @format_code='" + Trim(CStr(recTemp.Fields("format_code"))) + "', "
            strUpdate = strUpdate + " @format_characters='" + Trim(recTemp.Fields("format_characters")) + "', "
            strUpdate = strUpdate + " @indent_code='" + Trim(CStr(recTemp.Fields("indent_code"))) + "', "
            strUpdate = strUpdate + " @book_desc='" + SQLFixString(recTemp.Fields("book_desc")) + "', "
            strUpdate = strUpdate + " @metric_book_desc='" + SQLFixString(recTemp.Fields("metric_book_desc")) + "', "
            strUpdate = strUpdate + " @tech_desc='" + SQLFixString(recTemp.Fields("tech_desc")) + "', "
            strUpdate = strUpdate + " @metric_tech_desc='" + SQLFixString(recTemp.Fields("metric_tech_desc")) + "', "
            strUpdate = strUpdate + " @assembly_book_desc='" + SQLFixString(recTemp.Fields("assembly_book_desc")) + "', "
            strUpdate = strUpdate + " @metric_assembly_book_desc='" + SQLFixString(recTemp.Fields("metric_assembly_book_desc")) + "', "
            strUpdate = strUpdate + " @index_code='" + Trim(CStr(recTemp.Fields("index_code"))) + "', "
            strUpdate = strUpdate + " @index_desc='" + SQLFixString(recTemp.Fields("index_desc")) + "', "
            strUpdate = strUpdate + " @crew_qty='" + ConvertGeneralString(recTemp.Fields("crew_qty")) + "', "
            strUpdate = strUpdate + " @crew_id='" + Trim(recTemp.Fields("crew_id")) + "', "
            strUpdate = strUpdate + " @unit='" + Trim(recTemp.Fields("unit")) + "', "
            strUpdate = strUpdate + " @daily_output='" + ConvertGeneralString(recTemp.Fields("daily_output")) + "', "
            strUpdate = strUpdate + " @std_labor_hour='" + ConvertGeneralString(recTemp.Fields("std_labor_hour")) + "', "
            strUpdate = strUpdate + " @std_mat_cost='" + ConvertGeneralString(recTemp.Fields("std_mat_cost")) + "', "
            strUpdate = strUpdate + " @std_labor_cost='" + ConvertGeneralString(recTemp.Fields("std_labor_cost")) + "', "
            strUpdate = strUpdate + " @std_equip_cost='" + ConvertGeneralString(recTemp.Fields("std_equip_cost")) + "', "
            strUpdate = strUpdate + " @std_total_cost='" + ConvertGeneralString(recTemp.Fields("std_total_cost")) + "', "
            strUpdate = strUpdate + " @std_mat_cost_op='" + ConvertGeneralString(recTemp.Fields("std_mat_cost_op")) + "', "
            strUpdate = strUpdate + " @std_labor_cost_op='" + ConvertGeneralString(recTemp.Fields("std_labor_cost_op")) + "', "
            strUpdate = strUpdate + " @std_equip_cost_op='" + ConvertGeneralString(recTemp.Fields("std_equip_cost_op")) + "', "
            strUpdate = strUpdate + " @std_total_cost_op='" + ConvertGeneralString(recTemp.Fields("std_total_cost_op")) + "', "
            strUpdate = strUpdate + " @opn_labor_hour='" + ConvertGeneralString(recTemp.Fields("opn_labor_hour")) + "', "
            strUpdate = strUpdate + " @opn_mat_cost='" + ConvertGeneralString(recTemp.Fields("opn_mat_cost")) + "', "
            strUpdate = strUpdate + " @opn_labor_cost='" + ConvertGeneralString(recTemp.Fields("opn_labor_cost")) + "', "
            strUpdate = strUpdate + " @opn_equip_cost='" + ConvertGeneralString(recTemp.Fields("opn_equip_cost")) + "', "
            strUpdate = strUpdate + " @opn_total_cost='" + ConvertGeneralString(recTemp.Fields("opn_total_cost")) + "', "
            strUpdate = strUpdate + " @opn_mat_cost_op='" + ConvertGeneralString(recTemp.Fields("opn_mat_cost_op")) + "', "
            strUpdate = strUpdate + " @opn_labor_cost_op='" + ConvertGeneralString(recTemp.Fields("opn_labor_cost_op")) + "', "
            strUpdate = strUpdate + " @opn_equip_cost_op='" + ConvertGeneralString(recTemp.Fields("opn_equip_cost_op")) + "', "
            strUpdate = strUpdate + " @opn_total_cost_op='" + ConvertGeneralString(recTemp.Fields("opn_total_cost_op")) + "', "
            strUpdate = strUpdate + " @rr_labor_hour='" + ConvertGeneralString(recTemp.Fields("rr_labor_hour")) + "', "
            strUpdate = strUpdate + " @rr_mat_cost='" + ConvertGeneralString(recTemp.Fields("rr_mat_cost")) + "', "
            strUpdate = strUpdate + " @rr_labor_cost='" + ConvertGeneralString(recTemp.Fields("rr_labor_cost")) + "', "
            strUpdate = strUpdate + " @rr_equip_cost='" + ConvertGeneralString(recTemp.Fields("rr_equip_cost")) + "', "
            strUpdate = strUpdate + " @rr_total_cost='" + ConvertGeneralString(recTemp.Fields("rr_total_cost")) + "', "
            strUpdate = strUpdate + " @rr_mat_cost_op='" + ConvertGeneralString(recTemp.Fields("rr_mat_cost_op")) + "', "
            strUpdate = strUpdate + " @rr_labor_cost_op='" + ConvertGeneralString(recTemp.Fields("rr_labor_cost_op")) + "', "
            strUpdate = strUpdate + " @rr_equip_cost_op='" + ConvertGeneralString(recTemp.Fields("rr_equip_cost_op")) + "', "
            strUpdate = strUpdate + " @rr_total_cost_op='" + ConvertGeneralString(recTemp.Fields("rr_total_cost_op")) + "', "
            strUpdate = strUpdate + " @metric_unit='" + Trim(recTemp.Fields("metric_unit")) + "', "
            strUpdate = strUpdate + " @metric_daily_output='" + ConvertGeneralString(recTemp.Fields("metric_daily_output")) + "', "
            strUpdate = strUpdate + " @metric_labor_hour='" + ConvertGeneralString(recTemp.Fields("metric_labor_hour")) + "', "
            strUpdate = strUpdate + " @metric_mat_cost='" + ConvertGeneralString(recTemp.Fields("metric_mat_cost")) + "', "
            strUpdate = strUpdate + " @metric_labor_cost='" + ConvertGeneralString(recTemp.Fields("metric_labor_cost")) + "', "
            strUpdate = strUpdate + " @metric_equip_cost='" + ConvertGeneralString(recTemp.Fields("metric_equip_cost")) + "', "
            strUpdate = strUpdate + " @metric_total_cost='" + ConvertGeneralString(recTemp.Fields("metric_total_cost")) + "', "
            strUpdate = strUpdate + " @metric_mat_cost_op='" + ConvertGeneralString(recTemp.Fields("metric_mat_cost_op")) + "', "
            strUpdate = strUpdate + " @metric_labor_cost_op='" + ConvertGeneralString(recTemp.Fields("metric_labor_cost_op")) + "', "
            strUpdate = strUpdate + " @metric_equip_cost_op='" + ConvertGeneralString(recTemp.Fields("metric_equip_cost_op")) + "', "
            strUpdate = strUpdate + " @metric_total_cost_op='" + ConvertGeneralString(recTemp.Fields("metric_total_cost_op")) + "', "
            strUpdate = strUpdate + " @percent_flag='" + recTemp.Fields("percent_flag") + "', "
            strUpdate = strUpdate + " @comment='" + recTemp.Fields("comment") + "', "
            strUpdate = strUpdate + " @last_update_person='" + recTemp.Fields("last_update_person") + "', "
            strUpdate = strUpdate + " @ucd_last_update_id=" + CStr(recTemp.Fields("ucd_last_update_id")) + ", "
            strUpdate = strUpdate + " @cstw_last_update_id=" + CStr(recTemp.Fields("cstw_last_update_id")) + ", "
            strUpdate = strUpdate + " @update_material_usage_ind=1" + ", "
            strUpdate = strUpdate + " @cost_change_ind=1,"
            strUpdate = strUpdate + " @bypass_ucd_ind=0, "
            
            ' Added resi support 3/2011.  These are used by both sp's.
            strUpdate = strUpdate + " @res_labor_hour='" + ConvertGeneralString(recTemp.Fields("res_labor_hour")) + "', "
            strUpdate = strUpdate + " @res_mat_cost='" + ConvertGeneralString(recTemp.Fields("res_mat_cost")) + "', "
            strUpdate = strUpdate + " @res_labor_cost='" + ConvertGeneralString(recTemp.Fields("res_labor_cost")) + "', "
            strUpdate = strUpdate + " @res_equip_cost='" + ConvertGeneralString(recTemp.Fields("res_equip_cost")) + "', "
            strUpdate = strUpdate + " @res_total_cost='" + ConvertGeneralString(recTemp.Fields("res_total_cost")) + "', "
            strUpdate = strUpdate + " @res_mat_cost_op='" + ConvertGeneralString(recTemp.Fields("res_mat_cost_op")) + "', "
            strUpdate = strUpdate + " @res_labor_cost_op='" + ConvertGeneralString(recTemp.Fields("res_labor_cost_op")) + "', "
            strUpdate = strUpdate + " @res_equip_cost_op='" + ConvertGeneralString(recTemp.Fields("res_equip_cost_op")) + "', "
            strUpdate = strUpdate + " @res_total_cost_op='" + ConvertGeneralString(recTemp.Fields("res_total_cost_op")) + "'"
            
            If iMasterFormat = UCD_MASTERFORMAT_VERSION Then
                ' This version doesn't use any more args.
            Else
                ' This version also needs the inhouse args.
                strUpdate = strUpdate + ", "
                strUpdate = strUpdate + " @inhouse_total_cost_op='" + ConvertGeneralString(recTemp.Fields("inhouse_total_cost_op")) + "', "
                strUpdate = strUpdate + " @inhouse_equip_cost_op='" + ConvertGeneralString(recTemp.Fields("inhouse_equip_cost_op")) + "', "
                strUpdate = strUpdate + " @inhouse_mat_cost_op='" + ConvertGeneralString(recTemp.Fields("inhouse_mat_cost_op")) + "', "
                strUpdate = strUpdate + " @inhouse_labor_cost_op='" + ConvertGeneralString(recTemp.Fields("inhouse_labor_cost_op")) + "'"
            End If
          
            blnRet = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
            If blnRet = False Then
                MsgBox strError, vbExclamation
            Else    'Find associated assembly IDs from UC Usage
                strSelect = "select distinct assembly_id from unit_cost_usage inner join " + _
                    "assembly_detail on unit_cost_usage.parent_skey = assembly_detail.assembly_skey where " + _
                    "skey_type = 'A' and unit_cost_skey = " + CStr(recTemp.Fields("unit_cost_skey"))
                rec.Close
                blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rec)
                If blnReturn = True Then
                    If rec.RecordCount > 0 Then
                        Do Until rec.EOF
                            SaveAssemblyID colAssemblyID, rec.Fields("assembly_id")
                            rec.MoveNext
                        Loop
                    End If
                End If
            End If
        End If
    
    Next varUnitCostID

    'Rollup assembly cost
    If colAssemblyID.Count > 0 Then
        UpdateAssembly colAssemblyID
    End If
'End If

End Sub

Public Sub UpdateAssembly(colAssemblyID As Collection)

Dim recTemp As New ADODB.RecordSet
Dim varAssemblyID
Dim strSelect As String
Dim strUpdate As String
Dim blnReturn As Boolean
Dim strError As String
Dim blnRet As Boolean

On Error Resume Next

For Each varAssemblyID In colAssemblyID
'Select only commercial
    strSelect = "exec sp_select_assembly @assembly_id = '" + CStr(varAssemblyID) + "', @assembly_type = '0', @tech_desc='%'"

    recTemp.Close
    ' Use DAL to perform select
    blnReturn = g_objDAL.GetRecordset(vbNullString, strSelect, recTemp)
    If blnReturn = True Then
    
        strUpdate = "exec sp_update_assembly_driver"
       
        strUpdate = strUpdate + " @assembly_skey=" + CStr(recTemp.Fields("assembly_skey")) + ", "
        strUpdate = strUpdate + " @assembly_id='" + CStr(varAssemblyID) + "', "
        strUpdate = strUpdate + " @alt_assembly_id='" + recTemp.Fields("alt_assembly_id") + "', "
        
        strUpdate = strUpdate + " @type_code='" + Trim(CStr(recTemp.Fields("type_code"))) + "', "
        
        strUpdate = strUpdate + " @book_desc='" + SQLFixString(recTemp.Fields("book_desc")) + "', "
        strUpdate = strUpdate + " @metric_book_desc='" + SQLFixString(recTemp.Fields("metric_book_desc")) + "', "
        strUpdate = strUpdate + " @tech_desc='" + SQLFixString(recTemp.Fields("tech_desc")) + "', "
        strUpdate = strUpdate + " @metric_tech_desc='" + SQLFixString(recTemp.Fields("metric_tech_desc")) + "', "
        
        strUpdate = strUpdate + " @unit='" + Trim(recTemp.Fields("unit")) + "', "
        strUpdate = strUpdate + " @metric_unit='" + Trim(recTemp.Fields("metric_unit")) + "', "
        strUpdate = strUpdate + " @std_labor_hour='" + ConvertGeneralString(recTemp.Fields("std_labor_hour")) + "', "
        strUpdate = strUpdate + " @metric_labor_hour='" + ConvertGeneralString(recTemp.Fields("metric_labor_hour")) + "', "
        strUpdate = strUpdate + " @opn_labor_hour='" + ConvertGeneralString(recTemp.Fields("opn_labor_hour")) + "', "
        strUpdate = strUpdate + " @rr_labor_hour='" + ConvertGeneralString(recTemp.Fields("rr_labor_hour")) + "', "
        
        strUpdate = strUpdate + " @std_mat_cost='" + ConvertGeneralString(recTemp.Fields("std_mat_cost")) + "', "
        strUpdate = strUpdate + " @std_labor_cost='" + ConvertGeneralString(recTemp.Fields("std_labor_cost")) + "', "
         strUpdate = strUpdate + " @std_equip_cost='" + ConvertGeneralString(recTemp.Fields("std_equip_cost")) + "', "
        strUpdate = strUpdate + " @std_inst_cost='" + ConvertGeneralString(recTemp.Fields("std_inst_cost")) + "', "
        strUpdate = strUpdate + " @std_total_cost='" + ConvertGeneralString(recTemp.Fields("std_total_cost")) + "', "
        strUpdate = strUpdate + " @std_mat_cost_op='" + ConvertGeneralString(recTemp.Fields("std_mat_cost_op")) + "', "
        strUpdate = strUpdate + " @std_labor_cost_op='" + ConvertGeneralString(recTemp.Fields("std_labor_cost_op")) + "', "
        strUpdate = strUpdate + " @std_equip_cost_op='" + ConvertGeneralString(recTemp.Fields("std_equip_cost_op")) + "', "
        strUpdate = strUpdate + " @std_inst_cost_op='" + ConvertGeneralString(recTemp.Fields("std_inst_cost_op")) + "', "
        strUpdate = strUpdate + " @std_total_cost_op='" + ConvertGeneralString(recTemp.Fields("std_total_cost_op")) + "', "
        
        strUpdate = strUpdate + " @opn_mat_cost='" + ConvertGeneralString(recTemp.Fields("opn_mat_cost")) + "', "
        strUpdate = strUpdate + " @opn_labor_cost='" + ConvertGeneralString(recTemp.Fields("opn_labor_cost")) + "', "
        strUpdate = strUpdate + " @opn_equip_cost='" + ConvertGeneralString(recTemp.Fields("opn_equip_cost")) + "', "
        strUpdate = strUpdate + " @opn_inst_cost='" + ConvertGeneralString(recTemp.Fields("opn_inst_cost")) + "', "
        strUpdate = strUpdate + " @opn_total_cost='" + ConvertGeneralString(recTemp.Fields("opn_total_cost")) + "', "
        strUpdate = strUpdate + " @opn_mat_cost_op='" + ConvertGeneralString(recTemp.Fields("opn_mat_cost_op")) + "', "
        strUpdate = strUpdate + " @opn_labor_cost_op='" + ConvertGeneralString(recTemp.Fields("opn_labor_cost_op")) + "', "
        strUpdate = strUpdate + " @opn_equip_cost_op='" + ConvertGeneralString(recTemp.Fields("opn_equip_cost_op")) + "', "
        strUpdate = strUpdate + " @opn_inst_cost_op='" + ConvertGeneralString(recTemp.Fields("opn_inst_cost_op")) + "', "
        strUpdate = strUpdate + " @opn_total_cost_op='" + ConvertGeneralString(recTemp.Fields("opn_total_cost_op")) + "', "
        
        strUpdate = strUpdate + " @rr_mat_cost='" + ConvertGeneralString(recTemp.Fields("rr_mat_cost")) + "', "
        strUpdate = strUpdate + " @rr_labor_cost='" + ConvertGeneralString(recTemp.Fields("rr_labor_cost")) + "', "
        strUpdate = strUpdate + " @rr_equip_cost='" + ConvertGeneralString(recTemp.Fields("rr_equip_cost")) + "', "
        strUpdate = strUpdate + " @rr_inst_cost='" + ConvertGeneralString(recTemp.Fields("rr_inst_cost")) + "', "
        strUpdate = strUpdate + " @rr_total_cost='" + ConvertGeneralString(recTemp.Fields("rr_total_cost")) + "', "
        strUpdate = strUpdate + " @rr_mat_cost_op='" + ConvertGeneralString(recTemp.Fields("rr_mat_cost_op")) + "', "
        strUpdate = strUpdate + " @rr_labor_cost_op='" + ConvertGeneralString(recTemp.Fields("rr_labor_cost_op")) + "', "
        strUpdate = strUpdate + " @rr_equip_cost_op='" + ConvertGeneralString(recTemp.Fields("rr_equip_cost_op")) + "', "
        strUpdate = strUpdate + " @rr_inst_cost_op='" + ConvertGeneralString(recTemp.Fields("rr_inst_cost_op")) + "', "
        strUpdate = strUpdate + " @rr_total_cost_op='" + ConvertGeneralString(recTemp.Fields("rr_total_cost_op")) + "', "
        
        strUpdate = strUpdate + " @metric_mat_cost='" + ConvertGeneralString(recTemp.Fields("metric_mat_cost")) + "', "
        strUpdate = strUpdate + " @metric_labor_cost='" + ConvertGeneralString(recTemp.Fields("metric_labor_cost")) + "', "
        strUpdate = strUpdate + " @metric_equip_cost='" + ConvertGeneralString(recTemp.Fields("metric_equip_cost")) + "', "
        strUpdate = strUpdate + " @metric_inst_cost='" + ConvertGeneralString(recTemp.Fields("metric_inst_cost")) + "', "
        strUpdate = strUpdate + " @metric_total_cost='" + ConvertGeneralString(recTemp.Fields("metric_total_cost")) + "', "
        strUpdate = strUpdate + " @metric_mat_cost_op='" + ConvertGeneralString(recTemp.Fields("metric_mat_cost_op")) + "', "
        strUpdate = strUpdate + " @metric_labor_cost_op='" + ConvertGeneralString(recTemp.Fields("metric_labor_cost_op")) + "', "
        strUpdate = strUpdate + " @metric_equip_cost_op='" + ConvertGeneralString(recTemp.Fields("metric_equip_cost_op")) + "', "
        strUpdate = strUpdate + " @metric_inst_cost_op='" + ConvertGeneralString(recTemp.Fields("metric_inst_cost_op")) + "', "
        strUpdate = strUpdate + " @metric_total_cost_op='" + ConvertGeneralString(recTemp.Fields("metric_total_cost_op")) + "', "
        
        strUpdate = strUpdate + " @pct_ind=" + CStr(CInt(recTemp.Fields("pct_ind"))) + ", "
        strUpdate = strUpdate + " @coml_ind='" + ConvertGeneralString(recTemp.Fields("coml_ind")) + "', "
        strUpdate = strUpdate + " @resi_ind='" + ConvertGeneralString(recTemp.Fields("resi_ind")) + "', "
        strUpdate = strUpdate + " @labor_equip_ind='" + ConvertGeneralString(recTemp.Fields("labor_equip_ind")) + "', "
        
        strUpdate = strUpdate + " @comment='" + recTemp.Fields("comment") + "', "
        strUpdate = strUpdate + " @last_update_person='" + recTemp.Fields("last_update_person") + "', "
        strUpdate = strUpdate + " @ad_last_update_id=" + CStr(recTemp.Fields("ad_last_update_id")) + ", "
        strUpdate = strUpdate + " @std_last_update_id=" + CStr(recTemp.Fields("std_last_update_id")) + ", "
        strUpdate = strUpdate + " @opn_last_update_id=" + CStr(recTemp.Fields("opn_last_update_id")) + ", "
        strUpdate = strUpdate + " @rr_last_update_id=" + CStr(recTemp.Fields("rr_last_update_id")) + ", "
        strUpdate = strUpdate + " @update_unitcost_usage_ind=1, "
        strUpdate = strUpdate + " @cost_change_ind=1, "
        strUpdate = strUpdate + " @ad_change_ind= 0, "
        strUpdate = strUpdate + " @std_change_ind= 0, "
        strUpdate = strUpdate + " @opn_change_ind= 0, "
        strUpdate = strUpdate + " @rr_change_ind='" + ConvertGeneralString(recTemp.Fields("rr_change_ind")) + "'"
        
'        strUpdate = strUpdate + " @ucd_last_update_id=" + CStr(recTemp.Fields("ucd_last_update_id")) + ", "
'        strUpdate = strUpdate + " @daily_output='" + ConvertGeneralString(recTemp.Fields("daily_output")) + "', "
        
        
        blnRet = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
        If blnRet = False Then
            MsgBox strError
        End If

    End If

Next varAssemblyID

End Sub

Public Sub Status(strStatus As String)
    fMainForm.sbStatusBar.Panels(1).Text = strStatus
End Sub

' Compares values on form to values in recordset to see if user changed any
Public Function IsControlChanged_old(frm As Form, rec As ADODB.RecordSet) As Boolean
    On Error Resume Next
    Dim ctr As Control
    
    IsControlChanged_old = False
    
    For Each ctr In frm.Controls
        ' Only check if control is not locked
        If ctr.Locked = False Then
            If TypeOf ctr Is TextBox Then
                Select Case rec.Fields(ctr.Name).Type
                Case adChar, adVarChar
                    If Not ctr.Text = rec.Fields(ctr.Name).Value Or (ctr.Text = "" Xor rec.Fields(ctr.Name).Value = "") Then
                        IsControlChanged_old = True
                        Exit For
                    End If
                Case adInteger, adSmallInt, adUnsignedInt, adUnsignedSmallInt, adTinyInt, adUnsignedTinyInt, adDouble
                    If Not ctr.Text = Format(rec.Fields(ctr.Name).Value) Or (ctr.Text = "" Xor rec.Fields(ctr.Name).Value = "") Then
                        IsControlChanged_old = True
                        Exit For
                    End If
                Case adCurrency
                    If Not ctr.Text = Format(rec.Fields(ctr.Name).Value, "#,###,##0.00") Or (ctr.Text = "" Xor rec.Fields(ctr.Name).Value = "") Then
                        IsControlChanged_old = True
                        Exit For
                    End If
                End Select
            ElseIf TypeOf ctr Is ComboBox Then
                If Not ctr.Text = rec.Fields(ctr.Name).Value Or (ctr.Text = "" Xor (IsNull(rec.Fields(ctr.Name).Value) Or rec.Fields(ctr.Name).Value = "")) Then
                    IsControlChanged_old = True
                    Exit For
                End If
            ElseIf TypeOf ctr Is CheckBox Then
                If (ctr.Value = 0 And rec.Fields(ctr.Name).Value = True) Or (ctr.Value = 1 And rec.Fields(ctr.Name).Value = False) Then
                    IsControlChanged_old = True
                    Exit For
                End If
            End If
        End If
    Next
End Function

Public Function IsControlChanged(frm As Form, rec As ADODB.RecordSet) As Boolean
    On Error Resume Next
    Dim recClone As ADODB.RecordSet
    Dim fld As ADODB.Field
    Dim ctr As Control

    Set recClone = rec.Clone
    recClone.AddNew

    UpdateRecordsetFromForm frm, recClone

    For Each fld In rec.Fields
        Set ctr = Nothing
        Set ctr = frm.Controls(fld.Name)
        If Not ctr Is Nothing Then
            If ctr.Locked = False Then
                If TypeOf ctr Is CheckBox Then
                    If Not IsNull(fld.Value) Then   'Check for incompatable types - 1/0 in recordset, true/false on fld
                        If fld.Value And recClone.Fields(fld.Name).Value = 0 Then
                            IsControlChanged = True
                            Exit For
                        End If
                        If Not fld.Value And recClone.Fields(fld.Name).Value = 1 Then
                            IsControlChanged = True
                            Exit For
                        End If
                        If Not fld.Value = recClone.Fields(fld.Name).Value Then
'                        If Not (Fld.Value <> recClone.Fields(Fld.Name).Value = 0) Or (Fld.Value And recClone.Fields(Fld.Name) = 1) Then
                            IsControlChanged = True
                            Exit For
                        End If
                    End If
                Else
                    If Not Trim(UCase(fld.Value)) = Trim(UCase(recClone.Fields(fld.Name).Value)) Or (IsNull(fld.Value) And Not Trim(recClone.Fields(fld.Name).Value) = "") Then
                        If IsNumeric(Trim(UCase(fld.Value))) Then
                            If Val(Trim(UCase(fld.Value))) <> 0 Then
                                IsControlChanged = True
                                Exit For
                            End If
                        Else
                            IsControlChanged = True
                            Exit For
                        End If
                    End If
                End If
            End If
        End If
    Next

    ' Cancel the AddNew above
    recClone.CancelUpdate
End Function

' Builds an update statement from the form specified
' strTagNumber is the first digit of the Tag property for the control
Public Sub BuildUpdateSQL(frm As Form, ByRef strUpdateClause As String, Optional strTagNumber As String = "1")
    Dim ctr As Control
    For Each ctr In frm.Controls
        If Left(ctr.Tag, 1) = strTagNumber Then
            If TypeOf ctr Is TextBox Then
                Select Case Right(ctr.Tag, 1)
                Case "S"
                    strUpdateClause = strUpdateClause + ctr.Name + "="
                    strUpdateClause = strUpdateClause + "'" + ctr.Text + "', "
                Case "N"
                    If IsNumeric(ctr.Text) = True Then
                        strUpdateClause = strUpdateClause + ctr.Name + "="
                        strUpdateClause = strUpdateClause + ctr.Text + ", "
                    End If
                Case "D"
                    If IsDate(ctr.Text) = True Then
                        strUpdateClause = strUpdateClause + ctr.Name + "="
                        strUpdateClause = strUpdateClause + "'" + ctr.Text + "', "
                    End If
                End Select
            ElseIf TypeOf ctr Is CheckBox Then
                strUpdateClause = strUpdateClause + ctr.Name + "="
                strUpdateClause = strUpdateClause + str(ctr.Value) + ", "
            End If
        End If
    Next ctr
    
    strUpdateClause = strUpdateClause + "last_update_date='" + Format(Now, "mm/dd/yyyy") + "', last_update_person='" + strUserName + "'"
End Sub

Public Sub BuildInsertSQL(frm As Form, ByRef strInsertClause As String, Optional strTagNumber As String = "1")
    Dim strCols As String
    Dim strValues As String
    Dim ctr As Control
    For Each ctr In frm.Controls
        If Left(ctr.Tag, 1) = strTagNumber Then
            
            If TypeOf ctr Is TextBox Then
                Select Case Right(ctr.Tag, 1)
                Case "S"
                    strCols = strCols + ctr.Name + ", "
                    strValues = strValues + "'" + ctr.Text + "', "
                Case "N"
                    If IsNumeric(ctr.Text) = True Then
                        strCols = strCols + ctr.Name + ", "
                        strValues = strValues + ctr.Text + ", "
                    End If
                Case "D"
                    If IsDate(ctr.Text) = True Then
                        strCols = strCols + ctr.Name + ", "
                        strValues = strValues + "'" + ctr.Text + "', "
                    End If
                End Select
            ElseIf TypeOf ctr Is CheckBox Then
                strCols = strCols + ctr.Name + ", "
                strValues = strValues + str(ctr.Value) + ", "
            End If
        End If
    Next ctr
    
    strCols = strCols + "last_update_date, last_update_person"
    strValues = strValues + "'" + Format(Now, "mm/dd/yyyy") + "', '" + strUserName + "'"
    strInsertClause = strInsertClause + " (" + strCols + ") values(" + strValues + ")"
End Sub

Public Sub BuildStoredProcSQL(frm As Form, ByRef strStoredProcClause As String, Optional strTagNumber As String = "0", _
                              Optional rec As ADODB.RecordSet = vbEmpty, Optional excludeList As Collection)
    Dim strCols As String
    Dim strValues As String
    Dim ctr As Control
    On Error GoTo Error_Processing
    For Each ctr In frm.Controls

        If strTagNumber = "0" Or Left(ctr.Tag, 1) = strTagNumber Then
            ' Process it unless it is supposed to be excluded.
            If (Not ExistsInCollection(excludeList, ctr.Name)) Then
            
                 If TypeOf ctr Is TextBox Then
                     If Not rec Is Nothing Then
                         Select Case rec.Fields(ctr.Name).Type
                         Case adInteger, adDouble, adUnsignedTinyInt
                             If ctr.Text = "" Then
                                 strStoredProcClause = strStoredProcClause + " @" + ctr.Name + "=0,"
                             Else
                                 strStoredProcClause = strStoredProcClause + " @" + ctr.Name + "=" + SQLFixString(ctr.Text) + ","
                             End If
                         Case adCurrency
                             If ctr.Text = "" Then
                                 strStoredProcClause = strStoredProcClause + " @" + ctr.Name + "=0,"
                             Else
                                 strStoredProcClause = strStoredProcClause + " @" + ctr.Name + "=" + Format(ctr.Text, "General Number") + ","
                             End If
                         Case adCurrency, adDecimal, adVarNumeric, adNumeric, adSmallInt
                             If ctr.Text = "" Then
                                 strStoredProcClause = strStoredProcClause + " @" + ctr.Name + "=0,"
                             Else
                                 strStoredProcClause = strStoredProcClause + " @" + ctr.Name + "=" + Format(ctr.Text, "General Number") + ","
                             End If
                         Case adVarChar, adChar
                             strStoredProcClause = strStoredProcClause + " @" + ctr.Name + "='" + SQLFixString(ctr.Text) + "',"
                         Case adDBTimeStamp
                             strStoredProcClause = strStoredProcClause + " @" + ctr.Name + "='" + Format(ctr.Text, "mm/dd/yyyy") + "',"
                         Case adBoolean
                             If ctr.Text = True Then
                                 strStoredProcClause = strStoredProcClause + " @" + ctr.Name + "=1,"
                             Else
                                 strStoredProcClause = strStoredProcClause + " @" + ctr.Name + "=0,"
                             End If
                         Case Else
                             MsgBox rec.Fields(ctr.Name).Type
                         End Select
                     Else
                         Select Case Right(ctr.Tag, 1)
                         Case "S"
                             strStoredProcClause = strStoredProcClause + " @" + ctr.Name + "='" + SQLFixString(ctr.Text) + "',"
                         Case "G"        'General Number
                             If IsNumeric(Trim(ctr.Text)) = True Then
                                 strStoredProcClause = strStoredProcClause + " @" + ctr.Name + "='" + CStr(SQLFixString(Format(ctr.Text, "General Number"))) + "',"
                             Else
                                 strStoredProcClause = strStoredProcClause + " @" + ctr.Name + "='0',"
                             End If
                         Case "N"
                             If IsNumeric(Trim(ctr.Text)) = True Then
                                 strStoredProcClause = strStoredProcClause + " @" + ctr.Name + "=" + Format(ctr.Text, "######0.#####") + ","
                             Else
                                 strStoredProcClause = strStoredProcClause + " @" + ctr.Name + "=0,"
                             End If
                         Case "D"
                             If IsDate(ctr.Text) = True Then
                                 strStoredProcClause = strStoredProcClause + " @" + ctr.Name + "='" + ctr.Text + "',"
                             End If
                         End Select
                     End If
                 ElseIf TypeOf ctr Is CheckBox Then
                     strStoredProcClause = strStoredProcClause + " @" + ctr.Name + "=" + str(ctr.Value) + ","
                 ElseIf TypeOf ctr Is ComboBox Then
                     strStoredProcClause = strStoredProcClause + " @" + ctr.Name + "='" + ctr.Text + "',"
                 ElseIf TypeOf ctr Is DTPicker Then
                     If IsNull(ctr.Value) Then
                         strStoredProcClause = strStoredProcClause + " @" + ctr.Name + "=Null,"
                     Else
                         strStoredProcClause = strStoredProcClause + " @" + ctr.Name + "='" + Format(ctr.Value, "mm/dd/yyyy") + "',"
                     End If
                 End If
                 
            End If      ' not exists in collection
        End If      ' tag = 0 or left(1) = tag
    Next ctr
    
'    strStoredProcClause = strStoredProcClause + "@last_update_person='" + strUserName + "'"
Exit_Sub:
    Exit Sub

Error_Processing:
    MsgBox Error$
    Resume Next
    'Resume Exit_Sub
    Resume 0
    
End Sub

' Returns true if the key item exists in the collection, else false.
' Also returns false if the collection is nothing.
Public Function ExistsInCollection(colCollection As Collection, vKey As Variant) As Boolean

    ' If the caller didn't pass in an exclude collection, or if it is empty, then return false.
    If (colCollection Is Nothing) Then
        ExistsInCollection = False
        Exit Function
    End If
    
    If (colCollection.Count = 0) Then
        ExistsInCollection = False
        Exit Function
    End If
    
    On Error GoTo ErrHandler
    
    Dim Value As Variant
    
    Value = colCollection.Item(vKey)
    ExistsInCollection = True

    Exit Function
    
ErrHandler:
    ExistsInCollection = False
    
End Function

' Builds an update statement from the RecordSet specified
Public Sub BuildUpdateSQLFromRS(rs As ADODB.RecordSet, ByRef strUpdateClause As String)
    Dim aField
    
    If Not rs.EOF And Not rs.BOF Then
        For Each aField In rs.Fields
            Select Case aField.Type
            Case adChar
                strUpdateClause = strUpdateClause + "@" + aField.Name + "='"
                strUpdateClause = strUpdateClause + aField.Value + "', "
            Case adVarChar
                strUpdateClause = strUpdateClause + "@" + aField.Name + "='"
                strUpdateClause = strUpdateClause + aField.Value + "', "
            Case adInteger
                strUpdateClause = strUpdateClause + "@" + aField.Name + "="
                strUpdateClause = strUpdateClause + str(aField.Value) + ", "
            Case adDouble
                strUpdateClause = strUpdateClause + "@" + aField.Name + "="
                strUpdateClause = strUpdateClause + str(aField.Value) + ", "
            Case adBoolean
                strUpdateClause = strUpdateClause + "@" + aField.Name + "="
                If aField.Value = True Then
                    strUpdateClause = strUpdateClause + str(1) + ", "
                Else
                    strUpdateClause = strUpdateClause + str(0) + ", "
                End If
            Case adDBTimeStamp
                strUpdateClause = strUpdateClause + "@" + aField.Name + "='"
                strUpdateClause = strUpdateClause + Format(aField.Value, "mm/dd/yyyy") + "', "
            End Select
        Next
        strUpdateClause = strUpdateClause + "@last_update_person='" + strUserName + "'"
    End If
End Sub

Public Sub UpdateRecordsetFromForm(frm As Form, ByRef rec As ADODB.RecordSet)
    On Error Resume Next
    Dim ctr As Control
    
    For Each ctr In frm.Controls
        If ctr.Tag = "ignore" Then
            'do nothing
        ElseIf TypeOf ctr Is TextBox Then
            If ctr.Name = "unit_cost_id" Or ctr.Name = "alt_unit_cost_id" Or ctr.Name = "ext_unit_cost_id" Or ctr.Name = "mat_id" Then
                rec.Fields(ctr.Name) = Compress_String(ctr.Text)
            Else
                rec.Fields(ctr.Name) = ctr.Text
            End If
'            Select Case Right(ctr.Tag, 1)
'            Case "S"
'                strStoredProcClause = strStoredProcClause + " @" + ctr.Name + "='" + ctr.Text + "',"
'            Case "N"
'                If IsNumeric(ctr.Text) = True Then
'                    strStoredProcClause = strStoredProcClause + " @" + ctr.Name + "=" + ctr.Text + ","
'                Else
'                    strStoredProcClause = strStoredProcClause + " @" + ctr.Name + "=Null,"
'                End If
'            Case "D"
'                If IsDate(ctr.Text) = True Then
'                    strStoredProcClause = strStoredProcClause + " @" + ctr.Name + "='" + ctr.Text + "',"
'                End If
'            End Select
        ElseIf TypeOf ctr Is CheckBox Then
            rec.Fields(ctr.Name) = str(ctr.Value)
        ElseIf TypeOf ctr Is ComboBox Then
            If Len(Trim(ctr.Text)) = 0 Then
                rec.Fields(ctr.Name) = Null
            Else
                rec.Fields(ctr.Name) = ctr.Text
            End If
        ElseIf TypeOf ctr Is DTPicker Then
            rec.Fields(ctr.Name) = ctr.Value
        End If
    Next ctr

End Sub

Public Sub UpdateFormFromRecordset(frm As Form, ByRef rec As ADODB.RecordSet)
    On Error Resume Next
    Dim ctr As Control
    Dim i As Integer
    ' Loop through all controls on form
    For Each ctr In frm.Controls
        ' Check type of control
        If TypeOf ctr Is TextBox Then
            If rec.Fields(ctr.Name).Type = adDBTimeStamp Then
                ctr = Format(rec.Fields(ctr.Name), "mm/dd/yyyy HH:MM:ss")
            ElseIf rec.Fields(ctr.Name).Type = adDBDate Then
                ctr = Format(rec.Fields(ctr.Name), "mm/dd/yyyy")
            ElseIf rec.Fields(ctr.Name).Type = adCurrency Then
                ctr = Format(rec.Fields(ctr.Name), "#,###,##0.00")
            Else
                If IsNull(rec.Fields(ctr.Name)) Then
                    ctr = ""
                Else
                    ctr = Trim(rec.Fields(ctr.Name))
                End If
            End If
        ElseIf TypeOf ctr Is CheckBox Then
            ' Convert from True/False to 1/0
            If IsNull(rec.Fields(ctr.Name)) Then
                ctr = 0
            ElseIf CBool(rec.Fields(ctr.Name)) Then
                    ctr = 1
                Else
                    ctr = 0
                End If
        ElseIf TypeOf ctr Is ComboBox Then
            For i = 0 To ctr.listcount - 1
                If ctr.List(i) = rec.Fields(ctr.Name) Then
                    ctr.ListIndex = i
                End If
            Next i
            If ctr.ListIndex = -1 Then
                ctr.Text = rec.Fields(ctr.Name)
            End If
        ElseIf TypeOf ctr Is DTPicker Then
            If IsNull(rec.Fields(ctr.Name)) Then
                ctr.Value = Null
            Else
                ctr.Value = rec.Fields(ctr.Name)
            End If
        End If
    Next ctr
End Sub

Public Sub CheckValueForNumber(str As String, Cancel As Boolean)
    If Not IsNumeric(str) Then
        Cancel = True
        MsgBox "This value must be numeric."
    End If
End Sub

Public Function SQLFixString(str As String) As String
    Dim strTemp As String
    strTemp = ReplaceStr(str, "'", "''")
    SQLFixString = ReplaceStr(strTemp, "Chr(34)", "Chr(34)Chr(34)")
End Function

Public Function SQLChangeWildcard(str As String) As String
    SQLChangeWildcard = ReplaceStr(str, "*", "%")
End Function

Function ReplaceStr(TextIn, ByVal SearchStr As String, _
                        ByVal Replacement As String)
    Dim WorkText As String, Pointer As Integer
    Dim CompMode As VbCompareMethod
    
    CompMode = vbBinaryCompare
    If IsNull(TextIn) Then
        ReplaceStr = Null
    Else
        WorkText = TextIn
        Pointer = InStr(1, WorkText, SearchStr, 0)
        Do While Pointer > 0
            WorkText = Left(WorkText, Pointer - 1) & Replacement & _
                       Mid(WorkText, Pointer + Len(SearchStr))
            Pointer = InStr(Pointer + Len(Replacement), WorkText, _
                            SearchStr, CompMode)
        Loop
        ReplaceStr = WorkText
     End If
End Function

Public Sub CopyRSFieldsAndData(ByRef recDest As ADODB.RecordSet, recSource As ADODB.RecordSet, colCols As Collection)
    On Error Resume Next
    Dim intItem As Integer
    Dim fld As ADODB.Field
    
    ' Copy all of the fields
    For Each fld In recSource.Fields
        recDest.Fields.Append fld.Name, fld.Type, fld.definedSize, fld.Attributes
        recDest.Fields.Item(fld.Name).NumericScale = fld.NumericScale
        recDest.Fields.Item(fld.Name).Precision = fld.Precision
    Next
    
    ' Ready a new record to accept the data
    recDest.Open
    recDest.AddNew
    
    ' Copy the data for fields in the collection
    For Each fld In recSource.Fields
        intItem = 0
        intItem = colCols(fld.Name)
        If Not intItem = 0 Then
            recDest.Fields(fld.Name).Value = fld.Value
        End If
    Next
End Sub

Public Sub CopyRSFields(ByRef recDest As ADODB.RecordSet, recSource As ADODB.RecordSet)
    On Error Resume Next
    Dim fld As ADODB.Field
    
    ' Copy all of the fields
    For Each fld In recSource.Fields
        recDest.Fields.Append fld.Name, fld.Type, fld.definedSize, fld.Attributes
        recDest.Fields.Item(fld.Name).NumericScale = fld.NumericScale
        recDest.Fields.Item(fld.Name).Precision = fld.Precision
    Next
    
    ' Ready a new record to accept the data
    recDest.Open
    recDest.AddNew
End Sub

Public Function BuildINFromListbox(lb As ListBox) As String
    Dim i As Integer
    Dim blnOne As Boolean
    Dim str As String
    blnOne = False
    str = "("
    
    For i = 0 To lb.listcount - 1
        If lb.Selected(i) Then
            If blnOne Then
                str = str + ",'"
            Else
                str = str + "'"
            End If
            str = str + lb.List(i)
            str = str + "'"
            blnOne = True
        End If
    Next i
    str = str + ")"
    BuildINFromListbox = str
End Function

Public Function FormOpen(strFormName As String, frmReturn As Form, blnVisible As Boolean) As Boolean
    Dim frm As Object

    ' See if the dialog is open
    For Each frm In Forms
        If frm.Name = strFormName Then
            Set frmReturn = frm
            blnVisible = frm.Visible
            FormOpen = True
            Exit For
        End If
    Next
End Function

Public Function CheckNumericField(sOrigValue As String, curKey As Integer, iSelStart As Integer, iSelLength As Integer, TotDecimals As Integer) As Boolean
 ' this function allows only #, decimal only once, allows how many ever decimal places the user requires
 ' the start point determines if the user is trying to type before the decimal key
 ' added on 8/12/99 siva
    Dim i As Integer
    Dim strValueWOutSel As String   'Contains the original string minus the selected text
    'strValueWOutSel = left(sOrigValue, Len(sOrigValue) - iSelStart) + right(sOrigValue, Len(sOrigValue) - iSelStart + SelLength)
    strValueWOutSel = Left(sOrigValue, iSelStart) + _
    Right(sOrigValue, IIf(Len(sOrigValue) - (iSelLength + iSelStart) > 0, Len(sOrigValue) - (iSelLength + iSelStart), 0))
    
    If TotDecimals = 0 And Chr(curKey) = "." Then
        CheckNumericField = False
        Exit Function
    End If
    If InStr(1, strValueWOutSel, ".") > 0 And Chr(curKey) = "." Then
            CheckNumericField = False
    Else
        If (curKey >= 48 And curKey <= 57) Then
            ' allow the key it is 0 to 9
            CheckNumericField = True
            If TotDecimals > 0 Then
                i = InStr(1, sOrigValue, ".")
                If i > 0 Then
                    If Len(Mid(strValueWOutSel, i + 1)) > TotDecimals - 1 Then
                        'the user is trying to enter more than required decimal places
                        ' so don't allow
                        If iSelStart >= i Then CheckNumericField = False
                    End If
                End If
            End If
        ElseIf Chr(curKey) = "." Then
            '   allow the key
            If TotDecimals = 0 Then
                CheckNumericField = False
            Else
                CheckNumericField = True
            End If
        ElseIf curKey = vbKeyBack Then
            '   allow the key
            CheckNumericField = True
        Else
            ' don't allow the key
            CheckNumericField = False
        End If
    End If
End Function

Public Function Change_Format_To_Numbers(myOldTxt As String, TotDecimals As Integer) As String
' this function changes the format back to regular numbers
' added on 8/12/99 siva
    Dim i As Integer, myNewTxt As String
    myNewTxt = ""
    For i = 1 To Len(RTrim(myOldTxt))
        If CheckNumericField(myNewTxt, Asc(Mid(myOldTxt, i, 1)), 1, 0, TotDecimals) = True Then
            myNewTxt = myNewTxt + Mid(myOldTxt, i, 1)
        End If
    Next i
    Change_Format_To_Numbers = myNewTxt
End Function

Public Sub doh()
    Dim SoundName As String
    Dim Result As Long
    SoundName$ = "c:\doh.wav"
    Result = sndPlaySound(SoundName$, cSndASYNCH Or cSndNODEFAULT)
End Sub

Public Sub final_answer()
    Dim SoundName As String
    Dim Result As Long
    SoundName$ = "c:\final_answer.wav"
    Result = sndPlaySound(SoundName$, cSndASYNCH Or cSndNODEFAULT)
End Sub

Public Sub puzzling()
    Dim SoundName As String
    Dim Result As Long
    SoundName$ = "c:\puzzling.wav"
    Result = sndPlaySound(SoundName$, cSndASYNCH Or cSndNODEFAULT)
End Sub

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
    Dim i As Integer
    
    sTemp = Replace(sText, vbTab, " ")
    For i = 1 To 31
        Do While InStr(sTemp, Chr(i)) > 0
            sTemp = Replace(sTemp, Chr(i), "")
        Loop
    Next
    StripControlCharacters = sTemp

End Function

Public Function GetFileVersion(sDriverFile As String) As String
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
                 Format$(LoWord(FI.dwFileVersionMS), "0") & "."
        If FI.dwFileVersionLS > 0 Then
           tmpVer = tmpVer & Format$(HiWord(FI.dwFileVersionLS), "0") & "." & _
                             Format$(LoWord(FI.dwFileVersionLS), "0")
        Else
           tmpVer = tmpVer & Format$(FI.dwFileVersionLS, "0")
        End If
    End If
    GetFileVersion = tmpVer
   
End Function

Public Function GetSpecialFolderLocation(CSIDL As Long) As String
' RETURN THE PATH TO A SPECIAL WINDOWS SHELL FOLDER (e.g., "My Documents")
    Dim sPath As String
    Dim pidl As Long
    
    If SHGetSpecialFolderLocation(fMainForm.hWnd, CSIDL, pidl) = 0 Then
        sPath = Space$(MAX_PATH)
        If SHGetPathFromIDList(ByVal pidl, ByVal sPath) Then
            sPath = Left(sPath, InStr(sPath, Chr$(0)) - 1)
            If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
            GetSpecialFolderLocation = sPath
        End If
        Call CoTaskMemFree(pidl)
    End If
   
End Function

Public Function LoadMasterFormatCombo(ByRef Combo1 As ComboBox, Optional bNoAltIDSelection As Boolean = False) As Long
' LOAD GIVEN COMBOBOX WITH MASTERFORMAT INFORMATION
' SELECT THE ITEM THAT CORRESPONDS WITH USER'S DEFAULT MASTERFORMAT SETTING
    Dim iIndex As Long
    Dim sDefaultMF As String
    
    Combo1.Clear
    
    ' Get Default MasterFormat and select in Combo
    sDefaultMF = QueryRegistryKey(HKEY_CURRENT_USER, CCD_KEY & "\Defaults\MasterFormat", "Value", CStr(UCD_MASTERFORMAT_VERSION))

    Combo1.AddItem "MF-" & EXT_MASTERFORMAT_VERSION
    Combo1.ItemData(Combo1.NewIndex) = EXT_MASTERFORMAT_VERSION
    If sDefaultMF = EXT_MASTERFORMAT_VERSION Then
        Combo1.ListIndex = Combo1.NewIndex
    End If
    
    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    '
    ' FLIP THIS SWITCH (MF95_ENABLED) AND YOU CAN TOGGLE BACK AND FORTH
    ' FROM SUPPORT/NO SUPPORT FOR MF95 PROCESSING!!!
    '
    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    If SUPER_USER_SUPPORT Then
    Else
        MF95_ENABLED = False     'rlh  02/19/2009
    End If
    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
       
    
    If MF95_ENABLED Then    'rlh  02/19/2009
        Combo1.AddItem "MF-" & UCD_MASTERFORMAT_VERSION
        Combo1.ItemData(Combo1.NewIndex) = UCD_MASTERFORMAT_VERSION
        If sDefaultMF = UCD_MASTERFORMAT_VERSION Then
            Combo1.ListIndex = Combo1.NewIndex
        End If
    
        If Not bNoAltIDSelection Then
            Combo1.AddItem "MF-" & ALT_MASTERFORMAT_VERSION
            Combo1.ItemData(Combo1.NewIndex) = ALT_MASTERFORMAT_VERSION
            If sDefaultMF = ALT_MASTERFORMAT_VERSION Then
                Combo1.ListIndex = Combo1.NewIndex
            End If
        End If
    End If 'rlh
    LoadMasterFormatCombo = iIndex

End Function

Public Function debugrs(ByVal rs As ADODB.RecordSet)
    Dim f As ADODB.Field
    For Each f In rs.Fields
        Select Case f.Type
        Case adBoolean
            Debug.Print vbTab & f.Name & vbTab & " bit"
        Case adInteger
            Debug.Print vbTab & f.Name & vbTab & " int"
        Case adSmallInt
            Debug.Print vbTab & f.Name & vbTab & " smallint"
        Case adUnsignedTinyInt
            Debug.Print vbTab & f.Name & vbTab & " tinyint"
        Case adBigInt
            Debug.Print vbTab & f.Name & vbTab & " bigint"
        Case adDate
            Debug.Print vbTab & f.Name & vbTab & " date"
        Case adDBTimeStamp
            Debug.Print vbTab & f.Name & vbTab & " datetime"
        Case adCurrency
            Debug.Print vbTab & f.Name & vbTab & " money"
        Case adNumeric
            Debug.Print vbTab & f.Name & vbTab & " decimal(" & f.definedSize & "," & f.Precision & ")"
        Case adSingle
            Debug.Print vbTab & f.Name & vbTab & " real(" & f.definedSize & ")"
        Case adDouble
            Debug.Print vbTab & f.Name & vbTab & " float(" & f.definedSize & ")"
        Case adChar
            Debug.Print vbTab & f.Name & vbTab & " char(" & f.definedSize & ")"
        Case adVarChar
            Debug.Print vbTab & f.Name & vbTab & " varchar(" & f.definedSize & ")"
        Case adVarWChar
            Debug.Print vbTab & f.Name & vbTab & " nvarchar(" & f.definedSize & ")"
        Case Else
            Debug.Print vbTab & f.Name & vbTab & " ?"
        End Select
    Next
End Function

Public Sub CenterFormInParent(ByRef cfrm As Form, ByVal pfrm As Form)
    
    On Error Resume Next
    cfrm.Move (pfrm.Width - cfrm.Width) / 2, (pfrm.Height - cfrm.Height) / 3

End Sub

Public Function LaunchBrowser(ByVal URL As String) As Boolean
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

Public Sub SetColors()
'SET THE GLOBAL COLOR VARIABLES
'USED BY THE GRIDS AND LOCKED CONTROLS
    Dim iColor1 As Long
    Dim iColor2 As Long
    
    'CALCULATE THE ALTERNATE GRID ROW COLOR --
    'AVERAGE BETWEEN FORM COLOR AND THE WINDOW BACKGROUND COLOR (USUALLY WHITE)
    iColor1 = GetSysColor(COLOR_BTNFACE)
    iColor2 = GetSysColor(COLOR_WINDOW)
    g_intAlternateRowColor = (iColor1 + iColor2 + ((iColor1 Xor iColor2) And &H10101)) \ 2

End Sub

Public Function FormatUnitCost(ByVal sUnitCostId As String, ByVal iMasterFormatVersion As Long) As String
'RETURN THE UNIT COST ID WITH PROPER CSI MASTERFORMAT SPACING
'ADDED 8/23/2005 RTD
    Dim sFormatted As String
    
    sFormatted = Compress_String(sUnitCostId)
    Select Case iMasterFormatVersion
        Case 2004
            sFormatted = Format(sFormatted, FORMAT_UNIT_COST_04_SRV)
        Case 1995
            sFormatted = Format(sFormatted, FORMAT_UNIT_COST_SRV)
    End Select
    FormatUnitCost = sFormatted

End Function

Public Function FindComboItemDataIndex(ByVal Combo1 As ComboBox, ByVal ItemData As Long) As Long
'ADDED 9/8/2005 RTD
'RETURN THE LISTINDEX OF THE FIRST ITEM IN THE COMBOBOX WHOSE ITEMDATA PROPERTY MATCHES
    Dim i As Long
    Dim Index As Long
    
    Index = -1
    For i = 0 To Combo1.listcount - 1
        If Combo1.ItemData(i) = ItemData Then
            Index = i
            Exit For
        End If
    Next
    FindComboItemDataIndex = Index

End Function

Public Function Login() As Boolean
'ADDED 9/13/2005 RTD
'UPDATE USER STATS IN THE CCD USER_NAME TABLE
    Dim user As New cUserInfo
    
    user.UserID = strUserName
    'user.GetData
    If user.Login = True Then
        g_blnIsUserAdmin = user.isAdmin
        Login = True
    Else
        Set user = Nothing
        Login = False
    End If
    
End Function

Public Function Update_Tree_With_Unit_Cost_Id(ByVal ID As String, ByVal alt_id As String) As Boolean

    If (g_intMasterFormat = 2004) Then
    
        Dim blnRetVal As Boolean
        Update_Tree_With_Unit_Cost_Id = False
        'create missing Headers for the Tree based on the unit_cost_id (if any)
        On Error GoTo Err_Handler
                 
        Dim cmd As New ADODB.Command
        Dim param As ADODB.Parameter
         
        'pass info into stored proc and update hier tree
                                    
        If Mid(ID, 1, 1) = "M" Then
            ID = Mid(ID, 2)
        End If
        If Mid(alt_id, 1, 1) = "M" Then
            alt_id = Mid(alt_id, 2)
        End If
        
                                    
        cmd.CommandText = "sp_update_tree_with_mat_or_cost_id"
        cmd.CommandType = CommandTypeEnum.adCmdStoredProc
        
        Set param = cmd.CreateParameter("id_passed", adVarChar, adParamInput, 12, Compress_String(ID))
        cmd.Parameters.Append param
        Set param = cmd.CreateParameter("alt_id_passed", adVarChar, adParamInput, 12, Compress_String(alt_id))
        cmd.Parameters.Append param
        Set param = cmd.CreateParameter("last_update_person_passed", adVarChar, adParamInput, 15, Mid(strUserName, 1, 15))
        cmd.Parameters.Append param
        
        ' Assuming a connection has been established and a recordset has
        ' been created
        Set cmd.ActiveConnection = g_cnShared
        Dim RecordSet As ADODB.RecordSet
        Set RecordSet = cmd.Execute()

    End If
    
    Update_Tree_With_Unit_Cost_Id = True

    Exit Function
Err_Handler:
    Dim errMessage As String
    errMessage = "MainModule:Update_Tree_with_Unit_Cost_Id - " + Err.Description
    Debug.Print errMessage
    

End Function


Public Function Update_MasterFormat04_ID_Hierarchy_Totals_Only() As Boolean

    Dim blnRetVal As Boolean
    Update_MasterFormat04_ID_Hierarchy_Totals_Only = False

    On Error GoTo Err_Handler
             
    Dim cmd As New ADODB.Command
    Dim param As ADODB.Parameter
     
    'pass info into stored proc and update hier tree
                                
                                
    cmd.CommandText = "sp_update_masterformat04_id_hierarchy_totals_only"
    cmd.CommandType = CommandTypeEnum.adCmdStoredProc
    
    ' Assuming a connection has been established and a recordset has
    ' been created
    Set cmd.ActiveConnection = g_cnShared
    Dim RecordSet As ADODB.RecordSet
    Set RecordSet = cmd.Execute()

    Update_MasterFormat04_ID_Hierarchy_Totals_Only = True

    Exit Function
Err_Handler:
    Dim errMessage As String
    errMessage = "MainModule:Update_MasterFormat04_ID_Hierarchy_Totals_Only - " + Err.Description
    Debug.Print errMessage
    

End Function
