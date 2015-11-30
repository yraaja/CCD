VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Construction Cost Database"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "About ConstructionCostDatabase"
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   930
      Left            =   120
      Picture         =   "frmAbout.frx":0442
      ScaleHeight     =   870
      ScaleWidth      =   3000
      TabIndex        =   9
      Top             =   480
      Width           =   3060
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   6000
      TabIndex        =   0
      Tag             =   "OK"
      Top             =   3165
      Width           =   1467
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   345
      Left            =   6000
      TabIndex        =   1
      Tag             =   "&System Info..."
      Top             =   3615
      Width           =   1452
   End
   Begin VB.Label lblDALVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Access Layer ccddal.dll Version"
      Height          =   195
      Left            =   3360
      TabIndex        =   11
      Top             =   2280
      Width           =   2640
   End
   Begin VB.Label lblMeansCtrlVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Means Control meansctrl.ocx Version"
      Height          =   195
      Left            =   3360
      TabIndex        =   10
      Top             =   2040
      Width           =   2610
   End
   Begin VB.Label lblConnection 
      Height          =   255
      Left            =   3360
      TabIndex        =   8
      Top             =   2640
      Width           =   4215
   End
   Begin VB.Label lblLastChange 
      Height          =   255
      Left            =   1140
      TabIndex        =   7
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Last Change:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   945
   End
   Begin VB.Label lblDescription 
      Caption         =   "App Description"
      ForeColor       =   &H00000000&
      Height          =   930
      Left            =   3390
      TabIndex        =   5
      Tag             =   "App Description"
      Top             =   945
      Width           =   4095
   End
   Begin VB.Label lblTitle 
      Caption         =   "Construction Cost Database"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   3390
      TabIndex        =   4
      Tag             =   "Application Title"
      Top             =   60
      Width           =   4095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   120
      X2              =   7560
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      Index           =   0
      X1              =   135
      X2              =   7555
      Y1              =   3015
      Y2              =   3015
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3390
      TabIndex        =   3
      Tag             =   "Version"
      Top             =   600
      Width           =   4095
   End
   Begin VB.Label lblDisclaimer 
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   240
      TabIndex        =   2
      Tag             =   "Warning: ..."
      Top             =   3165
      Width           =   5250
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Reg Key Security Options...
Const KEY_ALL_ACCESS = &H2003F
                                          

' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number


Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"


Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private m_blnDoubleClick As Boolean

Private Sub Form_Activate()
    
    OutputView False

End Sub

Private Sub Form_Load()
    Dim Found As Boolean
    Dim szFileName As String
    Dim strTimeStamp As String

    'szFileName = "C:\Program Files\Construction Cost Database\ConstructionCostDatabase.exe"
    szFileName = App.Path
    If Right(szFileName, 1) <> "\" Then szFileName = szFileName & "\"
    szFileName = szFileName & App.EXEName & ".exe"
    Found = FileExist(szFileName)
    If Found Then
        strTimeStamp = FileDateTime(szFileName)
    Else
        strTimeStamp = "n/a"
    End If
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.title
    lblDescription.Caption = App.FileDescription
    lblConnection.Caption = "Connection: " + strConnectServer + ":" + strConnectDatabase
    lblDisclaimer.Caption = App.LegalCopyright
    lblLastChange = strTimeStamp
    'ADDED 6/23/2005 RTD
    lblMeansCtrlVersion.Caption = "RSMeans Tree Control (meansctrl.ocx) Version " & GetFileVersion("meansctrl.ocx")
    lblDALVersion.Caption = "Data Access Layer (ccddal.dll) Version " & GetFileVersion("ccddal.dll")
    
End Sub



Private Sub cmdSysInfo_Click()
        Call StartSysInfo
End Sub


Private Sub cmdOK_Click()
        Unload Me
End Sub


Public Sub StartSysInfo()
    On Error GoTo SysInfoErr


        Dim rc As Long
        Dim SysInfoPath As String
        

        ' Try To Get System Info Program Path\Name From Registry...
        If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
        ' Try To Get System Info Program Path Only From Registry...
        ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
                ' Validate Existance Of Known 32 Bit File Version
                If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
                        SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
                        

                ' Error - File Can Not Be Found...
                Else
                        GoTo SysInfoErr
                End If
        ' Error - Registry Entry Can Not Be Found...
        Else
                GoTo SysInfoErr
        End If
        

        Call Shell(SysInfoPath, vbNormalFocus)
        

        Exit Sub
SysInfoErr:
        MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub


Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
        Dim i As Long                                           ' Loop Counter
        Dim rc As Long                                          ' Return Code
        Dim hKey As Long                                        ' Handle To An Open Registry Key
        Dim hDepth As Long                                      '
        Dim KeyValType As Long                                  ' Data Type Of A Registry Key
        Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
        Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
        '------------------------------------------------------------
        ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
        '------------------------------------------------------------
        rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
        

        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
        

        tmpVal = String$(1024, 0)                             ' Allocate Variable Space
        KeyValSize = 1024                                       ' Mark Variable Size
        

        '------------------------------------------------------------
        ' Retrieve Registry Key Value...
        '------------------------------------------------------------
        rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                                                

        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
        

        tmpVal = VBA.Left(tmpVal, InStr(tmpVal, VBA.Chr(0)) - 1)
        '------------------------------------------------------------
        ' Determine Key Value Type For Conversion...
        '------------------------------------------------------------
        Select Case KeyValType                                  ' Search Data Types...
        Case REG_SZ                                             ' String Registry Key Data Type
                KeyVal = tmpVal                                     ' Copy String Value
        Case REG_DWORD                                          ' Double Word Registry Key Data Type
                For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
                        KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
                Next
                KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
        End Select
        

        GetKeyValue = True                                      ' Return Success
        rc = RegCloseKey(hKey)                                  ' Close Registry Key
        Exit Function                                           ' Exit
        

GetKeyError:    ' Cleanup After An Error Has Occured...
        KeyVal = ""                                             ' Set Return Val To Empty String
        GetKeyValue = False                                     ' Return Failure
        rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Private Sub picIcon_DblClick()
    ' Signal that double-click has occurred
    m_blnDoubleClick = True
End Sub

Private Sub picIcon_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If this is the mouse-up form a double click
    If m_blnDoubleClick Then
        ' Make sure it is the left button
        If Button = vbRightButton Then
            ' Same function as clicking Unit Cost button, open single record view
            MsgBox "Originally developed by Douglyss Giuliana.", vbOKOnly, "You Found The Easter Egg!"
        End If
    End If
    m_blnDoubleClick = False
End Sub

Private Function FileExist(ByVal szFileName As String) As Boolean
Dim nFileNumber As Integer
On Error Resume Next
nFileNumber = FreeFile

'Try to open the file
Open szFileName For Input As nFileNumber

'If it fails the file doesn't exist
If Err.Number <> 0 Then
    FileExist = False
Else
    FileExist = True
End If

Close nFileNumber
End Function

