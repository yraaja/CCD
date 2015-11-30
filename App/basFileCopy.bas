Attribute VB_Name = "basFileCopy"
Option Explicit

Private Type SHFILEOPSTRUCT
   hWnd        As Long
   wFunc       As Long
   pFrom       As String
   pTo         As String
   fFlags      As Integer
   fAborted    As Boolean
   hNameMaps   As Long
   sProgress   As String
 End Type
  
Private Const FO_MOVE As Long = &H1
Private Const FO_COPY As Long = &H2
Private Const FO_DELETE As Long = &H3
Private Const FO_RENAME As Long = &H4

Private Const FOF_SILENT As Long = &H4
Private Const FOF_RENAMEONCOLLISION As Long = &H8
Private Const FOF_NOCONFIRMATION As Long = &H10
Private Const FOF_ALLOWUNDO As Long = &H40
Private Const FOF_FILESONLY As Long = &H80
Private Const FOF_SIMPLEPROGRESS As Long = &H100
Private Const FOF_NOCONFIRMMKDIR As Long = &H200

Private Declare Function SHFileOperation Lib "shell32" Alias "SHFileOperationA" _
    (lpFileOp As SHFILEOPSTRUCT) As Long
  
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" _
    (ByVal pidl As Long, ByVal pszPath As String) As Long


Public Function PerformShellAction(sSource As String, sDestination As String) As Long
    Dim SHFileOp As SHFILEOPSTRUCT
    
    'terminate the folder string with a pair of nulls
    sSource = sSource & Chr$(0) & Chr$(0)
    sDestination = sDestination & Chr$(0) & Chr$(0)
    
    'set up the options
    With SHFileOp
       .wFunc = FO_COPY
       .pFrom = sSource
       .pTo = sDestination
       .fFlags = FOF_NOCONFIRMATION
       .sProgress = "Upgrading Application Files..."
    End With
    PerformShellAction = SHFileOperation(SHFileOp)
        
End Function

Public Function CopyFolderToAppPath(ByVal sSourcePath As String) As Boolean
' COPY ALL FILES FROM SOURCE PATH INTO THE APPLICATION FOLDER
    Dim sAppPath As String
    Dim iReturn As Long
    
    sAppPath = App.Path
    If Right(sAppPath, 1) <> "\" Then sAppPath = sAppPath & "\"
    If Right(sSourcePath, 1) <> "\" Then sSourcePath = sSourcePath & "\"

    sSourcePath = sSourcePath & "*.*"

    iReturn = PerformShellAction(sSourcePath, sAppPath)
    CopyFolderToAppPath = (iReturn = 0)
    
End Function

