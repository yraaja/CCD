VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRSMDataAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Title   - Construction Cost Database Database Access Layer
'Date    - 12/17/98
'Author  - Douglyss Giuliana, Edgewater Technology, Inc.
'
'Comment - This ActiveX DLL provides data access functionality through a COM object
'          to the client components of the CCD. This class uses ADO.

'Updated - 8/16/2005 Rob Durfee - Added properties to report error information
'          back to CCD: LastErrorDescription and LastErrorNumber
'

Option Explicit

Private Const OBJNAME = "CRSMDataAccess"

Private m_oCachedConnection As ADODB.Connection
Private m_bInsideTransaction As Boolean
Private m_iLastErrorNum As Long         ' 8/16/2005 RTD
Private m_sLastErrorDesc As String      ' 8/16/2005 RTD

Public Enum OPENTRECORDSETOPTIONS
    OO_UseTable = ADODB.adCmdTable
    OO_UseText = ADODB.adCmdText
    OO_UseStoredProc = ADODB.adCmdStoredProc
    OO_UseUnknown = ADODB.adCmdUnknown
    OO_UseDefault = -1
End Enum

Public Enum LOCKTYPES
    LT_ReadOnly = ADODB.adLockReadOnly
    LT_Optimistic = ADODB.adLockOptimistic
    LT_Pessimistic = ADODB.adLockPessimistic
    LT_BatchOptimistic = ADODB.adLockBatchOptimistic
End Enum

Public Enum TRANSTYPES
    TT_DropConnection = 0
    TT_KeepConnection = 1
End Enum

'=======================================================================
'Synopsis:              This function begins a new transaction
'Function input:        Connect string which is a registered DSN
'Function output:       Returns TRUE if the function succeeds
'Remarks
'                       If Connect is passed in then a connection is made
'                       to that DSN, else the previously established
'                       connection is used
'=======================================================================
Public Function BeginTrans(Optional ByVal Connect As String) As Boolean
Dim lErrNo      As Long
Dim sErrDesc    As String

    'assume failure
    BeginTrans = False
    
    'enable error handler
    On Error GoTo ErrorHandler
    
    'establish the transaction if DSN specified
    If Connect <> vbNullString Then
        If Not m_oCachedConnection Is Nothing Then
            m_oCachedConnection.Close
            Set m_oCachedConnection = Nothing
        End If
        Set m_oCachedConnection = New ADODB.Connection
        m_oCachedConnection.Open Connect
    End If
    
    'signal beginning of transaction
    m_oCachedConnection.BeginTrans
    
    'signal start of transaction, success of function and we're out of here
    m_bInsideTransaction = True
    BeginTrans = True
    Exit Function
    
'if we're here there then's been an error so process
ErrorHandler:
    'store incoming values and raise error
    lErrNo = Err.Number
    sErrDesc = Err.Description
    LastErrorNumber = lErrNo
    LastErrorDescription = sErrDesc
    On Error GoTo 0
    Err.Raise lErrNo, OBJNAME, sErrDesc

End Function

'=======================================================================
'Synopsis:              This function either opens or closes a cached ADO connection.
'                       If Connect is supplied then a connection is made and cached.
'                       If it is not supplied then the connection is closed
'Function input:        optional Connection string
'Function output:       Returns TRUE if the function succeeds
'=======================================================================
Public Function CacheConnection(Optional ByVal Connect As String) As Boolean
Dim lErrNo      As Long
Dim sErrDesc    As String

    'assume failure
    CacheConnection = False
    
    'enable error handler
    On Error GoTo ErrorHandler

    'close any existing connection
    If Not m_oCachedConnection Is Nothing Then
        m_oCachedConnection.Close
        Set m_oCachedConnection = Nothing
    End If

    'if parameter not empty then make connection
    If Connect <> vbNullString Then
        Set m_oCachedConnection = New ADODB.Connection
        m_oCachedConnection.Open Connect
    End If
    
    'signal success of function and we're out of here
    CacheConnection = True
    Exit Function
    
'if we're here there then's been an error so process
ErrorHandler:
    'store incoming values
    lErrNo = Err.Number
    sErrDesc = Err.Description
    LastErrorNumber = lErrNo
    LastErrorDescription = sErrDesc
    
    'close connection, and raise error
    On Error Resume Next
    Set m_oCachedConnection = Nothing
    On Error GoTo 0
    Err.Raise lErrNo, OBJNAME, sErrDesc
    
End Function


'=======================================================================
'Synopsis:              This function commits a previously begun transaction
'Function input:        Connect string which is a registered DSN. If supplied
'                       then previously established connection is closed
'Function output:       Returns TRUE if the function succeeds
'=======================================================================
Public Function CommitTrans(Optional ByVal KeepOpen As TRANSTYPES = TT_DropConnection) As Boolean
Dim lErrNo      As Long
Dim sErrDesc    As String

    'assume failure
    CommitTrans = False
    
    'enable error handler
    On Error GoTo ErrorHandler
    
    'commit the transaction, closing down the connection if DSN specified
    m_oCachedConnection.CommitTrans
    If KeepOpen = TT_DropConnection Then
        m_oCachedConnection.Close
        Set m_oCachedConnection = Nothing
    End If
    
    'signal closure of transaction, success of function and we're out of here
    m_bInsideTransaction = False
    CommitTrans = True
    Exit Function
    
'if we're here there then's been an error so process
ErrorHandler:
    'store incoming values and raise error
    lErrNo = Err.Number
    sErrDesc = Err.Description
    LastErrorNumber = lErrNo
    LastErrorDescription = sErrDesc
    On Error GoTo 0
    Err.Raise lErrNo, OBJNAME, sErrDesc

End Function

'=======================================================================
'Synopsis:              This function opens a connection using ADO and
'                       executes the action query passed to it.
'Function input:        Connect string which is a registered DSN
'                       Query string to be executed.
'Function output:       Returns TRUE if the function succeeds
'=======================================================================
Public Function ExecQuery( _
                    ByVal Connect As String, _
                    ByVal SQLQuery As String, Optional ByRef ErrorDesc As String = "") As Boolean
Dim oConn       As ADODB.Connection
Dim lErrNo      As Long
Dim sErrDesc    As String

    'assume failure
    ExecQuery = False
    
    'enable error handler
    On Error GoTo ErrorHandler
    
    'Get connection
    If Connect <> vbNullString Then
        Set oConn = New ADODB.Connection
        oConn.Open Connect
    Else
        Set oConn = m_oCachedConnection
    End If
    
    'set up for transaction, execute query, commit transaction, and clean up
    oConn.Execute SQLQuery
    If Connect <> vbNullString Then
        oConn.Close
    End If
    Set oConn = Nothing
    
    'signal success of function and we're out of here
    ExecQuery = True
    Exit Function
    
'if we're here there then's been an error so process
ErrorHandler:
    oConn.Cancel
    'store incoming values
    lErrNo = Err.Number
    sErrDesc = Err.Description
    ErrorDesc = Err.Description
    LastErrorNumber = lErrNo
    LastErrorDescription = sErrDesc
    
    'close connection, and raise error
    On Error Resume Next
    If Connect <> vbNullString Then
        oConn.Close
    End If
    On Error GoTo 0
    Err.Raise lErrNo, OBJNAME, sErrDesc
    
End Function

'=======================================================================
'Synopsis:              Function opens a connection using ADO and executes
'                       a query passed to it. This function returns
'                       records in a recordset.
'Function input:        Connect string which is a registered DSN
'                       Query string to be executed.
'                       A recordset to return records in
'Function output:       Returns TRUE if the function succeeds
'=======================================================================
Public Function GetRecordset( _
                    ByVal Connect As String, _
                    ByVal SQLQuery As String, _
                    ByRef Recordset As ADODB.Recordset) As Boolean
Dim oConn       As ADODB.Connection
Dim lErrNo      As Long
Dim sErrDesc    As String

    'assume failure
    GetRecordset = False
    
    'enable error handler
    On Error GoTo ErrorHandler
    
    'close recordset if open
    '9/1/2005 RTD - Corrects issue reported by K. Rodriguez/J. Chiang on 8/31/05
    If Not (Recordset Is Nothing) Then
        If Recordset.State = adStateOpen Then
            Recordset.Close
        End If
    End If
    
    'Get connection and set up for recordset
    If Connect <> vbNullString Then
        Set oConn = New ADODB.Connection
        oConn.Open Connect
    Else
        Set oConn = m_oCachedConnection
    End If
    oConn.CommandTimeout = 300  'AKD - 08/20/2006 - to avoid early termination of Query execution
    'MsgBox "DAL:GetRecordSet - Command Time Out: " & oConn.CommandTimeout
    
    'execute query and get recordset
    If Recordset Is Nothing Then
        Set Recordset = New ADODB.Recordset
    End If
    Recordset.CursorLocation = adUseClient
    Recordset.Open _
        Source:=SQLQuery, _
        ActiveConnection:=oConn, _
        CursorType:=adOpenStatic, _
        LockType:=adLockBatchOptimistic
        
    'go to first record if records returns
    If Recordset.RecordCount Then
        Recordset.MoveFirst
    End If
        
    'we're finished so clean up, signal succes, and exit
    Set Recordset.ActiveConnection = Nothing
    If Connect <> vbNullString Then
        oConn.Close
    End If
    Set oConn = Nothing
    
    'signal success of function and we're out of here
    GetRecordset = True
    Exit Function
    
'if we're here there then's been an error so process
ErrorHandler:
    'store incoming values
    lErrNo = Err.Number
    sErrDesc = Err.Description
    LastErrorNumber = lErrNo
    LastErrorDescription = sErrDesc
    
    'close down what we can and raise error
    On Error Resume Next
    Set Recordset = Nothing
    If Connect <> vbNullString Then
        oConn.Close
    End If
    On Error GoTo 0
    Err.Raise lErrNo, OBJNAME, sErrDesc

End Function

'=======================================================================
'Synopsis:              Function opens a connection using ADO and executes
'                       a query passed to it. This function returns
'                       records in a recordset.
'Function input:        Connect string which is a registered DSN
'                       Query string to be executed.
'                       A recordset to return records in
'Function output:       Returns TRUE if the function succeeds
'=======================================================================
Public Function GetSPRecordset( _
                    ByVal Connect As String, _
                    ByVal SQLQuery As String, _
                    ByRef Recordset As ADODB.Recordset) As Boolean
Dim oConn       As ADODB.Connection
Dim oCmd        As New ADODB.Command
Dim lErrNo      As Long
Dim sErrDesc    As String

    'assume failure
    GetSPRecordset = False
    
    'enable error handler
    On Error GoTo ErrorHandler
    
    'Get connection and set up for recordset
    If Connect <> vbNullString Then
        Set oConn = New ADODB.Connection
        oConn.Open Connect
    Else
        Set oConn = m_oCachedConnection
    End If
    
    oConn.CommandTimeout = 300      'AKD - 08/20/2006
    oConn.ConnectionTimeout = 300   'AKD - 08/20/2006
    oCmd.ActiveConnection = oConn
    oCmd.CommandText = SQLQuery
    oCmd.CommandType = adCmdStoredProc
    
    'MsgBox "Command Time Out: " & oConn.CommandTimeout
        
    Set Recordset = oCmd.Execute()
        
    'go to first record if records returns
    If Recordset.RecordCount Then
        Recordset.MoveFirst
    End If
        
    'we're finished so clean up, signal succes, and exit
    Set oCmd.ActiveConnection = Nothing
    If Connect <> vbNullString Then
        oConn.Close
    End If
    Set oConn = Nothing
    
    'signal success of function and we're out of here
    GetSPRecordset = True
    Exit Function
    
'if we're here there then's been an error so process
ErrorHandler:
    'store incoming values
    lErrNo = Err.Number
    sErrDesc = Err.Description
    LastErrorNumber = lErrNo
    LastErrorDescription = sErrDesc
    
    'close down what we can and raise error
    On Error Resume Next
    Set Recordset = Nothing
    If Connect <> vbNullString Then
        oConn.Close
    End If
    On Error GoTo 0
    Err.Raise lErrNo, OBJNAME, sErrDesc

End Function


'=======================================================================
'Synopsis:              Function returns a boolean flag indicating whether
'                       a transacted connection is in progress
'Function output:       Returns TRUE if transacted connection
'=======================================================================
Public Property Get IsInTransaction() As Boolean
Attribute IsInTransaction.VB_Description = "Returns whether a transacted connection is in progress"
    
    IsInTransaction = m_bInsideTransaction

End Property

'=======================================================================
'Synopsis:              This function opens a transacted recordset on a
'                       previously established connection
'Function input:        Query string to be executed.
'                       A recordset to return records in
'                       optional record locking specifier
'                       optional options for ADO
'Function output:       Returns TRUE if the function succeeds
'=======================================================================
Public Function OpenTRecordset( _
                    ByVal SQLQuery As String, _
                    ByRef Recordset As ADODB.Recordset, _
                    Optional ByVal LockType As LOCKTYPES = LT_Optimistic, _
                    Optional ByVal Options As OPENTRECORDSETOPTIONS = OO_UseDefault) _
                    As Boolean
Dim lErrNo      As Long
Dim sErrDesc    As String

    'assume failure
    OpenTRecordset = False
    
    'enable error handler
    On Error GoTo ErrorHandler
    
    'lock as a batch if no lock type specified
    If LockType = 0 Then
        LockType = adLockBatchOptimistic
    End If
    
    'execute query and get SERVER SIDE recordset
    Set Recordset = New ADODB.Recordset
    Recordset.ActiveConnection = m_oCachedConnection
    Recordset.CursorType = adOpenKeyset
    Recordset.LockType = LockType
    Recordset.Open SQLQuery, , , , Options

    'go to first record if records returns
    If Recordset.RecordCount Then
        Recordset.MoveFirst
    End If
        
    'signal success of function and we're out of here
    OpenTRecordset = True
    Exit Function
    
'if we're here there then's been an error so process
ErrorHandler:
    'store incoming values
    lErrNo = Err.Number
    sErrDesc = Err.Description
    LastErrorNumber = lErrNo
    LastErrorDescription = sErrDesc
    
    'close down what we can and raise error
    On Error Resume Next
    Set Recordset = Nothing
    On Error GoTo 0
    Err.Raise lErrNo, OBJNAME, sErrDesc

End Function


'=======================================================================
'Synopsis:              Function opens a connection using ADO and updates the database
'                       using a passed in recordset
'Function input:        Connect string which is a registered DSN
'                       A recordset to update database with
'Function output:       Returns TRUE if the function succeeds
'=======================================================================
Public Function PutRecordset( _
                    ByVal Connect As String, _
                    Recordset As ADODB.Recordset) As Boolean
Dim oConn       As ADODB.Connection
Dim lErrNo      As Long
Dim sErrDesc    As String

    'assume failure
    PutRecordset = False
    
    'enable error handler
    On Error GoTo ErrorHandler
    
    'Get connection and set up for recordset
    If Connect <> vbNullString Then
        Set oConn = New ADODB.Connection
        oConn.Open Connect
    Else
        Set oConn = m_oCachedConnection
    End If
    
    'connect to database and update records
    Set Recordset.ActiveConnection = oConn
    Recordset.UpdateBatch
    Set Recordset.ActiveConnection = Nothing
    If Connect <> vbNullString Then
        oConn.Close
    End If
    Set oConn = Nothing
    
    'signal success of function and we're out of here
    PutRecordset = True
    Exit Function

'if we're here there then's been an error so process
ErrorHandler:
    'store incoming values
    lErrNo = Err.Number
    sErrDesc = Err.Description
    LastErrorNumber = lErrNo
    LastErrorDescription = sErrDesc
    
    'close down what we can and raise error
    On Error Resume Next
    If Connect <> vbNullString Then
        oConn.Close
    End If
    On Error GoTo 0
    Err.Raise lErrNo, OBJNAME, sErrDesc

End Function
'=======================================================================
'Synopsis:              This function begins a new transaction
'
'Function input:        Connect string which is a registered DSN. If not supplied
'                       then previously established connection is used
'Function output:       Returns TRUE if the function succeeds
'=======================================================================
Public Function CloseTRecordset(Recordset As ADODB.Recordset) As Boolean
Dim lErrNo      As Long
Dim sErrDesc    As String

    'assume failure
    CloseTRecordset = False
    
    'enable error handler
    On Error GoTo ErrorHandler
    
    'should be simple enough
    Recordset.Close

    'signal success of function and we're out of here
    CloseTRecordset = True
    Exit Function

'if we're here there then's been an error so process
ErrorHandler:
    'store incoming values
    lErrNo = Err.Number
    sErrDesc = Err.Description
    LastErrorNumber = lErrNo
    LastErrorDescription = sErrDesc
    
    'raise error
    On Error GoTo 0
    Err.Raise lErrNo, OBJNAME, sErrDesc

End Function


'=======================================================================
'Synopsis:              This function rolls back a previously begun transaction
'
'Function input:        Connect string which is a registered DSN. If supplied
'                       then previously established connection is closed
'Function output:       Returns TRUE if the function succeeds
'=======================================================================
Public Function RollbackTrans(Optional ByVal KeepOpen As TRANSTYPES = TT_DropConnection) As Boolean
Dim lErrNo      As Long
Dim sErrDesc    As String

    'assume failure
    RollbackTrans = False
    
    'enable error handler
    On Error GoTo ErrorHandler
    
    'roll back the transaction, closing down the connection if DSN specified
    m_oCachedConnection.RollbackTrans
    If KeepOpen = TT_DropConnection Then
        m_oCachedConnection.Close
        Set m_oCachedConnection = Nothing
    End If

    'signal closure of transaction, success of function and we're out of here
    m_bInsideTransaction = False
    RollbackTrans = True
    Exit Function
    
'if we're here there then's been an error so process
ErrorHandler:
    'store incoming values
    lErrNo = Err.Number
    sErrDesc = Err.Description
    LastErrorNumber = lErrNo
    LastErrorDescription = sErrDesc
    
    'raise error
    On Error GoTo 0
    Err.Raise lErrNo, OBJNAME, sErrDesc

End Function

Private Sub Class_Terminate()

    'clean up any outstanding connections
    If Not m_oCachedConnection Is Nothing Then
        m_oCachedConnection.Close
        Set m_oCachedConnection = Nothing
    End If
    If Not m_oCachedConnection Is Nothing Then
        m_oCachedConnection.Close
        Set m_oCachedConnection = Nothing
    End If

End Sub


'=======================================================================
'Synopsis:              Function verifies data access component installation by
'                       attempting to instantiate an ado connection and recordset
'                       attempts to open connection connection string is supplied
'Function output:       Returns TRUE if successful
'=======================================================================
Public Function DataComponentsInstalled(Optional ByVal Connect As String) As Boolean
Dim oConn       As Object
Dim Recordset   As ADODB.Recordset
Dim lErrNo      As Long
Dim sErrDesc    As String
    
    'assume failure
    DataComponentsInstalled = False

    'enable error handler
    On Error GoTo ErrorHandler
    
    'default to false for function
    DataComponentsInstalled = False
    
    'see if we can create objects
    Set oConn = New ADODB.Connection
    Set Recordset = New ADODB.Recordset
    
    'if connection string passed in then attempt connection (BEXT DETECTION METHOD)
    If Connect <> vbNullString Then
        oConn.Open Connect
        oConn.Close
    End If

    'objects instantiated without error so signal succes, clean up and exit
    Set Recordset = Nothing
    Set oConn = Nothing

    DataComponentsInstalled = True
    Exit Function

ErrorHandler:
    'store incoming values
    lErrNo = Err.Number
    sErrDesc = Err.Description
    LastErrorNumber = lErrNo
    LastErrorDescription = sErrDesc
    
    'close down what we can and raise error
    On Error Resume Next
    Set Recordset = Nothing
    Set oConn = Nothing

    On Error GoTo 0
    Err.Raise lErrNo, OBJNAME, sErrDesc
End Function

'=======================================================================
'Synopsis:              Gets/Sets the last database error text that occurred
'                       Allows CCD to receive/process the exact error message
'Function output:       N/A
'=======================================================================
Public Property Get LastErrorDescription() As String
Attribute LastErrorDescription.VB_Description = "Returns/Sets the error description of the last database error that occurred"
    LastErrorDescription = m_sLastErrorDesc
End Property
Public Property Let LastErrorDescription(NewValue As String)
    m_sLastErrorDesc = NewValue
End Property

'=======================================================================
'Synopsis:              Gets/Sets the last database error number that occurred
'                       Allows CCD to receive/process the exact error message
'Function output:       N/A
'=======================================================================
Public Property Get LastErrorNumber() As Long
Attribute LastErrorNumber.VB_Description = "Returns/Sets the error number of the last database error that occurred"
    LastErrorNumber = m_iLastErrorNum
End Property
Public Property Let LastErrorNumber(NewValue As Long)
    m_iLastErrorNum = NewValue
End Property

