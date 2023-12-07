VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cadwADOWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "This class handles all ADODB.Recordset and ADODB.Command accesses using ADO 2.5."

'--------------------------------------------------------------------------
' Object:               cadwADOWrapper
' Object Description:   Complete database
' Revision History:
'    10/04/02 - Betsy - Cloned from TRS.
'
' Procedures  :
'   Private     Class_Initialize()
'   Private     Class_Terminate()
'   Public      Property Get ADOCommand() As ADODB.Command
'   Public      CommandInitialize()
'   Public      CommandSetConnection(Optional ByRef pconIn As cconConnection) As Boolean
'   Public      CommandSetSproc(strSprocName As String, _
'                   Optional ByVal lngParamCount As Long = 0) As Boolean
'   Public      Execute_SQL_AsRST(ByRef pconIn As cconConnection, ByVal strSQL As String) _
'                   As ADODB.Recordset
'   Public      Execute_SQL_UpdateableDisconnectedRST(ByRef pconIn As cconConnection, _
'                   ByVal strSQL As String) As ADODB.Recordset
'   Public      GetMetaData_Columns(ByVal strTableName As String, _
'                   ByRef prstInOut As ADODB.Recordset) As Boolean
'   Public      GetMetaData_PrimaryKeys(ByVal strTableName As String, _
'                   ByRef prstInOut As ADODB.Recordset) As Boolean
'   Public      MoveFirst(ByRef prstTemp As ADODB.Recordset) As Boolean
'   Public      MoveLast(ByRef prstTemp As ADODB.Recordset) As Boolean
'   Public      MoveNext(ByRef prstTemp As ADODB.Recordset) As Boolean
'   Public      MovePrev(ByRef prstTemp As ADODB.Recordset) As Boolean
'
'
'-----------------------------------------------------------------------------
Option Explicit
Option Compare Binary

Private Const mcstrName As String = "cadwADOWrapper."

Private m_cmdADOCommand          As ADODB.Command



'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'\                                                                  /
'|          CLASS_INITIALIZE / CLASS_TERMINATE    Procedures        |
'/                                                                  \
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/

'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub Class_Initialize()
    '--------------------------------------------------------------------------
    ' Procedure:   Class_Initialize
    ' Description: Starting point for the class.
    ' Params:      None
    ' Returns:     Boolean
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc       As String = "Class_Initialize"

    On Error GoTo PROC_ERR

    Set m_cmdADOCommand = New ADODB.Command
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub Class_Terminate()
    '--------------------------------------------------------------------------
    ' Procedure:   Terminate
    ' Description: kill the class
    ' Returns:
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc       As String = "Class_Terminate"

    On Error GoTo PROC_ERR

    If fnIsObject(m_cmdADOCommand) Then
        Set m_cmdADOCommand.ActiveConnection = Nothing
        fnFreeObject m_cmdADOCommand
    End If
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'\                                                                  /
'|                PROPERTY GET/LET    Procedures                    |
'/                                                                  \
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/

'////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Get ADOCommand() As ADODB.Command
    '--------------------------------------------------------------------------
    ' Procedure:   Get ADOCommand
    ' Description: Get the ADODB Command object (used to define and
    '              retrieve parameters and recordsets when calling
    '              stored procedures)
    ' Returns:     Boolean
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc   As String = "Property Get ADOCommand"

    On Error GoTo PROC_ERR

    Set ADOCommand = m_cmdADOCommand
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Property
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Property


'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'\                                                                  /
'|                        PUBLIC  Procedures                        |
'/                                                                  \
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/

'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function CommandInitialize() As Boolean
    '--------------------------------------------------------------------------
    ' Procedure:   CommandInitialize
    ' Description: Initializes the ADODB.Command object of this class
    '              so the caller knows it is ready for reuse. If any
    '              previous use defined parameters or established
    '              recordsets, they will be deleted.
    ' Returns:     True if successfully initialized; False otherwise
    ' Params:      N/A
    '
    ' Called by:   CommandSetSproc( )
    '
    ' Modified:
    '   05/17/02 BAW Removed the call to CommandSetConnection since this
    '                was interfering with the LogOn screen (as of 2C) using
    '                the Authenticate object (which calls a sproc)...before
    '                the gconAppActive connection object was instanatiated.
    '   04/23/02 BAW Parameter deletion was getting an error. Had to change it to
    '                count backward to avoid this.
    '--------------------------------------------------------------------------
    Const cstrCurrentProc       As String = "CommandInitialize"
    Dim intI As Integer

    On Error GoTo PROC_ERR

    If fnIsObject(m_cmdADOCommand) Then
        With m_cmdADOCommand
            fnFreeObject .ActiveConnection
            ' Setting ActiveConnection to Nothing still leaves Parameters collection intact,
            ' so make sure they're wiped out for the next use of this Command object.
            With .Parameters
                intI = .Count
                Do While intI > 0
                    intI = intI - 1
                    .DELETE intI
                Loop
            End With
            .CommandText = vbNullString
        End With
    Else
        Set m_cmdADOCommand = New ADODB.Command
    End If

    CommandInitialize = True
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function CommandSetConnection(Optional ByRef pconIn As cconConnection) As Boolean
    '--------------------------------------------------------------------------
    ' Procedure:   CommandSetConnection
    ' Description: Sets the ActiveConnection for the Command object. By default,
    '              the global Connection object (gconAppActive) will be used.
    ' Returns:     True if successful; False otherwise
    ' Params:
    '     pconIn (in/out) - a pointer to the cconConnection object to use, e.g., the
    '                       active or archive database.
    '--------------------------------------------------------------------------
    Const cstrCurrentProc       As String = "CommandSetConnection"

    On Error GoTo PROC_ERR

    ' Default to using the global Connection object that points to the active DB,
    ' unless the caller has specified another.
    With m_cmdADOCommand
        If (pconIn Is Nothing) Then
            .ActiveConnection = gconAppActive.ADOConn
        Else
            .ActiveConnection = pconIn.ADOConn
        End If

        ' Make sure Connection is open.
        ' (It should be for the entire session, once the user logs on)
        Debug.Assert (.ActiveConnection.State = adStateOpen)
    End With

    CommandSetConnection = True
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function CommandSetSproc(ByVal strSprocName As String, _
    Optional ByRef pconIn As cconConnection) As Boolean
    '--------------------------------------------------------------------------
    ' Procedure:   CommandSetSproc
    ' Description: Defines the stored procedure that will be called
    '              with this ADODB.Command object and establishes it
    '              as a method of that object.
    '
    ' Params:
    '   strSprocName (in)  - the name of the stored procedure to execute
    '   lngParamCount (in) - how many input, output or other parameters
    '                        will be used by this sproc
    '   pconIn (in/out)    - the name of the cconConnection object to use
    '                        when executing the sproc. If not supplied,
    '                        the gconAppActive object will be used.
    '
    ' Returns:     True if successful; False otherwise
    '
    ' Modifications:
    '   05/17/02 BAW - Added call to CommandInitialize and CommandSetConnection
    '                  to allow this to work for all existing code without
    '                  modification, but also support the new-to-Phase2C
    '                  execution of this proc by the LogOn screen (before the
    '                  gconAppActive object is instantiated).
    '--------------------------------------------------------------------------
    Const cstrCurrentProc       As String = "CommandSetSproc"

    On Error GoTo PROC_ERR

    ' Wipe out existing parameters and close the ActiveConnection
    CommandInitialize
    ' Reestablish a connection to the specified Connection or (default) gconAppActive
    CommandSetConnection pconIn

    With m_cmdADOCommand
        .CommandType = adCmdStoredProc
        .CommandText = strSprocName
    End With

    CommandSetSproc = True
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function


'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function Execute_SQL_AsRST(ByRef pconIn As cconConnection, ByVal strSQL As String) _
    As ADODB.Recordset
    '--------------------------------------------------------------------------
    ' Procedure:   Execute_SQL_AsRST
    '
    '              NOTE: This should **only** be used to apply selection criteria
    '                    to a "SELECT" statement, e.g., for reporting purposes.
    '                    Even then, that should really be made into a stored procedure
    '                    that accepts parameters so this function can be made
    '                    obsolete and the prospective caller can use
    '                    Execute_Sproc_AsRST( ) instead.
    '
    '                    !TODO! Create sprocs and change each caller of this proc
    '                           per the above.
    '
    ' Description: Executes the specified SQL to build a recordset
    ' Returns:     ADODB.Recordset that was built using strSQL
    ' Params:      pconIn  input/output)   The Connection object to use
    '              strSQL  input           What is to be executed against the DB
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc       As String = "Execute_SQL_AsRST"

    On Error GoTo PROC_ERR

    ' Check the Status on the Global Connection object to ensure that it is still open.
    If pconIn.ADOConn.State = adStateClosed Then
        '!TODO! Either do an app-specific error or just restablish the connection
        GoTo PROC_ERR
    End If

    ' Build the recordset, by executing the SQL statement
    Set Execute_SQL_AsRST = pconIn.ADOConn.Execute(strSQL)
PROC_EXIT:
    On Error GoTo 0     ' disable error handler

    ' Clean-up statements go here

    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If

    Exit Function
PROC_ERR:
    ' - - - - - - - - - - - Keep these in here!!! - - - - - - - - - - -
    Debug.Print "In " & mcstrName & cstrCurrentProc & ", this error was generated: " & vbCr & _
        Err.Number & " - " & Err.Description
    Debug.Print "    strSQL=" & strSQL
    ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
        
    fnFreeObject Execute_SQL_AsRST

    Select Case Err.Number
        'Case statements for expected errors go here
        Case -2147217865   ' Invalid object name
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_FERR_SQL_STMT_OBJECT_NOT_FOUND, _
                                       mcstrName & cstrCurrentProc, _
                                       strSQL
            Resume PROC_EXIT
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function Execute_SQL_UpdateableDisconnectedRST(ByRef pconIn As cconConnection, ByVal strSQL As String) _
    As ADODB.Recordset
    '--------------------------------------------------------------------------
    ' Procedure:   Execute_SQL_UpdateableDisconnectedRST
    '
    ' Description: Executes the specified SQL to build an updateable recordset
    '              and then disconnects it. This procedure should **ONLY**
    '              be used by the Select Reports screen, which has to "manufacture"
    '              data for some of the complex PPVUL reports using a
    '              recordset (typically created from a view) as merely a starting
    '              point.
    '
    ' Returns:     Disconnected ADODB.Recordset that was built using strSQL
    ' Params:      pconIn  input/output)   The Connection object to use
    '              strSQL  input           What is to be executed against the DB
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc       As String = "Execute_SQL_UpdateableDisconnectedRST"

    On Error GoTo PROC_ERR

    ' Check the Status on the Global Connection object to ensure that it is still open.
    If pconIn.ADOConn.State = adStateClosed Then
        '!TODO! Either do an app-specific error or just restablish the connection
        GoTo PROC_ERR
    End If

    Set Execute_SQL_UpdateableDisconnectedRST = New ADODB.Recordset
    
    With Execute_SQL_UpdateableDisconnectedRST
        .CursorLocation = adUseClient
        .Open strSQL, pconIn.ADOConn, adOpenDynamic, adLockOptimistic, adCmdText
        .ActiveConnection = Nothing '!TODO! Does this need a SET???
    End With
PROC_EXIT:
    On Error GoTo 0     ' disable error handler

    ' Clean-up statements go here

    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If

    Exit Function
PROC_ERR:
    ' - - - - - - - - - - - Keep these in here!!! - - - - - - - - - - -
    Debug.Print "In " & mcstrName & cstrCurrentProc & ", this error was generated: " & vbCr & _
        Err.Number & " - " & Err.Description
    Debug.Print "    strSQL=" & strSQL
    ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
        
    fnFreeObject Execute_SQL_UpdateableDisconnectedRST

    Select Case Err.Number
        'Case statements for expected errors go here
        Case -2147217865   ' Invalid object name
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_FERR_SQL_STMT_OBJECT_NOT_FOUND, _
                                       mcstrName & cstrCurrentProc, _
                                       strSQL
            Resume PROC_EXIT
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function GetMetaData_Columns(ByVal strTableName As String, _
    ByRef prstInOut As ADODB.Recordset) As Boolean
    '--------------------------------------------------------------------------
    ' Procedure:   GetMetaData_Columns
    ' Description: Use the OpenSchema method to get metadata about each
    '              table column. One inherent assumption is that the app's
    '              active and archive databases are based on an identical
    '              schema, hence we only need to ever look at the Active DB's
    '              meta data.
    ' Returns:     True if successful; False otherwise
    ' Params:      N/A
    ' Date:        01/07/2002
    '-----------------------------------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc       As String = "GetMetaData_Columns"
    Dim strLoggedOnDB As String

    If prstInOut Is Nothing Then
        Set prstInOut = New ADODB.Recordset
    End If

    strLoggedOnDB = gapsApp.ActiveDatabase(gconAppActive.LastLogonEnviron)

    Set prstInOut = gconAppActive.ADOConn.Execute("sp_columns '" & strTableName & "'")
    '2008 Update
    'Set prstInOut = gconAppActive.ADOConn.OpenSchema(adSchemaColumns, _
    '                    Array(strLoggedOnDB, Empty, strTableName, Empty))

    GetMetaData_Columns = True
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler

    ' Disconnect the Recordset
    If fnIsObject(prstInOut) Then
        fnFreeObject prstInOut.ActiveConnection
    End If

    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function GetMetaData_PrimaryKeys(ByVal strTableName As String, _
    ByRef prstInOut As ADODB.Recordset) As Boolean
    '--------------------------------------------------------------------------
    ' Procedure:   GetMetaData_PrimaryKeys
    ' Description: Use the OpenSchema method to get metadata about each
    '              table column. One inherent assumption is that the app's
    '              active and archive databases are based on an identical
    '              schema, hence we only need to ever look at the Active DB's
    '              meta data.
    ' Returns:     True if successful; False otherwise
    ' Params:      N/A
    ' Date:        01/07/2002
    '-----------------------------------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc       As String = "GetMetaData_PrimaryKeys"
    Dim strLoggedOnDB As String

    If prstInOut Is Nothing Then
        Set prstInOut = New ADODB.Recordset
    ElseIf prstInOut.State = adStateOpen Then
        prstInOut.Close
    End If

    strLoggedOnDB = gapsApp.ActiveDatabase(gconAppActive.LastLogonEnviron)

    'SQL 2008 Update
    Set prstInOut = gconAppActive.ADOConn.Execute("SELECT column_name " _
    & "From INFORMATION_SCHEMA.KEY_COLUMN_USAGE " _
    & "WHERE OBJECTPROPERTY(OBJECT_ID(constraint_name), 'IsPrimaryKey') = 1 " _
    & "AND table_name = '" & strTableName & "'")
    
    'Set prstInOut = gconAppActive.ADOConn.OpenSchema(adSchemaPrimaryKeys, _
    '                    Array(strLoggedOnDB, Empty, strTableName))

    GetMetaData_PrimaryKeys = True
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler

    ' Disconnect the Recordset
    If fnIsObject(prstInOut) Then
        fnFreeObject prstInOut.ActiveConnection
    End If

    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub MoveFirst(ByRef prstIn As ADODB.Recordset)
    '--------------------------------------------------------------------------
    ' Procedure:   MoveFirst
    ' Description: Repositions the specified recordset to its first record.
    '
    '              NOTE: The caller must be diligent about checking for BOF
    '                    and/or EOF being True upon returning from this proc
    '                    rather than assuming the recordset is positioned to
    '                    a valid record (or that the .RecordCount is necessarily > 0).
    '
    ' Returns:     N/A
    ' Params:      prstIn (input/output) Recordset to be repositioned
    '--------------------------------------------------------------------------
    Const cstrCurrentProc       As String = "MoveFirst"

    On Error GoTo PROC_ERR

    prstIn.MoveFirst
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        Case 3021
            ' Either BOF or EOF is True or the current record has been deleted. Requested
            ' operation requires a current record.
            ' Ignore these; the caller must be diligent about checking for BOF and/or EOF
            ' rather than assuming the recordset is positioned to a valid record (or that
            ' the .RecordCount is necessarily > 0.
            Resume Next
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub MoveLast(ByRef prstIn As ADODB.Recordset)
    '--------------------------------------------------------------------------
    ' Procedure:   MoveLast
    ' Description: Repositions the specified recordset to its last record
    '
    '              NOTE: The caller must be diligent about checking for BOF
    '                    and/or EOF being True upon returning from this proc
    '                    rather than assuming the recordset is positioned to
    '                    a valid record (or that the .RecordCount is necessarily > 0).
    '
    ' Returns:     N/A
    ' Params:      prstIn (input/output) Recordset to be repositioned
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc       As String = "MoveLast"

    On Error GoTo PROC_ERR

    prstIn.MoveLast
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        Case 3021
            ' Either BOF or EOF is True or the current record has been deleted. Requested
            ' operation requires a current record.
            ' Ignore these; the caller must be diligent about checking for BOF and/or EOF
            ' rather than assuming the recordset is positioned to a valid record (or that
            ' the .RecordCount is necessarily > 0.
            Resume Next
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub MoveNext(ByRef prstIn As ADODB.Recordset)
    '--------------------------------------------------------------------------
    ' Procedure:   MoveNext
    ' Description: Try to go to the next record; if doing so moves us past the
    '              last record, then back up to the last record.
    '
    '              NOTE: The caller must be diligent about checking for BOF
    '                    and/or EOF being True upon returning from this proc
    '                    rather than assuming the recordset is positioned to
    '                    a valid record (or that the .RecordCount is necessarily > 0).
    '
    ' Returns:     N/A
    ' Params:      prstIn (input/output) Recordset to be repositioned
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc   As String = "MoveNext"

    On Error GoTo PROC_ERR

    With prstIn
        .MoveNext
        ' Commented out -- no way to tell that we've processed all records, when looping through a rst
        'If .EOF Then
        '    .MoveLast
        'End If
    End With
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        Case 3021
            ' Either BOF or EOF is True or the current record has been deleted. Requested
            ' operation requires a current record.
            ' Ignore these; the caller must be diligent about checking for BOF and/or EOF
            ' rather than assuming the recordset is positioned to a valid record (or that
            ' the .RecordCount is necessarily > 0.
            Resume Next
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub MovePrev(ByRef prstIn As ADODB.Recordset)
    '--------------------------------------------------------------------------
    ' Procedure:   MovePrev
    ' Description: Try to go to the previous record; if doing so moves us past
    '              the first record, then back up to the first record.
    '
    '              NOTE: The caller must be diligent about checking for BOF
    '                    and/or EOF being True upon returning from this proc
    '                    rather than assuming the recordset is positioned to
    '                    a valid record (or that the .RecordCount is necessarily > 0).
    '
    ' Returns:     N/A
    ' Params:      prstIn (input/output) Recordset to be repositioned
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc   As String = "MovePrev"
    On Error GoTo PROC_ERR

    With prstIn
        .MovePrevious
        ' Commented out -- no way to tell that we've processed all records, when looping through a rst
        'If .BOF Then
        '    .MoveFirst
        'End If
    End With
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        Case 3021
            ' Either BOF or EOF is True or the current record has been deleted. Requested
            ' operation requires a current record.
            ' Ignore these; the caller must be diligent about checking for BOF and/or EOF
            ' rather than assuming the recordset is positioned to a valid record (or that
            ' the .RecordCount is necessarily > 0.
            Resume Next
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub





'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'\                                                                  /
'|                        PRIVATE  Procedures                       |
'/                                                                  \
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/

