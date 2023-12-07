VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cconConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "This class handles ADODB.Connection accesses  using ADO 2.5, including connecting/disconnecting to/from a database and handling transaction blocks."
'--------------------------------------------------------------------------
' Module     : cconConnection
' Description: Instantiated by modStartup.Sub_Main
'              This object WILL know the ErrorLogger.cls
'
' Procedures :
'    Private   Class_Initialize()
'    Private   Class_Terminate()
'    Public    Property Get ADOConn() As ADODB.Connection
'    Public    Property Get LastLogonEnviron() As String
'    Public    Property Get LastLogonIsSpecialUser() As Boolean
'    Public    Property Get LastLogonPassword() As String
'    Public    Property Get State() As ObjectStateEnum
'    Public    BeginTrans() As Boolean
'    Public    CommitTrans() as boolean
'    Public    Connect(ByVal strEnviron As String, ByVal strUserID As String, _
'                  ByVal strPassword As String) As Boolean
'    Public    Disconnect() As Boolean
'    Public    RollbackTrans()
'
' Revision History:
'    10/04/02 - Betsy - Cloned from TRS.
'--------------------------------------------------------------------------
Option Explicit
Option Compare Binary

#Const DEBUG_ERH = False
#Const DEBUG_RST = False

Private Const mcstrName             As String = "cconConnection."
Private m_Connection                As ADODB.Connection
Private m_LastLogonIsSpecialUser    As Boolean
Private m_LastLogonUserID           As String
Private m_LastLogonPassword         As String
Private m_LastLogonEnviron          As String
Private m_Initialized               As Boolean
Private m_IsTransactionActive       As Boolean




'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'\                                                                  /
'|         CLASS_INITIALIZE / CLASS_TERMINATE   Procedures         |
'/                                                                  \
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/

'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub Class_Initialize()
    ' **************************************************************************
    ' Function  : Class_Initialize
    ' Purpose   : Starting point for the object
    ' Parameters: N/A
    ' Returns   : True/False
    ' SXS 08/04/2004  Error4048   Added support for error 4048 to trap SQL error -2147217871 (Timeout exceeded).
'                     Also Also default ADO connection Timeout changed from 30 seconds to 90 seconds
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Class_Initialize"
    On Error GoTo PROC_ERR

    Set m_Connection = New ADODB.Connection
        
    'Error4048
    m_Connection.CommandTimeout = 90
    m_Initialized = True
    m_IsTransactionActive = False
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
    ' **************************************************************************
    ' Function  : Class_Terminate
    ' Purpose   : Close the object
    ' Parameters: N/A
    ' Returns   : N/A
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Class_Terminate"
    On Error GoTo PROC_ERR

    Disconnect
    fnFreeObject m_Connection

    m_Initialized = False
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
Public Property Get ADOConn() As ADODB.Connection
    ' **************************************************************************
    ' Function  : GetADOConn
    ' Purpose   :
    ' Parameters: N/A
    ' Returns   : True/False
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get ADOConn"
    On Error GoTo PROC_ERR

    Set ADOConn = m_Connection
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



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Get LastLogonEnviron() As String
    ' **************************************************************************
    ' Function  : LastLogonEnviron
    ' Purpose   :
    ' Parameters: N/A
    ' Returns   :
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get LastLogonEnviron"
    On Error GoTo PROC_ERR

    LastLogonEnviron = m_LastLogonEnviron
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



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Get LastLogonIsSpecialUser() As Boolean
    ' **************************************************************************
    ' Function  : LastLogonIsSpecialUser
    ' Purpose   :
    ' Parameters: N/A
    ' Returns   :
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get LastLogonIsSpecialUser"
    On Error GoTo PROC_ERR

    LastLogonIsSpecialUser = m_LastLogonIsSpecialUser
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



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Get LastLogonPassword() As String
    ' **************************************************************************
    ' Function  : LastLogonPassword
    ' Purpose   :
    ' Parameters: N/A
    ' Returns   :
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get LastLogonPassword"
    On Error GoTo PROC_ERR

    LastLogonPassword = m_LastLogonPassword
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



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Get LastLogOnUserID() As String
    ' **************************************************************************
    ' Function  : LastLogonUserID
    ' Purpose   :
    ' Parameters: N/A
    ' Returns   :
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get LastLogonUserID"
    On Error GoTo PROC_ERR

    LastLogOnUserID = m_LastLogonUserID
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



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Get State() As ObjectStateEnum
    ' **************************************************************************
    ' Function  : GetState
    ' Purpose   :
    ' Parameters: N/A
    ' Returns   : True/False
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get State"
    On Error GoTo PROC_ERR

    State = m_Connection.State
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
Public Function BeginTrans() As Boolean
    '--------------------------------------------------------------------------
    ' Procedure:   BeginTrans
    ' Description: Will start transaction processing on any complex database
    '              activity. This allows groups of transactions, such as might
    '              be done within a stored procedure, to be rolled back or
    '              committed as a group.
    ' Returns:     True if successful; False otherwise
    ' Params:      N/A
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc   As String = "BeginTrans"
    On Error GoTo PROC_ERR

    m_Connection.BeginTrans
    
    ' Set flag indicating we've started a transaction
    m_IsTransactionActive = True
    
    BeginTrans = True
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    BeginTrans = False

    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function CommitTrans() As Boolean
    '--------------------------------------------------------------------------
    ' Procedure:   CommitTrans
    ' Description: Will end transaction processing on any complex database
    '              activity and commit changes made since BeginTrans( ) was called.
    ' Returns:     True if successful; False otherwise
    ' Params:      N/A
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc   As String = "CommitTrans"

    On Error GoTo PROC_ERR

    If m_IsTransactionActive Then
        m_Connection.CommitTrans
        m_IsTransactionActive = False
    End If
    
    CommitTrans = True
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    CommitTrans = False

    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function Connect(ByVal strEnviron As String, _
                        ByVal strUserID As String, _
                        ByVal strPassword As String, _
                        Optional ByVal bActiveDB As Boolean = True) As Boolean
    ' **************************************************************************
    ' Function  : Connect
    ' Purpose   : Will connect to the specified environment based on the user info.
    ' Parameters: Parameters:
    '               strEnviron  (input) = User's selection from the Log On screen's
    '                                     Environment combo box
    '               strUserID   (input) = will contain the ACF2 of the user.
    '               strPassword (input) = related to UserID, password for that user, or Claims app.
    ' Returns   : True if successful; False otherwise
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Connect"
    Dim strConnectPart      As String

    On Error GoTo PROC_ERR

    If m_Connection.State = adStateOpen Then
        m_Connection.Close
    End If
    fnFreeObject m_Connection   ' SQL_INTEGRATED_SECURITY

    ' Make the Connection to the database
    With gapsApp
        If (LenB(.ActiveServer(strEnviron)) = 0) Or _
            (LenB(.ActiveDatabase(strEnviron)) = 0) Or _
            (LenB(.ArchiveDatabase(strEnviron)) = 0) Or _
            (LenB(.ArchiveServer(strEnviron)) = 0) Or _
            (LenB(.Port(strEnviron)) = 0) Then
                ' gcRES_NERR_ENV_REG_ENTRIES_MISSING (4010) = One or more registry entries
                '     that define how to connect to the selected
                '     Environment (@@1) are missing. Without all of these entries,
                '     the app cannot connect to the database.
                gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_ENV_REG_ENTRIES_MISSING, _
                    mcstrName & cstrCurrentProc, strEnviron
                    GoTo PROC_EXIT
        End If
    End With
    
    Set m_Connection = New ADODB.Connection ' SQL_INTEGRATED_SECURITY
    With m_Connection
        .Provider = "SQLOLEDB"
        .Mode = adModeReadWrite
        .CursorLocation = adUseClient
        .Properties("Prompt").value = adPromptNever
        ' OLE DB Services = -2 added to disable connection pooling. Having it enabled It causes problems
        ' with application roles. See <http://support.microsoft.com/search/preview.aspx?scid=kb;en-us;Q229564>
        ' for more info.
        .Properties("OLE DB Services").value = -2
        ' The following logic was replaced with building the full .ConnectionString property
        ' since VB Watch Profiler identified these lines as being inefficient. This noticeably sped
        ' up the Log On screen.
        '       If bActiveDB Then
        '            .Properties("Data Source").value = gapsApp.ActiveServer(strEnviron)
        '            .Properties("Initial Catalog").value = gapsApp.ActiveDatabase(strEnviron)
        '       Else
        '           .Properties("Data Source").value = gapsApp.ArchiveServer(strEnviron)
        '             .Properties("Initial Catalog").value = gapsApp.ArchiveDatabase(strEnviron)
        '       End If
        '       .Properties("User Id").value = strUserID
        '       .Properties("Password").value = strPassword

        'SQL_INTEGRATED_SECURITY
        Select Case gapsApp.UsesWindowsAuthentication(strEnviron)
            Case True
                strConnectPart = "Integrated Security=SSPI" & ";"
            Case Else
                strConnectPart = "Password=" & strPassword & ";User ID=" & strUserID & ";"
        End Select
        'SQL_INTEGRATED_SECURITY

        If bActiveDB Then
            .ConnectionString = strConnectPart & "Initial Catalog=" & gapsApp.ActiveDatabase(strEnviron) & _
                ";Data Source=" & gapsApp.ActiveServer(strEnviron) & "," & gapsApp.Port(strEnviron)
        Else
            .ConnectionString = strConnectPart & "Initial Catalog=" & gapsApp.ArchiveDatabase(strEnviron) & _
                ";Data Source=" & gapsApp.ArchiveServer(strEnviron) & "," & gapsApp.Port(strEnviron)
        End If
    End With

    m_Connection.Open

    If m_Connection.State <> adStateOpen Then
        gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_CONNECTION_FAILURE, _
                                   mcstrName & cstrCurrentProc, _
                                   strEnviron, m_Connection.State
    Else
        Connect = True

        If bActiveDB Then
            ' Save the ID, Password and Environment of the logged on user. This
            ' will be used, if necessary, to log the user on to the corresponding
            ' ArchiveDB
            m_LastLogonUserID = strUserID
            m_LastLogonPassword = strPassword
            m_LastLogonEnviron = strEnviron
            m_LastLogonIsSpecialUser = fnIsSpecialUser()
        End If
    End If
PROC_EXIT:
    On Error GoTo 0     ' disable error handler

    ' Clean-up statements go here

    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Connect = False
    Select Case Err.Number
        Case -2147467259, -2147217843, -2147467259
            ' -2147467259 = Cannot open database requested in login 'indppvul_pr'. Login fails.
            ' -2147217843 = Not a valid password'
            ' -2147467259 = SQL Server does not exist or access denied.
             gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_LOGON_FAILURE, _
                                           mcstrName & cstrCurrentProc, _
                                           Err.Number
        Case 3704 ' Operation is not allowed
            Resume Next
        Case Else
            ' Any error stemming from this procedure is manifested as
            ' a Log On failure that cites a bad User ID, Password or permissions
            ' as the likely cause. Therefore, any caller of this procedure
            ' doesn't have to check the return value; it will always be True
            ' since a False would be raised to the caller's error handler
            ' due to the propagation done in this proc's PROC_EXIT.
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_LOGON_FAILURE, _
                                       mcstrName & cstrCurrentProc, _
                                       Err.Number
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function Disconnect() As Boolean
    ' **************************************************************************
    ' Function  : Disconnect
    ' Purpose   : Will disconnect the ADO connection object.
    ' Parameters: N/A
    ' Returns   : True if successful; False otherwise
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Disconnect"
    On Error GoTo PROC_ERR

    If m_Connection.State = adStateClosed Then
        GoTo PROC_EXIT
    End If

    m_Connection.Close
    Disconnect = True
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Disconnect = False
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Function fnIsSpecialUser() As Boolean
    '--------------------------------------------------------------------------
    ' Procedure:   fnIsSpecialUser
    ' Description: This method determines whether the user is a member of
    '              the Support or UserAdmin roles. Membership in this roles
    '              indicate the user has access to restricted areas of the
    '              application (the Current Rate and State Rule screens).
    '
    ' Params:
    '
    ' Returns:     True if the user is a member of these roles; false otherwise.
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc            As String = "fnIsSpecialUser"
    Const clngSprocParamCount        As Long = 2                ' # of input or output params sproc expects
    Const cstrSproc                  As String = "sp_helpuser"  ' Stored procedure to execute
    Dim prmReturnValue               As ADODB.Parameter
    Dim prmName_in_DB                As ADODB.Parameter
    Dim rstTemp                      As ADODB.Recordset
    Dim adwTemp                      As New cadwADOWrapper
    Dim strGroupName                 As String

    On Error GoTo PROC_ERR
    
    ' Connect to the specified environment using the Dummy App ID,
    ' then execute the sp_helpuser sproc.
    If Not (adwTemp.CommandSetSproc(cstrSproc, Me)) Then
        GoTo PROC_EXIT
    End If
    
    With adwTemp.ADOCommand
        ' ---Parameter #1---
        ' Define the return value that represents the error code (i.e. reason) why
        ' the stored procedure failed.
        Set prmReturnValue = .CreateParameter(Name:="@return_value", _
                                              Type:=adInteger, _
                                              Direction:=adParamReturnValue, _
                                              value:=Null)
        .Parameters.Append prmReturnValue

        ' ---Parameter #2---
        ' Define the name_in_db input parameter, which represents the User ID being checked
        Set prmName_in_DB = .CreateParameter(Name:="@name_in_db", _
                                             Type:=adVarChar, _
                                             Direction:=adParamInput, _
                                             Size:=255, _
                                             value:=fnNullIfZLS(varIn:=m_LastLogonUserID, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmName_in_DB
            
        Set rstTemp = .Execute()
    End With
                
    With rstTemp
        If Not (.BOF And .EOF) Then
            .MoveFirst
            Do While Not .EOF
                strGroupName = UCase$(.Fields("GroupName").value)
                If strGroupName = "SUPPORT" Or strGroupName = "USERADMIN" Then
                    fnIsSpecialUser = True
                End If
                .MoveNext
            Loop
        End If
    End With
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    
    ' Clean-up statements go here
    fnFreeObject prmReturnValue
    fnFreeObject prmName_in_DB
    fnFreeRecordset rstTemp
    fnFreeObject adwTemp
    
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        Case vbObjectError + gcRES_NERR_LOGON_FAILURE ' 4013
            ' This environment will be considered "not authorized"
            Err.Clear
            gerhApp.Clear
            GoTo PROC_EXIT
            'Resume Next
        Case -2147217900 ' The name supplied (xxx) is not a user, role or aliased login
            ' This environment will be considered "not authorized"
            Err.Clear
            gerhApp.Clear
            GoTo PROC_EXIT
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub RollbackTrans()
    '--------------------------------------------------------------------------
    ' Procedure:   RollbackTrans
    ' Description: Will undo any DB changes made since BeginTrans( ) was called
    ' Params:      N/A
    ' Returns:     N/A
    ' Date:        01/09/2002
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc   As String = "RollbackTrans"

    ' NOTE:  NO ERROR HANDLER should be active here since we want
    '        the Rollback to proceed even if an error has been logged but not
    '        yet reported to the user!

    If m_IsTransactionActive Then
        m_Connection.RollbackTrans
        m_IsTransactionActive = False
    End If
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function SetAppRole(ByVal strRoleName As String, ByVal strRolePassword As String) As Boolean
    '--------------------------------------------------------------------------
    ' Procedure:   SetAppRole
    ' Description: Puts the specified App Role into effect, so that the role's
    '              permissions will override that of the logged on user.
    '              This proc should be called **after** the user has been logged on.
    ' Returns:     True if successful; False otherwise
    ' Params:      N/A
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc         As String = "SetAppRole"
    Const cstrSproc               As String = "sp_setapprole"  ' Stored procedure to execute
    Dim adwTemp                   As cadwADOWrapper
    Dim prmReturnValue            As ADODB.Parameter
    Dim prmRoleName               As ADODB.Parameter
    Dim prmPassword               As ADODB.Parameter
    Dim rstTemp                   As ADODB.Recordset

    On Error GoTo PROC_ERR
    
    Set adwTemp = New cadwADOWrapper
    
    ' Set the sproc name and set **this** connection object as the active connection
    If Not (adwTemp.CommandSetSproc(cstrSproc, Me)) Then
        GoTo PROC_EXIT
    End If
        
    With adwTemp.ADOCommand
        ' ---Parameter #1---
        ' Define the return value that represents the error code (i.e. reason) why
        ' the stored procedure failed.
        Set prmReturnValue = .CreateParameter(Name:="@return_value", _
                                              Type:=adInteger, _
                                              Direction:=adParamReturnValue, _
                                              value:=Null)
        .Parameters.Append prmReturnValue

        ' ---Parameter #2---
        ' Define the RoleName input parameter, which reflects *which*
        ' application role should be put into effect
        Set prmRoleName = .CreateParameter(Name:="@rolename", _
                                             Type:=adVarChar, _
                                             Direction:=adParamInput, _
                                             Size:=255, _
                                             value:=fnNullIfZLS(varIn:=strRoleName, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmRoleName
        
        ' ---Parameter #3---
        ' Define the Password input parameter, which reflects the password
        ' for the specified application role
        Set prmPassword = .CreateParameter(Name:="@password", _
                                             Type:=adVarChar, _
                                             Direction:=adParamInput, _
                                             Size:=255, _
                                             value:=fnNullIfZLS(varIn:=strRolePassword, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmPassword
        
        .Execute
    End With
                
    SetAppRole = True
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    
    ' Clean-up statements go here
    fnFreeObject adwTemp
    fnFreeObject prmReturnValue
    fnFreeObject prmRoleName
    fnFreeObject prmPassword
    fnFreeRecordset rstTemp
    
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    SetAppRole = False

    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function

