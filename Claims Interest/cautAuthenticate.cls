VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cautAuthenticate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'!TODO! Add support for determining whether user has access to restricted areas of app
'       (e.g., is a member of Support or UserAdmin roles). Make this a property that
'       frmLogOn will query and use to set a global variable that other forms
'       can reference.

'--------------------------------------------------------------------------
' Module     : cautAuthenticate
' Description: Instantiated by frmLogon to determine if the user
'              is authenticated to each possible environment's
'              application SQL Server database
'
' Procedures :
'    Public    AuthenticateAll(ByVal strUserID As String) As String()
'    Private   fnIsAuthorized(ByVal strUserID, ByVal strEnv As String) As Boolean
'
' Revision History: 1.0 BAW 05/16/02 Initial Creation
'--------------------------------------------------------------------------
Option Explicit
Option Compare Binary

#Const DEBUG_ERH = False
#Const DEBUG_RST = False

Private Const mcstrName             As String = "cautAuthenticate."



'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'\                                                                  /
'|                        PUBLIC  Procedures                        |
'/                                                                  \
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/

'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function AuthenticateEnvironments(ByVal strUserID As String) As String()
    '--------------------------------------------------------------------------
    ' Procedure:   AuthenticateEnvironments
    ' Description: Will call the SP_USER stored procedure for each SQL server/database
    '              listed in the Environments section of the registry. It will return
    '              an array of Environment Names that the Log On screen can use
    '              to populate the Environments combo box (so the user is only
    '              presented with a list of Environments for which she/he is
    '              authorized.
    '
    ' Called By:   cmdOK_Click of frmLogon
    '
    ' Params:
    '    strUserID (in) - the User ID under which user is known by SQL Server
    '
    ' Returns:     An array of strings representing the Environments for which
    '              the user is authorized.
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc            As String = "AuthenticateEnvironments"
    Dim astrAllEnvironments()        As String
    Dim astrAuthorizedEnvironments() As String
    Dim intAllIndex                  As Integer
    Dim intAuthIndex                 As Integer

    On Error GoTo PROC_ERR
    
    ' An error is generated during app startup if there are no
    ' environment names defined, hence we should always have at least one.
    astrAllEnvironments = gapsApp.EnvironmentNames()
    
    ' Resize our authorized Environments array based on the max number
    ' of possibly authorized environments. Shouldn't be much waste.
    ReDim astrAuthorizedEnvironments(UBound(astrAllEnvironments) + 1)
    
    For intAllIndex = LBound(astrAllEnvironments) To UBound(astrAllEnvironments)
        'SQL_INTEGRATED_SECURITY If fnIsAuthorized(strUserID, astrAllEnvironments(intAllIndex)) Then
            astrAuthorizedEnvironments(intAuthIndex) = astrAllEnvironments(intAllIndex)
            intAuthIndex = intAuthIndex + 1
        'SQL_INTEGRATED_SECURITY End If
    Next
    
    AuthenticateEnvironments = astrAuthorizedEnvironments
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
Public Function AuthenticateUser(ByVal strEnvironIn As String, ByVal strUserIDIn As String, _
    ByVal strPasswordIn As String, ByRef pconIn As cconConnection, _
    Optional ByVal bActiveDBIn = True) As Boolean
    '--------------------------------------------------------------------------
    ' Procedure:   AuthenticateUser
    ' Description: Logs the user on to the specified Environment, using the
    '              specified Connection object. From the frmLogon screen,
    '              this will be the global connection object (gconAppActive).
    '              From other screens, such as those that need to verify a
    '              record can be deleted without impacting ArchiveDB records,
    '              it is done with a local Connection object that is intended
    '              to point to the Archive DB.
    ' Params:      strUserID   (input)        - the UserID of the user
    '              strPassword (input)        - the Password of the user
    '              pconIn      (input/output) - a pointer to the cconConnection object
    '                                           to use for the logging on
    '              strEnviron  (input)        - the name of the Environment
    '                                           to log on. (This may come from the
    '                                           Environments combobox on the Log On
    '                                           screen or it may be the .LoggedOnEnviron
    '                                           property of the cconConnection object)
    '              bArchive    (input)        - True to log onto the Archive DB; False
    '                                           to log onto the Active DB
    '
    ' Called By:   cmdOK_Click of frmLogon, to log on to the Active DB
    '              Other screens, when validations or processing must be
    '                     done against the Archive DB
    '
    ' Returns:     True if the logon was successful; False otherwise
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc          As String = "AuthenticateUser"
    Const cstrAppRoleClaims_UserId    As String = "AppRoleClaims"
    Const cstrAppRoleClaims_Password  As String = "claims"
    
    On Error GoTo PROC_ERR

    pconIn.Connect strEnviron:=strEnvironIn, _
                   strUserID:=strUserIDIn, _
                   strPassword:=strPasswordIn, _
                   bActiveDB:=bActiveDBIn
    
    ' NOTE: All sprocs must be DBO-owned if the application role is put into effect. This means
    '       that the following line may need to be commented out at times while application
    '       development is going on.
    pconIn.SetAppRole strRoleName:=cstrAppRoleClaims_UserId, strRolePassword:=cstrAppRoleClaims_Password
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



'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'\                                                                  /
'|                       PRIVATE  Procedures                        |
'/                                                                  \
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/


'////////////////////////////////////////////////////////////////////////////////////////////////
Private Function fnIsAuthorized(ByVal strUserID As String, ByVal strEnvironment As String) As Boolean
    '--------------------------------------------------------------------------
    ' Procedure:   fnIsAuthorized
    ' Description: This method uses the "sp_helpuser" system stored
    '              procedure to determine whether the specified
    '              UserID is an authorized user of the database
    '              referenced by the specified Environment.
    '
    ' Params:
    '    strUserID      (in) - the User ID under which user is known by SQL Server
    '    strEnvironment (in) - the Environment Name (as established by
    '                          the gapsApp AppSettings object.
    '
    ' Returns:     True if the user is authorized; false otherwise.
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc            As String = "fnIsAuthorized"
    
    ' NOTE: The following Dummy App ID/Password must be set up on every server
    '       on which the application will run
    Const cstrDummyApp_UserId        As String = "CLAIMAPP"     ' "claimapp"
    Const cstrDummyApp_Password      As String = "CLAIMAPP"     ' "claimapp"
    
    Const clngSprocParamCount        As Long = 2                ' # of input or output params sproc expects
    Const cstrSproc                  As String = "sp_helpuser"  ' Stored procedure to execute
    Dim prmReturnValue               As ADODB.Parameter
    Dim prmName_in_DB                As ADODB.Parameter
    Dim rstTemp                      As ADODB.Recordset
    Dim adwTemp                      As New cadwADOWrapper
    Dim conTemp                      As New cconConnection

    On Error GoTo PROC_ERR
    
    ' Connect to the specified environment using the Dummy App ID,
    ' then execute the sp_helpuser sproc.
    If conTemp.Connect(strEnvironment, cstrDummyApp_UserId, cstrDummyApp_Password) Then
        
        If Not (adwTemp.CommandSetSproc(cstrSproc, conTemp)) Then
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
                                                 value:=fnNullIfZLS(varIn:=strUserID, bHandleEmbeddedQuotes:=True))
            .Parameters.Append prmName_in_DB

            Set rstTemp = .Execute()
        End With
                
        If rstTemp.RecordCount > 0 Then
            fnIsAuthorized = True
        End If
    End If
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    
    ' Clean-up statements go here
    fnFreeObject prmReturnValue
    fnFreeObject prmName_in_DB
    fnFreeRecordset rstTemp
    If fnIsObject(conTemp) Then
        conTemp.Disconnect
        fnFreeObject conTemp
    End If
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
