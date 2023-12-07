VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "capsAppSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "This class stores and retrieves app-specific settings from the Windows registry."
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"No"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'
' =-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-=
' Class       : capsAppSettings
' Description : Simplified registry access routines used
'               for saving AppBranch settings
' Source      : Total Visual SourceBook 2000
'
' Procedures  :
'   Private     Class_Initialize()
'   Private     Class_Terminate()
'   Public      Property Get ActiveDatabase(strEnvName As String) As String
'   Public      Property Get ActiveServer(strEnvName As String) As String
'   Public      Property Get AppBranch() As String
'   Public      Property Let AppBranch(ByVal strValue As String)
'   Public      Property Get ArchiveDatabase(strEnvName As String) As String
'   Public      Property Get ArchiveServer(strEnvName As String) As String
'   Public      Property Get CompanyBranch() As String
'   Public      Property Let CompanyBranch(ByVal strValue As String)
'   Public      Property Get EnvironmentKey(strEnvName As String) As String
'   Public      Property Get EnvironmentNames() As String()
'   Public      Property Get LastLogOnUserID() As String
'   Public      Property Let LastLogOnUserID(ByVal strValue as String)
'   Public      Property Get LastLogOnUserPassword() As String
'   Public      Property Let LastLogOnUserPassword(ByVal strValue as String)
'   Public      Property Get MainBranch() As String
'   Public      Property Let MainBranch(strValue As String)
'   Public      Property Get Port(ByVal strEnvName As String) As String
'   Public      Public Property Get TaxFileFolder() As String
'   Public      Public Property Let TaxFileFolder(ByVal strValue As String)
'   Public      fnLoadEnvironments()
'   Private     fnShellSortAny(ByRef varArray As Variant, ByVal lngNbrOfElements As Long, _
'                  Optional ByVal bSortDescending As Boolean = False)
'   Public      LoadCbo_EnvironmentNames(ByRef cboIn As ComboBox)
'   Public      ReadEntry(ByVal eHKRoot As EnumRegistryRootKeys, ByVal strRegKey As String, _
'                         ByVal strEntry As String, Optional ByVal strDefault As String = "EMPTY") As String
'   Public      RestoreForm(ByRef frmIn As Form, Optional ByVal bForceVisible As Boolean = False) As Boolean
'   Public      SaveForm(ByRef frmIn As Form)
'   Public      WriteEntry(ByVal eHKRoot As EnumRegistryRootKeys, ByVal strRegKey As String, _
'                         ByVal strEntry As String, ByVal strValue As String)
'
' Modified:
'
'   Date     Who   What
'   -------- ---   -------------------------------------------------------------------
'   10/04/02 BAW   Cloned from TRS.
'   06/22/03 BAW   Added Port and support for same, since the new Prod environment under
'                  SQL Server 2000 utilizes nonstandard Port assignments and hence it must
'                  be specified in the ADO .ConnectionString property.
' =-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-=
Option Explicit
Option Compare Binary
Private Const mcstrName As String = "capsAppSettings."

' Local variables to hold Public Property values
Private m_strMainBranch As String
Private m_strCompanyBranch As String
Private m_strAppBranch As String
Private m_strVersionBranch As String
Private m_strLastLogOnUserID As String
Private m_strLastLogOnPassword As String
Private m_strLastLogOnEnvironment As String
Private m_audtEnvironments() As udtEnvironment
Private m_strTaxFileFolder As String


Private Const mcstrSlash As String = "\"
Private Const mcstrEMPTY As String = "EMPTY"

' Sections in the registry
Private Const mcstrScreenPreferences As String = "ScreenPreferences"

 ' Keys or Key Values in the registry
Private Const mcstrEnvironments                     As String = "Environments"
Private Const mcstrEnv_ActiveDatabase               As String = "ActiveDB"
Private Const mcstrEnv_ActiveServer                 As String = "ActiveServer"
Private Const mcstrEnv_ArchiveDatabase              As String = "ArchiveDB"
Private Const mcstrEnv_ArchiveServer                As String = "ArchiveServer"
Private Const mcstrEnv_EnvironmentName              As String = "EnvName"
Private Const mcstrLastLogOnUserID                  As String = "LastLogOnUserID"
Private Const mcstrEnv_Port                         As String = "Port"
Private Const mcstrEnv_UsesWindowsAuthentication    As String = "UsesWindowsAuthentication" 'SQL_INTEGRATED_SECURITY
Private Const mcstrTaxFileFolder                    As String = "TaxFileFolder"


Private Type udtEnvironment
    ActiveDatabase              As String
    ActiveServer                As String
    ArchiveDatabase             As String
    ArchiveServer               As String
    Port                        As String
    UsesWindowsAuthentication   As Boolean  'SQL_INTEGRATED_SECURITY
    EnvironmentName             As String
    EnvironmentKey              As String
End Type




'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'\                                                                  /
'|         CLASS_INITIALIZE / CLASS_TERMINATE   Procedures         |
'/                                                                  \
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/

'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub Class_Initialize()
    ' Set initial values to defaults which may be overridden with property settings
    '
    ' /\/\/\/\/\/\/\/\/\/\/\/\  WARNING /\/\/\/\/\/\/\/\/\/\/\/\
    ' Use the fnConstructor_cerhErrorHandler( ) procedure in
    ' modConstructors.bas to instantiate this object !!

    ' See that procedure for additional comments as to why this is
    ' necessary
    ' /\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
    '
    Const cstrCurrentProc As String = "Class_Initialize"
    On Error GoTo 0

    MainBranch = "SOFTWARE"
    With App
        AppBranch = .Title
        VersionBranch = CStr(.Major & gcstrDOT & .Minor & gcstrDOT & .Revision)
        CompanyBranch = "Sun Life Financial"
        LastLogonPassword = vbNullString
    End With

    ' Remainder of the intialization is done by
    ' fnConstructor_cerhErrorHandler( ) in modConstructors.bas.
End Sub


Private Sub Class_Terminate()
    ' Free up resources allocated in this class
    Const cstrCurrentProc As String = "Class_Terminate"
    On Error GoTo PROC_ERR

    Erase m_audtEnvironments
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
Public Property Get ActiveDatabase(ByVal strEnvName As String) As String
    ' Returns the active database name associated with the specified Environment
    Const cstrCurrentProc As String = "Get ActiveDatabase"
    Dim lngIndex As Long
    On Error GoTo PROC_ERR

    For lngIndex = LBound(m_audtEnvironments) To UBound(m_audtEnvironments)
        If m_audtEnvironments(lngIndex).EnvironmentName = strEnvName Then
            ActiveDatabase = m_audtEnvironments(lngIndex).ActiveDatabase
        End If
    Next
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
Public Property Get ActiveServer(ByVal strEnvName As String) As String
    ' Returns the active Server name associated with the specified Environment
    Const cstrCurrentProc As String = "Get ActiveServer"
    Dim lngIndex As Long
    On Error GoTo PROC_ERR

    For lngIndex = LBound(m_audtEnvironments) To UBound(m_audtEnvironments)
        If m_audtEnvironments(lngIndex).EnvironmentName = strEnvName Then
            ActiveServer = m_audtEnvironments(lngIndex).ActiveServer
        End If
    Next
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
Public Property Get AppBranch() As String
    ' Returns the current value of the AppBranch property
    Const cstrCurrentProc As String = "Get AppBranch"
    On Error GoTo PROC_ERR
    AppBranch = m_strAppBranch
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
End Property


'////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Let AppBranch(ByVal strValue As String)
    ' Sets the AppBranch property to the value specified by strValue
    Const cstrCurrentProc As String = "Let AppBranch"
    On Error GoTo PROC_ERR

    If Right$(strValue, 1) = mcstrSlash Then
        strValue = Left$(strValue, Len(strValue) - 1)
    End If

    m_strAppBranch = strValue
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
Public Property Get ArchiveDatabase(ByVal strEnvName As String) As String
    ' Returns the archive database name associated with the specified Environment
    Const cstrCurrentProc As String = "Get ArchiveDatabase"
    Dim lngIndex As Long
    On Error GoTo PROC_ERR

    For lngIndex = LBound(m_audtEnvironments) To UBound(m_audtEnvironments)
        If m_audtEnvironments(lngIndex).EnvironmentName = strEnvName Then
            ArchiveDatabase = m_audtEnvironments(lngIndex).ArchiveDatabase
        End If
    Next
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
Public Property Get ArchiveServer(ByVal strEnvName As String) As String
    ' Returns the archive Server name associated with the specified Environment
    Const cstrCurrentProc As String = "Get ArchiveServer"
    Dim lngIndex As Long
    On Error GoTo PROC_ERR

    For lngIndex = LBound(m_audtEnvironments) To UBound(m_audtEnvironments)
        If m_audtEnvironments(lngIndex).EnvironmentName = strEnvName Then
            ArchiveServer = m_audtEnvironments(lngIndex).ArchiveServer
        End If
    Next
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
Public Property Get CompanyBranch() As String
    ' Returns the current value of CompanyBranch
    Const cstrCurrentProc As String = "Get CompanyBranch"
    On Error GoTo PROC_ERR

    CompanyBranch = m_strCompanyBranch
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
Public Property Let CompanyBranch(ByVal strValue As String)
    ' Sets the CompanyBranch property to the value specified by strValue
    Const cstrCurrentProc As String = "Let CompanyBranch"
    On Error GoTo PROC_ERR

    If Right$(strValue, 1) = mcstrSlash Then
        strValue = Left$(strValue, Len(strValue) - 1)
    End If

    m_strCompanyBranch = strValue
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
Public Property Get EnvironmentNames() As String()
    ' Comments:   Retrieve an array of database Environment Names
    ' Parameters: N/A
    ' Returns:    An array of strings that list the environment name for each
    '             environment defined in the registry
    ' Called by : No one yet  :(
    Const cstrCurrentProc As String = "Get EnvironmentName"
    Dim lngIndex As Long

    On Error GoTo PROC_ERR

    ReDim astrNames(UBound(m_audtEnvironments)) As String

    For lngIndex = LBound(m_audtEnvironments) To (UBound(m_audtEnvironments))
        astrNames(lngIndex) = m_audtEnvironments(lngIndex).EnvironmentName
    Next

    ' Don't sort, since this will influence the order in which they appear...We want it sorted by
    ' KEY not Name, e.g., "1_Prd2A_COLI" versus "COLI Development Phase 2A"
    '       Commented out:  fnShellSortAny astrNames, UBound(astrNames) - 1, False

    EnvironmentNames = astrNames
PROC_EXIT:
    On Error GoTo 0     ' disable error handler

    ' Clean-up statements go here
    Erase astrNames

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
Public Property Get LastLogonEnvironment() As String
    ' Comments:   This property is NOT persisted across sessions. It is saved merely to allow
    '             the status bar text to be updated in its entirety whenever the MDI Form
    '             is activated or the user logs in.
    ' Parameters: N/A
    ' Returns:    The current value of LastLogOnEnvironment.
    ' Called by :
    Const cstrCurrentProc As String = "Get LastLogOnEnvironment"

    On Error GoTo PROC_ERR

    LastLogonEnvironment = m_strLastLogOnEnvironment
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
Public Property Let LastLogonEnvironment(ByVal strValue As String)
    ' Comments:   Set the value of the LastLogOnEnvironment property (called after a successful logon)
    ' Parameters: N/A
    ' Returns:    The current value of LastLogOnEnvironment. Its default value is a null string ("").
    ' Called by : cmdOK_Click() of frmLogOn
    Const cstrCurrentProc As String = "Let LastLogOnEnvironment"

    On Error GoTo PROC_ERR

    m_strLastLogOnEnvironment = strValue
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
Public Property Let LastLogonPassword(ByVal strValue As String)
    ' Comments:   Set the value of the LastLogOnPassword property (called after a successful logon)
    ' Parameters: N/A
    ' Returns:    The current value of LastLogOnPassword. Its default value is a null string ("").
    ' Called by : cmdOK_Click() of frmLogOn
    Const cstrCurrentProc As String = "Let LastLogOnPassword"

    On Error GoTo PROC_ERR

    m_strLastLogOnPassword = strValue
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
    ' Comments:   Retrieve the ACF2 under which the current user last logged on
    ' Parameters: N/A
    ' Returns:    The current value of LastLogOnUserID. Its default value is a null string ("").
    ' Called by : fnInit_gerhApp of modConstructors
    '             Form_Load() of frmLogOn
    Const cstrCurrentProc As String = "Get LastLogOnUserID"
    Dim strRegKey As String

    On Error GoTo PROC_ERR

    If m_strLastLogOnUserID = vbNullString Then
        strRegKey = m_strMainBranch & mcstrSlash & _
                    m_strCompanyBranch & mcstrSlash & _
                    m_strAppBranch
        m_strLastLogOnUserID = ReadEntry(rrkHKEY_CURRENT_USER, strRegKey, mcstrLastLogOnUserID, vbNullString)
        LastLogOnUserID = m_strLastLogOnUserID
    Else
        LastLogOnUserID = m_strLastLogOnUserID
    End If
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
Public Property Let LastLogOnUserID(ByVal strValue As String)
    ' Comments:   Set the value of the LastLogOnUserID property (called after a successful logon)
    ' Parameters: N/A
    ' Returns:    The current value of LastLogOnUserID. Its default value is a null string ("").
    ' Called by : cmdOK_Click() of frmLogOn
    Const cstrCurrentProc As String = "Let LastLogOnUserID"
    Dim strRegKey As String

    On Error GoTo PROC_ERR

    strRegKey = m_strMainBranch & mcstrSlash & _
                m_strCompanyBranch & mcstrSlash & _
                m_strAppBranch

    m_strLastLogOnUserID = strValue
    WriteEntry rrkHKEY_CURRENT_USER, strRegKey, mcstrLastLogOnUserID, strValue
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
Private Property Let MainBranch(ByVal strValue As String)
    ' Sets the MainBranch property to the value specified by strValue
    Const cstrCurrentProc As String = "Let MainBranch"
    On Error GoTo PROC_ERR

    If Right$(strValue, 1) = mcstrSlash Then
        strValue = Left$(strValue, Len(strValue) - 1)
    End If
    m_strMainBranch = strValue
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
Public Property Get Port(ByVal strEnvName As String) As String
    ' Returns the Port number associated with the specified Environment
    Const cstrCurrentProc As String = "Get Port"
    Dim lngIndex As Long
    On Error GoTo PROC_ERR

    For lngIndex = LBound(m_audtEnvironments) To UBound(m_audtEnvironments)
        If m_audtEnvironments(lngIndex).EnvironmentName = strEnvName Then
            Port = m_audtEnvironments(lngIndex).Port
        End If
    Next
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
Public Property Get TaxFileFolder() As String
    ' Comments:   Retrieve the ACF2 under which the current user last logged on
    ' Parameters: N/A
    ' Returns:    The current value of TaxFileFolder. Its default value is the user's
    '             My Documents folder.
    ' Called by : fnInit_gerhApp of modConstructors
    '             Form_Load() of frmGenerateTaxFile
    Const cstrCurrentProc                 As String = "Get TaxFileFolder"
    Dim fso                               As New Scripting.FileSystemObject
    Dim strPerUserTaxFileFolder           As String
    Dim strRegKey                         As String

    On Error GoTo PROC_ERR

    ' If the TaxFile Folder is empty or merely contains a slash...
    If m_strTaxFileFolder = vbNullString Then
        strRegKey = m_strMainBranch & mcstrSlash & _
                    m_strCompanyBranch & mcstrSlash & _
                    m_strAppBranch
        ' 1. Is a Per User TaxFile folder defined?
        '    a. If so, use it.
        '    b. If not, then use the Per User Non-Roaming folder (the user's
        '       My Documents folder)
        strPerUserTaxFileFolder = ReadEntry(rrkHKEY_CURRENT_USER, strRegKey, mcstrTaxFileFolder, vbNullString)
        If strPerUserTaxFileFolder <> vbNullString Then
            If gbLogVerbose Then
                fnLogWrite "TaxFileFolder defined in HKCU is: " & strPerUserTaxFileFolder, cstrCurrentProc
            End If
        Else
            ' Get the path to where per user non-roaming data is stored. This path
            ' will be created if it doesn't already exist.
            strPerUserTaxFileFolder = fnGetSpecialFolder(0, CSIDL_PERSONAL Or CSIDL_FLAG_CREATE)
            If gbLogVerbose Then
                fnLogWrite "TaxFileFolder defaulting to user's My Documents folder: " & strPerUserTaxFileFolder, cstrCurrentProc
            End If
        End If
            
        ' If directory doesn't exist, create it
        If Not fso.FolderExists(strPerUserTaxFileFolder) Then
            ' If an error occurs during the create, ignore for now. frmGenerateTaxFile's fnValidData procedure
            ' will generate an error if the folder doesn't exist by the time a generate is kicked off.
            On Error Resume Next
            fso.CreateFolder strPerUserTaxFileFolder
            On Error GoTo PROC_EXIT
        End If

        ' Using Property Let, save to Per User (HKCU). This will also add a trailing slash
        ' if one doesn't already exist.
        Me.TaxFileFolder = strPerUserTaxFileFolder
    End If
    
    TaxFileFolder = m_strTaxFileFolder
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    
    ' Clean-up statements go here
    fnFreeObject fso
    
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
Public Property Let TaxFileFolder(ByVal strValue As String)
    ' Comments:   Set the value of the TaxFileFolder property
    ' Parameters: N/A
    ' Returns:    The current value of TaxFileFolder. Its default value is the user's My Documents
    '             folder (the location for Per User Non-Roaming data).
    ' Called by : cmdOK_Click() of frmGenerateTaxFile
    Const cstrCurrentProc As String = "Let TaxFileFolder"
    Dim strRegKey As String

    On Error GoTo PROC_ERR

    strRegKey = m_strMainBranch & mcstrSlash & _
                m_strCompanyBranch & mcstrSlash & _
                m_strAppBranch

    m_strTaxFileFolder = UCase$(fnAddBackslash(Trim$(strValue)))
    WriteEntry rrkHKEY_CURRENT_USER, strRegKey, mcstrTaxFileFolder, strValue
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



'SQL_INTEGRATED_SECURITY - Added
'////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Get UsesWindowsAuthentication(ByVal strEnvName As String) As Boolean
    ' Returns True if the SQL Server of the specified Environment used windows authentication
    ' rather than SQL Server authentication; False otherwise.
    Const cstrCurrentProc As String = "Get UsesWindowsAuthentication"
    Dim lngIndex As Long
    On Error GoTo PROC_ERR

    For lngIndex = LBound(m_audtEnvironments) To UBound(m_audtEnvironments)
        If m_audtEnvironments(lngIndex).EnvironmentName = strEnvName Then
            UsesWindowsAuthentication = m_audtEnvironments(lngIndex).UsesWindowsAuthentication
        End If
    Next
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
'SQL_INTEGRATED_SECURITY - Added


'////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Get VersionBranch() As String
    ' Returns the current value of the VersionBranch property
    Const cstrCurrentProc As String = "Get VersionBranch"
    On Error GoTo PROC_ERR
    VersionBranch = m_strVersionBranch
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
Public Property Let VersionBranch(ByVal strValue As String)
    ' Sets the VersionBranch property to the value specified by strValue
    Const cstrCurrentProc   As String = "Let VersionBranch"
    On Error GoTo PROC_ERR

    m_strVersionBranch = strValue
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


'SQL_INTEGRATED_SECURITY
'////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub LoadCbo_EnvironmentNames(ByRef cboIn As ComboBox)
    ' Comments:   Populates the combobox parameter with a list of Environment Names
    ' Parameters: cboIn   (in)   combobox to populate
    ' Returns:    N/A
    ' Called by : Form_Load() of frmLogOn
    Const cstrCurrentProc As String = "LoadCbo_EnvironmentNames"
    Dim lngIndex As Long

    On Error GoTo PROC_ERR

    For lngIndex = LBound(m_audtEnvironments) To UBound(m_audtEnvironments)
        cboIn.AddItem m_audtEnvironments(lngIndex).EnvironmentName
    Next
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
'SQL_INTEGRATED_SECURITY


'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function ReadEntry(ByVal eHKRoot As EnumRegistryRootKeys, ByVal strRegKey As String, _
                          ByVal strEntry As String, Optional ByVal strDefault As String = mcstrEMPTY) As String
    ' Comments  : Reads a string value from the location in the
    '             registry specified by the class properties and/or parameters
    ' Parameters: eHKRoot    - (input) the hive from which to read a value, e.g., HKLM, HKCU, etc.
    '             strRegKey  - (input) the path under eHKRoot under which to find the key containing strEntry
    '             strEntry   - (input) The value under EHKRoot\strRegKey to retrieve
    '             strDefault - (input) The default value to return if strEntry isn't found
    ' Returns   : Either the registry value, or the default value
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "ReadEntry"
    Dim strValue As String

    On Error GoTo PROC_ERR

    strValue = RegistryGetKeyValue(eHKRoot, strRegKey, strEntry)

    If strValue = vbNullString Then
        ReadEntry = strDefault
    Else
        ReadEntry = strValue
    End If
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
Public Function RestoreForm(ByRef frmIn As Form, Optional ByVal bForceVisible As Boolean = True) As Boolean
    ' Comments  : Restores the form to its last-saved position (stored
    '             by the SaveForm method).
    ' Parameters:
    '             frmIn         - object pointer to the Form object whose size and position should
    '                             be restored
    '             bForceVisible - if true, and if restoring form to its previous location would
    '                             result in the form being off-screen, then put the form in the
    '                             upper-left-hand corner of the screen.
    ' Returns   : True if the form was restored; False if it was not restored
    '
    ' Called by : All forms' Form_Load( ) procedure
    '
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "RestoreForm"
    Dim varCurrent As Variant
    Dim lngPos As Long
    Dim lngLocation As Long
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim lngWindowState As Long
    Dim strRegKey As String

    On Error GoTo PROC_ERR

    strRegKey = m_strMainBranch & mcstrSlash & _
                m_strCompanyBranch & mcstrSlash & _
                m_strAppBranch & mcstrSlash & _
                "4.0.0" & mcstrSlash & _
                mcstrScreenPreferences
'    strRegKey = m_strMainBranch & mcstrSlash & _
'                m_strCompanyBranch & mcstrSlash & _
'                m_strAppBranch & mcstrSlash & _
'                m_strVersionBranch & mcstrSlash & _
'                mcstrScreenPreferences

    varCurrent = RegistryGetKeyValue(rrkHKEY_CURRENT_USER, strRegKey, frmIn.Name)

    If (varCurrent & vbNullString) = vbNullString Then
        varCurrent = vbNullString
    End If

    ' Test for missing value (Indicates that form's position was not previously saved)
    If LenB(varCurrent) = 0 Then
        RestoreForm = False
    Else
        ' Find saved Left position
        lngPos = InStr(varCurrent, "L=")
        If lngPos Then
            lngLocation = Val(Mid$(varCurrent, lngPos + 2))
            lngLeft = lngLocation
        End If

        ' Find saved Top position
        lngPos = InStr(varCurrent, "T=")
        If lngPos Then
            lngLocation = Val(Mid$(varCurrent, lngPos + 2))
            lngTop = lngLocation
        End If

        ' Find saved Width
        lngPos = InStr(varCurrent, "W=")
        If lngPos Then
            lngLocation = Val(Mid$(varCurrent, lngPos + 2))
            lngWidth = lngLocation
        End If

        ' Find saved Height
        lngPos = InStr(varCurrent, "H=")
        If lngPos Then
            lngLocation = Val(Mid$(varCurrent, lngPos + 2))
            lngHeight = lngLocation
        End If

        ' Find saved WindowState value (minimized, maximized, normal)
        lngPos = InStr(varCurrent, "S=")
        If lngPos Then
            lngLocation = Val(Mid$(varCurrent, lngPos + 2))
            lngWindowState = lngLocation
        End If

        ' If form was saved minimized or maximized, change the state only
        If lngWindowState = vbMinimized Or lngWindowState = vbMaximized Then
            frmIn.WindowState = lngWindowState
        Else
            If bForceVisible Then
                If (lngLeft >= Screen.Width) Then
                    lngLeft = Screen.Width - lngWidth
                End If

                If (lngTop >= Screen.Height) Then
                    lngTop = Screen.Height - lngHeight
                End If

                If (lngTop < 0) Then
                    lngTop = 0
                End If

                If (lngLeft < 0) Then
                    lngLeft = 0
                End If
            End If

            ' Move the form in a single statement instead of setting the
            ' properties individually
            frmIn.Move lngLeft, lngTop, lngWidth, lngHeight
        End If
        RestoreForm = True
    End If
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
Public Sub SaveForm(ByRef frmIn As Form)
    ' Comments  : Saves the current size and position of the named form so that it can be
    '             subsequently restored via the RestoreForm method.
    ' Parameters:
    '             frmIn         - object pointer to the Form object whose size and position should
    '                             be saved
    ' Returns   : nothing
    '
    ' Called by : All forms' Form_Unload( ) procedure
    '
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "SaveForm"
    Dim strNewValue As String
    Dim strRegKey As String

    On Error GoTo PROC_ERR

    strRegKey = m_strMainBranch & mcstrSlash & _
                m_strCompanyBranch & mcstrSlash & _
                m_strAppBranch & mcstrSlash & _
                m_strVersionBranch & mcstrSlash & _
                mcstrScreenPreferences

    With frmIn
        strNewValue = "L=" & .Left & ";T=" & .Top & ";W=" & .Width & ";H=" & .Height & ";S=" & .WindowState & ";"
        RegistrySetKeyValue rrkHKEY_CURRENT_USER, strRegKey, .Name, strNewValue, rrkRegSZ
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
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub WriteEntry(ByVal eHKRoot As EnumRegistryRootKeys, ByVal strRegKey As String, _
                      ByVal strEntry As String, ByVal strValue As String)
    ' Comments  : Writes a string entry into the registry
    '             at the location specified by the class properties and/or parameters
    ' Parameters: eHKRoot    - the hive to which to write a value, e.g., HKLM, HKCU, etc.
    '             strRegKey  - the path under eHKRoot under which to write the strEntry key and its
    '                          strValue value
    '             strEntry   - The key under eHKRoot\strRegKey to which to write the strValue value
    '             strValue   - The string value under eHKRoot\strRegKey\strEntry to write
    '
    ' Returns   : Nothing
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "WriteEntry"
    On Error GoTo PROC_ERR

    RegistrySetKeyValue eHKRoot, strRegKey, strEntry, strValue, rrkRegSZ
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
'|                        FRIEND Procedures                         |
'/                                                                  \
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/


'////////////////////////////////////////////////////////////////////////////////////////////////
Friend Sub fnLoadEnvironments()
    ' Comments  : Retrieves all Environment-related registry key values and loads
    '             them to an array member variable of type udtEnvironments
    ' Parameters: N/A
    ' Returns   : N/A
    ' Called by : fnInit_gerhApp of modConstructors
    Const cstrCurrentProc As String = "fnLoadEnvironments"
    Dim astrEnvKeys() As String
    Dim lngEnvs As Long
    Dim lngNbrOfEnvs As Long
    Dim strRegKey As String
    Dim strTempKey As String

    On Error GoTo PROC_ERR

    strRegKey = m_strMainBranch & mcstrSlash & _
                m_strCompanyBranch & mcstrSlash & _
                m_strAppBranch & mcstrSlash & _
                "4.0.0" & mcstrSlash & _
                mcstrEnvironments
'    strRegKey = m_strMainBranch & mcstrSlash & _
'                m_strCompanyBranch & mcstrSlash & _
'                m_strAppBranch & mcstrSlash & _
'                m_strVersionBranch & mcstrSlash & _
'                mcstrEnvironments

    ' Build an array (astrEnvKeys) containing each defined Environment key in the
    ' registry, e.g., Prod and IDev
    RegistryEnumerateSubKeys rrkHKEY_LOCAL_MACHINE, strRegKey, astrEnvKeys(), lngNbrOfEnvs

    If lngNbrOfEnvs < 1 Then
        gerhApp.SaveAppSpecificErr vbObjectError + gcRES_FERR_NO_ENVS, _
                                   mcstrName & cstrCurrentProc
        GoTo PROC_EXIT
    End If

    ' Resize the UDT array (m_audtEnvironments) according to how many entries were defined in the
    ' registry, then load the values associated with each such Environment key (e.g., database,
    ' server, and environment name) to that array.
    ReDim m_audtEnvironments(lngNbrOfEnvs - 1) As udtEnvironment

    For lngEnvs = 0 To lngNbrOfEnvs - 1
        m_audtEnvironments(lngEnvs).EnvironmentKey = astrEnvKeys(lngEnvs)
        strTempKey = strRegKey & mcstrSlash & astrEnvKeys(lngEnvs)
        m_audtEnvironments(lngEnvs).ActiveDatabase = RegistryGetKeyValue(rrkHKEY_LOCAL_MACHINE, _
                                               strTempKey, mcstrEnv_ActiveDatabase)
        m_audtEnvironments(lngEnvs).ActiveServer = RegistryGetKeyValue(rrkHKEY_LOCAL_MACHINE, _
                                               strTempKey, mcstrEnv_ActiveServer)
        m_audtEnvironments(lngEnvs).ArchiveDatabase = RegistryGetKeyValue(rrkHKEY_LOCAL_MACHINE, _
                                               strTempKey, mcstrEnv_ArchiveDatabase)
        m_audtEnvironments(lngEnvs).ArchiveServer = RegistryGetKeyValue(rrkHKEY_LOCAL_MACHINE, _
                                               strTempKey, mcstrEnv_ArchiveServer)
        m_audtEnvironments(lngEnvs).Port = RegistryGetKeyValue(rrkHKEY_LOCAL_MACHINE, _
                                               strTempKey, mcstrEnv_Port)
        m_audtEnvironments(lngEnvs).EnvironmentName = RegistryGetKeyValue(rrkHKEY_LOCAL_MACHINE, _
                                               strTempKey, mcstrEnv_EnvironmentName)
        'SQL_INTEGRATED_SECURITY
        If UCase$(RegistryGetKeyValue(rrkHKEY_LOCAL_MACHINE, strTempKey, mcstrEnv_UsesWindowsAuthentication)) = "TRUE" Then
            m_audtEnvironments(lngEnvs).UsesWindowsAuthentication = True
        Else
            m_audtEnvironments(lngEnvs).UsesWindowsAuthentication = False
        End If
        'SQL_INTEGRATED_SECURITY
    Next lngEnvs
PROC_EXIT:
    On Error GoTo 0     ' disable error handler

    ' Clean-up statements go here
    Erase astrEnvKeys

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
'|                      PRIVATE   Procedures                        |
'/                                                                  \
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/

