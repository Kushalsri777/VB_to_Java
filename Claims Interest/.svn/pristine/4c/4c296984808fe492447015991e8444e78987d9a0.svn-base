Attribute VB_Name = "modRegistrySettings"
'******************************************************************************
' Module     : modRegistryAndSettings
' Description: This standard module contains procedures relating to
'              retrieving and setting application-related values
'              from/to the registry and global variables.
' Procedures :
'              fnGetDriveFromPath(ByVal strPath As String)
'              fnIsDBPathValid() As Boolean
'              fnRegGetAppSettings()
'              fnRegGetClerkCode() As String
'              fnRegGetDBName() As String
'              fnRegGetDBPath() As String
'              fnRegInitializeForApp()
'              fnRegSetClerkCode(ByVal strClerkCode As String)
'              fnRegSetDBName(ByVal strDBName As String)
'              fnRegSetDBPath(ByVal strDBPath As String)
' Modified   :
' 03/26/01 BAW Cleaned with Total Visual CodeTools 2000
' 01/2002  BAW Optimized per Project Analyzer (removing dead code, adding "$" to Mid/Space, etc.).
'              Changed all calls to procs in CRegSettings to include a "Root" parameter, indicating
'              whether HKLM or HKCU should be accessed. This is so any registry writes are done
'              to HKCU, so they'll be successful on a Win2K PC where the user is non-Administrator.
' --------------------------------------------------
Option Explicit
Option Compare Binary

Private Const mcstrName             As String = "modRegistryAndSettings."

'-----------------------------------------------------------------------
' The following are used by procedures interacting with cRegSettings
' to read/write items in the registry. They should always "jive" with
' what the Claims Interest setup programs uses!
'-----------------------------------------------------------------------
Public Const gcRegCompany As String = "Sun Life Financial"
Public Const gcRegAppName As String = "Claims Interest"
' Default values of 'missing' registry settings
Public Const gcEmpty As String = "EMPTY"
Public Const gcDefaultDBPath = "L:\CLAIMSINTEREST"

'-----------------------------------------------------------------------
' The following are used throughout the app when it needs to
' determine the name of the database or its location (per registry
' settings for the app). If the registry's entry re: DBName doesn't exist,
' then the value of gcDefaultDBName will be used as the database name.
'-----------------------------------------------------------------------
'Public gstrPath As String
Public gstrDBPath As String
Public gstrDBName As String
Public gstrDBPathAndName As String
Public gstrClerkCode As String
Public Const gcDefaultDBName As String = "CLAIMS.MDB"
Public Const gcClaimsManagerClerkCode = "A1GMW"
Public Const gcstrPassword = "ireland"       ' old one = purple

Public gcReg As CRegSettings

' 01/31/2002 BAW - Added the following constants to support having default values in HKLM but per-user settings in HKCU
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function fnGetDriveFromPath(ByVal strPath As String) As String
    ' Comments  : Returns the drive letter part of the path
    ' Parameters: pstrPath - path containing the drive letter
    ' Returns   : the drive letter
    ' Source    : Total Visual SourceBook 2000
    '
    ' Called by : Form_Load( ) in frmSetDatabaseLocation
    '
    Dim intPos As Integer
    Dim strTmp As String
    Const cstrDelimiter As String = ":\"
    Const cstrCurrentProc As String = "fnGetDriveFromPath"

    On Error GoTo PROC_ERR

    ' Initialize the return value
    strTmp = vbNullString

    ' See of the colon and backslash exist
    intPos = InStr(strPath, cstrDelimiter)

    If intPos > 0 Then
        ' They exist, so return the remainder
        strTmp = Left$(strPath, intPos)
    Else
        ' Look for the colon
        intPos = InStr(strPath, ":")
        If intPos > 0 Then
            ' It exists so return the remainder
            strTmp = Left$(strPath, intPos)
        Else
            ' No drive letter information, so return a zero-length string
            strTmp = vbNullString
        End If
    End If

    fnGetDriveFromPath = strTmp
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
Public Function fnIsDBPathValid() As Boolean
    '----------------------------------------------------------------------------
    ' Procedure   :  Function fnIsDBPathValid
    ' Created by  :  BAW on 04-26-2001 11:18
    '
    ' Comments    :
    ' Called by   : Form_Load( ) in frmSetDatabaseLocation
    '               Main( ) in modStartup
    '
    ' Parameters  : None
    '
    ' Return value: True if global vars (gstrDBPath, gstrDBName, gstrDBPathAndName)
    '               point to a valid location in which CLAIMS.MDB exists
    ' Modified     :
    '----------------------------------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnIsDBPathValid"
    Dim fso As FileSystemObject

    fnIsDBPathValid = True

    If (gstrDBPath = gcEmpty) Or (gstrDBName = gcEmpty) Then
        fnIsDBPathValid = False
    Else
        Set fso = CreateObject("Scripting.FileSystemObject")
        If Not (fso.FileExists(gstrDBPathAndName)) Then
            fnIsDBPathValid = False
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
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub fnRegGetAppSettings()
    ' Comments  : Retrieves current app settings from registry, e.g,
    '             DBPath, DBName, ClerkCode and stores them in their
    '             corresponding global variables (gstrDBPath,
    '             gstrDBName and gstrClerkCode)
    '
    ' Called by : Sub Main( ) in modStartup
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnRegGetAppSettings"

    fnRegGetDBPath
    fnRegGetDBName
    fnRegGetClerkCode

    gstrDBPathAndName = gstrDBPath & "\" & gstrDBName
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
Public Sub fnRegGetClerkCode()
    ' Comments  : Retrieves the Clerk Code app-related entry from HKCU in the registry:
    '             Sun Life Financial\Claims Interest\System\ClerkCode, building a
    '             default value for that key if it wasn't found.
    ' Parameters: None
    ' Returns   : N/A
    ' Modified  :
    '   01/31/02 BAW  Changed to retrieve the key from HKCU if possible. Otherwise, get
    '                 the default value from HKLM, if present, and save it to HKCU.
    '                 This is so the app no longer writes to registry keys that may
    '                 not be accessible to a non-Administrator user under Win2K.
    ' Called by : fnRegGetAppSettings( ) in modRegistrySettings
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnRegGetClerkCode"

    gstrClerkCode = gcReg.GetSetting(HKEY_CURRENT_USER, "System", "ClerkCode", gcEmpty)
    
    If gstrClerkCode = gcEmpty Then
        gstrClerkCode = gcReg.GetSetting(HKEY_LOCAL_MACHINE, "System", "ClerkCode", gcEmpty)
        If gstrClerkCode = gcEmpty Then
            ' Build default key in HKLM
            fnRegSetClerkCode gcClaimsManagerClerkCode, HKEY_LOCAL_MACHINE
        End If
        ' Build key in HKCU
        fnRegSetClerkCode gstrClerkCode
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



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub fnRegGetDBName()
    ' Comments  : Retrieves the DBName app-related entry from HKCU in the registry:
    '             Sun Life Financial\Claims Interest\System\DBName, building a
    '             default value for that key if it wasn't found.
    ' Parameters: None
    ' Returns   : a string containing the value of that registry key
    ' Modified  :
    '   01/31/02 BAW  Changed to retrieve the key from HKCU if possible. Otherwise, get
    '                 the default value from HKLM, if present, and save it to HKCU.
    '                 This is so the app no longer writes to registry keys that may
    '                 not be accessible to a non-Administrator user under Win2K.
    '
    ' Called by : fnRegGetAppSettings( ) in modRegistrySettings
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnRegGetDBName"

    gstrDBName = gcReg.GetSetting(HKEY_CURRENT_USER, "System", "DBName", gcEmpty)
    
    If gstrDBName = gcEmpty Then
        gstrDBName = gcReg.GetSetting(HKEY_LOCAL_MACHINE, "System", "DBName", gcEmpty)
        If gstrDBName = gcEmpty Then
            ' Build default key in HKLM
            fnRegSetDBName gcDefaultDBName, HKEY_LOCAL_MACHINE
        End If
        ' Build key in HKCU
        fnRegSetDBName gstrDBName
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



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub fnRegGetDBPath()
    ' Comments  : Retrieves the DBPath app-related entry from HKCU in the registry:
    '             Sun Life Financial\Claims Interest\System\DBPath, building a
    '             default value for that key if it wasn't found.
    ' Parameters: None
    ' Modified  :
    '   01/31/02 BAW  Changed to retrieve the key from HKCU if possible. Otherwise, get
    '                 the default value from HKLM, if present, and save it to HKCU.
    '                 This is so the app no longer writes to registry keys that may
    '                 not be accessible to a non-Administrator user under Win2K.
    '
    ' Called by : fnRegGetAppSettings( ) in modRegistrySettings
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnRegGetDBPath"

    gstrDBPath = gcReg.GetSetting(HKEY_CURRENT_USER, "System", "DBPath", gcEmpty)
    
    If gstrDBPath = gcEmpty Then
        gstrDBPath = gcReg.GetSetting(HKEY_LOCAL_MACHINE, "System", "DBPath", gcEmpty)
        If gstrDBPath = gcEmpty Then
            ' Build default key in HKLM
            fnRegSetDBPath gcDefaultDBPath, HKEY_LOCAL_MACHINE
        End If
        ' Store key in HKCU
        fnRegSetDBPath gstrDBPath
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



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub fnRegInitializeForApp()
    ' Comments  : Initializes the CRegSettings object with
    '             app-specific values for Company Name and App Name.
    ' Called by : Sub Main() in modStartup
    ' Parameters: None
    ' Modified  :
    ' Called by : fnRegInitializeForApp( ) in modRegistrySettings
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnRegInitializeForApp"

    ' Establish and initialize a global object pointer to an instance of a
    ' CRegSettings class whose methods will be used to read/write
    ' registry settings.
    Set gcReg = New CRegSettings
    gcReg.Company = gcRegCompany
    gcReg.AppName = gcRegAppName
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
Public Sub fnRegSetClerkCode(ByVal strClerkCode As String, Optional ByVal Root As Long = HKEY_CURRENT_USER)
    ' Comments  : Stores the Clerk Code in the appropriate app-related
    '             entry in the registry:
    '             Sun Life Financial\Claims Interest\System\ClerkCode
    ' Called by : cmdUpdate_Click( ) in frmInsured
    '             fnRegGetAppSettings( ) in modRegistrySettings
    ' Parameters: strClerkCode, the value to store
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnRegSetClerkCode"

    gcReg.SaveSetting Root, "System", "ClerkCode", UCase$(strClerkCode)
    gstrClerkCode = strClerkCode
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
Public Sub fnRegSetDBName(ByVal strDBName As String, Optional ByVal Root As Long = HKEY_CURRENT_USER)
    ' Comments  : Stores the DBName in the appropriate app-related
    '             entry in the registry:
    '             Sun Life Financial\Claims Interest\System\DBName
    '             and updates related global variables to ensure
    '             they're always kept in synch.
    ' Called by : cmdApply_Click( ) in frmSetDatabaseLocation
    '             fnRegGetAppSettings( ) in modRegistrySettings
    '
    ' Parameters: strDBName, the value to store
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnRegSetDBName"

    strDBName = UCase$(strDBName)
    gcReg.SaveSetting Root, "System", "DBName", strDBName
    gstrDBName = strDBName
    gstrDBPathAndName = gstrDBPath & "\" & gstrDBName
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
Public Sub fnRegSetDBPath(ByVal strDBPath As String, Optional ByVal Root As Long = HKEY_CURRENT_USER)
    ' Comments  : Stores the DBPath in the appropriate app-related
    '             entry in the registry:
    '             Sun Life Financial\Claims Interest\System\DBPath
    '             NOTE: A trailing slash will be deleted if it exists.
    '
    ' Called by : cmdApply_Click( ) in frmSetDatabaseLocation
    '             fnRegGetAppSettings( ) in modRegistrySettings
    '
    ' Parameters: strDBPath, the value to store
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnRegSetDBPath"

    If Right$(strDBPath, 1) = "\" Then
        strDBPath = Left$(strDBPath, Len(strDBPath) - 1)
    End If

    strDBPath = UCase$(strDBPath)
    gcReg.SaveSetting Root, "System", "DBPath", strDBPath
    gstrDBPath = strDBPath
    gstrDBPathAndName = gstrDBPath & "\" & gstrDBName
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
