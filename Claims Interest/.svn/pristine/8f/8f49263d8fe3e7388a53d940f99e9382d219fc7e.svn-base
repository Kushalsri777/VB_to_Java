VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRegSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' --------------------       Modification History       --------------------
'
'  01/2002 BAW - Added "$" to some functions like Trim to optimize it, plus
'                other minor optimizations per Project Analyzer
'                Changed all calls to procs in CRegSettings to include a "Root"
'                parameter, indicating whether HKLM or HKCU should be accessed.
'                This is so any registry writes are done to HKCU, so they'll
'                be successful on a Win2K PC where the user is non-Administrator.
' --------------------------------------------------------------------------


' *********************************************************************
'  Copyright ©1997-99 Karl E. Peterson, All Rights Reserved
' *********************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code without prior written consent.
' *********************************************************************
Option Explicit
Option Compare Binary

Private Const mcstrName             As String = "cRegSettings."

'
' Win32 Registry functions
'
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
'
' Constants for Windows 32-bit Registry API
'
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003
Private Const HKEY_PERFORMANCE_DATA = &H80000004
Private Const HKEY_CURRENT_CONFIG = &H80000005
Private Const HKEY_DYN_DATA = &H80000006
'
' Reg result codes
'
Private Const REG_CREATED_NEW_KEY = &H1                      ' New Registry Key created
Private Const REG_OPENED_EXISTING_KEY = &H2                      ' Existing Key opened
'
' Reg Create Type Values...
'
Private Const REG_OPTION_RESERVED = 0           ' Parameter is reserved
Private Const REG_OPTION_NON_VOLATILE = 0       ' Key is preserved when system is rebooted
Private Const REG_OPTION_VOLATILE = 1           ' Key is not preserved when system is rebooted
Private Const REG_OPTION_CREATE_LINK = 2        ' Created key is a symbolic link
Private Const REG_OPTION_BACKUP_RESTORE = 4     ' open for backup or restore
'
' Reg Key Security Options
'
Private Const DELETE = &H10000
Private Const READ_CONTROL = &H20000
Private Const WRITE_DAC = &H40000
Private Const WRITE_OWNER = &H80000
Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const SPECIFIC_RIGHTS_ALL = &HFFFF
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Private Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))

Private Const ERROR_SUCCESS = 0&
Private Const ERROR_MORE_DATA = 234
Private Const ERROR_NO_MORE_ITEMS = 259

Private Const REG_SZ = 1                         ' Unicode nul terminated string
'
' Private member variables
'
Private m_Company As String
Private m_AppName As String
'
' Private class constants
'
Private Const defCompany As String = "VB and VBA Program Settings"

' ********************************************
'  Initialize and Terminate
' ********************************************
Private Sub Class_Initialize()
   m_Company = defCompany
   m_AppName = App.ProductName
End Sub

' ********************************************
'  Public Properties
' ********************************************
Public Property Let Company(ByVal NewVal As String)
    ' Called by : fnRegInitializeForApp( ) in modRegistrySettings
   If Len(NewVal) Then
      m_Company = Trim$(NewVal)
   Else
      m_Company = defCompany
   End If
End Property




Public Property Let AppName(ByVal NewVal As String)
    ' Called by : fnRegGetAppSettings( ) in modRegistrySettings
   If Len(NewVal) Then
      m_AppName = Trim$(NewVal)
   Else
      m_AppName = App.ProductName
   End If
End Property




' ********************************************
'  Public Methods
' ********************************************



'//////////////////////////////////////////////////////////////////////////////
Public Function GetSetting(ByVal Root As Long, ByVal Section As String, ByVal Key As String, Optional ByVal Default As String = "") As String
    ' Section   Required. String expression containing the name of the section where the key setting is found.
    '           If omitted, key setting is assumed to be in default subkey.
    ' Key       Required. String expression containing the name of the key setting to return.
    ' Default   Optional. Expression containing the value to return if no value is set in the key setting.
    '           If omitted, default is assumed to be a zero-length string ("").
    ' Called by fnRegGetClerkCode( ) in modRegistrySettings
    '           fnRegGetDBName( ) in modRegistrySettings
    '           fnRegGetDBPath( ) in modRegistrySettings
    Const cstrCurrentProc   As String = "GetSettings"
    On Error GoTo PROC_ERR
        
    Dim nRet As Long
    Dim hKey As Long
    Dim nType As Long
    Dim nBytes As Long
    Dim Buffer As String
   
    ' Assume failure and set return to Default
    GetSetting = Default

    ' Open key
    nRet = RegOpenKeyEx(Root, SubKey(Section), 0&, KEY_ALL_ACCESS, hKey)
    If nRet = ERROR_SUCCESS Then
        ' Set appropriate value for default query
        If Key = "*" Then
            Key = vbNullString
        End If
      
        ' Determine how large the buffer needs to be
        nRet = RegQueryValueEx(hKey, Key, 0&, nType, ByVal Buffer, nBytes)
        If nRet = ERROR_SUCCESS Then
            ' Build buffer and get data
            If nBytes > 0 Then
                Buffer = Space$(nBytes)
                nRet = RegQueryValueEx(hKey, Key, 0&, nType, ByVal Buffer, Len(Buffer))
                If nRet = ERROR_SUCCESS Then
                    ' Trim NULL and return successful query!
                    GetSetting = Left$(Buffer, nBytes - 1)
                End If
            End If
        RegCloseKey hKey
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



'//////////////////////////////////////////////////////////////////////////////
Public Function SaveSetting(ByVal Root As Long, ByVal Section As String, ByVal Key As String, ByVal Setting As String) As Boolean
    ' Section   Required. String expression containing the name of the section where the key setting is being saved.
    ' Key       Required. String expression containing the name of the key setting being saved.
    ' Setting   Required. Expression containing the value that key is being set to.
    ' Called by fnRegSetClerkCode( ) in modRegistrySettings
    '           fnRegSetDBName( ) in modRegistrySettings
    '           fnRegSetDBPath( ) in modRegistrySettings
    Const cstrCurrentProc   As String = "SaveSetting"
    On Error GoTo PROC_ERR
   
    Dim nRet As Long
    Dim hKey As Long
    Dim nResult As Long
   
    ' Open (or create and open) key
    nRet = RegCreateKeyEx(Root, SubKey(Section), 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, hKey, nResult)
    If nRet = ERROR_SUCCESS Then
        ' Set appropriate value for default query
        If Key = "*" Then
            Key = vbNullString
        End If
        ' Null-terminate setting, in case it's empty.
        ' Strange mirroring can occur otherwise.
        Setting = Setting & vbNullChar
        ' Write new value to registry
        nRet = RegSetValueEx(hKey, Key, 0&, REG_SZ, ByVal Setting, Len(Setting))
        RegCloseKey hKey
    End If
    SaveSetting = (nRet = ERROR_SUCCESS)
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



' ********************************************
'  Private Methods
' ********************************************

'//////////////////////////////////////////////////////////////////////////////
Private Function SubKey(Optional ByVal Section As String = "") As String
    ' Build SubKey from known values
    ' Called by  DeleteSetting( ) in CRegSettings
    '            GetAllSettings( ) in CRegSettings
    '            GetSetting( ) in CRegSettings
    '            SaveSetting( ) in CRegSettings
    Const cstrCurrentProc   As String = "Property Get Initialized"
    On Error GoTo PROC_ERR
   
    SubKey = "Software\" & m_Company & "\" & m_AppName
    If Len(Section) Then
        SubKey = SubKey & "\" & Section
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

