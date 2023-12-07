Attribute VB_Name = "modRegistry"
'
' Module      : modRegistry
' Description : This module Implements routines for manipulating the registry.
' Source      : Total Visual SourceBook 2000
'
' Procedures  :
'   Private
'   Public      RegistryCreateNewKey(eRootKey As EnumRegistryRootKeys, strKeyName As String)
'   Public      RegistryDeleteKey(eRootKey As EnumRegistryRootKeys, strKeyName As String)
'   Public      RegistryDeleteValue(eRootKey As EnumRegistryRootKeys, strKeyName As String, _
'                                   strValueName As String)
'   Public      RegistryEnumerateSubKeys(eRootKey As EnumRegistryRootKeys, strKeyName As String, _
'                                   astrKeys() As String, lngKeyCount As Long)
'   Public      RegistryEnumerateValues(eRootKey As EnumRegistryRootKeys, strKeyName As String, _
'                                   astrValues() As String, lngValueCount As Long)
'   Public      RegistryGetKeyValue(eRootKey As EnumRegistryRootKeys, strKeyName As String, _
'                                   strValueName As String) As Variant
'   Public      RegistrySetKeyValue(eRootKey As EnumRegistryRootKeys, strKeyName As String, _
'                                   strValueName As String, varData As Variant, eDataType As EnumRegistryValueType)
'
' Modified    :
' 03/03/02 BAW (Phase2A) Added support for new global error handler
'
Option Explicit
Option Compare Binary
Private Const mcstrName As String = "modRegistry."

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Enum EnumRegistryRootKeys
    rrkHKEY_CLASSES_ROOT = &H80000000
    rrkHKEY_CURRENT_USER = &H80000001
    rrkHKEY_LOCAL_MACHINE = &H80000002
    rrkHKEY_USERS = &H80000003
End Enum

Public Enum EnumRegistryValueType
    rrkRegSZ = 1
    rrkregBinary = 3
    rrkRegDWord = 4
End Enum

Private Declare Function RegCloseKey Lib "advapi32.dll" _
    (ByVal lngHKey As Long) As Long

Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" _
    (ByVal lngHKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, _
     ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, _
     ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) _
     As Long

Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
    (ByVal lngHKey As Long, ByVal lpSubKey As String) As Long

Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
    (ByVal lngHKey As Long, ByVal lpValueName As String) As Long

Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" _
    (ByVal lngHKey As Long, ByVal dwIndex As Long, ByVal lpName As String, _
     ByVal cbName As Long) As Long

Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" _
    (ByVal lngHKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
     lpcbValueName As Long, ByVal lpReserved As Long, ByVal lpType As Long, _
     ByVal lpData As Long, ByVal lpcbData As Long) As Long

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    (ByVal lngHKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
     ByVal samDesired As Long, phkResult As Long) As Long

Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" _
    (ByVal lngHKey As Long, ByVal lpClass As String, ByVal lpcbClass As Long, _
     ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, _
     ByVal lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, _
     ByVal lpcbMaxValueLen As Long, ByVal lpcbSecurityDescriptor As Long, _
     lpftLastWriteTime As FILETIME) As Long

Private Declare Function RegQueryValueExBinary Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal lngHKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
     lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long

Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal lngHKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
     lpType As Long, lpData As Long, lpcbData As Long) As Long

Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
     lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long

Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal lngHKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
     lpType As Long, ByVal lpData As String, lpcbData As Long) As Long

Private Declare Function RegSetValueExBinary Lib "advapi32.dll" Alias "RegSetValueExA" _
    (ByVal lngHKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
     ByVal dwType As Long, ByVal lpValue As Long, ByVal cbData As Long) As Long

Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" _
    (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
     ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long

Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" _
    (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
     ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long


Private Const mcregOptionNonVolatile = 0

Private Const mcregErrorNone = 0
Private Const mcregErrorBadDB = 1
Private Const mcregErrorBadKey = 2
Private Const mcregErrorCantOpen = 3
Private Const mcregErrorCantRead = 4
Private Const mcregErrorCantWrite = 5
Private Const mcregErrorOutOfMemory = 6
Private Const mcregErrorInvalidParameter = 7
Private Const mcregErrorAccessDenied = 8
Private Const mcregErrorInvalidParameterS = 87
Private Const mcregErrorNoMoreItems = 259

Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const SYNCHRONIZE = &H100000

Private Const KEY_CREATE_LINK = &H20
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))



'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'\                                                                  /
'|                        PUBLIC  Procedures                        |
'/                                                                  \
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub RegistryCreateNewKey(ByVal eRootKey As EnumRegistryRootKeys, ByVal strKeyName As String)
    ' Comments  : Creates a new key in the system registry
    ' Parameters: eRootKey - The root key
    '             strKeyName - The name of the key to create
    ' Returns   : Nothing
    '
    ' Called by :
    '
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "RegistryCreateNewKey"
    Dim lngRetVal As Long
    Dim lngHKey As Long

    On Error GoTo PROC_ERR

    ' Create the key
    lngRetVal = RegCreateKeyEx(eRootKey, strKeyName, 0&, vbNullString, _
        mcregOptionNonVolatile, KEY_WRITE, 0&, lngHKey, 0&)

    ' if the key was created, then close it
    If lngRetVal = mcregErrorNone Then
        RegCloseKey (lngHKey)
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



'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub RegistryDeleteKey(ByVal eRootKey As EnumRegistryRootKeys, ByVal strKeyName As String)
    ' Comments  : Deletes a key from the system registry
    ' Parameters: eRootKey - The root key
    '             strKeyName - The name of the key to delete
    ' Returns   : Nothing
    '
    ' Called by :
    '
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "RegistryDeleteKey"
    Dim lngRetVal As Long

    On Error GoTo PROC_ERR

    ' Delete the key
    lngRetVal = RegDeleteKey(eRootKey, strKeyName)
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



'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub RegistryDeleteValue(ByVal eRootKey As EnumRegistryRootKeys, ByVal strKeyName As String, _
                              ByVal strValueName As String)
    ' Comments  : Deletes a value from the system registry
    ' Parameters: eRootKey - The root key
    '             strKeyName - The name of the key to delete
    '             strValueName - The name of the value to delete
    ' Returns   : Nothing
    '
    ' Called by :
    '
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "RegistryDeleteValue"
    Dim lngRetVal As Long
    Dim lngHKey As Long

    On Error GoTo PROC_ERR

    ' Open the key
    lngRetVal = RegOpenKeyEx(eRootKey, strKeyName, 0, KEY_WRITE, _
        lngHKey)

    ' If the key was opened successfully, then delete it
    If lngRetVal = mcregErrorNone Then
        lngRetVal = RegDeleteValue(lngHKey, strValueName)
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



'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub RegistryEnumerateSubKeys(ByVal eRootKey As EnumRegistryRootKeys, ByVal strKeyName As String, _
                                    ByRef astrKeys() As String, ByRef lngKeyCount As Long)
    ' Comments  : Enumerates the sub keys of the specified key
    ' Parameters: eRootKey - The root key
    '             strKeyName - The name of the key to enumerate
    '             astrKeys - An array of strings to fill with sub key names
    '             lngKeyCount - The number of sub keys returned in the parameter
    '             astrKeys
    ' Returns   : Nothing
    '
    ' Called by :
    '
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "RegistryEnumerateSubKeys"
    Dim lngHKey         As Long
    Dim lngKeyIndex     As Long
    Dim lngMaxKeyLen    As Long
    Dim lngRetVal       As Long
    Dim lngSubkeyCount  As Long
    Dim strSubKeyName   As String
    Dim typFT           As FILETIME

    On Error GoTo PROC_ERR

    ' Open the key
    lngRetVal = RegOpenKeyEx(eRootKey, strKeyName, 0, KEY_READ, _
        lngHKey)

    If lngRetVal = mcregErrorNone Then
        'find the number of subkeys, and redim the return string array
        lngRetVal = RegQueryInfoKey(lngHKey, vbNullString, 0, 0, lngSubkeyCount, _
            lngMaxKeyLen, 0, 0, 0, 0, 0, typFT)
        If mcregErrorNone = lngRetVal Then
            If lngSubkeyCount > 0 Then
                ReDim astrKeys(lngSubkeyCount - 1) As String

                'set up the while loop
                lngKeyIndex = 0
                ' Pad the string to the maximum length of a sub key, plus 1 for null
                ' termination
                lngMaxKeyLen = lngMaxKeyLen + 1
                strSubKeyName = Space$(lngMaxKeyLen)

                Do While RegEnumKey(lngHKey, lngKeyIndex, strSubKeyName, lngMaxKeyLen + 1) = 0

                    ' Set the string array to the key name, removing null termination
                    If InStr(1, strSubKeyName, vbNullChar) > 0 Then
                        astrKeys(lngKeyIndex) = Left$(strSubKeyName, InStr(1, strSubKeyName, _
                            vbNullChar) - 1)
                    End If
                    ' Increment the key index for the return string array
                    lngKeyIndex = lngKeyIndex + 1

                Loop
            End If
            ' return the new dimension of the return string array
            lngKeyCount = lngSubkeyCount
        End If

        ' Close the key
        RegCloseKey (lngHKey)
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



'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub RegistryEnumerateValues(ByVal eRootKey As EnumRegistryRootKeys, ByVal strKeyName As String, _
                                   ByRef astrValues() As String, ByRef lngValueCount As Long)
    ' Comments  : Enumerates the values of the specified key
    ' Parameters: eRootKey - The root key
    '             strKeyName - The name of the key to enumerate
    '             astrValues - An array of strings to fill with value names
    '             lngValueCount - The number of values returned in the parameter astrValues
    ' Returns   : Nothing
    '
    ' Called by :
    '
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "RegistryEnumerateValues"
    Dim lngHKey             As Long
    Dim lngKeyIndex         As Long
    Dim lngMaxValueLen      As Long
    Dim lngRetVal           As Long
    Dim lngTempValueCount   As Long
    Dim strValueName        As String
    Dim typFT               As FILETIME

    On Error GoTo PROC_ERR

    ' Open the key
    lngRetVal = RegOpenKeyEx(eRootKey, strKeyName, 0, KEY_READ, _
        lngHKey)

    If lngRetVal = mcregErrorNone Then
        'find the number of subkeys, and redim the return string array
        lngRetVal = RegQueryInfoKey(lngHKey, vbNullString, 0, 0, 0, _
            0, 0, lngTempValueCount, lngMaxValueLen, 0, 0, typFT)
        If mcregErrorNone = lngRetVal Then
            If lngTempValueCount > 0 Then
                ReDim astrValues(lngTempValueCount - 1) As String

                'set up the while loop
                lngKeyIndex = 0
                ' Pad the string to the maximum length of a sub key, plus 1 for null
                ' termination
                lngMaxValueLen = lngMaxValueLen + 1
                strValueName = Space$(lngMaxValueLen)

                Do While RegEnumValue(lngHKey, lngKeyIndex, strValueName, _
                    lngMaxValueLen + 1, 0, 0, 0, 0) = 0

                    ' Set the string array to the key name, removing null termination
                    If InStr(1, strValueName, vbNullChar) > 0 Then
                        astrValues(lngKeyIndex) = Left$(strValueName, InStr(1, strValueName, _
                            vbNullChar) - 1)
                    End If
                    ' Increment the key index for the return string array
                    lngKeyIndex = lngKeyIndex + 1

                Loop
            End If
            ' return the new dimension of the return string array
            lngValueCount = lngTempValueCount
        End If

        ' Close the key
        RegCloseKey (lngHKey)
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



'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Function RegistryGetKeyValue(ByVal eHKeyRoot As EnumRegistryRootKeys, _
                                    ByVal strKeyName As String, ByVal strValueName As String) As Variant
    ' Comments  : Returns a value from the system registry
    ' Parameters: eHKeyRoot - The root key
    '             strKeyName - The name of the key
    '             strValueName - The name of the value
    ' Returns   : The data in the registry value
    '
    ' Called by : ReadEntry( ) in CAppSettings.cls
    '             RestoreForm( ) in CAppSettings.cls
    '
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "RegistryGetKeyValue"
    Dim abytValueData()   As Byte
    Dim lngDataSize       As Long
    Dim lngHKey           As Long
    Dim lngRetVal         As Long
    Dim lngValueData      As Long
    Dim lngValueType      As Long
    Dim strValueData      As String
    Dim varValue          As Variant

    On Error GoTo PROC_ERR

    varValue = Empty

    lngRetVal = RegOpenKeyEx(eHKeyRoot, strKeyName, 0&, KEY_READ, _
        lngHKey)

    If mcregErrorNone = lngRetVal Then

        lngRetVal = RegQueryValueExNULL(lngHKey, strValueName, 0&, lngValueType, _
            0&, lngDataSize)

        If lngRetVal = mcregErrorNone Then

            Select Case lngValueType

            ' String type

                Case rrkRegSZ:
                    If lngDataSize > 0 Then
                        strValueData = String(lngDataSize, 0)
                        lngRetVal = RegQueryValueExString(lngHKey, strValueName, 0&, _
                            lngValueType, strValueData, lngDataSize)
                        If InStr(strValueData, vbNullChar) > 0 Then
                            strValueData = Mid$(strValueData, 1, InStr(strValueData, _
                                vbNullChar) - 1)
                        End If
                    End If
                    If mcregErrorNone = lngRetVal Then
                        varValue = Left$(strValueData, lngDataSize)
                    Else
                        varValue = Empty
                    End If

                ' Long type
                Case rrkRegDWord:
                    lngRetVal = RegQueryValueExLong(lngHKey, strValueName, 0&, _
                        lngValueType, lngValueData, lngDataSize)
                    If mcregErrorNone = lngRetVal Then
                        varValue = lngValueData
                    End If

                ' Binary type
                Case rrkregBinary
                    If lngDataSize > 0 Then
                        ReDim abytValueData(lngDataSize - 1) As Byte
                        lngRetVal = RegQueryValueExBinary(lngHKey, strValueName, 0&, _
                            lngValueType, VarPtr(abytValueData(0)), lngDataSize)
                    End If
                    If mcregErrorNone = lngRetVal Then
                        varValue = abytValueData
                    Else
                        varValue = Empty
                    End If

                Case Else
                    'No other data types supported
                    lngRetVal = -1
                    gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_UNEXPECTED_VAL_SELECT_CASE, _
                        mcstrName & cstrCurrentProc
                    GoTo PROC_EXIT
            End Select
        End If

        RegCloseKey (lngHKey)
    End If

    'Return varValue
    RegistryGetKeyValue = varValue
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
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_CANTOPEN_REGKEY, _
                                   mcstrName & cstrCurrentProc, _
                                   strValueName, strKeyName
    End Select
    Resume PROC_EXIT
End Function



'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub RegistrySetKeyValue(ByVal eHKeyRoot As EnumRegistryRootKeys, ByVal strKeyName As String, _
                               ByVal strValueName As String, ByVal varData As Variant, ByVal eDataType As EnumRegistryValueType)
    ' Comments  : This procedure sets a key value, creating the key if it doesn't exist.
    ' Parameters: eHKeyRoot - The root key
    '             strKeyName - The name of the key to open
    '             strValueName - The name of the value to open (vbNulLString will open the default value).
    '             varData - The data to store in the value
    '             eDataType - The type of data to store in the value
    ' Returns   : Nothing
    '
    ' Called by : SaveForm( ) in CAppSettings.cls
    '             WriteEntry( ) in CAppSettings.cls
    '
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "RegistrySetKeyValue"
    Dim abytData()    As Byte
    Dim lngData       As Long
    Dim lngHKey       As Long
    Dim lngRetVal     As Long
    Dim strData       As String

    On Error GoTo PROC_ERR

    ' Open the specified key. If it does not exist, then create it
    lngRetVal = RegCreateKeyEx(eHKeyRoot, strKeyName, 0&, vbNullString, _
        mcregOptionNonVolatile, KEY_READ Or KEY_WRITE, 0&, lngHKey, 0&)

    ' Determine the data type of the key
    Select Case eDataType
        Case rrkRegSZ       ' String
            strData = varData & vbNullChar
            lngRetVal = RegSetValueExString(lngHKey, strValueName, 0&, eDataType, _
                strData, Len(strData))
        Case rrkRegDWord    ' DWord
            lngData = varData
            lngRetVal = RegSetValueExLong(lngHKey, strValueName, 0&, eDataType, _
                lngData, Len(lngData))
        Case rrkregBinary   ' Binary
            abytData = varData
            lngRetVal = RegSetValueExBinary(lngHKey, strValueName, 0&, eDataType, _
                VarPtr(abytData(0)), UBound(abytData) + 1)
        'Case Else
            ' Do nothing
    End Select

    RegCloseKey (lngHKey)
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
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_CANTSAVE_REGKEY, _
                                   mcstrName & cstrCurrentProc, _
                                   strValueName & "=" & CStr(varData), strKeyName
    End Select
    Resume PROC_EXIT
End Sub
