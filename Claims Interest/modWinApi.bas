Attribute VB_Name = "modWinApi"
' Module     : modWinApi
' Description: Declarations and constants associated with Windows API functions called
'              throughout the app
' Procedures :
'              fnEnumWindowsProc(ByVal hwnd As Long, ByVal NotUsed As Long) As Boolean
'              fnEnumAllWindows(ByVal strSessionID As String, ByVal strSearchString As String, _
'                  ByRef hwndFound As Long) As Boolean
'              fnWindowText(ByVal hwnd As Long) As String
'
' Uses       : USER32.DLL, to get EnumWindows() and GetWindowText()
' Modified   :
' 03/16/01 DAS Cleaned with Total Visual CodeTools 2000
' --------------------------------------------------
Option Explicit
Option Compare Binary

Private Const mcstrName             As String = "modWinApi."


'------------------------------------------------------------------------
'            Prototypes for Win API functions used by more than 1 module
'
'       SHGetFolderPath - used by fnGetSpecialFolder( ) in modWinApi
'       ShellExecute    - used by fnOpenFileInDefaultApp() in modGeneral.bas
'------------------------------------------------------------------------
Public Declare Function SHGetFolderPath Lib "shfolder.dll" Alias "SHGetFolderPathA" _
    (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, _
     ByVal dwFlags As Long, ByVal pszPath As String) As Long
    
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    
    
' The following are used by the SHGetFolderPath function that's wrapped
' by fnGetSpecialFolder
Public Const CSIDL_ADMINTOOLS As Long = &H30             '{user}\Start Menu _
                                                         '\Programs\Administrative Tools
Public Const CSIDL_ALTSTARTUP As Long = &H1D             'non localized startup
Public Const CSIDL_APPDATA As Long = &H1A                '{user}\Application Data
Public Const CSIDL_BITBUCKET As Long = &HA               '{desktop}\Recycle Bin
Public Const CSIDL_CONTROLS As Long = &H3                'My Computer\Control Panel
Public Const CSIDL_COOKIES As Long = &H21
Public Const CSIDL_DESKTOP As Long = &H0                 '{namespace root}
Public Const CSIDL_DESKTOPDIRECTORY As Long = &H10       '{user}\Desktop
Public Const CSIDL_FAVORITES As Long = &H6               '{user}\Favourites
Public Const CSIDL_FONTS As Long = &H14                  'windows\fonts
Public Const CSIDL_HISTORY As Long = &H22
Public Const CSIDL_INTERNET As Long = &H1                'Internet virtual folder
Public Const CSIDL_INTERNET_CACHE As Long = &H20         'Internet Cache folder
Public Const CSIDL_LOCAL_APPDATA  As Long = &H1C&        '{user}\Local Settings\
                                                             '_Application Data (non roaming)
                                                             
Public Const CSIDL_DRIVES As Long = &H11                 'My Computer
Public Const CSIDL_MYPICTURES As Long = &H27             'C:\Program Files\My Pictures
Public Const CSIDL_NETHOOD As Long = &H13                '{user}\nethood
Public Const CSIDL_NETWORK As Long = &H12                'Network Neighbourhood

Public Const CSIDL_PRINTERS As Long = &H4                'My Computer\Printers
Public Const CSIDL_PRINTHOOD As Long = &H1B              '{user}\PrintHood
Public Const CSIDL_PERSONAL As Long = &H5                'My Documents

Public Const CSIDL_PROGRAM_FILES As Long = &H26          'Program Files folder
Public Const CSIDL_PROGRAM_FILESX86 As Long = &H2A       'Program Files folder for x86 apps (Alpha)
Public Const CSIDL_PROGRAMS As Long = &H2                'Start Menu\Programs
Public Const CSIDL_PROGRAM_FILES_COMMON As Long = &H2B   'Program Files\Common
Public Const CSIDL_PROGRAM_FILES_COMMONX86 As Long = &H2C 'x86 \Program Files\Common on RISC
Public Const CSIDL_RECENT As Long = &H8                  '{user}\Recent
Public Const CSIDL_SENDTO As Long = &H9                  '{user}\SendTo
Public Const CSIDL_STARTMENU As Long = &HB               '{user}\Start Menu
Public Const CSIDL_STARTUP As Long = &H7                 'Start Menu\Programs\Startup
Public Const CSIDL_SYSTEM As Long = &H25                 'system folder
Public Const CSIDL_SYSTEMX86 As Long = &H29              'system folder for x86 apps (Alpha)
Public Const CSIDL_TEMPLATES As Long = &H15
Public Const CSIDL_PROFILE As Long = &H28                'user's profile folder
Public Const CSIDL_WINDOWS As Long = &H24                'Windows directory or SYSROOT()

Public Const CSIDL_COMMON_ADMINTOOLS As Long = &H2F      '(all users)\Start Menu\ _
                                                         'Programs\Administrative Tools
Public Const CSIDL_COMMON_ALTSTARTUP As Long = &H1E      'non localized common startup
Public Const CSIDL_COMMON_APPDATA As Long = &H23         '(all users)\Application Data
Public Const CSIDL_COMMON_DESKTOPDIRECTORY As Long = &H19 '(all users)\Desktop
Public Const CSIDL_COMMON_DOCUMENTS As Long = &H2E       '(all users)\Documents
Public Const CSIDL_COMMON_FAVORITES As Long = &H1F       '(all users)\Favourites
Public Const CSIDL_COMMON_PROGRAMS As Long = &H17        '(all users)\Programs
Public Const CSIDL_COMMON_STARTMENU As Long = &H16       '(all users)\Start Menu
Public Const CSIDL_COMMON_STARTUP As Long = &H18         '(all users)\Startup
Public Const CSIDL_COMMON_TEMPLATES As Long = &H2D       '(all users)\Templates

Public Const CSIDL_FLAG_CREATE = &H8000&          'combine with CSIDL_ value to force
                                                  'create on SHGetSpecialFolderLocation()
Public Const CSIDL_FLAG_DONT_VERIFY = &H4000      'combine with CSIDL_ value to force
                                                  'create on SHGetSpecialFolderLocation()
Public Const CSIDL_FLAG_MASK = &HFF00             'mask for all possible flag values
Public Const SHGFP_TYPE_CURRENT = 0               'current value for user, verify it exists
Public Const SHGFP_TYPE_DEFAULT = 1
Public Const MAX_PATH = 260
Public Const S_OK = &H0                           'Success
Public Const S_FALSE = &H1                        'The folder is valid, but does not exist
Private Const E_INVALIDARG = &H80070057           'Invalid CSIDL Value

'SQL_INTEGRATED_SECURITY
'
' Win32 APIs to determine OS information.
'
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type
Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

'
' Win32 NetAPIs.
'
Private Const USERNAME_LENGTH = 256         ' Maximum username length
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserNameW Lib "advapi32.dll" (lpBuffer As Byte, nSize As Long) As Long
'SQL_INTEGRATED_SECURITY



'SQL_INTEGRATED_SECURITY  - Added
'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function fnGetNetworkUser() As String
    ' Comments  : Gets the User ID that is currently logged on to the network
    ' Parameters: None
    ' Returns   : User ID
    ' Source    : Karl Peterson's Classic VB site
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc       As String = "fnGetNetworkUser"
    Const clngNameLength = USERNAME_LENGTH + 1
    Dim objOS                   As OSVERSIONINFO
    Dim strBuffer               As String
    Dim bytBuffer()             As Byte
    Dim lngRetVal               As Long
    Dim lngLength               As Long

   objOS.dwOSVersionInfoSize = Len(objOS)
   Call GetVersionEx(objOS)
   
   If objOS.dwPlatformId = VER_PLATFORM_WIN32_NT Then
        lngLength = clngNameLength * 2
        ReDim bytBuffer(0 To lngLength - 1) As Byte
        If GetUserNameW(bytBuffer(0), lngLength) Then
            strBuffer = bytBuffer
            fnGetNetworkUser = Left(strBuffer, lngLength - 1)
        End If
    End If
PROC_EXIT:
    On Error GoTo 0     ' disablfe error handler

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
'SQL_INTEGRATED_SECURITY - Added



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function fnGetSpecialFolder(hwndOwner As Long, CSIDL As Long) As String
    ' Comments  : Get the fully qualified path to one of Windows' special folders.
    '             This approach is what is recommended for Win2K and all previous
    '             versions of Windows. Be sure to include " Or CSIDL_FLAG_CREATE"
    '             if you want the folder created if it doesn't already exist.
    '
    '             Using it requires that SHFOLDER.DLL be distributed if the app
    '             will be used on pre-Win2K versions of the Windows OS. This is
    '             available as a redistributable within the Platform SDK.
    '
    '             Much of this code was lifted from MSKB article Q252652.
    '
    ' Parameters: hWndOwner - handle to a window (0 if not needed)
    '             CSIDL - the CSIDL indicating which folder path to return.
    '
    ' Called by : fnLogOpen( ) in modAppLog
    '             fnLogPrune( ) in modAppLog
    '
    ' Returns   : Directory name, with appended "\" if necessary
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnGetSpecialFolder"
    Dim strPath As String
    Dim lngRetVal As Long
  
    ' Fill our string buffer
    strPath = String(MAX_PATH, 0)
   
    lngRetVal = SHGetFolderPath(hwndOwner, CSIDL, 0&, SHGFP_TYPE_CURRENT, strPath)
   
    Select Case lngRetVal
        Case S_OK
            ' We retrieved the folder successfully.
            ' All C strings are null-terminated, so return the string up to the
            ' first null character
            fnGetSpecialFolder = Left$(strPath, InStr(1, strPath, Chr$(0)) - 1)
    Case S_FALSE
            ' The CSIDL in the 2nd argument is valid, but the folder does not exist.
            ' Use CSIDL_FLAG_CREATE to have it created automatically
'!TODO! Gen msg via frmMsgBox
            'fnProcessFatalError Err.Source, _
            '                    fte_OtherErrType, Err.Number, _
            '                    Err.Description, Err.Source, _
            '                    Err.HelpFile, Err.HelpContext, _
            '                    "The specified folder ( " & CStr(CSIDL) & ") does not exist. " & _
            '                    "Add the CSIDL_FLAG_CREATE flag to create it. RC = " & CStr(lngRetVal)
    Case Else
            ' E_INVALIDARG...CSIDL in the 2nd argument is invalid
'!TODO! Gen msg via frmMsgBox
            'fnProcessFatalError Err.Source, _
            '                    fte_OtherErrType, Err.Number, _
            '                    Err.Description, Err.Source, _
            '                    Err.HelpFile, Err.HelpContext, _
            '                    "An invalid CSIDL argument (" & CStr(CSIDL) & ") was " & _
            '                    "passed to this function. RC = " & CStr(lngRetVal)
    End Select
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
