VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cbrfBrowseFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' cbrfBrowseFolder
' 2000/12/17 Copyright © 2000, Larry Rebich, using the VAIO
' 2000/12/17 larry@buygold.net, www.buygold.net, 760.771.4730
'            Some parts of this code from Randy Birch and others.
'
' Use this class module in conjunction with module: modBrowseFolder.bas

    Option Explicit
    DefLng A-Z

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Some parts copyright ©1996-2000 VBnet, Randy Birch, All Rights Reserved.
' http://www.mvps.org/vbnet/index.html
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
Private Const mcstrName             As String = "cbrfBrowseFolder."

Private Type BrowseInfo
   hOwner           As Long
   pidlRoot         As Long
   pszDisplayName   As String
   lpszTitle        As String
   uFlags           As Long
   lpfn             As Long
   lParam           As Long
   iImage           As Long
End Type

Private BrowseInfo As BrowseInfo    'allocate and name storage for the structure

Private Declare Function SHGetPathFromIDList _
    Lib "shell32.dll" Alias "SHGetPathFromIDListA" _
    (ByVal pidl As Long, ByVal pszPath As String) As Long
    
Private Declare Function SHBrowseForFolder Lib "shell32.dll" _
    Alias "SHBrowseForFolderA" (lpBrowseInfo As BrowseInfo) As Long
    
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
   
Private Declare Function SHSimpleIDListFromPath Lib "shell32" Alias "#162" (ByVal szPath As String) As Long
    
Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
    
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
    
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
    
Private Const LMEM_FIXED = &H0
Private Const LMEM_ZEROINIT = &H40
Private Const LPTR = (LMEM_FIXED Or LMEM_ZEROINIT)
   
Private m_lngHwnd      As Long
Private m_strTitle     As String
Private m_strFolder    As String
Private m_bCancelled   As Boolean
Private m_lngFlags     As Long


'////////////////////////////////////////////////////////////////////////////////////////
Public Property Let Flags(ByVal Flags As Long)
    m_lngFlags = Flags
End Property



'////////////////////////////////////////////////////////////////////////////////////////
Public Property Get Cancelled() As Boolean
    Cancelled = m_bCancelled
End Property



'////////////////////////////////////////////////////////////////////////////////////////
Public Property Let hWnd(ByVal lnghWnd As Long)
    m_lngHwnd = lnghWnd
End Property



'////////////////////////////////////////////////////////////////////////////////////////
Public Property Let Title(ByVal strValue As String)
    m_strTitle = strValue
End Property



'////////////////////////////////////////////////////////////////////////////////////////
Public Property Let Folder(ByVal strValue As String)
    m_strFolder = strValue
End Property



'////////////////////////////////////////////////////////////////////////////////////////
Public Function ShowBrowse(Optional ByVal varHwnd As Variant, Optional ByVal varTitle As Variant, _
    Optional ByVal varFolder As Variant) As String
    '--------------------------------------------------------------------------
    ' Procedure:   ShowBrowse
    ' Description: Opens a Browse for Folder dialog box so the user can
    '              select an existing folder.
    ' Returns:     Folder name that was selected
    ' Params:      varHwnd - handle to the calling form
    '              varTitle - title to show on the Browse for Folder screen
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc         As String = "ShowBrowse"
    Dim lngSelPath    As Long
    Dim lngRtn        As Long
    Dim lngPIDL       As Long
    Dim intPosition   As Integer
    Dim strFolder     As String * MAX_PATH
    
    On Error GoTo PROC_ERR
    
    ' Allow user to specify params in the call
    If Not IsMissing(varHwnd) Then
        m_lngHwnd = varHwnd
    End If
    If Not IsMissing(varTitle) Then
        m_strTitle = varTitle
    End If
    If Not IsMissing(varFolder) Then
        m_strFolder = varFolder
    End If
    
    With BrowseInfo
        .hOwner = m_lngHwnd                                   'owner's hWnd
        .uFlags = m_lngFlags                                  'flags, default is BIF_RETURNONLYFSDIRS
        .lpszTitle = m_strTitle                               'dialog's title
        .lpfn = FARPROC(AddressOf BrowseCallbackProcStr)      'set to pass an address into a structure
        lngSelPath = LocalAlloc(LPTR, Len(m_strFolder))
        MoveMemory ByVal lngSelPath, ByVal m_strFolder, Len(m_strFolder)
        .lParam = lngSelPath                                  'now into structure
    End With
    
    ' Show Browse for Folder window
    lngPIDL = SHBrowseForFolder(BrowseInfo)
    
    If lngPIDL Then
        ' If not cancelled, translate the PIDL into a folder name, drop the characters in that
        ' folder name that occur after the Null character, and then free memory.
        If SHGetPathFromIDList(lngPIDL, strFolder) Then
            strFolder = Left$(strFolder, InStr(strFolder, vbNullChar) - 1)
        End If
        Call CoTaskMemFree(lngPIDL)
    Else
        m_bCancelled = True
    End If
    
    ' Free memory allocated for SelPath
    Call LocalFree(lngSelPath)

    ' Update member variable with user's selection
    m_strFolder = strFolder
    
    If m_strFolder = vbNullString Then
        ' If the user clicked Cancel in the Browse For Folder window, then inform caller of that
        m_bCancelled = True
    Else
        ' Otherwise, return the name of the new selected folder
        ShowBrowse = m_strFolder
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



'////////////////////////////////////////////////////////////////////////////////////////
Private Sub Class_Initialize()
    '--------------------------------------------------------------------------
    ' Procedure:   Class_Initialize
    ' Description: Instantiates this class, setting default values as appropriate
    '
    ' Returns:     N/A
    ' Params:      N/A
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc         As String = "Class_Initialize"
    
    On Error GoTo PROC_ERR
    
    m_lngFlags = BIF_RETURNONLYFSDIRS    ' Set to default: return only real drives, no virtuals
    m_strTitle = "Select a Folder"
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
