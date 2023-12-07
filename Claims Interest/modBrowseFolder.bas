Attribute VB_Name = "modBrowseFolder"
' modBrowseFolder
'
' Use this module in conjunction with class module: cbrfBrowseFolder.cls

Option Explicit

' Browse Flags
Public Const BIF_RETURNONLYFSDIRS = &H1
Public Const BIF_DONTGOBELOWDOMAIN = &H2
Public Const BIF_STATUSTEXT = &H4
Public Const BIF_RETURNFSANCESTORS = &H8
Public Const BIF_BROWSEFORCOMPUTER = &H1000
Public Const BIF_BROWSEFORPRINTER = &H2000

'From MSDN help on BROWSEINFO:

' Flags specifying the options for the dialog box. This member can include zero or a
' combination of the following values:
'    BIF_BROWSEFORCOMPUTER  Only return computers. If the user selects anything other
'                           than a computer, the OK button is grayed.
'    BIF_BROWSEFORPRINTER   Only return printers. If the user selects anything other
'                           than a printer, the OK button is grayed.
'    BIF_BROWSEINCLUDEFILES Version 4.71. The browse dialog will display files as well as folders.
'    BIF_BROWSEINCLUDEURLS  Version 5.0. The browse dialog box can display URLs.
'                           The BIF_USENEWUI and BIF_BROWSEINCLUDEFILES flags must also be set.
'                           If these three flags are not set, the browser dialog box will reject URLs.
'                           Even when these flags are set, the browse dialog box will only display URLs
'                           if the folder that contains the selected item supports them.
'                           When the folder's IShellFolder::GetAttributesOf method is called to request
'                           the selected item's attributes, the folder must set the SFGAO_FOLDER
'                           attribute flag. Otherwise, the browse dialog box will not display the URL.
'    BIF_DONTGOBELOWDOMAIN  Do not include network folders below the domain level in the
'                           dialog box's tree view control.
'    BIF_EDITBOX            Version 4.71. Include an edit control in the browse dialog box
'                           that allows the user to type the name of an item.
'    BIF_NEWDIALOGSTYLE     Version 5.0. Use the new user interface. Setting this flag provides
'                           the user with a larger dialog box that can be resized.
'                           The dialog box has several new capabilities including:
'                           drag and drop capability within the dialog box, reordering,
'                           context menus, new folders, delete, and other context menu commands.
'                           To use this flag, you must call OleInitialize or CoInitialize
'                           before calling SHBrowseForFolder.
'    BIF_RETURNFSANCESTORS  Only return file system ancestors. An ancestor is a subfolder
'                           that is beneath the root folder in the namespace hierarchy.
'                           If the user selects an ancestor of the root folder that is not
'                           part of the file system, the OK button is grayed.
'    BIF_RETURNONLYFSDIRS   Only return file system directories. If the user selects folders
'                           that are not part of the file system, the OK button is grayed.
'    BIF_SHAREABLE          Version 5.0. The browse dialog box can display shareable resources
'                           on remote systems. It is intended for applications that want to
'                           expose remote shares on a local system. The BIF_USENEWUI flag must also be set.
'    BIF_STATUSTEXT         Include a status area in the dialog box. The callback function
'                           can set the status text by sending messages to the dialog box.
'    BIF_USENEWUI           Version 5.0. Use the new user interface, including an edit box.
'                           This flag is equivalent to BIF_EDITBOX | BIF_NEWDIALOGSTYLE.
'                           To use BIF_USENEWUI, you must call OleInitialize or CoInitialize
'                           before calling SHBrowseForFolder.
'    BIF_VALIDATE           Version 4.71. If the user types an invalid name into the edit box,
'                           the browse dialog will call the application's BrowseCallbackProc
'                           with the BFFM_VALIDATEFAILED message.
'                           This flag is ignored if BIF_EDITBOX is not specified.

Private Const BFFM_INITIALIZED = 1
Private Const WM_USER = &H400
Private Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
Private Const BFFM_SETSELECTIONW As Long = (WM_USER + 103)
    
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    
    
'////////////////////////////////////////////////////////////////////////////////////////
Public Function BrowseCallbackProcStr(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
    'Callback for the Browse STRING method.
     ' On initialization, set the dialog's pre-selected folder from the pointer
    ' to the path allocated as bi.lParam, passed back to the callback as lpData param.
 
    Select Case uMsg
        Case BFFM_INITIALIZED
            Call SendMessage(hWnd, BFFM_SETSELECTIONA, True, ByVal lpData)
    End Select
End Function



'////////////////////////////////////////////////////////////////////////////////////////
Public Function FARPROC(ByVal pfn As Long) As Long
    ' A dummy procedure that receives and returns
    ' the value of the AddressOf operator.
 
    ' Obtain and set the address of the callback
    ' This workaround is needed as you can't assign
    ' AddressOf directly to a member of a user-
    ' defined type, but you can assign it to another
    ' long and use that (as returned here)

    ' From Randy Birch 2000/12/17
    ' Matt (Curland) correctly pointed out that in passing the addressof via a
    ' wrapper routine, we really *do* want to pass the real address, and not a
    ' reference. Added ByVal to above function [ByVal pfn As Long]
 
    FARPROC = pfn
End Function
