Attribute VB_Name = "modTextBox"
'******************************************************************************
' Module      : modTextBox
' Procedures  :
'               fnCanUndoTextBox(txtIn As TextBox) As Boolean
'               fnCRToTab( ByVal intKeyAscii As Integer) As Integer
'               fnFilterKeyPressAlpha( ByVal intKeyAscii As Integer) As Integer
'               fnFilterKeyPressNumeric( ByVal intKeyAscii As Integer, _
'                   Optional ByVal fIntegerOnly As Boolean = False) As Integer
'               fnFilterKeyPressValues( ByVal intKeyAscii As Integer, _
'                   ParamArray varpaKeys() As Variant) As Integer
'               fnGetTextBoxLine( txtIn As TextBox, _
'                   ByVal intLineNumber As Integer) As String
'               fnGetTextBoxLineCount(txtIn As TextBox) As Integer
'               fnSelectText(txtIn As TextBox)
'               fnSetTextBoxCase( txtIn As TextBox, Optional ByVal intMode As Integer = 1)
'               fnSetTextBoxRect( txtIn As TextBox, ByVal lngLeft As Long, _
'                   ByVal lngTop As Long, ByVal lngRight As Long, ByVal lngBottom As Long)
'               fnSetTextBoxTabStops( txtIn As TextBox, ParamArray varpaTabStops() As Variant)
'               fnTextBoxFromDisk( txtIn As TextBox, strFileName As String)
'               fnTextBoxToDisk( txtIn As TextBox, strFileName As String)
'               fnUndoTextBox(txtIn As TextBox)
'
' Description : Routines to extend the functionality of the
'               standard VB TextBox control
' Source      : Total Visual SourceBook 2000
' --------------------------------------------------
Option Explicit
Option Compare Binary

Private Const mcstrName             As String = "modTextBox."

Public Const gcintForceLowerCase As Integer = 0
Public Const gcintForceUpperCase As Integer = 1

Private Declare Function SendMessage _
    Lib "user32" _
    Alias "SendMessageA" _
    (ByVal hwnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        lParam As Any) _
    As Long

Private Declare Function SetWindowLong _
    Lib "user32" _
    Alias "SetWindowLongA" _
    (ByVal hwnd As Long, _
        ByVal nIndex As Long, _
        ByVal dwNewLong As Long) _
    As Long

Private Declare Function GetWindowLong _
    Lib "user32" _
    Alias "GetWindowLongA" _
    (ByVal hwnd As Long, _
        ByVal nIndex As Long) _
    As Long

Private Const GWL_EXSTYLE As Long = (-20)
Private Const GWL_STYLE As Long = (-16)
Private Const GWL_WNDPROC As Long = (-4)

Private Const ES_AUTOHSCROLL As Long = &H80&
Private Const ES_AUTOVSCROLL As Long = &H40&
Private Const ES_CENTER As Long = &H1&
Private Const ES_LEFT As Long = &H0&
Private Const ES_LOWERCASE As Long = &H10&
Private Const ES_UPPERCASE As Long = &H8&
Private Const ES_RIGHT As Long = &H2&
Private Const ES_WANTRETURN As Long = &H1000& 'won't work

Private Const WM_USER = &H400
Private Const EM_GETLINECOUNT = &HBA
Private Const EM_LINEINDEX = &HBB
Private Const EM_LINELENGTH = &HC1
Private Const EM_GETLINE = &HC4
Private Const EM_CANUNDO = &HC6
Private Const EM_EMPTYUNDOBUFFER = &HCD
Private Const EM_UNDO = &HC7
Private Const EM_SETTABSTOPS = &HCB

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Const EM_SETRECT = &HB3



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function fnCanUndoTextBox(txtIn As TextBox) As Boolean
    ' Comments  : Determines whether or not there is something to 'undo'
    '             in the specified text box control. Useful for disabling
    '             Undo buttons or menu choices if there is nothing to undo.
    ' Parameters: txtIn - text box to check
    ' Returns   : true if there is something to undo
    '             false if there is nothing to undo
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "fnCanUndoTextBox"
    Dim fResult As Boolean

    On Error GoTo PROC_ERR

    fResult = CBool(SendMessage(txtIn.hwnd, EM_CANUNDO, 0&, ByVal 0))
    fnCanUndoTextBox = fResult
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
Public Function fnCRToTab(ByVal intKeyAscii As Integer) As Integer
    ' Comments  : Allows user to press the Return key to tab
    '             between text boxes
    ' Parameters: intKeyAscii - character keycode to test
    ' Returns   : If the passed value is a carriage return,
    '             eat the keystroke, and send a tab character to
    '             the input stream to move focus to the next
    '             control in the tab order.
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "fnCRToTab"
    On Error GoTo PROC_ERR

    If intKeyAscii = vbKeyReturn Then
        SendKeys "{Tab}"
        fnCRToTab = 0
    Else
        fnCRToTab = intKeyAscii
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
Public Function fnFilterKeyPressAlpha(ByVal intKeyAscii As Integer) As Integer
    ' Comments  : Test passed key value to see whether it is an alphabetic
    '             character. If it is, return the actual value. If it is not,
    '             return a zero.
    '             Assign the return value of this function to the
    '             KeyAscii argument of the xx_KeyPress() event
    ' Parameters: intKeyAscii - character keycode to test
    ' Returns   : if the passed value is alphabetic, returns 0
    '             otherwise it returns the same value passed in
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "fnFilterKeyPressAlpha"
    Dim lngKeyReturn As Long

    On Error GoTo PROC_ERR

    lngKeyReturn = intKeyAscii

    'allow backspace key
    If intKeyAscii <> vbKeyBack Then
        If intKeyAscii < vbKeyA Then

            lngKeyReturn = 0

        Else
            If intKeyAscii > vbKeyZ And intKeyAscii < (vbKeyA + 32) Then
                lngKeyReturn = 0
            Else
                If intKeyAscii > (vbKeyZ + 32) Then
                    lngKeyReturn = 0
                End If
            End If

        End If
    End If

    fnFilterKeyPressAlpha = lngKeyReturn
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
Public Function fnFilterKeyPressNumeric(ByVal intKeyAscii As Integer, _
    Optional ByVal fIntegerOnly As Boolean = False) As Integer
    ' Comments  : Test passed key value to see whether it is a numeric
    '             character. If it is, return the actual value. If it is not,
    '             return a zero.
    '             Assign the return value of this function to the
    '             KeyAscii argument of the xx_KeyPress() event
    ' Parameters: intKeyAscii - character keycode to test
    '             fIntegerOnly - if true, allows only actual
    '             numbers between 0 and 9. if false, allows - and .
    ' Returns   : if the passed value is numeric, returns 0
    '             otherwise it returns the same value passed in
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "fnFilterKeyPressNumeric"
    Dim intKeyReturn As Integer

    On Error GoTo PROC_ERR

    intKeyReturn = intKeyAscii

    'allow backspace key
    If intKeyAscii <> vbKeyBack Then

        If fIntegerOnly = False Then
            If intKeyAscii <> vbKeyInsert And intKeyAscii <> vbKeyDelete Then
                If intKeyAscii < vbKey0 Or intKeyAscii > vbKey9 Then
                    intKeyReturn = 0
                End If
            End If
        Else
            If intKeyAscii < vbKey0 Or intKeyAscii > vbKey9 Then
                intKeyReturn = 0
            End If

        End If
    End If

    fnFilterKeyPressNumeric = intKeyReturn
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
Public Function fnFilterKeyPressValues(ByVal intKeyAscii As Integer, _
    ParamArray varpaKeys() As Variant) As Integer
    ' Comments  : Test passed key value and a list of valid key codes
    '             to see whether the character is in that list.
    '             If it is, return the actual value. If it is not,
    '             return a zero.
    '             Assign the return value of this function to the
    '             KeyAscii argument of the xx_KeyPress() event
    ' Parameters: intKeyAscii - character keycode to test
    '             varpaKeys() - array of keycodes to allow
    ' Returns   : if the passed value is in the list of valid values,
    '             returns 0. Otherwise it returns the same value passed in
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "fnFilterKeyPressValues"
    Dim varCheckKey As Variant
    Dim intKeyReturn As Integer

    On Error GoTo PROC_ERR

    'return 0 if key code is not in the list
    intKeyReturn = 0

    'test each element
    For Each varCheckKey In varpaKeys

        'compare passed key value to current element of paramarray
        If intKeyAscii = varCheckKey Then

            'return the original key value
            intKeyReturn = intKeyAscii

            'no need to check further
            Exit For
        End If
    Next varCheckKey

    fnFilterKeyPressValues = intKeyReturn
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
Public Function fnGetTextBoxLine(txtIn As TextBox, _
    ByVal intLineNumber As Integer) As String
    ' Comments  : Return a line of text from a multi-line text box
    ' Parameters: txtIn - Text box to check
    '             intLineNumber - the desired line number to return
    ' Returns   : A string containing the specified line
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "fnGetTextBoxLine"
    Dim lngNumLines As Long
    Dim strTest As String
    Dim lngTestStringLength As Long
    Dim intLineIndex As Integer
    Dim intLineLength As Integer

    On Error GoTo PROC_ERR

    'Get the number of lines in the TextBox
    lngNumLines = SendMessage(txtIn.hwnd, EM_GETLINECOUNT, 0, ByVal 0&)

    If intLineNumber >= 0 And intLineNumber < lngNumLines Then

        'Get character offset of the first character of the line
        intLineIndex = _
            SendMessage(txtIn.hwnd, EM_LINEINDEX, intLineNumber, ByVal 0&)

        'Get the line length of that line
        intLineLength = _
            SendMessage(txtIn.hwnd, EM_LINELENGTH, intLineIndex, ByVal 0&) + 1

        'Initialize a String buffer to hold the line. Allow 2 characters
        'for the line length information
        strTest = String$(intLineLength + 2, 0)                             '

        'Put the line length into the first two bytes of the string
        Mid$(strTest, 1, 1) = Chr$(intLineLength And &HFF)
        Mid$(strTest, 2, 1) = Chr$(intLineLength / &H100)

        'Get selected line, save the length of the returned string
        lngTestStringLength = _
            SendMessage(txtIn.hwnd, EM_GETLINE, intLineNumber, ByVal strTest)

        'remove trailing nulls
        strTest = Left$(strTest, lngTestStringLength)

        fnGetTextBoxLine = strTest
    Else
        Err.Raise vbObjectError + 1, _
            App.Title, "Invalid line number specified."
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
Public Function fnGetTextBoxLineCount(txtIn As TextBox) As Integer
    ' Comments  : Determines the number of lines in a multi-line text box
    ' Parameters: txtIn - the text box to check
    ' Returns   : Number of lines
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "fnGetTextBoxLineCount"
    Dim lngNumLines As Long

    On Error GoTo PROC_ERR

    lngNumLines = SendMessage(txtIn.hwnd, EM_GETLINECOUNT, 0, ByVal 0&)
    fnGetTextBoxLineCount = lngNumLines
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
Public Sub fnSelectText(txtIn As TextBox)
    ' Comments  : Selects all text in the passed text box. Call this
    '             sub from the GotFocus event of the text box
    ' Parameters: txtIn - reference to text box
    ' Returns   : Nothing
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "fnSelectText"
    On Error GoTo PROC_ERR

    With txtIn
        .SelStart = 0
        .SelLength = Len(.Text)
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
Public Sub fnSetTextBoxCase(txtIn As TextBox, Optional ByVal intMode As Integer = 1)
    ' Comments  : Forces all new entry into the text box to be either
    '             uppercase or lower-case
    ' Parameters: txtIn - the text box to change
    '             intMode - 1 (gcintForceUpperCase) to force new text to uppercase.
    '                       0 (gcintForceLowerCase) to force new text to lowercase
    ' Returns   : Nothing
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "fnSetTextBoxCase"
    Dim lngResult As Long
    Dim lngStyle As Long
    Dim lnghWnd As Long

    On Error GoTo PROC_ERR

    lnghWnd = txtIn.hwnd
    lngStyle = GetWindowLong(lnghWnd, GWL_STYLE)

    Select Case intMode
        Case 0
            lngStyle = lngStyle And Not ES_UPPERCASE
            lngStyle = lngStyle Or ES_LOWERCASE

        Case 1
            'add upper case bits
            lngStyle = lngStyle And Not ES_LOWERCASE
            lngStyle = lngStyle Or ES_UPPERCASE

    End Select

    lngResult = SetWindowLong(lnghWnd, GWL_STYLE, lngStyle)
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
Public Sub fnSetTextBoxRect(txtIn As TextBox, ByVal lngLeft As Long, _
    ByVal lngTop As Long, ByVal lngRight As Long, ByVal lngBottom As Long)
    ' Comments  : Create a bounding rectangle for a text box that is
    '             different from its actual dimensions. Useful for
    '             controlling line wrapping.
    '             Textbox must be set to multi-line mode
    '             The dimensions are specified in PIXELS
    ' Parameters: txtIn - text box to modify
    '             lngLeft - Left position of text box in pixels
    '             lngTop  - Top position of text box in pixels
    '             lngRight - Right position of text box in pixels
    '             lngBottom - Bottom position of text box in pixels
    ' Returns   : Nothing
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "fnSetTextBoxRect"
    Dim lngReturn As Long
    Dim rectSize As RECT

    On Error GoTo PROC_ERR

    With rectSize
        .Left = lngLeft
        .Top = lngTop
        .Right = lngRight
        .Bottom = lngBottom
    End With

    lngReturn = SendMessage(txtIn.hwnd, EM_SETRECT, 0&, rectSize)
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
Public Sub fnSetTextBoxTabStops(txtIn As TextBox, ParamArray varpaTabStops() As Variant)
    ' Comments  : Set tab stops for a text box
    ' Parameters: txttIn - the text box to modify
    '             varpaTabStops() - a list of tab stop positions, specified
    '             in pixels
    ' Returns   : Nothing
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "fnSetTextBoxTabStops"
    Dim varItem As Variant
    Dim intCount As Integer
    Dim lngResult As Long

    On Error GoTo PROC_ERR

    ' get number of param array items passed
    intCount = UBound(varpaTabStops)

    ' size the array which will be passed to SendMessage
    ReDim alngTabStops(intCount) As Long

    ' populate the array with the values in the Param Array
    intCount = 0
    For Each varItem In varpaTabStops
        alngTabStops(intCount) = CLng(varItem)
        intCount = intCount + 1
    Next varItem

    ' set the tab stops.
    lngResult = SendMessage( _
        txtIn.hwnd, _
        EM_SETTABSTOPS, _
        UBound(alngTabStops), _
        alngTabStops(0))

    txtIn.Refresh
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
Public Sub fnTextBoxFromDisk(txtIn As TextBox, strFileName As String)
    ' Comments  : Loads the contents of a text file into a text box.
    '             If the file is too long to put into a text box, it
    '             is truncated.
    ' Parameters: txtIn - The text box to load
    '             strFileName - The file name to load
    ' Returns   : Nothing
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "fnTextBoxFromDisk"
    Dim lngFileHandle As Long
    Dim lngLOF As Long
    Dim strTemp As String

    On Error GoTo PROC_ERR

    lngFileHandle = FreeFile

    ' Open the specified file
    Open strFileName For Input Access Read Shared As lngFileHandle

    ' Test its length
    lngLOF = LOF(lngFileHandle)

    ' Truncate if too long to hold in a Windows 95 text box
    If lngLOF > 32767 Then
        lngLOF = 32767
    End If

    ' Get the text from the file
    strTemp = Input(lngLOF, lngFileHandle)

    ' Assign it to the text box
    txtIn.Text = strTemp

    Close lngFileHandle
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
Public Sub fnTextBoxToDisk(txtIn As TextBox, strFileName As String)
    ' Comments  : Writes the contents of a text box to a disk file
    ' Parameters: txtIn - The text box to use for the text
    '             strFileName - The full path name to the file you
    '             wish to create. The file will be overwritten if it
    '             already exists
    ' Returns   : Nothing
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "fnTextBoxToDisk"
    Dim lngFileHandle As Long

    On Error GoTo PROC_ERR

    ' Get handle for file operation
    lngFileHandle = FreeFile

    ' Open the output file
    Open strFileName For Output As lngFileHandle

    ' Send contents to the file
    Print #lngFileHandle, txtIn.Text

    Close lngFileHandle
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
Public Sub fnUndoTextBox(txtIn As TextBox)
    ' Comments  : Undoes the last typing action in a text box
    ' Parameters: txtIn - the text box to undo
    ' Returns   : Nothing
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "fnUndoTextBox"
    Dim lngResult As Long

    On Error GoTo PROC_ERR

    lngResult = SendMessage(txtIn.hwnd, EM_UNDO, 0&, ByVal 0)
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
