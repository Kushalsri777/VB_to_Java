VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cetbExtendedTextbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------
' Procedure:   cetbExtendedTextbox
' Description: Uses multicasting to ensure TextBoxes bound to this class
'              have standard behavior.
'
' Revisions:
'   1.0 BAW Initial creation
'
' Procedures  :
'   Private     Class_Initialize()
'   Public      Property Let Format(ByVal strIn As String)
'   Public      Property Let IsNumeric(ByVal bIn As Boolean)
'   Public      Property Let ShouldForceToUppercase(ByVal bIn As Boolean)
'   Public      Property Let ValidKeys(ByVal strIn As String)
'   Public      BindControl(ByRef txtIn As TextBox)
'   Private     myTextBox_GotFocus()
'   Private     myTextBox_KeyDown(ByRef pintKeyCode As Integer, ByRef pintShift As Integer)
'   Private     myTextBox_KeyPress(ByRef pintKeyAscii As Integer)
'   Private     myTextBox_LostFocus()
'
'-----------------------------------------------------------------------------
Option Explicit
Option Compare Binary

Private Const mcstrName As String = "cetbExtendedTextbox."

'...............................................................................................
' These are the private variables corresponding to PUBLIC properties.
'...............................................................................................
Private m_bIsNumeric As Boolean
Private m_bShouldForceToUppercase As Boolean
Private m_strFormat As String
Private m_strValidKeys As String

'...............................................................................................
' Other private variables that do NOT correspond to PUBLIC properties.
'...............................................................................................
' myTextBox is the object that will intercept the events from a form's "real" TextBox control
Private WithEvents myTextBox As VB.TextBox
Attribute myTextBox.VB_VarHelpID = -1



'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'\                                                                  /
'|          CLASS_INITIALIZE / CLASS_TERMINATE    Procedures        |
'/                                                                  \
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/


'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub Class_Initialize()
    '----------------------------------------------------------------------------
    ' Procedure   :  Class_Initialize (Private)
    '
    ' Comments    : Initializes class properties to their default values.
    '
    ' Parameters  : N/A
    '
    ' Modified
    '
    '----------------------------------------------------------------------------
    Const cstrCurrentProc As String = "Class_Initialize"
    On Error GoTo PROC_ERR
    
    m_bIsNumeric = False
    m_bShouldForceToUppercase = False
    m_strValidKeys = vbNullString
    m_strFormat = vbNullString
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
Public Property Let Format(ByVal strIn As String)
    '----------------------------------------------------------------------------
    ' Procedure   :  Property Format   (Public)
    '
    ' Comments    : The Format property indicates how the contents of a
    '               textbox bound to this class should be formatted for
    '               display purposes.
    '
    ' Parameters  :
    '    strIn (in) = a string, as a VB control's Format property could use
    '
    ' Called by   : The Form_Load event handler of any form that has a control
    '               of type TextBox
    '
    ' Modified
    '
    '----------------------------------------------------------------------------
    Const cstrCurrentProc As String = "Property Let Format"
    On Error GoTo PROC_ERR

    m_strFormat = strIn
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
Public Property Let IsNumeric(ByVal bIn As Boolean)
    '----------------------------------------------------------------------------
    ' Procedure   :  Property IsNumeric   (Public)
    '
    ' Comments    : The IsNumeric property indicates whether the contents of a
    '               textbox bound to this class should be a numeric value. This info
    '               is used to drive how the Format property is applied in the
    '               LostFocus event.
    '
    ' Parameters  :
    '    bIn (in) = True if the contents should be a numeric value; False otherwise
    '
    ' Called by   : The Form_Load event handler of any form that has a control
    '               of type TextBox
    '
    ' Modified
    '
    '----------------------------------------------------------------------------
    Const cstrCurrentProc As String = "Property Let IsNumeric"
    On Error GoTo PROC_ERR

    m_bIsNumeric = bIn
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
Public Property Let ShouldForceToUppercase(ByVal bIn As Boolean)
    '----------------------------------------------------------------------------
    ' Procedure   :  Property ShouldForceToUppercase   (Public)
    '
    ' Comments    : The ShouldForceToUppercase property indicates whether the contents of a
    '               textbox bound to this class should be forced to uppercase when its
    '               value has changed.
    '
    ' Parameters  :
    '    bIn (in) = True if the contents should be forced to uppercase; False otherwise
    '
    ' Called by   : The Form_Load event handler of any form that has a control
    '               of type TextBox
    '
    ' Modified
    '
    '----------------------------------------------------------------------------
    Const cstrCurrentProc As String = "Property Let ShouldForceToUppercase"
    On Error GoTo PROC_ERR

    m_bShouldForceToUppercase = bIn
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
Public Property Let ValidKeys(ByVal strIn As String)
    '----------------------------------------------------------------------------
    ' Procedure   :  Property ValidKeys   (Public)
    '
    ' Comments    : The ValidKeys property indicates which characters can be entered
    '               into a textbox bound to this class.
    '
    ' Parameters  :
    '    strIn (in) = the value to assign to this property
    '
    ' Called by   : The Form_Load event handler of any form that has a control
    '               of type TextBox
    '
    ' Modified
    '
    '----------------------------------------------------------------------------
    Const cstrCurrentProc As String = "Property Let ValidKeys"
    On Error GoTo PROC_ERR

    m_strValidKeys = strIn
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
Public Sub BindControl(ByRef txtIn As TextBox)
Attribute BindControl.VB_Description = "This procedure is used by the form to ""bind"" one of its textboxes to this class, so this class (through multicasting) can extend the default functionality of that textbox without bogging down each form with the same code for each control of type TextBox."
    '----------------------------------------------------------------------------
    ' Procedure   :  BindControl
    '
    ' Comments    : This procedure is used by the form to "bind" one of its textboxes
    '               to this class, so this class (through multicasting) can extend the
    '               default functionality of that textbox without bogging down each form
    '               with the same code for each control of type TextBox.
    ' Parameters  :
    '    txtIn (in) = a TextBox control which should be "handled" by this class in
    '                  order to gain extended default functionality.
    '
    ' Called by   : The Form_Load event handler of any form that has a control
    '               of type TextBox
    '
    ' Modified    :
    '
    '----------------------------------------------------------------------------
    Const cstrCurrentProc As String = "BindControl"
    On Error GoTo PROC_ERR

    Set myTextBox = txtIn
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
'|                        PRIVATE  Procedures                       |
'/                                                                  \
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/


'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub myTextBox_GotFocus()
    '----------------------------------------------------------------------------
    ' Procedure   :  myTextBox_GotFocus
    '
    ' Comments    : Ensures the entire contents of the textbox is selected
    '               when the control receives focus. This facilitates the
    '               user emptying out or replacing the current contents.
    ' Parameters  :
    '    txtIn (in) = a TextBox control which should be "handled" by this class in
    '                  order to gain extended default functionality.
    '
    ' Called by   : the GotFocus event handler belonging to each TextBox control
    '               (of a form) that is bound to this class
    '
    ' Modified
    ':
    '----------------------------------------------------------------------------
    Const cstrCurrentProc As String = "myTextBox_GotFocus"
    On Error GoTo PROC_ERR

    myTextBox.SelStart = 0
    myTextBox.SelLength = Len(myTextBox.Text)
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
Private Sub myTextBox_KeyDown(ByRef pintKeyCode As Integer, ByRef pintShift As Integer)
    '----------------------------------------------------------------------------
    ' Procedure   :  myTextBox_KeyDown
    '
    ' Comments    : If the user presses Enter while in the TextBox, it treats it
    '               as a Tab press and thus moves focus to the next control
    '
    ' Parameters  : N/A
    '
    ' Called by   : the KeyDown event handler belonging to each TextBox control
    '               (of a form) that is bound to this class
    '
    ' Modified
    ':
    '----------------------------------------------------------------------------
    Const cstrCurrentProc As String = "myTextBox_KeyDown"
    On Error GoTo PROC_ERR

    If pintKeyCode = vbKeyReturn Then
        pintKeyCode = 0
        SendKeys vbTab
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
Private Sub myTextBox_KeyPress(ByRef pintKeyAscii As Integer)
    '----------------------------------------------------------------------------
    ' Procedure   :  myTextBox_KeyPress
    '
    ' Comments    : Accepts or rejects typed-in characters depending on the
    '               ValidKeys property setting.
    '
    ' Parameters  : N/A
    '
    ' Called by   : the KeyPress event handler belonging to each TextBox control
    '               (of a form) that is bound to this class
    '
    ' Modified
    ':
    '----------------------------------------------------------------------------
    Const cstrCurrentProc As String = "myTextBox_KeyPress"
    On Error GoTo PROC_ERR

    ' If the ValidKeys property hasn't been set by the form, just skip this event.
    If m_strValidKeys = vbNullString Then
        GoTo PROC_EXIT
    End If
    
    ' Test to see if the key that was pressed is allowed. Also, force the
    ' "back" key to be valid.
    If InStr(m_strValidKeys & vbBack, Chr(pintKeyAscii)) = 0 Then
        ' No, the key wasn't valid, so reject it
        pintKeyAscii = 0
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
Private Sub myTextBox_LostFocus()
    '----------------------------------------------------------------------------
    ' Procedure   :  myTextBox_LostFocus
    '
    ' Comments    : Displays the control's contents per the Format property
    '
    '               NOTE: Dependening on the m_strFormat value, repeatedly hitting
    '                     this event for a non-numeric field could corrupt the contents
    '                     of the TextBox this is operating against. For this app, the
    '                     Format is generally going to be very simple for non-numeric
    '                     fields, e.g., empty, or "any character" or "make it uppercase",
    '                     which should have no side effects if repeated LostFocus events
    '                     are triggered. If the m_strFormat said, for instance,
    '                     to show a literal character and do something else, then *that*
    '                     would have a side effect!
    '
    ' Parameters  : N/A
    '
    ' Called by   : the LostFocus event handler belonging to each TextBox control
    '               (of a form) that is bound to this class
    '
    ' Modified
    ':
    '----------------------------------------------------------------------------
    Const cstrCurrentProc As String = "myTextBox_LostFocus"
    On Error GoTo PROC_ERR

    If m_bShouldForceToUppercase Then
        myTextBox.Text = UCase$(myTextBox.Text)
    End If
    
    If Len(m_strFormat) Then
        If m_bIsNumeric Then
            ' Can't do a "Format$" or "Val" on a numeric field or you lose the value eventually !!!!
            ' If the numeric field is empty, can't do CDec( ).
            ' Have to use CDec() not CDbl() since our amount fields are HUGE; otherwise we lose
            ' decimal positions when applying the Format.
            If Len(myTextBox.Text) > 0 Then
                myTextBox.Text = VBA.Format(CDec(myTextBox.Text), m_strFormat)
            Else
                myTextBox.Text = VBA.Format$(myTextBox.Text, m_strFormat)
            End If
        Else
            myTextBox.Text = VBA.Format$(myTextBox.Text, m_strFormat)
        End If
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
