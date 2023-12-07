VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "chrgHourglass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "This class is used to ensure any change to the Screen.Pointer is restored when the procedure that set it goes out of scope."
'******************************************************************************
' Module     : chrgHourglass
' Description: This class implements an alternative way to show an hourglass
' Procedures :
'              Property Get Value() - public
'              Property Let Value() - public
'              Class_Terminate()
' Source     : Total Visual SourceBook 2000
' Modified   :
' 03/03/02 BAW (Phase2A) Added support for new global error handler'
' -------------------------------------------------------------------------------
Option Explicit
Option Compare Binary
Private Const mcstrName As String = "chrgHourglass."

'------------------------------------------
'            MEMBER VARIABLES
'------------------------------------------
Private m_bValue As Boolean 'local copy


'------------------------------------------
'           PROPERTY GET / LET
'------------------------------------------
Public Property Get value() As Boolean
    ' Comments  : Returns True if the cursor is shown as
    '             an hourglass; False otherwise
    ' Parameters: None
    ' Modified  :
    ' Source    : Total Visual SourceBook 2000
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "Property Get Value"

    value = m_bValue
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



Public Property Let value(ByVal bValue As Boolean)
    ' Comments  : Changes the cursor to/from an hourglass
    ' Parameters: bValue (in) - True to turn the cursor into
    '                  an hourglass; False to set it to the
    '                  default
    ' Modified  :
    ' Source    : Total Visual SourceBook 2000
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "Property Let Value"

    m_bValue = bValue
    If m_bValue Then
        Screen.MousePointer = vbHourglass
    Else
        Screen.MousePointer = vbDefault
    End If
PROC_EXIT:
    ' BAW 03/25/2002 - We might get here IN THE COURSE OF reporting an error from
    '                  an event handler, e.g., as a long-winded process has ended
    '                  due to an error. So, ignore all errors so we don't
    '                  get into a sort of loop trying to propagate an error back
    '                  to the Event Handler that called us. This is why
    '                  On Error Resume Next is used instead of On Error GoTo 0.
    On Error Resume Next
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



' ********************************************
'  Initialize and Terminate
' ********************************************
Private Sub Class_Terminate()
    ' Comments  : Reset the cursor to the default
    ' Parameters: None
    ' Modified  :
    ' Source    : Total Visual SourceBook 2000
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "Class_Terminate"

    Screen.MousePointer = vbDefault
PROC_EXIT:
    ' BAW 03/25/2002 - We might get here IN THE COURSE OF reporting an error from
    '                  an event handler, e.g., as a long-winded process has ended
    '                  due to an error. So, ignore all errors so we don't
    '                  get into a sort of loop trying to propagate an error back
    '                  to the Event Handler that called us. This is why
    '                  On Error Resume Next is used instead of On Error GoTo 0.
    On Error Resume Next
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
