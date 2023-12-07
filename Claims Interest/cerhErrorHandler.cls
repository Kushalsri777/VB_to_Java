VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cerhErrorHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "This class serves as an application-wide error handler."
' =-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-=
' Class       : cerhErrorHandler
' Description : Used to support a centralized transaction-oriented error handling
'               scheme across the application
'
'               Note that this class does NOT CONTAIN ERROR HANDLERS, since it
'               is itself an error handler. If any errors are encountered,
'               then VB itself will report them...as fatal errors that cause the
'               app to terminate.
'
' Procedures  :
'   Private     Class_Initialize()
'   Public      Property Get ErrNum() As Long
'   Public      Property Let ScreenName(ByVal strValue As String)
'   Public      AddSubstitution(ByVal strValue As String)
'   Public      Clear()
'   Public      PropagateError()
'   Public      ReportFatalError(ByVal strScreenName As String) As Integer
'   Public      ReportNonFatal(ByVal lngErrNumIn As Long, ByVal strErrContextIn As String, _
                    ParamArray varParmsIn() As Variant) As Integer
'   Public      SaveAppSpecificErr(ByVal lngErrNumIn As Long, ByVal strErrContextIn As _
'                   String, ParamArray varParmsIn() As Variant)
'   Public      SaveErrObjectData(ByVal strAppContext As String) As Boolean
'   Private     fnGetSubstitutionsLBound() As Long
'   Private     fnGetResString(ByVal intID As Integer) As String
'   Private     fnGetSubstitutionsUBound() As Long
'
' Modified:
'
'   Version Date     Who   What
'   ------- -------- ---   -------------------------------------------------------------------
'   4.0     03/20/02 BAW   (Phase2A) Commented out procedures that aren't used and possibly
'                          could be deleted.
'   3.0     03/11/02 BAW   (Phase2A) Made SaveAppSpecificError( ) raise the error that was just
'                          recorded so that it could be handled by the local error handler
'                          immediately (for non-error situations) or propogated (for errors).
'                          Also added TranslatedErrNum as a public property.
'   2.0     03/07/02 BAW   (Phase2A) Updated comments at top of module. Added conditional
'                          compilation code (#If DEBUG_ERH) with which to debug error
'                          handling code.
' =-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-=
Option Explicit
Option Compare Binary

' Local variables to hold Public Property values
Private m_lngErrNum           As Long
Private m_strErrDesc          As String
Private m_strErrContext       As String
Private m_strScreenName       As String
Private m_lngTranslatedErrNum As Long

' Local variables that are *NOT* Public Properties
Private m_strSubstitutions()  As String

' Following constants are used with procedures accessing the m_strSubstitutions array
Private Const mclngSubscriptOutOfRange As Long = 9
Private Const mclngArrayNotInitialized As Long = -1

' The following constants are used to initialize the public properties
' of this object, or to determine whether the properties are at their
' initialized values.
Private Const mcstrErrDescDefault As String = vbNullString
Private Const mcstrErrContextDefault As String = vbNullString

' The following defines how substitution placeholders are identified in message text
' defined in the CLAIM.RES resource file
Private Const mcstrSubstitutionDelimiter As String = "@@"




'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'\                                                                  /
'|         CLASS_INITIALIZE / CLASS_TERMINATE   Procedures         |
'/                                                                  \
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub Class_Initialize()
    ' Set initial values to defaults which may be overridden with property settings
    Clear
End Sub


Private Sub Class_Terminate()
    ' Free up resources allocated in this class

    End
    'Erase m_strSubstitutions
End Sub




'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'\                                                                  /
'|                PROPERTY GET/LET    Procedures                    |
'/                                                                  \
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Property Get ErrNum() As Long
    ' Returns the current value of the ErrNum property

    ErrNum = m_lngErrNum
End Property


Public Property Let ScreenName(ByVal strValue As String)
    ' Sets the ScreenName property to the value specified by strValue

    m_strScreenName = strValue
End Property







'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'\                                                                  /
'|                        PUBLIC  Procedures                        |
'/                                                                  \
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub AddSubstitution(ByVal strValue As String)
    ' Comments:   Resizes the current array and then adds the entry
    ' Parameters: strValue   (in)   Entry to add to the array
    ' Returns:    N/A
    ' Called by : Any application procedure that needs to generate
    '             a message that utilizes substitution values in
    '             its text, e.g., @@1, @@2, etc.
    Dim lngLastEntry As Long

    Debug.Assert strValue <> vbNullString   ' Validate the input

    lngLastEntry = fnGetSubstitutionsUBound()
    If lngLastEntry = mclngArrayNotInitialized Then
        lngLastEntry = 0
    Else
        lngLastEntry = lngLastEntry + 1
    End If

    ReDim Preserve m_strSubstitutions(lngLastEntry)
    m_strSubstitutions(lngLastEntry) = strValue
End Sub



'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub Clear()
    ' Comments:  Resets all properties to their intialalized values
    '
    ' Parameters: None
    ' Returns:    N/A
    ' Called by : ReportFatalError, Class_Initialize

    m_lngErrNum = gclngERR_NUM_DEFAULT      ' Fatal Error Code, meant to indicate it hasn't been set
    m_strErrDesc = mcstrErrDescDefault
    m_strErrContext = mcstrErrContextDefault
    Erase m_strSubstitutions
End Sub


'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub PropagateError(ByVal strErrContextIn As String)
    ' Comments  : This public subroutine is used when an error must be raised
    '             from a standard module, class method, or non-event handler
    '             back up to the event handler that initiated it. It assumes
    '             the error has already been saved to this class' member variables
    '             and thus raises the error based on those member variables and
    '             NOT the VBA.Err object.
    '
    ' Parameters: N/A
    ' Returns   : N/A
    ' Modified  :
    ' --------------------------------------------------

    #If DEBUG_ERH Then
        Debug.Print "PropagateError called by " & strErrContextIn & _
                    ". Error#=" & CStr(m_lngErrNum) & "(or " & CStr(m_lngErrNum - vbObjectError) & _
                    ") Desc=" & m_strErrDesc & " Context=" & strErrContextIn
    #End If

    Err.Raise m_lngErrNum, m_strErrContext, m_strErrDesc
End Sub



'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Function ReportFatalError(ByVal strScreenName As String) As Integer
    ' Comments  : This public subroutine should be used INSTEAD OF the VBA.MsgBox
    '             function to report *errors*, whether they be PROCESS fatal
    '             (i.e. the event terminates, but not the app) or APP fatal (the
    '             app terminates). All of the information it needs to report the error
    '             is based on public properties that should
    '             be set prior to calling this procedure.
    '
    '             It is assumed that ErrNum reflect this scheme in the CLAIM.RES resource file:
    '                   ID values in the range of:
    '                      4000-4999 = non-fatal error messages
    '                      9000-9999 = fatal error messages
    '
    '             Anything outside this range (such as VB and ADO errors) will be
    '             treated as a fatal error
    '
    '             Note that ErrNums in the following ranges should be reported
    '             via the ReportNotFatal( ) method of this class:
    '                      1000-1999 = informational messages
    '                      2000-2999 = warning messages
    '                      3000-3999 = alert messages
    '
    ' Parameters: strScreenName (in) - the name of the screen (as it should appear in
    '                      the MsgBox's title
    '
    ' Returns   : Integer indicating the button the user clicked in the MsgBox,
    '             e.g., vbOK, vbYes, vbNo, etc.
    ' Modified  :
    ' --------------------------------------------------
    Const cstrCurrentProc As String = "ReportFatalError"
    Dim intI                        As Integer
    Dim lngButtons                  As Long
    Dim lngTranslatedErrNum         As Long
    Dim strErrDesc                  As String
    Dim strMsgText                  As String
    Const cstrFatalErrorPrefix      As String = "An error has occurred from which the application cannot recover. " & _
                                                "The application will now be terminated." & vbCrLf & vbCrLf

    On Error Resume Next

    #If DEBUG_ERH Then
        Debug.Print "ReportFatalError called by " & strScreenName & _
                    ". Error#=" & CStr(m_lngErrNum) & "(or " & CStr(m_lngErrNum - vbObjectError) & _
                    ") Desc=" & m_strErrDesc & " Context=" & m_strErrContext
    #End If

    ' Make sure Cursor reverts back to normal, in case it was left in an hourglass
    Screen.MousePointer = vbDefault

    ' Remove the vbObjectError value, so we're left with an app-specific error code, as is
    ' used in the Resource File. If this lowers it too much, then revert back to the
    ' specified Error Code (probably is an ADO or VB error).
    lngTranslatedErrNum = m_lngErrNum - vbObjectError
    If lngTranslatedErrNum < gcRES_LOWEST_APP_ERROR Or lngTranslatedErrNum > gcRES_HIGHEST_APP_ERROR Then
        lngTranslatedErrNum = m_lngErrNum
        strMsgText = m_strErrDesc
    Else
        strMsgText = fnGetResString(lngTranslatedErrNum)

        ' Replace Carriage Return / Line Feed tokens with VB-equivalent
        '        @@CRLF ==> vbCrLf
        '        @@CR   ==> vbCr
        '        @@LF   ==> vbLf
        strMsgText = Replace(strMsgText, mcstrSubstitutionDelimiter & "CRLF", vbCrLf)
        strMsgText = Replace(strMsgText, mcstrSubstitutionDelimiter & "CR", vbCr)
        strMsgText = Replace(strMsgText, mcstrSubstitutionDelimiter & "LF", vbLf)

        ' Replace parameters in message text with substituted values, e.g.,
        '      @@1  ... for the value of varParms(0)
        '      @@2  ... for the value of varParms(1), etc.
        If fnGetSubstitutionsUBound() <> mclngArrayNotInitialized Then
            For intI = LBound(m_strSubstitutions) To UBound(m_strSubstitutions)
                strMsgText = Replace(strMsgText, mcstrSubstitutionDelimiter & CStr(intI + 1), m_strSubstitutions(intI))
            Next
            ' Do assertion if any "@@" still remain after doing the substitution. It
            ' indicates the call to this function and Error text in the .RES file are out of synch.
            Debug.Assert (InStr(1, strMsgText, mcstrSubstitutionDelimiter) = 0)
        End If
    End If

    Select Case lngTranslatedErrNum
        Case gcRES_NERR_START To gcRES_NERR_END
            lngButtons = vbOKOnly & vbExclamation
        Case Else
            ' Intended to be 9000 to 9999 ... or a VB or ADO error code
            lngButtons = vbOKOnly & vbExclamation
            gbAmProcessingAnAppFatalError = True
    End Select

    If gbAmProcessingAnAppFatalError Then
        fnLogWrite "Application Fatal Error in " & m_strErrContext & ": " & strMsgText, cstrCurrentProc
        ' Display the errror via a modal frmMsgBox window
        With frmMsgBox
            .ScreenName = strScreenName
            .ErrorCode = lngTranslatedErrNum
            .MsgText = cstrFatalErrorPrefix & strMsgText
            .ErrorContext = m_strErrContext
            .Show vbModal
            ReportFatalError = .ButtonClicked
        End With
        Unload frmMsgBox

        ' Initialize this class' member variables to acknowledge the error has been reported
        ' and thus no longer needs to be Propagated or reported again, e.g., in the
        ' Unload events triggered by the following FOR loop
        Clear

        Debug.Assert False      ' If in IDE, force the debugger to stop here
        For intI = Forms.Count - 1 To 0 Step -1
            Unload Forms(intI)
        Next intI

        fnDeallocateGlobalObjects
        
        gbAmProcessingAnAppFatalError = False

        ' !!!!!!!  TERMINATE THE APP !!!!!!!
        End
    Else
        fnLogWrite "Error in " & m_strErrContext & ": " & strMsgText, cstrCurrentProc
        With frmMsgBox
            .ScreenName = strScreenName
            .ErrorCode = lngTranslatedErrNum
            .MsgText = strMsgText
            .ErrorContext = m_strErrContext
            
            ' Initialize this class' member variables to acknowledge the error has been reported
            ' and thus no longer needs to be Propagated or reported again. Do it now,
            ' before showing frmMsgBox, because if we somehow get into a procedure's PROC_EXIT
            ' with this cerhErrorHandler object still showing the remains of an error, there
            ' is nothing to propagate the error back up to and the app will die with an
            ' unhandled error. (Example: cetbExtendedTextbox.Lost_Focus will be called if an
            ' error occurs on a maintenance screen that has TextBoxes tied to that extended
            ' textbox class and that LostFocus event handler will treat the as-yet-unreported-error
            ' (if we didn't call .Clear first) as an unhandled error.
            Clear
            
            .Show vbModal
            ReportFatalError = .ButtonClicked
        End With
        Unload frmMsgBox

        ' Unload the splash screen, if it is still loaded
        fnUnloadSplash
    End If
End Function



'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Function ReportNonFatal(ByVal lngErrNumIn As Long, ByVal strErrContextIn As String, _
                    ParamArray varParmsIn() As Variant) As Integer
    ' Comments  : This public subroutine should be used whenever a Warning,
    '             Informational or Alert type of message should be displayed.
    '             It can be called anywhere in the system, e.g, form or std
    '             module, or event handler or non-event handler.
    '
    '             Use this function instead of VBA.MsgBox( ).
    '
    '             The messages displayed by this function are *not* recorded
    '             in the class' data.
    '
    '             It is assumed that ErrNum reflects this scheme in the CLAIM.RES resource file:
    '                   ID values in the range of:
    '                      1000-1999 = informational messages
    '                      2000-2999 = warning messages
    '                      3000-3999 = alert messages
    '             Anything outside this range will raise an assert while in
    '             Debug mode.
    '
    ' Parameters:
    '             lngErrNumIn (in)     - the error code (in the range of gcRES_LOWEST_APP_ERROR
    '                                    to gcRES_HIGHEST_APP_ERROR)
    '             strErrContextIn (in) - the module.procname in which the error occurred
    '             varParmsIn (in)      - (optional) a variable number of parameters which
    '                                    will eventually be substituted for  @@1, @@2, etc.
    '                                    placeholders in the message text.
    '
    ' Returns   : Integer indicating the button the user clicked in the MsgBox,
    '             e.g., vbOK, vbYes, vbNo, etc.
    ' Modified  :
    '    04/23/02 BAW  Added logic to clear the Err object.
    ' --------------------------------------------------
    Const cstrCurrentProc As String = "ReportNonFatal"
    Dim intI                 As Integer
    Dim lngTranslatedErrNum  As Long
    Dim strMsgText           As String
    On Error Resume Next

    #If DEBUG_ERH Then
        Debug.Print "ReportNonFatal called by " & strErrContextIn & _
                    ". Error#=" & CStr(lngErrNumIn) & "(or " & CStr(m_lngErrNum - vbObjectError) & ")"
    #End If

    ' Clear Err object in case an error is left in it. This could occur in the case of a stored procedure
    ' returning a 4029 (Dependent Records exist) error that the form, at times, will turn into an "info" message.
    Err.Clear

    ' Remove the vbObjectError value, so we're left with an app-specific error code, as is
    ' used in the Resource File.
    lngTranslatedErrNum = lngErrNumIn - vbObjectError

    ' Verify we've got an lngErrNumIn in the appropriate range
    Debug.Assert (lngTranslatedErrNum >= gcRES_INFO_START)
    Debug.Assert (lngTranslatedErrNum <= gcRES_ALRT_END)

    ' If the translated error number is too low (which shouldn't happen in the
    ' production environment due to the previous Assert bringing it to the
    ' developer's attention during the development process), then revert back
    ' to the specified Error Code.
    If lngTranslatedErrNum < gcRES_LOWEST_APP_ERROR Or lngTranslatedErrNum > gcRES_HIGHEST_APP_ERROR Then
        lngTranslatedErrNum = lngErrNumIn
        strMsgText = "The following error was encountered but its description " & _
                     "could not be located: " & vbCr & m_strErrDesc
    Else
        strMsgText = fnGetResString(lngTranslatedErrNum)

        ' Replace Carriage Return / Line Feed tokens with VB-equivalent
        '        @@CRLF ==> vbCrLf
        '        @@CR   ==> vbCr
        '        @@LF   ==> vbLf
        strMsgText = Replace(strMsgText, mcstrSubstitutionDelimiter & "CRLF", vbCrLf)
        strMsgText = Replace(strMsgText, mcstrSubstitutionDelimiter & "CR", vbCr)
        strMsgText = Replace(strMsgText, mcstrSubstitutionDelimiter & "LF", vbLf)

        ' Replace parameters in message text with substituted values, e.g.,
        '      @@1  ... for the value of varParms(0)
        '      @@2  ... for the value of varParms(1), etc.

        For intI = LBound(varParmsIn) To UBound(varParmsIn)
            strMsgText = Replace(strMsgText, mcstrSubstitutionDelimiter & CStr(intI + 1), varParmsIn(intI))
        Next
        ' Do assertion if any "@@" still remain after doing the substitution. It
        ' indicates the call to this function and Error text in the .RES file are out of synch.
        Debug.Assert (InStr(1, strMsgText, mcstrSubstitutionDelimiter) = 0)
    End If

    ' May need to make this log file write conditional, or not do it at all,
    ' if the log file starts filling up too quickly...
    fnLogWrite "Non-Fatal Msg/Error in " & strErrContextIn & ": " & strMsgText, cstrCurrentProc
    
    ' Display the errror via a modal frmMsgBox window
    With frmMsgBox
        .ScreenName = m_strScreenName
        .ErrorCode = lngTranslatedErrNum
        .MsgText = strMsgText
        .ErrorContext = strErrContextIn
        .Show vbModal
        ReportNonFatal = .ButtonClicked
    End With
    Unload frmMsgBox
End Function



Public Sub SaveAppSpecificErr(ByVal lngErrNumIn As Long, ByVal strErrContextIn As String, _
                              ParamArray varParmsIn() As Variant)
    ' Comments  : This public subroutine loads this object's public properties
    '             with the specified values. The ErrDesc property is set when
    '             the error is actually reported via ReportFatalError()
    ' Parameters: lngErrNumIn (in)     - the error code in non-fatal or fatal error range
    '             strErrContextIn (in) - the module.procname in which the error occurred
    '             varParmsIn (in)      - (optional) a variable number of parameters which
    '                                    will eventually be substituted for  @@1, @@2, etc.
    '                                    placeholders in the message text.
    '
    ' Returns   : N/A
    '
    ' Modified  :
    ' --------------------------------------------------
    Dim intI As Integer
    Dim lngTranslatedErrNum As Long

    #If DEBUG_ERH Then
        Debug.Print "SaveAppSpecificError called by " & strErrContextIn & _
                    ". Error#=" & CStr(lngErrNumIn) & "(or " & CStr(m_lngErrNum - vbObjectError) & ")"
    #End If

    ' Inform developer if error is being recorded prior to reporting a previously
    ' recorded error
    Debug.Assert (m_lngErrNum = gclngERR_NUM_DEFAULT)

    ' Verify we've got an lngErrNumIn in the appropriate range
    lngTranslatedErrNum = lngErrNumIn - vbObjectError
    Debug.Assert (lngTranslatedErrNum >= gcRES_NERR_START)
    Debug.Assert (lngTranslatedErrNum <= gcRES_FERR_END)

    m_lngErrNum = lngErrNumIn
    m_strErrDesc = "App-specific error " & CStr(lngErrNumIn - vbObjectError)
    m_strErrContext = strErrContextIn
    m_lngTranslatedErrNum = lngTranslatedErrNum

    Debug.Print "Recording this error: Error#=" & CStr(m_lngErrNum - vbObjectError) & " Context=" & strErrContextIn
    For intI = 0 To UBound(varParmsIn)
        AddSubstitution varParmsIn(intI)
    Next intI

    'Err.Raise m_lngErrNum, m_strErrContext, m_strErrDesc  ' NEW 03/11/2002
End Sub



'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Function SaveErrObjectData(ByVal strAppContextIn As String) As Boolean
    ' Comments  : This public subroutine loads this object's public properties
    '             with the contents of the Err object *unless* this object's
    '             properties have already been set (and not cleared/reported)
    '             after a previous Err.Raise. This conditional setting of the
    '             properties guards against corruption of the Err object when
    '             errors are passed back up the call chain.
    '
    '             This should be called at the beginning of an error handler
    '             that will be raising an error back up the call chain.
    '
    '             WARNING: Do **NOT** include an On Error or Resume statement
    '                      in this procedure as it will clear the Err object.
    '
    ' Parameters: strAppContextIn (in) - indicates the application context in
    '                                  which the error occured (to augment
    '                                  what the Err object automatically has)
    '
    ' Returns   : True if the properties were saved; False otherwise
    '
    ' Modified  :
    ' --------------------------------------------------
    Dim lngTempErrNum As Long
    Dim strTempErrDesc As String
    Dim strTempErrContext As String

    ' The following statement calls a function (fnGetSubstitutions)
    ' which must use an error handler. However, its use of an error
    ' handler resets the Err object, hence we must save the Err object's
    ' contents to local variables before calling that function.
    lngTempErrNum = Err.Number
    strTempErrDesc = Err.Description
    strTempErrContext = Err.Source

    On Error GoTo 0

    #If DEBUG_ERH Then
        Debug.Print "SaveErrObjectData called by " & strAppContextIn & _
                    ". Error#=" & CStr(lngTempErrNum) & "(or " & CStr(lngTempErrNum - vbObjectError) & _
                    ") Desc=" & strTempErrDesc & " Context=" & strTempErrContext
    #End If

    If m_lngErrNum = gclngERR_NUM_DEFAULT And _
       m_strErrDesc = mcstrErrDescDefault And _
       m_strErrContext = mcstrErrContextDefault And _
       fnGetSubstitutionsUBound = mclngArrayNotInitialized Then
            m_lngErrNum = lngTempErrNum
            m_strErrDesc = strTempErrDesc
            m_strErrContext = strAppContextIn & "/" & strTempErrContext
            m_lngTranslatedErrNum = lngTempErrNum
            SaveErrObjectData = True
    End If
End Function



'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'\                                                                  /
'|                       PRIVATE  Procedures                        |
'/                                                                  \
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/


'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Function fnGetResString(ByVal intID As Integer) As String
    ' Comments  : This public subroutine is called when
    '             displaying a message box. It retrieves the text for
    '             the specified error code (intID) from the string table
    '             in the CLAIM.RES resource file.
    ' Parameters: intID    = (input) the error code whose description to retrieve
    ' Returns   : A string containing the description of the error.
    ' Modified  :
    ' --------------------------------------------------
    Dim intDefaultID As Integer

    On Error Resume Next
    fnGetResString = LoadResString(intID)
    If Err Then
        ' If the named ID cannot be found, strip off the final digits and append 000
        ' (i.e. 1699 becomes 1000) so the default text for that message range can
        ' be shown.
        intDefaultID = CInt(Left$(CStr(intID), 1) & "000")
        fnGetResString = LoadResString(intDefaultID)
        If Err Then
            fnGetResString = "Error " & CStr(intID) & " not found in resource file."     ' catch-all
        End If
    End If

    On Error GoTo 0
End Function



'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Function fnGetSubstitutionsUBound() As Long
    ' Comments:   Returns the highest possible index to the
    '             m_Substitutions array.
    '
    ' Parameters: None
    ' Returns:    The UBound() index, or -1 if the array has not been initialized
    ' Called by : GetSubstitution(), AddSubstitution(), ReportFatalError()
    Const cstrCurrentProc As String = "fnGetSubstitutionsUBound"

    On Error Resume Next
    fnGetSubstitutionsUBound = UBound(m_strSubstitutions)
    If Err.Number = mclngSubscriptOutOfRange Then
        Err.Clear   ' Lose the error 9 so the error will be ignored
        fnGetSubstitutionsUBound = mclngArrayNotInitialized
    End If
End Function
