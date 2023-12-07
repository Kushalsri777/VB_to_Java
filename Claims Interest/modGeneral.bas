Attribute VB_Name = "modGeneral"
'******************************************************************************
' Module     : modGeneral
' Description:
' Procedures :
'              fnAddBackslash(ByVal strPathIn As String) As String
'              fnAddColumnToGrid(ByRef vfgIn As VSFlexGrid, ByVal strColumnName As String, Optional ByVal bHidden As Boolean = False)
'              fnAreChildFormsOpen() As Boolean
'              fnBoolToYN(ByVal bIn As Boolean) As String
'              fnBuildQualifiedFileName(ByVal strDir As String, strFileName As String) As String
'              fnCenterFormOnMDI(ByVal frmMDIParent As Form, ByRef frmMDIChild As Form)
'              fnCenterFormOnScreen(ByRef frmIn As Form)
'              fnConnectToArchiveDB(ByRef conIn As cconConnection)
'              fnCopyFieldToRST(ByVal strColNm As String, ByRef rstSource As ADODB.Recordset, _
'              fnCopyRSTAsUpdateable(ByVal rstSource As ADODB.Recordset) As ADODB.Recordset
'              fnCopyRSTAsUpdateable2(ByVal rstSource As ADODB.Recordset) As ADODB.Recordset
'              fnCstrDate(ByRef strDateIn As String, Optional ByRef strFormatIn As String = "MM/DD/YYYY") As Date
'              fnDeallocateGlobalObjects()
'              fnEnableDisableControl(ByVal ctlIn As Control, Optional ByVal bEnable As Boolean = True)
'              fnFirstDayOfMonth(ByVal dteIn As Date) As Date
'              fnFixDecimal(ByVal dblAmount As Double, ByVal intPosition As Integer, _
'              fnFormatMMDDYYYYDate(ByVal strDateIn As String) As String
'              fnFormatYYYYMMDDDate(ByVal strDateIn As String) As String
'              fnFreeObject(ByRef pObj As Object)
'              fnFreeRecordset(ByRef pRST As ADODB.Recordset)
'              fnGetExtPart(pstrIn As String) As String
'              fnGetStateInfo(ByVal strStateIn As String, ByVal strLOBIn As String, _
'              fnGetStateRule(ByVal strStateIn As String, ByVal strLOBIn As String, _
' MME START WRUS 4999
'              fnGetStateTierInfo(ByVal strStateIn As String, ByVal strLOBIn As String, _
' MME END WRUS 4999
'              fnHighlightText(ctlIn As Control)
'              fnIfNull(varValueIn As Variant, Optional varNullValueIn As Variant = "") As Variant
'              fnInitializeAppConnectionObject()
'              fnInitializeMenuItems()
'              fnInitializeStateInfo(ByRef siInOut As StateInfo)
'              fnIsFormLoaded(ByVal strFormName As String, Optional ByRef frmFound As Form) As Boolean
'              fnLastDayOfMonth(ByVal dteIn As Date) As Date
'              fnLimitChange(ByRef pctlIn As Control, ByRef pintMaxLen As Integer)
'              fnLimitKeyPress(ByRef pctlIn As Control, _
'              fnLongStateToShortState(ByVal strStateIn As String) As String
'              fnMakeWeekday(ByVal dteIn As Date, ByVal intDirection As EnumPrevNext) As Date
'              fnMaxDate(ByVal dte1 As Date, ByVal dte2 As Date) As Date
'              fnMaxDouble(ByVal dblOne As Double, ByVal dblTwo As Double) As Double
'              fnMinDate(ByVal dte1 As Date, ByVal dte2 As Date) As Date
'              fnMinDouble(ByVal dblOne As Double, ByVal dblTwo As Double) As Double
'              fnPadRightString(ByVal strIn As String, ByVal lngStrLen As Long, _
'              fnPersistRecordsetToCSV(ByRef rstIn As ADODB.Recordset, ByVal strFileNm As String)
'              fnPersistRecordsetToXML(ByRef rstIn As ADODB.Recordset, ByVal strFileNm As String)
'              fnPhoneNumber_AddDash(ByVal strIn As String) As String
'              fnQuoted(ByVal strIn As String) As String
'              fnQuotedOrNull(ByVal strIn As String, Optional ByVal bRTrim As Boolean = False, _
'              fnRemoveCloseButton(ByVal frmIn As Form)
'              fnRemoveUnderScoresFromFieldName(ByVal strIn As String) As String
'              fnRound(ByVal dblAmountIn As Double, ByVal intSignIn As Integer) As Double
'              fnRoundToNextWholeDollar(ByVal dblNumber As Double) As Double
' MME START WRUS 4999
'              fnSelectRecord(ByVal lngKey1 As Long) As ADODB.Recordset
' MME END WRUS 4999
'              fnSetTopmostWindow(ByVal frm As Form, Optional ByVal bTopmost As Boolean = True)
'              fnShortStateToLongState(ByVal strStateIn As String) As String
'              fnShowFormsCollection()
'              fnShowRecordPosition(rstIn As ADODB.Recordset) As String
'              fnSSNTIN_AddDash(ByVal strIn As String, Optional bIsTin As Boolean = False) As String
'              fnTerminateTheApp()
'              fnTranslateToMaxValue(ByVal intDollarPositions As Integer, ByVal intDecimalPositions As Integer) As Double
'              fnUnloadSplash()
'              fnYNToBool(ByVal strIn As String) As Boolean
'              Sub fnWindowLock(ByVal hWnd As Long)
'              Sub fnWindowUnlock()
'              TestStub_fnGetStateInfo()
'
' Modified   :
' 03/26/01 BAW Cleaned with Total Visual CodeTools 2000
' 01/2002  BAW Removed Scope-related fields & logic; added fnLongStateToShortState( )
'              and fnShortStateToLongState( ). Also, optimized per Project Analyzer,
'              removing dead code, adding "$" to Mid/Space, etc. Also, added the
'              fnBuildQualifiedFileName( ) and fnPadRightString( ) procs. Also
'              corrected a latent bug in fnIsFormLoaded( ).
' --------------------------------------------------
Option Explicit
Option Compare Binary
Private Const mcstrName As String = "modGeneral."

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'                                        CONDITIONAL COMPILER CONSTANTS
'                                      Set to 1 to enable or 0 to disable.
'
' DEBUG_ERH - Shows when and by whom errors are recorded, propagated and reported.
' DEBUG_RST - Shows how many records are in each recordset created (to determine if additional
'             tuning is warranted).
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
#Const DEBUG_ERH = 0
#Const DEBUG_RST = 0

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'                                        DEBUGGING CONSTANTS
'                                      Set to True to enable or False to disable.
'
' bDebugAppTermination - Shows information about how forms are getting unloaded and global objects
'                        are getting deallocated  (Doesn't work right if this is defined as a
'                        conditional compiler constant!)
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Const bDebugAppTermination As Boolean = False


'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'                                        GLOBAL VARIABLES
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Each of these must be set to Nothing when app ends (due to fatal error or the user's choice) !
Public gapsApp                As capsAppSettings        ' Accesses app settings stored in the registry
Public gerhApp                As cerhErrorHandler       ' Global error handler
Public gadwApp                As cadwADOWrapper         ' Use ADO Wrapper to replace DataService class (Can delete in 2C)
Public gconAppActive          As cconConnection         ' Handles ADO connection to the active app database

' Indicates that an application-fatal error is being processed and thus the app will soon be forcibly shut down.
' This is set by the ReportFatalError method of the cerhErrorhanlder class, but is queried by each form's
' Form_QueryUnload to ensure ALL requests to unload will be honored without user prompts if this switch is set to True.
Public gbAmProcessingAnAppFatalError As Boolean

' Indicates that the app is trying to be shut down. This could occur if an application-fatal error is being
' processed, but also if the user chose File | Exit from the MDI screen or pressed the equivalent keystroke
' (Alt-F4).  When this indicator is true, a form's QueryUnload event should not set pintCancel to True if
' the user has opted to discard their pending changes.
Public gbAmTryingToTerminateTheApp As Boolean

Public Const gcstrDoubleQuote As String = """"
Public Const gcstrSingleQuote As String = "'"

' gclngNoSelection is used to indicate there is no selected entry in a ComboBox, ListBox
' or fpComboAdo control.
Private Const gclngNoSelection As Long = -1

' gcstrBlankEntry is used in ComboBoxes used for Nullable fields, so
' the user can select and the screen can successfully display Nulls.
Public Const gcstrBlankEntry As String = " "

' gcstrAllEntry used in combo box population
Public Const gcstrAllEntry As String = "--All--"

' gcstrNullEntry used in combo box population
Public Const gcstrNullEntry As String = "<NULL>"

' gcintClickedCloseButton used when gerhApp.ReportNonFatal is called, to handle situations
' where the user clicked the Close ("X") button rather than Yes, No, OK, Cancel, etc. to
' dismiss the screen.
Public Const gcintClickedCloseButton As Integer = 0

Private Const mcintZero As Integer = 0

' The following boolean indicates whether the application log file entries should be verbose (i.e. extra
' loggin) as well as wider (i.e. more text visible). Currently this isn't used. If/when added, add
' code in the startup object to parse the command line to see if the /v switch (verbose mode) was
' specified and set this boolean accordingly. See Spuds/Scuds for an example.
Public gbLogVerbose    As Boolean      ' Indicates whether log file should be terse (default) or verbose


'-----------------------------------------------------------------------
' The following defines selected columns from the State98 table. It is
' used by the fnGetStateInfo( ) function in frmPayee.
'-----------------------------------------------------------------------
Public Type StateInfo
    ' Fields from State98 table
   LobCd As String                  ' not null
   StCd As String                   ' not null
   StrlEffDt As Date                ' not null
   CalcIdtypCd As String            ' not null
   ReqdIdtypCd As String            ' not null
   IruleCd As String                ' not null
   StrlEndDt As Variant             ' nullable
   StrlIntRptgFlrAmt As Currency    ' not null  decimal(11,2)
   StrlIntCalcOfstNum As Integer    ' not null  smallint
   StrlIntReqdOfstNum As Integer    ' not null  smalling
   StrlIntRuleAmt As Variant        ' nullable  decimal(11,5)
   StrlSpclInstrTxt As String       ' nullable
   ' Fields used in doing Calculation
   FiguredFromDate As Date
   PayablePeriodEndDate As Date
   NbrOfDaysToPayInterest As Integer
   InterestRateToUse As Double
   ClaimInterestAmt As Currency
   WithheldAmt As Currency
   TotalForThisPayee As Currency
   CalculationInfo As String
End Type

' The following 3 functions are used by fnWindowLock( ) and fnWindowUnlock( )
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWnd As Long) As Long

' The following 3 functions are used by fnRemoveCloseButton( )
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long

' The following function is used by fnSetTopmostWindow( )
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long


' The following Enum is used by maintenance screens, to aid in referencing
' the navigation buttons
Public Enum EnumNavigationButtons
    navFirst = 0
    navPrev = 1
    navNext = 2
    navLast = 3
End Enum

' The following Enum is used by the fnMakeWeekday( ) method.
Public Enum EnumPrevNext
    epnPrev = 0
    epnNext = 1
End Enum

' The enumPositionDirection enum is used by the GetRelativeRecord( ) method.
Public Enum enumPositionDirection
    epdPreviousRecord = 0
    epdNextRecord = 1
    epdSameRecord = 2
    epdFirstRecord = 3
    epdLastRecord = 4
End Enum

' The enumWhatOperationIsBeingAttempted enum is used by the CheckForAnotherUsersChanges() method.
Public Enum enumWhatOperationIsBeingAttempted
    ewoUpdate = 0
    ewoDelete = 1
End Enum

' The following UDT is used by all table wrapper classes, to define the "standard" propererties and values
' retained for each public property that corresponds to a column in that class' underlying SQL Server table.
Public Type udtColumn
    ColName As String                 ' Corresponds to COLUMN_NAME meta data from Column schema info
    DataType As DataTypeEnum          ' Corresponds to DATA_TYPE meta data from Column schema info
    IsKey As Boolean                  ' Corresponds to XXX from PrimaryKeys schema info
    IsNullable As Boolean             ' Corresponds to IS_NULLABLE  meta data from Column schema info
    HasDefault As Boolean             ' Corresponds to COLUMN_HASDEFAULT meta data from Column schema info
    DefaultValue As Variant           ' Corresponds to COLUMN_DEFAULT meta data from Column schema info
    DollarPositions As Integer        ' Calculated from PRECISION meta data from Column schema info, but which could be overriden
    DecimalPositions As Integer       ' Corresponds to SCALE meta data from Column schema info, but which could be overriden
    Precision As Integer              ' Corresponds to original PRECISION from DBMS. SHOULD NOT be overriden!
    NumericScale As Integer           ' Corresponds to original SCALE from DBMS. SHOULD NOT be overriden!
    MaxCharacters As Integer          ' Correspond to CHARACTER_MAXIMUM_LENGTH meta data from Column schema info
    Format As String                  ' Initially set based on DataType, but form can override
    Mask As String                    ' Initially set based on DataType, but form can override
    AllowableCharacters As String     ' Initially set based on DataType and DecimalPositions, but form can override
    ShouldForceToUppercase As Boolean ' Does *not* correspond to DBMS meta data.
    value As Variant                  ' Initially set based on DefaultValue, if present, but form can override
End Type


Global Const dbChar As Integer = 1
Global Const dbDecimal As Integer = 3
Global Const dbInteger As Integer = 4
Global Const dbDateTime As Integer = 11
Global Const dbVarChar As Integer = 12

'////////////////////////////////////////////////////////////////////////////////////////
Public Function fnAddBackslash(ByVal strPathIn As String) As String
    ' Add a backslash to strPathIn, if needed
    ' Returns a path with a backslash
    On Error GoTo PROC_ERR
    Const cstrCurrentProc   As String = "fnAddBackslash"
    Const cstrBackSlash    As String = "\"
    
    strPathIn = Trim$(strPathIn)
    
    If Right$(strPathIn, 1) <> cstrBackSlash Then
        strPathIn = strPathIn + cstrBackSlash
    End If
    
    fnAddBackslash = strPathIn
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
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
Public Function fnAreChildFormsOpen() As Boolean
    ' Comments  : Determines if any child form is open within this MDI app.
    ' Parameters: N/A
    ' Returns   : True, if one or more child forms are open; False otherwise
    '
    On Error GoTo PROC_ERR

    Const cstrCurrentProc As String = "fnAreChildFormsOpen"
    Dim frm As Form

    fnAreChildFormsOpen = False

    For Each frm In Forms
        If Not frm Is frmMDIMain Then
            fnAreChildFormsOpen = True
            'Debug.Print frm.Name & "is still in Forms collection"
        End If
    Next frm
PROC_EXIT:
    On Error GoTo 0     ' disable error handler

    ' Clean-up statements go here
    fnFreeObject frm
    
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
Public Function fnBoolToTF(ByVal bIn As Boolean) As String
    ' Comments  : Translates True to "T" and False to "F"
    ' Parameters: bIn (in) the boolean to translate
    '
    ' Returns   : "T" or "F"
    '
    ' Modified  : Berry Kropiwka 2019-11-04
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnBoolToTF"

    If bIn Then
        fnBoolToTF = "T"
    Else
        fnBoolToTF = "F"
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
Public Function fnBoolToYN(ByVal bIn As Boolean) As String
    ' Comments  : Translates True to "Y" and False to "N"
    ' Parameters: bIn (in) the boolean to translate
    '
    ' Returns   : "Y" or "N"
    '
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnBoolToYN"

    If bIn Then
        fnBoolToYN = "Y"
    Else
        fnBoolToYN = "N"
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
Public Function fnBuildQualifiedFileName(ByVal strDir As String, ByVal strFileName As String) As String
    ' Comments  : Returns the fully qualified filename, by joining the strDir and
    '             and strFile parameters...with a slash if appropriate
    ' Parameters: strDir - fully qualified folder name
    '             strFileName - file name
    '
    ' Called By : fnLogOpen( ) of modAppLog
    '             fnLogPrune( ) of modAppLog
    '
    ' Modified  :
    '  01/2002 BAW  Copied from SPUDS/SCUDS
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnBuildQualifiedFileName"
    Dim fso      As Scripting.FileSystemObject
    
    Set fso = New Scripting.FileSystemObject
    fnBuildQualifiedFileName = fso.BuildPath(strDir, strFileName)
PROC_EXIT:
    On Error GoTo 0     ' disable error handler

    ' Clean-up statements go here
    fnFreeObject fso
    
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
Public Sub fnCenterFormOnMDI(ByVal frmMDIParent As Form, ByRef frmMDIChild As Form)
    ' Comments  : Centers the form on the MDI parent
    ' Parameters: none
    ' Returns   : Nothing
    ' Source    : www.vbexplorer.com
    '
    Const cstrCurrentProc As String = "fnCenterFormOnMDI"
    Dim intTop As Integer
    Dim intLeft As Integer
    On Error GoTo PROC_ERR

    If frmMDIParent.WindowState = vbNormal Then
        intTop = ((frmMDIParent.Height - frmMDIChild.Height) \ 2)
        intLeft = ((frmMDIParent.Width - frmMDIChild.Width) \ 2)
        frmMDIChild.Move intLeft, intTop
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
Public Sub fnCenterFormOnScreen(ByRef frmIn As Form)
    ' Comments  : Centers the form on the screen
    ' Parameters: none
    ' Returns   : Nothing
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "fnCenterFormOnScreen"
    On Error GoTo PROC_ERR

    frmIn.Move (Screen.Width - frmIn.Width) / 2, (Screen.Height - frmIn.Height) / 2
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
Public Sub fnDeallocateGlobalObjects()
    ' Comments  : This procedure releases the memory allocated for global objects. It should
    '             be called before ANY application termination (whether due to a fatal error
    '             or the user's inititiation.
    '             IMPORTANT: The Global Error Handler should be the last object deallocated,
    '                        hence it is deallocated when the MDI Main form is unloaded.
    '
    '             NOTE:  This procedure (and possibly sub Main in modStartup.bas)
    '                    should be updated as global object variables are added to
    '                    or removed from the application!
    '
    ' Parameters: none
    ' Returns   : Nothing
    '
    ' Called by : fnTerminateTheApp() of modGeneral.bas   (user-initiated app termination)
    '             ReportFatalError() of cerhErrorHandler.bas   (fatal error)
    '
    ' Source    : Total Visual SourceBook 2000
    '
    ' Modified  :
    ' 04/30/02 BAW (Phase 2B, but 2B004) Made the Crystal Application object a global variable:
    '              defined in modReporting; instantiated in modStartup; deallocated in
    '              fnDeallocateGlobalObjects. This avoids "Out of memory" errors
    '              when the frmReportViewer screen is displayed.
    '
    Const cstrCurrentProc As String = "fnDeallocateGlobalObjects"
    On Error GoTo PROC_EXIT

    If bDebugAppTermination Then
        Debug.Print "   Freeing gadwApp from fnDeallocateGlobalObjects"
    End If
    fnFreeObject gadwApp
    
    If bDebugAppTermination Then
        Debug.Print "   Freeing gconAppActive from fnDeallocateGlobalObjects"
    End If
    fnFreeObject gconAppActive
    
    If bDebugAppTermination Then
        Debug.Print "   Freeing gapsApp from fnDeallocateGlobalObjects"
    End If
    fnFreeObject gapsApp
    
    If bDebugAppTermination Then
        Debug.Print "   Freeing gcrxApp from fnDeallocateGlobalObjects"
    End If
    fnFreeObject gcrxApp
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
Public Sub fnEnableDisableControl(ByVal ctlIn As Control, Optional ByVal bEnable As Boolean = True)
    '--------------------------------------------------------------------------
    ' Procedure:   fnEnableDisableControl
    ' Description: Given a control, either make it look and behave Enabled or
    '              Disabled, depending on the bEnable parameter
    '
    ' Params:      n/a
    '    ctlIn   (in) The control to enable/disable
    '    bEnable (in) True to enable the specified control; False to disable it.
    '
    ' Returns:     N/A
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc       As String = "fnEnableDisableControl"
    
    On Error GoTo PROC_ERR

    With ctlIn
        Select Case bEnable
            Case True
                ' Temporary turn off error handling, since some controls do not support the .Locked property
                On Error Resume Next
                .Locked = False
                On Error GoTo PROC_ERR
                .TabStop = True
                .BackColor = vbWindowBackground
                .ForeColor = vbWindowText
                .Enabled = True
            Case False
                ' Temporary turn off error handling, since some controls do not support the .Locked property
                On Error Resume Next
                .Locked = True
                On Error GoTo PROC_ERR
                .TabStop = False
                .BackColor = vbButtonFace
                .ForeColor = vbButtonText
                .Enabled = False
        End Select
    End With
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
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
Public Function fnFirstDayOfMonth(ByVal dteIn As Date) As Date
    ' Comments  : Calculates the first day of the month for the specified date.
    ' Parameters: dteIn - Date for which first DOM will be determined
    '
    ' Returns   : The date representing the first day of that month
    ' Source    : <http://www.vb-world.net/misc/tip479.html>
    ' Modified  :
    '
    ' --------------------------------------------------
    Const cstrCurrentProc As String = "fnFirstDayOfMonth"
    Dim intMonth          As Integer
    Dim intYear           As Integer
    On Error GoTo PROC_ERR
    
    intMonth = Month(dteIn)
    intYear = Year(dteIn)

    fnFirstDayOfMonth = DateSerial(intYear, intMonth, 1)
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
Public Sub fnFreeObject(ByRef pObj As Object)
    ' Comments  : Safely frees memory used by an object
    '
    '             NOTE: Use fnFreeRecordset( ) for objects
    '                   of type "ADODB.Recordset" since this
    '                   will ensure its DBMS resources are
    '                   released.
    '
    ' Parameters: pObj (in/out) - pointer to the object to free
    '
    ' Called by : Lots of places (usually in the PROC_EXIT block)
    '
    ' Returns   : N/A
    '
    ' Modified  :
    '
    ' --------------------------------------------------
    If fnIsObject(pObj) Then
        Set pObj = Nothing
    End If
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub fnFreeRecordset(ByRef pRST As ADODB.Recordset)
    ' Comments  : Safely frees memory used by an ADODB.Recordset
    '             object after first ensuring it is closed.
    '
    ' Parameters: pRST (in/out) - pointer to the object to free
    '
    ' Called by : Lots of places (usually in the PROC_EXIT block)
    '
    ' Returns   : N/A
    '
    ' Modified  :
    '
    ' --------------------------------------------------
    If fnIsObject(pRST) Then
        If pRST.State = adStateOpen Then
            On Error Resume Next    ' Guard against 3219 "Operation not allowed in this context" error
            pRST.Close
            On Error GoTo 0
        End If
        Set pRST = Nothing
    End If
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function fnGetExtPart(pstrIn As String) As String
    ' Comments  : Returns the extension of a fully qualified file name
    ' Parameters: strIn - path and name to parse
    ' Returns   : file extension
    ' Source    : Shamelessly plagurized from Total Visual SourceBook 2000
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnGetExtPart"
    
    Dim intCounter As Integer
    Dim strTmp As String
       
    ' Parse the string
    For intCounter = Len(pstrIn) To 1 Step -1
        ' It its a slash, grab the sub string
        If Mid$(pstrIn, intCounter, 1) <> "." Then
            strTmp = Mid$(pstrIn, intCounter, 1) & strTmp
        Else
            Exit For
        End If
    Next intCounter
    
    ' Return the value
    fnGetExtPart = UCase$(strTmp)
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

'MME START - ADDED InsuredClmID and PayeDthbPmtAmt to paramater list

Public Sub fnGetStateInfo(ByVal strStateIn As String, ByVal strLOBIn As String, _
    ByVal dtePmtDt As Date, ByVal InsuredClmID As Long, ByVal PayeDthbPmtAmt As Long, ByRef siInOut As StateInfo)
    ' Comments  : Retrieves info from the State98 table and returns data
    '             from the selected row in a UDT called StateInfo
    ' Parameters: strWhereIn (in) - the WHERE clause for the SQL query that selects a
    '                               particular row, e.g., "[State] = 'Alabama'"
    '             siIn (in/out)   - a StateInfo (UDT) structure that will hold the contents of
    '                               the selected row
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc           As String = "fnGetStateInfo"
    Const mcstrGroupLOB             As String = "G"
    Const mcstrIndividualLOB        As String = "I"
    Dim rstTemp                     As ADODB.Recordset


    'MME START - WRUS 4999
    
    Const mcstrGroupTier2LOB        As String = "H"
    Const mcstrIndividualTier2LOB   As String = "J"
    Const mcstrPROOFDEATH           As String = "PROOFDTH"
    Const mcstrPROOF                As String = "PROOF"
    Const mcstrDEATH                As String = "DEATH"
    Const mcstrAMOUNT               As String = "AMOUNT"
    Dim mcstrStrltIdtypCd           As String
    Dim rstTierTemp                 As ADODB.Recordset
    Dim DtResultDate                As Date
    Dim DtProofDate                 As Date
    Dim DtDeathDate                 As Date
    Dim dblAmount                   As Double
    Dim DblCompareVal               As Double
    Dim DtDthProofDifference        As Long
    Dim rstSingleRecord_Fresh       As ADODB.Recordset

 
    'fnLogWrite "      In fnGetStateInfo, getting State Tier info: " & strStateIn & " " & strLOBIn & " " & CStr(DateValue(dtePmtDt)), cstrCurrentProc
 
    Set rstTierTemp = fnGetStateTierInfo(strStateIn, strLOBIn, dtePmtDt)
 
    If rstTierTemp.RecordCount <> 0 Then
    
       Set rstSingleRecord_Fresh = fnSelectRecord(InsuredClmID)
      
       If rstSingleRecord_Fresh.RecordCount <> 0 Then
       
          mcstrStrltIdtypCd = Trim(rstTierTemp.Fields(3))
          
          DtProofDate = rstSingleRecord_Fresh.Fields(8).value
          DtDeathDate = rstSingleRecord_Fresh.Fields(2).value
          dblAmount = rstSingleRecord_Fresh.Fields(10).value
          
          If rstTierTemp.Fields(4) < 0 Then
             DblCompareVal = rstTierTemp.Fields(4) * -1
          Else
             DblCompareVal = rstTierTemp.Fields(4)
          End If
          
          Select Case mcstrStrltIdtypCd
       
                 Case mcstrPROOFDEATH
                    DtDthProofDifference = DateDiff("d", DtDeathDate, DtProofDate)
                    If rstTierTemp.Fields(4) < 0 Then                              'IF NEGATIVE VALUE, THEN REVERSE E.G. WITHIN TIMEFRAME, THEN TIER2
                       If DtDthProofDifference <= DblCompareVal Then               'DtDthProofDifference is difference in days
                          If strLOBIn = mcstrGroupLOB Then
                             strLOBIn = mcstrGroupTier2LOB
                          Else
                             strLOBIn = mcstrIndividualTier2LOB
                          End If
                       End If
                    Else
                       If DtDthProofDifference > DblCompareVal Then                'DtDthProofDifference is difference in days
                          If strLOBIn = mcstrGroupLOB Then
                             strLOBIn = mcstrGroupTier2LOB
                          Else
                             strLOBIn = mcstrIndividualTier2LOB
                          End If
                       End If
                    End If
                    
                Case mcstrPROOF
                    DtResultDate = (DateAdd("d", DblCompareVal, DtProofDate))
                    If rstTierTemp.Fields(4) < 0 Then                              'IF NEGATIVE VALUE, THEN REVERSE E.G. WITHIN TIMEFRAME, THEN TIER2
                       If DateValue(dtePmtDt) <= DateValue(DtResultDate) Then      'DtResultDate is cutoff date
                          If strLOBIn = mcstrGroupLOB Then
                             strLOBIn = mcstrGroupTier2LOB
                          Else
                             strLOBIn = mcstrIndividualTier2LOB
                          End If
                       End If
                    Else
                       If DateValue(dtePmtDt) > DateValue(DtResultDate) Then        'NOT WITHIN TIMEFRAME, THEN TIER2
                          If strLOBIn = mcstrGroupLOB Then
                             strLOBIn = mcstrGroupTier2LOB
                          Else
                             strLOBIn = mcstrIndividualTier2LOB
                          End If
                       End If
                    End If
                
                 Case mcstrDEATH
                    DtResultDate = (DateAdd("d", DblCompareVal, DtDeathDate))
                    If rstTierTemp.Fields(4) < 0 Then                               'IF NEGATIVE VALUE, THEN REVERSE E.G. WITHIN TIMEFRAME, THEN TIER2
                       If DateValue(dtePmtDt) <= DateValue(DtResultDate) Then       'DtResultDate is cutoff date
                          If strLOBIn = mcstrGroupLOB Then
                             strLOBIn = mcstrGroupTier2LOB
                          Else
                             strLOBIn = mcstrIndividualTier2LOB
                          End If
                       End If
                    Else
                       If DateValue(dtePmtDt) > DateValue(DtResultDate) Then        'NOT WITHIN TIMEFRAME, THEN TIER2
                          If strLOBIn = mcstrGroupLOB Then
                             strLOBIn = mcstrGroupTier2LOB
                          Else
                             strLOBIn = mcstrIndividualTier2LOB
                          End If
                       End If
                    End If
          
                 Case mcstrAMOUNT
                    If rstTierTemp.Fields(4) < 0 Then                               'IF NEGATIVE VALUE, THEN REVERSE E.G. WITHIN TOLERANCE, THEN TIER2
                       If PayeDthbPmtAmt > DblCompareVal Then
                          If strLOBIn = mcstrGroupLOB Then
                             strLOBIn = mcstrGroupTier2LOB
                          Else
                             strLOBIn = mcstrIndividualTier2LOB
                          End If
                       End If
                    Else
                       If PayeDthbPmtAmt <= DblCompareVal Then                      'NOT WITHIN TOLERANCE, THEN TIER2
                          If strLOBIn = mcstrGroupLOB Then
                             strLOBIn = mcstrGroupTier2LOB
                          Else
                             strLOBIn = mcstrIndividualTier2LOB
                          End If
                       End If
                    End If
         
                 Case Else
                    ' Invalid record found on table STATE_RULE_TIER_T (4012) -
                    ' for the state of @@1 as of @@2. The calculations cannot be done.
                    gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_INVALID_ENTRY_RULE_TIER_T, _
                                       mcstrName & cstrCurrentProc, _
                                       strStateIn, CStr(DateValue(dtePmtDt))
                    GoTo PROC_EXIT
         End Select
      Else
         If (rstSingleRecord_Fresh Is Nothing) Or (rstSingleRecord_Fresh.RecordCount = 0) Then
            ' Claim has been deleted by another user -
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_INFO_ANOTHER_USER_DELETED_NOT_SAVED, _
            mcstrName & cstrCurrentProc, _
            InsuredClmID
            GoTo PROC_EXIT
        End If
      End If
   End If
   
  'MME END - WRUS 4999


    'fnLogWrite "      In fnGetStateInfo, getting new rule: " & strStateIn & " " & strLOBIn & " " & CStr(DateValue(dtePmtDt)), cstrCurrentProc

    Set rstTemp = fnGetStateRule(strStateIn, strLOBIn, dtePmtDt)

    ' If we didn't find such a row:
    '    * if we had been looking for a Group LOB (a common situation), then
    '      try for the row that has an Individual LOB (I). There should be one!
    '    * if we had been looking for an Individual LOB, then something is
    '      very wrong. Every state should have an Individual row, but there
    '      will most likely only be Group rows for a small handful of states
    '      (like Georgia).
    ' Group is supposed to default to using Individual rates if no
    ' Group-specific rates are defined for a given state.
    If rstTemp.RecordCount = 0 Then
        If strLOBIn = mcstrGroupLOB Then
            Set rstTemp = fnGetStateRule(strStateIn, mcstrIndividualLOB, dtePmtDt)
            If rstTemp.RecordCount = 0 Then
                
                ' gcRES_NERR_STATE_RATES_NOT_FOUND (4006) - Neither Group nor Individual
                ' rates were found for the state of @@1 as of @@2. The calculations cannot be done.
                gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_STATE_RATES_NOT_FOUND, _
                                           mcstrName & cstrCurrentProc, _
                                           strStateIn, CStr(DateValue(dtePmtDt))
                GoTo PROC_EXIT
            End If
        Else
            ' gcRES_NERR_INDV_STATE_RATES_NOT_FOUND (4007) - Individual rates were not found
            ' for the state of @@1 as of @@2. The calculations cannot be done.
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_INDV_STATE_RATES_NOT_FOUND, _
                                       mcstrName & cstrCurrentProc, _
                                       strStateIn, CStr(DateValue(dtePmtDt))
            GoTo PROC_EXIT
        End If
    End If
    
    With rstTemp
        siInOut.StCd = !st_cd
        siInOut.LobCd = !lob_cd
        siInOut.StrlEffDt = !strl_eff_dt
        siInOut.CalcIdtypCd = !calc_idtyp_cd
        siInOut.ReqdIdtypCd = !reqd_idtyp_cd
        siInOut.IruleCd = !irule_cd
        siInOut.StrlEndDt = !strl_end_dt
        siInOut.StrlIntRptgFlrAmt = !strl_int_rptg_flr_amt
        siInOut.StrlIntCalcOfstNum = !strl_int_calc_ofst_num
        siInOut.StrlIntReqdOfstNum = !strl_int_reqd_ofst_num
        siInOut.StrlIntRuleAmt = !strl_int_rule_amt
        siInOut.StrlSpclInstrTxt = !strl_spcl_instr_txt
        'fnLogWrite "      In fnGetStateInfo, got: " & !st_cd & " " & !lob_cd & " " & CStr(DateValue(!strl_eff_dt)), cstrCurrentProc
        .Close
    End With
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    
    ' Clean-up statements go here
    fnFreeRecordset rstTemp
    
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


' MME START WRUS 4999

'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function fnGetStateTierInfo(ByVal strStateIn As String, ByVal strLOBIn As String, _
    ByVal dtePmtDt As Date) As ADODB.Recordset
    ' Comments  : Retrieves info from the State_rule_tier_t and returns a value
    '             based on cals performed against the data on the selected row.
    ' Parameters: strWhereIn (in) - the WHERE clause for the SQL query that selects a
    '                               particular row, e.g., "[State] = 'Alabama'"
    '             siIn (in/out)   - a String that will hold pass back a value of 'G', 'H', 'I', or 'J'
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc           As String = "fnGetStateTierInfo"
    Const cstrSproc                 As String = "dbo.proc_state_rule_tier_t"  ' Stored procedure to execute
 
    Dim adwTemp                     As cadwADOWrapper
    Dim rstTemp                     As ADODB.Recordset
    Dim prmReturnValue              As ADODB.Parameter
    Dim prmStCd                     As ADODB.Parameter
    Dim prmLobCd                    As ADODB.Parameter
    Dim prmPayePmtDt                As ADODB.Parameter
    
    'fnLogWrite "      In fnGetStateTierInfo, getting new rule: " & strStateIn & " " & strLOBIn & " " & CStr(DateValue(dtePmtDt)), cstrCurrentProc
   
    
    Set adwTemp = New cadwADOWrapper
    
    If Not (adwTemp.CommandSetSproc(cstrSproc)) Then
        GoTo PROC_EXIT
    End If

    With adwTemp.ADOCommand
        ' ---Parameter #1---
        ' Define the return value that represents the error code (i.e. reason) why
        ' the stored procedure failed.
        Set prmReturnValue = .CreateParameter(Name:="@return_value", _
                                              Type:=adInteger, _
                                              Direction:=adParamReturnValue, _
                                              value:=Null)
        .Parameters.Append prmReturnValue

        ' ---Parameter #2---
        ' Define the LOB_CD parameter
        Set prmLobCd = .CreateParameter(Name:="@lob_cd", _
                                        Type:=adChar, _
                                        Direction:=adParamInput, _
                                        Size:=1, _
                                        value:=strLOBIn)
        .Parameters.Append prmLobCd

        ' ---Parameter #3---
        ' Define the ST_CD parameter
        Set prmStCd = .CreateParameter(Name:="@st_cd", _
                                        Type:=adChar, _
                                        Direction:=adParamInput, _
                                        Size:=2, _
                                        value:=strStateIn)
        .Parameters.Append prmStCd

        ' ---Parameter #4---
        ' Define the PAYE_PMT_DT parameter
        Set prmPayePmtDt = .CreateParameter(Name:="@paye_pmt_dt", _
                                         Type:=adDBTimeStamp, _
                                         Direction:=adParamInput, _
                                         Size:=8, _
                                         value:=dtePmtDt)
        .Parameters.Append prmPayePmtDt

        Set rstTemp = .Execute()
        rstTemp.ActiveConnection = Nothing
               
        ' The rstTemp recordset may well be empty. That's okay though since the caller
        ' (fnGetStateRule) can accomodate this...either by looking for a different LOB's row
        ' or by generating an error of its own accord.
    End With

    Set fnGetStateTierInfo = rstTemp
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    
    ' Clean-up statements go here
    
    ' DO NOT do "fnFreeRecordset rstTemp" since this will cause the recordset
    ' returned by this function to be wiped out as well!
    fnFreeObject prmStCd
    fnFreeObject prmLobCd
    fnFreeObject prmPayePmtDt
    fnFreeObject prmReturnValue
    
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case prmReturnValue
        Case gcRES_NERR_REC_NOT_FOUND       ' 4027 -  The @@1 is invalid. @@2
            ' Note that the following error is presented as an ATYPICAL 4027 error!
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_INVALID_DATA, _
                                       mcstrName & cstrCurrentProc, _
                                       "State, Line of Business, or Date of Payment", _
                                       "The Calculation Rule cannot be retrieved " & _
                                       "when any of these fields are NULL or if the State Code cannot be found in the State table. It may also " & _
                                       "be that no Calculation Rule is in effect " & _
                                       "for the State for the given Date of Payment. State=[" & strStateIn & "], LOB=[" & strLOBIn & _
                                       "], Payment Date=[" & FormatDateTime(dtePmtDt, vbShortDate) & "]"
            Resume PROC_EXIT
        Case gcRES_NERR_ERR_WHILE_TRYING_TO   ' 4028
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_ERR_WHILE_TRYING_TO, _
                                       mcstrName & cstrCurrentProc, _
                                       "locate"
            Resume PROC_EXIT '
    End Select

    ' If any other errors exist, i.e. in Err object, then let it fall through into default error handling.

    Select Case Err.Number
        Case -2147217900 ' Object not found
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_FERR_SPROC_NOT_FOUND, _
                                       mcstrName & cstrCurrentProc, _
                                       cstrSproc
            Resume PROC_EXIT
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    
    Resume PROC_EXIT
End Function

' MME END WRUS 4999


'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function fnGetStateRule(ByVal strStateIn As String, ByVal strLOBIn As String, _
    dtePayePmtDt As Date) As ADODB.Recordset
    ' Comments  : Retrieves info from the STATE_RULE_T and returns data
    '             from the selected row in
    ' Parameters: strStateIn (in)   - the desired state code
    '             strLOBIn (in)     - the desired line-of-business
    '             dtePayePmtDt (in) - the date as of which to retrieve the state rule info
    ' Returns   : An ADODB.Recordset
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc           As String = "fnGetStateRule"
    Const cstrSproc                 As String = "dbo.proc_state_rule_select"  ' Stored procedure to execute
    Dim adwTemp                     As cadwADOWrapper
    Dim rstTemp                     As ADODB.Recordset
    Dim prmReturnValue              As ADODB.Parameter
    Dim prmStCd                     As ADODB.Parameter
    Dim prmLobCd                    As ADODB.Parameter
    Dim prmPayePmtDt                As ADODB.Parameter
    
    Set adwTemp = New cadwADOWrapper
    
    If Not (adwTemp.CommandSetSproc(cstrSproc)) Then
        GoTo PROC_EXIT
    End If

    With adwTemp.ADOCommand
        ' ---Parameter #1---
        ' Define the return value that represents the error code (i.e. reason) why
        ' the stored procedure failed.
        Set prmReturnValue = .CreateParameter(Name:="@return_value", _
                                              Type:=adInteger, _
                                              Direction:=adParamReturnValue, _
                                              value:=Null)
        .Parameters.Append prmReturnValue

        ' ---Parameter #2---
        ' Define the LOB_CD parameter
        Set prmLobCd = .CreateParameter(Name:="@lob_cd", _
                                        Type:=adChar, _
                                        Direction:=adParamInput, _
                                        Size:=1, _
                                        value:=strLOBIn)
        .Parameters.Append prmLobCd

        ' ---Parameter #3---
        ' Define the ST_CD parameter
        Set prmStCd = .CreateParameter(Name:="@st_cd", _
                                        Type:=adChar, _
                                        Direction:=adParamInput, _
                                        Size:=2, _
                                        value:=strStateIn)
        .Parameters.Append prmStCd

        ' ---Parameter #4---
        ' Define the PAYE_PMT_DT parameter
        Set prmPayePmtDt = .CreateParameter(Name:="@paye_pmt_dt", _
                                         Type:=adDBTimeStamp, _
                                         Direction:=adParamInput, _
                                         Size:=8, _
                                         value:=dtePayePmtDt)
        .Parameters.Append prmPayePmtDt

        Set rstTemp = .Execute()
        rstTemp.ActiveConnection = Nothing
               
        ' The rstTemp recordset may well be empty. That's okay though since the caller
        ' (fnGetStateRule) can accomodate this...either by looking for a different LOB's row
        ' or by generating an error of its own accord.
    End With

    Set fnGetStateRule = rstTemp
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    
    ' Clean-up statements go here
    
    ' DO NOT do "fnFreeRecordset rstTemp" since this will cause the recordset
    ' returned by this function to be wiped out as well!
    fnFreeObject prmStCd
    fnFreeObject prmLobCd
    fnFreeObject prmPayePmtDt
    fnFreeObject prmReturnValue
    
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case prmReturnValue
        Case gcRES_NERR_REC_NOT_FOUND       ' 4027 -  The @@1 is invalid. @@2
            ' Note that the following error is presented as an ATYPICAL 4027 error!
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_INVALID_DATA, _
                                       mcstrName & cstrCurrentProc, _
                                       "State, Line of Business, or Date of Payment", _
                                       "The Calculation Rule cannot be retrieved " & _
                                       "when any of these fields are NULL or if the State Code cannot be found in the State table. It may also " & _
                                       "be that no Calculation Rule is in effect " & _
                                       "for the State for the given Date of Payment. State=[" & strStateIn & "], LOB=[" & strLOBIn & _
                                       "], Payment Date=[" & FormatDateTime(dtePayePmtDt, vbShortDate) & "]"
            Resume PROC_EXIT
        Case gcRES_NERR_ERR_WHILE_TRYING_TO   ' 4028
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_ERR_WHILE_TRYING_TO, _
                                       mcstrName & cstrCurrentProc, _
                                       "locate"
            Resume PROC_EXIT '
    End Select

    ' If any other errors exist, i.e. in Err object, then let it fall through into default error handling.

    Select Case Err.Number
        Case -2147217900 ' Object not found
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_FERR_SPROC_NOT_FOUND, _
                                       mcstrName & cstrCurrentProc, _
                                       cstrSproc
            Resume PROC_EXIT
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    
    Resume PROC_EXIT
End Function


'////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub fnInitializeStateInfo(ByRef siInOut As StateInfo)
    ' Comments  : Initializes the specified UDT called StateInfo
    ' Parameters:
    '       siIn (in/out)   - a StateInfo (UDT) structure that will hold the contents of
    '                         the selected row
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc           As String = "fnInitializeStateInfo"

    With siInOut
        .StCd = gcstrBlankEntry
        .StrlEffDt = CDate(Now)
        .CalcIdtypCd = vbNullString
        .ReqdIdtypCd = vbNullString
        .IruleCd = vbNullString
        .StrlEndDt = vbNull
        .StrlIntRptgFlrAmt = mcintZero
        .StrlIntCalcOfstNum = mcintZero
        .StrlIntReqdOfstNum = mcintZero
        .StrlIntRuleAmt = mcintZero
        .StrlSpclInstrTxt = vbNullString
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
Public Function fnIsFormLoaded(ByVal strFormName As String, Optional ByRef frmFound As Form) As Boolean
    '----------------------------------------------------------------------------
    ' Comments  : Tests to see whether the default instance
    '             of a form is loaded
    ' Parameters: strFormName (in) - name of form to search for
    '             frmFound (out)   - pointer to the searched-for form, if found
    '
    ' Returns   : True if the form is loaded, false otherwise
    ' Source    : Total Visual SourceBook 2000
    '----------------------------------------------------------------------------
    Dim frm As Form
    Dim bResult As Boolean
    Const cstrCurrentProc As String = "fnIsFormLoaded"

    On Error GoTo PROC_ERR

    ' If a form is loaded, it will be in the Forms collection.
    ' Search this collection to see if the specified form
    ' is present.
    For Each frm In Forms
        If UCase$(frm.Name) = UCase$(strFormName) Then
            bResult = True
            Set frmFound = frm
            Exit For
        End If
    Next frm

    fnIsFormLoaded = bResult
PROC_EXIT:
    On Error GoTo 0     ' disable error handler

    ' Clean-up statements go here
    fnFreeObject frm

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
Public Function fnIsObject(ByVal objIn As Object) As Boolean
    '----------------------------------------------------------------------------
    ' Comments  : Safe test to see if object exists (better than "If IsObject()"
    ' Parameters: objIn (in) - Object reference
    '
    ' Returns   : True if the object has been initialized, false otherwise
    '----------------------------------------------------------------------------
    Const cstrCurrentProc As String = "fnIsObject"

    fnIsObject = Not (objIn Is Nothing)
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function fnLastDayOfMonth(ByVal dteIn As Date) As Date
    ' Comments  : Calculates the last day of the month for the specified date.
    ' Parameters: dteIn - Date for which last DOM will be determined
    '
    ' Returns   : The date representing the last day of that month
    ' Source    : <http://www.vb-world.net/misc/tip479.html>
    ' Modified  :
    '
    ' --------------------------------------------------
    Const cstrCurrentProc  As String = "fnLastDayOfMonth"
    Dim intLastDay         As Integer
    
    On Error GoTo PROC_ERR

    intLastDay = DatePart("d", _
                       DateAdd("d", -1, _
                       DateAdd("m", 1, _
                       DateAdd("d", -DatePart("d", dteIn) + 1, dteIn))))

    fnLastDayOfMonth = DateSerial(Year(dteIn), Month(dteIn), intLastDay)
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
Public Function fnMakeWeekday(ByVal dteIn As Date, ByVal intDirection As EnumPrevNext) As Date
    ' Comments  : Returns the input date, coerced to the next or previous weekday (based
    '             on the intDirection paramter) if the input date fell on a weekend.
    '             Date returned is always between Monday and Friday.
    ' Parameters: dteIn        (in) - Date to coerce
    '             intDirection (in) - indicates whether to move to the next or previous weekday
    ' Returns   : Coerced date
    ' Source    : Based on Total Visual SourceBook 2000's PriorWeekday and NextWeekday functions
    '
    Const cstrCurrentProc As String = "fnMakeWeekday"
    Dim dteTemp           As Date
    
    On Error GoTo PROC_ERR
  
    If Weekday(dteIn) = vbSaturday Or Weekday(dteIn) = vbSunday Then
        Select Case intDirection
            Case epnPrev
                dteTemp = dteIn - 1
                While Weekday(dteTemp) = vbSunday Or Weekday(dteTemp) = vbSaturday
                    dteTemp = dteTemp - 1
                Wend
            Case epnNext
                dteTemp = dteIn + 1
                While Weekday(dteTemp) = vbSunday Or Weekday(dteTemp) = vbSaturday
                    dteTemp = dteTemp + 1
                Wend
        End Select
        fnMakeWeekday = dteTemp
    Else
        fnMakeWeekday = dteIn
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


'//////////////////////////////////////////////////////////////////////////////////////////
Public Function fnOpenFileInDefaultApp(ByVal strFile As String) As Boolean
    ' Comments  : Opens the specified file in the application that is
    '             associated with that kind of file.
    ' Parameters: strFile - the fully-qualified file to open
    ' Called by : mnuHelpViewApplicationLogFile_Click() of frmMDIMain
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc   As String = "fnOpenFileInDefaultApp"
    Dim lngReturnCode       As Long

    lngReturnCode = ShellExecute(0&, "open", strFile, _
        vbNullString, vbNullString, vbNormalFocus)

    fnOpenFileInDefaultApp = (lngReturnCode > 32)
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
            ' gcRES_INFO_CANT_OPEN_FILE (1014)
            ' Unable to open <@@1>. The file either does not exist or no application is associated with files of type .TXT.
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_INFO_CANT_OPEN_FILE, _
                                    mcstrName & cstrCurrentProc, _
                                    strFile
    End Select
    Resume PROC_EXIT
End Function



'!TODO! Make version that pads with left with leading zeroes
'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function fnPadRightString(ByVal strIn As String, ByVal lngStrLen As Long, _
    Optional ByVal strPadCharIn As String = " ") As String
    ' Comments  : Right pads a string for left justification
    ' Parameters: strIn - String to pad
    '             strPadCharIn - Character to use for padding
    '             lngStrLen - Desired length of string
    '
    ' Returns   : right padded string
    ' Source    : Total Visual SourceBook 2000
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnPadRightString"
  
    fnPadRightString = Left$(strIn & String$(lngStrLen, Left$(strPadCharIn, 1)), _
        lngStrLen)
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
Public Sub fnPersistRecordsetToCSV(ByRef rstIn As ADODB.Recordset, ByVal strFileNm As String)
    ' Comments  : Persists the specified ADO recordset to the specified tab-delimited file.
    '
    ' Parameters: rstIn     (in) - an ADO recordset
    '             strFileNm (in) - the fully qualified filename to which to persist the rst
    '
    ' Returns   : N/A
    '
    ' Modified  :
    '
    ' --------------------------------------------------
    Const cintBlockSize As Integer = 100
    Dim strTemp         As String
    Dim lngCounter      As Long
    Dim fld             As ADODB.Field
    On Error Resume Next
    
    ' Delete the file, if it exists, effectively overwriting it.
    Kill strFileNm
    
    Open strFileNm For Output As #1
    
    With rstIn
        ' Print line with field names and headings
        For Each fld In .Fields
            strTemp = strTemp & fld.Name & ","
        Next fld
        ' Drop trailling comma
        strTemp = Left$(strTemp, Len(strTemp) - 1)
        Print #1, strTemp
            
        lngCounter = 1
        
        If Not (.BOF And .EOF) Then
            .MoveFirst
        End If
        
        Do Until .EOF
            If lngCounter <> 1 Then
                strTemp = .GetString(, cintBlockSize, """, """, """" & vbCrLf & """", "")
            Else
                ' Prepend the opening quotes for the first field of the first row
                strTemp = """" & .GetString(, cintBlockSize, """, """, """" & vbCrLf & """", "")
            End If
            If .EOF Then
                ' Drop the double quote character printed in excess
                strTemp = Left$(strTemp, Len(strTemp) - 1)
            End If
            Print #1, strTemp;
        Loop
    
        ' Go back to the first record, so subsequent accesses won't be starting at .EOF
        If Not (.BOF And .EOF) Then
            .MoveFirst
        End If
    End With
    
    Close #1
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub fnPersistRecordsetToXML(ByRef rstIn As ADODB.Recordset, ByVal strFileNm As String)
    ' Comments  : Persists the specified ADO recordset to the specified file.
    '
    ' Parameters: rstIn     (in) - an ADO recordset
    '             strFileNm (in) - the fully qualified filename to which to persist the rst
    '
    ' Returns   : N/A
    '
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error Resume Next
    
    ' Delete the file, if it exists, effectively overwriting it.
    Kill strFileNm
    
    rstIn.Save strFileNm, adPersistXML
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function fnQuoted(ByVal strIn As String) As String
    ' Comments  : Returns the input string, surrounded by single quotes,
    '             for use with building SQL statements or values that
    '             will be used as parameters to stored procdures.
    ' Parameters: strIn (in) the string to surround
    '
    ' Returns   : quoted string, e.g., xxx  ==> 'xxx'
    '
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnQuoted"

    'fnQuoted = gcstrDoubleQuote & strIn & gcstrDoubleQuote
    fnQuoted = gcstrSingleQuote & strIn & gcstrSingleQuote
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
Public Sub fnRemoveCloseButton(ByVal frmIn As Form)
    '----------------------------------------------------------------------------
    ' Procedure :  Sub fnRemoveCloseButton
    '
    ' Comments  :  Remove the Close command from the system menu and disable
    '              the use of Alt-F4, for the specified form.
    ' Called by :  frmLogOn
    ' Parameters:  frmIn (in) - pointer to the Form object to process
    '
    ' Modified  :
    '----------------------------------------------------------------------------
    Const cstrCurrentProc As String = "fnRemoveCloseButton"
    Const clngMF_BYPOSITION = &H400&
    Dim lngHmenu As Long
    Dim lngItemCount As Long

    On Error Resume Next

    ' Get the handle of the system menu
    lngHmenu = GetSystemMenu(frmIn.hWnd, 0)

    ' Remove the system menu Close menu item
    RemoveMenu lngHmenu, 6, clngMF_BYPOSITION
    ' Remove the system menu separator line
    RemoveMenu lngHmenu, 5, clngMF_BYPOSITION
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function fnRemoveUnderScoresFromFieldName(ByVal strIn As String) As String
    ' Comments  : Removes underscores from the specified string
    ' Parameters:
    '   strIn (in) - string from which to remove underscore characters
    '
    ' Returns   : string without underscores
    '
    ' Called by :
    '
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnRemoveUnderScoresFromFieldName"
   
   fnRemoveUnderScoresFromFieldName = Replace(strIn, "_", "")
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
Public Function fnRound(ByVal dblAmountIn As Double, ByVal intSignIn As Integer) As Double
    ' Comments  :
    ' Parameters: dblAmountIn - the amount to round
    '             intSignIn -
    ' Returns   : Double - the rounded version of dblAmountIn
    '
    ' Called by :
    '
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnRound"
    Dim lngSignChange As Long

    If (dblAmountIn >= 0) Then
        lngSignChange = 1
    Else
        lngSignChange = -1
    End If

    fnRound = dblAmountIn + (0.5 * 0.1 ^ intSignIn) * lngSignChange
    fnRound = fnRound * 10 ^ intSignIn
    fnRound = Fix(fnRound)
    fnRound = fnRound / 10 ^ intSignIn
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



' MME START WRUS 4999

'////////////////////////////////////////////////////////////////////////////////////////////////
Private Function fnSelectRecord(ByVal lngKey1 As Long) As ADODB.Recordset
    '--------------------------------------------------------------------------
    ' Procedure:   fnSelectRecord
    ' Description: Selects a single record based on the value(s) in the
    '              properties that correspond to the table's key(s)
    '
    '              NOTE: For each table key, there should be a parameter
    '                    of the appropriate data type!
    '
    ' Parameters:
    '     lngKey1 (in) - the key to the table that should be retrieved
    '
    ' Returns:     A disconnected ADODB.Recordset containing all table columns
    '              for the specified key
    '-----------------------------------------------------------------------------
    
    '!CUSTOMIZE!  This proc and all calls to it must be customized to reflect
    '             one parameter for each key column of the table. Make sure the
    '             parameter is defined to be of the right data type. Also,
    '             the way the recordset's .Find property is set must be changed
    '             to reflect each key column so the right record will be located.
    '             Also make sure that the substitution values passed to
    '             SaveAppSpecificError are correct and TRIM'd if appropriate.
    
    Const cstrCurrentProc          As String = "fnSelectRecord"
    Const cstrSproc                As String = "dbo.proc_claim_select"  ' Stored procedure to execute
    Dim rstTemp                    As ADODB.Recordset
    Dim prmReturnValue             As ADODB.Parameter
    Dim prmClmId                   As ADODB.Parameter
        Dim adwTemp                     As cadwADOWrapper

    On Error GoTo PROC_ERR

    Set adwTemp = New cadwADOWrapper
    
    If Not (adwTemp.CommandSetSproc(cstrSproc)) Then
        GoTo PROC_EXIT
    End If

    ' For Char/VarChar fields,
    '     * Use fnNullIfZLS to ensure Nulls are appropriately handled.
    '     * Do *not* set the optional 2nd parameter to fnNullIfZLS to True.
    ' For numeric fields,
    '     * Use fnNullIfZero to ensure Nulls are appropriately handled.
    ' For Y/N fields,
    '     * Use fnBoolToYN to ensure True/False is appropriately translated.

    With adwTemp.ADOCommand
        ' ---Parameter #1---
        ' Define the return value that represents the error code (i.e. reason) why
        ' the stored procedure failed.
        Set prmReturnValue = .CreateParameter(Name:="@return_value", _
                                              Type:=adInteger, _
                                              Direction:=adParamReturnValue, _
                                              value:=Null)
        .Parameters.Append prmReturnValue

        ' ---Parameter #2---
        ' Define the CLM_ID parameter
        Set prmClmId = .CreateParameter(Name:="@clm_id", _
                                         Type:=adInteger, _
                                         Direction:=adParamInput, _
                                         value:=fnNullIfZero(lngKey1))
        .Parameters.Append prmClmId

        Set rstTemp = .Execute()
    End With
    
    rstTemp.ActiveConnection = Nothing
    Set fnSelectRecord = rstTemp
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
    
    ' Clean-up statements go here
    
    ' Do *not* do "fnFreeRecordset rstTemp" since this will cause the recordset returned
    ' by this function to be wiped out as well!
    fnFreeObject prmReturnValue
    fnFreeObject prmClmId
    
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case prmReturnValue
        Case gcRES_NERR_REC_NOT_FOUND
            ' 4027 = The specified record was not found in the database (@@1).
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_REC_NOT_FOUND, _
                                       mcstrName & cstrCurrentProc, _
                                       "Claim ID " & RTrim$(lngKey1)
            Resume PROC_EXIT
        Case gcRES_NERR_ERR_WHILE_TRYING_TO
            ' 4028 = An error occurred while attempting to @@1 this record.
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_ERR_WHILE_TRYING_TO, _
                                       mcstrName & cstrCurrentProc, _
                                       "locate"
            Resume PROC_EXIT
    End Select

    ' If any other errors exist, i.e. in Err object, then let it fall through into default error handling.

    Select Case Err.Number
        Case -2147217900 ' Object not found
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_FERR_SPROC_NOT_FOUND, _
                                       mcstrName & cstrCurrentProc, _
                                       cstrSproc
            Resume PROC_EXIT
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function

' MME END WRUS 4999



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub fnSetTopmostWindow(ByVal frm As Form, Optional ByVal bTopmost As Boolean = True)
    '----------------------------------------------------------------------------
    ' Procedure :  Function fnSetTopmostWindow
    '
    ' Comments  : Makes a form the topmost window or reverts it to normal status
    ' Source    : VBMaximizer code library
    ' Called by : cmdOK_Click( ) in the frmPrintReport form
    ' Parameters: hWnd (in) - window handle to form to operate upon
    '             bTopmost (in) - True if window should be made topmost; False
    '                  otherwise
    '
    ' Modified  :
    '----------------------------------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnSetTopmostWindow"

    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    Const SWP_NOSIZE = &H1
    Const SWP_NOMOVE = &H2
    Const SWP_NOACTIVATE = &H10

    SetWindowPos frm.hWnd, _
                 IIf(bTopmost, HWND_TOPMOST, HWND_NOTOPMOST), _
                 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
PROC_EXIT:
    ' Use Resume Next rather than GoTo 0 since Close could error if TS wasn't successfully opened
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


'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function fnShowRecordPosition(rstIn As ADODB.Recordset) As String
    '----------------------------------------------------------------------------
    ' Procedure :  Function fnShowRecordPosition
    '
    ' Comments  : Used to build "Record x of y" label on the screen to denote
    '             current record position
    ' Called by : cmdNavigate_Click( ), Form_Load( ) in Insured and Payee forms
    ' Parameters: N/A
    '
    ' Modified  :
    '----------------------------------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnShowRecordPosition"
    Dim strPos As String

    ' It would be nice if this function could just use the table wrapper's Lookup
    ' recordset's public functions, but that's not possible since this function
    ' must support *all* table wrapper's Lookup recordsets. So, we must
    ' continue to receive an ADO Recordset as input (i.e. <tablewrapper>.LookupData)
    ' and go from there.

    Select Case rstIn.AbsolutePosition
        Case adPosBOF
            strPos = "0"
        Case adPosEOF
            strPos = CStr(rstIn.RecordCount)
        Case adPosUnknown
            strPos = "?"
        Case Else
            strPos = CStr(rstIn.AbsolutePosition)
    End Select
    
    fnShowRecordPosition = "Record " & strPos & " of " & rstIn.RecordCount
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
Public Function fnSSNTIN_AddDash(ByVal strIn As String, Optional bIsTin As Boolean = False) As String
    ' Comments  : Returns the input string with a dash added between:
    '             * characters 3 and 4 and 5 and 6, if bIsTin = True
    '             * characters 2 and 3, if bIsTin = False
    ' Parameters: strIn (in) - a 9-digit Social Security Number or
    '                          Taxpayer Identification Number
    '
    ' Returns   : string, e.g., 123456789  ==> '123-45-6789' or '12-3456789'
    '
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnSSNTIN_AddDash"
    Const cstrDash        As String = "-"

    If bIsTin Then
        fnSSNTIN_AddDash = Left$(strIn, 2) & cstrDash & Right$(strIn, 7)
    Else
        fnSSNTIN_AddDash = Left$(strIn, 3) & cstrDash & _
                           Mid$(strIn, 4, 2) & cstrDash & Right$(strIn, 4)
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
Sub fnTerminateTheApp()
    '----------------------------------------------------------------------------
    ' Procedure :  Sub fnTerminateTheApp
    '
    ' Comments  :  Unloads all loaded forms, in the reverse order from which
    '              they were originally ordered
    ' Parameters:  N/A
    '
    ' Called by : cmdCancel_Click() of frmLogOn
    '             MDIForm_Unload() of frmMDIMain
    '
    ' Modified  :
    '----------------------------------------------------------------------------
    Const cstrCurrentProc       As String = "fnTerminateTheApp"
    Const cstrMDIForm           As String = "frmMDIMain"
    Dim intIndex                As Integer
    Dim intFormsCount           As Integer

    On Error GoTo 0     ' No error handler. Nowhere to go :(

    gbAmTryingToTerminateTheApp = True
    
    If bDebugAppTermination Then
        Debug.Print "Entering fnTerminateTheApp"
    End If

   
    ' Unload the Splash screen if it's still around. Otherwise, the attempt to unload
    ' the MDI form will fail and the app won't truly be shut down.
    fnUnloadSplash
    
    intFormsCount = Forms.Count
    
    If bDebugAppTermination Then
        Debug.Print "   Number of forms in memory = " & CStr(intFormsCount)
    End If
        
    If intFormsCount > 0 Then
        For intIndex = intFormsCount - 1 To 0 Step -1
            If bDebugAppTermination Then
                Debug.Print "   Trying to unload " & Forms(intIndex).Name & " from fnTerminateTheApp"
            End If
            ' Only attempt to unload the MDI Main form if it is now the only form left
            ' in memory.
            '
            ' Since the forms are unloaded in the reverse order from which
            ' they were loaded, the MDI Form should always be the last form unloaded.
            ' Therefore, if all other forms that were open have been unloaded, it should
            ' be okay to unload the MDI. If any of those forms were NOT unloaded (probably
            ' because the user said "No I don't want to lose my pending changes"), then
            ' don't unload the MDI form.  This conditionality is needed to avoid the user
            ' being prompted twice (or more) about losing their pending changes on the same
            ' form.
            If Forms(intIndex).Name = cstrMDIForm Then
                If Forms.Count = 1 Then
                    Unload Forms(intIndex)
                End If
            Else
                Unload Forms(intIndex)
            End If
        Next intIndex
    End If

    gbAmTryingToTerminateTheApp = False
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function fnTranslateToMaxValue(ByVal intDollarPositions As Integer, ByVal intDecimalPositions As Integer) As Double
    ' Comments  : Translates meta data (i.e. number of dollar and decimal positions) into a numeric value
    '             that represents the Maximum Value allowed.
    '             Example:  fnTranslateToMaxValue(5,4) would return 99999.9999.
    '                       fnTranslateToMaxValue(3,0) would return 999 (equivalent to 999.0)
    '
    ' Parameters: intDollarPositions  (in) - the number of dollar positions allowed
    '             intDecimalPositions (in) - the number of decimal positions allowed
    '
    ' Returns   : Double representing the maximum value
    '
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnSSNTIN_AddDash"
    Dim strMaxValue         As String
    Dim intI                As Integer
    Const cstrAnotherDigit  As String = "9"
    Const cstrZero          As String = "0"
    Const cstrDecimalPoint  As String = "."
    
    For intI = 1 To intDollarPositions
        strMaxValue = strMaxValue & cstrAnotherDigit
    Next intI
    
    strMaxValue = strMaxValue & cstrDecimalPoint
    
    If intDecimalPositions = 0 Then
        strMaxValue = strMaxValue + cstrZero
    Else
        For intI = 1 To intDecimalPositions
            strMaxValue = strMaxValue + cstrAnotherDigit
        Next intI
    End If
    
    fnTranslateToMaxValue = CDbl(strMaxValue)
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
Public Sub fnUnloadSplash()
    '----------------------------------------------------------------------------
    ' Procedure :  Sub fnUnloadSplash
    '
    ' Comments  :  Get rid of the splash screen, in case the error occurs while
    '              the splash screen is still being displayed. Otherwise,
    '              the splash screen can obscure any message box that is displayed.
    ' Called by :  modStartup's sub Main( )
    ' Parameters:  N/A
    '
    ' Modified  :
    '----------------------------------------------------------------------------
    'On Error GoTo 0
    Const cstrCurrentProc As String = "fnUnloadSplash"

    If fnIsFormLoaded("frmSplash") Then
        Unload frmSplash
        fnFreeObject frmSplash
    End If
'PROC_EXIT:
'    On Error GoTo 0     ' disable error handler
'    ' Clean-up statements go here
'    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
'        gerhApp.PropagateError mcstrName & cstrCurrentProc
'    End If
'    Exit Sub
'PROC_ERR:
'    Select Case Err.Number
'        'Case statements for expected errors go here
'        Case Else
'            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
'    End Select
'    Resume PROC_EXIT
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Sub fnWindowLock(ByVal hWnd As Long)
    '----------------------------------------------------------------------------
    ' Procedure :  Function fnWindowLock
    '
    ' Comments  : To avoid screen flicker caused by excessive repainting, use
    '             this before making a lot of screen changes and then
    '             call its companion procedure (fnWindowUnlock) afterward.
    '
    ' Called by : cmdDetailCollapse_Click( ) in the frmMsgBox form
    '             cmdDetailExpand_Click( ) in the frmMsgBox form
    ' Parameters: hWnd (in) - window handle of the form to operate upon
    '
    ' Modified  :
    '----------------------------------------------------------------------------
    ' No error handler since this could be called prior to the app startup being completed
    On Error Resume Next

    LockWindowUpdate hWnd
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Sub fnWindowUnlock()
    '----------------------------------------------------------------------------
    ' Procedure :  Function fnWindowUnlock
    '
    ' Comments  : To avoid screen flicker caused by excessive repainting, call
    '             this procedure's companion procedure (fnWindowLock)
    '             before making a lot of screen changes and then
    '             call this procedure (fnWindowUnlock) afterward.
    '
    ' Called by : cmdDetailCollapse_Click( ) in the frmMsgBox form
    '             cmdDetailExpand_Click( ) in the frmMsgBox form
    ' Parameters: N/A
    '
    ' Modified  :
    '----------------------------------------------------------------------------
    ' No error handler since this could be called prior to the app startup being completed
    On Error Resume Next

    LockWindowUpdate 0
End Sub

'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function fnTFToBool(ByVal strIn As String) As Boolean
    ' Comments  : Translates "T" to True and everything else to False
    ' Parameters: strIn (in) the string expression to translate
    '
    ' Returns   : True or False
    '
    ' Modified  : Berry Kropiwka
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnTFToBool"

    If UCase$(strIn) = "T" Then
        fnTFToBool = True
    Else
        fnTFToBool = False
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
Public Function fnYNToBool(ByVal strIn As String) As Boolean
    ' Comments  : Translates "Y" to True and everything else to False
    ' Parameters: strIn (in) the string expression to translate
    '
    ' Returns   : True or False
    '
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnYNToBool"

    If UCase$(strIn) = "Y" Then
        fnYNToBool = True
    Else
        fnYNToBool = False
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



' %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
' %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'
'   The following procedures exist only to facilitate testing. They should
'   ONLY be called from the Immediate window and not from other procedures
'   in this form or project.
'
'
'   To use these, set a breakpoint at the top of the Form_Initialize event
'   handler. Then, once you've stopped at the breakpoint, type the function
'   name in the Immediate window.
'       Correct:    TestStub_fnGetStateInfo
'                   modGeneral.TestStub_fnGetStateInfo
'
'       Incorrect:  ? TestStub1
'                   TestStub1()
'
' %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
' %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Private Sub TestStub_fnGetStateInfo()
    On Error GoTo PROC_ERR
    Const cstrCurrentProc   As String = "TestStub_fnGetStateInfo"
    Dim siTemp              As StateInfo
    
' MME START WRUS 4999 - ADDED CLAIMID AND PayeDthbPmtAmt PARAMATERS - CHECK THIS EXISTS IN DEV AND PROD..

    fnGetStateInfo "FL", "I", CDate(DateValue("1/1/1979")), 45418, 2, siTemp

    Debug.Print "State = " & siTemp.StCd
    Debug.Print "StrlEffDt = " & siTemp.StrlEffDt
    Debug.Print "CalcIdtypCd = " & siTemp.CalcIdtypCd
    Debug.Print "ReqdIdtypCd = " & siTemp.ReqdIdtypCd
    Debug.Print "IruleCd = " & siTemp.IruleCd
    Debug.Print "StrlEndDt = " & siTemp.StrlEndDt
    Debug.Print "StrlIntRptgFlrAmt = " & siTemp.StrlIntRptgFlrAmt
    Debug.Print "StrlIntCalcOfstNum = " & siTemp.StrlIntCalcOfstNum
    Debug.Print "StrlIntReqdOfstNum = " & siTemp.StrlIntReqdOfstNum
    Debug.Print "StrlIntRuleAmt = " & siTemp.StrlIntRuleAmt
    Debug.Print "StrlSpclInstrTxt = "; siTemp.StrlSpclInstrTxt
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

