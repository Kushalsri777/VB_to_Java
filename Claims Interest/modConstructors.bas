Attribute VB_Name = "modConstructors"
' =-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-=
' Class       : modConstructors
' Description : Used to instantiate and initialize the cerhErrorHandler class. This is
'               necessary so the object will be properly instantiated (i.e. as an empty
'               object) prior to filling it with default or registry-based settings,
'               since the latter can encounter errors and needs its own methods to
'               be able to report errors! Otherwise the error propogation gets
'               screwed up.
' Source      :
'
' Procedures  :
'   Public      Init_cerhErrorHandler() as cerhErrorHandler
'
' Modified:
'
'   Version Date     Who   What
'   ------- -------- ---   -------------------------------------------------------------------
'   1.0     03/07/02 BAW   (Phase2A) Created.
' =-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-=

Option Explicit
Option Compare Binary
Private Const mcstrName As String = "modConstructors."

'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'\                                                                  /
'|                        PUBLIC  Procedures                        |
'/                                                                  \
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function fnInit_gapsApp()
    ' Comments:   Initializes the capsAppSettings object. This
    '             procedure should be called immediately after
    '             instantiating an object of this class to
    '             pre-populate those settings that should be set
    '             at app startup:
    '                  Set gapsApp = New capsAppSettings
    '                  fnInit_gapsApp
    '             The Class_Initialize() method cannot do the
    '             initialization itself due to the
    '             possibility of hitting an error during
    '             the initialization. By keeping this class'
    '             Class_Initialize() to a minimum, we are assured
    '             of having a valid object before there is any
    '             possibility of hitting an error.
    ' Parameters: N/A
    ' Returns:    N/A
    ' Called by : Sub Main of modStartup.bas
    Const cstrCurrentProc As String = "fnInit_gapsApp"
    Dim strThrowaway As String

    On Error GoTo PROC_ERR

    ' Pre-populate those settings that should be retrieved at app startup
    strThrowaway = gapsApp.LastLogOnUserID
    strThrowaway = gapsApp.TaxFileFolder
    gapsApp.fnLoadEnvironments
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

