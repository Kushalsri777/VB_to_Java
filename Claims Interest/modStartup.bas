Attribute VB_Name = "modStartup"
'******************************************************************************
' Module     : Startup
' Description:
' Procedures :
'              fnInitialize()
'              Main()

' Modified   :
' 10/25/01 BAW Changed logic dealing with detecting that another instance of the app
'              is already running. This in essence applies the same fix to that logic
'              as was made to SPUDS/SCUDS a couple months ago.
' 03/26/01 BAW Cleaned with Total Visual CodeTools 2000
' --------------------------------------------------
Option Explicit
Option Compare Binary
Private Const mcstrName As String = "modStartup."

Public gstrDebug As String




'////////////////////////////////////////////////////////////////////////////////////////////////
Sub Main()
    ' Comments  : This starts up the app, displaying the splash
    '             screen, instantiating global objects and then
    '             displaying the main MDI form.
    '
    '             NOTE:  This procedure (and fnDeallocateGlobalObjects in modGeneral.bas)
    '                    should be updated as global object variables are added to
    '                    or removed from the application!
    '
    ' Parameters:  -
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "Sub Main"

    ' Instantiate the global error handler
    '    NOTE: The Error Handler must be instantiated *immediately*
    '          upon app startup, since it is used EVERYWHERE !
    Set gerhApp = New cerhErrorHandler

    ' Do not allow a 2nd instance of the app to be started up. Activate
    ' the current (1st) instance instead and terminate this (the 2nd)
    ' instance.
    If App.PrevInstance Then
        AppActivate frmMDIMain.Caption, False
        End
    End If


    ' Display the splash screen, pre-load frequently used form(s),
    ' and then display the main MDI form.
    '
    ' Note using the following:
    '           Dim frm As frmSplash
    '           Set frm = New frmSplash
    '           frm.fnShowAsSplashScreen
    ' seems to make fnUnloadSplash( ) not be able to unload that form.
    ' So, be sure to use the "frmSplash.xxx" notation instead.
    frmSplash.fnShowAsSplashScreen
    DoEvents

    ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    ' Instantiate the remaining global objects used by most/all of the app
    ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

    Set gapsApp = New capsAppSettings               ' Accesses app settings stored in the registry
    Set gadwApp = New cadwADOWrapper                '!TODO! Obsolete?
    Set gconAppActive = New cconConnection          ' Handles ADO connection to the active app database
    
    ' Moved the instantiation of the Crystal object to frmSelectReports as a conditional instantiation
    ' since this CreateObject invocation is such a pig per VB Watch Profiler.
    '           Set gcrxApp = CreateObject("CrystalRuntime.Application")

    ' Initialize the AppSettings object using the initialization procedure defined
    ' in modConstructors so those properties that should be set at app startup are set.
    ' This must be done after all of the global objects since control doesn't return to here
    ' if an error occurs in this function.
    fnInit_gapsApp
    
    ' Set the log to verbose mode. In the future, this may be conditioned on
    ' a command line parameter.
    gbLogVerbose = True
    
    DoEvents

    ' Pre-load (without showing) the MDIMain and Login screens. The Splash screen will "show"
    ' the MDIMain form when it reaches 100% and unloads the splash screen.
    Load frmMDIMain
    Load frmLogOn
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.ReportFatalError App.Title   ' no screen name, so use App name
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
