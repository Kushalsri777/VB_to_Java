VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.MDIForm frmMDIMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Claims Interest"
   ClientHeight    =   8085
   ClientLeft      =   360
   ClientTop       =   555
   ClientWidth     =   11010
   Icon            =   "frmMDIMain.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar sbrStatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   7830
      Width           =   11010
      _ExtentX        =   19420
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdlCommonDialog 
      Left            =   4000
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileLogon 
         Caption         =   "&Log On"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrintSetup 
         Caption         =   "Print Set&up..."
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewInsured 
         Caption         =   "&Insured..."
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu mnuReportsPrintReport 
         Caption         =   "&Print Report..."
      End
      Begin VB.Menu mnuReportsGenerateTaxFile 
         Caption         =   "&Generate Tax File..."
      End
   End
   Begin VB.Menu mnuWindowTop 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindow 
         Caption         =   "&Cascade"
         Index           =   0
      End
      Begin VB.Menu mnuWindowArrange 
         Caption         =   "Tile &Horizontal"
         Index           =   1
      End
      Begin VB.Menu mnuWindowArrange 
         Caption         =   "Tile &Vertical"
         Index           =   2
      End
      Begin VB.Menu mnuWindowArrange 
         Caption         =   "&Arrange Icons"
         Index           =   3
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpTechnicalSupport 
         Caption         =   "Technical &Support"
      End
      Begin VB.Menu mnuHelpViewApplicationLogFile 
         Caption         =   "&View Application Log File"
      End
      Begin VB.Menu mnuHelpAboutClaimsInterest 
         Caption         =   "&About Claims Interest..."
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmMDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
' Module     : frmMDIMain
' Description:
' Procedures:
'              fnGetNbrOfDataIntegrityIssues() As Long
'              fnOpenURLInBrowser(ByVal strURL As String) As Boolean
'              fnShowInsuredForm()
'              MDIForm_Activate()
'              MDIForm_Load()
'              MDIForm_Resize()
'              mnuFileExit_Click()
'              mnuHelpAbout_Click()
'              mnuReportsGenerateCheckFreeFile_Click()
'              mnuReportsPrintReport_Click()
'              mnuViewInsured_Click()
'              mnuWindow_Click

' Modified   :
' 10/25/01 BAW Added Help | Technical Support menu option
' 03/26/01 BAW Cleaned with Total Visual CodeTools 2000
' 01/2002  BAW Optimized per Project Analyzer (removing dead code, adding "$" to Mid/Space, etc.)
'              Also added extra DoEvents in Form_Load to make the Splash screen's progress bar
'              move more smoothly.
' --------------------------------------------------
Option Explicit
Option Compare Binary
Private mstrScreenName As String
Private Const mclngMinFormWidth  As Long = 14625
Private Const mclngMinFormHeight As Long = 10260

' The ShellExecute API is used by mnuHelpTechnicalSupport
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Function fnGetNbrOfDataIntegrityIssues() As Long
    ' Description: Executes a view to identify the number of potential
    '              data integrity issues in the database
    '              WARNING: Both frmPrintReport.fnGetData_DataIntegrityIssuesReport()
    '                       and frmMDIMain.fnGetNbrOfDataIntegrityIssues() run
    '                       this view!!!
    '
    ' Parameters: N/A
    '
    ' Called by :
    '
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnGetNbrOfDataIntegrityIssues"
    Const cstrSQLView         As String = "dbo.DataIntegrityIssuesReport_v"
    Dim strSQL                As String
    Dim rstTemp               As ADODB.Recordset

    On Error GoTo PROC_ERR
    
    ' If the user hasn't logged on yet, just return 0.
    If gconAppActive.ADOConn.State <> adStateClosed Then
        strSQL = "SELECT * from " & cstrSQLView & " WHERE calcReason <> ''"
    
        Set rstTemp = gadwApp.Execute_SQL_AsRST(gconAppActive, strSQL)
        With rstTemp
            #If DEBUG_RST Then
                Debug.Print "In " & cstrCurrentProc & ", " & CStr(.RecordCount) & " records were retrieved in the rst."
                Debug.Print "SQL statement is: " & vbCr & strSQL
            #End If
    
            ' Disconnect the recordset
            .ActiveConnection = Nothing
    
            fnGetNbrOfDataIntegrityIssues = .RecordCount
        End With
    End If
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
    
    ' Clean-up statements go here
    fnFreeRecordset rstTemp
    
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            ' Save Err object data, if not already saved
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Function fnOpenURLInBrowser(ByVal strURL As String) As Boolean
    ' Comments  : Opens the default browser on the specified URL
    ' Parameters: strURL - the URL to display in the browser window
    ' Called by : mnuHelpTechnicalSupport_Click() of frmMDIMain
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnOpenURLInBrowser"

    Dim lngReturnCode As Long

    ' Make sure the URL is prefixed with http:// or https://
    If InStr(1, UCase$(strURL), "HTTP", vbBinaryCompare) <> 1 Then
        strURL = "http://" & strURL
    End If

    lngReturnCode = ShellExecute(0&, "open", strURL, _
        vbNullString, vbNullString, vbNormalFocus)

    fnOpenURLInBrowser = (lngReturnCode > 32)
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub fnShowInsuredForm()
    ' Comments  : This is a PUBLIC procedure that opens a new instance of the Insured screen
    ' Called by : frmSplash.tmrTimer_Timer( ) and mnuViewInsured_Click( )
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnShowInsuredForm"
    Dim frm As frmInsured
    
    If Not (fnIsFormLoaded("frmInsured", frm)) Then
        Set frm = New frmInsured
    End If

    frm.Show
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0

    ' Clean-up statements go here
    fnFreeObject frm
    
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



' This routine exists only to support programmer testing of new table wrappers.
'////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub fnTestTableWrapper()
    ' Comments  : Opens the Test Table Wrapper screen
    '
    '             WARNING:  This code should be REM'd out once the table
    '                       wrappers have been fully tested!
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnTestTableWrapper"
    Dim frm               As frmTestTableWrapper

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName
    
    If Not (fnIsFormLoaded("frmTestTableWrapper", frm)) Then
        Set frm = New frmTestTableWrapper
    End If

    frm.Show
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0

    ' Clean-up statements go here
    fnFreeObject frm
    
    ' Report the error, since this is an event handler
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.ReportFatalError mstrScreenName
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            ' Save Err object data, if not already saved
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'//////////////////////////////////////////////////////////////////////////////////////////
Public Sub fnUpdateStatusBar(ByVal strUserID As String, ByVal strEnv As String)
    ' Comments  : Updates the Status Bar so it reflects the
    '             current information per the Log On screen
    '
    ' Parameters:
    '             strUserID - the User ID of the logged on user, if any (as specified on the Log On screen)
    '             strEnv - the Environment to which the user logged on, if any (as selected on the Log On screen)
    '
    ' Called by : frmMDIForm_Load()
    '             This is also called by frmLogOn's cmdOK_Click event.
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR

    Const cstrCurrentProc As String = "fnUpdateStatusBar"
    Const cstrLabelPanel1 As String = "User ID: "
    Const cstrLabelPanel2 As String = "Environment: "
    Const cstrLabelPanel3 As String = "# of Possible Errors in the DB: "

    Me.sbrStatusBar.Panels(1).Text = cstrLabelPanel1 & strUserID
    Me.sbrStatusBar.Panels(2).Text = cstrLabelPanel2 & strEnv
    
    ' If the User ID and Environment are empty, then the user must not be logged on, so don't
    ' attempt to open the view.
    If strUserID = vbNullString And strEnv = vbNullString Then
        Me.sbrStatusBar.Panels(3).Text = cstrLabelPanel3 & "?"
    Else
        Me.sbrStatusBar.Panels(3).Text = cstrLabelPanel3 & CStr(fnGetNbrOfDataIntegrityIssues())
    End If
    DoEvents
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0

    ' Clean-up statements go here

    ' Report the error, since this is an event handler
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.ReportFatalError mstrScreenName
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            ' Save Err object data, if not already saved
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub MDIForm_Activate()
    ' Comments  : Refreshes the Status Bar text to show
    '             if there are errors
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "MDIForm_Activate"
    
    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName
    
    
    With gapsApp
        frmMDIMain.fnUpdateStatusBar strUserID:=.LastLogOnUserID, _
                                     strEnv:=.LastLogonEnvironment
    End With
    DoEvents
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here

    ' Report the error, since this is an event handler
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.ReportFatalError mstrScreenName
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            ' Save Err object data, if not already saved
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub


'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub MDIForm_Load()
    ' Comments  : Opens a ADO Connection object and sets the
    '             Status Bar text to show if there are errors
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "MDIForm_Load"
    Dim pnlAdd As Panel

    ' Set the screen name that will be used to form the Title on message boxes
    mstrScreenName = Me.Caption

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName
    
    ' Open the Claimslog.log log file to track application events during this session.
    fnLogOpen
    DoEvents

    Me.Icon = LoadResPicture(gcRES_ICON_MAINAPP, vbResIcon)

    ' If the user has ever opened this form before, restore its size & placement.
    ' If the restore would result in the form being off-screen, just center it instead.
    If gapsApp.RestoreForm(Me) = False Then
        With Me
            .Width = mclngMinFormWidth
            .Height = mclngMinFormHeight
        End With
        fnCenterFormOnScreen Me
    End If
    
    ' Set properties of Status Bar, defining 3 panels:
    '   1. User ID: <acf2>
    '   2. Environment: <environment name from cconClaimsActive>
    '   3. # of Possible Errors in the DB: <##>
    '
    ' Make the status bar panels proportionate to the width of the form,
    ' with the left panel being much bigger since it contains the fully
    ' qualified path to the database.
    ' PANEL 1
    Me.sbrStatusBar.Panels(1).AutoSize = sbrContents
    Me.sbrStatusBar.Panels(1).Alignment = sbrLeft
    Me.sbrStatusBar.Panels(1).MinWidth = Me.Width * 0.33
    ' PANEL 2
    Set pnlAdd = Me.sbrStatusBar.Panels.Add(Index:=2, Style:=sbrText)
    pnlAdd.AutoSize = sbrContents
    pnlAdd.Alignment = sbrLeft
    pnlAdd.MinWidth = Me.Width * 0.33
    ' PANEL 3
    Set pnlAdd = Me.sbrStatusBar.Panels.Add(Index:=3, Style:=sbrText)
    pnlAdd.AutoSize = sbrContents
    pnlAdd.Alignment = sbrLeft
    pnlAdd.MinWidth = Me.Width * 0.34
    
    ' Initialize status bar text in all panels so it looks okay if the MDIForm
    ' is displayed while it calls the frmLogOn form.
    fnUpdateStatusBar strUserID:=vbNullString, strEnv:=vbNullString
    DoEvents
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0

    ' Clean-up statements go here
    fnFreeObject pnlAdd
    
    ' Report the error, since this is an event handler
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.ReportFatalError mstrScreenName
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            ' Save Err object data, if not already saved
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub


'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' Comments  : Inhibit closing the app if there are child forms open.
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "MDIForm_QueryUnload"

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    If bDebugAppTermination Then
        Debug.Print "Entering " & mstrScreenName & gcstrDOT & cstrCurrentProc
    End If

    If gbAmProcessingAnAppFatalError Then
        ' ALWAYS let the form be unloaded, with no prompts to the user, if shutting
        ' down the app due to an application fatal error having been hit.
        GoTo PROC_EXIT
    End If
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here

    ' Report the error, since this is an event handler
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.ReportFatalError mstrScreenName
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            ' Save Err object data, if not already saved
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub


'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub MDIForm_Resize()
    ' Comments  : Don't let the MDI form be resized such that it
    '             is too small to fit the largest form.
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "MDIForm_Resize"

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName
    
    If Me.WindowState = vbNormal Then
        ' Bypass if vbMinimized or vbMaximized, to avoid run-time error 384
        ' which says" "a form can't be moved or sized while minimized or maximized"
        If Me.Height < mclngMinFormHeight Then
            Me.Height = mclngMinFormHeight
        End If
        If Me.Width < mclngMinFormWidth Then
            Me.Width = mclngMinFormWidth
        End If
    End If
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here

    ' Report the error, since this is an event handler
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.ReportFatalError mstrScreenName
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            ' Save Err object data, if not already saved
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub


'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub MDIForm_Unload(ByRef pintCancel As Integer)
    ' Comments  : Close the log file and exit the application
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "MDIForm_Unload"
    'Dim frm As Form

    If bDebugAppTermination Then
        Debug.Print "Now in " & mstrScreenName & gcstrDOT & cstrCurrentProc
    End If

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

'!TODO! Put the following code in fnTerminateTheApp ???
    fnLogWrite "Exiting application.", cstrCurrentProc
    fnLogClose
'!TODO! End

    ' The following "If" check is needed in case of a fatal error being hit upon app start,
    ' such that the capsAppSettings object (gapsApp) didn't make it through its Class_Initialize
    ' event. Without the IF, a VB error 91 (Object variable or With block not set) is reported.
    If Not (gapsApp Is Nothing) Then
        gapsApp.SaveForm Me
    End If

    ' Can't do "Set frmMDIMain = Nothing"...causes a crash when fnTerminateTheApp tries to
    ' deallocate the global objects
    'Unload Me

    Debug.Print mstrScreenName & gcstrDOT & cstrCurrentProc & " is calling fnTerminateTheApp..."
    fnTerminateTheApp
    
        
    Debug.Print mstrScreenName & gcstrDOT & cstrCurrentProc & " is calling fnDeallocateGlobalObjects..."
    fnDeallocateGlobalObjects
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here

    ' Report the error, since this is an event handler
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.ReportFatalError mstrScreenName
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            ' Save Err object data, if not already saved
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub mnuFile_Click()
    ' Comments  :
    ' Parameters: N/A
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "mnuFile_Click"

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    With Forms
        ' Enable the File | Log On option only if the MDIMain form is the sole open form
        mnuFileLogon.Enabled = (.Count = 1)

        mnuFileExit.Enabled = True
    End With
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0

    ' Clean-up statements go here

    ' Report the error, since this is an event handler
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.ReportFatalError mstrScreenName
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            ' Save Err object data, if not already saved
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub mnuFileExit_Click()
    ' Comments  : Terminates the app. Note that this menu option should be disabled
    '             if any forms besides the MDI form are open.
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "mnuFileExit_Click"
    Dim frm As Form

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName
    
    If bDebugAppTermination Then
        Debug.Print mstrScreenName & gcstrDOT & cstrCurrentProc & " is calling fnTerminateTheApp"
    End If
    fnTerminateTheApp   '   !!!!!!!!!!!!!!!!!!!!!!!!
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0

    ' Clean-up statements go here
    fnFreeObject frm
    
    ' Report the error, since this is an event handler
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.ReportFatalError mstrScreenName
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            ' Save Err object data, if not already saved
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub mnuFileLogOn_Click()
    ' Comments  : Note that this menu option should not be available unless
    '             all forms besides the MDI form are closed.
    ' Parameters: N/A
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "mnuFileLogOn_Click"
    Dim frm As frmLogOn

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    ' If the user has an ADO connection currently open, then
    ' terminate the connection before proceeding. This ensures the connection
    ' will always represent the logged-on user.
    gconAppActive.Disconnect

    If Not (fnIsFormLoaded("frmLogOn", frm)) Then
        Set frm = New frmLogOn
    End If

    frm.Show vbModal
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0

    ' Clean-up statements go here
    fnFreeObject frm

    ' Report the error, since this is an event handler
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.ReportFatalError mstrScreenName
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            ' Save Err object data, if not already saved
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub


'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub mnuFilePrintSetup_Click()
    ' Comments  :
    ' Parameters: N/A
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "mnuFilePrintSetup_Click"

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    With cdlCommonDialog
        .PrinterDefault = True
        .Flags = cdlPDPrintSetup
        .ShowPrinter
    End With
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0

    ' Clean-up statements go here

    ' Report the error, since this is an event handler
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.ReportFatalError mstrScreenName
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            ' Save Err object data, if not already saved
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub mnuHelpAboutClaimsInterest_Click()
    ' Comments  : Displays the splash screen as a Help | About box
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "mnuHelpAboutClaimsInterest_Click"
    Dim frm               As frmSplash

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

'!TODO! frm is not initialized if Else condition is hit!
    If Not (fnIsFormLoaded("frmSplash", frm)) Then
        Set frm = New frmSplash
        frm.fnShowAsAboutBox
    Else
        frm.fnShowAsAboutBox
    End If
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    
    ' Clean-up statements go here
    fnFreeObject frm

    ' Report the error, since this is an event handler
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.ReportFatalError mstrScreenName
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            ' Save Err object data, if not already saved
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub


'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub mnuHelpTechnicalSupport_Click()
    ' Comments  : Display the Claims Interest Technical Support page on The Source
    ' Parameters:  -
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "mnuHelpTechnicalSupport_Click"
    Const cstrURL As String = "http://intranet/showcontext.cfm?context=1826"

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    If Not (fnOpenURLInBrowser(cstrURL)) Then
        gerhApp.ReportNonFatal vbObjectError + gcRES_INFO_CANT_LAUNCH_URL, _
                               mstrScreenName & gcstrDOT & cstrCurrentProc, _
                               cstrURL
    End If
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here

    ' Report the error, since this is an event handler
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.ReportFatalError mstrScreenName
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            ' Save Err object data, if not already saved
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub


' The creation of maintenance screens for the Current Rate and State Rule tables
' is deferred in the v2.4 release. As such, the Tools menu bites the dust...for now.
#If False Then
'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub mnuTools_Click()
    ' Comments  : Opens the Current Rate screen as a modal window
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "mnuTools_Click"

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    If gconAppActive.LastLogonIsSpecialUser Then
        mnuToolsCurrentRate.Enabled = True
        mnuToolsStateRule.Enabled = True
    Else
        mnuToolsCurrentRate.Enabled = False
        mnuToolsStateRule.Enabled = False
    End If
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here

    ' Report the error, since this is an event handler
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.ReportFatalError mstrScreenName
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            ' Save Err object data, if not already saved
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub


'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub mnuToolsCurrentRate_Click()
    ' Comments  : Opens the Current Rate screen as a modal window
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "mnuToolsCurrentRate_Click"
    Dim intResponse       As Integer
    Dim frm               As frmCurrentRate

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName
    
    If Not (fnIsFormLoaded("frmCurrentRate", frm)) Then
        Set frm = New frmCurrentRate
    End If

    frm.Show vbModal
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here

    ' Report the error, since this is an event handler
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.ReportFatalError mstrScreenName
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            ' Save Err object data, if not already saved
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub


'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub mnuToolsStateRule_Click()
    ' Comments  : Opens the State Rules screen as a modal window
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "mnuToolsStateRules_Click"
    Dim intResponse       As Integer
    Dim frm               As frmStateRule

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName
    
    If Not (fnIsFormLoaded("frmStateRule", frm)) Then
        Set frm = New frmStateRule
    End If

    frm.Show vbModal
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here

    ' Report the error, since this is an event handler
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.ReportFatalError mstrScreenName
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            ' Save Err object data, if not already saved
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub
#End If



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub mnuHelpViewApplicationLogFile_Click()
    ' Comments  : Show the splash screen as an About box.
    ' Parameters: N/A
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc   As String = "mnuHelpViewApplicationLogFile_Click"
    Dim strLogFileNm        As String
    Dim strLogFileExt       As String

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    strLogFileNm = fnGetAppLogFileFQ()
    strLogFileExt = fnGetExtPart(strLogFileNm)
    
    If Not fnOpenFileInDefaultApp(strLogFileNm) Then
        ' gcRES_INFO_CANT_OPEN_FILE (1014)
        ' Unable to open @@1. The file either does not exist or no application is associated with files of type @@2.
        gerhApp.ReportNonFatal vbObjectError + gcRES_INFO_CANT_OPEN_FILE, _
                               mstrScreenName & gcstrDOT & cstrCurrentProc, _
                               "the application log file (" & strLogFileNm & ")", UCase$(strLogFileExt)
    End If
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0

    ' Clean-up statements go here

    ' Report the error, since this is an event handler
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.ReportFatalError mstrScreenName
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            ' Save Err object data, if not already saved
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    
    Resume PROC_EXIT
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub mnuReportsGenerateTaxFile_Click()
    ' Comments  : Opens the Generate Tax File screen
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "mnuReportsGenerateTaxFile_Click"
    Dim frm               As frmGenerateTaxFile

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName
    
    If Not (fnIsFormLoaded("frmGenerateTaxFile", frm)) Then
        Set frm = New frmGenerateTaxFile
    End If

    frm.Show
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0

    ' Clean-up statements go here
    fnFreeObject frm
    
    ' Report the error, since this is an event handler
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.ReportFatalError mstrScreenName
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            ' Save Err object data, if not already saved
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub


'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub mnuReportsPrintReport_Click()
    ' Comments  : Opens the Print Report screen
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "mnuReportsPrintReport_Click"
    Dim frm               As frmPrintReport

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName
    
    If Not (fnIsFormLoaded("frmPrintReport", frm)) Then
        Set frm = New frmPrintReport
    End If

    frm.Show
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0

    ' Clean-up statements go here
    fnFreeObject frm
    
    ' Report the error, since this is an event handler
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.ReportFatalError mstrScreenName
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            ' Save Err object data, if not already saved
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub


'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub mnuViewInsured_Click()
    ' Comments  : Opens a new instance of the Insured screen
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "mnuViewInsured_Click"

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    fnShowInsuredForm
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here

    ' Report the error, since this is an event handler
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.ReportFatalError mstrScreenName
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            ' Save Err object data, if not already saved
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
'   Don't need a  mnuWindow_Click() event handler; all menu options are enabled as of 2A
'////////////////////////////////////////////////////////////////////////////////////////////////



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub mnuWindowArrange_Click(ByRef pintIndex As Integer)
    ' Comments  : The "arrangement" items in the Window menu
    '             are a control array, all controlled by
    '             this control event.
    ' Parameters: pintIndex (in) - indicates which Window menu option was selected
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "mnuWindowArrange_Click"

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    Select Case pintIndex
        Case 0
            Me.Arrange vbCascade
        Case 1
            Me.Arrange vbTileHorizontal
        Case 2
            Me.Arrange vbTileVertical
        Case 3
            Me.Arrange vbArrangeIcons
    End Select
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here

    ' Report the error, since this is an event handler
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.ReportFatalError mstrScreenName
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            ' Save Err object data, if not already saved
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub
