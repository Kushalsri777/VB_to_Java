VERSION 5.00
Begin VB.Form frmLogOn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Log On"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4875
   Icon            =   "frmLogon.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Log On"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   4875
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1350
      PasswordChar    =   "*"
      TabIndex        =   5
      ToolTipText     =   "Type your password (Note: Each character typed will appear as an asterisk.)"
      Top             =   945
      Width           =   3255
   End
   Begin VB.TextBox txtUserId 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1350
      TabIndex        =   3
      ToolTipText     =   "Type your User ID"
      Top             =   570
      Width           =   3255
   End
   Begin VB.ComboBox cboEnvironment 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmLogon.frx":000C
      Left            =   1350
      List            =   "frmLogon.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Select the Environment to which to log on"
      Top             =   180
      Width           =   3255
   End
   Begin VB.CommandButton cmdExitApplication 
      Cancel          =   -1  'True
      Caption         =   "E&xit Application"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2500
      TabIndex        =   7
      ToolTipText     =   "Do not log on to SQL Server. (Note: This will terminate the application.)"
      Top             =   1650
      Width           =   1350
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1026
      TabIndex        =   6
      ToolTipText     =   "Proceed with logging on"
      Top             =   1650
      Width           =   1350
   End
   Begin VB.Label lblEnvironment 
      AutoSize        =   -1  'True
      Caption         =   "*&Environment:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   238
      TabIndex        =   0
      Top             =   255
      Width           =   1050
   End
   Begin VB.Label lblUserId 
      AutoSize        =   -1  'True
      Caption         =   "*&User ID:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   238
      TabIndex        =   2
      Top             =   645
      Width           =   690
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      Caption         =   "*Password:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   1020
      Width           =   840
   End
End
Attribute VB_Name = "frmLogOn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
' Module     : frmLogOn
' Description:
' Procedures :
'    Private   cmdExitApplication_Click()
'    Private   cmdOK_Click()
'    Private   fnValidData() As Boolean
'    Private   fnWarningData()
'    Private   Form_Activate()
'    Private   Form_KeyDown(ByRef pintKeyCode As Integer, ByRef pintShift As Integer)
'    Private   Form_Load()
'    Private   Form_Unload(ByRef pintCancel As Integer)
'    Private   txtUserId_LostFocus()
'
' Modified   :
' 03/03/02 BAW (Phase2A) Added support for new global error handler
' 08/31/01 BAW (Phase2A) Added standardized error handlers
' 09/25/00 JG  (Phase2A) Cleaned with Total Visual CodeTools 2000
'
' --------------------------------------------------
Option Explicit
Option Compare Binary
Private mstrScreenName                   As String
Private Const mclngMinFormWidth          As Long = 4965
Private Const mclngMinFormHeight         As Long = 2715

Private Const mcstrTxtUserIdLabel        As String = "User ID"
Private Const mcstrTxtPasswordLabel      As String = "Password"
Private Const mcstrCboEnvironmentLabel   As String = "Environment"

' The following is used to determine whether the user has actually changed
' the User ID TextBox. If not, then the LostFocus event should not revalidate.
Private mstrSaveUserID              As String
Private mstrNetworkUserID           As String   'SQL_INTEGRATED_SECURITY

Dim m_autAuthenticate           As cautAuthenticate


'SQL_INTEGRATED_SECURITY - Added
'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub cboEnvironment_Click()
    ' Comments  : Make sure to default User ID to logged on network User ID if the
    '             selected Environment uses Integrated Security (i.e. Windows Authentication)
    ' Parameters: N/A
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "cboEnvironment_Click"
    
    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName
    
    If gapsApp.UsesWindowsAuthentication(cboEnvironment.Text) Then
        txtUserId.Text = fnGetNetworkUser()
        txtPassword.Text = vbNullString
        fnEnableDisableControl txtUserId, False
        fnEnableDisableControl txtPassword, False
    Else
        txtPassword.Text = vbNullString
        fnEnableDisableControl txtUserId, True
        fnEnableDisableControl txtPassword, True
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
'SQL_INTEGRATED_SECURITY - Added



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub cmdExitApplication_Click()
    ' Comments  : Exit Application button
    ' Parameters: N/A
    ' Modified  :
    '   BAW 09/10/2001 - Changed "End" to call a procedure terminates the app by unloading all forms
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "cmdExitApplication_Click"

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    If bDebugAppTermination Then
        Debug.Print "frmLogOn.cmdExitApplication_Click is calling fnTerminateTheApp..."
    End If
    fnTerminateTheApp
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
Private Sub cmdOK_Click()
    ' Comments  : OK button
    ' Parameters: N/A
    ' Modified  :
    '   09/10/2001 - Changed the behavior following an unsuccessful logon
    '                to call the fatal error handler, rather than just "End"
    ' --------------------------------------------------
    Const cstrCurrentProc As String = "cmdOK_Click"
    Dim hrgHourglass      As chrgHourglass
    On Error GoTo PROC_ERR
    
    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    ' Change mouse pointer into an hourglass
    Set hrgHourglass = New chrgHourglass
    hrgHourglass.value = True

    ' Close the ADO Connection since the user is probably logging on to a different environment,
    ' or under a different User ID. Update the status bar accordingly.
    frmMDIMain.fnUpdateStatusBar strUserID:=vbNullString, strEnv:=vbNullString
    gconAppActive.Disconnect

    If fnValidData() Then
        ' Connect to the selected Environment and put the desired App Role into effect.
        ' .Connect( ) raises an error if it is unsuccessful, so if we get to the
        ' subsequent statement it means "success".
        m_autAuthenticate.AuthenticateUser strEnvironIn:=cboEnvironment.Text, _
                                           strUserIDIn:=txtUserId.Text, _
                                           strPasswordIn:=txtPassword.Text, _
                                           pconIn:=gconAppActive, _
                                           bActiveDBIn:=True
        
        'Debug.Print "This user was authenticated: " & txtUserId.Text & _
        '            " (password=" & txtPassword.Text & _
        '            " Environment=" & cboEnvironment.Text
        
        With gapsApp
            ' Save logged-on User ID, so it can be shown the next time the user comes to this screen
            .LastLogOnUserID = txtUserId.Text
            ' Save logged-on User's Password, so it can reused -- but only within this session.
            .LastLogonPassword = txtPassword.Text
            ' Save Environment name
            .LastLogonEnvironment = cboEnvironment.Text
        
            frmMDIMain.fnUpdateStatusBar strUserID:=.LastLogOnUserID, _
                                         strEnv:=.LastLogonEnvironment
        End With
        Unload Me
    End If
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0

    ' Report the error, since this is an event handler
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.ReportFatalError mstrScreenName
    End If

    ' Clean-up statements go here
    hrgHourglass.value = False
    fnFreeObject hrgHourglass

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



' SQL_INTEGRATED_SECURITY - Added
'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub fnEnableDisableControl(ByVal ctlIn As Control, Optional bEnable As Boolean = True)
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
                ' Next 3 lines commented out since this app doesn't use the VSFlexGrid control
                ' If (TypeOf ctlIn Is VSFlexGrid) Then
                '     ' Do nothing
                ' ElseIf (TypeOf ctlIn Is DTPicker) Then
                If (TypeOf ctlIn Is DTPicker) Then
                    .TabStop = True
                    .Enabled = True
                Else
                    .Locked = False
                    .TabStop = True
                    .BackColor = vbWindowBackground
                    .ForeColor = vbWindowText
                    .Enabled = True
                End If
            Case False
                ' Next 3 lines commented out since this app doesn't use the VSFlexGrid control
                ' If (TypeOf ctlIn Is VSFlexGrid) Then
                '     ' Do nothing
                ' ElseIf (TypeOf ctlIn Is DTPicker) Then
                If (TypeOf ctlIn Is DTPicker) Then
                    .TabStop = False
                    .Enabled = False
                Else
                    .Locked = True
                    .TabStop = False
                    .BackColor = vbButtonFace
                    .ForeColor = vbButtonText
                    .Enabled = False
                End If
        End Select
    End With
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0

    ' Clean-up statements go here

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
' SQL_INTEGRATED_SECURITY - Added



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Function fnValidData() As Boolean
    ' Comments  : Determines if all data is valid, including
    '             whether all required fields have been input.
    '             If a data error is found, it returns False
    '             which directs the caller to stop processing.
    '             It also generates warnings, by calling
    '             WarningData(), but only if no errors were
    '             found up to that point.
    ' Parameters: N/A
    '
    ' Called By : cmdOK_Click() in frmLogOn
    '
    ' Returns   : True if all data is valid; False otherwise
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnValidData"

    Dim bErrorFound As Boolean
    Dim ctlFirstToFail As Control
    Dim intFailures As Integer
    Dim strFieldList As String

    ' Check the fields in a left-to-right, top-to-bottom screen sequence.
    '     1. User ID
    '     2. Password
    '     3. Environment
    '
    ' ------------- 1.  Verify required fields are missing --------------

    ' Verify fields are necessary to connect to requested Environment

    'SQL_INTEGRATED_SECURITY
    ' Only validate User ID and Password if Windows Authentication is NOT used.
    ' (If it IS used, we get it from how the user is logged on to the network;
    ' disregarding and not allowing them to input the User ID and Password
    ' on this screen.
    If Not gapsApp.UsesWindowsAuthentication(cboEnvironment) Then
        If IsNull(txtUserId.Text) Or Len(txtUserId.Text) = 0 Then
            If intFailures = 0 Then
                strFieldList = vbCrLf & mcstrTxtUserIdLabel
                Set ctlFirstToFail = txtUserId
            Else
                strFieldList = strFieldList & vbCrLf & mcstrTxtUserIdLabel
            End If
            intFailures = intFailures + 1
        End If

    'SQL_INTEGRATED_SECURITY ' Need to do Environment next, even though it's not the next consecutive field
    'SQL_INTEGRATED_SECURITY ' on the screen, since its value determines whether Password will be checked
    'SQL_INTEGRATED_SECURITY If IsNull(cboEnvironment.Text) Or Len(cboEnvironment.Text) = 0 Then
    'SQL_INTEGRATED_SECURITY     If intFailures = 0 Then
    'SQL_INTEGRATED_SECURITY         strFieldList = vbCrLf & mcstrCboEnvironmentLabel
    'SQL_INTEGRATED_SECURITY         Set ctlFirstToFail = cboEnvironment
    'SQL_INTEGRATED_SECURITY     Else
    'SQL_INTEGRATED_SECURITY         strFieldList = strFieldList & vbCrLf & mcstrCboEnvironmentLabel
    'SQL_INTEGRATED_SECURITY     End If
    'SQL_INTEGRATED_SECURITY     intFailures = intFailures + 1
    'SQL_INTEGRATED_SECURITY End If

        ' Password is required in the SQL Server environment
        If IsNull(txtPassword.Text) Or Len(txtPassword.Text) = 0 Then
            If intFailures = 0 Then
                strFieldList = vbCrLf & mcstrTxtPasswordLabel
                Set ctlFirstToFail = txtPassword
            Else
                strFieldList = strFieldList & vbCrLf & mcstrTxtPasswordLabel
            End If
            intFailures = intFailures + 1
        End If
    End If
    'SQL_INTEGRATED_SECURITY

    If intFailures <> 0 Then
        bErrorFound = True
        fnValidData = False
        If ctlFirstToFail.Visible Then
            ctlFirstToFail.SetFocus
        End If
        gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_REQD_FIELDS_MISSING, _
                                   mstrScreenName & gcstrDOT & cstrCurrentProc, _
                                   strFieldList
        GoTo PROC_EXIT
    End If

    ' If no errors found, continue with checking for warnings
    ' NOTE: We won't get here if any errors were raised by preceding lines in Section 3.
    If Not bErrorFound Then
        fnValidData = True
        fnWarningData
    End If
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0

    ' Clean-up statements go here
    fnFreeObject ctlFirstToFail

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
Private Sub fnWarningData()
    ' Comments  : Validates fields, generating warnings if appropriate.
    '             It should NOT cause ValidData (this procedure's caller)
    '             to return False, since we want updates to proceed.
    ' Parameters: N/A
    '
    ' Called By : fnValidData() in frmMain
    '
    ' Returns   : N/A
    ' Modified  :
    ' --------------------------------------------------
    Const cstrCurrentProc As String = "fnWarningData"
    On Error GoTo PROC_EXIT

    ' ***   Currently there are no warnings  :(   ***

PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here
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



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
    ' Comments  :
    ' Parameters: N/A
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "Form_Activate"

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    ' SQL_INTEGRATED_SECURITY
    If cboEnvironment.Visible Then
        cboEnvironment.SetFocus
    End If
    ' SQL_INTEGRATED_SECURITY
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
Private Sub Form_KeyDown(ByRef pintKeyCode As Integer, ByRef pintShift As Integer)
    ' Comments  :
    ' Parameters: pintKeyCode
    '             pintShift -
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "Form_KeyDown"

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    If pintKeyCode = vbKeyEscape Then
        If bDebugAppTermination Then
            Debug.Print mstrScreenName & gcstrDOT & cstrCurrentProc & "  is calling fnTerminateTheApp..."
        End If
        fnTerminateTheApp
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
Private Sub Form_Load()
    ' Comments  : For Phase2A, we can populate the Environment combo box based on the
    '             gapsApp.LoadCbo_EnvironmentNames( ) method, and then filter out
    '             Dev environments for unauthorized users (per a hard-coded list of User IDs).
    '
    '             For Phase2C, it will be based on what the Authenticate object
    '             determines to be valid environments for the user.
    ' Parameters: N/A
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "Form_Load"

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    mstrScreenName = Me.Caption
    gerhApp.ScreenName = mstrScreenName

    ' Disable means of closing this form *other than* the OK or Exit Application buttons...such as
    ' the Close button in the upper righthand corner of the screen and Alt-F4.
    fnRemoveCloseButton Me

    Me.Icon = LoadResPicture(gcRES_ICON_MAINAPP, vbResIcon)

    ' If the user has ever opened this form before, restore its size & placement.
    ' If the restore would result in the form being off-screen, just center it instead.
    If gapsApp.RestoreForm(Me) = False Then
        With Me
            .Width = mclngMinFormWidth
            .Height = mclngMinFormHeight
        End With
        fnCenterFormOnScreen Me ' Not an MDI child, so it must be centered on the screen (not MDI parent)
    End If

    Set m_autAuthenticate = New cautAuthenticate

    'SQL_INTEGRATED_SECURITY ' Initialize saved User ID to blank so initial txtUserId_LostFocus trigger will
    'SQL_INTEGRATED_SECURITY ' validate environments
    mstrNetworkUserID = fnGetNetworkUser()
    
    'SQL_INTEGRATED_SECURITY mstrSaveUserID = gcstrBlankEntry
    mstrSaveUserID = mstrNetworkUserID

    ' Initialize the User ID to the User ID under which the user logged on to the network.
    txtUserId.Text = mstrNetworkUserID 'SQL_INTEGRATED_SECURITY

    ' Load Environment combobox with list of all available SQL environments, regardless
    ' of whether user is authorized for them.
    gapsApp.LoadCbo_EnvironmentNames cboEnvironment
    'SQL_INTEGRATED_SECURITY ' The Environments combo box will appear with an empty entry until the user types
    'SQL_INTEGRATED_SECURITY ' in a User ID.
    'SQL_INTEGRATED_SECURITY cboEnvironment.Clear
    'SQL_INTEGRATED_SECURITY cboEnvironment.AddItem gcstrBlankEntry
    cboEnvironment.ListIndex = 0            ' Select 1st entry as default selection
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
Private Sub Form_Unload(ByRef pintCancel As Integer)
    ' Comments  :
    ' Parameters: N/A
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "Form_Unload"

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    If bDebugAppTermination Then
        Debug.Print "Entering " & mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    
    gapsApp.SaveForm Me
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    
    ' Clean-up statements go here
    fnFreeObject frmLogOn
    fnFreeObject m_autAuthenticate

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
Private Sub txtUserId_LostFocus()
    ' Comments  : Whenever the user leaves this field, see if the Environment combobox
    '             needs to be updated. In Phase2C, this will be done via a call to
    '             the Authenticate object.
    ' Parameters: N/A
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "txtUserId_LostFocus"
    Dim intIndex          As Integer
    'SQL_INTEGRATED_SECURITY Dim aEnvs()           As String

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    'SQL_INTEGRATED_SECURITY If txtUserId.Text <> mstrSaveUserID Then
    'SQL_INTEGRATED_SECURITY    ' Repopulate the Environment combo box
    'SQL_INTEGRATED_SECURITY    cboEnvironment.Clear
    'SQL_INTEGRATED_SECURITY    aEnvs = m_autAuthenticate.AuthenticateEnvironments(txtUserId.Text)
    'SQL_INTEGRATED_SECURITY    For intIndex = LBound(aEnvs) To UBound(aEnvs)
    'SQL_INTEGRATED_SECURITY        If Len(Trim$(aEnvs(intIndex))) <> 0 Then
    'SQL_INTEGRATED_SECURITY            'Debug.Print "This user was authenticated: " & txtUserId.Text & _
    'SQL_INTEGRATED_SECURITY            '            " Environment=" & aEnvs(intIndex)
    'SQL_INTEGRATED_SECURITY            cboEnvironment.AddItem aEnvs(intIndex)
    'SQL_INTEGRATED_SECURITY        End If
    'SQL_INTEGRATED_SECURITY    Next
    
    'SQL_INTEGRATED_SECURITY    ' If there are no environments for which the specified User ID is authorized,
    'SQL_INTEGRATED_SECURITY    ' then disable the OK button so the user can only click the Exit Application
    'SQL_INTEGRATED_SECURITY    ' button unless they specify a different User ID.
        If cboEnvironment.ListCount < 1 Then
            'SQL_INTEGRATED_SECURITY gerhApp.ReportNonFatal vbObjectError + gcRES_INFO_NO_AUTHENTICATED_ENVIRONMENTS, _
            'SQL_INTEGRATED_SECURITY                        mstrScreenName & gcstrDOT & cstrCurrentProc, _
            'SQL_INTEGRATED_SECURITY                        mcstrTxtUserIdLabel, mcstrCboEnvironmentLabel
            cmdOK.Enabled = False
            cmdExitApplication.SetFocus
        Else
            'SQL_INTEGRATED_SECURITY cboEnvironment.ListIndex = 0            ' Select 1st entry as default selection
            cmdOK.Enabled = True
        End If
        
        mstrSaveUserID = txtUserId.Text
    'SQL_INTEGRATED_SECURITY End If
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
