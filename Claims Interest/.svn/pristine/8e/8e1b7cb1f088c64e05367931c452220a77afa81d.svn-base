VERSION 5.00
Object = "{B0A5E263-9338-11D4-ABD9-004F4904FC81}#2.0#0"; "CpbOcx.ocx"
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "frmSplash"
   ClientHeight    =   4755
   ClientLeft      =   2520
   ClientTop       =   2400
   ClientWidth     =   7200
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
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
      Height          =   495
      Left            =   3112
      TabIndex        =   0
      Top             =   3900
      Visible         =   0   'False
      Width           =   975
   End
   Begin pCoolProjectBar.CoolProgressBar cpbProgressBar 
      Height          =   195
      Left            =   503
      TabIndex        =   6
      Top             =   4020
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   344
      Value           =   0
   End
   Begin VB.TextBox txtProgressBorder 
      BackColor       =   &H00600000&
      Height          =   315
      Left            =   443
      TabIndex        =   4
      Top             =   3960
      Width           =   6315
   End
   Begin VB.Timer tmrTimer 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   480
      Top             =   1740
   End
   Begin VB.Label lblApplicationDeveloper 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Application Developer: Sun Life Financial, Individual Systems"
      BeginProperty Font 
         Name            =   "GiovanniEFBook"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   270
      Left            =   420
      TabIndex        =   5
      Top             =   3000
      Width           =   6375
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblApplicationName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Claims Interest"
      BeginProperty Font 
         Name            =   "GiovanniEFBold"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   495
      Left            =   555
      TabIndex        =   3
      Top             =   2460
      Width           =   6090
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "GiovanniEFBook"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   285
      Left            =   3180
      TabIndex        =   2
      Top             =   3240
      Width           =   840
   End
   Begin VB.Label lblProgress 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please wait while the application loads..."
      BeginProperty Font 
         Name            =   "GiovanniEFBook"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   225
      Left            =   450
      TabIndex        =   1
      Top             =   3660
      Width           =   3450
   End
   Begin VB.Image imgSplashBackground 
      Height          =   4755
      Left            =   0
      Picture         =   "frmSplash.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7200
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
' Module     : frmSplash
' Description:
' Procedures :
'   Private    cmdOK_Click()
'   Private    fnResetColors()
'   Public     fnShowAsAboutBox()
'   Public     fnShowAsSplashScreen()
'   Private    Form_KeyPress(ByRef pintKeyAscii As Integer)
'   Private    Form_Load()
'   Private    Form_Unload(ByRef pintCancel As Integer)
'   Private    tmrTimer_Timer()

'
' Modified   :
' 03/03/02 BAW (Phase2A) Added support for new global error handler
' 08/31/01 BAW (Phase2A) Added standardized error handlers
' 09/25/00 JG  (Phase2A) Cleaned with Total Visual CodeTools 2000
'
' --------------------------------------------------
Option Explicit
Option Compare Binary
Private mstrScreenName As String

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const mcSWP_NOMOVE = &H2
Private Const mcSWP_NOREDRAW = &H8
Private Const mcSWP_NOSIZE = &H1
Private Const mcHWND_TOPMOST = -1
Private Const mcHWND_NOTOPMOST = -2

'Define the colors for the statusbar
Private Const mcSLFBlue = 4865792
Private Const mcSLFOrange = 1026784
Private Const mcSLFWhite = 16777215
Private Const mcSLFBlack = 0

Private msng_EstProgress As Single



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub cmdOK_Click()
    ' Comments  : When shown as a Help | About screen, the OK button is visible
    '             and unloads this screen when clicked.
    ' Parameters:  -
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "cmdOK_Click"

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    Unload Me
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
Private Sub fnResetColors()
    ' Comments  : Adjusts the colors of the progress meter per the
    '             user's current screen resolution
    ' Parameters:  -
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnResetColors"

    Dim bm As BITMAP

    AutoRedraw = True
    GetObject Image, Len(bm), bm

    If bm.bmBitsPixel = 8 Then
        cpbProgressBar.Color1 = mcSLFWhite
        cpbProgressBar.Color2 = mcSLFWhite
        ' Following 3 lines added by BAW to ensure text is visible
        ' on top of mcSLFBlue background
        lblApplicationName.ForeColor = mcSLFWhite
        lblApplicationDeveloper.ForeColor = mcSLFWhite
        lblVersion.ForeColor = mcSLFWhite
    Else
        cpbProgressBar.Color1 = mcSLFWhite
        cpbProgressBar.Color2 = mcSLFBlue
        ' Following 3 lines added by BAW to ensure text is visible
        ' on top of mcSLFBlue background
        lblApplicationName.ForeColor = mcSLFWhite
        lblApplicationDeveloper.ForeColor = mcSLFWhite
        lblVersion.ForeColor = mcSLFWhite
    End If

    cpbProgressBar.BackColor = mcSLFBlack
    txtProgressBorder.BackColor = mcSLFBlack
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
            ' Save Err object data, if not already saved
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub fnShowAsAboutBox()
    ' Comments  : This function is called from the Help | About menu choice
    '             and displays this form with an OK button.
    ' Parameters:  -
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnShowAsAboutBox"

    ' The following line triggers the Form_Load event if the form has not yet
    ' been loaded
    lblProgress.Visible = False
    cpbProgressBar.Visible = False
    txtProgressBorder.Visible = False

    cmdOK.Visible = True
    Show vbModal
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
            ' Save Err object data, if not already saved
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub fnShowAsSplashScreen()
    ' Comments  : Called by sub Main, this function displays this
    '             form as a splash screen
    ' Parameters:  -
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnShowAsSplashScreen"

    cpbProgressBar.Width = 6195
    cpbProgressBar.value = 0
    msng_EstProgress = 0
    tmrTimer.Enabled = True
    ' CMP set the interval to be the number of seconds to display the splashscreen * 100
    ' 10000 milliseconds = 10 seconds
    tmrTimer.Interval = 55    ' was 150
    Me.Show vbModeless

    ' Put this form "on top"
    SetWindowPos Me.hWnd, mcHWND_TOPMOST, 0, 0, 0, 0, mcSWP_NOMOVE Or mcSWP_NOSIZE
    ' Now let it float behind other windows (like our own app's message boxes!) if warranted.
    SetWindowPos Me.hWnd, mcHWND_NOTOPMOST, 0, 0, 0, 0, mcSWP_NOMOVE Or mcSWP_NOSIZE
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
            ' Save Err object data, if not already saved
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub Form_KeyPress(ByRef pintKeyAscii As Integer)
    ' Comments  : If the user presses any key, unload this form.
    ' Parameters: pintKeyAscii -
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "Form_KeyPress"

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    Unload Me
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
    ' Comments  : Displays the form with the current version information. The
    '             colors of the progress meter reflect the user's selected
    '             screen resolution.
    ' Parameters:  -
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "Form_Load"

    mstrScreenName = Me.Caption

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    Me.Icon = LoadResPicture(gcRES_ICON_MAINAPP, vbResIcon)

    ' Don't restore previously displayed form size & position

    ' Adjust the colors per the user's current screen resolution
    fnResetColors

    lblApplicationName = App.ProductName
    lblVersion.Caption = "Application Version " & App.Major & "." & App.Minor & "." & App.Revision

    cpbProgressBar.Width = 6195
    cpbProgressBar.value = 0
    cpbProgressBar.Refresh
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

    ' Don't save current form size & position
    fnFreeObject frmSplash
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
Private Sub tmrTimer_Timer()
    ' Comments  : This event is triggered each time the Timer's interval
    '             has elapsed. When it detects the progress meter has
    '             gotten to a certain point (currently >= 98%), it
    '             unloads the splash screen.
    ' Parameters:  -
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "tmrTimer_Timer"

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    msng_EstProgress = msng_EstProgress + 5

    'cpbProgressBar.Width = 6195 * (msng_EstProgress / 100)
    cpbProgressBar.value = msng_EstProgress
    cpbProgressBar.Refresh
    If cpbProgressBar.Visible Then
        cpbProgressBar.SetFocus
    End If
    DoEvents

    If msng_EstProgress >= 98 Then
        Me.Hide
        frmMDIMain.Show
        ' Display the Log On screen, which forces the user to either log on to the
        ' application or, alternatively, EXIT the application via a call to fnTerminateTheApp().
        frmLogOn.Show vbModal, frmMDIMain
        
        ' Uncomment out the next line if you want to test a table wrapper.
        'frmMDIMain.fnTestTableWrapper
        
        ' Don't try showing the Insured form since it's hard to tell that the user successfully
        ' logged on vs. clicked the Exit Application button (unless yet another global boolean
        ' is set.)  Per Michelle Wilkosky, it's fine to just skip automatically displaying
        ' the Insured screen. Hence, the following lines were commented out.
        '       ' If we get here, the user successfully logged on, so now display
        '       ' the Insured screen.
        '       frmMDIMain.fnShowInsuredForm
        Unload Me
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
        Case 5      ' Invalid Procedure Call or Argument
            ' Don't know why it's occurring but it's harmless so ignore it.
            Resume Next
        Case 402    ' Must close or hide topmost modal form first
            ' If we detect an invalid DB path upon starting up the app, the frmSetDatabaseLocation
            ' form is opened modally so the user can set a valid path. However, we can't open
            ' that form modally until the splash screen has been unloaded.
            Unload Me
            Resume PROC_EXIT
        Case Else
            ' Save Err object data, if not already saved
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub
