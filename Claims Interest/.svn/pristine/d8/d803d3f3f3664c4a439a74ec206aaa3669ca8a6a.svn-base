VERSION 5.00
Begin VB.Form frmSearchForClaimNumber 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Search for Claim Number"
   ClientHeight    =   675
   ClientLeft      =   3015
   ClientTop       =   2910
   ClientWidth     =   4155
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   675
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2820
      TabIndex        =   1
      ToolTipText     =   "Display the first record found with the specified Claim Number"
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtClmNum 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "txtClmNum"
      ToolTipText     =   "Type in all or part of a Claim Number"
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmSearchForClaimNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
' Module     : frmSearchForClaimNumber
' Description:
' Procedures:
'              cmdOK_Click()
'              Form_Load()
'              Form_Unload(ByRef pintCancel As Integer)
'              txtClmNum_KeyPress(ByRef pintKeyAscii As Integer)
' Modified   :
' 06/18/01 BAW Added comments and standardized variable names
' --------------------------------------------------
Option Explicit
Option Compare Binary

Private mstrScreenName As String

Private Const mclngMinFormWidth As Long = 4245
Private Const mclngMinFormHeight As Long = 1005



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub cmdOK_Click()
    ' Comments  : Return to the caller (frmInsured), where the user's
    '             input will be saved and this form will be unloaded
    ' Parameters: N/A
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "cmdOK_Click"

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

   frmSearchForClaimNumber.Hide
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
    ' Comments  : Open the requested report in a modal
    '             Crystal Report 8 viewer window.
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "Form_Load"

    ' Set the screen name that will be used to form the Title on message boxes
    mstrScreenName = Me.Caption

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    ' Identify the icons that will be used for the form and the picture next to the Lookup ComboBox
    Me.Icon = LoadResPicture(gcRES_ICON_MAINAPP, vbResIcon)

    ' If the user has ever opened this form before, restore its size & placement.
    ' If the restore would result in the form being off-screen, just center it instead.
    If gapsApp.RestoreForm(Me) = False Then
        With Me
            .Width = mclngMinFormWidth
            .Height = mclngMinFormHeight
        End With
        fnCenterFormOnMDI frmMDIMain, Me
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
Private Sub Form_Unload(ByRef pintCancel As Integer)
    ' Comments  : Unloads the form
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "Form_Unload"

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    gapsApp.SaveForm Me
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
Private Sub txtClmNum_KeyPress(ByRef pintKeyAscii As Integer)
    ' Comments  : Allow/disallow keyboard entry as appropriate
    ' Parameters: pintKeyAscii - ASCII key code of key just pressed
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "txtClmNum_KeyPress"

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    ' Interpret the user pressing Escape as a request to cancel the search
    If pintKeyAscii = vbKeyEscape Then
       frmSearchForClaimNumber.Hide
       txtClmNum.Text = "blankval"
    End If

    ' Ignore characters other than 0-9, A-Z, a-z, and whatever 1-31 is...
    If Not ((pintKeyAscii > 0 And pintKeyAscii < 32) Or (pintKeyAscii > 47 And pintKeyAscii < 58) Or (pintKeyAscii > 64 And pintKeyAscii < 91) Or (pintKeyAscii > 96 And pintKeyAscii < 123)) Then
       pintKeyAscii = 0
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
