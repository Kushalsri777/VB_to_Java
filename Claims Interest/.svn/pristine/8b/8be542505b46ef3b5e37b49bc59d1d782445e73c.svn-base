VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmReportViewer 
   Caption         =   "Report Viewer"
   ClientHeight    =   10095
   ClientLeft      =   1425
   ClientTop       =   870
   ClientWidth     =   10785
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   10095
   ScaleWidth      =   10785
   Begin CRVIEWERLibCtl.CRViewer crxViewer 
      Height          =   9915
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10605
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   -1  'True
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmReportViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
' Module     : frmReportViewer
' Description:
' Procedures:
'              Form_Load)
'              Form_Resize()
'              Form_Unload(ByRef pintCancel As Integer)
'
' Modified   :
'
' --------------------------------------------------
Option Explicit
Option Compare Binary

Private mstrScreenName As String
Private Const mclngMinFormWidth As Long = 10905
Private Const mclngMinFormHeight As Long = 10530




' ' member variable for ReportToPrint property
' Private m_ReportToPrint As Object

' ' Used by other forms (such as frmInsured and frmPrintReport) to print a Crystal Report .RPT file
' Property Get ReportToPrint() As Object
'     Set ReportToPrint = m_ReportToPrint
' End Property
' Property Set ReportToPrint(ByVal newValue As Object)
'     Set m_ReportToPrint = newValue
' End Property


Private Sub Form_Load()
    ' Comments  : Open the requested report in a modal
    '             Crystal Report 8 viewer window.
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "Form_Load"
    Dim hrgHourglass As chrgHourglass

    ' Set the screen name that will be used to form the Title on message boxes
    mstrScreenName = Me.Caption

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    Me.Icon = LoadResPicture(gcRES_ICON_MAINAPP, vbResIcon)

    ' If the user has ever opened this form before, restore its size & placement.
    ' If the restore would result in the form being off-screen, just center it instead.
    If gapsApp.RestoreForm(Me) = False Then
        'fnSetFormSize
        With Me
            .Width = mclngMinFormWidth
            .Height = mclngMinFormHeight
        End With
        'fnCenterFormOnScreen Me
        fnCenterFormOnMDI frmMDIMain, Me
    End If

    Set hrgHourglass = New chrgHourglass
    hrgHourglass.value = True

    ' Set ReportSource using a Property Get on frmPrintReports2
    crxViewer.ReportSource = gcReportToPrint
    crxViewer.EnableCloseButton = True

    ' View the report
    crxViewer.ViewReport
    ' Or, print it without viewing...
    '       crxViewer.PrintReport

    hrgHourglass.value = False
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



Private Sub Form_Resize()
    ' Comments  : Resize the viewer control so it fills
    '             the form window
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "Form_Resize"

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    'With Me
    '    If .Width < mclngMinFormWidth Then
    '        .Width = mclngMinFormWidth
    '    End If
    '    If .Height < mclngMinFormHeight Then
    '        .Height = mclngMinFormHeight
    '    End If
    'End With

    crxViewer.Top = 0
    crxViewer.Left = 0
    crxViewer.Height = ScaleHeight
    crxViewer.Width = ScaleWidth
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
