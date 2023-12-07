VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmPrintReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Report"
   ClientHeight    =   4350
   ClientLeft      =   2850
   ClientTop       =   2070
   ClientWidth     =   6315
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraSpecifyCriteria 
      Caption         =   "Specify criteria"
      Height          =   1815
      Left            =   270
      TabIndex        =   4
      Top             =   1680
      Width           =   5775
      Begin VB.Frame fraLineOfBusiness 
         Caption         =   "Line of Business"
         Height          =   1275
         Left            =   3000
         TabIndex        =   9
         Top             =   240
         Width           =   2355
         Begin VB.OptionButton optLOB 
            Caption         =   "All of the a&bove"
            Height          =   315
            Index           =   2
            Left            =   240
            TabIndex        =   12
            ToolTipText     =   "All lines-of-business"
            Top             =   840
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton optLOB 
            Caption         =   "&Individual"
            Height          =   315
            Index           =   0
            Left            =   240
            TabIndex        =   10
            ToolTipText     =   "Individual line-of-business"
            Top             =   210
            Width           =   1095
         End
         Begin VB.OptionButton optLOB 
            Caption         =   "&Group"
            Height          =   315
            Index           =   1
            Left            =   240
            TabIndex        =   11
            ToolTipText     =   "Group line-of-business"
            Top             =   525
            Width           =   795
         End
      End
      Begin MSComCtl2.DTPicker dtpFromDt 
         Height          =   375
         Left            =   1200
         TabIndex        =   6
         ToolTipText     =   "The date from which data should appear on the report"
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "MM/dd/yyy"
         Format          =   55246851
         CurrentDate     =   37025
         MinDate         =   21916
      End
      Begin MSComCtl2.DTPicker dtpToDt 
         Height          =   375
         Left            =   1200
         TabIndex        =   8
         ToolTipText     =   "The date through which data should appear on the report"
         Top             =   780
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "MM/dd/yyy"
         Format          =   55246851
         CurrentDate     =   37025
         MinDate         =   21916
      End
      Begin VB.Label lblFromDt 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fro&m Date:"
         Height          =   195
         Left            =   255
         TabIndex        =   5
         Top             =   450
         Width           =   810
      End
      Begin VB.Label lblToDt 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "T&o Date:"
         Height          =   195
         Left            =   255
         TabIndex        =   7
         Top             =   870
         Width           =   630
      End
   End
   Begin VB.Frame fraSelectReport 
      Caption         =   "Select a report"
      Height          =   1335
      Left            =   270
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      Begin VB.OptionButton optReport 
         Caption         =   "&Data Integrity Issues Report"
         Height          =   315
         Index           =   2
         Left            =   300
         TabIndex        =   3
         ToolTipText     =   "Information about possible bad data in the database"
         Top             =   840
         Width           =   2820
      End
      Begin VB.OptionButton optReport 
         Caption         =   "&Custom Claim Payment Report"
         Height          =   315
         Index           =   1
         Left            =   300
         TabIndex        =   2
         ToolTipText     =   "Information about amounts paid to each Payee on each claim"
         Top             =   540
         Width           =   2820
      End
      Begin VB.OptionButton optReport 
         Caption         =   "&State Interest Report"
         Height          =   315
         Index           =   0
         Left            =   300
         TabIndex        =   1
         ToolTipText     =   "Information about total amounts paid across all claims for each state"
         Top             =   240
         Value           =   -1  'True
         Width           =   2820
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "C&lose"
      Height          =   375
      Left            =   3300
      TabIndex        =   14
      ToolTipText     =   "Close this screen"
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "O&K"
      Default         =   -1  'True
      Height          =   375
      Left            =   1800
      TabIndex        =   13
      ToolTipText     =   "View/Print the selected report"
      Top             =   3720
      Width           =   1215
   End
End
Attribute VB_Name = "frmPrintReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'!TODO! Consider whether any controls s/b changed to listboxes or comboboxes
'******************************************************************************
' Module     : frmPrintReport
' Description:
' Procedures:
'              cmdClose_Click()
'              cmdOK_Click()
'              fnClearControls()
'              fnEnableLOB(ByVal bEnable As Boolean)
'              Function fnGetReportFile() As String
'              fnInspectObjects()                           (DEBUGGING USE ONLY)
'              fnPrepare_CustomClaimPaymentReport
'              fnPrepare_DataIntegrityIssuesReport
'              fnPrepare_StateInterestReport()
'              fnSetFocusToFirstUpdateableField()
'              fnValidData() As Boolean
'              fnWarningData()
'              Form_Load()
'              Form_Unload(ByRef pintCancel as Integer)
' Modified   :
'
' --------------------------------------------------
Option Explicit
Option Compare Binary

Private mstrScreenName As String

' mrstReportData is the recordset of data that, along with formula fields and parameter fields (if appropriate),
' that is sent to the frmReportViewer form in order to preview or print a report. This recordset is restruck
' in cmdPrintPreview_click, by each fnPrepare_XXX( ) method.
Private mrstReportData        As ADODB.Recordset

' The following dates are used to set the default From/To dates when the user selects a report
Private mdteFirstDayOfPrevMonth As Date
Private mdteLastDayOfPrevMonth  As Date

#If False Then
' mconArchiveDB points to the Archive SQL Server database corresponding to the "active" database to which the
' user is currently logged on. This is created anew during Form_Load and destroyed in Form_Unload.
Private mconArchiveDB         As cconConnection
#End If

Private Const mclngMinFormWidth As Long = 8760  '!TODO! - Check these values
Private Const mclngMinFormHeight As Long = 5055


' Define a constant for each field that may get an error or warning. This
' should match the text of that control's associated Label control.
'!TODO! Add new Data Integrity report
Private Const mcstrOptStateReportLabel      As String = "State Report"
Private Const mcstrOptCustomDateReportLabel As String = "Custom Date Report"
Private Const mcstrDtpFromDateLabel         As String = "From Date"
Private Const mcstrDtToDateLabel            As String = "To Date"

Dim mctlFirstEditableField As Control

'-----------------------------------------------------------------------
' The following Enum represents which Report option button
' was selected
'-----------------------------------------------------------------------
Public Enum EnumReport
    erpt_StateInterestReport = 0
    erpt_CustomClaimPaymentReport = 1
    erpt_DataIntegrityIssuesReport = 2
End Enum

'-----------------------------------------------------------------------
' The following Enum represents which Line of Business option button
' was selected
'-----------------------------------------------------------------------
Public Enum EnumLOB
    elob_Individual = 0
    elob_Group = 1
    elob_AllOfTheAbove = 2
End Enum



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub cmdClose_Click()
    ' Comments  : Closes this form
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "cmdClose_Click"

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
Private Sub cmdOK_Click()
    ' Comments  : Open the requested report in a modal
    '             Crystal Report 8 viewer window.
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "cmdOK_Click"
    Dim hrgHourglass As chrgHourglass

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    If fnValidData Then
        Set hrgHourglass = New chrgHourglass
        hrgHourglass.value = True

        Select Case True
            Case optReport(erpt_StateInterestReport).value
                fnPrepare_StateInterestReport
            Case optReport(erpt_CustomClaimPaymentReport).value
                fnPrepare_CustomClaimPaymentReport
            Case optReport(erpt_DataIntegrityIssuesReport).value
                fnPrepare_DataIntegrityIssuesReport
            Case Else
                gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_UNEXPECTED_VAL_SELECT_CASE, _
                    mstrScreenName & gcstrDOT & cstrCurrentProc
                GoTo PROC_EXIT

        End Select

        hrgHourglass.value = False

        ' DEBUGDEBUG -- uncomment out the next line to investigate the data sent to the report -- DEBUGDEBUG
        fnPersistRecordsetToCSV mrstReportData, "c:\ReportData.csv"

        ' Print report to modal Viewer window
        fnViewReport

        ' Initialize controls, in case user wants to do another report
        fnClearControls

        ' Make sure this window is shown on top of all other windows in the app
        ' after the Viewer window is closed
        fnSetTopmostWindow Me, bTopmost:=True
    End If      ' if fnValidData returned False, indicating it found errors
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here
    fnFreeObject hrgHourglass
    ' Close the recordset, but don't bother to set to Nothing; This will be done when the
    ' form is unloaded.
    If Not (mrstReportData Is Nothing) Then
        If mrstReportData.State = adStateOpen Then
            mrstReportData.Close
        End If
    End If

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
Private Sub fnClearControls()
    ' Comments  : Initializes controls to their default settings
    ' Called by : Form_Initialize, cmdOK_Click
    ' Parameters: N/A
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnClearControls"

    optLOB(elob_AllOfTheAbove) = True        ' Select the "All of the above" LOB button

    ' Select the State Interest Report option button. The optReport_Click event
    ' handler will ensure all controls appropriate for that report are enabled
    ' and others, if any, are disabled.
    optReport(erpt_StateInterestReport) = True
    
    ' DateTimePicker controls (dtpFromDt and dtpToDt) will
    ' automatically be set to today's date. Cannot set them to Null
    ' unless their CheckBox property is set to True.
    dtpFromDt.value = mdteFirstDayOfPrevMonth
    dtpToDt.value = mdteLastDayOfPrevMonth
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
Private Sub fnEnableLOB(ByVal bEnable As Boolean)
    ' Comments  : Enables/Disables entry to the Line-of-Business
    '             controls
    ' Parameters: bEnable=True to enable/unlock it; False otherwise
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnEnableLOB"
    Dim ctl As Control

    For Each ctl In Controls
        If ctl.Container.Name = fraLineOfBusiness.Name Then
            ctl.Enabled = bEnable
        End If
    Next ctl
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
Private Sub fnPrepare_CustomClaimPaymentReport()
    ' Comments  : Prepares the Report object to produce the
    '             Custom Claim Payment Report
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc   As String = "fnPrepare_CustomClaimPaymentReport"
    Const cstrSQLView       As String = "dbo.CustomClaimPaymentReport_v"
    Dim strSQL              As String
    Dim strWhereDate        As String
    Dim strWhereLOB         As String
    Dim strOrderBy          As String
    Dim strLOBDesc          As String
    Dim crDB                As CRAXDRT.Database

    Set gcReportToPrint = gcrxApp.OpenReport(fnGetReportFile())
    Set crDB = gcReportToPrint.Database

    ' Build an ADODB.Recordset containing the info to appear on the report
    strWhereDate = " WHERE paye_pmt_dt BETWEEN '" & dtpFromDt.value & "' AND '" & dtpToDt.value & "'" & vbCr

    ' Build SQL string for optional report criteria:  Line of Business
    Select Case True
        Case optLOB(elob_Individual).value
            strWhereLOB = " AND lob_cd = 'I'"
            strLOBDesc = "Individual"
        Case optLOB(elob_Group).value
            strWhereLOB = " AND lob_cd = 'G'"
            strLOBDesc = "Group"
        Case Else
            ' Get all of them
            strWhereLOB = " "
            strLOBDesc = "[All]"
    End Select

    strOrderBy = " ORDER BY clm_num, paye_full_nm"

    strSQL = "SELECT * from " & cstrSQLView & strWhereDate & strWhereLOB & strOrderBy

    Set mrstReportData = gadwApp.Execute_SQL_AsRST(gconAppActive, strSQL)
    #If DEBUG_RST Then
        Debug.Print "In " & cstrCurrentProc & ", " & CStr(mrstReportData.RecordCount) & " records were retrieved in the rst."
        Debug.Print "SQL statement is: " & vbCr & strSQL
    #End If

    ' Disconnect the recordset
    mrstReportData.ActiveConnection = Nothing

    ' ...............................................................................
    ' Set formula field(s) in the report that supply additional info that
    ' is not in the recordset (typically singularly-occuring data)
    ' ...............................................................................
    fnSetFormulaField "formulaReportName", "Custom Claim Payment Report"
    fnSetFormulaField "formulaReportPeriodDescript", "Report Criteria:" & _
        " Date of Payment between " & CStr(dtpFromDt.value) & " and " & CStr(dtpToDt.value) & _
        " and LOB=" & strLOBDesc

    ' ...............................................................................
    ' Tell the report where the data is coming from (overriding whatever might
    ' have been set at design-time). All of the following is necessary since
    ' the location and Connect string set within the .RPT itself may not be
    ' accurate in a production environment (or even on another developer's PC)
    ' ...............................................................................
    With crDB
        .SetDataSource mrstReportData
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



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub fnPrepare_DataIntegrityIssuesReport()
    ' Comments  : Prepares the Report object to produce the
    '             State Interest Report
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc   As String = "fnPrepare_DataIntegrityIssuesReport"
    Const cstrSQLView       As String = "dbo.DataIntegrityIssuesReport_v"
    Dim strSQL              As String
    Dim strWhereDate        As String
    Dim strWhereReason      As String
    Dim strOrderBy          As String
    Dim crDB                As CRAXDRT.Database

    Set gcReportToPrint = gcrxApp.OpenReport(fnGetReportFile())
    Set crDB = gcReportToPrint.Database

    ' Build an ADODB.Recordset containing the info to appear on the report
    strWhereDate = " WHERE paye_pmt_dt BETWEEN '" & dtpFromDt.value & "' AND '" & dtpToDt.value & "'" & vbCr
    strWhereReason = " AND calcReason <> ''"
    strOrderBy = " ORDER BY clm_num, paye_full_nm"

    strSQL = "SELECT * from " & cstrSQLView & strWhereDate & strWhereReason & strOrderBy

    Set mrstReportData = gadwApp.Execute_SQL_AsRST(gconAppActive, strSQL)
    #If DEBUG_RST Then
        Debug.Print "In " & cstrCurrentProc & ", " & CStr(mrstReportData.RecordCount) & " records were retrieved in the rst."
        Debug.Print "SQL statement is: " & vbCr & strSQL
    #End If

    ' Disconnect the recordset
    mrstReportData.ActiveConnection = Nothing

    ' ...............................................................................
    ' Set formula field(s) in the report that supply additional info that
    ' is not in the recordset (typically singularly-occuring data)
    ' ...............................................................................
    fnSetFormulaField "formulaReportName", "Data Integrity Issues Report"
    fnSetFormulaField "formulaReportPeriodDescript", "Reported Period: " & CStr(dtpFromDt.value) & " to " & CStr(dtpToDt.value)

    ' ...............................................................................
    ' Tell the report where the data is coming from (overriding whatever might
    ' have been set at design-time). All of the following is necessary since
    ' the location and Connect string set within the .RPT itself may not be
    ' accurate in a production environment (or even on another developer's PC)
    ' ...............................................................................
    With crDB
        .SetDataSource mrstReportData
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


'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub fnPrepare_StateInterestReport()
    ' Comments  : Prepares the Report object to produce the
    '             State Interest Report
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc   As String = "fnPrepare_StateInterestReport"
    Const cstrSQLView       As String = "dbo.StateInterestReport_v"
    Dim strSQL              As String
    Dim strWhereDate        As String
    Dim strWhereLOB         As String
    Dim strOrderBy          As String
    Dim strLOBDesc          As String
    Dim crDB                As CRAXDRT.Database

    Set gcReportToPrint = gcrxApp.OpenReport(fnGetReportFile())
    Set crDB = gcReportToPrint.Database

    ' Build an ADODB.Recordset containing the info to appear on the report
    strWhereDate = " WHERE paye_pmt_dt BETWEEN '" & dtpFromDt.value & "' AND '" & dtpToDt.value & "'" & vbCr

    ' Build SQL string for optional report criteria:  Line of Business
    Select Case True
        Case optLOB(elob_Individual).value
            strWhereLOB = " AND lob_cd = 'I'"
            strLOBDesc = "Individual"
        Case optLOB(elob_Group).value
            strWhereLOB = " AND lob_cd = 'G'"
            strLOBDesc = "Group"
        Case Else
            ' Get all of them
            strWhereLOB = " "
            strLOBDesc = "[All]"
    End Select

    strOrderBy = " ORDER BY paye_st_cd"

    strSQL = "SELECT * from " & cstrSQLView & strWhereDate & strWhereLOB & strOrderBy

    Set mrstReportData = gadwApp.Execute_SQL_AsRST(gconAppActive, strSQL)
    #If DEBUG_RST Then
        Debug.Print "In " & cstrCurrentProc & ", " & CStr(mrstReportData.RecordCount) & " records were retrieved in the rst."
        Debug.Print "SQL statement is: " & vbCr & strSQL
    #End If

    ' Disconnect the recordset
    mrstReportData.ActiveConnection = Nothing

    ' ...............................................................................
    ' Set formula field(s) in the report that supply additional info that
    ' is not in the recordset (typically singularly-occuring data)
    ' ...............................................................................
    fnSetFormulaField "formulaReportName", "State Interest Report"
    fnSetFormulaField "formulaReportPeriodDescript", "Report Criteria:" & _
        " Date of Payment between " & CStr(dtpFromDt.value) & " and " & CStr(dtpToDt.value) & _
        " and LOB=" & strLOBDesc

    ' ...............................................................................
    ' Tell the report where the data is coming from (overriding whatever might
    ' have been set at design-time). All of the following is necessary since
    ' the location and Connect string set within the .RPT itself may not be
    ' accurate in a production environment (or even on another developer's PC)
    ' ...............................................................................
    With crDB
        .SetDataSource mrstReportData
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


' Dead code per Project Analyzer
'Private Sub fnSetFocusToFirstUpdateableField()
'    '----------------------------------------------------------------------------
'    ' Procedure :  Sub fnSetFocusToFirstUpdateableField
'    ' Created by:  BAW on 04-26-2001 08:55
'    '
'    ' Comments  : Moves the focus to the first editable field on the screen
'    ' Called by :
'    ' Parameters: N/A
'    '
'    ' Modified  :
'    '----------------------------------------------------------------------------
'    On Error GoTo PROC_ERR
'    Const cstrCurrentProc As String = "fnSetFocusToFirstUpdateableField"
'
'    ' Set focus to first editable field, by default
'    If mctlFirstEditableField.Visible Then
'        mctlFirstEditableField.SetFocus
'    End If
'PROC_EXIT:
'    On Error Resume Next
'    Exit Sub
'PROC_ERR:
'    Select Case Err.Number
'    'Case statements for expected errors go here
'    Case Else
'        ' Display msgbox re: fatal error and terminate the app
'        fnProcessFatalError mcstrCurrentModule & "." & cstrCurrentProc, _
'                                 fte_DefaultErrType, Err.Number, _
'                                 Err.Description, Err.Source, _
'                                 Err.HelpFile, Err.HelpContext
'    End Select
'    Resume PROC_EXIT
'End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Function fnValidData() As Boolean
    ' Comments  : Determines if all data is valid, including
    '             whether all required fields have been input.
    '             This function is called by cmdOK_Click.
    '             If a data error is found, it returns False
    '             which directs the caller to stop processing.
    '             It also generates warnings, by calling
    '             WarningData(), but only if no errors were
    '             found up to that point.
    ' Parameters: N/A
    ' Returns   : True if all data is valid; False otherwise
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnValidData"
    Dim bErrorFound As Boolean
    Dim ctlFirstToFail As Control
    Dim intFailures As Integer
    Dim strMsgText As String

    fnValidData = True

    ' Check the fields in a left-to-right, top-to-bottom screen sequence.
    '     1. optStateReport
    '     2. Custom Date Report
    '     3. Data Integrity Issues Report
    '     4. dtpFromDt
    '     5. dtpToDt

    ' ------------- 2.  Verify other characteristics are valid --------------

    ' Disallow a future-dated Start Date
    If DateValue(dtpFromDt.value) > Date Then
        intFailures = intFailures + 1
        Set ctlFirstToFail = dtpFromDt
        strMsgText = strMsgText & vbCrLf & _
                     "The " & mcstrDtpFromDateLabel & " (" & dtpFromDt.value & _
                     ") cannot be in the future."
    End If

    ' Disallow an End Date more than 5 days future dated. (This used to just be "today", but
    ' now that Michelle wants the Date of Paymment to support being up to 5 days future-dated,
    ' then this logic had to also be adjusted.
    If DateValue(dtpToDt.value) > DateAdd("d", 5, Date) Then
        intFailures = intFailures + 1
        Set ctlFirstToFail = dtpToDt
        strMsgText = strMsgText & vbCrLf & _
                     "The " & mcstrDtToDateLabel & " (" & dtpToDt.value & _
                     ") cannot be more than 5 days in the future."
    End If

    ' End Date must on or after Start Date
    If DateValue(dtpToDt.value) < DateValue(dtpFromDt.value) Then
        intFailures = intFailures + 1
        Set ctlFirstToFail = dtpToDt
        strMsgText = strMsgText & vbCrLf & _
                     "The " & mcstrDtToDateLabel & " (" & dtpToDt.value & _
                     ") must be on or after the " & mcstrDtpFromDateLabel & _
                     " (" & dtpFromDt.value & ")."
    End If

    If intFailures <> 0 Then
        bErrorFound = True
        fnValidData = False
        If ctlFirstToFail.Visible Then
            ctlFirstToFail.SetFocus
        End If
        gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_CROSS_FLD_VALIDATIONS_FAILED, _
                                   mstrScreenName & gcstrDOT & cstrCurrentProc, _
                                   "the report can be produced", strMsgText
        GoTo PROC_EXIT
    End If

    ' If no errors found, continue with checking for warnings
    If Not bErrorFound Then
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
    ' Returns   : N/A
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnWarningData"

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


#If False Then
'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub Form_Initialize()
    ' Comments  : Initializes the form
    ' Parameters: N/A
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "Form_Initialize"

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    ' Set the Start Date as the first editable control, e.g., the one
    ' which will get the initial focus.
    Set mctlFirstEditableField = optReport(0)

    ' Initialize all controls to their default settings
    fnClearControls
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
Private Sub Form_Load()
    ' Comments  : Initializes the form
    ' Parameters: N/A
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "Form_Initialize"

    ' Set the screen name that will be used to form the Title on message boxes
    mstrScreenName = Me.Caption

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    ' Identify the icons that will be used for the form and the picture next to the Lookup ComboBox
    Me.Icon = LoadResPicture(gcRES_ICON_MAINAPP, vbResIcon)

    ' Moved the instantiation of the Crystal object here (from modStartup) as a conditional instantiation
    ' since this CreateObject invocation is such a pig per VB Watch Profiler.
    If (gcrxApp Is Nothing) Then
        Set gcrxApp = CreateObject("CrystalRuntime.Application")
    End If

    mdteFirstDayOfPrevMonth = fnFirstDayOfMonth(DateAdd("m", -1, Date))
    mdteLastDayOfPrevMonth = fnLastDayOfMonth(DateAdd("m", -1, Date))
    
    ' Set the Start Date as the first editable control, e.g., the one
    ' which will get the initial focus.
    Set mctlFirstEditableField = optReport(0)

    ' Initialize all controls to their default settings
    fnClearControls

    ' Set availability of Data Integrity Issues report based on whether the user is a member
    ' of the USERADMIN or SUPPORT user roles. If so, enable it; otherwise disable it.
    optReport(erpt_DataIntegrityIssuesReport).Enabled = gconAppActive.LastLogonIsSpecialUser
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
    ' Comments  : Closes this form
    ' Parameters: pintCancel (in/out), if set to TRUE
    '             then the unload is aborted
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "Form_Unload"

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    #If False Then
    ' DO NOT disconnect gconAppActive() or set it Nothing!!
    With mconArchiveDB
        If Not (mconArchiveDB Is Nothing) Then
            If .State = adStateOpen Then
                .Disconnect
            End If
            Set mconArchiveDB = Nothing
        End If
    End With
    #End If

    Unload Me

    ' Following needed to ensure this form will be deleted from the Forms collection
    ' This may not work as intended. (Might set the wrong form reference, or
    ' might not actually "take" (i.e. releasing all memory) if there are
    ' other variables that reference it.
    fnFreeObject frmPrintReport
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
Private Sub optReport_Click(ByRef pintIndex As Integer)
    ' Comments  : Enables/disables criteria as appropriate,
    '             given the user's report selection
    ' Parameters: pintIndex (in), indicates which option
    '             button was selected
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "optReport_Click"

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    Select Case pintIndex
        Case erpt_StateInterestReport
            fnEnableLOB True
            dtpFromDt.value = mdteFirstDayOfPrevMonth
            dtpToDt.value = mdteLastDayOfPrevMonth
        Case erpt_CustomClaimPaymentReport
            fnEnableLOB True
            dtpFromDt.value = mdteFirstDayOfPrevMonth
            dtpToDt.value = mdteLastDayOfPrevMonth
        Case erpt_DataIntegrityIssuesReport
            fnEnableLOB False
            ' Set default From/To date range to be wide open so *all* issues will be shown
            dtpFromDt.value = dtpFromDt.MinDate
            dtpToDt.value = Date
        Case Else
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_UNEXPECTED_VAL_SELECT_CASE, _
                mstrScreenName & gcstrDOT & cstrCurrentProc
            GoTo PROC_EXIT
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



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Function fnGetReportFile() As String
    ' Comments  : Using the selection in the Select A Report ListBox,
    '             this proc retrieves the corresponding .RPT's filename
    '             from the Report Meta Data array.
    ' Parameters: N/A
    ' Returns   : String - the name of the .RPT file for that report
    ' Modified  :
    '
    ' --------------------------------------------------
    Const cstrCurrentProc As String = "fnGetReportFile"
    Const cstrUnknown     As String = "Unknown"
    Dim fso               As Scripting.FileSystemObject

    On Error GoTo PROC_ERR

    Set fso = New Scripting.FileSystemObject

    Select Case True
        Case optReport(erpt_StateInterestReport).value
            fnGetReportFile = fso.BuildPath(App.Path, "StateInterest_CR8.RPT")
        Case optReport(erpt_CustomClaimPaymentReport).value
            fnGetReportFile = fso.BuildPath(App.Path, "CustomClaimPayment_CR8.RPT")
        Case optReport(erpt_DataIntegrityIssuesReport).value
            fnGetReportFile = fso.BuildPath(App.Path, "DataIntegrityIssues_cr8.rpt")
        Case Else
            fnGetReportFile = cstrUnknown
     End Select


    ' Non-fatal error if .RPT doesn't exist or if we couldn't determine the .RPT filename
    If Not (fso.FileExists(fnGetReportFile)) Or fnGetReportFile = cstrUnknown Then
        gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_RPTFILE_NOT_FOUND, _
                                   mstrScreenName & gcstrDOT & cstrCurrentProc, _
                                   fnGetReportFile
        GoTo PROC_EXIT
    End If
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here
    fnFreeObject fso

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
