VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmGenerateTaxFile 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generate Tax File"
   ClientHeight    =   4455
   ClientLeft      =   2925
   ClientTop       =   3135
   ClientWidth     =   8070
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGenerateTaxFile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdViewPRTaxFile 
      Caption         =   "View &Puerto Rico Tax File"
      Default         =   -1  'True
      Height          =   375
      Left            =   5520
      TabIndex        =   16
      ToolTipText     =   "Generate the tax file"
      Top             =   3960
      Width           =   2415
   End
   Begin VB.CommandButton cmdViewNonPRTaxFile 
      Caption         =   "View &Non-Puerto Rico Tax File"
      Height          =   375
      Left            =   5520
      TabIndex        =   15
      ToolTipText     =   "Generate the tax file"
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Frame fraSelectionCriteria 
      Caption         =   "Selection Criteria"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      Begin MSComCtl2.DTPicker dtpToDate 
         Height          =   315
         Left            =   1575
         TabIndex        =   4
         ToolTipText     =   "The end of the date range for which tax reporting should be done"
         Top             =   720
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MM/dd/yyy"
         Format          =   59572227
         CurrentDate     =   37986
         MinDate         =   21916
      End
      Begin MSComCtl2.DTPicker dtpFromDate 
         Height          =   315
         Left            =   1575
         TabIndex        =   2
         ToolTipText     =   "The start of the date range for which tax reporting should be done"
         Top             =   360
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MM/dd/yyy"
         Format          =   59572227
         CurrentDate     =   37622
         MinDate         =   21916
      End
      Begin VB.Label lblToDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&To Date:"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   780
         Width           =   630
      End
      Begin VB.Label lblFromDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fro&m Date:"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   420
         Width           =   810
      End
   End
   Begin VB.Frame fraTaxFiles 
      Caption         =   "Tax Files"
      Height          =   1695
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   7815
      Begin VB.TextBox txtTaxFileFolder 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   1935
         TabIndex        =   11
         Top             =   1080
         Width           =   4755
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Browse"
         Height          =   312
         Left            =   6840
         TabIndex        =   12
         ToolTipText     =   "Select a different Download Folder"
         Top             =   1080
         Width           =   795
      End
      Begin VB.TextBox txtFileNamePR 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   1935
         TabIndex        =   7
         Text            =   "taxfile_pr.txt"
         ToolTipText     =   "The fully qualified file name that should be generated"
         Top             =   360
         Width           =   3195
      End
      Begin VB.TextBox txtFileNameNonPR 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   1920
         TabIndex        =   8
         Text            =   "taxfile.txt"
         ToolTipText     =   "The fully qualified file name that should be generated"
         Top             =   720
         Width           =   3195
      End
      Begin VB.Label lblTaxFileFolder 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Save files to:"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   1140
         Width           =   945
      End
      Begin VB.Label lblFileNamePRData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Puerto Rico Data:"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   420
         Width           =   1275
      End
      Begin VB.Label lblFileNameNonPR 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Non-Puerto Rico Data:"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   780
         Width           =   1620
      End
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0C0C0&
      Caption         =   "C&lose"
      Height          =   375
      Left            =   3825
      TabIndex        =   14
      ToolTipText     =   "Close this screen"
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2430
      TabIndex        =   13
      ToolTipText     =   "Generate the tax file"
      Top             =   3840
      Width           =   1215
   End
End
Attribute VB_Name = "frmGenerateTaxFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'!TODO! This isn't a maintenance screen, so make sure it is set up appropriately,
'      e.g., Dirty, saving/restoring position.
'!TODO! Be sure to test a date range that should get no data, to ensure formerly fatal error is handled correctly.

'******************************************************************************
' Module     : frmGenerateTaxFile
' Description:
' Procedures : cmdClose_Click()
'              cmdOK_Click()
'              fnFillRecord()
'              fnGetTaxFileData(ByVal dteFromDt As date, Byval dteToDt as date) As ADODB.Recordset
'              fnValidData() as Boolean
'              fnWarningData()
'              Form_Initialize()
'              Form_Load()
'              Form_QueryUnload(ByRef pintCancel As Integer, ByRef pintUnloadMode As Integer)
'              Form_Unload(ByRef pintCancel As Integer)

' Modified   :
' 03/26/01 BAW Cleaned with Total Visual CodeTools 2000
' 01/xx/02 BAW Optimized per Project Analyzer (Space/Mid, etc. => Space$/Mid$)
' 03/14/03 BAW Revamped to get all info from sproc
' 12/06/04 BAW (YE2004) Added totals to "tax files generated" message generated by cmdOK_Click, and
'              corrected its logic to identify Puerto Rico data based on Field9, not Field11.
' --------------------------------------------------
Option Explicit
Option Compare Binary

Private mstrScreenName As String

Private mbHasFinishedSuccessfully           As Boolean
Private mbIsInProgress                      As Boolean

Private Const mclngMinFormWidth             As Long = 8160
Private Const mclngMinFormHeight            As Long = 4830

Private Const mcstrDtpFromDateLabel         As String = "From Date"
Private Const mcstrDtpToDateLabel           As String = "To Date"
Private Const mcstrTxtTaxFileFolderLabel    As String = "Save Files To"

Private mctlFirstEditableField              As Control

Private m_strTaxFile_PR                     As String
Private m_strTaxFile_NonPR                  As String
Private m_tsTaxFile_PR                      As Scripting.TextStream
Private m_tsTaxFile_NonPR                   As Scripting.TextStream
'-----------------------------------------------------------------------
' The following Enum represents which type of tax file is currently
' being worked with and is used by the fnTaxFile_XXX procedures.
'-----------------------------------------------------------------------
Private Enum EnumTaxFile
    etf_PuertoRico = 0
    etf_NonPuertoRico = 1
End Enum



'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'\                                                                  /
'|                Procedures and Event Handlers                     |
'/                                                                  \
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/



'////////////////////////////////////////////////////////////////////////////////////////
Private Sub cmdBrowse_Click()
    ' Browse for Drive/Folder
    On Error GoTo PROC_ERR
    Const cstrCurrentProc   As String = "cmdBrowse_Click"
    Dim brfFolder           As New cbrfBrowseFolder
    Dim strFolderName       As String
    
    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName
    
    With brfFolder
        .hWnd = Me.hWnd                          ' Owner of the BrowseFolder window
        .Title = "Select a Folder to which the tax files will be saved"    ' Title
        .Folder = txtTaxFileFolder.Text        ' Initial folder
        .Flags = BIF_RETURNONLYFSDIRS          ' Default flags
        strFolderName = .ShowBrowse()          ' Go get it
        If Not .Cancelled Then
            txtTaxFileFolder.Text = UCase$(fnAddBackslash(strFolderName))
            Me.Refresh
        End If
    End With
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    
    ' Clean-up statements go here
    fnFreeObject brfFolder

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
Private Sub cmdClose_Click()
    ' Comments  : Returns user to previous screen upon exiting
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
    ' Comments  : Creates the ASCII file that will be sent to
    '             the Tax system to do yearend tax
    '             reporting
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc       As String = "cmdOK_Click"
    Const cstrPuertoRico        As String = "PR"
    Const cstrUnused            As String = "!"
    Const cstrGeneralNumber     As String = "#,###,##0"                 ' "General Number"
    Const cstrCurrency          As String = "Currency"
    Dim dblTotInt_PR            As Double
    Dim dblTotIntWthld_PR       As Double
    Dim dblTotInt_NonPR         As Double
    Dim dblTotIntWthld_NonPR    As Double
    Dim fld                     As ADODB.Field
    Dim hrgHourglass            As chrgHourglass
    Dim intRecCtr_PR            As Integer
    Dim intRecCtr_NonPR         As Integer
    Dim rstTaxFileData          As ADODB.Recordset


    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName
    
    mbHasFinishedSuccessfully = False
    mbIsInProgress = True
    
    ' Disables/Hides controls so user cannot do anything while batch is in progress
    fnSetAvailabilityOfControls

    If fnValidData Then
        Set hrgHourglass = New chrgHourglass
        hrgHourglass.value = True
    
        ' Save user's current setting to the registry as a Per User setting
        gapsApp.TaxFileFolder = txtTaxFileFolder.Text
    
        fnTaxFile_Open (etf_PuertoRico)
        fnTaxFile_Open (etf_NonPuertoRico)
        
        Set rstTaxFileData = fnGetTaxFileData(dtpFromDate.value, dtpToDate.value)

        With rstTaxFileData
            If Not (.BOF And .EOF) Then
                .MoveFirst
                Do While Not .EOF
                    If !Fld09 <> cstrPuertoRico Then                                      ' YE2004
                        For Each fld In .Fields
                            If fld.value <> cstrUnused Then
                                m_tsTaxFile_NonPR.Write fld.value
                            End If
                        Next fld
                        m_tsTaxFile_NonPR.WriteLine
                        intRecCtr_NonPR = intRecCtr_NonPR + 1
                        ' Amounts have implied decimals, so divide by 100 to make them explicit
                        dblTotInt_NonPR = dblTotInt_NonPR + (!Fld19.value / 100)           ' YE2004
                        dblTotIntWthld_NonPR = dblTotIntWthld_NonPR + (!Fld22.value / 100) ' YE2004
                    Else
                        For Each fld In .Fields
                            If fld.value <> cstrUnused Then
                                m_tsTaxFile_PR.Write fld.value
                            End If
                        Next fld
                        m_tsTaxFile_PR.WriteLine
                        intRecCtr_PR = intRecCtr_PR + 1
                        ' Amounts have implied decimals, so divide by 100 to make them explicit
                        dblTotInt_PR = dblTotInt_PR + (!Fld19.value / 100)                  ' YE2004
                        dblTotIntWthld_PR = dblTotIntWthld_PR + (!Fld22.value / 100)        ' YE2004

                    End If
                    .MoveNext
                Loop    ' while not rstTaxFileData.EOF
                ' Indicate success
                ' gcRES_INFO_TAX_FILE_GEND (1004) = @@1 record(s) were written to the @@2 tax file. The total interest
                '        (Box 1) amount was @@3 and the total Interest Withheld (Box 4) amount was @@4.@@CRLF
                '        @@5 record(s) were written to the @@6 tax file. The total interest (Box 1) amount
                '        was @@7 and the total Interest Withheld (Box 4) amount was @@8.
                gerhApp.ReportNonFatal vbObjectError + gcRES_INFO_TAX_FILE_GEND, _
                                       mstrScreenName & gcstrDOT & cstrCurrentProc, _
                                       Format(intRecCtr_PR, cstrGeneralNumber), UCase$(txtFileNamePR.Text), _
                                       Format(dblTotInt_PR, cstrCurrency), Format(dblTotIntWthld_PR, cstrCurrency), _
                                       Format(intRecCtr_NonPR, cstrGeneralNumber), UCase$(txtFileNameNonPR.Text), _
                                       Format(dblTotInt_NonPR, cstrCurrency), Format(dblTotIntWthld_NonPR, cstrCurrency)
            Else
                ' Tell the user no records met the selection criteria
                ' gcRES_NERR_NO_RECS_WERE_FOUND (4004) = No records were found with @@1.
                gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_NO_RECS_WERE_FOUND, _
                                       mstrScreenName & gcstrDOT & cstrCurrentProc, _
                                       "Payment Dates on or within the specified date"
                GoTo PROC_EXIT
            End If
        End With
    End If
    
    mbHasFinishedSuccessfully = True
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0

    ' Clean-up statements go here
    mbIsInProgress = False
    ' Disables/Hides controls so user cannot do anything while batch is in progress
    fnSetAvailabilityOfControls
    
    If Not (hrgHourglass Is Nothing) Then
        fnFreeObject hrgHourglass
    End If
    If Not (m_tsTaxFile_PR Is Nothing) Then
        m_tsTaxFile_PR.Close
    End If
    If Not (m_tsTaxFile_NonPR Is Nothing) Then
        m_tsTaxFile_NonPR.Close
    End If
    fnFreeRecordset rstTaxFileData
    fnFreeObject m_tsTaxFile_PR
    fnFreeObject m_tsTaxFile_NonPR
    
    ' Report the error, since this is an event handler
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.ReportFatalError mstrScreenName
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case 76 ' Path Not Found
            ' 4003 = The drive or path specified does not exist. Please be sure to specify an existing drive and directory.
            gerhApp.ReportNonFatal vbObjectError + gcRES_NERR_DRIVE_OR_PATH_NOT_FOUND, _
                                   mstrScreenName & gcstrDOT & cstrCurrentProc
        Case Else
            ' Save Err object data, if not already saved
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub





'////////////////////////////////////////////////////////////////////////////////////////////////
Private Function fnGetTaxFileData(ByVal dteFromDt As Date, ByVal dteToDt As Date) As ADODB.Recordset
    '--------------------------------------------------------------------------
    ' Procedure:   fnGetTaxFileData
    ' Description: This stored procedure builds a recordset of Payee & Claim
    '              information for use in preparing a tax file in the necessary
    '              I.R.S. TVTAXFORM layout for the PC.
    '
    '
    ' Parameters:
    '     dteFromDt (in) - the earliest Payee Date of Payment to select
    '     dteToDt   (in) - the latest Payee Date of Payment to select
    '
    ' Returns:     A disconnected ADODB.Recordset
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc          As String = "fnGetTaxFileData"
    Const cstrSproc                As String = "dbo.proc_tax_file_layout_generate"
    Dim rstTemp                    As ADODB.Recordset
    Dim prmReturnValue             As ADODB.Parameter
    Dim prmPayePmtDtFromDt         As ADODB.Parameter
    Dim prmPayePmtDtToDt           As ADODB.Parameter
    Dim adwTemp                    As cadwADOWrapper

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
        ' Define the PAYE_PMT_DT_FROM_DATE parameter
        Set prmPayePmtDtFromDt = .CreateParameter(Name:="paye_pmt_dt_from_date", _
                                         Type:=adDBDate, _
                                         Direction:=adParamInput, _
                                         value:=dteFromDt)
        .Parameters.Append prmPayePmtDtFromDt

        ' ---Parameter #3---
        ' Define the PAYE_PMT_DT_TO_DATE parameter
        Set prmPayePmtDtToDt = .CreateParameter(Name:="paye_pmt_dt_to_date", _
                                         Type:=adDBDate, _
                                         Direction:=adParamInput, _
                                         value:=dteToDt)
        .Parameters.Append prmPayePmtDtToDt

        Set rstTemp = .Execute()
    End With
    
    rstTemp.ActiveConnection = Nothing
    Set fnGetTaxFileData = rstTemp
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
    
    ' Clean-up statements go here
    
    ' Do *not* do "fnFreeRecordset rstTemp" since this will cause the recordset returned
    ' by this function to be wiped out as well!
    fnFreeObject prmReturnValue
    fnFreeObject prmPayePmtDtFromDt
    fnFreeObject prmPayePmtDtToDt
    
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case prmReturnValue
        Case gcRES_NERR_REC_NOT_FOUND       ' 4027 -  The @@1 is invalid. @@2
            ' Note that the following error is presented as an ATYPICAL 4027 error!
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_INVALID_DATA, _
                                       mstrScreenName & gcstrDOT & cstrCurrentProc, _
                                       "From Date or To Date", _
                                       "The tax data cannot be retrieved " & _
                                       "when any of these fields are NULL. FromDt=[" & _
                                       FormatDateTime(dteFromDt, vbShortDate) & "], ToDt=[" & _
                                       FormatDateTime(dteToDt, vbShortDate) & "]"
            Resume PROC_EXIT
        Case gcRES_NERR_ERR_WHILE_TRYING_TO   ' 4028
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_ERR_WHILE_TRYING_TO, _
                                       mstrScreenName & gcstrDOT & cstrCurrentProc, _
                                       "locate"
            Resume PROC_EXIT
    End Select

    ' If any other errors exist, i.e. in Err object, then let it fall through into default error handling.

    Select Case Err.Number
        Case -2147217900 ' Object not found
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_FERR_SPROC_NOT_FOUND, _
                                       mstrScreenName & gcstrDOT & cstrCurrentProc, _
                                       cstrSproc
            Resume PROC_EXIT
        Case Else
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub fnSetAvailabilityOfControls()
    ' Comments  :  Sets up file names and on-screen controls based on the select As Of Date.
    ' Parameters:
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc               As String = "fnSetAvailabilityOfControls"
    
    ' If we're in the process of generating the tax files, then disable/hide the View buttons;
    ' Otherwise, enable/show them
    
    If mbHasFinishedSuccessfully And (Not mbIsInProgress) Then
        cmdViewNonPRTaxFile.Enabled = True
        cmdViewPRTaxFile.Enabled = True
    Else
        cmdViewNonPRTaxFile.Enabled = False
        cmdViewPRTaxFile.Enabled = False
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
Private Sub fnTaxFile_Open(ByVal intTaxFile As EnumTaxFile)
    ' Comments  : Determines if all data is valid, including
    '             whether all required fields have been input.
    '             This function is called by cmdOK_Click.
    '             If a data error is found, it returns False
    '             which directs the caller to stop processing.
    '             It also generates warnings, by calling
    '             WarningData(), but only if no errors were
    '             found up to that point.
    ' Parameters: Enum representing tax file to open
    ' Returns   : N/A
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc   As String = "fnTaxFile_Open"
    Dim fso                 As Scripting.FileSystemObject

    Set fso = New Scripting.FileSystemObject
    
    Select Case intTaxFile
        Case etf_NonPuertoRico
            m_strTaxFile_NonPR = fnBuildQualifiedFileName(gapsApp.TaxFileFolder, txtFileNameNonPR.Text)
            Set m_tsTaxFile_NonPR = fso.CreateTextFile(m_strTaxFile_NonPR, True)
        Case etf_PuertoRico
            m_strTaxFile_PR = fnBuildQualifiedFileName(gapsApp.TaxFileFolder, txtFileNamePR.Text)
            Set m_tsTaxFile_PR = fso.CreateTextFile(m_strTaxFile_PR, True)
        Case Else
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_UNEXPECTED_VAL_SELECT_CASE, _
                mstrScreenName & gcstrDOT & cstrCurrentProc
            GoTo PROC_EXIT
    End Select
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0

    ' Clean-up statements go here
    fnFreeObject fso
    
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
    Const cstrCurrentProc   As String = "fnValidData"
    Dim bErrorFound         As Boolean
    Dim ctlFirstToFail      As Control
    Dim fso                 As New Scripting.FileSystemObject
    Dim intFailures         As Integer
    Dim strFieldList        As String
    Dim strMsgText          As String

    fnValidData = True

    ' Check the fields in a left-to-right, top-to-bottom screen sequence.
    '     1. From Date
    '     2. To Date
    '     3. Save files to (folder)

    ' ------------- First, verify required fields are missing --------------

    ' --No required fields at this time--
    'If IsNull(txtFileNameNonPR) Or txtFileNameNonPR = vbNullString Then
    '    If intFailures = 0 Then
    '        strFieldList = vbCrLf & cstrTxtFileNameLabel
    '        Set ctlFirstToFail = txtFileNameNonPR
    '    Else
    '        strFieldList = strFieldList & vbCrLf & cstrTxtFileNameLabel
    '    End If
    '    intFailures = intFailures + 1
    'End If

    If intFailures <> 0 Then
        bErrorFound = True
        fnValidData = False
        If ctlFirstToFail.Visible Then
            ctlFirstToFail.SetFocus
        End If
        ' 4041 = The following required fields must be supplied before your request can be processed:@@CRLF@@1
        gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_REQD_FIELDS_MISSING, _
                                   mstrScreenName & gcstrDOT & cstrCurrentProc, _
                                   strFieldList
        GoTo PROC_EXIT
    End If



    ' ------------------- Now, do cross-field validations --------------------


    intFailures = 0     ' Reset for this section of error validations

    If DateValue(dtpToDate.value) < DateValue(dtpFromDate.value) Then
        intFailures = intFailures + 1
        Set ctlFirstToFail = dtpToDate
        strMsgText = strMsgText & vbCrLf & _
                     "The " & mcstrDtpToDateLabel & " (" & dtpToDate.value & _
                     ") must be on or after the " & mcstrDtpFromDateLabel & " (" & _
                     dtpFromDate.value & ")."
    End If

    ' At app startup, capsAppSettings ensured that a Per User (HKCU) registry entry was built
    ' to define where to place downloaded files and it also tried to create that folder if
    ' it didn't already exist since we can't create files there until it exists!. Now,
    ' let's verify that folder actually exists. If it doesn't, then do an error here.
    ' (It's more logical to the user to generate the error when we actually **need** the
    ' folder, than at app startup.
    If Not fso.FolderExists(txtTaxFileFolder.Text) Then
        intFailures = intFailures + 1
        Set ctlFirstToFail = mctlFirstEditableField
        strMsgText = strMsgText & vbCrLf & _
                     "The " & mcstrTxtTaxFileFolderLabel & " folder (" & txtTaxFileFolder.Text & _
                     ") does not exist. " & "That folder must be created or a different one selected " & _
                     "before the tax files can be generated."
    End If

    If intFailures <> 0 Then
        bErrorFound = True
        fnValidData = False
        If ctlFirstToFail.Visible Then
            ctlFirstToFail.SetFocus
        End If
        ' 4034 = Cross-field validation errors were found. These must be corrected before @@1:@@2
        gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_CROSS_FLD_VALIDATIONS_FAILED, _
                                   mstrScreenName & gcstrDOT & cstrCurrentProc, _
                                   "your request can be processed", strMsgText
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

    ' Add logic to display warning messages, if any
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
Private Sub cmdViewNonPRTaxFile_Click()
    ' Comments  :
    ' Parameters:
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc   As String = "cmdViewNonPRTaxFile_Click"
    Dim lngReturnCode       As Long
    Dim strFileNm           As String
    
    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName
    
    strFileNm = fnBuildQualifiedFileName(gapsApp.TaxFileFolder, txtFileNameNonPR.Text)
    lngReturnCode = Shell("notepad.exe " & strFileNm, vbNormalFocus)
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
        Case -2147022987
            ' gcRES_WARN_POSSIBLE_FILESYS_PERM_PROBLEM (2010)
            ' An error was encountered while trying to @@1. This may be due to
            ' network unavailability or insufficient authorizations. @@2
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_WARN_POSSIBLE_FILESYS_PERM_PROBLEM, _
                                       mstrScreenName & gcstrDOT & cstrCurrentProc, _
                                       "view the Non-Puerto Rico tax file", vbCr & Err.Description
            Resume PROC_EXIT
        Case Else
            ' Save Err object data, if not already saved
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub


'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub cmdViewPRTaxFile_Click()
    ' Comments  :
    ' Parameters:
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc   As String = "cmdViewPRTaxFile_Click"
    Dim lngReturnCode       As Long
    Dim strFileNm           As String
    
    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName
    
    strFileNm = fnBuildQualifiedFileName(gapsApp.TaxFileFolder, txtFileNamePR.Text)
    lngReturnCode = Shell("notepad.exe " & strFileNm, vbNormalFocus)
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
        Case -2147022987
            ' gcRES_WARN_POSSIBLE_FILESYS_PERM_PROBLEM (2010)
            ' An error was encountered while trying to @@1. This may be due to
            ' network unavailability or insufficient authorizations. @@2
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_WARN_POSSIBLE_FILESYS_PERM_PROBLEM, _
                                       mstrScreenName & gcstrDOT & cstrCurrentProc, _
                                       "view the Puerto Rico tax file", vbCr & Err.Description
            Resume PROC_EXIT
        Case Else
            ' Save Err object data, if not already saved
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub




'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    ' Comments  : Adds password to Connect string and initializes
    '             bound controls
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

    ' Set defaults for controls, including:
    '   * From Date = Jan 1st of previous year
    '   * To Date = Dec 31st of previous year
    dtpFromDate.value = DateSerial(Year(Date) - 1, 1, 1)
    dtpToDate.value = DateSerial(Year(Date) - 1, 12, 31)
    txtFileNamePR.Text = "taxfile_pr.txt"
    txtFileNameNonPR.Text = "taxfile.txt"
    txtTaxFileFolder.Text = gapsApp.TaxFileFolder
    
    ' Disables/Hides controls so user cannot do anything while batch is in progress
    fnSetAvailabilityOfControls
    
    ' Set the control to receive the focus after errors (the first editable field
    ' on the screen)
    Set mctlFirstEditableField = dtpFromDate
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
Private Sub Form_QueryUnload(ByRef pintCancel As Integer, ByRef pintUnloadMode As Integer)
    ' Comments  : If the user clicks the Close button ("X" in the upper
    '             right corner of the form), close the form.
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "Form_QueryUnload"

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    If pintUnloadMode = vbFormControlMenu Then
        pintCancel = True
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
        'Case statements for expected errors go here
        Case Else
            ' Save Err object data, if not already saved
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub


'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub Form_Unload(ByRef pintCancel As Integer)
    ' Comments  : Close the form
    ' Parameters: pintCancel (in/out), if set to True
    '             the unload is aborted
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "Form_Unload"

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName
    
    gapsApp.SaveForm Me

    Unload Me

    ' Following needed to ensure this form will be deleted from the Forms collection
    ' This may not work as intended. (Might set the wrong form reference, or
    ' might not actually "take" (i.e. releasing all memory) if there are
    ' other variables that reference it.
    fnFreeObject frmGenerateTaxFile
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


