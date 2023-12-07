Attribute VB_Name = "modReporting"
'******************************************************************************
' Module     : modReporting
' Description:
' Procedures :
'
'
' Modified   :
' 04/30/02 BAW Made the Crystal Application object a global variable:
'              defined in modReporting; instantiated in modStartup; deallocated in
'              fnDeallocateGlobalObjects. This avoids "Out of memory" errors
'              when the frmReportViewer screen is displayed.
' --------------------------------------------------
Option Explicit
Option Compare Binary
Private Const mcstrName As String = "modReporting."

    
Public gcReportToPrint As CRAXDRT.Report
Public gcrxApp As CRAXDRT.Application


'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function fnSetFormulaField(ByVal strFormulaName As String, ByVal strFormulaText As String) As Boolean
    ' Comments  : Sets the value of the named Crystal .RPT formula field. Derived from
    '             p592 of George Peck's "Crystal Reports 8: The Complete Reference" book.
    '
    '             NOTE: Assumes the caller set gcReportToPrint to point to the
    '                   correct .RPT file prior to calling this procedure.
    '
    '             NOTE 2: The formulae names are CASE-SENSITIVE ! ! !
    '
    ' Parameters: strFormulaName (in) = the name of the formula field (without a "@")
    '             strFormulaText (in) = the text of the formula, in Crystal syntax
    '
    ' Returns   : True if named formula was found and updated; False otherwise
    '
    ' Called by : fnPrepare_xxx( ) of frmPrintReports
    '             cmdOK_Click( ) of frmInsured
    '
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc   As String = "fnSetFormulaField"
    Dim intCounter          As Integer

    With gcReportToPrint
        For intCounter = 1 To .FormulaFields.Count
            If .FormulaFields(intCounter).FormulaFieldName = strFormulaName Then
                .FormulaFields(intCounter).Text = fnQuoted(strFormulaText)
                fnSetFormulaField = True
                GoTo PROC_EXIT
            End If
        Next intCounter
    End With
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
            ' Save Err object data, if not already saved
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function





'////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub fnViewReport()
    ' Comments  : Displays the report in the Report Viewer window.
    '             It assumes that the all of the Report's properties
    '             were set appropriately (e.g., RecordSelectionFormula,
    '             SetDataSource, etc.) before this routine was called.
    ' Parameters: N/A
    '
    ' Called by : cmdOK_Click( ) in frmSelectReports
    '
    ' Returns   : N/A
    '
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnViewReport"
    Dim frmChild As Form

    ' Report Viewer Form's Load event will automatically display report.
    Set frmChild = New frmReportViewer
    frmChild.Show vbModal
PROC_EXIT:
    On Error GoTo 0     ' disable error handler

    ' Clean-up statements go here
   ' Terminate the Crystal Report Viewer window, removing it from the Forms collection
    fnFreeObject frmChild

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
