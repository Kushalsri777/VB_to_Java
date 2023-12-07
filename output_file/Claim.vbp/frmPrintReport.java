
import java.util.Date;

public class frmPrintReport {

  //!TODO! Consider whether any controls s/b changed to listboxes or comboboxes
  //******************************************************************************
  // Module     : frmPrintReport
  // Description:
  // Procedures:
  //              cmdClose_Click()
  //              cmdOK_Click()
  //              fnClearControls()
  //              fnEnableLOB(ByVal bEnable As Boolean)
  //              Function fnGetReportFile() As String
  //              fnInspectObjects()                           (DEBUGGING USE ONLY)
  //              fnPrepare_CustomClaimPaymentReport
  //              fnPrepare_DataIntegrityIssuesReport
  //              fnPrepare_StateInterestReport()
  //              fnSetFocusToFirstUpdateableField()
  //              fnValidData() As Boolean
  //              fnWarningData()
  //              Form_Load()
  //              Form_Unload(ByRef pintCancel as Integer)
  // Modified   :
  //
  // --------------------------------------------------
//Option Explicit
  *Option Compare Binary

  private String mstrScreenName = "";

  // mrstReportData is the recordset of data that, along with formula fields and parameter fields (if appropriate),
  // that is sent to the frmReportViewer form in order to preview or print a report. This recordset is restruck
  // in cmdPrintPreview_click, by each fnPrepare_XXX( ) method.
  private DBRecordSet mrstReportData;

  // The following dates are used to set the default From/To dates when the user selects a report
  private Date mdteFirstDayOfPrevMonth = null;
  private Date mdteLastDayOfPrevMonth = null;

  *#If False Then
  // mconArchiveDB points to the Archive SQL Server database corresponding to the "active" database to which the
  // user is currently logged on. This is created anew during Form_Load and destroyed in Form_Unload.
  private cconConnection mconArchiveDB;
  *#End If

  //'!TODO! - Check these values
  private static final Long MCLNGMINFORMWIDTH = 8760;
  private static final Long MCLNGMINFORMHEIGHT = 5055;


  // Define a constant for each field that may get an error or warning. This
  // should match the text of that control's associated Label control.
  //!TODO! Add new Data Integrity report
  private static final String MCSTROPTSTATEREPORTLABEL = "State Report";
  private static final String MCSTROPTCUSTOMDATEREPORTLABEL = "Custom Date Report";
  private static final String MCSTRDTPFROMDATELABEL = "From Date";
  private static final String MCSTRDTTODATELABEL = "To Date";

  Control mctlFirstEditableField = null;

  //-----------------------------------------------------------------------
  // The following Enum represents which Report option button
  // was selected
  //-----------------------------------------------------------------------
//*TODO:** enum is translated as a new class at the end of the file Public Enum EnumReport

  //-----------------------------------------------------------------------
  // The following Enum represents which Line of Business option button
  // was selected
  //-----------------------------------------------------------------------
//*TODO:** enum is translated as a new class at the end of the file Public Enum EnumLOB



  //////////////////////////////////////////////////////////////////////////////////////////////////
  private void cmdClose_Click() {
    // Comments  : Closes this form
    // Parameters: None
    // Modified  :
    // --------------------------------------------------
    try {
      "cmdClose_Click"
.equals(Const cstrCurrentProc As String);

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      Unload(this);
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here

    // Report the error, since this is an event handler
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.reportFatalError(mstrScreenName);
    }
    return;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        // Save Err object data, if not already saved
        modGeneral.gerhApp.saveErrObjectData(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  private void cmdOK_Click() {
    // Comments  : Open the requested report in a modal
    //             Crystal Report 8 viewer window.
    // Parameters: None
    // Modified  :
    // --------------------------------------------------
    try {
      "cmdOK_Click"
.equals(Const cstrCurrentProc As String);
      chrgHourglass hrgHourglass = null;

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      if (fnValidData()) {
        hrgHourglass = new chrgHourglass();
        hrgHourglass.setValue(true);

        switch (true) {
          case  optReport(EnumReport.eRPT_STATEINTERESTREPORT).chrgHourglass.getValue():
            fnPrepare_StateInterestReport();
            break;

          case  optReport(EnumReport.eRPT_CUSTOMCLAIMPAYMENTREPORT).chrgHourglass.getValue():
            fnPrepare_CustomClaimPaymentReport();
            break;

          case  optReport(EnumReport.eRPT_DATAINTEGRITYISSUESREPORT).chrgHourglass.getValue():
            fnPrepare_DataIntegrityIssuesReport();
            break;

          default:
            modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_UNEXPECTED_VAL_SELECT_CASE, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
            // **TODO:** goto found: GoTo PROC_EXIT;

            break;
        }

        hrgHourglass.setValue(false);

        // DEBUGDEBUG -- uncomment out the next line to investigate the data sent to the report -- DEBUGDEBUG
        modGeneral.fnPersistRecordsetToCSV(mrstReportData, "c:\\ReportData.csv");

        // Print report to modal Viewer window
        modReporting.fnViewReport();

        // Initialize controls, in case user wants to do another report
        fnClearControls();

        // Make sure this window is shown on top of all other windows in the app
        // after the Viewer window is closed
        modGeneral.fnSetTopmostWindow(this, bTopmost:=True);
      //' if fnValidData returned False, indicating it found errors
      }
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    modGeneral.fnFreeObject(hrgHourglass);
    // Close the recordset, but don't bother to set to Nothing; This will be done when the
    // form is unloaded.
    if (!(mrstReportData == null)) {
      if (mrstReportData.State == adStateOpen) {
        mrstReportData.Close;
      }
    }

    // Report the error, since this is an event handler
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.reportFatalError(mstrScreenName);
    }
    return;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        // Save Err object data, if not already saved
        modGeneral.gerhApp.saveErrObjectData(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }
}


//////////////////////////////////////////////////////////////////////////////////////////////////
  private void fnClearControls() {
    // Comments  : Initializes controls to their default settings
    // Called by : Form_Initialize, cmdOK_Click
    // Parameters: N/A
    // Modified  :
    // --------------------------------------------------
    try {
      "fnClearControls"
.equals(Const cstrCurrentProc As String);

      //' Select the "All of the above" LOB button
      optLOB(EnumLOB.eLOB_ALLOFTHEABOVE) = true;

      // Select the State Interest Report option button. The optReport_Click event
      // handler will ensure all controls appropriate for that report are enabled
      // and others, if any, are disabled.
      optReport(EnumReport.eRPT_STATEINTERESTREPORT) = true;

      // DateTimePicker controls (dtpFromDt and dtpToDt) will
      // automatically be set to today's date. Cannot set them to Null
      // unless their CheckBox property is set to True.
      dtpFromDt.chrgHourglass.setValue(mdteFirstDayOfPrevMonth);
      dtpToDt.chrgHourglass.setValue(mdteLastDayOfPrevMonth);
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
    }
    return;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  private void fnEnableLOB(boolean bEnable) {
    // Comments  : Enables/Disables entry to the Line-of-Business
    //             controls
    // Parameters: bEnable=True to enable/unlock it; False otherwise
    // Modified  :
    // --------------------------------------------------
    try {
      "fnEnableLOB"
.equals(Const cstrCurrentProc As String);
      Control ctl = null;

      for (int _i = 0; _i < Controls.size(); _i++) {
        ctl = Controls.item(_i);
        if (ctl.Container.Name == fraLineOfBusiness.Name) {
          ctl.Enabled = bEnable;
        }
      }
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
    }
    return;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }
}


//////////////////////////////////////////////////////////////////////////////////////////////////
  private void fnPrepare_CustomClaimPaymentReport() {
    // Comments  : Prepares the Report object to produce the
    //             Custom Claim Payment Report
    // Parameters: None
    // Modified  :
    // --------------------------------------------------
    try {
      "fnPrepare_CustomClaimPaymentReport"
.equals(Const cstrCurrentProc As String);
      "dbo.CustomClaimPaymentReport_v"
.equals(Const cstrSQLView As String);
      String strSQL = "";
      String strWhereDate = "";
      String strWhereLOB = "";
      String strOrderBy = "";
      String strLOBDesc = "";
      CRAXDRT.Database crDB = null;

      modReporting.gcReportToPrint = modReporting.gcrxApp.OpenReport(fnGetReportFile());
      crDB = modReporting.gcReportToPrint.Database;

      // Build an ADODB.Recordset containing the info to appear on the report
      strWhereDate = " WHERE paye_pmt_dt BETWEEN '"+ ((Boolean) dtpFromDt.chrgHourglass.getValue()).toString()+ "' AND '"+ ((Boolean) dtpToDt.chrgHourglass.getValue()).toString()+ "'"+ "\\n";

      // Build SQL string for optional report criteria:  Line of Business
      switch (true) {
        case  optLOB(EnumLOB.eLOB_INDIVIDUAL).chrgHourglass.getValue():
          strWhereLOB = " AND lob_cd = 'I'";
          strLOBDesc = "Individual";
          break;

        case  optLOB(EnumLOB.eLOB_GROUP).chrgHourglass.getValue():
          strWhereLOB = " AND lob_cd = 'G'";
          strLOBDesc = "Group";
          break;

        default:
          // Get all of them
          strWhereLOB = " ";
          strLOBDesc = "[All]";
          break;
      }

      strOrderBy = " ORDER BY clm_num, paye_full_nm";

      strSQL = "SELECT * from "+ cstrSQLView+ strWhereDate+ strWhereLOB+ strOrderBy;

      mrstReportData = modGeneral.gadwApp.execute_SQL_AsRST(modGeneral.gconAppActive, strSQL);
      #If DEBUG_RST Then;
      Debug.Print("In "+ cstrCurrentProc+ ", "+ CStr(mrstReportData.RecordCount)+ " records were retrieved in the rst.");
      Debug.Print("SQL statement is: "+ "\\n"+ strSQL);
      #End If;

      // Disconnect the recordset
      mrstReportData.ActiveConnection = null;

      // ...............................................................................
      // Set formula field(s) in the report that supply additional info that
      // is not in the recordset (typically singularly-occuring data)
      // ...............................................................................
      modReporting.fnSetFormulaField("formulaReportName", "Custom Claim Payment Report");
      modReporting.fnSetFormulaField("formulaReportPeriodDescript", "Report Criteria:"+ " Date of Payment between "+ CStr(dtpFromDt.chrgHourglass.getValue())+ " and "+ CStr(dtpToDt.chrgHourglass.getValue())+ " and LOB="+ strLOBDesc);

      // ...............................................................................
      // Tell the report where the data is coming from (overriding whatever might
      // have been set at design-time). All of the following is necessary since
      // the location and Connect string set within the .RPT itself may not be
      // accurate in a production environment (or even on another developer's PC)
      // ...............................................................................
      crDB.SetDataSource(mrstReportData);
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here

    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
    }
    return;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  private void fnPrepare_DataIntegrityIssuesReport() {
    // Comments  : Prepares the Report object to produce the
    //             State Interest Report
    // Parameters: None
    // Modified  :
    // --------------------------------------------------
    try {
      "fnPrepare_DataIntegrityIssuesReport"
.equals(Const cstrCurrentProc As String);
      "dbo.DataIntegrityIssuesReport_v"
.equals(Const cstrSQLView As String);
      String strSQL = "";
      String strWhereDate = "";
      String strWhereReason = "";
      String strOrderBy = "";
      CRAXDRT.Database crDB = null;

      modReporting.gcReportToPrint = modReporting.gcrxApp.OpenReport(fnGetReportFile());
      crDB = modReporting.gcReportToPrint.Database;

      // Build an ADODB.Recordset containing the info to appear on the report
      strWhereDate = " WHERE paye_pmt_dt BETWEEN '"+ ((Boolean) dtpFromDt.chrgHourglass.getValue()).toString()+ "' AND '"+ ((Boolean) dtpToDt.chrgHourglass.getValue()).toString()+ "'"+ "\\n";
      strWhereReason = " AND calcReason <> ''";
      strOrderBy = " ORDER BY clm_num, paye_full_nm";

      strSQL = "SELECT * from "+ cstrSQLView+ strWhereDate+ strWhereReason+ strOrderBy;

      mrstReportData = modGeneral.gadwApp.execute_SQL_AsRST(modGeneral.gconAppActive, strSQL);
      #If DEBUG_RST Then;
      Debug.Print("In "+ cstrCurrentProc+ ", "+ CStr(mrstReportData.RecordCount)+ " records were retrieved in the rst.");
      Debug.Print("SQL statement is: "+ "\\n"+ strSQL);
      #End If;

      // Disconnect the recordset
      mrstReportData.ActiveConnection = null;

      // ...............................................................................
      // Set formula field(s) in the report that supply additional info that
      // is not in the recordset (typically singularly-occuring data)
      // ...............................................................................
      modReporting.fnSetFormulaField("formulaReportName", "Data Integrity Issues Report");
      modReporting.fnSetFormulaField("formulaReportPeriodDescript", "Reported Period: "+ CStr(dtpFromDt.chrgHourglass.getValue())+ " to "+ CStr(dtpToDt.chrgHourglass.getValue()));

      // ...............................................................................
      // Tell the report where the data is coming from (overriding whatever might
      // have been set at design-time). All of the following is necessary since
      // the location and Connect string set within the .RPT itself may not be
      // accurate in a production environment (or even on another developer's PC)
      // ...............................................................................
      crDB.SetDataSource(mrstReportData);
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here

    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
    }
    return;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }
}


//////////////////////////////////////////////////////////////////////////////////////////////////
  private void fnPrepare_StateInterestReport() {
    // Comments  : Prepares the Report object to produce the
    //             State Interest Report
    // Parameters: None
    // Modified  :
    // --------------------------------------------------
    try {
      "fnPrepare_StateInterestReport"
.equals(Const cstrCurrentProc As String);
      "dbo.StateInterestReport_v"
.equals(Const cstrSQLView As String);
      String strSQL = "";
      String strWhereDate = "";
      String strWhereLOB = "";
      String strOrderBy = "";
      String strLOBDesc = "";
      CRAXDRT.Database crDB = null;

      modReporting.gcReportToPrint = modReporting.gcrxApp.OpenReport(fnGetReportFile());
      crDB = modReporting.gcReportToPrint.Database;

      // Build an ADODB.Recordset containing the info to appear on the report
      strWhereDate = " WHERE paye_pmt_dt BETWEEN '"+ ((Boolean) dtpFromDt.chrgHourglass.getValue()).toString()+ "' AND '"+ ((Boolean) dtpToDt.chrgHourglass.getValue()).toString()+ "'"+ "\\n";

      // Build SQL string for optional report criteria:  Line of Business
      switch (true) {
        case  optLOB(EnumLOB.eLOB_INDIVIDUAL).chrgHourglass.getValue():
          strWhereLOB = " AND lob_cd = 'I'";
          strLOBDesc = "Individual";
          break;

        case  optLOB(EnumLOB.eLOB_GROUP).chrgHourglass.getValue():
          strWhereLOB = " AND lob_cd = 'G'";
          strLOBDesc = "Group";
          break;

        default:
          // Get all of them
          strWhereLOB = " ";
          strLOBDesc = "[All]";
          break;
      }

      strOrderBy = " ORDER BY paye_st_cd";

      strSQL = "SELECT * from "+ cstrSQLView+ strWhereDate+ strWhereLOB+ strOrderBy;

      mrstReportData = modGeneral.gadwApp.execute_SQL_AsRST(modGeneral.gconAppActive, strSQL);
      #If DEBUG_RST Then;
      Debug.Print("In "+ cstrCurrentProc+ ", "+ CStr(mrstReportData.RecordCount)+ " records were retrieved in the rst.");
      Debug.Print("SQL statement is: "+ "\\n"+ strSQL);
      #End If;

      // Disconnect the recordset
      mrstReportData.ActiveConnection = null;

      // ...............................................................................
      // Set formula field(s) in the report that supply additional info that
      // is not in the recordset (typically singularly-occuring data)
      // ...............................................................................
      modReporting.fnSetFormulaField("formulaReportName", "State Interest Report");
      modReporting.fnSetFormulaField("formulaReportPeriodDescript", "Report Criteria:"+ " Date of Payment between "+ CStr(dtpFromDt.chrgHourglass.getValue())+ " and "+ CStr(dtpToDt.chrgHourglass.getValue())+ " and LOB="+ strLOBDesc);

      // ...............................................................................
      // Tell the report where the data is coming from (overriding whatever might
      // have been set at design-time). All of the following is necessary since
      // the location and Connect string set within the .RPT itself may not be
      // accurate in a production environment (or even on another developer's PC)
      // ...............................................................................
      crDB.SetDataSource(mrstReportData);
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here

    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
    }
    return;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }
}


// Dead code per Project Analyzer
//Private Sub fnSetFocusToFirstUpdateableField()
//    '----------------------------------------------------------------------------
//    ' Procedure :  Sub fnSetFocusToFirstUpdateableField
//    ' Created by:  BAW on 04-26-2001 08:55
//    '
//    ' Comments  : Moves the focus to the first editable field on the screen
//    ' Called by :
//    ' Parameters: N/A
//    '
//    ' Modified  :
//    '----------------------------------------------------------------------------
//    On Error GoTo PROC_ERR
//    Const cstrCurrentProc As String = "fnSetFocusToFirstUpdateableField"
//
//    ' Set focus to first editable field, by default
//    If mctlFirstEditableField.Visible Then
//        mctlFirstEditableField.SetFocus
//    End If
//PROC_EXIT:
//    On Error Resume Next
//    Exit Sub
//PROC_ERR:
//    Select Case Err.Number
//    'Case statements for expected errors go here
//    Case Else
//        ' Display msgbox re: fatal error and terminate the app
//        fnProcessFatalError mcstrCurrentModule & "." & cstrCurrentProc, _
//                                 fte_DefaultErrType, Err.Number, _
//                                 Err.Description, Err.Source, _
//                                 Err.HelpFile, Err.HelpContext
//    End Select
//    Resume PROC_EXIT
//End Sub



//////////////////////////////////////////////////////////////////////////////////////////////////
  private boolean fnValidData() {
    boolean _rtn = false;
    // Comments  : Determines if all data is valid, including
    //             whether all required fields have been input.
    //             This function is called by cmdOK_Click.
    //             If a data error is found, it returns False
    //             which directs the caller to stop processing.
    //             It also generates warnings, by calling
    //             WarningData(), but only if no errors were
    //             found up to that point.
    // Parameters: N/A
    // Returns   : True if all data is valid; False otherwise
    // Modified  :
    // --------------------------------------------------
    try {
      "fnValidData"
.equals(Const cstrCurrentProc As String);
      boolean bErrorFound = false;
      Control ctlFirstToFail = null;
      int intFailures = 0;
      String strMsgText = "";

      _rtn = true;

      // Check the fields in a left-to-right, top-to-bottom screen sequence.
      //     1. optStateReport
      //     2. Custom Date Report
      //     3. Data Integrity Issues Report
      //     4. dtpFromDt
      //     5. dtpToDt

      // ------------- 2.  Verify other characteristics are valid --------------

      // Disallow a future-dated Start Date
      if (DateValue(dtpFromDt.chrgHourglass.getValue()) > Date) {
        intFailures = intFailures + 1;
        ctlFirstToFail = dtpFromDt;
        strMsgText = strMsgText+ "\\r\\n"+ "The "+ MCSTRDTPFROMDATELABEL+ " ("+ ((Boolean) dtpFromDt.chrgHourglass.getValue()).toString()+ ") cannot be in the future.";
      }

      // Disallow an End Date more than 5 days future dated. (This used to just be "today", but
      // now that Michelle wants the Date of Paymment to support being up to 5 days future-dated,
      // then this logic had to also be adjusted.
      if (DateValue(dtpToDt.chrgHourglass.getValue()) > DateAdd("d", 5, Date)) {
        intFailures = intFailures + 1;
        ctlFirstToFail = dtpToDt;
        strMsgText = strMsgText+ "\\r\\n"+ "The "+ MCSTRDTTODATELABEL+ " ("+ ((Boolean) dtpToDt.chrgHourglass.getValue()).toString()+ ") cannot be more than 5 days in the future.";
      }

      // End Date must on or after Start Date
      if (DateValue(dtpToDt.chrgHourglass.getValue()) < DateValue(dtpFromDt.chrgHourglass.getValue())) {
        intFailures = intFailures + 1;
        ctlFirstToFail = dtpToDt;
        strMsgText = strMsgText+ "\\r\\n"+ "The "+ MCSTRDTTODATELABEL+ " ("+ ((Boolean) dtpToDt.chrgHourglass.getValue()).toString()+ ") must be on or after the "+ MCSTRDTPFROMDATELABEL+ " ("+ ((Boolean) dtpFromDt.chrgHourglass.getValue()).toString()+ ").";
      }

      if (intFailures != 0) {
        bErrorFound = true;
        _rtn = false;
        if (ctlFirstToFail.Visible) {
          ctlFirstToFail.SetFocus;
        }
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_CROSS_FLD_VALIDATIONS_FAILED, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc, "the report can be produced", strMsgText);
        // **TODO:** goto found: GoTo PROC_EXIT;
      }

      // If no errors found, continue with checking for warnings
      if (!bErrorFound) {
        fnWarningData();
      }
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    modGeneral.fnFreeObject(ctlFirstToFail);

    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  private void fnWarningData() {
    // Comments  : Validates fields, generating warnings if appropriate.
    //             It should NOT cause ValidData (this procedure's caller)
    //             to return False, since we want updates to proceed.
    // Parameters: N/A
    // Returns   : N/A
    // Modified  :
    // --------------------------------------------------
    try {
      "fnWarningData"
.equals(Const cstrCurrentProc As String);

      // ***   Currently there are no warnings  :(   ***
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
    }
    return;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }
}


  *#If False Then
  //////////////////////////////////////////////////////////////////////////////////////////////////
  private void form_Initialize() {
    // Comments  : Initializes the form
    // Parameters: N/A
    // Modified  :
    // --------------------------------------------------
    try {
      "Form_Initialize"
.equals(Const cstrCurrentProc As String);

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      // Set the Start Date as the first editable control, e.g., the one
      // which will get the initial focus.
      mctlFirstEditableField = optReport(0);

      // Initialize all controls to their default settings
      fnClearControls();
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here

    // Report the error, since this is an event handler
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.reportFatalError(mstrScreenName);
    }
    return;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        // Save Err object data, if not already saved
        modGeneral.gerhApp.saveErrObjectData(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }
}
  *#End If


  //////////////////////////////////////////////////////////////////////////////////////////////////
  private void form_Load() {
    // Comments  : Initializes the form
    // Parameters: N/A
    // Modified  :
    // --------------------------------------------------
    try {
      "Form_Initialize"
.equals(Const cstrCurrentProc As String);

      // Set the screen name that will be used to form the Title on message boxes
      mstrScreenName = Me.Caption;

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      // Identify the icons that will be used for the form and the picture next to the Lookup ComboBox
      Me.Icon = LoadResPicture(modResConstants.gCRES_ICON_MAINAPP, vbResIcon);

      // Moved the instantiation of the Crystal object here (from modStartup) as a conditional instantiation
      // since this CreateObject invocation is such a pig per VB Watch Profiler.
      if ((modReporting.gcrxApp == null)) {
        modReporting.gcrxApp = CreateObject("CrystalRuntime.Application");
      }

      mdteFirstDayOfPrevMonth = modGeneral.fnFirstDayOfMonth(DateAdd("m", -1, Date));
      mdteLastDayOfPrevMonth = modGeneral.fnLastDayOfMonth(DateAdd("m", -1, Date));

      // Set the Start Date as the first editable control, e.g., the one
      // which will get the initial focus.
      mctlFirstEditableField = optReport(0);

      // Initialize all controls to their default settings
      fnClearControls();

      // Set availability of Data Integrity Issues report based on whether the user is a member
      // of the USERADMIN or SUPPORT user roles. If so, enable it; otherwise disable it.
      optReport(EnumReport.eRPT_DATAINTEGRITYISSUESREPORT).Enabled = modGeneral.gconAppActive.getLastLogonIsSpecialUser();
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here

    // Report the error, since this is an event handler
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.reportFatalError(mstrScreenName);
    }
    return;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        // Save Err object data, if not already saved
        modGeneral.gerhApp.saveErrObjectData(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }
}

//////////////////////////////////////////////////////////////////////////////////////////////////
  private void form_Unload(int pintCancel) { // TODO: Use of ByRef founded Private Sub Form_Unload(ByRef pintCancel As Integer)
    // Comments  : Closes this form
    // Parameters: pintCancel (in/out), if set to TRUE
    //             then the unload is aborted
    // Modified  :
    // --------------------------------------------------
    try {
      "Form_Unload"
.equals(Const cstrCurrentProc As String);

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      #If False Then;
      // DO NOT disconnect gconAppActive() or set it Nothing!!
      if (!(mconArchiveDB == null)) {
        if (mconArchiveDB.getState() == adStateOpen) {
          mconArchiveDB.disconnect();
        }
        mconArchiveDB = null;
      }
      #End If;

      Unload(this);

      // Following needed to ensure this form will be deleted from the Forms collection
      // This may not work as intended. (Might set the wrong form reference, or
      // might not actually "take" (i.e. releasing all memory) if there are
      // other variables that reference it.
      modGeneral.fnFreeObject(frmPrintReport);
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here

    // Report the error, since this is an event handler
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.reportFatalError(mstrScreenName);
    }
    return;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        // Save Err object data, if not already saved
        modGeneral.gerhApp.saveErrObjectData(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  private void optReport_Click(int pintIndex) { // TODO: Use of ByRef founded Private Sub optReport_Click(ByRef pintIndex As Integer)
    // Comments  : Enables/disables criteria as appropriate,
    //             given the user's report selection
    // Parameters: pintIndex (in), indicates which option
    //             button was selected
    // Modified  :
    // --------------------------------------------------
    try {
      "optReport_Click"
.equals(Const cstrCurrentProc As String);

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      switch (pintIndex) {
        case  EnumReport.eRPT_STATEINTERESTREPORT:
          fnEnableLOB(true);
          dtpFromDt.chrgHourglass.setValue(mdteFirstDayOfPrevMonth);
          dtpToDt.chrgHourglass.setValue(mdteLastDayOfPrevMonth);
          break;

        case  EnumReport.eRPT_CUSTOMCLAIMPAYMENTREPORT:
          fnEnableLOB(true);
          dtpFromDt.chrgHourglass.setValue(mdteFirstDayOfPrevMonth);
          dtpToDt.chrgHourglass.setValue(mdteLastDayOfPrevMonth);
          break;

        case  EnumReport.eRPT_DATAINTEGRITYISSUESREPORT:
          fnEnableLOB(false);
          // Set default From/To date range to be wide open so *all* issues will be shown
          dtpFromDt.chrgHourglass.setValue(dtpFromDt.MinDate);
          dtpToDt.chrgHourglass.setValue(Date);
          break;

        default:
          modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_UNEXPECTED_VAL_SELECT_CASE, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
          // **TODO:** goto found: GoTo PROC_EXIT;
          break;
      }
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here

    // Report the error, since this is an event handler
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.reportFatalError(mstrScreenName);
    }
    return;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        // Save Err object data, if not already saved
        modGeneral.gerhApp.saveErrObjectData(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  private String fnGetReportFile() {
    String _rtn = "";
    // Comments  : Using the selection in the Select A Report ListBox,
    //             this proc retrieves the corresponding .RPT's filename
    //             from the Report Meta Data array.
    // Parameters: N/A
    // Returns   : String - the name of the .RPT file for that report
    // Modified  :
    //
    // --------------------------------------------------
    "fnGetReportFile"
.equals(Const cstrCurrentProc As String);
    "Unknown"
.equals(Const cstrUnknown As String);
    Scripting.FileSystemObject fso = null;

    try {

      fso = new Scripting.FileSystemObject();

      switch (true) {
        case  optReport(EnumReport.eRPT_STATEINTERESTREPORT).chrgHourglass.getValue():
          _rtn = fso.BuildPath(App.Path, "StateInterest_CR8.RPT");
          break;

        case  optReport(EnumReport.eRPT_CUSTOMCLAIMPAYMENTREPORT).chrgHourglass.getValue():
          _rtn = fso.BuildPath(App.Path, "CustomClaimPayment_CR8.RPT");
          break;

        case  optReport(EnumReport.eRPT_DATAINTEGRITYISSUESREPORT).chrgHourglass.getValue():
          _rtn = fso.BuildPath(App.Path, "DataIntegrityIssues_cr8.rpt");
          break;

        default:
          _rtn = cstrUnknown;
          break;
      }


      // Non-fatal error if .RPT doesn't exist or if we couldn't determine the .RPT filename
      if (!(fso.FileExists(fnGetReportFile())) || fnGetReportFile().equals(cstrUnknown)) {
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_RPTFILE_NOT_FOUND, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc, fnGetReportFile());
        // **TODO:** goto found: GoTo PROC_EXIT;
      }
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    modGeneral.fnFreeObject(fso);

    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        // Save Err object data, if not already saved
        modGeneral.gerhApp.saveErrObjectData(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}
}

public class EnumReport {
    public static final int ERPT_STATEINTERESTREPORT = 0;
    public static final int ERPT_CUSTOMCLAIMPAYMENTREPORT = 1;
    public static final int ERPT_DATAINTEGRITYISSUESREPORT = 2;
}


public class EnumLOB {
    public static final int ELOB_INDIVIDUAL = 0;
    public static final int ELOB_GROUP = 1;
    public static final int ELOB_ALLOFTHEABOVE = 2;
}




// Controller 

package controllers.logged.modules.general

import controllers._
import play.api.mvc._
import play.api.data._
import play.api.data.Forms._
import actions._
import play.api.Logger
import play.api.libs.json._
import models.cairo.modules.general._
import models.cairo.system.security.CairoSecurity
import models.cairo.system.database.DBHelper


case class RmprintreportData(
              id: Option[Int],

              )

object Rmprintreports extends Controller with ProvidesUser {

  val rmprintreportForm = Form(
    mapping(
      "id" -> optional(number),

  )(RmprintreportData.apply)(RmprintreportData.unapply))

  implicit val rmprintreportWrites = new Writes[Rmprintreport] {
    def writes(rmprintreport: Rmprintreport) = Json.obj(
      "id" -> Json.toJson(rmprintreport.id),
      C.ID -> Json.toJson(rmprintreport.id),

    )
  }

  def get(id: Int) = GetAction { implicit request =>
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.LIST_RMPRINTREPORT), { user =>
      Ok(Json.toJson(Rmprintreport.get(user, id)))
    })
  }

  def update(id: Int) = PostAction { implicit request =>
    Logger.debug("in rmprintreports.update")
    rmprintreportForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      rmprintreport => {
        Logger.debug(s"form: ${rmprintreport.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.EDIT_RMPRINTREPORT), { user =>
          Ok(
            Json.toJson(
              Rmprintreport.update(user,
                Rmprintreport(
                       id,

                ))))
        })
      }
    )
  }

  def create = PostAction { implicit request =>
    Logger.debug("in rmprintreports.create")
    rmprintreportForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      rmprintreport => {
        Logger.debug(s"form: ${rmprintreport.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.NEW_RMPRINTREPORT), { user =>
          Ok(
            Json.toJson(
              Rmprintreport.create(user,
                Rmprintreport(

                ))))
        })
      }
    )
  }

  def delete(id: Int) = PostAction { implicit request =>
    Logger.debug("in rmprintreports.delete")
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.DELETE_RMPRINTREPORT), { user =>
      Rmprintreport.delete(user, id)
      // Backbonejs requires at least an empty json object in the response
      // if not it will call errorHandler even when we responded with 200 OK :P
      Ok(JsonUtil.emptyJson)
    })
  }

}

// Model

package models.cairo.modules.general

import java.sql.{Connection, CallableStatement, ResultSet, Types, SQLException}
import anorm.SqlParser._
import anorm._
import services.DateUtil
import services.db.DB
import models.cairo.system.database.{DBHelper, Register, Field, FieldType, SaveResult}
import play.api.Play.current
import models.domain.CompanyUser
import java.util.Date
import play.api.Logger
import play.api.libs.json._
import scala.util.control.NonFatal

case class Rmprintreport(
              id: Int,
,
              createdAt: Date,
              updatedAt: Date,
              updatedBy: Int) {

  def this(
      id: Int,
) = {

    this(
      id,
,
      DateUtil.currentTime,
      DateUtil.currentTime,
      DBHelper.NoId)
  }

  def this(
) = {

    this(
      DBHelper.NoId,
)

  }

}

object Rmprintreport {

  lazy val emptyRmprintreport = Rmprintreport(
)

  def apply(
      id: Int,
) = {

    new Rmprintreport(
      id,
)
  }

  def apply(
) = {

    new Rmprintreport(
)
  }

  private val rmprintreportParser: RowParser[Rmprintreport] = {
      SqlParser.get[Int](C.ID) ~
      SqlParser.get[Date](DBHelper.CREATED_AT) ~
      SqlParser.get[Date](DBHelper.UPDATED_AT) ~
      SqlParser.get[Int](DBHelper.UPDATED_BY) map {
      case
              id ~
 ~
              createdAt ~
              updatedAt ~
              updatedBy =>
        Rmprintreport(
              id,
,
              createdAt,
              updatedAt,
              updatedBy)
    }
  }

  def create(user: CompanyUser, rmprintreport: Rmprintreport): Rmprintreport = {
    save(user, rmprintreport, true)
  }

  def update(user: CompanyUser, rmprintreport: Rmprintreport): Rmprintreport = {
    save(user, rmprintreport, false)
  }

  private def save(user: CompanyUser, rmprintreport: Rmprintreport, isNew: Boolean): Rmprintreport = {
    def getFields = {
      List(

      )
    }
    def throwException = {
      throw new RuntimeException(s"Error when saving ${C.RMPRINTREPORT}")
    }

    DBHelper.saveEx(
      user,
      Register(
        C.RMPRINTREPORT,
        C.ID,
        rmprintreport.id,
        false,
        true,
        true,
        getFields),
      isNew,
      C.CODE
    ) match {
      case SaveResult(true, id) => load(user, id).getOrElse(throwException)
      case SaveResult(false, id) => throwException
    }
  }

  def load(user: CompanyUser, id: Int): Option[Rmprintreport] = {
    loadWhere(user, s"${C.ID} = {id}", 'id -> id)
  }

  def loadWhere(user: CompanyUser, where: String, args : scala.Tuple2[scala.Any, anorm.ParameterValue[_]]*) = {
    DB.withConnection(user.database.database) { implicit connection =>
      SQL(s"SELECT t1.*, t2.${C.FK_NAME} FROM ${C.RMPRINTREPORT} t1 INNER JOIN ${C.???} t2 ON t1.${C.FK_ID} = t2.${C.FK_ID} WHERE $where")
        .on(args: _*)
        .as(rmprintreportParser.singleOpt)
    }
  }

  def delete(user: CompanyUser, id: Int) = {
    DB.withConnection(user.database.database) { implicit connection =>
      try {
        SQL(s"DELETE FROM ${C.RMPRINTREPORT} WHERE ${C.ID} = {id}")
        .on('id -> id)
        .executeUpdate
      } catch {
        case NonFatal(e) => {
          Logger.error(s"can't delete a ${C.RMPRINTREPORT}. ${C.ID} id: $id. Error ${e.toString}")
          throw e
        }
      }
    }
  }

  def get(user: CompanyUser, id: Int): Rmprintreport = {
    load(user, id) match {
      case Some(p) => p
      case None => emptyRmprintreport
    }
  }
}


// Router

GET     /api/v1/general/rmprintreport/:id              controllers.logged.modules.general.Rmprintreports.get(id: Int)
POST    /api/v1/general/rmprintreport                  controllers.logged.modules.general.Rmprintreports.create
PUT     /api/v1/general/rmprintreport/:id              controllers.logged.modules.general.Rmprintreports.update(id: Int)
DELETE  /api/v1/general/rmprintreport/:id              controllers.logged.modules.general.Rmprintreports.delete(id: Int)




/**/
