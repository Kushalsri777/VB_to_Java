public class frmMDIMain {

  //******************************************************************************
  // Module     : frmMDIMain
  // Description:
  // Procedures:
  //              fnGetNbrOfDataIntegrityIssues() As Long
  //              fnOpenURLInBrowser(ByVal strURL As String) As Boolean
  //              fnShowInsuredForm()
  //              MDIForm_Activate()
  //              MDIForm_Load()
  //              MDIForm_Resize()
  //              mnuFileExit_Click()
  //              mnuHelpAbout_Click()
  //              mnuReportsGenerateCheckFreeFile_Click()
  //              mnuReportsPrintReport_Click()
  //              mnuViewInsured_Click()
  //              mnuWindow_Click

  // Modified   :
  // 10/25/01 BAW Added Help | Technical Support menu option
  // 03/26/01 BAW Cleaned with Total Visual CodeTools 2000
  // 01/2002  BAW Optimized per Project Analyzer (removing dead code, adding "$" to Mid/Space, etc.)
  //              Also added extra DoEvents in Form_Load to make the Splash screen's progress bar
  //              move more smoothly.
  // --------------------------------------------------
//Option Explicit
  *Option Compare Binary
  private String mstrScreenName = "";
  private static final Long MCLNGMINFORMWIDTH = 14625;
  private static final Long MCLNGMINFORMHEIGHT = 10260;

  // The ShellExecute API is used by mnuHelpTechnicalSupport
*TODO: API Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long



  //////////////////////////////////////////////////////////////////////////////////////////////////
  private int fnGetNbrOfDataIntegrityIssues() {
    int _rtn = 0;
    // Description: Executes a view to identify the number of potential
    //              data integrity issues in the database
    //              WARNING: Both frmPrintReport.fnGetData_DataIntegrityIssuesReport()
    //                       and frmMDIMain.fnGetNbrOfDataIntegrityIssues() run
    //                       this view!!!
    //
    // Parameters: N/A
    //
    // Called by :
    //
    // Modified  :
    //
    // --------------------------------------------------
    try {
      "fnGetNbrOfDataIntegrityIssues"
.equals(Const cstrCurrentProc As String);
      "dbo.DataIntegrityIssuesReport_v"
.equals(Const cstrSQLView As String);
      String strSQL = "";
      DBRecordSet rstTemp = null;

  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // If the user hasn't logged on yet, just return 0.
    if (modGeneral.gconAppActive.getADOConn().State != adStateClosed) {
      strSQL = "SELECT * from "+ cstrSQLView+ " WHERE calcReason <> ''";

      rstTemp = modGeneral.gadwApp.execute_SQL_AsRST(modGeneral.gconAppActive, strSQL);
      #If DEBUG_RST Then;
      Debug.Print("In "+ cstrCurrentProc+ ", "+ CStr(rstTemp.RecordCount)+ " records were retrieved in the rst.");
      Debug.Print("SQL statement is: "+ "\\n"+ strSQL);
      #End If;

      // Disconnect the recordset
      rstTemp.ActiveConnection = null;

      _rtn = rstTemp.RecordCount;
    }
    // **TODO:** label found: PROC_EXIT:;
//' Disable error handler
}
//*TODO:** the error label PROC_ERR: couldn't be found
  try {

  // Clean-up statements go here
  modGeneral.fnFreeRecordset(rstTemp);

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



//////////////////////////////////////////////////////////////////////////////////////////////////
  private boolean fnOpenURLInBrowser(String strURL) {
    boolean _rtn = false;
    // Comments  : Opens the default browser on the specified URL
    // Parameters: strURL - the URL to display in the browser window
    // Called by : mnuHelpTechnicalSupport_Click() of frmMDIMain
    // Modified  :
    //
    // --------------------------------------------------
    try {
      "fnOpenURLInBrowser"
.equals(Const cstrCurrentProc As String);

      int lngReturnCode = 0;

      // Make sure the URL is prefixed with http:// or https://
      if (!(strURL.toUpperCase().indexOf("HTTP", 1).equals(1))) {
        strURL = "http://"+ strURL;
      }

      lngReturnCode = ShellExecute(0&, "open", strURL, "", "", vbNormalFocus);

      _rtn = (lngReturnCode > 32);
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
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
  public void fnShowInsuredForm() {
    // Comments  : This is a PUBLIC procedure that opens a new instance of the Insured screen
    // Called by : frmSplash.tmrTimer_Timer( ) and mnuViewInsured_Click( )
    // Parameters: None
    // Modified  :
    // --------------------------------------------------
    try {
      "fnShowInsuredForm"
.equals(Const cstrCurrentProc As String);
      frmInsured frm = null;

      if (!(modGeneral.fnIsFormLoaded("frmInsured", frm))) {
        frm = new frmInsured();
      }

      frm.Show;
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    modGeneral.fnFreeObject(frm);

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



// This routine exists only to support programmer testing of new table wrappers.
//////////////////////////////////////////////////////////////////////////////////////////////////
  public void fnTestTableWrapper() {
    // Comments  : Opens the Test Table Wrapper screen
    //
    //             WARNING:  This code should be REM'd out once the table
    //                       wrappers have been fully tested!
    // Parameters: None
    // Modified  :
    // --------------------------------------------------
    try {
      "fnTestTableWrapper"
.equals(Const cstrCurrentProc As String);
      frmTestTableWrapper frm = null;

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      if (!(modGeneral.fnIsFormLoaded("frmTestTableWrapper", frm))) {
        frm = new frmTestTableWrapper();
      }

      frm.Show;
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    modGeneral.fnFreeObject(frm);

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



////////////////////////////////////////////////////////////////////////////////////////////
  public void fnUpdateStatusBar(String strUserID, String strEnv) {
    // Comments  : Updates the Status Bar so it reflects the
    //             current information per the Log On screen
    //
    // Parameters:
    //             strUserID - the User ID of the logged on user, if any (as specified on the Log On screen)
    //             strEnv - the Environment to which the user logged on, if any (as selected on the Log On screen)
    //
    // Called by : frmMDIForm_Load()
    //             This is also called by frmLogOn's cmdOK_Click event.
    // Parameters: None
    // Modified  :
    // --------------------------------------------------
    try {

      "fnUpdateStatusBar"
.equals(Const cstrCurrentProc As String);
      "User ID: "
.equals(Const cstrLabelPanel1 As String);
      "Environment: "
.equals(Const cstrLabelPanel2 As String);
      "# of Possible Errors in the DB: "
.equals(Const cstrLabelPanel3 As String);

      Me.sbrStatusBar.Panels(1).Text = cstrLabelPanel1+ strUserID;
      Me.sbrStatusBar.Panels(2).Text = cstrLabelPanel2+ strEnv;

      // If the User ID and Environment are empty, then the user must not be logged on, so don't
      // attempt to open the view.
      if (strUserID.equals("") && strEnv.equals("")) {
        Me.sbrStatusBar.Panels(3).Text = cstrLabelPanel3+ "?";
      } 
      else {
        Me.sbrStatusBar.Panels(3).Text = cstrLabelPanel3+ CStr(fnGetNbrOfDataIntegrityIssues());
      }
      DoEvents;
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
  private void mDIForm_Activate() {
    // Comments  : Refreshes the Status Bar text to show
    //             if there are errors
    // Parameters: None
    // Modified  :
    // --------------------------------------------------
    try {
      "MDIForm_Activate"
.equals(Const cstrCurrentProc As String);

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);


      frmMDIMain.fnUpdateStatusBar(strUserID:=.capsAppSettings.getLastLogOnUserID(), strEnv:=.capsAppSettings.getLastLogonEnvironment());
      DoEvents;
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
  private void mDIForm_Load() {
    // Comments  : Opens a ADO Connection object and sets the
    //             Status Bar text to show if there are errors
    // Parameters: None
    // Modified  :
    // --------------------------------------------------
    try {
      "MDIForm_Load"
.equals(Const cstrCurrentProc As String);
      Panel pnlAdd = null;

      // Set the screen name that will be used to form the Title on message boxes
      mstrScreenName = Me.Caption;

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      // Open the Claimslog.log log file to track application events during this session.
      modAppLog.fnLogOpen();
      DoEvents;

      Me.Icon = LoadResPicture(modResConstants.gCRES_ICON_MAINAPP, vbResIcon);

      // If the user has ever opened this form before, restore its size & placement.
      // If the restore would result in the form being off-screen, just center it instead.
      if (modGeneral.gapsApp.restoreForm(this) == false) {
        //*TODO:** can't found type for with block
        //*With this
        __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = this;
        w___TYPE_NOT_FOUND.Width = MCLNGMINFORMWIDTH;
        w___TYPE_NOT_FOUND.Height = MCLNGMINFORMHEIGHT;
        modGeneral.fnCenterFormOnScreen(this);
      }

      // Set properties of Status Bar, defining 3 panels:
      //   1. User ID: <acf2>
      //   2. Environment: <environment name from cconClaimsActive>
      //   3. # of Possible Errors in the DB: <##>
      //
      // Make the status bar panels proportionate to the width of the form,
      // with the left panel being much bigger since it contains the fully
      // qualified path to the database.
      // PANEL 1
      Me.sbrStatusBar.Panels(1).AutoSize = sbrContents;
      Me.sbrStatusBar.Panels(1).Alignment = sbrLeft;
      Me.sbrStatusBar.Panels(1).MinWidth = Me.Width * 0.33;
      // PANEL 2
      pnlAdd = Me.sbrStatusBar.Panels.Add(Index:=2, Style:=sbrText);
      pnlAdd.AutoSize = sbrContents;
      pnlAdd.Alignment = sbrLeft;
      pnlAdd.MinWidth = Me.Width * 0.33;
      // PANEL 3
      pnlAdd = Me.sbrStatusBar.Panels.Add(Index:=3, Style:=sbrText);
      pnlAdd.AutoSize = sbrContents;
      pnlAdd.Alignment = sbrLeft;
      pnlAdd.MinWidth = Me.Width * 0.34;

      // Initialize status bar text in all panels so it looks okay if the MDIForm
      // is displayed while it calls the frmLogOn form.
      fnUpdateStatusBar(strUserID:=vbNullString, strEnv:=vbNullString);
      DoEvents;
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    modGeneral.fnFreeObject(pnlAdd);

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
  private void mDIForm_QueryUnload(int cancel, int unloadMode) {
    // Comments  : Inhibit closing the app if there are child forms open.
    // Parameters: None
    // Modified  :
    // --------------------------------------------------
    try {
      "MDIForm_QueryUnload"
.equals(Const cstrCurrentProc As String);

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      if (modGeneral.bDEBUGAPPTERMINATION) {
        Debug.Print("Entering "+ mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
      }

      if (modGeneral.gbAmProcessingAnAppFatalError) {
        // ALWAYS let the form be unloaded, with no prompts to the user, if shutting
        // down the app due to an application fatal error having been hit.
        // **TODO:** goto found: GoTo PROC_EXIT;
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
  private void mDIForm_Resize() {
    // Comments  : Don't let the MDI form be resized such that it
    //             is too small to fit the largest form.
    // Parameters: None
    // Modified  :
    // --------------------------------------------------
    try {
      "MDIForm_Resize"
.equals(Const cstrCurrentProc As String);

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      if (Me.WindowState == vbNormal) {
        // Bypass if vbMinimized or vbMaximized, to avoid run-time error 384
        // which says" "a form can't be moved or sized while minimized or maximized"
        if (Me.Height < MCLNGMINFORMHEIGHT) {
          Me.Height = MCLNGMINFORMHEIGHT;
        }
        if (Me.Width < MCLNGMINFORMWIDTH) {
          Me.Width = MCLNGMINFORMWIDTH;
        }
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
  private void mDIForm_Unload(int pintCancel) { // TODO: Use of ByRef founded Private Sub MDIForm_Unload(ByRef pintCancel As Integer)
    // Comments  : Close the log file and exit the application
    // Parameters: None
    // Modified  :
    // --------------------------------------------------
    try {
      "MDIForm_Unload"
.equals(Const cstrCurrentProc As String);
      //Dim frm As Form

      if (modGeneral.bDEBUGAPPTERMINATION) {
        Debug.Print("Now in "+ mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
      }

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      //!TODO! Put the following code in fnTerminateTheApp ???
      modAppLog.fnLogWrite("Exiting application.", cstrCurrentProc);
      modAppLog.fnLogClose();
      //!TODO! End

      // The following "If" check is needed in case of a fatal error being hit upon app start,
      // such that the capsAppSettings object (gapsApp) didn't make it through its Class_Initialize
      // event. Without the IF, a VB error 91 (Object variable or With block not set) is reported.
      if (!(modGeneral.gapsApp == null)) {
        modGeneral.gapsApp.saveForm(this);
      }

      // Can't do "Set frmMDIMain = Nothing"...causes a crash when fnTerminateTheApp tries to
      // deallocate the global objects
      //Unload Me

      Debug.Print(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc+ " is calling fnTerminateTheApp...");
      fnTerminateTheApp;


      Debug.Print(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc+ " is calling fnDeallocateGlobalObjects...");
      modGeneral.fnDeallocateGlobalObjects();
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
  private void mnuFile_Click() {
    // Comments  :
    // Parameters: N/A
    // Modified  :
    //
    // --------------------------------------------------
    try {
      "mnuFile_Click"
.equals(Const cstrCurrentProc As String);

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      //*TODO:** can't found type for with block
      //*With Forms
      __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = Forms;
      // Enable the File | Log On option only if the MDIMain form is the sole open form
      mnuFileLogon.Enabled = (w___TYPE_NOT_FOUND.Count == 1);

      mnuFileExit.Enabled = true;
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
  private void mnuFileExit_Click() {
    // Comments  : Terminates the app. Note that this menu option should be disabled
    //             if any forms besides the MDI form are open.
    // Parameters: None
    // Modified  :
    // --------------------------------------------------
    try {
      "mnuFileExit_Click"
.equals(Const cstrCurrentProc As String);
      Form frm = null;

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      if (modGeneral.bDEBUGAPPTERMINATION) {
        Debug.Print(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc+ " is calling fnTerminateTheApp");
      }
      //'   !!!!!!!!!!!!!!!!!!!!!!!!
      fnTerminateTheApp;
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    modGeneral.fnFreeObject(frm);

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
  private void mnuFileLogOn_Click() {
    // Comments  : Note that this menu option should not be available unless
    //             all forms besides the MDI form are closed.
    // Parameters: N/A
    // Modified  :
    //
    // --------------------------------------------------
    try {
      "mnuFileLogOn_Click"
.equals(Const cstrCurrentProc As String);
      frmLogOn frm = null;

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      // If the user has an ADO connection currently open, then
      // terminate the connection before proceeding. This ensures the connection
      // will always represent the logged-on user.
      modGeneral.gconAppActive.disconnect();

      if (!(modGeneral.fnIsFormLoaded("frmLogOn", frm))) {
        frm = new frmLogOn();
      }

      frm.Show(vbModal);
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    modGeneral.fnFreeObject(frm);

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
  private void mnuFilePrintSetup_Click() {
    // Comments  :
    // Parameters: N/A
    // Modified  :
    //
    // --------------------------------------------------
    try {
      "mnuFilePrintSetup_Click"
.equals(Const cstrCurrentProc As String);

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      //*TODO:** can't found type for with block
      //*With cdlCommonDialog
      __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = cdlCommonDialog;
      w___TYPE_NOT_FOUND.PrinterDefault = true;
      w___TYPE_NOT_FOUND.Flags = cdlPDPrintSetup;
      w___TYPE_NOT_FOUND.ShowPrinter;
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
  private void mnuHelpAboutClaimsInterest_Click() {
    // Comments  : Displays the splash screen as a Help | About box
    // Parameters: None
    // Modified  :
    // --------------------------------------------------
    try {
      "mnuHelpAboutClaimsInterest_Click"
.equals(Const cstrCurrentProc As String);
      frmSplash frm = null;

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      //!TODO! frm is not initialized if Else condition is hit!
      if (!(modGeneral.fnIsFormLoaded("frmSplash", frm))) {
        frm = new frmSplash();
        frm.fnShowAsAboutBox();
      } 
      else {
        frm.fnShowAsAboutBox();
      }
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    modGeneral.fnFreeObject(frm);

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
  private void mnuHelpTechnicalSupport_Click() {
    // Comments  : Display the Claims Interest Technical Support page on The Source
    // Parameters:  -
    // Modified  :
    //
    // --------------------------------------------------
    try {
      "mnuHelpTechnicalSupport_Click"
.equals(Const cstrCurrentProc As String);
      "http://intranet/showcontext.cfm?context=1826"
.equals(Const cstrURL As String);

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      if (!(fnOpenURLInBrowser(cstrURL))) {
        modGeneral.gerhApp.reportNonFatal(vbObjectError + modResConstants.gCRES_INFO_CANT_LAUNCH_URL, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc, cstrURL);
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


// The creation of maintenance screens for the Current Rate and State Rule tables
// is deferred in the v2.4 release. As such, the Tools menu bites the dust...for now.
  *#If False Then
  //////////////////////////////////////////////////////////////////////////////////////////////////
  private void mnuTools_Click() {
    // Comments  : Opens the Current Rate screen as a modal window
    // Parameters: None
    // Modified  :
    // --------------------------------------------------
    try {
      "mnuTools_Click"
.equals(Const cstrCurrentProc As String);

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      if (modGeneral.gconAppActive.getLastLogonIsSpecialUser()) {
        mnuToolsCurrentRate.Enabled = true;
        mnuToolsStateRule.Enabled = true;
      } 
      else {
        mnuToolsCurrentRate.Enabled = false;
        mnuToolsStateRule.Enabled = false;
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
  private void mnuToolsCurrentRate_Click() {
    // Comments  : Opens the Current Rate screen as a modal window
    // Parameters: None
    // Modified  :
    // --------------------------------------------------
    try {
      "mnuToolsCurrentRate_Click"
.equals(Const cstrCurrentProc As String);
      int intResponse = 0;
      frmCurrentRate frm = null;

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      if (!(modGeneral.fnIsFormLoaded("frmCurrentRate", frm))) {
        frm = new frmCurrentRate();
      }

      frm.Show(vbModal);
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
  private void mnuToolsStateRule_Click() {
    // Comments  : Opens the State Rules screen as a modal window
    // Parameters: None
    // Modified  :
    // --------------------------------------------------
    try {
      "mnuToolsStateRules_Click"
.equals(Const cstrCurrentProc As String);
      int intResponse = 0;
      frmStateRule frm = null;

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      if (!(modGeneral.fnIsFormLoaded("frmStateRule", frm))) {
        frm = new frmStateRule();
      }

      frm.Show(vbModal);
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
  private void mnuHelpViewApplicationLogFile_Click() {
    // Comments  : Show the splash screen as an About box.
    // Parameters: N/A
    // Modified  :
    //
    // --------------------------------------------------
    try {
      "mnuHelpViewApplicationLogFile_Click"
.equals(Const cstrCurrentProc As String);
      String strLogFileNm = "";
      String strLogFileExt = "";

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      strLogFileNm = modAppLog.fnGetAppLogFileFQ();
      strLogFileExt = modGeneral.fnGetExtPart(strLogFileNm);

      if (!modGeneral.fnOpenFileInDefaultApp(strLogFileNm)) {
        // gcRES_INFO_CANT_OPEN_FILE (1014)
        // Unable to open @@1. The file either does not exist or no application is associated with files of type @@2.
        modGeneral.gerhApp.reportNonFatal(vbObjectError + modResConstants.gCRES_INFO_CANT_OPEN_FILE, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc, "the application log file ("+ strLogFileNm+ ")", strLogFileExt.toUpperCase());
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
  private void mnuReportsGenerateTaxFile_Click() {
    // Comments  : Opens the Generate Tax File screen
    // Parameters: None
    // Modified  :
    // --------------------------------------------------
    try {
      "mnuReportsGenerateTaxFile_Click"
.equals(Const cstrCurrentProc As String);
      frmGenerateTaxFile frm = null;

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      if (!(modGeneral.fnIsFormLoaded("frmGenerateTaxFile", frm))) {
        frm = new frmGenerateTaxFile();
      }

      frm.Show;
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    modGeneral.fnFreeObject(frm);

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
  private void mnuReportsPrintReport_Click() {
    // Comments  : Opens the Print Report screen
    // Parameters: None
    // Modified  :
    // --------------------------------------------------
    try {
      "mnuReportsPrintReport_Click"
.equals(Const cstrCurrentProc As String);
      frmPrintReport frm = null;

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      if (!(modGeneral.fnIsFormLoaded("frmPrintReport", frm))) {
        frm = new frmPrintReport();
      }

      frm.Show;
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    modGeneral.fnFreeObject(frm);

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
  private void mnuViewInsured_Click() {
    // Comments  : Opens a new instance of the Insured screen
    // Parameters: None
    // Modified  :
    // --------------------------------------------------
    try {
      "mnuViewInsured_Click"
.equals(Const cstrCurrentProc As String);

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      fnShowInsuredForm();
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
//   Don't need a  mnuWindow_Click() event handler; all menu options are enabled as of 2A
//////////////////////////////////////////////////////////////////////////////////////////////////



//////////////////////////////////////////////////////////////////////////////////////////////////
  private void mnuWindowArrange_Click(int pintIndex) { // TODO: Use of ByRef founded Private Sub mnuWindowArrange_Click(ByRef pintIndex As Integer)
    // Comments  : The "arrangement" items in the Window menu
    //             are a control array, all controlled by
    //             this control event.
    // Parameters: pintIndex (in) - indicates which Window menu option was selected
    // Modified  :
    //
    // --------------------------------------------------
    try {
      "mnuWindowArrange_Click"
.equals(Const cstrCurrentProc As String);

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      switch (pintIndex) {
        case  0:
          Me.Arrange(vbCascade);
          break;

        case  1:
          Me.Arrange(vbTileHorizontal);
          break;

        case  2:
          Me.Arrange(vbTileVertical);
          break;

        case  3:
          Me.Arrange(vbArrangeIcons);
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


case class RmmdimainData(
              id: Option[Int],

              )

object Rmmdimains extends Controller with ProvidesUser {

  val rmmdimainForm = Form(
    mapping(
      "id" -> optional(number),

  )(RmmdimainData.apply)(RmmdimainData.unapply))

  implicit val rmmdimainWrites = new Writes[Rmmdimain] {
    def writes(rmmdimain: Rmmdimain) = Json.obj(
      "id" -> Json.toJson(rmmdimain.id),
      C.ID -> Json.toJson(rmmdimain.id),

    )
  }

  def get(id: Int) = GetAction { implicit request =>
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.LIST_RMMDIMAIN), { user =>
      Ok(Json.toJson(Rmmdimain.get(user, id)))
    })
  }

  def update(id: Int) = PostAction { implicit request =>
    Logger.debug("in rmmdimains.update")
    rmmdimainForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      rmmdimain => {
        Logger.debug(s"form: ${rmmdimain.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.EDIT_RMMDIMAIN), { user =>
          Ok(
            Json.toJson(
              Rmmdimain.update(user,
                Rmmdimain(
                       id,

                ))))
        })
      }
    )
  }

  def create = PostAction { implicit request =>
    Logger.debug("in rmmdimains.create")
    rmmdimainForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      rmmdimain => {
        Logger.debug(s"form: ${rmmdimain.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.NEW_RMMDIMAIN), { user =>
          Ok(
            Json.toJson(
              Rmmdimain.create(user,
                Rmmdimain(

                ))))
        })
      }
    )
  }

  def delete(id: Int) = PostAction { implicit request =>
    Logger.debug("in rmmdimains.delete")
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.DELETE_RMMDIMAIN), { user =>
      Rmmdimain.delete(user, id)
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

case class Rmmdimain(
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

object Rmmdimain {

  lazy val emptyRmmdimain = Rmmdimain(
)

  def apply(
      id: Int,
) = {

    new Rmmdimain(
      id,
)
  }

  def apply(
) = {

    new Rmmdimain(
)
  }

  private val rmmdimainParser: RowParser[Rmmdimain] = {
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
        Rmmdimain(
              id,
,
              createdAt,
              updatedAt,
              updatedBy)
    }
  }

  def create(user: CompanyUser, rmmdimain: Rmmdimain): Rmmdimain = {
    save(user, rmmdimain, true)
  }

  def update(user: CompanyUser, rmmdimain: Rmmdimain): Rmmdimain = {
    save(user, rmmdimain, false)
  }

  private def save(user: CompanyUser, rmmdimain: Rmmdimain, isNew: Boolean): Rmmdimain = {
    def getFields = {
      List(

      )
    }
    def throwException = {
      throw new RuntimeException(s"Error when saving ${C.RMMDIMAIN}")
    }

    DBHelper.saveEx(
      user,
      Register(
        C.RMMDIMAIN,
        C.ID,
        rmmdimain.id,
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

  def load(user: CompanyUser, id: Int): Option[Rmmdimain] = {
    loadWhere(user, s"${C.ID} = {id}", 'id -> id)
  }

  def loadWhere(user: CompanyUser, where: String, args : scala.Tuple2[scala.Any, anorm.ParameterValue[_]]*) = {
    DB.withConnection(user.database.database) { implicit connection =>
      SQL(s"SELECT t1.*, t2.${C.FK_NAME} FROM ${C.RMMDIMAIN} t1 INNER JOIN ${C.???} t2 ON t1.${C.FK_ID} = t2.${C.FK_ID} WHERE $where")
        .on(args: _*)
        .as(rmmdimainParser.singleOpt)
    }
  }

  def delete(user: CompanyUser, id: Int) = {
    DB.withConnection(user.database.database) { implicit connection =>
      try {
        SQL(s"DELETE FROM ${C.RMMDIMAIN} WHERE ${C.ID} = {id}")
        .on('id -> id)
        .executeUpdate
      } catch {
        case NonFatal(e) => {
          Logger.error(s"can't delete a ${C.RMMDIMAIN}. ${C.ID} id: $id. Error ${e.toString}")
          throw e
        }
      }
    }
  }

  def get(user: CompanyUser, id: Int): Rmmdimain = {
    load(user, id) match {
      case Some(p) => p
      case None => emptyRmmdimain
    }
  }
}


// Router

GET     /api/v1/general/rmmdimain/:id              controllers.logged.modules.general.Rmmdimains.get(id: Int)
POST    /api/v1/general/rmmdimain                  controllers.logged.modules.general.Rmmdimains.create
PUT     /api/v1/general/rmmdimain/:id              controllers.logged.modules.general.Rmmdimains.update(id: Int)
DELETE  /api/v1/general/rmmdimain/:id              controllers.logged.modules.general.Rmmdimains.delete(id: Int)




/**/
