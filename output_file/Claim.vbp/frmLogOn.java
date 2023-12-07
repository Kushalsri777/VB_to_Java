public class frmLogOn {

  //******************************************************************************
  // Module     : frmLogOn
  // Description:
  // Procedures :
  //    Private   cmdExitApplication_Click()
  //    Private   cmdOK_Click()
  //    Private   fnValidData() As Boolean
  //    Private   fnWarningData()
  //    Private   Form_Activate()
  //    Private   Form_KeyDown(ByRef pintKeyCode As Integer, ByRef pintShift As Integer)
  //    Private   Form_Load()
  //    Private   Form_Unload(ByRef pintCancel As Integer)
  //    Private   txtUserId_LostFocus()
  //
  // Modified   :
  // 03/03/02 BAW (Phase2A) Added support for new global error handler
  // 08/31/01 BAW (Phase2A) Added standardized error handlers
  // 09/25/00 JG  (Phase2A) Cleaned with Total Visual CodeTools 2000
  //
  // --------------------------------------------------
//Option Explicit
  *Option Compare Binary
  private String mstrScreenName = "";
  private static final Long MCLNGMINFORMWIDTH = 4965;
  private static final Long MCLNGMINFORMHEIGHT = 2715;

  private static final String MCSTRTXTUSERIDLABEL = "User ID";
  private static final String MCSTRTXTPASSWORDLABEL = "Password";
  private static final String MCSTRCBOENVIRONMENTLABEL = "Environment";

  // The following is used to determine whether the user has actually changed
  // the User ID TextBox. If not, then the LostFocus event should not revalidate.
  private String mstrSaveUserID = "";
  //'SQL_INTEGRATED_SECURITY
  private String mstrNetworkUserID = "";

  cautAuthenticate m_autAuthenticate = null;


  //SQL_INTEGRATED_SECURITY - Added
  //////////////////////////////////////////////////////////////////////////////////////////////////
  private void cboEnvironment_Click() {
    // Comments  : Make sure to default User ID to logged on network User ID if the
    //             selected Environment uses Integrated Security (i.e. Windows Authentication)
    // Parameters: N/A
    // --------------------------------------------------
    try {
      "cboEnvironment_Click"
.equals(Const cstrCurrentProc As String);

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      if (modGeneral.gapsApp.getUsesWindowsAuthentication(cboEnvironment.Text)) {
        txtUserId.Text = modWinApi.fnGetNetworkUser();
        txtPassword.Text = "";
        fnEnableDisableControl(txtUserId, false);
        fnEnableDisableControl(txtPassword, false);
      } 
      else {
        txtPassword.Text = "";
        fnEnableDisableControl(txtUserId, true);
        fnEnableDisableControl(txtPassword, true);
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
//SQL_INTEGRATED_SECURITY - Added



//////////////////////////////////////////////////////////////////////////////////////////////////
  private void cmdExitApplication_Click() {
    // Comments  : Exit Application button
    // Parameters: N/A
    // Modified  :
    //   BAW 09/10/2001 - Changed "End" to call a procedure terminates the app by unloading all forms
    // --------------------------------------------------
    try {
      "cmdExitApplication_Click"
.equals(Const cstrCurrentProc As String);

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      if (modGeneral.bDEBUGAPPTERMINATION) {
        Debug.Print("frmLogOn.cmdExitApplication_Click is calling fnTerminateTheApp...");
      }
      fnTerminateTheApp;
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
    // Comments  : OK button
    // Parameters: N/A
    // Modified  :
    //   09/10/2001 - Changed the behavior following an unsuccessful logon
    //                to call the fatal error handler, rather than just "End"
    // --------------------------------------------------
    "cmdOK_Click"
.equals(Const cstrCurrentProc As String);
    chrgHourglass hrgHourglass = null;
    try {

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      // Change mouse pointer into an hourglass
      hrgHourglass = new chrgHourglass();
      hrgHourglass.setValue(true);

      // Close the ADO Connection since the user is probably logging on to a different environment,
      // or under a different User ID. Update the status bar accordingly.
      frmMDIMain.frmMDIMain.fnUpdateStatusBar(strUserID:=vbNullString, strEnv:=vbNullString);
      modGeneral.gconAppActive.disconnect();

      if (fnValidData()) {
        // Connect to the selected Environment and put the desired App Role into effect.
        // .Connect( ) raises an error if it is unsuccessful, so if we get to the
        // subsequent statement it means "success".
        m_autAuthenticate.cautAuthenticate.authenticateUser(strEnvironIn:=cboEnvironment.Text, strUserIDIn:=txtUserId.Text, strPasswordIn:=txtPassword.Text, pconIn:=gconAppActive, bActiveDBIn:=True);

        //Debug.Print "This user was authenticated: " & txtUserId.Text & _
        //            " (password=" & txtPassword.Text & _
        //            " Environment=" & cboEnvironment.Text

        // Save logged-on User ID, so it can be shown the next time the user comes to this screen
        gapsApp.setLastLogOnUserID(txtUserId.Text);
        // Save logged-on User's Password, so it can reused -- but only within this session.
        gapsApp.setLastLogonPassword(txtPassword.Text);
        // Save Environment name
        gapsApp.setLastLogonEnvironment(cboEnvironment.Text);

        frmMDIMain.frmMDIMain.fnUpdateStatusBar(strUserID:=.capsAppSettings.getLastLogOnUserID(), strEnv:=.capsAppSettings.getLastLogonEnvironment());
        Unload(this);
      }
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Report the error, since this is an event handler
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.reportFatalError(mstrScreenName);
    }

    // Clean-up statements go here
    hrgHourglass.setValue(false);
    modGeneral.fnFreeObject(hrgHourglass);

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



// SQL_INTEGRATED_SECURITY - Added
//////////////////////////////////////////////////////////////////////////////////////////////////
  private void fnEnableDisableControl(Control ctlIn, boolean bEnable) {
    //--------------------------------------------------------------------------
    // Procedure:   fnEnableDisableControl
    // Description: Given a control, either make it look and behave Enabled or
    //              Disabled, depending on the bEnable parameter
    //
    // Params:      n/a
    //    ctlIn   (in) The control to enable/disable
    //    bEnable (in) True to enable the specified control; False to disable it.
    //
    // Returns:     N/A
    //-----------------------------------------------------------------------------
    "fnEnableDisableControl"
.equals(Const cstrCurrentProc As String);

    try {

      switch (bEnable) {
        case  True:
          // Next 3 lines commented out since this app doesn't use the VSFlexGrid control
          // If (TypeOf ctlIn Is VSFlexGrid) Then
          //     ' Do nothing
          // ElseIf (TypeOf ctlIn Is DTPicker) Then
          if ((TypeOf ctlIn Is DTPicker)) {
            ctlIn.TabStop = true;
            ctlIn.Enabled = true;
          } 
          else {
            ctlIn.Locked = false;
            ctlIn.TabStop = true;
            ctlIn.BackColor = vbWindowBackground;
            ctlIn.ForeColor = vbWindowText;
            ctlIn.Enabled = true;
          }
          break;

        case  False:
          // Next 3 lines commented out since this app doesn't use the VSFlexGrid control
          // If (TypeOf ctlIn Is VSFlexGrid) Then
          //     ' Do nothing
          // ElseIf (TypeOf ctlIn Is DTPicker) Then
          if ((TypeOf ctlIn Is DTPicker)) {
            ctlIn.TabStop = false;
            ctlIn.Enabled = false;
          } 
          else {
            ctlIn.Locked = true;
            ctlIn.TabStop = false;
            ctlIn.BackColor = vbButtonFace;
            ctlIn.ForeColor = vbButtonText;
            ctlIn.Enabled = false;
          }
          break;
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
// SQL_INTEGRATED_SECURITY - Added



//////////////////////////////////////////////////////////////////////////////////////////////////
  private boolean fnValidData() {
    boolean _rtn = false;
    // Comments  : Determines if all data is valid, including
    //             whether all required fields have been input.
    //             If a data error is found, it returns False
    //             which directs the caller to stop processing.
    //             It also generates warnings, by calling
    //             WarningData(), but only if no errors were
    //             found up to that point.
    // Parameters: N/A
    //
    // Called By : cmdOK_Click() in frmLogOn
    //
    // Returns   : True if all data is valid; False otherwise
    // Modified  :
    // --------------------------------------------------
    try {
      "fnValidData"
.equals(Const cstrCurrentProc As String);

      boolean bErrorFound = false;
      Control ctlFirstToFail = null;
      int intFailures = 0;
      String strFieldList = "";

      // Check the fields in a left-to-right, top-to-bottom screen sequence.
      //     1. User ID
      //     2. Password
      //     3. Environment
      //
      // ------------- 1.  Verify required fields are missing --------------

      // Verify fields are necessary to connect to requested Environment

      //SQL_INTEGRATED_SECURITY
      // Only validate User ID and Password if Windows Authentication is NOT used.
      // (If it IS used, we get it from how the user is logged on to the network;
      // disregarding and not allowing them to input the User ID and Password
      // on this screen.
      if (!modGeneral.gapsApp.getUsesWindowsAuthentication(cboEnvironment)) {
        if (txtUserId.Text == null || txtUserId.Text.length() == 0) {
          if (intFailures == 0) {
            strFieldList = "\\r\\n"+ MCSTRTXTUSERIDLABEL;
            ctlFirstToFail = txtUserId;
          } 
          else {
            strFieldList = strFieldList+ "\\r\\n"+ MCSTRTXTUSERIDLABEL;
          }
          intFailures = intFailures + 1;
        }

        //SQL_INTEGRATED_SECURITY ' Need to do Environment next, even though it's not the next consecutive field
        //SQL_INTEGRATED_SECURITY ' on the screen, since its value determines whether Password will be checked
        //SQL_INTEGRATED_SECURITY If IsNull(cboEnvironment.Text) Or Len(cboEnvironment.Text) = 0 Then
        //SQL_INTEGRATED_SECURITY     If intFailures = 0 Then
        //SQL_INTEGRATED_SECURITY         strFieldList = vbCrLf & mcstrCboEnvironmentLabel
        //SQL_INTEGRATED_SECURITY         Set ctlFirstToFail = cboEnvironment
        //SQL_INTEGRATED_SECURITY     Else
        //SQL_INTEGRATED_SECURITY         strFieldList = strFieldList & vbCrLf & mcstrCboEnvironmentLabel
        //SQL_INTEGRATED_SECURITY     End If
        //SQL_INTEGRATED_SECURITY     intFailures = intFailures + 1
        //SQL_INTEGRATED_SECURITY End If

        // Password is required in the SQL Server environment
        if (txtPassword.Text == null || txtPassword.Text.length() == 0) {
          if (intFailures == 0) {
            strFieldList = "\\r\\n"+ MCSTRTXTPASSWORDLABEL;
            ctlFirstToFail = txtPassword;
          } 
          else {
            strFieldList = strFieldList+ "\\r\\n"+ MCSTRTXTPASSWORDLABEL;
          }
          intFailures = intFailures + 1;
        }
      }
      //SQL_INTEGRATED_SECURITY

      if (intFailures != 0) {
        bErrorFound = true;
        _rtn = false;
        if (ctlFirstToFail.Visible) {
          ctlFirstToFail.SetFocus;
        }
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_REQD_FIELDS_MISSING, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc, strFieldList);
        // **TODO:** goto found: GoTo PROC_EXIT;
      }

      // If no errors found, continue with checking for warnings
      // NOTE: We won't get here if any errors were raised by preceding lines in Section 3.
      if (!bErrorFound) {
        _rtn = true;
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
  private void fnWarningData() {
    // Comments  : Validates fields, generating warnings if appropriate.
    //             It should NOT cause ValidData (this procedure's caller)
    //             to return False, since we want updates to proceed.
    // Parameters: N/A
    //
    // Called By : fnValidData() in frmMain
    //
    // Returns   : N/A
    // Modified  :
    // --------------------------------------------------
    "fnWarningData"
.equals(Const cstrCurrentProc As String);
    try {

      // ***   Currently there are no warnings  :(   ***

    } catch (Exception ex) {
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
    }
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
  private void form_Activate() {
    // Comments  :
    // Parameters: N/A
    // Modified  :
    //
    // --------------------------------------------------
    try {
      "Form_Activate"
.equals(Const cstrCurrentProc As String);

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      // SQL_INTEGRATED_SECURITY
      if (cboEnvironment.Visible) {
        cboEnvironment.SetFocus;
      }
      // SQL_INTEGRATED_SECURITY
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
  private void form_KeyDown(int pintKeyCode, int pintShift) { // TODO: Use of ByRef founded Private Sub Form_KeyDown(ByRef pintKeyCode As Integer, ByRef pintShift As Integer)
    // Comments  :
    // Parameters: pintKeyCode
    //             pintShift -
    // Modified  :
    //
    // --------------------------------------------------
    try {
      "Form_KeyDown"
.equals(Const cstrCurrentProc As String);

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      if (pintKeyCode == vbKeyEscape) {
        if (modGeneral.bDEBUGAPPTERMINATION) {
          Debug.Print(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc+ "  is calling fnTerminateTheApp...");
        }
        fnTerminateTheApp;
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
  private void form_Load() {
    // Comments  : For Phase2A, we can populate the Environment combo box based on the
    //             gapsApp.LoadCbo_EnvironmentNames( ) method, and then filter out
    //             Dev environments for unauthorized users (per a hard-coded list of User IDs).
    //
    //             For Phase2C, it will be based on what the Authenticate object
    //             determines to be valid environments for the user.
    // Parameters: N/A
    // Modified  :
    //
    // --------------------------------------------------
    try {
      "Form_Load"
.equals(Const cstrCurrentProc As String);

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      mstrScreenName = Me.Caption;
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      // Disable means of closing this form *other than* the OK or Exit Application buttons...such as
      // the Close button in the upper righthand corner of the screen and Alt-F4.
      modGeneral.fnRemoveCloseButton(this);

      Me.Icon = LoadResPicture(modResConstants.gCRES_ICON_MAINAPP, vbResIcon);

      // If the user has ever opened this form before, restore its size & placement.
      // If the restore would result in the form being off-screen, just center it instead.
      if (modGeneral.gapsApp.restoreForm(this) == false) {
        //*TODO:** can't found type for with block
        //*With this
        __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = this;
        w___TYPE_NOT_FOUND.Width = MCLNGMINFORMWIDTH;
        w___TYPE_NOT_FOUND.Height = MCLNGMINFORMHEIGHT;
        //' Not an MDI child, so it must be centered on the screen (not MDI parent)
        modGeneral.fnCenterFormOnScreen(this);
      }

      m_autAuthenticate = new cautAuthenticate();

      //SQL_INTEGRATED_SECURITY ' Initialize saved User ID to blank so initial txtUserId_LostFocus trigger will
      //SQL_INTEGRATED_SECURITY ' validate environments
      mstrNetworkUserID = modWinApi.fnGetNetworkUser();

      //SQL_INTEGRATED_SECURITY mstrSaveUserID = gcstrBlankEntry
      mstrSaveUserID = mstrNetworkUserID;

      // Initialize the User ID to the User ID under which the user logged on to the network.
      //'SQL_INTEGRATED_SECURITY
      txtUserId.Text = mstrNetworkUserID;

      // Load Environment combobox with list of all available SQL environments, regardless
      // of whether user is authorized for them.
      modGeneral.gapsApp.loadCbo_EnvironmentNames(cboEnvironment);
      //SQL_INTEGRATED_SECURITY ' The Environments combo box will appear with an empty entry until the user types
      //SQL_INTEGRATED_SECURITY ' in a User ID.
      //SQL_INTEGRATED_SECURITY cboEnvironment.Clear
      //SQL_INTEGRATED_SECURITY cboEnvironment.AddItem gcstrBlankEntry
      //' Select 1st entry as default selection
      cboEnvironment.ListIndex = 0;
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
    // Comments  :
    // Parameters: N/A
    // Modified  :
    //
    // --------------------------------------------------
    try {
      "Form_Unload"
.equals(Const cstrCurrentProc As String);

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      if (modGeneral.bDEBUGAPPTERMINATION) {
        Debug.Print("Entering "+ mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
      }

      modGeneral.gapsApp.saveForm(this);
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    modGeneral.fnFreeObject(frmLogOn);
    modGeneral.fnFreeObject(m_autAuthenticate);

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
  private void txtUserId_LostFocus() {
    // Comments  : Whenever the user leaves this field, see if the Environment combobox
    //             needs to be updated. In Phase2C, this will be done via a call to
    //             the Authenticate object.
    // Parameters: N/A
    // Modified  :
    //
    // --------------------------------------------------
    try {
      "txtUserId_LostFocus"
.equals(Const cstrCurrentProc As String);
      int intIndex = 0;
      //SQL_INTEGRATED_SECURITY Dim aEnvs()           As String

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      //SQL_INTEGRATED_SECURITY If txtUserId.Text <> mstrSaveUserID Then
      //SQL_INTEGRATED_SECURITY    ' Repopulate the Environment combo box
      //SQL_INTEGRATED_SECURITY    cboEnvironment.Clear
      //SQL_INTEGRATED_SECURITY    aEnvs = m_autAuthenticate.AuthenticateEnvironments(txtUserId.Text)
      //SQL_INTEGRATED_SECURITY    For intIndex = LBound(aEnvs) To UBound(aEnvs)
      //SQL_INTEGRATED_SECURITY        If Len(Trim$(aEnvs(intIndex))) <> 0 Then
      //SQL_INTEGRATED_SECURITY            'Debug.Print "This user was authenticated: " & txtUserId.Text & _
      //SQL_INTEGRATED_SECURITY            '            " Environment=" & aEnvs(intIndex)
      //SQL_INTEGRATED_SECURITY            cboEnvironment.AddItem aEnvs(intIndex)
      //SQL_INTEGRATED_SECURITY        End If
      //SQL_INTEGRATED_SECURITY    Next

      //SQL_INTEGRATED_SECURITY    ' If there are no environments for which the specified User ID is authorized,
      //SQL_INTEGRATED_SECURITY    ' then disable the OK button so the user can only click the Exit Application
      //SQL_INTEGRATED_SECURITY    ' button unless they specify a different User ID.
      if (cboEnvironment.ListCount < 1) {
        //SQL_INTEGRATED_SECURITY gerhApp.ReportNonFatal vbObjectError + gcRES_INFO_NO_AUTHENTICATED_ENVIRONMENTS, _
        //SQL_INTEGRATED_SECURITY                        mstrScreenName & gcstrDOT & cstrCurrentProc, _
        //SQL_INTEGRATED_SECURITY                        mcstrTxtUserIdLabel, mcstrCboEnvironmentLabel
        cmdOK.Enabled = false;
        cmdExitApplication.SetFocus;
      } 
      else {
        //SQL_INTEGRATED_SECURITY cboEnvironment.ListIndex = 0            ' Select 1st entry as default selection
        cmdOK.Enabled = true;
      }

      mstrSaveUserID = txtUserId.Text;
      //SQL_INTEGRATED_SECURITY End If
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


case class RmlogonData(
              id: Option[Int],

              )

object Rmlogons extends Controller with ProvidesUser {

  val rmlogonForm = Form(
    mapping(
      "id" -> optional(number),

  )(RmlogonData.apply)(RmlogonData.unapply))

  implicit val rmlogonWrites = new Writes[Rmlogon] {
    def writes(rmlogon: Rmlogon) = Json.obj(
      "id" -> Json.toJson(rmlogon.id),
      C.ID -> Json.toJson(rmlogon.id),

    )
  }

  def get(id: Int) = GetAction { implicit request =>
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.LIST_RMLOGON), { user =>
      Ok(Json.toJson(Rmlogon.get(user, id)))
    })
  }

  def update(id: Int) = PostAction { implicit request =>
    Logger.debug("in rmlogons.update")
    rmlogonForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      rmlogon => {
        Logger.debug(s"form: ${rmlogon.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.EDIT_RMLOGON), { user =>
          Ok(
            Json.toJson(
              Rmlogon.update(user,
                Rmlogon(
                       id,

                ))))
        })
      }
    )
  }

  def create = PostAction { implicit request =>
    Logger.debug("in rmlogons.create")
    rmlogonForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      rmlogon => {
        Logger.debug(s"form: ${rmlogon.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.NEW_RMLOGON), { user =>
          Ok(
            Json.toJson(
              Rmlogon.create(user,
                Rmlogon(

                ))))
        })
      }
    )
  }

  def delete(id: Int) = PostAction { implicit request =>
    Logger.debug("in rmlogons.delete")
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.DELETE_RMLOGON), { user =>
      Rmlogon.delete(user, id)
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

case class Rmlogon(
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

object Rmlogon {

  lazy val emptyRmlogon = Rmlogon(
)

  def apply(
      id: Int,
) = {

    new Rmlogon(
      id,
)
  }

  def apply(
) = {

    new Rmlogon(
)
  }

  private val rmlogonParser: RowParser[Rmlogon] = {
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
        Rmlogon(
              id,
,
              createdAt,
              updatedAt,
              updatedBy)
    }
  }

  def create(user: CompanyUser, rmlogon: Rmlogon): Rmlogon = {
    save(user, rmlogon, true)
  }

  def update(user: CompanyUser, rmlogon: Rmlogon): Rmlogon = {
    save(user, rmlogon, false)
  }

  private def save(user: CompanyUser, rmlogon: Rmlogon, isNew: Boolean): Rmlogon = {
    def getFields = {
      List(

      )
    }
    def throwException = {
      throw new RuntimeException(s"Error when saving ${C.RMLOGON}")
    }

    DBHelper.saveEx(
      user,
      Register(
        C.RMLOGON,
        C.ID,
        rmlogon.id,
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

  def load(user: CompanyUser, id: Int): Option[Rmlogon] = {
    loadWhere(user, s"${C.ID} = {id}", 'id -> id)
  }

  def loadWhere(user: CompanyUser, where: String, args : scala.Tuple2[scala.Any, anorm.ParameterValue[_]]*) = {
    DB.withConnection(user.database.database) { implicit connection =>
      SQL(s"SELECT t1.*, t2.${C.FK_NAME} FROM ${C.RMLOGON} t1 INNER JOIN ${C.???} t2 ON t1.${C.FK_ID} = t2.${C.FK_ID} WHERE $where")
        .on(args: _*)
        .as(rmlogonParser.singleOpt)
    }
  }

  def delete(user: CompanyUser, id: Int) = {
    DB.withConnection(user.database.database) { implicit connection =>
      try {
        SQL(s"DELETE FROM ${C.RMLOGON} WHERE ${C.ID} = {id}")
        .on('id -> id)
        .executeUpdate
      } catch {
        case NonFatal(e) => {
          Logger.error(s"can't delete a ${C.RMLOGON}. ${C.ID} id: $id. Error ${e.toString}")
          throw e
        }
      }
    }
  }

  def get(user: CompanyUser, id: Int): Rmlogon = {
    load(user, id) match {
      case Some(p) => p
      case None => emptyRmlogon
    }
  }
}


// Router

GET     /api/v1/general/rmlogon/:id              controllers.logged.modules.general.Rmlogons.get(id: Int)
POST    /api/v1/general/rmlogon                  controllers.logged.modules.general.Rmlogons.create
PUT     /api/v1/general/rmlogon/:id              controllers.logged.modules.general.Rmlogons.update(id: Int)
DELETE  /api/v1/general/rmlogon/:id              controllers.logged.modules.general.Rmlogons.delete(id: Int)




/**/
