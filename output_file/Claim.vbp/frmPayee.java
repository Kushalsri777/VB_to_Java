
import java.util.Date;

public class frmPayee {

  //!TODO! Add override capability
  //******************************************************************************
  // Module     : frmPayee
  // Description:
  // Procedures :
  //              cboCalcStCd_Click()
  //              cboPayeSsnTinTypCd_Click()
  //              cboPayeStCd_Click()
  //              chkPayeDfltOvrdInd_Click()
  //              cmdAdd_Click()
  //              cmdCloneThisPayee_Click()
  //              cmdClose_Click()
  //              cmdDelete_Click()
  //              cmdNavigate_Click(ByRef pintIndex As Integer)
  //              cmdUpdate_Click()
  //              dtpPayePmtDt_Change()
  //              fnAddRecord()
  //              fnBindControlsToTableWrapper()
  //              fnCalcAndLogMSIInfo(ByRef msiIn As StateInfo, strDesc As String)
  //              fnCalcClaimForState(ByRef msiStateIn As StateInfo) As Boolean
  //              fnCalcClaimInterest() As Boolean
  //              fnClearControls()
  //              fnGetCurrentIntRate(ByVal dtePayePmtDt As Date) As Double
  //              fnGetFieldLabel(ByVal strControlName As String) As String
  //              fnGetInterestRate(ByRef siRatesIn As StateInfo) As Currency
  //              fnGetListOfStates() As ADODB.Recordset
  //              fnGetStateInfo_InsdDthResStCd()
  //              fnGetStateInfo_IssStCd()
  //              fnGetStateInfo_Override()
  //              fnGetStateInfo_PayeStCd()
  //              fnInitializeCalcInfo(ByRef msiIn As StateInfo)
  //              fnInitializeEditMode()
  //              fnLoadCboPayeSsnTinTypCd()
  //              fnLoadCbosForStates()
  //              fnLoadControls()
  //              fnLoadLpcLookup()
  //              fnPerformLookup(ByRef lpcIn As LPLib.fpCombo)
  //              fnPromptForRate(ByVal strPromptText As String) As Double
  //              fnRefreshAllCombos()
  //              fnResetStateRules()
  //              fnSetAvailabilityOfControls(Optional ByVal bChangeFocus = True)
  //              fnSetCommandButtons(ByVal bEnable As Boolean)
  //              fnSetDefaultControlProperties()
  //              fnSetFocusToFirstUpdateableField()
  //              fnSetNavigationButtons(Optional ByVal bUnconditionalDisable As Boolean = False)
  //              fnSetupScreenControls()
  //              fnValidData() As Boolean
  //              fnWarningData()
  //              Form_Activate()
  //              Form_Initialize()
  //              Form_Load()
  //              Form_QueryUnload(ByRef pintCancel As Integer, ByRef pintUnloadMode As Integer)
  //              Form_Unload(ByRef pintCancel As Integer)
  //              ipcPayeDthbPmtAmt_Change()
  //              ipdPayeClmIntRt_Change()
  //              ipdPayeWthldRt_Change()
  //              ipmPayeSsnTinNum_Change()
  //              iptPayeAddrLn1Txt_Change()
  //              iptPayeAddrLn2Txt_Change()
  //              iptPayeCareOfTxt_Change()
  //              iptPayeCityNmTxt_Change()
  //              iptPayeFullNm_Change()
  //              iptPayeZipCd_Change()
  //              lpcLookupName_Click()
  //              lpcLookupName_GotFocus()
  //              lpcLookupName_KeyDown(ByRef intKeyCode As Integer, ByRef intShift As Integer)
  //              lpcLookupName_LostFocus()
  //              TestStub1()
  //              TestStub1Sub(siIn As StateInfo, curRate As Currency)
  // Modified   :
  // 03/26/01 BAW Cleaned with Total Visual CodeTools 2000
  // 06/18/01 BAW Updated to avoid ADO Error 3001 when doing a Find on a Payee name
  //              that contains an embedded single quote, e.g., O'Dell
  // 10/11/01 BAW Additional changes to accommodate single quotes in Payee Name
  // 01/2002  BAW Updated calcs to ignore scope, and calc interest 3 ways, using the method that
  //              resulted in the highest interest as the "final way." This involved added a
  //              Contract Issue State to the Payee screen as well. Also, removed
  //              "#If gcfLOOKUP" stuff since we definitely want Lookup capability. (At one
  //              time before v2.2 was released, we thought the performance might be too bad to keep it.)
  //              Also, optimized per Project Analyzer (removing dead code, adding "$" to Mid/Space, etc.)
  // Modified  : Berry Kropiwka 2019-09-27, added code from compact calc
  // --------------------------------------------------
//Option Explicit
  *Option Compare Binary

  private String mstrScreenName = "";

  private static final Long MCLNGMINFORMWIDTH = 10275;
  private static final Long MCLNGMINFORMHEIGHT = 8955;

  // The following constants identify, for fpCombo controls used as Lookups,
  // which column is displayed in the Edit portion of the control (index = mcintDisplayCol_xxxx,
  // where xxxx is the fpCombo control's name).
  private static final Integer MCINTDISPLAYCOL_LPCLOOKUPNAME = 0;

  // These constants define the valid entries in the cboPayeSsnTinTypCd combobox.
  private static final String MCSTRPAYEEISABUSINESS = "B";
  private static final String MCSTRPAYEEISAPERSON = "P";

  // These constants define the masks for the ipmPayeSsnTinNum control
  private static final String MCSTRTINMASK = "##-#######";
  private static final String MCSTRSSNMASK = "###-##-####";
  private static final String MCSTRUNKNOWNTINTYPEMASK = "#########";

  // These constants define the columns within the Lookup/Multi-column combo boxes.
  // These are used to give a name to a given column of the fpCombo control so
  // it can be referenced by name, not by number.
  private static final String MCSTRDISPLAYCOL = "DISPLAY_COL";
  private static final String MCSTRPAYEID = "PAYE_ID";
  private static final String MCSTRPAYEFULLNM = "PAYE_FULL_NM";

  // mtWrapper is an instance of the table wrapper corresponding to the main table maintained by this form.
  private ctpyePayee mtWrapper;

  // Define a constant for each field that may get an error. This should match
  // the text of that control's associated Label control.
  //' Editable only upon an Add
  private static final String MCSTRIPTPAYEFULLNMLABEL = "Name";
  private static final String MCSTRIPTPAYECAREOFTXTLABEL = "Care Of";
  private static final String MCSTRIPTPAYEADDRLN1TXTLABEL = "Address1";
  private static final String MCSTRIPTPAYEADDRLN2TXTLABEL = "Address2";
  private static final String MCSTRIPTPAYECITYNMTXTLABEL = "City";
  //' Payee's Residence State at Insured's Death (short name)
  private static final String MCSTRCBOPAYESTCDLABEL = "State";
  private static final String MCSTRIPTPAYEZIPCDLABEL = "Zip";
  private static final String MCSTRIPTPAYEZIP4CDLABEL = "Zip";
  private static final String MCSTRIPMPAYESSNTINNUMLABEL = "TIN/SSN";
  private static final String MCSTRCBOPAYESSNTINTYPCDLABEL = "TIN Type";
  //'' BZ4999 October 2013 Non US payee - SXS
  private static final String MCSTRCHKPAYE1099INDLABEL = "1099INT";
  //OBSOLETE Private Const mcstrCboContractIssueStateLabel As String = "Contract Issue State"    ' Insured's Residence State at Issue (short name)
  private static final String MCSTRIPDPAYEWTHLDRTLABEL = "Withholding Rate";
  private static final String MCSTRCHKPAYEDFLTOVRDINDLABEL = "Override";
  private static final String MCSTRCBOCALCSTCDLABEL = "Calc State";
  private static final String MCSTRIPDPAYECLMINTRTLABEL = "Interest Rate";
  private static final String MCSTRDTPPAYEPMTDTLABEL = "Date of Payment";
  private static final String MCSTRIPDPAYEINTDAYSPDNUMLABEL = "Days of Interest Paid";
  private static final String MCSTRIPCPAYEDTHBPMTAMTLABEL = "DB Payment";
  private static final String MCSTRIPCPAYECLMINTAMTLABEL = "Claim Interest";
  private static final String MCSTRIPCPAYEWTHLDAMTLABEL = "Interest Withheld";
  private static final String MCSTRIPCPAYECLMPDAMTLABEL = "Total";

  private static final String MCSTRTXTINSDDTHRESSTCD_USEDINAUTOCALCLABEL = "Issue State";
  // Labels from Insured screen
  private static final String MCSTRDTPCLMINSDDTHDTLABEL = "Date Of Death";
  private static final String MCSTRDTPCLMPROOFDTLABEL = "Date Of Proof";

  private static final String MCSTRTXTPAYEIDLABEL = "Payee ID";

  //Dim mrstLookup As ADODB.Recordset
  //Dim mrstPayee As ADODB.Recordset
  Form mfrmMyInsuredForm = null;

  // mbInLookupMode determines whether the user is in the process of doing a search using the Lookup ComboBox
  private boolean mbInLookupMode = false;

  // mbInAddMode determines whether the user has begun the process of adding a new record to the table.
  // Note that Add mode is independent of Update mode
  private boolean mbInAddMode = false;

  private Control mctlFirstUpdateableField_Add;
  private Control mctlFirstUpdateableField_Upd;

  // The following field (mcurTotalWithheld) is a "cousin" to ipcPayeWthldAmt
  // that appears on-screen. ipcPayeWthldAmt is formatted with the Format( )
  // function to display as ($$$.$$) since it reduces the total amount paid
  // for a claim. However, mcurTotalWithheld is the unformatted equivalent,
  // unformatted so that the value -- with its sign preserved -- can be
  // stored.  When the Format( ) function adds "(" and ")" around a string and
  // that string is stored, it's regarded as a negative number. Yech!
  Currency mcurTotalWithheld = null;

  StateInfo msiInsdDthResStCd = null;
  StateInfo msiPayeStCd = null;
  StateInfo msiIssStCd = null;
  StateInfo msiOverride = null;
  StateInfo msiCalcStCd = null;
  StateInfo msiCompactCalc = null;


  // m_bIsDirty corresponds to the public property called IsDirty.
  // All maintenance screens should have this field and that property! When True, it indicates
  // that the user has made --but not yet saved-- changes to a record. The MDI form will query
  // this property if the user opens the File menu, since the Exit option should be disabled if
  // any form has outstanding changes.
  // Be sure to use this variable's corresponding Property Let to change its value.
  // Do **NOT** set m_bIsDirty itself, as this will using the Property Let proc will
  // ensure the Close button caption is always synchronized with the value of the property.
  private boolean m_bIsDirty = false;

  // The following UDT is used to suport the "Clone This Payee" functionality
//*TODO:** type is translated as a new class at the end of the file Private Type udtPayeeClone
  // Calc-related fields

  //MME START WRUS 4999
  public double DblScreenDBPaymentValue = 0;
  //MME END WRUS 4999

  private String m_admPolicySystem = "";
  private udtPayeeClone m_upcPayeeClone;
  private static final String M_GROUP_ADMIN_SYS = "GROUP";

  //Private Const variable for the compact filling state code.  When setting the state to this variable in fnCalcClaimInterest when we calcuate the msiCompactCalc it
  //   will using the state rule for compact filling.
  private static final String CSTCOMPACTFILLING = "YY";


  ///\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
  //\                                                                  /
  //|                PROPERTY GET/LET    Procedures                    |
  ///                                                                  \
  //\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/


  //////////////////////////////////////////////////////////////////////////////////////////////////
  public boolean getIsDirty() {
    // Returns True if the record displayed in the form has been
    // edited; False otherwise.
    "Property Get IsDirty"
.equals(Const cstrCurrentProc As String);
    try {

      return m_bIsDirty;
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
    }
    Exit Property;
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
  private void setIsDirty(boolean bValue) {
    // Sets the value of the IsDirty property. This should ONLY be set by this form itself.
    //
    // Be sure to use this Property Let to change the value of the m_bIsDirty variable.
    // Do **NOT** set m_bIsDirty itself, since using the Property Let proc will ensure
    // that the Close button caption is always synchronized with the value of this property.
    "Let IsDirty"
.equals(Const cstrCurrentProc As String);
    "&Cancel"
.equals(Const cstrCancel As String);
    "&Close"
.equals(Const cstrClose As String);

    try {

      m_bIsDirty = bValue;

      // Adjust Close button caption accordingly. Do it conditionally, to avoid
      // flickering when the user does a lot of quick scrolling.
      if (bValue) {
        if (cmdClose.Caption != cstrCancel) {
          cmdClose.Caption = cstrCancel;
        }
      } 
      else {
        if (cmdClose.Caption != cstrClose) {
          cmdClose.Caption = cstrClose;
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
    Exit Property;
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

///\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
//\                                                                  /
//|                Procedures and Event Handlers                     |
///                                                                  \
//\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/


//////////////////////////////////////////////////////////////////////////////////////////////////
  private void cboCalcStCd_Click() {
    // Comments  : Sets a flag to indicate the current record has been
    //             edited, and thus Update button becomes enabled
    // Parameters: N/A
    // Modified  :
    // --------------------------------------------------
    try {
      "cboCalcStCd_Click"
.equals(Const cstrCurrentProc As String);


      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      fnInitializeEditMode();

      // Populate msiOverride structure with data from the STATE_RULE_T
      // row that matches the Calc State.
      fnGetStateInfo_Override();
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
  private void cboPayeSsnTinTypCd_Click() {
    // Comments  :
    // Parameters:  -
    // Modified  :
    // --------------------------------------------------
    try {
      "cboPayeSsnTinTypCd_Click"
.equals(Const cstrCurrentProc As String);
      String strSavedSSN = "";

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      fnInitializeEditMode();

      // Reset the mask and format the value accordingly
      strSavedSSN = ipmPayeSsnTinNum.UnFmtText;
      ipmPayeSsnTinNum.ctclmClaim.setMask("");
      if (MCSTRPAYEEISABUSINESS
.equals(cboPayeSsnTinTypCd.Text)) {
        ipmPayeSsnTinNum.ctclmClaim.setMask(MCSTRTINMASK);
        ipmPayeSsnTinNum.UnFmtText = modGeneral.fnSSNTIN_AddDash(strIn:=strSavedSSN, bIsTin:=True);
      } 
      else if (MCSTRPAYEEISAPERSON
.equals(cboPayeSsnTinTypCd.Text)) {
        ipmPayeSsnTinNum.ctclmClaim.setMask(MCSTRSSNMASK);
        ipmPayeSsnTinNum.UnFmtText = modGeneral.fnSSNTIN_AddDash(strIn:=strSavedSSN, bIsTin:=False);
      } 
      else {
        ipmPayeSsnTinNum.ctclmClaim.setMask(MCSTRUNKNOWNTINTYPEMASK);
        ipmPayeSsnTinNum.UnFmtText = strSavedSSN;
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
  private void cboPayeStCd_Click() {
    // Comments  : Sets a flag to indicate the current record has been
    //             edited, and thus Update button becomes enabled
    // Parameters: N/A
    // Modified  :
    // --------------------------------------------------
    try {
      "cboPayeStCd_Click"
.equals(Const cstrCurrentProc As String);


      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      fnInitializeEditMode();

      txtPayeStCd_UsedInAutoCalc.Text = cboPayeStCd.Text;
      // Populate msiPayeStCd structure with data from the STATE_RULE_T
      // row that matches the Payee State.
      //' BZ4999 October 2013 Non US payee - SXS
      ChkPaye1099Ind.chrgHourglass.setValue("1");
      if ("ZZ"
.equals(cboPayeStCd.Text)) {
        ChkPaye1099Ind.chrgHourglass.setValue("0");
        //''Button Shadow
        ChkPaye1099Ind.ForeColor = -2147483632;
        iptPayeZip4Cd.Text = "     ";
        iptPayeZipCd.Text = "    ";
        ChkPaye1099Ind.Enabled = false;
      } 
      else {
        //'' Button Text
        ChkPaye1099Ind.ForeColor = -2147483630;
        fnGetStateInfo_PayeStCd();
        ChkPaye1099Ind.Enabled = true;
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
//''''''''''''''''''''''''''''  '' BZ4999 October 2013 Non US payee - SXS
//////////////////////////////////////////////////////////////////////////////////////////////////
  private void chkPaye1099Ind_Click() {

    // Comments  : Sets a flag to indicate the current record has been
    //             edited, and thus Update button becomes enabled
    // Parameters: N/A
    // Modified  :
    // --------------------------------------------------
    try {
      "ChkPaye_1099INd_Click"
.equals(Const cstrCurrentProc As String);

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      fnInitializeEditMode();

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


//'''''''''''''''''''''''''''''''''''''
//////////////////////////////////////////////////////////////////////////////////////////////////
  private void chkPayeDfltOvrdInd_Click() {
    // Comments  : Sets a flag to indicate the current record has been
    //             edited, and thus Update button becomes enabled
    // Parameters: N/A
    // Modified  :
    // --------------------------------------------------
    try {
      "chkPayeDfltOvrdInd_Click"
.equals(Const cstrCurrentProc As String);

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      fnInitializeEditMode();

      if (chkPayeDfltOvrdInd.chrgHourglass.getValue() == vbChecked) {
        modGeneral.fnEnableDisableControl(ctlIn:=cboCalcStCd, bEnable:=True);
        modGeneral.fnEnableDisableControl(ctlIn:=ipdPayeClmIntRt, bEnable:=True);
      } 
      else {
        modGeneral.fnEnableDisableControl(ctlIn:=cboCalcStCd, bEnable:=False);
        modGeneral.fnEnableDisableControl(ctlIn:=ipdPayeClmIntRt, bEnable:=False);
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
  private void cmdAdd_Click() {
    // Comments  : Handles the adding of a new record.
    // Parameters:  -
    // Modified  :
    // --------------------------------------------------
    try {
      "cmdAdd_Click"
.equals(Const cstrCurrentProc As String);

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      fnAddRecord();
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
  private void cmdCloneThisPayee_Click() {
    // Comments  : If the user clicked this button, save current fields
    //             to a udt, simulate an cmdAdd_Click event, and use the
    //             udt to pre-fill the new record.
    //
    // Parameters: N/A
    // Modified  :
    // --------------------------------------------------
    try {
      "cmdCloneThisPayee_Click"
.equals(Const cstrCurrentProc As String);

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      // Save current values
      m_upcPayeeClone.payeFullNm = iptPayeFullNm.Text;
      m_upcPayeeClone.payeCareOfTxt = iptPayeCareOfTxt.Text;
      m_upcPayeeClone.payeAddrLn1Txt = iptPayeAddrLn1Txt.Text;
      m_upcPayeeClone.payeAddrLn2Txt = iptPayeAddrLn2Txt.Text;
      m_upcPayeeClone.payeCityNmTxt = iptPayeCityNmTxt.Text;
      m_upcPayeeClone.payeStCd = cboPayeStCd.Text;
      m_upcPayeeClone.payeZipCd = iptPayeZipCd.Text;
      m_upcPayeeClone.payeZip4Cd = iptPayeZip4Cd.Text;
      m_upcPayeeClone.payeSsnTinNum = ipmPayeSsnTinNum.UnFmtText;
      m_upcPayeeClone.payeSsnTinTypCd = cboPayeSsnTinTypCd.Text;
      //'' BZ4999 October 2013 Non US payee - SXS
      m_upcPayeeClone.paye_1099int_ind = (ChkPaye1099Ind.chrgHourglass.getValue() == vbChecked);
      m_upcPayeeClone.payePmtDt = dtpPayePmtDt.chrgHourglass.getValue();
      m_upcPayeeClone.payeDthbPmtAmt = ipcPayeDthbPmtAmt.chrgHourglass.getValue();
      m_upcPayeeClone.clmId = mtWrapper.getClmId();
      // Calc-related fields
      m_upcPayeeClone.payeStCd_UsedInAutoCalc = txtPayeStCd_UsedInAutoCalc.Text;
      m_upcPayeeClone.payeStCdSpecialInstructions_UsedInAutoCalc = txtPayeStCdSpecialInstructions_UsedInAutoCalc.Text;
      m_upcPayeeClone.bClmForResDthInd_UsedInAutoCalc = (chkClmForResDthInd_UsedInAutoCalc.chrgHourglass.getValue() == vbChecked);
      m_upcPayeeClone.insdDthResStCd_UsedInAutoCalc = txtInsdDthResStCd_UsedInAutoCalc.Text;
      m_upcPayeeClone.insdDthResStCdSpecialInstructions_UsedInAutoCalc = txtInsdDthResStCdSpecialInstructions_UsedInAutoCalc.Text;
      m_upcPayeeClone.issStCd_UsedInAutoCalc = txtIssStCd_UsedInAutoCalc.Text;
      m_upcPayeeClone.issStCdSpecialInstructions_UsedInAutoCalc = txtIssStCdSpecialInstructions_UsedInAutoCalc.Text;

      // Hide updates to the window until we're done. This avoids ugly screen flickering
      fnWindowLock(Me.cbrfBrowseFolder.setHWnd());

      fnAddRecord();

      iptPayeFullNm.Text = m_upcPayeeClone.payeFullNm;
      iptPayeCareOfTxt.Text = m_upcPayeeClone.payeCareOfTxt;
      iptPayeAddrLn1Txt.Text = m_upcPayeeClone.payeAddrLn1Txt;
      iptPayeAddrLn2Txt.Text = m_upcPayeeClone.payeAddrLn2Txt;
      iptPayeCityNmTxt.Text = m_upcPayeeClone.payeCityNmTxt;
      cboPayeStCd.Text = m_upcPayeeClone.payeStCd;
      iptPayeZipCd.Text = m_upcPayeeClone.payeZipCd;
      iptPayeZip4Cd.Text = m_upcPayeeClone.payeZip4Cd;
      if (m_upcPayeeClone.paye_1099int_ind) {
        ChkPaye1099Ind.chrgHourglass.setValue(vbChecked);
      } 
      else {
        ChkPaye1099Ind.chrgHourglass.setValue(vbUnchecked);
      }
      ipmPayeSsnTinNum.UnFmtText = m_upcPayeeClone.payeSsnTinNum;
      cboPayeSsnTinTypCd.Text = m_upcPayeeClone.payeSsnTinTypCd;
      dtpPayePmtDt.chrgHourglass.setValue(m_upcPayeeClone.payePmtDt);
      ipcPayeDthbPmtAmt.chrgHourglass.setValue(m_upcPayeeClone.payeDthbPmtAmt);
      mtWrapper.setClmId(m_upcPayeeClone.clmId);
      // Calc-related fields
      txtPayeStCd_UsedInAutoCalc.Text = m_upcPayeeClone.payeStCd_UsedInAutoCalc;
      txtPayeStCdSpecialInstructions_UsedInAutoCalc.Text = m_upcPayeeClone.payeStCdSpecialInstructions_UsedInAutoCalc;
      if (m_upcPayeeClone.bClmForResDthInd_UsedInAutoCalc) {
        chkClmForResDthInd_UsedInAutoCalc.chrgHourglass.setValue(vbChecked);
      } 
      else {
        chkClmForResDthInd_UsedInAutoCalc.chrgHourglass.setValue(vbUnchecked);
      }
      txtInsdDthResStCd_UsedInAutoCalc.Text = m_upcPayeeClone.insdDthResStCd_UsedInAutoCalc;
      txtInsdDthResStCdSpecialInstructions_UsedInAutoCalc.Text = m_upcPayeeClone.insdDthResStCdSpecialInstructions_UsedInAutoCalc;
      txtIssStCd_UsedInAutoCalc.Text = m_upcPayeeClone.issStCd_UsedInAutoCalc;
      txtIssStCdSpecialInstructions_UsedInAutoCalc.Text = m_upcPayeeClone.issStCdSpecialInstructions_UsedInAutoCalc;
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    fnWindowUnlock;

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
  private void cmdClose_Click() {
    // Comments  : If the user clicked the Close button, see if
    //             there are outstanding data changes that have not been saved.
    //             If so, instruct the user how to proceed depending on whether
    //             they want to save or lose their changes.
    //
    //             NOTE: The logic in this function should closely resemble that
    //                   in the Form_QueryUnload event handler!
    // Parameters: N/A
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
  private void cmdDelete_Click() {
    // Comments  : Deletes the current record. Note: This button
    //             will be disabled if any children to this
    //             record (i.e. Payees to this Insured) exist,
    //             forcing the user to first delete those children
    //             and then delete the parent.
    // Parameters: N/A
    // Modified  :
    // --------------------------------------------------
    try {
      "cmdDelete_Click"
.equals(Const cstrCurrentProc As String);
      int intButtonClicked = 0;
      int lngReturnValue = 0;
      String strACF2 = "";
      chrgHourglass hrgHourglass = null;

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      // .......................................................................
      // Make sure the user really, really, really wants to delete this record.
      // .......................................................................
      // 3002 = Are you sure you want to delete this record?
      intButtonClicked = modGeneral.gerhApp.reportNonFatal(vbObjectError + modResConstants.gCRES_ALRT_OK_TO_DELETE_RECORD, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);

      Me.Refresh;

      if ((intButtonClicked == vbNo) || (intButtonClicked == modGeneral.gCINTCLICKEDCLOSEBUTTON)) {
        // **TODO:** goto found: GoTo PROC_EXIT;
      }

      // .......................................................................
      // Proceed with the Delete.
      // *  If another user has updated the record, we don't care. No message
      //    should be generated, and the delete should proceed.
      // *  If another user has deleted the record, display a message to
      //    that effect and then show the record whose key value immediately
      //    preceeds the record we wanted to delete (or if there are now
      //    no other records in the table, go into Add mode).
      // *  If no other user did anything with this record, then just
      //    delete it and then show the record whose key value immediately
      //    preceeds the record we wanted to delete (or if there are now
      //    no other records in the table, go into Add mode).
      //
      // Note that .GetRelativeRecord( ) can be called directly but is also
      // called via the .DeleteRecord( ) method. In both cases, it
      // refreshes the Lookup recordset (m_rstLookup) before positioning
      // to the desired relative record.
      //
      // Anytime the Lookup recordset is refreshed, we need to reload the
      // vfgLookup VSFlexGrid control.
      // .......................................................................
      hrgHourglass = new chrgHourglass();
      hrgHourglass.setValue(true);

      // Hide updates to the window until we're done. This avoids ugly screen flickering
      fnWindowLock(Me.cbrfBrowseFolder.setHWnd());

      lngReturnValue = mtWrapper.checkForAnotherUsersChanges(ewoDelete, strACF2);
      if (lngReturnValue == vbObjectError + modResConstants.gCRES_INFO_ANOTHER_USER_DELETED) {
        // Another user has deleted the record that *this* user is trying to
        // delete. So, display a message to that effect, refresh the Lookup
        // recordset and then show the record whose key value immediately
        // preceeds the record this user wanted to delete.
        modGeneral.gerhApp.reportNonFatal(vbObjectError + modResConstants.gCRES_INFO_ANOTHER_USER_DELETED, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
        // Can't use the GetPayeFullNmFromPayeID( ) method since the PAYEE_T row
        // doesn't exist and hence a "-2147217900" (Payee ID does not exist) error
        // throws things off.
        mtWrapper.getRelativeRecord(mtWrapper.getPayeFullNm(), epdPreviousRecord);
      } 
      else {
        // If another user updated the record *this* user is trying to delete,
        // we don't care. No message should be generated and the delete should
        // proceed as if no other user did anything to this record.
        //
        // If no other user did anything with this record, then delete it,
        // refresh the Lookup recordset, and then show the record whose
        // key value immediately preceeds the record this user wanted to
        // delete.
        mtWrapper.deleteRecord();
      }

      // Repopulate the all Lookup and ComboBox controls so
      // they reflects this and other users' changes.
      fnRefreshAllCombos();


      // If there are no records now in the table (based on this user's or
      // another user's actions), then go into Add mode. Otherwise, display
      // the now-current record. fnLoadControls will set the navigation buttons
      // and "record x of y" label as appropriate.
      if (mtWrapper.getLookupRecordCount() > 0) {
        fnLoadControls();
        fnSetCommandButtons(true);
      } 
      else {
        fnAddRecord();
      }
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    if (!(hrgHourglass == null)) {
      hrgHourglass.setValue(false);
    }
    modGeneral.fnFreeObject(hrgHourglass);
    fnWindowUnlock;

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
  private void cmdNavigate_Click(int pintIndex) { // TODO: Use of ByRef founded Private Sub cmdNavigate_Click(ByRef pintIndex As Integer)
    // Comments  : Enables/Disables the navigation buttons
    //             which is a control array:
    //             (0) = go to first record
    //             (1) = go to prev  record
    //             (2) = go to next  record
    //             (3) = go to last  record
    // Parameters:  -
    // Modified  :
    // --------------------------------------------------
    "cmdNavigate_Click"
.equals(Const cstrCurrentProc As String);
    try {

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      switch (pintIndex) {
        case  navFirst:
          mtWrapper.goToFirstRecord();
          break;

        case  navPrev:
          mtWrapper.goToPreviousRecord();
          break;

        case  navNext:
          mtWrapper.goToNextRecord();
        //' Go to Last
          break;

        default:
          mtWrapper.goToLastRecord();
          break;
      }

      if (modGeneral.bDEBUGAPPTERMINATION) {
        Debug.Print("   Turning off Update mode (#1) in "+ mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
      }
      setIsDirty(false);

      if ((mtWrapper.getCurrentLookupRecordNumber() == adPosBOF) || (mtWrapper.getCurrentLookupRecordNumber() == adPosEOF) || (mtWrapper.getCurrentLookupRecordNumber() == adPosUnknown)) {
        modGeneral.gerhApp.reportNonFatal(vbObjectError + modResConstants.gCRES_INFO_TABLE_IS_EMPTY, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
        fnAddRecord();
      } 
      else {
        // Note that the Lookup controls' selection is no longer synchronized
        // with the table wrapper's CurrentLookupRecordNumber. In other words,
        // the CurrentLookupRecordNumber may indicate we're on the 5th record and,
        // by virtue of fnLoadControls being called following each navigation, that should
        // the same record that is currently displayed on-screen. However, the Lookup
        // controls themselves are not necessarily *itself* positioned to the 5th record.
        // The total number of entries in that control, however, should jive with the
        // table wrapper's LookupRecordCount property.

        // Load current record's properties to form's controls, reset navigation buttons and set "rec x of y" label
        fnLoadControls();
        if (modGeneral.bDEBUGAPPTERMINATION) {
          Debug.Print("   Turning off Update mode (#2) in "+ mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
        }

        setIsDirty(false);
        fnSetCommandButtons(true);
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
  private void cmdUpdate_Click() {
    // Comments:    This function handles updating an existing record or, if in Add mode,
    //              the adding of a new record. It is called when the user clicks the
    //              Update button, as well as by Form_QueryUnload when the user
    //              attempts to close the form while edits are outstanding.
    // Parameters:  -
    // Modified  :
    // --------------------------------------------------
    "cmdUpdate_Click"
.equals(Const cstrCurrentProc As String);
    int lngReturnValue = 0;
    String strACF2 = "";
    chrgHourglass hrgHourglass = null;

    try {

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      hrgHourglass = new chrgHourglass();
      hrgHourglass.setValue(true);

      if (!(fnValidData())) {
        // **TODO:** goto found: GoTo PROC_EXIT;
      }

      // Update the Table wrapper's properties with the screen values
      mtWrapper.setPayeFullNm(iptPayeFullNm.Text);
      mtWrapper.setPayeCareOfTxt(iptPayeCareOfTxt.Text);
      mtWrapper.setPayeAddrLn1Txt(iptPayeAddrLn1Txt.Text);
      mtWrapper.setPayeAddrLn2Txt(iptPayeAddrLn2Txt.Text);
      mtWrapper.setPayeCityNmTxt(iptPayeCityNmTxt.Text);
      mtWrapper.setPayeStCd(cboPayeStCd.Text);
      mtWrapper.setPayeZipCd(iptPayeZipCd.Text);
      mtWrapper.setPayeZip4Cd(iptPayeZip4Cd.Text);

      //' Use .UnFmtText to get rid of mask characters in fpMask control
      mtWrapper.setPayeSsnTinNum(ipmPayeSsnTinNum.UnFmtText);
      // cboPayeSsnTinTypCd corresponds to a Nullable field, so accommodate Nulls
      if (cboPayeSsnTinTypCd.Text == modGeneral.gCSTRBLANKENTRY) {
        mtWrapper.setPayeSsnTinTypCd("");
      } 
      else {
        mtWrapper.setPayeSsnTinTypCd(cboPayeSsnTinTypCd.Text);
      }
      //' BZ4999 October 2013 Non US payee - SXS
      mtWrapper.setPaye1099INTInd((ChkPaye1099Ind.chrgHourglass.getValue() == vbChecked));
      mtWrapper.setPayeWthldRt(ipdPayeWthldRt.chrgHourglass.getValue());

      mtWrapper.setPayeDfltOvrdInd((chkPayeDfltOvrdInd.chrgHourglass.getValue() == vbChecked));
      mtWrapper.setPayePmtDt(dtpPayePmtDt.chrgHourglass.getValue());
      mtWrapper.setPayeDthbPmtAmt(ipcPayeDthbPmtAmt.chrgHourglass.getValue());

      mtWrapper.setLstUpdtUserId(modGeneral.gconAppActive.getLastLogOnUserID());
      mtWrapper.setLstUpdtDtm(Now);

      // These will propagate back an error if the Insert/Update failed.
      if (mbInAddMode) {
        fnCalcClaimInterest();

        // Update wrapper with values calculated or possibly affected by fnCalcClaimInterest( )
        mtWrapper.setPayeIntDaysPdNum(ipdPayeIntDaysPdNum.chrgHourglass.getValue());
        mtWrapper.setPayeClmIntAmt(ipcPayeClmIntAmt.chrgHourglass.getValue());
        mtWrapper.setPayeWthldAmt(ipcPayeWthldAmt.chrgHourglass.getValue());
        mtWrapper.setPayeClmPdAmt(ipcPayeClmPdAmt.chrgHourglass.getValue());
        mtWrapper.setCalcStCd(cboCalcStCd.Text);
        mtWrapper.setPayeClmIntRt(ipdPayeClmIntRt.chrgHourglass.getValue());

        // Add the record, refresh the lookup recordset and reposition
        // to the record just added
        mtWrapper.addRecord();
        // Turn off Add mode since the Add was successful
        mbInAddMode = false;
        // Repopulate the all Lookup and ComboBox controls so
        // they reflects this and other users' changes.
        fnRefreshAllCombos();
        // This **must** be done as the user leaves Add mode, so that the key fields
        // will now be protected to prevent the user from being able to edit them.
        // Editing a key field is allowed only when in Add mode.
        fnSetAvailabilityOfControls();
      } 
      else {
        // Determine whether another user updated or deleted the record about to be updated.
        // Note: this multi-user checking is performed on an Update but not an Add.
        lngReturnValue = mtWrapper.checkForAnotherUsersChanges(ewoUpdate, strACF2);

        if (lngReturnValue == vbObjectError + modResConstants.gCRES_INFO_ANOTHER_USER_DELETED_NOT_SAVED) {
          modGeneral.gerhApp.reportNonFatal(vbObjectError + modResConstants.gCRES_INFO_ANOTHER_USER_DELETED_NOT_SAVED, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
          // Discard *this* user's pending changes and show the previous record.
          // Can't use the GetPayeFullNmFromPayeID( ) method since the PAYEE_T row
          // doesn't exist and hence a "-2147217900" (Payee ID does not exist) error
          // throws things off.
          mtWrapper.getRelativeRecord(mtWrapper.getPayeFullNm(), epdPreviousRecord);

        } 
        else if (lngReturnValue == vbObjectError + modResConstants.gCRES_INFO_ANOTHER_USER_UPDATED_DISCARDED) {
          modGeneral.gerhApp.reportNonFatal(vbObjectError + modResConstants.gCRES_INFO_ANOTHER_USER_UPDATED_DISCARDED, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc, strACF2.trim());
          // Discard *this* user's pending changes by re-retrieving the current record
          // as it currently looks on the database and refreshing the lookup recordset
          // Can't use the GetPayeFullNmFromPayeID( ) method since the PAYEE_T row
          // doesn't exist and hence a "-2147217900" (Payee ID does not exist) error
          // throws things off.
          mtWrapper.getRelativeRecord(mtWrapper.getPayeFullNm(), epdSameRecord);
        } 
        else {
          fnCalcClaimInterest();

          // Update wrapper with values calculated or possibly affected by fnCalcClaimInterest( )
          mtWrapper.setPayeIntDaysPdNum(ipdPayeIntDaysPdNum.chrgHourglass.getValue());
          mtWrapper.setPayeClmIntAmt(ipcPayeClmIntAmt.chrgHourglass.getValue());
          mtWrapper.setPayeWthldAmt(ipcPayeWthldAmt.chrgHourglass.getValue());
          mtWrapper.setPayeClmPdAmt(ipcPayeClmPdAmt.chrgHourglass.getValue());
          mtWrapper.setCalcStCd(cboCalcStCd.Text);
          mtWrapper.setPayeClmIntRt(ipdPayeClmIntRt.chrgHourglass.getValue());

          // Update the record with this user's pending changes, refresh the lookup
          // recordset and reposition to the record just updated
          mtWrapper.updateRecord();
        }

        // Turn off Update mode since the Update was either successful or abandoned
        if (modGeneral.bDEBUGAPPTERMINATION) {
          Debug.Print("   Turning off Update mode in "+ mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
        }
        setIsDirty(false);

        // Repopulate the all Lookup and ComboBox controls so
        // they reflects this and other users' changes.
        fnRefreshAllCombos();
      }

      // Do an immediate repaint. This allows the Insured screen to be redrawn BEFORE all
      // the work of requerying and repainting is started. When the requerying/repainting is done,
      // only small parts of the screen (not the whole screen) will need to be repainted. This
      // eliminates the user seeing a very slow repainting.
      Me.Refresh;

      if (mtWrapper.getLookupRecordCount() > 0) {
        // Ensure the on-screen controls reflect the record just added/updated, in case the
        // DBMS altered it in some way, e.g., determining an Identity column value and
        // getting the most up-to-date Last Updated info. This also sets the navigation
        // buttons and updates the "record x of y" label
        fnLoadControls();

        // Display the message indicating the calc was overriden if the Override
        // checkbox is still selected.
        lblWarningAboutOverride.Visible = (chkPayeDfltOvrdInd.chrgHourglass.getValue() == vbChecked);

        fnSetCommandButtons(true);

        if (((ipmPayeSsnTinNum.UnFmtText.length() == 0) || ("000000000".equals(ipmPayeSsnTinNum.UnFmtText))) && ((CDbl(ipdPayeWthldRt.chrgHourglass.getValue()) == 0) || ipdPayeWthldRt.chrgHourglass.getValue().length() == 0) && (Double.parseDouble(ipcPayeClmIntAmt.UnFmtText) >= msiCalcStCd.StrlIntRptgFlrAmt)) {
          // gcRES_WARN_GET_TIN_BEFORE_PAYING_INT (2004) = This claims requires a certified @@1 to avoid withholding.
          //                                               Make sure you don't pay interest until it has been received.
          // Per Michelle Wilkosky, this warning should gen'd if the calculated interest equals or exceeds the
          //      state reporting floor AND (either the TIN was not supplied or was set to all zeroes) AND
          //      (either the Withholding Rate was not supplied or was set to 0)
          modGeneral.gerhApp.reportNonFatal(vbObjectError + modResConstants.gCRES_WARN_GET_TIN_BEFORE_PAYING_INT, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc, MCSTRIPMPAYESSNTINNUMLABEL);
        }
      } 
      else {
        fnAddRecord();
      }

      Me.Refresh;
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    if (!(hrgHourglass == null)) {
      hrgHourglass.setValue(false);
    }
    modGeneral.fnFreeObject(hrgHourglass);

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
  private void dtpPayePmtDt_Change() {
    // Comments  : Since this field was just changed, reset
    //             Enabled property on command and navigation
    //             buttons as appropriate given that the user
    //             is in the middle of updating a record.
    // Parameters:  -
    // Modified  :
    // --------------------------------------------------
    try {
      "dtpPayePmtDt_Change"
.equals(Const cstrCurrentProc As String);

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      fnInitializeEditMode();

      fnResetStateRules();
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
  private void fnAddRecord() {
    // Comments  : This function handles adding a new record. It is called
    //             by cmdAdd_Click (when the user clicks the Add button)
    //             and by cmdDelete_Click (when the last record in the
    //             recordset is deleted)
    // Parameters:  -
    // Returns   :  -
    // Modified  :
    // --------------------------------------------------
    "fnAddRecord"
.equals(Const cstrCurrentProc As String);

    try {

      // All we do here is display an empty record. The cmdUpdate_Click event
      // handler actually does the add when it sees that mbInAddMode=True.
      // Adds and Updates are treated very nearly the same in that event handler!

      mbInAddMode = true;

      // Display empty or initialized values for on-screen controls
      fnClearControls();

      // We can't populate the StateInfo structures associated with the various State Codes
      // used in the calculations since, with the 06/2003 release of the system, the state
      // rules now can vary by Date of Payment. So, the dtpPayePmtDt_Change event handler
      // populates them. Here, however, we can set the State Codes themselves and empty out
      // their associated Special Instructions.
      // 1. Insured State of Residence at time of death (carried over from Insured screen)
      txtInsdDthResStCd_UsedInAutoCalc.Text = mfrmMyInsuredForm.frmInsured.getInsuredInsdDthResStCd();
      txtInsdDthResStCdSpecialInstructions_UsedInAutoCalc.Text = "";
      if (mfrmMyInsuredForm.frmInsured.getInsuredClmForResDthInd()) {
        Me.chkClmForResDthInd_UsedInAutoCalc = vbChecked;
      } 
      else {
        Me.chkClmForResDthInd_UsedInAutoCalc = vbUnchecked;
      }
      // 2. Contract Issue State (carried over from Insured screen)
      txtIssStCd_UsedInAutoCalc.Text = mfrmMyInsuredForm.frmInsured.getInsuredIssStCd();
      txtIssStCdSpecialInstructions_UsedInAutoCalc.Text = "";
      // 3. Payee Residence State at time of death
      txtPayeStCd_UsedInAutoCalc.Text = cboPayeStCd.Text;
      txtPayeStCdSpecialInstructions_UsedInAutoCalc.Text = "";

      if (modGeneral.bDEBUGAPPTERMINATION) {
        Debug.Print("   Turning off Update mode (#1) in "+ mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
      }
      setIsDirty(false);

      // Enable and set focus to key field(s) so the user can specify a value.
      // This **must** be done as the user goes into Add mode, so they can specify
      // the key(s) for the record they're adding.
      fnSetAvailabilityOfControls();

      // Restrike "Record x of y" to reflect pending Add. Can't call fnShowRecordPosition
      // since it is based on a recordset's AbsolutePosition which, in unbound /disconnected mode,
      // isn't set appropriately.
      lblRecordPosition = "Record ? of "+ ((Integer) mtWrapper.getLookupRecordCount()).toString();

      if (modGeneral.bDEBUGAPPTERMINATION) {
        Debug.Print("   Turning off Update mode (#2) in "+ mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
      }
      setIsDirty(false);
      fnSetCommandButtons(false);

      fnSetNavigationButtons(bUnconditionalDisable:=True);

      // Make sure first field gets the focus. Note, when Add mode is triggered
      // from Form_Load, this statement accomplishes nothing: the control isn't yet visible,
      // so it can't receive the focus. This is why Form_Activate must also call this function.
      fnSetFocusToFirstUpdateableField();
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
  private void fnBindControlsToTableWrapper() {
    //--------------------------------------------------------------------------
    // Procedure:   fnBindControls
    // Description: Binds the on-screen controls to the table wrapper class
    //              properties with which they are associated. This is done so
    //              various control properties can be set based on meta data
    //              gathered by the table wrapper class.
    //
    // Params:      N/A
    // Returns:     N/A
    // Date:        04/04/2002
    //-----------------------------------------------------------------------------
    "fnBindControlsToTableWrapper"
.equals(Const cstrCurrentProc As String);
    try {


      iptPayeFullNm.Tag = "PayeFullNm";
      iptPayeCareOfTxt.Tag = "PayeCareOfTxt";
      iptPayeAddrLn1Txt.Tag = "PayeAddrLn1Txt";
      iptPayeAddrLn2Txt.Tag = "PayeAddrLn2Txt";
      iptPayeCityNmTxt.Tag = "PayeCityNmTxt";
      cboPayeStCd.Tag = "PayeStCd";
      iptPayeZipCd.Tag = "PayeZipCd";
      iptPayeZip4Cd.Tag = "PayeZip4Cd";
      ChkPaye1099Ind.Tag = "Paye1099Ind";
      ipmPayeSsnTinNum.UnFmtText = "PayeSsnTinNum";
      cboPayeSsnTinTypCd.Tag = "PayeSsnTinTypCd";
      ipdPayeWthldRt.Tag = "PayeWthldRt";
      ipdPayeClmIntRt.Tag = "PayeClmIntRt";
      chkPayeDfltOvrdInd.Tag = "PayeDfltOvrdInd";
      cboCalcStCd.Tag = "CalcStCd";
      dtpPayePmtDt.Tag = "PayePmtDt";
      ipdPayeIntDaysPdNum.Tag = "PayeIntDaysPdNum";
      ipcPayeDthbPmtAmt.Tag = "PayeDthbPmtAmt";
      ipcPayeClmIntAmt.Tag = "PayeClmIntAmt";
      ipcPayeWthldAmt.Tag = "PayeWthldAmt";
      ipcPayeClmPdAmt.Tag = "PayeClmPdAmt";

      //!TODO! - ClmId too?
      //!TODO! - PayeId too?

      // LstUpdDtm     isn't shown on-screen
      // LstUpdUserId  isn't shown on-screen
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
  private void fnCalcAndLogMSIInfo(StateInfo msiIn, String strDesc, boolean bUseSuppliedRate) { // TODO: Use of ByRef founded Private Sub fnCalcAndLogMSIInfo(ByRef msiIn As StateInfo, ByVal strDesc As String, Optional ByVal bUseSuppliedRate As Boolean = False)
    // Comments  : This function will calculate and log info from the specified State Info structure
    //             to the application log file
    // Parameters: None
    // Returns   : True, if successful; False otherwise
    // Modified  :
    // --------------------------------------------------
    try {
      "fnCalcAndLogMSIInfo"
.equals(Const cstrCurrentProc As String);
      "#####0.00000"
.equals(Const cstrDec11_5 As String);
      "Currency"
.equals(Const cstrCurrency As String);

      msiIn.CalculationInfo = "The claim interest was calculated based on rates "+ "for the "+ strDesc+ ". ";
      // The next statement does the calc, and updates the msiIn structure with the results of the calc
      fnCalcClaimForState(msiIn, strDesc, bUseSuppliedRate);
      modAppLog.fnLogWrite(" ", cstrCurrentProc);
      modAppLog.fnLogWrite(strDesc+ ":", cstrCurrentProc);
      modAppLog.fnLogWrite("  State:                          "+ msiIn.StCd, cstrCurrentProc);
      modAppLog.fnLogWrite("  Line-of-business:               "+ msiIn.LobCd, cstrCurrentProc);
      modAppLog.fnLogWrite("  Rule Effective Date:            "+ modDataConversion.fnZLSIfNull(msiIn.StrlEffDt), cstrCurrentProc);
      modAppLog.fnLogWrite("  Rule End Date:                  "+ modDataConversion.fnZLSIfNull(msiIn.StrlEndDt), cstrCurrentProc);
      modAppLog.fnLogWrite("  Interest Required Date Type:    "+ msiIn.ReqdIdtypCd, cstrCurrentProc);
      modAppLog.fnLogWrite("  Interest Required Offset:       "+ msiIn.StrlIntReqdOfstNum, cstrCurrentProc);
      modAppLog.fnLogWrite("  Interest Calculation Date Type: "+ msiIn.CalcIdtypCd, cstrCurrentProc);
      modAppLog.fnLogWrite("  Interest Calculation Offset:    "+ msiIn.StrlIntCalcOfstNum, cstrCurrentProc);
      modAppLog.fnLogWrite("  Interest Rule Code:             "+ msiIn.IruleCd, cstrCurrentProc);
      modAppLog.fnLogWrite("  Interest Rule Amount:           "+ Format$(modDataConversion.fnZeroIfNull(msiIn.StrlIntRuleAmt), cstrDec11_5), cstrCurrentProc);
      modAppLog.fnLogWrite("  Reporting Floor Amount:         "+ Format$(msiIn.StrlIntRptgFlrAmt, "Currency"), cstrCurrentProc);
      modAppLog.fnLogWrite(modGeneral.gCSTRBLANKENTRY, cstrCurrentProc);
      modAppLog.fnLogWrite("  SpecialInstructions:            "+ msiIn.StrlSpclInstrTxt, cstrCurrentProc);
      modAppLog.fnLogWrite("  Figured From Date:              "+ CStr(DateValue(msiIn.FiguredFromDate)), cstrCurrentProc);
      modAppLog.fnLogWrite("  PayablePeriodEndDate:           "+ CStr(DateValue(msiIn.PayablePeriodEndDate)), cstrCurrentProc);
      modAppLog.fnLogWrite("  InterestRateToUse:              "+ Format$(msiIn.InterestRateToUse, cstrDec11_5), cstrCurrentProc);
      modAppLog.fnLogWrite(modGeneral.gCSTRBLANKENTRY, cstrCurrentProc);
      modAppLog.fnLogWrite("  NbrOfDaysToPayInterest:         "+ msiIn.NbrOfDaysToPayInterest, cstrCurrentProc);
      modAppLog.fnLogWrite("  ClaimInterest:                  "+ Format$(msiIn.ClaimInterestAmt, cstrCurrency), cstrCurrentProc);
      modAppLog.fnLogWrite("  Withheld:                       "+ Format$(msiIn.WithheldAmt, cstrCurrency), cstrCurrentProc);
      modAppLog.fnLogWrite("  TotalForThisPayee:              "+ Format$(msiIn.TotalForThisPayee, cstrCurrency), cstrCurrentProc);
      modAppLog.fnLogWrite("  CalculationInfo:                "+ msiIn.CalculationInfo, cstrCurrentProc);
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
  private boolean fnCalcClaimInterest() {
    boolean _rtn = false;
    // Comments  : This function will calculate Claim Interest
    //             for this Payee
    // Parameters: None
    // Returns   : True, if successful; False otherwise
    // Modified  :
    // Modified  : Berry Kropiwka 2019-09-27, added code for compact calc
    // --------------------------------------------------
    try {
      "fnCalcClaimInterest"
.equals(Const cstrCurrentProc As String);

      modAppLog.fnLogWrite(modGeneral.gCSTRBLANKENTRY, cstrCurrentProc);
      modAppLog.fnLogWrite("---New Calc---", cstrCurrentProc);
      modAppLog.fnLogWrite("Inputs:", cstrCurrentProc);
      modAppLog.fnLogWrite("  Claim Number:                   "+ mfrmMyInsuredForm.frmInsured.getInsuredClmNum(), cstrCurrentProc);
      modAppLog.fnLogWrite("  Payee:                          "+ iptPayeFullNm.Text, cstrCurrentProc);
      modAppLog.fnLogWrite("  Insured Foreign Res. at Death:  "+ (chkClmForResDthInd_UsedInAutoCalc.chrgHourglass.getValue() == vbChecked), cstrCurrentProc);
      modAppLog.fnLogWrite("  Insured Residence State:        "+ mfrmMyInsuredForm.frmInsured.getInsuredInsdDthResStCd(), cstrCurrentProc);
      modAppLog.fnLogWrite("  Contract Issue State:           "+ mfrmMyInsuredForm.frmInsured.getInsuredIssStCd(), cstrCurrentProc);
      modAppLog.fnLogWrite("  Date of Proof:                  "+ CStr(DateValue(mfrmMyInsuredForm.frmInsured.getInsuredClmProofDt())), cstrCurrentProc);
      modAppLog.fnLogWrite("  Date of Death:                  "+ CStr(DateValue(mfrmMyInsuredForm.frmInsured.getInsuredClmInsdDthDt())), cstrCurrentProc);
      modAppLog.fnLogWrite("  Date of Payment:                "+ CStr(DateValue(dtpPayePmtDt.chrgHourglass.getValue())), cstrCurrentProc);
      modAppLog.fnLogWrite("  Payment:                        "+ ipcPayeDthbPmtAmt.Text, cstrCurrentProc);
      modAppLog.fnLogWrite("  Withholding Percent:            "+ ipdPayeWthldRt.Text, cstrCurrentProc);
      modAppLog.fnLogWrite("  Calculation Override:           "+ (chkPayeDfltOvrdInd.chrgHourglass.getValue() == vbChecked)+ "    CalcState="+ cboCalcStCd.Text+ "    InterestRate="+ CStr(ipdPayeClmIntRt.chrgHourglass.getValue()), cstrCurrentProc);

      // Initialize the calculated amounts in each StateInfo structure
      fnInitializeCalcInfo(msiCalcStCd);
      fnInitializeCalcInfo(msiInsdDthResStCd);
      fnInitializeCalcInfo(msiPayeStCd);
      fnInitializeCalcInfo(msiIssStCd);
      if (mfrmMyInsuredForm.chkClmCmpCalInd.chrgHourglass.getValue() == vbChecked) {
        // This is Compact Calcatution
        fnInitializeCalcInfo(msiCompactCalc);
      }

      //Debug.Print "msiCalcStCd.StrlEffDt, msiInsdDthResStCd.StrlEffDt, msiPayeStCd.StrlEffDt, msiIssStCd.StrlEffDt, msiOverride.StrlEffDt:"
      //Debug.Print msiCalcStCd.StrlEffDt, msiInsdDthResStCd.StrlEffDt, msiPayeStCd.StrlEffDt, msiIssStCd.StrlEffDt, msiOverride.StrlEffDt

      // If a Calculation Override is being done, then ONLY do a calc using the Calc St/Interest Rate. (1-way)
      if (chkPayeDfltOvrdInd.chrgHourglass.getValue() == vbChecked) {
        msiOverride.StCd = cboCalcStCd.Text;
        msiOverride.InterestRateToUse = ipdPayeClmIntRt.chrgHourglass.getValue();
        fnCalcAndLogMSIInfo(msiOverride, "Calc State", true);
        msiCalcStCd = msiOverride;
      } 
      else {
        // Only do a calc using Insured Residence State At Time Of Death if the Insured lived within the
        // United States and its territories (3-way)
        if (chkClmForResDthInd_UsedInAutoCalc.chrgHourglass.getValue() == vbUnchecked) {
          fnCalcAndLogMSIInfo(msiInsdDthResStCd, "Insured's Residence State");
          //' BZ4999 October 2013 Non US payee sxs
          if (!("ZZ".equals(cboPayeStCd))) {
            fnCalcAndLogMSIInfo(msiPayeStCd, "Payee's Residence State");
          }
          fnCalcAndLogMSIInfo(msiIssStCd, "Contract Issue State");
        } 
        else {
          // Otherwise do a 2-way calc
          //' BZ4999 October 2013 Non US payee sxs
          if (!("ZZ".equals(cboPayeStCd))) {
            fnCalcAndLogMSIInfo(msiPayeStCd, "Payee's Residence State");
          }
          fnCalcAndLogMSIInfo(msiIssStCd, "Contract Issue State");
        }
        if (mfrmMyInsuredForm.chkClmCmpCalInd.chrgHourglass.getValue() == vbChecked) {
          // This is Compact Calcatution
          //if death to payment is less than 31 days then use current rate
          //if over pay from proof to payment, with an 10% rate
          msiCompactCalc.StCd = CSTCOMPACTFILLING;
          fnCalcAndLogMSIInfo(msiCompactCalc, "Compact Calculation");
        }
        // Pick the one that calculated the highest Claim Interest
        // (Be sure not to look at msiInsdDthResStCd first since it would have an empty StCd value
        //  if the Foreign Residence at Death checkbox is selected and hence the Update could fail
        //  due to "blank" not being defined on STATE_T.)
        //' BZ4999 October 2013 Non US payee - SXS
        if ("ZZ"
.equals(cboPayeStCd)) {
          msiCalcStCd = msiIssStCd;
        } 
        else {
          msiCalcStCd = msiPayeStCd;
        }
        //Y027 07-Nov-2012
        //If its a GROUP Policy and any of these five states are involved then ignore that state
        if (m_admPolicySystem.equals(M_GROUP_ADMIN_SYS)) {
          //Assign it to a state which is not to be ignored
          //' BZ4999 October 2013 Non US payee - SXS
          if (fnAnomolyState(msiPayeStCd.StCd) == false  && !("ZZ".equals(cboPayeStCd))) {
            msiCalcStCd = msiPayeStCd;
          } 
          else if (fnAnomolyState(msiInsdDthResStCd.StCd) == false) {
            msiCalcStCd = msiInsdDthResStCd;
          } 
          else {
            msiCalcStCd = msiIssStCd;
          }
          if (msiIssStCd.ClaimInterestAmt > msiCalcStCd.ClaimInterestAmt && fnAnomolyState(msiIssStCd.StCd) == false) {
            msiCalcStCd = msiIssStCd;
          }
          if (msiInsdDthResStCd.ClaimInterestAmt > msiCalcStCd.ClaimInterestAmt && fnAnomolyState(msiInsdDthResStCd.StCd) == false) {
            msiCalcStCd = msiInsdDthResStCd;
          }
          if (msiPayeStCd.ClaimInterestAmt > msiCalcStCd.ClaimInterestAmt && fnAnomolyState(msiPayeStCd.StCd) == false) {
            msiCalcStCd = msiPayeStCd;
          }
          if (mfrmMyInsuredForm.chkClmCmpCalInd.chrgHourglass.getValue() == vbChecked) {
            // This is Compact Calcatution
            if (msiCompactCalc.ClaimInterestAmt > msiCalcStCd.ClaimInterestAmt && fnAnomolyState(msiCompactCalc.StCd) == false) {
              msiCalcStCd = msiCompactCalc;
            }
          }
          //At this point we have the highest interest rate of a non anomolous state or
          //we have an anomolous state
        } 
        else {
          //Non group policy
          if (msiInsdDthResStCd.ClaimInterestAmt > msiCalcStCd.ClaimInterestAmt) {
            msiCalcStCd = msiInsdDthResStCd;
          }
          if (msiIssStCd.ClaimInterestAmt > msiCalcStCd.ClaimInterestAmt) {
            msiCalcStCd = msiIssStCd;
          }
          if (mfrmMyInsuredForm.chkClmCmpCalInd.chrgHourglass.getValue() == vbChecked) {
            // This is Compact Calcatution
            if (msiCompactCalc.ClaimInterestAmt > msiCalcStCd.ClaimInterestAmt) {
              msiCalcStCd = msiCompactCalc;
            }
          }
        }
      }

      modAppLog.fnLogWrite(modGeneral.gCSTRBLANKENTRY, cstrCurrentProc);
      modAppLog.fnLogWrite("The state selected was "+ msiCalcStCd.StCd+ ".", cstrCurrentProc);

      cboCalcStCd.Text = msiCalcStCd.StCd;
      txtCalcStCdSpecialInstructions_UsedInAutoCalc.Text = msiCalcStCd.StrlSpclInstrTxt;

      if (fnAnomolyState(msiCalcStCd.StCd) == true  && m_admPolicySystem.equals(M_GROUP_ADMIN_SYS)) {

        lblCalculationInfo = "Note: This is a group policy - an interest rate of 0% applies.";
        ipcPayeClmIntAmt.Text = 0;
        ipcPayeWthldAmt.Text = 0;
        mcurTotalWithheld = 0;
        ipdPayeIntDaysPdNum.Text = 0;
        ipdPayeClmIntRt.UnFmtText = 0;
        ipcPayeClmPdAmt.Text = ipcPayeDthbPmtAmt.Text;

      } 
      else {

        // Initialize on-screen label that shows some info about how the
        // calculation was done.
        lblCalculationInfo = msiCalcStCd.CalculationInfo;
        ipcPayeClmIntAmt.Text = msiCalcStCd.ClaimInterestAmt;
        ipcPayeWthldAmt.Text = msiCalcStCd.WithheldAmt;
        mcurTotalWithheld = msiCalcStCd.WithheldAmt;
        ipdPayeIntDaysPdNum.Text = msiCalcStCd.NbrOfDaysToPayInterest;
        ipcPayeClmPdAmt.Text = msiCalcStCd.TotalForThisPayee;

        ipdPayeClmIntRt.UnFmtText = msiCalcStCd.InterestRateToUse;


      }



      _rtn = true;
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


  private boolean fnAnomolyState(String strStateCode) { // TODO: Use of ByRef founded Private Function fnAnomolyState(ByRef strStateCode As String) As Boolean
    boolean _rtn = false;
    if (!(strStateCode.equals("AR")) && !(strStateCode.equals("FL")) && !(strStateCode.equals("IN")) && !(strStateCode.equals("IL")) && !(strStateCode.equals("MA")) && !(strStateCode.equals(" "))) {
      _rtn = false;
    } 
    else {
      _rtn = true;
    }
  
    return _rtn;
  }

  //////////////////////////////////////////////////////////////////////////////////////////////////
  private boolean fnCalcClaimForState(StateInfo msiStateIn, String strDesc, boolean bUseSuppliedRate) { // TODO: Use of ByRef founded Private Function fnCalcClaimForState(ByRef msiStateIn As StateInfo, ByVal strDesc As String, ByVal bUseSuppliedRate As Boolean) As Boolean
    boolean _rtn = false;
    // Comments  : This function will calculate Claim Interest
    //             for the specified State.
    // Parameters:
    //             msiStateIn (in/out)   Pointer to a StateInfo structure, containing state to calculate
    //             strDesc (in)          Descriptive text to appear in prompts to explain why rate is needed
    //             bUseSuppliedRate (in) If True, indicates to use supplied rate (as if the IRULE_CD = "SPECAMT")
    //                                   and to suppress all rate-related prompts for that state.
    // Returns   : True, if successful; False otherwise
    // Modified  :
    //  07/17/03 K758  For bug 2455, modified the calc of # of Days Of Interest To Be Paid
    //                 to floor it at zero, so negative numbers won't come through and
    //                 adversely affect the Claims Interest Amount's calculation.
    // --------------------------------------------------
    try {

      "fnCalcClaimForState"
.equals(Const cstrCurrentProc As String);
      Const(cdblFloorOfZero As Double == 0);
      Date dteDateOfPayment = null;

      dteDateOfPayment = DateValue(dtpPayePmtDt.chrgHourglass.getValue());

      //       **************************************************
      //       **************************************************
      //         The steps listed below are described in the
      //         ClaimsInterest_HowToDoManualCalc.Doc document
      //         on \\500ip03\Vol2\DesktopTechnology\Deploy\Docs
      //       **************************************************
      //       **************************************************

      // ................................................................
      //    Step 1. No interest will be paid if CalcIdtypCd=None
      // ................................................................
      if (msiStateIn.CalcIdtypCd.toUpperCase().equals("NONE    ")) {
        msiStateIn.CalculationInfo = msiStateIn.CalculationInfo+ "No interest was paid due "+ "to that state's Calculation Interest Date Type Code specification. ";
        msiStateIn.NbrOfDaysToPayInterest = 0;
        // **TODO:** goto found: GoTo STEP8;
      }

      // ................................................................
      //    Step 2. No interest will be paid if ReqdIdtypCd=None
      // ................................................................
      if (msiStateIn.ReqdIdtypCd.toUpperCase().equals("NONE    ")) {
        msiStateIn.CalculationInfo = msiStateIn.CalculationInfo+ "No interest was paid due "+ "to that state's Required Interest Date Type Code specification. ";
        msiStateIn.NbrOfDaysToPayInterest = 0;
        // **TODO:** goto found: GoTo STEP8;
      }

      // ................................................................
      //    Step 3. Calculate the Payable Period End Date. If the
      //            claim is being paid ON or BEFORE that date,
      //            we do NOT pay interest on the claim. If the
      //            claim is being paid AFTER that date, we may have
      //            to pay interest on the claim (unless Step 7 says
      //            otherwise).
      // ................................................................
      if (msiStateIn.ReqdIdtypCd.toUpperCase().equals("PROOF   ")) {
        msiStateIn.PayablePeriodEndDate = DateValue(mfrmMyInsuredForm.frmInsured.getInsuredClmProofDt()) + msiStateIn.StrlIntReqdOfstNum;
      } 
      else {
        msiStateIn.PayablePeriodEndDate = DateValue(mfrmMyInsuredForm.frmInsured.getInsuredClmInsdDthDt()) + msiStateIn.StrlIntReqdOfstNum;
      }

      // ................................................................
      //    Step 4. No interest will be paid if the claim is being
      //            paid on or before the Payable Period End Date.
      // ................................................................
      if (dteDateOfPayment <= msiStateIn.PayablePeriodEndDate) {
        msiStateIn.CalculationInfo = msiStateIn.CalculationInfo+ "No interest was paid due "+ "to the Date Of Payment being within the Payable Period. ";
        msiStateIn.NbrOfDaysToPayInterest = 0;
        // **TODO:** goto found: GoTo STEP8;
      }

      // ................................................................
      //    Step 5. Calculate the Figured From Date. This will be
      //            used to calculate the number of days of interest
      //            to pay.
      // ................................................................
      if (msiStateIn.CalcIdtypCd.toUpperCase().equals("PROOF   ")) {
        msiStateIn.FiguredFromDate = DateValue(mfrmMyInsuredForm.frmInsured.getInsuredClmProofDt()) + msiStateIn.StrlIntCalcOfstNum;
      } 
      else {
        msiStateIn.FiguredFromDate = DateValue(mfrmMyInsuredForm.frmInsured.getInsuredClmInsdDthDt()) + msiStateIn.StrlIntCalcOfstNum;
      }

      // ................................................................
      //    Step 6. Calculate the number of days to pay interest on
      //            the claim.
      //            NOTE: If this is set to 0, then no interest or
      //            withholding will be paid when 0 is plugged into
      //            the formulaes in Steps 8 and 9.
      // ................................................................
      msiStateIn.NbrOfDaysToPayInterest = DateDiff("d", msiStateIn.FiguredFromDate, dteDateOfPayment);

      // Per bug 2455, floor the NbrOfDaysToPayInterest so negative numbers are turned into 0.
      msiStateIn.NbrOfDaysToPayInterest = modDataConversion.fnAtLeast(msiStateIn.NbrOfDaysToPayInterest, cdblFloorOfZero);

      msiStateIn.CalculationInfo = msiStateIn.CalculationInfo+ "The # of days ("+ msiStateIn.NbrOfDaysToPayInterest+ ") was based on "+ CStr(msiStateIn.FiguredFromDate)+ " to "+ CStr(dteDateOfPayment)+ ". ";

      // ................................................................
      //    Step 7. Determine the interest rate to use to calculate
      //            Claims Interest
      // ................................................................
      if (bUseSuppliedRate) {
        // Do nothing...the InterestRateToUse was previously set
      } 
      else {
        msiStateIn.InterestRateToUse = fnGetInterestRate(msiStateIn, strDesc);
      }


      // **TODO:** label found: STEP8:;
      // ................................................................
      //    Step 8. Calculate the Claims Interest Amount, rounded to
      //            2 decimal positions
      // ................................................................
      msiStateIn.ClaimInterestAmt = (Double.parseDouble(ipcPayeDthbPmtAmt.Text) * (msiStateIn.InterestRateToUse / 100));
      msiStateIn.ClaimInterestAmt = msiStateIn.ClaimInterestAmt * (msiStateIn.NbrOfDaysToPayInterest / 365);
      msiStateIn.ClaimInterestAmt = Round(msiStateIn.ClaimInterestAmt, 2);

      // ................................................................
      //    Step 9. Calculate the Withheld Amount, rounded to
      //            2 decimal positions. If the Claim Interest Amount
      //             is zero, then the Withheld Amount will be zero.
      // ................................................................
      msiStateIn.WithheldAmt = msiStateIn.ClaimInterestAmt * (Double.parseDouble(ipdPayeWthldRt.Text) / 100);
      msiStateIn.WithheldAmt = Round(msiStateIn.WithheldAmt, 2);

      // ................................................................
      //    Step 10. Calculate the Total Amount to be paid for this Payee.
      // ................................................................
      msiStateIn.TotalForThisPayee = Double.parseDouble(ipcPayeDthbPmtAmt.Text) + msiStateIn.ClaimInterestAmt - msiStateIn.WithheldAmt;

      _rtn = true;
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
  private void fnClearControls() {
    // Comments  : Initializes screen controls in order to add a new record
    // Parameters: None
    // Called by : fnAddRecord of frmPayee
    // Modified  :
    // --------------------------------------------------
    try {
      "fnClearControls"
.equals(Const cstrCurrentProc As String);
      Const(cintZero As Integer == 0);
      Control ctl = null;
      Object varDefaultValue = null;
      String strSavedMask = "";

      // Hide updates to the window until we're done. This avoids ugly screen flickering
      fnWindowLock(Me.cbrfBrowseFolder.setHWnd());

      iptPayeFullNm.Text = "";
      iptPayeCareOfTxt.Text = "";
      iptPayeAddrLn1Txt.Text = "";
      iptPayeAddrLn2Txt.Text = "";
      iptPayeCityNmTxt.Text = "";

      if (cboCalcStCd.ListCount > 0) {
        //' Select first (blank) entry
        cboCalcStCd.ListIndex = 0;
      } 
      else {
        modGeneral.gerhApp.reportNonFatal(vbObjectError + modResConstants.gCRES_WARN_CBO_IS_EMPTY, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc, MCSTRCBOCALCSTCDLABEL);
      }

      iptPayeZipCd.Text = "";
      iptPayeZip4Cd.Text = "";

      // NOTE: For MaskEdBox and fpMask controls, have to remove mask before clearing out the control
      //       since the vbNullString value doesn't match the mask specification.
      strSavedMask = ipmPayeSsnTinNum.ctclmClaim.getMask();
      ipmPayeSsnTinNum.ctclmClaim.setMask("");
      ipmPayeSsnTinNum.Text = "";
      ipmPayeSsnTinNum.ctclmClaim.setMask(strSavedMask);

      if (cboPayeStCd.ListCount > 0) {
        //' Select first (blank) entry
        cboPayeStCd.ListIndex = 0;
      } 
      else {
        modGeneral.gerhApp.reportNonFatal(vbObjectError + modResConstants.gCRES_WARN_CBO_IS_EMPTY, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc, MCSTRCBOPAYESTCDLABEL);
      }

      ipdPayeWthldRt.chrgHourglass.setValue(cintZero);


      // ------------------------------------------------------------------------
      //    The checkPayeDfltOvrdInd, cboCalcStCd and iptPayeClmIntRt and
      //   lblWarningAboutOverride controls are all tied to one another with
      //   regard to their availability and initialization.
      // ------------------------------------------------------------------------
      // Set their values
      chkPayeDfltOvrdInd.chrgHourglass.setValue(vbUnchecked);

      if (cboCalcStCd.ListCount > 0) {
        //' Select first (blank) entry
        cboCalcStCd.ListIndex = 0;
      } 
      else {
        modGeneral.gerhApp.reportNonFatal(vbObjectError + modResConstants.gCRES_WARN_CBO_IS_EMPTY, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc, MCSTRCBOCALCSTCDLABEL);
      }

      ipdPayeClmIntRt.chrgHourglass.setValue(cintZero);

      // Set their availability
      //'' BZ4999 October 2013 Non US payee - SXS
      modGeneral.fnEnableDisableControl(ctlIn:=ChkPaye1099Ind, bEnable:=True);
      modGeneral.fnEnableDisableControl(ctlIn:=chkPayeDfltOvrdInd, bEnable:=False);
      modGeneral.fnEnableDisableControl(ctlIn:=cboCalcStCd, bEnable:=False);
      modGeneral.fnEnableDisableControl(ctlIn:=ipdPayeClmIntRt, bEnable:=False);
      lblWarningAboutOverride.Visible = false;
      // ------------------------------------------------------------------------

      // DateTimePicker controls (dtpPayePmtDt) will
      // automatically be set to today's date. Cannot set them to Null
      // unless their CheckBox property is set to True.
      dtpPayePmtDt.chrgHourglass.setValue(Date);
      fnResetStateRules();

      ipdPayeIntDaysPdNum.chrgHourglass.setValue(cintZero);
      ipcPayeDthbPmtAmt.chrgHourglass.setValue(cintZero);
      ipcPayeClmIntAmt.chrgHourglass.setValue(cintZero);
      ipcPayeWthldAmt.chrgHourglass.setValue(cintZero);
      ipcPayeClmPdAmt.chrgHourglass.setValue(cintZero);

      // intitialize the label that describes how the calculation was done.
      lblCalculationInfo = "";

      // Initialize fields that will be set when calculation is done
      txtCalcStCdSpecialInstructions_UsedInAutoCalc = "";
      //' non-displayed version of ipcPayeWthldAmt
      mcurTotalWithheld = 0;

      // Skip initialization of Insured Residence State and Insured Residence State's
      // Special Instructions since these should not change from payee to payee as they
      // are based on the Insured screen to which all payees belong.

      // Skip initialization of Payee Residence State and Payee Residence State's
      // Special Instructions since they are set when the cboPayeStCd is changed
      // (as occurred when its ListIndex was set to 0 above).
      //       txtPayeStCd_UsedInAutoCalc = vbNullString
      //       txtPayeStCdSpecialInstructions_UsedInAutoCalc = vbNullString

      // Skip initialization of Contract Issue State's Special Instructions since
      // it is set when the Contract Issue State changed (as occurred when its
      // ListIndex was set to 0 above) .
      //       txtIssStCdSpecialInstructions_UsedInAutoCalc = vbNullString


      // Next, set each control to its default value per the meta data, if available.
      for (int _i = 0; _i < Me.Controls.size(); _i++) {
        ctl = Me.Controls.item(_i);
        // Debug.Print ctl.Name & vbTab & ctl.Tag
        // If control corresponds to a SQL Server table column, then try
        // to set its default properties. The Tag property contains
        // the name of its property within the table class.
        if (ctl.Tag.length() > 0) {
          // If there's a default value, use it
          varDefaultValue = mtWrapper.getDefaultValue(ctl.Tag);
          if (!(IsEmpty(varDefaultValue))) {
            if ((TypeOf ctl Is TextBox) || (TypeOf ctl Is fpText) || (TypeOf ctl Is ComboBox) || (TypeOf ctl Is ListBox)) {
              ctl.Text = varDefaultValue;
            } 
            else if ((TypeOf ctl Is fpCurrency)) {
              ctl.value = varDefaultValue;
              //ElseIf (TypeOf ctl Is MaskEdBox) Then
              //    .SelText = varDefaultValue
            } 
            else if ((TypeOf ctl Is fpMask) || (TypeOf ctl Is fpDoubleSingle)) {
              ctl.UnFmtText = varDefaultValue;
            } 
            else if (ctl(instanceOf CheckBox)) {
              // Bug thinks the default value is "Y" or "N" when really it's True or False
              if (varDefaultValue.toUpperCase().equals(true)) {
                ctl.value = vbChecked;
              } 
              else {
                ctl.value = vbUnchecked;
              }
            } 
            else if (ctl(instanceOf Label)) {
              ctl.Caption = varDefaultValue;
            } 
            else if (ctl(instanceOf fpCombo)) {
              modComboBox.fnSearchFPCombo(lpcIn:=ctl, strSearchText:=varDefaultValue, intSearchCol:=1);
            }
          }
        }
      }
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    fnWindowUnlock;

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
  private DBRecordSet fnGetListOfStates() {
    //--------------------------------------------------------------------------
    // Procedure:   fnGetListOfStates
    // Description: Builds an ADODB.Recordset containing state codes in STATE_T
    // Params:      N/A
    // Returns:     N/A
    // Modified:
    //-----------------------------------------------------------------------------
    "fnGetListOfStates"
.equals(Const cstrCurrentProc As String);
    //' Stored procedure to execute
    "dbo.proc_state_lu_select"
.equals(Const cstrSproc As String);
    ADODB.Parameter prmReturnValue = null;
    cadwADOWrapper adwTemp = null;

    try {

      adwTemp = new cadwADOWrapper();

      if (!(adwTemp.commandSetSproc(cstrSproc))) {
        // **TODO:** goto found: GoTo PROC_EXIT;
      }

      ADODB.Command w_aDOCommand = adwTemp.getADOCommand();
      // ---Parameter #1---
      // Define the return value that represents the error code (i.e. reason) why
      // the stored procedure failed.
      prmReturnValue = w_aDOCommand.CreateParameter(Name:="@return_value", Type:=adInteger, Direction:=adParamReturnValue, chrgHourglass.getValue():=Null);
      w_aDOCommand.Parameters.Append(prmReturnValue);

      return w_aDOCommand.Execute();
      // Do not close this recordset.
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    // Do not free the fnGetListOfStates recordset!
    modGeneral.fnFreeObject(adwTemp);
    modGeneral.fnFreeObject(prmReturnValue);

    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
    }
    return null;
    // **TODO:** label found: PROC_ERR:;
    switch (prmReturnValue) {
      //' 4028
      case  modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO  :
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc, "locate");
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
        break;
    }

    // If any other errors exist, i.e. in Err object, then let it fall through into default error handling.

    switch (VBA.ex.Number) {
      //' Object not found
      case  -2147217900:
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_FERR_SPROC_NOT_FOUND, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc, cstrSproc);
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
        break;

      default:
        modGeneral.gerhApp.saveErrObjectData(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }
}


//////////////////////////////////////////////////////////////////////////////////////////////////
  private double fnGetCurrentIntRate(Date dtePayePmtDt) {
    double _rtn = 0;
    //--------------------------------------------------------------------------
    // Procedure:   fnGetCurrentIntRate
    // Description: retrieves the Current Rate in effect on the Date of Payment
    // Params:      N/A
    // Returns:     double, representing the Current Rate
    // Modified:
    //-----------------------------------------------------------------------------
    "fnGetCurrentIntRate"
.equals(Const cstrCurrentProc As String);
    //' Stored procedure to execute
    "dbo.proc_current_rate_select"
.equals(Const cstrSproc As String);
    ADODB.Parameter prmReturnValue = null;
    ADODB.Parameter prmPayePmtDt = null;
    ADODB.Parameter prmCurrIntRt = null;
    cadwADOWrapper adwTemp = null;
    DBRecordSet rstTemp = null;

    try {

      adwTemp = new cadwADOWrapper();

      if (!(adwTemp.commandSetSproc(cstrSproc))) {
        // **TODO:** goto found: GoTo PROC_EXIT;
      }

      ADODB.Command w_aDOCommand = adwTemp.getADOCommand();
      // ---Parameter #1---
      // Define the return value that represents the error code (i.e. reason) why
      // the stored procedure failed.
      prmReturnValue = w_aDOCommand.CreateParameter(Name:="@return_value", Type:=adInteger, Direction:=adParamReturnValue, chrgHourglass.getValue():=Null);
      w_aDOCommand.Parameters.Append(prmReturnValue);

      // ---Parameter #2---
      prmPayePmtDt = w_aDOCommand.CreateParameter(Name:="@paye_pmt_dt", Type:=adDBTimeStamp, Direction:=adParamInput, Size:=16, chrgHourglass.getValue():=dtePayePmtDt);
      w_aDOCommand.Parameters.Append(prmPayePmtDt);

      // ---Parameter #3---
      prmCurrIntRt = w_aDOCommand.CreateParameter(Name:="@curr_int_rt", Type:=adNumeric, Direction:=adParamOutput, chrgHourglass.getValue():=Null);
      w_aDOCommand.Parameters.Append(prmCurrIntRt);
      // Have to hard-code the precision/scale since we have no meta data for this table
      prmCurrIntRt.Precision = 11;
      prmCurrIntRt.NumericScale = 5;

      rstTemp = w_aDOCommand.Execute();

      if (prmCurrIntRt.value == null) {
        //' -1
        _rtn = modComboBox.gCLNGNOSELECTION;
      } 
      else {
        _rtn = prmCurrIntRt.value;
      }
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    modGeneral.fnFreeObject(adwTemp);
    modGeneral.fnFreeRecordset(rstTemp);
    modGeneral.fnFreeObject(prmReturnValue);
    modGeneral.fnFreeObject(prmPayePmtDt);
    modGeneral.fnFreeObject(prmCurrIntRt);

    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (prmReturnValue) {
      case  modResConstants.gCRES_NERR_REC_NOT_FOUND:
        // 4027 = The specified record was not found in the database (@@1).
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_REC_NOT_FOUND, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc, "Current Rate effective on ["+ FormatDateTime(dtePayePmtDt, vbShortDate)+ "]");
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
      //' 4028
        break;

      case  modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO  :
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc, "locate");
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
        break;
    }

    // If any other errors exist, i.e. in Err object, then let it fall through into default error handling.

    switch (VBA.ex.Number) {
      //' Object not found
      case  -2147217900:
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_FERR_SPROC_NOT_FOUND, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc, cstrSproc);
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
        break;

      default:
        modGeneral.gerhApp.saveErrObjectData(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}


//////////////////////////////////////////////////////////////////////////////////////////////////
  private String fnGetFieldLabel(String strControlName) {
    String _rtn = "";
    //--------------------------------------------------------------------------
    // Procedure:   fnGetFieldLabel
    // Description: Given a control name, return the value of the control's label
    //
    // Params:      N/A
    //    strControlName  (in) A string containing the control's name
    //
    // Returns:     A string containing the controls' label.
    //-----------------------------------------------------------------------------

    //!CUSTOMIZE!  There should be one Case statement for each control that
    //             corresponds to a table column. Each Case statement should
    //             reference a Const literal that indicates how the control is
    //             labelled on-screen.

    "fnGetFieldLabel"
.equals(Const cstrCurrentProc As String);

    try {

      switch (strControlName) {
        case  "iptPayeFullNm":
          _rtn = MCSTRIPTPAYEFULLNMLABEL;
          break;

        case  "iptPayeCareOfTxt":
          _rtn = MCSTRIPTPAYECAREOFTXTLABEL;
          break;

        case  "iptPayeAddrLn1Txt":
          _rtn = MCSTRIPTPAYEADDRLN1TXTLABEL;
          break;

        case  "iptPayeAddrLn2Txt":
          _rtn = MCSTRIPTPAYEADDRLN2TXTLABEL;
          break;

        case  "iptPayeCityNmTxt":
          _rtn = MCSTRIPTPAYECITYNMTXTLABEL;
          break;

        case  "cboPayeStCd":
          _rtn = MCSTRCBOPAYESTCDLABEL;
          break;

        case  "iptPayeZipCd":
          _rtn = MCSTRIPTPAYEZIPCDLABEL;
          break;

        case  "iptPayeZip4Cd":
          _rtn = MCSTRIPTPAYEZIP4CDLABEL;
          break;

        case  "ChkPaye1099Ind":
          //'' BZ4999 October 2013 Non US payee - SXS
          _rtn = MCSTRCHKPAYE1099INDLABEL;
          break;

        case  "ipmPayeSsnTinNum":
          _rtn = MCSTRIPMPAYESSNTINNUMLABEL;
          break;

        case  "cboPayeSsnTinTypCd":
          _rtn = MCSTRCBOPAYESSNTINTYPCDLABEL;
          break;

        case  "ipdPayeWthldRt":
          _rtn = MCSTRIPDPAYEWTHLDRTLABEL;
          break;

        case  "ipdPayeClmIntRt":
          _rtn = MCSTRIPDPAYECLMINTRTLABEL;
          break;

        case  "chkPayeDfltOvrdInd":
          _rtn = MCSTRCHKPAYEDFLTOVRDINDLABEL;
          break;

        case  "cboCalcStCd":
          _rtn = MCSTRCBOCALCSTCDLABEL;
          break;

        case  "dtpPayePmtDt":
          _rtn = MCSTRDTPPAYEPMTDTLABEL;
          break;

        case  "ipdPayeIntDaysPdNum":
          _rtn = MCSTRIPDPAYEINTDAYSPDNUMLABEL;
          break;

        case  "ipcPayeDthbPmtAmt":
          _rtn = MCSTRIPCPAYEDTHBPMTAMTLABEL;
          break;

        case  "ipcPayeClmIntAmt":
          _rtn = MCSTRIPCPAYECLMINTAMTLABEL;
          break;

        case  "ipcPayeWthldAmt":
          _rtn = MCSTRIPCPAYEWTHLDAMTLABEL;
          break;

        case  "ipcPayeClmPdAmt":
          _rtn = MCSTRIPCPAYECLMPDAMTLABEL;
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
  private double fnGetInterestRate(StateInfo siRatesIn, String strDesc) { // TODO: Use of ByRef founded Private Function fnGetInterestRate(ByRef siRatesIn As StateInfo, ByVal strDesc As String) As Double
    double _rtn = 0;
    // Comments  : This function will return the interest
    //             rate. This rate is taken from the State98
    //             table directly if it is numeric. Otherwise
    //             the user is prompted for it.
    // Parameters: siRatesIn (in) - a StateInfo structure that
    //                              contains pertinent info from a row in the
    //                              STATE_RULE_T table.
    //             strDesc (in)   - Descriptive text that indicates why a rate is needed (to appear in prompts)
    // Returns   : The interest rate as a Double
    // Modified  :
    // --------------------------------------------------
    try {
      "fnGetInterestRate"
.equals(Const cstrCurrentProc As String);
      Const(cintMaxInterestRate As Integer == 12);
      "Please specify the "
.equals(Const cstrPleaseSpecify As String);
      " in effect on "
.equals(Const cstrInEffectOn As String);
      " for the "
.equals(Const cstrForThe As String);
      " of "
.equals(Const cstrOf As String);
      "Current Loan Rate"
.equals(Const cstrCurrLoanRt As String);
      "The derived Rate based on "
.equals(Const cstrDerivedRateBasedOn As String);
      "The Current Rate in effect on "
.equals(Const cstrCurrRtInEffectOn As String);
      " is not defined. Please supply the Rate to use "
.equals(Const cstrIsNotDefdSupplyARate As String);
      " is a negative number. Please supply the Rate to use "
.equals(Const cstrIsNegNbrSupplyRateToUse As String);
      "the Current Loan Rate("
.equals(Const cstrCurrLoanRtX As String);
      "the Current Rate("
.equals(Const cstrCurrRtX As String);
      ") - "
.equals(Const cstrCloseParenMinus As String);
      ") + "
.equals(Const cstrCloseParenPlus As String);
      "the State Rule Amount("
.equals(Const cstrStateRuleAmtX As String);
      ") "
.equals(Const cstrCloseParen As String);
      "the maximum of "
.equals(Const cstrTheMaxOf As String);
      "the minimum of "
.equals(Const cstrTheMinOf As String);
      "the greater of "
.equals(Const cstrTheGreaterOf As String);
      " and "
.equals(Const cstrAnd As String);
      double dblCurrIntRt = 0;
      double dblCurrLoanRt = 0;
      String strPayePmtDt = "";
      double dblRateToUse = 0;

      strPayePmtDt = FormatDateTime(dtpPayePmtDt.chrgHourglass.getValue(), vbShortDate);

      // **TODO:** label found: DETERMINE_RATE:;
      switch (siRatesIn.IruleCd) {
        case  "CLNW/MAX":
          dblCurrLoanRt = fnPromptForRate(cstrPleaseSpecify+ cstrCurrLoanRt+ cstrInEffectOn+ strPayePmtDt+ cstrForThe+ strDesc+ cstrOf+ siRatesIn.StCd);
          if (dblCurrLoanRt > siRatesIn.StrlIntRuleAmt) {
            dblRateToUse = siRatesIn.StrlIntRuleAmt;
          } 
          else {
            dblRateToUse = dblCurrLoanRt;
          }
          if (dblRateToUse < 0) {
            // If the derived rate is negative, ask the user to supply it
            dblRateToUse = fnPromptForRate(cstrDerivedRateBasedOn+ cstrTheMaxOf+ cstrCurrLoanRtX+ CStr(dblCurrLoanRt)+ cstrCloseParen+ cstrAnd+ cstrStateRuleAmtX+ CStr(siRatesIn.StrlIntRuleAmt)+ cstrCloseParen+ cstrIsNegNbrSupplyRateToUse+ cstrForThe+ strDesc+ cstrOf+ siRatesIn.StCd);
          }

          break;

        case  "CLNW/MIN":
          dblCurrLoanRt = fnPromptForRate(cstrPleaseSpecify+ cstrCurrLoanRt+ cstrInEffectOn+ strPayePmtDt+ cstrForThe+ strDesc+ cstrOf+ siRatesIn.StCd);
          if (dblCurrLoanRt < siRatesIn.StrlIntRuleAmt) {
            dblRateToUse = siRatesIn.StrlIntRuleAmt;
          } 
          else {
            dblRateToUse = dblCurrLoanRt;
          }
          if (dblRateToUse < 0) {
            // If the derived rate is negative, ask the user to supply it
            dblRateToUse = fnPromptForRate(cstrDerivedRateBasedOn+ cstrTheMinOf+ cstrCurrLoanRtX+ CStr(dblCurrLoanRt)+ cstrCloseParen+ cstrAnd+ cstrStateRuleAmtX+ CStr(siRatesIn.StrlIntRuleAmt)+ cstrCloseParen+ cstrIsNegNbrSupplyRateToUse+ cstrForThe+ strDesc+ cstrOf+ siRatesIn.StCd);
          }

          break;

        case  "CRTW/MAX":
          dblCurrIntRt = fnGetCurrentIntRate(dtpPayePmtDt.chrgHourglass.getValue());
          if (dblCurrIntRt == modComboBox.gCLNGNOSELECTION) {
            // If Current Rate was not found, ask the user to supply it
            dblCurrIntRt = fnPromptForRate(cstrCurrRtInEffectOn+ strPayePmtDt+ cstrIsNotDefdSupplyARate+ cstrForThe+ strDesc+ cstrOf+ siRatesIn.StCd);
          }
          if (dblCurrIntRt > siRatesIn.StrlIntRuleAmt) {
            dblRateToUse = siRatesIn.StrlIntRuleAmt;
          } 
          else {
            dblRateToUse = dblCurrIntRt;
          }
          if (dblRateToUse < 0) {
            // If the derived rate is negative, ask the user to supply it
            dblRateToUse = fnPromptForRate(cstrDerivedRateBasedOn+ cstrTheMaxOf+ cstrCurrRtX+ CStr(dblCurrIntRt)+ cstrCloseParen+ cstrAnd+ cstrStateRuleAmtX+ CStr(siRatesIn.StrlIntRuleAmt)+ cstrCloseParen+ cstrIsNegNbrSupplyRateToUse+ cstrForThe+ strDesc+ cstrOf+ siRatesIn.StCd);
          }

          break;

        case  "CRTW/MIN":
          dblCurrIntRt = fnGetCurrentIntRate(dtpPayePmtDt.chrgHourglass.getValue());
          // If Current Rate was not found, ask the user to supply it
          if (dblCurrIntRt == modComboBox.gCLNGNOSELECTION) {
            dblCurrIntRt = fnPromptForRate(cstrCurrRtInEffectOn+ strPayePmtDt+ cstrIsNotDefdSupplyARate+ cstrForThe+ strDesc+ cstrOf+ siRatesIn.StCd);
          }
          if (dblCurrIntRt < siRatesIn.StrlIntRuleAmt) {
            dblRateToUse = siRatesIn.StrlIntRuleAmt;
          } 
          else {
            dblRateToUse = dblCurrIntRt;
          }
          if (dblRateToUse < 0) {
            // If the derived rate is negative, ask the user to supply it
            dblRateToUse = fnPromptForRate(cstrDerivedRateBasedOn+ cstrTheMinOf+ cstrCurrRtX+ CStr(dblCurrIntRt)+ cstrCloseParen+ cstrAnd+ cstrStateRuleAmtX+ CStr(siRatesIn.StrlIntRuleAmt)+ cstrCloseParen+ cstrIsNegNbrSupplyRateToUse+ cstrForThe+ strDesc+ cstrOf+ siRatesIn.StCd);
          }

          break;

        case  "CURLN   ":
          dblRateToUse = fnPromptForRate(cstrPleaseSpecify+ cstrCurrLoanRt+ cstrInEffectOn+ strPayePmtDt+ cstrForThe+ strDesc+ cstrOf+ siRatesIn.StCd);

          break;

        case  "CURLN+X ":
          dblCurrLoanRt = fnPromptForRate(cstrPleaseSpecify+ cstrCurrLoanRt+ cstrInEffectOn+ strPayePmtDt+ cstrForThe+ strDesc+ cstrOf+ siRatesIn.StCd);
          dblRateToUse = dblCurrLoanRt + siRatesIn.StrlIntRuleAmt;
          if (dblRateToUse < 0) {
            // If the derived rate is negative, ask the user to supply it
            dblRateToUse = fnPromptForRate(cstrDerivedRateBasedOn+ cstrCurrLoanRtX+ CStr(dblCurrIntRt)+ cstrCloseParenPlus+ cstrStateRuleAmtX+ CStr(siRatesIn.StrlIntRuleAmt)+ cstrCloseParen+ cstrIsNegNbrSupplyRateToUse+ cstrForThe+ strDesc+ cstrOf+ siRatesIn.StCd);
          }

          break;

        case  "CURLN-X ":
          dblCurrLoanRt = fnPromptForRate(cstrPleaseSpecify+ cstrCurrLoanRt+ cstrInEffectOn+ strPayePmtDt+ cstrForThe+ strDesc+ cstrOf+ siRatesIn.StCd);
          dblRateToUse = dblCurrLoanRt - siRatesIn.StrlIntRuleAmt;
          if (dblRateToUse < 0) {
            // If the derived rate is negative, ask the user to supply it
            dblRateToUse = fnPromptForRate(cstrDerivedRateBasedOn+ cstrCurrLoanRtX+ CStr(dblCurrLoanRt)+ cstrCloseParenMinus+ cstrStateRuleAmtX+ CStr(siRatesIn.StrlIntRuleAmt)+ cstrCloseParen+ cstrIsNegNbrSupplyRateToUse+ cstrForThe+ strDesc+ cstrOf+ siRatesIn.StCd);
          }

          break;

        case  "CURRT   ":
          dblCurrIntRt = fnGetCurrentIntRate(dtpPayePmtDt.chrgHourglass.getValue());
          if (dblCurrIntRt == modComboBox.gCLNGNOSELECTION) {
            dblCurrIntRt = fnPromptForRate(cstrCurrRtInEffectOn+ strPayePmtDt+ cstrIsNotDefdSupplyARate+ cstrForThe+ strDesc+ cstrOf+ siRatesIn.StCd);
          } 
          else if (dblCurrIntRt < 0) {
            // If Current Rate was found but is negative, ask the user to supply it
            dblCurrIntRt = fnPromptForRate(cstrCurrRtInEffectOn+ strPayePmtDt+ cstrIsNegNbrSupplyRateToUse+ cstrForThe+ strDesc+ cstrOf+ siRatesIn.StCd);
          }
          dblRateToUse = dblCurrIntRt;

          break;

        case  "CURRT+X ":
          dblCurrIntRt = fnGetCurrentIntRate(dtpPayePmtDt.chrgHourglass.getValue());
          if (dblCurrIntRt == modComboBox.gCLNGNOSELECTION) {
            dblCurrIntRt = fnPromptForRate(cstrCurrRtInEffectOn+ strPayePmtDt+ cstrIsNotDefdSupplyARate+ cstrForThe+ strDesc+ cstrOf+ siRatesIn.StCd);
          }
          dblRateToUse = dblCurrIntRt + siRatesIn.StrlIntRuleAmt;
          if (dblRateToUse < 0) {
            // If the derived rate is negative, ask the user to supply it
            dblRateToUse = fnPromptForRate(cstrDerivedRateBasedOn+ cstrCurrRtX+ CStr(dblCurrIntRt)+ cstrCloseParenPlus+ cstrStateRuleAmtX+ CStr(siRatesIn.StrlIntRuleAmt)+ cstrCloseParen+ cstrIsNegNbrSupplyRateToUse+ cstrForThe+ strDesc+ cstrOf+ siRatesIn.StCd);
          }

          break;

        case  "CURRT-X ":
          dblCurrIntRt = fnGetCurrentIntRate(dtpPayePmtDt.chrgHourglass.getValue());
          if (dblCurrIntRt == modComboBox.gCLNGNOSELECTION) {
            dblCurrIntRt = fnPromptForRate(cstrCurrRtInEffectOn+ strPayePmtDt+ cstrIsNotDefdSupplyARate+ cstrForThe+ strDesc+ cstrOf+ siRatesIn.StCd);
          }
          dblRateToUse = dblCurrIntRt - siRatesIn.StrlIntRuleAmt;
          if (dblRateToUse < 0) {
            // If the derived rate is negative, ask the user to supply it
            dblRateToUse = fnPromptForRate(cstrDerivedRateBasedOn+ cstrCurrRtX+ CStr(dblCurrIntRt)+ cstrCloseParenMinus+ cstrStateRuleAmtX+ CStr(siRatesIn.StrlIntRuleAmt)+ cstrCloseParen+ cstrIsNegNbrSupplyRateToUse+ cstrForThe+ strDesc+ cstrOf+ siRatesIn.StCd);
          }

          break;

        case  "GTCLN&X ":
          dblCurrLoanRt = fnPromptForRate(cstrPleaseSpecify+ cstrCurrLoanRt+ cstrInEffectOn+ strPayePmtDt+ cstrForThe+ strDesc+ cstrOf+ siRatesIn.StCd);
          if (dblCurrLoanRt > siRatesIn.StrlIntRuleAmt) {
            dblRateToUse = dblCurrLoanRt;
          } 
          else {
            dblRateToUse = siRatesIn.StrlIntRuleAmt;
          }
          if (dblRateToUse < 0) {
            // If the derived rate is negative, ask the user to supply it
            dblRateToUse = fnPromptForRate(cstrDerivedRateBasedOn+ cstrTheGreaterOf+ cstrCurrLoanRtX+ CStr(dblCurrLoanRt)+ cstrCloseParen+ cstrAnd+ cstrStateRuleAmtX+ CStr(siRatesIn.StrlIntRuleAmt)+ cstrCloseParen+ cstrIsNegNbrSupplyRateToUse+ cstrForThe+ strDesc+ cstrOf+ siRatesIn.StCd);
          }

          break;

        case  "GTCRT&LN":
          dblCurrIntRt = fnGetCurrentIntRate(dtpPayePmtDt.chrgHourglass.getValue());
          if (dblCurrIntRt == modComboBox.gCLNGNOSELECTION) {
            dblCurrIntRt = fnPromptForRate(cstrCurrRtInEffectOn+ strPayePmtDt+ cstrIsNotDefdSupplyARate+ cstrForThe+ strDesc+ cstrOf+ siRatesIn.StCd);
          }
          dblCurrLoanRt = fnPromptForRate(cstrPleaseSpecify+ cstrCurrLoanRt+ cstrInEffectOn+ strPayePmtDt+ cstrForThe+ strDesc+ cstrOf+ siRatesIn.StCd);
          if (dblCurrIntRt > dblCurrLoanRt) {
            dblRateToUse = dblCurrIntRt;
          } 
          else {
            dblRateToUse = dblCurrLoanRt;
          }
          if (dblRateToUse < 0) {
            // If the derived rate is negative, ask the user to supply it
            dblRateToUse = fnPromptForRate(cstrDerivedRateBasedOn+ cstrTheGreaterOf+ cstrCurrRtX+ CStr(dblCurrIntRt)+ cstrCloseParen+ cstrAnd+ cstrCurrLoanRtX+ CStr(dblCurrLoanRt)+ cstrCloseParen+ cstrIsNegNbrSupplyRateToUse+ cstrForThe+ strDesc+ cstrOf+ siRatesIn.StCd);
          }

          break;

        case  "GTCRT&X ":
          dblCurrIntRt = fnGetCurrentIntRate(dtpPayePmtDt.chrgHourglass.getValue());
          if (dblCurrIntRt == modComboBox.gCLNGNOSELECTION) {
            dblCurrIntRt = fnPromptForRate(cstrCurrRtInEffectOn+ strPayePmtDt+ cstrIsNotDefdSupplyARate+ cstrForThe+ strDesc+ cstrOf+ siRatesIn.StCd);
          }
          if (dblCurrIntRt > siRatesIn.StrlIntRuleAmt) {
            dblRateToUse = dblCurrIntRt;
          } 
          else {
            dblRateToUse = siRatesIn.StrlIntRuleAmt;
          }
          if (dblRateToUse < 0) {
            // If the derived rate is negative, ask the user to supply it
            dblRateToUse = fnPromptForRate(cstrDerivedRateBasedOn+ cstrTheGreaterOf+ cstrCurrRtX+ CStr(dblCurrIntRt)+ cstrCloseParen+ cstrAnd+ cstrStateRuleAmtX+ CStr(siRatesIn.StrlIntRuleAmt)+ cstrCloseParen+ cstrIsNegNbrSupplyRateToUse+ cstrForThe+ strDesc+ cstrOf+ siRatesIn.StCd);
          }

          break;

        case  "PROMPT  ":
          dblRateToUse = fnPromptForRate(cstrPleaseSpecify+ "Interest Rate to use"+ cstrForThe+ strDesc+ cstrOf+ siRatesIn.StCd);

          break;

        case  "SPECAMT ":
          dblRateToUse = siRatesIn.StrlIntRuleAmt;
          if (dblRateToUse < 0) {
            // If the derived rate is negative, ask the user to supply it
            dblRateToUse = fnPromptForRate(cstrDerivedRateBasedOn+ cstrStateRuleAmtX+ CStr(siRatesIn.StrlIntRuleAmt)+ cstrCloseParen+ cstrIsNegNbrSupplyRateToUse+ cstrForThe+ strDesc+ cstrOf+ siRatesIn.StCd);
          }

          break;

        default:
          modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_UNEXPECTED_VAL_SELECT_CASE, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
          // **TODO:** goto found: GoTo PROC_EXIT;
          break;
      }

      // Per Michelle Wilkosky, the following check should be performed against all interest rates, whether user-supplied,
      // calculated, or obtained from the STATE_RULE_T table, and whether it represents an interest rate or loan rate.
      if (dblRateToUse < 0) {
        // (gcRES_WARN_RATE_IS_NEGATIVE (2007) = The Rate supplied or derived from the supplied Rate
        //                                       is a negative number (@@1). Please try again.
        modGeneral.gerhApp.reportNonFatal(vbObjectError + modResConstants.gCRES_WARN_RATE_IS_NEGATIVE, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc, CStr(dblRateToUse));
        // **TODO:** goto found: GoTo DETERMINE_RATE;
      }

      // MME WRUS 4999 - Per Dave O'Connor - the max interest rate check is no longer needed
      // Per Michelle Wilkosky, the following check should be performed against all interest rates, whether user-supplied,
      // calculated, or obtained from the STATE_RULE_T table, and whether it represents an interest rate or loan rate.
      if (G.isNumeric(dblRateToUse)) {
        //If Val(dblRateToUse) > cintMaxInterestRate And siRatesIn.StCd <> "ME" Then
        //    ' 4005 = The interest rate supplied is more than @@1%. This is only allowed when the @@2 is Maine. Please try again.
        //    gerhApp.ReportNonFatal vbObjectError + gcRES_NERR_INTEREST_RATE_TOO_HIGH, _
        //                           mstrScreenName & gcstrDOT & cstrCurrentProc, _
        //                           CStr(cintMaxInterestRate), mcstrCboCalcStCdLabel
        //    GoTo DETERMINE_RATE
        //End If
      } 
      else {
        // gcRES_WARN_NONNUMERIC_RATE (2006) = The Rate supplied or obtained from the STATE_RULE_T table is non-numeric (@@1). Please try again.
        modGeneral.gerhApp.reportNonFatal(vbObjectError + modResConstants.gCRES_WARN_NONNUMERIC_RATE, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc, CStr(dblRateToUse));
        // **TODO:** goto found: GoTo DETERMINE_RATE;
      }

      _rtn = dblRateToUse;
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
  private void fnGetStateInfo_Override() {
    // Comments  : This function retrieves the calculation rule info
    //             for the Overriden Calc State and/or Interest Rate
    // Parameters:  -
    // Returns   :  -
    // Modified  :
    // --------------------------------------------------
    "fnGetStateInfo_Override"
.equals(Const cstrCurrentProc As String);

    try {

      if (cboCalcStCd.ListIndex > 0) {

        // MME START WRUS 4999 - ADDED EXTRA PARAMATERS

        modGeneral.fnGetStateInfo(cboCalcStCd.Text, mfrmMyInsuredForm.frmInsured.getInsuredLobCd(), DateValue(dtpPayePmtDt.chrgHourglass.getValue()), mfrmMyInsuredForm.frmInsured.getInsuredClmID(), dblScreenDBPaymentValue, msiOverride);
      } 
      else {
        modGeneral.fnInitializeStateInfo(msiOverride);
      }
      // Store Interest Rate in StateInfo structure too
      if (G.isNumeric(ipdPayeClmIntRt.UnFmtText)) {
        msiOverride.InterestRateToUse = ipdPayeClmIntRt.UnFmtText;
      } 
      else {
        // If the user cleared the contents, set the StateInfo structure and
        // screen field to 0
        msiOverride.InterestRateToUse = 0;
        ipdPayeClmIntRt.UnFmtText = 0;
      }
      txtCalcStCdSpecialInstructions_UsedInAutoCalc.Text = msiOverride.StrlSpclInstrTxt;
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
  private void fnGetStateInfo_InsdDthResStCd() {
    // Comments  : This function retrieves the calculation rule info
    //             for the Insured Residence State at time of death
    // Parameters:  -
    // Returns   :  -
    // Modified  :
    // --------------------------------------------------
    "fnGetStateInfo_InsdDthResStCd"
.equals(Const cstrCurrentProc As String);

    try {

      // If the Foreign Residence At Death checkbox is selected, there's no InsdDthResStCd so bypass
      // the call to fnGetStateInfo( ) and just initalize the StateInfo structure
      if (mfrmMyInsuredForm.frmInsured.getInsuredClmForResDthInd()) {
        chkClmForResDthInd_UsedInAutoCalc.chrgHourglass.setValue(vbChecked);
        modGeneral.fnInitializeStateInfo(msiInsdDthResStCd);
      } 
      else {
        chkClmForResDthInd_UsedInAutoCalc.chrgHourglass.setValue(vbUnchecked);

        // MME START WRUS 4999 - ADDED EXTRA PARAMATERS

        modGeneral.fnGetStateInfo(mfrmMyInsuredForm.frmInsured.getInsuredInsdDthResStCd(), mfrmMyInsuredForm.frmInsured.getInsuredLobCd(), DateValue(dtpPayePmtDt.chrgHourglass.getValue()), mfrmMyInsuredForm.frmInsured.getInsuredClmID(), dblScreenDBPaymentValue, msiInsdDthResStCd);
      }

      //*TODO:** can't found type for with block
      //*With msiInsdDthResStCd
      __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = msiInsdDthResStCd;
      txtInsdDthResStCd_UsedInAutoCalc.Text = w___TYPE_NOT_FOUND.StCd;
      txtInsdDthResStCdSpecialInstructions_UsedInAutoCalc.Text = w___TYPE_NOT_FOUND.StrlSpclInstrTxt;
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
  private void fnGetStateInfo_IssStCd() {
    // Comments  : This function retrieves the calculation rule info
    //             for the Contract Issue State
    // Parameters:  -
    // Returns   :  -
    // Modified  :
    // --------------------------------------------------
    "fnGetStateInfo_IssStCd"
.equals(Const cstrCurrentProc As String);

    try {

      // MME START WRUS 4999 - ADDED EXTRA PARAMATERS

      modGeneral.fnGetStateInfo(mfrmMyInsuredForm.frmInsured.getInsuredIssStCd(), mfrmMyInsuredForm.frmInsured.getInsuredLobCd(), DateValue(dtpPayePmtDt.chrgHourglass.getValue()), mfrmMyInsuredForm.frmInsured.getInsuredClmID(), dblScreenDBPaymentValue, msiIssStCd);
      //*TODO:** can't found type for with block
      //*With msiIssStCd
      __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = msiIssStCd;
      txtIssStCd_UsedInAutoCalc.Text = w___TYPE_NOT_FOUND.StCd;
      txtIssStCdSpecialInstructions_UsedInAutoCalc.Text = w___TYPE_NOT_FOUND.StrlSpclInstrTxt;
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
  private void fnGetStateInfo_PayeStCd() {
    // Comments  : This function retrieves the calculation rule info
    //             for the Payee Residence State at time of death
    // Parameters:  -
    // Returns   :  -
    // Modified  :
    // --------------------------------------------------
    "fnGetStateInfo_PayeStCd"
.equals(Const cstrCurrentProc As String);

    try {

      if (cboPayeStCd.ListIndex > 0) {

        // MME START WRUS 4999 - ADDED EXTRA PARAMATERS
        //''' '' BZ4999 October 2013 Non US payee - SXS
        if (!("ZZ".equals(cboPayeStCd.Text))) {
          modGeneral.fnGetStateInfo(cboPayeStCd.Text, mfrmMyInsuredForm.frmInsured.getInsuredLobCd(), DateValue(dtpPayePmtDt.chrgHourglass.getValue()), mfrmMyInsuredForm.frmInsured.getInsuredClmID(), dblScreenDBPaymentValue, msiPayeStCd);
        }
      } 
      else {
        modGeneral.fnInitializeStateInfo(msiPayeStCd);
      }
      txtPayeStCdSpecialInstructions_UsedInAutoCalc.Text = msiPayeStCd.StrlSpclInstrTxt;
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
  private void fnGetStateInfo_Compact() {
    // Comments  : This function retrieves the calculation rule info
    //             for the Payee Residence State at time of death
    // Parameters:  -
    // Returns   :  -
    // Modified  : Berry Kropiwka -11-06-2019 - New for Compact Filling Calcuation
    // --------------------------------------------------
    "fnGetStateInfo_Compact"
.equals(Const cstrCurrentProc As String);

    try {

      if (cboPayeStCd.ListIndex > 0) {

        // MME START WRUS 4999 - ADDED EXTRA PARAMATERS
        //''' '' BZ4999 October 2013 Non US payee - SXS
        if (!("ZZ".equals(cboPayeStCd.Text))) {
          modGeneral.fnGetStateInfo(CSTCOMPACTFILLING, mfrmMyInsuredForm.frmInsured.getInsuredLobCd(), DateValue(dtpPayePmtDt.chrgHourglass.getValue()), mfrmMyInsuredForm.frmInsured.getInsuredClmID(), dblScreenDBPaymentValue, msiCompactCalc);
        }
      } 
      else {
        modGeneral.fnInitializeStateInfo(msiCompactCalc);
      }
      txtPayeStCdSpecialInstructions_UsedInAutoCalc.Text = msiCompactCalc.StrlSpclInstrTxt;
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
  private void fnInitializeCalcInfo(StateInfo msiIn) { // TODO: Use of ByRef founded Private Sub fnInitializeCalcInfo(ByRef msiIn As StateInfo)
    // Comments  : This function will initialize the calculated amount fields in
    //             the specified State Info structure
    // Parameters: None
    // Returns   : N/A
    // Modified  :
    // --------------------------------------------------
    try {
      "fnInitializeCalcInfo"
.equals(Const cstrCurrentProc As String);
      Const(cintZero As Integer == 0);

      msiIn.NbrOfDaysToPayInterest = cintZero;
      msiIn.InterestRateToUse = cintZero;
      msiIn.ClaimInterestAmt = cintZero;
      msiIn.WithheldAmt = cintZero;
      msiIn.TotalForThisPayee = cintZero;
      msiIn.CalculationInfo = "";
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
  private void fnInitializeEditMode() {
    // Description: Enabled/disables command and navigation buttons, as well
    //             as flips on a flag to indicate the record has been edited.
    // Parameters : N/A
    // Returns    : N/A
    // Modified  :
    // --------------------------------------------------
    try {
      "fnInitializeEditMode"
.equals(Const cstrCurrentProc As String);

      if (setIsDirty() == false) {
        setIsDirty(true);
        fnSetCommandButtons(false);
        fnSetNavigationButtons(bUnconditionalDisable:=True);
        // Clear the Calculation Override checkbox if the user is editing the record. They must re-select that
        // checkbox if they want to override it again.
        // If the checkbox is selected but the corresponding wrapper property is the equivalent of "unchecked"
        // then assume the user just made it go from Unchecked to Checked and thus don't change anything.
        // Otherwise, assume the previous calc had been overriden and thus deselect the indicator so the *next*
        // calc, by default, will not be overriden.
        if ((Not mbInAddMode) && (chkPayeDfltOvrdInd.chrgHourglass.getValue() == vbChecked) && (mtWrapper.getPayeDfltOvrdInd())) {
          chkPayeDfltOvrdInd.chrgHourglass.setValue(vbUnchecked);
          lblWarningAboutOverride.Visible = false;
          //'' BZ4999 October 2013 Non US payee - SXS
          modGeneral.fnEnableDisableControl(ctlIn:=ChkPaye1099Ind, bEnable:=True);
          modGeneral.fnEnableDisableControl(ctlIn:=chkPayeDfltOvrdInd, bEnable:=True);
          modGeneral.fnEnableDisableControl(ctlIn:=cboCalcStCd, bEnable:=False);
          modGeneral.fnEnableDisableControl(ctlIn:=ipdPayeClmIntRt, bEnable:=False);
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
  private void fnLoadCbosForStates() {
    //--------------------------------------------------------------------------
    // Procedure:   fnLoadCbosForStates
    // Description: Populates the Calc State and [Payee Residence] State
    //              comboboxes
    // Params:      N/A
    // Returns:     N/A
    // Modified:
    //-----------------------------------------------------------------------------
    "fnLoadCbosForStates"
.equals(Const cstrCurrentProc As String);
    DBRecordSet rstStates = null;

    try {

      rstStates = fnGetListOfStates();

      //*TODO:** can't found type for with block
      //*With cboCalcStCd
      __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = cboCalcStCd;
      w___TYPE_NOT_FOUND.Clear;

      // Add a blank entry as the first entry of the combobox. This will force the user to select
      // an entry (no default selection) since fnValidData will generate an error if the blank
      // entry is still selected when the user clicks Update.
      w___TYPE_NOT_FOUND.AddItem(modGeneral.gCSTRBLANKENTRY);

      modComboBox.fnADORecordSetToComboBox(rstIn:=rstStates, cboIn:=cboCalcStCd, strDisplayColumn:="st_cd", bClear:=False);

      //*TODO:** can't found type for with block
      //*With cboPayeStCd
      __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = cboPayeStCd;
      w___TYPE_NOT_FOUND.Clear;

      // Add a blank entry as the first entry of the combobox. This will force the user to select
      // an entry (no default selection) since fnValidData will generate an error if the blank
      // entry is still selected when the user clicks Update.
      w___TYPE_NOT_FOUND.AddItem(modGeneral.gCSTRBLANKENTRY);

      modComboBox.fnADORecordSetToComboBox(rstIn:=rstStates, cboIn:=cboPayeStCd, strDisplayColumn:="st_cd", bClear:=False);
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    modGeneral.fnFreeRecordset(rstStates);

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
  private void fnLoadCboPayeSsnTinTypCd() {
    // Comments  : Populates CboPayeSsnTinTypCd combo box
    // Parameters: None
    // Modified  :
    // --------------------------------------------------
    try {
      "fnLoadCboPayeSsnTinTypCd"
.equals(Const cstrCurrentProc As String);

      //*TODO:** can't found type for with block
      //*With cboPayeSsnTinTypCd
      __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = cboPayeSsnTinTypCd;
      w___TYPE_NOT_FOUND.Clear;
      //' blank (default)
      w___TYPE_NOT_FOUND.AddItem(modGeneral.gCSTRBLANKENTRY);
      //' P = Person
      w___TYPE_NOT_FOUND.AddItem(MCSTRPAYEEISAPERSON);
      //' B = Business
      w___TYPE_NOT_FOUND.AddItem(MCSTRPAYEEISABUSINESS);
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



  private void fnLoadControls() {
    // Comments  : Populates screen controls with data from recordset
    // Parameters: None
    // Modified  :
    // --------------------------------------------------
    try {
      "fnLoadControls"
.equals(Const cstrCurrentProc As String);
      "B"
.equals(Const cstrPayeeIsABusiness As String);

      // Deliberately load Date of Payment before other State Codes, so the StateInfo structures will be
      // set up correctly, with the fewest number of event-driven iterations
      dtpPayePmtDt.chrgHourglass.setValue(mtWrapper.getPayePmtDt());

      iptPayeFullNm.Text = mtWrapper.getPayeFullNm();
      iptPayeCareOfTxt.Text = modDataConversion.fnZLSIfNull(mtWrapper.getPayeCareOfTxt());
      iptPayeAddrLn1Txt.Text = mtWrapper.getPayeAddrLn1Txt();
      iptPayeAddrLn2Txt.Text = modDataConversion.fnZLSIfNull(mtWrapper.getPayeAddrLn2Txt());
      iptPayeCityNmTxt.Text = mtWrapper.getPayeCityNmTxt();
      cboPayeStCd.Text = mtWrapper.getPayeStCd();
      iptPayeZipCd.Text = mtWrapper.getPayeZipCd();
      iptPayeZip4Cd.Text = modDataConversion.fnZLSIfNull(mtWrapper.getPayeZip4Cd());
      //' BZ4999 October 2013 Non US payee - SXS
      if (mtWrapper.getPaye1099INTInd()) {
        ChkPaye1099Ind = vbChecked;
      } 
      else {
        ChkPaye1099Ind.chrgHourglass.setValue(vbUnchecked);
      }

      // Set the TIN Type before we set the SSN Tin Num. This makes
      // sure the ipmPayeSsnTinNum control's mask gets set appropriately
      // so the next code chunk (setting the value of the
      // ipvPayeSsnTinNum control) will work correctly.
      //
      // Note: Since this is a nullable field, the cbo must have
      // a blank entry in it, so the fnZLSIfNull( ) call will work.
      if (LenB(mtWrapper.getPayeSsnTinTypCd()) == 0) {
        cboPayeSsnTinTypCd.Text = modGeneral.gCSTRBLANKENTRY;
      } 
      else {
        cboPayeSsnTinTypCd.Text = modDataConversion.fnZLSIfNull(mtWrapper.getPayeSsnTinTypCd());
      }

      // NOTE: For MaskEdBox or Input Pro Mask controls, have to do special processing based on
      //       whether or not the field is empty, to avoid a 380 "invalid property value" runtime error.
      //       * If it's empty, temporarily delete the mask, set the value, and then restore
      //         the mask.
      //       * If it's not empty, format the value so it will be "valid" per the .Mask
      //         (for phone numbers, this means inserting a dash between characters 3 and 4).
      if (LenB(mtWrapper.getPayeSsnTinNum()) == 0) {
        ipmPayeSsnTinNum.ctclmClaim.setMask("");
        ipmPayeSsnTinNum.Text = "";
        ipmPayeSsnTinNum.ctclmClaim.setMask(MCSTRUNKNOWNTINTYPEMASK);
      } 
      else {
        if (mtWrapper.getPayeSsnTinTypCd().equals(cstrPayeeIsABusiness)) {
          ipmPayeSsnTinNum.Text = modGeneral.fnSSNTIN_AddDash(strIn:=.ctpyePayee.getPayeSsnTinNum(), bIsTin:=True);
        } 
        else {
          ipmPayeSsnTinNum.Text = modGeneral.fnSSNTIN_AddDash(strIn:=.ctpyePayee.getPayeSsnTinNum(), bIsTin:=False);
        }
      }

      ipdPayeWthldRt.Text = mtWrapper.getPayeWthldRt();

      // ------------------------------------------------------------------------
      //    The checkPayeDfltOvrdInd, cboCalcStCd and iptPayeClmIntRt and
      //   lblWarningAboutOverride controls are all tied to one another with
      //   regard to their availability and initialization.
      // ------------------------------------------------------------------------
      cboCalcStCd.Text = mtWrapper.getCalcStCd();
      ipdPayeClmIntRt.Text = mtWrapper.getPayeClmIntRt();

      if (mtWrapper.getPayeDfltOvrdInd()) {
        chkPayeDfltOvrdInd.Enabled = true;
        chkPayeDfltOvrdInd.chrgHourglass.setValue(vbChecked);
        lblWarningAboutOverride.Visible = true;
        modGeneral.fnEnableDisableControl(ctlIn:=cboCalcStCd, bEnable:=True);
        modGeneral.fnEnableDisableControl(ctlIn:=ipdPayeClmIntRt, bEnable:=True);
      } 
      else {
        chkPayeDfltOvrdInd.Enabled = true;
        chkPayeDfltOvrdInd.chrgHourglass.setValue(vbUnchecked);
        lblWarningAboutOverride.Visible = false;
        modGeneral.fnEnableDisableControl(ctlIn:=cboCalcStCd, bEnable:=False);
        modGeneral.fnEnableDisableControl(ctlIn:=ipdPayeClmIntRt, bEnable:=False);
      }
      // ------------------------------------------------------------------------


      ipdPayeIntDaysPdNum.Text = mtWrapper.getPayeIntDaysPdNum();
      ipcPayeDthbPmtAmt.Text = mtWrapper.getPayeDthbPmtAmt();
      ipcPayeClmIntAmt.Text = mtWrapper.getPayeClmIntAmt();
      ipcPayeWthldAmt.Text = mtWrapper.getPayeWthldAmt();
      ipcPayeClmPdAmt.Text = mtWrapper.getPayeClmPdAmt();

      // ClmId         isn't shown on-screen
      // PayeId        isn't shown on-screen
      // LstUpdDtm     isn't shown on-screen
      // LstUpdUserId  isn't shown on-screen

      // No need to load txtPayeStCd_UsedInAutoCalc and txtPayeStCdSpecialInstructions_UsedInAutoCalc,
      // since they is set when the Payee State (cboPayeStCd) is set.

      // No need to load txtInsdDthResStCd_UsedInAutoCalc and txtInsdDthResStCdSpecialInstructions_UsedInAutoCalc,
      // since they are set during Form_Load and don't need to be changed.

      // No need to set txtCalcStCdSpecialInstructions_UsedInAutoCalc; it is set whenever
      // the txtCalculationState changes.

      // Make sure Navigation buttons are enabled/disabled based on current record position in the Lookup recordset
      fnSetNavigationButtons(bUnconditionalDisable:=False);

      // Update the "record x of y" label
      lblRecordPosition = modGeneral.fnShowRecordPosition(mtWrapper.getLookupData());

      if (modGeneral.bDEBUGAPPTERMINATION) {
        Debug.Print("   Turning off Update mode in "+ mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
      }

      // Set to False to show there are no pending changes. Loading data to controls above
      // could trigger fnInitializeEditMode to falsely think there is a pending change.
      setIsDirty(false);
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
  private void fnLoadLpcLookup() {
    //--------------------------------------------------------------------------
    // Procedure:   fnLoadLpcLookup
    // Description: Populates the specified fpCombo Lookup control using
    //              the mtWrapper's lookup recordset
    // Params:      N/A
    // Returns:     N/A
    // Modified:
    //-----------------------------------------------------------------------------
    "fnLoadLpcLookup"
.equals(Const cstrCurrentProc As String);
    Const(cintRowDimension As Integer == 2);
    Object[] aRows() = null;
    int lngRow = 0;
    try {


      //*TODO:** can't found type for with block
      //*With lpcLookupName
      __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = lpcLookupName;
      w___TYPE_NOT_FOUND.Clear;
      w___TYPE_NOT_FOUND.SortState = SortStateSuspend;

      if (mtWrapper.getLookupRecordCount() != 0) {
        aRows = mtWrapper.getLookupData_Name();
      }

      // Add a blank entry as the first entry of the combobox. This will force the user to select
      // an entry (no default selection) since fnValidData will generate an error if the blank
      // entry is still selected when the user clicks Update.

      // Set .Row to -1 so insertion works okay whether or not the fpCombo is sorted
      w___TYPE_NOT_FOUND.Row = modComboBox.gCLNGNOSELECTION;
      // 3 columns: PAYE_FULL_NM, PAYE_ID and CLM_ID
      w___TYPE_NOT_FOUND.InsertRow = modGeneral.gCSTRBLANKENTRY+ vbTab+ modGeneral.gCSTRBLANKENTRY+ vbTab+ modGeneral.gCSTRBLANKENTRY;
      // Next statement gets a run-time error 9 (subscript out of range) if the aRows array is empty
      for (lngRow = 0; lngRow <= (aRows, cintRowDimension).length; lngRow++) {
        // There are 3 columns in the array and fpCombo control (indexed 0 thru 2).
        w___TYPE_NOT_FOUND.InsertRow = aRows[0, lngRow]+ vbTab+ aRows[1, lngRow]+ vbTab+ aRows[2, lngRow];
      }
      //.SortState = SortStateActiveReSort
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    Erase(aRows);

    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
    }
    return;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      //' subscript out of range
      case  9 :
        /**TODO:** resume found: Resume(Next)*/;
        break;

      default:
        modGeneral.gerhApp.saveErrObjectData(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  private void fnPerformLookup(LPLib.fpCombo lpcIn) { // TODO: Use of ByRef founded Private Sub fnPerformLookup(ByRef lpcIn As LPLib.fpCombo)
    // Comments  : Retrieves selected record
    // Parameters: N/A
    // Called by : lpcLookupClaim_Click
    //             lpcLookupClaim_KeyDown (if Enter was pressed)
    // Modified  :
    //
    // --------------------------------------------------
    "fnPerformLookup"
.equals(Const cstrCurrentProc As String);
    chrgHourglass hrgHourglass = null;
    int lngRecordKeyToRetrieve = 0;

    try {

      lpcIn.Col = 0;

      //.SetFocus
      lpcIn.ColFromName = MCSTRPAYEID;

      // If there are no records in the main table maintained by this form,
      // if the blank entry was selected, or if the user typed in nothing
      // (i.e. a blank entry in the Lookup box), then skip further processing.
      // There's nothing to do a lookup on!
      // If the LookupRecordCount = 0 then we should already be in Add mode
      // and thus should just stay as we are.
      if ((mtWrapper.getLookupRecordCount() == 0) || (lpcIn.ColText == modGeneral.gCSTRBLANKENTRY) || (lpcIn.ColText == "")) {
        // **TODO:** goto found: GoTo PROC_EXIT;
      }

      // Above GoTo avoids a run-time error 13 (type mismatch) on the next
      // statement if .ColText = gcstrBlankEntry
      lngRecordKeyToRetrieve = lpcIn.ColText;

      // Restore focus back to the display column
      lpcIn.ColFromName = MCSTRDISPLAYCOL;

      // Turn on hourglass, in case the lookup is slow
      hrgHourglass = new chrgHourglass();
      hrgHourglass.setValue(true);

      // If the user issues a lookup request while in Add mode or while there are
      // pending changes, then it is interpreted to mean that all pending changes
      // should be discarded. Hence, turn off Add mode and the IsDirty flag and then
      // retrieve the selected record.
      if (modGeneral.bDEBUGAPPTERMINATION) {
        Debug.Print("   Turning off Add mode in "+ mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
      }
      mbInAddMode = false;
      if (modGeneral.bDEBUGAPPTERMINATION) {
        Debug.Print("   Turning off Update mode in "+ mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
      }
      setIsDirty(false);
      fnSetAvailabilityOfControls(bChangeFocus:=False);
      mtWrapper.getSingleRecord(lngKey1:=lngRecordKeyToRetrieve, bSynchLookupRST:=True);
      Me.Refresh;
      // Load current record's properties to form's controls, reset navigation buttons
      // and set "rec x of y" label
      fnLoadControls();
      fnSetCommandButtons(true);
      Me.Refresh;
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    if (!(hrgHourglass == null)) {
      hrgHourglass.setValue(false);
    }
    modGeneral.fnFreeObject(hrgHourglass);

    // Report the error, since this is an event handler
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.reportFatalError(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
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
  private double fnPromptForRate(String strPromptText) {
    double _rtn = 0;
    //--------------------------------------------------------------------------
    // Procedure:   fnPromptForRate
    // Description: Prompts the user to supply the Current Loan rate effective on
    //              the Date of Payment
    // Params:      N/A
    // Returns:     double, representing the Current Loan Rate
    // Modified:
    //-----------------------------------------------------------------------------
    "fnPromptForRate"
.equals(Const cstrCurrentProc As String);
    String strTemp = "";
    double dblRate = 0;

    try {

      // **TODO:** label found: PROMPT_FOR_RATE:;
      strTemp = InputBox(strPromptText);
      if (strTemp.equals("")) {
        // vbNullString is returned if the user clicked Cancel in the InputBox. Call SaveAppSpecificError so the
        // error is reported upstream, effecting the cancellation of the Update process.
        // gcRES_NERR_CALC_WAS_CANCELLED (4010) = The calculation was halted since you clicked Cancel. Your changes have not been saved.
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_CALC_WAS_CANCELLED, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
        // **TODO:** goto found: GoTo PROC_EXIT;
      }

      if (G.isNumeric(strTemp)) {
        dblRate = Double.parseDouble(strTemp);
        if (dblRate < 0) {
          // gcRES_WARN_RATE_IS_NEGATIVE (2007) = The Rate supplied or derived from the supplied Rate is a negative number (@@1). Please try again.
          modGeneral.gerhApp.reportNonFatal(vbObjectError + modResConstants.gCRES_WARN_RATE_IS_NEGATIVE, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc, CStr(dblRate));
          // **TODO:** goto found: GoTo PROMPT_FOR_RATE;
        }
        // Silently round up to 5 decimals, to ensure input can be successfully stored in DB
        dblRate = Round(dblRate, 5);
        //intLengthOfRate = Len(strTemp)
        //varPositionOfDecimal = InStr(1, strTemp, ".", vbTextCompare)
        //If IsNull(varPositionOfDecimal) Then
        //    varPositionOfDecimal = 0
        //End If
        //intMaxDecimalsAllowed = mtWrapper.DecimalPositions(ipdPayeClmIntRt.Tag)
        //If (intLengthOfRate - varPositionOfDecimal) > intMaxDecimalsAllowed Then
        //    ' gcRES_WARN_TOO_MANY_DECIMALS (2008) = The Rate supplied cannot have more than @@1 decimal positions specified. Please try again.
        //    gerhApp.ReportNonFatal vbObjectError + gcRES_WARN_TOO_MANY_DECIMALS, _
        //                           mstrScreenName & gcstrDOT & cstrCurrentProc, _
        //                           intMaxDecimalsAllowed
        //    GoTo PROMPT_FOR_RATE
        //End If
      } 
      else {
        // gcRES_WARN_NONNUMERIC_RATE (2006) = The Rate supplied or obtained from the STATE_RULE_T table is non-numeric (@@1). Please try again.
        modGeneral.gerhApp.reportNonFatal(vbObjectError + modResConstants.gCRES_WARN_NONNUMERIC_RATE, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc, strTemp);
        // **TODO:** goto found: GoTo PROMPT_FOR_RATE;
      }

      _rtn = dblRate;
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
  private void fnRefreshAllCombos() {
    //--------------------------------------------------------------------------
    // Procedure:   fnRefreshAllCombos
    // Description: Repopulates each ComboBox or VSFlexGrid control
    //              so they reflect this and other users' changes. This proc
    //              should be called after each Add, Update or Delete.
    //
    // Params:      N/A
    // Called by:   cmdUpdate_Click() of frmFund
    //              cmdDelete_Click() of frmFund
    //              Form_Load() of frmFund
    //
    // Returns:     N/A
    //-----------------------------------------------------------------------------
    //!CUSTOMIZE!    This should call a function to load each ComboBox or
    //               VSFlexGrid control on the form. This will ensure that
    //               when one is refreshed (i.e. to make this and other
    //               user's changes visible), *all* will be.
    "fnRefreshAllCombos"
.equals(Const cstrCurrentProc As String);
    try {

      //' Payee Full Name (PAYE_FULL_NM, PAYE_ID, CLM_ID)
      fnLoadLpcLookup();
      //' SSN Tin Type (single column: P or B or blank)
      fnLoadCboPayeSsnTinTypCd();
      //' Calc State / Payee State (ST_CD)
      fnLoadCbosForStates();
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
  private void fnResetStateRules() {
    //--------------------------------------------------------------------------
    // Procedure:   fnResetStateRules
    // Description: Resets all rules. This function should be called if the
    //              Payee's Date of Payment has been changed.
    //
    // Params:      N/A
    // Called by:   fnClearControls
    //              dtpPayePmgDt_Change() of frmPayee
    //
    // Returns:     N/A
    // Modifed:     Berry Kropiwka - 11-06-2019 - Add fngetstateinfo_compact for compact filling
    //-----------------------------------------------------------------------------
    "fnResetStateRules"
.equals(Const cstrCurrentProc As String);
    try {

      // Populate StateInfo structures with data from the STATE_RULE_T
      // row that matches the various State Codes.
      fnGetStateInfo_InsdDthResStCd();
      fnGetStateInfo_IssStCd();
      fnGetStateInfo_PayeStCd();
      fnGetStateInfo_Override();
      if (mfrmMyInsuredForm.chkClmCmpCalInd.chrgHourglass.getValue() == vbChecked) {
        // This is Compact Calcatution
        fnGetStateInfo_Compact();
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
  private void fnSetAvailabilityOfControls(boolean bChangeFocus) {
    //--------------------------------------------------------------------------
    // Procedure:   fnSetAvailabilityOfControls
    // Description: Determines whether a control representing a lookup
    //              or a key field should be display-only.
    //
    // Params:      bChangeFocus - If True, moves the focus to the first updateable field.
    //
    // Called by:   cmdUpdate_Click
    //              fnAddRecord
    //              Form_Load
    //              Form_QueryUnload
    // Returns:     N/A
    //-----------------------------------------------------------------------------
    "fnSetAvailabilityOfControls"
.equals(Const cstrCurrentProc As String);
    Control ctl = null;

    try {

      for (int _i = 0; _i < Me.Controls.size(); _i++) {
        ctl = Me.Controls.item(_i);
        // Debug.Print ctl.Name & vbTab & ctl.Tag

        // If the control corresponds to a SQL Server table column that's a key field, then
        // only enable it if in Add mode.
        if (ctl.Tag.length() > 0) {
          // If it's a key, disable it unless we're in Add mode
          if (mtWrapper.getIsKey(ctl.Tag)) {
            //Debug.Print .Tag & " is a key field, per meta data"
            if (mbInAddMode) {
              modGeneral.fnEnableDisableControl(ctlIn:=ctl, bEnable:=True);
            } 
            else {
              modGeneral.fnEnableDisableControl(ctlIn:=ctl, bEnable:=False);
            }
          }
        }
      }

      if (bChangeFocus) {
        fnSetFocusToFirstUpdateableField();
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
  private void fnSetCommandButtons(boolean bEnable) {
    //----------------------------------------------------------------------------
    // Procedure : fnSetCommandButtons
    //
    // Comments  : Enables/Disables the command buttons, per boolean parameter
    //             Here's how the button enabling should work. Note it assumes
    //             that IsDirty and mbInAddMode have been set prior to
    //             calling this routine, e.g., they accurately reflect whether
    //             or not there are edits outstanding and/or the user is in
    //             Add mode, respectively.
    //             Remember, though: mbInAddMode and IsDirty are
    //             independent of one another!
    //
    //     State          ADD btn  UPD btn  DEL btn  CLOSE btn
    //    --------------  -------- -------- -------- ---------
    //    Add mode       disabled  enabled  disabled enabled
    //    (no edits yet)
    //
    //    Edits o/s      disabled  enabled  disabled enabled
    //
    //    No edits o/s   enabled   disabled enabled  enabled

    //
    // Called by : fnAddRecord and fnInitializeEditMode, with bEnable = False
    //
    //             lpcLookupName_Click, cmdDelete_Click, cmdNavigate_Click, cmdUpdate_Click
    //             (when updating existing record) and Form_Load, with
    //             bEnable = True
    //
    // Parameters: bEnable - indicates whether Add/Update buttons should be enabled
    //                       or disabled
    //
    // Modified  :
    //----------------------------------------------------------------------------
    "fnSetCommandButtons"
.equals(Const cstrCurrentProc As String);
    String strDependent_Table = "";
    boolean bHaveDependents = false;

    try {

      // Hide updates to the window until we're done. This avoids ugly screen flickering
      fnWindowLock(Me.cbrfBrowseFolder.setHWnd());

      cmdAdd.Enabled = bEnable;
      cmdUpdate.Enabled = !bEnable;

      if (mbInAddMode) {
        bHaveDependents = false;
      } 
      else {
        bHaveDependents = mtWrapper.haveDependents(mtWrapper.getClmId(), strDependent_Table);
      }

      // Can only delete a record when (a) when you're not in the middle of an Add or Update
      // and (b) there are no rows in dependent tables (i.e. children).
      if ((setIsDirty() || mbInAddMode)) {
        cmdDelete.Enabled = false;
      } 
      else {
        if ((bHaveDependents)) {
          cmdDelete.Enabled = false;
        } 
        else {
          cmdDelete.Enabled = true;
        }
      }

      // Can only use to the Clone This Payee button when you're NOT in the middle of
      // an Add or Update.
      if ((Not setIsDirty()) && (Not mbInAddMode)) {
        cmdCloneThisPayee.Enabled = true;
      } 
      else {
        cmdCloneThisPayee.Enabled = false;
      }
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    fnWindowUnlock;

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
  private void fnSetDefaultControlProperties() {
    //--------------------------------------------------------------------------
    // Procedure:   fnSetDefaultControlProperties
    // Description: Sets default properties of controls bound to table columns
    //              in the table wrapper class, using the meta data that class
    //              gathered.
    //
    //              These defaults are initially based on the data type
    //              of the column (see the table wrapper's fnGetColMetaData method)
    //              but then overriden, if desired, in the table wrapper's
    //              fnLoadColMetaData method.
    //
    //              NOTE: Tags should only be present if the control
    //                    is bound to a property of the table wrapper class.
    //                    Also, the entire contents of the Tag should be the
    //                    name of the public property in that class.
    //
    // Params:      N/A
    // Called by:   Form_Load
    // Returns:     N/A
    //-----------------------------------------------------------------------------
    "fnSetDefaultControlProperties"
.equals(Const cstrCurrentProc As String);
    Control ctl = null;
    boolean bSavedIsDirty = false;

    try {

      // This procedure can be called, among other places, by Form_Load before
      // the screen controls have been loaded with values. Setting some of
      // those controls' properties can trigger their Change event which
      // causes fnSetCommandButtons and ultimately the table wrapper's
      // HaveDependents procs to be called. The latter can fail with a
      // spurious error if the key hasn't been set. Since we really don't
      // care about this processing because it'll be hit again after
      // data *has* been loaded to the controls, let's just fake it out
      // here by making any Change event hit by this proc's code
      // think that "IsDirty" processing has already been done. We'll restore
      // the IsDirty flag when we're done.

      // Start of "fake out"
      if (modGeneral.bDEBUGAPPTERMINATION) {
        Debug.Print("   Saving then faking out Update mode in "+ mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
      }

      bSavedIsDirty = setIsDirty();
      setIsDirty(true);

      for (int _i = 0; _i < Me.Controls.size(); _i++) {
        ctl = Me.Controls.item(_i);
        // Debug.Print ctl.Name
        // If ctl.Name = "fraPayeeInfo" Then
        //    Debug.Print "yes"
        // End If
        // If control corresponds to a SQL Server table column, then try
        // to set its default properties. The Tag property contains
        // the name of its property within the table class.
        if (ctl.Tag.length() > 0) {
          // If it's a key, disable it unless we're in Add mode
          if (mtWrapper.getIsKey(ctl.Tag)) {
            ctl.Enabled = mbInAddMode;
          }

          if (ctl(instanceOf TextBox)) {
            ctl.MaxLength = mtWrapper.getMaxCharacters(ctl.Tag);
          }

          if (ctl(instanceOf fpText)) {
            ctl.MaxLength = mtWrapper.getMaxCharacters(ctl.Tag);
            // Convert VB True (-1) to 1 (AutoCaseUpper) and vb False (0) to 0 (AutoCaseNone)
            ctl.AutoCase = Abs(mtWrapper.getShouldForceToUppercase(ctl.Tag));
            ctl.CharValidationText = mtWrapper.getAllowableCharacters(ctl.Tag);
          }

          if (ctl(instanceOf fpDateTime)) {
            //' Show days from previous/next month if possible
            ctl.CalGrayAreaStyle(2);
            //' Allow user to scroll by clicking in gray area of calendar
            ctl.CalGrayAreaAllowScroll(true);
            //' Show invalid date w/ diff bkgrd color; don't auto-correct it
            ctl.InvalidOption = ShowData;
            ctl.PopUpType = PopCalendar;
            ctl.UserEntry = UserEntryFormatted;
            ctl.ButtonStyle = ButtonStyleDropDown;
            ctl.DateTimeFormat = IntlShortDate;
          }

          //If TypeOf ctl Is MaskEdBox Then
          //    .MaxLength = mtWrapper.MaxCharacters(.Tag)
          //    .Mask = mtWrapper.Mask(.Tag)
          //End If

          if (ctl(instanceOf fpMask)) {
            ctl.MaxLength = mtWrapper.getMaxCharacters(ctl.Tag);
            ctl.Mask = mtWrapper.getMask(ctl.Tag);
          }
        }
        // Make all fpCurrency controls have the same formatting  to start with.
        if (ctl(instanceOf fpCurrency)) {
          ctl.AlignTextH = AlignTextHRight;
          ctl.AllowNull = true;
          ctl.NullColor = vbRed;
          ctl.BackColor = vbWindowBackground;
          ctl.ForeColor = vbWindowText;
          ctl.InvalidColor = vbWindowText;
          //' ($1)
          ctl.CurrencyNegFormat = pd1p;
          ctl.LeadZero = NoLeadingZero;
          ctl.UseSeparator = true;
          ctl.OnFocusNoSelect = false;
          ctl.OnFocusAlignH = OnFocusAlignHRight;
          //' Cannot use "-" on numeric keypad to toggle btwn pos/neg
          ctl.NegToggle = false;
          //' Negatives not allowed
          ctl.MinValue = 0;
          ctl.NoSpecialKeys = AllKeysEnabled;
          // Next line is needed to avoid -2147217887 (Invalid character value for cast specification) error
          // when storing too big a value into a SQL column.
          ctl.MaxValue = modGeneral.fnTranslateToMaxValue(mtWrapper.getDollarPositions(ctl.Tag), mtWrapper.getDecimalPositions(ctl.Tag));
        }
        // Make all fpDoubleSingle controls have the same formatting  to start with.
        if (ctl(instanceOf fpDoubleSingle)) {
          ctl.AlignTextH = AlignTextHRight;
          ctl.AllowNull = true;
          ctl.NullColor = vbRed;
          ctl.BackColor = vbWindowBackground;
          ctl.ForeColor = vbWindowText;
          ctl.InvalidColor = vbWindowText;
          ctl.LeadZero = NoLeadingZero;
          ctl.UseSeparator = true;
          ctl.OnFocusNoSelect = false;
          ctl.OnFocusAlignH = OnFocusAlignHRight;
          ctl.DecimalPlaces = mtWrapper.getDecimalPositions(ctl.Tag);
          ctl.FixedPoint = true;
          //' (1)
          ctl.NegFormat = n1;
          //' Cannot use "-" on numeric keypad to toggle btwn pos/neg
          ctl.NegToggle = false;
          //' Negatives not allowed
          ctl.MinValue = 0;
          ctl.NoSpecialKeys = AllKeysEnabled;
          // Next line is needed to avoid -2147217887 (Invalid character value for cast specification) error
          // when storing too big a value into a SQL column.
          ctl.MaxValue = modGeneral.fnTranslateToMaxValue(mtWrapper.getDollarPositions(ctl.Tag), mtWrapper.getDecimalPositions(ctl.Tag));
        }
        // Make all fpMask controls have the same formatting to start with.
        if (ctl(instanceOf fpMask)) {
          ctl.AlignTextH = AlignTextHLeft;
          ctl.AllowNull = false;
          ctl.NullColor = vbRed;
          ctl.BackColor = vbWindowBackground;
          ctl.ForeColor = vbWindowText;
          ctl.InvalidColor = vbWindowText;
          ctl.HideSelection = true;
          ctl.OnFocusNoSelect = false;
          ctl.OnFocusAlignH = OnFocusAlignHLeft;
          //' Char to display in unfilled positions?
          ctl.PromptChar = "_";
          //' Include Prompt char when saving bound value to DB?
          ctl.PromptInclude = false;
          //' Trigger event if all prompt chars not supplied when ctl loses focus?
          ctl.RequireFill = 0;
        }
        // Make all fpText controls have the same formatting to start with.
        // NOTE: Be sure not to un-do any settings done earlier in this procedure, for fpText
        //       for controls bound to a table column!!
        if (ctl(instanceOf fpText)) {
          //' How to align horizontally?
          ctl.AlignTextH = AlignTextHLeft;
          //' Is Null a valid value? User can press Ctrl-N or F2 to insert a Null value.
          ctl.AllowNull = false;
          //' Color of contents when value is Null
          ctl.NullColor = vbRed;
          ctl.BackColor = vbWindowBackground;
          ctl.ForeColor = vbWindowText;
          ctl.InvalidColor = vbWindowText;
          //' contents selected when ctl loses focus?
          ctl.HideSelection = true;
          //' don't select contents when ctl receives focus?
          ctl.OnFocusNoSelect = false;
          ctl.MultiLine = false;
        }
      }

      // Per Michelle, the Date of Payment can be future-dated up to 5 days.
      //'''''''''''''''dtpPayePmtDt.MaxDate = DateAdd("d", 5, Now)  '''''' BZ 6495 SXS
      dtpPayePmtDt.MaxDate = DateAdd("d", 30, Now);

      //' 99.00000
      ipdPayeWthldRt.MaxValue = 99#;

      // End of "fake out"
      if (modGeneral.bDEBUGAPPTERMINATION) {
        Debug.Print("   Restoring saved Update mode afer fake out in "+ mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
      }

      setIsDirty(bSavedIsDirty);
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
  private void fnSetFocusToFirstUpdateableField() {
    //--------------------------------------------------------------------------
    // Procedure:   fnSetFocusToFirstUpdateableField
    // Description: Moves the focus to the first editable (i.e. updateable) field on the screen
    //
    // Params:      N/A
    // Called by:
    //
    // Returns:     N/A
    //-----------------------------------------------------------------------------
    try {
      "fnSetFocusToFirstUpdateableField"
.equals(Const cstrCurrentProc As String);

      // Set focus to first editable field, by default
      if (mbInAddMode) {
        if (mctlFirstUpdateableField_Add.Visible) {
          mctlFirstUpdateableField_Add.SetFocus;
        }
      } 
      else {
        if (mctlFirstUpdateableField_Upd.Visible) {
          mctlFirstUpdateableField_Upd.SetFocus;
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
      case  5:
        // Invalid Procedure Call or Argument  (See MSKB Article Q242347)
        /**TODO:** resume found: Resume(Next)*/;
        break;

      default:
        modGeneral.gerhApp.saveErrObjectData(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }
}


//////////////////////////////////////////////////////////////////////////////////////////////////
  private void fnSetNavigationButtons(boolean bUnconditionalDisable) {
    //----------------------------------------------------------------------------
    // Procedure  : fnSetNavigationButtons
    // Description: Enables/Disables the control array of navigation buttons, based
    //              on the bEnable input parameter
    //
    // Parameters:  bUnconditionalDisable (in) - indicates whether buttons should be disabled
    //                  regardless of where the current record position is in the recordset.
    //                  This will generally be set to True only via the
    //                  fnAddRecords( ) and fnInitializeEditMode( ) procs.
    //
    // Called by :
    //              cmdDelete_Click( )
    //              cndNavigate_Click( )
    //              fnAddRecord( )
    //              fnInitializeEditMode( )
    //              Form_Load( )
    //              lpcLookupName_Click( )
    //
    // Returns   :  N/A
    // Modified  :
    //----------------------------------------------------------------------------
    "fnSetNavigationButtons"
.equals(Const cstrCurrentProc As String);
    CommandButton cmd = null;
    boolean bHaveRecords = false;

    try {

      if (bUnconditionalDisable) {
        for (int _i = 0; _i < cmdNavigate.size(); _i++) {
          cmd = cmdNavigate.item(_i);
          cmd.Enabled = false;
        }
        // **TODO:** goto found: GoTo PROC_EXIT;
      }

      //...........................................................
      // Enable navigation buttons based on where we're currently
      // positioned in the Lookup recordset
      //...........................................................

      // Default to all buttons enabled if there are records in the Lookup recordset; Otherwise, disable them all.
      bHaveRecords = (mtWrapper.getLookupRecordCount() != 0);
      for (int _i = 0; _i < cmdNavigate.size(); _i++) {
        cmd = cmdNavigate.item(_i);
        cmd.Enabled = bHaveRecords;
      }

      // Now selectively disable if our current record position causes certain navigation to be unavailable/illogical.
      if (bHaveRecords) {
        if (mtWrapper.getCurrentLookupRecordNumber() == 1) {
          cmdNavigate(navFirst).Enabled = false;
          cmdNavigate(navPrev).Enabled = false;
        }

        if (mtWrapper.getCurrentLookupRecordNumber() == mtWrapper.getLookupRecordCount()) {
          cmdNavigate(navNext).Enabled = false;
          cmdNavigate(navLast).Enabled = false;
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
  private void fnSetupScreenControls() {
    //--------------------------------------------------------------------------
    // Procedure:   fnSetupScreenControls
    // Description: This procedure:
    //              * Binds the on-screen controls to the table wrapper class
    //                properties with which they are associated.
    //              * Sets default settings for those controls' properties
    //              * Binds editable TextBoxes controls to the Extended TextBox
    //                class so they will behave appropriately and in a consistent
    //                manner.
    //
    // Params:      N/A
    // Returns:     N/A
    // Date:        04/04/2002
    //-----------------------------------------------------------------------------
    "fnSetupScreenControls"
.equals(Const cstrCurrentProc As String);
    try {

      // Set each control's Tag property to identify the table class property to which it corresponds,
      // set its defaults attributes per the DBMS' meta data. In addition, for those columns that correspond
      // to editable TextBox controls, set properties of its associated ExtendedTextBox variable, so
      // it will behave appropriately and in a standard manner.
      fnBindControlsToTableWrapper();

      // Set default attributes for those controls, per the DBMS' meta data
      fnSetDefaultControlProperties();

      // Disable controls that are always "display-only"
      modGeneral.fnEnableDisableControl(ctlIn:=txtPayeStCd_UsedInAutoCalc, bEnable:=False);
      modGeneral.fnEnableDisableControl(ctlIn:=txtPayeStCdSpecialInstructions_UsedInAutoCalc, bEnable:=False);
      modGeneral.fnEnableDisableControl(ctlIn:=chkClmForResDthInd_UsedInAutoCalc, bEnable:=False);
      modGeneral.fnEnableDisableControl(ctlIn:=txtInsdDthResStCd_UsedInAutoCalc, bEnable:=False);
      modGeneral.fnEnableDisableControl(ctlIn:=txtInsdDthResStCdSpecialInstructions_UsedInAutoCalc, bEnable:=False);
      modGeneral.fnEnableDisableControl(ctlIn:=txtIssStCd_UsedInAutoCalc, bEnable:=False);
      modGeneral.fnEnableDisableControl(ctlIn:=txtIssStCdSpecialInstructions_UsedInAutoCalc, bEnable:=False);
      modGeneral.fnEnableDisableControl(ctlIn:=txtCalcStCdSpecialInstructions_UsedInAutoCalc, bEnable:=False);
      modGeneral.fnEnableDisableControl(ctlIn:=ipdPayeIntDaysPdNum, bEnable:=False);
      modGeneral.fnEnableDisableControl(ctlIn:=ipcPayeClmIntAmt, bEnable:=False);
      modGeneral.fnEnableDisableControl(ctlIn:=ipcPayeWthldAmt, bEnable:=False);
      modGeneral.fnEnableDisableControl(ctlIn:=ipcPayeClmPdAmt, bEnable:=False);
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
  private boolean fnValidData() {
    boolean _rtn = false;
    // Comments  : Determines if all data is valid, including
    //             whether all required fields have been input.
    //             This function is called by cmdUpdate_Click.
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
      Const(cintMinZipLength As Integer == 5);
      Const(cintMaxZipLength As Integer == 9);
      Const(cintTinLengthIfInput As Integer == 9);
      boolean bErrorFound = false;
      Control ctl = null;
      Control ctlFirstToFail = null;
      int intFailures = 0;
      String strFieldList = "";
      String strMsgText = "";
      int intLengthToTest = 0;


      _rtn = true;

      // Check the fields in a left-to-right, top-to-bottom screen sequence.
      //     1. iptPayeFullNm        9. ipmPayeSsnTinNum
      //     2. iptPayeCareOfTxt    10. cboPayeSsnTinTypCd
      //     3. iptPayeAddrLn1Txt   11. ipdPayeWthldRt
      //     4. iptPayeAddrLn2Txt   12. chkPayeDftOvrdInd
      //     5. iptPayeCityNmTxt    13. cboCalcStCd
      //     6. cboPayeStCd         14. ipdPayeClmIntRt
      //     7. iptPayeZipCd        15. dtpPayePmtDt
      //     8. iptPayeZip4Cd       16. ipcPayeDthbPmtAmt
      //                            17. chkpaye1099ind
      // ------------- 1.  Verify required fields are missing --------------
      // Check key fields too, although they should be absent only if
      // the user is in Add mode and neglected to specify their values.

      // Using Metadata, verify that all fields are populated if required to (i.e. IsNullable() is True)
      for (int _i = 0; _i < Me.Controls.size(); _i++) {
        ctl = Me.Controls.item(_i);
        // Debug.Print ctl.Name

        // If the control corresponds to a SQL Server table column, then determine if its
        // Not Nullable (i.e. required). The Tag property contains the name of its
        // property within the table class. ' If it's Not Nullable and not input...then
        // generate an error
        // Skip over the control that is bound to PAYE_ID since this is a hidden
        // field and thus the user shouldn't be informed if it hasn't been set yet.
        // Skip over the control that is bound to CALC_ST_CD since this, unless overriden,
        // is a disabled field that is set automatically during the Update processing.
        if (ctl.Tag.length() > 0 && (!("PayeId".equals(ctl.Tag))) && (!("CalcStCd".equals(ctl.Tag)))) {
          if (!(mtWrapper.getIsNullable(ctl.Tag))) {
            if (ctl(instanceOf fpCombo)) {
              // Special handling for fpCombo since its default property
              // isn't the one that must be checked
              if ((ctl.ColText.length() == 0) || (ctl.ColText == modGeneral.gCSTRBLANKENTRY)) {
                if (intFailures == 0) {
                  strFieldList = "\\r\\n"+ fnGetFieldLabel(ctl.Name);
                  ctlFirstToFail = ctl;
                } 
                else {
                  strFieldList = strFieldList+ "\\r\\n"+ fnGetFieldLabel(ctl.Name);
                }
                intFailures = intFailures + 1;
              }
            } 
            else {
              //' BZ4999 October 2013 Non US payee - SXS
              if ("Zip".equals(fnGetFieldLabel(ctl.Name)) && ChkPaye1099Ind == 0) {
              } 
              else {
                if ((ctl.length() == 0) || (ctl == modGeneral.gCSTRBLANKENTRY) || ((ctl == null) && "Zip".equals(fnGetFieldLabel(ctl.Name)))) {
                  if (intFailures == 0) {
                    strFieldList = "\\r\\n"+ fnGetFieldLabel(ctl.Name);

                    ctlFirstToFail = ctl;
                  } 
                  else {
                    strFieldList = strFieldList+ "\\r\\n"+ fnGetFieldLabel(ctl.Name);
                  }
                  intFailures = intFailures + 1;
                }
              }
            }
          }
        }
      }

      if (intFailures != 0) {
        bErrorFound = true;
        _rtn = false;
        if (ctlFirstToFail.Visible) {
          ctlFirstToFail.SetFocus;
        }
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_REQD_FIELDS_MISSING, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc, strFieldList);
        // **TODO:** goto found: GoTo PROC_EXIT;
      }



      // ------------- 2.  Verify other characteristics are valid --------------

      // Reset for this section of error validations
      strMsgText = "";
      intFailures = 0;

      // The Zip Code must be input, either as a 5-digit or 9-digit (zip+4) number
      intLengthToTest = ipmPayeSsnTinNum.UnFmtText.length();
      if ((intLengthToTest != 0) && (intLengthToTest == cintMinZipLength) || (intLengthToTest == cintMaxZipLength)) {
        // Do Nothing
      } 
      else {
        intFailures = intFailures + 1;
        ctlFirstToFail = iptPayeZipCd;
        strMsgText = strMsgText+ "\\r\\n"+ "The "+ MCSTRIPTPAYEZIPCDLABEL+ " must be either "+ cintMinZipLength+ " or "+ cintMaxZipLength+ " digits.";
      }
      // End If

      // If input, the SSN/Tin must be 9 digits
      intLengthToTest = ipmPayeSsnTinNum.UnFmtText.length();
      if ((intLengthToTest != 0) && (intLengthToTest != cintTinLengthIfInput)) {
        intFailures = intFailures + 1;
        ctlFirstToFail = ipmPayeSsnTinNum;
        strMsgText = strMsgText+ "\\r\\n"+ "The "+ MCSTRIPMPAYESSNTINNUMLABEL+ " must be "+ cintTinLengthIfInput+ " digits.";
      }

      // If either the SSN/TIN or SSN/TIN Type is input, then both must be input
      if (((ipmPayeSsnTinNum.UnFmtText.length() != 0) && (!(modGeneral.gCSTRBLANKENTRY.equals(cboPayeSsnTinTypCd.Text)))) || ((ipmPayeSsnTinNum.UnFmtText.length() == 0) && (modGeneral.gCSTRBLANKENTRY.equals(cboPayeSsnTinTypCd.Text)))) {
        // Okay
      } 
      else {
        intFailures = intFailures + 1;
        ctlFirstToFail = ipmPayeSsnTinNum;
        strMsgText = strMsgText+ "\\r\\n"+ "If either the "+ MCSTRIPMPAYESSNTINNUMLABEL+ " or "+ MCSTRCBOPAYESSNTINTYPCDLABEL+ " is input, then both must be input.";
      }

      // The Date of Payment must be on or after the Insured's Date of Death
      if (DateValue(dtpPayePmtDt.chrgHourglass.getValue()) < DateValue(mfrmMyInsuredForm.frmInsured.getInsuredClmInsdDthDt())) {
        intFailures = intFailures + 1;
        ctlFirstToFail = dtpPayePmtDt;
        strMsgText = strMsgText+ "\\r\\n"+ "The "+ MCSTRDTPPAYEPMTDTLABEL+ " ("+ ((Boolean) dtpPayePmtDt.chrgHourglass.getValue()).toString()+ ") must be on or after the Insured's "+ MCSTRDTPCLMINSDDTHDTLABEL+ " ("+ mfrmMyInsuredForm.frmInsured.getInsuredClmInsdDthDt()+ ").";
      }

      // The Date of Payment must be on or after the Date of Proof
      if (DateValue(dtpPayePmtDt.chrgHourglass.getValue()) < DateValue(mfrmMyInsuredForm.frmInsured.getInsuredClmProofDt())) {
        intFailures = intFailures + 1;
        ctlFirstToFail = dtpPayePmtDt;
        strMsgText = strMsgText+ "\\r\\n"+ "The "+ MCSTRDTPPAYEPMTDTLABEL+ " ("+ ((Boolean) dtpPayePmtDt.chrgHourglass.getValue()).toString()+ ") must be on or after the Insured's "+ MCSTRDTPCLMPROOFDTLABEL+ " ("+ mfrmMyInsuredForm.frmInsured.getInsuredClmProofDt()+ ").";
      }


      // If the user is overriding the Calc State or Interest Rate, then one but only one
      // of those fields can be input.
      if (chkPayeDfltOvrdInd.chrgHourglass.getValue() == vbChecked) {
        if ((modGeneral.gCSTRBLANKENTRY.equals(cboCalcStCd.Text)) && (ipdPayeClmIntRt.UnFmtText.length() == 0)) {
          intFailures = intFailures + 1;
          ctlFirstToFail = cboCalcStCd;
          strMsgText = strMsgText+ "\\r\\n"+ "If the "+ MCSTRCHKPAYEDFLTOVRDINDLABEL+ " is selected, then both the "+ MCSTRCBOCALCSTCDLABEL+ " and the "+ MCSTRIPDPAYECLMINTRTLABEL+ " must be input.";
        }
      }


      // DB Payment must be a positive non-zero amount.
      if (Double.parseDouble(ipcPayeDthbPmtAmt.UnFmtText) <= 0) {
        intFailures = intFailures + 1;
        ctlFirstToFail = ipcPayeDthbPmtAmt;
        strMsgText = strMsgText+ "\\r\\n"+ "The "+ MCSTRIPCPAYEDTHBPMTAMTLABEL+ " must be supplied as a positive non-zero amount.";
      }

      //!TODO! This can go away if the control enforces this upon data entry (and for pasting)
      if (!(G.isNumeric(ipdPayeClmIntRt.Text))) {
        intFailures = intFailures + 1;
        ctlFirstToFail = ipdPayeClmIntRt;
        strMsgText = strMsgText+ "\\r\\n"+ "The "+ MCSTRIPDPAYECLMINTRTLABEL+ " must be numeric.";
      }


      // Validation for Payee Interest Rate (> 12 and CalcState <> "ME") relocated to
      // fnGetInterestRate since this is a protected field and thus needs to be validated
      // when the user supplies it...from fnGetInterestRate.

      if (intFailures != 0) {
        bErrorFound = true;
        _rtn = false;
        if (ctlFirstToFail.Visible) {
          ctlFirstToFail.SetFocus;
        }
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_CROSS_FLD_VALIDATIONS_FAILED, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc, "your request can be processed", strMsgText);
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
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  5:
        // Invalid Procedure Call or Argument  (See MSKB Article Q242347)
        /**TODO:** resume found: Resume(Next)*/;
        break;

      default:
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
      Const(cdbl99Million As Double == 99999999#);

      if (ipcPayeDthbPmtAmt.chrgHourglass.getValue() > cdbl99Million) {
        //gcRES_WARN_DTHB_PMT_AMT_MAY_BE_TOO_HIGH (2005) = The @@1 exceeds @@2. Please verify this amount is correct.
        modGeneral.gerhApp.reportNonFatal(vbObjectError + modResConstants.gCRES_WARN_DTHB_PMT_AMT_MAY_BE_TOO_HIGH, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc, MCSTRIPCPAYEDTHBPMTAMTLABEL, "$99m");
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
  private void form_Activate() {
    // Comments  : I think this is an unnecessary event handler! (Betsy 05/14/2001)
    // Parameters: None
    // Modified  :
    // --------------------------------------------------
    try {
      "Form_Activate"
.equals(Const cstrCurrentProc As String);

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      if (Me.Visible) {
        fnSetFocusToFirstUpdateableField();
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
  private void form_Initialize() {
    // Comments  : Intializes the form.
    // Parameters: None
    // Modified  :
    // 01/2002 BAW - Populate the new Insured Residence State text box.
    // --------------------------------------------------
    try {
      "Form_Initialize"
.equals(Const cstrCurrentProc As String);

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      // Once the Payee screen becomes the active form, we lose the ability to reference
      // the Insured form using "frmInsured" (Remember, in an MDI environment, there
      // could be multiple instances of an Insured form loaded). So, while it still is
      // the active form, set a reference to it.
      mfrmMyInsuredForm = frmMDIMain.ActiveForm;

      // Set "Claim#" caption on the form
      lblClmNum = mfrmMyInsuredForm.frmInsured.getInsuredClmNum();

      //Y027 07-11-2012
      m_admPolicySystem = mfrmMyInsuredForm.frmInsured.getAdminSystemCode();

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
      //' Invalid procedure call or argument
      case  5     :
        // Caused by setting focus to a field that's not yet visible
        /**TODO:** resume found: Resume(Next)*/;
        break;

      default:
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
    // Comments  :
    // Parameters: None
    // Modified  :
    //   01/2002 BAW - Populate the new Insured Residence State Special Instructions
    //                 field, based on the state set on the Insured screen.
    // --------------------------------------------------
    try {
      "Form_Load"
.equals(Const cstrCurrentProc As String);

      // Set the screen name that will be used to form the Title on message boxes
      mstrScreenName = Me.Caption;

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      // Identify the icons that will be used for the form and the picture next to the Lookup ComboBox
      Me.Icon = LoadResPicture(modResConstants.gCRES_ICON_MAINAPP, vbResIcon);

      // If the user has ever opened this form before, restore its size & placement.
      // If the restore would result in the form being off-screen, just center it instead.
      if (modGeneral.gapsApp.restoreForm(this) == false) {
        //*TODO:** can't found type for with block
        //*With this
        __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = this;
        w___TYPE_NOT_FOUND.Width = MCLNGMINFORMWIDTH;
        w___TYPE_NOT_FOUND.Height = MCLNGMINFORMHEIGHT;
        modGeneral.fnCenterFormOnMDI(frmMDIMain, this);
      }

      //...............................................................................
      // Set our fpCombo control settings, for those used as Lookups. Since these
      // contain lots of rows (5000 or so, currently), they are loaded with sorted data
      // rather than having the control itself sort its contents. This GREATLY improves
      // the time it takes to display the form and refresh the control!
      //...............................................................................
      // 1. Name Lookup
      //*TODO:** can't found type for with block
      //*With lpcLookupName
      __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = lpcLookupName;
      modComboBox.fnInitializefpCombo(lpcIn:=lpcLookupName, bShowColHeaders:=False, bSortable:=False, lngNbrOfCols:=2, lngEditCol:=mcintDisplayCol_lpcLookupName, lngNbrOfRowsInDropdown:=8);
      // Column definitions
      //' First column, Primary sort
      w___TYPE_NOT_FOUND.Col = 0;
      w___TYPE_NOT_FOUND.ColHeaderText = MCSTRIPTPAYEFULLNMLABEL;
      w___TYPE_NOT_FOUND.ColName = MCSTRDISPLAYCOL;
      //' Second column
      w___TYPE_NOT_FOUND.Col = 1;
      w___TYPE_NOT_FOUND.ColHeaderText = MCSTRTXTPAYEIDLABEL;
      w___TYPE_NOT_FOUND.ColName = MCSTRPAYEID;
      w___TYPE_NOT_FOUND.ColHide = true;
      w___TYPE_NOT_FOUND.ColumnSearch = MCINTDISPLAYCOL_LPCLOOKUPNAME;


      // Set the control to receive the focus after errors (the first editable field
      // on the screen), dependent upon whether we're in Add Mode or not. If in Add mode,
      // this control would typically be the first control that corresponds to a Key field.
      // If not in Add mode, this control would typically be the topmost/leftmost
      // "always updateable" control on the screen (excepting the Lookup ComboBox).
      mctlFirstUpdateableField_Add = iptPayeFullNm;
      mctlFirstUpdateableField_Upd = iptPayeFullNm;

      // Instantiate and initialize a table wrapper object for the appropriate table(s).
      mtWrapper = new ctpyePayee();
      mtWrapper.initPayee(mfrmMyInsuredForm.frmInsured.getInsuredClmID());

      // Populate the Insured's state of residence (at time of death) and its corresponding
      // Special Instructions.
      txtInsdDthResStCd_UsedInAutoCalc.Text = mfrmMyInsuredForm.frmInsured.getInsuredInsdDthResStCd();
      txtInsdDthResStCdSpecialInstructions_UsedInAutoCalc.Text = msiInsdDthResStCd.StrlSpclInstrTxt;


      // Bind the on-screen controls to the table wrapper class properties with which they
      // are associated. set default settings for those controls' properties, and
      // bind editable TextBoxes controls to the Extended TextBox class so they will
      // behave appropriately and in a consistent manner.
      fnSetupScreenControls();

      // Populate all ComboBoxes and ListPro controls
      fnRefreshAllCombos();

      // "And" condition added 05/22/01 for bug 0033, so this screen can be invoked
      // through the Insured screen's msgPayees grid control (to edit an existing Payee)
      // as identified by a non-empty InsuredCurrentPayeeName field or to add a new Payee
      // by the user clicking the Add Payees button on the Insured screen (identified
      // by an *empty* InsuredCurrentPayeeName field).
      if (mtWrapper.getLookupRecordCount() > 0 && !(mfrmMyInsuredForm.frmInsured.getInsuredCurrentPayeeName().equals(""))) {
        // Pull up the Payee on whose name they double-clicked in the Payee grid of the Insured screen
        mtWrapper.goToFirstRecord();
        mtWrapper.getSingleRecord(lngKey1:=mfrmMyInsuredForm.frmInsured.getInsuredCurrentPayeeID(), bSynchLookupRST:=True);
        fnLoadControls();

        // Populate StateInfo structures with data from the STATE_RULE_T
        // row that matches the State Codes as set on the Insured screen.
        fnGetStateInfo_InsdDthResStCd();
        fnGetStateInfo_IssStCd();

        fnSetCommandButtons(true);
      } 
      else {
        // Populate StateInfo structures with data from the STATE_RULE_T
        // row that matches the State Codes as set on the Insured screen.
        fnGetStateInfo_InsdDthResStCd();
        fnGetStateInfo_IssStCd();

        // Go into "Add" mode. (The user clicked the Add Payee button on the Insured screen)
        fnAddRecord();
      }

      mbInLookupMode = false;
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
  private void form_QueryUnload(int pintCancel, int pintUnloadMode) { // TODO: Use of ByRef founded Private Sub Form_QueryUnload(ByRef pintCancel As Integer, ByRef pintUnloadMode As Integer)
    // Comments  :
    // Parameters:
    //    pintCancel     (in/out) - if set to True, refuses to honor the unload request.
    //    pintUnloadMode (in/out) - Identifies what triggered the unload request
    //
    // --------------------------------------------------------------------------------------------
    int intButtonClicked = 0;
    "Form_QueryUnload"
.equals(Const cstrCurrentProc As String);
    try {

      // Set screen name in case errors are reported here or in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      if (modGeneral.bDEBUGAPPTERMINATION) {
        Debug.Print("Entering "+ mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
      }

      if (modGeneral.gbAmProcessingAnAppFatalError) {
        // ALWAYS let the form be unloaded, with no prompts to the user, if shutting
        // down the app due to an application fatal error having been hit.
        if (modGeneral.bDEBUGAPPTERMINATION) {
          Debug.Print("   Early exit from "+ mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc+ " since processing a fatal error.");
        }
        // **TODO:** goto found: GoTo PROC_EXIT;
      }

      if ((Not mbInAddMode) && (Not setIsDirty())) {
        // Let the form be closed if the user is in neither Add nor Update mode.
        // Let the form be closed if the user is in neither Add nor Update mode.
        if (modGeneral.bDEBUGAPPTERMINATION) {
          Debug.Print("   Early exit from "+ mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc+ " since not in Add or Update mode.");
        }
        // **TODO:** goto found: GoTo PROC_EXIT;
      }

      // Since Update (IsDirty) mode can be True while in Add mode, we must check for Add mode first.
      // Otherwise, Adds where the user has started typing (thus setting IsDirty to True) will be
      // treated like an Update, when it should be treated like an Add.
      if (mbInAddMode) {
        if (setIsDirty()) {
          if (modGeneral.bDEBUGAPPTERMINATION) {
            Debug.Print("   Add/Update mode: Prompt the user re: okay to discard pending changes, in "+ mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
          }
          intButtonClicked = modGeneral.gerhApp.reportNonFatal(vbObjectError + modResConstants.gCRES_ALRT_CHANGES_PENDING, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
        } 
        else {
          if (modGeneral.bDEBUGAPPTERMINATION) {
            Debug.Print("   Add mode only (not Update): Do not prompt the user re: okay to discard pending changes, in "+ mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
          }
          intButtonClicked = vbYes;
        }
        if (intButtonClicked == vbYes) {
          if (modGeneral.bDEBUGAPPTERMINATION) {
            Debug.Print("      User opted to discard pending changes in "+ mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
          }

          // If they want to abandon an Add before they started data entry, let them!
          // Redisplay the form with the *first* record now showing
          mtWrapper.getLookupData(mfrmMyInsuredForm.frmInsured.getInsuredClmID());
          if (mtWrapper.getLookupIsAtBOF() && mtWrapper.getLookupIsAtEOF()) {
            // There are no records in the table, so let the form close (If we went into Add
            // mode, the user would never be able to exit the screen!)
          } 
          else {
            if (!modGeneral.gbAmTryingToTerminateTheApp) {
              pintCancel = true;
              mtWrapper.goToFirstRecord();
              //!TODO!: Have to code for the situation where the user is abandoning the
              //        Add of the table's first record...e.g., go into Add mode.
              // Load current record's properties to form's controls, reset
              // navigation buttons and set "rec x of y" label
              fnLoadControls();
              if (modGeneral.bDEBUGAPPTERMINATION) {
                Debug.Print("         Turn off Add mode in "+ mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
              }

              mbInAddMode = false;
              fnSetCommandButtons(true);
              // This **must** be done as the user leaves Add mode, so that the key fields
              // will now be protected to prevent the user from being able to edit them.
              // Editing a key field is allowed only when in Add mode.
              fnSetAvailabilityOfControls();
            }
          }
          mbInLookupMode = false;
        } 
        else {
          if (modGeneral.bDEBUGAPPTERMINATION) {
            Debug.Print("      User opted NOT to discard pending changes in "+ mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
          }

          // User doesn't want to abandon the Add that's still in progress, so ignore the request
          // to close the form and redisplay the form with the same data and with the user's Add
          // still in progress.
          pintCancel = true;
        }
      //' IsDirty (a.k.a. in Update mode)
      } 
      else {
        if (modGeneral.bDEBUGAPPTERMINATION) {
          Debug.Print("   Update mode only (not Add): Prompt the user re: okay to discard pending changes, in "+ mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
        }

        intButtonClicked = modGeneral.gerhApp.reportNonFatal(vbObjectError + modResConstants.gCRES_ALRT_CHANGES_PENDING, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
        if (intButtonClicked == vbYes) {
          if (modGeneral.bDEBUGAPPTERMINATION) {
            Debug.Print("      User opted to discard pending changes in "+ mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
          }

          if (!modGeneral.gbAmTryingToTerminateTheApp) {
            // Abandon their pending changes and redisplay the same record as it *now* appears in
            // the database
            pintCancel = true;
            mtWrapper.getRelativeRecord(mtWrapper.getPayeFullNm(), epdSameRecord);
            //!TODO!: Have to code for the situation where another user deleted the record whose
            //        edits *this* user is abandoning....e.g., go into Add mode
            fnLoadControls();
            if (modGeneral.bDEBUGAPPTERMINATION) {
              Debug.Print("         Turn off Update mode in "+ mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
            }
            setIsDirty(false);

            fnSetCommandButtons(true);
          }
        } 
        else {
          if (modGeneral.bDEBUGAPPTERMINATION) {
            Debug.Print("      User opted NOT to discard pending changes in "+ mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
          }

          // User wants to keep pending changes, so ignore the request to close the form and redisplay
          // the form with the same record showing and with the user's pending changes still pending.
          pintCancel = true;
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
// No resize event handler. This is a non-resizable form.
//////////////////////////////////////////////////////////////////////////////////////////////////


//////////////////////////////////////////////////////////////////////////////////////////////////
  private void form_Unload(int pintCancel) { // TODO: Use of ByRef founded Private Sub Form_Unload(ByRef pintCancel As Integer)
    // Comments  : Close the form
    // Parameters: pvarLastRow
    //             pintLastCol -
    // Modified  :
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

      setIsDirty(false);

      modGeneral.fnFreeObject(mfrmMyInsuredForm);
      modGeneral.fnFreeObject(mtWrapper);
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


  private void ipdPayeClmIntRt_Change() {
    // Comments  : Sets a flag to indicate the current record has been
    //             edited, and thus Update button becomes enabled
    // Parameters: N/A
    // Modified  :
    // --------------------------------------------------
    try {
      "ipdPayeClmIntRt_Change"
.equals(Const cstrCurrentProc As String);


      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      fnInitializeEditMode();

      // Populate msiOverride structure with data from the STATE_RULE_T
      // row that matches the Calc State and Interest Rate
      fnGetStateInfo_Override();
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
  private void ipmPayeSsnTinNum_Change() {
    // Comments  :
    // Parameters: None
    // Modified  :
    // --------------------------------------------------
    try {
      "ipmPayeSsnTinNum_Change"
.equals(Const cstrCurrentProc As String);

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      fnInitializeEditMode();
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
  private void lpcLookupName_Click() {
    // Comments  : Retrieve selected record
    // Parameters: N/A
    //
    // --------------------------------------------------
    "lpcLookupName_Click"
.equals(Const cstrCurrentProc As String);

    try {

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      fnPerformLookup(lpcLookupName);
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
  private void lpcLookupName_GotFocus() {
    //-----------------------------------------------------------------------------
    // Function     lpcLookupName_GotFocus
    // Purpose      Turn on Lookup Mode and drop down the list now that the user has entered this control.
    // Parameters   N/A
    // Returns      N/A
    // Date:        12/19/2001
    //-----------------------------------------------------------------------------
    "lpcLookupName_GotFocus"
.equals(Const cstrCurrentProc As String);
    try {

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      //lpcLookupName.ListDown = True

      mbInLookupMode = true;
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
  private void lpcLookupName_KeyDown(int intKeyCode, int intShift) { // TODO: Use of ByRef founded Private Sub lpcLookupName_KeyDown(ByRef intKeyCode As Integer, ByRef intShift As Integer)
    //-----------------------------------------------------------------------------
    // Function     lpcLookupName_KeyDown
    // Purpose      If the user presses Enter, make it do just what the Click event does
    //              (i.e. display the selected record)
    // Parameters   intKeyCode - ASCII code of key that was pressed
    //              intShift - indicates whether the Shift key was pressed
    // Returns      N/A
    //-----------------------------------------------------------------------------
    "lpcLookupName_KeyDown"
.equals(Const cstrCurrentProc As String);
    try {

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      if (intKeyCode == vbKeyReturn) {
        fnPerformLookup(lpcLookupName);
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
  private void lpcLookupName_LostFocus() {
    //-----------------------------------------------------------------------------
    // Function     lpcLookupName_LostFocus
    // Purpose      Turn off Lookup Mode now that the user has left that control.
    // Parameters   N/A
    // Returns      N/A
    // Date:        12/19/2001
    //-----------------------------------------------------------------------------
    "lpcLookupName_LostFocus"
.equals(Const cstrCurrentProc As String);
    Const(clngFirstRow As Long == 0);
    try {

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      // Display the first (blank) entry in the Lookup control so the
      // user doesn't get confused. Without this code, the Lookup box continues to display
      // the value last selected for lookup purposes, even when the user has since positioned
      // to a different record by virtue of doing a Delete or Add or using the navigation buttons.
      //*TODO:** can't found type for with block
      //*With lpcLookupName
      __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = lpcLookupName;
      w___TYPE_NOT_FOUND.Row = clngFirstRow;
      w___TYPE_NOT_FOUND.ListIndex = clngFirstRow;
      w___TYPE_NOT_FOUND.Action = ActionClearSearchBuffer;

      //fnSearchFPCombo lpcLookupName, gcstrBlankEntry, mcintDisplayCol_lpcLookupName
      lpcLookupName.Refresh;

      mbInLookupMode = false;
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
  private void iptPayeAddrLn1Txt_Change() {
    // Comments  : Limits the number of characters input to that able to
    //             be stored on the CheckFree file
    // Parameters: None
    // Modified  :
    // --------------------------------------------------
    try {
      "iptPayeAddrLn1Txt_Change"
.equals(Const cstrCurrentProc As String);

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      fnInitializeEditMode();
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
  private void iptPayeAddrLn2Txt_Change() {
    // Comments  : Limits the number of characters input to that able to
    //             be stored on the CheckFree file
    // Parameters: None
    // Modified  :
    // --------------------------------------------------
    try {
      "iptPayeAddrLn2Txt_Change"
.equals(Const cstrCurrentProc As String);

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      fnInitializeEditMode();
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
  private void iptPayeCareOfTxt_Change() {
    // Comments  : Limits the number of characters input to that able to
    //             be stored on the CheckFree file
    // Parameters: None
    // Modified  :
    // --------------------------------------------------
    try {
      "iptPayeCareOfTxt_Change"
.equals(Const cstrCurrentProc As String);

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      fnInitializeEditMode();
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
  private void iptPayeCityNmTxt_Change() {
    // Comments  : Limits the number of characters input to that able to
    //             be stored on the CheckFree file
    // Parameters: None
    // Modified  :
    // --------------------------------------------------
    try {
      "iptPayeCityNmTxt_Change"
.equals(Const cstrCurrentProc As String);

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      fnInitializeEditMode();
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
  private void iptPayeFullNm_Change() {
    // Comments  : Limits the number of characters input to that able to
    //             be stored on the CheckFree file
    // Parameters: None
    // Modified  :
    // --------------------------------------------------
    try {
      "iptPayeFullNm_Change"
.equals(Const cstrCurrentProc As String);

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      fnInitializeEditMode();
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
  private void ipcPayeDthbPmtAmt_Change() {
    // Comments  : Set a flag indicating some change has been made.
    //             The formatting of this numeric field won't be
    //             done until the LostFocus event. If done here,
    //             the user's input moves from right to left.
    // Parameters: None
    // Modified  :
    // --------------------------------------------------
    try {
      "ipcPayeDthbPmtAmt_Change"
.equals(Const cstrCurrentProc As String);

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      fnInitializeEditMode();

      //MME START WRUS 4999 - Added the following to ensure that tier 2 (state_rule_t) entries are used if needed.

      if (!IsNumeric(ipcPayeDthbPmtAmt.UnFmtText)) {
        dblScreenDBPaymentValue = 0;
      } 
      else {
        dblScreenDBPaymentValue = ipcPayeDthbPmtAmt.UnFmtText;
      }

      fnResetStateRules();

      //MME END WRUS 4999

      if (!IsNumeric(ipcPayeDthbPmtAmt.UnFmtText)) {
        // If the user cleared the contents, set the screen field to 0
        ipcPayeDthbPmtAmt.UnFmtText = 0;
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
  private void ipdPayeWthldRt_Change() {
    // Comments  :
    // Parameters: None
    // Modified  :
    // --------------------------------------------------
    try {
      "ipdPayeWthldRt_Change"
.equals(Const cstrCurrentProc As String);

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      fnInitializeEditMode();

      if (!IsNumeric(ipdPayeWthldRt.UnFmtText)) {
        // If the user cleared the contents, set the screen field to 0
        ipdPayeWthldRt.UnFmtText = 0;
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
  private void iptPayeZipCd_Change() {
    // Comments  : Limits the number of characters input to that able to
    //             be stored on the CheckFree file
    // Parameters: None
    // Modified  :
    // --------------------------------------------------
    try {
      "txtPayZipCd_Change"
.equals(Const cstrCurrentProc As String);

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      fnInitializeEditMode();
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

//''''''''''''''
//////////////////////////////////////////////////////////////////////////////////////////////////
  private void chkpaye1099ind_Change() {
    // Comments  : Limits the number of characters input to that able to
    //             be stored on the CheckFree file
    // Parameters: None
    // Modified  :
    // --------------------------------------------------
    try {
      "chkpaye1099ind_Change"
.equals(Const cstrCurrentProc As String);

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      fnInitializeEditMode();
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






// %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
//
//   The following procedures exist only to facilitate testing. They should
//   ONLY be called from the Immediate window and not from other procedures
//   in this form or project.
//
//
//   To use these, set a breakpoint at the top of the Form_Initialize event
//   handler. Then, once you've stopped at the breakpoint, type the routine
//   name in the Immediate window.
//       Correct:   TestStub2         Incorrect:  ? TestStub2
//                                                TestStub2()
//
// %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
  private void testStub1() {
    try {
      StateInfo siTemp = null;
      Currency curRate = null;

      siTemp.StrlIntRuleAmt = 0;
      curRate = fnGetInterestRate(siTemp, "test");
      testStub1Sub(siTemp, curRate);

      siTemp.StrlIntRuleAmt = "8";
      curRate = fnGetInterestRate(siTemp, "test");
      testStub1Sub(siTemp, curRate);

      siTemp.StrlIntRuleAmt = "LOAN RATE - just hit Enter";
      curRate = fnGetInterestRate(siTemp, "test");
      testStub1Sub(siTemp, curRate);

      siTemp.StrlIntRuleAmt = "Current Rate - enter a numeric rate";
      curRate = fnGetInterestRate(siTemp, "test");
      testStub1Sub(siTemp, curRate);

      siTemp.StrlIntRuleAmt = "  > of current or 6%";
      curRate = fnGetInterestRate(siTemp, "test");
      testStub1Sub(siTemp, curRate);

      siTemp.StrlIntRuleAmt = "Rate Condition";
      curRate = fnGetInterestRate(siTemp, "test");
      testStub1Sub(siTemp, curRate);
      // **TODO:** label found: PROC_EXIT:;
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    return;
    // **TODO:** label found: PROC_ERR:;
    Debug.Print("Error at line "+ Erl);
    Debug.Print("Error "+ VBA.ex.Number+ ": "+ VBA.ex.Description);
    Debug.Assert(false);
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
}
}



  private void testStub1Sub(StateInfo siIn, Currency curRate) {
    try {
      String strScope = "";

      Debug.Print("RateIn=["+ siIn.StrlIntRuleAmt+ "]    RateUsed=["+ curRate+ "]");
      // **TODO:** label found: PROC_EXIT:;
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    return;
    // **TODO:** label found: PROC_ERR:;
    Debug.Print("Error at line "+ Erl);
    Debug.Print("Error "+ VBA.ex.Number+ ": "+ VBA.ex.Description);
    Debug.Assert(false);
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
}
}

}

private class udtPayeeClone {
    public String payeFullNm;
    public String payeCareOfTxt;
    public String payeAddrLn1Txt;
    public String payeAddrLn2Txt;
    public String payeCityNmTxt;
    public String payeStCd;
    public String payeZipCd;
    public String payeZip4Cd;
    public String payeSsnTinNum;
    public String payeSsnTinTypCd;
    public Date payePmtDt;
    public Double payeDthbPmtAmt;
    public Long clmId;
    public String paye_1099int_ind;//'' BZ4999 October 2013 Non US payee - SXS
    public String payeStCd_UsedInAutoCalc;
    public String payeStCdSpecialInstructions_UsedInAutoCalc;
    public Boolean bClmForResDthInd_UsedInAutoCalc;
    public String insdDthResStCd_UsedInAutoCalc;
    public String insdDthResStCdSpecialInstructions_UsedInAutoCalc;
    public String issStCd_UsedInAutoCalc;
    public String issStCdSpecialInstructions_UsedInAutoCalc;
    public Boolean bClmForCompactCalc_UsedInAutoCalc;
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


case class RmpayeeData(
              id: Option[Int],

              )

object Rmpayees extends Controller with ProvidesUser {

  val rmpayeeForm = Form(
    mapping(
      "id" -> optional(number),

  )(RmpayeeData.apply)(RmpayeeData.unapply))

  implicit val rmpayeeWrites = new Writes[Rmpayee] {
    def writes(rmpayee: Rmpayee) = Json.obj(
      "id" -> Json.toJson(rmpayee.id),
      C.ID -> Json.toJson(rmpayee.id),

    )
  }

  def get(id: Int) = GetAction { implicit request =>
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.LIST_RMPAYEE), { user =>
      Ok(Json.toJson(Rmpayee.get(user, id)))
    })
  }

  def update(id: Int) = PostAction { implicit request =>
    Logger.debug("in rmpayees.update")
    rmpayeeForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      rmpayee => {
        Logger.debug(s"form: ${rmpayee.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.EDIT_RMPAYEE), { user =>
          Ok(
            Json.toJson(
              Rmpayee.update(user,
                Rmpayee(
                       id,

                ))))
        })
      }
    )
  }

  def create = PostAction { implicit request =>
    Logger.debug("in rmpayees.create")
    rmpayeeForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      rmpayee => {
        Logger.debug(s"form: ${rmpayee.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.NEW_RMPAYEE), { user =>
          Ok(
            Json.toJson(
              Rmpayee.create(user,
                Rmpayee(

                ))))
        })
      }
    )
  }

  def delete(id: Int) = PostAction { implicit request =>
    Logger.debug("in rmpayees.delete")
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.DELETE_RMPAYEE), { user =>
      Rmpayee.delete(user, id)
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

case class Rmpayee(
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

object Rmpayee {

  lazy val emptyRmpayee = Rmpayee(
)

  def apply(
      id: Int,
) = {

    new Rmpayee(
      id,
)
  }

  def apply(
) = {

    new Rmpayee(
)
  }

  private val rmpayeeParser: RowParser[Rmpayee] = {
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
        Rmpayee(
              id,
,
              createdAt,
              updatedAt,
              updatedBy)
    }
  }

  def create(user: CompanyUser, rmpayee: Rmpayee): Rmpayee = {
    save(user, rmpayee, true)
  }

  def update(user: CompanyUser, rmpayee: Rmpayee): Rmpayee = {
    save(user, rmpayee, false)
  }

  private def save(user: CompanyUser, rmpayee: Rmpayee, isNew: Boolean): Rmpayee = {
    def getFields = {
      List(

      )
    }
    def throwException = {
      throw new RuntimeException(s"Error when saving ${C.RMPAYEE}")
    }

    DBHelper.saveEx(
      user,
      Register(
        C.RMPAYEE,
        C.ID,
        rmpayee.id,
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

  def load(user: CompanyUser, id: Int): Option[Rmpayee] = {
    loadWhere(user, s"${C.ID} = {id}", 'id -> id)
  }

  def loadWhere(user: CompanyUser, where: String, args : scala.Tuple2[scala.Any, anorm.ParameterValue[_]]*) = {
    DB.withConnection(user.database.database) { implicit connection =>
      SQL(s"SELECT t1.*, t2.${C.FK_NAME} FROM ${C.RMPAYEE} t1 INNER JOIN ${C.???} t2 ON t1.${C.FK_ID} = t2.${C.FK_ID} WHERE $where")
        .on(args: _*)
        .as(rmpayeeParser.singleOpt)
    }
  }

  def delete(user: CompanyUser, id: Int) = {
    DB.withConnection(user.database.database) { implicit connection =>
      try {
        SQL(s"DELETE FROM ${C.RMPAYEE} WHERE ${C.ID} = {id}")
        .on('id -> id)
        .executeUpdate
      } catch {
        case NonFatal(e) => {
          Logger.error(s"can't delete a ${C.RMPAYEE}. ${C.ID} id: $id. Error ${e.toString}")
          throw e
        }
      }
    }
  }

  def get(user: CompanyUser, id: Int): Rmpayee = {
    load(user, id) match {
      case Some(p) => p
      case None => emptyRmpayee
    }
  }
}


// Router

GET     /api/v1/general/rmpayee/:id              controllers.logged.modules.general.Rmpayees.get(id: Int)
POST    /api/v1/general/rmpayee                  controllers.logged.modules.general.Rmpayees.create
PUT     /api/v1/general/rmpayee/:id              controllers.logged.modules.general.Rmpayees.update(id: Int)
DELETE  /api/v1/general/rmpayee/:id              controllers.logged.modules.general.Rmpayees.delete(id: Int)




/**/
