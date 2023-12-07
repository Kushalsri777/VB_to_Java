
import java.util.Date;

public class frmInsured {

  //******************************************************************************
  // Module     : frmInsured
  // Description:
  // Procedures:
  //              ClmCmpCalInd_Click()
  //              cboInsdDthResStCd_Click()
  //              cboIssStCd_Click()
  //              chkClmForResDthInd_Click()
  //              cmdAdd_Click()
  //              cmdAddPayee_Click()
  //              cmdClose_Click()
  //              cmdDelete_Click()
  //              cmdNavigate_Click(ByRef pintIndex As Integer)
  //              cmdPrintReport_Click()
  //              cmdUpdate_Click()
  //              dtpClmInsdDthDt_Change()
  //              dtpClmProofDt_Change()
  //              fnAddRecord()
  //              fnBindControlsToTableWrapper()
  //              fnCalcTotalsForAllPayees()
  //              fnCalcTotalsForAllPayees(ByVal lngClmID As Long) As ADODB.Recordset
  //              fnClearControls()
  //              fnFillPayeeGrid()
  //              fnGetChildren()
  //              fnGetData_IndividualReport() As ADODB.Recordset
  //              fnGetFieldLabel(ByVal strControlName As String) As String
  //              fnGetLobCd() As String
  //              fnGetPayeesNeedingRecalcDueToDeath(lngClmIdIn As Long, dteClmInsdDthDtIn As Date) As Long
  //              fnGetPayeesNeedingRecalcDueToProof(lngClmIdIn As Long, dteClmProofDtIn As Date) As Long
  //              fnGetReportFile() As String
  //              fnInitializeEditMode()
  //              fnGetDefaultPayorCompany(ByVal strClmPolNum As String, ByVal strAdmnSystCd As String) As String
  //              fnLoadCboInsdDthResStCd()
  //              fnLoadCboIssStCd()
  //              fnLoadControls()
  //              fnLoadLpcAdmnSystCd()
  //              fnLoadLpcLookup(ByRef lpcIn As LPLib.fpCombo, ByVal lngLookupType As EnumLookupType)
  //              fnLoadLpcPycoTypCd()
  //              fnLoadRecordWithCalculatedControls()
  //              fnPerformLookup(ByRef lpcIn As LPLib.fpCombo)
  //              fnRefreshAllCombos()
  //              fnSetAvailabilityOfControls(Optional ByVal bChangeFocus = True)
  //              fnSetCommandButtons(ByVal bEnable As Boolean)
  //              fnSetDefaultControlProperties()
  //              fnSetFocusToFirstUpdateableField()
  //              fnSetInsdDthResStCdAvailability()
  //              fnSetNavigationButtons(Optional ByVal bUnconditionalDisable As Boolean = False)
  //              fnSetPropertiesForPayeeScreen(bSendEmptyName As Boolean)
  //              fnSetTxtClmNum()
  //              fnSetupScreenControls()
  //              fnValidData() As Boolean
  //              fnWarningData()
  //              Form_Activate()
  //              Form_Load()
  //              Form_QueryUnload(ByRef pintCancel As Integer, ByRef pintUnloadMode As Integer)
  //              Form_Resize()
  //              Form_Unload(ByRef pintCancel As Integer)
  //              ipmClmInsdSsnNum_Change()
  //              iptClmInsdFirstNm_Change()
  //              iptClmInsdLastNm_Change()
  //              iptClmPolNum_Change()
  //              lpcAdmnSystCd_Change()
  //              lpcAdmnSystCd_GotFocus()
  //              lpcLookupClaim_Click()
  //              lpcLookupClaim_GotFocus()
  //              lpcLookupClaim_KeyDown(ByRef intKeyCode As Integer, ByRef intShift As Integer)
  //              lpcLookupClaim_LostFocus()
  //              lpcLookupName_Click()
  //              lpcLookupName_GotFocus()
  //              lpcLookupName_KeyDown(ByRef intKeyCode As Integer, ByRef intShift As Integer)
  //              lpcLookupName_LostFocus()
  //              lpcLookupSSN_Click()
  //              lpcLookupSSN_GotFocus()
  //              lpcLookupSSN_KeyDown(ByRef intKeyCode As Integer, ByRef intShift As Integer)
  //              lpcLookupSSN_LostFocus()
  //              lpcPycoTypCd_Change()
  //              lpcPycoTypCd_GotFocus()
  //              msgPayees_DblClick()
  //              ClmCmpCalInd_Click()
  //               fnSetCompactFillingCheckBox()
  //
  // Modified   :
  // 03/26/01 BAW Cleaned with Total Visual CodeTools 2000
  // 01/2002  BAW Removed "#If gcfLOOKUP" stuff since we definitely want Lookup capability. (At one
  //              time before v2.2 was released, we thought the performance might be too bad to keep it.)
  //              Also optimized per Project Analyzer (removing dead code, adding "$" to Mid/Space, etc.).
  //              Also updated the cboLoadCboInsdDthResStCd and fnLoadCboLookupClaim to improve performance.
  // Modified:     Berry Kropiwka - Added Compact Calc - 2019-09-27
  // --------------------------------------------------
//Option Explicit
  *Option Compare Binary

  private String mstrScreenName = "";

  private static final Long MCLNGMINFORMWIDTH = 12465;
  private static final Long MCLNGMINFORMHEIGHT = 7500;
  // The following constants identify, for fpCombo controls used as multi-column comboboxes,
  // which column is displayed in the Edit portion of the control (index = mcintDisplayCol_xxxx)
  // and which is saved to a column in its corresponding SQL table (index = mcintStoreCol_xxxx),
  // where xxxx is the fpCombo control's name.
  private static final Integer MCINTDISPLAYCOL_LPCADMNSYSTCD = 0;
  private static final Integer MCINTSTORECOL_LPCADMNSYSTCD = 1;
  private static final Integer MCINTDISPLAYCOL_LPCPYCOTYPCD = 0;
  private static final Integer MCINTSTORECOL_LPCPYCOTYPCD = 1;

  private static final String MCSTRGROUPLOB = "G";
  private static final String MCSTRINDIVIDUALLOB = "I";

  private static final String MCSTRPYCO_SUBSIDIARY = "Subsidiary";
  private static final String MCSTRPYCO_PARENT = "Parent";

  // Leverage/Claimbuilder Project - K723 - 07/05/2014
  private static final String MCSTRPYCO_SLHIC = "SLHIC";
  private static final String MCSTRADMNSYSTSOLAR = "24";

  // The following constants identify, for fpCombo controls used as Lookups,
  // which column is displayed in the Edit portion of the control (index = mcintDisplayCol_xxxx,
  // where xxxx is the fpCombo control's name).
  private static final Integer MCINTDISPLAYCOL_LPCLOOKUPCLAIM = 0;
  private static final Integer MCINTDISPLAYCOL_LPCLOOKUPSSN = 0;
  private static final Integer MCINTDISPLAYCOL_LPCLOOKUPNAME = 0;

  // These constants define the columns within the Lookup/Multi-column combo boxes.
  // These are used to give a name to a given column of the fpCombo control so
  // it can be referenced by name, not by number.
  private static final String MCSTRDISPLAYCOL = "DISPLAY_COL";
  private static final String MCSTRCLMID = "CLM_ID";
  private static final String MCSTRCLMINSDDTHDT = "CLM_INSD_DTH_DT";
  private static final String MCSTRCLMINSDFIRSTNM = "CLM_INSD_FIRST_NM";
  private static final String MCSTRCLMINSDLASTNM = "CLM_INSD_LAST_NM";
  private static final String MCSTRCLMINSDSSNNUM = "CLM_INSD_SSN_NUM";
  private static final String MCSTRCLMNUM = "CLM_NUM";
  private static final String MCSTRADMNSYSTDSC = "ADMN_SYST_DSC";
  private static final String MCSTRADMNSYSTCD = "ADMN_SYST_CD";
  private static final String MCSTRPYCOTYPDSC = "PYCO_TYP_DSC";
  private static final String MCSTRPYCOTYPCD = "PYCO_TYP_CD";

  //-----------------------------------------------------------------------
  // The following Enum is used by fnLoadLpcLookup and denotes which
  // lookup is being populated.
  //-----------------------------------------------------------------------
//*TODO:** enum is translated as a new class at the end of the file Public Enum EnumLookupType


  // mtWrapper is an instance of the table wrapper corresponding to the main table maintained by this form.
  private ctclmClaim mtWrapper;
  // mtPayee is an instance of the table wrapper corresponding to the table that is SUBORDINATE to the
  // main table maintained by this form
  private ctpyePayee mtPayee;


  // Define a constant for each field that may get an error or warning. This
  // should match the text of that control's associated Label control.
  private static final String MCSTRLPCADMNSYSTCDLABEL = "Admin System";
  private static final String MCSTRIPTCLMPOLNUMLABEL = "Policy Number";
  private static final String MCSTRLPCPYCOTYPCDLABEL = "Company Type";
  private static final String MCSTRIPTCLMINSDFIRSTNMLABEL = "First Name";
  private static final String MCSTRIPTCLMINSDLASTNMLABEL = "Last Name";
  private static final String MCSTRCBOISSSTCDLABEL = "Issue State";
  private static final String MCSTRCBOINSDDTHRESSTCDLABEL = "Residence State";
  private static final String MCSTRCHKCLMCMPCALINDLABEL = "Compact Filling";
  private static final String MCSTRCHKCLMFORRESDTHINDLABEL = "Foreign Residence at Death";
  private static final String MCSTRDTPCLMINSDDTHDTLABEL = "Date of Death";
  private static final String MCSTRDTPCLMPROOFDTLABEL = "Date of Proof";
  private static final String MCSTRIPMCLMINSDSSNNUMLABEL = "SSN";
  private static final String MCSTRIPCCLMTOTDTHBPMTAMTLABEL = "Total DB Payment";
  private static final String MCSTRIPCCLMTOTINTAMTLABEL = "Total Claim Interest";
  private static final String MCSTRIPCCLMTOTWTHLDAMTLABEL = "Total Interest Withheld";
  private static final String MCSTRIPCCLMTOTCLMPDAMTLABEL = "Total";

  private static final String MCSTRTXTCLMNUMLABEL = "Claim Number";
  private static final String MCSTRTXTCLMIDLABEL = "Claim ID";


  DBRecordSet mrstPayees = null;

  // mbInLookupMode determines whether the user is in the process of doing a search using the Lookup ComboBox
  private boolean mbInLookupMode = false;

  // mbInAddMode determines whether the user has begun the process of adding a new record to the table.
  // Note that Add mode is independent of Update mode
  private boolean mbInAddMode = false;

  private Control mctlFirstUpdateableField_Add;
  private Control mctlFirstUpdateableField_Upd;

  private String mstrOrigDateOfDeath = "";
  private String mstrOrigDateOfProof = "";

  private int mintAdmnSyst_MinPolNumLength = 0;
  private int mintAdmnSyst_MaxPolNumLength = 0;
  private String mstrAdmnSyst_DfltPycoTypDsc = "";
  private String mstrAdmnSyst_TaxRptgInd = "";


  //------------------------------------------
  //            MEMBER VARIABLES
  //
  // These are used by the Payee screen.
  //------------------------------------------
  // member variable for InsuredClmForResDthInd property
  private boolean m_bInsuredClmForResDthInd = false;
  // member variable for InsuredClmID property
  private int m_lngInsuredClmID = 0;
  // member variable for InsuredClmNum property
  private String m_strInsuredClmNum = "";
  // member variable for InsuredCurrentPayeeName property
  private String m_strInsuredCurrentPayeeName = "";
  // member variable for InsuredCurrentPayeeID property
  private int m_lngInsuredCurrentPayeeID = 0;
  // member variable for InsuredClmInsdDthDt property
  private Date m_dteInsuredClmInsdDthDt = null;
  // member variable for InsuredClmProofDt property
  private Date m_dteInsuredClmProofDt = null;
  // member variable for InsuredLobCd property
  private String m_strInsuredLobCd = "";
  // member variable for InsuredInsdDthResStCd property
  private String m_strInsuredInsdDthResStCd = "";
  // member variable for InsuredIssStCd property
  private String m_strInsuredIssStCd = "";

  // m_bIsDirty corresponds to the public property called IsDirty.
  // All maintenance screens should have this field and that property! When True, it indicates
  // that the user has made --but not yet saved-- changes to a record. The MDI form will query
  // this property if the user opens the File menu, since the Exit option should be disabled if
  // any form has outstanding changes.
  // Be sure to use this variable's corresponding Property Let to change its value.
  // Do **NOT** set m_bIsDirty itself, as this will using the Property Let proc will
  // ensure the Close button caption is always synchronized with the value of the property.
  private boolean m_bIsDirty = false;

  //Private Const variable for the compact filling state code. Used when filter the Issued State and Residence State.
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


//////////////////////////////////////////////////////////////////////////////////////////////////
  public boolean getInsuredClmForResDthInd() {
    "Property Get InsuredClmForResDthInd"
.equals(Const cstrCurrentProc As String);
    try {

      return m_bInsuredClmForResDthInd;
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
  private void setInsuredClmForResDthInd(boolean bValue) {
    "Property Let InsuredClmForResDthInd"
.equals(Const cstrCurrentProc As String);
    try {

      m_bInsuredClmForResDthInd = bValue;
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
  public int getInsuredClmID() {
    "Property Get InsuredClmID"
.equals(Const cstrCurrentProc As String);
    try {

      return m_lngInsuredClmID;
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
  private void setInsuredClmID(int lngValue) {
    "Property Let InsuredClmID"
.equals(Const cstrCurrentProc As String);
    try {

      m_lngInsuredClmID = lngValue;
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

//Y027 07-11-2012
//////////////////////////////////////////////////////////////////////////////////////////////////
  public String getAdminSystemCode() {
    String _rtn = "";
    "Property Get InsuredClmNum"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = lpcAdmnSystCd.Text;
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

  return _rtn;
}


//////////////////////////////////////////////////////////////////////////////////////////////////
  public String getInsuredClmNum() {
    "Property Get InsuredClmNum"
.equals(Const cstrCurrentProc As String);
    try {

      return m_strInsuredClmNum;
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
  private void setInsuredClmNum(String strValue) {
    "Property Let InsuredClmNum"
.equals(Const cstrCurrentProc As String);
    try {

      m_strInsuredClmNum = strValue;
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
  public Date getInsuredClmProofDt() {
    "Property Get InsuredClmProofDt"
.equals(Const cstrCurrentProc As String);
    try {

      return m_dteInsuredClmProofDt;
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
  private void setInsuredClmProofDt(Date dteValue) {
    "Property Let InsuredClmProofDt"
.equals(Const cstrCurrentProc As String);
    try {

      m_dteInsuredClmProofDt = dteValue;
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
  public int getInsuredCurrentPayeeID() {
    "Property Get InsuredCurrentPayeeID"
.equals(Const cstrCurrentProc As String);
    try {

      return m_lngInsuredCurrentPayeeID;
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
  private void setInsuredCurrentPayeeID(int lngValue) {
    "Property Let InsuredCurrentPayeeID"
.equals(Const cstrCurrentProc As String);
    try {

      m_lngInsuredCurrentPayeeID = lngValue;
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
  public String getInsuredCurrentPayeeName() {
    "Property Get InsuredCurrentPayeeName"
.equals(Const cstrCurrentProc As String);
    try {

      return m_strInsuredCurrentPayeeName;
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
  private void setInsuredClmInsdDthDt(Date dteValue) {
    "Property Let InsuredClmInsdDthDt"
.equals(Const cstrCurrentProc As String);
    try {

      m_dteInsuredClmInsdDthDt = dteValue;
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
  public Date getInsuredClmInsdDthDt() {
    "Property Get InsuredClmInsdDthDt"
.equals(Const cstrCurrentProc As String);
    try {

      return m_dteInsuredClmInsdDthDt;
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
  private void setInsuredCurrentPayeeName(String strValue) {
    "Property Let InsuredCurrentPayeeName"
.equals(Const cstrCurrentProc As String);
    try {

      m_strInsuredCurrentPayeeName = strValue;
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
  public String getInsuredInsdDthResStCd() {
    "Property Get InsuredInsdDthResStCd"
.equals(Const cstrCurrentProc As String);
    try {

      return m_strInsuredInsdDthResStCd;
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
  private void setInsuredInsdDthResStCd(String strValue) {
    "Property Let InsuredInsdDthResStCd"
.equals(Const cstrCurrentProc As String);
    try {

      m_strInsuredInsdDthResStCd = strValue;
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
  public String getInsuredIssStCd() {
    "Property Get InsuredIssStCd"
.equals(Const cstrCurrentProc As String);
    try {

      return m_strInsuredIssStCd;
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
  private void setInsuredIssStCd(String strValue) {
    "Property Let InsuredIssStCd"
.equals(Const cstrCurrentProc As String);
    try {

      m_strInsuredIssStCd = strValue;
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
  public String getInsuredLobCd() {
    "Property Get InsuredLobCd"
.equals(Const cstrCurrentProc As String);
    try {

      return m_strInsuredLobCd;
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
  private void setInsuredLobCd(String strValue) {
    "Property Let InsuredLobCd"
.equals(Const cstrCurrentProc As String);
    try {

      m_strInsuredLobCd = strValue;
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
//|                      PRIVATE    Procedures                       |
///                                                                  \
//\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/


//////////////////////////////////////////////////////////////////////////////////////////////////
  private void cboInsdDthResStCd_Click() {
    // Comments  : Sets a flag to indicate the current record has been
    //             edited, and thus Update button becomes enabled
    // Parameters: N/A
    // Modified  :
    // --------------------------------------------------
    try {
      "cboInsdDthResStCd_Click"
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
  private void cboIssStCd_Click() {
    // Comments  : Sets a flag to indicate the current record has been
    //             edited, and thus Update button becomes enabled
    // Parameters: N/A
    // Modified  :
    // --------------------------------------------------
    try {
      "cboIssStCd_Click"
.equals(Const cstrCurrentProc As String);


      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      fnInitializeEditMode();

      // Enable or Disable the Compact Filling check box based on Admin System
      fnSetCompactFillingCheckBox(Me.lpcAdmnSystCd.Text, Me.cboIssStCd.Text);

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
  private void chkClmCmpCalInd_Click() {
    // Comments  : Sets a flag to indicate the current record has been
    //             edited, and thus Update button becomes enabled
    // Parameters: N/A
    // Modified:     Berry Kropiwka - Added Compact Calc - 2019-09-27
    // --------------------------------------------------
    try {
      "chkClmCmpCalInd_Click"
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
  private void chkClmForResDthInd_Click() {
    // Comments  : Sets a flag to indicate the current record has been
    //             edited, and thus Update button becomes enabled
    // Parameters: N/A
    // Modified  :
    // --------------------------------------------------
    try {
      "chkClmForResDthInd_Click"
.equals(Const cstrCurrentProc As String);

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      fnInitializeEditMode();

      // Disable the Residence State combobox if this checkbox is selected; otherwise enable it.
      fnSetInsdDthResStCdAvailability();

      // Blank out the Residence State's selection if that control is now disabled.
      if (chkClmForResDthInd.chrgHourglass.getValue() == vbUnchecked) {
        cboInsdDthResStCd.Text = modGeneral.gCSTRBLANKENTRY;
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
  private void cmdAddPayee_Click() {
    // Comments  : Opens the Payee maintenance screen
    // Parameters:  -
    // Modified  :
    // --------------------------------------------------
    try {
      "cmdAddPayee_Click"
.equals(Const cstrCurrentProc As String);
      Form frmChild = null;
      String strSaveClaimNumber = "";
      chrgHourglass hrgHourglass = null;
      int lngReturnValue = 0;
      String strACF2 = "";

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);


      if (mtWrapper.getLookupRecordCount() > 0) {
        fnSetPropertiesForPayeeScreen(bSendEmptyName:=True);
        // Following statement triggers the Form_Initialize & Form_Load events in frmPayee
        frmChild = new frmPayee();
        // Following statement triggers the Form_Activate event in frmPayee
        frmChild.Show(vbModal);

        hrgHourglass = new chrgHourglass();
        hrgHourglass.setValue(true);

        // Note: You *must* requery the Insured and Payee recordsets to accomodate the possibility
        //       that another user (a) add/changed/deleted one more Payees for the
        //       current Insured and (b) returned to the Insured screen which triggered an update
        //       to the Insured record for the claim-wide totals it carries. If you don't do the
        //       requeries then a -2147217864 "row cannot be located for updating..." error could
        //       occur. So, we'll do the requerying automatically with no visible indication to the
        //       user that it occured unless the requerying revealed that another user deleted the
        //       current claim number and hence the Insured with the next higher claim number will
        //       be displayed (otherwise the same claim remains being displayed).

        strSaveClaimNumber = iptClmPolNum.Text;

        // Do an immediate repaint. This allows the Insured screen to be redrawn BEFORE all
        // the work of requerying and repainting is started. When the requerying/repainting is done,
        // only small parts of the screen (not the whole screen) will need to be repainted. This
        // eliminates the user seeing a very slow repainting.
        Me.Refresh;

        //!TODO! The following looks like unnecessary (i.e. dead) code
        //If txtClmNum <> strSaveClaimNumber Then
        //MsgBox "Another user has deleted the Claim Number (" & strSaveClaimNumber & ") you were viewing.", _
        //       vbOKOnly + vbInformation, mcstrDialogTitle
        //End If

        hrgHourglass.setValue(true);

        fnGetChildren();

        // 01/31/2001 BAW - Add another Refresh to speed up repainting
        Me.Refresh;

        // Totals may have changed. Update the Insured record just in case.
        fnLoadRecordWithCalculatedControls();

        // 01/31/2001 BAW - Add another Refresh to speed up repainting
        Me.Refresh;

        // 01/31/2001 BAW - Add another Refresh to speed up repainting
        Me.Refresh;

        // Determine whether another user updated or deleted the record about to be updated.
        // Note: this multi-user checking is performed on an Update but not an Add.
        lngReturnValue = mtWrapper.checkForAnotherUsersChanges(ewoUpdate, strACF2);

        if (lngReturnValue == vbObjectError + modResConstants.gCRES_INFO_ANOTHER_USER_DELETED_NOT_SAVED) {
          modGeneral.gerhApp.reportNonFatal(vbObjectError + modResConstants.gCRES_INFO_ANOTHER_USER_DELETED_NOT_SAVED, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
          // Discard *this* user's pending changes and show the previous record.
          // Can't use the GetClmNumFromClmId( ) method since the CLAIM_T row
          // doesn't exist and hence a "-2147217900" (Claim ID does not exist) error
          // throws things off.
          mtWrapper.getRelativeRecord(mtWrapper.getClmNum(), epdPreviousRecord);
          // Do NOT bother to check for another UPDATING the record, since all we're doing is
          // updating the total fields. Let the totals update go through.
          //   ElseIf lngReturnValue = vbObjectError + gcRES_INFO_ANOTHER_USER_UPDATED_DISCARDED Then
          //       gerhApp.ReportNonFatal vbObjectError + gcRES_INFO_ANOTHER_USER_UPDATED_DISCARDED, _
          //                           mstrScreenName & gcstrDOT & cstrCurrentProc, _
          //                           Trim$(strACF2)
          //       ' Discard *this* user's pending changes by re-retrieving the current record
          //       ' as it currently looks on the database and refreshing the lookup recordset.
          //       ' Can't use the GetClmNumFromClmId( ) method since the CLAIM_T row
          //       ' doesn't exist and hence a "-2147217900" (Claim ID does not exist) error
          //       ' throws things off.
          //       .GetRelativeRecord .ClmNum, epdSameRecord
        } 
        else {
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
        // they reflects this and other users' changes. Then call fnLoadControls
        // to make sure comboboxes' selection is reset as appropriate
        fnRefreshAllCombos();

        // Have to call fnLoadControls here, like in cmdAdd_Click and cmdDelete_Click and cmdUpdate_Click,
        // to ensure refreshed comboboxes have their previous value still selected.
        if (mtWrapper.getLookupRecordCount() > 0) {
          // Ensure the on-screen controls reflect the record just added/updated, in case the
          // DBMS altered it in some way, e.g., determining an Identity column value and
          // getting the most up-to-date Last Updated info. This also sets the navigation
          // buttons and updates the "record x of y" label
          fnLoadControls();
          fnSetCommandButtons(true);
        } 
        else {
          fnAddRecord();
        }
      } 
      else {
        // 2003 = There is no current Insured record. The Payee screen cannot be opened.
        modGeneral.gerhApp.reportNonFatal(vbObjectError + modResConstants.gCRES_WARN_NO_CURR_INSURED, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
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
    // Terminate the Payee form, removing it from the Forms collection
    modGeneral.fnFreeObject(frmChild);

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
    // Purpose     Will close the screen
    //
    //             NOTE: The logic in this function should closely resemble that
    //                   in the Form_QueryUnload event handler!
    // Parameters: N/A
    // Returns:    N/A
    // Modified:
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
    "cmdDelete_Click"
.equals(Const cstrCurrentProc As String);
    int intButtonClicked = 0;
    int lngReturnValue = 0;
    String strACF2 = "";
    chrgHourglass hrgHourglass = null;
    try {

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      // .......................................................................
      // Make sure the user really, really, really wants to delete this record.
      // .......................................................................
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
        // Can't use the GetClmNumFromClmId( ) method since the CLAIM_T row
        // doesn't exist and hence a "-2147217900" (Claim ID does not exist) error
        // throws things off.
        mtWrapper.getRelativeRecord(mtWrapper.getClmNum(), epdPreviousRecord);
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
  private void cmdPrintReport_Click() {
    // Comments  :
    // Parameters:  -
    // Modified  :
    // --------------------------------------------------
    try {
      "cmdPrintReport_Click"
.equals(Const cstrCurrentProc As String);
      CRAXDRT.Database crDB = null;
      chrgHourglass hrgHourglass = null;
      New rstReportData = null; ADODB.Recordset

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      hrgHourglass = new chrgHourglass();
      hrgHourglass.setValue(true);

      // Moved the instantiation of the Crystal object here (from modStartup) as a conditional instantiation
      // since this CreateObject invocation is such a pig per VB Watch Profiler.
      if ((modReporting.gcrxApp == null)) {
        modReporting.gcrxApp = CreateObject("CrystalRuntime.Application");
      }

      modReporting.gcReportToPrint = modReporting.gcrxApp.OpenReport(fnGetReportFile());
      crDB = modReporting.gcReportToPrint.Database;

      // Build an ADODB.Recordset containing the info to appear on the report
      rstReportData = fnGetData_IndividualReport();

      // Tell the report the where its data is coming from, e.g., the
      // ADODB.Recordset just created
      modReporting.gcReportToPrint.Database.SetDataSource(Data:=rstReportData, dataTag:=3, tableNumber:=1);

      // ...............................................................................
      // Set formula field(s) in the report that supply additional info that
      // is not in the recordset (typically singularly-occuring data)
      // ...............................................................................
      modReporting.fnSetFormulaField("formulaReportName", "Individual Report");
      //' No criteria for this report
      modReporting.fnSetFormulaField("formulaReportPeriodDescript", "");

      // ...............................................................................
      // Tell the report where the data is coming from (overriding whatever might
      // have been set at design-time). All of the following is necessary since
      // the location and Connect string set within the .RPT itself may not be
      // accurate in a production environment (or even on another developer's PC)
      // ...............................................................................
      crDB.SetDataSource(rstReportData);
      //With crDB.Tables.Item(1)
      //    .SetLogOnInfo pServerName:=strDBPath, pDatabaseName:=vbNullString, _
      //                       pUserID:=vbNullString, pPassword:=vbNullString
      //End With

      if (!(hrgHourglass == null)) {
        hrgHourglass.setValue(false);
      }

      // Print report to modal Viewer window
      modReporting.fnViewReport();

      // Make sure this window is shown on top of all other windows in the app
      // after the Viewer window is closedlkj  lkj klj
      modGeneral.fnSetTopmostWindow(this, bTopmost:=True);
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    modGeneral.fnFreeObject(hrgHourglass);
    modGeneral.fnFreeObject(crDB);

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
    // Params:      N/A
    // Returns:     N/A
    // Modified  :
    // Modified:     Berry Kropiwka - Added Compact Calc - 2019-09-27
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
      mtWrapper.setClmNum(txtClmNum);

      // Just get the ADMN_SYST_CD (column 1) of the selected row in lpcAdmnSystCd
      lpcAdmnSystCd.Col = MCINTSTORECOL_LPCADMNSYSTCD;
      mtWrapper.setAdmnSystCd(lpcAdmnSystCd.ColText);

      mtWrapper.setClmPolNum(iptClmPolNum.Text);
      mtWrapper.setClmForResDthInd((chkClmForResDthInd.chrgHourglass.getValue() == vbChecked));
      mtWrapper.setClmCompactClcnInd((chkClmCmpCalInd.chrgHourglass.getValue() == vbChecked));

      // Just get the PYCO_TYP_CD (column 1) of the selected row in lpcPycoTypCd
      lpcPycoTypCd.Col = MCINTSTORECOL_LPCPYCOTYPCD;
      mtWrapper.setPycoTypCd(lpcPycoTypCd.ColText);

      mtWrapper.setClmInsdFirstNm(iptClmInsdFirstNm.Text);
      mtWrapper.setClmInsdLastNm(iptClmInsdLastNm.Text);

      // cboIssStCd corresponds to a Nullable field, so accommodate Nulls
      if (cboIssStCd.Text == modGeneral.gCSTRBLANKENTRY) {
        mtWrapper.setIssStCd("");
      } 
      else {
        mtWrapper.setIssStCd(cboIssStCd.Text);
      }
      // cboInsdDthResStCd corresponds to a Nullable field, so accommodate Nulls
      if (cboInsdDthResStCd.Text == modGeneral.gCSTRBLANKENTRY) {
        mtWrapper.setInsdDthResStCd("");
      } 
      else {
        mtWrapper.setInsdDthResStCd(cboInsdDthResStCd.Text);
      }

      mtWrapper.setClmInsdDthDt(dtpClmInsdDthDt.chrgHourglass.getValue());
      mtWrapper.setClmProofDt(dtpClmProofDt.chrgHourglass.getValue());
      //' Use .UnFmtText to get rid of mask characters in fpMask control
      mtWrapper.setClmInsdSsnNum(ipmClmInsdSsnNum.UnFmtText);

      //' Use .Value to get unformatted value of fpCurrency control
      mtWrapper.setClmTotDthbPmtAmt(ipcClmTotDthbPmtAmt.chrgHourglass.getValue());
      mtWrapper.setClmTotIntAmt(ipcClmTotIntAmt.chrgHourglass.getValue());
      mtWrapper.setClmTotWthldAmt(ipcClmTotWthldAmt.chrgHourglass.getValue());
      mtWrapper.setClmTotClmPdAmt(ipcClmTotClmPdAmt.chrgHourglass.getValue());

      mtWrapper.setLstUpdtUserId(modGeneral.gconAppActive.getLastLogOnUserID());
      mtWrapper.setLstUpdtDtm(Now);

      // These will propagate back an error if the Insert/Update failed.
      if (mbInAddMode) {
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
          // Can't use the GetClmNumFromClmId( ) method since the CLAIM_T row
          // doesn't exist and hence a "-2147217900" (Claim ID does not exist) error
          // throws things off.
          mtWrapper.getRelativeRecord(mtWrapper.getClmNum(), epdPreviousRecord);
        } 
        else if (lngReturnValue == vbObjectError + modResConstants.gCRES_INFO_ANOTHER_USER_UPDATED_DISCARDED) {
          modGeneral.gerhApp.reportNonFatal(vbObjectError + modResConstants.gCRES_INFO_ANOTHER_USER_UPDATED_DISCARDED, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc, strACF2.trim());
          // Discard *this* user's pending changes by re-retrieving the current record
          // as it currently looks on the database and refreshing the lookup recordset
          // Can't use the GetClmNumFromClmId( ) method since the CLAIM_T row
          // doesn't exist and hence a "-2147217900" (Claim ID does not exist) error
          // throws things off.
          mtWrapper.getRelativeRecord(mtWrapper.getClmNum(), epdSameRecord);
        } 
        else {
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
        fnSetCommandButtons(true);
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
  private void dtpClmInsdDthDt_Change() {
    // Comments  : Since this field was just changed, reset
    //             Enabled property on command and navigation
    //             buttons as appropriate given that the user
    //             is in the middle of updating a record.
    // Parameters:  -
    // Modified  :
    // --------------------------------------------------
    try {
      "dtpClmInsdDthDt_Change"
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
  private void dtpClmProofDt_Change() {
    // Comments  : Since this field was just changed, reset
    //             Enabled property on command and navigation
    //             buttons as appropriate given that the user
    //             is in the middle of updating a record.
    // Parameters:  -
    // Modified  :
    // --------------------------------------------------
    try {
      "dtpClmProofDt_Change"
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
  private void fnAddRecord() {
    // Comments  : This function handles adding a new record. It is called
    //             by cmdAdd_Click (when the user clicks the Add button)
    //             and by cmdDelete_Click (when the last record in the
    //             recordset is deleted)
    // Parameters: N/A
    // Returns   : N/A
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

      if (modGeneral.bDEBUGAPPTERMINATION) {
        Debug.Print("   Turning off Update mode (#1) in "+ mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
      }
      setIsDirty(false);

      // Initialize Payee recordset to avoid run-time error 91 - object variable or with block variable not set
      mrstPayees = new ADODB.Recordset();
      // Only show the 1st row (column headers) in Payee Grid
      msgPayees.Rows = 1;

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
    // Modified:     Berry Kropiwka - Added Compact Calc - 2019-09-27
    //-----------------------------------------------------------------------------
    "fnBindControlsToTableWrapper"
.equals(Const cstrCurrentProc As String);
    try {

      lpcAdmnSystCd.Tag = "AdmnSystCd";
      iptClmPolNum.Tag = "ClmPolNum";
      lpcPycoTypCd.Tag = "PycoTypCd";
      iptClmInsdFirstNm.Tag = "ClmInsdFirstNm";
      iptClmInsdLastNm.Tag = "ClmInsdLastNm";
      cboIssStCd.Tag = "IssStCd";
      cboInsdDthResStCd.Tag = "InsdDthResStCd";
      dtpClmInsdDthDt.Tag = "ClmInsdDthDt";
      dtpClmProofDt.Tag = "ClmProofDt";
      ipmClmInsdSsnNum.Tag = "ClmInsdSsnNum";
      ipcClmTotDthbPmtAmt.Tag = "ClmTotDthbPmtAmt";
      ipcClmTotIntAmt.Tag = "ClmTotIntAmt";
      ipcClmTotWthldAmt.Tag = "ClmTotWthldAmt";
      ipcClmTotClmPdAmt.Tag = "ClmTotClmPdAmt";
      txtClmNum.Tag = "ClmNum";
      chkClmForResDthInd.Tag = "ClmForResDthInd";
      chkClmCmpCalInd.Tag = "ClmCompactClcnInd";

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
  private DBRecordSet fnCalcTotalsForAllPayees(int lngClmID) {
    // Comments  : This function will add up all of the Payee for each
    //             policy/claim to produce totals
    // Parameters:
    //     lngClmId (in) - the CLM_ID of the desired Claim
    // Returns:     A disconnected ADODB.Recordset containing calculated
    //              columns for the specified key
    // Modified  :
    // --------------------------------------------------
    "fnCalcTotalsForAllPayees"
.equals(Const cstrCurrentProc As String);
    //' Stored procedure to execute
    "dbo.proc_payee_totals_for_claim"
.equals(Const cstrSproc As String);
    DBRecordSet rstTemp = null;
    ADODB.Parameter prmReturnValue = null;
    ADODB.Parameter prmClmId = null;
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

      // ---Parameter #2---
      // Define the CLM_ID parameter
      prmClmId = w_aDOCommand.CreateParameter(Name:="@clm_id", Type:=adInteger, Direction:=adParamInput, chrgHourglass.getValue():=lngClmID);
      w_aDOCommand.Parameters.Append(prmClmId);

      rstTemp = w_aDOCommand.Execute();

      rstTemp.ActiveConnection = null;
      return rstTemp;

      // Use fnZeroIfNull to accommodate Nulls in case there are no Payees yet defined for this claim
      ipcClmTotDthbPmtAmt.Text = modDataConversion.fnZeroIfNull(!ctclmClaim.getClmTotDthbPmtAmt());
      ipcClmTotWthldAmt.Text = modDataConversion.fnZeroIfNull(!ctclmClaim.getClmTotWthldAmt());
      ipcClmTotIntAmt.Text = modDataConversion.fnZeroIfNull(!ctclmClaim.getClmTotIntAmt());
      ipcClmTotClmPdAmt.Text = modDataConversion.fnZeroIfNull(!ctclmClaim.getClmTotClmPdAmt());
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    modGeneral.fnFreeRecordset(rstTemp);
    modGeneral.fnFreeObject(prmReturnValue);
    modGeneral.fnFreeObject(prmClmId);
    modGeneral.fnFreeObject(adwTemp);

    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
    }
    return null;
    // **TODO:** label found: PROC_ERR:;
    switch (prmReturnValue) {
      case  modResConstants.gCRES_NERR_REC_NOT_FOUND:
        // 4027 = The specified record was not found in the database (@@1).
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_REC_NOT_FOUND, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc, "Claim ID "+ RTrim$(lngClmID)+ "/Claim Number "+ mtWrapper.getClmNumFromClmID(lngClmID));
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
        break;

      case  modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO:
        // 4028 = An error occurred while attempting to @@1 this record.
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
  private void fnClearControls() {
    // Comments:   Initializes screen controls in order to add a new record
    // Parameters: N/A
    // Returns:    N/A
    // Modified:
    // Modified:     Berry Kropiwka - Added Compact Calc - 2019-09-27
    // --------------------------------------------------
    "fnClearControls"
.equals(Const cstrCurrentProc As String);
    Const(cintZero As Integer == 0);
    Const(clngFirstEntry As Long == 0);
    Control ctl = null;
    Object varDefaultValue = null;
    String strSavedMask = "";

    try {

      // Hide updates to the window until we're done. This avoids ugly screen flickering
      fnWindowLock(Me.cbrfBrowseFolder.setHWnd());

      chkClmForResDthInd.chrgHourglass.setValue(vbUnchecked);
      chkClmCmpCalInd.chrgHourglass.setValue(vbUnchecked);

      // Enable or Disable the Compact Filling check box based on Admin System
      fnSetCompactFillingCheckBox(Me.lpcAdmnSystCd.Text, Me.cboIssStCd.Text);

      // Select the first entry (the blank row) in the Admin System's fpCombo control.
      modComboBox.fnSearchFPCombo(lpcAdmnSystCd, modGeneral.gCSTRBLANKENTRY, MCINTSTORECOL_LPCADMNSYSTCD);

      iptClmPolNum.Text = "";

      // Select the first entry (the blank row) in the Company Type's fpCombo control.
      modComboBox.fnSearchFPCombo(lpcPycoTypCd, modGeneral.gCSTRBLANKENTRY, MCINTSTORECOL_LPCPYCOTYPCD);

      iptClmInsdFirstNm.Text = "";
      iptClmInsdLastNm.Text = "";

      if (cboIssStCd.ListCount > 0) {
        //' Select first (blank) entry
        cboIssStCd.ListIndex = 0;
      } 
      else {
        modGeneral.gerhApp.reportNonFatal(vbObjectError + modResConstants.gCRES_WARN_CBO_IS_EMPTY, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc, MCSTRCBOISSSTCDLABEL);
      }

      if (cboInsdDthResStCd.ListCount > cintZero) {
        //' Select first (blank) entry
        cboInsdDthResStCd.ListIndex = cintZero;
      } 
      else {
        modGeneral.gerhApp.reportNonFatal(vbObjectError + modResConstants.gCRES_WARN_CBO_IS_EMPTY, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc, MCSTRCBOINSDDTHRESSTCDLABEL);
      }

      // DateTimePicker controls (dtpClmInsdDthDt and dtpClmProofDt) will
      // automatically be set to today's date. Cannot set them to Null
      // unless their CheckBox property is set to True.
      dtpClmInsdDthDt.chrgHourglass.setValue(Date);
      dtpClmProofDt.chrgHourglass.setValue(Date);


      // NOTE: For MaskEdBox controls, have to remove mask before clearing out the control
      //       since the vbNullString value doesn't match the mask specification.
      strSavedMask = ipmClmInsdSsnNum.ctclmClaim.getMask();
      ipmClmInsdSsnNum.ctclmClaim.setMask("");
      ipmClmInsdSsnNum.Text = "";
      ipmClmInsdSsnNum.ctclmClaim.setMask(strSavedMask);

      // ' Select the "VUL" entry in the Product Family ComboBox, if present, otherwise select
      // ' the first entry. If the ComboBox is empty, display a message to the user to warn them of
      // ' unpredictible behavior.
      // If cboPfamCd.ListCount > 0 Then
      //     lngEntryFoundSlot = fnFindStringComboBox(cboIn:=cboPfamCd, strSearchIn:="VUL     ", bDoExactSearch:=True)
      //     If lngEntryFoundSlot = clngNotFound Then
      //         cboPfamCd.ListIndex = clngFirstEntry
      //     Else
      //         cboPfamCd.ListIndex = lngEntryFoundSlot
      //     End If
      // Else
      //     gerhApp.ReportNonFatal vbObjectError + gcRES_WARN_CBO_IS_EMPTY, _
      //                            mstrScreenName & gcstrDOT & cstrCurrentProc, _
      //                            mcstrCboPfamCdLabel
      // End If

      ipcClmTotDthbPmtAmt.Text = cintZero;
      ipcClmTotIntAmt.Text = cintZero;
      ipcClmTotWthldAmt.Text = cintZero;
      ipcClmTotClmPdAmt.Text = cintZero;

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
            else if ((TypeOf ctl Is fpMask)) {
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
  private void fnFillPayeeGrid() {
    // Comments  : Loads the MSFlexGrid control with
    //             Payee data for the current Insured
    // Called By : fnGetChildren()
    //
    // Parameters:  -
    // Modified  :
    // --------------------------------------------------
    try {
      "fnFillPayeeGrid"
.equals(Const cstrCurrentProc As String);
      int intRecordCounter = 0;


      //!TODO! Change to use vsflexgrid
      //*TODO:** can't found type for with block
      //*With msgPayees
      __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = msgPayees;
      // Set Rows to reflect # of records +1 (for header row)
      w___TYPE_NOT_FOUND.Rows = mrstPayees.RecordCount + 1;
      // Fill in columns of grid per current recordset row
      for (intRecordCounter = 1; intRecordCounter <= mrstPayees.RecordCount; intRecordCounter++) {
        w___TYPE_NOT_FOUND.Row = intRecordCounter;

        // Column 1 - Counter
        w___TYPE_NOT_FOUND.Col = 0;
        w___TYPE_NOT_FOUND.Text = intRecordCounter;
        // Column 2 - Payee Full Name
        w___TYPE_NOT_FOUND.Col = 1;
        w___TYPE_NOT_FOUND.Text = mrstPayees!paye_full_nm;
        // Column 3 - Address Line 1
        w___TYPE_NOT_FOUND.Col = 2;
        w___TYPE_NOT_FOUND.Text = modDataConversion.fnZLSIfNull(mrstPayees!paye_addr_ln1_txt);
        // Column 4 - Address Line2
        w___TYPE_NOT_FOUND.Col = 3;
        w___TYPE_NOT_FOUND.Text = modDataConversion.fnZLSIfNull(mrstPayees!paye_addr_ln2_txt);
        // Column 5 - Payee Residence State
        w___TYPE_NOT_FOUND.Col = 4;
        w___TYPE_NOT_FOUND.Text = modDataConversion.fnZLSIfNull(mrstPayees!calc_st_cd);
        // Column 6 - Date Of Payment
        w___TYPE_NOT_FOUND.Col = 5;
        w___TYPE_NOT_FOUND.Text = mrstPayees!paye_pmt_dt;
        // Column 7 - TIN/SSN
        w___TYPE_NOT_FOUND.Col = 6;
        w___TYPE_NOT_FOUND.Text = modDataConversion.fnZLSIfNull(mrstPayees!paye_ssn_tin_num);
        //!TODO! Change to use meta data for formatting!
        // Column 8 - Interest Amt
        w___TYPE_NOT_FOUND.Col = 7;
        w___TYPE_NOT_FOUND.Text = Format$(mrstPayees!paye_clm_int_amt, "###,###,##0.00");
        // Column 9 - Total Claim Amt for Payee
        w___TYPE_NOT_FOUND.Col = 8;
        w___TYPE_NOT_FOUND.Text = Format$(mrstPayees!paye_clm_pd_amt, "###,###,##0.00");
        // Column 10 - DB Payment
        w___TYPE_NOT_FOUND.Col = 9;
        w___TYPE_NOT_FOUND.Text = Format$(mrstPayees!paye_dthb_pmt_amt, "###,###,##0.00");
        // Column 11 - Interest Rate
        w___TYPE_NOT_FOUND.Col = 10;
        w___TYPE_NOT_FOUND.Text = Format$(mrstPayees!paye_clm_int_rt, "###,##0.00000");
        // Column 12 - Withholding Rate
        w___TYPE_NOT_FOUND.Col = 11;
        w___TYPE_NOT_FOUND.Text = Format$(mrstPayees!paye_wthld_rt, "###,##0.00000");
        // Column 13 - Interest Withheld
        // TotalAmt is reduced by the Withheld Amt, so show Withheld
        // Amt as a negative number. (It is stored as a positive number.)
        w___TYPE_NOT_FOUND.Col = 12;
        w___TYPE_NOT_FOUND.Text = Format$(mrstPayees!paye_wthld_amt, "(###,###,##0.00)");
        // Column 14 - Payee ID
        // This is needed so the Insured screen can tell the Payee
        // screen which Payee to display
        w___TYPE_NOT_FOUND.Col = 13;
        w___TYPE_NOT_FOUND.Text = mrstPayees!paye_id;
        // Make the width=0 to effectively hide it
        w___TYPE_NOT_FOUND.ColWidth(13) = 0;

        // Read next record in recordset and loop
        mrstPayees.cadwADOWrapper.moveNext();
      }

      //' 1st (non-column header) row
      w___TYPE_NOT_FOUND.Row = 1;
      //' 2nd column - Payee name
      w___TYPE_NOT_FOUND.Col = 1;

      fnCalcTotalsForAllPayees(mtWrapper.getClmId());
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
  private void fnGetAdminSysMetadata() {
    // Comments  : This function looks up admin system metadata based on the current
    //             value selected in the Admin System combobox.
    // Parameters: N/A
    // Returns:     N/A. However it sets some module-level variables.
    // Modified  :
    // --------------------------------------------------
    "fnGetAdminSysMetadata"
.equals(Const cstrCurrentProc As String);
    //' Stored procedure to execute
    "dbo.proc_admin_system_select2"
.equals(Const cstrSproc As String);
    ADODB.Parameter prmReturnValue = null;
    ADODB.Parameter prmAdmnSystCd = null;
    ADODB.Parameter prmMinLength = null;
    ADODB.Parameter prmMaxLength = null;
    ADODB.Parameter prmDfltPycoTypDsc = null;
    ADODB.Parameter prmTaxRptgInd = null;
    cadwADOWrapper adwTemp = null;

    try {

      // Set default values in case of error
      mintAdmnSyst_MinPolNumLength = 1;
      mintAdmnSyst_MaxPolNumLength = mtWrapper.getMaxCharacters(iptClmPolNum.Tag);
      mstrAdmnSyst_TaxRptgInd = "";
      mstrAdmnSyst_DfltPycoTypDsc = "";

      // Just get the ADMN_SYST_CD (column 1) of the selected row in lpcAdmnSystCd.
      // If it hasn't been input yet (i.e. it is still blank) then just accept default
      // values set above and bypass the sproc call. This avoids a SQL error hit when
      // passing an invalid value into the @admn_syst_cd parameter.
      lpcAdmnSystCd.Col = MCINTSTORECOL_LPCADMNSYSTCD;
      if ((lpcAdmnSystCd.ColText == "") || (modGeneral.gCSTRBLANKENTRY.equals(lpcAdmnSystCd.ColText))) {
        // **TODO:** goto found: GoTo PROC_EXIT;
      }

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
      // Define the ADMN_SYST_CD input parameter
      prmAdmnSystCd = w_aDOCommand.CreateParameter(Name:="@admn_syst_cd", Type:=adChar, Direction:=adParamInput, chrgHourglass.getValue():=lpcAdmnSystCd.ColText, Size:=2);
      w_aDOCommand.Parameters.Append(prmAdmnSystCd);

      // ---Parameter #3---
      // Define the MIN_LENGTH output parameter
      prmMinLength = w_aDOCommand.CreateParameter(Name:="@MinLength", Type:=adSmallInt, Direction:=adParamOutput, Size:=2);
      w_aDOCommand.Parameters.Append(prmMinLength);

      // ---Parameter #4---
      // Define the MAX_LENGTH output parameter
      prmMaxLength = w_aDOCommand.CreateParameter(Name:="@MaxLength", Type:=adSmallInt, Direction:=adParamOutput, Size:=2);
      w_aDOCommand.Parameters.Append(prmMaxLength);

      // ---Parameter #5---
      // Define the DFLT_PYCO_TYP_DSC output parameter
      prmDfltPycoTypDsc = w_aDOCommand.CreateParameter(Name:="@DfltPycoTypDsc", Type:=adVarChar, Direction:=adParamOutput, Size:=60);
      w_aDOCommand.Parameters.Append(prmDfltPycoTypDsc);

      // ---Parameter #6---
      // Define the TAX_RPTG_IND output parameter
      prmTaxRptgInd = w_aDOCommand.CreateParameter(Name:="@TaxRptgInd", Type:=adChar, Direction:=adParamOutput, Size:=1);
      w_aDOCommand.Parameters.Append(prmTaxRptgInd);

      w_aDOCommand.Execute;


      mintAdmnSyst_MinPolNumLength = prmMinLength.value;
      mintAdmnSyst_MaxPolNumLength = prmMaxLength.value;
      mstrAdmnSyst_TaxRptgInd = prmTaxRptgInd.value;
      mstrAdmnSyst_DfltPycoTypDsc = prmDfltPycoTypDsc.value;
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    modGeneral.fnFreeObject(prmReturnValue);
    modGeneral.fnFreeObject(prmAdmnSystCd);
    modGeneral.fnFreeObject(prmMinLength);
    modGeneral.fnFreeObject(prmMaxLength);
    modGeneral.fnFreeObject(prmDfltPycoTypDsc);
    modGeneral.fnFreeObject(prmTaxRptgInd);

    modGeneral.fnFreeObject(adwTemp);

    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
    }
    return;
    // **TODO:** label found: PROC_ERR:;
    switch (prmReturnValue) {
      case  modResConstants.gCRES_NERR_REC_NOT_FOUND:
        // 4027 = The specified record was not found in the database (@@1).
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_REC_NOT_FOUND, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc, MCSTRLPCADMNSYSTCDLABEL+ ": "+ lpcAdmnSystCd.ColText);
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
        break;

      case  modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO:
        // 4028 = An error occurred while attempting to @@1 this record.
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc, "locate the "+ MCSTRLPCADMNSYSTCDLABEL+ " metadata for");
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
  private void fnGetChildren() {
    // Comments  : Loads data associated from tables that are
    //             subordinate (i.e. children) to the table
    //             supplying the main data for this form
    // Parameters:  -
    // Modified  :
    // --------------------------------------------------
    try {
      "fnGetChildren"
.equals(Const cstrCurrentProc As String);


      // --- Build the Recordset object for Payee data (mrstPayees) ---
      //     that's associated with the current Insured.

      mrstPayees = mtPayee.getPayeesForClaim(mtWrapper.getClmId());

      // Load MSFlexGrid with Payee records, if any. Disallow Delete
      // of Insured/Claim record if there are Payee records. (The user
      // must delete Payees before attempting to delete the Insured/Claim.)
      if (mrstPayees.RecordCount > 0) {
        fnFillPayeeGrid();
      } 
      else {
        // Only show the 1st row (column headers)
        msgPayees.Rows = 1;
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
  private DBRecordSet fnGetData_IndividualReport() {
    //--------------------------------------------------------------------------
    // Procedure:   fnGetData_IndividualReport
    // Description: Builds a recordset containing data needed to send
    //              to the .RPT file associated with a report.
    //
    //
    // Parameters:  N/A
    //
    // Returns:     A disconnected ADODB.Recordset
    //-----------------------------------------------------------------------------
    "fnGetData_IndividualReport"
.equals(Const cstrCurrentProc As String);
    "dbo.IndividualReport_v"
.equals(Const cstrSQLView As String);
    String strSQL = "";
    String strWhereClmID = "";
    String strOrderBy = "";
    DBRecordSet rstTemp = null;

    try {

      strWhereClmID = " WHERE clm_id = "+ ((Integer) mtWrapper.getClmId()).toString()+ "\\n";
      strOrderBy = " ORDER BY clm_id, paye_full_nm";

      strSQL = "SELECT * from "+ cstrSQLView+ strWhereClmID+ strOrderBy;

      rstTemp = modGeneral.gadwApp.execute_SQL_AsRST(modGeneral.gconAppActive, strSQL);
      #If DEBUG_RST Then;
      Debug.Print("In "+ cstrCurrentProc+ ", "+ CStr(rstTemp.RecordCount)+ " records were retrieved in the rst.");
      Debug.Print("SQL statement is: "+ "\\n"+ strSQL);
      #End If;

      // Disconnect the recordset
      rstTemp.ActiveConnection = null;

      return rstTemp;
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here

    // DO NOT do "fnFreeRecordset rstTemp" since this will cause the recordset
    // returned by this function to be wiped out as well!

    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
    }
    return null;
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
    // Modified:     Berry Kropiwka - Added Compact Calc - 2019-09-27

    //!CUSTOMIZE!  There should be one Case statement for each control that
    //             corresponds to a table column. Each Case statement should
    //             reference a Const literal that indicates how the control is
    //             labelled on-screen.

    "fnGetFieldLabel"
.equals(Const cstrCurrentProc As String);

    try {

      switch (strControlName) {
        case  "lpcAdmnSystCd":
          _rtn = MCSTRLPCADMNSYSTCDLABEL;
          break;

        case  "iptClmPolNum":
          _rtn = MCSTRIPTCLMPOLNUMLABEL;
          break;

        case  "lpcPycoTypCd":
          _rtn = MCSTRLPCPYCOTYPCDLABEL;
          break;

        case  "iptClmInsdFirstNm":
          _rtn = MCSTRIPTCLMINSDFIRSTNMLABEL;
          break;

        case  "iptClmInsdLastNm":
          _rtn = MCSTRIPTCLMINSDLASTNMLABEL;
          break;

        case  "cboIssStCd":
          _rtn = MCSTRCBOISSSTCDLABEL;
          break;

        case  "cboInsdDthResStCd":
          _rtn = MCSTRCBOINSDDTHRESSTCDLABEL;
          break;

        case  "dtpClmInsdDthDt":
          _rtn = MCSTRDTPCLMINSDDTHDTLABEL;
          break;

        case  "dtpClmProofDt":
          _rtn = MCSTRDTPCLMPROOFDTLABEL;
          break;

        case  "ipmClmInsdSsnNum":
          _rtn = MCSTRIPMCLMINSDSSNNUMLABEL;
          break;

        case  "ipcClmTotDthbPmtAmt":
          _rtn = MCSTRIPCCLMTOTDTHBPMTAMTLABEL;
          break;

        case  "ipcClmTotIntAmt":
          _rtn = MCSTRIPCCLMTOTINTAMTLABEL;
          break;

        case  "ipcClmTotWthldAmt":
          _rtn = MCSTRIPCCLMTOTWTHLDAMTLABEL;
          break;

        case  "ipcClmTotClmPdAmt":
          _rtn = MCSTRIPCCLMTOTCLMPDAMTLABEL;
          break;

        case  "txtClmNum":
          _rtn = MCSTRTXTCLMNUMLABEL;
          break;

        case  "chkClmForResDthInd":
          _rtn = MCSTRCHKCLMFORRESDTHINDLABEL;
          break;

        case  "chkClmCmpCalInd":
          _rtn = MCSTRCHKCLMCMPCALINDLABEL;
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
  private String fnGetLobCd() {
    String _rtn = "";
    //--------------------------------------------------------------------------
    // Procedure:   fnGetLobCd
    // Description: This procedure gets the Line-of-business, given what the
    //              Admin System fpCombo box is currently set to. It defaults to
    //              "I" (for Individual).
    //
    // Params:      N/A
    // Returns:     "G" if a Group-based Admin System is selected; "I" otherwise
    //-----------------------------------------------------------------------------
    "fnGetLobCd"
.equals(Const cstrCurrentProc As String);
    try {

      if ((lpcAdmnSystCd.ColText != "") && (!(modGeneral.gCSTRBLANKENTRY.equals(lpcAdmnSystCd.ColText)))) {
        _rtn = mtWrapper.getLobCdFromAdmnSystCd(lpcAdmnSystCd.ColText);
      } 
      else {
        _rtn = MCSTRINDIVIDUALLOB;
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
  private int fnGetPayeesNeedingRecalcDueToDeath(int lngClmIdIn, Date dteClmInsdDthDtIn) {
    int _rtn = 0;
    //--------------------------------------------------------------------------
    // Procedure:   fnGetPayeesNeedingRecalcDueToDeath
    // Description: Returns the number of Payees for the claim that have a
    //              Date of Payment prior to the Date of Death
    // Params:      N/A
    // Returns:     N/A
    // Date:        04/12/2002
    //-----------------------------------------------------------------------------
    "fnGetPayeesNeedingRecalcDueToDeath"
.equals(Const cstrCurrentProc As String);
    //' Stored procedure to execute
    "dbo.proc_payee_select3"
.equals(Const cstrSproc As String);
    DBRecordSet rstTemp = null;
    ADODB.Parameter prmReturnValue = null;
    ADODB.Parameter prmClmId = null;
    ADODB.Parameter prmClmInsdDthDt = null;
    ADODB.Parameter prmNbrOfPayees = null;
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

      // ---Parameter #2---
      // Define the CLM_ID parameter
      prmClmId = w_aDOCommand.CreateParameter(Name:="@clm_id", Type:=adInteger, Direction:=adParamInput, chrgHourglass.getValue():=lngClmIdIn);
      w_aDOCommand.Parameters.Append(prmClmId);

      // ---Parameter #3---
      // Define the CLM_PROOF_DT parameter
      prmClmInsdDthDt = w_aDOCommand.CreateParameter(Name:="@clm_insd_dth_dt", Type:=adDBTimeStamp, Direction:=adParamInput, Size:=16, chrgHourglass.getValue():=dteClmInsdDthDtIn);
      w_aDOCommand.Parameters.Append(prmClmInsdDthDt);

      // ---Parameter #4---
      // Define the NBR_OF_PAYEES parameter
      prmNbrOfPayees = w_aDOCommand.CreateParameter(Name:="@nbr_of_payees", Type:=adInteger, Direction:=adParamInputOutput, chrgHourglass.getValue():=Null);
      w_aDOCommand.Parameters.Append(prmNbrOfPayees);

      rstTemp = w_aDOCommand.Execute();

      _rtn = prmNbrOfPayees.value;
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    modGeneral.fnFreeRecordset(rstTemp);
    modGeneral.fnFreeObject(adwTemp);
    modGeneral.fnFreeObject(prmClmId);
    modGeneral.fnFreeObject(prmClmInsdDthDt);
    modGeneral.fnFreeObject(prmNbrOfPayees);

    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (prmReturnValue) {
      //' 4027
      case  modResConstants.gCRES_NERR_REC_NOT_FOUND      :
        // Note that this error is presented as a 4027 rather than a 4037!
        // 4037 = The @@1 is invalid. @@2
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_INVALID_DATA, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc, "Claim ID or Date of Proof", "The need for any Payee recalculation cannot be determined "+ "when any of these fields are NULL.");
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
        break;

      case  modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO:
        // 4028 = An error occurred while attempting to @@1 this record.
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
  private int fnGetPayeesNeedingRecalcDueToProof(int lngClmIdIn, Date dteClmProofDtIn) {
    int _rtn = 0;
    //--------------------------------------------------------------------------
    // Procedure:   fnGetPayeesNeedingRecalcDueToProof
    // Description: Returns the number of Payees for the claim that have a
    //              Date of Payment prior to the Date of Proof
    // Params:      N/A
    // Returns:     N/A
    // Date:        04/12/2002
    //-----------------------------------------------------------------------------
    "fnGetPayeesNeedingRecalcDueToProof"
.equals(Const cstrCurrentProc As String);
    //' Stored procedure to execute
    "dbo.proc_payee_select2"
.equals(Const cstrSproc As String);
    DBRecordSet rstTemp = null;
    ADODB.Parameter prmReturnValue = null;
    ADODB.Parameter prmClmId = null;
    ADODB.Parameter prmClmProofDt = null;
    ADODB.Parameter prmNbrOfPayees = null;
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

      // ---Parameter #2---
      // Define the CLM_ID parameter
      prmClmId = w_aDOCommand.CreateParameter(Name:="@clm_id", Type:=adInteger, Direction:=adParamInput, chrgHourglass.getValue():=lngClmIdIn);
      w_aDOCommand.Parameters.Append(prmClmId);

      // ---Parameter #3---
      // Define the CLM_PROOF_DT parameter
      prmClmProofDt = w_aDOCommand.CreateParameter(Name:="@clm_proof_dt", Type:=adDBTimeStamp, Direction:=adParamInput, Size:=16, chrgHourglass.getValue():=dteClmProofDtIn);
      w_aDOCommand.Parameters.Append(prmClmProofDt);

      // ---Parameter #4---
      // Define the NBR_OF_PAYEES parameter
      prmNbrOfPayees = w_aDOCommand.CreateParameter(Name:="@nbr_of_payees", Type:=adInteger, Direction:=adParamInputOutput, chrgHourglass.getValue():=Null);
      w_aDOCommand.Parameters.Append(prmNbrOfPayees);

      rstTemp = w_aDOCommand.Execute();

      _rtn = prmNbrOfPayees.value;
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    modGeneral.fnFreeRecordset(rstTemp);
    modGeneral.fnFreeObject(adwTemp);
    modGeneral.fnFreeObject(prmClmId);
    modGeneral.fnFreeObject(prmClmProofDt);
    modGeneral.fnFreeObject(prmNbrOfPayees);

    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (prmReturnValue) {
      //' 4027
      case  modResConstants.gCRES_NERR_REC_NOT_FOUND      :
        // Note that this error is presented as a 4027 rather than a 4037!
        // 4037 = The @@1 is invalid. @@2
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_INVALID_DATA, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc, "Claim ID or Date of Proof", "The need for any Payee recalculation cannot be determined "+ "when any of these fields are NULL.");
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
        break;

      case  modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO:
        // 4028 = An error occurred while attempting to @@1 this record.
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
    Scripting.FileSystemObject fso = null;

    try {

      fso = new Scripting.FileSystemObject();

      _rtn = fso.BuildPath(App.Path, "Individual_CR8.rpt");

      // Non-fatal error if .RPT doesn't exist or if we couldn't determine the .RPT filename
      if (!(fso.FileExists(fnGetReportFile()))) {
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
  private String fnGetDefaultPayorCompany(String strClmPolNum, String strAdmnSystCd) {
    String _rtn = "";
    // Description: Parses supplied Policy Number and/or uses metadata to determine whether it represents a
    //              a special policy - one whose Parent Company should be
    //              set to a particular value. This includes:
    //                  * AdmnSystCd = 02 (ALIS)  -- all policies
    //                  *              22 (CYBER) -- all policies
    //                  *              37 (VPAS)  -- all policies
    //                  *              SOLAR policies beginning with 'UL'
    // Parameters : N/A
    // Returns    : Default Payor Company value
    // Modified  :
    // --------------------------------------------------
    try {
      "fnGetDefaultPayorCompany"
.equals(Const cstrCurrentProc As String);
      String strChars1Thru2 = "";

      strChars1Thru2 = strClmPolNum.substring(0, 2).toUpperCase();

      _rtn = mstrAdmnSyst_DfltPycoTypDsc;

      // Override admin system default, if policy # is "special" SOLAR range
      if ((strChars1Thru2.equals("UL"))) {
        _rtn = MCSTRPYCO_SUBSIDIARY;
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
  private void fnLoadCboInsdDthResStCd() {
    //--------------------------------------------------------------------------
    // Procedure:   fnLoadcboInsdDthResStCd
    // Description: Populates the Residence State combobox using a sproc
    // Params:      N/A
    // Returns:     N/A
    // Modified:
    //-----------------------------------------------------------------------------
    "fnLoadCboInsdDthResStCd"
.equals(Const cstrCurrentProc As String);
    //' Stored procedure to execute
    "dbo.proc_state_lu_select"
.equals(Const cstrSproc As String);
    DBRecordSet rstTemp = null;
    ADODB.Parameter prmReturnValue = null;
    cadwADOWrapper adwTemp = null;

    try {

      adwTemp = new cadwADOWrapper();

      cboInsdDthResStCd.cerhErrorHandler.clear();

      // Add a blank entry as the first entry of the combobox. This will force the user to select
      // an entry (no default selection) since fnValidData will generate an error if the blank
      // entry is still selected when the user clicks Update.
      cboInsdDthResStCd.AddItem(modGeneral.gCSTRBLANKENTRY);

      if (!(adwTemp.commandSetSproc(cstrSproc))) {
        // **TODO:** goto found: GoTo PROC_EXIT;
      }

      ADODB.Command w_aDOCommand = adwTemp.getADOCommand();
      // ---Parameter #1---
      // Define the return value that represents the error code (i.e. reason) why
      // the stored procedure failed.
      prmReturnValue = w_aDOCommand.CreateParameter(Name:="@return_value", Type:=adInteger, Direction:=adParamReturnValue, chrgHourglass.getValue():=Null);
      w_aDOCommand.Parameters.Append(prmReturnValue);

      rstTemp = w_aDOCommand.Execute();

      //Filter out Compact Filling State
      rstTemp.Filter = "st_cd <> '"+ CSTCOMPACTFILLING+ "'";

      // Add the following, if the combobox contains a hidden ID column associated with the
      // column that *is* displayed:      varItemDataColumn:="co_id",
      modComboBox.fnADORecordSetToComboBox(rstIn:=rstTemp, cboIn:=cboInsdDthResStCd, strDisplayColumn:="st_cd", bClear:=False);
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    rstTemp.Close;
    modGeneral.fnFreeRecordset(rstTemp);
    modGeneral.fnFreeObject(adwTemp);
    modGeneral.fnFreeObject(prmReturnValue);

    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
    }
    return;
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
  private void fnLoadCboIssStCd() {
    //--------------------------------------------------------------------------
    // Procedure:   fnLoadcboIssStCd
    // Description: Populates the Issue State combobox using a sproc
    // Params:      N/A
    // Returns:     N/A
    // Modified:
    //-----------------------------------------------------------------------------
    "fnLoadCboIssStCd"
.equals(Const cstrCurrentProc As String);
    //' Stored procedure to execute
    "dbo.proc_state_lu_select"
.equals(Const cstrSproc As String);
    DBRecordSet rstTemp = null;
    ADODB.Parameter prmReturnValue = null;
    cadwADOWrapper adwTemp = null;

    try {

      adwTemp = new cadwADOWrapper();

      cboIssStCd.cerhErrorHandler.clear();

      // Add a blank entry as the first entry of the combobox. This will force the user to select
      // an entry (no default selection) since fnValidData will generate an error if the blank
      // entry is still selected when the user clicks Update.
      cboIssStCd.AddItem(modGeneral.gCSTRBLANKENTRY);

      if (!(adwTemp.commandSetSproc(cstrSproc))) {
        // **TODO:** goto found: GoTo PROC_EXIT;
      }

      ADODB.Command w_aDOCommand = adwTemp.getADOCommand();
      // ---Parameter #1---
      // Define the return value that represents the error code (i.e. reason) why
      // the stored procedure failed.
      prmReturnValue = w_aDOCommand.CreateParameter(Name:="@return_value", Type:=adInteger, Direction:=adParamReturnValue, chrgHourglass.getValue():=Null);
      w_aDOCommand.Parameters.Append(prmReturnValue);

      rstTemp = w_aDOCommand.Execute();

      //Filter out Compact Filling State
      rstTemp.Filter = "st_cd <> '"+ CSTCOMPACTFILLING+ "'";

      // Add the following, if the combobox contains a hidden ID column associated with the
      // column that *is* displayed:      varItemDataColumn:="co_id",
      modComboBox.fnADORecordSetToComboBox(rstIn:=rstTemp, cboIn:=cboIssStCd, strDisplayColumn:="st_cd", bClear:=False);
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    rstTemp.Close;
    modGeneral.fnFreeRecordset(rstTemp);
    modGeneral.fnFreeObject(adwTemp);
    modGeneral.fnFreeObject(prmReturnValue);

    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
    }
    return;
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
  private void fnLoadControls() {
    // Description: Will take applicable value from the recordset and put them into
    //              the screen controls
    // Params:      N/A
    // Returns:     N/A
    // Modified  :
    // --------------------------------------------------
    // Modified:     Berry Kropiwka - Added Compact Calc - 2019-09-27
    "fnLoadControls"
.equals(Const cstrCurrentProc As String);
    String strSavedMask = "";

    try {

      // The AdmnSystCd is displayed via a fpCombo control since it must display a
      // Description but store a Code. Search the control's contents for the Code,
      // so that the row with that Code's corresponding Description will be selected
      modComboBox.fnSearchFPCombo(lpcAdmnSystCd, mtWrapper.getAdmnSystCd(), MCINTSTORECOL_LPCADMNSYSTCD);

      // The following will trigger iptClmPolNum_Change( ) which sets the line-of-business (LOB)
      // and then repopulates the Admin System combo box based on the LOB
      iptClmPolNum.Text = mtWrapper.getClmPolNum().trim();

      if (mtWrapper.getClmForResDthInd()) {
        chkClmForResDthInd.chrgHourglass.setValue(vbChecked);
      } 
      else {
        chkClmForResDthInd.chrgHourglass.setValue(vbUnchecked);
      }

      if (mtWrapper.getClmCompactClcnInd()) {
        chkClmCmpCalInd.chrgHourglass.setValue(vbChecked);
      } 
      else {
        chkClmCmpCalInd.chrgHourglass.setValue(vbUnchecked);
      }

      // Set the availability of the InsdDthResStCd based on the Foreign Residence at Death checkbox selection.
      fnSetInsdDthResStCdAvailability();

      // The PycoTypCd is displayed via a fpCombo control since it must display a
      // Description but store a Code. Search the control's contents for the Code,
      // so that the row with that Code's corresponding Description will be selected
      modComboBox.fnSearchFPCombo(lpcPycoTypCd, mtWrapper.getPycoTypCd(), MCINTSTORECOL_LPCPYCOTYPCD);

      iptClmInsdFirstNm = mtWrapper.getClmInsdFirstNm();
      iptClmInsdLastNm = mtWrapper.getClmInsdLastNm();

      // cboIssStCd corresponds to a Nullable field, so accommodate Nulls
      if (mtWrapper.getIssStCd().equals("")) {
        cboIssStCd.Text = modGeneral.gCSTRBLANKENTRY;
      } 
      else {
        cboIssStCd.Text = mtWrapper.getIssStCd();
      }

      // cboInsdDthResStCd corresponds to a Nullable field, so accommodate Nulls
      if (mtWrapper.getInsdDthResStCd().equals("")) {
        cboInsdDthResStCd.Text = modGeneral.gCSTRBLANKENTRY;
      } 
      else {
        cboInsdDthResStCd.Text = mtWrapper.getInsdDthResStCd();
      }

      dtpClmInsdDthDt.chrgHourglass.setValue(mtWrapper.getClmInsdDthDt());
      dtpClmProofDt.chrgHourglass.setValue(mtWrapper.getClmProofDt());

      //!TODO! Should these Original dates be of type Date?
      // Save the original value of these 2 dates fields. If they change and Payees
      // exist at that time, a warning should be issued to indicate the change may
      // necessitate a recalculation of the Payee's values.
      mstrOrigDateOfDeath = mtWrapper.getClmInsdDthDt();
      mstrOrigDateOfProof = mtWrapper.getClmProofDt();

      // NOTE: For MaskEdBox controls, have to do special processing based on whether or not the
      //       field is empty, to avoid a 380 "invalid property value" runtime error.
      //       * If it's empty, temporarily delete the mask, set the value, and then restore
      //         the mask.
      //       * If it's not empty, format the value so it will be "valid" per the .Mask
      //         (for phone numbers, this means inserting a dash between characters 3 and 4).
      if (LenB(mtWrapper.getClmInsdSsnNum()) == 0) {
        strSavedMask = ipmClmInsdSsnNum.ctclmClaim.getMask();
        ipmClmInsdSsnNum.ctclmClaim.setMask("");
        ipmClmInsdSsnNum.Text = "";
        ipmClmInsdSsnNum.ctclmClaim.setMask(strSavedMask);
      } 
      else {
        ipmClmInsdSsnNum.Text = modGeneral.fnSSNTIN_AddDash(strIn:=.ctclmClaim.getClmInsdSsnNum(), bIsTin:=False);
      }

      ipcClmTotDthbPmtAmt.Text = mtWrapper.getClmTotDthbPmtAmt();
      ipcClmTotIntAmt.Text = mtWrapper.getClmTotIntAmt();
      ipcClmTotWthldAmt.Text = mtWrapper.getClmTotWthldAmt();
      ipcClmTotClmPdAmt.Text = mtWrapper.getClmTotClmPdAmt();
      txtClmNum.Text = mtWrapper.getClmNum();

      // ClmId         isn't shown on-screen
      // LstUpdDtm     isn't shown on-screen
      // LstUpdUserId  isn't shown on-screen

      // Get the Payees associated with the claim and populate the Payees grid
      fnGetChildren();

      // Make sure Navigation buttons are enabled/disabled based on current record position in the Lookup recordset
      fnSetNavigationButtons(bUnconditionalDisable:=False);

      // Update the "record x of y" label
      lblRecordPosition = modGeneral.fnShowRecordPosition(mtWrapper.getLookupData());

      // Enable or Disable the Compact Filling check box based on Admin System
      fnSetCompactFillingCheckBox(Me.lpcAdmnSystCd.Text, Me.cboIssStCd.Text);

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
  private void fnLoadLpcAdmnSystCd() {
    //--------------------------------------------------------------------------
    // Procedure:   fnLoadLpcAdmnSystCd
    // Description: Populates the Admin System fpCombo control using a sproc
    // Params:      N/A
    // Returns:     N/A
    // Modified:
    //-----------------------------------------------------------------------------
    "fnLoadLpcAdmnSystCd"
.equals(Const cstrCurrentProc As String);
    //' Stored procedure to execute
    "dbo.proc_admin_system_lu_select"
.equals(Const cstrSproc As String);
    DBRecordSet rstTemp = null;
    ADODB.Parameter prmReturnValue = null;
    cadwADOWrapper adwTemp = null;

    try {

      adwTemp = new cadwADOWrapper();

      //*TODO:** can't found type for with block
      //*With lpcAdmnSystCd
      __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = lpcAdmnSystCd;
      w___TYPE_NOT_FOUND.Clear;
      // Add a blank entry as the first entry of the combobox. This will force the user to select
      // an entry (no default selection) since fnValidData will generate an error if the blank
      // entry is still selected when the user clicks Update.

      // Set .Row to -1 so insertion works okay whether or not the fpCombo is sorted
      w___TYPE_NOT_FOUND.Row = modComboBox.gCLNGNOSELECTION;
      w___TYPE_NOT_FOUND.InsertRow = modGeneral.gCSTRBLANKENTRY+ vbTab+ modGeneral.gCSTRBLANKENTRY;

      if (!(adwTemp.commandSetSproc(cstrSproc))) {
        // **TODO:** goto found: GoTo PROC_EXIT;
      }

      ADODB.Command w_aDOCommand = adwTemp.getADOCommand();
      // ---Parameter #1---
      // Define the return value that represents the error code (i.e. reason) why
      // the stored procedure failed.
      prmReturnValue = w_aDOCommand.CreateParameter(Name:="@return_value", Type:=adInteger, Direction:=adParamReturnValue, chrgHourglass.getValue():=Null);
      w_aDOCommand.Parameters.Append(prmReturnValue);

      rstTemp = w_aDOCommand.Execute();

      if (rstTemp.RecordCount != 0) {
        rstTemp.MoveFirst;
        do Until .EOF          // Set .Row to -1 so insertion works okay whether or not the fpCombo is sorted
          lpcAdmnSystCd.Row = modComboBox.gCLNGNOSELECTION;
          lpcAdmnSystCd.InsertRow = rstTemp.Fields(MCSTRADMNSYSTDSC)((Boolean) rstTemp.value).toString()+ vbTab+ rstTemp.Fields(MCSTRADMNSYSTCD).chrgHourglass.getValue();
          rstTemp.MoveNext;
        }
      } 
      else {
        modGeneral.gerhApp.reportNonFatal(vbObjectError + modResConstants.gCRES_WARN_CBO_IS_EMPTY, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc, MCSTRLPCADMNSYSTCDLABEL);
      }
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    modGeneral.fnFreeRecordset(rstTemp);
    modGeneral.fnFreeObject(adwTemp);
    modGeneral.fnFreeObject(prmReturnValue);

    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
    }
    return;
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
  private void fnLoadLpcLookup(LPLib.fpCombo lpcIn, EnumLookupType lngLookupType) { // TODO: Use of ByRef founded Private Sub fnLoadLpcLookup(ByRef lpcIn As LPLib.fpCombo, ByVal lngLookupType As EnumLookupType)
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


      lpcIn.Clear;
      lpcIn.Row = modComboBox.gCLNGNOSELECTION;
      lpcIn.SortState = SortStateSuspend;

      switch (lngLookupType) {
        case  EnumLookupType.eLT_CLAIM:
          aRows = mtWrapper.getLookupData_Claim();
          lpcIn.InsertRow = modGeneral.gCSTRBLANKENTRY+ vbTab+ modGeneral.gCSTRBLANKENTRY;
          for (lngRow = 0; lngRow <= (aRows, cintRowDimension).length; lngRow++) {
            // There are 2 columns in the array and fpCombo control (indexed 0 thru 1).
            lpcIn.InsertRow = aRows[0, lngRow]+ vbTab+ aRows[1, lngRow];
          }
          // Reset property to ensure whole width displays
          lpcIn.DataAutoSizeCols = DataAutoSizeColsMaxColWidth;
          break;

        case  EnumLookupType.eLT_NAME:
          aRows = mtWrapper.getLookupData_Name();
          lpcIn.InsertRow = modGeneral.gCSTRBLANKENTRY+ vbTab+ modGeneral.gCSTRBLANKENTRY+ vbTab+ modGeneral.gCSTRBLANKENTRY+ vbTab+ modGeneral.gCSTRBLANKENTRY;
          for (lngRow = 0; lngRow <= (aRows, cintRowDimension).length; lngRow++) {
            // There are 4 columns in the array and fpCombo control (indexed 0 thru 3)
            lpcIn.InsertRow = aRows[0, lngRow]+ vbTab+ aRows[1, lngRow]+ vbTab+ aRows[2, lngRow]+ vbTab+ aRows[3, lngRow];
          }
          // Reset property to ensure whole width displays
          lpcIn.DataAutoSizeCols = DataAutoSizeColsMaxColWidth;
          break;

        case  EnumLookupType.eLT_SSN:
          aRows = mtWrapper.getLookupData_SSN();
          lpcIn.InsertRow = modGeneral.gCSTRBLANKENTRY+ vbTab+ modGeneral.gCSTRBLANKENTRY+ vbTab+ modGeneral.gCSTRBLANKENTRY;
          for (lngRow = 0; lngRow <= (aRows, cintRowDimension).length; lngRow++) {
            // There are 3 columns in the array and fpCombo control (indexed 0 thru 2)
            lpcIn.InsertRow = aRows[0, lngRow]+ vbTab+ aRows[1, lngRow]+ vbTab+ aRows[2, lngRow];
          }
          // Reset property to ensure whole width displays
          lpcIn.DataAutoSizeCols = DataAutoSizeColsMaxColWidth;
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
    Erase(aRows);

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
  private void fnLoadLpcPycoTypCd() {
    // Comments  : Populates Company Type fpCombo control using a sproc
    // Parameters: N/A
    // Returns   : N/A
    // Modified  :
    // --------------------------------------------------
    "fnLoadLpcPycoTypCd"
.equals(Const cstrCurrentProc As String);
    //' Stored procedure to execute
    "dbo.proc_payor_company_type_lu_select"
.equals(Const cstrSproc As String);
    DBRecordSet rstTemp = null;
    ADODB.Parameter prmReturnValue = null;
    cadwADOWrapper adwTemp = null;

    try {

      adwTemp = new cadwADOWrapper();

      //*TODO:** can't found type for with block
      //*With lpcPycoTypCd
      __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = lpcPycoTypCd;
      w___TYPE_NOT_FOUND.Clear;
      // Add a blank entry as the first entry of the combobox. This will be the default entry
      // until the user specifies a Policy Number, since the true default is based on
      // whether the first digits of the Policy Number begin with UL, UV, UZ or 222.
      // Note too that fnValidData will generate an error if the blank
      // entry is still selected when the user clicks Update.

      // Set .Row to -1 so insertion works okay whether or not the fpCombo is sorted
      w___TYPE_NOT_FOUND.Row = modComboBox.gCLNGNOSELECTION;
      w___TYPE_NOT_FOUND.InsertRow = modGeneral.gCSTRBLANKENTRY+ vbTab+ modGeneral.gCSTRBLANKENTRY;

      if (!(adwTemp.commandSetSproc(cstrSproc))) {
        // **TODO:** goto found: GoTo PROC_EXIT;
      }

      ADODB.Command w_aDOCommand = adwTemp.getADOCommand();
      // ---Parameter #1---
      // Define the return value that represents the error code (i.e. reason) why
      // the stored procedure failed.
      prmReturnValue = w_aDOCommand.CreateParameter(Name:="@return_value", Type:=adInteger, Direction:=adParamReturnValue, chrgHourglass.getValue():=Null);
      w_aDOCommand.Parameters.Append(prmReturnValue);

      rstTemp = w_aDOCommand.Execute();

      if (rstTemp.RecordCount != 0) {
        rstTemp.MoveFirst;
        do Until .EOF          // Set .Row to -1 so insertion works okay whether or not the fpCombo is sorted
          lpcPycoTypCd.Row = modComboBox.gCLNGNOSELECTION;
          lpcPycoTypCd.InsertRow = rstTemp.Fields(MCSTRPYCOTYPDSC)((Boolean) rstTemp.value).toString()+ vbTab+ rstTemp.Fields(MCSTRPYCOTYPCD).chrgHourglass.getValue();
          rstTemp.MoveNext;
        }
      } 
      else {
        modGeneral.gerhApp.reportNonFatal(vbObjectError + modResConstants.gCRES_WARN_CBO_IS_EMPTY, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc, MCSTRLPCPYCOTYPCDLABEL);

      }
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    modGeneral.fnFreeRecordset(rstTemp);
    modGeneral.fnFreeObject(adwTemp);
    modGeneral.fnFreeObject(prmReturnValue);

    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
    }
    return;
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
  private void fnLoadRecordWithCalculatedControls() {
    // Comments  : Populates DB record with data from screen controls
    //             that are calculated
    // Parameters: None
    // Modified  :
    // --------------------------------------------------
    try {
      "fnLoadRecordWithCalculatedControls"
.equals(Const cstrCurrentProc As String);

      if (!mbInAddMode) {
        // Extra precaution...to always calc totals across Payees
        // before doing a save. This will also ensure the totals
        // are 0 for an Add.  Can't call this prodedure on an Add
        // since there is no current record and it will get a
        // ADO 3021 error: "Either BOF or EOF is true or the current
        // record has been deleted. Requested operation requires a
        // current record."
        fnCalcTotalsForAllPayees(mtWrapper.getClmId());
      }

      // The following fields cannot be edited by the user but are calculated
      // by the program
      mtWrapper.setClmTotDthbPmtAmt(ipcClmTotDthbPmtAmt.UnFmtText);
      mtWrapper.setClmTotIntAmt(ipcClmTotIntAmt.UnFmtText);
      mtWrapper.setClmTotWthldAmt(ipcClmTotWthldAmt.UnFmtText);
      mtWrapper.setClmTotClmPdAmt(ipcClmTotClmPdAmt.UnFmtText);
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
  private void fnSetCompactFillingCheckBox(String strAdminSystem, String strState) {
    // Comments  : Retrieves the state record for the current claim record
    // Parameters: N/A
    // Called by : lpcAdmnSystCd_Change
    //             fnLoadControls
    // Modified  : Berry Kropiwka 11-06-2019
    //
    // --------------------------------------------------
    String strSQL = "";
    DBRecordSet rstTemp = null;
    String cstrCurrentProc = "";
    cstrCurrentProc = "fnSetCompactFillingCheckBox";
    try {
      //'Or strAdminSystem = "LEVERAGE" Then
      if (strAdminSystem.equals("SOLAR")) {
        //Now check to make sure the claim is in a state that allows Compact Filling
        strSQL = "SELECT st_compact_clcn_allow_ind from dbo.state_t WHERE st_cd = '"+ strState+ "'";
        rstTemp = modGeneral.gadwApp.execute_SQL_AsRST(modGeneral.gconAppActive, strSQL);
        if (rstTemp.RecordCount > 0) {
          if ("T"
.equals(rstTemp!st_compact_clcn_allow_ind)) {
            Me.chkClmCmpCalInd.Enabled = true;
          } 
          else {
            Me.chkClmCmpCalInd.Enabled = false;
            Me.chkClmCmpCalInd.chrgHourglass.setValue(vbUnchecked);
          }
        } 
        else {
          Me.chkClmCmpCalInd.Enabled = false;
          Me.chkClmCmpCalInd.chrgHourglass.setValue(vbUnchecked);
        }
        rstTemp.ActiveConnection = null;
      } 
      else {
        if (strAdminSystem.isEmpty() && strState.isEmpty()) {
          Me.chkClmCmpCalInd.Enabled = true;
        } 
        else {
          Me.chkClmCmpCalInd.Enabled = false;
          Me.chkClmCmpCalInd.chrgHourglass.setValue(vbUnchecked);
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
      lpcIn.ColFromName = MCSTRCLMID;

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
  private void fnRefreshAllCombos() {
    //--------------------------------------------------------------------------
    // Procedure:   fnRefreshAllCombos
    // Description: Repopulates each ComboBox or VSFlexGrid control
    //              so they reflect this and other users' changes. This proc
    //              should be called after each Add, Update or Delete.
    //
    // Params:      N/A
    // Called by:   cmdUpdate_Click of frmFund
    //              cmdDelete_Click of frmFund
    //              Form_Load of frmFund
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

      //' #1 = Claim Number (CLM_NUM, CLM_ID)
      fnLoadLpcLookup(lpcLookupClaim, EnumLookupType.eLT_CLAIM);
      //' #2 = Insured Name (CLM_INSD_LAST_NM, CLM_INSD_FIRST_NM, CLM_NUM, CLM_ID)
      fnLoadLpcLookup(lpcLookupName, EnumLookupType.eLT_NAME);
      //' #3 = Insured SSN (CLM_INSD_SSN_NUM, CLM_ID)
      fnLoadLpcLookup(lpcLookupSSN, EnumLookupType.eLT_SSN);
      //' #4 = Admin System (ADMN_SYST_DSC, ADMN_SYST_CD)
      fnLoadLpcAdmnSystCd();
      //' #5 = Payor Company Type (PYCO_TYP_DSC, PYCO_TYP_CD)
      fnLoadLpcPycoTypCd();
      //' #6 = Issue State (ISS_ST_CD)
      fnLoadCboIssStCd();
      //' #7 = Insured State of Residence at time of Death (INSD_DTH_RES_ST_CD)
      fnLoadCboInsdDthResStCd();
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
    //             calling this function, e.g., they accurately reflect whether
    //             or not there are edits outstanding and/or the user is in
    //             Add mode, respectively.
    //             Remember, though: mbInAddMode and IsDirty are
    //             independent of one another!
    //
    //     State          ADD btn  UPD btn  DEL btn  CLOSE btn PAYEE btn PRTRPT btn
    //    --------------  -------- -------- -------- --------- --------- ----------
    //    Add mode       disabled  enabled  disabled enabled   disabled  disabled
    //    (no edits yet)
    //
    //    Edits o/s      disabled  enabled  disabled enabled   disabled  disabled
    //
    //    No edits o/s   enabled   disabled enabled  enabled   enabled   enabled
    //    & #Children = 0
    //
    //    No edits o/s   enabled   disabled disabled enabled   enabled   enabled
    //    & #Children > 0
    //
    // Called by : fnAddRecord and fnInitializeEditMode, with bEnable = False
    //
    //             lpcLookupClaim_Click, lpcLookupName_Click, lpcLookupSSN_Click,
    //             cmdDelete_Click, cmdNavigate_Click, cmdUpdate_Click
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

      // Can only go to the Payees or Print an Individual Report when you're NOT in the middle of
      // an Add or Update. It doesn't matter whether you have Payees though!
      if ((Not setIsDirty()) && (Not mbInAddMode)) {
        cmdAddPayee.Enabled = true;
        cmdPrintReport.Enabled = true;
      } 
      else {
        cmdAddPayee.Enabled = false;
        cmdPrintReport.Enabled = false;
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

      // Disallow future-dated Date of Death or Date of Proof
      dtpClmInsdDthDt.MaxDate = Now;
      dtpClmProofDt.MaxDate = Now;

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
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }
}


//////////////////////////////////////////////////////////////////////////////////////////////////
  private void fnSetInsdDthResStCdAvailability() {
    //--------------------------------------------------------------------------
    // Procedure:   fnSetInsdDthResStCdAvailability
    // Description: Sets the availability of the InsdDthResStCd based on
    //              whether the Foreign Residence at Death checkbox is selected.
    //
    // Params:      n/a
    // Called by:   Form_Load of frmInsured
    //              fnLoadControls of frmInsured
    // Returns:     n/a
    //-----------------------------------------------------------------------------
    "fnSetInsdDthResStCdAvailability"
.equals(Const cstrCurrentProc As String);
    try {

      if (chkClmForResDthInd.chrgHourglass.getValue() == vbChecked) {
        cboInsdDthResStCd.Text = modGeneral.gCSTRBLANKENTRY;
        lblInsdDthResStCd.Enabled = false;
        cboInsdDthResStCd.Enabled = false;
      } 
      else {
        lblInsdDthResStCd.Enabled = true;
        cboInsdDthResStCd.Enabled = true;
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
    //              lpcLookupClaim_Click( )
    //              lpcLookupName_Click( )
    //              lpcLookupSSN_Click( )
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
  private void fnSetPropertiesForPayeeScreen(boolean bSendEmptyName) {
    //----------------------------------------------------------------------------
    // Procedure :  Sub fnSetPropertiesForPayeeScreen
    // Created by:  BAW on 04-26-2001 08:55
    //
    // Comments  : Sets member variables so they can be accessed from/by Payee screen
    // Called by : msgPayees_DblClick and cmdAddPayee_Click
    // Parameters: N/A
    //
    // Modified  :
    //----------------------------------------------------------------------------
    try {
      "fnSetPropertiesForPayeeScreen"
.equals(Const cstrCurrentProc As String);

      //*TODO:** can't found type for with block
      //*With msgPayees
      __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = msgPayees;
      // Note: If there are no Payees, msgPayees.Row will be set to 0 (the header row)
      //' Payee Name column (2nd column, current row)
      w___TYPE_NOT_FOUND.Col = 1;

      if (bSendEmptyName) {
        setInsuredCurrentPayeeName("");
        setInsuredCurrentPayeeID(0);
      } 
      else {
        setInsuredCurrentPayeeName(w___TYPE_NOT_FOUND.Text);
        // Get Payee ID from same row, different column
        w___TYPE_NOT_FOUND.Col = 13;
        setInsuredCurrentPayeeID(w___TYPE_NOT_FOUND.Text);
      }

      setInsuredClmID(mtWrapper.getClmId());
      setInsuredClmForResDthInd(mtWrapper.getClmForResDthInd());
      setInsuredClmInsdDthDt(mtWrapper.getClmInsdDthDt());
      setInsuredClmNum(mtWrapper.getClmNum());
      setInsuredClmProofDt(mtWrapper.getClmProofDt());
      setInsuredInsdDthResStCd(mtWrapper.getInsdDthResStCd());
      setInsuredIssStCd(mtWrapper.getIssStCd());
      setInsuredLobCd(fnGetLobCd());
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
  private void fnSetTxtClmNum() {
    //--------------------------------------------------------------------------
    // Procedure:   fnSetTxtClmNum
    // Description: This procedure sets the hidden txtClmNum control, based
    //              on AdmnSystCd, ClmNum and, for Group, ClmInsdSSNNum
    //
    // Params:      N/A
    // Returns:     N/A
    //-----------------------------------------------------------------------------
    "fnSetTxtClmNum"
.equals(Const cstrCurrentProc As String);
    try {

      //*TODO:** can't found type for with block
      //*With txtClmNum
      __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = txtClmNum;
      if (fnGetLobCd() == MCSTRGROUPLOB) {
        w___TYPE_NOT_FOUND.Text = iptClmPolNum.Text+ MCSTRGROUPLOB+ ipmClmInsdSsnNum.UnFmtText;
      } 
      else {
        w___TYPE_NOT_FOUND.Text = iptClmPolNum.Text;
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
      modGeneral.fnEnableDisableControl(ctlIn:=ipcClmTotDthbPmtAmt, bEnable:=False);
      modGeneral.fnEnableDisableControl(ctlIn:=ipcClmTotIntAmt, bEnable:=False);
      modGeneral.fnEnableDisableControl(ctlIn:=ipcClmTotWthldAmt, bEnable:=False);
      modGeneral.fnEnableDisableControl(ctlIn:=ipcClmTotClmPdAmt, bEnable:=False);
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
      Const(cintClmInsdSsnNumMinLgth As Integer == 9);
      boolean bErrorFound = false;
      Control ctl = null;
      Control ctlFirstToFail = null;
      int intFailures = 0;
      String strFieldList = "";
      String strMsgText = "";
      int intLengthToTest = 0;
      String strLOB = "";
      String strDefaultPycoTypDsc = "";

      _rtn = true;

      // Check the fields in a left-to-right, top-to-bottom screen sequence.
      //     1. cboAdmnSystCd         7. chkClmForResDthInd
      //     2. iptClmPolNum          8. cboInsdDthResStCd
      //     3. lpcPycoTypCd          9. dtpInsdDthDt
      //     4. iptClmInsdFirstNm    10. dtpClmProofDt
      //     5. iptClmInsdLastNm     11. ipmClmInsdSsnNum
      //     6. cboIssStCd

      // ------------- First, verify required fields are missing --------------

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
        // Skip over the control that is bound to CLM_NUM since this is a hidden
        // field and thus the user shouldn't be informed if it hasn't been set yet.
        if (ctl.Tag.length() > 0 && (!("ClmNum".equals(ctl.Tag)))) {
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
              if ((ctl.length() == 0) || (ctl == modGeneral.gCSTRBLANKENTRY)) {
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

      // Check the Issue State, which is a nullable column. This may not have been input
      // if the claim was entered prior to 2002 or so. However, if the user modifies the
      // claim once the backend has been ported to SQL Server, or on new claims, then they **will** have to supply it.
      if ((LenB(cboIssStCd.Text) == 0) || (modGeneral.gCSTRBLANKENTRY.equals(cboIssStCd.Text))) {
        if (intFailures == 0) {
          strFieldList = "\\r\\n"+ MCSTRCBOISSSTCDLABEL;
          ctlFirstToFail = cboIssStCd;
        } 
        else {
          strFieldList = strFieldList+ "\\r\\n"+ MCSTRCBOISSSTCDLABEL;
        }
        intFailures = intFailures + 1;
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


      // ------------------- Now, do cross-field validations --------------------



      //' Reset for this section of error validations
      intFailures = 0;

      intLengthToTest = iptClmPolNum.Text.length();
      // Min/Max lengths were set upon each change to the Admin System control.
      if ((intLengthToTest < mintAdmnSyst_MinPolNumLength) || (intLengthToTest > mintAdmnSyst_MaxPolNumLength)) {
        intFailures = intFailures + 1;
        ctlFirstToFail = iptClmPolNum;
        strMsgText = strMsgText+ "\\r\\n"+ "For the selected Admin System, the "+ MCSTRIPTCLMPOLNUMLABEL+ " must be between "+ mintAdmnSyst_MinPolNumLength+ " and "+ mintAdmnSyst_MaxPolNumLength+ " characters long, inclusive.";
      }

      // If the CLM_NUM (the logical key to this table) has changed, verify it has not been changed to
      // one that already exists.
      //If txtClmNum.Text <> mtWrapper.ClmNum Then


      // Verify the Payor Company Type is set to an appropriate default value
      strDefaultPycoTypDsc = fnGetDefaultPayorCompany(iptClmPolNum.Text, lpcAdmnSystCd.ColText);

      if (modGeneral.gapsApp.getLastLogonEnvironment().indexOf("Sun") == 1) {
        if (!(strDefaultPycoTypDsc.equals(lpcPycoTypCd.Text))) {
          if (!(MCSTRADMNSYSTSOLAR.equals(lpcAdmnSystCd.ColText) && MCSTRPYCO_SLHIC.equals(lpcPycoTypCd.Text) && !(strDefaultPycoTypDsc.equals(MCSTRPYCO_SUBSIDIARY)))) {
            //report the error
            intFailures = intFailures + 1;
            ctlFirstToFail = lpcPycoTypCd;
            strMsgText = strMsgText+ "\\r\\n"+ "The selected "+ MCSTRLPCPYCOTYPCDLABEL+ " is invalid for this "+ MCSTRLPCADMNSYSTCDLABEL+ " or "+ MCSTRIPTCLMPOLNUMLABEL+ ".";
          }
        }
      } 
      else {
        if (!(strDefaultPycoTypDsc.equals(lpcPycoTypCd.Text))) {
          intFailures = intFailures + 1;
          ctlFirstToFail = lpcPycoTypCd;
          strMsgText = strMsgText+ "\\r\\n"+ "The selected "+ MCSTRLPCPYCOTYPCDLABEL+ " is invalid for this "+ MCSTRLPCADMNSYSTCDLABEL+ " or "+ MCSTRIPTCLMPOLNUMLABEL+ ".";
        }
      }

      // Make sure the user selected a non-blank entry in the Residence State & Issue State controls.
      if (chkClmForResDthInd.chrgHourglass.getValue() == vbUnchecked) {
        if ((LenB(cboInsdDthResStCd.Text) == 0) || (modGeneral.gCSTRBLANKENTRY.equals(cboInsdDthResStCd.Text))) {
          intFailures = intFailures + 1;
          ctlFirstToFail = cboInsdDthResStCd;
          strMsgText = strMsgText+ "\\r\\n"+ "Unless "+ MCSTRCHKCLMFORRESDTHINDLABEL+ " is selected, the "+ MCSTRCBOINSDDTHRESSTCDLABEL+ " must be supplied.";
        }
      }

      // Verify the Date of Proof is on or after the Date of Death
      if (DateValue(dtpClmProofDt.chrgHourglass.getValue()) < DateValue(dtpClmInsdDthDt.chrgHourglass.getValue())) {
        intFailures = intFailures + 1;
        ctlFirstToFail = dtpClmProofDt;
        strMsgText = strMsgText+ "\\r\\n"+ "The "+ MCSTRDTPCLMPROOFDTLABEL+ " ("+ ((Boolean) dtpClmProofDt.chrgHourglass.getValue()).toString()+ ") must be on or after the "+ MCSTRDTPCLMINSDDTHDTLABEL+ " ("+ ((Boolean) dtpClmInsdDthDt.chrgHourglass.getValue()).toString()+ ").";
      }

      // Determine whether any Payees exist with a Date Of Payment earlier than the
      // Insured's Date of PROOF.  Skip if in Add mode, since there would be no Payees and
      // the ClmId would be invalid.
      if (!mbInAddMode) {
        if (fnGetPayeesNeedingRecalcDueToProof(mtWrapper.getClmId(), mtWrapper.getClmProofDt()) > 0) {
          intFailures = intFailures + 1;
          ctlFirstToFail = dtpClmProofDt;
          strMsgText = strMsgText+ "\\r\\n"+ "One or more Payees exist with a Date Of Payment "+ "earlier than the "+ MCSTRDTPCLMPROOFDTLABEL+ ".";
        }

        // Determine whether any Payees exist with a Date Of Payment earlier than the
        // Insured's Date of DEATH.
        if (fnGetPayeesNeedingRecalcDueToDeath(mtWrapper.getClmId(), mtWrapper.getClmInsdDthDt()) > 0) {
          intFailures = intFailures + 1;
          ctlFirstToFail = dtpClmInsdDthDt;
          strMsgText = strMsgText+ "\\r\\n"+ "One or more Payees exist with a Date Of Payment "+ "earlier than the "+ MCSTRDTPCLMINSDDTHDTLABEL+ ".";
        }
      }

      // Verify that a 9-character SSN was input, if anything was input to that field
      intLengthToTest = ipmClmInsdSsnNum.UnFmtText.length();
      if (intLengthToTest != 0  && intLengthToTest != cintClmInsdSsnNumMinLgth) {
        intFailures = intFailures + 1;
        ctlFirstToFail = iptClmPolNum;
        strMsgText = strMsgText+ "\\r\\n"+ "If input, the "+ MCSTRIPMCLMINSDSSNNUMLABEL+ " must be "+ CStr(cintClmInsdSsnNumMinLgth)+ " characters long.";
      }

      if (intFailures != 0) {
        bErrorFound = true;
        _rtn = false;
        if (ctlFirstToFail.Visible) {
          ctlFirstToFail.SetFocus;
        }
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_CROSS_FLD_VALIDATIONS_FAILED, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc, "this record can be updated", strMsgText);
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

      if ((Not mbInAddMode)) {
        if ((mrstPayees.RecordCount > 0)) {
          if (DateValue(mstrOrigDateOfDeath) != DateValue(dtpClmInsdDthDt.chrgHourglass.getValue())) {
            // 1008 = The @@1 has changed. This change may affect the calculations for existing Payees. Please review and, if necessary, recalculate each Payee.
            modGeneral.gerhApp.reportNonFatal(vbObjectError + modResConstants.gCRES_INFO_DT_CHG_MAY_AFFECT_PAYEES, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc, MCSTRDTPCLMINSDDTHDTLABEL);
          }
          if (DateValue(mstrOrigDateOfProof) != DateValue(dtpClmProofDt.chrgHourglass.getValue())) {
            // 1008 = The @@1 has changed. This change may affect the calculations for existing Payees. Please review and, if necessary, recalculate each Payee.
            modGeneral.gerhApp.reportNonFatal(vbObjectError + modResConstants.gCRES_INFO_DT_CHG_MAY_AFFECT_PAYEES, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc, MCSTRDTPCLMPROOFDTLABEL);
          }
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
  private void form_Activate() {
    // Comments  :
    // Parameters:  -
    // Modified  :
    // --------------------------------------------------
    try {
      "Form_Activate"
.equals(Const cstrCurrentProc As String);

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      // Since this form is hidden which the Payee form is visible, clicking on the Payee
      // form can trigger the frmInsured's Form_Activate event. Therefore, the bulk
      // of the processing in this event is conditioned on whether it (frmInsured)
      // is visible or not. If not visible, we don't want to mess up the Payee-related
      // values that could mess up the processing in the Payee form.
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
      //' Invalid procedure call or argument
      case  5 :
        // Caused by setting the focus to a field that's not yet visible
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
    // Parameters:  -
    // Modified  :
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
      // 1. Claim Lookup
      //*TODO:** can't found type for with block
      //*With lpcLookupClaim
      __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = lpcLookupClaim;
      modComboBox.fnInitializefpCombo(lpcIn:=lpcLookupClaim, bShowColHeaders:=False, bSortable:=False, lngNbrOfCols:=2, lngEditCol:=mcintDisplayCol_lpcLookupClaim, lngNbrOfRowsInDropdown:=8);
      // Column definitions
      //' First column, Primary sort
      w___TYPE_NOT_FOUND.Col = 0;
      w___TYPE_NOT_FOUND.ColHeaderText = MCSTRTXTCLMNUMLABEL;
      w___TYPE_NOT_FOUND.ColName = MCSTRDISPLAYCOL;
      w___TYPE_NOT_FOUND.ColWidth = 20;
      //' Second column
      w___TYPE_NOT_FOUND.Col = 1;
      w___TYPE_NOT_FOUND.ColHeaderText = MCSTRTXTCLMIDLABEL;
      w___TYPE_NOT_FOUND.ColName = MCSTRCLMID;
      w___TYPE_NOT_FOUND.ColHide = true;
      w___TYPE_NOT_FOUND.ColumnSearch = MCINTDISPLAYCOL_LPCLOOKUPCLAIM;
      // 2. Name Lookup
      //*TODO:** can't found type for with block
      //*With lpcLookupName
      __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = lpcLookupName;
      modComboBox.fnInitializefpCombo(lpcIn:=lpcLookupName, bShowColHeaders:=False, bSortable:=False, lngNbrOfCols:=4, lngEditCol:=mcintDisplayCol_lpcLookupName, lngNbrOfRowsInDropdown:=8);
      // Since there are multiple visible columns, show lines on this one
      w___TYPE_NOT_FOUND.ListApplyTo = ListApplyToAllCols;
      w___TYPE_NOT_FOUND.LineStyle = LineStyleLowered;
      // Column definitions
      //' First column, Primary sort
      w___TYPE_NOT_FOUND.Col = 0;
      w___TYPE_NOT_FOUND.ColHeaderText = MCSTRIPTCLMINSDLASTNMLABEL;
      w___TYPE_NOT_FOUND.ColName = MCSTRDISPLAYCOL;
      //' Second column, first Secondary sort
      w___TYPE_NOT_FOUND.Col = 1;
      w___TYPE_NOT_FOUND.ColHeaderText = MCSTRIPTCLMINSDFIRSTNMLABEL;
      w___TYPE_NOT_FOUND.ColName = MCSTRCLMINSDFIRSTNM;
      //' Third column, second Secondary sort
      w___TYPE_NOT_FOUND.Col = 2;
      w___TYPE_NOT_FOUND.ColHeaderText = MCSTRTXTCLMNUMLABEL;
      w___TYPE_NOT_FOUND.ColName = MCSTRCLMNUM;
      w___TYPE_NOT_FOUND.ColWidth = 20;
      //' Fourth column
      w___TYPE_NOT_FOUND.Col = 3;
      w___TYPE_NOT_FOUND.ColHeaderText = MCSTRTXTCLMIDLABEL;
      w___TYPE_NOT_FOUND.ColName = MCSTRCLMID;
      w___TYPE_NOT_FOUND.ColHide = true;
      w___TYPE_NOT_FOUND.ColumnSearch = MCINTDISPLAYCOL_LPCLOOKUPNAME;
      // 3. SSN Lookup
      //*TODO:** can't found type for with block
      //*With lpcLookupSSN
      __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = lpcLookupSSN;
      modComboBox.fnInitializefpCombo(lpcIn:=lpcLookupSSN, bShowColHeaders:=False, bSortable:=False, lngNbrOfCols:=3, lngEditCol:=mcintDisplayCol_lpcLookupSSN, lngNbrOfRowsInDropdown:=8);
      // Since there are multiple visible columns, show lines on this one
      w___TYPE_NOT_FOUND.ListApplyTo = ListApplyToAllCols;
      w___TYPE_NOT_FOUND.LineStyle = LineStyleLowered;
      // Column definitions
      //' First column, Primary sort
      w___TYPE_NOT_FOUND.Col = 0;
      w___TYPE_NOT_FOUND.ColHeaderText = MCSTRIPMCLMINSDSSNNUMLABEL;
      w___TYPE_NOT_FOUND.ColName = MCSTRDISPLAYCOL;
      //' Second column, second Secondary sort
      w___TYPE_NOT_FOUND.Col = 1;
      w___TYPE_NOT_FOUND.ColHeaderText = MCSTRTXTCLMNUMLABEL;
      w___TYPE_NOT_FOUND.ColName = MCSTRCLMNUM;
      w___TYPE_NOT_FOUND.ColWidth = 20;
      //' Third column
      w___TYPE_NOT_FOUND.Col = 2;
      w___TYPE_NOT_FOUND.ColHeaderText = MCSTRTXTCLMIDLABEL;
      w___TYPE_NOT_FOUND.ColName = MCSTRCLMID;
      w___TYPE_NOT_FOUND.ColHide = true;
      w___TYPE_NOT_FOUND.ColumnSearch = MCINTDISPLAYCOL_LPCLOOKUPSSN;

      //...............................................................................
      // Set our fpCombo control settings, for those used as multi-column comboboxes.
      //...............................................................................
      // 1. ADMN_SYST_CD
      //*TODO:** can't found type for with block
      //*With lpcAdmnSystCd
      __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = lpcAdmnSystCd;
      modComboBox.fnInitializefpCombo(lpcIn:=lpcAdmnSystCd, bShowColHeaders:=False, bSortable:=True, lngNbrOfCols:=2, lngEditCol:=mcintDisplayCol_lpcAdmnSystCd, lngNbrOfRowsInDropdown:=8);
      // Column definitions
      //' 1st column (description), Primary sort
      w___TYPE_NOT_FOUND.Col = MCINTDISPLAYCOL_LPCADMNSYSTCD;
      w___TYPE_NOT_FOUND.ColName = MCSTRADMNSYSTDSC;
      w___TYPE_NOT_FOUND.ColSortSeq = 0;
      w___TYPE_NOT_FOUND.ColSorted = SortedAscending;
      //' 2nd column (code)
      w___TYPE_NOT_FOUND.Col = MCINTSTORECOL_LPCADMNSYSTCD;
      w___TYPE_NOT_FOUND.ColName = MCSTRADMNSYSTCD;
      w___TYPE_NOT_FOUND.ColHide = true;
      // 2. PYCO_TYP_CD
      //*TODO:** can't found type for with block
      //*With lpcPycoTypCd
      __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = lpcPycoTypCd;
      modComboBox.fnInitializefpCombo(lpcIn:=lpcPycoTypCd, bShowColHeaders:=False, bSortable:=True, lngNbrOfCols:=2, lngEditCol:=mcintDisplayCol_lpcPycoTypCd, lngNbrOfRowsInDropdown:=8);
      // Column definitions
      //' 1st column (description), Primary sort
      w___TYPE_NOT_FOUND.Col = MCINTDISPLAYCOL_LPCPYCOTYPCD;
      w___TYPE_NOT_FOUND.ColName = MCSTRPYCOTYPDSC;
      w___TYPE_NOT_FOUND.ColSortSeq = 0;
      w___TYPE_NOT_FOUND.ColSorted = SortedAscending;
      //' 2nd column (code)
      w___TYPE_NOT_FOUND.Col = MCINTSTORECOL_LPCPYCOTYPCD;
      w___TYPE_NOT_FOUND.ColName = MCSTRPYCOTYPCD;
      w___TYPE_NOT_FOUND.ColHide = true;

      // Set the control to receive the focus after errors (the first editable field
      // on the screen), dependent upon whether we're in Add Mode or not. If in Add mode,
      // this control would typically be the first control that corresponds to a Key field.
      // If not in Add mode, this control would typically be the topmost/leftmost
      // "always updateable" control on the screen (excepting the Lookup ComboBox).
      mctlFirstUpdateableField_Add = lpcAdmnSystCd;
      mctlFirstUpdateableField_Upd = lpcAdmnSystCd;

      // NOTE: The next IF block probably isn't necessary now that the Insured
      // screen is no longer automatically displayed after the user initially
      // logs on.
      // Allow the progress meter on the splash screen to get updated
      if (modGeneral.fnIsFormLoaded("frmSplash")) {
        DoEvents;
      }

      // Instantiate and initialize a table wrapper object for the appropriate table(s).
      mtWrapper = new ctclmClaim();
      // Instantiate and initialize a table wrapper object for the Payee table. This will be used
      // to get data associated with the current claim.
      mtPayee = new ctpyePayee();

      // Bind the on-screen controls to the table wrapper class properties with which they
      // are associated. set default settings for those controls' properties, and
      // bind editable TextBoxes controls to the Extended TextBox class so they will
      // behave appropriately and in a consistent manner.
      fnSetupScreenControls();

      // Populate all ComboBoxes and ListPro controls
      fnRefreshAllCombos();

      // Always go into Add mode, per the user, to ensure they don't inadvertently start
      // editing that first record.
      fnAddRecord();

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
          mtWrapper.getLookupData();
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
            mtWrapper.getRelativeRecord(mtWrapper.getClmNum(), epdSameRecord);
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
  private void form_Resize() {
    // Comments  : Resize the form
    // Parameters: N/A
    // Modified  :
    // --------------------------------------------------
    try {
      "Form_Resize"
.equals(Const cstrCurrentProc As String);
      //'855
      Const(cintCmdAddPayees_Orig_FormHeightLessBtnTop As Integer == 940);
      Const(cintFraPayees_Orig_FormWidthLessFrameWidth As Integer == 330);
      Const(cintFraPayees_Orig_Height As Integer == 1980);
      Const(cintSpacerBorderAroundAllEdgesOfForm As Integer == 15);
      Const(cintMsgPayees_Orig_Width As Integer == 11925);
      Const(cintMsgPayees_Orig_Height As Integer == 1515);
      Const(cintLblGridInstructions_Orig_Left As Integer == 7320);

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      if (Me.WindowState == vbNormal  && Me.Visible) {
        // Bypass if vbMinimized or vbMaximized, to avoid run-time error 384
        // which says" "a form can't be moved or sized while minimized or maximized"
        if ((Me.Height < MCLNGMINFORMHEIGHT)) {
          Me.Height = MCLNGMINFORMHEIGHT;
        }
        if ((Me.Width < MCLNGMINFORMWIDTH)) {
          Me.Width = MCLNGMINFORMWIDTH;
        }

        cmdAddPayee.Left = (Me.Width - cmdAddPayee.Width) / 2;
        cmdAddPayee.Top = Me.Height - cintCmdAddPayees_Orig_FormHeightLessBtnTop;
        fraPayees.Width = Me.Width - cintFraPayees_Orig_FormWidthLessFrameWidth;
        fraPayees.Height = cintFraPayees_Orig_Height + Me.Height - (MCLNGMINFORMHEIGHT + cintSpacerBorderAroundAllEdgesOfForm);
        msgPayees.Width = cintMsgPayees_Orig_Width + Me.Width - (MCLNGMINFORMWIDTH + (cintSpacerBorderAroundAllEdgesOfForm * 2));
        msgPayees.Height = cintMsgPayees_Orig_Height + Me.Height - (MCLNGMINFORMHEIGHT + cintSpacerBorderAroundAllEdgesOfForm);
        lblGridInstructions.Left = Me.Width - (cintLblGridInstructions_Orig_Left - (cintSpacerBorderAroundAllEdgesOfForm * 2));
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
  private void form_Unload(int pintCancel) { // TODO: Use of ByRef founded Private Sub Form_Unload(ByRef pintCancel As Integer)
    // Comments  : Close the form
    // Parameters: pintCancel (in/out), if set to True
    //             the unload is aborted
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

      modGeneral.fnFreeObject(mtWrapper);
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
  private void ipmClmInsdSsnNum_Change() {
    // Comments  : Sets a flag to indicate the current record has been
    //             edited, and thus Update button becomes enabled
    // Parameters: N/A
    // Modified  :
    // --------------------------------------------------
    try {
      "fpmClmInsdSsnNum_Change"
.equals(Const cstrCurrentProc As String);

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      // Ensure availability of navigation & command buttons is set appropriately
      fnInitializeEditMode();

      // Set the hidden Claim Number field.
      fnSetTxtClmNum();
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
  private void iptClmInsdFirstNm_Change() {
    // Comments  : Sets a flag to indicate the current record has been
    //             edited, and thus Update button becomes enabled
    // Parameters: N/A
    // Modified  :
    // --------------------------------------------------
    try {
      "fptClmInsdFirstNm_Change"
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
  private void iptClmInsdLastNm_Change() {
    // Comments  : Sets a flag to indicate the current record has been
    //             edited, and thus Update button becomes enabled
    // Parameters: N/A
    // Modified  :
    // --------------------------------------------------
    try {
      "fptClmInsdFirstNm_Change"
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
  private void iptClmPolNum_Change() {
    // Comments  : Sets a flag to indicate the current record has been
    //             edited, and thus Update button becomes enabled
    // Parameters: N/A
    // Modified  :
    // --------------------------------------------------
    try {
      "fptClmPolNum_Change"
.equals(Const cstrCurrentProc As String);
      String strDefaultPycoTypDsc = "";

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      // Ensure availability of navigation & command buttons is set appropriately
      fnInitializeEditMode();

      // Use the store column (ADMN_SYST_CD) column of the Admin System combobox,
      // then determine if the Company Type should change based on the new input
      lpcAdmnSystCd.Col = MCINTSTORECOL_LPCADMNSYSTCD;

      strDefaultPycoTypDsc = fnGetDefaultPayorCompany(iptClmPolNum.Text, lpcAdmnSystCd.ColText);
      modComboBox.fnSearchFPCombo(lpcPycoTypCd, strDefaultPycoTypDsc, MCINTDISPLAYCOL_LPCPYCOTYPCD);

      // Set the hidden Claim Number field.
      fnSetTxtClmNum();
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
  private void lpcAdmnSystCd_Change() {
    // Comments  : Sets a flag to indicate the current record has been
    //             edited, and thus Update button becomes enabled. It also
    //             dictates how long the Policy Number (ClmPolNum) can be, given
    //             the Admin System chosen.
    // Parameters: N/A
    // Modified  : Berry Kropiwka - 2019-11-06 - Added code to enable or disable the Compact Filling check box bases on admin system and state
    // --------------------------------------------------
    try {
      "lpcAdmnSystCd_Change"
.equals(Const cstrCurrentProc As String);
      String strDefaultPycoTypDsc = "";

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      fnInitializeEditMode();

      // Get metadata (min/max policy # allowed, tax rptg ind, default pyco type code) and store in module-level variables
      fnGetAdminSysMetadata();

      //*TODO:** can't found type for with block
      //*With iptClmPolNum
      __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = iptClmPolNum;
      // There is no .MinLength property on this control  :-)
      w___TYPE_NOT_FOUND.MaxLength = mintAdmnSyst_MaxPolNumLength;

      // Use the store column (ADMN_SYST_CD) column of the Admin System combobox,
      // then determine if the Company Type should change based on the new input
      lpcAdmnSystCd.Col = MCINTSTORECOL_LPCADMNSYSTCD;

      strDefaultPycoTypDsc = fnGetDefaultPayorCompany(iptClmPolNum.Text, lpcAdmnSystCd.ColText);
      modComboBox.fnSearchFPCombo(lpcPycoTypCd, strDefaultPycoTypDsc, MCINTDISPLAYCOL_LPCPYCOTYPCD);

      // Set the hidden Claim Number field.
      fnSetTxtClmNum();

      // enable or disable the Compact Filling check box based on Admin System
      fnSetCompactFillingCheckBox(Me.lpcAdmnSystCd.Text, Me.cboIssStCd.Text);

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
  private void lpcAdmnSystCd_GotFocus() {
    //-----------------------------------------------------------------------------
    // Function     lpcAdmnSystCd_GotFocus
    // Purpose      Display the drop down list now that the user has entered this control.
    // Parameters   N/A
    // Returns      N/A
    // Date:        12/19/2001
    //-----------------------------------------------------------------------------
    "lpcAdmnSystCd_GotFocus"
.equals(Const cstrCurrentProc As String);
    try {

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      //lpcAdmnSystCd.ListDown = True
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
  private void lpcLookupClaim_Click() {
    // Comments  : Retrieve selected record
    // Parameters: N/A
    //
    // --------------------------------------------------
    "lpcLookupClaim_Click"
.equals(Const cstrCurrentProc As String);

    try {

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      fnPerformLookup(lpcLookupClaim);
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
  private void lpcLookupClaim_GotFocus() {
    //-----------------------------------------------------------------------------
    // Function     lpcLookupClaim_GotFocus
    // Purpose      Turn on Lookup Mode and drop down the list now that the user has entered this control.
    // Parameters   N/A
    // Returns      N/A
    // Date:        12/19/2001
    //-----------------------------------------------------------------------------
    "lpcLookupClaim_GotFocus"
.equals(Const cstrCurrentProc As String);
    try {

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      //lpcLookupClaim.ListDown = True

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
  private void lpcLookupClaim_KeyDown(int intKeyCode, int intShift) { // TODO: Use of ByRef founded Private Sub lpcLookupClaim_KeyDown(ByRef intKeyCode As Integer, ByRef intShift As Integer)
    //-----------------------------------------------------------------------------
    // Function     lpcLookupClaim_KeyDown
    // Purpose      If the user presses Enter, make it do just what the Click event does
    //              (i.e. display the selected record)
    // Parameters   intKeyCode - ASCII code of key that was pressed
    //              intShift - indicates whether the Shift key was pressed
    // Returns      N/A
    //-----------------------------------------------------------------------------
    "lpcLookupClaim_KeyDown"
.equals(Const cstrCurrentProc As String);
    try {

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      if (intKeyCode == vbKeyReturn) {
        fnPerformLookup(lpcLookupClaim);
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
  private void lpcLookupClaim_LostFocus() {
    //-----------------------------------------------------------------------------
    // Function     lpcLookupClaim_LostFocus
    // Purpose      Turn off Lookup Mode now that the user has left that control.
    // Parameters   N/A
    // Returns      N/A
    // Date:        12/19/2001
    //-----------------------------------------------------------------------------
    "lpcLookupClaim_LostFocus"
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
      //*With lpcLookupClaim
      __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = lpcLookupClaim;
      w___TYPE_NOT_FOUND.Row = clngFirstRow;
      w___TYPE_NOT_FOUND.ListIndex = clngFirstRow;
      w___TYPE_NOT_FOUND.Action = ActionClearSearchBuffer;

      //fnSearchFPCombo lpcLookupClaim, gcstrBlankEntry, mcintDisplayCol_lpcLookupClaim
      lpcLookupClaim.Refresh;

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
  private void lpcLookupName_Click() {
    // Comments  : Retrieve selected record
    // Parameters: N/A
    // Modified  : CMP 4/27/2002
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
    "lpcLookupClaim_KeyDown"
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
  private void lpcLookupSSN_Click() {
    // Comments  : Retrieve selected record
    // Parameters: N/A
    // Modified  :
    //
    // --------------------------------------------------
    "lpcLookupSSN_Click"
.equals(Const cstrCurrentProc As String);

    try {

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      fnPerformLookup(lpcLookupSSN);
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
  private void lpcLookupSSN_GotFocus() {
    //-----------------------------------------------------------------------------
    // Function     lpcLookupSSN_GotFocus
    // Purpose      Turn on Lookup Mode and drop down the list now that the user has entered this control.
    // Parameters   N/A
    // Returns      N/A
    // Date:        12/19/2001
    //-----------------------------------------------------------------------------
    "lpcLookupSSN_GotFocus"
.equals(Const cstrCurrentProc As String);
    try {

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      //lpcLookupSSN.ListDown = True

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
  private void lpcLookupSSN_KeyDown(int intKeyCode, int intShift) { // TODO: Use of ByRef founded Private Sub lpcLookupSSN_KeyDown(ByRef intKeyCode As Integer, ByRef intShift As Integer)
    //-----------------------------------------------------------------------------
    // Function     lpcLookupSSN_KeyDown
    // Purpose      If the user presses Enter, make it do just what the Click event does
    //              (i.e. display the selected record)
    // Parameters   intKeyCode - ASCII code of key that was pressed
    //              intShift - indicates whether the Shift key was pressed
    // Returns      N/A
    //-----------------------------------------------------------------------------
    "lpcLookupSSN_KeyDown"
.equals(Const cstrCurrentProc As String);
    try {

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      if (intKeyCode == vbKeyReturn) {
        fnPerformLookup(lpcLookupSSN);
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
  private void lpcLookupSSN_LostFocus() {
    //-----------------------------------------------------------------------------
    // Function     lpcLookupSSN_LostFocus
    // Purpose      Turn off Lookup Mode now that the user has left that control.
    // Parameters   N/A
    // Returns      N/A
    // Date:        12/19/2001
    //-----------------------------------------------------------------------------
    "lpcLookupSSN_LostFocus"
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
      //*With lpcLookupSSN
      __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = lpcLookupSSN;
      w___TYPE_NOT_FOUND.Row = clngFirstRow;
      w___TYPE_NOT_FOUND.ListIndex = clngFirstRow;
      w___TYPE_NOT_FOUND.Action = ActionClearSearchBuffer;
      //fnSearchFPCombo lpcLookupSSN, gcstrBlankEntry, mcintDisplayCol_lpcLookupSSN
      lpcLookupSSN.Refresh;

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
  private void lpcPycoTypCd_Change() {
    // Comments  : Sets a flag to indicate the current record has been
    //             edited, and thus Update button becomes enabled
    // Parameters: N/A
    // Modified  :
    // --------------------------------------------------
    try {
      "lpcPycoTypCd_Change"
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
  private void lpcPycoTypCd_GotFocus() {
    //-----------------------------------------------------------------------------
    // Function     lpcPycoTypCd_GotFocus
    // Purpose      Display the drop down list now that the user has entered this control.
    // Parameters   N/A
    // Returns      N/A
    // Date:        12/19/2001
    //-----------------------------------------------------------------------------
    "lpcPycoTypCd_GotFocus"
.equals(Const cstrCurrentProc As String);
    try {

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      //lpcPycoTypCd.ListDown = True
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
  private void msgPayees_DblClick() {
    // Comments  : This event handler is triggered when the user double-clicks in
    //             the Payee grid to indicate they want to edit that Payee
    // Parameters:  -
    // Modified  :
    // --------------------------------------------------
    try {
      "msgPayee_DblClick"
.equals(Const cstrCurrentProc As String);
      Form frmChild = null;
      String strSaveClaimNumber = "";
      chrgHourglass hrgHourglass = null;
      int lngReturnValue = 0;
      String strACF2 = "";

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      // Set the current column to 1 (the Payee Name)
      fnSetPropertiesForPayeeScreen(bSendEmptyName:=False);
      // Following statement triggers the Form_Initialize & Form_Load events in frmPayee
      frmChild = new frmPayee();
      // Following statement triggers the Form_Activate event in frmPayee
      frmChild.Show(vbModal);

      // Do DoEvents to allow the Payee screen to fully disappear.
      DoEvents;

      fnWindowLock(Me.cbrfBrowseFolder.setHWnd());

      hrgHourglass = new chrgHourglass();
      hrgHourglass.setValue(true);

      // Update the Payees Recordset to reflect any Payees just added, changed or deleted
      // when the Payees screen was open. Then, update the msgPayees grid and recalculate
      // totals across all Payees.
      // Note: You *must* requery the Insured and Payee recordsets to accomodate the possibility
      //       that another user (a) add/changed/deleted one more Payees for the
      //       current Insured and (b) returned to the Insured screen which triggered an update
      //       to the Insured record for the claim-wide totals it carries. If you don't do the
      //       requeries then a -2147217864 "row cannot be located for updating..." error could
      //       occur. So, we'll do the requerying automatically with no visible indication to the
      //       user that it occured unless the requerying revealed that another user deleted the
      //       current claim number and hence the Insured with the next higher claim number will
      //       be displayed (otherwise the same claim remains being displayed).
      strSaveClaimNumber = iptClmPolNum.Text;
      hrgHourglass.setValue(false);

      //!TODO! The following looks like unnecessary (i.e. dead) code
      //   If iptClmPolNum <> strSaveClaimNumber Then
      //   '!TODO! Gen msg via frmMsgBox
      //        'MsgBox "Another user has deleted the Claim Number (" & strSaveClaimNumber & ") you were viewing.", _
      //        '       vbOKOnly + vbInformation, mcstrDialogTitle
      //    End If

      hrgHourglass.setValue(true);

      fnGetChildren();

      // 01/31/2001 BAW - Add another Refresh to speed up repainting
      Me.Refresh;

      // Totals may have changed. Update the Insured record just in case.
      fnLoadRecordWithCalculatedControls();

      // 01/31/2001 BAW - Add another Refresh to speed up repainting
      Me.Refresh;

      // Determine whether another user updated or deleted the record about to be updated.
      // Note: this multi-user checking is performed on an Update but not an Add.
      lngReturnValue = mtWrapper.checkForAnotherUsersChanges(ewoUpdate, strACF2);

      if (lngReturnValue == vbObjectError + modResConstants.gCRES_INFO_ANOTHER_USER_DELETED_NOT_SAVED) {
        modGeneral.gerhApp.reportNonFatal(vbObjectError + modResConstants.gCRES_INFO_ANOTHER_USER_DELETED_NOT_SAVED, mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
        // Discard *this* user's pending changes and show the previous record.
        // Can't use the GetClmNumFromClmId( ) method since the CLAIM_T row
        // doesn't exist and hence a "-2147217900" (Claim ID does not exist) error
        // throws things off.
        mtWrapper.getRelativeRecord(mtWrapper.getClmNum(), epdPreviousRecord);
        // Do NOT bother to check for another UPDATING the record, since all we're doing is
        // updating the total fields. Let the totals update go through.
        //   ElseIf lngReturnValue = vbObjectError + gcRES_INFO_ANOTHER_USER_UPDATED_DISCARDED Then
        //           gerhApp.ReportNonFatal vbObjectError + gcRES_INFO_ANOTHER_USER_UPDATED_DISCARDED, _
        //                                   mstrScreenName & gcstrDOT & cstrCurrentProc, _
        //                                   Trim$(strACF2)
        //       ' Discard *this* user's pending changes by re-retrieving the current record
        //       ' as it currently looks on the database and refreshing the lookup recordset.
        //       ' Can't use the GetClmNumFromClmId( ) method since the CLAIM_T row
        //       ' doesn't exist and hence a "-2147217900" (Claim ID does not exist) error
        //       ' throws things off.
        //       .GetRelativeRecord .ClmNum, epdSameRecord
      } 
      else {
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

      // Have to call fnLoadControls here, like in cmdAdd_Click and cmdDelete_Click and cmdUpdate_Click,
      // to ensure refreshed comboboxes have their previous value still selected.
      if (mtWrapper.getLookupRecordCount() > 0) {
        // Ensure the on-screen controls reflect the record just added/updated, in case the
        // DBMS altered it in some way, e.g., determining an Identity column value and
        // getting the most up-to-date Last Updated info. This also sets the navigation
        // buttons and updates the "record x of y" label
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
    // Terminate the Payee form, removing it from the Forms collection
    modGeneral.fnFreeObject(frmChild);
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

}

public class EnumLookupType {
    public static final int ELT_CLAIM = 0;
    public static final int ELT_NAME = 1;
    public static final int ELT_SSN = 2;
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


case class RminsuredData(
              id: Option[Int],

              )

object Rminsureds extends Controller with ProvidesUser {

  val rminsuredForm = Form(
    mapping(
      "id" -> optional(number),

  )(RminsuredData.apply)(RminsuredData.unapply))

  implicit val rminsuredWrites = new Writes[Rminsured] {
    def writes(rminsured: Rminsured) = Json.obj(
      "id" -> Json.toJson(rminsured.id),
      C.ID -> Json.toJson(rminsured.id),

    )
  }

  def get(id: Int) = GetAction { implicit request =>
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.LIST_RMINSURED), { user =>
      Ok(Json.toJson(Rminsured.get(user, id)))
    })
  }

  def update(id: Int) = PostAction { implicit request =>
    Logger.debug("in rminsureds.update")
    rminsuredForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      rminsured => {
        Logger.debug(s"form: ${rminsured.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.EDIT_RMINSURED), { user =>
          Ok(
            Json.toJson(
              Rminsured.update(user,
                Rminsured(
                       id,

                ))))
        })
      }
    )
  }

  def create = PostAction { implicit request =>
    Logger.debug("in rminsureds.create")
    rminsuredForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      rminsured => {
        Logger.debug(s"form: ${rminsured.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.NEW_RMINSURED), { user =>
          Ok(
            Json.toJson(
              Rminsured.create(user,
                Rminsured(

                ))))
        })
      }
    )
  }

  def delete(id: Int) = PostAction { implicit request =>
    Logger.debug("in rminsureds.delete")
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.DELETE_RMINSURED), { user =>
      Rminsured.delete(user, id)
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

case class Rminsured(
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

object Rminsured {

  lazy val emptyRminsured = Rminsured(
)

  def apply(
      id: Int,
) = {

    new Rminsured(
      id,
)
  }

  def apply(
) = {

    new Rminsured(
)
  }

  private val rminsuredParser: RowParser[Rminsured] = {
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
        Rminsured(
              id,
,
              createdAt,
              updatedAt,
              updatedBy)
    }
  }

  def create(user: CompanyUser, rminsured: Rminsured): Rminsured = {
    save(user, rminsured, true)
  }

  def update(user: CompanyUser, rminsured: Rminsured): Rminsured = {
    save(user, rminsured, false)
  }

  private def save(user: CompanyUser, rminsured: Rminsured, isNew: Boolean): Rminsured = {
    def getFields = {
      List(

      )
    }
    def throwException = {
      throw new RuntimeException(s"Error when saving ${C.RMINSURED}")
    }

    DBHelper.saveEx(
      user,
      Register(
        C.RMINSURED,
        C.ID,
        rminsured.id,
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

  def load(user: CompanyUser, id: Int): Option[Rminsured] = {
    loadWhere(user, s"${C.ID} = {id}", 'id -> id)
  }

  def loadWhere(user: CompanyUser, where: String, args : scala.Tuple2[scala.Any, anorm.ParameterValue[_]]*) = {
    DB.withConnection(user.database.database) { implicit connection =>
      SQL(s"SELECT t1.*, t2.${C.FK_NAME} FROM ${C.RMINSURED} t1 INNER JOIN ${C.???} t2 ON t1.${C.FK_ID} = t2.${C.FK_ID} WHERE $where")
        .on(args: _*)
        .as(rminsuredParser.singleOpt)
    }
  }

  def delete(user: CompanyUser, id: Int) = {
    DB.withConnection(user.database.database) { implicit connection =>
      try {
        SQL(s"DELETE FROM ${C.RMINSURED} WHERE ${C.ID} = {id}")
        .on('id -> id)
        .executeUpdate
      } catch {
        case NonFatal(e) => {
          Logger.error(s"can't delete a ${C.RMINSURED}. ${C.ID} id: $id. Error ${e.toString}")
          throw e
        }
      }
    }
  }

  def get(user: CompanyUser, id: Int): Rminsured = {
    load(user, id) match {
      case Some(p) => p
      case None => emptyRminsured
    }
  }
}


// Router

GET     /api/v1/general/rminsured/:id              controllers.logged.modules.general.Rminsureds.get(id: Int)
POST    /api/v1/general/rminsured                  controllers.logged.modules.general.Rminsureds.create
PUT     /api/v1/general/rminsured/:id              controllers.logged.modules.general.Rminsureds.update(id: Int)
DELETE  /api/v1/general/rminsured/:id              controllers.logged.modules.general.Rminsureds.delete(id: Int)




/**/
