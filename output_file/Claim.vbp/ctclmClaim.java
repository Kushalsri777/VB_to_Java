
import java.util.Date;

public class ctclmClaim {

  //!TODO! Add methods (or screen code??) to load m_rstLookup to appropriate comboboxes
  //       and translate selected cbo entry to corresponding m_rstLookup row.
  //!TODO! GUI Code - make SSN required if LOB = group; optional otherwise
  //!TODO! Add code to set/get clm_num, for group, based on clm_pol_num and clm_insd_ssn_num;
  //       otherwise set it to clm_pol_num only.  Changes to Property Set/Get methods??

  //--------------------------------------------------------------------------
  // Procedure:   ctclmClaim
  // Description: Provides properties and methods to support the claim_t table values
  //
  //              Do a Find on "!CUSTOMIZE!" to locate places where the table
  //              wrapper must be changed to work for a different table.
  //
  //              NOTE: PUBLIC PROPERTIES corresponding to table columns should be
  //                    named so that they reflect the table column name with
  //                    underscores eliminated and each word beginning with a
  //                    capital letter.
  //
  //                    The naming convention for MEMBER VARIABLES that store
  //                    public properties should be follow this standard:
  //                    "m_dddPPPP" where "m_" identifies it as a member variable,
  //                    "ddd" indicates the data type (i.e. lng, str, dte) and PPPP
  //                    is the name of the public property to which it corresponds.
  //
  //                 Examples:
  //
  //                    TableCol         PublicProperty  MemberVariable     Constant
  //                    ---------------  --------------  -----------------  ----------------
  //                    clm_num          ClmNum          m_strClmNum        mcstrClmNum
  //                    clm_proof_dt     ClmProofDt      m_dteClmProofDt    mcstrClmProofDt
  //
  //
  //                 NOTE also that navigation should be done via the **ADO Wrapper's**
  //                 navigation methods instead of directly referencing the navigation
  //                 methods on a ADODB.Recordset object!
  //
  // Revisions:
  //   Modified:     Berry Kropiwka - Added Compact Calc - 2019-09-27
  //
  //
  // Procedures  :
  //   Private     Class_Initialize()
  //   Private     Class_Terminate()
  //   Public      Property Get AdmnSystCd() As String
  //   Public      Property Let AdmnSystCd(ByVal strValue As String)
  //   Public      Property Get AllowableCharacters(ByVal strTagIn As String) As String
  //   Public      Property Get ClmCompactClcnInd() As Boolean
  //   Public      Property Let ClmCompactClcnInd(ByVal bValue As Boolean)
  //   Public      Property Get ClmForResDthInd() As Boolean
  //   Public      Property Let ClmForResDthInd(ByVal bValue As Boolean)
  //   Public      Property Get ClmId() As Long
  //   Public      Property Let ClmId(ByVal lngValue As Long)
  //   Public      Property Get ClmInsdDthDt() As Date
  //   Public      Property Let ClmInsdDthDt(ByVal dteValue As Date)
  //   Public      Property Get ClmInsdFirstNm() As String
  //   Public      Property Let ClmInsdFirstNm(ByVal strValue As String)
  //   Public      Property Get ClmInsdLastNm() As String
  //   Public      Property Let ClmInsdLastNm(ByVal strValue As String)
  //   Public      Property Get ClmInsdSsnNum() As String
  //   Public      Property Let ClmInsdSsnNum(ByVal strValue As String)
  //   Public      Property Get ClmNum() As String
  //   Public      Property Let ClmNum(ByVal strValue As String)
  //   Public      Property Get ClmPolNum() As String
  //   Public      Property Let ClmPolNum(ByVal strValue As String)
  //   Public      Property Get ClmProofDt() As Date
  //   Public      Property Let ClmProofDt(ByVal dteValue As Date)
  //   Public      Property Get ClmTotClmPdAmt() As Double
  //   Public      Property Let ClmTotClmPdAmt(ByVal dblValue As Double)
  //   Public      Property Get ClmTotDthbPmtAmt() As Double
  //   Public      Property Let ClmTotDthbPmtAmt(ByVal dblValue As Double)
  //   Public      Property Get ClmTotIntAmt() As Double
  //   Public      Property Let ClmTotIntAmt(ByVal dblValue As Double)
  //   Public      Property Get ClmTotWthldAmt() As Double
  //   Public      Property Let ClmTotWthldAmt(ByVal dblValue As Double)
  //   Public      Property Get CurrentLookupRecordNumber() As Long
  //   Public      Property Get DecimalPositions(ByVal strTagIn As String) As Integer
  //   Public      Property Get DefaultValue(ByVal strTagIn As String) As Variant
  //   Public      Property Get DollarPositions(ByVal strTagIn As String) As Integer
  //   Public      Property Get Format(ByVal strTagIn As String) As String
  //   Public      Property Get InsdDthResStCd() As String
  //   Public      Property Let InsdDthResStCd(ByVal strValue As String)
  //   Public      Property Get IsKey(ByVal strTagIn As String) As Boolean
  //   Public      Property Get IsNullable(ByVal strTagIn As String) As Boolean
  //   Public      Property Get IssStCd() As String
  //   Public      Property Let IssStCd(ByVal strValue As String)
  //   Public      Property Get LookupData() As ADODB.Recordset
  //   Public      Property Get LookupData_Claim() As ADODB.Recordset
  //   Public      Property Get LookupData_Name() As ADODB.Recordset
  //   Public      Property Get LookupData_SSN() As ADODB.Recordset
  //   Public      Property Get LookupIsAtBOF() As Boolean
  //   Public      Property Get LookupIsAtEOF() As Boolean
  //   Public      Property Get LookupRecordCount() As Long
  //   Public      Property Get LstUpdtDtm() As Date
  //   Public      Property Let LstUpdtDtm(ByVal NewValue As String)
  //   Public      Property Get LstUpdtUserId() As String
  //   Public      Property Let LstUpdtUserId(ByVal strValue As String)
  //   Public      Property Get Mask(ByVal strTagIn As String) As String
  //   Public      Property Get MaxCharacters(ByVal strTagIn As String) As Long
  //   Public      Property Get PycoTypCd() As String
  //   Public      Property Let PycoTypCd(ByVal strValue As String)
  //   Public      Property Get ShouldForceToUppercase(ByVal strTagIn As String) As Boolean
  //   Public      AddRecord() as Boolean
  //   Public      CheckForAnotherUsersChanges(ByVal lngWhatOperation As enumWhatOperationIsBeingAttempted, _
  //                   ByRef strACF2 As String) As Long
  //   Public      DeleteRecord() As Boolean
  //   Public      GetClmIdFromClmNum(ByVal strClmNum As String) As Integer
  //   Public      GetClmNumFromClmID(ByVal lngClmID As Long) As Variant
  //   Public      GetLobCdFromAdmnSystCd(ByVal strAdmnSystCd As String) As String
  //   Public      GetLookupData() As Boolean
  //   Public      GetSingleRecord(ByVal strKey1 As String) As Boolean
  //   Public      GoToFirstRecord()
  //   Public      GoToLastRecord()
  //   Public      GoToNextRecord()
  //   Public      GoToPreviousRecord()
  //   Public      UpdateRecord() As Boolean
  //   Private     fnGetColMetaData(ByRef pudtCol As udtColumn, _
  //                   ByRef prstIn As ADODB.Recordset) As Boolean
  //   Private     fnGetProperty(ByVal strTagIn As String) As udtColumn
  //   Private     GetRelativeRecord(ByVal strKey1 As String, _
  //                   ByVal lngPositionDirection As enumPositionDirection) As Boolean
  //   Private     fnLoadColMetaData() As Boolean
  //   Private     fnSelectRecord(ByVal strKey1 As String) As ADODB.Recordset
  //
  //-----------------------------------------------------------------------------
//Option Explicit
  *Option Compare Binary

  //!CUSTOMIZE! Change both the filename and class name to represent the main table.
  //!CUSTOMIZE! Change mcstrName to reflect the class name, followed by a period.
  private static final String MCSTRNAME = "ctclmClaim.";

  //...............................................................................................
  //!CUSTOMIZE!
  // These are the private variables corresponding to PUBLIC properties.
  // There should be one (of type udtColumn) for each column in the table that this class accesses.
  //...............................................................................................
  private udtColumn m_strAdmnSystCd;
  private udtColumn m_lngClmId;
  private udtColumn m_dteClmInsdDthDt;
  private udtColumn m_strClmInsdFirstNm;
  private udtColumn m_strClmInsdLastNm;
  private udtColumn m_strClmInsdSsnNum;
  private udtColumn m_strClmNum;
  private udtColumn m_strClmPolNum;
  private udtColumn m_dteClmProofDt;
  private udtColumn m_varClmTotClmPdAmt;
  private udtColumn m_varClmTotDthbPmtAmt;
  private udtColumn m_varClmTotIntAmt;
  private udtColumn m_varClmTotWthldAmt;
  private udtColumn m_strInsdDthResStCd;
  private udtColumn m_strIssStCd;
  private udtColumn m_dteLstUpdtDtm;
  private udtColumn m_strLstUpdtUserId;
  private udtColumn m_strPycoTypCd;
  private udtColumn m_bClmForResDthInd;
  private udtColumn m_bClmCompactClcnInd;

  //...............................................................................................
  //!CUSTOMIZE!
  // Create one Const for each column in the table, defining the table column to which it refers.
  //...............................................................................................
  private static final String MCSTRADMNSYSTCD = "ADMN_SYST_CD";
  private static final String MCSTRCLMID = "CLM_ID";
  private static final String MCSTRCLMINSDDTHDT = "CLM_INSD_DTH_DT";
  private static final String MCSTRCLMINSDFIRSTNM = "CLM_INSD_FIRST_NM";
  private static final String MCSTRCLMINSDLASTNM = "CLM_INSD_LAST_NM";
  private static final String MCSTRCLMINSDSSNNUM = "CLM_INSD_SSN_NUM";
  private static final String MCSTRCLMNUM = "CLM_NUM";
  private static final String MCSTRCLMPOLNUM = "CLM_POL_NUM";
  private static final String MCSTRCLMPROOFDT = "CLM_PROOF_DT";
  private static final String MCSTRCLMTOTCLMPDAMT = "CLM_TOT_CLM_PD_AMT";
  private static final String MCSTRCLMTOTDTHBPMTAMT = "CLM_TOT_DTHB_PMT_AMT";
  private static final String MCSTRCLMTOTINTAMT = "CLM_TOT_INT_AMT";
  private static final String MCSTRCLMTOTWTHLDAMT = "CLM_TOT_WTHLD_AMT";
  private static final String MCSTRINSDDTHRESSTCD = "INSD_DTH_RES_ST_CD";
  private static final String MCSTRISSSTCD = "ISS_ST_CD";
  private static final String MCSTRLSTUPDTDTM = "LST_UPDT_DTM";
  private static final String MCSTRLSTUPDTUSERID = "LST_UPDT_USER_ID";
  private static final String MCSTRPYCOTYPCD = "PYCO_TYP_CD";
  private static final String MCSTRCLMFORRESDTHIND = "CLM_FOR_RES_DTH_IND";
  private static final String MCSTRCLMCOMPACTCLCNIND = "CLM_COMPACT_CLCN_IND";

  //...............................................................................................
  // Other private variables that do NOT correspond to PUBLIC properties.
  //...............................................................................................
  // m_adwADO is a private instantiation of the ADO Wrapper, used to do ADO things like
  // navigation, executing a stored procedure, etc.
  private cadwADOWrapper m_adwADO;

  // The next 2 vars (m_dteLstUpdtDtm_Original and m_strLstUpdtUserId_Original) are used by
  // the CheckForAnotherUsersChanges method to determine if another user affected the
  // record since *this* user originally retrieved the record.
  private Date m_dteLstUpdtDtm_Original = null;
  private String m_strLstUpdtUserId_Original = "";

  // m_rstLookup contains selected columns for each row in the table and is used by the form
  // to populate its 3 Lookup VSFlexGrid controls that the user uses to hop directly to a desired record.
  // m_rstLookup should be PRIVATE! If anyone besides this class needs to reference properties of this
  // Recordset, then those properties should be exposed as public properties of this class.
  private DBRecordSet m_rstLookup;


  ///\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
  //\                                                                  /
  //|          CLASS_INITIALIZE / CLASS_TERMINATE    Procedures        |
  ///                                                                  \
  //\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/

  //////////////////////////////////////////////////////////////////////////////////////////////////
  private void class_Initialize() {
    // **************************************************************************
    // Function  : Class_Initialize
    // Purpose   : Starting Point for Object
    //             >GetLookupData (Recordset of KEy Columns for every row in table)
    //             >Populate Object Field Properties with Table's First Record

    // Parameters: N/A
    // Returns   : Boolean
    // **************************************************************************
    "Class_Initialize"
.equals(Const cstrCurrentProc As String);
    try {

      m_adwADO = new cadwADOWrapper();

      // Refresh lookup RST and set LookupRecordCount / CurrentLookupRecNbr properties
      getLookupData();

      // Get all columns for the 1st record in the Lookup RST and load to member vars.
      // If there are no records (m_rstLookup is Nothing), then initialize the
      // properties that correspond to table columns. (Caller must take action if
      // m_rstLookup Is Nothing!!!)
      if (m_rstLookup.RecordCount != 0) {
        //!CUSTOMIZE! The constant referenced below should refer to the key field.
        getSingleRecord(m_rstLookup.Fields(MCSTRCLMID).value);
      } 
      else {
        fnClearPropertyValues();
      }

      // Obtain meta data about each table column from the DBMS and load it to the
      // properties that correspond to those table columns
      fnLoadColMetaData();
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  private void class_Terminate() {
    // **************************************************************************
    // Function  : Class_Terminate
    // Purpose   : Closes the private recordset variable, then frees members
    //             associated with internal objects
    // Parameters: N/A
    // Returns   : Int
    // **************************************************************************
    "Class_Terminate"
.equals(Const cstrCurrentProc As String);
    try {

      modGeneral.fnFreeRecordset(m_rstLookup);
      modGeneral.fnFreeObject(m_adwADO);
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }
}





///\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
//\                                                                  /
//|                PROPERTY GET/LET    Procedures                    |
///                                                                  \
//\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/

//!CUSTOMIZE! so that there is a Property Get and Let for each table column.

//////////////////////////////////////////////////////////////////////////////////////////////////
  public String getAdmnSystCd() {
    String _rtn = "";
    // **************************************************************************
    // Function  : Property Get AdmnSystCd
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : string
    // **************************************************************************
    "Property Get AdmnSystCd"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = CStr(m_strAdmnSystCd.value);
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}


//////////////////////////////////////////////////////////////////////////////////////////////////
  public void setAdmnSystCd(String strValue) {
    String _rtn = null;
    // **************************************************************************
    // Function  : Property Let AdmnSystCd
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal NewValue As String
    // Returns   :
    // **************************************************************************
    "Property Let AdmnSystCd"
.equals(Const cstrCurrentProc As String);
    try {

      m_strAdmnSystCd.value = strValue;
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public String getAllowableCharacters(String strTagIn) {
    String _rtn = "";
    // **************************************************************************
    // Function  : Property Get AllowableCharacters
    // Purpose   : Retrieves the default Format Mask (i.e. #####.###) from property
    // Parameters: N/A
    // Returns   : String
    // **************************************************************************
    "Property Get AllowableCharacters"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = fnGetProperty(strTagIn).allowableCharacters;
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}


//////////////////////////////////////////////////////////////////////////////////////////////////
  public boolean getClmForResDthInd() {
    boolean _rtn = false;
    // **************************************************************************
    // Function  : Property Get ClmForResDthInd
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : Boolean
    // **************************************************************************
    "Property Get ClmForResDthInd"
.equals(Const cstrCurrentProc As String);
    try {

      // Note that this field is stored as a Y or N.
      _rtn = CBool(m_bClmForResDthInd.value);
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}


//////////////////////////////////////////////////////////////////////////////////////////////////
  public void setClmForResDthInd(boolean bValue) {
    boolean _rtn = null;
    // **************************************************************************
    // Function  : Property Let ClmForResDthInd
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal bValue As Boolean
    // Returns   :
    // **************************************************************************
    "Property Let ClmForResDthInd"
.equals(Const cstrCurrentProc As String);
    try {

      // Note that this field is stored as a Y or N.
      m_bClmForResDthInd.value = bValue;
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}

//////////////////////////////////////////////////////////////////////////////////////////////////
  public boolean getClmCompactClcnInd() {
    boolean _rtn = false;
    // **************************************************************************
    // Function  : Property Get ClmCompactClcnInd
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : Boolean
    // Created   : Berry Kropiwka 2019-09-27
    // **************************************************************************
    "Property Get ClmCompactClcnInd"
.equals(Const cstrCurrentProc As String);
    try {

      // Note that this field is stored as a T or F.
      _rtn = CBool(m_bClmCompactClcnInd.value);
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}


//////////////////////////////////////////////////////////////////////////////////////////////////
  public void setClmCompactClcnInd(boolean bValue) {
    boolean _rtn = null;
    // **************************************************************************
    // Function  : Property Let ClmCompactClcnInd
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal bValue As Boolean
    // Returns   :
    // Created   : Berry Kropiwka 2019-09-27
    // **************************************************************************
    "Property Let ClmCompactClcnInd"
.equals(Const cstrCurrentProc As String);
    try {

      // Note that this field is stored as a T or F.
      m_bClmCompactClcnInd.value = bValue;
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}

//////////////////////////////////////////////////////////////////////////////////////////////////
  public int getClmId() {
    int _rtn = 0;
    // **************************************************************************
    // Function  : Property Get ClmId
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : double
    // **************************************************************************
    "Property Get ClmId"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = Long.parseLong(m_lngClmId.value);
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}


//////////////////////////////////////////////////////////////////////////////////////////////////
  public void setClmId(int lngValue) {
    int _rtn = null;
    // **************************************************************************
    // Function  : Property Let ClmId
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal lngValue As Long
    // Returns   :
    // **************************************************************************
    "Property Let ClmId"
.equals(Const cstrCurrentProc As String);
    try {

      m_lngClmId.value = lngValue;
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}


//////////////////////////////////////////////////////////////////////////////////////////////////
  public Date getClmInsdDthDt() {
    Date _rtn = null;
    // **************************************************************************
    // Function  : Property Get ClmInsdDthDt
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : string
    // **************************************************************************
    "Property Get ClmInsdDthDt"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = G.parseDate(m_dteClmInsdDthDt.value);
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}


//////////////////////////////////////////////////////////////////////////////////////////////////
  public void setClmInsdDthDt(Date dteValue) {
    Date _rtn = null;
    // **************************************************************************
    // Function  : Property Let ClmInsdDthDt
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal NewValue As String
    // Returns   :
    // **************************************************************************
    "Property Let ClmInsdDthDt"
.equals(Const cstrCurrentProc As String);
    try {

      m_dteClmInsdDthDt.value = dteValue;
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public String getClmInsdFirstNm() {
    String _rtn = "";
    // **************************************************************************
    // Function  : Property Get ClmInsdFirstNm
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : string
    // **************************************************************************
    "Property Get ClmInsdFirstNm"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = CStr(m_strClmInsdFirstNm.value);
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}


//////////////////////////////////////////////////////////////////////////////////////////////////
  public void setClmInsdFirstNm(String strValue) {
    String _rtn = null;
    // **************************************************************************
    // Function  : Property Let ClmInsdFirstNm
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal NewValue As String
    // Returns   :
    // **************************************************************************
    "Property Let ClmInsdFirstNm"
.equals(Const cstrCurrentProc As String);
    try {

      m_strClmInsdFirstNm.value = strValue;
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public String getClmInsdLastNm() {
    String _rtn = "";
    // **************************************************************************
    // Function  : Property Get ClmInsdLastNm
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : string
    // **************************************************************************
    "Property Get ClmInsdLastNm"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = CStr(m_strClmInsdLastNm.value);
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}


//////////////////////////////////////////////////////////////////////////////////////////////////
  public void setClmInsdLastNm(String strValue) {
    String _rtn = null;
    // **************************************************************************
    // Function  : Property Let ClmInsdLastNm
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal NewValue As String
    // Returns   :
    // **************************************************************************
    "Property Let ClmInsdLastNm"
.equals(Const cstrCurrentProc As String);
    try {

      m_strClmInsdLastNm.value = strValue;
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public String getClmInsdSsnNum() {
    String _rtn = "";
    // **************************************************************************
    // Function  : Property Get ClmInsdSsnNum
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : string
    // **************************************************************************
    "Property Get ClmInsdSsnNum"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = CStr(m_strClmInsdSsnNum.value);
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}


//////////////////////////////////////////////////////////////////////////////////////////////////
  public void setClmInsdSsnNum(String strValue) {
    String _rtn = null;
    // **************************************************************************
    // Function  : Property Let ClmInsdSsnNum
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal NewValue As String
    // Returns   :
    // **************************************************************************
    "Property Let ClmInsdSsnNum"
.equals(Const cstrCurrentProc As String);
    try {

      m_strClmInsdSsnNum.value = strValue;
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public String getClmNum() {
    String _rtn = "";
    // **************************************************************************
    // Function  : Property Get ClmNum
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : string
    // **************************************************************************
    "Property Get ClmNum"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = CStr(m_strClmNum.value);
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}


//////////////////////////////////////////////////////////////////////////////////////////////////
  public void setClmNum(String strValue) {
    String _rtn = null;
    // **************************************************************************
    // Function  : Property Let ClmNum
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal NewValue As String
    // Returns   :
    // **************************************************************************
    "Property Let ClmNum"
.equals(Const cstrCurrentProc As String);
    try {

      m_strClmNum.value = strValue;
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public String getClmPolNum() {
    String _rtn = "";
    // **************************************************************************
    // Function  : Property Get ClmPolNum
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : string
    // **************************************************************************
    "Property Get ClmPolNum"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = CStr(m_strClmPolNum.value);
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}


//////////////////////////////////////////////////////////////////////////////////////////////////
  public void setClmPolNum(String strValue) {
    String _rtn = null;
    // **************************************************************************
    // Function  : Property Let ClmPolNum
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal NewValue As String
    // Returns   :
    // **************************************************************************
    "Property Let ClmPolNum"
.equals(Const cstrCurrentProc As String);
    try {

      m_strClmPolNum.value = strValue;
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public Date getClmProofDt() {
    Date _rtn = null;
    // **************************************************************************
    // Function  : Property Get ClmProofDt
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : string
    // **************************************************************************
    "Property Get ClmProofDt"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = G.parseDate(m_dteClmProofDt.value);
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}


//////////////////////////////////////////////////////////////////////////////////////////////////
  public void setClmProofDt(Date dteValue) {
    Date _rtn = null;
    // **************************************************************************
    // Function  : Property Let ClmProofDt
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal NewValue As String
    // Returns   :
    // **************************************************************************
    "Property Let ClmProofDt"
.equals(Const cstrCurrentProc As String);
    try {

      m_dteClmProofDt.value = dteValue;
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public Object getClmTotClmPdAmt() {
    Object _rtn = null;
    // **************************************************************************
    // Function  : Property Get ClmTotClmPdAmt
    // Purpose   : Retrieves current value (which could be NULL) from property
    // Parameters: N/A
    // Returns   : double
    // **************************************************************************
    "Property Get ClmTotClmPdAmt"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = m_varClmTotClmPdAmt.value;
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public void setClmTotClmPdAmt(Object varValue) {
    Object _rtn = null;
    // **************************************************************************
    // Function  : Property Let ClmTotClmPdAmt
    // Purpose   : Assigns new Value (which could be NULL) to property
    // Parameters: ByVal varValue
    // **************************************************************************
    "Property Let ClmTotClmPdAmt"
.equals(Const cstrCurrentProc As String);
    try {

      m_varClmTotClmPdAmt.value = varValue;
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public Object getClmTotDthbPmtAmt() {
    Object _rtn = null;
    // **************************************************************************
    // Function  : Property Get ClmTotDthbPmtAmt
    // Purpose   : Retrieves current value (which could be NULL) from property
    // Parameters: N/A
    // Returns   : double
    // **************************************************************************
    "Property Get ClmTotDthbPmtAmt"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = m_varClmTotDthbPmtAmt.value;
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public void setClmTotDthbPmtAmt(Object varValue) {
    Object _rtn = null;
    // **************************************************************************
    // Function  : Property Let ClmTotDthbPmtAmt
    // Purpose   : Assigns new Value (which could be NULL) to property
    // Parameters: ByVal varValue
    // **************************************************************************
    "Property Let ClmTotDthbPmtAmt"
.equals(Const cstrCurrentProc As String);
    try {

      m_varClmTotDthbPmtAmt.value = varValue;
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public Object getClmTotIntAmt() {
    Object _rtn = null;
    // **************************************************************************
    // Function  : Property Get ClmTotIntAmt
    // Purpose   : Retrieves current value (which could be NULL) from property
    // Parameters: N/A
    // Returns   : double
    // **************************************************************************
    "Property Get ClmTotIntAmt"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = m_varClmTotIntAmt.value;
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public void setClmTotIntAmt(Object varValue) {
    Object _rtn = null;
    // **************************************************************************
    // Function  : Property Let ClmTotIntAmt
    // Purpose   : Assigns new Value (which could be NULL) to property
    // Parameters: ByVal varValue
    // **************************************************************************
    "Property Let ClmTotIntAmt"
.equals(Const cstrCurrentProc As String);
    try {

      m_varClmTotIntAmt.value = varValue;
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public Object getClmTotWthldAmt() {
    Object _rtn = null;
    // **************************************************************************
    // Function  : Property Get ClmTotWthldAmt
    // Purpose   : Retrieves current value (which could be NULL) from property
    // Parameters: N/A
    // Returns   : double
    // **************************************************************************
    "Property Get ClmTotWthldAmt"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = m_varClmTotWthldAmt.value;
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public void setClmTotWthldAmt(Object varValue) {
    Object _rtn = null;
    // **************************************************************************
    // Function  : Property Let ClmTotWthldAmt
    // Purpose   : Assigns new Value (which could be NULL) to property
    // Parameters: ByVal varValue
    // **************************************************************************
    "Property Let ClmTotWthldAmt"
.equals(Const cstrCurrentProc As String);
    try {

      m_varClmTotWthldAmt.value = varValue;
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public int getCurrentLookupRecordNumber() {
    int _rtn = 0;
    //--------------------------------------------------------------------------
    // Procedure:   Property Get CurrentLookupRecordNumber
    // Description: Retrieve the record number of the record currently in context
    // Returns:     record position as Long
    //-----------------------------------------------------------------------------
    try {
      "Property Get CurrentLookupRecordNumber"
.equals(Const cstrCurrentProc As String);

      _rtn = m_rstLookup.AbsolutePosition;
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public int getDecimalPositions(String strTagIn) {
    int _rtn = 0;
    // **************************************************************************
    // Function  : Property Get DecimalPositions
    // Purpose   : Retrieves the max number of decimal positions from the
    //             named property
    // Parameters: N/A
    // Returns   : Integer
    // **************************************************************************
    "Property Get DecimalPositions"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = fnGetProperty(strTagIn).decimalPositions;
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}


//////////////////////////////////////////////////////////////////////////////////////////////////
  public Object getDefaultValue(String strTagIn) {
    Object _rtn = null;
    // **************************************************************************
    // Function  : Property Get DefaultValue
    // Purpose   : Retrieves the Default Value from the
    //             named property
    //
    //             It's up to the CALLER to see if DefaultValue = Empty and,
    //             if so, not to use the return value. It's also the
    //             caller's responsibility to do any data type conversion
    //             that might be necessary.
    //
    // Parameters: N/A
    // Returns   : Variant
    // **************************************************************************
    "Property Get DefaultValue"
.equals(Const cstrCurrentProc As String);
    try {
      udtColumn udtTemp = null;
      Object varTemp = null;

      //' for efficiency
      udtTemp = fnGetProperty(strTagIn);

      if (udtTemp.hasDefault) {
        varTemp = udtTemp.defaultValue;

        // Strip leading single quote
        if ("'".equals(varTemp.substring(0, 1)) && varTemp.length() > 1) {
          varTemp = varTemp.substring(varTemp.length() - varTemp.length() - 1);
        }
        // Strip trailing single quote
        if ("'".equals(varTemp.substring(varTemp.length() - 1)) && varTemp.length() > 1) {
          varTemp = varTemp.substring(0, varTemp.length() - 1);
        }

        _rtn = varTemp;
        // **TODO:** goto found: GoTo PROC_EXIT;
      } 
      else {
        _rtn = Empty;
      }
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public int getDollarPositions(String strTagIn) {
    int _rtn = 0;
    // **************************************************************************
    // Function  : Property Get DollarPositions
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : Integer
    // **************************************************************************
    "Property Get DollarPositions"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = fnGetProperty(strTagIn).dollarPositions;
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public String getFormat(String strTagIn) {
    String _rtn = "";
    // **************************************************************************
    // Function  : Property Get Format
    // Purpose   : Retrieves the default Format Mask (i.e. #####.###) from property
    // Parameters: N/A
    // Returns   : String
    // **************************************************************************
    "Property Get Format"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = fnGetProperty(strTagIn).format;
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public String getInsdDthResStCd() {
    String _rtn = "";
    // **************************************************************************
    // Function  : Property Get InsdDthResStCd
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : string
    // **************************************************************************
    "Property Get InsdDthResStCd"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = CStr(m_strInsdDthResStCd.value);
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}


//////////////////////////////////////////////////////////////////////////////////////////////////
  public void setInsdDthResStCd(String strValue) {
    String _rtn = null;
    // **************************************************************************
    // Function  : Property Let InsdDthResStCd
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal NewValue As String
    // Returns   :
    // **************************************************************************
    "Property Let InsdDthResStCd"
.equals(Const cstrCurrentProc As String);
    try {

      m_strInsdDthResStCd.value = strValue;
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public boolean getIsKey(String strTagIn) {
    boolean _rtn = false;
    // **************************************************************************
    // Function  : Property Get IsKey
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : Int
    // **************************************************************************
    "Property Get IsKey"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = fnGetProperty(strTagIn).isKey;
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public boolean getIsNullable(String strTagIn) {
    boolean _rtn = false;
    // **************************************************************************
    // Function  : Property Get IsNullable
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : Int
    // **************************************************************************
    "Property Get IsNullable"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = fnGetProperty(strTagIn).isNullable;
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public String getIssStCd() {
    String _rtn = "";
    // **************************************************************************
    // Function  : Property Get IssStCd
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : string
    // **************************************************************************
    "Property Get IssStCd"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = CStr(m_strIssStCd.value);
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}


//////////////////////////////////////////////////////////////////////////////////////////////////
  public void setIssStCd(String strValue) {
    String _rtn = null;
    // **************************************************************************
    // Function  : Property Let IssStCd
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal NewValue As String
    // Returns   :
    // **************************************************************************
    "Property Let IssStCd"
.equals(Const cstrCurrentProc As String);
    try {

      m_strIssStCd.value = strValue;
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public DBRecordSet getLookupData() {
    //--------------------------------------------------------------------------
    // Procedure:   Get_LookupData
    // Description: Get a copy of the objects Lookup Recordset
    // Returns:     ADODB.Recordset
    //-----------------------------------------------------------------------------
    try {
      "Get_LookupData"
.equals(Const cstrCurrentProc As String);

      return m_rstLookup;
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public Object getLookupData_Claim() {
    Object _rtn = null;
    //--------------------------------------------------------------------------
    // Procedure:   Get LookupData_Claim
    // Description: Return an array containing just the desired columns
    //              that should be populated in the Lookup fpCombo control
    // Returns:     variant array
    //-----------------------------------------------------------------------------
    try {
      "Get_LookupData_Claim"
.equals(Const cstrCurrentProc As String);
      Object[] aRows() = null;
      String strOriginalClmNum = "";

      // The .GetRows method changes the positioning within the m_rstLookup, which
      // causes the "record x of y" label to incorrectly get set to the
      // .RecordCount value. So, save the key of the current record, issue the
      // .GetRows and then navigate back to the original position within the rst.

      //!TODO! 06/28/03 BAW - The following code needs to be tightened up.
      // If the m_rstLookup recordset is empty, the first line within the WITH
      // block gets runtime error 3021 "Either BOF or EOF is True, or the current
      // record has been deleted. Requested operation requires a current record."
      // so the logic should be improved to return an initialized array (or one
      // with a pseudo blank row) if the m_rstLookup is at BOF or EOF
      // so the caller (fnLoadLpcLookup) will work.
      strOriginalClmNum = m_rstLookup.Fields(MCSTRCLMNUM).value;

      // m_rstLookup is already sorted in the order we want: clm_num
      aRows = m_rstLookup.GetRows(Rows:=adGetRowsRest, Start:=adBookmarkFirst);

      getRelativeRecord(strOriginalClmNum, enumPositionDirection.ePDSAMERECORD);

      _rtn = aRows;
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    Erase(aRows);

    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public Object getLookupData_Name() {
    Object _rtn = null;
    //--------------------------------------------------------------------------
    // Procedure:   Get LookupData_Name
    // Description: Retrieves a sorted ADO Recordset, loads it to an array
    //              containing for fast population to the Name Lookup
    //              fpCombo control on the Insured screen
    // Returns:     variant array
    //-----------------------------------------------------------------------------
    try {
      "Get_LookupData_Name"
.equals(Const cstrCurrentProc As String);
      "dbo.proc_claim_lu_select2"
.equals(Const cstrSproc As String);
      Object[] aRows() = null;
      DBRecordSet rstTemp = null;
      ADODB.Parameter prmReturnValue = null;

      if (!(m_adwADO.commandSetSproc(cstrSproc))) {
        // **TODO:** goto found: GoTo PROC_EXIT;
      }

      ADODB.Command w_aDOCommand = m_adwADO.getADOCommand();
      // ---Parameter #1---
      // Define the return value that represents the error code (i.e. reason) why
      // the stored procedure failed.
      prmReturnValue = w_aDOCommand.CreateParameter(Name:="@return_value", Type:=adInteger, Direction:=adParamReturnValue, .value:=Null);
      w_aDOCommand.Parameters.Append(prmReturnValue);

      rstTemp = w_aDOCommand.Execute();

      aRows = rstTemp.GetRows(Rows:=adGetRowsRest, Start:=adBookmarkFirst);

      _rtn = aRows;
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    Erase(aRows);
    modGeneral.fnFreeObject(rstTemp);
    modGeneral.fnFreeObject(prmReturnValue);

    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public Object getLookupData_SSN() {
    Object _rtn = null;
    //--------------------------------------------------------------------------
    // Procedure:   Get LookupData_SSN
    // Description: Retrieves a sorted ADO Recordset, loads it to an array
    //              containing for fast population to the SSN Lookup
    //              fpCombo control on the Insured screen
    // Returns:     variant array
    //-----------------------------------------------------------------------------
    try {
      "Get_LookupData_SSN"
.equals(Const cstrCurrentProc As String);
      "dbo.proc_claim_lu_select3"
.equals(Const cstrSproc As String);
      Object[] aRows() = null;
      DBRecordSet rstTemp = null;
      ADODB.Parameter prmReturnValue = null;

      if (!(m_adwADO.commandSetSproc(cstrSproc))) {
        // **TODO:** goto found: GoTo PROC_EXIT;
      }

      ADODB.Command w_aDOCommand = m_adwADO.getADOCommand();
      // ---Parameter #1---
      // Define the return value that represents the error code (i.e. reason) why
      // the stored procedure failed.
      prmReturnValue = w_aDOCommand.CreateParameter(Name:="@return_value", Type:=adInteger, Direction:=adParamReturnValue, .value:=Null);
      w_aDOCommand.Parameters.Append(prmReturnValue);

      rstTemp = w_aDOCommand.Execute();

      aRows = rstTemp.GetRows(Rows:=adGetRowsRest, Start:=adBookmarkFirst);

      _rtn = aRows;
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    Erase(aRows);
    modGeneral.fnFreeObject(rstTemp);
    modGeneral.fnFreeObject(prmReturnValue);

    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public boolean getLookupIsAtBOF() {
    boolean _rtn = false;
    //--------------------------------------------------------------------------
    // Procedure:   Property Get LookupIsAtBOF
    // Description: Indicates whether the Lookup recordset is at BOF
    //              (i.e. prior to the last record in the recordset).
    //
    //              Both LookupIsAtBOF() and LookupIsAtEOF() will return True
    //              if there are no records in the m_rstLookup recordset.
    //
    // Returns:     True if it is at BOF; False otherwise
    //-----------------------------------------------------------------------------
    try {
      "Property Get LookupIsAtBOF"
.equals(Const cstrCurrentProc As String);

      if (!(m_rstLookup == null)) {
        _rtn = m_rstLookup.BOF;
      } 
      else {
        _rtn = true;
      }
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public boolean getLookupIsAtEOF() {
    boolean _rtn = false;
    //--------------------------------------------------------------------------
    // Procedure:   Property Get LookupIsAtEOF
    // Description: Indicates whether the Lookup recordset is at EOF
    //              (i.e. beyond the last record in the recordset)
    //
    //              Both LookupIsAtBOF() and LookupIsAtEOF() will return True
    //              if there are no records in the m_rstLookup recordset.
    //
    // Returns:     True if it is at EOF; False otherwise
    //-----------------------------------------------------------------------------
    try {
      "Property Get LookupIsAtEOF"
.equals(Const cstrCurrentProc As String);

      if (!(m_rstLookup == null)) {
        _rtn = m_rstLookup.EOF;
      } 
      else {
        _rtn = true;
      }
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public int getLookupRecordCount() {
    int _rtn = 0;
    // **************************************************************************
    // Function  : Property Get LookupRecordCount
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : string
    // **************************************************************************
    "Property Get LookupRecordCount"
.equals(Const cstrCurrentProc As String);
    try {

      if (!(m_rstLookup == null)) {
        _rtn = m_rstLookup.RecordCount;
      } 
      else {
        _rtn = 0;
      }
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public Date getLstUpdtDtm() {
    Date _rtn = null;
    // **************************************************************************
    // Function  : Property Get LstUpdtDtm
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : Int
    // **************************************************************************
    "Property Get LstUpdtDtm"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = G.parseDate(m_dteLstUpdtDtm.value);
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public void setLstUpdtDtm(Date dteValue) {
    Date _rtn = null;
    // **************************************************************************
    // Function  : Property Let LstUpdtDtm
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal NewValue As Integer
    // Returns   :
    // **************************************************************************
    "Property Let LstUpdtDtm"
.equals(Const cstrCurrentProc As String);
    try {

      m_dteLstUpdtDtm.value = dteValue;
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public String getLstUpdtUserId() {
    String _rtn = "";
    // **************************************************************************
    // Function  : Property Get LstUpdtUserId
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : Int
    // **************************************************************************
    "Property Get LstUpdtUserId"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = CStr(m_strLstUpdtUserId.value);
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public void setLstUpdtUserId(String strValue) {
    String _rtn = null;
    // **************************************************************************
    // Function  : Property Let LstUpdtUserId
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal NewValue As Integer
    // Returns   :
    // **************************************************************************
    "Property Let LstUpdtUserId"
.equals(Const cstrCurrentProc As String);
    try {

      m_strLstUpdtUserId.value = strValue;
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public String getMask(String strTagIn) {
    String _rtn = "";
    // **************************************************************************
    // Function  : Property Get Mask
    // Purpose   : Retrieves the default Mask (i.e. #####.###) from property
    // Parameters: N/A
    // Returns   : String
    // **************************************************************************
    "Property Get Mask"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = fnGetProperty(strTagIn).mask;
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public int getMaxCharacters(String strTagIn) {
    int _rtn = 0;
    // **************************************************************************
    // Function  : Property Get MaxCharacters
    // Purpose   : Retrieves the number of allowable characters from property
    // Parameters: N/A
    // Returns   : Long
    // **************************************************************************
    "Property Get MaxCharacters"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = fnGetProperty(strTagIn).maxCharacters;
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}


//////////////////////////////////////////////////////////////////////////////////////////////////
  public String getPycoTypCd() {
    String _rtn = "";
    // **************************************************************************
    // Function  : Property Get PycoTypCd
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : string
    // **************************************************************************
    "Property Get PycoTypCd"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = CStr(m_strPycoTypCd.value);
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}


//////////////////////////////////////////////////////////////////////////////////////////////////
  public void setPycoTypCd(String strValue) {
    String _rtn = null;
    // **************************************************************************
    // Function  : Property Let PycoTypCd
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal NewValue As String
    // Returns   :
    // **************************************************************************
    "Property Let PycoTypCd"
.equals(Const cstrCurrentProc As String);
    try {

      m_strPycoTypCd.value = strValue;
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public boolean getShouldForceToUppercase(String strTagIn) {
    boolean _rtn = false;
    // **************************************************************************
    // Function  : Property Get ShouldForceToUppercase
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : Int
    // **************************************************************************
    "Property Get ShouldForceToUppercase"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = fnGetProperty(strTagIn).shouldForceToUppercase;
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}





///\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
//\                                                                  /
//|                        PUBLIC  Procedures                        |
///                                                                  \
//\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/

//////////////////////////////////////////////////////////////////////////////////////////////////
  public boolean addRecord() {
    boolean _rtn = false;
    //--------------------------------------------------------------------------
    // Procedure:   AddRecord
    // Description: Adds a single record based on key value
    //              selection.
    // Returns:     boolean
    // Params:      Not necessary, they will be derived from properties the form
    //              should have already set
    // Date:        04/11/2002
    //-----------------------------------------------------------------------------
    // Modified:     Berry Kropiwka - Added Compact Calc - 2019-09-27

    //!CUSTOMIZE!  Customize the name of the stored procedure, the number and names
    //             of the parameters and, perhaps, the return values trapped in the
    //             error handler.

    "AddRecord"
.equals(Const cstrCurrentProc As String);
    //' Stored procedure to execute
    "dbo.proc_claim_insert"
.equals(Const cstrSproc As String);
    DBRecordSet rstTemp = null;
    boolean bSuccessful = false;
    ADODB.Parameter prmReturnValue = null;
    ADODB.Parameter prmAdmnSystCd = null;
    ADODB.Parameter prmClmInsdDthDt = null;
    ADODB.Parameter prmClmInsdFirstNm = null;
    ADODB.Parameter prmClmInsdLastNm = null;
    ADODB.Parameter prmClmInsdSsnNum = null;
    ADODB.Parameter prmClmNum = null;
    ADODB.Parameter prmClmPolNum = null;
    ADODB.Parameter prmClmProofDt = null;
    ADODB.Parameter prmClmTotClmPdAmt = null;
    ADODB.Parameter prmClmTotDthbPmtAmt = null;
    ADODB.Parameter prmClmTotIntAmt = null;
    ADODB.Parameter prmClmTotWthldAmt = null;
    ADODB.Parameter prmInsdDthResStCd = null;
    ADODB.Parameter prmIssStCd = null;
    ADODB.Parameter prmPycoTypCd = null;
    ADODB.Parameter prmClmForResDthInd = null;
    ADODB.Parameter prmClmCompactClcnInd = null;
    ADODB.Parameter prmInvalid_Key = null;
    ADODB.Parameter prmNew_Id = null;

    try {

      //...........................................................................
      // No need to check to see if another user updated or deleted this record
      // since we're doing an Add.
      //...........................................................................

      if (!(m_adwADO.commandSetSproc(cstrSproc))) {
        // **TODO:** goto found: GoTo PROC_EXIT;
      }

      // For Char/VarChar fields,
      //     * Use fnNullIfZLS to ensure Nulls are appropriately handled.
      //     * Do *not* set the optional 2nd parameter to fnNullIfZLS to True.
      // For numeric fields,
      //     * Use fnNullIfZero to ensure Nulls are appropriately handled.
      // For Y/N fields,
      //     * Use fnBoolToYN to ensure True/False is appropriately translated.

      ADODB.Command w_aDOCommand = m_adwADO.getADOCommand();
      // ---Parameter #1---
      // Define the return value that represents the error code (i.e. reason) why
      // the stored procedure failed.
      prmReturnValue = w_aDOCommand.CreateParameter(Name:="@return_value", Type:=adInteger, Direction:=adParamReturnValue, .value:=Null);
      w_aDOCommand.Parameters.Append(prmReturnValue);

      // ---Parameter #2---
      // Define the ADMN_SYST_CD parameter
      prmAdmnSystCd = w_aDOCommand.CreateParameter(Name:="@admn_syst_cd", Type:=adChar, Direction:=adParamInput, Size:=2, .value:=fnNullIfZLS(varIn:=AdmnSystCd, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmAdmnSystCd);

      // ---Parameter #3---
      // Define the CLM_INSD_DTH_DT parameter
      prmClmInsdDthDt = w_aDOCommand.CreateParameter(Name:="@clm_insd_dth_dt", Type:=adDBTimeStamp, Direction:=adParamInput, Size:=16, .value:=ClmInsdDthDt);
      w_aDOCommand.Parameters.Append(prmClmInsdDthDt);

      // ---Parameter #4---
      // Define the CLM_INSD_FIRST_NM parameter
      prmClmInsdFirstNm = w_aDOCommand.CreateParameter(Name:="@clm_insd_first_nm", Type:=adVarChar, Direction:=adParamInput, Size:=50, .value:=fnNullIfZLS(varIn:=ClmInsdFirstNm, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmClmInsdFirstNm);

      // ---Parameter #5---
      // Define the CLM_INSD_LAST_NM parameter
      prmClmInsdLastNm = w_aDOCommand.CreateParameter(Name:="@clm_insd_last_nm", Type:=adVarChar, Direction:=adParamInput, Size:=50, .value:=fnNullIfZLS(varIn:=ClmInsdLastNm, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmClmInsdLastNm);

      // ---Parameter #6---
      // Define the CLM_INSD_SSN_NUM parameter
      prmClmInsdSsnNum = w_aDOCommand.CreateParameter(Name:="@clm_insd_ssn_num", Type:=adChar, Direction:=adParamInput, Size:=9, .value:=fnNullIfZLS(varIn:=ClmInsdSsnNum, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmClmInsdSsnNum);

      // ---Parameter #7---
      // Define the CLM_NUM parameter
      prmClmNum = w_aDOCommand.CreateParameter(Name:="@clm_num", Type:=adVarChar, Direction:=adParamInput, Size:=20, .value:=fnNullIfZLS(varIn:=ClmNum, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmClmNum);

      // ---Parameter #8---
      // Define the CLM_POL_NUM parameter
      prmClmPolNum = w_aDOCommand.CreateParameter(Name:="@clm_pol_num", Type:=adChar, Direction:=adParamInput, Size:=15, .value:=fnNullIfZLS(varIn:=ClmPolNum, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmClmPolNum);

      // ---Parameter #9---
      // Define the CLM_PROOF_DT parameter
      prmClmProofDt = w_aDOCommand.CreateParameter(Name:="@clm_proof_dt", Type:=adDBTimeStamp, Direction:=adParamInput, Size:=16, .value:=ClmProofDt);
      w_aDOCommand.Parameters.Append(prmClmProofDt);

      // ---Parameter #10---
      // Define the CLM_TOT_CLM_PD_AMT parameter (nullable)
      prmClmTotClmPdAmt = w_aDOCommand.CreateParameter(Name:="@clm_tot_clm_pd_amt", Type:=adNumeric, Direction:=adParamInput, .value:=ClmTotClmPdAmt);
      prmClmTotClmPdAmt.Precision = m_varClmTotClmPdAmt.precision;
      prmClmTotClmPdAmt.NumericScale = m_varClmTotClmPdAmt.numericScale;
      w_aDOCommand.Parameters.Append(prmClmTotClmPdAmt);

      // ---Parameter #11---
      // Define the CLM_TOT_DTHB_PMT_AMT parameter (nullable)
      prmClmTotDthbPmtAmt = w_aDOCommand.CreateParameter(Name:="@clm_tot_dthb_pmt_amt", Type:=adNumeric, Direction:=adParamInput, .value:=ClmTotDthbPmtAmt);
      prmClmTotDthbPmtAmt.Precision = m_varClmTotDthbPmtAmt.precision;
      prmClmTotDthbPmtAmt.NumericScale = m_varClmTotDthbPmtAmt.numericScale;
      w_aDOCommand.Parameters.Append(prmClmTotDthbPmtAmt);

      // ---Parameter #12---
      // Define the CLM_TOT_INT_AMT parameter (nullable)
      prmClmTotIntAmt = w_aDOCommand.CreateParameter(Name:="@clm_tot_int_amt", Type:=adNumeric, Direction:=adParamInput, .value:=ClmTotIntAmt);
      prmClmTotIntAmt.Precision = m_varClmTotIntAmt.precision;
      prmClmTotIntAmt.NumericScale = m_varClmTotIntAmt.numericScale;
      w_aDOCommand.Parameters.Append(prmClmTotIntAmt);

      // ---Parameter #13---
      // Define the CLM_TOT_WTHLD_AMT parameter (nullable)
      prmClmTotWthldAmt = w_aDOCommand.CreateParameter(Name:="@clm_tot_wthld_amt", Type:=adNumeric, Direction:=adParamInput, .value:=ClmTotWthldAmt);
      prmClmTotWthldAmt.Precision = m_varClmTotWthldAmt.precision;
      prmClmTotWthldAmt.NumericScale = m_varClmTotWthldAmt.numericScale;
      w_aDOCommand.Parameters.Append(prmClmTotWthldAmt);

      // ---Parameter #14---
      // Define the INSD_DTH_RES_ST_CD parameter
      prmInsdDthResStCd = w_aDOCommand.CreateParameter(Name:="@insd_dth_res_st_cd", Type:=adChar, Direction:=adParamInput, Size:=2, .value:=fnNullIfZLS(varIn:=InsdDthResStCd, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmInsdDthResStCd);

      // ---Parameter #15---
      // Define the ISS_ST_CD parameter
      prmIssStCd = w_aDOCommand.CreateParameter(Name:="@iss_st_cd", Type:=adChar, Direction:=adParamInput, Size:=2, .value:=fnNullIfZLS(varIn:=IssStCd, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmIssStCd);

      // ---Parameter #16---
      // Define the PYCO_TYP_CD parameter
      prmPycoTypCd = w_aDOCommand.CreateParameter(Name:="@pyco_typ_cd", Type:=adChar, Direction:=adParamInput, Size:=8, .value:=fnNullIfZLS(varIn:=PycoTypCd, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmPycoTypCd);

      // ---Parameter #17---
      // Define the clm_for_res_dth_ind parameter
      prmClmForResDthInd = w_aDOCommand.CreateParameter(Name:="@clm_for_res_dth_ind", Type:=adChar, Direction:=adParamInput, Size:=1, .value:=fnBoolToYN(bIn:=ClmForResDthInd));
      w_aDOCommand.Parameters.Append(prmClmForResDthInd);

      // ---Parameter #18---
      // Define the clm_compact_clcn_ind parameter
      prmClmCompactClcnInd = w_aDOCommand.CreateParameter(Name:="@clm_compact_clcn_ind", Type:=adChar, Direction:=adParamInput, Size:=1, .value:=fnBoolToTF(bIn:=ClmCompactClcnInd));
      w_aDOCommand.Parameters.Append(prmClmCompactClcnInd);

      // ---Parameter #19---
      // Define the Invalid_Key output parameter, which reflects *which* foreign
      // key violation was encountered.
      prmInvalid_Key = w_aDOCommand.CreateParameter(Name:="@Invalid_Key", Type:=adVarChar, Size:=255, Direction:=adParamOutput);
      w_aDOCommand.Parameters.Append(prmInvalid_Key);

      // ---Parameter #20---
      // Define the output parameter that represents the CLM_ID value that
      // was assigned to the record we're trying to insert
      prmNew_Id = w_aDOCommand.CreateParameter(Name:="@New_Id", Type:=adInteger, Direction:=adParamOutput, .value:=Null);
      w_aDOCommand.Parameters.Append(prmNew_Id);

      // Do the Add
      w_aDOCommand.Execute;

      //...........................................................................
      // Refresh the Lookup recordset, re-retrieve the just-added record so that
      // record is *still* the current record, and load its data to the
      // table wrapper's class properties so all table columns (including
      // those set by the DBMS like identity and Last Updated columns) are
      // up-to-date.
      //...........................................................................
      bSuccessful = getRelativeRecord(getClmNumFromClmID(prmNew_Id.value), enumPositionDirection.ePDSAMERECORD);

      _rtn = true;
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    modGeneral.fnFreeRecordset(rstTemp);
    modGeneral.fnFreeObject(prmReturnValue);
    modGeneral.fnFreeObject(prmAdmnSystCd);
    modGeneral.fnFreeObject(prmClmInsdDthDt);
    modGeneral.fnFreeObject(prmClmInsdFirstNm);
    modGeneral.fnFreeObject(prmClmInsdLastNm);
    modGeneral.fnFreeObject(prmClmInsdSsnNum);
    modGeneral.fnFreeObject(prmClmNum);
    modGeneral.fnFreeObject(prmClmPolNum);
    modGeneral.fnFreeObject(prmClmProofDt);
    modGeneral.fnFreeObject(prmClmTotClmPdAmt);
    modGeneral.fnFreeObject(prmClmTotDthbPmtAmt);
    modGeneral.fnFreeObject(prmClmTotIntAmt);
    modGeneral.fnFreeObject(prmClmTotWthldAmt);
    modGeneral.fnFreeObject(prmInsdDthResStCd);
    modGeneral.fnFreeObject(prmIssStCd);
    modGeneral.fnFreeObject(prmPycoTypCd);
    modGeneral.fnFreeObject(prmClmForResDthInd);
    modGeneral.fnFreeObject(prmClmCompactClcnInd);
    modGeneral.fnFreeObject(prmInvalid_Key);
    modGeneral.fnFreeObject(prmNew_Id);

    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (prmReturnValue) {
      case  modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO:
        // 4028 = An error occurred while attempting to @@1 this record.
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO, MCSTRNAME+ cstrCurrentProc, "add");
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
        break;

      case  modResConstants.gCRES_NERR_ADD_WITH_NONUNIQUE_KEY:
        // 4031 = A record with the specified key (@@1) already exists. Please specify a unique key.
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_ADD_WITH_NONUNIQUE_KEY, MCSTRNAME+ cstrCurrentProc, "Claim Number "+ RTrim$(getClmNum()));
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
        break;

      case  modResConstants.gCRES_NERR_KEY_NOT_FOUND:
        // 4032 = The @@1 specified (@@2) is no longer defined in the @@3 table. Please choose a different value.
        if ("ADMN_SYST_CD"
.equals(prmInvalid_Key.toUpperCase())) {
          modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_KEY_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, "Admin System", RTrim$(getAdmnSystCd()), "ADMIN_SYSTEM_T");
          /**TODO:** resume found: Resume(PROC_EXIT)*/;
        } 
        else if ("INSD_DTH_RES_ST_CD"
.equals(prmInvalid_Key.toUpperCase())) {
          modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_KEY_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, "Residence State", RTrim$(getInsdDthResStCd()), "STATE_T");
          /**TODO:** resume found: Resume(PROC_EXIT)*/;
        } 
        else if ("ISS_ST_CD"
.equals(prmInvalid_Key.toUpperCase())) {
          modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_KEY_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, "Issue State", RTrim$(getIssStCd()), "STATE_T");
          /**TODO:** resume found: Resume(PROC_EXIT)*/;
        } 
        else {
          // PYCO_TYP_CD
          modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_KEY_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, "Company Type", RTrim$(getPycoTypCd()), "PAYOR_COMPANY_TYPE_T");
          /**TODO:** resume found: Resume(PROC_EXIT)*/;
        }
        break;
    }

    // If any other errors exist, i.e. in Err object, then let it fall through into default error handling.

    switch (VBA.ex.Number) {
      //' Object not found
      case  -2147217900:
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_FERR_SPROC_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, cstrSproc);
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
        break;

      default:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public int checkForAnotherUsersChanges(enumWhatOperationIsBeingAttempted lngWhatOperation, String strACF2) { // TODO: Use of ByRef founded Public Function CheckForAnotherUsersChanges(ByVal lngWhatOperation As enumWhatOperationIsBeingAttempted, ByRef strACF2 As String) As Long
    int _rtn = 0;
    //--------------------------------------------------------------------------
    // Procedure:   CheckForAnotherUsersChanges
    // Description: Check to see if another user has altered or deleted
    //              the record that is about to be operated upon. This is called
    //              (directly or indirectly) by each form's cmdDelete_Click and
    //              cmdUpdate_Click event handlers.
    //
    //              NOTE: The caller must check for every possible return value
    //                    that a given lngWhatOperation value could hit!
    //
    // Returns:     A return code indicating what has occured, so the form
    //              can determine what to do about it. A zero value means the form
    //              doesn't have to do anything.
    //
    // Params:
    //    lngWhatOperation (in) - indicates whether an Update or Delete is being attempted
    //    strACF2 (out)         - for some errors, reflects the ACF2 id of the user who updated
    //                            the record
    //
    // Date:        04/27/2002
    //-----------------------------------------------------------------------------
    "CheckForAnotherUsersChanges"
.equals(Const cstrCurrentProc As String);
    // rstSingleRecord_Fresh contains all columns of a single row in this class' underlying table. It reflects
    // the now-current contents of the record that is about to be updated or deleted, so the CheckForAnotherUsersChanges
    // process can determine if another user updated the record since it was originally retrieved.
    DBRecordSet rstSingleRecord_Fresh = null;

    try {

      //...........................................................................
      // See if another user deleted or updated the record since we last retrieved it...
      //...........................................................................

      // The following statement will raise a 4027 if the specified record isn't found. PROC_ERR does
      // a Resume Next so the first validation (to see if another user deleted the record) needs
      // to check for both .RecordCount=0 --or-- rst=Nothing; otherwise a runtime error 91
      // (Object variable or With block not set) is raised.
      //!CUSTOMIZE! fnSelectRecord call should pass the key column(s)
      rstSingleRecord_Fresh = fnSelectRecord(getClmId());

      // Disconnect the recordset so we can edit the data, if desired, for testing purposes
      modGeneral.fnFreeRecordset(rstSingleRecord_Fresh.ActiveConnection);

      if ((rstSingleRecord_Fresh == null) || (rstSingleRecord_Fresh.RecordCount == 0)) {
        switch (lngWhatOperation) {
          case  enumWhatOperationIsBeingAttempted.eWOUPDATE:
            _rtn = vbObjectError + modResConstants.gCRES_INFO_ANOTHER_USER_DELETED_NOT_SAVED;
          //' ewoDelete
            break;

          default:
            _rtn = vbObjectError + modResConstants.gCRES_INFO_ANOTHER_USER_DELETED;
            break;
        }
        // **TODO:** goto found: GoTo PROC_EXIT;
      }

      // Note: A "<>" comparison on the date/time values reports a false positive. Use DateDiff( ) instead.
      // Convert dates to string using CStr( ) to avoid ADO's millisecond rounding which could result in a false positive.
      if (lngWhatOperation == enumWhatOperationIsBeingAttempted.eWOUPDATE) {
        if ((DateDiff("s", CStr(m_dteLstUpdtDtm_Original), CStr(!lst_updt_dtm)) != 0) || (!lst_updt_user_id != m_strLstUpdtUserId_Original)) {
          strACF2 = !lst_updt_user_id;
          _rtn = vbObjectError + modResConstants.gCRES_INFO_ANOTHER_USER_UPDATED_DISCARDED;
        }
      }
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    modGeneral.fnFreeRecordset(rstSingleRecord_Fresh);

    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
      //' Object variable or With block variable not set
      case  91    :
        // This error will be encountered if the call to .SelectRecord didn't find the specified
        // record, e.g., another user deleted it. Ignore it so the logic that generates the
        // desired "transformed" error code will be hit.
        /**TODO:** resume found: Resume(Next)*/;
      //' 4027
        break;

      case  vbObjectError + modResConstants.gCRES_NERR_REC_NOT_FOUND  :
        // If we got record not found from the call to SelectRecord(), then wipe out traces
        // of that error and do a Resume Next. This will allow this proc to
        // transform *that* error into the one we really want:
        // gcRES_INFO_ANOTHER_USER_DELETED_NOT_SAVED vs. gcRES_INFO_ANOTHER_USER_DELETED
        modGeneral.gerhApp.clear();
        /**TODO:** resume found: Resume(Next)*/;
        break;

      default:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public boolean deleteRecord() {
    boolean _rtn = false;
    //--------------------------------------------------------------------------
    // Procedure:   DeleteRecord
    // Description: Deletes a single record based on the value(s) in the
    //              properties that correspond to the table's key(s)
    // Returns:     True if successful, False otherwise
    //-----------------------------------------------------------------------------

    //!CUSTOMIZE!  Customize the name of the stored procedure, the number and names
    //             of the parameters and, perhaps, the return values trapped in the
    //             error handler.

    "DeleteRecord"
.equals(Const cstrCurrentProc As String);
    //' Stored procedure to execute
    "dbo.proc_claim_delete"
.equals(Const cstrSproc As String);
    boolean bSuccessful = false;
    cadwADOWrapper adwTemp = null;
    DBRecordSet rstTemp = null;
    ADODB.Parameter prmReturnValue = null;
    ADODB.Parameter prmClmId = null;
    ADODB.Parameter prmDependent_Table = null;
    String strClmNum = "";

    try {

      adwTemp = new cadwADOWrapper();
      adwTemp.commandSetSproc(strSprocName:=cstrSproc);

      // Save the Claim Number associated with the Claim that's going to be deleted.
      // We'll need it afterwards to position to the previous record.
      strClmNum = getClmNumFromClmID(getClmId());

      // For Char/VarChar fields,
      //     * Use fnNullIfZLS to ensure Nulls are appropriately handled.
      //     * Do *not* set the optional 2nd parameter to fnNullIfZLS to True.
      // For numeric fields,
      //     * Use fnNullIfZero to ensure Nulls are appropriately handled.
      // For Y/N fields,
      //     * Use fnBoolToYN to ensure True/False is appropriately translated.

      ADODB.Command w_aDOCommand = adwTemp.getADOCommand();
      // ---Parameter #1---
      // Define the return value that represents the error code (i.e. reason) why
      // the stored procedure failed.
      prmReturnValue = w_aDOCommand.CreateParameter(Name:="@return_value", Type:=adInteger, Direction:=adParamReturnValue, .value:=Null);
      w_aDOCommand.Parameters.Append(prmReturnValue);

      // ---Parameter #2---
      // Define the CLM_ID parameter
      prmClmId = w_aDOCommand.CreateParameter(Name:="@clm_id", Type:=adInteger, Direction:=adParamInput, .value:=ClmId);
      w_aDOCommand.Parameters.Append(prmClmId);

      // ---Parameter #3---
      // Define the Dependent_Table output parameter, which reflects *which* dependent table
      // contains rows with a foreign key equal to the key being deleted.
      prmDependent_Table = w_aDOCommand.CreateParameter(Name:="@Dependent_Table", Type:=adVarChar, Size:=255, Direction:=adParamOutput);
      w_aDOCommand.Parameters.Append(prmDependent_Table);

      // Do the Delete
      w_aDOCommand.Execute;

      //...........................................................................
      // Refresh the Lookup recordset, reposition the Lookup data so the record
      // prior to the one just deleted is now the current record. Load that
      // record's data to the table wrapper's class properties.
      //...........................................................................
      bSuccessful = getRelativeRecord(strClmNum, enumPositionDirection.ePDPREVIOUSRECORD);

      _rtn = true;
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    modGeneral.fnFreeRecordset(rstTemp);
    modGeneral.fnFreeObject(adwTemp);
    modGeneral.fnFreeObject(prmReturnValue);
    modGeneral.fnFreeObject(prmClmId);
    modGeneral.fnFreeObject(prmDependent_Table);

    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (prmReturnValue) {
      case  modResConstants.gCRES_NERR_REC_NOT_FOUND:
        // 4027 = The specified record was not found in the database (@@1).
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_REC_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, "Claim ID "+ RTrim$(getClmId()));
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
        break;

      case  modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO:
        // 4028 = An error occurred while attempting to @@1 this record.
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO, MCSTRNAME+ cstrCurrentProc, "delete");
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
        break;

      case  modResConstants.gCRES_NERR_DEPENDENT_RECS_EXIST:
        // 4029 = This @@1 is associated with one or more records on the @@2 table and cannot be deleted until those records themselves are deleted.
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_DEPENDENT_RECS_EXIST, MCSTRNAME+ cstrCurrentProc, "Claim ID ("+ RTrim$(getClmId())+ ")", prmDependent_Table.toUpperCase());
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
        break;
    }

    // If any other errors exist, i.e. in Err object, then let it fall through into default error handling.

    switch (VBA.ex.Number) {
      //' Object not found
      case  -2147217900:
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_FERR_SPROC_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, cstrSproc);
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
        break;

      default:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public Object getClmIdFromClmNum(String strClmNum) {
    Object _rtn = null;
    //--------------------------------------------------------------------------
    // Procedure:   GetClmIdFromClmNum
    // Description: Query database for the Claim ID for a specified Claim Number
    // Params:
    //               strClmNum  (in)  The Claim Number to translate.
    //-----------------------------------------------------------------------------
    "GetClmIdFromClmNum"
.equals(Const cstrCurrentProc As String);
    "dbo.proc_clm_id_lu_select"
.equals(Const cstrSproc As String);
    DBRecordSet rstTemp = null;
    ADODB.Parameter prmReturnValue = null;
    ADODB.Parameter prmClmNum = null;
    ADODB.Parameter prmClmId = null;

    try {

      if (!(m_adwADO.commandSetSproc(cstrSproc))) {
        // **TODO:** goto found: GoTo PROC_EXIT;
      }

      ADODB.Command w_aDOCommand = m_adwADO.getADOCommand();
      // ---Parameter #1---
      // Define the return value that represents the error code (i.e. reason) why
      // the stored procedure failed.
      prmReturnValue = w_aDOCommand.CreateParameter(Name:="@return_value", Type:=adInteger, Direction:=adParamReturnValue, .value:=Null);
      w_aDOCommand.Parameters.Append(prmReturnValue);

      // ---Parameter #2---
      prmClmNum = w_aDOCommand.CreateParameter(Name:="clm_num", Type:=adVarChar, Direction:=adParamInput, Size:=20, .value:=strClmNum);
      w_aDOCommand.Parameters.Append(prmClmNum);

      // ---Parameter #2---
      prmClmId = w_aDOCommand.CreateParameter(Name:="clm_id", Type:=adInteger, Direction:=adParamInputOutput, .value:=Null);
      w_aDOCommand.Parameters.Append(prmClmId);

      rstTemp = w_aDOCommand.Execute();

      _rtn = prmClmId.value;
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
      modGeneral.fnFreeObject(prmReturnValue);
      modGeneral.fnFreeObject(prmClmNum);
      modGeneral.fnFreeObject(prmClmId);

      modGeneral.fnFreeRecordset(rstTemp);

  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public Object getClmNumFromClmID(int lngClmID) {
    Object _rtn = null;
    //--------------------------------------------------------------------------
    // Procedure:   GetClmIdFromClmNum
    // Description: Query database for the Claim ID for a specified Claim Number
    // Params:
    //               strClmID  (in)  The Claim ID to translate.
    //-----------------------------------------------------------------------------
    "GetClmNumFromClmID"
.equals(Const cstrCurrentProc As String);
    "dbo.proc_clm_num_lu_select"
.equals(Const cstrSproc As String);
    DBRecordSet rstTemp = null;
    ADODB.Parameter prmReturnValue = null;
    ADODB.Parameter prmClmNum = null;
    ADODB.Parameter prmClmId = null;

    try {

      if (!(m_adwADO.commandSetSproc(cstrSproc))) {
        // **TODO:** goto found: GoTo PROC_EXIT;
      }

      ADODB.Command w_aDOCommand = m_adwADO.getADOCommand();
      // ---Parameter #1---
      // Define the return value that represents the error code (i.e. reason) why
      // the stored procedure failed.
      prmReturnValue = w_aDOCommand.CreateParameter(Name:="@return_value", Type:=adInteger, Direction:=adParamReturnValue, .value:=Null);
      w_aDOCommand.Parameters.Append(prmReturnValue);

      // ---Parameter #2---
      prmClmId = w_aDOCommand.CreateParameter(Name:="clm_id", Type:=adInteger, Direction:=adParamInput, .value:=lngClmID);
      w_aDOCommand.Parameters.Append(prmClmId);

      // ---Parameter #3---
      prmClmNum = w_aDOCommand.CreateParameter(Name:="clm_num", Type:=adVarChar, Direction:=adParamInputOutput, Size:=20, .value:=Null);
      w_aDOCommand.Parameters.Append(prmClmNum);


      rstTemp = w_aDOCommand.Execute();

      _rtn = prmClmNum.value;
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
      modGeneral.fnFreeObject(prmReturnValue);
      modGeneral.fnFreeObject(prmClmId);
      modGeneral.fnFreeObject(prmClmNum);

      modGeneral.fnFreeRecordset(rstTemp);

  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public String getLobCdFromAdmnSystCd(String strAdmnSystCd) {
    String _rtn = "";
    //--------------------------------------------------------------------------
    // Procedure:   GetLobCdFromAdmnSystCd
    // Description: Query database for the Line-of-Business Code for a specified
    //              Admin System Code
    // Params:
    //              strAdmnSystCd  (in)  The Admin System Code to translate.
    //-----------------------------------------------------------------------------
    "GetLobCdFromAdmnSystCd"
.equals(Const cstrCurrentProc As String);
    "dbo.proc_lob_cd_lu_select"
.equals(Const cstrSproc As String);
    DBRecordSet rstTemp = null;
    ADODB.Parameter prmReturnValue = null;
    ADODB.Parameter prmAdmnSystCd = null;
    ADODB.Parameter prmLobCd = null;

    try {

      if (!(m_adwADO.commandSetSproc(cstrSproc))) {
        // **TODO:** goto found: GoTo PROC_EXIT;
      }

      ADODB.Command w_aDOCommand = m_adwADO.getADOCommand();
      // ---Parameter #1---
      // Define the return value that represents the error code (i.e. reason) why
      // the stored procedure failed.
      prmReturnValue = w_aDOCommand.CreateParameter(Name:="@return_value", Type:=adInteger, Direction:=adParamReturnValue, .value:=Null);
      w_aDOCommand.Parameters.Append(prmReturnValue);

      // ---Parameter #2---
      prmAdmnSystCd = w_aDOCommand.CreateParameter(Name:="admn_syst_cd", Type:=adChar, Direction:=adParamInput, Size:=2, .value:=strAdmnSystCd);
      w_aDOCommand.Parameters.Append(prmAdmnSystCd);

      // ---Parameter #2---
      prmLobCd = w_aDOCommand.CreateParameter(Name:="lob_cd", Type:=adChar, Direction:=adParamOutput, Size:=1, .value:=Null);
      w_aDOCommand.Parameters.Append(prmLobCd);

      rstTemp = w_aDOCommand.Execute();

      _rtn = prmLobCd.value;
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
      modGeneral.fnFreeObject(prmReturnValue);
      modGeneral.fnFreeObject(prmAdmnSystCd);
      modGeneral.fnFreeObject(prmLobCd);

      modGeneral.fnFreeRecordset(rstTemp);

  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (prmReturnValue) {
      case  modResConstants.gCRES_NERR_REC_NOT_FOUND:
        // 4027 = The specified record was not found in the database (@@1).
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_REC_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, "Claim ID "+ RTrim$(getClmId()));
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
        break;

      case  modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO:
        // 4028 = An error occurred while attempting to @@1 this record.
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO, MCSTRNAME+ cstrCurrentProc, "locate");
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
        break;
    }

    // If any other errors exist, i.e. in Err object, then let it fall through into default error handling.

    switch (VBA.ex.Number) {
      //' Object not found
      case  -2147217900:
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_FERR_SPROC_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, cstrSproc);
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
        break;

      default:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public boolean getLookupData() {
    //--------------------------------------------------------------------------
    // Procedure:   GetLookupData
    // Description: Get all rows but only particular columns
    // Returns:     True if successful; False otherwise
    // Params:      None
    // Date:        01/07/2002
    //-----------------------------------------------------------------------------

    //!CUSTOMIZE!  Customize the name of the stored procedure, the number and names
    //             of the parameters and, perhaps, the return values trapped in the
    //             error handler.

    "GetLookupData"
.equals(Const cstrCurrentProc As String);
    "dbo.proc_claim_lu_select"
.equals(Const cstrSproc As String);
    ADODB.Parameter prmReturnValue = null;

    try {

      if (!(m_adwADO.commandSetSproc(cstrSproc))) {
        // **TODO:** goto found: GoTo PROC_EXIT;
      }

      ADODB.Command w_aDOCommand = m_adwADO.getADOCommand();
      // ---Parameter #1---
      // Define the return value that represents the error code (i.e. reason) why
      // the stored procedure failed.
      prmReturnValue = w_aDOCommand.CreateParameter(Name:="@return_value", Type:=adInteger, Direction:=adParamReturnValue, .value:=Null);
      w_aDOCommand.Parameters.Append(prmReturnValue);

      m_rstLookup = w_aDOCommand.Execute();
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    // Have to check for "Not (m_rstLookup Is Nothing)" to avoid a
    // "91 - Object variable or With block variable not set" runtime error
    if (!(m_rstLookup == null)) {
      //' Disconnect the Recordset
      modGeneral.fnFreeObject(m_rstLookup.ActiveConnection);
    }
    modGeneral.fnFreeObject(prmReturnValue);

    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return null;
    // **TODO:** label found: PROC_ERR:;
    switch (prmReturnValue) {
      case  modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO:
        // 4028 = An error occurred while attempting to @@1 this record.
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO, MCSTRNAME+ cstrCurrentProc, "locate");
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
        break;
    }

    // If any other errors exist, i.e. in Err object, then let it fall through into default error handling.

    switch (VBA.ex.Number) {
      //' Object not found
      case  -2147217900:
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_FERR_SPROC_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, cstrSproc);
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
        break;

      default:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }
}


//////////////////////////////////////////////////////////////////////////////////////////////////
  public boolean getRelativeRecord(String strKey1, enumPositionDirection lngPositionDirection) {
    //--------------------------------------------------------------------------
    // Procedure:   GetRelativeRecord
    // Description: Refreshes the Lookup recordset and repositions it to
    //              the record relative to the specified **logical** key value.
    //              Then, it resets each of the class properties that correspond
    //              to columns in the underlying table so the form is able
    //              to load that newly-positioned-to record's data to
    //              itrs on-screen controls.
    //
    //              NOTE: For each table key, there should be an input parameter
    //                    and a local var (i.e. strKey1ForNewRec) of the
    //                    appropriate data type! Also, the setting of
    //                    the .Filter property below must reflect each
    //                    table key.
    //
    //                    This table's true key is CLM_ID, but the Insured screen
    //                    shows claims in CLM_NUM sequence. Hence, the key this
    //                    proc uses is CLM_NUM.
    //
    // Params:
    //     strKey1              (in) = ClmNum value from which to do the relative
    //                                 repositioning
    //     lngPositionDirection (in) = Indicates to which relative record the
    //                                 recordset should be positioned (relative
    //                                 to the strKey1 parameter value).
    //
    //
    // Called By:   cmdDelete_Click( ) of frmclaim.frm
    //              cmdUpdate_Click( ) of frmclaim.frm
    //              cmdNavigate_Click( ) of frmclaim.frm
    //              fnAddRecord( ) of frmclaim.frm
    //
    // Returns:     True if successful; False otherwise
    //-----------------------------------------------------------------------------

    //!CUSTOMIZE!  This proc and all calls to it must be customized to reflect
    //             one parameter for each key column of the table. Make sure the
    //             parameter is defined to be of the right data type. Also,
    //             the way the recordset's .Find property is set must be changed
    //             to reflect each key column so the right record will be located.
    //             Also make sure that the substitution values passed to
    //             SaveAppSpecificError are correct and TRIM'd if appropriate.

    "GetRelativeRecord"
.equals(Const cstrCurrentProc As String);
    DBRecordSet rstTemp = null;
    String strKey1ForNewRec = "";

    try {

      //...........................................................................
      // Refresh the lookup data (m_rstLookupData) so other's changes
      // --and our own-- are now reflected in it. This resets the Lookup data,
      // record count, and current record number, and leaves the Lookup recordset
      // positioned to the first record (if there are records) or BOF (if there are
      // no records).
      //...........................................................................
      getLookupData();

      switch (lngPositionDirection) {
        case  enumPositionDirection.ePDPREVIOUSRECORD:
          // Make visible only those rows with keys prior to the specified key
          m_rstLookup.Filter = "clm_num < '"+ strKey1+ "'";
          if (m_rstLookup.RecordCount != 0) {
            // Grab the last record. The one with the highest key less than the
            // specified key is the one we want.
            m_adwADO.moveLast(m_rstLookup);
            //!CUSTOMIZE! The constant referenced below should refer to the key field.
            strKey1ForNewRec = m_rstLookup.Fields(MCSTRCLMNUM).value;
            // Okay, we got it. Now make all records in the Lookup recordset
            // visible again, and then reposition to the new record.
            m_rstLookup.Filter = adFilterNone;
            m_adwADO.moveFirst(m_rstLookup);
            m_rstLookup.Find("clm_num = '"+ strKey1ForNewRec+ "'");
            // If the new record wasn't found, generate an error. We should
            // never hit this error, except due to bad program logic, since
            // it means the new record whose key value was just identified
            // could not be found.
            if (m_rstLookup.EOF) {
              // Should never hit this code. It means the new record whose key
              // value was just identified could not be found.
              modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO, MCSTRNAME+ cstrCurrentProc, "locate");
              /**TODO:** resume found: Resume(PROC_EXIT)*/;
            }
          } 
          else {
            // No records meet the criteria.
            // If there are any records, show the first one.
            // If there are no records, the caller (the form) should go into Add mode
            // upon seeing that the m_rstLookup.LookupRecordCount = 0.
            m_rstLookup.Filter = adFilterNone;
            if (m_rstLookup.RecordCount != 0) {
              m_adwADO.moveFirst(m_rstLookup);
            }
          }


          break;

        case  enumPositionDirection.ePDNEXTRECORD:
          // Make visible only those rows with keys prior to the specified key
          m_rstLookup.Filter = "clm_num > '"+ strKey1+ "'";
          if (m_rstLookup.RecordCount != 0) {
            // Grab the first record. The one with the lowest key higher than the
            // specified key is the one we want.
            m_adwADO.moveFirst(m_rstLookup);
            //!CUSTOMIZE! The constant referenced below should refer to the key field.
            strKey1ForNewRec = m_rstLookup.Fields(MCSTRCLMNUM).value;
            // Okay, we got it. Now make all records in the Lookup recordset
            // visible again, and then reposition to the new record.
            m_rstLookup.Filter = adFilterNone;
            m_adwADO.moveFirst(m_rstLookup);
            m_rstLookup.Find("clm_num = '"+ strKey1ForNewRec+ "'");
            // If the new record wasn't found, generate an error. We should
            // never hit this error, except due to bad program logic, since
            // it means the new record whose key value was just identified
            // could not be found.
            if (m_rstLookup.EOF) {
              modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO, MCSTRNAME+ cstrCurrentProc, "locate");
              /**TODO:** resume found: Resume(PROC_EXIT)*/;
            }
          } 
          else {
            // No records meet the criteria.
            // If there are any records, show the last one.
            // If there are no records, the caller (the form) should go into Add mode
            // upon seeing that the m_rstLookup.LookupRecordCount = 0.
            m_rstLookup.Filter = adFilterNone;
            if (m_rstLookup.RecordCount != 0) {
              m_adwADO.moveLast(m_rstLookup);
            }
          }



          break;

        case  enumPositionDirection.ePDSAMERECORD:
          // This is used by Update processing, where we just want to
          // stay on the just-updated record but make sure its current
          // data (from the DBMS) is loaded to the class properties and
          // other's changes to any record in the Lookup recordset are
          // visible.
          m_rstLookup.Find("clm_num = '"+ strKey1+ "'");
          // If the record wasn't found, generate an error. We should
          // never hit this error, except due to bad program logic, since
          // it means the current record could not be found.
          if (m_rstLookup.EOF) {
            modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO, MCSTRNAME+ cstrCurrentProc, "locate");
            /**TODO:** resume found: Resume(PROC_EXIT)*/;
          }


          break;

        case  enumPositionDirection.ePDFIRSTRECORD:
          // This operation is used by the form's cmdNavigate_Click.
          // It ignores the passed-in key parameter(s).
          // Do *not* generate an error if we hit BOF since that event
          // handler will use see that the Lookup recordset's position
          // is at BOF and throw the form into Add mode.
          m_adwADO.moveFirst(m_rstLookup);


          break;

        case  enumPositionDirection.ePDLASTRECORD:
          // This operation is used by the form's cmdNavigate_Click.
          // It ignores the passed-in key parameter(s).
          // Do *not* generate an error if we hit BOF since that event
          // handler will use see that the Lookup recordset's position
          // is at EOF and throw the form into Add mode.
          m_adwADO.moveLast(m_rstLookup);


          break;

        default:
          modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_UNEXPECTED_VAL_SELECT_CASE, MCSTRNAME+ cstrCurrentProc);
          // **TODO:** goto found: GoTo PROC_EXIT;
          break;
      }

      //...........................................................................
      // Get the column data for the just-repositioned-to record and load it to the
      // class properties corresponding to those columns.
      //...........................................................................
      if (m_rstLookup.BOF && m_rstLookup.EOF) {
        fnClearPropertyValues();
      } 
      else {
        //!CUSTOMIZE! The constant referenced below should refer to the key field.
        getSingleRecord(lngKey1:=.Fields(MCSTRCLMID).value, bSynchLookupRST:=True);
      }
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    modGeneral.fnFreeRecordset(rstTemp);

    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return null;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public boolean getSingleRecord(int lngKey1, boolean bSynchLookupRST) {
    boolean _rtn = false;
    //--------------------------------------------------------------------------
    // Procedure:   GetSingleRecord
    // Description: Obtains data from the database for the specified key(s).
    //              It then loads its columnar values to the class properties
    //              that correspond to those columns. It also saves the
    //              Last Updated info to separate member variables so it can
    //              be used (when/if the user tries to update or delete
    //              the record) to determine if another user affected this
    //              record since this function retrieved it.
    //
    //              This proc should not *refresh* the Lookup recordset. It
    //              is merely retrieving all of the table columns for the
    //              specified record. The Lookup recordset only contains
    //              a subset of the columns for that key.
    //
    //
    //              NOTE: For each table key, there should be an input parameter
    //                    of the appropriate data type!
    //
    // Returns:     Boolean
    // Params:
    //    lngKey1         (in) = represents the primary key for the table (clm_id)
    //    bSynchLookupRST (in) = indicates whether the Lookup recordset should be
    //                           repositioned to the record this function just
    //                           retrieved. This would be set to True by the
    //                           form's vfgLookup_ChangeEdit event handler to ensure
    //                           the "record x of y" will be set appropriately when
    //                           it calls fnLoadControls.
    //-----------------------------------------------------------------------------
    // Modified:     Berry Kropiwka - Added Compact Calc - 2019-09-27

    //!CUSTOMIZE!  This proc and all calls to it must be customized to reflect
    //             one parameter for each key column of the table. Make sure the
    //             parameter is defined to be of the right data type. The
    //             "With Me" block must be updated to reflect the current set of
    //             wrapper properties and table column names. Also, the way
    //             the recordset's .Find property is set must be changed
    //             to reflect each key column so the right record will be located.
    //             Also make sure that the substitution values passed to
    //             SaveAppSpecificError are correct and TRIM'd if appropriate.

    "GetSingleRecord"
.equals(Const cstrCurrentProc As String);
    DBRecordSet rstTemp = null;

    try {

      _rtn = false;

      //!CUSTOMIZE! fnSelectRecord call should pass the key column(s)
      rstTemp = fnSelectRecord(lngKey1);

      //*TODO:** can't found type for with block
      //*With this
      __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = this;
      // To ensure Nulls and/or True/False are appropriately translated, use:
      //     adChar        (Boolean, regardless of nullability)    = fnYNToBool
      //     adChar        (Other, regardless of nullability)      = fnZLSIfNull
      //     adDBTimeStamp (regardless of nullability)             = N/A
      //     adInteger     (not nullable)                          = fnZeroIfNull
      //     adInteger     (nullable)                              = N/A
      //     adNumeric     (not nullable)                          = fnZeroIfNull
      //     adNumeric     (nullable)                              = N/A
      //     adVarChar     (regardless of nullability)             = fnZLSIfNull
      w___TYPE_NOT_FOUND.AdmnSystCd = modDataConversion.fnZLSIfNull(rstTemp!admn_syst_cd);
      w___TYPE_NOT_FOUND.ClmId = modDataConversion.fnZeroIfNull(rstTemp!clm_id);
      w___TYPE_NOT_FOUND.ClmInsdDthDt = rstTemp!clm_insd_dth_dt;
      w___TYPE_NOT_FOUND.ClmInsdFirstNm = modDataConversion.fnZLSIfNull(rstTemp!clm_insd_first_nm);
      w___TYPE_NOT_FOUND.ClmInsdLastNm = modDataConversion.fnZLSIfNull(rstTemp!clm_insd_last_nm);
      w___TYPE_NOT_FOUND.ClmInsdSsnNum = modDataConversion.fnZLSIfNull(rstTemp!clm_insd_ssn_num);
      w___TYPE_NOT_FOUND.ClmNum = modDataConversion.fnZLSIfNull(rstTemp!clm_num);
      w___TYPE_NOT_FOUND.ClmPolNum = modDataConversion.fnZLSIfNull(rstTemp!clm_pol_num);
      w___TYPE_NOT_FOUND.ClmProofDt = rstTemp!clm_proof_dt;
      w___TYPE_NOT_FOUND.ClmTotClmPdAmt = rstTemp!clm_tot_clm_pd_amt;
      w___TYPE_NOT_FOUND.ClmTotDthbPmtAmt = rstTemp!clm_tot_dthb_pmt_amt;
      w___TYPE_NOT_FOUND.ClmTotIntAmt = rstTemp!clm_tot_int_amt;
      w___TYPE_NOT_FOUND.ClmTotWthldAmt = rstTemp!clm_tot_wthld_amt;
      w___TYPE_NOT_FOUND.InsdDthResStCd = modDataConversion.fnZLSIfNull(rstTemp!insd_dth_res_st_cd);
      w___TYPE_NOT_FOUND.IssStCd = modDataConversion.fnZLSIfNull(rstTemp!iss_st_cd);
      w___TYPE_NOT_FOUND.LstUpdtDtm = rstTemp!lst_updt_dtm;
      w___TYPE_NOT_FOUND.LstUpdtUserId = modDataConversion.fnZLSIfNull(rstTemp!lst_updt_user_id);
      w___TYPE_NOT_FOUND.PycoTypCd = modDataConversion.fnZLSIfNull(rstTemp!pyco_typ_cd);
      w___TYPE_NOT_FOUND.ClmForResDthInd = modGeneral.fnYNToBool(rstTemp!clm_for_res_dth_ind);
      w___TYPE_NOT_FOUND.ClmCompactClcnInd = modGeneral.fnTFToBool(rstTemp!clm_compact_clcn_ind);

      // Save original Last Updated info, to be used during UpdateRecord( ) and DeleteRecord( )
      // to determine if another user updated the record since it was retrieved.
      m_dteLstUpdtDtm_Original = modDataConversion.fnZLSIfNull(rstTemp!lst_updt_dtm);
      m_strLstUpdtUserId_Original = modDataConversion.fnZLSIfNull(rstTemp!lst_updt_user_id);

      if (bSynchLookupRST) {
        m_adwADO.moveFirst(m_rstLookup);
        m_rstLookup.Find("clm_id = "+ ((Integer) getClmId()).toString());
        if (m_rstLookup.EOF) {
          modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO, MCSTRNAME+ cstrCurrentProc, "locate");
          /**TODO:** resume found: Resume(PROC_EXIT)*/;
        }
      }

      _rtn = true;
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    modGeneral.fnFreeRecordset(rstTemp);

    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    _rtn = false;

    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public void goToFirstRecord() {
    // **************************************************************************
    // Function  : GoToFirstRecord
    // Purpose   : Moves to the First record in the table
    // Parameters: N/A
    // Returns   : Int
    // **************************************************************************
    "GoToFirstRecord"
.equals(Const cstrCurrentProc As String);
    try {

      if (m_rstLookup.RecordCount > 0) {
        m_adwADO.moveFirst(m_rstLookup);
        // Get the requested record and reposition the Lookup recordset to that record
        //!CUSTOMIZE! The constant referenced below should refer to the key field.
        getSingleRecord(m_rstLookup.Fields(MCSTRCLMID).value);
      }
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public void goToLastRecord() {
    // **************************************************************************
    // Function  : GoToLastRecord
    // Purpose   : Moves to the Last record in the table
    // Parameters: N/A
    // Returns   : Int
    // **************************************************************************
    "GoToLastRecord"
.equals(Const cstrCurrentProc As String);
    try {

      if (m_rstLookup.RecordCount > 0) {
        m_adwADO.moveLast(m_rstLookup);
        // Get the requested record and reposition the Lookup recordset to that record
        //!CUSTOMIZE! The constant referenced below should refer to the key field.
        getSingleRecord(m_rstLookup.Fields(MCSTRCLMID).value);
      }
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public void goToNextRecord() {
    // **************************************************************************
    // Function  : GoToNextRecord
    // Purpose   : Moves to the next record in the table
    // Parameters: N/A
    // Returns   : Int
    // **************************************************************************
    "GoToNextRecord"
.equals(Const cstrCurrentProc As String);
    try {

      if (m_rstLookup.RecordCount > 0) {
        m_adwADO.moveNext(m_rstLookup);
        if (getLookupIsAtBOF() || getLookupIsAtEOF()) {
          getLookupData();
          getRelativeRecord(getClmNum(), enumPositionDirection.ePDNEXTRECORD);
        } 
        else {
          // Get the requested record and reposition the Lookup recordset to that record
          //!CUSTOMIZE! The constant referenced below should refer to the key field.
          getSingleRecord(m_rstLookup.Fields(MCSTRCLMID).value);
        }
      }
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public void goToPreviousRecord() {
    // **************************************************************************
    // Function  : GoToPreviousRecord
    // Purpose   : Moves to the previous record in the table
    // Parameters: N/A
    // Returns   : Int
    // **************************************************************************
    "GoToPreviousRecord"
.equals(Const cstrCurrentProc As String);
    try {

      if (m_rstLookup.RecordCount > 0) {
        m_adwADO.movePrev(m_rstLookup);
        if (getLookupIsAtBOF() || getLookupIsAtEOF()) {
          getLookupData();
          getRelativeRecord(getClmNum(), enumPositionDirection.ePDPREVIOUSRECORD);
        } 
        else {
          // Get the requested record and reposition the Lookup recordset to that record
          //!CUSTOMIZE! The constant referenced below should refer to the key field.
          getSingleRecord(m_rstLookup.Fields(MCSTRCLMID).value);
        }
      }
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public boolean haveDependents(int lngKey1, String strDependentTable) { // TODO: Use of ByRef founded Public Function HaveDependents(ByVal lngKey1 As Long, ByRef strDependentTable As String) As Boolean
    boolean _rtn = false;
    // Comments  : Determines whether the current record can be deleted without
    //             hitting a referential integrity violation due to either:
    //             a. row(s) existing in other tables that use the current key value
    //                as a foreign key
    //             b. (for tables with circular references only) row(s) existing in the same table
    //                which has a circular reference to to the current key value.
    //             The calling form should look at the return value. If True, then
    //             the form's Delete button should be disabled.
    //
    // Parameters:
    //   lngKey1
    //
    // Called by : fnSetCommandButtons( ) in each maintenance screen
    //
    // Returns   : True if there are children or other dependencies; False otherwise
    // Modified  :
    // --------------------------------------------------

    //!CUSTOMIZE!  This proc must be customized to return True unconditionally, if
    //             the table has no dependencies to any other tables.
    //
    //             Otherwise, it must be customized to have an input parameter of
    //             the correct data type for each key to the table, call the
    //             correct stored procedure with the correct number and type of
    //             paraemters, and interpret its return values correctly.
    "HaveDependents"
.equals(Const cstrCurrentProc As String);
    //' Stored procedure to execute
    "dbo.proc_claim_verify_dependents"
.equals(Const cstrSproc As String);
    cadwADOWrapper adwTemp = null;
    DBRecordSet rstTemp = null;
    ADODB.Parameter prmReturnValue = null;
    ADODB.Parameter prmClmId = null;
    ADODB.Parameter prmDependent_Table = null;

    try {

      adwTemp = new cadwADOWrapper();
      adwTemp.commandSetSproc(strSprocName:=cstrSproc);

      ADODB.Command w_aDOCommand = adwTemp.getADOCommand();
      // ---Parameter #1---
      // Define the return value that represents the error code (i.e. reason) why
      // the stored procedure failed.
      prmReturnValue = w_aDOCommand.CreateParameter(Name:="@return_value", Type:=adInteger, Direction:=adParamReturnValue, .value:=Null);
      w_aDOCommand.Parameters.Append(prmReturnValue);

      // ---Parameter #2---
      // Define the input parameter that represents the key value to see
      // if it exists as a foreign key on dependent tables
      prmClmId = w_aDOCommand.CreateParameter(Name:="@clm_id", Type:=adInteger, Direction:=adParamInput, .value:=lngKey1);
      w_aDOCommand.Parameters.Append(prmClmId);

      // ---Parameter #3---
      // Define the output parameter that indicates whether **any** dependent table
      // has children. If True, we need to look at prm2 and report a 4029 error.
      prmDependent_Table = w_aDOCommand.CreateParameter(Name:="@Dependent_Table", Type:=adVarChar, Size:=255, Direction:=adParamOutput);
      w_aDOCommand.Parameters.Append(prmDependent_Table);

      // Now execute the sproc...and you get access to those output parameters
      // as well as, if applicable, the recordset/resultset it returns
      rstTemp = w_aDOCommand.Execute;
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    modGeneral.fnFreeObject(adwTemp);
    modGeneral.fnFreeRecordset(rstTemp);
    modGeneral.fnFreeObject(prmReturnValue);
    modGeneral.fnFreeObject(prmClmId);
    modGeneral.fnFreeObject(prmDependent_Table);

    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (prmReturnValue) {
        // The 4027 return code should only occur if a multi-user situation has occurred,
        // such as User A going into Update mode on a record that another user just
        // deleted. For this reason, we'll remove any trace that this error occurred
        // and return True. When and if the user clicks Update, then they'll get a
        // message that another user deleted the record.
      case  modResConstants.gCRES_NERR_REC_NOT_FOUND:
        // 4027 = The specified record was not found in the database (@@1).
        _rtn = true;
        //' This is actually ignored by the caller
        strDependentTable = "Unknown";
        // Remove any trace that this error occurred since we're not going to report it as an error.
        VBA.ex.Clear;
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
        break;

      case  modResConstants.gCRES_NERR_DEPENDENT_RECS_EXIST:
        // 4029 = This @@1 is associated with one or more records on the @@2 table and cannot be deleted until those records themselves are deleted.
        _rtn = true;
        strDependentTable = prmDependent_Table;
        // Remove any trace that this error occurred since we're not going to report it as an error.
        VBA.ex.Clear;
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
        break;
    }

    // If any other errors exist, i.e. in Err object, then let it fall through into default error handling.

    switch (VBA.ex.Number) {
      //' Object not found
      case  -2147217900:
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_FERR_SPROC_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, cstrSproc);
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
        break;

      default:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public boolean updateRecord() {
    boolean _rtn = false;
    //--------------------------------------------------------------------------
    // Procedure:   UpdateRecord
    // Description: Updates a single record based on current values stored in
    //              the class properties corresponding to table columns.
    //
    // Returns:     True if successful; False otherwise
    // Params:      N/A
    // Date:        01/07/2002
    //-----------------------------------------------------------------------------
    // Modified:     Berry Kropiwka - Added Compact Calc - 2019-09-27

    //!CUSTOMIZE!  Customize the name of the stored procedure, the number and names
    //             of the parameters and, perhaps, the return values trapped in the
    //             error handler.

    "UpdateRecord"
.equals(Const cstrCurrentProc As String);
    //' Stored procedure to execute
    "dbo.proc_claim_update"
.equals(Const cstrSproc As String);
    DBRecordSet rstTemp = null;
    boolean bSuccessful = false;
    ADODB.Parameter prmReturnValue = null;
    ADODB.Parameter prmClmId = null;
    ADODB.Parameter prmAdmnSystCd = null;
    ADODB.Parameter prmClmInsdDthDt = null;
    ADODB.Parameter prmClmInsdFirstNm = null;
    ADODB.Parameter prmClmInsdLastNm = null;
    ADODB.Parameter prmClmInsdSsnNum = null;
    ADODB.Parameter prmClmNum = null;
    ADODB.Parameter prmClmPolNum = null;
    ADODB.Parameter prmClmProofDt = null;
    ADODB.Parameter prmClmTotClmPdAmt = null;
    ADODB.Parameter prmClmTotDthbPmtAmt = null;
    ADODB.Parameter prmClmTotIntAmt = null;
    ADODB.Parameter prmClmTotWthldAmt = null;
    ADODB.Parameter prmInsdDthResStCd = null;
    ADODB.Parameter prmIssStCd = null;
    ADODB.Parameter prmPycoTypCd = null;
    ADODB.Parameter prmClmForResDthInd = null;
    ADODB.Parameter prmClmCompactClcnInd = null;
    ADODB.Parameter prmInvalid_Key = null;

    try {

      if (!(m_adwADO.commandSetSproc(cstrSproc))) {
        // **TODO:** goto found: GoTo PROC_EXIT;
      }

      // For Char/VarChar fields,
      //     * Use fnNullIfZLS to ensure Nulls are appropriately handled.
      //     * Do *not* set the optional 2nd parameter to fnNullIfZLS to True.
      // For numeric fields,
      //     * Use fnNullIfZero to ensure Nulls are appropriately handled.
      // For Y/N fields,
      //     * Use fnBoolToYN to ensure True/False is appropriately translated.


      ADODB.Command w_aDOCommand = m_adwADO.getADOCommand();
      // ---Parameter #1---
      // Define the return value that represents the error code (i.e. reason) why
      // the stored procedure failed.
      prmReturnValue = w_aDOCommand.CreateParameter(Name:="@return_value", Type:=adInteger, Direction:=adParamReturnValue, .value:=Null);
      w_aDOCommand.Parameters.Append(prmReturnValue);

      // ---Parameter #2---
      // Define the CLM_ID parameter
      prmClmId = w_aDOCommand.CreateParameter(Name:="@clm_id", Type:=adInteger, Direction:=adParamInput, .value:=ClmId);
      w_aDOCommand.Parameters.Append(prmClmId);

      // ---Parameter #3---
      // Define the ADMN_SYST_CD parameter
      prmAdmnSystCd = w_aDOCommand.CreateParameter(Name:="@admn_syst_cd", Type:=adChar, Direction:=adParamInput, Size:=2, .value:=fnNullIfZLS(varIn:=AdmnSystCd, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmAdmnSystCd);

      // ---Parameter #4---
      // Define the CLM_INSD_DTH_DT parameter
      prmClmInsdDthDt = w_aDOCommand.CreateParameter(Name:="@clm_insd_dth_dt", Type:=adDBTimeStamp, Direction:=adParamInput, Size:=16, .value:=ClmInsdDthDt);
      w_aDOCommand.Parameters.Append(prmClmInsdDthDt);

      // ---Parameter #5---
      // Define the CLM_INSD_FIRST_NM parameter
      prmClmInsdFirstNm = w_aDOCommand.CreateParameter(Name:="@clm_insd_first_nm", Type:=adVarChar, Direction:=adParamInput, Size:=50, .value:=fnNullIfZLS(varIn:=ClmInsdFirstNm, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmClmInsdFirstNm);

      // ---Parameter #6---
      // Define the CLM_INSD_LAST_NM parameter
      prmClmInsdLastNm = w_aDOCommand.CreateParameter(Name:="@clm_insd_last_nm", Type:=adVarChar, Direction:=adParamInput, Size:=50, .value:=fnNullIfZLS(varIn:=ClmInsdLastNm, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmClmInsdLastNm);

      // ---Parameter #7---
      // Define the CLM_INSD_SSN_NUM parameter
      prmClmInsdSsnNum = w_aDOCommand.CreateParameter(Name:="@clm_insd_ssn_num", Type:=adChar, Direction:=adParamInput, Size:=9, .value:=fnNullIfZLS(varIn:=ClmInsdSsnNum, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmClmInsdSsnNum);

      // ---Parameter #8---
      // Define the CLM_NUM parameter
      prmClmNum = w_aDOCommand.CreateParameter(Name:="@clm_num", Type:=adVarChar, Direction:=adParamInput, Size:=20, .value:=fnNullIfZLS(varIn:=ClmNum, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmClmNum);

      // ---Parameter #9---
      // Define the CLM_POL_NUM parameter
      prmClmPolNum = w_aDOCommand.CreateParameter(Name:="@clm_pol_num", Type:=adChar, Direction:=adParamInput, Size:=15, .value:=fnNullIfZLS(varIn:=ClmPolNum, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmClmPolNum);

      // ---Parameter #10---
      // Define the CLM_PROOF_DT parameter
      prmClmProofDt = w_aDOCommand.CreateParameter(Name:="@clm_proof_dt", Type:=adDBTimeStamp, Direction:=adParamInput, Size:=16, .value:=ClmProofDt);
      w_aDOCommand.Parameters.Append(prmClmProofDt);

      // ---Parameter #11---
      // Define the CLM_TOT_CLM_PD_AMT parameter
      prmClmTotClmPdAmt = w_aDOCommand.CreateParameter(Name:="@clm_tot_clm_pd_amt", Type:=adNumeric, Direction:=adParamInput, .value:=ClmTotClmPdAmt);
      prmClmTotClmPdAmt.Precision = m_varClmTotClmPdAmt.precision;
      prmClmTotClmPdAmt.NumericScale = m_varClmTotClmPdAmt.numericScale;
      w_aDOCommand.Parameters.Append(prmClmTotClmPdAmt);

      // ---Parameter #12---
      // Define the CLM_TOT_DTHB_PMT_AMT parameter
      prmClmTotDthbPmtAmt = w_aDOCommand.CreateParameter(Name:="@clm_tot_dthb_pmt_amt", Type:=adNumeric, Direction:=adParamInput, .value:=ClmTotDthbPmtAmt);
      prmClmTotDthbPmtAmt.Precision = m_varClmTotDthbPmtAmt.precision;
      prmClmTotDthbPmtAmt.NumericScale = m_varClmTotDthbPmtAmt.numericScale;
      w_aDOCommand.Parameters.Append(prmClmTotDthbPmtAmt);

      // ---Parameter #13---
      // Define the CLM_TOT_INT_AMT parameter
      prmClmTotIntAmt = w_aDOCommand.CreateParameter(Name:="@clm_tot_int_amt", Type:=adNumeric, Direction:=adParamInput, .value:=ClmTotIntAmt);
      prmClmTotIntAmt.Precision = m_varClmTotIntAmt.precision;
      prmClmTotIntAmt.NumericScale = m_varClmTotIntAmt.numericScale;
      w_aDOCommand.Parameters.Append(prmClmTotIntAmt);

      // ---Parameter #14---
      // Define the CLM_TOT_WTHLD_AMT parameter
      prmClmTotWthldAmt = w_aDOCommand.CreateParameter(Name:="@clm_tot_wthld_amt", Type:=adNumeric, Direction:=adParamInput, .value:=ClmTotWthldAmt);
      prmClmTotWthldAmt.Precision = m_varClmTotWthldAmt.precision;
      prmClmTotWthldAmt.NumericScale = m_varClmTotWthldAmt.numericScale;
      w_aDOCommand.Parameters.Append(prmClmTotWthldAmt);

      // ---Parameter #15---
      // Define the INSD_DTH_RES_ST_CD parameter
      prmInsdDthResStCd = w_aDOCommand.CreateParameter(Name:="@insd_dth_res_st_cd", Type:=adChar, Direction:=adParamInput, Size:=2, .value:=fnNullIfZLS(varIn:=InsdDthResStCd, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmInsdDthResStCd);

      // ---Parameter #16---
      // Define the ISS_ST_CD parameter
      prmIssStCd = w_aDOCommand.CreateParameter(Name:="@iss_st_cd", Type:=adChar, Direction:=adParamInput, Size:=2, .value:=fnNullIfZLS(varIn:=IssStCd, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmIssStCd);

      // ---Parameter #17---
      // Define the PYCO_TYP_CD parameter
      prmPycoTypCd = w_aDOCommand.CreateParameter(Name:="@pyco_typ_cd", Type:=adChar, Direction:=adParamInput, Size:=8, .value:=fnNullIfZLS(varIn:=PycoTypCd, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmPycoTypCd);

      // ---Parameter #18---
      // Define the CLM_FOR_RES_DTH_IND parameter
      prmClmForResDthInd = w_aDOCommand.CreateParameter(Name:="clm_for_res_dth_ind", Type:=adChar, Direction:=adParamInput, Size:=1, .value:=fnBoolToYN(getClmForResDthInd()));
      w_aDOCommand.Parameters.Append(prmClmForResDthInd);

      // ---Parameter #19---
      // Define the clm_compact_clcn_ind parameter
      prmClmCompactClcnInd = w_aDOCommand.CreateParameter(Name:="clm_compact_clcn_ind", Type:=adChar, Direction:=adParamInput, Size:=1, .value:=fnBoolToTF(getClmCompactClcnInd()));
      w_aDOCommand.Parameters.Append(prmClmCompactClcnInd);

      // ---Parameter #20---
      // Define the Invalid_Key output parameter, which reflects *which* foreign
      // key violation was encountered.
      prmInvalid_Key = w_aDOCommand.CreateParameter(Name:="@Invalid_Key", Type:=adVarChar, Size:=255, Direction:=adParamOutput);
      w_aDOCommand.Parameters.Append(prmInvalid_Key);

      // Do the Update
      w_aDOCommand.Execute;

      //...........................................................................
      // Refresh the Lookup recordset, re-retrieve the just-updated record so that
      // record is *still* the current record, and load its data to the table
      // wrapper's class properties so all table columns (including those set by
      // the DBMS like identity columns and Last Updated columns) are up-to-date.
      //...........................................................................
      bSuccessful = getRelativeRecord(getClmNum(), enumPositionDirection.ePDSAMERECORD);

      _rtn = true;
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    modGeneral.fnFreeRecordset(rstTemp);
    modGeneral.fnFreeObject(prmReturnValue);
    modGeneral.fnFreeObject(prmClmId);
    modGeneral.fnFreeObject(prmAdmnSystCd);
    modGeneral.fnFreeObject(prmClmInsdDthDt);
    modGeneral.fnFreeObject(prmClmInsdFirstNm);
    modGeneral.fnFreeObject(prmClmInsdLastNm);
    modGeneral.fnFreeObject(prmClmInsdSsnNum);
    modGeneral.fnFreeObject(prmClmNum);
    modGeneral.fnFreeObject(prmClmPolNum);
    modGeneral.fnFreeObject(prmClmProofDt);
    modGeneral.fnFreeObject(prmClmTotClmPdAmt);
    modGeneral.fnFreeObject(prmClmTotDthbPmtAmt);
    modGeneral.fnFreeObject(prmClmTotIntAmt);
    modGeneral.fnFreeObject(prmClmTotWthldAmt);
    modGeneral.fnFreeObject(prmInsdDthResStCd);
    modGeneral.fnFreeObject(prmIssStCd);
    modGeneral.fnFreeObject(prmPycoTypCd);
    modGeneral.fnFreeObject(prmClmForResDthInd);
    modGeneral.fnFreeObject(prmClmCompactClcnInd);
    modGeneral.fnFreeObject(prmInvalid_Key);

    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (prmReturnValue) {
      case  modResConstants.gCRES_NERR_REC_NOT_FOUND:
        // 4027 = The specified record was not found in the database (@@1).
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_REC_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, "Claim Number "+ RTrim$(getClmNum()));
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
        break;

      case  modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO:
        // 4028 = An error occurred while attempting to @@1 this record.
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO, MCSTRNAME+ cstrCurrentProc, "update");
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
        break;

      case  modResConstants.gCRES_NERR_KEY_NOT_FOUND:
        // 4032 = The @@1 specified (@@2) is no longer defined in the @@3 table. Please choose a different value.
        if ("ADMN_SYST_CD"
.equals(prmInvalid_Key.toUpperCase())) {
          modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_KEY_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, "Admin System", RTrim$(getAdmnSystCd()), "ADMIN_SYSTEM_T");
          /**TODO:** resume found: Resume(PROC_EXIT)*/;
        } 
        else if ("INSD_DTH_RES_ST_CD"
.equals(prmInvalid_Key.toUpperCase())) {
          modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_KEY_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, "Residence State", RTrim$(getInsdDthResStCd()), "STATE_T");
          /**TODO:** resume found: Resume(PROC_EXIT)*/;
        } 
        else if ("ISS_ST_CD"
.equals(prmInvalid_Key.toUpperCase())) {
          modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_KEY_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, "Issue State", RTrim$(getIssStCd()), "STATE_T");
          /**TODO:** resume found: Resume(PROC_EXIT)*/;
        } 
        else {
          // PYCO_TYP_CD
          modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_KEY_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, "Company Type", RTrim$(getPycoTypCd()), "PAYOR_COMPANY_TYPE_T");
          /**TODO:** resume found: Resume(PROC_EXIT)*/;
        }
        break;
    }

    // If any other errors exist, i.e. in Err object, then let it fall through into default error handling.

    switch (VBA.ex.Number) {
      //' Object not found
      case  -2147217900:
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_FERR_SPROC_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, cstrSproc);
        /**TODO:** resume found: Resume(PROC_EXIT)*/;

      //' Cannot insert duplicate key row in object 'xxx' with unique index 'yyy'
        break;

      case  -2147217873:
        // 4031 = A record with the specified key (@@1) already exists. Please specify a unique key.
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_ADD_WITH_NONUNIQUE_KEY, MCSTRNAME+ cstrCurrentProc, "Claim Number "+ RTrim$(getClmNum()));
        /**TODO:** resume found: Resume(PROC_EXIT)*/;

        // Added -2147217887 check to fix Claims Interest bug 2454.
      //' Invalid Character Value for Cast Specification
        break;

      case  -2147217887:
        // (Internally manifested in sproc as "arithmetic overflow error converting numeric to data type numeric")
        // gcRES_NERR_NUMERIC_FLD_TOO_LARGE as Integer (4008) = One or more numeric fields are too large to be stored in the database. Your changes cannot be saved.
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_NUMERIC_FLD_TOO_LARGE, MCSTRNAME+ cstrCurrentProc);
        /**TODO:** resume found: Resume(PROC_EXIT)*/;

        break;

      default:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}





///\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
//\                                                                  /
//|                        PRIVATE  Procedures                       |
///                                                                  \
//\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/

//////////////////////////////////////////////////////////////////////////////////////////////////
  private void fnClearPropertyValues() {
    //--------------------------------------------------------------------------
    // Procedure:   fnClearPropertyValues
    // Description: Sets the value of each class property that corresponds to a table column
    //              so it is wiped out. The remaining attributes of that UDTColumn structure
    //              are left intact. This proc is used when a navigation or refreshing
    //              of the Lookup recordset resulted in having no records (.BOF and/or .EOF
    //              is True) or the AbsolutePosition is invalid). Without calling this
    //              proc when those situations occur, we'd still have the previous record's
    //              value displayed.
    //
    // Returns:     N/A
    //-----------------------------------------------------------------------------
    // Modified:     Berry Kropiwka - Added Compact Calc - 2019-09-27

    //!CUSTOMIZE!  This proc should set each class property (of type UDTColumn) that
    //             corresponds to a table column. What it is set to depends on its
    //             data type: Strings => vbNullString
    //                        Numeric => 0
    //                        Nullable numerics => Null
    //                        Booleans => False
    //                        Dates => Now

    "fnClearPropertyValues"
.equals(Const cstrCurrentProc As String);
    Const(clngZero As Long == 0);

    try {

      //*TODO:** can't found type for with block
      //*With this
      __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = this;
      w___TYPE_NOT_FOUND.AdmnSystCd = "";
      w___TYPE_NOT_FOUND.ClmId = clngZero;
      w___TYPE_NOT_FOUND.ClmInsdDthDt = Now;
      w___TYPE_NOT_FOUND.ClmInsdFirstNm = "";
      w___TYPE_NOT_FOUND.ClmInsdLastNm = "";
      w___TYPE_NOT_FOUND.ClmInsdSsnNum = "";
      w___TYPE_NOT_FOUND.ClmNum = "";
      w___TYPE_NOT_FOUND.ClmPolNum = "";
      w___TYPE_NOT_FOUND.ClmProofDt = Now;
      w___TYPE_NOT_FOUND.ClmTotClmPdAmt = Null;
      w___TYPE_NOT_FOUND.ClmTotDthbPmtAmt = Null;
      w___TYPE_NOT_FOUND.ClmTotIntAmt = Null;
      w___TYPE_NOT_FOUND.ClmTotWthldAmt = Null;
      w___TYPE_NOT_FOUND.InsdDthResStCd = "";
      w___TYPE_NOT_FOUND.IssStCd = "";
      w___TYPE_NOT_FOUND.LstUpdtDtm = Now;
      w___TYPE_NOT_FOUND.LstUpdtUserId = "";
      w___TYPE_NOT_FOUND.PycoTypCd = "";
      w___TYPE_NOT_FOUND.ClmForResDthInd = false;
      w___TYPE_NOT_FOUND.ClmCompactClcnInd = false;

      // Also reset the saved "original" values for the Last Updated info
      m_dteLstUpdtDtm_Original = Now;
      m_strLstUpdtUserId_Original = "";
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here

    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }
}



//////////////////////////////////////////////////////////////////////////////////////////////////
//UPDATED FOR SQL 2008
//////////////////////////////////////////////////////////////////////////////////////////////////
  private boolean fnGetColMetaData(udtColumn pudtCol, DBRecordSet prstIn) { // TODO: Use of ByRef founded Private Function fnGetColMetaData(ByRef pudtCol As udtColumn, ByRef prstIn As ADODB.Recordset) As Boolean
    boolean _rtn = false;
    //--------------------------------------------------------------------------
    // Procedure:   fnGetColMetaData
    // Description: Given a variable that represents a table column,
    //              load its meta data to its udtColumn-defined properties,
    //              setting default values based on the data type of that
    //              table column.
    //
    //              NOTE: If the default properties aren't right for a
    //                    particular column, then the table wrapper should
    //                    have override values coded in its fnLoadColMetaData
    //                    method.
    //
    //                    **THIS** fnGETColMetaData method should be identical
    //                    in all table wrappers!
    //
    // Returns:     True if successful; False otherwise
    //
    // Params:      pudtCol  (in/out)  - mbr var associated with a table column
    //              prstIn   (in/out)  - recordset containing meta data, positioned
    //                                   to the row that corresponds to the
    //                                   specified column (pudtCol)
    // Date:        04/03/2002
    //-----------------------------------------------------------------------------
    try {
      "fnGetColMetaData"
.equals(Const cstrCurrentProc As String);
      "TRUE"
.equals(Const cstrTrue As String);
      "#"
.equals(Const cstrNumericChar As String);
      "&"
.equals(Const cstrAnyCharChar As String);
      String strDomainNameToParse = "";
      String strDefaultValueToParse = "";
      String strEditedDefaultValue = "";

      // Do NOT reposition the prstIn recordset as this will mess up the caller (fnGetMetaData)
      // who is calling *this* proc for each table column (i.e. row in that recordset).


      //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
      // The following data types are used within the TRS tables:
      //       Identity fields   = adInteger
      //       Char fields       = adChar
      //       Varchar fields    = adChar
      //       Dates/lst_updt_dtm= adDBTimeStamp   (e.g., lst_updt_dtm or eff_dt)
      //       Area codes        = adNumeric
      //       Indicators        = adChar          (e.g., Yes/No or other values)
      //       Percents/Factors  = adNumeric
      //       Monetary amounts  = adNumeric

      // Certain TYPE_NAMEs are typically used to control default values or indicate
      // certain usage, such as:
      //    a. dom_id_key........for identity columns
      //    b. dom_pct...........for percentages
      //    c. dom_ind...........for indicator fields (bound to rule_ind, this is a Y, N
      //       dom_indyn.........for indicator fields (bound to rul_indyn, this is a Y, N or Null
      //       dom_char1_ind.....for indicator fields (bound to rule_char1_ind, this is a Y or N (*not* null)
      //    d. dom_dt_nn.........for dates (not nullable)
      //       dom_dt-null.......for dates (nullable)
      //       dom_lst_updt_dtm..for dates, this sets the system date (getdate()) as the default value on an Insert
      //    e. dom_lst_updt_id...to set the user's ACF2 (suser_sname()) as the default value on an Insert
      //
      // The following meta data appears to be available for these data types:
      //
      // Property:  (ignored) (HasDefault) (Default (Is       (Dollar    (Decimal    (Max
      //                                    Value)  Nullable)  Positions) Positions)  Characters)
      //
      //                                                                             CHARACTER_
      //            DATETIME_  COLUMN_     COLUMN_  IS_       NUMERIC_   NUMERIC_    MAXIMUM_     DOMAIN_
      // DATA_TYPE  PRECISION  HASDEFAULT  DEFAULT  NULLABLE  PRECISION  SCALE       LENGTH       NAME
      // ---------  ---------  ----------  -------  --------  ---------  --------    ----------   -------
      // adInteger     No         Yes       Yes-1       Yes       Yes       No          No         Yes-2
      // adChar        No         Yes       Yes-1       Yes       No        No          Yes        Yes-2
      // adNumeric     No         Yes       Yes-1       Yes       Yes       Yes         No         Yes-2
      // adDBTimeStamp Yes        Yes       Yes-1       Yes       No        No          No         Yes-2
      //
      // Legend:
      //   Yes-1 - COLUMN_DEF is present only when COLUMN_HASDEFAULT is present and is set to TRUE
      //   Yes-2 - TYPE_NAME is present only when a domain name has been assigned. It appears to be
      //          able to be present on any data type.
      //
      // DATETIME_PRECISION appears to be meaningless in the Sun Life environment. Danny Khoury thinks
      // it refers to "smalldatetime" versus <regular> "datetime", and we only use the latter. Hence,
      // we won't bother collecting this piece of meta data for dates.
      //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
      pudtCol.colName = prstIn("COLUMN_NAME").value.toUpperCase();
      pudtCol.dataType = prstIn("DATA_TYPE").value;

      strDefaultValueToParse = modDataConversion.fnZLSIfNull(prstIn("COLUMN_DEF").value).toUpperCase();
      strDomainNameToParse = modDataConversion.fnZLSIfNull(prstIn("TYPE_NAME").value).toUpperCase();

      // ~~~~~~~~~~~~~ Set default value, if applicable ~~~~~~~~~~~~
      if (prstIn("COLUMN_DEF").value == null == false) {
        pudtCol.hasDefault = true;
      } 
      else {
        pudtCol.hasDefault = false;
      }
      // If column has a default value, grab it. The COLUMN_DEF column is absent if
      // there is no default
      if (pudtCol.hasDefault) {
        // Char (not VarChar) fields may have a default value that is narrower than the column's width.
        // This is okay in a SQL environment, because SQL will pad it with trailing spaces
        // when doing an Inserts. Our GUI, however, doesn't like having a default value
        // narrower than the column width if the column in question is used as the selectable
        // column in a ComboBox. (The fnClearControls function fails with the equivalent
        // of the user selecting a value that's not in the list.) Avoid this by manually
        // adding trailing spaces here.
        if (pudtCol.dataType == adChar) {
          strEditedDefaultValue = prstIn("COLUMN_DEF").value;
          // Strip leading single quote
          if (strEditedDefaultValue.substring(0, 1).equals("'") && strEditedDefaultValue.length() > 1) {
            strEditedDefaultValue = strEditedDefaultValue.substring(strEditedDefaultValue.length() - strEditedDefaultValue.length() - 1);
          }
          // Strip trailing single quote
          if ("'".equals(strEditedDefaultValue.substring(strEditedDefaultValue.length() - 1)) && strEditedDefaultValue.length() > 1) {
            strEditedDefaultValue = strEditedDefaultValue.substring(0, strEditedDefaultValue.length() - 1);
          }
          pudtCol.defaultValue = modGeneral.fnPadRightString(strEditedDefaultValue, CInt(prstIn("LENGTH").value));
        } 
        else {
          pudtCol.defaultValue = prstIn("COLUMN_DEF").value;
        }
      } 
      else {
        // "default" DefaultValue value (cute, huh?)...may be overriden in next code chunks
        pudtCol.defaultValue = Empty;
      }

      // In some cases, the COLUMN_DEF value is not a value, per se,
      // but SQL text that indicates which Rule or Default should be applied.
      // If this is the case, then override the actual COLUMN_DEF value
      // with an interpreted value equivalent to what the DBMS would have set.

      // DEF_LST_UPD_USER is typically used for the Lst_Updt_User_Id column,
      // indicating to set it to the logged on user.
      if (strDefaultValueToParse.indexOf("DEF_LST_UPD_ID", 1) > 0) {
        pudtCol.defaultValue = modGeneral.gconAppActive.getLastLogOnUserID();
      }

      // DEF_LST_UPD_DTM is typically used for the Lst_Updt_Dtm column,
      // indicating to set it to the System Date.
      //!TODO! I think this default value is meaningless but harmless.
      //       The app shouldn't even reference this field on an INSERT to ensure that the
      //       DBMS sets it itself based on the exact date/time that the INSERT occurs.
      //       On an UPDATE statement, the app *should* (and *must*) reference this
      //       column to ensure it is updated, but it should be set by the form immediately
      //       prior to issuing the UPDATE.  If my thoughts are correct, maybe the
      //       following IF should be deleted.
      if (strDefaultValueToParse.indexOf("DEF_LST_UPD_DTM", 1) > 0) {
        pudtCol.defaultValue = Date;
      }

      // All Claims Interest tables use only the "DOM_IND" domain name for indicator columns.
      // This domain name indicates the column must be valued Y or N. Columns
      // bound to this domain name have a default constraint set so its default
      // value is "N".
      // For indicator columns, transform its default value from a literal "N" or "Y"
      // to its corresponding Boolean value since it will typically be represented
      // on forms as a checkbox.
      if ((strDomainNameToParse.indexOf("DOM_IND", 1) > 0)) {
        if ("Y"
.equals(pudtCol.defaultValue)) {
          pudtCol.defaultValue = true;
        } 
        else {
          pudtCol.defaultValue = false;
        }
      }

      if (prstIn("NULLABLE").value) {
        pudtCol.isNullable = true;
      } 
      else {
        pudtCol.isNullable = false;
      }

      // If the data model indicates, for instance, that a column's Numeric Scale
      // is (9,7), that means there are 9 numeric positions --excluding the decimal point--
      // of which 7 are decimal positions....i.e. "99.9999999"
      // Debug.Print .ColName
      switch (pudtCol.dataType) {
        case  dbDecimal:
          // Save original values (will be used to code sproc parameters).
          pudtCol.numericScale = CByte(prstIn("SCALE").value);
          pudtCol.precision = CByte(prstIn("PRECISION").value);
          // Save interpreted equivalents. These may be overriden in fnLoadColMetaData( ).
          pudtCol.decimalPositions = Integer.parseInt(prstIn("SCALE").value);
          pudtCol.dollarPositions = Integer.parseInt(prstIn("PRECISION").value) - pudtCol.decimalPositions;
          pudtCol.maxCharacters = 0;
          if ("ZIP_CD"
.equals(pudtCol.colName.substring(pudtCol.colName.length() - 7))) {
            // Area Codes require 5 numeric positions if input.
            // They should be displayed via a TextBox control.
            pudtCol.format = "&&&&&";
            pudtCol.mask = "";
            pudtCol.maxCharacters = 5;
          } 
          else if ("ZIP4_CD"
.equals(pudtCol.colName.substring(pudtCol.colName.length() - 7))) {
            // Area Codes require 4 numeric positions if input.
            // They should be displayed via a TextBox control.
            pudtCol.format = "&&&&";
            pudtCol.mask = "";
            pudtCol.maxCharacters = 4;
          } 
          else {
            // These should be displayed via a TextBox control.
            //
            // NOTE: Other numeric items will need to overridden
            // on a table-specific basis in fnLoadColMetaData( ) since there is no
            // easy way to recognize and process these fields. For instance,
            // amounts (ending in _AMT) could have a varying number of dollar or
            // decimal positions.
            pudtCol.format = String(pudtCol.dollarPositions, cstrNumericChar);
            if (pudtCol.decimalPositions > 0) {
              pudtCol.format = pudtCol.format+ "."+ String(pudtCol.decimalPositions, cstrNumericChar);
            }
            pudtCol.mask = "";
          }
          pudtCol.allowableCharacters = ".0123456789";
          pudtCol.shouldForceToUppercase = false;

          break;

        case  dbInteger:
          pudtCol.decimalPositions = 0;
          pudtCol.dollarPositions = Integer.parseInt(prstIn("PRECISION").value) - pudtCol.decimalPositions;
          pudtCol.maxCharacters = 0;
          pudtCol.format = String(pudtCol.dollarPositions, cstrNumericChar);
          pudtCol.mask = "";
          pudtCol.allowableCharacters = "0123456789";
          pudtCol.shouldForceToUppercase = false;

          break;

        case  dbChar:
        case  dbVarChar:
          pudtCol.decimalPositions = 0;
          pudtCol.dollarPositions = 0;
          pudtCol.maxCharacters = Integer.parseInt(prstIn("LENGTH").value);
          if ("SSN_NUM"
.equals(pudtCol.colName.substring(pudtCol.colName.length() - 7))) {
            // Social Security Numbers should be displayed via a MaskEdBox control.
            pudtCol.format = "";
            pudtCol.mask = "###-##-####";
          } 
          else {
            // All other adChar fields should be displayed via a TextBox control.
            pudtCol.format = String(pudtCol.maxCharacters, cstrAnyCharChar);
            pudtCol.mask = "";
          }
          pudtCol.allowableCharacters = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789@!#$%^&*()-+_=~:;.,<>\\|/?' ";
          pudtCol.shouldForceToUppercase = false;

          break;

        case  dbDateTime:
          pudtCol.decimalPositions = 0;
          pudtCol.dollarPositions = 0;
          pudtCol.maxCharacters = 0;
          // These should typically be displayed via  DTPicker control.
          // For a Grid,     .Format should be "MM/DD/YYYY"
          // For a DTPicker, .Format should be "MM/dd/yyy"
          // For a TextBox, .Format should be "mm/dd/yyyy"
          pudtCol.format = "MM/dd/yyy";
          pudtCol.mask = "";
          pudtCol.allowableCharacters = "0123456789/-";
          pudtCol.shouldForceToUppercase = false;
          break;

        default:
          modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_UNEXPECTED_VAL_SELECT_CASE, MCSTRNAME+ cstrCurrentProc);
          // **TODO:** goto found: GoTo PROC_EXIT;
          break;
      }

      _rtn = true;
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  private udtColumn fnGetProperty(String strTagIn) {
    udtColumn _rtn = null;
    //--------------------------------------------------------------------------
    // Procedure:   fnGetTableProperty
    // Description: Given the tag name, it returns a pointer to the specified
    //              table class's public property.
    //
    // Params:      N/A
    //    strTagIn  (in)     A string containing the Property Name, typically
    //                       from the form control's Tag property
    //
    // Returns:     A pointer to the public property
    //-----------------------------------------------------------------------------
    // Modified:     Berry Kropiwka - Added Compact Calc - 2019-09-27

    //!CUSTOMIZE!  There should be one Case statement for each table column.
    //             Each Case statement should reference a class property that
    //             corresponds to a table column, and the fnGetProperty return
    //             value should be set to the private variable (of type UDTColumn)
    //             that corresponds to that table column/class property.

    "fnGetProperty"
.equals(Const cstrCurrentProc As String);

    try {

      strTagIn = strTagIn.toUpperCase();
      switch (strTagIn) {
        case  "ADMNSYSTCD":
          _rtn = m_strAdmnSystCd;
          break;

        case  "CLMID":
          _rtn = m_lngClmId;
          break;

        case  "CLMINSDDTHDT":
          _rtn = m_dteClmInsdDthDt;
          break;

        case  "CLMINSDFIRSTNM":
          _rtn = m_strClmInsdFirstNm;
          break;

        case  "CLMINSDLASTNM":
          _rtn = m_strClmInsdLastNm;
          break;

        case  "CLMINSDSSNNUM":
          _rtn = m_strClmInsdSsnNum;
          break;

        case  "CLMNUM":
          _rtn = m_strClmNum;
          break;

        case  "CLMPOLNUM":
          _rtn = m_strClmPolNum;
          break;

        case  "CLMPROOFDT":
          _rtn = m_dteClmProofDt;
          break;

        case  "CLMTOTCLMPDAMT":
          _rtn = m_varClmTotClmPdAmt;
          break;

        case  "CLMTOTDTHBPMTAMT":
          _rtn = m_varClmTotDthbPmtAmt;
          break;

        case  "CLMTOTINTAMT":
          _rtn = m_varClmTotIntAmt;
          break;

        case  "CLMTOTWTHLDAMT":
          _rtn = m_varClmTotWthldAmt;
          break;

        case  "INSDDTHRESSTCD":
          _rtn = m_strInsdDthResStCd;
          break;

        case  "ISSSTCD":
          _rtn = m_strIssStCd;
          break;

        case  "LSTUPDTDTM":
          _rtn = m_dteLstUpdtDtm;
          break;

        case  "LSTUPDTUSERID":
          _rtn = m_strLstUpdtUserId;
          break;

        case  "PYCOTYPCD":
          _rtn = m_strPycoTypCd;
          break;

        case  "CLMFORRESDTHIND":
          _rtn = m_bClmForResDthInd;
          break;

        case  "CLMCOMPACTCLCNIND":
          _rtn = m_bClmCompactClcnInd;
          break;

        default:
          modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_UNEXPECTED_VAL_SELECT_CASE, MCSTRNAME+ cstrCurrentProc);
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
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  private boolean fnLoadColMetaData() {
    boolean _rtn = false;
    //--------------------------------------------------------------------------
    // Procedure:   fnLoadColMetaData
    // Description: For each table column, collect its meta data
    //              and load it to its corresponding UDT.
    //
    //              NOTE: Here is where you would override that meta data if it's
    //                    warranted for a given column. For instance, a column
    //                    that holds a numeric column that is allowed to have a
    //                    negative value should have its default
    //                    a. .AllowableCharacters property overriden to allow a "-" sign
    //                    b. .Format property overriden to specify that a negative
    //                       value should be enclosed within parentheses.
    //
    // See fnGetColMetaData( ) in this module to see what the defaults are set to
    // for a given data type.

    //
    // Returns:     True if successful; False otherwise
    // Params:      None
    // Date:        01/07/2002
    //-----------------------------------------------------------------------------
    // Modified:     Berry Kropiwka - Added Compact Calc - 2019-09-27

    //!CUSTOMIZE!  The call to GetMetaData_Column should pass in the correct table name.
    //
    //             In the first Select Case block,
    //             * there should be one Case statement for each table column.
    //             * Each Case statement should reference the constant that
    //               corresponds to the table column, and it should pass the private
    //               variable (of type UDTColumn) that corresponds to that
    //               table column/class property to fnGetColMetaData.
    //
    //             In the second Select Case block (getting info re: primary keys),
    //             * there should be one Case statement for each table column.
    //             * Each Case statement should reference the constant that
    //               corresponds to the table column, and it should call the IsKey
    //               method of the priva variable (of type UDTColumn) that
    //               corresponds to that table column/class property.


    "fnLoadColMetaData"
.equals(Const cstrCurrentProc As String);
    DBRecordSet rstMetaData = null;

    try {

      //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
      //       Get meta data, like nullability, data type, default value, etc.
      //       and set override values
      //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
      m_adwADO.getMetaData_Columns("claim_t", rstMetaData);

      // NOTES:
      //    a. Phone Numbers and SSNs are Char fields, but we have to override the meta data to
      //       ensure only numbers can be input to them.
      //    b. If the GUI is imposing a default value where the DBMS does not have one,
      //       you must set .HasDefault to True and *then* set the .DefaultValue value.
      //
      //       WARNING: Char (not VarChar) fields may have a default value that is narrower than the column's width.
      //                This is okay in a SQL environment, because SQL will pad it with trailing spaces
      //                when doing an Inserts. Our GUI, however, doesn't like having a default value
      //                narrower than the column width if the column in question is used as the selectable
      //                column in a ComboBox. (The fnClearControls function fails with the equivalent
      //                of the user selecting a value that's not in the list.) Avoid this by making sure
      //                any override default value you specify here is as wide as the column to which
      //                it should be applied.

      do Until .EOF        // Make sure SELECT CASE lists all table columns, including the LST_UPDT_xxx!
        //Debug.Print UCase$(rstMetaData("COLUMN_NAME").value)
        switch (rstMetaData("COLUMN_NAME").value.toUpperCase()) {
          case  MCSTRADMNSYSTCD:
            fnGetColMetaData(m_strAdmnSystCd, rstMetaData);
            break;

          case  MCSTRCLMID:
            fnGetColMetaData(m_lngClmId, rstMetaData);
            break;

          case  MCSTRCLMINSDDTHDT:
            fnGetColMetaData(m_dteClmInsdDthDt, rstMetaData);
            break;

          case  MCSTRCLMINSDFIRSTNM:
            fnGetColMetaData(m_strClmInsdFirstNm, rstMetaData);
            m_strClmInsdFirstNm.shouldForceToUppercase = true;
            break;

          case  MCSTRCLMINSDLASTNM:
            fnGetColMetaData(m_strClmInsdLastNm, rstMetaData);
            m_strClmInsdLastNm.shouldForceToUppercase = true;
            break;

          case  MCSTRCLMINSDSSNNUM:
            fnGetColMetaData(m_strClmInsdSsnNum, rstMetaData);
            m_strClmInsdSsnNum.allowableCharacters = "0123456789";
            break;

          case  MCSTRCLMNUM:
            fnGetColMetaData(m_strClmNum, rstMetaData);
            m_strClmNum.shouldForceToUppercase = true;
            break;

          case  MCSTRCLMPOLNUM:
            fnGetColMetaData(m_strClmPolNum, rstMetaData);
            m_strClmPolNum.shouldForceToUppercase = true;
            break;

          case  MCSTRCLMPROOFDT:
            fnGetColMetaData(m_dteClmProofDt, rstMetaData);
            break;

          case  MCSTRCLMTOTCLMPDAMT:
            fnGetColMetaData(m_varClmTotClmPdAmt, rstMetaData);
            break;

          case  MCSTRCLMTOTDTHBPMTAMT:
            fnGetColMetaData(m_varClmTotDthbPmtAmt, rstMetaData);
            break;

          case  MCSTRCLMTOTINTAMT:
            fnGetColMetaData(m_varClmTotIntAmt, rstMetaData);
            break;

          case  MCSTRCLMTOTWTHLDAMT:
            fnGetColMetaData(m_varClmTotWthldAmt, rstMetaData);
            break;

          case  MCSTRINSDDTHRESSTCD:
            fnGetColMetaData(m_strInsdDthResStCd, rstMetaData);
            break;

          case  MCSTRISSSTCD:
            fnGetColMetaData(m_strIssStCd, rstMetaData);
            break;

          case  MCSTRLSTUPDTDTM:
            fnGetColMetaData(m_dteLstUpdtDtm, rstMetaData);
            break;

          case  MCSTRLSTUPDTUSERID:
            fnGetColMetaData(m_strLstUpdtUserId, rstMetaData);
            break;

          case  MCSTRPYCOTYPCD:
            fnGetColMetaData(m_strPycoTypCd, rstMetaData);
            break;

          case  MCSTRCLMFORRESDTHIND:
            fnGetColMetaData(m_bClmForResDthInd, rstMetaData);
            break;

          case  MCSTRCLMCOMPACTCLCNIND:
            fnGetColMetaData(m_bClmCompactClcnInd, rstMetaData);
            break;

          default:
            modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_UNEXPECTED_VAL_SELECT_CASE, MCSTRNAME+ cstrCurrentProc);
            // **TODO:** goto found: GoTo PROC_EXIT;
            break;
        }
        m_adwADO.moveNext(rstMetaData);
      }
      //' Close now so the recordset can be reused
      rstMetaData.Close;

      //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
      //             Get meta data concerning which columns are key fields
      //
      //     If a given COLUMN_NAME is returned in the recordset, it is a primary key.
      //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
      m_adwADO.getMetaData_PrimaryKeys("claim_t", rstMetaData);

      do Until .EOF        // The SELECT CASE should list all table columns
        // (though you could skip the LST_UPDT_xxx columns if you change the Case Else,
        //  since these would never be a key)
        switch (rstMetaData("COLUMN_NAME").value.toUpperCase()) {
          case  MCSTRADMNSYSTCD:
            m_strAdmnSystCd.isKey = true;
            break;

          case  MCSTRCLMID:
            m_lngClmId.isKey = true;
            break;

          case  MCSTRCLMINSDDTHDT:
            m_dteClmInsdDthDt.isKey = true;
            break;

          case  MCSTRCLMINSDFIRSTNM:
            m_strClmInsdFirstNm.isKey = true;
            break;

          case  MCSTRCLMINSDLASTNM:
            m_strClmInsdLastNm.isKey = true;
            break;

          case  MCSTRCLMINSDSSNNUM:
            m_strClmInsdSsnNum.isKey = true;
            break;

          case  MCSTRCLMNUM:
            m_strClmNum.isKey = true;
            break;

          case  MCSTRCLMPOLNUM:
            m_strClmPolNum.isKey = true;
            break;

          case  MCSTRCLMPROOFDT:
            m_dteClmProofDt.isKey = true;
            break;

          case  MCSTRCLMTOTCLMPDAMT:
            m_varClmTotClmPdAmt.isKey = true;
            break;

          case  MCSTRCLMTOTDTHBPMTAMT:
            m_varClmTotDthbPmtAmt.isKey = true;
            break;

          case  MCSTRCLMTOTINTAMT:
            m_varClmTotIntAmt.isKey = true;
            break;

          case  MCSTRCLMTOTWTHLDAMT:
            m_varClmTotWthldAmt.isKey = true;
            break;

          case  MCSTRINSDDTHRESSTCD:
            m_strInsdDthResStCd.isKey = true;
            break;

          case  MCSTRISSSTCD:
            m_strIssStCd.isKey = true;
            break;

          case  MCSTRLSTUPDTDTM:
            m_dteLstUpdtDtm.isKey = true;
            break;

          case  MCSTRLSTUPDTUSERID:
            m_strLstUpdtUserId.isKey = true;
            break;

          case  MCSTRPYCOTYPCD:
            m_strPycoTypCd.isKey = true;
            break;

          case  MCSTRCLMFORRESDTHIND:
            m_bClmForResDthInd.isKey = true;
            break;

          case  MCSTRCLMCOMPACTCLCNIND:
            m_bClmCompactClcnInd.isKey = true;
            break;

          default:
            modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_UNEXPECTED_VAL_SELECT_CASE, MCSTRNAME+ cstrCurrentProc);
            // **TODO:** goto found: GoTo PROC_EXIT;
            break;
        }
        m_adwADO.moveNext(rstMetaData);
      }
      //' Close now so the recordset can be reused
      rstMetaData.Close;

      _rtn = true;
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    modGeneral.fnFreeRecordset(rstMetaData);

    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  private DBRecordSet fnSelectRecord(int lngKey1) {
    //--------------------------------------------------------------------------
    // Procedure:   fnSelectRecord
    // Description: Selects a single record based on the value(s) in the
    //              properties that correspond to the table's key(s)
    //
    //              NOTE: For each table key, there should be a parameter
    //                    of the appropriate data type!
    //
    // Parameters:
    //     lngKey1 (in) - the key to the table that should be retrieved
    //
    // Returns:     A disconnected ADODB.Recordset containing all table columns
    //              for the specified key
    //-----------------------------------------------------------------------------

    //!CUSTOMIZE!  This proc and all calls to it must be customized to reflect
    //             one parameter for each key column of the table. Make sure the
    //             parameter is defined to be of the right data type. Also,
    //             the way the recordset's .Find property is set must be changed
    //             to reflect each key column so the right record will be located.
    //             Also make sure that the substitution values passed to
    //             SaveAppSpecificError are correct and TRIM'd if appropriate.

    "fnSelectRecord"
.equals(Const cstrCurrentProc As String);
    //' Stored procedure to execute
    "dbo.proc_claim_select"
.equals(Const cstrSproc As String);
    DBRecordSet rstTemp = null;
    ADODB.Parameter prmReturnValue = null;
    ADODB.Parameter prmClmId = null;

    try {

      if (!(m_adwADO.commandSetSproc(cstrSproc))) {
        // **TODO:** goto found: GoTo PROC_EXIT;
      }

      // For Char/VarChar fields,
      //     * Use fnNullIfZLS to ensure Nulls are appropriately handled.
      //     * Do *not* set the optional 2nd parameter to fnNullIfZLS to True.
      // For numeric fields,
      //     * Use fnNullIfZero to ensure Nulls are appropriately handled.
      // For Y/N fields,
      //     * Use fnBoolToYN to ensure True/False is appropriately translated.

      ADODB.Command w_aDOCommand = m_adwADO.getADOCommand();
      // ---Parameter #1---
      // Define the return value that represents the error code (i.e. reason) why
      // the stored procedure failed.
      prmReturnValue = w_aDOCommand.CreateParameter(Name:="@return_value", Type:=adInteger, Direction:=adParamReturnValue, .value:=Null);
      w_aDOCommand.Parameters.Append(prmReturnValue);

      // ---Parameter #2---
      // Define the CLM_ID parameter
      prmClmId = w_aDOCommand.CreateParameter(Name:="@clm_id", Type:=adInteger, Direction:=adParamInput, .value:=fnNullIfZero(lngKey1));
      w_aDOCommand.Parameters.Append(prmClmId);

      rstTemp = w_aDOCommand.Execute();

      rstTemp.ActiveConnection = null;
      return rstTemp;
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here

    // Do *not* do "fnFreeRecordset rstTemp" since this will cause the recordset returned
    // by this function to be wiped out as well!
    modGeneral.fnFreeObject(prmReturnValue);
    modGeneral.fnFreeObject(prmClmId);

    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return null;
    // **TODO:** label found: PROC_ERR:;
    switch (prmReturnValue) {
      case  modResConstants.gCRES_NERR_REC_NOT_FOUND:
        // 4027 = The specified record was not found in the database (@@1).
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_REC_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, "Claim ID "+ RTrim$(lngKey1));
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
        break;

      case  modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO:
        // 4028 = An error occurred while attempting to @@1 this record.
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO, MCSTRNAME+ cstrCurrentProc, "locate");
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
        break;
    }

    // If any other errors exist, i.e. in Err object, then let it fall through into default error handling.

    switch (VBA.ex.Number) {
      //' Object not found
      case  -2147217900:
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_FERR_SPROC_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, cstrSproc);
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
        break;

      default:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
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


case class TclmclaimData(
              id: Option[Int],

              )

object Tclmclaims extends Controller with ProvidesUser {

  val tclmclaimForm = Form(
    mapping(
      "id" -> optional(number),

  )(TclmclaimData.apply)(TclmclaimData.unapply))

  implicit val tclmclaimWrites = new Writes[Tclmclaim] {
    def writes(tclmclaim: Tclmclaim) = Json.obj(
      "id" -> Json.toJson(tclmclaim.id),
      C.ID -> Json.toJson(tclmclaim.id),

    )
  }

  def get(id: Int) = GetAction { implicit request =>
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.LIST_TCLMCLAIM), { user =>
      Ok(Json.toJson(Tclmclaim.get(user, id)))
    })
  }

  def update(id: Int) = PostAction { implicit request =>
    Logger.debug("in tclmclaims.update")
    tclmclaimForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      tclmclaim => {
        Logger.debug(s"form: ${tclmclaim.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.EDIT_TCLMCLAIM), { user =>
          Ok(
            Json.toJson(
              Tclmclaim.update(user,
                Tclmclaim(
                       id,

                ))))
        })
      }
    )
  }

  def create = PostAction { implicit request =>
    Logger.debug("in tclmclaims.create")
    tclmclaimForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      tclmclaim => {
        Logger.debug(s"form: ${tclmclaim.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.NEW_TCLMCLAIM), { user =>
          Ok(
            Json.toJson(
              Tclmclaim.create(user,
                Tclmclaim(

                ))))
        })
      }
    )
  }

  def delete(id: Int) = PostAction { implicit request =>
    Logger.debug("in tclmclaims.delete")
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.DELETE_TCLMCLAIM), { user =>
      Tclmclaim.delete(user, id)
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

case class Tclmclaim(
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

object Tclmclaim {

  lazy val emptyTclmclaim = Tclmclaim(
)

  def apply(
      id: Int,
) = {

    new Tclmclaim(
      id,
)
  }

  def apply(
) = {

    new Tclmclaim(
)
  }

  private val tclmclaimParser: RowParser[Tclmclaim] = {
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
        Tclmclaim(
              id,
,
              createdAt,
              updatedAt,
              updatedBy)
    }
  }

  def create(user: CompanyUser, tclmclaim: Tclmclaim): Tclmclaim = {
    save(user, tclmclaim, true)
  }

  def update(user: CompanyUser, tclmclaim: Tclmclaim): Tclmclaim = {
    save(user, tclmclaim, false)
  }

  private def save(user: CompanyUser, tclmclaim: Tclmclaim, isNew: Boolean): Tclmclaim = {
    def getFields = {
      List(

      )
    }
    def throwException = {
      throw new RuntimeException(s"Error when saving ${C.TCLMCLAIM}")
    }

    DBHelper.saveEx(
      user,
      Register(
        C.TCLMCLAIM,
        C.ID,
        tclmclaim.id,
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

  def load(user: CompanyUser, id: Int): Option[Tclmclaim] = {
    loadWhere(user, s"${C.ID} = {id}", 'id -> id)
  }

  def loadWhere(user: CompanyUser, where: String, args : scala.Tuple2[scala.Any, anorm.ParameterValue[_]]*) = {
    DB.withConnection(user.database.database) { implicit connection =>
      SQL(s"SELECT t1.*, t2.${C.FK_NAME} FROM ${C.TCLMCLAIM} t1 INNER JOIN ${C.???} t2 ON t1.${C.FK_ID} = t2.${C.FK_ID} WHERE $where")
        .on(args: _*)
        .as(tclmclaimParser.singleOpt)
    }
  }

  def delete(user: CompanyUser, id: Int) = {
    DB.withConnection(user.database.database) { implicit connection =>
      try {
        SQL(s"DELETE FROM ${C.TCLMCLAIM} WHERE ${C.ID} = {id}")
        .on('id -> id)
        .executeUpdate
      } catch {
        case NonFatal(e) => {
          Logger.error(s"can't delete a ${C.TCLMCLAIM}. ${C.ID} id: $id. Error ${e.toString}")
          throw e
        }
      }
    }
  }

  def get(user: CompanyUser, id: Int): Tclmclaim = {
    load(user, id) match {
      case Some(p) => p
      case None => emptyTclmclaim
    }
  }
}


// Router

GET     /api/v1/general/tclmclaim/:id              controllers.logged.modules.general.Tclmclaims.get(id: Int)
POST    /api/v1/general/tclmclaim                  controllers.logged.modules.general.Tclmclaims.create
PUT     /api/v1/general/tclmclaim/:id              controllers.logged.modules.general.Tclmclaims.update(id: Int)
DELETE  /api/v1/general/tclmclaim/:id              controllers.logged.modules.general.Tclmclaims.delete(id: Int)




/**/
