
import java.util.Date;

public class ctpyePayee {

  //--------------------------------------------------------------------------
  // Procedure:   ctpyePayee
  // Description: Provides properties and methods to support the payee_t table values
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
  //
  //
  // Procedures  :
  //   Private     Class_Initialize()
  //   Private     Class_Terminate()
  //   Public      Property Get AllowableCharacters(ByVal strTagIn As String) As String
  //   Public      Property Get CalcStCd() As String
  //   Public      Property Let CalcStCd(ByVal strValue As String)
  //   Public      Property Get ClmId() As Long
  //   Public      Property Let ClmId(ByVal lngValue As Long)
  //   Public      Property Get CurrentLookupRecordNumber() As Long
  //   Public      Property Get DecimalPositions(ByVal strTagIn As String) As Integer
  //   Public      Property Get DefaultValue(ByVal strTagIn As String) As Variant
  //   Public      Property Get DollarPositions(ByVal strTagIn As String) As Integer
  //   Public      Property Get Format(ByVal strTagIn As String) As String
  //   Public      InitPayee(lngClmIdIn As Long)
  //   Public      Property Get IsKey(ByVal strTagIn As String) As Boolean
  //   Public      Property Get IsNullable(ByVal strTagIn As String) As Boolean
  //   Public      Property Get LookupData() As ADODB.Recordset
  //   Public      Property Get LookupData_Name() As Variant
  //   Public      Property Get LookupIsAtBOF() As Boolean
  //   Public      Property Get LookupIsAtEOF() As Boolean
  //   Public      Property Get LookupRecordCount() As Long
  //   Public      Property Get LstUpdtDtm() As Date
  //   Public      Property Let LstUpdtDtm(ByVal NewValue As String)
  //   Public      Property Get LstUpdtUserId() As String
  //   Public      Property Let LstUpdtUserId(ByVal strValue As String)
  //   Public      Property Get Mask(ByVal strTagIn As String) As String
  //   Public      Property Get MaxCharacters(ByVal strTagIn As String) As Long
  //   Public      Property Get PayeAddrLn1Txt() As String
  //   Public      Property Let PayeAddrLn1Txt(ByVal strValue As String)
  //   Public      Property Get PayeAddrLn2Txt() As String
  //   Public      Property Let PayeAddrLn2Txt(ByVal strValue As String)
  //   Public      Property Get PayeCareOfTxt() As String
  //   Public      Property Let PayeCareOfTxt(ByVal strValue As String)
  //   Public      Property Get PayeCityNmTxt() As String
  //   Public      Property Let PayeCityNmTxt(ByVal strValue As String)
  //   Public      Property Get PayeClmIntAmt() As Double
  //   Public      Property Let PayeClmIntAmt(ByVal dblValue As Double)
  //   Public      Property Get PayeClmIntRt() As Double
  //   Public      Property Let PayeClmIntRt(ByVal dblValue As Double)
  //   Public      Property Get PayeClmPdAmt() As Double
  //   Public      Property Let PayeClmPdAmt(ByVal dblValue As Double)
  //   Public      Property Get PayeDfltOvrdInd() As Boolean
  //   Public      Property Let PayeDfltOvrdInd(ByVal bValue As Boolean)
  //   Public      Property Get PayeDthbPmtAmt() As Double
  //   Public      Property Let PayeDthbPmtAmt(ByVal dblValue As Double)
  //   Public      Property Get PayeFullNm() As String
  //   Public      Property Let PayeFullNm(ByVal strValue As String)
  //   Public      Property Get PayeId() As Long
  //   Public      Property Let PayeId(ByVal lngValue As Long)
  //   Public      Property Get PayeIntDaysPdNum() As Integer
  //   Public      Property Let PayeIntDaysPdNum(ByVal intValue As Integer)
  //   Public      Property Get PayePmtDt() As Date
  //   Public      Property Let PayePmtDt(ByVal dteValue As Date)
  //   Public      Property Get PayeSsnTinNum() As String
  //   Public      Property Let PayeSsnTinNum(ByVal strValue As String)
  //   Public      Property Get PayeSsnTinTypCd() As String
  //   Public      Property Let PayeSsnTinTypCd(ByVal strValue As String)
  //   Public      Property Get PayeStCd() As String
  //   Public      Property Let PayeStCd(ByVal strValue As String)
  //   Public      Property Get PayeWthldAmt() As Double
  //   Public      Property Let PayeWthldAmt(ByVal dblValue As Double)
  //   Public      Property Get PayeWthldRt() As Double
  //   Public      Property Let PayeWthldRt(ByVal dblValue As Double)
  //   Public      Property Get PayeZip4Cd() As String
  //   Public      Property Let PayeZip4Cd(ByVal strValue As String)
  //   Public      Property Get PayeZipCd() As String
  //   Public      Property Let PayeZipCd(ByVal strValue As String)
  //   Public      Property Get ShouldForceToUppercase(ByVal strTagIn As String) As Boolean
  //   Public      AddRecord() as Boolean
  //   Public      CheckForAnotherUsersChanges(ByVal lngWhatOperation As enumWhatOperationIsBeingAttempted, _
  //                   ByRef strACF2 As String) As Long
  //   Public      DeleteRecord() As Boolean
  //   Public      GetClmNumFromClmID(ByVal lngClmID As Long) As Variant
  //   Public      GetLookupData() As Boolean
  //   Public      GetPayeFullNmFromPayeID(ByVal lngPayeID As Long) As Variant
  //   Public      GetPayeesForClaim(ByVal lngClmId As Long) As ADODB.Recordset
  //   Public      GetClaimForPayeeClaim(ByVal lngClmId As Long) As ADODB.Recordset
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
  private static final String MCSTRNAME = "ctpyePayee.";

  //...............................................................................................
  //!CUSTOMIZE!
  // These are the private variables corresponding to PUBLIC properties.
  // There should be one (of type udtColumn) for each column in the table that this class accesses.
  //...............................................................................................
  private udtColumn m_strCalcStCd;
  private udtColumn m_lngClmId;
  private udtColumn m_dteLstUpdtDtm;
  private udtColumn m_strLstUpdtUserId;
  private udtColumn m_strPayeAddrLn1Txt;
  private udtColumn m_strPayeAddrLn2Txt;
  private udtColumn m_strPayeCareOfTxt;
  private udtColumn m_strPayeCityNmTxt;
  private udtColumn m_dblPayeClmIntAmt;
  private udtColumn m_dblPayeClmIntRt;
  private udtColumn m_dblPayeClmPdAmt;
  private udtColumn m_bPayeDfltOvrdInd;
  //'' BZ4999 October 2013 Non US payee - SXS
  private udtColumn m_bstrPaye1099INTInd;
  private udtColumn m_dblPayeDthbPmtAmt;
  private udtColumn m_strPayeFullNm;
  private udtColumn m_lngPayeId;
  private udtColumn m_intPayeIntDaysPdNum;
  private udtColumn m_dtePayePmtDt;
  private udtColumn m_strPayeSsnTinNum;
  private udtColumn m_strPayeSsnTinTypCd;
  private udtColumn m_strPayeStCd;
  private udtColumn m_dblPayeWthldAmt;
  private udtColumn m_dblPayeWthldRt;
  private udtColumn m_strPayeZip4Cd;
  private udtColumn m_strPayeZipCd;


  //...............................................................................................
  //!CUSTOMIZE!
  // Create one Const for each column in the table, defining the table column to which it refers.
  //...............................................................................................
  private static final String MCSTRCALCSTCD = "CALC_ST_CD";
  private static final String MCSTRCLMID = "CLM_ID";
  private static final String MCSTRLSTUPDTDTM = "LST_UPDT_DTM";
  private static final String MCSTRLSTUPDTUSERID = "LST_UPDT_USER_ID";
  private static final String MCSTRPAYEADDRLN1TXT = "PAYE_ADDR_LN1_TXT";
  private static final String MCSTRPAYEADDRLN2TXT = "PAYE_ADDR_LN2_TXT";
  private static final String MCSTRPAYECAREOFTXT = "PAYE_CARE_OF_TXT";
  private static final String MCSTRPAYECITYNMTXT = "PAYE_CITY_NM_TXT";
  private static final String MCSTRPAYECLMINTAMT = "PAYE_CLM_INT_AMT";
  private static final String MCSTRPAYECLMINTRT = "PAYE_CLM_INT_RT";
  private static final String MCSTRPAYECLMPDAMT = "PAYE_CLM_PD_AMT";
  private static final String MCSTRPAYEDFLTOVRDIND = "PAYE_DFLT_OVRD_IND";
  //'' BZ4999 October 2013 Non US payee - SXS
  private static final String MCSTRPAYE1099INTIND = "PAYE_1099INT_IND";
  private static final String MCSTRPAYEDTHBPMTAMT = "PAYE_DTHB_PMT_AMT";
  private static final String MCSTRPAYEFULLNM = "PAYE_FULL_NM";
  private static final String MCSTRPAYEID = "PAYE_ID";
  private static final String MCSTRPAYEINTDAYSPDNUM = "PAYE_INT_DAYS_PD_NUM";
  private static final String MCSTRPAYEPMTDT = "PAYE_PMT_DT";
  private static final String MCSTRPAYESSNTINNUM = "PAYE_SSN_TIN_NUM";
  private static final String MCSTRPAYESSNTINTYPCD = "PAYE_SSN_TIN_TYP_CD";
  private static final String MCSTRPAYESTCD = "PAYE_ST_CD";
  private static final String MCSTRPAYEWTHLDAMT = "PAYE_WTHLD_AMT";
  private static final String MCSTRPAYEWTHLDRT = "PAYE_WTHLD_RT";
  private static final String MCSTRPAYEZIP4CD = "PAYE_ZIP4_CD";
  private static final String MCSTRPAYEZIPCD = "PAYE_ZIP_CD";


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
  public String getCalcStCd() {
    String _rtn = "";
    // **************************************************************************
    // Function  : Property Get CalcStCd
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : string
    // **************************************************************************
    "Property Get CalcStCd"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = CStr(m_strCalcStCd.value);
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
  public void setCalcStCd(String strValue) {
    String _rtn = null;
    // **************************************************************************
    // Function  : Property Let CalcStCd
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal strValue As String
    // Returns   :
    // **************************************************************************
    "Property Let CalcStCd"
.equals(Const cstrCurrentProc As String);
    try {

      m_strCalcStCd.value = strValue;
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
  public int getClmId() {
    int _rtn = 0;
    // **************************************************************************
    // Function  : Property Get ClmId
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : Long
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
    //             that might be necessary, such as turning a
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

      //!TODO! Consider whether another variation of this procedure should
      //       be created...where it accepts a parameter of type Control
      //       which it can then set if the column has a default value
      //       and do nothing to if the column has no default value.

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
  public Object getLookupData_Name() {
    Object _rtn = null;
    //--------------------------------------------------------------------------
    // Procedure:   Get LookupData_Name
    // Description: Return an array containing just the desired columns
    //              that should be populated in the Lookup fpCombo control
    // Returns:     variant array
    //-----------------------------------------------------------------------------
    try {
      "Get_LookupData_Name"
.equals(Const cstrCurrentProc As String);
      Object[] aRows() = null;
      String strOriginalPayeFullNum = "";

      // The .GetRows method changes the positioning within the m_rstLookup, which
      // causes the "record x of y" label to incorrectly get set to the
      // .RecordCount value. So, save the key of the current record, issue the
      // .GetRows and then navigate back to the original position within the rst.
      strOriginalPayeFullNum = m_rstLookup.Fields(MCSTRPAYEFULLNM).value;

      // m_rstLookup is already sorted in the order we want: clm_num
      aRows = m_rstLookup.GetRows(Rows:=adGetRowsRest, Start:=adBookmarkFirst);

      ctclmClaim.getRelativeRecord(strOriginalPayeFullNum, enumPositionDirection.ePDSAMERECORD);
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
    // Returns   : Date
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
    // Parameters: ByVal dteValue As Date
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
    // Returns   : String
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
    // Parameters: ByVal strValue As String
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
  public String getPayeAddrLn1Txt() {
    String _rtn = "";
    // **************************************************************************
    // Function  : Property Get PayeAddrLn1Txt
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : string
    // **************************************************************************
    "Property Get PayeAddrLn1Txt"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = CStr(m_strPayeAddrLn1Txt.value);
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
  public void setPayeAddrLn1Txt(String strValue) {
    String _rtn = null;
    // **************************************************************************
    // Function  : Property Let PayeAddrLn1Txt
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal strValue As String
    // Returns   :
    // **************************************************************************
    "Property Let PayeAddrLn1Txt"
.equals(Const cstrCurrentProc As String);
    try {

      m_strPayeAddrLn1Txt.value = strValue;
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
  public String getPayeAddrLn2Txt() {
    String _rtn = "";
    // **************************************************************************
    // Function  : Property Get PayeAddrLn2Txt
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : string
    // **************************************************************************
    "Property Get PayeAddrLn2Txt"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = CStr(m_strPayeAddrLn2Txt.value);
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
  public void setPayeAddrLn2Txt(String strValue) {
    String _rtn = null;
    // **************************************************************************
    // Function  : Property Let PayeAddrLn2Txt
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal strValue As String
    // Returns   :
    // **************************************************************************
    "Property Let PayeAddrLn2Txt"
.equals(Const cstrCurrentProc As String);
    try {

      m_strPayeAddrLn2Txt.value = strValue;
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
  public String getPayeCareOfTxt() {
    String _rtn = "";
    // **************************************************************************
    // Function  : Property Get PayeCareOfTxt
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : string
    // **************************************************************************
    "Property Get PayeCareOfTxt"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = CStr(m_strPayeCareOfTxt.value);
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
  public void setPayeCareOfTxt(String strValue) {
    String _rtn = null;
    // **************************************************************************
    // Function  : Property Let PayeCareOfTxt
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal strValue As String
    // Returns   :
    // **************************************************************************
    "Property Let PayeCareOfTxt"
.equals(Const cstrCurrentProc As String);
    try {

      m_strPayeCareOfTxt.value = strValue;
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
  public String getPayeCityNmTxt() {
    String _rtn = "";
    // **************************************************************************
    // Function  : Property Get PayeCityNmTxt
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : string
    // **************************************************************************
    "Property Get PayeCityNmTxt"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = CStr(m_strPayeCityNmTxt.value);
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
  public void setPayeCityNmTxt(String strValue) {
    String _rtn = null;
    // **************************************************************************
    // Function  : Property Let PayeCityNmTxt
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal strValue As String
    // Returns   :
    // **************************************************************************
    "Property Let PayeCityNmTxt"
.equals(Const cstrCurrentProc As String);
    try {

      m_strPayeCityNmTxt.value = strValue;
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
  public double getPayeClmIntAmt() {
    double _rtn = 0;
    // **************************************************************************
    // Function  : Property Get PayeClmIntAmt
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : double
    // **************************************************************************
    "Property Get PayeClmIntAmt"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = m_dblPayeClmIntAmt.value;
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
  public void setPayeClmIntAmt(double dblValue) {
    double _rtn = null;
    // **************************************************************************
    // Function  : Property Let PayeClmIntAmt
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal dblValue
    // **************************************************************************
    "Property Let PayeClmIntAmt"
.equals(Const cstrCurrentProc As String);
    try {

      m_dblPayeClmIntAmt.value = dblValue;
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
  public double getPayeClmIntRt() {
    double _rtn = 0;
    // **************************************************************************
    // Function  : Property Get PayeClmIntRt
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : double
    // **************************************************************************
    "Property Get PayeClmIntRt"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = m_dblPayeClmIntRt.value;
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
  public void setPayeClmIntRt(double dblValue) {
    double _rtn = null;
    // **************************************************************************
    // Function  : Property Let PayeClmIntRt
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal dblValue
    // **************************************************************************
    "Property Let PayeClmIntRt"
.equals(Const cstrCurrentProc As String);
    try {

      m_dblPayeClmIntRt.value = dblValue;
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
  public double getPayeClmPdAmt() {
    double _rtn = 0;
    // **************************************************************************
    // Function  : Property Get PayeClmPdAmt
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : double
    // **************************************************************************
    "Property Get PayeClmPdAmt"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = m_dblPayeClmPdAmt.value;
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
  public void setPayeClmPdAmt(double dblValue) {
    double _rtn = null;
    // **************************************************************************
    // Function  : Property Let PayeClmPdAmt
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal dblValue
    // **************************************************************************
    "Property Let PayeClmPdAmt"
.equals(Const cstrCurrentProc As String);
    try {

      m_dblPayeClmPdAmt.value = dblValue;
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
//''''''''''''''''''''''''''''''''''' '' BZ4999 October 2013 Non US payee - SXS
//////////////////////////////////////////////////////////////////////////////////////////////////
  public boolean getPaye1099INTInd() {
    boolean _rtn = false;
    // **************************************************************************
    // Function  : Property Get PayeDfltOvrdInd
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : Boolean
    // **************************************************************************
    "Property Get Paye1099INTInd"
.equals(Const cstrCurrentProc As String);
    try {

      // Note that this field is stored as a Y or N.
      _rtn = CBool(m_bstrPaye1099INTInd.value);
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
//' BZ4999 October 2013 Non US payee - SXS
  public void setPaye1099INTInd(boolean bValue) {
    boolean _rtn = null;
    // **************************************************************************
    // Function  : Property Let Paye1099INTInd
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal bValue As Boolean
    // Returns   :
    // **************************************************************************
    "Property Let Paye1099INTInd"
.equals(Const cstrCurrentProc As String);
    try {

      // Note that this field is stored as a Y or N.
      m_bstrPaye1099INTInd.value = bValue;
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
//'''''''''''''''''''''''''''''''''''


//////////////////////////////////////////////////////////////////////////////////////////////////
  public boolean getPayeDfltOvrdInd() {
    boolean _rtn = false;
    // **************************************************************************
    // Function  : Property Get PayeDfltOvrdInd
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : Boolean
    // **************************************************************************
    "Property Get PayeDfltOvrdInd"
.equals(Const cstrCurrentProc As String);
    try {

      // Note that this field is stored as a Y or N.
      _rtn = CBool(m_bPayeDfltOvrdInd.value);
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
  public void setPayeDfltOvrdInd(boolean bValue) {
    boolean _rtn = null;
    // **************************************************************************
    // Function  : Property Let PayeDfltOvrdInd
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal bValue As Boolean
    // Returns   :
    // **************************************************************************
    "Property Let PayeDfltOvrdInd"
.equals(Const cstrCurrentProc As String);
    try {

      // Note that this field is stored as a Y or N.
      m_bPayeDfltOvrdInd.value = bValue;
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
  public double getPayeDthbPmtAmt() {
    double _rtn = 0;
    // **************************************************************************
    // Function  : Property Get PayeDthbPmtAmt
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : double
    // **************************************************************************
    "Property Get PayeDthbPmtAmt"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = Double.parseDouble(m_dblPayeDthbPmtAmt.value);
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
  public void setPayeDthbPmtAmt(double dblValue) {
    double _rtn = null;
    // **************************************************************************
    // Function  : Property Let PayeDthbPmtAmt
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal dblValue As Double
    // Returns   :
    // **************************************************************************
    "Property Let PayeDthbPmtAmt"
.equals(Const cstrCurrentProc As String);
    try {

      m_dblPayeDthbPmtAmt.value = dblValue;
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
  public String getPayeFullNm() {
    String _rtn = "";
    // **************************************************************************
    // Function  : Property Get PayeFullNm
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : string
    // **************************************************************************
    "Property Get PayeFullNm"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = CStr(m_strPayeFullNm.value);
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
  public void setPayeFullNm(String strValue) {
    String _rtn = null;
    // **************************************************************************
    // Function  : Property Let PayeFullNm
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal strValue As String
    // Returns   :
    // **************************************************************************
    "Property Let PayeFullNm"
.equals(Const cstrCurrentProc As String);
    try {

      m_strPayeFullNm.value = strValue;
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
  public int getPayeId() {
    int _rtn = 0;
    // **************************************************************************
    // Function  : Property Get PayeId
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : Long
    // **************************************************************************
    "Property Get PayeId"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = Long.parseLong(m_lngPayeId.value);
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
  public void setPayeId(int lngValue) {
    int _rtn = null;
    // **************************************************************************
    // Function  : Property Let PayeId
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal lngValue As Long
    // Returns   :
    // **************************************************************************
    "Property Let PayeId"
.equals(Const cstrCurrentProc As String);
    try {

      m_lngPayeId.value = lngValue;
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
  public int getPayeIntDaysPdNum() {
    int _rtn = 0;
    // **************************************************************************
    // Function  : Property Get PayeIntDaysPdNum
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : Integer
    // **************************************************************************
    "Property Get PayeIntDaysPdNum"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = m_intPayeIntDaysPdNum.value;
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
  public void setPayeIntDaysPdNum(int intValue) {
    int _rtn = null;
    // **************************************************************************
    // Function  : Property Let PayeIntDaysPdNum
    // Purpose   : Assigns new Value (which could be NULL) to property
    // Parameters: ByVal intValue
    // **************************************************************************
    "Property Let PayeIntDaysPdNum"
.equals(Const cstrCurrentProc As String);
    try {

      m_intPayeIntDaysPdNum.value = intValue;
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
  public Date getPayePmtDt() {
    Date _rtn = null;
    // **************************************************************************
    // Function  : Property Get PayePmtDt
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : Date
    // **************************************************************************
    "Property Get PayePmtDt"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = G.parseDate(m_dtePayePmtDt.value);
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
  public void setPayePmtDt(Date dteValue) {
    Date _rtn = null;
    // **************************************************************************
    // Function  : Property Let PayePmtDt
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal dteValue As Date
    // Returns   :
    // **************************************************************************
    "Property Let PayePmtDt"
.equals(Const cstrCurrentProc As String);
    try {

      m_dtePayePmtDt.value = dteValue;
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
  public String getPayeSsnTinNum() {
    String _rtn = "";
    // **************************************************************************
    // Function  : Property Get PayeSsnTinNum
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : string
    // **************************************************************************
    "Property Get PayeSsnTinNum"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = CStr(m_strPayeSsnTinNum.value);
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
  public void setPayeSsnTinNum(String strValue) {
    String _rtn = null;
    // **************************************************************************
    // Function  : Property Let PayeSsnTinNum
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal strValue As String
    // Returns   :
    // **************************************************************************
    "Property Let PayeSsnTinNum"
.equals(Const cstrCurrentProc As String);
    try {

      m_strPayeSsnTinNum.value = strValue;
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
  public String getPayeSsnTinTypCd() {
    String _rtn = "";
    // **************************************************************************
    // Function  : Property Get PayeSsnTinTypCd
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : string
    // **************************************************************************
    "Property Get PayeSsnTinTypCd"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = CStr(m_strPayeSsnTinTypCd.value);
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
  public void setPayeSsnTinTypCd(String strValue) {
    String _rtn = null;
    // **************************************************************************
    // Function  : Property Let PayeSsnTinTypCd
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal strValue As String
    // Returns   :
    // **************************************************************************
    "Property Let PayeSsnTinTypCd"
.equals(Const cstrCurrentProc As String);
    try {

      m_strPayeSsnTinTypCd.value = strValue;
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
  public String getPayeStCd() {
    String _rtn = "";
    // **************************************************************************
    // Function  : Property Get PayeStCd
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : string
    // **************************************************************************
    "Property Get PayeStCd"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = CStr(m_strPayeStCd.value);

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
  public void setPayeStCd(String strValue) {
    String _rtn = null;
    // **************************************************************************
    // Function  : Property Let PayeStCd
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal strValue As String
    // Returns   :
    // **************************************************************************
    "Property Let PayeStCd"
.equals(Const cstrCurrentProc As String);
    try {

      m_strPayeStCd.value = strValue;
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
  public double getPayeWthldAmt() {
    double _rtn = 0;
    // **************************************************************************
    // Function  : Property Get PayeWthldAmt
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : Double
    // **************************************************************************
    "Property Get PayeWthldAmt"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = m_dblPayeWthldAmt.value;
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
  public void setPayeWthldAmt(double dblValue) {
    double _rtn = null;
    // **************************************************************************
    // Function  : Property Let PayeWthldAmt
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal dblValue
    // **************************************************************************
    "Property Let PayeWthldAmt"
.equals(Const cstrCurrentProc As String);
    try {

      m_dblPayeWthldAmt.value = dblValue;
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
  public double getPayeWthldRt() {
    double _rtn = 0;
    // **************************************************************************
    // Function  : Property Get PayeWthldRt
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : Double
    // **************************************************************************
    "Property Get PayeWthldRt"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = m_dblPayeWthldRt.value;
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
  public void setPayeWthldRt(double dblValue) {
    double _rtn = null;
    // **************************************************************************
    // Function  : Property Let PayeWthldRt
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal dblValue
    // **************************************************************************
    "Property Let PayeWthldRt"
.equals(Const cstrCurrentProc As String);
    try {

      m_dblPayeWthldRt.value = dblValue;
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
  public String getPayeZip4Cd() {
    String _rtn = "";
    // **************************************************************************
    // Function  : Property Get PayeZip4Cd
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : string
    // **************************************************************************
    "Property Get PayeZip4Cd"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = CStr(m_strPayeZip4Cd.value);
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
  public void setPayeZip4Cd(String strValue) {
    String _rtn = null;
    // **************************************************************************
    // Function  : Property Let PayeZip4Cd
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal strValue As String
    // Returns   :
    // **************************************************************************
    "Property Let PayeZip4Cd"
.equals(Const cstrCurrentProc As String);
    try {

      m_strPayeZip4Cd.value = strValue;
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
  public String getPayeZipCd() {
    String _rtn = "";
    // **************************************************************************
    // Function  : Property Get PayeZipCd
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : string
    // **************************************************************************
    "Property Get PayeZipCd"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = CStr(m_strPayeZipCd.value);
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
  public void setPayeZipCd(String strValue) {
    String _rtn = null;
    // **************************************************************************
    // Function  : Property Let PayeZipCd
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal strValue As String
    // Returns   :
    // **************************************************************************
    "Property Let PayeZipCd"
.equals(Const cstrCurrentProc As String);
    try {

      m_strPayeZipCd.value = strValue;
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
    // Returns   : Boolean
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

    //!CUSTOMIZE!  Customize the name of the stored procedure, the number and names
    //             of the parameters and, perhaps, the return values trapped in the
    //             error handler.

    "AddRecord"
.equals(Const cstrCurrentProc As String);
    //' Stored procedure to execute
    "dbo.proc_payee_insert"
.equals(Const cstrSproc As String);
    DBRecordSet rstTemp = null;
    boolean bSuccessful = false;
    ADODB.Parameter prmReturnValue = null;
    ADODB.Parameter prmInvalid_Key = null;
    ADODB.Parameter prmCalcStCd = null;
    ADODB.Parameter prmClmId = null;
    ADODB.Parameter prmPayeAddrLn1Txt = null;
    ADODB.Parameter prmPayeAddrLn2Txt = null;
    ADODB.Parameter prmPayeCareOfTxt = null;
    ADODB.Parameter prmPayeCityNmTxt = null;
    ADODB.Parameter prmPayeClmIntAmt = null;
    ADODB.Parameter prmPayeClmIntRt = null;
    ADODB.Parameter prmPayeClmPdAmt = null;
    ADODB.Parameter prmPayeDfltOvrdInd = null;
    ADODB.Parameter prmPayeDthbPmtAmt = null;
    ADODB.Parameter prmPayeFullNm = null;
    ADODB.Parameter prmPayeIntDaysPdNum = null;
    ADODB.Parameter prmPayePmtDt = null;
    ADODB.Parameter prmPayeSsnTinNum = null;
    ADODB.Parameter prmPayeSsnTinTypCd = null;
    ADODB.Parameter prmPayeStCd = null;
    ADODB.Parameter prmPayeWthldAmt = null;
    ADODB.Parameter prmPayeWthldRt = null;
    ADODB.Parameter prmPayeZip4Cd = null;
    ADODB.Parameter prmPayeZipCd = null;
    ADODB.Parameter prmNew_Id = null;
    //'' BZ4999 October 2013 Non US payee - SXS
    ADODB.Parameter prmPaye1099INTInd = null;

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
      // Define the CALC_ST_CD parameter
      prmCalcStCd = w_aDOCommand.CreateParameter(Name:="@calc_st_cd", Type:=adChar, Direction:=adParamInput, Size:=2, .value:=fnNullIfZLS(varIn:=CalcStCd, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmCalcStCd);

      // ---Parameter #3---
      // Define the CLM_ID parameter
      prmClmId = w_aDOCommand.CreateParameter(Name:="@clm_id", Type:=adInteger, Direction:=adParamInput, .value:=ClmId);
      w_aDOCommand.Parameters.Append(prmClmId);

      // ---Parameter #4---
      // Define the PAYE_ADR_LN1_TXT parameter
      prmPayeAddrLn1Txt = w_aDOCommand.CreateParameter(Name:="@paye_addr_ln1_txt", Type:=adVarChar, Direction:=adParamInput, Size:=40, .value:=fnNullIfZLS(varIn:=PayeAddrLn1Txt, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmPayeAddrLn1Txt);

      // ---Parameter #5---
      // Define the PAYE_ADR_LN2_TXT parameter
      prmPayeAddrLn2Txt = w_aDOCommand.CreateParameter(Name:="@paye_addr_ln2_txt", Type:=adVarChar, Direction:=adParamInput, Size:=40, .value:=fnNullIfZLS(varIn:=PayeAddrLn2Txt, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmPayeAddrLn2Txt);

      // ---Parameter #6---
      // Define the PAYE_ADR_LN1_TXT parameter
      prmPayeCareOfTxt = w_aDOCommand.CreateParameter(Name:="@paye_care_of_txt", Type:=adVarChar, Direction:=adParamInput, Size:=40, .value:=fnNullIfZLS(varIn:=PayeCareOfTxt, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmPayeCareOfTxt);

      // ---Parameter #7---
      // Define the PAYE_CITY_NM_TXT parameter
      prmPayeCityNmTxt = w_aDOCommand.CreateParameter(Name:="@paye_city_nm_txt", Type:=adVarChar, Direction:=adParamInput, Size:=25, .value:=fnNullIfZLS(varIn:=PayeCityNmTxt, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmPayeCityNmTxt);

      // ---Parameter #8---
      // Define the PAYE_CLM_INT_AMT parameter
      prmPayeClmIntAmt = w_aDOCommand.CreateParameter(Name:="@paye_clm_int_amt", Type:=adNumeric, Direction:=adParamInput, .value:=PayeClmIntAmt);
      prmPayeClmIntAmt.Precision = m_dblPayeClmIntAmt.precision;
      prmPayeClmIntAmt.NumericScale = m_dblPayeClmIntAmt.numericScale;
      w_aDOCommand.Parameters.Append(prmPayeClmIntAmt);

      // ---Parameter #9---
      // Define the PAYE_CLM_INT_RT parameter
      prmPayeClmIntRt = w_aDOCommand.CreateParameter(Name:="@paye_clm_int_rt", Type:=adNumeric, Direction:=adParamInput, .value:=PayeClmIntRt);
      prmPayeClmIntRt.Precision = m_dblPayeClmIntRt.precision;
      prmPayeClmIntRt.NumericScale = m_dblPayeClmIntRt.numericScale;
      w_aDOCommand.Parameters.Append(prmPayeClmIntRt);

      // ---Parameter #10---
      // Define the PAYE_CLM_PD_AMT parameter
      prmPayeClmPdAmt = w_aDOCommand.CreateParameter(Name:="@paye_clm_pd_amt", Type:=adNumeric, Direction:=adParamInput, .value:=PayeClmPdAmt);
      prmPayeClmPdAmt.Precision = m_dblPayeClmPdAmt.precision;
      prmPayeClmPdAmt.NumericScale = m_dblPayeClmPdAmt.numericScale;
      w_aDOCommand.Parameters.Append(prmPayeClmPdAmt);

      // ---Parameter #11---
      // Define the PAYE_DFLT_OVRD_IND parameter
      prmPayeDfltOvrdInd = w_aDOCommand.CreateParameter(Name:="@paye_dflt_ovrd_ind", Type:=adChar, Direction:=adParamInput, Size:=1, .value:=fnBoolToYN(bIn:=PayeDfltOvrdInd));
      w_aDOCommand.Parameters.Append(prmPayeDfltOvrdInd);

      // ---Parameter #12---
      // Define the PAYE_DTHB_PMT_AMT parameter
      prmPayeDthbPmtAmt = w_aDOCommand.CreateParameter(Name:="@paye_dthb_pmt_amt", Type:=adNumeric, Direction:=adParamInput, .value:=PayeDthbPmtAmt);
      prmPayeDthbPmtAmt.Precision = m_dblPayeDthbPmtAmt.precision;
      prmPayeDthbPmtAmt.NumericScale = m_dblPayeDthbPmtAmt.numericScale;
      w_aDOCommand.Parameters.Append(prmPayeDthbPmtAmt);

      // ---Parameter #13---
      // Define the PAYE_FULL_NM parameter
      prmPayeFullNm = w_aDOCommand.CreateParameter(Name:="@paye_full_nm", Type:=adVarChar, Direction:=adParamInput, Size:=40, .value:=fnNullIfZLS(varIn:=PayeFullNm, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmPayeFullNm);

      // ---Parameter #14---
      // Define the PAYE_INT_DAYS_PD_NUM parameter
      prmPayeIntDaysPdNum = w_aDOCommand.CreateParameter(Name:="@paye_int_days_pd_num", Type:=adInteger, Direction:=adParamInput, .value:=PayeIntDaysPdNum);
      w_aDOCommand.Parameters.Append(prmPayeIntDaysPdNum);

      // ---Parameter #15---
      // Define the PAYE_PMT_DT parameter
      prmPayePmtDt = w_aDOCommand.CreateParameter(Name:="@paye_pmt_dt", Type:=adDBTimeStamp, Direction:=adParamInput, Size:=16, .value:=PayePmtDt);
      w_aDOCommand.Parameters.Append(prmPayePmtDt);

      // ---Parameter #16---
      // Define the PAYE_SSN_TIN_NUM parameter
      prmPayeSsnTinNum = w_aDOCommand.CreateParameter(Name:="@paye_ssn_tin_num", Type:=adChar, Direction:=adParamInput, Size:=9, .value:=fnNullIfZLS(varIn:=PayeSsnTinNum, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmPayeSsnTinNum);

      // ---Parameter #17---
      // Define the PAYE_SSN_TIN_TYP_CD parameter
      prmPayeSsnTinTypCd = w_aDOCommand.CreateParameter(Name:="@paye_ssn_tin_typ_cd", Type:=adChar, Direction:=adParamInput, Size:=1, .value:=fnNullIfZLS(varIn:=PayeSsnTinTypCd, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmPayeSsnTinTypCd);

      // ---Parameter #18---
      // Define the PAYE_ST_CD parameter
      prmPayeStCd = w_aDOCommand.CreateParameter(Name:="@paye_st_cd", Type:=adChar, Direction:=adParamInput, Size:=2, .value:=fnNullIfZLS(varIn:=PayeStCd, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmPayeStCd);

      // ---Parameter #19---
      // Define the PAY_WTHLD_AMT parameter
      prmPayeWthldAmt = w_aDOCommand.CreateParameter(Name:="@paye_wthld_amt", Type:=adNumeric, Direction:=adParamInput, .value:=PayeWthldAmt);
      prmPayeWthldAmt.Precision = m_dblPayeWthldAmt.precision;
      prmPayeWthldAmt.NumericScale = m_dblPayeWthldAmt.numericScale;
      w_aDOCommand.Parameters.Append(prmPayeWthldAmt);

      // ---Parameter #20---
      // Define the PAY_WTHLD_RT parameter
      prmPayeWthldRt = w_aDOCommand.CreateParameter(Name:="@paye_wthld_rt", Type:=adNumeric, Direction:=adParamInput, .value:=PayeWthldRt);
      prmPayeWthldRt.Precision = m_dblPayeWthldRt.precision;
      prmPayeWthldRt.NumericScale = m_dblPayeWthldRt.numericScale;
      w_aDOCommand.Parameters.Append(prmPayeWthldRt);


      // ---Parameter #21---
      // Define the PAYE_ZIP4_cd parameter
      prmPayeZip4Cd = w_aDOCommand.CreateParameter(Name:="@paye_zip4_cd", Type:=adChar, Direction:=adParamInput, Size:=4, .value:=fnNullIfZLS(varIn:=PayeZip4Cd, bHandleEmbeddedQuotes:=True));
      //' BZ4999 October 2013 Non US payee - SXS
      if ("ZZ"
.equals(prmPayeStCd)) {
        prmPayeZip4Cd = " ";
      }
      w_aDOCommand.Parameters.Append(prmPayeZip4Cd);

      // ---Parameter #22---
      // Define the PAYE_ZIP_CD parameter
      prmPayeZipCd = w_aDOCommand.CreateParameter(Name:="@paye_zip_cd", Type:=adChar, Direction:=adParamInput, Size:=5, .value:=fnNullIfZLS(varIn:=PayeZipCd, bHandleEmbeddedQuotes:=True));
      //' BZ4999 October 2013 Non US payee - SXS
      if ("ZZ"
.equals(prmPayeStCd)) {
        prmPayeZipCd = " ";
      }
      w_aDOCommand.Parameters.Append(prmPayeZipCd);

      // ---Parameter #23---  '' BZ4999 October 2013 Non US payee - SXS
      // Define the PAYE_DFLT_OVRD_IND parameter
      prmPaye1099INTInd = w_aDOCommand.CreateParameter(Name:="@Paye_1099INT_Ind", Type:=adChar, Direction:=adParamInput, Size:=1, .value:=fnBoolToYN(bIn:=Paye1099INTInd));
      w_aDOCommand.Parameters.Append(prmPaye1099INTInd);


      // ---Parameter #24---
      // Define the Invalid_Key output parameter, which reflects *which* foreign
      // key violation was encountered.
      prmInvalid_Key = w_aDOCommand.CreateParameter(Name:="@Invalid_Key", Type:=adVarChar, Size:=255, Direction:=adParamOutput);
      w_aDOCommand.Parameters.Append(prmInvalid_Key);

      // ---Parameter #25---
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
      bSuccessful = ctclmClaim.getRelativeRecord(getPayeFullNmFromPayeID(prmNew_Id.value), enumPositionDirection.ePDSAMERECORD);

      _rtn = true;
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    modGeneral.fnFreeRecordset(rstTemp);
    modGeneral.fnFreeObject(prmReturnValue);
    modGeneral.fnFreeObject(prmInvalid_Key);
    modGeneral.fnFreeObject(prmCalcStCd);
    modGeneral.fnFreeObject(prmClmId);
    modGeneral.fnFreeObject(prmPayeAddrLn1Txt);
    modGeneral.fnFreeObject(prmPayeAddrLn2Txt);
    modGeneral.fnFreeObject(prmPayeCareOfTxt);
    modGeneral.fnFreeObject(prmPayeCityNmTxt);
    modGeneral.fnFreeObject(prmPayeClmIntAmt);
    modGeneral.fnFreeObject(prmPayeClmIntRt);
    modGeneral.fnFreeObject(prmPayeClmPdAmt);
    modGeneral.fnFreeObject(prmPayeDfltOvrdInd);
    //'' BZ4999 October 2013 Non US payee - SXS
    modGeneral.fnFreeObject(prmPaye1099INTInd);
    modGeneral.fnFreeObject(prmPayeDthbPmtAmt);
    modGeneral.fnFreeObject(prmPayeFullNm);
    modGeneral.fnFreeObject(prmPayeIntDaysPdNum);
    modGeneral.fnFreeObject(prmPayePmtDt);
    modGeneral.fnFreeObject(prmPayeSsnTinNum);
    modGeneral.fnFreeObject(prmPayeSsnTinTypCd);
    modGeneral.fnFreeObject(prmPayeStCd);
    modGeneral.fnFreeObject(prmPayeWthldAmt);
    modGeneral.fnFreeObject(prmPayeWthldRt);
    modGeneral.fnFreeObject(prmPayeZip4Cd);
    modGeneral.fnFreeObject(prmPayeZipCd);
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
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_ADD_WITH_NONUNIQUE_KEY, MCSTRNAME+ cstrCurrentProc, "Payee Name "+ RTrim$(getPayeFullNm())+ "/Claim Number "+ ctclmClaim.getClmNumFromClmID(ctclmClaim.getClmId()));
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
        break;

      case  modResConstants.gCRES_NERR_KEY_NOT_FOUND:
        // 4032 = The @@1 specified (@@2) is no longer defined in the @@3 table. Please choose a different value.
        if ("CALC_ST_CD"
.equals(prmInvalid_Key.toUpperCase())) {
          modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_KEY_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, "Calc State", RTrim$(getCalcStCd()), "STATE_T");
          /**TODO:** resume found: Resume(PROC_EXIT)*/;
        } 
        else if ("PAYE_ST_CD"
.equals(prmInvalid_Key.toUpperCase())) {
          // PAYE_ST_CD
          modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_KEY_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, "Payee State ", RTrim$(getPayeStCd()), "STATE_T");
          /**TODO:** resume found: Resume(PROC_EXIT)*/;
        } 
        else {
          // CLM_ID
          modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_KEY_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, "Claim ID ", CStr(ctclmClaim.getClmId()), "CLAIM_T");
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
      rstSingleRecord_Fresh = fnSelectRecord(getPayeId());

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
    "dbo.proc_payee_delete"
.equals(Const cstrSproc As String);
    boolean bSuccessful = false;
    cadwADOWrapper adwTemp = null;
    DBRecordSet rstTemp = null;
    ADODB.Parameter prmReturnValue = null;
    ADODB.Parameter prmPayeId = null;
    ADODB.Parameter prmDependent_Table = null;
    String strPayeFullNm = "";

    try {


      adwTemp = new cadwADOWrapper();
      adwTemp.commandSetSproc(strSprocName:=cstrSproc);

      // Save the Payee Full name associated with the Payee that's going to be deleted.
      // We'll need it afterwards to position to the previous record.
      strPayeFullNm = getPayeFullNmFromPayeID(getPayeId());

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
      // Define the paye_id parameter
      prmPayeId = w_aDOCommand.CreateParameter(Name:="@paye_id", Type:=adInteger, Direction:=adParamInput, .value:=PayeId);
      w_aDOCommand.Parameters.Append(prmPayeId);

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
      bSuccessful = ctclmClaim.getRelativeRecord(strPayeFullNm, enumPositionDirection.ePDPREVIOUSRECORD);

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
    modGeneral.fnFreeObject(prmPayeId);
    modGeneral.fnFreeObject(prmDependent_Table);

    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (prmReturnValue) {
      case  modResConstants.gCRES_NERR_REC_NOT_FOUND:
        // 4027 = The specified record was not found in the database (@@1).
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_REC_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, "Payee ID "+ RTrim$(getPayeId()));
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
        break;

      case  modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO:
        // 4028 = An error occurred while attempting to @@1 this record.
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO, MCSTRNAME+ cstrCurrentProc, "delete");
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
      prmClmNum = w_aDOCommand.CreateParameter(Name:="clm_num", Type:=adChar, Direction:=adParamInputOutput, Size:=15, .value:=Null);
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
  public boolean getLookupData(int lngClmIdIn) {
    boolean _rtn = false;
    //--------------------------------------------------------------------------
    // Procedure:   fnGetLookupData_ClaimNum
    // Description: Get all rows with the specified Claim ID but only particular columns
    // Returns:     True if successful; False otherwise
    // Params:      None
    // Date:        01/07/2002
    //-----------------------------------------------------------------------------

    //!CUSTOMIZE!  Customize the name of the stored procedure, the number and names
    //             of the parameters and, perhaps, the return values trapped in the
    //             error handler.

    "fnGetLookupData_ClaimNum"
.equals(Const cstrCurrentProc As String);
    "dbo.proc_payee_lu_select"
.equals(Const cstrSproc As String);
    ADODB.Parameter prmReturnValue = null;
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
      // Define the CLM_ID parameter
      prmClmId = w_aDOCommand.CreateParameter(Name:="@clm_id", Type:=adInteger, Direction:=adParamInput, .value:=lngClmIdIn);
      w_aDOCommand.Parameters.Append(prmClmId);

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
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (prmReturnValue) {
      case  modResConstants.gCRES_NERR_REC_NOT_FOUND:
        // 4027 = The specified record was not found in the database (@@1).
        // Wipe out any trace of this error, but return False so the caller
        // knows to go into Add mode if desired. NOTE: The caller can also
        // identify this by looking at the LookupRecordCount public property.
        VBA.ex.Clear;
        _rtn = false;
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
  public Object getPayeFullNmFromPayeID(int lngPayeID) {
    Object _rtn = null;
    //--------------------------------------------------------------------------
    // Procedure:   GetPayeFullNmFromPayeID
    // Description: Query database for the Payee Full Name for a specified Payee ID
    // Params:
    //               lngPayeID  (in)  The Payee ID to translate.
    //-----------------------------------------------------------------------------
    "GetPayeFullNmFromPayeID"
.equals(Const cstrCurrentProc As String);
    "dbo.proc_paye_full_nm_lu_select"
.equals(Const cstrSproc As String);
    DBRecordSet rstTemp = null;
    ADODB.Parameter prmReturnValue = null;
    ADODB.Parameter prmPayeFullNm = null;
    ADODB.Parameter prmPayeId = null;

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
      prmPayeId = w_aDOCommand.CreateParameter(Name:="paye_id", Type:=adInteger, Direction:=adParamInput, .value:=lngPayeID);
      w_aDOCommand.Parameters.Append(prmPayeId);

      // ---Parameter #3---
      prmPayeFullNm = w_aDOCommand.CreateParameter(Name:="paye_full_nm", Type:=adVarChar, Direction:=adParamInputOutput, Size:=40, .value:=Null);
      w_aDOCommand.Parameters.Append(prmPayeFullNm);


      rstTemp = w_aDOCommand.Execute();

      _rtn = prmPayeFullNm.value;
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
      modGeneral.fnFreeObject(prmReturnValue);
      modGeneral.fnFreeObject(prmPayeId);
      modGeneral.fnFreeObject(prmPayeFullNm);

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
  public DBRecordSet getPayeesForClaim(int lngClmID) {
    //--------------------------------------------------------------------------
    // Procedure:   GetPayeesForClaim
    // Description: Returns a recordset containing selected data on each Payee
    //              associated with the specified CLM_ID
    //
    // Parameters:
    //     lngClmId (in) - the CLM_ID of the desired Claim
    //
    // Returns:     A disconnected ADODB.Recordset containing selected table
    //              columns for the specified key
    //-----------------------------------------------------------------------------
    "GetPayeesForClaim"
.equals(Const cstrCurrentProc As String);
    //' Stored procedure to execute
    "dbo.proc_payee_select4"
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
      // Define the paye_id parameter
      prmClmId = w_aDOCommand.CreateParameter(Name:="@clm_id", Type:=adInteger, Direction:=adParamInput, .value:=fnNullIfZero(lngClmID));
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
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_REC_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, "Claim ID "+ RTrim$(lngClmID)+ "/Claim Number "+ ctclmClaim.getClmNumFromClmID(lngClmID));
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

//'////////////////////////////////////////////////////////////////////////////////////////////////
//Public Function GetClaimForPayeeClaim(ByVal lngClmID As Long) As ADODB.Recordset
//    '--------------------------------------------------------------------------
//    ' Procedure:   GetClaimForPayeeClaim
//    ' Description: Returns a recordset containing selected data on the Claim for the Payee
//    '              associated with the specified CLM_ID
//    '
//    ' Parameters:
//    '     lngClmId (in) - the CLM_ID of the desired Claim
//    '
//    ' Returns:     A disconnected ADODB.Recordset containing selected table
//    '              columns for the specified key
//    '
//    ' Created:     Berry Kropiwka 2019-09-27
//    '-----------------------------------------------------------------------------
//    Const cstrCurrentProc          As String = "GetClaimForPayeeClaim"
//    Const cstrSproc                As String = "dbo.proc_claim_select"  ' Stored procedure to execute
//    Dim rstTemp                    As ADODB.Recordset
//    Dim prmReturnValue             As ADODB.Parameter
//    Dim prmClmId                   As ADODB.Parameter
//
//    On Error GoTo PROC_ERR
//
//    If Not (m_adwADO.CommandSetSproc(cstrSproc)) Then
//        GoTo PROC_EXIT
//    End If
//
//    ' For Char/VarChar fields,
//    '     * Use fnNullIfZLS to ensure Nulls are appropriately handled.
//    '     * Do *not* set the optional 2nd parameter to fnNullIfZLS to True.
//    ' For numeric fields,
//    '     * Use fnNullIfZero to ensure Nulls are appropriately handled.
//    ' For Y/N fields,
//    '     * Use fnBoolToYN to ensure True/False is appropriately translated.
//
//    With m_adwADO.ADOCommand
//        ' ---Parameter #1---
//        ' Define the return value that represents the error code (i.e. reason) why
//        ' the stored procedure failed.
//        Set prmReturnValue = .CreateParameter(Name:="@return_value", _
//                                              Type:=adInteger, _
//                                              Direction:=adParamReturnValue, _
//                                              value:=Null)
//        .Parameters.Append prmReturnValue
//
//        ' ---Parameter #2---
//        ' Define the paye_id parameter
//        Set prmClmId = .CreateParameter(Name:="@clm_id", _
//                                         Type:=adInteger, _
//                                         Direction:=adParamInput, _
//                                         value:=fnNullIfZero(lngClmID))
//        .Parameters.Append prmClmId
//
//        Set rstTemp = .Execute()
//    End With
//
//    rstTemp.ActiveConnection = Nothing
//    Set GetClaimForPayeeClaim = rstTemp
//PROC_EXIT:
//    On Error GoTo 0     ' Disable error handler
//
//    ' Clean-up statements go here
//
//    ' Do *not* do "fnFreeRecordset rstTemp" since this will cause the recordset returned
//    ' by this function to be wiped out as well!
//    fnFreeObject prmReturnValue
//    fnFreeObject prmClmId
//
//    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
//        gerhApp.PropagateError mcstrName & cstrCurrentProc
//    End If
//    Exit Function
//PROC_ERR:
//    Select Case prmReturnValue
//        Case gcRES_NERR_REC_NOT_FOUND
//            ' 4027 = The specified record was not found in the database (@@1).
//            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_REC_NOT_FOUND, _
//                                       mcstrName & cstrCurrentProc, _
//                                       "Claim ID " & RTrim$(lngClmID) & "/Claim Number " & GetClmNumFromClmID(lngClmID)
//            Resume PROC_EXIT
//        Case gcRES_NERR_ERR_WHILE_TRYING_TO
//            ' 4028 = An error occurred while attempting to @@1 this record.
//            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_ERR_WHILE_TRYING_TO, _
//                                       mcstrName & cstrCurrentProc, _
//                                       "locate"
//            Resume PROC_EXIT
//    End Select
//
//    ' If any other errors exist, i.e. in Err object, then let it fall through into default error handling.
//
//    Select Case Err.Number
//        Case -2147217900 ' Object not found
//            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_FERR_SPROC_NOT_FOUND, _
//                                       mcstrName & cstrCurrentProc, _
//                                       cstrSproc
//            Resume PROC_EXIT
//        Case Else
//            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
//    End Select
//    Resume PROC_EXIT
//End Function

//////////////////////////////////////////////////////////////////////////////////////////////////
  public boolean getRelativeRecord(String strKey1, enumPositionDirection lngPositionDirection) {
    //--------------------------------------------------------------------------
    // Procedure:   GetRelativeRecord
    // Description: Refreshes the Lookup recordset and repositions it to
    //              the record relative to the specified key value. Then,
    //              it resets each of the class properties that correspond
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
    //                    This table's true key is PAYE_ID, but the Insured screen
    //                    shows claims in PAYE_FULL_NM sequence. Hence, the key this
    //                    proc uses is PAYE_FULL_NM.
    //
    // Params:
    //     strKey1              (in) = PayeFullNm value from which to do the relative
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

      // In case the Payee Name contains an embedded single quote (e.g. O'Dell),
      // change it to 2 single quotes so SQL will interpret it correctly
      strKey1 = strKey1.replace("'", "''");

      //...........................................................................
      // Refresh the lookup data (m_rstLookupData) so other's changes
      // --and our own-- are now reflected in it. This resets the Lookup data,
      // record count, and current record number, and leaves the Lookup recordset
      // positioned to the first record (if there are records) or BOF (if there are
      // no records).
      //...........................................................................
      ctclmClaim.getLookupData(ctclmClaim.getClmId());

      switch (lngPositionDirection) {
        case  enumPositionDirection.ePDPREVIOUSRECORD:
          // Make visible only those rows with keys prior to the specified key
          m_rstLookup.Filter = "paye_full_nm < '"+ strKey1+ "'";
          if (m_rstLookup.RecordCount != 0) {
            // Grab the last record. The one with the highest key less than the
            // specified key is the one we want.
            m_adwADO.moveLast(m_rstLookup);

            //!CUSTOMIZE! The constant referenced below should refer to the key field.
            strKey1ForNewRec = m_rstLookup.Fields(MCSTRPAYEFULLNM).value;
            // Accommodate embedded single quotes (e.g. O'Dell) by changing it to 2 single quotes
            strKey1ForNewRec = strKey1ForNewRec.replace("'", "''");

            // Okay, we got it. Now make all records in the Lookup recordset
            // visible again, and then reposition to the new record.
            m_rstLookup.Filter = adFilterNone;
            m_adwADO.moveFirst(m_rstLookup);
            m_rstLookup.Find("paye_full_nm = '"+ strKey1ForNewRec+ "'");
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
          m_rstLookup.Filter = "paye_full_nm > '"+ strKey1+ "'";
          if (m_rstLookup.RecordCount != 0) {
            // Grab the first record. The one with the lowest key higher than the
            // specified key is the one we want.
            m_adwADO.moveFirst(m_rstLookup);

            //!CUSTOMIZE! The constant referenced below should refer to the key field.
            strKey1ForNewRec = m_rstLookup.Fields(MCSTRPAYEFULLNM).value;
            // Accommodate embedded single quotes (e.g. O'Dell) by changing it to 2 single quotes
            strKey1ForNewRec = strKey1ForNewRec.replace("'", "''");

            // Okay, we got it. Now make all records in the Lookup recordset
            // visible again, and then reposition to the new record.
            m_rstLookup.Filter = adFilterNone;
            m_adwADO.moveFirst(m_rstLookup);
            m_rstLookup.Find("paye_full_nm = '"+ strKey1ForNewRec+ "'");
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
          m_rstLookup.Find("paye_full_nm = '"+ strKey1+ "'");
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
        ctclmClaim.getSingleRecord(lngKey1:=.Fields(MCSTRPAYEID).value, bSynchLookupRST:=True);
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
    //    lngKey1         (in) = represents the primary key for the table (paye_id)
    //    bSynchLookupRST (in) = indicates whether the Lookup recordset should be
    //                           repositioned to the record this function just
    //                           retrieved. This would be set to True by the
    //                           form's vfgLookup_ChangeEdit event handler to ensure
    //                           the "record x of y" will be set appropriately when
    //                           it calls fnLoadControls.
    //-----------------------------------------------------------------------------

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
      w___TYPE_NOT_FOUND.CalcStCd = modDataConversion.fnZLSIfNull(rstTemp!calc_st_cd);
      w___TYPE_NOT_FOUND.ClmId = modDataConversion.fnZeroIfNull(rstTemp!clm_id);
      w___TYPE_NOT_FOUND.LstUpdtDtm = rstTemp!lst_updt_dtm;
      w___TYPE_NOT_FOUND.LstUpdtUserId = modDataConversion.fnZLSIfNull(rstTemp!lst_updt_user_id);
      w___TYPE_NOT_FOUND.PayeAddrLn1Txt = modDataConversion.fnZLSIfNull(rstTemp!paye_addr_ln1_txt);
      w___TYPE_NOT_FOUND.PayeAddrLn2Txt = modDataConversion.fnZLSIfNull(rstTemp!paye_addr_ln2_txt);
      w___TYPE_NOT_FOUND.PayeCareOfTxt = modDataConversion.fnZLSIfNull(rstTemp!paye_care_of_txt);
      w___TYPE_NOT_FOUND.PayeCityNmTxt = modDataConversion.fnZLSIfNull(rstTemp!paye_city_nm_txt);
      w___TYPE_NOT_FOUND.PayeClmIntAmt = modDataConversion.fnZeroIfNull(rstTemp!paye_clm_int_amt);
      w___TYPE_NOT_FOUND.PayeClmIntRt = modDataConversion.fnZeroIfNull(rstTemp!paye_clm_int_rt);
      w___TYPE_NOT_FOUND.PayeClmPdAmt = modDataConversion.fnZeroIfNull(rstTemp!paye_clm_pd_amt);
      w___TYPE_NOT_FOUND.PayeDfltOvrdInd = modGeneral.fnYNToBool(rstTemp!paye_dflt_ovrd_ind);
      w___TYPE_NOT_FOUND.PayeDthbPmtAmt = modDataConversion.fnZeroIfNull(rstTemp!paye_dthb_pmt_amt);
      w___TYPE_NOT_FOUND.PayeFullNm = modDataConversion.fnZLSIfNull(rstTemp!paye_full_nm);
      w___TYPE_NOT_FOUND.PayeId = modDataConversion.fnZeroIfNull(rstTemp!paye_id);
      w___TYPE_NOT_FOUND.PayeIntDaysPdNum = modDataConversion.fnZeroIfNull(rstTemp!paye_int_days_pd_num);
      w___TYPE_NOT_FOUND.PayePmtDt = rstTemp!paye_pmt_dt;
      w___TYPE_NOT_FOUND.PayeSsnTinNum = modDataConversion.fnZLSIfNull(rstTemp!paye_ssn_tin_num);
      w___TYPE_NOT_FOUND.PayeSsnTinTypCd = modDataConversion.fnZLSIfNull(rstTemp!paye_ssn_tin_typ_cd);
      w___TYPE_NOT_FOUND.PayeStCd = modDataConversion.fnZLSIfNull(rstTemp!paye_St_cd);
      w___TYPE_NOT_FOUND.PayeWthldAmt = rstTemp!paye_wthld_amt;
      w___TYPE_NOT_FOUND.PayeWthldRt = rstTemp!paye_wthld_rt;
      w___TYPE_NOT_FOUND.PayeZip4Cd = modDataConversion.fnZLSIfNull(rstTemp!paye_zip4_cd);
      w___TYPE_NOT_FOUND.PayeZipCd = modDataConversion.fnZLSIfNull(rstTemp!paye_zip_cd);
      //'' BZ4999 October 2013 Non US payee - SXS
      w___TYPE_NOT_FOUND.Paye1099INTInd = modGeneral.fnYNToBool(rstTemp!paye_1099int_ind);

      // Save original Last Updated info, to be used during UpdateRecord( ) and DeleteRecord( )
      // to determine if another user updated the record since it was retrieved.
      m_dteLstUpdtDtm_Original = modDataConversion.fnZLSIfNull(rstTemp!lst_updt_dtm);
      m_strLstUpdtUserId_Original = modDataConversion.fnZLSIfNull(rstTemp!lst_updt_user_id);

      if (bSynchLookupRST) {
        m_adwADO.moveFirst(m_rstLookup);
        m_rstLookup.Find("paye_id = "+ ((Integer) getPayeId()).toString());
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
        ctclmClaim.getSingleRecord(m_rstLookup.Fields(MCSTRPAYEID).value);
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
        ctclmClaim.getSingleRecord(m_rstLookup.Fields(MCSTRPAYEID).value);
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
        if (ctclmClaim.getLookupIsAtBOF() || ctclmClaim.getLookupIsAtEOF()) {
          ctclmClaim.getLookupData(ctclmClaim.getClmId());
          ctclmClaim.getRelativeRecord(getPayeFullNm(), enumPositionDirection.ePDNEXTRECORD);
        } 
        else {
          // Get the requested record and reposition the Lookup recordset to that record
          //!CUSTOMIZE! The constant referenced below should refer to the key field.
          ctclmClaim.getSingleRecord(m_rstLookup.Fields(MCSTRPAYEID).value);
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
        if (ctclmClaim.getLookupIsAtBOF() || ctclmClaim.getLookupIsAtEOF()) {
          ctclmClaim.getLookupData(ctclmClaim.getClmId());
          ctclmClaim.getRelativeRecord(getPayeFullNm(), enumPositionDirection.ePDPREVIOUSRECORD);
        } 
        else {
          // Get the requested record and reposition the Lookup recordset to that record
          //!CUSTOMIZE! The constant referenced below should refer to the key field.
          ctclmClaim.getSingleRecord(m_rstLookup.Fields(MCSTRPAYEID).value);
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
    //             b. (for payee_t table only, I think) row(s) existing in the same table
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
    "dbo.proc_payee_verify_dependents"
.equals(Const cstrSproc As String);
    cadwADOWrapper adwTemp = null;
    DBRecordSet rstTemp = null;
    ADODB.Parameter prmReturnValue = null;
    ADODB.Parameter prmPayeId = null;
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
      prmPayeId = w_aDOCommand.CreateParameter(Name:="@paye_id", Type:=adInteger, Direction:=adParamInput, .value:=lngKey1);
      w_aDOCommand.Parameters.Append(prmPayeId);

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
    modGeneral.fnFreeObject(prmPayeId);
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
        _rtn = false;
        //' This is actually ignored by the caller
        strDependentTable = "Unknown";
        // Remove any trace that this error occurred since we're not going to report it as an error.
        VBA.ex.Clear;
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
        break;

      case  modResConstants.gCRES_NERR_DEPENDENT_RECS_EXIST:
        // 4029 = This @@1 is associated with one or more records on the @@2 table and cannot be deleted until those records themselves are deleted.
        // NOTE: Currently this table has no dependencies and hence the sproc doesn't ever return a 4029.
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
  public void initPayee(int lngClmIdIn) {
    // Comments:   Completes the initialization of the Payee object that
    //             was started in the Class_Initialize procedure. This
    //             procedure should be called immediately after
    //             instantiating an object of this class to
    //             load lookup data associated with the
    //             specified Claim ID.
    //                  Set mtWrapper = New Payee
    //                  InitPayee
    //             The Class_Initialize() method cannot do the
    //             initialization itself since it requires a parameter.
    // Parameters: lngClmId = the Claim ID with which the desired payees are associated
    // Returns:    N/A
    // Called by :
    "InitPayee"
.equals(Const cstrCurrentProc As String);

    try {

      // Store the Claim ID!!!
      ctclmClaim.setClmId(lngClmIdIn);

      // Refresh lookup RST and set LookupRecordCount / CurrentLookupRecNbr properties
      ctclmClaim.getLookupData(lngClmIdIn);

      // Get all columns for the 1st record in the Lookup RST and load to member vars.
      // If there are no records (m_rstLookup is Nothing), then initialize the
      // properties that correspond to table columns. (Caller must take action if
      // m_rstLookup Is Nothing!!!)
      if (m_rstLookup.RecordCount != 0) {
        //!CUSTOMIZE! The constant referenced below should refer to the key field.
        ctclmClaim.getSingleRecord(m_rstLookup.Fields(MCSTRPAYEID).value);
      } 
      else {
        fnClearPropertyValues();
      }

      // Obtain meta data about each table column from the DBMS and load it to the
      // properties that correspond to those table columns
      fnLoadColMetaData();
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
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

    //!CUSTOMIZE!  Customize the name of the stored procedure, the number and names
    //             of the parameters and, perhaps, the return values trapped in the
    //             error handler.

    "UpdateRecord"
.equals(Const cstrCurrentProc As String);
    //' Stored procedure to execute
    "dbo.proc_payee_update"
.equals(Const cstrSproc As String);
    DBRecordSet rstTemp = null;
    boolean bSuccessful = false;
    ADODB.Parameter prmReturnValue = null;
    ADODB.Parameter prmInvalid_Key = null;
    ADODB.Parameter prmPayeId = null;
    ADODB.Parameter prmCalcStCd = null;
    ADODB.Parameter prmClmId = null;
    ADODB.Parameter prmPayeAddrLn1Txt = null;
    ADODB.Parameter prmPayeAddrLn2Txt = null;
    ADODB.Parameter prmPayeCareOfTxt = null;
    ADODB.Parameter prmPayeCityNmTxt = null;
    ADODB.Parameter prmPayeClmIntAmt = null;
    ADODB.Parameter prmPayeClmIntRt = null;
    ADODB.Parameter prmPayeClmPdAmt = null;
    ADODB.Parameter prmPayeDfltOvrdInd = null;
    ADODB.Parameter prmPayeDthbPmtAmt = null;
    ADODB.Parameter prmPayeFullNm = null;
    ADODB.Parameter prmPayeIntDaysPdNum = null;
    ADODB.Parameter prmPayePmtDt = null;
    ADODB.Parameter prmPayeSsnTinNum = null;
    ADODB.Parameter prmPayeSsnTinTypCd = null;
    ADODB.Parameter prmPayeStCd = null;
    ADODB.Parameter prmPayeWthldAmt = null;
    ADODB.Parameter prmPayeWthldRt = null;
    ADODB.Parameter prmPayeZip4Cd = null;
    ADODB.Parameter prmPayeZipCd = null;
    //'' BZ4999 October 2013 Non US payee - SXS
    ADODB.Parameter prmPaye1099INTInd = null;
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
      // Define the PAYE_ID parameter
      prmPayeId = w_aDOCommand.CreateParameter(Name:="@paye_id", Type:=adInteger, Direction:=adParamInput, .value:=PayeId);
      w_aDOCommand.Parameters.Append(prmPayeId);

      // ---Parameter #3---
      // Define the CALC_ST_CD parameter
      prmCalcStCd = w_aDOCommand.CreateParameter(Name:="@calc_st_cd", Type:=adChar, Direction:=adParamInput, Size:=2, .value:=fnNullIfZLS(varIn:=CalcStCd, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmCalcStCd);

      // ---Parameter #4---
      // Define the CLM_ID parameter
      prmClmId = w_aDOCommand.CreateParameter(Name:="@clm_id", Type:=adInteger, Direction:=adParamInput, .value:=ClmId);
      w_aDOCommand.Parameters.Append(prmClmId);

      // ---Parameter #5---
      // Define the PAYE_ADR_LN1_TXT parameter
      prmPayeAddrLn1Txt = w_aDOCommand.CreateParameter(Name:="@paye_addr_ln1_txt", Type:=adVarChar, Direction:=adParamInput, Size:=40, .value:=fnNullIfZLS(varIn:=PayeAddrLn1Txt, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmPayeAddrLn1Txt);

      // ---Parameter #6---
      // Define the PAYE_ADR_LN2_TXT parameter
      prmPayeAddrLn2Txt = w_aDOCommand.CreateParameter(Name:="@paye_addr_ln2_txt", Type:=adVarChar, Direction:=adParamInput, Size:=40, .value:=fnNullIfZLS(varIn:=PayeAddrLn2Txt, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmPayeAddrLn2Txt);

      // ---Parameter #7---
      // Define the PAYE_CARE_OF_TXT parameter
      prmPayeCareOfTxt = w_aDOCommand.CreateParameter(Name:="@paye_care_of_txt", Type:=adVarChar, Direction:=adParamInput, Size:=40, .value:=fnNullIfZLS(varIn:=PayeCareOfTxt, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmPayeCareOfTxt);

      // ---Parameter #8---
      // Define the PAYE_CITY_NM_TXT parameter
      prmPayeCityNmTxt = w_aDOCommand.CreateParameter(Name:="@paye_city_nm_txt", Type:=adVarChar, Direction:=adParamInput, Size:=25, .value:=fnNullIfZLS(varIn:=PayeCityNmTxt, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmPayeCityNmTxt);

      // ---Parameter #9---
      // Define the PAYE_CLM_INT_AMT parameter
      prmPayeClmIntAmt = w_aDOCommand.CreateParameter(Name:="@paye_clm_int_amt", Type:=adNumeric, Direction:=adParamInput, .value:=PayeClmIntAmt);
      prmPayeClmIntAmt.Precision = m_dblPayeClmIntAmt.precision;
      prmPayeClmIntAmt.NumericScale = m_dblPayeClmIntAmt.numericScale;
      w_aDOCommand.Parameters.Append(prmPayeClmIntAmt);

      // ---Parameter #10---
      // Define the PAYE_CLM_INT_RT parameter
      prmPayeClmIntRt = w_aDOCommand.CreateParameter(Name:="@paye_clm_int_rt", Type:=adNumeric, Direction:=adParamInput, .value:=PayeClmIntRt);
      prmPayeClmIntRt.Precision = m_dblPayeClmIntRt.precision;
      prmPayeClmIntRt.NumericScale = m_dblPayeClmIntRt.numericScale;
      w_aDOCommand.Parameters.Append(prmPayeClmIntRt);

      // ---Parameter #11---
      // Define the PAYE_CLM_PD_AMT parameter
      prmPayeClmPdAmt = w_aDOCommand.CreateParameter(Name:="@paye_clm_pd_amt", Type:=adNumeric, Direction:=adParamInput, .value:=PayeClmPdAmt);
      prmPayeClmPdAmt.Precision = m_dblPayeClmPdAmt.precision;
      prmPayeClmPdAmt.NumericScale = m_dblPayeClmPdAmt.numericScale;
      w_aDOCommand.Parameters.Append(prmPayeClmPdAmt);

      // ---Parameter #12---
      // Define the PAYE_DFLT_OVRD_IND parameter
      prmPayeDfltOvrdInd = w_aDOCommand.CreateParameter(Name:="@paye_dflt_ovrd_ind", Type:=adChar, Direction:=adParamInput, Size:=1, .value:=fnBoolToYN(bIn:=PayeDfltOvrdInd));
      w_aDOCommand.Parameters.Append(prmPayeDfltOvrdInd);

      // ---Parameter #13---
      // Define the PAYE_DTHB_PMT_AMT parameter
      prmPayeDthbPmtAmt = w_aDOCommand.CreateParameter(Name:="@paye_dthb_pmt_amt", Type:=adNumeric, Direction:=adParamInput, .value:=PayeDthbPmtAmt);
      prmPayeDthbPmtAmt.Precision = m_dblPayeDthbPmtAmt.precision;
      prmPayeDthbPmtAmt.NumericScale = m_dblPayeDthbPmtAmt.numericScale;
      w_aDOCommand.Parameters.Append(prmPayeDthbPmtAmt);

      // ---Parameter #14---
      // Define the PAYE_FULL_NM parameter
      prmPayeFullNm = w_aDOCommand.CreateParameter(Name:="@paye_full_nm", Type:=adVarChar, Direction:=adParamInput, Size:=40, .value:=fnNullIfZLS(varIn:=PayeFullNm, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmPayeFullNm);

      // ---Parameter #15---
      // Define the PAYE_INT_DAYS_PD_NUM parameter
      prmPayeIntDaysPdNum = w_aDOCommand.CreateParameter(Name:="@paye_int_days_pd_num", Type:=adInteger, Direction:=adParamInput, .value:=PayeIntDaysPdNum);
      w_aDOCommand.Parameters.Append(prmPayeIntDaysPdNum);

      // ---Parameter #16---
      // Define the PAYE_PMT_DT parameter
      prmPayePmtDt = w_aDOCommand.CreateParameter(Name:="@paye_pmt_dt", Type:=adDBTimeStamp, Direction:=adParamInput, Size:=16, .value:=PayePmtDt);
      w_aDOCommand.Parameters.Append(prmPayePmtDt);

      // ---Parameter #17---
      // Define the PAYE_SSN_TIN_NUM parameter
      prmPayeSsnTinNum = w_aDOCommand.CreateParameter(Name:="@paye_ssn_tin_num", Type:=adChar, Direction:=adParamInput, Size:=9, .value:=fnNullIfZLS(varIn:=PayeSsnTinNum, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmPayeSsnTinNum);

      // ---Parameter #18---
      // Define the PAYE_SSN_TIN_TYP_CD parameter
      prmPayeSsnTinTypCd = w_aDOCommand.CreateParameter(Name:="@paye_ssn_tin_typ_cd", Type:=adChar, Direction:=adParamInput, Size:=1, .value:=fnNullIfZLS(varIn:=PayeSsnTinTypCd, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmPayeSsnTinTypCd);

      // ---Parameter #19---
      // Define the PAYE_ST_CD parameter
      prmPayeStCd = w_aDOCommand.CreateParameter(Name:="@paye_st_cd", Type:=adChar, Direction:=adParamInput, Size:=2, .value:=fnNullIfZLS(varIn:=PayeStCd, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmPayeStCd);

      // ---Parameter #20---
      // Define the PAY_WTHLD_AMT parameter
      prmPayeWthldAmt = w_aDOCommand.CreateParameter(Name:="@paye_wthld_amt", Type:=adNumeric, Direction:=adParamInput, .value:=PayeWthldAmt);
      prmPayeWthldAmt.Precision = m_dblPayeWthldAmt.precision;
      prmPayeWthldAmt.NumericScale = m_dblPayeWthldAmt.numericScale;
      w_aDOCommand.Parameters.Append(prmPayeWthldAmt);

      // ---Parameter #21---
      // Define the PAY_WTHLD_RT parameter
      prmPayeWthldRt = w_aDOCommand.CreateParameter(Name:="@paye_wthld_rt", Type:=adNumeric, Direction:=adParamInput, .value:=PayeWthldRt);
      prmPayeWthldRt.Precision = m_dblPayeWthldRt.precision;
      prmPayeWthldRt.NumericScale = m_dblPayeWthldRt.numericScale;
      w_aDOCommand.Parameters.Append(prmPayeWthldRt);


      // ---Parameter #22---
      // Define the PAYE_ZIP4_cd parameter
      prmPayeZip4Cd = w_aDOCommand.CreateParameter(Name:="@paye_zip4_cd", Type:=adChar, Direction:=adParamInput, Size:=4, .value:=fnNullIfZLS(varIn:=PayeZip4Cd, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmPayeZip4Cd);

      // ---Parameter #23---
      // Define the PAYE_ZIP_CD parameter
      prmPayeZipCd = w_aDOCommand.CreateParameter(Name:="@paye_zip_cd", Type:=adChar, Direction:=adParamInput, Size:=5, .value:=fnNullIfZLS(varIn:=PayeZipCd, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmPayeZipCd);

      // ---Parameter #24---  '' BZ4999 October 2013 Non US payee - SXS
      // Define the Paye1099INTInd parameter
      prmPaye1099INTInd = w_aDOCommand.CreateParameter(Name:="@Paye_1099INT_Ind", Type:=adChar, Direction:=adParamInput, Size:=1, .value:=fnBoolToYN(bIn:=Paye1099INTInd));
      w_aDOCommand.Parameters.Append(prmPaye1099INTInd);

      // ---Parameter #25---
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
      bSuccessful = ctclmClaim.getRelativeRecord(getPayeFullNm(), enumPositionDirection.ePDSAMERECORD);

      _rtn = true;
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    modGeneral.fnFreeRecordset(rstTemp);
    modGeneral.fnFreeObject(prmReturnValue);
    modGeneral.fnFreeObject(prmInvalid_Key);
    modGeneral.fnFreeObject(prmPayeId);
    modGeneral.fnFreeObject(prmCalcStCd);
    modGeneral.fnFreeObject(prmClmId);
    modGeneral.fnFreeObject(prmPayeAddrLn1Txt);
    modGeneral.fnFreeObject(prmPayeAddrLn2Txt);
    modGeneral.fnFreeObject(prmPayeCareOfTxt);
    modGeneral.fnFreeObject(prmPayeCityNmTxt);
    modGeneral.fnFreeObject(prmPayeClmIntAmt);
    modGeneral.fnFreeObject(prmPayeClmIntRt);
    modGeneral.fnFreeObject(prmPayeClmPdAmt);
    modGeneral.fnFreeObject(prmPayeDfltOvrdInd);
    //'' BZ4999 October 2013 Non US payee - SXS
    modGeneral.fnFreeObject(prmPaye1099INTInd);
    modGeneral.fnFreeObject(prmPayeDthbPmtAmt);
    modGeneral.fnFreeObject(prmPayeFullNm);
    modGeneral.fnFreeObject(prmPayeIntDaysPdNum);
    modGeneral.fnFreeObject(prmPayePmtDt);
    modGeneral.fnFreeObject(prmPayeSsnTinNum);
    modGeneral.fnFreeObject(prmPayeSsnTinTypCd);
    modGeneral.fnFreeObject(prmPayeStCd);
    modGeneral.fnFreeObject(prmPayeWthldAmt);
    modGeneral.fnFreeObject(prmPayeWthldRt);
    modGeneral.fnFreeObject(prmPayeZip4Cd);
    modGeneral.fnFreeObject(prmPayeZipCd);


    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    // Can get -2147217887 "Invalid character value for cast specification" if too
    // many characters are defined for a numeric field, e.g., 1111111.22222 to store to
    // a field defined as 99.99999.

    switch (prmReturnValue) {
      case  modResConstants.gCRES_NERR_REC_NOT_FOUND:
        // 4027 = The specified record was not found in the database (@@1).
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_REC_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, "Payee ID "+ RTrim$(getPayeId()));
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
        break;

      case  modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO:
        // 4028 = An error occurred while attempting to @@1 this record.
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO, MCSTRNAME+ cstrCurrentProc, "update");
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
        break;

      case  modResConstants.gCRES_NERR_KEY_NOT_FOUND:
        // 4032 = The @@1 specified (@@2) is no longer defined in the @@3 table. Please choose a different value.
        if ("CLM_ID"
.equals(prmInvalid_Key.toUpperCase())) {
          modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_KEY_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, "Claim ID", RTrim$(ctclmClaim.getClmId()), "CLAIM_T");
          /**TODO:** resume found: Resume(PROC_EXIT)*/;
        } 
        else if ("CALC_ST_CD"
.equals(prmInvalid_Key.toUpperCase())) {
          modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_KEY_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, "Calc State", RTrim$(getCalcStCd()), "STATE_T");
          /**TODO:** resume found: Resume(PROC_EXIT)*/;
        } 
        else {
          // PAYE_ST_CD
          modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_KEY_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, "Payee State", RTrim$(getPayeStCd()), "STATE_T");
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
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_ADD_WITH_NONUNIQUE_KEY, MCSTRNAME+ cstrCurrentProc, "Payee Name "+ RTrim$(getPayeFullNm())+ "/Claim Number "+ ctclmClaim.getClmNumFromClmID(ctclmClaim.getClmId()));
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
      w___TYPE_NOT_FOUND.CalcStCd = "";
      w___TYPE_NOT_FOUND.LstUpdtDtm = Now;
      w___TYPE_NOT_FOUND.LstUpdtUserId = "";
      w___TYPE_NOT_FOUND.PayeAddrLn1Txt = "";
      w___TYPE_NOT_FOUND.PayeAddrLn2Txt = "";
      w___TYPE_NOT_FOUND.PayeCareOfTxt = "";
      w___TYPE_NOT_FOUND.PayeCityNmTxt = "";
      w___TYPE_NOT_FOUND.PayeClmIntAmt = clngZero;
      w___TYPE_NOT_FOUND.PayeClmIntRt = clngZero;
      w___TYPE_NOT_FOUND.PayeClmPdAmt = clngZero;
      w___TYPE_NOT_FOUND.PayeDfltOvrdInd = false;
      //'' BZ4999 October 2013 Non US payee - SXS
      w___TYPE_NOT_FOUND.Paye1099INTInd = true;
      w___TYPE_NOT_FOUND.PayeDthbPmtAmt = clngZero;
      w___TYPE_NOT_FOUND.PayeFullNm = "";
      w___TYPE_NOT_FOUND.PayeId = clngZero;
      w___TYPE_NOT_FOUND.PayeIntDaysPdNum = clngZero;
      w___TYPE_NOT_FOUND.PayePmtDt = Now;
      w___TYPE_NOT_FOUND.PayeSsnTinNum = "";
      w___TYPE_NOT_FOUND.PayeSsnTinTypCd = "";
      w___TYPE_NOT_FOUND.PayeStCd = "";
      w___TYPE_NOT_FOUND.PayeWthldAmt = clngZero;
      w___TYPE_NOT_FOUND.PayeWthldRt = clngZero;
      w___TYPE_NOT_FOUND.PayeZip4Cd = "";
      w___TYPE_NOT_FOUND.PayeZipCd = "";

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
      // The following data types are used within the Claims Interest tables:
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
      if (strDefaultValueToParse.indexOf("DEF_LST_UPD_USER", 1) > 0) {
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

      switch (pudtCol.dataType) {
        case  dbDecimal:
          // Save original values (will be used to code sproc parameters).
          pudtCol.numericScale = CByte(prstIn("SCALE").value);
          pudtCol.precision = CByte(prstIn("PRECISION").value);
          // Save interpreted equivalents. These may be overriden in fnLoadColMetaData( ).
          pudtCol.decimalPositions = Integer.parseInt(prstIn("SCALE").value);
          pudtCol.dollarPositions = Integer.parseInt(prstIn("PRECISION").value) - pudtCol.decimalPositions;
          pudtCol.maxCharacters = 0;
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
          // Establish default values that will usually be in effect if a
          // TextBox control is used
          pudtCol.maxCharacters = Integer.parseInt(prstIn("LENGTH").value);
          pudtCol.allowableCharacters = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789@!#$%^&*()-+_=~:;.,<>\\|/?' ";
          pudtCol.format = String(pudtCol.maxCharacters, cstrAnyCharChar);
          pudtCol.mask = "";
          pudtCol.shouldForceToUppercase = false;

          if ("SSN_NUM"
.equals(pudtCol.colName.substring(pudtCol.colName.length() - 8))) {
            // Social Security Numbers should be displayed via a MaskEdBox control.
            pudtCol.format = "";
            pudtCol.mask = "###-##-####";
            pudtCol.allowableCharacters = "0123456789";
          }
          if ("ZIP_CD"
.equals(pudtCol.colName.substring(pudtCol.colName.length() - 6))) {
            // Zip Codes require 5 numeric positions if input.
            // They should be displayed via a TextBox control.
            pudtCol.mask = "";
            pudtCol.maxCharacters = 5;
            pudtCol.format = String(pudtCol.maxCharacters, cstrNumericChar);
            pudtCol.allowableCharacters = "0123456789";
          }
          if ("ZIP4_CD"
.equals(pudtCol.colName.substring(pudtCol.colName.length() - 7))) {
            // Zip Codes require 4 numeric positions if input.
            // They should be displayed via a TextBox control.
            pudtCol.mask = "";
            pudtCol.maxCharacters = 4;
            pudtCol.format = String(pudtCol.maxCharacters, cstrNumericChar);
            pudtCol.allowableCharacters = "0123456789";
          }

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
        case  "CALCSTCD":
          _rtn = m_strCalcStCd;
          break;

        case  "CLMID":
          _rtn = m_lngClmId;
          break;

        case  "LSTUPDTDTM":
          _rtn = m_dteLstUpdtDtm;
          break;

        case  "LSTUPDTUSERID":
          _rtn = m_strLstUpdtUserId;
          break;

        case  "PAYEADDRLN1TXT":
          _rtn = m_strPayeAddrLn1Txt;
          break;

        case  "PAYEADDRLN2TXT":
          _rtn = m_strPayeAddrLn2Txt;
          break;

        case  "PAYECAREOFTXT":
          _rtn = m_strPayeCareOfTxt;
          break;

        case  "PAYECITYNMTXT":
          _rtn = m_strPayeCityNmTxt;
          break;

        case  "PAYECLMINTAMT":
          _rtn = m_dblPayeClmIntAmt;
          break;

        case  "PAYECLMINTRT":
          _rtn = m_dblPayeClmIntRt;
          break;

        case  "PAYECLMPDAMT":
          _rtn = m_dblPayeClmPdAmt;
          break;

        case  "PAYEDFLTOVRDIND":
          _rtn = m_bPayeDfltOvrdInd;
        //'' BZ4999 October 2013 Non US payee - SXS
          break;

        case  "PAYE1099IND"  :
          _rtn = m_bstrPaye1099INTInd;
          break;

        case  "PAYEDTHBPMTAMT":
          _rtn = m_dblPayeDthbPmtAmt;
          break;

        case  "PAYEFULLNM":
          _rtn = m_strPayeFullNm;
          break;

        case  "PAYEID":
          _rtn = m_lngPayeId;
          break;

        case  "PAYEINTDAYSPDNUM":
          _rtn = m_intPayeIntDaysPdNum;
          break;

        case  "PAYEPMTDT":
          _rtn = m_dtePayePmtDt;
          break;

        case  "PAYESSNTINNUM":
          _rtn = m_strPayeSsnTinNum;
          break;

        case  "PAYESSNTINTYPCD":
          _rtn = m_strPayeSsnTinTypCd;
          break;

        case  "PAYESTCD":
          _rtn = m_strPayeStCd;
          break;

        case  "PAYEWTHLDAMT":
          _rtn = m_dblPayeWthldAmt;
          break;

        case  "PAYEWTHLDRT":
          _rtn = m_dblPayeWthldRt;
          break;

        case  "PAYEZIP4CD":
          _rtn = m_strPayeZip4Cd;
          break;

        case  "PAYEZIPCD":
          _rtn = m_strPayeZipCd;
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
      m_adwADO.getMetaData_Columns("payee_t", rstMetaData);

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
        switch (rstMetaData("COLUMN_NAME").value.toUpperCase()) {
          case  MCSTRCALCSTCD:
            fnGetColMetaData(m_strCalcStCd, rstMetaData);
            break;

          case  MCSTRCLMID:
            fnGetColMetaData(m_lngClmId, rstMetaData);
            break;

          case  MCSTRLSTUPDTDTM:
            fnGetColMetaData(m_dteLstUpdtDtm, rstMetaData);
            break;

          case  MCSTRLSTUPDTUSERID:
            fnGetColMetaData(m_strLstUpdtUserId, rstMetaData);
            break;

          case  MCSTRPAYEADDRLN1TXT:
            fnGetColMetaData(m_strPayeAddrLn1Txt, rstMetaData);
            m_strPayeAddrLn1Txt.shouldForceToUppercase = true;
            break;

          case  MCSTRPAYEADDRLN2TXT:
            fnGetColMetaData(m_strPayeAddrLn2Txt, rstMetaData);
            m_strPayeAddrLn2Txt.shouldForceToUppercase = true;
            break;

          case  MCSTRPAYECAREOFTXT:
            fnGetColMetaData(m_strPayeCareOfTxt, rstMetaData);
            m_strPayeCareOfTxt.shouldForceToUppercase = true;
            break;

          case  MCSTRPAYECITYNMTXT:
            fnGetColMetaData(m_strPayeCityNmTxt, rstMetaData);
            m_strPayeCityNmTxt.shouldForceToUppercase = true;
            break;

          case  MCSTRPAYECLMINTAMT:
            fnGetColMetaData(m_dblPayeClmIntAmt, rstMetaData);
            break;

          case  MCSTRPAYECLMINTRT:
            fnGetColMetaData(m_dblPayeClmIntRt, rstMetaData);
            break;

          case  MCSTRPAYECLMPDAMT:
            fnGetColMetaData(m_dblPayeClmPdAmt, rstMetaData);
            break;

          case  MCSTRPAYEDFLTOVRDIND:
            fnGetColMetaData(m_bPayeDfltOvrdInd, rstMetaData);
            break;

          case  MCSTRPAYEDTHBPMTAMT:
            fnGetColMetaData(m_dblPayeDthbPmtAmt, rstMetaData);
            // GUI enforces this default value of zero, not DBMS
            m_dblPayeDthbPmtAmt.defaultValue = 0;
            break;

          case  MCSTRPAYEFULLNM:
            fnGetColMetaData(m_strPayeFullNm, rstMetaData);
            m_strPayeFullNm.shouldForceToUppercase = true;
            break;

          case  MCSTRPAYEID:
            fnGetColMetaData(m_lngPayeId, rstMetaData);
            break;

          case  MCSTRPAYEINTDAYSPDNUM:
            fnGetColMetaData(m_intPayeIntDaysPdNum, rstMetaData);
            break;

          case  MCSTRPAYEPMTDT:
            fnGetColMetaData(m_dtePayePmtDt, rstMetaData);
            break;

          case  MCSTRPAYESSNTINNUM:
            fnGetColMetaData(m_strPayeSsnTinNum, rstMetaData);
            m_strPayeSsnTinNum.allowableCharacters = "0123456789";
            break;

          case  MCSTRPAYESSNTINTYPCD:
            fnGetColMetaData(m_strPayeSsnTinTypCd, rstMetaData);
            m_strPayeSsnTinTypCd.defaultValue = modGeneral.gCSTRBLANKENTRY;
            break;

          case  MCSTRPAYESTCD:
            fnGetColMetaData(m_strPayeStCd, rstMetaData);
            break;

          case  MCSTRPAYEWTHLDAMT:
            fnGetColMetaData(m_dblPayeWthldAmt, rstMetaData);
            break;

          case  MCSTRPAYEWTHLDRT:
            fnGetColMetaData(m_dblPayeWthldRt, rstMetaData);
            break;

          case  MCSTRPAYEZIP4CD:
            fnGetColMetaData(m_strPayeZip4Cd, rstMetaData);
            break;

          case  MCSTRPAYEZIPCD:
            fnGetColMetaData(m_strPayeZipCd, rstMetaData);
            break;

          case  MCSTRPAYE1099INTIND:
            fnGetColMetaData(m_bstrPaye1099INTInd, rstMetaData);
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
      m_adwADO.getMetaData_PrimaryKeys("payee_t", rstMetaData);

      do Until .EOF        // The SELECT CASE should list all table columns
        // (though you could skip the LST_UPDT_xxx columns if you change the Case Else,
        //  since these would never be a key)
        switch (rstMetaData("COLUMN_NAME").value.toUpperCase()) {
          case  MCSTRCALCSTCD:
            m_strCalcStCd.isKey = true;
            break;

          case  MCSTRCLMID:
            m_lngClmId.isKey = true;
            break;

          case  MCSTRLSTUPDTDTM:
            m_dteLstUpdtDtm.isKey = true;
            break;

          case  MCSTRLSTUPDTUSERID:
            m_strLstUpdtUserId.isKey = true;
            break;

          case  MCSTRPAYEADDRLN1TXT:
            m_strPayeAddrLn1Txt.isKey = true;
            break;

          case  MCSTRPAYEADDRLN2TXT:
            m_strPayeAddrLn2Txt.isKey = true;
            break;

          case  MCSTRPAYECAREOFTXT:
            m_strPayeCareOfTxt.isKey = true;
            break;

          case  MCSTRPAYECITYNMTXT:
            m_strPayeCityNmTxt.isKey = true;
            break;

          case  MCSTRPAYECLMINTAMT:
            m_dblPayeClmIntAmt.isKey = true;
            break;

          case  MCSTRPAYECLMINTRT:
            m_dblPayeClmIntRt.isKey = true;
            break;

          case  MCSTRPAYECLMPDAMT:
            m_dblPayeClmPdAmt.isKey = true;
            break;

          case  MCSTRPAYEDFLTOVRDIND:
            m_bPayeDfltOvrdInd.isKey = true;
            break;

          case  MCSTRPAYEDTHBPMTAMT:
            m_dblPayeDthbPmtAmt.isKey = true;
            break;

          case  MCSTRPAYEFULLNM:
            m_strPayeFullNm.isKey = true;
            break;

          case  MCSTRPAYEID:
            m_lngPayeId.isKey = true;
            break;

          case  MCSTRPAYEINTDAYSPDNUM:
            m_intPayeIntDaysPdNum.isKey = true;
            break;

          case  MCSTRPAYEPMTDT:
            m_dtePayePmtDt.isKey = true;
            break;

          case  MCSTRPAYESSNTINNUM:
            m_strPayeSsnTinNum.isKey = true;
            break;

          case  MCSTRPAYESSNTINTYPCD:
            m_strPayeSsnTinTypCd.isKey = true;
            break;

          case  MCSTRPAYESTCD:
            m_strPayeStCd.isKey = true;
            break;

          case  MCSTRPAYEWTHLDAMT:
            m_dblPayeWthldAmt.isKey = true;
            break;

          case  MCSTRPAYEWTHLDRT:
            m_dblPayeWthldRt.isKey = true;
            break;

          case  MCSTRPAYEZIP4CD:
            m_strPayeZip4Cd.isKey = true;
            break;

          case  MCSTRPAYEZIPCD:
            m_strPayeZipCd.isKey = true;
            break;

          case  MCSTRPAYE1099INTIND:
            m_bstrPaye1099INTInd.isKey = true;
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
    "dbo.proc_payee_select"
.equals(Const cstrSproc As String);
    DBRecordSet rstTemp = null;
    ADODB.Parameter prmReturnValue = null;
    ADODB.Parameter prmPayeId = null;

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
      // Define the paye_id parameter
      prmPayeId = w_aDOCommand.CreateParameter(Name:="@paye_id", Type:=adInteger, Direction:=adParamInput, .value:=fnNullIfZero(lngKey1));
      w_aDOCommand.Parameters.Append(prmPayeId);

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
    modGeneral.fnFreeObject(prmPayeId);

    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return null;
    // **TODO:** label found: PROC_ERR:;
    switch (prmReturnValue) {
      case  modResConstants.gCRES_NERR_REC_NOT_FOUND:
        // 4027 = The specified record was not found in the database (@@1).
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_REC_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, "Payee ID "+ RTrim$(lngKey1));
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


case class TpyepayeeData(
              id: Option[Int],

              )

object Tpyepayees extends Controller with ProvidesUser {

  val tpyepayeeForm = Form(
    mapping(
      "id" -> optional(number),

  )(TpyepayeeData.apply)(TpyepayeeData.unapply))

  implicit val tpyepayeeWrites = new Writes[Tpyepayee] {
    def writes(tpyepayee: Tpyepayee) = Json.obj(
      "id" -> Json.toJson(tpyepayee.id),
      C.ID -> Json.toJson(tpyepayee.id),

    )
  }

  def get(id: Int) = GetAction { implicit request =>
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.LIST_TPYEPAYEE), { user =>
      Ok(Json.toJson(Tpyepayee.get(user, id)))
    })
  }

  def update(id: Int) = PostAction { implicit request =>
    Logger.debug("in tpyepayees.update")
    tpyepayeeForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      tpyepayee => {
        Logger.debug(s"form: ${tpyepayee.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.EDIT_TPYEPAYEE), { user =>
          Ok(
            Json.toJson(
              Tpyepayee.update(user,
                Tpyepayee(
                       id,

                ))))
        })
      }
    )
  }

  def create = PostAction { implicit request =>
    Logger.debug("in tpyepayees.create")
    tpyepayeeForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      tpyepayee => {
        Logger.debug(s"form: ${tpyepayee.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.NEW_TPYEPAYEE), { user =>
          Ok(
            Json.toJson(
              Tpyepayee.create(user,
                Tpyepayee(

                ))))
        })
      }
    )
  }

  def delete(id: Int) = PostAction { implicit request =>
    Logger.debug("in tpyepayees.delete")
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.DELETE_TPYEPAYEE), { user =>
      Tpyepayee.delete(user, id)
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

case class Tpyepayee(
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

object Tpyepayee {

  lazy val emptyTpyepayee = Tpyepayee(
)

  def apply(
      id: Int,
) = {

    new Tpyepayee(
      id,
)
  }

  def apply(
) = {

    new Tpyepayee(
)
  }

  private val tpyepayeeParser: RowParser[Tpyepayee] = {
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
        Tpyepayee(
              id,
,
              createdAt,
              updatedAt,
              updatedBy)
    }
  }

  def create(user: CompanyUser, tpyepayee: Tpyepayee): Tpyepayee = {
    save(user, tpyepayee, true)
  }

  def update(user: CompanyUser, tpyepayee: Tpyepayee): Tpyepayee = {
    save(user, tpyepayee, false)
  }

  private def save(user: CompanyUser, tpyepayee: Tpyepayee, isNew: Boolean): Tpyepayee = {
    def getFields = {
      List(

      )
    }
    def throwException = {
      throw new RuntimeException(s"Error when saving ${C.TPYEPAYEE}")
    }

    DBHelper.saveEx(
      user,
      Register(
        C.TPYEPAYEE,
        C.ID,
        tpyepayee.id,
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

  def load(user: CompanyUser, id: Int): Option[Tpyepayee] = {
    loadWhere(user, s"${C.ID} = {id}", 'id -> id)
  }

  def loadWhere(user: CompanyUser, where: String, args : scala.Tuple2[scala.Any, anorm.ParameterValue[_]]*) = {
    DB.withConnection(user.database.database) { implicit connection =>
      SQL(s"SELECT t1.*, t2.${C.FK_NAME} FROM ${C.TPYEPAYEE} t1 INNER JOIN ${C.???} t2 ON t1.${C.FK_ID} = t2.${C.FK_ID} WHERE $where")
        .on(args: _*)
        .as(tpyepayeeParser.singleOpt)
    }
  }

  def delete(user: CompanyUser, id: Int) = {
    DB.withConnection(user.database.database) { implicit connection =>
      try {
        SQL(s"DELETE FROM ${C.TPYEPAYEE} WHERE ${C.ID} = {id}")
        .on('id -> id)
        .executeUpdate
      } catch {
        case NonFatal(e) => {
          Logger.error(s"can't delete a ${C.TPYEPAYEE}. ${C.ID} id: $id. Error ${e.toString}")
          throw e
        }
      }
    }
  }

  def get(user: CompanyUser, id: Int): Tpyepayee = {
    load(user, id) match {
      case Some(p) => p
      case None => emptyTpyepayee
    }
  }
}


// Router

GET     /api/v1/general/tpyepayee/:id              controllers.logged.modules.general.Tpyepayees.get(id: Int)
POST    /api/v1/general/tpyepayee                  controllers.logged.modules.general.Tpyepayees.create
PUT     /api/v1/general/tpyepayee/:id              controllers.logged.modules.general.Tpyepayees.update(id: Int)
DELETE  /api/v1/general/tpyepayee/:id              controllers.logged.modules.general.Tpyepayees.delete(id: Int)




/**/
