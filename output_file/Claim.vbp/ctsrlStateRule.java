
import java.util.Date;

public class ctsrlStateRule {

  //--------------------------------------------------------------------------
  // Procedure:   ctsrlStateRule
  // Description: Provides properties and methods to support the state_rule_t table values
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
  //                    FUND_MGR_PRV_CD  FundMgrPrvCd    m_strFundMgrPrvCd  mcstrFundMgrPrvCd
  //                    FUND_SVS_IND     FundSvsInd      m_bFundSvsInd      mcstrFundSvsInd
  //
  //
  //                 NOTE also that navigation should be done via the **ADO Wrapper's**
  //                 navigation methods instead of directly referencing the navigation
  //                 methods on a ADODB.Recordset object!
  //
  // Revisions:   1.0 BAW 05/05/2002 Initial creation.
  //
  //
  // Procedures  :
  //   Private     Class_Initialize()
  //   Private     Class_Terminate()
  //   Public      Property Get AllowableCharacters(ByVal strTagIn As String) As String
  //   Public      Property Get CurrentLookupRecordNumber() As Long
  //   Public      Property Get DecimalPositions(ByVal strTagIn As String) As Integer
  //   Public      Property Get DefaultValue(ByVal strTagIn As String) As Variant
  //   Public      Property Get DollarPositions(ByVal strTagIn As String) As Integer
  //   Public      Property Get Format(ByVal strTagIn As String) As String
  //   Public      Property Get FundCd() As String
  //   Public      Property Let FundCd(ByVal strValue As String)
  //   Public      Property Get IsKey(ByVal strTagIn As String) As Boolean
  //   Public      Property Get IsNullable(ByVal strTagIn As String) As Boolean
  //   Public      Property Get LookupData() As ADODB.Recordset
  //   Public      Property Get LookupIsAtBOF() As Boolean
  //   Public      Property Get LookupIsAtEOF() As Boolean
  //   Public      Property Get LookupRecordCount() As Long
  //   Public      Property Get LstUpdDtm() As Date
  //   Public      Property Let LstUpdDtm(ByVal NewValue As String)
  //   Public      Property Get LstUpdUserId() As String
  //   Public      Property Let LstUpdUserId(ByVal strValue As String)
  //   Public      Property Get Mask(ByVal strTagIn As String) As String
  //   Public      Property Get MaxCharacters(ByVal strTagIn As String) As Long
  //   Public      Property Get ProdCd() As String
  //   Public      Property Let ProdCd(ByVal strValue As String)
  //   Public      Property Get PrfndAbbrNm() As String
  //   Public      Property Let PrfndAbbrNm(ByVal strValue As String)
  //   Public      Property Get PrfndFrzDt() As Date
  //   Public      Property Let PrfndFrzDt(ByVal dteValue As Date)
  //   Public      Property Get PrfndTypCd() As String
  //   Public      Property Let PrfndTypCd(ByVal strValue As String)

  //   Public      Property Get ShouldForceToUppercase(ByVal strTagIn As String) As Boolean
  //   Public      AddRecord() as Boolean
  //   Public      CheckForAnotherUsersChanges(ByVal lngWhatOperation As enumWhatOperationIsBeingAttempted, _
  //                   ByRef strACF2 As String) As Long
  //   Public      DeleteRecord() As Boolean
  //   Public      GetLookupData() As Boolean
  //   Public      GetSingleRecord(ByVal strFundCdKey As String,
  //                   ByValstrProdCdKey As String, Optional ByVal
  //                   bSynchLookupRST As Boolean = False ) As Boolean
  //   Public      GoToFirstRecord()
  //   Public      GoToLastRecord()
  //   Public      GoToNextRecord()
  //   Public      GoToPreviousRecord()
  //   Public      UpdateRecord() As Boolean
  //   Private     fnGetColMetaData(ByRef pudtCol As udtColumn, _
  //                   ByRef prstIn As ADODB.Recordset) As Boolean
  //   Private     fnGetProperty(ByVal strTagIn As String) As udtColumn
  //   Private     GetRelativeRecord(ByVal strFundCdKey As String, _
  //                   ByVal strProdCdKey As String, _
  //                   ByVal lngPositionDirection As enumPositionDirection) As Boolean
  //   Private     fnLoadColMetaData() As Boolean
  //   Private     fnSelectRecord(ByVal strFundCdKey As String, ByVal strProdCdKey As String)
  //                   As ADODB.Recordset
  //
  //-----------------------------------------------------------------------------
//Option Explicit
  *Option Compare Binary

  //!CUSTOMIZE! Change both the filename and class name to represent the main table.
  //!CUSTOMIZE! Change mcstrName to reflect the class name, followed by a period.
  private static final String MCSTRNAME = "ctsrlStateRule.";

  //...............................................................................................
  //!CUSTOMIZE!
  // These are the private variables corresponding to PUBLIC properties.
  // There should be one (of type udtColumn) for each column in the table that this class accesses.
  //...............................................................................................
  private udtColumn m_strFundCd;
  private udtColumn m_strProdCd;
  private udtColumn m_varPrfndFrzDt;
  private udtColumn m_strPrfndTypCd;
  private udtColumn m_strPrfndAbbrNm;
  private udtColumn m_dteLstUpdDtm;
  private udtColumn m_strLstUpdUserId;


  //...............................................................................................
  //!CUSTOMIZE!
  // Create one Const for each column in the table, defining the table column to which it refers.
  //...............................................................................................
  private static final String MCSTRFUNDCD = "FUND_CD";
  private static final String MCSTRPRODCD = "PROD_CD";
  private static final String MCSTRPRFNDFRZDT = "PRFND_FRZ_DT";
  private static final String MCSTRPRFNDTYPCD = "PRFND_TYP_CD";
  private static final String MCSTRPRFNDABBRNM = "PRFND_ABBR_NM";
  private static final String MCSTRLSTUPDDTM = "LST_UPD_DTM";
  private static final String MCSTRLSTUPDUSERID = "LST_UPD_USER_ID";



  //...............................................................................................
  // Other private variables that do NOT correspond to PUBLIC properties.
  //...............................................................................................
  // m_adwADO is a private instantiation of the ADO Wrapper, used to do ADO things like
  // navigation, executing a stored procedure, etc.
  private cadwADOWrapper m_adwADO;

  // The next 2 vars (m_dteLstUpdDtm_Original and m_strLstUpdUserId_Original) are used by
  // the CheckForAnotherUsersChanges method to determine if another user affected the
  // record since *this* user originally retrieved the record.
  private Date m_dteLstUpdDtm_Original = null;
  private String m_strLstUpdUserId_Original = "";

  // m_rstLookup contains selected columns for each row in the table and is used by the form
  // to populate its Lookup VSFlexGrid control that the user uses to hop directly to a desired record.
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
      ctclmClaim.getLookupData();

      // Get all columns for the 1st record in the Lookup RST and load to member vars.
      // If there are no records (m_rstLookup is Nothing), then initialize the
      // properties that correspond to table columns. (Caller must take action if
      // m_rstLookup Is Nothing!!!)
      if (m_rstLookup.RecordCount != 0) {
        //!CUSTOMIZE! The constant referenced below should refer to the key field.
        ctclmClaim.getSingleRecord(m_rstLookup.Fields(MCSTRFUNDCD).value, m_rstLookup.Fields(MCSTRPRODCD).value);
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
  public String getFundCd() {
    String _rtn = "";
    // **************************************************************************
    // Function  : Property Get FundCd
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : string
    // **************************************************************************
    "Property Get FundCd"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = CStr(m_strFundCd.value);
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
  public void setFundCd(String strValue) {
    String _rtn = null;
    // **************************************************************************
    // Function  : Property Let FundCd
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal NewValue As String
    // Returns   :
    // **************************************************************************
    "Property Let FundCd"
.equals(Const cstrCurrentProc As String);
    try {

      m_strFundCd.value = strValue;

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
  public Date getLstUpdDtm() {
    Date _rtn = null;
    // **************************************************************************
    // Function  : Property Get LstUpdDtm
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : Int
    // **************************************************************************
    "Property Get LstUpdDtm"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = G.parseDate(m_dteLstUpdDtm.value);
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
  public void setLstUpdDtm(Date dteValue) {
    Date _rtn = null;
    // **************************************************************************
    // Function  : Property Let LstUpdDtm
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal NewValue As Integer
    // Returns   :
    // **************************************************************************
    "Property Let LstUpdDtm"
.equals(Const cstrCurrentProc As String);
    try {

      m_dteLstUpdDtm.value = dteValue;
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
  public String getLstUpdUserId() {
    String _rtn = "";
    // **************************************************************************
    // Function  : Property Get LstUpdUserId
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : string
    // **************************************************************************
    "Property Get LstUpdUserId"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = CStr(m_strLstUpdUserId.value);
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
  public void setLstUpdUserId(String strValue) {
    String _rtn = null;
    // **************************************************************************
    // Function  : Property Let LstUpdUserId
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal NewValue As string
    // Returns   :
    // **************************************************************************
    "Property Let LstUpdUserId"
.equals(Const cstrCurrentProc As String);
    try {

      m_strLstUpdUserId.value = strValue;
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
  public String getPrfndAbbrNm() {
    String _rtn = "";
    // **************************************************************************
    // Function  : Property Get PrfndAbbrNm
    // Purpose   : Retrieves current setting from property
    // Parameters: N/A
    // Returns   : string
    // **************************************************************************
    "Property Get PrfndAbbrNm"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = CStr(m_strPrfndAbbrNm.value);
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
  public void setPrfndAbbrNm(String strValue) {
    String _rtn = null;
    // **************************************************************************
    // Function  : Property Let PrfndAbbrNm
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal strValue As String
    // Returns   :
    // **************************************************************************
    "Property Let PrfndAbbrNm"
.equals(Const cstrCurrentProc As String);
    try {

      m_strPrfndAbbrNm.value = strValue;
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
  public Object getPrfndFrzDt() {
    Object _rtn = null;
    // **************************************************************************
    // Function  : Property Get PrfndFrzDt
    // Purpose   : This is a nullable date field, so it must be defined as variant
    //             so it can contain a Null value without getting a runtime
    //             error 94 (invalid use of Null).
    // Parameters: N/A
    // Returns   : string
    // **************************************************************************
    "Property Get PrfndFrzDt"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = m_varPrfndFrzDt.value;
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
  public void setPrfndFrzDt(Object varValue) {
    Object _rtn = null;
    // **************************************************************************
    // Function  : Property Let PrfndFrzDt
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal NewValue As string
    // Returns   :
    // **************************************************************************
    "Property Let PrfndFrzDt"
.equals(Const cstrCurrentProc As String);
    try {

      m_varPrfndFrzDt.value = varValue;
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
  public String getPrfndTypCd() {
    String _rtn = "";
    // **************************************************************************
    // Function  : Property Get PrfndTypCd
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : string
    // **************************************************************************
    "Property Get PrfndTypCd"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = CStr(m_strPrfndTypCd.value);
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
  public void setPrfndTypCd(String strValue) {
    String _rtn = null;
    // **************************************************************************
    // Function  : Property Let PrfndTypCd
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal strValue As String
    // Returns   :
    // **************************************************************************
    "Property Let PrfndTypCd"
.equals(Const cstrCurrentProc As String);
    try {

      m_strPrfndTypCd.value = strValue;
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
  public String getProdCd() {
    String _rtn = "";
    // **************************************************************************
    // Function  : Property Get ProdCd
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : string
    // **************************************************************************
    "Property Get ProdCd"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = CStr(m_strProdCd.value);
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
  public void setProdCd(String strValue) {
    String _rtn = null;
    // **************************************************************************
    // Function  : Property Let ProdCd
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal strValue As String
    // Returns   :
    // **************************************************************************
    "Property Let ProdCd"
.equals(Const cstrCurrentProc As String);
    try {

      m_strProdCd.value = strValue;
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

    //!CUSTOMIZE!  Customize the name of the stored procedure, the number and names
    //             of the parameters and, perhaps, the return values trapped in the
    //             error handler.

    "AddRecord"
.equals(Const cstrCurrentProc As String);
    //' Stored procedure to execute
    "dbo.proc_product_fund_insert"
.equals(Const cstrSproc As String);
    DBRecordSet rstTemp = null;
    boolean bSuccessful = false;
    ADODB.Parameter prmReturnValue = null;
    ADODB.Parameter prmProdCd = null;
    ADODB.Parameter prmFundCd = null;
    ADODB.Parameter prmPrfndFrzDt = null;
    ADODB.Parameter prmPrfndTypCd = null;
    ADODB.Parameter prmPrfndAbbrNm = null;
    ADODB.Parameter prmInvalid_Key = null;

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
      // Define the PROD_CD parameter
      prmProdCd = w_aDOCommand.CreateParameter(Name:="prod_cd", Type:=adChar, Direction:=adParamInput, Size:=8, .value:=fnNullIfZLS(varIn:=ProdCd, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmProdCd);

      // ---Parameter #3---
      // Define the FUND_CD parameter
      prmFundCd = w_aDOCommand.CreateParameter(Name:="fund_cd", Type:=adChar, Direction:=adParamInput, Size:=8, .value:=fnNullIfZLS(varIn:=FundCd, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmFundCd);

      // ---Parameter #4---
      // Define the PRFND_FRZ_DT parameter
      // This is a nullable column coming from a control that can be valued Null. Don't call
      // fnNullIfZLS or fnNullIfZero as they shouldn't be necessary.
      prmPrfndFrzDt = w_aDOCommand.CreateParameter(Name:="PrfndFrzDt", Type:=adDBTimeStamp, Direction:=adParamInput, .value:=PrfndFrzDt);
      w_aDOCommand.Parameters.Append(prmPrfndFrzDt);

      // ---Parameter #5---
      // Define the PRFND_TYP_CD parameter
      prmPrfndTypCd = w_aDOCommand.CreateParameter(Name:="prfnd_typ_cd", Type:=adChar, Direction:=adParamInput, Size:=8, .value:=fnNullIfZLS(varIn:=PrfndTypCd, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmPrfndTypCd);

      // ---Parameter #6---
      // Define the prfnd_abbr_nm parameter
      prmPrfndAbbrNm = w_aDOCommand.CreateParameter(Name:="prfnd_abbr_nm", Type:=adChar, Direction:=adParamInput, Size:=4, .value:=fnNullIfZLS(varIn:=PrfndAbbrNm, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmPrfndAbbrNm);

      // ---Parameter #7---
      // Define the Invalid_Key output parameter, which reflects *which* foreign
      // key violation was encountered.
      prmInvalid_Key = w_aDOCommand.CreateParameter(Name:="@Invalid_Key", Type:=adVarChar, Size:=255, Direction:=adParamOutput);
      w_aDOCommand.Parameters.Append(prmInvalid_Key);

      // Do the Add
      w_aDOCommand.Execute;

      //...........................................................................
      // Refresh the Lookup recordset, re-retrieve the just-added record so that
      // record is *still* the current record, and load its data to the
      // table wrapper's class properties so all table columns (including
      // those set by the DBMS like identity and Last Updated columns) are
      // up-to-date.
      //...........................................................................
      bSuccessful = ctclmClaim.getRelativeRecord(getFundCd(), getProdCd(), enumPositionDirection.ePDSAMERECORD);

      _rtn = true;
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    modGeneral.fnFreeRecordset(rstTemp);
    modGeneral.fnFreeObject(prmReturnValue);
    modGeneral.fnFreeObject(prmProdCd);
    modGeneral.fnFreeObject(prmFundCd);
    modGeneral.fnFreeObject(prmPrfndFrzDt);
    modGeneral.fnFreeObject(prmPrfndTypCd);
    modGeneral.fnFreeObject(prmPrfndAbbrNm);
    modGeneral.fnFreeObject(prmInvalid_Key);

    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (prmReturnValue) {
      //' 4027
      case  modResConstants.gCRES_NERR_REC_NOT_FOUND      :
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_REC_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, "Product "+ RTrim$(getProdCd())+ "/ Fund "+ RTrim$(getFundCd()));
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
      //' 4028
        break;

      case  modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO  :
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO, MCSTRNAME+ cstrCurrentProc, "add");
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
      //' 4031
        break;

      case  modResConstants.gCRES_NERR_ADD_WITH_NONUNIQUE_KEY  :
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_ADD_WITH_NONUNIQUE_KEY, MCSTRNAME+ cstrCurrentProc, "Product "+ RTrim$(getProdCd())+ "/ Fund "+ RTrim$(getFundCd()));
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
      //' 4032
        break;

      case  modResConstants.gCRES_NERR_KEY_NOT_FOUND:
        if ("PROD_CD"
.equals(prmInvalid_Key.toUpperCase())) {
          modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_KEY_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, "Product Cd", RTrim$(getProdCd()), "Product");
          /**TODO:** resume found: Resume(PROC_EXIT)*/;
        } 
        else if ("FUND_CD"
.equals(prmInvalid_Key.toUpperCase())) {
          modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_KEY_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, "Fund Cd", RTrim$(getFundCd()), "Fund");
          /**TODO:** resume found: Resume(PROC_EXIT)*/;
        } 
        else {
          // PRFND_TYP_CD
          modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_KEY_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, "Product Fund Type Cd", RTrim$(getPrfndTypCd()), "Product Fund Type");
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
    Const(clngNoError As Long == 0);
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
      rstSingleRecord_Fresh = fnSelectRecord(getFundCd(), getProdCd());

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
        if ((DateDiff("s", CStr(m_dteLstUpdDtm_Original), CStr(!lst_upd_dtm)) != 0) || (!lst_upd_user_id != m_strLstUpdUserId_Original)) {
          strACF2 = !lst_upd_user_id;
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
    "dbo.proc_product_fund_delete"
.equals(Const cstrSproc As String);
    boolean bSuccessful = false;
    DBRecordSet rstTemp = null;
    ADODB.Parameter prmReturnValue = null;
    ADODB.Parameter prmFundCd = null;
    ADODB.Parameter prmProdCd = null;
    ADODB.Parameter prmDependent_Table = null;

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
      // Define the PROD_CD parameter
      prmProdCd = w_aDOCommand.CreateParameter(Name:="prod_cd", Type:=adChar, Direction:=adParamInput, Size:=8, .value:=fnNullIfZLS(varIn:=ProdCd, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmProdCd);

      // ---Parameter #3---
      // Define the FUND_CD parameter
      prmFundCd = w_aDOCommand.CreateParameter(Name:="fund_cd", Type:=adChar, Direction:=adParamInput, Size:=8, .value:=fnNullIfZLS(varIn:=FundCd, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmFundCd);

      // ---Parameter #4---
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
      bSuccessful = ctclmClaim.getRelativeRecord(getFundCd(), getProdCd(), enumPositionDirection.ePDPREVIOUSRECORD);

      _rtn = true;
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    modGeneral.fnFreeRecordset(rstTemp);
    modGeneral.fnFreeObject(prmReturnValue);
    modGeneral.fnFreeObject(prmProdCd);
    modGeneral.fnFreeObject(prmDependent_Table);

    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (prmReturnValue) {
      //' 4027
      case  modResConstants.gCRES_NERR_REC_NOT_FOUND      :
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_REC_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, "Product "+ RTrim$(getProdCd())+ "/ Fund "+ RTrim$(getFundCd()));
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
      //' 4028
        break;

      case  modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO  :
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO, MCSTRNAME+ cstrCurrentProc, "add");
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
      //' 4029
        break;

      case  modResConstants.gCRES_NERR_DEPENDENT_RECS_EXIST  :
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_DEPENDENT_RECS_EXIST, MCSTRNAME+ cstrCurrentProc, "Product "+ RTrim$(getProdCd())+ "/ Fund "+ RTrim$(getFundCd()), prmDependent_Table.toUpperCase());
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
    //' Stored procedure to execute
    "dbo.proc_product_fund_lu2_select"
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
      //' 4028
      case  modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO  :
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
  public boolean getRelativeRecord(String strFundCdKey, String strProdCdKey, enumPositionDirection lngPositionDirection) {
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
    //                    and a local var (i.e. strProdCdKeyForNewRec) of the
    //                    appropriate data type! Also, the setting of
    //                    the .Filter property below must reflect each
    //                    table key.
    //
    // Params:
    //     strProdCdKey         (in) = ProdCd value from which to do the relative
    //                                 repositioning
    //     strFundCdKey         (in) = FundCd value from which to do the relative
    //                                 repositioning
    //     lngPositionDirection (in) = Indicates to which relative record the
    //                                 recordset should be positioned (relative
    //                                 to the strProdCdKey parameter value).
    //
    //
    // Called By:   cmdDelete_Click( ) of frmProduct
    //              cmdUpdate_Click( ) of frmProduct
    //              cmdNavigate_Click( ) of frmProduct
    //              fnAddRecord( ) of frmProduct
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
    Const(cintNoRecords As Integer == 0);
    DBRecordSet rstTemp = null;
    String strFundCdKeyForNewRec = "";
    String strProdCdKeyForNewRec = "";

    try {

      //...........................................................................
      // Refresh the lookup data (m_rstLookupData) so other's changes
      // --and our own-- are now reflected in it. This resets the Lookup data,
      // record count, and current record number, and leaves the Lookup recordset
      // positioned to the first record (if there are records) or BOF (if there are
      // no records).
      //...........................................................................
      ctclmClaim.getLookupData();

      // Since this table has a 2-part key, this function is different than the one
      // in single-key-column table wrappers. epdPreviousRecord, for instance,
      // should try to go to the record with the same Key1 (i.e. same Key1 but
      // lesser Key2) if it exists, otherwise go to the record with a lesser Key1.
      //
      // WARNING: Make sure your filter is in-sync with the actual order of the data
      // per the sproc that GetLookupData method calls; otherwise you won't get
      // the right record!
      if (m_rstLookup.RecordCount > 0) {
        switch (lngPositionDirection) {
          case  enumPositionDirection.ePDPREVIOUSRECORD:
            // Walk backwards until we find a record with a lesser key. If none found,
            // show the first one (if there is one).
            m_rstLookup.MoveLast;
            do Until ((!fund_cd = strFundCdKey And !prod_cd < strProdCdKey) Or (!fund_cd < strFundCdKey)) Or .EOF Or .BOF              m_rstLookup.MovePrevious;
              if (m_rstLookup.BOF || m_rstLookup.EOF) {
                break;
              }
            }
            if ((m_rstLookup.EOF || m_rstLookup.BOF)) {
              // No records prior to the specified Fund Cd/Prod Cd.
              // Show the first record, if there is one.
              if (m_rstLookup.RecordCount != 0) {
                m_adwADO.moveFirst(m_rstLookup);
              }
              // Else
              //     we're already at the record we want, so do nothing.
            }

            break;

          case  enumPositionDirection.ePDNEXTRECORD:
            // Walk forwards until we find a record with a higher key. If none found,
            // show the last one (if there is one).
            m_rstLookup.MoveFirst;
            do Until ((!fund_cd = strFundCdKey And !prod_cd > strProdCdKey) Or (!fund_cd > strFundCdKey)) Or .EOF Or .BOF              m_rstLookup.MoveNext;
            }
            if ((m_rstLookup.EOF || m_rstLookup.BOF)) {
              // No records prior to the specified Fund Cd/Prod Cd.
              // Show the first record, if there is one.
              if (m_rstLookup.RecordCount != 0) {
                m_adwADO.moveLast(m_rstLookup);
              }
              // Else
              //     we're already at the record we want, so do nothing.
            }

            break;

          case  enumPositionDirection.ePDSAMERECORD:
            // This is used by Update processing, where we just want to
            // stay on the just-updated record but make sure its current
            // data (from the DBMS) is loaded to the class properties and
            // other's changes to any record in the Lookup recordset are
            // visible.

            // Walk forwards until we find a record with the desired "exact match" key.
            m_rstLookup.MoveFirst;
            do Until (!fund_cd = strFundCdKey And !prod_cd = strProdCdKey) Or .EOF Or .BOF              m_rstLookup.MoveNext;
            }
            if ((m_rstLookup.EOF || m_rstLookup.BOF)) {
              // If the record wasn't found, generate an error. We should
              // never hit this error, except due to bad program logic, since
              // it means the current record could not be found.
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
        ctclmClaim.getSingleRecord(m_rstLookup.Fields(MCSTRFUNDCD).value, m_rstLookup.Fields(MCSTRPRODCD).value);
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
  public boolean getSingleRecord(String strFundCdKey, String strProdCdKey, boolean bSynchLookupRST) {
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
    //    strFundCdKey    (in) = represents the FUND_CD primary key for the table
    //    strProdCdKey    (in) = represents the PROD_CD primary key for the table. If
    //                           this argument is omitted, a lookup by partial key
    //                           rather than full key will be done.
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
    "????????"
.equals(Const cstrMissingProdCd As String);
    DBRecordSet rstTemp = null;
    boolean bFoundIt = false;

    try {

      //!CUSTOMIZE! fnSelectRecord call should pass the key column(s)
      if (strProdCdKey.equals(cstrMissingProdCd)) {
        rstTemp = fnSelectRecordByPartialKey(strFundCdKey);
      } 
      else {
        rstTemp = fnSelectRecord(strFundCdKey, strProdCdKey);
      }

      if (rstTemp.RecordCount == 0) {
        // Clear properties and return False to make sure caller can tell this function failed to get a record
        fnClearPropertyValues();
        // **TODO:** goto found: GoTo PROC_EXIT;
      }

      // A recordset is always returned from the above sprocs. If it has no rows,
      // then the following "With Me" block will act just like fnClearPropertyValues( ).
      //*TODO:** can't found type for with block
      //*With this
      __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = this;
      // For Char/VarChar fields,
      //     * Use fnZLSIfNull to ensure Nulls are appropriately translated.
      // For Numeric fields,
      //     * Use fnZeroIfNull to ensure Nulls are appropriately translated.
      // For Boolean fields,
      //     * Use fnYNToBool to ensure True/False is appropriately translated.

      w___TYPE_NOT_FOUND.FundCd = modDataConversion.fnZLSIfNull(rstTemp!fund_cd);
      w___TYPE_NOT_FOUND.ProdCd = modDataConversion.fnZLSIfNull(rstTemp!prod_cd);
      // The following is a nullable date field. If it's Null, let it go through as Null (no fnZLSIfNull call)
      w___TYPE_NOT_FOUND.PrfndFrzDt = rstTemp!prfnd_frz_dt;
      w___TYPE_NOT_FOUND.PrfndTypCd = modDataConversion.fnZLSIfNull(rstTemp!prfnd_typ_cd);
      w___TYPE_NOT_FOUND.PrfndAbbrNm = modDataConversion.fnZLSIfNull(rstTemp!prfnd_abbr_nm);
      w___TYPE_NOT_FOUND.LstUpdDtm = modDataConversion.fnZLSIfNull(rstTemp!lst_upd_dtm);
      w___TYPE_NOT_FOUND.LstUpdUserId = modDataConversion.fnZLSIfNull(rstTemp!lst_upd_user_id);

      // Save original Last Updated info, to be used during UpdateRecord( ) and DeleteRecord( )
      // to determine if another user updated the record since it was retrieved.
      m_dteLstUpdDtm_Original = modDataConversion.fnZLSIfNull(rstTemp!lst_upd_dtm);
      m_strLstUpdUserId_Original = modDataConversion.fnZLSIfNull(rstTemp!lst_upd_user_id);

      if (bSynchLookupRST) {
        m_adwADO.moveFirst(m_rstLookup);
        // .Find method supports single-column searching only. Thus we have to do a find
        // on the Key1 value then loop through records until we find the one with
        // the Key2 value we want.
        m_rstLookup.Find("fund_cd = '"+ getFundCd()+ "'");
        // Start looping until we find a match. The only time we *shouldn't* find a match
        // or have no records is due to a multi-user scenario, so let it error.
        if (m_rstLookup.EOF) {
          modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO, MCSTRNAME+ cstrCurrentProc, "locate");
          /**TODO:** resume found: Resume(PROC_EXIT)*/;
        } 
        else {
          // Stop upon exact match on desired Key1 and Key2 values
          while ((!fund_cd == getFundCd()) && (Not bFoundIt)) {
            if (!prod_cd == getProdCd()) {
              bFoundIt = true;
            } 
            else {
              m_rstLookup.MoveNext;
              if (m_rstLookup.EOF) {
                modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO, MCSTRNAME+ cstrCurrentProc, "locate");
                /**TODO:** resume found: Resume(PROC_EXIT)*/;
              }
            }
          }
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


  *#If False Then
  //////////////////////////////////////////////////////////////////////////////////////////////////
  public boolean getSingleRecordByPartialKey(String strFundCdKey, boolean bSynchLookupRST) {
    boolean _rtn = false;
    //--------------------------------------------------------------------------
    // Procedure:   GetSingleRecordByPartialKey
    // Description: Obtains data from the database for the specified partial key(s).
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
    //              NOTE: This is very similar to GetSingleRecord
    //                    except it takes only a partial key.
    //
    // Returns:     Boolean
    // Params:
    //    strFundCdKey         (in) = represents the FUND_CD primary key for the table
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
    //             the recordset's .Find property is set and the logic in the
    //             "if bSynchLookupRST then" loop must be changed
    //             to reflect each key column so the right record will be located.
    //             Also make sure that the substitution values passed to
    //             SaveAppSpecificError are correct and TRIM'd if appropriate.

    "GetSingleRecordByPartialKey"
.equals(Const cstrCurrentProc As String);
    DBRecordSet rstTemp = null;
    boolean bFoundIt = false;

    try {

      _rtn = false;

      //!CUSTOMIZE! fnSelectRecordByPartialKey call should pass the key column referenced on parent's screen
      // Grab all records for the specified Fund Cd (ignoring the Product Cd key). Recordset will
      // be sorted by Fund Cd, then Product Cd
      rstTemp = fnSelectRecordByPartialKey(strFundCdKey);

      if (rstTemp.RecordCount == 0) {
        // If there aren't any records for the requested Fund, then the form should go
        // into Add mode.
        fnClearPropertyValues();
        // Let the function return False
        // **TODO:** goto found: GoTo PROC_EXIT;
      }


      // Go to the first Product Cd for the specified Fund Cd
      rstTemp.MoveFirst;
      //*TODO:** can't found type for with block
      //*With this
      __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = this;
      // For Char/VarChar fields,
      //     * Use fnZLSIfNull to ensure Nulls are appropriately translated.
      // For Numeric fields,
      //     * Use fnZeroIfNull to ensure Nulls are appropriately translated.
      // For Boolean fields,
      //     * Use fnYNToBool to ensure True/False is appropriately translated.
      w___TYPE_NOT_FOUND.FundCd = modDataConversion.fnZLSIfNull(rstTemp!fund_cd);
      w___TYPE_NOT_FOUND.ProdCd = modDataConversion.fnZLSIfNull(rstTemp!prod_cd);
      // The following is a nullable date field. If it's Null, let it go through as Null (no fnZLSIfNull call)
      w___TYPE_NOT_FOUND.PrfndFrzDt = rstTemp!prfnd_frz_dt;
      w___TYPE_NOT_FOUND.PrfndTypCd = modDataConversion.fnZLSIfNull(rstTemp!prfnd_typ_cd);
      w___TYPE_NOT_FOUND.LstUpdDtm = modDataConversion.fnZLSIfNull(rstTemp!lst_upd_dtm);
      w___TYPE_NOT_FOUND.LstUpdUserId = modDataConversion.fnZLSIfNull(rstTemp!lst_upd_user_id);

      // Save original Last Updated info, to be used during UpdateRecord( ) and DeleteRecord( )
      // to determine if another user updated the record since it was retrieved.
      m_dteLstUpdDtm_Original = modDataConversion.fnZLSIfNull(rstTemp!lst_upd_dtm);
      m_strLstUpdUserId_Original = modDataConversion.fnZLSIfNull(rstTemp!lst_upd_user_id);

      if (bSynchLookupRST) {
        m_adwADO.moveFirst(m_rstLookup);
        // .Find method supports single-column searching only. Thus we have to do a find
        // on the Key1 value then loop through records until we find the one with
        // the Key2 value we want.
        m_rstLookup.Find("fund_cd = '"+ getFundCd()+ "'");
        // Start looping until we find a match. The only time we *shouldn't* find a match
        // or have no records is due to a multi-user scenario, so let it error.
        if (m_rstLookup.EOF) {
          modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO, MCSTRNAME+ cstrCurrentProc, "locate");
          /**TODO:** resume found: Resume(PROC_EXIT)*/;
        } 
        else {
          // Stop upon exact match on desired Key1 and Key2 values
          while ((!fund_cd == getFundCd()) && (Not bFoundIt)) {
            if (!prod_cd == getProdCd()) {
              bFoundIt = true;
            } 
            else {
              m_rstLookup.MoveNext;
              if (m_rstLookup.EOF) {
                modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO, MCSTRNAME+ cstrCurrentProc, "locate");
                /**TODO:** resume found: Resume(PROC_EXIT)*/;
              }
            }
          }

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
  *#End If



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
        ctclmClaim.getSingleRecord(m_rstLookup.Fields(MCSTRFUNDCD).value, m_rstLookup.Fields(MCSTRPRODCD).value);
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
        ctclmClaim.getSingleRecord(m_rstLookup.Fields(MCSTRFUNDCD).value, m_rstLookup.Fields(MCSTRPRODCD).value);
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
          ctclmClaim.getLookupData();
          ctclmClaim.getRelativeRecord(getFundCd(), getProdCd(), enumPositionDirection.ePDNEXTRECORD);
        } 
        else {
          // Get the requested record and reposition the Lookup recordset to that record
          //!CUSTOMIZE! The constant referenced below should refer to the key field.
          ctclmClaim.getSingleRecord(m_rstLookup.Fields(MCSTRFUNDCD).value, m_rstLookup.Fields(MCSTRPRODCD).value);
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
          ctclmClaim.getLookupData();
          ctclmClaim.getRelativeRecord(getFundCd(), getProdCd(), enumPositionDirection.ePDPREVIOUSRECORD);
        } 
        else {
          // Get the requested record and reposition the Lookup recordset to that record
          //!CUSTOMIZE! The constant referenced below should refer to the key field.
          ctclmClaim.getSingleRecord(m_rstLookup.Fields(MCSTRFUNDCD).value, m_rstLookup.Fields(MCSTRPRODCD).value);
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
  public boolean haveDependents(String strFundCdKey, String strProdCdKey, String strDependentTable) { // TODO: Use of ByRef founded Public Function HaveDependents(ByVal strFundCdKey As String, ByVal strProdCdKey As String, ByRef strDependentTable As String) As Boolean
    boolean _rtn = false;
    // Comments  : Determines whether the current record can be deleted without
    //             hitting a referential integrity violation due to either:
    //             a. row(s) existing in other tables that use the current key value
    //                as a foreign key
    //             The calling form should look at the return value. If True, then
    //             the form's Delete button should be disabled.
    //
    // Parameters:
    //   strProdCdKey = value of the first key to this table (prod_cd)
    //   strFundCdKey = value of the second key to this table (fund_cd)
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
    "dbo.proc_product_fund_verify_dependents"
.equals(Const cstrSproc As String);
    cadwADOWrapper adwTemp = null;
    DBRecordSet rstTemp = null;
    ADODB.Parameter prmReturnValue = null;
    ADODB.Parameter prmFundCd = null;
    ADODB.Parameter prmProdCd = null;
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
      prmProdCd = w_aDOCommand.CreateParameter(Name:="@prod_cd", Type:=adChar, Direction:=adParamInput, Size:=8, .value:=strProdCdKey);
      w_aDOCommand.Parameters.Append(prmProdCd);

      // ---Parameter #3---
      // Define the input parameter that represents the key value to see
      // if it exists as a foreign key on dependent tables
      prmFundCd = w_aDOCommand.CreateParameter(Name:="@fund_cd", Type:=adChar, Direction:=adParamInput, Size:=8, .value:=strFundCdKey);
      w_aDOCommand.Parameters.Append(prmFundCd);

      // ---Parameter #4---
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
    modGeneral.fnFreeObject(prmProdCd);
    modGeneral.fnFreeObject(prmFundCd);
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
      //' 4027
      case  modResConstants.gCRES_NERR_REC_NOT_FOUND          :
        _rtn = true;
        //' This is actually ignored by the caller
        strDependentTable = "Unknown";
        // Remove any trace that this error occurred since we're not going to report it as an error.
        VBA.ex.Clear;
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
      //' 4029
        break;

      case  modResConstants.gCRES_NERR_DEPENDENT_RECS_EXIST   :
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

    //!CUSTOMIZE!  Customize the name of the stored procedure, the number and names
    //             of the parameters and, perhaps, the return values trapped in the
    //             error handler.

    "UpdateRecord"
.equals(Const cstrCurrentProc As String);
    //' Stored procedure to execute
    "dbo.proc_product_fund_update"
.equals(Const cstrSproc As String);
    DBRecordSet rstTemp = null;
    boolean bSuccessful = false;
    ADODB.Parameter prmReturnValue = null;
    ADODB.Parameter prmProdCd = null;
    ADODB.Parameter prmFundCd = null;
    ADODB.Parameter prmPrfndFrzDt = null;
    ADODB.Parameter prmPrfndTypCd = null;
    ADODB.Parameter prmPrfndAbbrNm = null;
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
      // Define the PROD_CD parameter
      prmProdCd = w_aDOCommand.CreateParameter(Name:="prod_cd", Type:=adChar, Direction:=adParamInput, Size:=8, .value:=fnNullIfZLS(varIn:=ProdCd, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmProdCd);

      // ---Parameter #3---
      // Define the FUND_CD parameter
      prmFundCd = w_aDOCommand.CreateParameter(Name:="fund_cd", Type:=adChar, Direction:=adParamInput, Size:=8, .value:=fnNullIfZLS(varIn:=FundCd, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmFundCd);

      // ---Parameter #4---
      // Define the PRFND_FRZ_DT parameter
      prmPrfndFrzDt = w_aDOCommand.CreateParameter(Name:="prfnd_frz_dt", Type:=adDBTimeStamp, Direction:=adParamInput, Size:=8, .value:=PrfndFrzDt);
      //value:=fnNullIfZLS(varIn:=PrfndFrzDt, bHandleEmbeddedQuotes:=True))
      w_aDOCommand.Parameters.Append(prmPrfndFrzDt);

      // ---Parameter #5---
      // Define the PRFND_TYP_CD parameter
      prmPrfndTypCd = w_aDOCommand.CreateParameter(Name:="prfnd_typ_cd", Type:=adChar, Direction:=adParamInput, Size:=8, .value:=fnNullIfZLS(varIn:=PrfndTypCd, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmPrfndTypCd);

      // ---Parameter #6---
      // Define the prfnd_abbr_nm parameter
      prmPrfndAbbrNm = w_aDOCommand.CreateParameter(Name:="prfnd_abbr_nm", Type:=adChar, Direction:=adParamInput, Size:=4, .value:=fnNullIfZLS(varIn:=PrfndAbbrNm, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmPrfndAbbrNm);

      // ---Parameter #7---
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
      bSuccessful = ctclmClaim.getRelativeRecord(getFundCd(), getProdCd(), enumPositionDirection.ePDSAMERECORD);

      _rtn = true;
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    modGeneral.fnFreeRecordset(rstTemp);
    modGeneral.fnFreeObject(prmReturnValue);
    modGeneral.fnFreeObject(prmProdCd);
    modGeneral.fnFreeObject(prmFundCd);
    modGeneral.fnFreeObject(prmPrfndFrzDt);
    modGeneral.fnFreeObject(prmPrfndTypCd);
    modGeneral.fnFreeObject(prmPrfndAbbrNm);
    modGeneral.fnFreeObject(prmInvalid_Key);

    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (prmReturnValue) {
      //' 4027
      case  modResConstants.gCRES_NERR_REC_NOT_FOUND      :
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_REC_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, "Product "+ RTrim$(getProdCd())+ "/ Fund "+ RTrim$(getFundCd()));
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
      //' 4028
        break;

      case  modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO  :
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO, MCSTRNAME+ cstrCurrentProc, "update");
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
      //' 4032
        break;

      case  modResConstants.gCRES_NERR_KEY_NOT_FOUND:
        if ("PROD_CD"
.equals(prmInvalid_Key.toUpperCase())) {
          modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_KEY_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, "Product Cd", RTrim$(getProdCd()), "Product");
          /**TODO:** resume found: Resume(PROC_EXIT)*/;
        } 
        else if ("FUND_CD"
.equals(prmInvalid_Key.toUpperCase())) {
          modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_KEY_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, "Fund Cd", RTrim$(getFundCd()), "Fund");
          /**TODO:** resume found: Resume(PROC_EXIT)*/;
        } 
        else {
          // PRFND_TYP_CD
          modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_KEY_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, "Product Fund Type Cd", RTrim$(getPrfndTypCd()), "Product Fund Type");
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
    //                        Booleans => False
    //                        Dates => Now

    "fnClearPropertyValues"
.equals(Const cstrCurrentProc As String);

    try {

      //*TODO:** can't found type for with block
      //*With this
      __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = this;
      w___TYPE_NOT_FOUND.FundCd = "";
      w___TYPE_NOT_FOUND.ProdCd = "";
      w___TYPE_NOT_FOUND.PrfndFrzDt = Null;
      w___TYPE_NOT_FOUND.PrfndTypCd = "";
      w___TYPE_NOT_FOUND.PrfndAbbrNm = "";
      w___TYPE_NOT_FOUND.LstUpdDtm = Now;
      w___TYPE_NOT_FOUND.LstUpdUserId = "";

      // Also reset the saved "original" values for the Last Updated info
      m_dteLstUpdDtm_Original = Now;
      m_strLstUpdUserId_Original = "";
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
      "0"
.equals(Const cstrZeroChar As String);
      "."
.equals(Const cstrDecimalChar As String);
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
      //       Dates/lst_upd_dtm = adDBTimeStamp   (e.g., lst_upd_dtm or eff_dt)
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
      //       dom_lst_upd_dtm...for dates, this sets the system date (getdate()) as the default value on an Insert
      //    e. dom_lst_upd_id....to set the user's ACF2 (suser_sname()) as the default value on an Insert
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

      // DEF_LST_UPD_ID is typically used for the Lst_Upd_User_Id column,
      // indicating to set it to the logged on user.
      if (strDefaultValueToParse.indexOf("DEF_LST_UPD_ID", 1) > 0) {
        pudtCol.defaultValue = modGeneral.gconAppActive.getLastLogOnUserID();
      }

      // DEF_LST_UPD_DTM is typically used for the Lst_Upd_Dtm column,
      // indicating to set it to the System Date.
      if (strDefaultValueToParse.indexOf("DEF_LST_UPD_DTM", 1) > 0) {
        pudtCol.defaultValue = Date;
      }

      // All TRS tables use only the "DOM_IND" domain name for indicator columns.
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
          if ("AREA_CD"
.equals(pudtCol.colName.substring(pudtCol.colName.length() - 7))) {
            // Area Codes require 3 numeric positions if input.
            // They should be displayed via a TextBox control.
            pudtCol.format = "###";
            pudtCol.mask = "";
            pudtCol.maxCharacters = 3;
          } 
          else if ("_AMT"
.equals(pudtCol.colName.substring(pudtCol.colName.length() - 4))) {
            // These are assumed to be currency with up to 14 dollar/2 decimal positions.
            // These should typically  be displayed via a TextBox control.
            // For a Grid, .Format should be "($##,###,###,###,##0.00)"
            // For any other control, .Format should be "$##,###,###,###,##0.00;($##,###,###,###,##0.00)"
            pudtCol.format = "$##,###,###,###,##0.00";
            pudtCol.mask = "";
          } 
          else if ("UNIT_QTY"
.equals(pudtCol.colName.substring(pudtCol.colName.length() - 8))) {
            // These should typically be displayed via a TextBox control.
            // These have 11 dollar/6 decimal positions and can be negative!
            // For a Grid, .Format should be "(##,###,###,##0.000000)"
            // For any other control, .Format should be "##,###,###,##0.000000;(##,###,###,##0.000000)"
            pudtCol.format = "##,###,###,##0.000000;(##,###,###,##0.000000)";
            pudtCol.mask = "";
          } 
          else {
            // These should be displayed via a TextBox control.
            //
            // NOTE: Other items like SHARES and PERCENTAGES will need to overridden
            // on a table-specific basis in fnLoadColMetaData( ) since there is no
            // easy way to recognize and process these fields.
            //
            // Example: Though Percentages should end in "_PCT", the number
            //          of decimal positions vary.
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
          if (("PHON_NUM".equals(pudtCol.colName.substring(pudtCol.colName.length() - 8))) || ("FAX_NUM".equals(pudtCol.colName.substring(pudtCol.colName.length() - 7)))) {
            // Phone and Fax Numbers should be displayed via a MaskEdBox control.
            pudtCol.format = "";
            pudtCol.mask = "###-####";
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
        case  "FUNDCD":
          _rtn = m_strFundCd;
          break;

        case  "PRODCD":
          _rtn = m_strProdCd;
          break;

        case  "PRFNDFRZDT":
          _rtn = m_varPrfndFrzDt;
          break;

        case  "PRFNDTYPCD":
          _rtn = m_strPrfndTypCd;
          break;

        case  "PRFNDABBRNM":
          _rtn = m_strPrfndAbbrNm;
          break;

        case  "LSTUPDDTM":
          _rtn = m_dteLstUpdDtm;
          break;

        case  "LSTUPDUSERID":
          _rtn = m_strLstUpdUserId;
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
      m_adwADO.getMetaData_Columns("state_rule_t", rstMetaData);

      // NOTES:
      //    a. Phone Numbers are Char fields, but we have to override the meta data to ensure
      //       only numbers can be input to them.
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


      do Until .EOF        // Make sure SELECT CASE lists all table columns, including the LST_UPD_xxx!
        switch (rstMetaData("COLUMN_NAME").value.toUpperCase()) {
          case  MCSTRFUNDCD:
            fnGetColMetaData(m_strFundCd, rstMetaData);
            break;

          case  MCSTRPRODCD:
            fnGetColMetaData(m_strProdCd, rstMetaData);
            break;

          case  MCSTRPRFNDFRZDT:
            fnGetColMetaData(m_varPrfndFrzDt, rstMetaData);
            break;

          case  MCSTRPRFNDTYPCD:
            fnGetColMetaData(m_strPrfndTypCd, rstMetaData);
            break;

          case  MCSTRPRFNDABBRNM:
            fnGetColMetaData(m_strPrfndAbbrNm, rstMetaData);
            // Force to be uppercase, per screen spec
            m_strPrfndAbbrNm.shouldForceToUppercase = true;
            break;

          case  MCSTRLSTUPDDTM:
            fnGetColMetaData(m_dteLstUpdDtm, rstMetaData);
            break;

          case  MCSTRLSTUPDUSERID:
            fnGetColMetaData(m_strLstUpdUserId, rstMetaData);
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
      m_adwADO.getMetaData_PrimaryKeys("state_rule_t", rstMetaData);

      do Until .EOF        // The SELECT CASE should list all table columns
        // (though you could skip the LST_UPD_xxx columns if you change the Case Else,
        //  since these would never be a key)
        switch (rstMetaData("COLUMN_NAME").value.toUpperCase()) {
          case  MCSTRFUNDCD:
            m_strFundCd.isKey = true;
            break;

          case  MCSTRPRODCD:
            m_strProdCd.isKey = true;
            break;

          case  MCSTRPRFNDFRZDT:
            m_varPrfndFrzDt.isKey = true;
            break;

          case  MCSTRPRFNDTYPCD:
            m_strPrfndTypCd.isKey = true;
            break;

          case  MCSTRPRFNDABBRNM:
            m_strPrfndAbbrNm.isKey = true;
            break;

          case  MCSTRLSTUPDDTM:
            m_dteLstUpdDtm.isKey = true;
            break;

          case  MCSTRLSTUPDUSERID:
            m_strLstUpdUserId.isKey = true;
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
  private DBRecordSet fnSelectRecord(String strFundCdKey, String strProdCdKey) {
    //--------------------------------------------------------------------------
    // Procedure:   fnSelectRecord
    // Description: Selects a single record based on the value(s) in the
    //              properties that correspond to the table's key(s)
    //
    //              NOTE: For each table key, there should be a parameter
    //                    of the appropriate data type!
    //
    // Parameters:
    //     strProdCdKey (in) - the PROD_CD key to the table that should be retrieved
    //     strFundCdKey (in) - the FUND_CD key to the table that should be retrieved
    //
    // Returns:     A disconnected ADODB.Recordset containing all table columns
    //              for the specified key
    //-----------------------------------------------------------------------------

    //!CUSTOMIZE!  This proc and all calls to it must be customized to reflect
    //             one parameter for each key column of the table. Make sure the
    //             parameter is defined to be of the right data type.
    //             Also make sure that the substitution values passed to
    //             SaveAppSpecificError are correct and TRIM'd if appropriate.

    "fnSelectRecord"
.equals(Const cstrCurrentProc As String);
    //' Stored procedure to execute
    "dbo.proc_product_fund_select"
.equals(Const cstrSproc As String);
    DBRecordSet rstTemp = null;
    ADODB.Parameter prmReturnValue = null;
    ADODB.Parameter prmProdCd = null;
    ADODB.Parameter prmFundCd = null;

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
      // Define the PROD_CD parameter
      prmProdCd = w_aDOCommand.CreateParameter(Name:="prod_cd", Type:=adChar, Direction:=adParamInput, Size:=8, .value:=fnNullIfZLS(varIn:=strProdCdKey, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmProdCd);

      // ---Parameter #3---
      // Define the FUND_CD parameter
      prmFundCd = w_aDOCommand.CreateParameter(Name:="fund_cd", Type:=adChar, Direction:=adParamInput, Size:=8, .value:=fnNullIfZLS(varIn:=strFundCdKey, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmFundCd);

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
    modGeneral.fnFreeObject(prmProdCd);
    modGeneral.fnFreeObject(prmFundCd);

    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return null;
    // **TODO:** label found: PROC_ERR:;
    switch (prmReturnValue) {
      //' 4027
      case  modResConstants.gCRES_NERR_REC_NOT_FOUND      :
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_REC_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, "Product "+ RTrim$(strProdCdKey)+ "/ Fund "+ RTrim$(strFundCdKey));
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
      //' 4028
        break;

      case  modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO  :
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
  private DBRecordSet fnSelectRecordByPartialKey(String strFundCdKey) {
    //--------------------------------------------------------------------------
    // Procedure:   fnSelectRecordByPartialKey
    // Description: Selects a single record based on the value(s) in the
    //              properties that correspond to only one of the table's key(s)
    //              (the one that the parent screen is on)
    //
    // Parameters:
    //     strFundCdKey (in) - the FUND_CD key to the table that should be retrieved
    //
    // Returns:     A disconnected ADODB.Recordset containing all table columns
    //              for the specified key
    //-----------------------------------------------------------------------------

    //!CUSTOMIZE!  This proc can be deleted in table wrappers whose maintenance screens
    //             will never be called from another screen. If it *will* be called
    //             from another screen, then this must be customized to reflect
    //             the key(s) which the parent screen supports, that determine which
    //             record on the subordinate screen to initially display.

    "fnSelectRecordByPartialKey"
.equals(Const cstrCurrentProc As String);
    //' Stored procedure to execute
    "dbo.proc_product_fund2_select"
.equals(Const cstrSproc As String);
    DBRecordSet rstTemp = null;
    ADODB.Parameter prmReturnValue = null;
    ADODB.Parameter prmFundCd = null;

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
      // Define the FUND_CD parameter
      prmFundCd = w_aDOCommand.CreateParameter(Name:="fund_cd", Type:=adChar, Direction:=adParamInput, Size:=8, .value:=fnNullIfZLS(varIn:=strFundCdKey, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmFundCd);

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
    modGeneral.fnFreeObject(prmFundCd);

    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return null;
    // **TODO:** label found: PROC_ERR:;
    switch (prmReturnValue) {
      //' 4028
      case  modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO  :
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


case class TsrlstateruleData(
              id: Option[Int],

              )

object Tsrlstaterules extends Controller with ProvidesUser {

  val tsrlstateruleForm = Form(
    mapping(
      "id" -> optional(number),

  )(TsrlstateruleData.apply)(TsrlstateruleData.unapply))

  implicit val tsrlstateruleWrites = new Writes[Tsrlstaterule] {
    def writes(tsrlstaterule: Tsrlstaterule) = Json.obj(
      "id" -> Json.toJson(tsrlstaterule.id),
      C.ID -> Json.toJson(tsrlstaterule.id),

    )
  }

  def get(id: Int) = GetAction { implicit request =>
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.LIST_TSRLSTATERULE), { user =>
      Ok(Json.toJson(Tsrlstaterule.get(user, id)))
    })
  }

  def update(id: Int) = PostAction { implicit request =>
    Logger.debug("in tsrlstaterules.update")
    tsrlstateruleForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      tsrlstaterule => {
        Logger.debug(s"form: ${tsrlstaterule.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.EDIT_TSRLSTATERULE), { user =>
          Ok(
            Json.toJson(
              Tsrlstaterule.update(user,
                Tsrlstaterule(
                       id,

                ))))
        })
      }
    )
  }

  def create = PostAction { implicit request =>
    Logger.debug("in tsrlstaterules.create")
    tsrlstateruleForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      tsrlstaterule => {
        Logger.debug(s"form: ${tsrlstaterule.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.NEW_TSRLSTATERULE), { user =>
          Ok(
            Json.toJson(
              Tsrlstaterule.create(user,
                Tsrlstaterule(

                ))))
        })
      }
    )
  }

  def delete(id: Int) = PostAction { implicit request =>
    Logger.debug("in tsrlstaterules.delete")
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.DELETE_TSRLSTATERULE), { user =>
      Tsrlstaterule.delete(user, id)
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

case class Tsrlstaterule(
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

object Tsrlstaterule {

  lazy val emptyTsrlstaterule = Tsrlstaterule(
)

  def apply(
      id: Int,
) = {

    new Tsrlstaterule(
      id,
)
  }

  def apply(
) = {

    new Tsrlstaterule(
)
  }

  private val tsrlstateruleParser: RowParser[Tsrlstaterule] = {
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
        Tsrlstaterule(
              id,
,
              createdAt,
              updatedAt,
              updatedBy)
    }
  }

  def create(user: CompanyUser, tsrlstaterule: Tsrlstaterule): Tsrlstaterule = {
    save(user, tsrlstaterule, true)
  }

  def update(user: CompanyUser, tsrlstaterule: Tsrlstaterule): Tsrlstaterule = {
    save(user, tsrlstaterule, false)
  }

  private def save(user: CompanyUser, tsrlstaterule: Tsrlstaterule, isNew: Boolean): Tsrlstaterule = {
    def getFields = {
      List(

      )
    }
    def throwException = {
      throw new RuntimeException(s"Error when saving ${C.TSRLSTATERULE}")
    }

    DBHelper.saveEx(
      user,
      Register(
        C.TSRLSTATERULE,
        C.ID,
        tsrlstaterule.id,
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

  def load(user: CompanyUser, id: Int): Option[Tsrlstaterule] = {
    loadWhere(user, s"${C.ID} = {id}", 'id -> id)
  }

  def loadWhere(user: CompanyUser, where: String, args : scala.Tuple2[scala.Any, anorm.ParameterValue[_]]*) = {
    DB.withConnection(user.database.database) { implicit connection =>
      SQL(s"SELECT t1.*, t2.${C.FK_NAME} FROM ${C.TSRLSTATERULE} t1 INNER JOIN ${C.???} t2 ON t1.${C.FK_ID} = t2.${C.FK_ID} WHERE $where")
        .on(args: _*)
        .as(tsrlstateruleParser.singleOpt)
    }
  }

  def delete(user: CompanyUser, id: Int) = {
    DB.withConnection(user.database.database) { implicit connection =>
      try {
        SQL(s"DELETE FROM ${C.TSRLSTATERULE} WHERE ${C.ID} = {id}")
        .on('id -> id)
        .executeUpdate
      } catch {
        case NonFatal(e) => {
          Logger.error(s"can't delete a ${C.TSRLSTATERULE}. ${C.ID} id: $id. Error ${e.toString}")
          throw e
        }
      }
    }
  }

  def get(user: CompanyUser, id: Int): Tsrlstaterule = {
    load(user, id) match {
      case Some(p) => p
      case None => emptyTsrlstaterule
    }
  }
}


// Router

GET     /api/v1/general/tsrlstaterule/:id              controllers.logged.modules.general.Tsrlstaterules.get(id: Int)
POST    /api/v1/general/tsrlstaterule                  controllers.logged.modules.general.Tsrlstaterules.create
PUT     /api/v1/general/tsrlstaterule/:id              controllers.logged.modules.general.Tsrlstaterules.update(id: Int)
DELETE  /api/v1/general/tsrlstaterule/:id              controllers.logged.modules.general.Tsrlstaterules.delete(id: Int)




/**/
