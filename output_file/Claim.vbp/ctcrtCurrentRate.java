
import java.util.Date;

public class ctcrtCurrentRate {

  //--------------------------------------------------------------------------
  // Procedure:   ctcrtCurrentRate
  // Description: Provides properties and methods to support the current_rate_t table values
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
  //                    current rate_MGR_PRV_CD  CurrentRateMgrPrvCd    m_strCurrentRateMgrPrvCd  mcstrCurrentRateMgrPrvCd
  //                    current rate_SVS_IND     CurrentRateSvsInd      m_bCurrentRateSvsInd      mcstrCurrentRateSvsInd
  //
  //
  //                 NOTE also that navigation should be done via the **ADO Wrapper's**
  //                 navigation methods instead of directly referencing the navigation
  //                 methods on a ADODB.Recordset object!
  //
  // Revisions:   1.0 ECG 01/02/2002 Initial creation.
  //              2.0 ECG 01/29/2002 Added Error handling used in Betsy's Prototype.
  //              3.0 BAW 05/02/2002 Added additional comments including instructions on
  //                                 where this class should be changed to work for a
  //                                 different table.
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
  //   Public      Property Get CurrentRateCd() As String
  //   Public      Property Let CurrentRateCd(ByVal strValue As String)
  //   Public      Property Get CurrentRateMgr() As String
  //   Public      Property Let CurrentRateMgr(ByVal strValue As String)
  //   Public      Property Get CurrentRateMgrPrvCd() As String
  //   Public      Property Let CurrentRateMgrPrvCd(ByVal strValue As String)
  //   Public      Property Get CurrentRateNm() As String
  //   Public      Property Let CurrentRateNm(ByVal strValue As String)
  //   Public      Property Get CurrentRateSvsInd() As Boolean
  //   Public      Property Let CurrentRateSvsInd(ByVal bValue As Boolean)
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
  //   Public      Property Get MktvalCurrentRateCd() As String
  //   Public      Property Let MktvalCurrentRateCd(ByVal strValue As String)
  //   Public      Property Get ShouldForceToUppercase(ByVal strTagIn As String) As Boolean
  //   Public      AddRecord() as Boolean
  //   Public      CheckForAnotherUsersChanges(ByVal lngWhatOperation As enumWhatOperationIsBeingAttempted, _
  //                   ByRef strACF2 As String) As Long
  //   Public      DeleteRecord() As Boolean
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
  private static final String MCSTRNAME = "ctcrtCurrentRate.";

  //...............................................................................................
  //!CUSTOMIZE!
  // These are the private variables corresponding to PUBLIC properties.
  // There should be one (of type udtColumn) for each column in the table that this class accesses.
  //...............................................................................................
  private udtColumn m_strCurrentRateCd;
  private udtColumn m_strCurrentRateNm;
  private udtColumn m_dteLstUpdDtm;
  private udtColumn m_strLstUpdUserId;
  private udtColumn m_bCurrentRateSVSInd;
  private udtColumn m_strMktValCurrentRateCd;
  private udtColumn m_strCurrentRateMgrPrvCd;
  private udtColumn m_strCurrentRateMgrCurrentRateCd;

  //...............................................................................................
  //!CUSTOMIZE!
  // Create one Const for each column in the table, defining the table column to which it refers.
  //...............................................................................................
  private static final String MCSTRCURRENTRATECD = "current rate_CD";
  private static final String MCSTRCURRENTRATENM = "current rate_NM";
  private static final String MCSTRLSTUPDDTM = "LST_UPD_DTM";
  private static final String MCSTRLSTUPDUSERID = "LST_UPD_USER_ID";
  private static final String MCSTRCURRENTRATESVSIND = "current rate_SVS_IND";
  private static final String MCSTRMKTVALCURRENTRATECD = "MKTVAL_current rate_CD";
  private static final String MCSTRCURRENTRATEMGRPRVCD = "current rate_MGR_PRV_CD";
  private static final String MCSTRCURRENTRATEMGRFUNDCD = "current rate_MGR_current rate_CD";


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
        ctclmClaim.getSingleRecord(m_rstLookup.Fields(MCSTRCURRENTRATECD).value);
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
  public String getCurrentRateCd() {
    String _rtn = "";
    // **************************************************************************
    // Function  : Property Get CurrentRateCd
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : string
    // **************************************************************************
    "Property Get CurrentRateCd"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = CStr(m_strCurrentRateCd.value);
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
  public void setCurrentRateCd(String strValue) {
    String _rtn = null;
    // **************************************************************************
    // Function  : Property Let CurrentRateCd
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal NewValue As string
    // Returns   :
    // **************************************************************************
    "Property Let CurrentRateCd"
.equals(Const cstrCurrentProc As String);
    try {

      m_strCurrentRateCd.value = strValue;
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
  public String getCurrentRateMgr() {
    String _rtn = "";
    // **************************************************************************
    // Function  : Property Get CurrentRateMgr
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : Int
    // **************************************************************************
    "Property Get CurrentRateMgr"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = CStr(m_strCurrentRateMgrCurrentRateCd.value);
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
  public void setCurrentRateMgr(String strValue) {
    String _rtn = null;
    // **************************************************************************
    // Function  : Property Let CurrentRateMgr
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal NewValue As Integer
    // Returns   :
    // **************************************************************************
    "Property Let CurrentRateMgr"
.equals(Const cstrCurrentProc As String);
    try {

      m_strCurrentRateMgrCurrentRateCd.value = strValue;
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
  public String getCurrentRateMgrPrvCd() {
    String _rtn = "";
    // **************************************************************************
    // Function  : Property Get CurrentRateMgrPrvCd
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : string
    // **************************************************************************
    "Property Get CurrentRateMgrPrvCd"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = CStr(m_strCurrentRateMgrPrvCd.value);
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
  public void setCurrentRateMgrPrvCd(String strValue) {
    String _rtn = null;
    // **************************************************************************
    // Function  : Property Let CurrentRateMgrPrvCd
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal NewValue As Integer
    // Returns   :
    // **************************************************************************
    "Property Let CurrentRateMgrPrvCd"
.equals(Const cstrCurrentProc As String);
    try {

      m_strCurrentRateMgrPrvCd.value = strValue;
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
  public String getCurrentRateNm() {
    String _rtn = "";
    // **************************************************************************
    // Function  : Property Get CurrentRateNm
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : Int
    // **************************************************************************
    "Property Get CurrentRateNm"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = CStr(m_strCurrentRateNm.value);
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
  public void setCurrentRateNm(String strValue) {
    String _rtn = null;
    // **************************************************************************
    // Function  : Property Let CurrentRateNm
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal NewValue As Integer
    // Returns   :
    // **************************************************************************
    "Property Let CurrentRateNm"
.equals(Const cstrCurrentProc As String);
    try {

      m_strCurrentRateNm.value = strValue;
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
  public boolean getCurrentRateSvsInd() {
    boolean _rtn = false;
    // **************************************************************************
    // Function  : Property Get CurrentRateSvsInd
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : Int
    // **************************************************************************
    "Property Get CurrentRateSvsInd"
.equals(Const cstrCurrentProc As String);
    try {

      // Note that this field is stored as a Y or N.

      _rtn = CBool(m_bCurrentRateSVSInd.value);
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
  public void setCurrentRateSvsInd(boolean bValue) {
    boolean _rtn = null;
    // **************************************************************************
    // Function  : Property Let CurrentRateSvsInd
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal NewValue As Integer
    // Returns   :
    // **************************************************************************
    "Property Let CurrentRateSvsInd"
.equals(Const cstrCurrentProc As String);
    try {

      // Note that this field is stored as a Y or N.

      m_bCurrentRateSVSInd.value = bValue;
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
    // Returns   : Int
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
    // Parameters: ByVal NewValue As Integer
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
  public String getMktValCurrentRateCd() {
    String _rtn = "";
    // **************************************************************************
    // Function  : Property Get MktvalCurrentRateCd
    // Purpose   : Retrieves current value from property
    // Parameters: N/A
    // Returns   : Int
    // **************************************************************************
    "Property Get MktvalCurrentRateCd"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = CStr(m_strMktValCurrentRateCd.value);
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
  public void setMktValCurrentRateCd(String strValue) {
    String _rtn = null;
    // **************************************************************************
    // Function  : Property Let MktvalCurrentRateCd
    // Purpose   : Assigns new Value to property
    // Parameters: ByVal NewValue As Integer
    // Returns   :
    // **************************************************************************
    "Property Let MktvalCurrentRateCd"
.equals(Const cstrCurrentProc As String);
    try {

      m_strMktValCurrentRateCd.value = strValue;
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
    "dbo.proc_current rate_insert"
.equals(Const cstrSproc As String);
    DBRecordSet rstTemp = null;
    boolean bSuccessful = false;
    ADODB.Parameter prmReturnValue = null;
    ADODB.Parameter prmCurrentRateCd = null;
    ADODB.Parameter prmCurrentRateNm = null;
    ADODB.Parameter prmCurrentRateSvsInd = null;
    ADODB.Parameter prmMktvalCurrentRateCd = null;
    ADODB.Parameter prmCurrentRateMgrPrvCd = null;
    ADODB.Parameter prmCurrentRateMgr = null;
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
      // Define the current rate_CD parameter
      prmCurrentRateCd = w_aDOCommand.CreateParameter(Name:="current rate_cd", Type:=adChar, Direction:=adParamInput, Size:=8, .value:=fnNullIfZLS(varIn:=CurrentRateCd, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmCurrentRateCd);

      // ---Parameter #3---
      // Define the current rate_NM parameter
      prmCurrentRateNm = w_aDOCommand.CreateParameter(Name:="current rate_nm", Type:=adVarChar, Direction:=adParamInput, Size:=50, .value:=fnNullIfZLS(varIn:=CurrentRateNm, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmCurrentRateNm);

      // ---Parameter #4---
      // Define the current rate_SVS_IND parameter
      prmCurrentRateSvsInd = w_aDOCommand.CreateParameter(Name:="current rate_svs_ind", Type:=adChar, Direction:=adParamInput, Size:=1, .value:=fnBoolToYN(getCurrentRateSvsInd()));
      w_aDOCommand.Parameters.Append(prmCurrentRateSvsInd);

      // ---Parameter #5---
      // Define the MKTVAL_current rate_CD parameter
      prmMktvalCurrentRateCd = w_aDOCommand.CreateParameter(Name:="mktval_current rate_cd", Type:=adChar, Direction:=adParamInput, Size:=8, .value:=fnNullIfZLS(varIn:=MktValCurrentRateCd, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmMktvalCurrentRateCd);

      // ---Parameter #6---
      // Define the current rate_MGR_PRV_CD parameter
      prmCurrentRateMgrPrvCd = w_aDOCommand.CreateParameter(Name:="current rate_mgr_prv_cd", Type:=adChar, Direction:=adParamInput, Size:=8, .value:=fnNullIfZLS(varIn:=CurrentRateMgrPrvCd, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmCurrentRateMgrPrvCd);

      // ---Parameter #7---
      // Define the current rate_MGR_current rate_CD parameter
      prmCurrentRateMgr = w_aDOCommand.CreateParameter(Name:="current rate_mgr_current rate_cd", Type:=adChar, Direction:=adParamInput, Size:=8, .value:=fnNullIfZLS(varIn:=CurrentRateMgr, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmCurrentRateMgr);

      // ---Parameter #8---
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
      bSuccessful = ctclmClaim.getRelativeRecord(getCurrentRateCd(), enumPositionDirection.ePDSAMERECORD);

      _rtn = true;
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    modGeneral.fnFreeRecordset(rstTemp);
    modGeneral.fnFreeObject(prmReturnValue);
    modGeneral.fnFreeObject(prmCurrentRateCd);
    modGeneral.fnFreeObject(prmCurrentRateNm);
    modGeneral.fnFreeObject(prmCurrentRateSvsInd);
    modGeneral.fnFreeObject(prmMktvalCurrentRateCd);
    modGeneral.fnFreeObject(prmCurrentRateMgrPrvCd);
    modGeneral.fnFreeObject(prmCurrentRateMgr);
    modGeneral.fnFreeObject(prmInvalid_Key);

    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (prmReturnValue) {
      //' 4028
      case  modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO  :
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO, MCSTRNAME+ cstrCurrentProc, "add");
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
      //' 4031
        break;

      case  modResConstants.gCRES_NERR_ADD_WITH_NONUNIQUE_KEY  :
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_ADD_WITH_NONUNIQUE_KEY, MCSTRNAME+ cstrCurrentProc, RTrim$(getCurrentRateCd()));
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
      //' 4032
        break;

      case  modResConstants.gCRES_NERR_KEY_NOT_FOUND:
        if ("MKTVAL_current rate_CD"
.equals(prmInvalid_Key.toUpperCase())) {
          modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_KEY_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, "Market Value current rate Cd", RTrim$(getMktValCurrentRateCd()), "current rate");
          /**TODO:** resume found: Resume(PROC_EXIT)*/;
        } 
        else {
          // current rate_MGR_PRV_CD
          modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_KEY_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, "current rate Mgr", RTrim$(getCurrentRateMgrPrvCd()), "Provider");
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
      //!CUSTOMIZE! fnSelectRecord call should pass the key column(s)
      rstSingleRecord_Fresh = fnSelectRecord(getCurrentRateCd());

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
    "dbo.proc_current rate_delete"
.equals(Const cstrSproc As String);
    boolean bSuccessful = false;
    DBRecordSet rstTemp = null;
    ADODB.Parameter prmReturnValue = null;
    ADODB.Parameter prmCurrentRateCd = null;
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
      // Define the current rate_CD parameter
      prmCurrentRateCd = w_aDOCommand.CreateParameter(Name:="current rate_cd", Type:=adChar, Direction:=adParamInput, Size:=8, .value:=fnNullIfZLS(varIn:=CurrentRateCd, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmCurrentRateCd);

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
      bSuccessful = ctclmClaim.getRelativeRecord(getCurrentRateCd(), enumPositionDirection.ePDPREVIOUSRECORD);

      _rtn = true;
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    modGeneral.fnFreeRecordset(rstTemp);
    modGeneral.fnFreeObject(prmReturnValue);
    modGeneral.fnFreeObject(prmCurrentRateCd);
    modGeneral.fnFreeObject(prmDependent_Table);

    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (prmReturnValue) {
      //' 4027
      case  modResConstants.gCRES_NERR_REC_NOT_FOUND      :
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_REC_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, "current rate Code "+ RTrim$(getCurrentRateCd()));
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
      //' 4028
        break;

      case  modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO  :
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO, MCSTRNAME+ cstrCurrentProc, "add");
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
      //' 4029
        break;

      case  modResConstants.gCRES_NERR_DEPENDENT_RECS_EXIST  :
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_DEPENDENT_RECS_EXIST, MCSTRNAME+ cstrCurrentProc, "current rate Cd ("+ RTrim$(getCurrentRateCd())+ ")", prmDependent_Table.toUpperCase());
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
    "dbo.proc_current rate_lu_select"
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
        //        Case gcRES_NERR_REC_NOT_FOUND       ' 4027
        //            ' Wipe out any trace of this error, but return False so the caller
        //            ' knows to go into Add mode if desired. NOTE: The caller can also
        //            ' identify this by looking at the LookupRecordCount public property.
        //            Err.Clear
        //            GetLookupData = False
        //            Resume PROC_EXIT
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
    // Params:
    //     strKey1              (in) = CurrentRateCd value from which to do the relative
    //                                 repositioning
    //     lngPositionDirection (in) = Indicates to which relative record the
    //                                 recordset should be positioned (relative
    //                                 to the strKey1 parameter value).
    //
    //
    // Called By:   cmdDelete_Click( ) of frmcurrent rate.frm
    //              cmdUpdate_Click( ) of frmcurrent rate.frm
    //              cmdNavigate_Click( ) of frmcurrent rate.frm
    //              fnAddRecord( ) of frmcurrent rate.frm
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
    String strKey1ForNewRec = "";

    try {

      //...........................................................................
      // Refresh the lookup data (m_rstLookupData) so other's changes
      // --and our own-- are now reflected in it. This resets the Lookup data,
      // record count, and current record number, and leaves the Lookup recordset
      // positioned to the first record (if there are records) or BOF (if there are
      // no records).
      //...........................................................................
      ctclmClaim.getLookupData();

      switch (lngPositionDirection) {
        case  enumPositionDirection.ePDPREVIOUSRECORD:
          // Make visible only those rows with keys prior to the specified key
          m_rstLookup.Filter = "current rate_cd < '"+ strKey1+ "'";
          if (m_rstLookup.RecordCount != 0) {
            // Grab the last record. The one with the highest key less than the
            // specified key is the one we want.
            m_adwADO.moveLast(m_rstLookup);
            //!CUSTOMIZE! The constant referenced below should refer to the key field.
            strKey1ForNewRec = m_rstLookup.Fields(MCSTRCURRENTRATECD).value;
            // Okay, we got it. Now make all records in the Lookup recordset
            // visible again, and then reposition to the new record.
            m_rstLookup.Filter = adFilterNone;
            m_adwADO.moveFirst(m_rstLookup);
            m_rstLookup.Find("current rate_cd = '"+ strKey1ForNewRec+ "'");
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
          m_rstLookup.Filter = "current rate_cd > '"+ strKey1+ "'";
          if (m_rstLookup.RecordCount != 0) {
            // Grab the first record. The one with the lowest key higher than the
            // specified key is the one we want.
            m_adwADO.moveFirst(m_rstLookup);
            //!CUSTOMIZE! The constant referenced below should refer to the key field.
            strKey1ForNewRec = m_rstLookup.Fields(MCSTRCURRENTRATECD).value;
            // Okay, we got it. Now make all records in the Lookup recordset
            // visible again, and then reposition to the new record.
            m_rstLookup.Filter = adFilterNone;
            m_adwADO.moveFirst(m_rstLookup);
            m_rstLookup.Find("current rate_cd = '"+ strKey1ForNewRec+ "'");
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
          m_rstLookup.Find("current rate_cd = '"+ strKey1+ "'");
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
        ctclmClaim.getSingleRecord(m_rstLookup.Fields(MCSTRCURRENTRATECD).value);
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
  public boolean getSingleRecord(String strKey1, boolean bSynchLookupRST) {
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
    //    strKey1         (in) = represents the primary key for the table (current rate_cd)
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
      rstTemp = fnSelectRecord(strKey1);

      //*TODO:** can't found type for with block
      //*With this
      __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = this;
      // For Char/VarChar fields,
      //     * Use fnZLSIfNull to ensure Nulls are appropriately translated.
      // For Numeric fields,
      //     * Use fnZeroIfNull to ensure Nulls are appropriately translated.
      // For Boolean fields,
      //     * Use fnYNToBool to ensure True/False is appropriately translated.
      w___TYPE_NOT_FOUND.CurrentRateCd = modDataConversion.fnZLSIfNull(rstTemp!current_rate_cd);
      w___TYPE_NOT_FOUND.CurrentRateNm = modDataConversion.fnZLSIfNull(rstTemp!current_rate_nm);
      w___TYPE_NOT_FOUND.LstUpdDtm = modDataConversion.fnZLSIfNull(rstTemp!lst_upd_dtm);
      w___TYPE_NOT_FOUND.LstUpdUserId = modDataConversion.fnZLSIfNull(rstTemp!lst_upd_user_id);
      w___TYPE_NOT_FOUND.CurrentRateSvsInd = modGeneral.fnYNToBool(rstTemp!current_rate_svs_ind);
      w___TYPE_NOT_FOUND.MktValCurrentRateCd = modDataConversion.fnZLSIfNull(rstTemp!mktval_current_rate_cd);
      w___TYPE_NOT_FOUND.CurrentRateMgrPrvCd = modDataConversion.fnZLSIfNull(rstTemp!current_rate_mgr_prv_cd);
      w___TYPE_NOT_FOUND.CurrentRateMgr = modDataConversion.fnZLSIfNull(rstTemp!current_rate_mgr_current_rate_cd);

      // Save original Last Updated info, to be used during UpdateRecord( ) and DeleteRecord( )
      // to determine if another user updated the record since it was retrieved.
      m_dteLstUpdDtm_Original = modDataConversion.fnZLSIfNull(rstTemp!lst_upd_dtm);
      m_strLstUpdUserId_Original = modDataConversion.fnZLSIfNull(rstTemp!lst_upd_user_id);

      if (bSynchLookupRST) {
        m_adwADO.moveFirst(m_rstLookup);
        m_rstLookup.Find("current rate_cd = '"+ getCurrentRateCd()+ "'");
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
        ctclmClaim.getSingleRecord(m_rstLookup.Fields(MCSTRCURRENTRATECD).value);
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
        ctclmClaim.getSingleRecord(m_rstLookup.Fields(MCSTRCURRENTRATECD).value);
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
          ctclmClaim.getRelativeRecord(getCurrentRateCd(), enumPositionDirection.ePDNEXTRECORD);
        } 
        else {
          // Get the requested record and reposition the Lookup recordset to that record
          //!CUSTOMIZE! The constant referenced below should refer to the key field.
          ctclmClaim.getSingleRecord(m_rstLookup.Fields(MCSTRCURRENTRATECD).value);
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
          ctclmClaim.getRelativeRecord(getCurrentRateCd(), enumPositionDirection.ePDPREVIOUSRECORD);
        } 
        else {
          // Get the requested record and reposition the Lookup recordset to that record
          //!CUSTOMIZE! The constant referenced below should refer to the key field.
          ctclmClaim.getSingleRecord(m_rstLookup.Fields(MCSTRCURRENTRATECD).value);
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
  public boolean haveDependents(String strKey1, String strDependentTable) { // TODO: Use of ByRef founded Public Function HaveDependents(ByVal strKey1 As String, ByRef strDependentTable As String) As Boolean
    boolean _rtn = false;
    // Comments  : Determines whether the current record can be deleted without
    //             hitting a referential integrity violation due to either:
    //             a. row(s) existing in other tables that use the current key value
    //                as a foreign key
    //             b. (for current_rate_t table only, I think) row(s) existing in the same table
    //                which has a circular reference to to the current key value.
    //             The calling form should look at the return value. If True, then
    //             the form's Delete button should be disabled.
    //
    // Parameters:
    //   strKey1
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
    "dbo.proc_current rate_verify_dependents"
.equals(Const cstrSproc As String);
    cadwADOWrapper adwTemp = null;
    DBRecordSet rstTemp = null;
    ADODB.Parameter prmReturnValue = null;
    ADODB.Parameter prmCurrentRateCd = null;
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
      prmCurrentRateCd = w_aDOCommand.CreateParameter(Name:="@current rate_cd", Type:=adChar, Direction:=adParamInput, Size:=8, .value:=strKey1);
      w_aDOCommand.Parameters.Append(prmCurrentRateCd);

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
    modGeneral.fnFreeObject(prmCurrentRateCd);
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
    "dbo.proc_current rate_update"
.equals(Const cstrSproc As String);
    DBRecordSet rstTemp = null;
    boolean bSuccessful = false;
    ADODB.Parameter prmReturnValue = null;
    ADODB.Parameter prmCurrentRateCd = null;
    ADODB.Parameter prmCurrentRateNm = null;
    ADODB.Parameter prmCurrentRateSvsInd = null;
    ADODB.Parameter prmMktvalCurrentRateCd = null;
    ADODB.Parameter prmCurrentRateMgrPrvCd = null;
    ADODB.Parameter prmCurrentRateMgr = null;
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
      // Define the current rate_CD parameter
      prmCurrentRateCd = w_aDOCommand.CreateParameter(Name:="current rate_cd", Type:=adChar, Direction:=adParamInput, Size:=8, .value:=fnNullIfZLS(varIn:=CurrentRateCd, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmCurrentRateCd);

      // ---Parameter #3---
      // Define the current rate_NM parameter
      prmCurrentRateNm = w_aDOCommand.CreateParameter(Name:="current rate_nm", Type:=adVarChar, Direction:=adParamInput, Size:=50, .value:=fnNullIfZLS(varIn:=CurrentRateNm, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmCurrentRateNm);

      // ---Parameter #4---
      // Define the current rate_SVS_IND parameter
      prmCurrentRateSvsInd = w_aDOCommand.CreateParameter(Name:="current rate_svs_ind", Type:=adChar, Direction:=adParamInput, Size:=1, .value:=fnBoolToYN(getCurrentRateSvsInd()));
      w_aDOCommand.Parameters.Append(prmCurrentRateSvsInd);

      // ---Parameter #5---
      // Define the MKTVAL_current rate_CD parameter
      prmMktvalCurrentRateCd = w_aDOCommand.CreateParameter(Name:="mktval_current rate_cd", Type:=adChar, Direction:=adParamInput, Size:=8, .value:=fnNullIfZLS(varIn:=MktValCurrentRateCd, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmMktvalCurrentRateCd);

      // ---Parameter #6---
      // Define the current rate_MGR_PRV_CD parameter
      prmCurrentRateMgrPrvCd = w_aDOCommand.CreateParameter(Name:="current rate_mgr_prv_cd", Type:=adChar, Direction:=adParamInput, Size:=8, .value:=fnNullIfZLS(varIn:=CurrentRateMgrPrvCd, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmCurrentRateMgrPrvCd);

      // ---Parameter #7---
      // Define the current rate_MGR_current rate_CD parameter
      prmCurrentRateMgr = w_aDOCommand.CreateParameter(Name:="current rate_mgr_current rate_cd", Type:=adChar, Direction:=adParamInput, Size:=8, .value:=fnNullIfZLS(varIn:=CurrentRateMgr, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmCurrentRateMgr);

      // ---Parameter #8---
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
      bSuccessful = ctclmClaim.getRelativeRecord(getCurrentRateCd(), enumPositionDirection.ePDSAMERECORD);

      _rtn = true;
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    modGeneral.fnFreeRecordset(rstTemp);
    modGeneral.fnFreeObject(prmReturnValue);
    modGeneral.fnFreeObject(prmCurrentRateCd);
    modGeneral.fnFreeObject(prmCurrentRateNm);
    modGeneral.fnFreeObject(prmCurrentRateSvsInd);
    modGeneral.fnFreeObject(prmMktvalCurrentRateCd);
    modGeneral.fnFreeObject(prmCurrentRateMgrPrvCd);
    modGeneral.fnFreeObject(prmCurrentRateMgr);
    modGeneral.fnFreeObject(prmInvalid_Key);

    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (prmReturnValue) {
      //' 4027
      case  modResConstants.gCRES_NERR_REC_NOT_FOUND      :
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_REC_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, "current rate Code "+ RTrim$(getCurrentRateCd()));
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
      //' 4028
        break;

      case  modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO  :
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO, MCSTRNAME+ cstrCurrentProc, "update");
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
      //' 4032
        break;

      case  modResConstants.gCRES_NERR_KEY_NOT_FOUND:
        if ("MKTVAL_current rate_CD"
.equals(prmInvalid_Key.toUpperCase())) {
          modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_KEY_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, "Market Value current rate Cd", RTrim$(getMktValCurrentRateCd()), "current rate");
          /**TODO:** resume found: Resume(PROC_EXIT)*/;
        } 
        else {
          // current rate_MGR_PRV_CD
          modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_KEY_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, "current rate Mgr", RTrim$(getCurrentRateMgrPrvCd()), "Provider");
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
      w___TYPE_NOT_FOUND.CurrentRateCd = "";
      w___TYPE_NOT_FOUND.CurrentRateNm = "";
      w___TYPE_NOT_FOUND.LstUpdDtm = Now;
      w___TYPE_NOT_FOUND.LstUpdUserId = "";
      w___TYPE_NOT_FOUND.CurrentRateSvsInd = false;
      w___TYPE_NOT_FOUND.MktValCurrentRateCd = "";
      w___TYPE_NOT_FOUND.CurrentRateMgrPrvCd = "";
      w___TYPE_NOT_FOUND.CurrentRateMgr = "";

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
        case  "CurrentRateCd":
          _rtn = m_strCurrentRateCd;
          break;

        case  "CurrentRateNm":
          _rtn = m_strCurrentRateNm;
          break;

        case  "LSTUPDDTM":
          _rtn = m_dteLstUpdDtm;
          break;

        case  "LSTUPDUSERID":
          _rtn = m_strLstUpdUserId;
          break;

        case  "CurrentRateSvsInd":
          _rtn = m_bCurrentRateSVSInd;
          break;

        case  "MKTVALCurrentRateCd":
          _rtn = m_strMktValCurrentRateCd;
          break;

        case  "CurrentRateMgrPrvCd":
          _rtn = m_strCurrentRateMgrPrvCd;
          break;

        case  "CurrentRateMgr":
          _rtn = m_strCurrentRateMgrCurrentRateCd;
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
      m_adwADO.getMetaData_Columns("current_rate_t", rstMetaData);

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
          case  MCSTRCURRENTRATECD:
            fnGetColMetaData(m_strCurrentRateCd, rstMetaData);
            // Per the screen spec, the following overrides
            // the DBMS max width of 8 until such time
            // that the Trades import files can all handle
            // a full-width/8-character current rate code.
            m_strCurrentRateCd.maxCharacters = 4;
            // Force to be uppercase, per screen spec
            m_strCurrentRateCd.shouldForceToUppercase = true;
            break;

          case  MCSTRCURRENTRATENM:
            fnGetColMetaData(m_strCurrentRateNm, rstMetaData);
            break;

          case  MCSTRLSTUPDDTM:
            fnGetColMetaData(m_dteLstUpdDtm, rstMetaData);
            break;

          case  MCSTRLSTUPDUSERID:
            fnGetColMetaData(m_strLstUpdUserId, rstMetaData);
            break;

          case  MCSTRCURRENTRATESVSIND:
            fnGetColMetaData(m_bCurrentRateSVSInd, rstMetaData);
            break;

          case  MCSTRMKTVALCURRENTRATECD:
            fnGetColMetaData(m_strMktValCurrentRateCd, rstMetaData);
            break;

          case  MCSTRCURRENTRATEMGRPRVCD:
            fnGetColMetaData(m_strCurrentRateMgrPrvCd, rstMetaData);
            break;

          case  MCSTRCURRENTRATEMGRFUNDCD:
            fnGetColMetaData(m_strCurrentRateMgrCurrentRateCd, rstMetaData);
            // Force to be uppercase, per screen spec
            m_strCurrentRateMgrCurrentRateCd.shouldForceToUppercase = true;
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
      m_adwADO.getMetaData_PrimaryKeys("current_rate_t", rstMetaData);

      do Until .EOF        // The SELECT CASE should list all table columns
        // (though you could skip the LST_UPD_xxx columns if you change the Case Else,
        //  since these would never be a key)
        switch (rstMetaData("COLUMN_NAME").value.toUpperCase()) {
          case  MCSTRCURRENTRATECD:
            m_strCurrentRateCd.isKey = true;
            break;

          case  MCSTRCURRENTRATENM:
            m_strCurrentRateNm.isKey = true;
            break;

          case  MCSTRLSTUPDDTM:
            m_dteLstUpdDtm.isKey = true;
            break;

          case  MCSTRLSTUPDUSERID:
            m_strLstUpdUserId.isKey = true;
            break;

          case  MCSTRCURRENTRATESVSIND:
            m_bCurrentRateSVSInd.isKey = true;
            break;

          case  MCSTRMKTVALCURRENTRATECD:
            m_strMktValCurrentRateCd.isKey = true;
            break;

          case  MCSTRCURRENTRATEMGRPRVCD:
            m_strCurrentRateMgrPrvCd.isKey = true;
            break;

          case  MCSTRCURRENTRATEMGRFUNDCD:
            m_strCurrentRateMgrCurrentRateCd.isKey = true;
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
  private DBRecordSet fnSelectRecord(String strKey1) {
    //--------------------------------------------------------------------------
    // Procedure:   fnSelectRecord
    // Description: Selects a single record based on the value(s) in the
    //              properties that correspond to the table's key(s)
    //
    //              NOTE: For each table key, there should be a parameter
    //                    of the appropriate data type!
    //
    // Parameters:
    //     strKey1 (in) - the key to the table that should be retrieved
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
    "dbo.proc_current rate_select"
.equals(Const cstrSproc As String);
    DBRecordSet rstTemp = null;
    ADODB.Parameter prmReturnValue = null;
    ADODB.Parameter prmCurrentRateCd = null;

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
      // Define the current rate_CD parameter
      prmCurrentRateCd = w_aDOCommand.CreateParameter(Name:="current rate_cd", Type:=adChar, Direction:=adParamInput, Size:=8, .value:=fnNullIfZLS(varIn:=strKey1, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmCurrentRateCd);

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
    modGeneral.fnFreeObject(prmCurrentRateCd);

    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return null;
    // **TODO:** label found: PROC_ERR:;
    switch (prmReturnValue) {
      //' 4027
      case  modResConstants.gCRES_NERR_REC_NOT_FOUND      :
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_REC_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, "current rate Code "+ RTrim$(strKey1));
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


case class TcrtcurrentrateData(
              id: Option[Int],

              )

object Tcrtcurrentrates extends Controller with ProvidesUser {

  val tcrtcurrentrateForm = Form(
    mapping(
      "id" -> optional(number),

  )(TcrtcurrentrateData.apply)(TcrtcurrentrateData.unapply))

  implicit val tcrtcurrentrateWrites = new Writes[Tcrtcurrentrate] {
    def writes(tcrtcurrentrate: Tcrtcurrentrate) = Json.obj(
      "id" -> Json.toJson(tcrtcurrentrate.id),
      C.ID -> Json.toJson(tcrtcurrentrate.id),

    )
  }

  def get(id: Int) = GetAction { implicit request =>
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.LIST_TCRTCURRENTRATE), { user =>
      Ok(Json.toJson(Tcrtcurrentrate.get(user, id)))
    })
  }

  def update(id: Int) = PostAction { implicit request =>
    Logger.debug("in tcrtcurrentrates.update")
    tcrtcurrentrateForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      tcrtcurrentrate => {
        Logger.debug(s"form: ${tcrtcurrentrate.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.EDIT_TCRTCURRENTRATE), { user =>
          Ok(
            Json.toJson(
              Tcrtcurrentrate.update(user,
                Tcrtcurrentrate(
                       id,

                ))))
        })
      }
    )
  }

  def create = PostAction { implicit request =>
    Logger.debug("in tcrtcurrentrates.create")
    tcrtcurrentrateForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      tcrtcurrentrate => {
        Logger.debug(s"form: ${tcrtcurrentrate.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.NEW_TCRTCURRENTRATE), { user =>
          Ok(
            Json.toJson(
              Tcrtcurrentrate.create(user,
                Tcrtcurrentrate(

                ))))
        })
      }
    )
  }

  def delete(id: Int) = PostAction { implicit request =>
    Logger.debug("in tcrtcurrentrates.delete")
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.DELETE_TCRTCURRENTRATE), { user =>
      Tcrtcurrentrate.delete(user, id)
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

case class Tcrtcurrentrate(
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

object Tcrtcurrentrate {

  lazy val emptyTcrtcurrentrate = Tcrtcurrentrate(
)

  def apply(
      id: Int,
) = {

    new Tcrtcurrentrate(
      id,
)
  }

  def apply(
) = {

    new Tcrtcurrentrate(
)
  }

  private val tcrtcurrentrateParser: RowParser[Tcrtcurrentrate] = {
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
        Tcrtcurrentrate(
              id,
,
              createdAt,
              updatedAt,
              updatedBy)
    }
  }

  def create(user: CompanyUser, tcrtcurrentrate: Tcrtcurrentrate): Tcrtcurrentrate = {
    save(user, tcrtcurrentrate, true)
  }

  def update(user: CompanyUser, tcrtcurrentrate: Tcrtcurrentrate): Tcrtcurrentrate = {
    save(user, tcrtcurrentrate, false)
  }

  private def save(user: CompanyUser, tcrtcurrentrate: Tcrtcurrentrate, isNew: Boolean): Tcrtcurrentrate = {
    def getFields = {
      List(

      )
    }
    def throwException = {
      throw new RuntimeException(s"Error when saving ${C.TCRTCURRENTRATE}")
    }

    DBHelper.saveEx(
      user,
      Register(
        C.TCRTCURRENTRATE,
        C.ID,
        tcrtcurrentrate.id,
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

  def load(user: CompanyUser, id: Int): Option[Tcrtcurrentrate] = {
    loadWhere(user, s"${C.ID} = {id}", 'id -> id)
  }

  def loadWhere(user: CompanyUser, where: String, args : scala.Tuple2[scala.Any, anorm.ParameterValue[_]]*) = {
    DB.withConnection(user.database.database) { implicit connection =>
      SQL(s"SELECT t1.*, t2.${C.FK_NAME} FROM ${C.TCRTCURRENTRATE} t1 INNER JOIN ${C.???} t2 ON t1.${C.FK_ID} = t2.${C.FK_ID} WHERE $where")
        .on(args: _*)
        .as(tcrtcurrentrateParser.singleOpt)
    }
  }

  def delete(user: CompanyUser, id: Int) = {
    DB.withConnection(user.database.database) { implicit connection =>
      try {
        SQL(s"DELETE FROM ${C.TCRTCURRENTRATE} WHERE ${C.ID} = {id}")
        .on('id -> id)
        .executeUpdate
      } catch {
        case NonFatal(e) => {
          Logger.error(s"can't delete a ${C.TCRTCURRENTRATE}. ${C.ID} id: $id. Error ${e.toString}")
          throw e
        }
      }
    }
  }

  def get(user: CompanyUser, id: Int): Tcrtcurrentrate = {
    load(user, id) match {
      case Some(p) => p
      case None => emptyTcrtcurrentrate
    }
  }
}


// Router

GET     /api/v1/general/tcrtcurrentrate/:id              controllers.logged.modules.general.Tcrtcurrentrates.get(id: Int)
POST    /api/v1/general/tcrtcurrentrate                  controllers.logged.modules.general.Tcrtcurrentrates.create
PUT     /api/v1/general/tcrtcurrentrate/:id              controllers.logged.modules.general.Tcrtcurrentrates.update(id: Int)
DELETE  /api/v1/general/tcrtcurrentrate/:id              controllers.logged.modules.general.Tcrtcurrentrates.delete(id: Int)




/**/
