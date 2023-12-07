public class capsAppSettings {

  //
  // =-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-=
  // Class       : capsAppSettings
  // Description : Simplified registry access routines used
  //               for saving AppBranch settings
  // Source      : Total Visual SourceBook 2000
  //
  // Procedures  :
  //   Private     Class_Initialize()
  //   Private     Class_Terminate()
  //   Public      Property Get ActiveDatabase(strEnvName As String) As String
  //   Public      Property Get ActiveServer(strEnvName As String) As String
  //   Public      Property Get AppBranch() As String
  //   Public      Property Let AppBranch(ByVal strValue As String)
  //   Public      Property Get ArchiveDatabase(strEnvName As String) As String
  //   Public      Property Get ArchiveServer(strEnvName As String) As String
  //   Public      Property Get CompanyBranch() As String
  //   Public      Property Let CompanyBranch(ByVal strValue As String)
  //   Public      Property Get EnvironmentKey(strEnvName As String) As String
  //   Public      Property Get EnvironmentNames() As String()
  //   Public      Property Get LastLogOnUserID() As String
  //   Public      Property Let LastLogOnUserID(ByVal strValue as String)
  //   Public      Property Get LastLogOnUserPassword() As String
  //   Public      Property Let LastLogOnUserPassword(ByVal strValue as String)
  //   Public      Property Get MainBranch() As String
  //   Public      Property Let MainBranch(strValue As String)
  //   Public      Property Get Port(ByVal strEnvName As String) As String
  //   Public      Public Property Get TaxFileFolder() As String
  //   Public      Public Property Let TaxFileFolder(ByVal strValue As String)
  //   Public      fnLoadEnvironments()
  //   Private     fnShellSortAny(ByRef varArray As Variant, ByVal lngNbrOfElements As Long, _
  //                  Optional ByVal bSortDescending As Boolean = False)
  //   Public      LoadCbo_EnvironmentNames(ByRef cboIn As ComboBox)
  //   Public      ReadEntry(ByVal eHKRoot As EnumRegistryRootKeys, ByVal strRegKey As String, _
  //                         ByVal strEntry As String, Optional ByVal strDefault As String = "EMPTY") As String
  //   Public      RestoreForm(ByRef frmIn As Form, Optional ByVal bForceVisible As Boolean = False) As Boolean
  //   Public      SaveForm(ByRef frmIn As Form)
  //   Public      WriteEntry(ByVal eHKRoot As EnumRegistryRootKeys, ByVal strRegKey As String, _
  //                         ByVal strEntry As String, ByVal strValue As String)
  //
  // Modified:
  //
  //   Date     Who   What
  //   -------- ---   -------------------------------------------------------------------
  //   10/04/02 BAW   Cloned from TRS.
  //   06/22/03 BAW   Added Port and support for same, since the new Prod environment under
  //                  SQL Server 2000 utilizes nonstandard Port assignments and hence it must
  //                  be specified in the ADO .ConnectionString property.
  // =-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-=
//Option Explicit
  *Option Compare Binary
  private static final String MCSTRNAME = "capsAppSettings.";

  // Local variables to hold Public Property values
  private String m_strMainBranch = "";
  private String m_strCompanyBranch = "";
  private String m_strAppBranch = "";
  private String m_strVersionBranch = "";
  private String m_strLastLogOnUserID = "";
  private String m_strLastLogOnPassword = "";
  private String m_strLastLogOnEnvironment = "";
  private udtEnvironment[] m_audtEnvironments;
  private String m_strTaxFileFolder = "";


  private static final String MCSTRSLASH = "\";
  private static final String MCSTREMPTY = "EMPTY";

  // Sections in the registry
  private static final String MCSTRSCREENPREFERENCES = "ScreenPreferences";

  // Keys or Key Values in the registry
  private static final String MCSTRENVIRONMENTS = "Environments";
  private static final String MCSTRENV_ACTIVEDATABASE = "ActiveDB";
  private static final String MCSTRENV_ACTIVESERVER = "ActiveServer";
  private static final String MCSTRENV_ARCHIVEDATABASE = "ArchiveDB";
  private static final String MCSTRENV_ARCHIVESERVER = "ArchiveServer";
  private static final String MCSTRENV_ENVIRONMENTNAME = "EnvName";
  private static final String MCSTRLASTLOGONUSERID = "LastLogOnUserID";
  private static final String MCSTRENV_PORT = "Port";
  //'SQL_INTEGRATED_SECURITY
  private static final String MCSTRENV_USESWINDOWSAUTHENTICATION = "UsesWindowsAuthentication";
  private static final String MCSTRTAXFILEFOLDER = "TaxFileFolder";


//*TODO:** type is translated as a new class at the end of the file Private Type udtEnvironment




  ///\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
  //\                                                                  /
  //|         CLASS_INITIALIZE / CLASS_TERMINATE   Procedures         |
  ///                                                                  \
  //\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/

  //////////////////////////////////////////////////////////////////////////////////////////////////
  private void class_Initialize() {
    // Set initial values to defaults which may be overridden with property settings
    //
    // /\/\/\/\/\/\/\/\/\/\/\/\  WARNING /\/\/\/\/\/\/\/\/\/\/\/\
    // Use the fnConstructor_cerhErrorHandler( ) procedure in
    // modConstructors.bas to instantiate this object !!

    // See that procedure for additional comments as to why this is
    // necessary
    // /\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
    //
    "Class_Initialize"
.equals(Const cstrCurrentProc As String);
    try {

      setMainBranch("SOFTWARE");
      //*TODO:** can't found type for with block
      //*With App
      __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = App;
      setAppBranch(w___TYPE_NOT_FOUND.Title);
      setVersionBranch(CStr(w___TYPE_NOT_FOUND.Major+ modResConstants.gCSTRDOT+ w___TYPE_NOT_FOUND.Minor+ modResConstants.gCSTRDOT+ w___TYPE_NOT_FOUND.Revision));
      setCompanyBranch("Sun Life Financial");
      setLastLogonPassword("");

      // Remainder of the intialization is done by
      // fnConstructor_cerhErrorHandler( ) in modConstructors.bas.
//*TODO:** the error label 0: couldn't be found
  }
}


  private void class_Terminate() {
    // Free up resources allocated in this class
    "Class_Terminate"
.equals(Const cstrCurrentProc As String);
    try {

      Erase(m_audtEnvironments);
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




///\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
//\                                                                  /
//|                PROPERTY GET/LET    Procedures                    |
///                                                                  \
//\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/

//////////////////////////////////////////////////////////////////////////////////////////////////
  public String getActiveDatabase(String strEnvName) {
    String _rtn = "";
    // Returns the active database name associated with the specified Environment
    "Get ActiveDatabase"
.equals(Const cstrCurrentProc As String);
    int lngIndex = 0;
    try {

      for (lngIndex = LBound(m_audtEnvironments); lngIndex <= m_audtEnvironments.length; lngIndex++) {
        if (m_audtEnvironments[lngIndex].environmentName.equals(strEnvName)) {
          _rtn = m_audtEnvironments[lngIndex].activeDatabase;
        }
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
  public String getActiveServer(String strEnvName) {
    String _rtn = "";
    // Returns the active Server name associated with the specified Environment
    "Get ActiveServer"
.equals(Const cstrCurrentProc As String);
    int lngIndex = 0;
    try {

      for (lngIndex = LBound(m_audtEnvironments); lngIndex <= m_audtEnvironments.length; lngIndex++) {
        if (m_audtEnvironments[lngIndex].environmentName.equals(strEnvName)) {
          _rtn = m_audtEnvironments[lngIndex].activeServer;
        }
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
  public String getAppBranch() {
    String _rtn = "";
    // Returns the current value of the AppBranch property
    "Get AppBranch"
.equals(Const cstrCurrentProc As String);
    try {
      _rtn = m_strAppBranch;
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

  return _rtn;
}


//////////////////////////////////////////////////////////////////////////////////////////////////
  public void setAppBranch(String strValue) {
    String _rtn = null;
    // Sets the AppBranch property to the value specified by strValue
    "Let AppBranch"
.equals(Const cstrCurrentProc As String);
    try {

      if (MCSTRSLASH
.equals(strValue.substring(strValue.length() - 1))) {
        strValue = strValue.substring(0, strValue.length() - 1);
      }

      m_strAppBranch = strValue;
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
  public String getArchiveDatabase(String strEnvName) {
    String _rtn = "";
    // Returns the archive database name associated with the specified Environment
    "Get ArchiveDatabase"
.equals(Const cstrCurrentProc As String);
    int lngIndex = 0;
    try {

      for (lngIndex = LBound(m_audtEnvironments); lngIndex <= m_audtEnvironments.length; lngIndex++) {
        if (m_audtEnvironments[lngIndex].environmentName.equals(strEnvName)) {
          _rtn = m_audtEnvironments[lngIndex].archiveDatabase;
        }
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
  public String getArchiveServer(String strEnvName) {
    String _rtn = "";
    // Returns the archive Server name associated with the specified Environment
    "Get ArchiveServer"
.equals(Const cstrCurrentProc As String);
    int lngIndex = 0;
    try {

      for (lngIndex = LBound(m_audtEnvironments); lngIndex <= m_audtEnvironments.length; lngIndex++) {
        if (m_audtEnvironments[lngIndex].environmentName.equals(strEnvName)) {
          _rtn = m_audtEnvironments[lngIndex].archiveServer;
        }
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
  public String getCompanyBranch() {
    String _rtn = "";
    // Returns the current value of CompanyBranch
    "Get CompanyBranch"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = m_strCompanyBranch;
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
  public void setCompanyBranch(String strValue) {
    String _rtn = null;
    // Sets the CompanyBranch property to the value specified by strValue
    "Let CompanyBranch"
.equals(Const cstrCurrentProc As String);
    try {

      if (MCSTRSLASH
.equals(strValue.substring(strValue.length() - 1))) {
        strValue = strValue.substring(0, strValue.length() - 1);
      }

      m_strCompanyBranch = strValue;
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
  public String() getEnvironmentNames() {
    String() _rtn = null;
    // Comments:   Retrieve an array of database Environment Names
    // Parameters: N/A
    // Returns:    An array of strings that list the environment name for each
    //             environment defined in the registry
    // Called by : No one yet  :(
    "Get EnvironmentName"
.equals(Const cstrCurrentProc As String);
    int lngIndex = 0;

    try {

      G.redimPreserve(m_audtEnvironments.length,  );

      for (lngIndex = LBound(m_audtEnvironments); lngIndex <= (m_audtEnvironments.length); lngIndex++) {
        astrNames(lngIndex) = m_audtEnvironments[lngIndex].environmentName;
      }

      // Don't sort, since this will influence the order in which they appear...We want it sorted by
      // KEY not Name, e.g., "1_Prd2A_COLI" versus "COLI Development Phase 2A"
      //       Commented out:  fnShellSortAny astrNames, UBound(astrNames) - 1, False

      _rtn = astrNames;
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    Erase(astrNames);

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
  public String getLastLogonEnvironment() {
    String _rtn = "";
    // Comments:   This property is NOT persisted across sessions. It is saved merely to allow
    //             the status bar text to be updated in its entirety whenever the MDI Form
    //             is activated or the user logs in.
    // Parameters: N/A
    // Returns:    The current value of LastLogOnEnvironment.
    // Called by :
    "Get LastLogOnEnvironment"
.equals(Const cstrCurrentProc As String);

    try {

      _rtn = m_strLastLogOnEnvironment;
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
  public void setLastLogonEnvironment(String strValue) {
    String _rtn = null;
    // Comments:   Set the value of the LastLogOnEnvironment property (called after a successful logon)
    // Parameters: N/A
    // Returns:    The current value of LastLogOnEnvironment. Its default value is a null string ("").
    // Called by : cmdOK_Click() of frmLogOn
    "Let LastLogOnEnvironment"
.equals(Const cstrCurrentProc As String);

    try {

      m_strLastLogOnEnvironment = strValue;
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
  public void setLastLogonPassword(String strValue) {
    // Comments:   Set the value of the LastLogOnPassword property (called after a successful logon)
    // Parameters: N/A
    // Returns:    The current value of LastLogOnPassword. Its default value is a null string ("").
    // Called by : cmdOK_Click() of frmLogOn
    "Let LastLogOnPassword"
.equals(Const cstrCurrentProc As String);

    try {

      m_strLastLogOnPassword = strValue;
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
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public String getLastLogOnUserID() {
    String _rtn = "";
    // Comments:   Retrieve the ACF2 under which the current user last logged on
    // Parameters: N/A
    // Returns:    The current value of LastLogOnUserID. Its default value is a null string ("").
    // Called by : fnInit_gerhApp of modConstructors
    //             Form_Load() of frmLogOn
    "Get LastLogOnUserID"
.equals(Const cstrCurrentProc As String);
    String strRegKey = "";

    try {

      if (m_strLastLogOnUserID.equals("")) {
        strRegKey = m_strMainBranch+ MCSTRSLASH+ m_strCompanyBranch+ MCSTRSLASH+ m_strAppBranch;
        m_strLastLogOnUserID = readEntry(EnumRegistryRootKeys.rRKHKEY_CURRENT_USER, strRegKey, MCSTRLASTLOGONUSERID, "");
        _rtn = m_strLastLogOnUserID;
      } 
      else {
        _rtn = m_strLastLogOnUserID;
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
  public void setLastLogOnUserID(String strValue) {
    String _rtn = null;
    // Comments:   Set the value of the LastLogOnUserID property (called after a successful logon)
    // Parameters: N/A
    // Returns:    The current value of LastLogOnUserID. Its default value is a null string ("").
    // Called by : cmdOK_Click() of frmLogOn
    "Let LastLogOnUserID"
.equals(Const cstrCurrentProc As String);
    String strRegKey = "";

    try {

      strRegKey = m_strMainBranch+ MCSTRSLASH+ m_strCompanyBranch+ MCSTRSLASH+ m_strAppBranch;

      m_strLastLogOnUserID = strValue;
      writeEntry(EnumRegistryRootKeys.rRKHKEY_CURRENT_USER, strRegKey, MCSTRLASTLOGONUSERID, strValue);
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
  private void setMainBranch(String strValue) {
    // Sets the MainBranch property to the value specified by strValue
    "Let MainBranch"
.equals(Const cstrCurrentProc As String);
    try {

      if (MCSTRSLASH
.equals(strValue.substring(strValue.length() - 1))) {
        strValue = strValue.substring(0, strValue.length() - 1);
      }
      m_strMainBranch = strValue;
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
}


//////////////////////////////////////////////////////////////////////////////////////////////////
  public String getPort(String strEnvName) {
    String _rtn = "";
    // Returns the Port number associated with the specified Environment
    "Get Port"
.equals(Const cstrCurrentProc As String);
    int lngIndex = 0;
    try {

      for (lngIndex = LBound(m_audtEnvironments); lngIndex <= m_audtEnvironments.length; lngIndex++) {
        if (m_audtEnvironments[lngIndex].environmentName.equals(strEnvName)) {
          _rtn = m_audtEnvironments[lngIndex].port;
        }
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
  public String getTaxFileFolder() {
    String _rtn = "";
    // Comments:   Retrieve the ACF2 under which the current user last logged on
    // Parameters: N/A
    // Returns:    The current value of TaxFileFolder. Its default value is the user's
    //             My Documents folder.
    // Called by : fnInit_gerhApp of modConstructors
    //             Form_Load() of frmGenerateTaxFile
    "Get TaxFileFolder"
.equals(Const cstrCurrentProc As String);
    New fso = null; Scripting.FileSystemObject
    String strPerUserTaxFileFolder = "";
    String strRegKey = "";

    try {

      // If the TaxFile Folder is empty or merely contains a slash...
      if (m_strTaxFileFolder.equals("")) {
        strRegKey = m_strMainBranch+ MCSTRSLASH+ m_strCompanyBranch+ MCSTRSLASH+ m_strAppBranch;
        // 1. Is a Per User TaxFile folder defined?
        //    a. If so, use it.
        //    b. If not, then use the Per User Non-Roaming folder (the user's
        //       My Documents folder)
        strPerUserTaxFileFolder = readEntry(EnumRegistryRootKeys.rRKHKEY_CURRENT_USER, strRegKey, MCSTRTAXFILEFOLDER, "");
        if (!(strPerUserTaxFileFolder.equals(""))) {
          if (modGeneral.gbLogVerbose) {
            modAppLog.fnLogWrite("TaxFileFolder defined in HKCU is: "+ strPerUserTaxFileFolder, cstrCurrentProc);
          }
        } 
        else {
          // Get the path to where per user non-roaming data is stored. This path
          // will be created if it doesn't already exist.
          strPerUserTaxFileFolder = modWinApi.fnGetSpecialFolder(0, modWinApi.cSIDL_PERSONAL || CSIDL_FLAG_CREATE);
          if (modGeneral.gbLogVerbose) {
            modAppLog.fnLogWrite("TaxFileFolder defaulting to user's My Documents folder: "+ strPerUserTaxFileFolder, cstrCurrentProc);
          }
        }

        // If directory doesn't exist, create it
        if (!fso.FolderExists(strPerUserTaxFileFolder)) {
          // If an error occurs during the create, ignore for now. frmGenerateTaxFile's fnValidData procedure
          // will generate an error if the folder doesn't exist by the time a generate is kicked off.
      }
      //*TODO:** the error label PROC_ERR: couldn't be found
        try {
        fso.CreateFolder(strPerUserTaxFileFolder);
    }
    try {
    }

    // Using Property Let, save to Per User (HKCU). This will also add a trailing slash
    // if one doesn't already exist.
    Me.setTaxFileFolder(strPerUserTaxFileFolder);
  }

  _rtn = m_strTaxFileFolder;
  } catch (Exception ex) {
  //' disable error handler
  }
  try {

    // Clean-up statements go here
    modGeneral.fnFreeObject(fso);

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
  public void setTaxFileFolder(String strValue) {
    String _rtn = null;
    // Comments:   Set the value of the TaxFileFolder property
    // Parameters: N/A
    // Returns:    The current value of TaxFileFolder. Its default value is the user's My Documents
    //             folder (the location for Per User Non-Roaming data).
    // Called by : cmdOK_Click() of frmGenerateTaxFile
    "Let TaxFileFolder"
.equals(Const cstrCurrentProc As String);
    String strRegKey = "";

    try {

      strRegKey = m_strMainBranch+ MCSTRSLASH+ m_strCompanyBranch+ MCSTRSLASH+ m_strAppBranch;

      m_strTaxFileFolder = modGeneral.fnAddBackslash(strValue.trim()).toUpperCase();
      writeEntry(EnumRegistryRootKeys.rRKHKEY_CURRENT_USER, strRegKey, MCSTRTAXFILEFOLDER, strValue);
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



//SQL_INTEGRATED_SECURITY - Added
//////////////////////////////////////////////////////////////////////////////////////////////////
  public boolean getUsesWindowsAuthentication(String strEnvName) {
    boolean _rtn = false;
    // Returns True if the SQL Server of the specified Environment used windows authentication
    // rather than SQL Server authentication; False otherwise.
    "Get UsesWindowsAuthentication"
.equals(Const cstrCurrentProc As String);
    int lngIndex = 0;
    try {

      for (lngIndex = LBound(m_audtEnvironments); lngIndex <= m_audtEnvironments.length; lngIndex++) {
        if (m_audtEnvironments[lngIndex].environmentName.equals(strEnvName)) {
          _rtn = m_audtEnvironments[lngIndex].usesWindowsAuthentication;
        }
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
//SQL_INTEGRATED_SECURITY - Added


//////////////////////////////////////////////////////////////////////////////////////////////////
  public String getVersionBranch() {
    String _rtn = "";
    // Returns the current value of the VersionBranch property
    "Get VersionBranch"
.equals(Const cstrCurrentProc As String);
    try {
      _rtn = m_strVersionBranch;
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
  public void setVersionBranch(String strValue) {
    String _rtn = null;
    // Sets the VersionBranch property to the value specified by strValue
    "Let VersionBranch"
.equals(Const cstrCurrentProc As String);
    try {

      m_strVersionBranch = strValue;
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



///\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
//\                                                                  /
//|                        PUBLIC  Procedures                        |
///                                                                  \
//\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/


//SQL_INTEGRATED_SECURITY
//////////////////////////////////////////////////////////////////////////////////////////////////
  public void loadCbo_EnvironmentNames(ComboBox cboIn) { // TODO: Use of ByRef founded Public Sub LoadCbo_EnvironmentNames(ByRef cboIn As ComboBox)
    // Comments:   Populates the combobox parameter with a list of Environment Names
    // Parameters: cboIn   (in)   combobox to populate
    // Returns:    N/A
    // Called by : Form_Load() of frmLogOn
    "LoadCbo_EnvironmentNames"
.equals(Const cstrCurrentProc As String);
    int lngIndex = 0;

    try {

      for (lngIndex = LBound(m_audtEnvironments); lngIndex <= m_audtEnvironments.length; lngIndex++) {
        cboIn.AddItem(m_audtEnvironments[lngIndex].environmentName);
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
//SQL_INTEGRATED_SECURITY


//////////////////////////////////////////////////////////////////////////////////////////////////
  public String readEntry(EnumRegistryRootKeys eHKRoot, String strRegKey, String strEntry, String strDefault) {
    String _rtn = "";
    // Comments  : Reads a string value from the location in the
    //             registry specified by the class properties and/or parameters
    // Parameters: eHKRoot    - (input) the hive from which to read a value, e.g., HKLM, HKCU, etc.
    //             strRegKey  - (input) the path under eHKRoot under which to find the key containing strEntry
    //             strEntry   - (input) The value under EHKRoot\strRegKey to retrieve
    //             strDefault - (input) The default value to return if strEntry isn't found
    // Returns   : Either the registry value, or the default value
    // Source    : Total Visual SourceBook 2000
    //
    "ReadEntry"
.equals(Const cstrCurrentProc As String);
    String strValue = "";

    try {

      strValue = modRegistry.registryGetKeyValue(eHKRoot, strRegKey, strEntry);

      if (strValue.equals("")) {
        _rtn = strDefault;
      } 
      else {
        _rtn = strValue;
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
  public boolean restoreForm(Form frmIn, boolean bForceVisible) { // TODO: Use of ByRef founded Public Function RestoreForm(ByRef frmIn As Form, Optional ByVal bForceVisible As Boolean = True) As Boolean
    boolean _rtn = false;
    // Comments  : Restores the form to its last-saved position (stored
    //             by the SaveForm method).
    // Parameters:
    //             frmIn         - object pointer to the Form object whose size and position should
    //                             be restored
    //             bForceVisible - if true, and if restoring form to its previous location would
    //                             result in the form being off-screen, then put the form in the
    //                             upper-left-hand corner of the screen.
    // Returns   : True if the form was restored; False if it was not restored
    //
    // Called by : All forms' Form_Load( ) procedure
    //
    // Source    : Total Visual SourceBook 2000
    //
    "RestoreForm"
.equals(Const cstrCurrentProc As String);
    Object varCurrent = null;
    int lngPos = 0;
    int lngLocation = 0;
    int lngLeft = 0;
    int lngTop = 0;
    int lngWidth = 0;
    int lngHeight = 0;
    int lngWindowState = 0;
    String strRegKey = "";

    try {

      strRegKey = m_strMainBranch+ MCSTRSLASH+ m_strCompanyBranch+ MCSTRSLASH+ m_strAppBranch+ MCSTRSLASH+ "4.0.0"+ MCSTRSLASH+ MCSTRSCREENPREFERENCES;
      //    strRegKey = m_strMainBranch & mcstrSlash & _
      //                m_strCompanyBranch & mcstrSlash & _
      //                m_strAppBranch & mcstrSlash & _
      //                m_strVersionBranch & mcstrSlash & _
      //                mcstrScreenPreferences

      varCurrent = modRegistry.registryGetKeyValue(EnumRegistryRootKeys.rRKHKEY_CURRENT_USER, strRegKey, frmIn.Name);

      if ((varCurrent+ "") == "") {
        varCurrent = "";
      }

      // Test for missing value (Indicates that form's position was not previously saved)
      if (LenB(varCurrent) == 0) {
        _rtn = false;
      } 
      else {
        // Find saved Left position
        lngPos = varCurrent.indexOf("L=");
        if (lngPos) {
          lngLocation = Val(varCurrent.substring(lngPos + 2));
          lngLeft = lngLocation;
        }

        // Find saved Top position
        lngPos = varCurrent.indexOf("T=");
        if (lngPos) {
          lngLocation = Val(varCurrent.substring(lngPos + 2));
          lngTop = lngLocation;
        }

        // Find saved Width
        lngPos = varCurrent.indexOf("W=");
        if (lngPos) {
          lngLocation = Val(varCurrent.substring(lngPos + 2));
          lngWidth = lngLocation;
        }

        // Find saved Height
        lngPos = varCurrent.indexOf("H=");
        if (lngPos) {
          lngLocation = Val(varCurrent.substring(lngPos + 2));
          lngHeight = lngLocation;
        }

        // Find saved WindowState value (minimized, maximized, normal)
        lngPos = varCurrent.indexOf("S=");
        if (lngPos) {
          lngLocation = Val(varCurrent.substring(lngPos + 2));
          lngWindowState = lngLocation;
        }

        // If form was saved minimized or maximized, change the state only
        if (lngWindowState == vbMinimized  || lngWindowState == vbMaximized) {
          frmIn.WindowState = lngWindowState;
        } 
        else {
          if (bForceVisible) {
            if ((lngLeft >= Screen.Width)) {
              lngLeft = Screen.Width - lngWidth;
            }

            if ((lngTop >= Screen.Height)) {
              lngTop = Screen.Height - lngHeight;
            }

            if ((lngTop < 0)) {
              lngTop = 0;
            }

            if ((lngLeft < 0)) {
              lngLeft = 0;
            }
          }

          // Move the form in a single statement instead of setting the
          // properties individually
          frmIn.Move(lngLeft, lngTop, lngWidth, lngHeight);
        }
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
  public void saveForm(Form frmIn) { // TODO: Use of ByRef founded Public Sub SaveForm(ByRef frmIn As Form)
    // Comments  : Saves the current size and position of the named form so that it can be
    //             subsequently restored via the RestoreForm method.
    // Parameters:
    //             frmIn         - object pointer to the Form object whose size and position should
    //                             be saved
    // Returns   : nothing
    //
    // Called by : All forms' Form_Unload( ) procedure
    //
    // Source    : Total Visual SourceBook 2000
    //
    "SaveForm"
.equals(Const cstrCurrentProc As String);
    String strNewValue = "";
    String strRegKey = "";

    try {

      strRegKey = m_strMainBranch+ MCSTRSLASH+ m_strCompanyBranch+ MCSTRSLASH+ m_strAppBranch+ MCSTRSLASH+ m_strVersionBranch+ MCSTRSLASH+ MCSTRSCREENPREFERENCES;

      strNewValue = "L="+ frmIn.Left+ ";T="+ frmIn.Top+ ";W="+ frmIn.Width+ ";H="+ frmIn.Height+ ";S="+ frmIn.WindowState+ ";";
      modRegistry.registrySetKeyValue(EnumRegistryRootKeys.rRKHKEY_CURRENT_USER, strRegKey, frmIn.Name, strNewValue, EnumRegistryValueType.rRKREGSZ);
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
  public void writeEntry(EnumRegistryRootKeys eHKRoot, String strRegKey, String strEntry, String strValue) {
    // Comments  : Writes a string entry into the registry
    //             at the location specified by the class properties and/or parameters
    // Parameters: eHKRoot    - the hive to which to write a value, e.g., HKLM, HKCU, etc.
    //             strRegKey  - the path under eHKRoot under which to write the strEntry key and its
    //                          strValue value
    //             strEntry   - The key under eHKRoot\strRegKey to which to write the strValue value
    //             strValue   - The string value under eHKRoot\strRegKey\strEntry to write
    //
    // Returns   : Nothing
    // Source    : Total Visual SourceBook 2000
    //
    "WriteEntry"
.equals(Const cstrCurrentProc As String);
    try {

      modRegistry.registrySetKeyValue(eHKRoot, strRegKey, strEntry, strValue, EnumRegistryValueType.rRKREGSZ);
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




///\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
//\                                                                  /
//|                        FRIEND Procedures                         |
///                                                                  \
//\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/


//////////////////////////////////////////////////////////////////////////////////////////////////
  public void fnLoadEnvironments() {
    // Comments  : Retrieves all Environment-related registry key values and loads
    //             them to an array member variable of type udtEnvironments
    // Parameters: N/A
    // Returns   : N/A
    // Called by : fnInit_gerhApp of modConstructors
    "fnLoadEnvironments"
.equals(Const cstrCurrentProc As String);
    String[] astrEnvKeys() = null;
    int lngEnvs = 0;
    int lngNbrOfEnvs = 0;
    String strRegKey = "";
    String strTempKey = "";

    try {

      strRegKey = m_strMainBranch+ MCSTRSLASH+ m_strCompanyBranch+ MCSTRSLASH+ m_strAppBranch+ MCSTRSLASH+ "4.0.0"+ MCSTRSLASH+ MCSTRENVIRONMENTS;
      //    strRegKey = m_strMainBranch & mcstrSlash & _
      //                m_strCompanyBranch & mcstrSlash & _
      //                m_strAppBranch & mcstrSlash & _
      //                m_strVersionBranch & mcstrSlash & _
      //                mcstrEnvironments

      // Build an array (astrEnvKeys) containing each defined Environment key in the
      // registry, e.g., Prod and IDev
      modRegistry.registryEnumerateSubKeys(EnumRegistryRootKeys.rRKHKEY_LOCAL_MACHINE, strRegKey, astrEnvKeys[], lngNbrOfEnvs);

      if (lngNbrOfEnvs < 1) {
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_FERR_NO_ENVS, MCSTRNAME+ cstrCurrentProc);
        // **TODO:** goto found: GoTo PROC_EXIT;
      }

      // Resize the UDT array (m_audtEnvironments) according to how many entries were defined in the
      // registry, then load the values associated with each such Environment key (e.g., database,
      // server, and environment name) to that array.
      G.redimPreserve(lngNbrOfEnvs - 1,  );

      for (lngEnvs = 0; lngEnvs <= lngNbrOfEnvs - 1; lngEnvs++) {
        m_audtEnvironments[lngEnvs].environmentKey = astrEnvKeys[lngEnvs];
        strTempKey = strRegKey+ MCSTRSLASH+ astrEnvKeys[lngEnvs];
        m_audtEnvironments[lngEnvs].activeDatabase = modRegistry.registryGetKeyValue(EnumRegistryRootKeys.rRKHKEY_LOCAL_MACHINE, strTempKey, MCSTRENV_ACTIVEDATABASE);
        m_audtEnvironments[lngEnvs].activeServer = modRegistry.registryGetKeyValue(EnumRegistryRootKeys.rRKHKEY_LOCAL_MACHINE, strTempKey, MCSTRENV_ACTIVESERVER);
        m_audtEnvironments[lngEnvs].archiveDatabase = modRegistry.registryGetKeyValue(EnumRegistryRootKeys.rRKHKEY_LOCAL_MACHINE, strTempKey, MCSTRENV_ARCHIVEDATABASE);
        m_audtEnvironments[lngEnvs].archiveServer = modRegistry.registryGetKeyValue(EnumRegistryRootKeys.rRKHKEY_LOCAL_MACHINE, strTempKey, MCSTRENV_ARCHIVESERVER);
        m_audtEnvironments[lngEnvs].port = modRegistry.registryGetKeyValue(EnumRegistryRootKeys.rRKHKEY_LOCAL_MACHINE, strTempKey, MCSTRENV_PORT);
        m_audtEnvironments[lngEnvs].environmentName = modRegistry.registryGetKeyValue(EnumRegistryRootKeys.rRKHKEY_LOCAL_MACHINE, strTempKey, MCSTRENV_ENVIRONMENTNAME);
        //SQL_INTEGRATED_SECURITY
        if ("TRUE"
.equals(modRegistry.registryGetKeyValue(EnumRegistryRootKeys.rRKHKEY_LOCAL_MACHINE, strTempKey, MCSTRENV_USESWINDOWSAUTHENTICATION).toUpperCase())) {
          m_audtEnvironments[lngEnvs].usesWindowsAuthentication = true;
        } 
        else {
          m_audtEnvironments[lngEnvs].usesWindowsAuthentication = false;
        }
        //SQL_INTEGRATED_SECURITY
      }
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    Erase(astrEnvKeys);

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
//|                      PRIVATE   Procedures                        |
///                                                                  \
//\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/

}

private class udtEnvironment {
    public String activeDatabase;
    public String activeServer;
    public String archiveDatabase;
    public String archiveServer;
    public String port;
    public Boolean usesWindowsAuthentication;//'SQL_INTEGRATED_SECURITY
    public String environmentName;
    public String environmentKey;
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


case class ApsappsettingsData(
              id: Option[Int],

              )

object Apsappsettingss extends Controller with ProvidesUser {

  val apsappsettingsForm = Form(
    mapping(
      "id" -> optional(number),

  )(ApsappsettingsData.apply)(ApsappsettingsData.unapply))

  implicit val apsappsettingsWrites = new Writes[Apsappsettings] {
    def writes(apsappsettings: Apsappsettings) = Json.obj(
      "id" -> Json.toJson(apsappsettings.id),
      C.ID -> Json.toJson(apsappsettings.id),

    )
  }

  def get(id: Int) = GetAction { implicit request =>
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.LIST_APSAPPSETTINGS), { user =>
      Ok(Json.toJson(Apsappsettings.get(user, id)))
    })
  }

  def update(id: Int) = PostAction { implicit request =>
    Logger.debug("in apsappsettingss.update")
    apsappsettingsForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      apsappsettings => {
        Logger.debug(s"form: ${apsappsettings.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.EDIT_APSAPPSETTINGS), { user =>
          Ok(
            Json.toJson(
              Apsappsettings.update(user,
                Apsappsettings(
                       id,

                ))))
        })
      }
    )
  }

  def create = PostAction { implicit request =>
    Logger.debug("in apsappsettingss.create")
    apsappsettingsForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      apsappsettings => {
        Logger.debug(s"form: ${apsappsettings.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.NEW_APSAPPSETTINGS), { user =>
          Ok(
            Json.toJson(
              Apsappsettings.create(user,
                Apsappsettings(

                ))))
        })
      }
    )
  }

  def delete(id: Int) = PostAction { implicit request =>
    Logger.debug("in apsappsettingss.delete")
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.DELETE_APSAPPSETTINGS), { user =>
      Apsappsettings.delete(user, id)
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

case class Apsappsettings(
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

object Apsappsettings {

  lazy val emptyApsappsettings = Apsappsettings(
)

  def apply(
      id: Int,
) = {

    new Apsappsettings(
      id,
)
  }

  def apply(
) = {

    new Apsappsettings(
)
  }

  private val apsappsettingsParser: RowParser[Apsappsettings] = {
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
        Apsappsettings(
              id,
,
              createdAt,
              updatedAt,
              updatedBy)
    }
  }

  def create(user: CompanyUser, apsappsettings: Apsappsettings): Apsappsettings = {
    save(user, apsappsettings, true)
  }

  def update(user: CompanyUser, apsappsettings: Apsappsettings): Apsappsettings = {
    save(user, apsappsettings, false)
  }

  private def save(user: CompanyUser, apsappsettings: Apsappsettings, isNew: Boolean): Apsappsettings = {
    def getFields = {
      List(

      )
    }
    def throwException = {
      throw new RuntimeException(s"Error when saving ${C.APSAPPSETTINGS}")
    }

    DBHelper.saveEx(
      user,
      Register(
        C.APSAPPSETTINGS,
        C.ID,
        apsappsettings.id,
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

  def load(user: CompanyUser, id: Int): Option[Apsappsettings] = {
    loadWhere(user, s"${C.ID} = {id}", 'id -> id)
  }

  def loadWhere(user: CompanyUser, where: String, args : scala.Tuple2[scala.Any, anorm.ParameterValue[_]]*) = {
    DB.withConnection(user.database.database) { implicit connection =>
      SQL(s"SELECT t1.*, t2.${C.FK_NAME} FROM ${C.APSAPPSETTINGS} t1 INNER JOIN ${C.???} t2 ON t1.${C.FK_ID} = t2.${C.FK_ID} WHERE $where")
        .on(args: _*)
        .as(apsappsettingsParser.singleOpt)
    }
  }

  def delete(user: CompanyUser, id: Int) = {
    DB.withConnection(user.database.database) { implicit connection =>
      try {
        SQL(s"DELETE FROM ${C.APSAPPSETTINGS} WHERE ${C.ID} = {id}")
        .on('id -> id)
        .executeUpdate
      } catch {
        case NonFatal(e) => {
          Logger.error(s"can't delete a ${C.APSAPPSETTINGS}. ${C.ID} id: $id. Error ${e.toString}")
          throw e
        }
      }
    }
  }

  def get(user: CompanyUser, id: Int): Apsappsettings = {
    load(user, id) match {
      case Some(p) => p
      case None => emptyApsappsettings
    }
  }
}


// Router

GET     /api/v1/general/apsappsettings/:id              controllers.logged.modules.general.Apsappsettingss.get(id: Int)
POST    /api/v1/general/apsappsettings                  controllers.logged.modules.general.Apsappsettingss.create
PUT     /api/v1/general/apsappsettings/:id              controllers.logged.modules.general.Apsappsettingss.update(id: Int)
DELETE  /api/v1/general/apsappsettings/:id              controllers.logged.modules.general.Apsappsettingss.delete(id: Int)




/**/
