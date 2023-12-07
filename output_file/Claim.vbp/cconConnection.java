
import java.sql.Connection;

public class cconConnection {

  //--------------------------------------------------------------------------
  // Module     : cconConnection
  // Description: Instantiated by modStartup.Sub_Main
  //              This object WILL know the ErrorLogger.cls
  //
  // Procedures :
  //    Private   Class_Initialize()
  //    Private   Class_Terminate()
  //    Public    Property Get ADOConn() As ADODB.Connection
  //    Public    Property Get LastLogonEnviron() As String
  //    Public    Property Get LastLogonIsSpecialUser() As Boolean
  //    Public    Property Get LastLogonPassword() As String
  //    Public    Property Get State() As ObjectStateEnum
  //    Public    BeginTrans() As Boolean
  //    Public    CommitTrans() as boolean
  //    Public    Connect(ByVal strEnviron As String, ByVal strUserID As String, _
  //                  ByVal strPassword As String) As Boolean
  //    Public    Disconnect() As Boolean
  //    Public    RollbackTrans()
  //
  // Revision History:
  //    10/04/02 - Betsy - Cloned from TRS.
  //--------------------------------------------------------------------------
//Option Explicit
  *Option Compare Binary

  *#Const DEBUG_ERH = False
  *#Const DEBUG_RST = False

  private static final String MCSTRNAME = "cconConnection.";
  private Connection m_connection;
  private boolean m_lastLogonIsSpecialUser = false;
  private String m_lastLogonUserID = "";
  private String m_lastLogonPassword = "";
  private String m_lastLogonEnviron = "";
  private boolean m_initialized = false;
  private boolean m_isTransactionActive = false;




  ///\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
  //\                                                                  /
  //|         CLASS_INITIALIZE / CLASS_TERMINATE   Procedures         |
  ///                                                                  \
  //\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/

  //////////////////////////////////////////////////////////////////////////////////////////////////
  private void class_Initialize() {
    // **************************************************************************
    // Function  : Class_Initialize
    // Purpose   : Starting point for the object
    // Parameters: N/A
    // Returns   : True/False
    // SXS 08/04/2004  Error4048   Added support for error 4048 to trap SQL error -2147217871 (Timeout exceeded).
    //                     Also Also default ADO connection Timeout changed from 30 seconds to 90 seconds
    // **************************************************************************
    "Class_Initialize"
.equals(Const cstrCurrentProc As String);
    try {

      m_connection = new ADODB.Connection();

      //Error4048
      m_connection.CommandTimeout = 90;
      m_initialized = true;
      m_isTransactionActive = false;
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
  private void class_Terminate() {
    // **************************************************************************
    // Function  : Class_Terminate
    // Purpose   : Close the object
    // Parameters: N/A
    // Returns   : N/A
    // **************************************************************************
    "Class_Terminate"
.equals(Const cstrCurrentProc As String);
    try {

      disconnect();
      modGeneral.fnFreeObject(m_connection);

      m_initialized = false;
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
  public Connection getADOConn() {
    // **************************************************************************
    // Function  : GetADOConn
    // Purpose   :
    // Parameters: N/A
    // Returns   : True/False
    // **************************************************************************
    "Property Get ADOConn"
.equals(Const cstrCurrentProc As String);
    try {

      return m_connection;
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
  public String getLastLogonEnviron() {
    String _rtn = "";
    // **************************************************************************
    // Function  : LastLogonEnviron
    // Purpose   :
    // Parameters: N/A
    // Returns   :
    // **************************************************************************
    "Property Get LastLogonEnviron"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = m_lastLogonEnviron;
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
  public boolean getLastLogonIsSpecialUser() {
    boolean _rtn = false;
    // **************************************************************************
    // Function  : LastLogonIsSpecialUser
    // Purpose   :
    // Parameters: N/A
    // Returns   :
    // **************************************************************************
    "Property Get LastLogonIsSpecialUser"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = m_lastLogonIsSpecialUser;
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
  public String getLastLogonPassword() {
    String _rtn = "";
    // **************************************************************************
    // Function  : LastLogonPassword
    // Purpose   :
    // Parameters: N/A
    // Returns   :
    // **************************************************************************
    "Property Get LastLogonPassword"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = m_lastLogonPassword;
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
  public String getLastLogOnUserID() {
    String _rtn = "";
    // **************************************************************************
    // Function  : LastLogonUserID
    // Purpose   :
    // Parameters: N/A
    // Returns   :
    // **************************************************************************
    "Property Get LastLogonUserID"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = m_lastLogonUserID;
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
  public ObjectStateEnum getState() {
    ObjectStateEnum _rtn = null;
    // **************************************************************************
    // Function  : GetState
    // Purpose   :
    // Parameters: N/A
    // Returns   : True/False
    // **************************************************************************
    "Property Get State"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = m_connection.State;
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

//////////////////////////////////////////////////////////////////////////////////////////////////
  public boolean beginTrans() {
    boolean _rtn = false;
    //--------------------------------------------------------------------------
    // Procedure:   BeginTrans
    // Description: Will start transaction processing on any complex database
    //              activity. This allows groups of transactions, such as might
    //              be done within a stored procedure, to be rolled back or
    //              committed as a group.
    // Returns:     True if successful; False otherwise
    // Params:      N/A
    //-----------------------------------------------------------------------------
    "BeginTrans"
.equals(Const cstrCurrentProc As String);
    try {

      m_connection.BeginTrans;

      // Set flag indicating we've started a transaction
      m_isTransactionActive = true;

      _rtn = true;
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
  public boolean commitTrans() {
    boolean _rtn = false;
    //--------------------------------------------------------------------------
    // Procedure:   CommitTrans
    // Description: Will end transaction processing on any complex database
    //              activity and commit changes made since BeginTrans( ) was called.
    // Returns:     True if successful; False otherwise
    // Params:      N/A
    //-----------------------------------------------------------------------------
    "CommitTrans"
.equals(Const cstrCurrentProc As String);

    try {

      if (m_isTransactionActive) {
        m_connection.CommitTrans;
        m_isTransactionActive = false;
      }

      _rtn = true;
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
  public boolean connect(String strEnviron, String strUserID, String strPassword, boolean bActiveDB) {
    boolean _rtn = false;
    // **************************************************************************
    // Function  : Connect
    // Purpose   : Will connect to the specified environment based on the user info.
    // Parameters: Parameters:
    //               strEnviron  (input) = User's selection from the Log On screen's
    //                                     Environment combo box
    //               strUserID   (input) = will contain the ACF2 of the user.
    //               strPassword (input) = related to UserID, password for that user, or Claims app.
    // Returns   : True if successful; False otherwise
    // **************************************************************************
    "Connect"
.equals(Const cstrCurrentProc As String);
    String strConnectPart = "";

    try {

      if (m_connection.State == adStateOpen) {
        m_connection.Close;
      }
      //' SQL_INTEGRATED_SECURITY
      modGeneral.fnFreeObject(m_connection);

      // Make the Connection to the database
      if ((LenB(gapsApp.getActiveServer(strEnviron)).equals(0)) || (LenB(gapsApp.getActiveDatabase(strEnviron)).equals(0)) || (LenB(gapsApp.getArchiveDatabase(strEnviron)).equals(0)) || (LenB(gapsApp.getArchiveServer(strEnviron)).equals(0)) || (LenB(gapsApp.getPort(strEnviron)).equals(0))) {
        // gcRES_NERR_ENV_REG_ENTRIES_MISSING (4010) = One or more registry entries
        //     that define how to connect to the selected
        //     Environment (@@1) are missing. Without all of these entries,
        //     the app cannot connect to the database.
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_ENV_REG_ENTRIES_MISSING, MCSTRNAME+ cstrCurrentProc, strEnviron);
        // **TODO:** goto found: GoTo PROC_EXIT;
      }

      //' SQL_INTEGRATED_SECURITY
      m_connection = new ADODB.Connection();
      m_connection.Provider = "SQLOLEDB";
      m_connection.Mode = adModeReadWrite;
      m_connection.CursorLocation = adUseClient;
      m_connection.Properties("Prompt").value = adPromptNever;
      // OLE DB Services = -2 added to disable connection pooling. Having it enabled It causes problems
      // with application roles. See <http://support.microsoft.com/search/preview.aspx?scid=kb;en-us;Q229564>
      // for more info.
      m_connection.Properties("OLE DB Services").value = -2;
      // The following logic was replaced with building the full .ConnectionString property
      // since VB Watch Profiler identified these lines as being inefficient. This noticeably sped
      // up the Log On screen.
      //       If bActiveDB Then
      //            .Properties("Data Source").value = gapsApp.ActiveServer(strEnviron)
      //            .Properties("Initial Catalog").value = gapsApp.ActiveDatabase(strEnviron)
      //       Else
      //           .Properties("Data Source").value = gapsApp.ArchiveServer(strEnviron)
      //             .Properties("Initial Catalog").value = gapsApp.ArchiveDatabase(strEnviron)
      //       End If
      //       .Properties("User Id").value = strUserID
      //       .Properties("Password").value = strPassword

      //SQL_INTEGRATED_SECURITY
      switch (gapsApp.getUsesWindowsAuthentication(strEnviron)) {
        case  True:
          strConnectPart = "Integrated Security=SSPI"+ ";";
          break;

        default:
          strConnectPart = "Password="+ strPassword+ ";User ID="+ strUserID+ ";";
          break;
      }
      //SQL_INTEGRATED_SECURITY

      if (bActiveDB) {
        m_connection.ConnectionString = strConnectPart+ "Initial Catalog="+ gapsApp.getActiveDatabase(strEnviron)+ ";Data Source="+ gapsApp.getActiveServer(strEnviron)+ ","+ gapsApp.getPort(strEnviron);
      } 
      else {
        m_connection.ConnectionString = strConnectPart+ "Initial Catalog="+ gapsApp.getArchiveDatabase(strEnviron)+ ";Data Source="+ gapsApp.getArchiveServer(strEnviron)+ ","+ gapsApp.getPort(strEnviron);
      }

      m_connection.Open;

      if (m_connection.State != adStateOpen) {
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_CONNECTION_FAILURE, MCSTRNAME+ cstrCurrentProc, strEnviron, m_connection.State);
      } 
      else {
        _rtn = true;

        if (bActiveDB) {
          // Save the ID, Password and Environment of the logged on user. This
          // will be used, if necessary, to log the user on to the corresponding
          // ArchiveDB
          m_lastLogonUserID = strUserID;
          m_lastLogonPassword = strPassword;
          m_lastLogonEnviron = strEnviron;
          m_lastLogonIsSpecialUser = fnIsSpecialUser();
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
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    _rtn = false;
    switch (VBA.ex.Number) {
      case  -2147467259:
      case  -2147217843:
      case  -2147467259:
        // -2147467259 = Cannot open database requested in login 'indppvul_pr'. Login fails.
        // -2147217843 = Not a valid password'
        // -2147467259 = SQL Server does not exist or access denied.
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_LOGON_FAILURE, MCSTRNAME+ cstrCurrentProc, VBA.ex.Number);
      //' Operation is not allowed
        break;

      case  3704:
        /**TODO:** resume found: Resume(Next)*/;
        break;

      default:
        // Any error stemming from this procedure is manifested as
        // a Log On failure that cites a bad User ID, Password or permissions
        // as the likely cause. Therefore, any caller of this procedure
        // doesn't have to check the return value; it will always be True
        // since a False would be raised to the caller's error handler
        // due to the propagation done in this proc's PROC_EXIT.
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_LOGON_FAILURE, MCSTRNAME+ cstrCurrentProc, VBA.ex.Number);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public boolean disconnect() {
    boolean _rtn = false;
    // **************************************************************************
    // Function  : Disconnect
    // Purpose   : Will disconnect the ADO connection object.
    // Parameters: N/A
    // Returns   : True if successful; False otherwise
    // **************************************************************************
    "Disconnect"
.equals(Const cstrCurrentProc As String);
    try {

      if (m_connection.State == adStateClosed) {
        // **TODO:** goto found: GoTo PROC_EXIT;
      }

      m_connection.Close;
      _rtn = true;
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
  private boolean fnIsSpecialUser() {
    boolean _rtn = false;
    //--------------------------------------------------------------------------
    // Procedure:   fnIsSpecialUser
    // Description: This method determines whether the user is a member of
    //              the Support or UserAdmin roles. Membership in this roles
    //              indicate the user has access to restricted areas of the
    //              application (the Current Rate and State Rule screens).
    //
    // Params:
    //
    // Returns:     True if the user is a member of these roles; false otherwise.
    //-----------------------------------------------------------------------------
    "fnIsSpecialUser"
.equals(Const cstrCurrentProc As String);
    //' # of input or output params sproc expects
    Const(clngSprocParamCount As Long == 2);
    //' Stored procedure to execute
    "sp_helpuser"
.equals(Const cstrSproc As String);
    ADODB.Parameter prmReturnValue = null;
    ADODB.Parameter prmName_in_DB = null;
    DBRecordSet rstTemp = null;
    New adwTemp = null; cadwADOWrapper
    String strGroupName = "";

    try {

      // Connect to the specified environment using the Dummy App ID,
      // then execute the sp_helpuser sproc.
      if (!(adwTemp.CommandSetSproc(cstrSproc, this))) {
        // **TODO:** goto found: GoTo PROC_EXIT;
      }

      //*TODO:** can't found type for with block
      //*With adwTemp.ADOCommand
      __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = adwTemp.ADOCommand;
      // ---Parameter #1---
      // Define the return value that represents the error code (i.e. reason) why
      // the stored procedure failed.
      prmReturnValue = w___TYPE_NOT_FOUND.CreateParameter(Name:="@return_value", Type:=adInteger, Direction:=adParamReturnValue, .value:=Null);
      w___TYPE_NOT_FOUND.Parameters.Append(prmReturnValue);

      // ---Parameter #2---
      // Define the name_in_db input parameter, which represents the User ID being checked
      prmName_in_DB = w___TYPE_NOT_FOUND.CreateParameter(Name:="@name_in_db", Type:=adVarChar, Direction:=adParamInput, Size:=255, .value:=fnNullIfZLS(varIn:=m_LastLogonUserID, bHandleEmbeddedQuotes:=True));
      w___TYPE_NOT_FOUND.Parameters.Append(prmName_in_DB);

      rstTemp = w___TYPE_NOT_FOUND.Execute();

      if (!(rstTemp.BOF && rstTemp.EOF)) {
        rstTemp.MoveFirst;
        while (!rstTemp.EOF) {
          strGroupName = rstTemp.Fields("GroupName").value.toUpperCase();
          if ("SUPPORT".equals(strGroupName) || "USERADMIN"
.equals(strGroupName)) {
            _rtn = true;
          }
          rstTemp.MoveNext;
        }
      }
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    modGeneral.fnFreeObject(prmReturnValue);
    modGeneral.fnFreeObject(prmName_in_DB);
    modGeneral.fnFreeRecordset(rstTemp);
    modGeneral.fnFreeObject(adwTemp);

    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
      //' 4013
      case  vbObjectError + modResConstants.gCRES_NERR_LOGON_FAILURE:
        // This environment will be considered "not authorized"
        VBA.ex.Clear;
        modGeneral.gerhApp.clear();
        // **TODO:** goto found: GoTo PROC_EXIT;
        //Resume Next
      //' The name supplied (xxx) is not a user, role or aliased login
        break;

      case  -2147217900:
        // This environment will be considered "not authorized"
        VBA.ex.Clear;
        modGeneral.gerhApp.clear();
        // **TODO:** goto found: GoTo PROC_EXIT;
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
  public void rollbackTrans() {
    //--------------------------------------------------------------------------
    // Procedure:   RollbackTrans
    // Description: Will undo any DB changes made since BeginTrans( ) was called
    // Params:      N/A
    // Returns:     N/A
    // Date:        01/09/2002
    //-----------------------------------------------------------------------------
    "RollbackTrans"
.equals(Const cstrCurrentProc As String);

    // NOTE:  NO ERROR HANDLER should be active here since we want
    //        the Rollback to proceed even if an error has been logged but not
    //        yet reported to the user!

    if (m_isTransactionActive) {
      m_connection.RollbackTrans;
      m_isTransactionActive = false;
    }
  }



  //////////////////////////////////////////////////////////////////////////////////////////////////
  public boolean setAppRole(String strRoleName, String strRolePassword) {
    boolean _rtn = false;
    //--------------------------------------------------------------------------
    // Procedure:   SetAppRole
    // Description: Puts the specified App Role into effect, so that the role's
    //              permissions will override that of the logged on user.
    //              This proc should be called **after** the user has been logged on.
    // Returns:     True if successful; False otherwise
    // Params:      N/A
    //-----------------------------------------------------------------------------
    "SetAppRole"
.equals(Const cstrCurrentProc As String);
    //' Stored procedure to execute
    "sp_setapprole"
.equals(Const cstrSproc As String);
    cadwADOWrapper adwTemp = null;
    ADODB.Parameter prmReturnValue = null;
    ADODB.Parameter prmRoleName = null;
    ADODB.Parameter prmPassword = null;
    DBRecordSet rstTemp = null;

    try {

      adwTemp = new cadwADOWrapper();

      // Set the sproc name and set **this** connection object as the active connection
      if (!(adwTemp.commandSetSproc(cstrSproc, this))) {
        // **TODO:** goto found: GoTo PROC_EXIT;
      }

      ADODB.Command w_aDOCommand = adwTemp.getADOCommand();
      // ---Parameter #1---
      // Define the return value that represents the error code (i.e. reason) why
      // the stored procedure failed.
      prmReturnValue = w_aDOCommand.CreateParameter(Name:="@return_value", Type:=adInteger, Direction:=adParamReturnValue, .value:=Null);
      w_aDOCommand.Parameters.Append(prmReturnValue);

      // ---Parameter #2---
      // Define the RoleName input parameter, which reflects *which*
      // application role should be put into effect
      prmRoleName = w_aDOCommand.CreateParameter(Name:="@rolename", Type:=adVarChar, Direction:=adParamInput, Size:=255, .value:=fnNullIfZLS(varIn:=strRoleName, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmRoleName);

      // ---Parameter #3---
      // Define the Password input parameter, which reflects the password
      // for the specified application role
      prmPassword = w_aDOCommand.CreateParameter(Name:="@password", Type:=adVarChar, Direction:=adParamInput, Size:=255, .value:=fnNullIfZLS(varIn:=strRolePassword, bHandleEmbeddedQuotes:=True));
      w_aDOCommand.Parameters.Append(prmPassword);

      w_aDOCommand.Execute;

      _rtn = true;
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    modGeneral.fnFreeObject(adwTemp);
    modGeneral.fnFreeObject(prmReturnValue);
    modGeneral.fnFreeObject(prmRoleName);
    modGeneral.fnFreeObject(prmPassword);
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


case class ConconnectionData(
              id: Option[Int],

              )

object Conconnections extends Controller with ProvidesUser {

  val conconnectionForm = Form(
    mapping(
      "id" -> optional(number),

  )(ConconnectionData.apply)(ConconnectionData.unapply))

  implicit val conconnectionWrites = new Writes[Conconnection] {
    def writes(conconnection: Conconnection) = Json.obj(
      "id" -> Json.toJson(conconnection.id),
      C.ID -> Json.toJson(conconnection.id),

    )
  }

  def get(id: Int) = GetAction { implicit request =>
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.LIST_CONCONNECTION), { user =>
      Ok(Json.toJson(Conconnection.get(user, id)))
    })
  }

  def update(id: Int) = PostAction { implicit request =>
    Logger.debug("in conconnections.update")
    conconnectionForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      conconnection => {
        Logger.debug(s"form: ${conconnection.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.EDIT_CONCONNECTION), { user =>
          Ok(
            Json.toJson(
              Conconnection.update(user,
                Conconnection(
                       id,

                ))))
        })
      }
    )
  }

  def create = PostAction { implicit request =>
    Logger.debug("in conconnections.create")
    conconnectionForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      conconnection => {
        Logger.debug(s"form: ${conconnection.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.NEW_CONCONNECTION), { user =>
          Ok(
            Json.toJson(
              Conconnection.create(user,
                Conconnection(

                ))))
        })
      }
    )
  }

  def delete(id: Int) = PostAction { implicit request =>
    Logger.debug("in conconnections.delete")
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.DELETE_CONCONNECTION), { user =>
      Conconnection.delete(user, id)
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

case class Conconnection(
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

object Conconnection {

  lazy val emptyConconnection = Conconnection(
)

  def apply(
      id: Int,
) = {

    new Conconnection(
      id,
)
  }

  def apply(
) = {

    new Conconnection(
)
  }

  private val conconnectionParser: RowParser[Conconnection] = {
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
        Conconnection(
              id,
,
              createdAt,
              updatedAt,
              updatedBy)
    }
  }

  def create(user: CompanyUser, conconnection: Conconnection): Conconnection = {
    save(user, conconnection, true)
  }

  def update(user: CompanyUser, conconnection: Conconnection): Conconnection = {
    save(user, conconnection, false)
  }

  private def save(user: CompanyUser, conconnection: Conconnection, isNew: Boolean): Conconnection = {
    def getFields = {
      List(

      )
    }
    def throwException = {
      throw new RuntimeException(s"Error when saving ${C.CONCONNECTION}")
    }

    DBHelper.saveEx(
      user,
      Register(
        C.CONCONNECTION,
        C.ID,
        conconnection.id,
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

  def load(user: CompanyUser, id: Int): Option[Conconnection] = {
    loadWhere(user, s"${C.ID} = {id}", 'id -> id)
  }

  def loadWhere(user: CompanyUser, where: String, args : scala.Tuple2[scala.Any, anorm.ParameterValue[_]]*) = {
    DB.withConnection(user.database.database) { implicit connection =>
      SQL(s"SELECT t1.*, t2.${C.FK_NAME} FROM ${C.CONCONNECTION} t1 INNER JOIN ${C.???} t2 ON t1.${C.FK_ID} = t2.${C.FK_ID} WHERE $where")
        .on(args: _*)
        .as(conconnectionParser.singleOpt)
    }
  }

  def delete(user: CompanyUser, id: Int) = {
    DB.withConnection(user.database.database) { implicit connection =>
      try {
        SQL(s"DELETE FROM ${C.CONCONNECTION} WHERE ${C.ID} = {id}")
        .on('id -> id)
        .executeUpdate
      } catch {
        case NonFatal(e) => {
          Logger.error(s"can't delete a ${C.CONCONNECTION}. ${C.ID} id: $id. Error ${e.toString}")
          throw e
        }
      }
    }
  }

  def get(user: CompanyUser, id: Int): Conconnection = {
    load(user, id) match {
      case Some(p) => p
      case None => emptyConconnection
    }
  }
}


// Router

GET     /api/v1/general/conconnection/:id              controllers.logged.modules.general.Conconnections.get(id: Int)
POST    /api/v1/general/conconnection                  controllers.logged.modules.general.Conconnections.create
PUT     /api/v1/general/conconnection/:id              controllers.logged.modules.general.Conconnections.update(id: Int)
DELETE  /api/v1/general/conconnection/:id              controllers.logged.modules.general.Conconnections.delete(id: Int)




/**/
