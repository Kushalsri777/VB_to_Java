public class cadwADOWrapper {


  //--------------------------------------------------------------------------
  // Object:               cadwADOWrapper
  // Object Description:   Complete database
  // Revision History:
  //    10/04/02 - Betsy - Cloned from TRS.
  //
  // Procedures  :
  //   Private     Class_Initialize()
  //   Private     Class_Terminate()
  //   Public      Property Get ADOCommand() As ADODB.Command
  //   Public      CommandInitialize()
  //   Public      CommandSetConnection(Optional ByRef pconIn As cconConnection) As Boolean
  //   Public      CommandSetSproc(strSprocName As String, _
  //                   Optional ByVal lngParamCount As Long = 0) As Boolean
  //   Public      Execute_SQL_AsRST(ByRef pconIn As cconConnection, ByVal strSQL As String) _
  //                   As ADODB.Recordset
  //   Public      Execute_SQL_UpdateableDisconnectedRST(ByRef pconIn As cconConnection, _
  //                   ByVal strSQL As String) As ADODB.Recordset
  //   Public      GetMetaData_Columns(ByVal strTableName As String, _
  //                   ByRef prstInOut As ADODB.Recordset) As Boolean
  //   Public      GetMetaData_PrimaryKeys(ByVal strTableName As String, _
  //                   ByRef prstInOut As ADODB.Recordset) As Boolean
  //   Public      MoveFirst(ByRef prstTemp As ADODB.Recordset) As Boolean
  //   Public      MoveLast(ByRef prstTemp As ADODB.Recordset) As Boolean
  //   Public      MoveNext(ByRef prstTemp As ADODB.Recordset) As Boolean
  //   Public      MovePrev(ByRef prstTemp As ADODB.Recordset) As Boolean
  //
  //
  //-----------------------------------------------------------------------------
//Option Explicit
  *Option Compare Binary

  private static final String MCSTRNAME = "cadwADOWrapper.";

  private ADODB.Command m_cmdADOCommand;



  ///\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
  //\                                                                  /
  //|          CLASS_INITIALIZE / CLASS_TERMINATE    Procedures        |
  ///                                                                  \
  //\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/

  //////////////////////////////////////////////////////////////////////////////////////////////////
  private void class_Initialize() {
    //--------------------------------------------------------------------------
    // Procedure:   Class_Initialize
    // Description: Starting point for the class.
    // Params:      None
    // Returns:     Boolean
    //-----------------------------------------------------------------------------
    "Class_Initialize"
.equals(Const cstrCurrentProc As String);

    try {

      m_cmdADOCommand = new ADODB.Command();
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
    //--------------------------------------------------------------------------
    // Procedure:   Terminate
    // Description: kill the class
    // Returns:
    //-----------------------------------------------------------------------------
    "Class_Terminate"
.equals(Const cstrCurrentProc As String);

    try {

      if (modGeneral.fnIsObject(m_cmdADOCommand)) {
        m_cmdADOCommand.ActiveConnection = null;
        modGeneral.fnFreeObject(m_cmdADOCommand);
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



///\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
//\                                                                  /
//|                PROPERTY GET/LET    Procedures                    |
///                                                                  \
//\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/

//////////////////////////////////////////////////////////////////////////////////////////////////
  public ADODB.Command getADOCommand() {
    //--------------------------------------------------------------------------
    // Procedure:   Get ADOCommand
    // Description: Get the ADODB Command object (used to define and
    //              retrieve parameters and recordsets when calling
    //              stored procedures)
    // Returns:     Boolean
    //-----------------------------------------------------------------------------
    "Property Get ADOCommand"
.equals(Const cstrCurrentProc As String);

    try {

      return m_cmdADOCommand;
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


///\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
//\                                                                  /
//|                        PUBLIC  Procedures                        |
///                                                                  \
//\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/

//////////////////////////////////////////////////////////////////////////////////////////////////
  public boolean commandInitialize() {
    boolean _rtn = false;
    //--------------------------------------------------------------------------
    // Procedure:   CommandInitialize
    // Description: Initializes the ADODB.Command object of this class
    //              so the caller knows it is ready for reuse. If any
    //              previous use defined parameters or established
    //              recordsets, they will be deleted.
    // Returns:     True if successfully initialized; False otherwise
    // Params:      N/A
    //
    // Called by:   CommandSetSproc( )
    //
    // Modified:
    //   05/17/02 BAW Removed the call to CommandSetConnection since this
    //                was interfering with the LogOn screen (as of 2C) using
    //                the Authenticate object (which calls a sproc)...before
    //                the gconAppActive connection object was instanatiated.
    //   04/23/02 BAW Parameter deletion was getting an error. Had to change it to
    //                count backward to avoid this.
    //--------------------------------------------------------------------------
    "CommandInitialize"
.equals(Const cstrCurrentProc As String);
    int intI = 0;

    try {

      if (modGeneral.fnIsObject(m_cmdADOCommand)) {
        modGeneral.fnFreeObject(m_cmdADOCommand.ActiveConnection);
        // Setting ActiveConnection to Nothing still leaves Parameters collection intact,
        // so make sure they're wiped out for the next use of this Command object.
        //*TODO:** can't found type for with block
        //*With .Parameters
        __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = m_cmdADOCommand.Parameters;
        intI = w___TYPE_NOT_FOUND.Count;
        while (intI > 0) {
          intI = intI - 1;
          w___TYPE_NOT_FOUND.DELETE(intI);
        }
        m_cmdADOCommand.CommandText = "";
      } 
      else {
        m_cmdADOCommand = new ADODB.Command();
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
  public boolean commandSetConnection(cconConnection pconIn) { // TODO: Use of ByRef founded Public Function CommandSetConnection(Optional ByRef pconIn As cconConnection) As Boolean
    boolean _rtn = false;
    //--------------------------------------------------------------------------
    // Procedure:   CommandSetConnection
    // Description: Sets the ActiveConnection for the Command object. By default,
    //              the global Connection object (gconAppActive) will be used.
    // Returns:     True if successful; False otherwise
    // Params:
    //     pconIn (in/out) - a pointer to the cconConnection object to use, e.g., the
    //                       active or archive database.
    //--------------------------------------------------------------------------
    "CommandSetConnection"
.equals(Const cstrCurrentProc As String);

    try {

      // Default to using the global Connection object that points to the active DB,
      // unless the caller has specified another.
      if ((pconIn == null)) {
        m_cmdADOCommand.ActiveConnection = modGeneral.gconAppActive.getADOConn();
      } 
      else {
        m_cmdADOCommand.ActiveConnection = pconIn.getADOConn();
      }

      // Make sure Connection is open.
      // (It should be for the entire session, once the user logs on)
      Debug.Assert(m_cmdADOCommand.ActiveConnection.cconConnection.getState() == adStateOpen);

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
  public boolean commandSetSproc(String strSprocName, cconConnection pconIn) { // TODO: Use of ByRef founded Public Function CommandSetSproc(ByVal strSprocName As String, Optional ByRef pconIn As cconConnection) As Boolean
    boolean _rtn = false;
    //--------------------------------------------------------------------------
    // Procedure:   CommandSetSproc
    // Description: Defines the stored procedure that will be called
    //              with this ADODB.Command object and establishes it
    //              as a method of that object.
    //
    // Params:
    //   strSprocName (in)  - the name of the stored procedure to execute
    //   lngParamCount (in) - how many input, output or other parameters
    //                        will be used by this sproc
    //   pconIn (in/out)    - the name of the cconConnection object to use
    //                        when executing the sproc. If not supplied,
    //                        the gconAppActive object will be used.
    //
    // Returns:     True if successful; False otherwise
    //
    // Modifications:
    //   05/17/02 BAW - Added call to CommandInitialize and CommandSetConnection
    //                  to allow this to work for all existing code without
    //                  modification, but also support the new-to-Phase2C
    //                  execution of this proc by the LogOn screen (before the
    //                  gconAppActive object is instantiated).
    //--------------------------------------------------------------------------
    "CommandSetSproc"
.equals(Const cstrCurrentProc As String);

    try {

      // Wipe out existing parameters and close the ActiveConnection
      commandInitialize();
      // Reestablish a connection to the specified Connection or (default) gconAppActive
      commandSetConnection(pconIn);

      m_cmdADOCommand.CommandType = adCmdStoredProc;
      m_cmdADOCommand.CommandText = strSprocName;

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
  public DBRecordSet execute_SQL_AsRST(cconConnection pconIn, String strSQL) { // TODO: Use of ByRef founded Public Function Execute_SQL_AsRST(ByRef pconIn As cconConnection, ByVal strSQL As String) As ADODB.Recordset
    //--------------------------------------------------------------------------
    // Procedure:   Execute_SQL_AsRST
    //
    //              NOTE: This should **only** be used to apply selection criteria
    //                    to a "SELECT" statement, e.g., for reporting purposes.
    //                    Even then, that should really be made into a stored procedure
    //                    that accepts parameters so this function can be made
    //                    obsolete and the prospective caller can use
    //                    Execute_Sproc_AsRST( ) instead.
    //
    //                    !TODO! Create sprocs and change each caller of this proc
    //                           per the above.
    //
    // Description: Executes the specified SQL to build a recordset
    // Returns:     ADODB.Recordset that was built using strSQL
    // Params:      pconIn  input/output)   The Connection object to use
    //              strSQL  input           What is to be executed against the DB
    //-----------------------------------------------------------------------------
    "Execute_SQL_AsRST"
.equals(Const cstrCurrentProc As String);

    try {

      // Check the Status on the Global Connection object to ensure that it is still open.
      if (pconIn.getADOConn().State == adStateClosed) {
        //!TODO! Either do an app-specific error or just restablish the connection
        // **TODO:** goto found: GoTo PROC_ERR;
      }

      // Build the recordset, by executing the SQL statement
      return pconIn.getADOConn().Execute(strSQL);
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here

    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }

    return null;
    // **TODO:** label found: PROC_ERR:;
    // - - - - - - - - - - - Keep these in here!!! - - - - - - - - - - -
    Debug.Print("In "+ MCSTRNAME+ cstrCurrentProc+ ", this error was generated: "+ "\\n"+ VBA.ex.Number+ " - "+ VBA.ex.Description);
    Debug.Print("    strSQL="+ strSQL);
    // - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

    modGeneral.fnFreeObject(execute_SQL_AsRST());

    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      //' Invalid object name
      case  -2147217865  :
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_FERR_SQL_STMT_OBJECT_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, strSQL);
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
  public DBRecordSet execute_SQL_UpdateableDisconnectedRST(cconConnection pconIn, String strSQL) { // TODO: Use of ByRef founded Public Function Execute_SQL_UpdateableDisconnectedRST(ByRef pconIn As cconConnection, ByVal strSQL As String) As ADODB.Recordset
    //--------------------------------------------------------------------------
    // Procedure:   Execute_SQL_UpdateableDisconnectedRST
    //
    // Description: Executes the specified SQL to build an updateable recordset
    //              and then disconnects it. This procedure should **ONLY**
    //              be used by the Select Reports screen, which has to "manufacture"
    //              data for some of the complex PPVUL reports using a
    //              recordset (typically created from a view) as merely a starting
    //              point.
    //
    // Returns:     Disconnected ADODB.Recordset that was built using strSQL
    // Params:      pconIn  input/output)   The Connection object to use
    //              strSQL  input           What is to be executed against the DB
    //-----------------------------------------------------------------------------
    "Execute_SQL_UpdateableDisconnectedRST"
.equals(Const cstrCurrentProc As String);

    try {

      // Check the Status on the Global Connection object to ensure that it is still open.
      if (pconIn.getADOConn().State == adStateClosed) {
        //!TODO! Either do an app-specific error or just restablish the connection
        // **TODO:** goto found: GoTo PROC_ERR;
      }

      return new ADODB.Recordset();

      DBRecordSet w_execute_SQL_UpdateableDisconnectedRST = execute_SQL_UpdateableDisconnectedRST();
      w_execute_SQL_UpdateableDisconnectedRST.CursorLocation = adUseClient;
      w_execute_SQL_UpdateableDisconnectedRST.Open(strSQL, pconIn.getADOConn(), adOpenDynamic, adLockOptimistic, adCmdText);
      //'!TODO! Does this need a SET???
      w_execute_SQL_UpdateableDisconnectedRST.ActiveConnection = null;
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here

    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }

    return null;
    // **TODO:** label found: PROC_ERR:;
    // - - - - - - - - - - - Keep these in here!!! - - - - - - - - - - -
    Debug.Print("In "+ MCSTRNAME+ cstrCurrentProc+ ", this error was generated: "+ "\\n"+ VBA.ex.Number+ " - "+ VBA.ex.Description);
    Debug.Print("    strSQL="+ strSQL);
    // - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

    modGeneral.fnFreeObject(execute_SQL_UpdateableDisconnectedRST());

    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      //' Invalid object name
      case  -2147217865  :
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_FERR_SQL_STMT_OBJECT_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, strSQL);
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
  public boolean getMetaData_Columns(String strTableName, DBRecordSet prstInOut) { // TODO: Use of ByRef founded Public Function GetMetaData_Columns(ByVal strTableName As String, ByRef prstInOut As ADODB.Recordset) As Boolean
    boolean _rtn = false;
    //--------------------------------------------------------------------------
    // Procedure:   GetMetaData_Columns
    // Description: Use the OpenSchema method to get metadata about each
    //              table column. One inherent assumption is that the app's
    //              active and archive databases are based on an identical
    //              schema, hence we only need to ever look at the Active DB's
    //              meta data.
    // Returns:     True if successful; False otherwise
    // Params:      N/A
    // Date:        01/07/2002
    //-----------------------------------------------------------------------------
    try {
      "GetMetaData_Columns"
.equals(Const cstrCurrentProc As String);
      String strLoggedOnDB = "";

      if (prstInOut == null) {
        prstInOut = new ADODB.Recordset();
      }

      strLoggedOnDB = modGeneral.gapsApp.getActiveDatabase(modGeneral.gconAppActive.getLastLogonEnviron());

      prstInOut = modGeneral.gconAppActive.getADOConn().Execute("sp_columns '"+ strTableName+ "'");
      //2008 Update
      //Set prstInOut = gconAppActive.ADOConn.OpenSchema(adSchemaColumns, _
      //                    Array(strLoggedOnDB, Empty, strTableName, Empty))

      _rtn = true;
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Disconnect the Recordset
    if (modGeneral.fnIsObject(prstInOut)) {
      modGeneral.fnFreeObject(prstInOut.ActiveConnection);
    }

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
  public boolean getMetaData_PrimaryKeys(String strTableName, DBRecordSet prstInOut) { // TODO: Use of ByRef founded Public Function GetMetaData_PrimaryKeys(ByVal strTableName As String, ByRef prstInOut As ADODB.Recordset) As Boolean
    boolean _rtn = false;
    //--------------------------------------------------------------------------
    // Procedure:   GetMetaData_PrimaryKeys
    // Description: Use the OpenSchema method to get metadata about each
    //              table column. One inherent assumption is that the app's
    //              active and archive databases are based on an identical
    //              schema, hence we only need to ever look at the Active DB's
    //              meta data.
    // Returns:     True if successful; False otherwise
    // Params:      N/A
    // Date:        01/07/2002
    //-----------------------------------------------------------------------------
    try {
      "GetMetaData_PrimaryKeys"
.equals(Const cstrCurrentProc As String);
      String strLoggedOnDB = "";

      if (prstInOut == null) {
        prstInOut = new ADODB.Recordset();
      } 
      else if (prstInOut.State == adStateOpen) {
        prstInOut.Close;
      }

      strLoggedOnDB = modGeneral.gapsApp.getActiveDatabase(modGeneral.gconAppActive.getLastLogonEnviron());

      //SQL 2008 Update
      prstInOut = modGeneral.gconAppActive.getADOConn().Execute("SELECT column_name "+ "From INFORMATION_SCHEMA.KEY_COLUMN_USAGE "+ "WHERE OBJECTPROPERTY(OBJECT_ID(constraint_name), 'IsPrimaryKey') = 1 "+ "AND table_name = '"+ strTableName+ "'");

      //Set prstInOut = gconAppActive.ADOConn.OpenSchema(adSchemaPrimaryKeys, _
      //                    Array(strLoggedOnDB, Empty, strTableName))

      _rtn = true;
      // **TODO:** label found: PROC_EXIT:;
  //' Disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Disconnect the Recordset
    if (modGeneral.fnIsObject(prstInOut)) {
      modGeneral.fnFreeObject(prstInOut.ActiveConnection);
    }

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
  public void moveFirst(DBRecordSet prstIn) { // TODO: Use of ByRef founded Public Sub MoveFirst(ByRef prstIn As ADODB.Recordset)
    //--------------------------------------------------------------------------
    // Procedure:   MoveFirst
    // Description: Repositions the specified recordset to its first record.
    //
    //              NOTE: The caller must be diligent about checking for BOF
    //                    and/or EOF being True upon returning from this proc
    //                    rather than assuming the recordset is positioned to
    //                    a valid record (or that the .RecordCount is necessarily > 0).
    //
    // Returns:     N/A
    // Params:      prstIn (input/output) Recordset to be repositioned
    //--------------------------------------------------------------------------
    "MoveFirst"
.equals(Const cstrCurrentProc As String);

    try {

      prstIn.MoveFirst;
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
      case  3021:
        // Either BOF or EOF is True or the current record has been deleted. Requested
        // operation requires a current record.
        // Ignore these; the caller must be diligent about checking for BOF and/or EOF
        // rather than assuming the recordset is positioned to a valid record (or that
        // the .RecordCount is necessarily > 0.
        /**TODO:** resume found: Resume(Next)*/;
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
  public void moveLast(DBRecordSet prstIn) { // TODO: Use of ByRef founded Public Sub MoveLast(ByRef prstIn As ADODB.Recordset)
    //--------------------------------------------------------------------------
    // Procedure:   MoveLast
    // Description: Repositions the specified recordset to its last record
    //
    //              NOTE: The caller must be diligent about checking for BOF
    //                    and/or EOF being True upon returning from this proc
    //                    rather than assuming the recordset is positioned to
    //                    a valid record (or that the .RecordCount is necessarily > 0).
    //
    // Returns:     N/A
    // Params:      prstIn (input/output) Recordset to be repositioned
    //-----------------------------------------------------------------------------
    "MoveLast"
.equals(Const cstrCurrentProc As String);

    try {

      prstIn.MoveLast;
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
      case  3021:
        // Either BOF or EOF is True or the current record has been deleted. Requested
        // operation requires a current record.
        // Ignore these; the caller must be diligent about checking for BOF and/or EOF
        // rather than assuming the recordset is positioned to a valid record (or that
        // the .RecordCount is necessarily > 0.
        /**TODO:** resume found: Resume(Next)*/;
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
  public void moveNext(DBRecordSet prstIn) { // TODO: Use of ByRef founded Public Sub MoveNext(ByRef prstIn As ADODB.Recordset)
    //--------------------------------------------------------------------------
    // Procedure:   MoveNext
    // Description: Try to go to the next record; if doing so moves us past the
    //              last record, then back up to the last record.
    //
    //              NOTE: The caller must be diligent about checking for BOF
    //                    and/or EOF being True upon returning from this proc
    //                    rather than assuming the recordset is positioned to
    //                    a valid record (or that the .RecordCount is necessarily > 0).
    //
    // Returns:     N/A
    // Params:      prstIn (input/output) Recordset to be repositioned
    //-----------------------------------------------------------------------------
    "MoveNext"
.equals(Const cstrCurrentProc As String);

    try {

      prstIn.MoveNext;
      // Commented out -- no way to tell that we've processed all records, when looping through a rst
      //If .EOF Then
      //    .MoveLast
      //End If
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
      case  3021:
        // Either BOF or EOF is True or the current record has been deleted. Requested
        // operation requires a current record.
        // Ignore these; the caller must be diligent about checking for BOF and/or EOF
        // rather than assuming the recordset is positioned to a valid record (or that
        // the .RecordCount is necessarily > 0.
        /**TODO:** resume found: Resume(Next)*/;
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
  public void movePrev(DBRecordSet prstIn) { // TODO: Use of ByRef founded Public Sub MovePrev(ByRef prstIn As ADODB.Recordset)
    //--------------------------------------------------------------------------
    // Procedure:   MovePrev
    // Description: Try to go to the previous record; if doing so moves us past
    //              the first record, then back up to the first record.
    //
    //              NOTE: The caller must be diligent about checking for BOF
    //                    and/or EOF being True upon returning from this proc
    //                    rather than assuming the recordset is positioned to
    //                    a valid record (or that the .RecordCount is necessarily > 0).
    //
    // Returns:     N/A
    // Params:      prstIn (input/output) Recordset to be repositioned
    //-----------------------------------------------------------------------------
    "MovePrev"
.equals(Const cstrCurrentProc As String);
    try {

      prstIn.MovePrevious;
      // Commented out -- no way to tell that we've processed all records, when looping through a rst
      //If .BOF Then
      //    .MoveFirst
      //End If
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
      case  3021:
        // Either BOF or EOF is True or the current record has been deleted. Requested
        // operation requires a current record.
        // Ignore these; the caller must be diligent about checking for BOF and/or EOF
        // rather than assuming the recordset is positioned to a valid record (or that
        // the .RecordCount is necessarily > 0.
        /**TODO:** resume found: Resume(Next)*/;
        break;

      default:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }
}





///\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
//\                                                                  /
//|                        PRIVATE  Procedures                       |
///                                                                  \
//\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/

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


case class AdwadowrapperData(
              id: Option[Int],

              )

object Adwadowrappers extends Controller with ProvidesUser {

  val adwadowrapperForm = Form(
    mapping(
      "id" -> optional(number),

  )(AdwadowrapperData.apply)(AdwadowrapperData.unapply))

  implicit val adwadowrapperWrites = new Writes[Adwadowrapper] {
    def writes(adwadowrapper: Adwadowrapper) = Json.obj(
      "id" -> Json.toJson(adwadowrapper.id),
      C.ID -> Json.toJson(adwadowrapper.id),

    )
  }

  def get(id: Int) = GetAction { implicit request =>
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.LIST_ADWADOWRAPPER), { user =>
      Ok(Json.toJson(Adwadowrapper.get(user, id)))
    })
  }

  def update(id: Int) = PostAction { implicit request =>
    Logger.debug("in adwadowrappers.update")
    adwadowrapperForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      adwadowrapper => {
        Logger.debug(s"form: ${adwadowrapper.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.EDIT_ADWADOWRAPPER), { user =>
          Ok(
            Json.toJson(
              Adwadowrapper.update(user,
                Adwadowrapper(
                       id,

                ))))
        })
      }
    )
  }

  def create = PostAction { implicit request =>
    Logger.debug("in adwadowrappers.create")
    adwadowrapperForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      adwadowrapper => {
        Logger.debug(s"form: ${adwadowrapper.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.NEW_ADWADOWRAPPER), { user =>
          Ok(
            Json.toJson(
              Adwadowrapper.create(user,
                Adwadowrapper(

                ))))
        })
      }
    )
  }

  def delete(id: Int) = PostAction { implicit request =>
    Logger.debug("in adwadowrappers.delete")
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.DELETE_ADWADOWRAPPER), { user =>
      Adwadowrapper.delete(user, id)
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

case class Adwadowrapper(
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

object Adwadowrapper {

  lazy val emptyAdwadowrapper = Adwadowrapper(
)

  def apply(
      id: Int,
) = {

    new Adwadowrapper(
      id,
)
  }

  def apply(
) = {

    new Adwadowrapper(
)
  }

  private val adwadowrapperParser: RowParser[Adwadowrapper] = {
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
        Adwadowrapper(
              id,
,
              createdAt,
              updatedAt,
              updatedBy)
    }
  }

  def create(user: CompanyUser, adwadowrapper: Adwadowrapper): Adwadowrapper = {
    save(user, adwadowrapper, true)
  }

  def update(user: CompanyUser, adwadowrapper: Adwadowrapper): Adwadowrapper = {
    save(user, adwadowrapper, false)
  }

  private def save(user: CompanyUser, adwadowrapper: Adwadowrapper, isNew: Boolean): Adwadowrapper = {
    def getFields = {
      List(

      )
    }
    def throwException = {
      throw new RuntimeException(s"Error when saving ${C.ADWADOWRAPPER}")
    }

    DBHelper.saveEx(
      user,
      Register(
        C.ADWADOWRAPPER,
        C.ID,
        adwadowrapper.id,
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

  def load(user: CompanyUser, id: Int): Option[Adwadowrapper] = {
    loadWhere(user, s"${C.ID} = {id}", 'id -> id)
  }

  def loadWhere(user: CompanyUser, where: String, args : scala.Tuple2[scala.Any, anorm.ParameterValue[_]]*) = {
    DB.withConnection(user.database.database) { implicit connection =>
      SQL(s"SELECT t1.*, t2.${C.FK_NAME} FROM ${C.ADWADOWRAPPER} t1 INNER JOIN ${C.???} t2 ON t1.${C.FK_ID} = t2.${C.FK_ID} WHERE $where")
        .on(args: _*)
        .as(adwadowrapperParser.singleOpt)
    }
  }

  def delete(user: CompanyUser, id: Int) = {
    DB.withConnection(user.database.database) { implicit connection =>
      try {
        SQL(s"DELETE FROM ${C.ADWADOWRAPPER} WHERE ${C.ID} = {id}")
        .on('id -> id)
        .executeUpdate
      } catch {
        case NonFatal(e) => {
          Logger.error(s"can't delete a ${C.ADWADOWRAPPER}. ${C.ID} id: $id. Error ${e.toString}")
          throw e
        }
      }
    }
  }

  def get(user: CompanyUser, id: Int): Adwadowrapper = {
    load(user, id) match {
      case Some(p) => p
      case None => emptyAdwadowrapper
    }
  }
}


// Router

GET     /api/v1/general/adwadowrapper/:id              controllers.logged.modules.general.Adwadowrappers.get(id: Int)
POST    /api/v1/general/adwadowrapper                  controllers.logged.modules.general.Adwadowrappers.create
PUT     /api/v1/general/adwadowrapper/:id              controllers.logged.modules.general.Adwadowrappers.update(id: Int)
DELETE  /api/v1/general/adwadowrapper/:id              controllers.logged.modules.general.Adwadowrappers.delete(id: Int)




/**/
