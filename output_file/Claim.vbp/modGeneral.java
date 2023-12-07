
import java.util.Date;

public class modGeneral {

  //******************************************************************************
  // Module     : modGeneral
  // Description:
  // Procedures :
  //              fnAddBackslash(ByVal strPathIn As String) As String
  //              fnAddColumnToGrid(ByRef vfgIn As VSFlexGrid, ByVal strColumnName As String, Optional ByVal bHidden As Boolean = False)
  //              fnAreChildFormsOpen() As Boolean
  //              fnBoolToYN(ByVal bIn As Boolean) As String
  //              fnBuildQualifiedFileName(ByVal strDir As String, strFileName As String) As String
  //              fnCenterFormOnMDI(ByVal frmMDIParent As Form, ByRef frmMDIChild As Form)
  //              fnCenterFormOnScreen(ByRef frmIn As Form)
  //              fnConnectToArchiveDB(ByRef conIn As cconConnection)
  //              fnCopyFieldToRST(ByVal strColNm As String, ByRef rstSource As ADODB.Recordset, _
  //              fnCopyRSTAsUpdateable(ByVal rstSource As ADODB.Recordset) As ADODB.Recordset
  //              fnCopyRSTAsUpdateable2(ByVal rstSource As ADODB.Recordset) As ADODB.Recordset
  //              fnCstrDate(ByRef strDateIn As String, Optional ByRef strFormatIn As String = "MM/DD/YYYY") As Date
  //              fnDeallocateGlobalObjects()
  //              fnEnableDisableControl(ByVal ctlIn As Control, Optional ByVal bEnable As Boolean = True)
  //              fnFirstDayOfMonth(ByVal dteIn As Date) As Date
  //              fnFixDecimal(ByVal dblAmount As Double, ByVal intPosition As Integer, _
  //              fnFormatMMDDYYYYDate(ByVal strDateIn As String) As String
  //              fnFormatYYYYMMDDDate(ByVal strDateIn As String) As String
  //              fnFreeObject(ByRef pObj As Object)
  //              fnFreeRecordset(ByRef pRST As ADODB.Recordset)
  //              fnGetExtPart(pstrIn As String) As String
  //              fnGetStateInfo(ByVal strStateIn As String, ByVal strLOBIn As String, _
  //              fnGetStateRule(ByVal strStateIn As String, ByVal strLOBIn As String, _
  // MME START WRUS 4999
  //              fnGetStateTierInfo(ByVal strStateIn As String, ByVal strLOBIn As String, _
  // MME END WRUS 4999
  //              fnHighlightText(ctlIn As Control)
  //              fnIfNull(varValueIn As Variant, Optional varNullValueIn As Variant = "") As Variant
  //              fnInitializeAppConnectionObject()
  //              fnInitializeMenuItems()
  //              fnInitializeStateInfo(ByRef siInOut As StateInfo)
  //              fnIsFormLoaded(ByVal strFormName As String, Optional ByRef frmFound As Form) As Boolean
  //              fnLastDayOfMonth(ByVal dteIn As Date) As Date
  //              fnLimitChange(ByRef pctlIn As Control, ByRef pintMaxLen As Integer)
  //              fnLimitKeyPress(ByRef pctlIn As Control, _
  //              fnLongStateToShortState(ByVal strStateIn As String) As String
  //              fnMakeWeekday(ByVal dteIn As Date, ByVal intDirection As EnumPrevNext) As Date
  //              fnMaxDate(ByVal dte1 As Date, ByVal dte2 As Date) As Date
  //              fnMaxDouble(ByVal dblOne As Double, ByVal dblTwo As Double) As Double
  //              fnMinDate(ByVal dte1 As Date, ByVal dte2 As Date) As Date
  //              fnMinDouble(ByVal dblOne As Double, ByVal dblTwo As Double) As Double
  //              fnPadRightString(ByVal strIn As String, ByVal lngStrLen As Long, _
  //              fnPersistRecordsetToCSV(ByRef rstIn As ADODB.Recordset, ByVal strFileNm As String)
  //              fnPersistRecordsetToXML(ByRef rstIn As ADODB.Recordset, ByVal strFileNm As String)
  //              fnPhoneNumber_AddDash(ByVal strIn As String) As String
  //              fnQuoted(ByVal strIn As String) As String
  //              fnQuotedOrNull(ByVal strIn As String, Optional ByVal bRTrim As Boolean = False, _
  //              fnRemoveCloseButton(ByVal frmIn As Form)
  //              fnRemoveUnderScoresFromFieldName(ByVal strIn As String) As String
  //              fnRound(ByVal dblAmountIn As Double, ByVal intSignIn As Integer) As Double
  //              fnRoundToNextWholeDollar(ByVal dblNumber As Double) As Double
  // MME START WRUS 4999
  //              fnSelectRecord(ByVal lngKey1 As Long) As ADODB.Recordset
  // MME END WRUS 4999
  //              fnSetTopmostWindow(ByVal frm As Form, Optional ByVal bTopmost As Boolean = True)
  //              fnShortStateToLongState(ByVal strStateIn As String) As String
  //              fnShowFormsCollection()
  //              fnShowRecordPosition(rstIn As ADODB.Recordset) As String
  //              fnSSNTIN_AddDash(ByVal strIn As String, Optional bIsTin As Boolean = False) As String
  //              fnTerminateTheApp()
  //              fnTranslateToMaxValue(ByVal intDollarPositions As Integer, ByVal intDecimalPositions As Integer) As Double
  //              fnUnloadSplash()
  //              fnYNToBool(ByVal strIn As String) As Boolean
  //              Sub fnWindowLock(ByVal hWnd As Long)
  //              Sub fnWindowUnlock()
  //              TestStub_fnGetStateInfo()
  //
  // Modified   :
  // 03/26/01 BAW Cleaned with Total Visual CodeTools 2000
  // 01/2002  BAW Removed Scope-related fields & logic; added fnLongStateToShortState( )
  //              and fnShortStateToLongState( ). Also, optimized per Project Analyzer,
  //              removing dead code, adding "$" to Mid/Space, etc. Also, added the
  //              fnBuildQualifiedFileName( ) and fnPadRightString( ) procs. Also
  //              corrected a latent bug in fnIsFormLoaded( ).
  // --------------------------------------------------
//Option Explicit
  *Option Compare Binary
  private static final String MCSTRNAME = "modGeneral.";

  //=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
  //                                        CONDITIONAL COMPILER CONSTANTS
  //                                      Set to 1 to enable or 0 to disable.
  //
  // DEBUG_ERH - Shows when and by whom errors are recorded, propagated and reported.
  // DEBUG_RST - Shows how many records are in each recordset created (to determine if additional
  //             tuning is warranted).
  //=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
  *#Const DEBUG_ERH = 0
  *#Const DEBUG_RST = 0

  //=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
  //                                        DEBUGGING CONSTANTS
  //                                      Set to True to enable or False to disable.
  //
  // bDebugAppTermination - Shows information about how forms are getting unloaded and global objects
  //                        are getting deallocated  (Doesn't work right if this is defined as a
  //                        conditional compiler constant!)
  //=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
  public static final Boolean BDEBUGAPPTERMINATION = False;


  //=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
  //                                        GLOBAL VARIABLES
  //=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
  // Each of these must be set to Nothing when app ends (due to fatal error or the user's choice) !
  //' Accesses app settings stored in the registry
  public static capsAppSettings gapsApp;
  //' Global error handler
  public static cerhErrorHandler gerhApp;
  //' Use ADO Wrapper to replace DataService class (Can delete in 2C)
  public static cadwADOWrapper gadwApp;
  //' Handles ADO connection to the active app database
  public static cconConnection gconAppActive;

  // Indicates that an application-fatal error is being processed and thus the app will soon be forcibly shut down.
  // This is set by the ReportFatalError method of the cerhErrorhanlder class, but is queried by each form's
  // Form_QueryUnload to ensure ALL requests to unload will be honored without user prompts if this switch is set to True.
  public static boolean gbAmProcessingAnAppFatalError = false;

  // Indicates that the app is trying to be shut down. This could occur if an application-fatal error is being
  // processed, but also if the user chose File | Exit from the MDI screen or pressed the equivalent keystroke
  // (Alt-F4).  When this indicator is true, a form's QueryUnload event should not set pintCancel to True if
  // the user has opted to discard their pending changes.
  public static boolean gbAmTryingToTerminateTheApp = false;

  public static final String GCSTRDOUBLEQUOTE = """";
  public static final String GCSTRSINGLEQUOTE = "'";

  // gclngNoSelection is used to indicate there is no selected entry in a ComboBox, ListBox
  // or fpComboAdo control.
  private static final Long GCLNGNOSELECTION = -1;

  // gcstrBlankEntry is used in ComboBoxes used for Nullable fields, so
  // the user can select and the screen can successfully display Nulls.
  public static final String GCSTRBLANKENTRY = " ";

  // gcstrAllEntry used in combo box population
  public static final String GCSTRALLENTRY = "--All--";

  // gcstrNullEntry used in combo box population
  public static final String GCSTRNULLENTRY = "<NULL>";

  // gcintClickedCloseButton used when gerhApp.ReportNonFatal is called, to handle situations
  // where the user clicked the Close ("X") button rather than Yes, No, OK, Cancel, etc. to
  // dismiss the screen.
  public static final Integer GCINTCLICKEDCLOSEBUTTON = 0;

  private static final Integer MCINTZERO = 0;

  // The following boolean indicates whether the application log file entries should be verbose (i.e. extra
  // loggin) as well as wider (i.e. more text visible). Currently this isn't used. If/when added, add
  // code in the startup object to parse the command line to see if the /v switch (verbose mode) was
  // specified and set this boolean accordingly. See Spuds/Scuds for an example.
  //' Indicates whether log file should be terse (default) or verbose
  public static boolean gbLogVerbose = false;


  //-----------------------------------------------------------------------
  // The following defines selected columns from the State98 table. It is
  // used by the fnGetStateInfo( ) function in frmPayee.
  //-----------------------------------------------------------------------
//*TODO:** type is translated as a new class at the end of the file Public Type StateInfo
  // Fields from State98 table
  // Fields used in doing Calculation

  // The following 3 functions are used by fnWindowLock( ) and fnWindowUnlock( )
*TODO: API Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWnd As Long) As Long

  // The following 3 functions are used by fnRemoveCloseButton( )
*TODO: API Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
*TODO: API Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
*TODO: API Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long

  // The following function is used by fnSetTopmostWindow( )
*TODO: API Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


  // The following Enum is used by maintenance screens, to aid in referencing
  // the navigation buttons
//*TODO:** enum is translated as a new class at the end of the file Public Enum EnumNavigationButtons

  // The following Enum is used by the fnMakeWeekday( ) method.
//*TODO:** enum is translated as a new class at the end of the file Public Enum EnumPrevNext

  // The enumPositionDirection enum is used by the GetRelativeRecord( ) method.
//*TODO:** enum is translated as a new class at the end of the file Public Enum enumPositionDirection

  // The enumWhatOperationIsBeingAttempted enum is used by the CheckForAnotherUsersChanges() method.
//*TODO:** enum is translated as a new class at the end of the file Public Enum enumWhatOperationIsBeingAttempted

  // The following UDT is used by all table wrapper classes, to define the "standard" propererties and values
  // retained for each public property that corresponds to a column in that class' underlying SQL Server table.
//*TODO:** type is translated as a new class at the end of the file Public Type udtColumn


  *Global Const dbChar As Integer = 1
  *Global Const dbDecimal As Integer = 3
  *Global Const dbInteger As Integer = 4
  *Global Const dbDateTime As Integer = 11
  *Global Const dbVarChar As Integer = 12

  //////////////////////////////////////////////////////////////////////////////////////////
  public static String fnAddBackslash(String strPathIn) {
    String _rtn = "";
    // Add a backslash to strPathIn, if needed
    // Returns a path with a backslash
    try {
      "fnAddBackslash"
.equals(Const cstrCurrentProc As String);
      "\\"
.equals(Const cstrBackSlash As String);

      strPathIn = strPathIn.trim();

      if (strPathIn.substring(strPathIn.length() - 1) != cstrBackSlash) {
        strPathIn = strPathIn + cstrBackSlash;
      }

      _rtn = strPathIn;
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public static boolean fnAreChildFormsOpen() {
    boolean _rtn = false;
    // Comments  : Determines if any child form is open within this MDI app.
    // Parameters: N/A
    // Returns   : True, if one or more child forms are open; False otherwise
    //
    try {

      "fnAreChildFormsOpen"
.equals(Const cstrCurrentProc As String);
      Form frm = null;

      _rtn = false;

      for (int _i = 0; _i < Forms.size(); _i++) {
        frm = Forms.item(_i);
        if (!frm(Is frmMDIMain)) {
          _rtn = true;
          //Debug.Print frm.Name & "is still in Forms collection"
        }
      }
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    fnFreeObject(frm);

    if (gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}


//////////////////////////////////////////////////////////////////////////////////////////////////
  public static String fnBoolToTF(boolean bIn) {
    String _rtn = "";
    // Comments  : Translates True to "T" and False to "F"
    // Parameters: bIn (in) the boolean to translate
    //
    // Returns   : "T" or "F"
    //
    // Modified  : Berry Kropiwka 2019-11-04
    //
    // --------------------------------------------------
    try {
      "fnBoolToTF"
.equals(Const cstrCurrentProc As String);

      if (bIn) {
        _rtn = "T";
      } 
      else {
        _rtn = "F";
      }
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}

//////////////////////////////////////////////////////////////////////////////////////////////////
  public static String fnBoolToYN(boolean bIn) {
    String _rtn = "";
    // Comments  : Translates True to "Y" and False to "N"
    // Parameters: bIn (in) the boolean to translate
    //
    // Returns   : "Y" or "N"
    //
    // Modified  :
    //
    // --------------------------------------------------
    try {
      "fnBoolToYN"
.equals(Const cstrCurrentProc As String);

      if (bIn) {
        _rtn = "Y";
      } 
      else {
        _rtn = "N";
      }
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}


//////////////////////////////////////////////////////////////////////////////////////////////////
  public static String fnBuildQualifiedFileName(String strDir, String strFileName) {
    String _rtn = "";
    // Comments  : Returns the fully qualified filename, by joining the strDir and
    //             and strFile parameters...with a slash if appropriate
    // Parameters: strDir - fully qualified folder name
    //             strFileName - file name
    //
    // Called By : fnLogOpen( ) of modAppLog
    //             fnLogPrune( ) of modAppLog
    //
    // Modified  :
    //  01/2002 BAW  Copied from SPUDS/SCUDS
    // --------------------------------------------------
    try {
      "fnBuildQualifiedFileName"
.equals(Const cstrCurrentProc As String);
      Scripting.FileSystemObject fso = null;

      fso = new Scripting.FileSystemObject();
      _rtn = fso.BuildPath(strDir, strFileName);
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    fnFreeObject(fso);

    if (gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}


//////////////////////////////////////////////////////////////////////////////////////////////////
  public static void fnCenterFormOnMDI(Form frmMDIParent, Form frmMDIChild) { // TODO: Use of ByRef founded Public Sub fnCenterFormOnMDI(ByVal frmMDIParent As Form, ByRef frmMDIChild As Form)
    // Comments  : Centers the form on the MDI parent
    // Parameters: none
    // Returns   : Nothing
    // Source    : www.vbexplorer.com
    //
    "fnCenterFormOnMDI"
.equals(Const cstrCurrentProc As String);
    int intTop = 0;
    int intLeft = 0;
    try {

      if (frmMDIParent.WindowState == vbNormal) {
        intTop = ((frmMDIParent.Height - frmMDIChild.Height) \ 2);
        intLeft = ((frmMDIParent.Width - frmMDIChild.Width) \ 2);
        frmMDIChild.Move(intLeft, intTop);
      }
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public static void fnCenterFormOnScreen(Form frmIn) { // TODO: Use of ByRef founded Public Sub fnCenterFormOnScreen(ByRef frmIn As Form)
    // Comments  : Centers the form on the screen
    // Parameters: none
    // Returns   : Nothing
    // Source    : Total Visual SourceBook 2000
    //
    "fnCenterFormOnScreen"
.equals(Const cstrCurrentProc As String);
    try {

      frmIn.Move(Screen.Width - frmIn.Width) / 2, (Screen.Height - frmIn.Height) / 2;
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }
}


//////////////////////////////////////////////////////////////////////////////////////////////////
  public static void fnDeallocateGlobalObjects() {
    // Comments  : This procedure releases the memory allocated for global objects. It should
    //             be called before ANY application termination (whether due to a fatal error
    //             or the user's inititiation.
    //             IMPORTANT: The Global Error Handler should be the last object deallocated,
    //                        hence it is deallocated when the MDI Main form is unloaded.
    //
    //             NOTE:  This procedure (and possibly sub Main in modStartup.bas)
    //                    should be updated as global object variables are added to
    //                    or removed from the application!
    //
    // Parameters: none
    // Returns   : Nothing
    //
    // Called by : fnTerminateTheApp() of modGeneral.bas   (user-initiated app termination)
    //             ReportFatalError() of cerhErrorHandler.bas   (fatal error)
    //
    // Source    : Total Visual SourceBook 2000
    //
    // Modified  :
    // 04/30/02 BAW (Phase 2B, but 2B004) Made the Crystal Application object a global variable:
    //              defined in modReporting; instantiated in modStartup; deallocated in
    //              fnDeallocateGlobalObjects. This avoids "Out of memory" errors
    //              when the frmReportViewer screen is displayed.
    //
    "fnDeallocateGlobalObjects"
.equals(Const cstrCurrentProc As String);
    try {

      if (BDEBUGAPPTERMINATION) {
        Debug.Print("   Freeing gadwApp from fnDeallocateGlobalObjects");
      }
      fnFreeObject(gadwApp);

      if (BDEBUGAPPTERMINATION) {
        Debug.Print("   Freeing gconAppActive from fnDeallocateGlobalObjects");
      }
      fnFreeObject(gconAppActive);

      if (BDEBUGAPPTERMINATION) {
        Debug.Print("   Freeing gapsApp from fnDeallocateGlobalObjects");
      }
      fnFreeObject(gapsApp);

      if (BDEBUGAPPTERMINATION) {
        Debug.Print("   Freeing gcrxApp from fnDeallocateGlobalObjects");
      }
      fnFreeObject(modReporting.gcrxApp);
    } catch (Exception ex) {
    //' disable error handler
    }
    try {
      // Clean-up statements go here
      if (gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
        gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
      }
      return;
      // **TODO:** label found: PROC_ERR:;
      switch (VBA.ex.Number) {
          //Case statements for expected errors go here
        case  Else:
          gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
          break;
      }
      /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public static void fnEnableDisableControl(Control ctlIn, boolean bEnable) {
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
          // Temporary turn off error handling, since some controls do not support the .Locked property
      }
      //*TODO:** the error label PROC_ERR: couldn't be found
        try {
        ctlIn.Locked = false;
    }
    try {
      ctlIn.TabStop = true;
      ctlIn.BackColor = vbWindowBackground;
      ctlIn.ForeColor = vbWindowText;
      ctlIn.Enabled = true;
      break;

    case  False:
      // Temporary turn off error handling, since some controls do not support the .Locked property
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    ctlIn.Locked = true;
}
try {
  ctlIn.TabStop = false;
  ctlIn.BackColor = vbButtonFace;
  ctlIn.ForeColor = vbButtonText;
  ctlIn.Enabled = false;
      break;
  }
  // **TODO:** label found: PROC_EXIT:;
  // Disable the error handler so errors hit here won't be handled by PROC_ERR
}
//*TODO:** the error label PROC_ERR: couldn't be found
  try {
  // Clean-up statements go here
  if (gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
    gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
  }
  return;
  // **TODO:** label found: PROC_ERR:;
  switch (VBA.ex.Number) {
      //Case statements for expected errors go here
    case  Else:
      gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
      break;
  }
  /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }
}


//////////////////////////////////////////////////////////////////////////////////////////////////
  public static Date fnFirstDayOfMonth(Date dteIn) {
    Date _rtn = null;
    // Comments  : Calculates the first day of the month for the specified date.
    // Parameters: dteIn - Date for which first DOM will be determined
    //
    // Returns   : The date representing the first day of that month
    // Source    : <http://www.vb-world.net/misc/tip479.html>
    // Modified  :
    //
    // --------------------------------------------------
    "fnFirstDayOfMonth"
.equals(Const cstrCurrentProc As String);
    int intMonth = 0;
    int intYear = 0;
    try {

      intMonth = Month(dteIn);
      intYear = Year(dteIn);

      _rtn = DateSerial(intYear, intMonth, 1);
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here

    if (gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}


//////////////////////////////////////////////////////////////////////////////////////////////////
  public static void fnFreeObject(Object pObj) { // TODO: Use of ByRef founded Public Sub fnFreeObject(ByRef pObj As Object)
    // Comments  : Safely frees memory used by an object
    //
    //             NOTE: Use fnFreeRecordset( ) for objects
    //                   of type "ADODB.Recordset" since this
    //                   will ensure its DBMS resources are
    //                   released.
    //
    // Parameters: pObj (in/out) - pointer to the object to free
    //
    // Called by : Lots of places (usually in the PROC_EXIT block)
    //
    // Returns   : N/A
    //
    // Modified  :
    //
    // --------------------------------------------------
    if (fnIsObject(pObj)) {
      pObj = null;
    }
  }



  //////////////////////////////////////////////////////////////////////////////////////////////////
  public static void fnFreeRecordset(DBRecordSet pRST) { // TODO: Use of ByRef founded Public Sub fnFreeRecordset(ByRef pRST As ADODB.Recordset)
    // Comments  : Safely frees memory used by an ADODB.Recordset
    //             object after first ensuring it is closed.
    //
    // Parameters: pRST (in/out) - pointer to the object to free
    //
    // Called by : Lots of places (usually in the PROC_EXIT block)
    //
    // Returns   : N/A
    //
    // Modified  :
    //
    // --------------------------------------------------
    if (fnIsObject(pRST)) {
      if (pRST.State == adStateOpen) {
        //' Guard against 3219 "Operation not allowed in this context" error
        try {
          pRST.Close;
      }
      try {
      }
      pRST = null;
    }
//*TODO:** the error label 0: couldn't be found
  }
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public static String fnGetExtPart(String pstrIn) {
    String _rtn = "";
    // Comments  : Returns the extension of a fully qualified file name
    // Parameters: strIn - path and name to parse
    // Returns   : file extension
    // Source    : Shamelessly plagurized from Total Visual SourceBook 2000
    // --------------------------------------------------
    try {
      "fnGetExtPart"
.equals(Const cstrCurrentProc As String);

      int intCounter = 0;
      String strTmp = "";

      // Parse the string
      for (intCounter = pstrIn.length(); intCounter >= 1; intCounter--) {
        // It its a slash, grab the sub string
        if (!(pstrIn.substring(intCounter, 1).equals("."))) {
          strTmp = pstrIn.substring(intCounter, 1)+ strTmp;
        } 
        else {
          break;
        }
      }

      // Return the value
      _rtn = strTmp.toUpperCase();
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}


//////////////////////////////////////////////////////////////////////////////////////////////////

//MME START - ADDED InsuredClmID and PayeDthbPmtAmt to paramater list

  public static void fnGetStateInfo(String strStateIn, String strLOBIn, Date dtePmtDt, int insuredClmID, int payeDthbPmtAmt, StateInfo siInOut) { // TODO: Use of ByRef founded Public Sub fnGetStateInfo(ByVal strStateIn As String, ByVal strLOBIn As String, ByVal dtePmtDt As Date, ByVal InsuredClmID As Long, ByVal PayeDthbPmtAmt As Long, ByRef siInOut As StateInfo)
    // Comments  : Retrieves info from the State98 table and returns data
    //             from the selected row in a UDT called StateInfo
    // Parameters: strWhereIn (in) - the WHERE clause for the SQL query that selects a
    //                               particular row, e.g., "[State] = 'Alabama'"
    //             siIn (in/out)   - a StateInfo (UDT) structure that will hold the contents of
    //                               the selected row
    // Modified  :
    // --------------------------------------------------
    try {
      "fnGetStateInfo"
.equals(Const cstrCurrentProc As String);
      "G"
.equals(Const mcstrGroupLOB As String);
      "I"
.equals(Const mcstrIndividualLOB As String);
      DBRecordSet rstTemp = null;


      //MME START - WRUS 4999

      "H"
.equals(Const mcstrGroupTier2LOB As String);
      "J"
.equals(Const mcstrIndividualTier2LOB As String);
      "PROOFDTH"
.equals(Const mcstrPROOFDEATH As String);
      "PROOF"
.equals(Const mcstrPROOF As String);
      "DEATH"
.equals(Const mcstrDEATH As String);
      "AMOUNT"
.equals(Const mcstrAMOUNT As String);
      String mcstrStrltIdtypCd = "";
      DBRecordSet rstTierTemp = null;
      Date dtResultDate = null;
      Date dtProofDate = null;
      Date dtDeathDate = null;
      double dblAmount = 0;
      double dblCompareVal = 0;
      int dtDthProofDifference = 0;
      DBRecordSet rstSingleRecord_Fresh = null;


      //fnLogWrite "      In fnGetStateInfo, getting State Tier info: " & strStateIn & " " & strLOBIn & " " & CStr(DateValue(dtePmtDt)), cstrCurrentProc

      rstTierTemp = fnGetStateTierInfo(strStateIn, strLOBIn, dtePmtDt);

      if (rstTierTemp.RecordCount != 0) {

        rstSingleRecord_Fresh = fnSelectRecord(insuredClmID);

        if (rstSingleRecord_Fresh.RecordCount != 0) {

          mcstrStrltIdtypCd = rstTierTemp.Fields(3).trim();

          dtProofDate = rstSingleRecord_Fresh.Fields(8).value;
          dtDeathDate = rstSingleRecord_Fresh.Fields(2).value;
          dblAmount = rstSingleRecord_Fresh.Fields(10).value;

          if (rstTierTemp.Fields(4) < 0) {
            dblCompareVal = rstTierTemp.Fields(4) * -1;
          } 
          else {
            dblCompareVal = rstTierTemp.Fields(4);
          }

          switch (mcstrStrltIdtypCd) {

            case  mcstrPROOFDEATH:
              dtDthProofDifference = DateDiff("d", dtDeathDate, dtProofDate);
              //'IF NEGATIVE VALUE, THEN REVERSE E.G. WITHIN TIMEFRAME, THEN TIER2
              if (rstTierTemp.Fields(4) < 0) {
                //'DtDthProofDifference is difference in days
                if (dtDthProofDifference <= dblCompareVal) {
                  if (strLOBIn.equals(mcstrGroupLOB)) {
                    strLOBIn = mcstrGroupTier2LOB;
                  } 
                  else {
                    strLOBIn = mcstrIndividualTier2LOB;
                  }
                }
              } 
              else {
                //'DtDthProofDifference is difference in days
                if (dtDthProofDifference > dblCompareVal) {
                  if (strLOBIn.equals(mcstrGroupLOB)) {
                    strLOBIn = mcstrGroupTier2LOB;
                  } 
                  else {
                    strLOBIn = mcstrIndividualTier2LOB;
                  }
                }
              }

              break;

            case  mcstrPROOF:
              dtResultDate = (DateAdd("d", dblCompareVal, dtProofDate));
              //'IF NEGATIVE VALUE, THEN REVERSE E.G. WITHIN TIMEFRAME, THEN TIER2
              if (rstTierTemp.Fields(4) < 0) {
                //'DtResultDate is cutoff date
                if (DateValue(dtePmtDt) <= DateValue(dtResultDate)) {
                  if (strLOBIn.equals(mcstrGroupLOB)) {
                    strLOBIn = mcstrGroupTier2LOB;
                  } 
                  else {
                    strLOBIn = mcstrIndividualTier2LOB;
                  }
                }
              } 
              else {
                //'NOT WITHIN TIMEFRAME, THEN TIER2
                if (DateValue(dtePmtDt) > DateValue(dtResultDate)) {
                  if (strLOBIn.equals(mcstrGroupLOB)) {
                    strLOBIn = mcstrGroupTier2LOB;
                  } 
                  else {
                    strLOBIn = mcstrIndividualTier2LOB;
                  }
                }
              }

              break;

            case  mcstrDEATH:
              dtResultDate = (DateAdd("d", dblCompareVal, dtDeathDate));
              //'IF NEGATIVE VALUE, THEN REVERSE E.G. WITHIN TIMEFRAME, THEN TIER2
              if (rstTierTemp.Fields(4) < 0) {
                //'DtResultDate is cutoff date
                if (DateValue(dtePmtDt) <= DateValue(dtResultDate)) {
                  if (strLOBIn.equals(mcstrGroupLOB)) {
                    strLOBIn = mcstrGroupTier2LOB;
                  } 
                  else {
                    strLOBIn = mcstrIndividualTier2LOB;
                  }
                }
              } 
              else {
                //'NOT WITHIN TIMEFRAME, THEN TIER2
                if (DateValue(dtePmtDt) > DateValue(dtResultDate)) {
                  if (strLOBIn.equals(mcstrGroupLOB)) {
                    strLOBIn = mcstrGroupTier2LOB;
                  } 
                  else {
                    strLOBIn = mcstrIndividualTier2LOB;
                  }
                }
              }

              break;

            case  mcstrAMOUNT:
              //'IF NEGATIVE VALUE, THEN REVERSE E.G. WITHIN TOLERANCE, THEN TIER2
              if (rstTierTemp.Fields(4) < 0) {
                if (payeDthbPmtAmt > dblCompareVal) {
                  if (strLOBIn.equals(mcstrGroupLOB)) {
                    strLOBIn = mcstrGroupTier2LOB;
                  } 
                  else {
                    strLOBIn = mcstrIndividualTier2LOB;
                  }
                }
              } 
              else {
                //'NOT WITHIN TOLERANCE, THEN TIER2
                if (payeDthbPmtAmt <= dblCompareVal) {
                  if (strLOBIn.equals(mcstrGroupLOB)) {
                    strLOBIn = mcstrGroupTier2LOB;
                  } 
                  else {
                    strLOBIn = mcstrIndividualTier2LOB;
                  }
                }
              }

              break;

            default:
              // Invalid record found on table STATE_RULE_TIER_T (4012) -
              // for the state of @@1 as of @@2. The calculations cannot be done.
              gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_INVALID_ENTRY_RULE_TIER_T, MCSTRNAME+ cstrCurrentProc, strStateIn, CStr(DateValue(dtePmtDt)));
              // **TODO:** goto found: GoTo PROC_EXIT;
              break;
          }
        } 
        else {
          if ((rstSingleRecord_Fresh == null) || (rstSingleRecord_Fresh.RecordCount == 0)) {
            // Claim has been deleted by another user -
            gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_INFO_ANOTHER_USER_DELETED_NOT_SAVED, MCSTRNAME+ cstrCurrentProc, insuredClmID);
            // **TODO:** goto found: GoTo PROC_EXIT;
          }
        }
      }

      //MME END - WRUS 4999


      //fnLogWrite "      In fnGetStateInfo, getting new rule: " & strStateIn & " " & strLOBIn & " " & CStr(DateValue(dtePmtDt)), cstrCurrentProc

      rstTemp = fnGetStateRule(strStateIn, strLOBIn, dtePmtDt);

      // If we didn't find such a row:
      //    * if we had been looking for a Group LOB (a common situation), then
      //      try for the row that has an Individual LOB (I). There should be one!
      //    * if we had been looking for an Individual LOB, then something is
      //      very wrong. Every state should have an Individual row, but there
      //      will most likely only be Group rows for a small handful of states
      //      (like Georgia).
      // Group is supposed to default to using Individual rates if no
      // Group-specific rates are defined for a given state.
      if (rstTemp.RecordCount == 0) {
        if (strLOBIn.equals(mcstrGroupLOB)) {
          rstTemp = fnGetStateRule(strStateIn, mcstrIndividualLOB, dtePmtDt);
          if (rstTemp.RecordCount == 0) {

            // gcRES_NERR_STATE_RATES_NOT_FOUND (4006) - Neither Group nor Individual
            // rates were found for the state of @@1 as of @@2. The calculations cannot be done.
            gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_STATE_RATES_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, strStateIn, CStr(DateValue(dtePmtDt)));
            // **TODO:** goto found: GoTo PROC_EXIT;
          }
        } 
        else {
          // gcRES_NERR_INDV_STATE_RATES_NOT_FOUND (4007) - Individual rates were not found
          // for the state of @@1 as of @@2. The calculations cannot be done.
          gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_INDV_STATE_RATES_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, strStateIn, CStr(DateValue(dtePmtDt)));
          // **TODO:** goto found: GoTo PROC_EXIT;
        }
      }

      siInOut.stCd = !st_cd;
      siInOut.lobCd = !lob_cd;
      siInOut.strlEffDt = !strl_eff_dt;
      siInOut.calcIdtypCd = !calc_idtyp_cd;
      siInOut.reqdIdtypCd = !reqd_idtyp_cd;
      siInOut.iruleCd = !irule_cd;
      siInOut.strlEndDt = !strl_end_dt;
      siInOut.strlIntRptgFlrAmt = !strl_int_rptg_flr_amt;
      siInOut.strlIntCalcOfstNum = !strl_int_calc_ofst_num;
      siInOut.strlIntReqdOfstNum = !strl_int_reqd_ofst_num;
      siInOut.strlIntRuleAmt = !strl_int_rule_amt;
      siInOut.strlSpclInstrTxt = !strl_spcl_instr_txt;
      //fnLogWrite "      In fnGetStateInfo, got: " & !st_cd & " " & !lob_cd & " " & CStr(DateValue(!strl_eff_dt)), cstrCurrentProc
      rstTemp.Close;
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    fnFreeRecordset(rstTemp);

    if (gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }
}


// MME START WRUS 4999

//////////////////////////////////////////////////////////////////////////////////////////////////
  public static DBRecordSet fnGetStateTierInfo(String strStateIn, String strLOBIn, Date dtePmtDt) {
    // Comments  : Retrieves info from the State_rule_tier_t and returns a value
    //             based on cals performed against the data on the selected row.
    // Parameters: strWhereIn (in) - the WHERE clause for the SQL query that selects a
    //                               particular row, e.g., "[State] = 'Alabama'"
    //             siIn (in/out)   - a String that will hold pass back a value of 'G', 'H', 'I', or 'J'
    // Modified  :
    // --------------------------------------------------
    try {
      "fnGetStateTierInfo"
.equals(Const cstrCurrentProc As String);
      //' Stored procedure to execute
      "dbo.proc_state_rule_tier_t"
.equals(Const cstrSproc As String);

      cadwADOWrapper adwTemp = null;
      DBRecordSet rstTemp = null;
      ADODB.Parameter prmReturnValue = null;
      ADODB.Parameter prmStCd = null;
      ADODB.Parameter prmLobCd = null;
      ADODB.Parameter prmPayePmtDt = null;

      //fnLogWrite "      In fnGetStateTierInfo, getting new rule: " & strStateIn & " " & strLOBIn & " " & CStr(DateValue(dtePmtDt)), cstrCurrentProc


      adwTemp = new cadwADOWrapper();

      if (!(adwTemp.commandSetSproc(cstrSproc))) {
        // **TODO:** goto found: GoTo PROC_EXIT;
      }

      ADODB.Command w_aDOCommand = adwTemp.getADOCommand();
      // ---Parameter #1---
      // Define the return value that represents the error code (i.e. reason) why
      // the stored procedure failed.
      prmReturnValue = w_aDOCommand.CreateParameter(Name:="@return_value", Type:=adInteger, Direction:=adParamReturnValue, .value:=Null);
      w_aDOCommand.Parameters.Append(prmReturnValue);

      // ---Parameter #2---
      // Define the LOB_CD parameter
      prmLobCd = w_aDOCommand.CreateParameter(Name:="@lob_cd", Type:=adChar, Direction:=adParamInput, Size:=1, .value:=strLOBIn);
      w_aDOCommand.Parameters.Append(prmLobCd);

      // ---Parameter #3---
      // Define the ST_CD parameter
      prmStCd = w_aDOCommand.CreateParameter(Name:="@st_cd", Type:=adChar, Direction:=adParamInput, Size:=2, .value:=strStateIn);
      w_aDOCommand.Parameters.Append(prmStCd);

      // ---Parameter #4---
      // Define the PAYE_PMT_DT parameter
      prmPayePmtDt = w_aDOCommand.CreateParameter(Name:="@paye_pmt_dt", Type:=adDBTimeStamp, Direction:=adParamInput, Size:=8, .value:=dtePmtDt);
      w_aDOCommand.Parameters.Append(prmPayePmtDt);

      rstTemp = w_aDOCommand.Execute();
      rstTemp.ActiveConnection = null;

      // The rstTemp recordset may well be empty. That's okay though since the caller
      // (fnGetStateRule) can accomodate this...either by looking for a different LOB's row
      // or by generating an error of its own accord.

      return rstTemp;
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here

    // DO NOT do "fnFreeRecordset rstTemp" since this will cause the recordset
    // returned by this function to be wiped out as well!
    fnFreeObject(prmStCd);
    fnFreeObject(prmLobCd);
    fnFreeObject(prmPayePmtDt);
    fnFreeObject(prmReturnValue);

    if (gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return null;
    // **TODO:** label found: PROC_ERR:;
    switch (prmReturnValue) {
      //' 4027 -  The @@1 is invalid. @@2
      case  modResConstants.gCRES_NERR_REC_NOT_FOUND      :
        // Note that the following error is presented as an ATYPICAL 4027 error!
        gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_INVALID_DATA, MCSTRNAME+ cstrCurrentProc, "State, Line of Business, or Date of Payment", "The Calculation Rule cannot be retrieved "+ "when any of these fields are NULL or if the State Code cannot be found in the State table. It may also "+ "be that no Calculation Rule is in effect "+ "for the State for the given Date of Payment. State=["+ strStateIn+ "], LOB=["+ strLOBIn+ "], Payment Date=["+ FormatDateTime(dtePmtDt, vbShortDate)+ "]");
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
      //' 4028
        break;

      case  modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO  :
        gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO, MCSTRNAME+ cstrCurrentProc, "locate");
        //'
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
        break;
    }

    // If any other errors exist, i.e. in Err object, then let it fall through into default error handling.

    switch (VBA.ex.Number) {
      //' Object not found
      case  -2147217900:
        gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_FERR_SPROC_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, cstrSproc);
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
        break;

      default:
        gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }

    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }
}

// MME END WRUS 4999


//////////////////////////////////////////////////////////////////////////////////////////////////
  public static DBRecordSet fnGetStateRule(String strStateIn, String strLOBIn, Date dtePayePmtDt) {
    // Comments  : Retrieves info from the STATE_RULE_T and returns data
    //             from the selected row in
    // Parameters: strStateIn (in)   - the desired state code
    //             strLOBIn (in)     - the desired line-of-business
    //             dtePayePmtDt (in) - the date as of which to retrieve the state rule info
    // Returns   : An ADODB.Recordset
    // Modified  :
    // --------------------------------------------------
    try {
      "fnGetStateRule"
.equals(Const cstrCurrentProc As String);
      //' Stored procedure to execute
      "dbo.proc_state_rule_select"
.equals(Const cstrSproc As String);
      cadwADOWrapper adwTemp = null;
      DBRecordSet rstTemp = null;
      ADODB.Parameter prmReturnValue = null;
      ADODB.Parameter prmStCd = null;
      ADODB.Parameter prmLobCd = null;
      ADODB.Parameter prmPayePmtDt = null;

      adwTemp = new cadwADOWrapper();

      if (!(adwTemp.commandSetSproc(cstrSproc))) {
        // **TODO:** goto found: GoTo PROC_EXIT;
      }

      ADODB.Command w_aDOCommand = adwTemp.getADOCommand();
      // ---Parameter #1---
      // Define the return value that represents the error code (i.e. reason) why
      // the stored procedure failed.
      prmReturnValue = w_aDOCommand.CreateParameter(Name:="@return_value", Type:=adInteger, Direction:=adParamReturnValue, .value:=Null);
      w_aDOCommand.Parameters.Append(prmReturnValue);

      // ---Parameter #2---
      // Define the LOB_CD parameter
      prmLobCd = w_aDOCommand.CreateParameter(Name:="@lob_cd", Type:=adChar, Direction:=adParamInput, Size:=1, .value:=strLOBIn);
      w_aDOCommand.Parameters.Append(prmLobCd);

      // ---Parameter #3---
      // Define the ST_CD parameter
      prmStCd = w_aDOCommand.CreateParameter(Name:="@st_cd", Type:=adChar, Direction:=adParamInput, Size:=2, .value:=strStateIn);
      w_aDOCommand.Parameters.Append(prmStCd);

      // ---Parameter #4---
      // Define the PAYE_PMT_DT parameter
      prmPayePmtDt = w_aDOCommand.CreateParameter(Name:="@paye_pmt_dt", Type:=adDBTimeStamp, Direction:=adParamInput, Size:=8, .value:=dtePayePmtDt);
      w_aDOCommand.Parameters.Append(prmPayePmtDt);

      rstTemp = w_aDOCommand.Execute();
      rstTemp.ActiveConnection = null;

      // The rstTemp recordset may well be empty. That's okay though since the caller
      // (fnGetStateRule) can accomodate this...either by looking for a different LOB's row
      // or by generating an error of its own accord.

      return rstTemp;
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here

    // DO NOT do "fnFreeRecordset rstTemp" since this will cause the recordset
    // returned by this function to be wiped out as well!
    fnFreeObject(prmStCd);
    fnFreeObject(prmLobCd);
    fnFreeObject(prmPayePmtDt);
    fnFreeObject(prmReturnValue);

    if (gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return null;
    // **TODO:** label found: PROC_ERR:;
    switch (prmReturnValue) {
      //' 4027 -  The @@1 is invalid. @@2
      case  modResConstants.gCRES_NERR_REC_NOT_FOUND      :
        // Note that the following error is presented as an ATYPICAL 4027 error!
        gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_INVALID_DATA, MCSTRNAME+ cstrCurrentProc, "State, Line of Business, or Date of Payment", "The Calculation Rule cannot be retrieved "+ "when any of these fields are NULL or if the State Code cannot be found in the State table. It may also "+ "be that no Calculation Rule is in effect "+ "for the State for the given Date of Payment. State=["+ strStateIn+ "], LOB=["+ strLOBIn+ "], Payment Date=["+ FormatDateTime(dtePayePmtDt, vbShortDate)+ "]");
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
      //' 4028
        break;

      case  modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO  :
        gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO, MCSTRNAME+ cstrCurrentProc, "locate");
        //'
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
        break;
    }

    // If any other errors exist, i.e. in Err object, then let it fall through into default error handling.

    switch (VBA.ex.Number) {
      //' Object not found
      case  -2147217900:
        gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_FERR_SPROC_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, cstrSproc);
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
        break;

      default:
        gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }

    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }
}


//////////////////////////////////////////////////////////////////////////////////////////////////
  public static void fnInitializeStateInfo(StateInfo siInOut) { // TODO: Use of ByRef founded Public Sub fnInitializeStateInfo(ByRef siInOut As StateInfo)
    // Comments  : Initializes the specified UDT called StateInfo
    // Parameters:
    //       siIn (in/out)   - a StateInfo (UDT) structure that will hold the contents of
    //                         the selected row
    // Modified  :
    // --------------------------------------------------
    try {
      "fnInitializeStateInfo"
.equals(Const cstrCurrentProc As String);

      siInOut.stCd = GCSTRBLANKENTRY;
      siInOut.strlEffDt = G.parseDate(Now);
      siInOut.calcIdtypCd = "";
      siInOut.reqdIdtypCd = "";
      siInOut.iruleCd = "";
      siInOut.strlEndDt = vbNull;
      siInOut.strlIntRptgFlrAmt = MCINTZERO;
      siInOut.strlIntCalcOfstNum = MCINTZERO;
      siInOut.strlIntReqdOfstNum = MCINTZERO;
      siInOut.strlIntRuleAmt = MCINTZERO;
      siInOut.strlSpclInstrTxt = "";
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here

    if (gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public static boolean fnIsFormLoaded(String strFormName, Form frmFound) { // TODO: Use of ByRef founded Public Function fnIsFormLoaded(ByVal strFormName As String, Optional ByRef frmFound As Form) As Boolean
    boolean _rtn = false;
    //----------------------------------------------------------------------------
    // Comments  : Tests to see whether the default instance
    //             of a form is loaded
    // Parameters: strFormName (in) - name of form to search for
    //             frmFound (out)   - pointer to the searched-for form, if found
    //
    // Returns   : True if the form is loaded, false otherwise
    // Source    : Total Visual SourceBook 2000
    //----------------------------------------------------------------------------
    Form frm = null;
    boolean bResult = false;
    "fnIsFormLoaded"
.equals(Const cstrCurrentProc As String);

    try {

      // If a form is loaded, it will be in the Forms collection.
      // Search this collection to see if the specified form
      // is present.
      for (int _i = 0; _i < Forms.size(); _i++) {
        frm = Forms.item(_i);
        if (frm.Name.toUpperCase().equals(strFormName.toUpperCase())) {
          bResult = true;
          frmFound = frm;
          break;
        }
      }

      _rtn = bResult;
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    fnFreeObject(frm);

    if (gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}


//////////////////////////////////////////////////////////////////////////////////////////////////
  public static boolean fnIsObject(Object objIn) {
    //----------------------------------------------------------------------------
    // Comments  : Safe test to see if object exists (better than "If IsObject()"
    // Parameters: objIn (in) - Object reference
    //
    // Returns   : True if the object has been initialized, false otherwise
    //----------------------------------------------------------------------------
    "fnIsObject"
.equals(Const cstrCurrentProc As String);

    return !(objIn == null);
  }



  //////////////////////////////////////////////////////////////////////////////////////////////////
  public static Date fnLastDayOfMonth(Date dteIn) {
    Date _rtn = null;
    // Comments  : Calculates the last day of the month for the specified date.
    // Parameters: dteIn - Date for which last DOM will be determined
    //
    // Returns   : The date representing the last day of that month
    // Source    : <http://www.vb-world.net/misc/tip479.html>
    // Modified  :
    //
    // --------------------------------------------------
    "fnLastDayOfMonth"
.equals(Const cstrCurrentProc As String);
    int intLastDay = 0;

    try {

      intLastDay = DatePart("d", DateAdd("d", -1, DateAdd("m", 1, DateAdd("d", -DatePart("d", dteIn) + 1, dteIn))));

      _rtn = DateSerial(Year(dteIn), Month(dteIn), intLastDay);
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here

    if (gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public static Date fnMakeWeekday(Date dteIn, EnumPrevNext intDirection) {
    Date _rtn = null;
    // Comments  : Returns the input date, coerced to the next or previous weekday (based
    //             on the intDirection paramter) if the input date fell on a weekend.
    //             Date returned is always between Monday and Friday.
    // Parameters: dteIn        (in) - Date to coerce
    //             intDirection (in) - indicates whether to move to the next or previous weekday
    // Returns   : Coerced date
    // Source    : Based on Total Visual SourceBook 2000's PriorWeekday and NextWeekday functions
    //
    "fnMakeWeekday"
.equals(Const cstrCurrentProc As String);
    Date dteTemp = null;

    try {

      if (Weekday(dteIn) == vbSaturday  || Weekday(dteIn) == vbSunday) {
        switch (intDirection) {
          case  EnumPrevNext.ePNPREV:
            dteTemp = dteIn - 1;
            while (Weekday(dteTemp) == vbSunday  || Weekday(dteTemp) == vbSaturday) {
              dteTemp = dteTemp - 1;
            }
            break;

          case  EnumPrevNext.ePNNEXT:
            dteTemp = dteIn + 1;
            while (Weekday(dteTemp) == vbSunday  || Weekday(dteTemp) == vbSaturday) {
              dteTemp = dteTemp + 1;
            }
            break;
        }
        _rtn = dteTemp;
      } 
      else {
        _rtn = dteIn;
      }
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here

    if (gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}


////////////////////////////////////////////////////////////////////////////////////////////
  public static boolean fnOpenFileInDefaultApp(String strFile) {
    boolean _rtn = false;
    // Comments  : Opens the specified file in the application that is
    //             associated with that kind of file.
    // Parameters: strFile - the fully-qualified file to open
    // Called by : mnuHelpViewApplicationLogFile_Click() of frmMDIMain
    // Modified  :
    //
    // --------------------------------------------------
    try {
      "fnOpenFileInDefaultApp"
.equals(Const cstrCurrentProc As String);
      int lngReturnCode = 0;

      lngReturnCode = ShellExecute(0&, "open", strFile, "", "", vbNormalFocus);

      _rtn = (lngReturnCode > 32);
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here

    if (gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        // gcRES_INFO_CANT_OPEN_FILE (1014)
        // Unable to open <@@1>. The file either does not exist or no application is associated with files of type .TXT.
        gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_INFO_CANT_OPEN_FILE, MCSTRNAME+ cstrCurrentProc, strFile);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//!TODO! Make version that pads with left with leading zeroes
//////////////////////////////////////////////////////////////////////////////////////////////////
  public static String fnPadRightString(String strIn, int lngStrLen, String strPadCharIn) {
    String _rtn = "";
    // Comments  : Right pads a string for left justification
    // Parameters: strIn - String to pad
    //             strPadCharIn - Character to use for padding
    //             lngStrLen - Desired length of string
    //
    // Returns   : right padded string
    // Source    : Total Visual SourceBook 2000
    // Modified  :
    //
    // --------------------------------------------------
    try {
      "fnPadRightString"
.equals(Const cstrCurrentProc As String);

      _rtn = (strIn+ String$(lngStrLen, strPadCharIn.substring(0, 1))).substring(0, lngStrLen);
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here

    if (gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public static void fnPersistRecordsetToCSV(DBRecordSet rstIn, String strFileNm) { // TODO: Use of ByRef founded Public Sub fnPersistRecordsetToCSV(ByRef rstIn As ADODB.Recordset, ByVal strFileNm As String)
    // Comments  : Persists the specified ADO recordset to the specified tab-delimited file.
    //
    // Parameters: rstIn     (in) - an ADO recordset
    //             strFileNm (in) - the fully qualified filename to which to persist the rst
    //
    // Returns   : N/A
    //
    // Modified  :
    //
    // --------------------------------------------------
    Const(cintBlockSize As Integer == 100);
    String strTemp = "";
    int lngCounter = 0;
    DBField fld = null;
    try {

      // Delete the file, if it exists, effectively overwriting it.
      Kill(strFileNm);

      Open(strFileNm For Output As #1);

      // Print line with field names and headings
      for (int _i = 0; _i < rstIn.Fields.size(); _i++) {
        fld = rstIn.Fields.item(_i);
        strTemp = strTemp+ fld.Name+ ",";
      }
      // Drop trailling comma
      strTemp = strTemp.substring(0, strTemp.length() - 1);
      Print #1, strTemp;

      lngCounter = 1;

      if (!(rstIn.BOF && rstIn.EOF)) {
        rstIn.MoveFirst;
      }

      do Until .EOF        if (lngCounter != 1) {
          strTemp = rstIn.GetString(, cintBlockSize, """, """, """"+ "\\r\\n"+ """", "");
        } 
        else {
          // Prepend the opening quotes for the first field of the first row
          strTemp = """"+ String.valueOf(rstIn.GetString(, cintBlockSize, """, """, """"+ "\\r\\n"+ """", ""));
        }
        if (rstIn.EOF) {
          // Drop the double quote character printed in excess
          strTemp = strTemp.substring(0, strTemp.length() - 1);
        }
        Print #1, strTemp;;
      }

      // Go back to the first record, so subsequent accesses won't be starting at .EOF
      if (!(rstIn.BOF && rstIn.EOF)) {
        rstIn.MoveFirst;
      }

      Close(#1);
}
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public static void fnPersistRecordsetToXML(DBRecordSet rstIn, String strFileNm) { // TODO: Use of ByRef founded Public Sub fnPersistRecordsetToXML(ByRef rstIn As ADODB.Recordset, ByVal strFileNm As String)
    // Comments  : Persists the specified ADO recordset to the specified file.
    //
    // Parameters: rstIn     (in) - an ADO recordset
    //             strFileNm (in) - the fully qualified filename to which to persist the rst
    //
    // Returns   : N/A
    //
    // Modified  :
    //
    // --------------------------------------------------
    try {

      // Delete the file, if it exists, effectively overwriting it.
      Kill(strFileNm);

      rstIn.Save(strFileNm, adPersistXML);
}
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public static String fnQuoted(String strIn) {
    String _rtn = "";
    // Comments  : Returns the input string, surrounded by single quotes,
    //             for use with building SQL statements or values that
    //             will be used as parameters to stored procdures.
    // Parameters: strIn (in) the string to surround
    //
    // Returns   : quoted string, e.g., xxx  ==> 'xxx'
    //
    // Modified  :
    //
    // --------------------------------------------------
    try {
      "fnQuoted"
.equals(Const cstrCurrentProc As String);

      //fnQuoted = gcstrDoubleQuote & strIn & gcstrDoubleQuote
      _rtn = GCSTRSINGLEQUOTE+ strIn+ GCSTRSINGLEQUOTE;
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here

    if (gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public static void fnRemoveCloseButton(Form frmIn) {
    //----------------------------------------------------------------------------
    // Procedure :  Sub fnRemoveCloseButton
    //
    // Comments  :  Remove the Close command from the system menu and disable
    //              the use of Alt-F4, for the specified form.
    // Called by :  frmLogOn
    // Parameters:  frmIn (in) - pointer to the Form object to process
    //
    // Modified  :
    //----------------------------------------------------------------------------
    "fnRemoveCloseButton"
.equals(Const cstrCurrentProc As String);
    Const(clngMF_BYPOSITION == &H400&);
    int lngHmenu = 0;
    int lngItemCount = 0;

    try {

      // Get the handle of the system menu
      lngHmenu = GetSystemMenu(frmIn.hWnd, 0);

      // Remove the system menu Close menu item
      RemoveMenu(lngHmenu, 6, clngMF_BYPOSITION);
      // Remove the system menu separator line
      RemoveMenu(lngHmenu, 5, clngMF_BYPOSITION);
}
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public static String fnRemoveUnderScoresFromFieldName(String strIn) {
    String _rtn = "";
    // Comments  : Removes underscores from the specified string
    // Parameters:
    //   strIn (in) - string from which to remove underscore characters
    //
    // Returns   : string without underscores
    //
    // Called by :
    //
    // Modified  :
    //
    // --------------------------------------------------
    try {
      "fnRemoveUnderScoresFromFieldName"
.equals(Const cstrCurrentProc As String);

      _rtn = strIn.replace("_", "");
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public static double fnRound(double dblAmountIn, int intSignIn) {
    double _rtn = 0;
    // Comments  :
    // Parameters: dblAmountIn - the amount to round
    //             intSignIn -
    // Returns   : Double - the rounded version of dblAmountIn
    //
    // Called by :
    //
    // Modified  :
    //
    // --------------------------------------------------
    try {
      "fnRound"
.equals(Const cstrCurrentProc As String);
      int lngSignChange = 0;

      if ((dblAmountIn >= 0)) {
        lngSignChange = 1;
      } 
      else {
        lngSignChange = -1;
      }

      _rtn = dblAmountIn + (0.5 * 0.1 ^ intSignIn) * lngSignChange;
      _rtn = fnRound() * 10 ^ intSignIn;
      _rtn = Fix(fnRound());
      _rtn = fnRound() / 10 ^ intSignIn;
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



// MME START WRUS 4999

//////////////////////////////////////////////////////////////////////////////////////////////////
  private static DBRecordSet fnSelectRecord(int lngKey1) {
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
    cadwADOWrapper adwTemp = null;

    try {

      adwTemp = new cadwADOWrapper();

      if (!(adwTemp.commandSetSproc(cstrSproc))) {
        // **TODO:** goto found: GoTo PROC_EXIT;
      }

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
    fnFreeObject(prmReturnValue);
    fnFreeObject(prmClmId);

    if (gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return null;
    // **TODO:** label found: PROC_ERR:;
    switch (prmReturnValue) {
      case  modResConstants.gCRES_NERR_REC_NOT_FOUND:
        // 4027 = The specified record was not found in the database (@@1).
        gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_REC_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, "Claim ID "+ RTrim$(lngKey1));
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
        break;

      case  modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO:
        // 4028 = An error occurred while attempting to @@1 this record.
        gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_ERR_WHILE_TRYING_TO, MCSTRNAME+ cstrCurrentProc, "locate");
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
        break;
    }

    // If any other errors exist, i.e. in Err object, then let it fall through into default error handling.

    switch (VBA.ex.Number) {
      //' Object not found
      case  -2147217900:
        gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_FERR_SPROC_NOT_FOUND, MCSTRNAME+ cstrCurrentProc, cstrSproc);
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
        break;

      default:
        gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }
}

// MME END WRUS 4999



//////////////////////////////////////////////////////////////////////////////////////////////////
  public static void fnSetTopmostWindow(Form frm, boolean bTopmost) {
    //----------------------------------------------------------------------------
    // Procedure :  Function fnSetTopmostWindow
    //
    // Comments  : Makes a form the topmost window or reverts it to normal status
    // Source    : VBMaximizer code library
    // Called by : cmdOK_Click( ) in the frmPrintReport form
    // Parameters: hWnd (in) - window handle to form to operate upon
    //             bTopmost (in) - True if window should be made topmost; False
    //                  otherwise
    //
    // Modified  :
    //----------------------------------------------------------------------------
    try {
      "fnSetTopmostWindow"
.equals(Const cstrCurrentProc As String);

      Const(HWND_TOPMOST == -1);
      Const(HWND_NOTOPMOST == -2);
      Const(SWP_NOSIZE == &H1);
      Const(SWP_NOMOVE == &H2);
      Const(SWP_NOACTIVATE == &H10);

      SetWindowPos(frm.hWnd, bTopmost ? HWND_TOPMOST : HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE || SWP_NOSIZE || SWP_NOACTIVATE);
      // **TODO:** label found: PROC_EXIT:;
      // Use Resume Next rather than GoTo 0 since Close could error if TS wasn't successfully opened
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here

    if (gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
}
}


//////////////////////////////////////////////////////////////////////////////////////////////////
  public static String fnShowRecordPosition(DBRecordSet rstIn) {
    String _rtn = "";
    //----------------------------------------------------------------------------
    // Procedure :  Function fnShowRecordPosition
    //
    // Comments  : Used to build "Record x of y" label on the screen to denote
    //             current record position
    // Called by : cmdNavigate_Click( ), Form_Load( ) in Insured and Payee forms
    // Parameters: N/A
    //
    // Modified  :
    //----------------------------------------------------------------------------
    try {
      "fnShowRecordPosition"
.equals(Const cstrCurrentProc As String);
      String strPos = "";

      // It would be nice if this function could just use the table wrapper's Lookup
      // recordset's public functions, but that's not possible since this function
      // must support *all* table wrapper's Lookup recordsets. So, we must
      // continue to receive an ADO Recordset as input (i.e. <tablewrapper>.LookupData)
      // and go from there.

      switch (rstIn.AbsolutePosition) {
        case  adPosBOF:
          strPos = "0";
          break;

        case  adPosEOF:
          strPos = CStr(rstIn.RecordCount);
          break;

        case  adPosUnknown:
          strPos = "?";
          break;

        default:
          strPos = CStr(rstIn.AbsolutePosition);
          break;
      }

      _rtn = "Record "+ strPos+ " of "+ rstIn.RecordCount;
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here

    if (gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public static String fnSSNTIN_AddDash(String strIn, boolean bIsTin) {
    String _rtn = "";
    // Comments  : Returns the input string with a dash added between:
    //             * characters 3 and 4 and 5 and 6, if bIsTin = True
    //             * characters 2 and 3, if bIsTin = False
    // Parameters: strIn (in) - a 9-digit Social Security Number or
    //                          Taxpayer Identification Number
    //
    // Returns   : string, e.g., 123456789  ==> '123-45-6789' or '12-3456789'
    //
    // Modified  :
    //
    // --------------------------------------------------
    try {
      "fnSSNTIN_AddDash"
.equals(Const cstrCurrentProc As String);
      "-"
.equals(Const cstrDash As String);

      if (bIsTin) {
        _rtn = strIn.substring(0, 2)+ cstrDash+ strIn.substring(strIn.length() - 7);
      } 
      else {
        _rtn = strIn.substring(0, 3)+ cstrDash+ strIn.substring(4, 2)+ cstrDash+ strIn.substring(strIn.length() - 4);
      }
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here

    if (gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  private static void fnTerminateTheApp() {
    //----------------------------------------------------------------------------
    // Procedure :  Sub fnTerminateTheApp
    //
    // Comments  :  Unloads all loaded forms, in the reverse order from which
    //              they were originally ordered
    // Parameters:  N/A
    //
    // Called by : cmdCancel_Click() of frmLogOn
    //             MDIForm_Unload() of frmMDIMain
    //
    // Modified  :
    //----------------------------------------------------------------------------
    "fnTerminateTheApp"
.equals(Const cstrCurrentProc As String);
    "frmMDIMain"
.equals(Const cstrMDIForm As String);
    int intIndex = 0;
    int intFormsCount = 0;

    //' No error handler. Nowhere to go :(
    try {

      gbAmTryingToTerminateTheApp = true;

      if (BDEBUGAPPTERMINATION) {
        Debug.Print("Entering fnTerminateTheApp");
      }


      // Unload the Splash screen if it's still around. Otherwise, the attempt to unload
      // the MDI form will fail and the app won't truly be shut down.
      fnUnloadSplash();

      intFormsCount = Forms.LinkedMap.size();

      if (BDEBUGAPPTERMINATION) {
        Debug.Print("   Number of forms in memory = "+ CStr(intFormsCount));
      }

      if (intFormsCount > 0) {
        for (intIndex = intFormsCount - 1; intIndex >= 0; intIndex--) {
          if (BDEBUGAPPTERMINATION) {
            Debug.Print("   Trying to unload "+ Forms(intIndex).Name+ " from fnTerminateTheApp");
          }
          // Only attempt to unload the MDI Main form if it is now the only form left
          // in memory.
          //
          // Since the forms are unloaded in the reverse order from which
          // they were loaded, the MDI Form should always be the last form unloaded.
          // Therefore, if all other forms that were open have been unloaded, it should
          // be okay to unload the MDI. If any of those forms were NOT unloaded (probably
          // because the user said "No I don't want to lose my pending changes"), then
          // don't unload the MDI form.  This conditionality is needed to avoid the user
          // being prompted twice (or more) about losing their pending changes on the same
          // form.
          if (Forms(intIndex).Name == cstrMDIForm) {
            if (Forms.LinkedMap.size() == 1) {
              Unload(Forms(intIndex));
            }
          } 
          else {
            Unload(Forms(intIndex));
          }
        }
      }

      gbAmTryingToTerminateTheApp = false;
//*TODO:** the error label 0: couldn't be found
  }
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public static double fnTranslateToMaxValue(int intDollarPositions, int intDecimalPositions) {
    double _rtn = 0;
    // Comments  : Translates meta data (i.e. number of dollar and decimal positions) into a numeric value
    //             that represents the Maximum Value allowed.
    //             Example:  fnTranslateToMaxValue(5,4) would return 99999.9999.
    //                       fnTranslateToMaxValue(3,0) would return 999 (equivalent to 999.0)
    //
    // Parameters: intDollarPositions  (in) - the number of dollar positions allowed
    //             intDecimalPositions (in) - the number of decimal positions allowed
    //
    // Returns   : Double representing the maximum value
    //
    // Modified  :
    //
    // --------------------------------------------------
    try {
      "fnSSNTIN_AddDash"
.equals(Const cstrCurrentProc As String);
      String strMaxValue = "";
      int intI = 0;
      "9"
.equals(Const cstrAnotherDigit As String);
      "0"
.equals(Const cstrZero As String);
      "."
.equals(Const cstrDecimalPoint As String);

      for (intI = 1; intI <= intDollarPositions; intI++) {
        strMaxValue = strMaxValue+ cstrAnotherDigit;
      }

      strMaxValue = strMaxValue+ cstrDecimalPoint;

      if (intDecimalPositions == 0) {
        strMaxValue = strMaxValue + cstrZero;
      } 
      else {
        for (intI = 1; intI <= intDecimalPositions; intI++) {
          strMaxValue = strMaxValue + cstrAnotherDigit;
        }
      }

      _rtn = Double.parseDouble(strMaxValue);
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here

    if (gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public static void fnUnloadSplash() {
    //----------------------------------------------------------------------------
    // Procedure :  Sub fnUnloadSplash
    //
    // Comments  :  Get rid of the splash screen, in case the error occurs while
    //              the splash screen is still being displayed. Otherwise,
    //              the splash screen can obscure any message box that is displayed.
    // Called by :  modStartup's sub Main( )
    // Parameters:  N/A
    //
    // Modified  :
    //----------------------------------------------------------------------------
    //On Error GoTo 0
    "fnUnloadSplash"
.equals(Const cstrCurrentProc As String);

    if (fnIsFormLoaded("frmSplash")) {
      Unload(frmSplash);
      fnFreeObject(frmSplash);
    }
    //PROC_EXIT:
    //    On Error GoTo 0     ' disable error handler
    //    ' Clean-up statements go here
    //    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
    //        gerhApp.PropagateError mcstrName & cstrCurrentProc
    //    End If
    //    Exit Sub
    //PROC_ERR:
    //    Select Case Err.Number
    //        'Case statements for expected errors go here
    //        Case Else
    //            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    //    End Select
    //    Resume PROC_EXIT
  }



  //////////////////////////////////////////////////////////////////////////////////////////////////
  private static void fnWindowLock(int hWnd) {
    //----------------------------------------------------------------------------
    // Procedure :  Function fnWindowLock
    //
    // Comments  : To avoid screen flicker caused by excessive repainting, use
    //             this before making a lot of screen changes and then
    //             call its companion procedure (fnWindowUnlock) afterward.
    //
    // Called by : cmdDetailCollapse_Click( ) in the frmMsgBox form
    //             cmdDetailExpand_Click( ) in the frmMsgBox form
    // Parameters: hWnd (in) - window handle of the form to operate upon
    //
    // Modified  :
    //----------------------------------------------------------------------------
    // No error handler since this could be called prior to the app startup being completed
    try {

      LockWindowUpdate(hWnd);
}
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  private static void fnWindowUnlock() {
    //----------------------------------------------------------------------------
    // Procedure :  Function fnWindowUnlock
    //
    // Comments  : To avoid screen flicker caused by excessive repainting, call
    //             this procedure's companion procedure (fnWindowLock)
    //             before making a lot of screen changes and then
    //             call this procedure (fnWindowUnlock) afterward.
    //
    // Called by : cmdDetailCollapse_Click( ) in the frmMsgBox form
    //             cmdDetailExpand_Click( ) in the frmMsgBox form
    // Parameters: N/A
    //
    // Modified  :
    //----------------------------------------------------------------------------
    // No error handler since this could be called prior to the app startup being completed
    try {

      LockWindowUpdate(0);
}
}

//////////////////////////////////////////////////////////////////////////////////////////////////
  public static boolean fnTFToBool(String strIn) {
    boolean _rtn = false;
    // Comments  : Translates "T" to True and everything else to False
    // Parameters: strIn (in) the string expression to translate
    //
    // Returns   : True or False
    //
    // Modified  : Berry Kropiwka
    //
    // --------------------------------------------------
    try {
      "fnTFToBool"
.equals(Const cstrCurrentProc As String);

      if (strIn.toUpperCase().equals("T")) {
        _rtn = true;
      } 
      else {
        _rtn = false;
      }
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here

    if (gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}

//////////////////////////////////////////////////////////////////////////////////////////////////
  public static boolean fnYNToBool(String strIn) {
    boolean _rtn = false;
    // Comments  : Translates "Y" to True and everything else to False
    // Parameters: strIn (in) the string expression to translate
    //
    // Returns   : True or False
    //
    // Modified  :
    //
    // --------------------------------------------------
    try {
      "fnYNToBool"
.equals(Const cstrCurrentProc As String);

      if (strIn.toUpperCase().equals("Y")) {
        _rtn = true;
      } 
      else {
        _rtn = false;
      }
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here

    if (gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
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
//   handler. Then, once you've stopped at the breakpoint, type the function
//   name in the Immediate window.
//       Correct:    TestStub_fnGetStateInfo
//                   modGeneral.TestStub_fnGetStateInfo
//
//       Incorrect:  ? TestStub1
//                   TestStub1()
//
// %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
  private static void testStub_fnGetStateInfo() {
    try {
      "TestStub_fnGetStateInfo"
.equals(Const cstrCurrentProc As String);
      StateInfo siTemp = null;

      // MME START WRUS 4999 - ADDED CLAIMID AND PayeDthbPmtAmt PARAMATERS - CHECK THIS EXISTS IN DEV AND PROD..

      fnGetStateInfo("FL", "I", CDate(DateValue("1/1/1979")), 45418, 2, siTemp);

      Debug.Print("State = "+ siTemp.stCd);
      Debug.Print("StrlEffDt = "+ siTemp.strlEffDt);
      Debug.Print("CalcIdtypCd = "+ siTemp.calcIdtypCd);
      Debug.Print("ReqdIdtypCd = "+ siTemp.reqdIdtypCd);
      Debug.Print("IruleCd = "+ siTemp.iruleCd);
      Debug.Print("StrlEndDt = "+ siTemp.strlEndDt);
      Debug.Print("StrlIntRptgFlrAmt = "+ siTemp.strlIntRptgFlrAmt);
      Debug.Print("StrlIntCalcOfstNum = "+ siTemp.strlIntCalcOfstNum);
      Debug.Print("StrlIntReqdOfstNum = "+ siTemp.strlIntReqdOfstNum);
      Debug.Print("StrlIntRuleAmt = "+ siTemp.strlIntRuleAmt);
      Debug.Print("StrlSpclInstrTxt = "; siTemp.strlSpclInstrTxt);
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }
}

}

public class StateInfo {
    public String lobCd;//' not null
    public String stCd;//' not null
    public Date strlEffDt;//' not null
    public String calcIdtypCd;//' not null
    public String reqdIdtypCd;//' not null
    public String iruleCd;//' not null
    public Variant strlEndDt;//' nullable
    public Currency strlIntRptgFlrAmt;//' not null  decimal(11,2)
    public Integer strlIntCalcOfstNum;//' not null  smallint
    public Integer strlIntReqdOfstNum;//' not null  smalling
    public Variant strlIntRuleAmt;//' nullable  decimal(11,5)
    public String strlSpclInstrTxt;//' nullable
    public Date figuredFromDate;
    public Date payablePeriodEndDate;
    public Integer nbrOfDaysToPayInterest;
    public Double interestRateToUse;
    public Currency claimInterestAmt;
    public Currency withheldAmt;
    public Currency totalForThisPayee;
    public String calculationInfo;
}


public class udtColumn {
    public String colName;//' Corresponds to COLUMN_NAME meta data from Column schema info
    public DataTypeEnum dataType;//' Corresponds to DATA_TYPE meta data from Column schema info
    public Boolean isKey;//' Corresponds to XXX from PrimaryKeys schema info
    public Boolean isNullable;//' Corresponds to IS_NULLABLE  meta data from Column schema info
    public Boolean hasDefault;//' Corresponds to COLUMN_HASDEFAULT meta data from Column schema info
    public Variant defaultValue;//' Corresponds to COLUMN_DEFAULT meta data from Column schema info
    public Integer dollarPositions;//' Calculated from PRECISION meta data from Column schema info, but which could be overriden
    public Integer decimalPositions;//' Corresponds to SCALE meta data from Column schema info, but which could be overriden
    public Integer precision;//' Corresponds to original PRECISION from DBMS. SHOULD NOT be overriden!
    public Integer numericScale;//' Corresponds to original SCALE from DBMS. SHOULD NOT be overriden!
    public Integer maxCharacters;//' Correspond to CHARACTER_MAXIMUM_LENGTH meta data from Column schema info
    public String format;//' Initially set based on DataType, but form can override
    public String mask;//' Initially set based on DataType, but form can override
    public String allowableCharacters;//' Initially set based on DataType and DecimalPositions, but form can override
    public Boolean shouldForceToUppercase;//' Does *not* correspond to DBMS meta data.
    public Variant value;//' Initially set based on DefaultValue, if present, but form can override
}


public class EnumNavigationButtons {
    public static final int NAVFIRST = 0;
    public static final int NAVPREV = 1;
    public static final int NAVNEXT = 2;
    public static final int NAVLAST = 3;
}


public class EnumPrevNext {
    public static final int EPNPREV = 0;
    public static final int EPNNEXT = 1;
}


public class enumPositionDirection {
    public static final int EPDPREVIOUSRECORD = 0;
    public static final int EPDNEXTRECORD = 1;
    public static final int EPDSAMERECORD = 2;
    public static final int EPDFIRSTRECORD = 3;
    public static final int EPDLASTRECORD = 4;
}


public class enumWhatOperationIsBeingAttempted {
    public static final int EWOUPDATE = 0;
    public static final int EWODELETE = 1;
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


case class OdgeneralData(
              id: Option[Int],

              )

object Odgenerals extends Controller with ProvidesUser {

  val odgeneralForm = Form(
    mapping(
      "id" -> optional(number),

  )(OdgeneralData.apply)(OdgeneralData.unapply))

  implicit val odgeneralWrites = new Writes[Odgeneral] {
    def writes(odgeneral: Odgeneral) = Json.obj(
      "id" -> Json.toJson(odgeneral.id),
      C.ID -> Json.toJson(odgeneral.id),

    )
  }

  def get(id: Int) = GetAction { implicit request =>
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.LIST_ODGENERAL), { user =>
      Ok(Json.toJson(Odgeneral.get(user, id)))
    })
  }

  def update(id: Int) = PostAction { implicit request =>
    Logger.debug("in odgenerals.update")
    odgeneralForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      odgeneral => {
        Logger.debug(s"form: ${odgeneral.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.EDIT_ODGENERAL), { user =>
          Ok(
            Json.toJson(
              Odgeneral.update(user,
                Odgeneral(
                       id,

                ))))
        })
      }
    )
  }

  def create = PostAction { implicit request =>
    Logger.debug("in odgenerals.create")
    odgeneralForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      odgeneral => {
        Logger.debug(s"form: ${odgeneral.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.NEW_ODGENERAL), { user =>
          Ok(
            Json.toJson(
              Odgeneral.create(user,
                Odgeneral(

                ))))
        })
      }
    )
  }

  def delete(id: Int) = PostAction { implicit request =>
    Logger.debug("in odgenerals.delete")
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.DELETE_ODGENERAL), { user =>
      Odgeneral.delete(user, id)
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

case class Odgeneral(
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

object Odgeneral {

  lazy val emptyOdgeneral = Odgeneral(
)

  def apply(
      id: Int,
) = {

    new Odgeneral(
      id,
)
  }

  def apply(
) = {

    new Odgeneral(
)
  }

  private val odgeneralParser: RowParser[Odgeneral] = {
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
        Odgeneral(
              id,
,
              createdAt,
              updatedAt,
              updatedBy)
    }
  }

  def create(user: CompanyUser, odgeneral: Odgeneral): Odgeneral = {
    save(user, odgeneral, true)
  }

  def update(user: CompanyUser, odgeneral: Odgeneral): Odgeneral = {
    save(user, odgeneral, false)
  }

  private def save(user: CompanyUser, odgeneral: Odgeneral, isNew: Boolean): Odgeneral = {
    def getFields = {
      List(

      )
    }
    def throwException = {
      throw new RuntimeException(s"Error when saving ${C.ODGENERAL}")
    }

    DBHelper.saveEx(
      user,
      Register(
        C.ODGENERAL,
        C.ID,
        odgeneral.id,
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

  def load(user: CompanyUser, id: Int): Option[Odgeneral] = {
    loadWhere(user, s"${C.ID} = {id}", 'id -> id)
  }

  def loadWhere(user: CompanyUser, where: String, args : scala.Tuple2[scala.Any, anorm.ParameterValue[_]]*) = {
    DB.withConnection(user.database.database) { implicit connection =>
      SQL(s"SELECT t1.*, t2.${C.FK_NAME} FROM ${C.ODGENERAL} t1 INNER JOIN ${C.???} t2 ON t1.${C.FK_ID} = t2.${C.FK_ID} WHERE $where")
        .on(args: _*)
        .as(odgeneralParser.singleOpt)
    }
  }

  def delete(user: CompanyUser, id: Int) = {
    DB.withConnection(user.database.database) { implicit connection =>
      try {
        SQL(s"DELETE FROM ${C.ODGENERAL} WHERE ${C.ID} = {id}")
        .on('id -> id)
        .executeUpdate
      } catch {
        case NonFatal(e) => {
          Logger.error(s"can't delete a ${C.ODGENERAL}. ${C.ID} id: $id. Error ${e.toString}")
          throw e
        }
      }
    }
  }

  def get(user: CompanyUser, id: Int): Odgeneral = {
    load(user, id) match {
      case Some(p) => p
      case None => emptyOdgeneral
    }
  }
}


// Router

GET     /api/v1/general/odgeneral/:id              controllers.logged.modules.general.Odgenerals.get(id: Int)
POST    /api/v1/general/odgeneral                  controllers.logged.modules.general.Odgenerals.create
PUT     /api/v1/general/odgeneral/:id              controllers.logged.modules.general.Odgenerals.update(id: Int)
DELETE  /api/v1/general/odgeneral/:id              controllers.logged.modules.general.Odgenerals.delete(id: Int)




/**/
