public class cerhErrorHandler {

  // =-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-=
  // Class       : cerhErrorHandler
  // Description : Used to support a centralized transaction-oriented error handling
  //               scheme across the application
  //
  //               Note that this class does NOT CONTAIN ERROR HANDLERS, since it
  //               is itself an error handler. If any errors are encountered,
  //               then VB itself will report them...as fatal errors that cause the
  //               app to terminate.
  //
  // Procedures  :
  //   Private     Class_Initialize()
  //   Public      Property Get ErrNum() As Long
  //   Public      Property Let ScreenName(ByVal strValue As String)
  //   Public      AddSubstitution(ByVal strValue As String)
  //   Public      Clear()
  //   Public      PropagateError()
  //   Public      ReportFatalError(ByVal strScreenName As String) As Integer
  //   Public      ReportNonFatal(ByVal lngErrNumIn As Long, ByVal strErrContextIn As String, _
  *ParamArray varParmsIn() As Variant) As Integer
  //   Public      SaveAppSpecificErr(ByVal lngErrNumIn As Long, ByVal strErrContextIn As _
  //                   String, ParamArray varParmsIn() As Variant)
  //   Public      SaveErrObjectData(ByVal strAppContext As String) As Boolean
  //   Private     fnGetSubstitutionsLBound() As Long
  //   Private     fnGetResString(ByVal intID As Integer) As String
  //   Private     fnGetSubstitutionsUBound() As Long
  //
  // Modified:
  //
  //   Version Date     Who   What
  //   ------- -------- ---   -------------------------------------------------------------------
  //   4.0     03/20/02 BAW   (Phase2A) Commented out procedures that aren't used and possibly
  //                          could be deleted.
  //   3.0     03/11/02 BAW   (Phase2A) Made SaveAppSpecificError( ) raise the error that was just
  //                          recorded so that it could be handled by the local error handler
  //                          immediately (for non-error situations) or propogated (for errors).
  //                          Also added TranslatedErrNum as a public property.
  //   2.0     03/07/02 BAW   (Phase2A) Updated comments at top of module. Added conditional
  //                          compilation code (#If DEBUG_ERH) with which to debug error
  //                          handling code.
  // =-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-=
//Option Explicit
  *Option Compare Binary

  // Local variables to hold Public Property values
  private int m_lngErrNum = 0;
  private String m_strErrDesc = "";
  private String m_strErrContext = "";
  private String m_strScreenName = "";
  private int m_lngTranslatedErrNum = 0;

  // Local variables that are *NOT* Public Properties
  private String[] m_strSubstitutions = "";

  // Following constants are used with procedures accessing the m_strSubstitutions array
  private static final Long MCLNGSUBSCRIPTOUTOFRANGE = 9;
  private static final Long MCLNGARRAYNOTINITIALIZED = -1;

  // The following constants are used to initialize the public properties
  // of this object, or to determine whether the properties are at their
  // initialized values.
  private static final String MCSTRERRDESCDEFAULT = vbNullString;
  private static final String MCSTRERRCONTEXTDEFAULT = vbNullString;

  // The following defines how substitution placeholders are identified in message text
  // defined in the CLAIM.RES resource file
  private static final String MCSTRSUBSTITUTIONDELIMITER = "@@";




  ///\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
  //\                                                                  /
  //|         CLASS_INITIALIZE / CLASS_TERMINATE   Procedures         |
  ///                                                                  \
  //\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/

  //=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
  private void class_Initialize() {
    // Set initial values to defaults which may be overridden with property settings
    clear();
  }


  private void class_Terminate() {
    // Free up resources allocated in this class

    End;
    //Erase m_strSubstitutions
  }




  ///\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
  //\                                                                  /
  //|                PROPERTY GET/LET    Procedures                    |
  ///                                                                  \
  //\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/

  //=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
  public int getErrNum() {
    // Returns the current value of the ErrNum property

    return m_lngErrNum;
  }


  public void setScreenName(String strValue) {
    // Sets the ScreenName property to the value specified by strValue

    m_strScreenName = strValue;
  }







  ///\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
  //\                                                                  /
  //|                        PUBLIC  Procedures                        |
  ///                                                                  \
  //\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/

  //=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
  public void addSubstitution(String strValue) {
    // Comments:   Resizes the current array and then adds the entry
    // Parameters: strValue   (in)   Entry to add to the array
    // Returns:    N/A
    // Called by : Any application procedure that needs to generate
    //             a message that utilizes substitution values in
    //             its text, e.g., @@1, @@2, etc.
    int lngLastEntry = 0;

    //' Validate the input
    Debug.Assert(strValue != "");

    lngLastEntry = fnGetSubstitutionsUBound();
    if (lngLastEntry == MCLNGARRAYNOTINITIALIZED) {
      lngLastEntry = 0;
    } 
    else {
      lngLastEntry = lngLastEntry + 1;
    }

    G.redimPreserve(m_strSubstitutions, lngLastEntry);
    m_strSubstitutions[lngLastEntry] = strValue;
  }



  //=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
  public void clear() {
    // Comments:  Resets all properties to their intialalized values
    //
    // Parameters: None
    // Returns:    N/A
    // Called by : ReportFatalError, Class_Initialize

    //' Fatal Error Code, meant to indicate it hasn't been set
    m_lngErrNum = modResConstants.gCLNGERR_NUM_DEFAULT;
    m_strErrDesc = MCSTRERRDESCDEFAULT;
    m_strErrContext = MCSTRERRCONTEXTDEFAULT;
    Erase(m_strSubstitutions);
  }


  //=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
  public void propagateError(String strErrContextIn) {
    // Comments  : This public subroutine is used when an error must be raised
    //             from a standard module, class method, or non-event handler
    //             back up to the event handler that initiated it. It assumes
    //             the error has already been saved to this class' member variables
    //             and thus raises the error based on those member variables and
    //             NOT the VBA.Err object.
    //
    // Parameters: N/A
    // Returns   : N/A
    // Modified  :
    // --------------------------------------------------

    #If DEBUG_ERH Then;
    Debug.Print("PropagateError called by "+ strErrContextIn+ ". Error#="+ CStr(m_lngErrNum)+ "(or "+ String.valueOf(CStr(m_lngErrNum - vbObjectError))+ ") Desc="+ m_strErrDesc+ " Context="+ strErrContextIn);
    #End If;

    VBA.ex.Raise(m_lngErrNum, m_strErrContext, m_strErrDesc);
  }



  //=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
  public int reportFatalError(String strScreenName) {
    int _rtn = 0;
    // Comments  : This public subroutine should be used INSTEAD OF the VBA.MsgBox
    //             function to report *errors*, whether they be PROCESS fatal
    //             (i.e. the event terminates, but not the app) or APP fatal (the
    //             app terminates). All of the information it needs to report the error
    //             is based on public properties that should
    //             be set prior to calling this procedure.
    //
    //             It is assumed that ErrNum reflect this scheme in the CLAIM.RES resource file:
    //                   ID values in the range of:
    //                      4000-4999 = non-fatal error messages
    //                      9000-9999 = fatal error messages
    //
    //             Anything outside this range (such as VB and ADO errors) will be
    //             treated as a fatal error
    //
    //             Note that ErrNums in the following ranges should be reported
    //             via the ReportNotFatal( ) method of this class:
    //                      1000-1999 = informational messages
    //                      2000-2999 = warning messages
    //                      3000-3999 = alert messages
    //
    // Parameters: strScreenName (in) - the name of the screen (as it should appear in
    //                      the MsgBox's title
    //
    // Returns   : Integer indicating the button the user clicked in the MsgBox,
    //             e.g., vbOK, vbYes, vbNo, etc.
    // Modified  :
    // --------------------------------------------------
    "ReportFatalError"
.equals(Const cstrCurrentProc As String);
    int intI = 0;
    int lngButtons = 0;
    int lngTranslatedErrNum = 0;
    String strErrDesc = "";
    String strMsgText = "";
    Const(cstrFatalErrorPrefix As String == "An error has occurred from which the application cannot recover. "+ "The application will now be terminated."+ "\\r\\n"+ "\\r\\n");

    try {

      #If DEBUG_ERH Then;
      Debug.Print("ReportFatalError called by "+ strScreenName+ ". Error#="+ CStr(m_lngErrNum)+ "(or "+ String.valueOf(CStr(m_lngErrNum - vbObjectError))+ ") Desc="+ m_strErrDesc+ " Context="+ m_strErrContext);
      #End If;

      // Make sure Cursor reverts back to normal, in case it was left in an hourglass
      Screen.MousePointer = vbDefault;

      // Remove the vbObjectError value, so we're left with an app-specific error code, as is
      // used in the Resource File. If this lowers it too much, then revert back to the
      // specified Error Code (probably is an ADO or VB error).
      lngTranslatedErrNum = m_lngErrNum - vbObjectError;
      if (lngTranslatedErrNum < modResConstants.gCRES_LOWEST_APP_ERROR || lngTranslatedErrNum > modResConstants.gCRES_HIGHEST_APP_ERROR) {
        lngTranslatedErrNum = m_lngErrNum;
        strMsgText = m_strErrDesc;
      } 
      else {
        strMsgText = fnGetResString(lngTranslatedErrNum);

        // Replace Carriage Return / Line Feed tokens with VB-equivalent
        //        @@CRLF ==> vbCrLf
        //        @@CR   ==> vbCr
        //        @@LF   ==> vbLf
        strMsgText = strMsgText.replace(MCSTRSUBSTITUTIONDELIMITER+ "CRLF", "\\r\\n");
        strMsgText = strMsgText.replace(MCSTRSUBSTITUTIONDELIMITER+ "CR", "\\n");
        strMsgText = strMsgText.replace(MCSTRSUBSTITUTIONDELIMITER+ "LF", "\\r");

        // Replace parameters in message text with substituted values, e.g.,
        //      @@1  ... for the value of varParms(0)
        //      @@2  ... for the value of varParms(1), etc.
        if (fnGetSubstitutionsUBound() != MCLNGARRAYNOTINITIALIZED) {
          for (intI = LBound(m_strSubstitutions); intI <= m_strSubstitutions.length; intI++) {
            strMsgText = strMsgText.replace(MCSTRSUBSTITUTIONDELIMITER+ String.valueOf(CStr(intI + 1)), m_strSubstitutions[intI]);
          }
          // Do assertion if any "@@" still remain after doing the substitution. It
          // indicates the call to this function and Error text in the .RES file are out of synch.
          Debug.Assert(strMsgText.indexOf(MCSTRSUBSTITUTIONDELIMITER, 1) == 0);
        }
      }

      switch (lngTranslatedErrNum) {
        case  modResConstants.gCRES_NERR_START To modResConstants.gCRES_NERR_END:
          lngButtons = vbOKOnly+ vbExclamation;
          break;

        default:
          // Intended to be 9000 to 9999 ... or a VB or ADO error code
          lngButtons = vbOKOnly+ vbExclamation;
          modGeneral.gbAmProcessingAnAppFatalError = true;
          break;
      }

      if (modGeneral.gbAmProcessingAnAppFatalError) {
        modAppLog.fnLogWrite("Application Fatal Error in "+ m_strErrContext+ ": "+ strMsgText, cstrCurrentProc);
        // Display the errror via a modal frmMsgBox window
        //*TODO:** can't found type for with block
        //*With frmMsgBox
        __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = frmMsgBox;
        w___TYPE_NOT_FOUND.ScreenName = strScreenName;
        w___TYPE_NOT_FOUND.ErrorCode = lngTranslatedErrNum;
        w___TYPE_NOT_FOUND.MsgText = cstrFatalErrorPrefix+ strMsgText;
        w___TYPE_NOT_FOUND.ErrorContext = m_strErrContext;
        w___TYPE_NOT_FOUND.Show(vbModal);
        _rtn = w___TYPE_NOT_FOUND.ButtonClicked;
        Unload(frmMsgBox);

        // Initialize this class' member variables to acknowledge the error has been reported
        // and thus no longer needs to be Propagated or reported again, e.g., in the
        // Unload events triggered by the following FOR loop
        clear();

        //' If in IDE, force the debugger to stop here
        Debug.Assert(false);
        for (intI = Forms.LinkedMap.size() - 1; intI >= 0; intI--) {
          Unload(Forms(intI));
        }

        modGeneral.fnDeallocateGlobalObjects();

        modGeneral.gbAmProcessingAnAppFatalError = false;

        // !!!!!!!  TERMINATE THE APP !!!!!!!
        End;
      } 
      else {
        modAppLog.fnLogWrite("Error in "+ m_strErrContext+ ": "+ strMsgText, cstrCurrentProc);
        //*TODO:** can't found type for with block
        //*With frmMsgBox
        __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = frmMsgBox;
        w___TYPE_NOT_FOUND.ScreenName = strScreenName;
        w___TYPE_NOT_FOUND.ErrorCode = lngTranslatedErrNum;
        w___TYPE_NOT_FOUND.MsgText = strMsgText;
        w___TYPE_NOT_FOUND.ErrorContext = m_strErrContext;

        // Initialize this class' member variables to acknowledge the error has been reported
        // and thus no longer needs to be Propagated or reported again. Do it now,
        // before showing frmMsgBox, because if we somehow get into a procedure's PROC_EXIT
        // with this cerhErrorHandler object still showing the remains of an error, there
        // is nothing to propagate the error back up to and the app will die with an
        // unhandled error. (Example: cetbExtendedTextbox.Lost_Focus will be called if an
        // error occurs on a maintenance screen that has TextBoxes tied to that extended
        // textbox class and that LostFocus event handler will treat the as-yet-unreported-error
        // (if we didn't call .Clear first) as an unhandled error.
        clear();

        w___TYPE_NOT_FOUND.Show(vbModal);
        _rtn = w___TYPE_NOT_FOUND.ButtonClicked;
        Unload(frmMsgBox);

        // Unload the splash screen, if it is still loaded
        modGeneral.fnUnloadSplash();
      }
}

  return _rtn;
}



//=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
  public int reportNonFatal(int lngErrNumIn, String strErrContextIn, Object[] varParmsIn) {
    int _rtn = 0;
    // Comments  : This public subroutine should be used whenever a Warning,
    //             Informational or Alert type of message should be displayed.
    //             It can be called anywhere in the system, e.g, form or std
    //             module, or event handler or non-event handler.
    //
    //             Use this function instead of VBA.MsgBox( ).
    //
    //             The messages displayed by this function are *not* recorded
    //             in the class' data.
    //
    //             It is assumed that ErrNum reflects this scheme in the CLAIM.RES resource file:
    //                   ID values in the range of:
    //                      1000-1999 = informational messages
    //                      2000-2999 = warning messages
    //                      3000-3999 = alert messages
    //             Anything outside this range will raise an assert while in
    //             Debug mode.
    //
    // Parameters:
    //             lngErrNumIn (in)     - the error code (in the range of gcRES_LOWEST_APP_ERROR
    //                                    to gcRES_HIGHEST_APP_ERROR)
    //             strErrContextIn (in) - the module.procname in which the error occurred
    //             varParmsIn (in)      - (optional) a variable number of parameters which
    //                                    will eventually be substituted for  @@1, @@2, etc.
    //                                    placeholders in the message text.
    //
    // Returns   : Integer indicating the button the user clicked in the MsgBox,
    //             e.g., vbOK, vbYes, vbNo, etc.
    // Modified  :
    //    04/23/02 BAW  Added logic to clear the Err object.
    // --------------------------------------------------
    "ReportNonFatal"
.equals(Const cstrCurrentProc As String);
    int intI = 0;
    int lngTranslatedErrNum = 0;
    String strMsgText = "";
    try {

      #If DEBUG_ERH Then;
      Debug.Print("ReportNonFatal called by "+ strErrContextIn+ ". Error#="+ CStr(lngErrNumIn)+ "(or "+ String.valueOf(CStr(m_lngErrNum - vbObjectError))+ ")");
      #End If;

      // Clear Err object in case an error is left in it. This could occur in the case of a stored procedure
      // returning a 4029 (Dependent Records exist) error that the form, at times, will turn into an "info" message.
      VBA.ex.Clear;

      // Remove the vbObjectError value, so we're left with an app-specific error code, as is
      // used in the Resource File.
      lngTranslatedErrNum = lngErrNumIn - vbObjectError;

      // Verify we've got an lngErrNumIn in the appropriate range
      Debug.Assert(lngTranslatedErrNum >= modResConstants.gCRES_INFO_START);
      Debug.Assert(lngTranslatedErrNum <= modResConstants.gCRES_ALRT_END);

      // If the translated error number is too low (which shouldn't happen in the
      // production environment due to the previous Assert bringing it to the
      // developer's attention during the development process), then revert back
      // to the specified Error Code.
      if (lngTranslatedErrNum < modResConstants.gCRES_LOWEST_APP_ERROR || lngTranslatedErrNum > modResConstants.gCRES_HIGHEST_APP_ERROR) {
        lngTranslatedErrNum = lngErrNumIn;
        strMsgText = "The following error was encountered but its description "+ "could not be located: "+ "\\n"+ m_strErrDesc;
      } 
      else {
        strMsgText = fnGetResString(lngTranslatedErrNum);

        // Replace Carriage Return / Line Feed tokens with VB-equivalent
        //        @@CRLF ==> vbCrLf
        //        @@CR   ==> vbCr
        //        @@LF   ==> vbLf
        strMsgText = strMsgText.replace(MCSTRSUBSTITUTIONDELIMITER+ "CRLF", "\\r\\n");
        strMsgText = strMsgText.replace(MCSTRSUBSTITUTIONDELIMITER+ "CR", "\\n");
        strMsgText = strMsgText.replace(MCSTRSUBSTITUTIONDELIMITER+ "LF", "\\r");

        // Replace parameters in message text with substituted values, e.g.,
        //      @@1  ... for the value of varParms(0)
        //      @@2  ... for the value of varParms(1), etc.

        for (intI = LBound(varParmsIn); intI <= varParmsIn.length; intI++) {
          strMsgText = strMsgText.replace(MCSTRSUBSTITUTIONDELIMITER+ String.valueOf(CStr(intI + 1)), varParmsIn(intI));
        }
        // Do assertion if any "@@" still remain after doing the substitution. It
        // indicates the call to this function and Error text in the .RES file are out of synch.
        Debug.Assert(strMsgText.indexOf(MCSTRSUBSTITUTIONDELIMITER, 1) == 0);
      }

      // May need to make this log file write conditional, or not do it at all,
      // if the log file starts filling up too quickly...
      modAppLog.fnLogWrite("Non-Fatal Msg/Error in "+ strErrContextIn+ ": "+ strMsgText, cstrCurrentProc);

      // Display the errror via a modal frmMsgBox window
      //*TODO:** can't found type for with block
      //*With frmMsgBox
      __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = frmMsgBox;
      w___TYPE_NOT_FOUND.ScreenName = m_strScreenName;
      w___TYPE_NOT_FOUND.ErrorCode = lngTranslatedErrNum;
      w___TYPE_NOT_FOUND.MsgText = strMsgText;
      w___TYPE_NOT_FOUND.ErrorContext = strErrContextIn;
      w___TYPE_NOT_FOUND.Show(vbModal);
      _rtn = w___TYPE_NOT_FOUND.ButtonClicked;
      Unload(frmMsgBox);
}

  return _rtn;
}



  public void saveAppSpecificErr(int lngErrNumIn, String strErrContextIn, Object[] varParmsIn) {
    // Comments  : This public subroutine loads this object's public properties
    //             with the specified values. The ErrDesc property is set when
    //             the error is actually reported via ReportFatalError()
    // Parameters: lngErrNumIn (in)     - the error code in non-fatal or fatal error range
    //             strErrContextIn (in) - the module.procname in which the error occurred
    //             varParmsIn (in)      - (optional) a variable number of parameters which
    //                                    will eventually be substituted for  @@1, @@2, etc.
    //                                    placeholders in the message text.
    //
    // Returns   : N/A
    //
    // Modified  :
    // --------------------------------------------------
    int intI = 0;
    int lngTranslatedErrNum = 0;

    #If DEBUG_ERH Then;
    Debug.Print("SaveAppSpecificError called by "+ strErrContextIn+ ". Error#="+ CStr(lngErrNumIn)+ "(or "+ String.valueOf(CStr(m_lngErrNum - vbObjectError))+ ")");
    #End If;

    // Inform developer if error is being recorded prior to reporting a previously
    // recorded error
    Debug.Assert(m_lngErrNum == modResConstants.gCLNGERR_NUM_DEFAULT);

    // Verify we've got an lngErrNumIn in the appropriate range
    lngTranslatedErrNum = lngErrNumIn - vbObjectError;
    Debug.Assert(lngTranslatedErrNum >= modResConstants.gCRES_NERR_START);
    Debug.Assert(lngTranslatedErrNum <= modResConstants.gCRES_FERR_END);

    m_lngErrNum = lngErrNumIn;
    m_strErrDesc = "App-specific error "+ String.valueOf(CStr(lngErrNumIn - vbObjectError));
    m_strErrContext = strErrContextIn;
    m_lngTranslatedErrNum = lngTranslatedErrNum;

    Debug.Print("Recording this error: Error#="+ String.valueOf(CStr(m_lngErrNum - vbObjectError))+ " Context="+ strErrContextIn);
    for (intI = 0; intI <= varParmsIn.length; intI++) {
      addSubstitution(varParmsIn(intI));
    }

    //Err.Raise m_lngErrNum, m_strErrContext, m_strErrDesc  ' NEW 03/11/2002
  }



  //=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
  public boolean saveErrObjectData(String strAppContextIn) {
    boolean _rtn = false;
    // Comments  : This public subroutine loads this object's public properties
    //             with the contents of the Err object *unless* this object's
    //             properties have already been set (and not cleared/reported)
    //             after a previous Err.Raise. This conditional setting of the
    //             properties guards against corruption of the Err object when
    //             errors are passed back up the call chain.
    //
    //             This should be called at the beginning of an error handler
    //             that will be raising an error back up the call chain.
    //
    //             WARNING: Do **NOT** include an On Error or Resume statement
    //                      in this procedure as it will clear the Err object.
    //
    // Parameters: strAppContextIn (in) - indicates the application context in
    //                                  which the error occured (to augment
    //                                  what the Err object automatically has)
    //
    // Returns   : True if the properties were saved; False otherwise
    //
    // Modified  :
    // --------------------------------------------------
    int lngTempErrNum = 0;
    String strTempErrDesc = "";
    String strTempErrContext = "";

    // The following statement calls a function (fnGetSubstitutions)
    // which must use an error handler. However, its use of an error
    // handler resets the Err object, hence we must save the Err object's
    // contents to local variables before calling that function.
    lngTempErrNum = VBA.ex.Number;
    strTempErrDesc = VBA.ex.Description;
    strTempErrContext = VBA.ex.Source;

    try {

      #If DEBUG_ERH Then;
      Debug.Print("SaveErrObjectData called by "+ strAppContextIn+ ". Error#="+ CStr(lngTempErrNum)+ "(or "+ String.valueOf(CStr(lngTempErrNum - vbObjectError))+ ") Desc="+ strTempErrDesc+ " Context="+ strTempErrContext);
      #End If;

      if (m_lngErrNum == modResConstants.gCLNGERR_NUM_DEFAULT  && m_strErrDesc.equals(MCSTRERRDESCDEFAULT) && m_strErrContext.equals(MCSTRERRCONTEXTDEFAULT) && fnGetSubstitutionsUBound() == MCLNGARRAYNOTINITIALIZED) {
        m_lngErrNum = lngTempErrNum;
        m_strErrDesc = strTempErrDesc;
        m_strErrContext = strAppContextIn+ "/"+ strTempErrContext;
        m_lngTranslatedErrNum = lngTempErrNum;
        _rtn = true;
      }
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



///\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
//\                                                                  /
//|                       PRIVATE  Procedures                        |
///                                                                  \
//\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/


//=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
  private String fnGetResString(int intID) {
    String _rtn = "";
    // Comments  : This public subroutine is called when
    //             displaying a message box. It retrieves the text for
    //             the specified error code (intID) from the string table
    //             in the CLAIM.RES resource file.
    // Parameters: intID    = (input) the error code whose description to retrieve
    // Returns   : A string containing the description of the error.
    // Modified  :
    // --------------------------------------------------
    int intDefaultID = 0;

    try {
      _rtn = LoadResString(intID);
      if (VBA.ex) {
        // If the named ID cannot be found, strip off the final digits and append 000
        // (i.e. 1699 becomes 1000) so the default text for that message range can
        // be shown.
        intDefaultID = Integer.parseInt((CStr(intID).substring(0, 1)+ "000"));
        _rtn = LoadResString(intDefaultID);
        if (VBA.ex) {
          //' catch-all
          _rtn = "Error "+ CStr(intID)+ " not found in resource file.";
        }
      }

  }
  try {
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
  private int fnGetSubstitutionsUBound() {
    int _rtn = 0;
    // Comments:   Returns the highest possible index to the
    //             m_Substitutions array.
    //
    // Parameters: None
    // Returns:    The UBound() index, or -1 if the array has not been initialized
    // Called by : GetSubstitution(), AddSubstitution(), ReportFatalError()
    "fnGetSubstitutionsUBound"
.equals(Const cstrCurrentProc As String);

    try {
      _rtn = m_strSubstitutions.length;
      if (VBA.ex.Number == MCLNGSUBSCRIPTOUTOFRANGE) {
        //' Lose the error 9 so the error will be ignored
        VBA.ex.Clear;
        _rtn = MCLNGARRAYNOTINITIALIZED;
      }
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


case class ErherrorhandlerData(
              id: Option[Int],

              )

object Erherrorhandlers extends Controller with ProvidesUser {

  val erherrorhandlerForm = Form(
    mapping(
      "id" -> optional(number),

  )(ErherrorhandlerData.apply)(ErherrorhandlerData.unapply))

  implicit val erherrorhandlerWrites = new Writes[Erherrorhandler] {
    def writes(erherrorhandler: Erherrorhandler) = Json.obj(
      "id" -> Json.toJson(erherrorhandler.id),
      C.ID -> Json.toJson(erherrorhandler.id),

    )
  }

  def get(id: Int) = GetAction { implicit request =>
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.LIST_ERHERRORHANDLER), { user =>
      Ok(Json.toJson(Erherrorhandler.get(user, id)))
    })
  }

  def update(id: Int) = PostAction { implicit request =>
    Logger.debug("in erherrorhandlers.update")
    erherrorhandlerForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      erherrorhandler => {
        Logger.debug(s"form: ${erherrorhandler.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.EDIT_ERHERRORHANDLER), { user =>
          Ok(
            Json.toJson(
              Erherrorhandler.update(user,
                Erherrorhandler(
                       id,

                ))))
        })
      }
    )
  }

  def create = PostAction { implicit request =>
    Logger.debug("in erherrorhandlers.create")
    erherrorhandlerForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      erherrorhandler => {
        Logger.debug(s"form: ${erherrorhandler.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.NEW_ERHERRORHANDLER), { user =>
          Ok(
            Json.toJson(
              Erherrorhandler.create(user,
                Erherrorhandler(

                ))))
        })
      }
    )
  }

  def delete(id: Int) = PostAction { implicit request =>
    Logger.debug("in erherrorhandlers.delete")
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.DELETE_ERHERRORHANDLER), { user =>
      Erherrorhandler.delete(user, id)
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

case class Erherrorhandler(
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

object Erherrorhandler {

  lazy val emptyErherrorhandler = Erherrorhandler(
)

  def apply(
      id: Int,
) = {

    new Erherrorhandler(
      id,
)
  }

  def apply(
) = {

    new Erherrorhandler(
)
  }

  private val erherrorhandlerParser: RowParser[Erherrorhandler] = {
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
        Erherrorhandler(
              id,
,
              createdAt,
              updatedAt,
              updatedBy)
    }
  }

  def create(user: CompanyUser, erherrorhandler: Erherrorhandler): Erherrorhandler = {
    save(user, erherrorhandler, true)
  }

  def update(user: CompanyUser, erherrorhandler: Erherrorhandler): Erherrorhandler = {
    save(user, erherrorhandler, false)
  }

  private def save(user: CompanyUser, erherrorhandler: Erherrorhandler, isNew: Boolean): Erherrorhandler = {
    def getFields = {
      List(

      )
    }
    def throwException = {
      throw new RuntimeException(s"Error when saving ${C.ERHERRORHANDLER}")
    }

    DBHelper.saveEx(
      user,
      Register(
        C.ERHERRORHANDLER,
        C.ID,
        erherrorhandler.id,
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

  def load(user: CompanyUser, id: Int): Option[Erherrorhandler] = {
    loadWhere(user, s"${C.ID} = {id}", 'id -> id)
  }

  def loadWhere(user: CompanyUser, where: String, args : scala.Tuple2[scala.Any, anorm.ParameterValue[_]]*) = {
    DB.withConnection(user.database.database) { implicit connection =>
      SQL(s"SELECT t1.*, t2.${C.FK_NAME} FROM ${C.ERHERRORHANDLER} t1 INNER JOIN ${C.???} t2 ON t1.${C.FK_ID} = t2.${C.FK_ID} WHERE $where")
        .on(args: _*)
        .as(erherrorhandlerParser.singleOpt)
    }
  }

  def delete(user: CompanyUser, id: Int) = {
    DB.withConnection(user.database.database) { implicit connection =>
      try {
        SQL(s"DELETE FROM ${C.ERHERRORHANDLER} WHERE ${C.ID} = {id}")
        .on('id -> id)
        .executeUpdate
      } catch {
        case NonFatal(e) => {
          Logger.error(s"can't delete a ${C.ERHERRORHANDLER}. ${C.ID} id: $id. Error ${e.toString}")
          throw e
        }
      }
    }
  }

  def get(user: CompanyUser, id: Int): Erherrorhandler = {
    load(user, id) match {
      case Some(p) => p
      case None => emptyErherrorhandler
    }
  }
}


// Router

GET     /api/v1/general/erherrorhandler/:id              controllers.logged.modules.general.Erherrorhandlers.get(id: Int)
POST    /api/v1/general/erherrorhandler                  controllers.logged.modules.general.Erherrorhandlers.create
PUT     /api/v1/general/erherrorhandler/:id              controllers.logged.modules.general.Erherrorhandlers.update(id: Int)
DELETE  /api/v1/general/erherrorhandler/:id              controllers.logged.modules.general.Erherrorhandlers.delete(id: Int)




/**/
