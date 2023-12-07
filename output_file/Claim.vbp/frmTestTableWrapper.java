
import java.util.Date;

public class frmTestTableWrapper {

  // % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % %
  //#If False Then


  ///\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
  // PAYEE_T WRAPPER
  ///\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\


  //--------------------------------------------------------------------------------------------
  // Customize for the table wrapper you want to test, by doing:
  //   a. An Edit/Replace All on "ctclmClaim"   (the name of the table wrapper class)
  //   b. Edit the fnDisplayRecord proc so it reflects each table column
  //   c. Look for !CUSTOMIZE! tags for other places where editing is needed.
  //--------------------------------------------------------------------------------------------





//Option Explicit
  *Option Compare Binary

  //--------------------------------------------------------------------------------------------
  // Just need declarations from some maintenance screen
  //--------------------------------------------------------------------------------------------
  private String mstrScreenName = "";

  private static final Long MCLNGMINFORMWIDTH = 4800;
  private static final Long MCLNGMINFORMHEIGHT = 3600;

  //' Make sure this CLM_ID value exists!
  private static final Long MCLNGCLMIDTOUSE = 279742;

  private Scripting.TextStream m_tsOutput;

  private Control mctlFirstUpdateableField_Add;
  private Control mctlFirstUpdateableField_Upd;


  // mtWrapper is an instance of the table wrapper corresponding to the main table maintained by this form.
  private ctpyePayee mtWrapper;




  //////////////////////////////////////////////////////////////////////////////////////////////////
  private void cmdClose_Click() {
    Unload(this);
  }




  //////////////////////////////////////////////////////////////////////////////////////////////////
  private void form_Load() {
    "Form_Load"
.equals(Const cstrCurrentProc As String);
    "C:\\WrapperTest.txt"
.equals(Const cstrOutputFile As String);
    DBRecordSet rst = null;
    DBField fld = null;
    Date dteToday = null;
    Scripting.FileSystemObject fso = null;

    try {

      //--------------------------------------------------------------------------------------------
      // Do not change
      //--------------------------------------------------------------------------------------------
      mstrScreenName = Me.Caption;
      modGeneral.gerhApp.setScreenName(mstrScreenName);
      Me.Icon = LoadResPicture(modResConstants.gCRES_ICON_MAINAPP, vbResIcon);
      //*TODO:** can't found type for with block
      //*With this
      __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = this;
      w___TYPE_NOT_FOUND.Width = MCLNGMINFORMWIDTH;
      w___TYPE_NOT_FOUND.Height = MCLNGMINFORMHEIGHT;
      modGeneral.fnCenterFormOnMDI(frmMDIMain, this);
      mctlFirstUpdateableField_Add = cmdClose;
      mctlFirstUpdateableField_Upd = cmdClose;
      fso = new Scripting.FileSystemObject();
      m_tsOutput = fso.CreateTextFile(cstrOutputFile, true);


      //--------------------------------------------------------------------------------------------
      // 1.  Verify the table wrapper was instantiated successfully.
      //--------------------------------------------------------------------------------------------
      mtWrapper = new ctpyePayee();
      //' -- using 279742 as the CLM_ID
      mtWrapper.initPayee(MCLNGCLMIDTOUSE);

      //!CUSTOMIZE! Include a call to fnDisplayMetaData for each property associated with a table column
      m_tsOutput.WriteLine("----------------------------------------------------");
      m_tsOutput.WriteLine(" 1. Meta Data");
      m_tsOutput.WriteLine("----------------------------------------------------");
      fnDisplayMetaData("CalcStCd");
      fnDisplayMetaData("ClmId");
      fnDisplayMetaData("LstUpdtDtm");
      fnDisplayMetaData("LstUpdtUserId");
      fnDisplayMetaData("PayeAddrLn1Txt");
      fnDisplayMetaData("PayeAddrLn2Txt");
      fnDisplayMetaData("PayeCareOfTxt");
      fnDisplayMetaData("PayeCityNmTxt");
      fnDisplayMetaData("PayeClmIntAmt");
      fnDisplayMetaData("PayeClmIntRt");
      fnDisplayMetaData("PayeDfltOvrdInd");
      fnDisplayMetaData("PayeDthbPmtAmt");
      fnDisplayMetaData("PayeFullNm");
      fnDisplayMetaData("PayeId");
      fnDisplayMetaData("PayeIntDaysPdNum");
      fnDisplayMetaData("PayePmtDt");
      fnDisplayMetaData("PayeSsnTinNum");
      fnDisplayMetaData("PayeSsnTinTypCd");
      fnDisplayMetaData("PayeWthldAmt");
      fnDisplayMetaData("PayeWthldRt");
      fnDisplayMetaData("PayeZip4Cd");
      fnDisplayMetaData("PayeZipCd");

      //--------------------------------------------------------------------------------------------
      // 2.  Verify that the Lookup recordset was successfully populated (right columns, and that the
      //     number of rows retrieved is appropriate)
      //--------------------------------------------------------------------------------------------
      m_tsOutput.WriteLine("\\r\\n");
      m_tsOutput.WriteLine("----------------------------------------------------");
      m_tsOutput.WriteLine(" 2. Lookup Recordset Population");
      m_tsOutput.WriteLine("----------------------------------------------------");
      m_tsOutput.WriteLine("The m_rstLookup recordset was populated as follows: ");
      m_tsOutput.WriteLine("   Record Count = "+ CStr(mtWrapper.getLookupRecordCount()));
      m_tsOutput.WriteLine("   Current Rec# = "+ CStr(mtWrapper.getCurrentLookupRecordNumber()));

      rst = mtWrapper.getLookupData();

      for (int _i = 0; _i < rst.Fields.size(); _i++) {
        fld = rst.Fields.item(_i);
        m_tsOutput.WriteLine("   For the first record in the m_rstLookup recordset:");
        m_tsOutput.WriteLine("     column = "+ fld.Name+ ", value = ["+ CStr(modDataConversion.fnZLSIfNull(fld.value))+ "]");
      }

      //--------------------------------------------------------------------------------------------
      // 3. Test navigation methods
      //--------------------------------------------------------------------------------------------
      m_tsOutput.WriteLine("\\r\\n");
      m_tsOutput.WriteLine("----------------------------------------------------");
      m_tsOutput.WriteLine(" 3. Navigation");
      m_tsOutput.WriteLine("----------------------------------------------------");
      if (mtWrapper.getLookupRecordCount() > 0) {
        mtWrapper.goToLastRecord();
        m_tsOutput.WriteLine("Navigated to last record, rec #"+ ((Integer) mtWrapper.getCurrentLookupRecordNumber()).toString());
        mtWrapper.goToFirstRecord();
        m_tsOutput.WriteLine("Navigated to first record, rec #"+ ((Integer) mtWrapper.getCurrentLookupRecordNumber()).toString());
      }
      if (mtWrapper.getLookupRecordCount() > 1) {
        mtWrapper.goToNextRecord();
        m_tsOutput.WriteLine("Navigated to next record, rec #"+ ((Integer) mtWrapper.getCurrentLookupRecordNumber()).toString());
        mtWrapper.goToPreviousRecord();
        m_tsOutput.WriteLine("Navigated to previous record, rec #"+ ((Integer) mtWrapper.getCurrentLookupRecordNumber()).toString());
      }

      //--------------------------------------------------------------------------------------------
      // 4. Test ADD capability
      //--------------------------------------------------------------------------------------------
      m_tsOutput.WriteLine("\\r\\n");
      m_tsOutput.WriteLine("----------------------------------------------------");
      m_tsOutput.WriteLine(" 4. Add");
      m_tsOutput.WriteLine("----------------------------------------------------");
      //!CUSTOMIZE!
      // Set properties, as if the user input them. Be sure to choose values that
      // should result in a successful Add.
      dteToday = Now;
      mtWrapper.setClmId(MCLNGCLMIDTOUSE);
      mtWrapper.setCalcStCd("FL");
      mtWrapper.setLstUpdtDtm(Now);
      mtWrapper.setLstUpdtUserId("TEST    ");
      mtWrapper.setPayeAddrLn1Txt("ADDRESS LINE 1");
      mtWrapper.setPayeAddrLn2Txt("ADDRESS LINE 2");
      mtWrapper.setPayeCareOfTxt("CARE OF INFO");
      mtWrapper.setPayeCityNmTxt("JACKSONVILLE");
      mtWrapper.setPayeClmIntAmt(17.75);
      mtWrapper.setPayeClmIntRt(3.5);
      mtWrapper.setPayeClmPdAmt(100017.75);
      mtWrapper.setPayeDfltOvrdInd(false);
      mtWrapper.setPayeDthbPmtAmt(10000);
      mtWrapper.setPayeFullNm("BETSY TESTCASE");
      //-- identity .PayeId =
      mtWrapper.setPayeIntDaysPdNum(7);
      mtWrapper.setPayePmtDt(G.parseDate("01/03/2003"));
      mtWrapper.setPayeSsnTinNum(123445555);
      mtWrapper.setPayeSsnTinTypCd("P");
      mtWrapper.setPayeStCd("FL");
      mtWrapper.setPayeWthldAmt(0);
      mtWrapper.setPayeWthldRt(0);
      mtWrapper.setPayeZip4Cd("4444");
      mtWrapper.setPayeZipCd("12345");
      // Call the Add method
      mtWrapper.addRecord();
      // Display the values of the current just-added record
      m_tsOutput.WriteLine("Just added this record: ");
      fnDisplayRecord();

      //--------------------------------------------------------------------------------------------
      // 5. Test UPDATE capability
      //--------------------------------------------------------------------------------------------
      m_tsOutput.WriteLine("\\r\\n");
      m_tsOutput.WriteLine("----------------------------------------------------");
      m_tsOutput.WriteLine(" 5. Update");
      m_tsOutput.WriteLine("----------------------------------------------------");
      //!CUSTOMIZE!
      // Set properties, to simulate the user editing the record just added. Be sure to choose
      // values that should result in a successful Update.  Be sure to NOT change key fields!
      mtWrapper.setPayeIntDaysPdNum(12);
      mtWrapper.setLstUpdtDtm(Now);
      mtWrapper.setLstUpdtUserId("K000    ");
      // Call the Update method
      mtWrapper.updateRecord();
      // Display the values of the current just-updated record
      m_tsOutput.WriteLine("After updating, this record now looks like this: ");
      fnDisplayRecord();

      //--------------------------------------------------------------------------------------------
      // 6. Test DELETE capability
      //--------------------------------------------------------------------------------------------
      m_tsOutput.WriteLine("\\r\\n");
      m_tsOutput.WriteLine("----------------------------------------------------");
      m_tsOutput.WriteLine(" 6. Delete");
      m_tsOutput.WriteLine("----------------------------------------------------");
      // Using properties already set by the Add and Update tests, call the
      // Delete method to delete the record.
      mtWrapper.deleteRecord();
      // Display the values of the record previous to the one just deleted
      m_tsOutput.WriteLine("After deleting, the new current record is now this: ");
      fnDisplayRecord();

      m_tsOutput.Close;
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    modGeneral.fnFreeObject(m_tsOutput);
    modGeneral.fnFreeObject(fso);

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
  private void fnDisplayMetaData(String strTagIn) {
    "fnDisplayMetaData"
.equals(Const cstrCurrentProc As String);
    try {

      m_tsOutput.WriteLine("The "+ strTagIn+ " property: ");
      m_tsOutput.WriteLine("  AllowableCharacters="+ mtWrapper.getAllowableCharacters(strTagIn));
      m_tsOutput.WriteLine("  DefaultValue=["+ CStr(mtWrapper.getDefaultValue(strTagIn))+ "]");
      m_tsOutput.WriteLine("  DollarPositions="+ CStr(mtWrapper.getDollarPositions(strTagIn)));
      m_tsOutput.WriteLine("  DecimalPositions="+ CStr(mtWrapper.getDecimalPositions(strTagIn)));
      m_tsOutput.WriteLine("  IsNullable="+ ((Boolean) mtWrapper.getIsNullable(strTagIn)).toString());
      m_tsOutput.WriteLine("  Format="+ mtWrapper.getFormat(strTagIn));
      m_tsOutput.WriteLine("  Mask="+ mtWrapper.getMask(strTagIn));
      m_tsOutput.WriteLine("  IsKey="+ ((Boolean) mtWrapper.getIsKey(strTagIn)).toString());
      m_tsOutput.WriteLine("  ShouldForceToUppercase="+ ((Boolean) mtWrapper.getShouldForceToUppercase(strTagIn)).toString());
      m_tsOutput.WriteLine("  MaxCharacters="+ CStr(mtWrapper.getMaxCharacters(strTagIn)));
    } catch (Exception ex) {
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
    }
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




//!CUSTOMIZE! Edit this procedure so it lists each column of your table.
  private void fnDisplayRecord() {
    "fnDisplayRecord"
.equals(Const cstrCurrentProc As String);
    try {

      m_tsOutput.WriteLine("PayeId = ["+ ((Integer) mtWrapper.getPayeId()).toString()+ "]");
      m_tsOutput.WriteLine("CalcStCd = ["+ mtWrapper.getCalcStCd()+ "]");
      m_tsOutput.WriteLine("LstUpdtDtm = ["+ mtWrapper.getLstUpdtDtm()+ "]");
      m_tsOutput.WriteLine("LstUpdtUserId = ["+ mtWrapper.getLstUpdtUserId()+ "]");
      m_tsOutput.WriteLine("PayeAddrLn1Txt = ["+ mtWrapper.getPayeAddrLn1Txt()+ "]");
      m_tsOutput.WriteLine("PayeAddrLn2Txt = ["+ mtWrapper.getPayeAddrLn2Txt()+ "]");
      m_tsOutput.WriteLine("PayeCareOfTxt = ["+ mtWrapper.getPayeCareOfTxt()+ "]");
      m_tsOutput.WriteLine("PayeCityNmTxt = ["+ mtWrapper.getPayeCityNmTxt()+ "]");
      m_tsOutput.WriteLine("PayeClmIntAmt = ["+ mtWrapper.getPayeClmIntAmt()+ "]");
      m_tsOutput.WriteLine("PayeClmIntRt = ["+ mtWrapper.getPayeClmIntRt()+ "]");
      m_tsOutput.WriteLine("PayeClmPdAmt = ["+ mtWrapper.getPayeClmPdAmt()+ "]");
      m_tsOutput.WriteLine("PayeDfltOvrdInd = ["+ ((Boolean) mtWrapper.getPayeDfltOvrdInd()).toString()+ "]");
      m_tsOutput.WriteLine("PayeDthbPmtAmt = ["+ mtWrapper.getPayeDthbPmtAmt()+ "]");
      m_tsOutput.WriteLine("PayeFullNm = ["+ mtWrapper.getPayeFullNm()+ "]");
      //'-- identity .PayeId =
      m_tsOutput.WriteLine;
      m_tsOutput.WriteLine("PayeIntDaysPdNum = ["+ ((Integer) mtWrapper.getPayeIntDaysPdNum()).toString()+ "]");
      m_tsOutput.WriteLine("PayePmtDt = ["+ mtWrapper.getPayePmtDt()+ "]");
      m_tsOutput.WriteLine("PayeSsnTinNum = ["+ mtWrapper.getPayeSsnTinNum()+ "]");
      m_tsOutput.WriteLine("PayeSsnTinTypCd = ["+ mtWrapper.getPayeSsnTinTypCd()+ "]");
      m_tsOutput.WriteLine("PayeStCd = ["+ mtWrapper.getPayeStCd()+ "]");
      m_tsOutput.WriteLine("PayeWthldAmt = ["+ mtWrapper.getPayeWthldAmt()+ "]");
      m_tsOutput.WriteLine("PayeWthldRt = ["+ mtWrapper.getPayeWthldRt()+ "]");
      m_tsOutput.WriteLine("PayeZip4Cd = ["+ mtWrapper.getPayeZip4Cd()+ "]");
      m_tsOutput.WriteLine("PayeZipCd = ["+ mtWrapper.getPayeZipCd()+ "]");
    } catch (Exception ex) {
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
    }
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


//#End If
// % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % %






// % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % %
  *#If False Then


  ///\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
  // CLAIM_T WRAPPER
  ///\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\


  //--------------------------------------------------------------------------------------------
  // Customize for the table wrapper you want to test, by doing:
  //   a. An Edit/Replace All on "ctclmClaim"   (the name of the table wrapper class)
  //   b. Edit the fnDisplayRecord proc so it reflects each table column
  //   c. Look for !CUSTOMIZE! tags for other places where editing is needed.
  //--------------------------------------------------------------------------------------------





//Option Explicit
  *Option Compare Binary

  //--------------------------------------------------------------------------------------------
  // Just need declarations from some maintenance screen
  //--------------------------------------------------------------------------------------------
  private String mstrScreenName = "";

  private static final Long MCLNGMINFORMWIDTH = 4800;
  private static final Long MCLNGMINFORMHEIGHT = 3600;

  private Scripting.TextStream m_tsOutput;

  private Control mctlFirstUpdateableField_Add;
  private Control mctlFirstUpdateableField_Upd;
  // mbInLookupMode determines whether the user is in the process of doing a search using the Lookup ComboBox
  private boolean mbInLookupMode = false;

  // mbInAddMode determines whether the user has begun the process of adding a new record to the table.
  // Note that Add mode is independent of Update mode
  private boolean mbInAddMode = false;

  // mtWrapper is an instance of the table wrapper corresponding to the main table maintained by this form.
  private ctclmClaim mtWrapper;

  // m_bIsDirty corresponds to the public property called IsDirty.
  // All maintenance screens should have this field and that property! When True, it indicates
  // that the user has made --but not yet saved-- changes to a record. The MDI form will query
  // this property if the user opens the File menu, since the Exit option should be disabled if
  // any form has outstanding changes.
  // Be sure to use this variable's corresponding Property Let to change its value.
  // Do **NOT** set m_bIsDirty itself, as this will using the Property Let proc will
  // ensure the Close button caption is always synchronized with the value of the property.
  private boolean m_bIsDirty = false;




  //////////////////////////////////////////////////////////////////////////////////////////////////
  private void cmdClose_Click() {
    Unload(this);
  }




  //////////////////////////////////////////////////////////////////////////////////////////////////
  private void form_Load() {
    "Form_Load"
.equals(Const cstrCurrentProc As String);
    Const(cintNoRowsInTable As Integer == 0);
    "C:\\WrapperTest.txt"
.equals(Const cstrOutputFile As String);
    DBRecordSet rst = null;
    DBField fld = null;
    String strProperty1 = "";
    String strProperty2 = "";
    String strProperty3 = "";
    Date dteToday = null;
    Scripting.FileSystemObject fso = null;

    try {

      //--------------------------------------------------------------------------------------------
      // Do not change
      //--------------------------------------------------------------------------------------------
      mstrScreenName = Me.Caption;
      modGeneral.gerhApp.setScreenName(mstrScreenName);
      Me.Icon = LoadResPicture(modResConstants.gCRES_ICON_MAINAPP, vbResIcon);
      //*TODO:** can't found type for with block
      //*With this
      __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = this;
      w___TYPE_NOT_FOUND.Width = MCLNGMINFORMWIDTH;
      w___TYPE_NOT_FOUND.Height = MCLNGMINFORMHEIGHT;
      modGeneral.fnCenterFormOnMDI(frmMDIMain, this);
      mctlFirstUpdateableField_Add = cmdClose;
      mctlFirstUpdateableField_Upd = cmdClose;
      fso = new Scripting.FileSystemObject();
      m_tsOutput = fso.CreateTextFile(cstrOutputFile, true);


      //--------------------------------------------------------------------------------------------
      // 1.  Verify the table wrapper was instantiated successfully.
      //--------------------------------------------------------------------------------------------
      mtWrapper = new ctclmClaim();

      //!CUSTOMIZE! Include a call to fnDisplayMetaData for each property associated with a table column
      m_tsOutput.WriteLine("----------------------------------------------------");
      m_tsOutput.WriteLine(" 1. Meta Data");
      m_tsOutput.WriteLine("----------------------------------------------------");
      fnDisplayMetaData("AdmnSystCd");
      fnDisplayMetaData("ClmId");
      fnDisplayMetaData("ClmInsdDthDt");
      fnDisplayMetaData("ClmInsdFirstNm");
      fnDisplayMetaData("ClmInsdLastNm");
      fnDisplayMetaData("ClmInsdSsnNum");
      fnDisplayMetaData("ClmNum");
      fnDisplayMetaData("ClmPolNum");
      fnDisplayMetaData("ClmProofDt");
      fnDisplayMetaData("ClmTotClmPdAmt");
      fnDisplayMetaData("ClmTotDthbPmtAmt");
      fnDisplayMetaData("ClmTotIntAmt");
      fnDisplayMetaData("ClmTotWthldAmt");
      fnDisplayMetaData("InsdDthResStCd");
      fnDisplayMetaData("IssStCd");
      fnDisplayMetaData("LstUpdtDtm");
      fnDisplayMetaData("LstUpdtUserId");
      fnDisplayMetaData("PycoTypCd");

      //--------------------------------------------------------------------------------------------
      // 2.  Verify that the Lookup recordset was successfully populated (right columns, and that the
      //     number of rows retrieved is appropriate)
      //--------------------------------------------------------------------------------------------
      m_tsOutput.WriteLine("\\r\\n");
      m_tsOutput.WriteLine("----------------------------------------------------");
      m_tsOutput.WriteLine(" 2. Lookup Recordset Population");
      m_tsOutput.WriteLine("----------------------------------------------------");
      m_tsOutput.WriteLine("The m_rstLookup recordset was populated as follows: ");
      m_tsOutput.WriteLine("   Record Count = "+ CStr(mtWrapper.getLookupRecordCount()));
      m_tsOutput.WriteLine("   Current Rec# = "+ CStr(mtWrapper.getCurrentLookupRecordNumber()));

      rst = mtWrapper.getLookupData();

      for (int _i = 0; _i < rst.Fields.size(); _i++) {
        fld = rst.Fields.item(_i);
        m_tsOutput.WriteLine("   For the first record in the m_rstLookup recordset:");
        m_tsOutput.WriteLine("     column = "+ fld.Name+ ", value = ["+ CStr(modDataConversion.fnZLSIfNull(fld.value))+ "]");
      }

      //--------------------------------------------------------------------------------------------
      // 3. Test navigation methods
      //--------------------------------------------------------------------------------------------
      m_tsOutput.WriteLine("\\r\\n");
      m_tsOutput.WriteLine("----------------------------------------------------");
      m_tsOutput.WriteLine(" 3. Navigation");
      m_tsOutput.WriteLine("----------------------------------------------------");
      if (mtWrapper.getLookupRecordCount() > 0) {
        mtWrapper.goToLastRecord();
        m_tsOutput.WriteLine("Navigated to last record, rec #"+ ((Integer) mtWrapper.getCurrentLookupRecordNumber()).toString());
        mtWrapper.goToFirstRecord();
        m_tsOutput.WriteLine("Navigated to first record, rec #"+ ((Integer) mtWrapper.getCurrentLookupRecordNumber()).toString());
      }
      if (mtWrapper.getLookupRecordCount() > 1) {
        mtWrapper.goToNextRecord();
        m_tsOutput.WriteLine("Navigated to next record, rec #"+ ((Integer) mtWrapper.getCurrentLookupRecordNumber()).toString());
        mtWrapper.goToPreviousRecord();
        m_tsOutput.WriteLine("Navigated to previous record, rec #"+ ((Integer) mtWrapper.getCurrentLookupRecordNumber()).toString());
      }

      //--------------------------------------------------------------------------------------------
      // 4. Test ADD capability
      //--------------------------------------------------------------------------------------------
      m_tsOutput.WriteLine("\\r\\n");
      m_tsOutput.WriteLine("----------------------------------------------------");
      m_tsOutput.WriteLine(" 4. Add");
      m_tsOutput.WriteLine("----------------------------------------------------");
      //!CUSTOMIZE!
      // Set properties, as if the user input them. Be sure to choose values that
      // should result in a successful Add.
      dteToday = Now;
      mtWrapper.AdmnSystCd = "02";
      mtWrapper.ClmInsdDthDt = G.parseDate("01/03/2003");
      mtWrapper.ClmInsdFirstNm = "BETSY";
      mtWrapper.ClmInsdLastNm = "TESTCASE";
      mtWrapper.ClmInsdSsnNum = "265802222";
      mtWrapper.ClmNum = "1234567";
      mtWrapper.ClmPolNum = "1234567";
      mtWrapper.ClmProofDt = G.parseDate("01/03/2003");
      mtWrapper.ClmTotClmPdAmt = 0;
      mtWrapper.ClmTotDthbPmtAmt = 0;
      mtWrapper.ClmTotIntAmt = 0;
      mtWrapper.ClmTotWthldAmt = 0;
      mtWrapper.InsdDthResStCd = "FL";
      mtWrapper.IssStCd = "FL";
      mtWrapper.setLstUpdtDtm(Now);
      mtWrapper.setLstUpdtUserId("TEST    ");
      mtWrapper.PycoTypCd = "I1";
      // Call the Add method
      mtWrapper.addRecord();
      // Display the values of the current just-added record
      m_tsOutput.WriteLine("Just added this record: ");
      fnDisplayRecord();

      //--------------------------------------------------------------------------------------------
      // 5. Test UPDATE capability
      //--------------------------------------------------------------------------------------------
      m_tsOutput.WriteLine("\\r\\n");
      m_tsOutput.WriteLine("----------------------------------------------------");
      m_tsOutput.WriteLine(" 5. Update");
      m_tsOutput.WriteLine("----------------------------------------------------");
      //!CUSTOMIZE!
      // Set properties, to simulate the user editing the record just added. Be sure to choose
      // values that should result in a successful Update.  Be sure to NOT change key fields!
      mtWrapper.InsdDthResStCd = "MA";
      mtWrapper.IssStCd = "MA";
      mtWrapper.setLstUpdtDtm(Now);
      mtWrapper.setLstUpdtUserId("K000    ");
      // Call the Update method
      mtWrapper.updateRecord();
      // Display the values of the current just-updated record
      m_tsOutput.WriteLine("After updating, this record now looks like this: ");
      fnDisplayRecord();

      //--------------------------------------------------------------------------------------------
      // 6. Test DELETE capability
      //--------------------------------------------------------------------------------------------
      m_tsOutput.WriteLine("\\r\\n");
      m_tsOutput.WriteLine("----------------------------------------------------");
      m_tsOutput.WriteLine(" 6. Delete");
      m_tsOutput.WriteLine("----------------------------------------------------");
      // Using properties already set by the Add and Update tests, call the
      // Delete method to delete the record.
      mtWrapper.deleteRecord();
      // Display the values of the record previous to the one just deleted
      m_tsOutput.WriteLine("After deleting, the new current record is now this: ");
      fnDisplayRecord();

      m_tsOutput.Close;
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    modGeneral.fnFreeObject(m_tsOutput);
    modGeneral.fnFreeObject(fso);

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
  private void fnDisplayMetaData(String strTagIn) {
    "fnDisplayMetaData"
.equals(Const cstrCurrentProc As String);
    try {

      m_tsOutput.WriteLine("The "+ strTagIn+ " property: ");
      m_tsOutput.WriteLine("  AllowableCharacters="+ mtWrapper.getAllowableCharacters(strTagIn));
      m_tsOutput.WriteLine("  DefaultValue=["+ CStr(mtWrapper.getDefaultValue(strTagIn))+ "]");
      m_tsOutput.WriteLine("  DollarPositions="+ CStr(mtWrapper.getDollarPositions(strTagIn)));
      m_tsOutput.WriteLine("  DecimalPositions="+ CStr(mtWrapper.getDecimalPositions(strTagIn)));
      m_tsOutput.WriteLine("  IsNullable="+ ((Boolean) mtWrapper.getIsNullable(strTagIn)).toString());
      m_tsOutput.WriteLine("  Format="+ mtWrapper.getFormat(strTagIn));
      m_tsOutput.WriteLine("  Mask="+ mtWrapper.getMask(strTagIn));
      m_tsOutput.WriteLine("  IsKey="+ ((Boolean) mtWrapper.getIsKey(strTagIn)).toString());
      m_tsOutput.WriteLine("  ShouldForceToUppercase="+ ((Boolean) mtWrapper.getShouldForceToUppercase(strTagIn)).toString());
      m_tsOutput.WriteLine("  MaxCharacters="+ CStr(mtWrapper.getMaxCharacters(strTagIn)));
    } catch (Exception ex) {
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
    }
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




//!CUSTOMIZE! Edit this procedure so it lists each column of your table.
  private void fnDisplayRecord() {
    "fnDisplayRecord"
.equals(Const cstrCurrentProc As String);
    try {

      m_tsOutput.WriteLine("ClmId = ["+ ((Integer) mtWrapper.getClmId()).toString()+ "]");
      m_tsOutput.WriteLine("AdmnSystCd = ["+ mtWrapper.AdmnSystCd+ "]");
      m_tsOutput.WriteLine("ClmInsdDthDt = ["+ mtWrapper.ClmInsdDthDt+ "]");
      m_tsOutput.WriteLine("ClmInsdFirstNm = ["+ mtWrapper.ClmInsdFirstNm+ "]");
      m_tsOutput.WriteLine("ClmInsdLastNm = ["+ mtWrapper.ClmInsdLastNm+ "]");
      m_tsOutput.WriteLine("ClmInsdSsnNum = ["+ mtWrapper.ClmInsdSsnNum+ "]");
      m_tsOutput.WriteLine("ClmNum = ["+ mtWrapper.ClmNum+ "]");
      m_tsOutput.WriteLine("ClmPolNum = ["+ mtWrapper.ClmPolNum+ "]");
      m_tsOutput.WriteLine("ClmProofDt = ["+ mtWrapper.ClmProofDt+ "]");
      m_tsOutput.WriteLine("ClmTotClmPdAmt = ["+ mtWrapper.ClmTotClmPdAmt+ "]");
      m_tsOutput.WriteLine("ClmTotDthbPmtAmt = ["+ mtWrapper.ClmTotDthbPmtAmt+ "]");
      m_tsOutput.WriteLine("ClmTotIntAmt = ["+ mtWrapper.ClmTotIntAmt+ "]");
      m_tsOutput.WriteLine("ClmTotWthldAmt = ["+ mtWrapper.ClmTotWthldAmt+ "]");
      m_tsOutput.WriteLine("InsdDthResStCd = ["+ mtWrapper.InsdDthResStCd+ "]");
      m_tsOutput.WriteLine("IssStCd = ["+ mtWrapper.IssStCd+ "]");
      m_tsOutput.WriteLine("LstUpdtDtm = ["+ mtWrapper.getLstUpdtDtm()+ "]");
      m_tsOutput.WriteLine("LstUpdtUserId = ["+ mtWrapper.getLstUpdtUserId()+ "]");
      m_tsOutput.WriteLine("PycoTypCd = ["+ mtWrapper.PycoTypCd+ "]");
    } catch (Exception ex) {
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
    }
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


  *#End If
  // % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % %
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


case class RmtesttablewrapperData(
              id: Option[Int],

              )

object Rmtesttablewrappers extends Controller with ProvidesUser {

  val rmtesttablewrapperForm = Form(
    mapping(
      "id" -> optional(number),

  )(RmtesttablewrapperData.apply)(RmtesttablewrapperData.unapply))

  implicit val rmtesttablewrapperWrites = new Writes[Rmtesttablewrapper] {
    def writes(rmtesttablewrapper: Rmtesttablewrapper) = Json.obj(
      "id" -> Json.toJson(rmtesttablewrapper.id),
      C.ID -> Json.toJson(rmtesttablewrapper.id),

    )
  }

  def get(id: Int) = GetAction { implicit request =>
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.LIST_RMTESTTABLEWRAPPER), { user =>
      Ok(Json.toJson(Rmtesttablewrapper.get(user, id)))
    })
  }

  def update(id: Int) = PostAction { implicit request =>
    Logger.debug("in rmtesttablewrappers.update")
    rmtesttablewrapperForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      rmtesttablewrapper => {
        Logger.debug(s"form: ${rmtesttablewrapper.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.EDIT_RMTESTTABLEWRAPPER), { user =>
          Ok(
            Json.toJson(
              Rmtesttablewrapper.update(user,
                Rmtesttablewrapper(
                       id,

                ))))
        })
      }
    )
  }

  def create = PostAction { implicit request =>
    Logger.debug("in rmtesttablewrappers.create")
    rmtesttablewrapperForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      rmtesttablewrapper => {
        Logger.debug(s"form: ${rmtesttablewrapper.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.NEW_RMTESTTABLEWRAPPER), { user =>
          Ok(
            Json.toJson(
              Rmtesttablewrapper.create(user,
                Rmtesttablewrapper(

                ))))
        })
      }
    )
  }

  def delete(id: Int) = PostAction { implicit request =>
    Logger.debug("in rmtesttablewrappers.delete")
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.DELETE_RMTESTTABLEWRAPPER), { user =>
      Rmtesttablewrapper.delete(user, id)
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

case class Rmtesttablewrapper(
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

object Rmtesttablewrapper {

  lazy val emptyRmtesttablewrapper = Rmtesttablewrapper(
)

  def apply(
      id: Int,
) = {

    new Rmtesttablewrapper(
      id,
)
  }

  def apply(
) = {

    new Rmtesttablewrapper(
)
  }

  private val rmtesttablewrapperParser: RowParser[Rmtesttablewrapper] = {
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
        Rmtesttablewrapper(
              id,
,
              createdAt,
              updatedAt,
              updatedBy)
    }
  }

  def create(user: CompanyUser, rmtesttablewrapper: Rmtesttablewrapper): Rmtesttablewrapper = {
    save(user, rmtesttablewrapper, true)
  }

  def update(user: CompanyUser, rmtesttablewrapper: Rmtesttablewrapper): Rmtesttablewrapper = {
    save(user, rmtesttablewrapper, false)
  }

  private def save(user: CompanyUser, rmtesttablewrapper: Rmtesttablewrapper, isNew: Boolean): Rmtesttablewrapper = {
    def getFields = {
      List(

      )
    }
    def throwException = {
      throw new RuntimeException(s"Error when saving ${C.RMTESTTABLEWRAPPER}")
    }

    DBHelper.saveEx(
      user,
      Register(
        C.RMTESTTABLEWRAPPER,
        C.ID,
        rmtesttablewrapper.id,
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

  def load(user: CompanyUser, id: Int): Option[Rmtesttablewrapper] = {
    loadWhere(user, s"${C.ID} = {id}", 'id -> id)
  }

  def loadWhere(user: CompanyUser, where: String, args : scala.Tuple2[scala.Any, anorm.ParameterValue[_]]*) = {
    DB.withConnection(user.database.database) { implicit connection =>
      SQL(s"SELECT t1.*, t2.${C.FK_NAME} FROM ${C.RMTESTTABLEWRAPPER} t1 INNER JOIN ${C.???} t2 ON t1.${C.FK_ID} = t2.${C.FK_ID} WHERE $where")
        .on(args: _*)
        .as(rmtesttablewrapperParser.singleOpt)
    }
  }

  def delete(user: CompanyUser, id: Int) = {
    DB.withConnection(user.database.database) { implicit connection =>
      try {
        SQL(s"DELETE FROM ${C.RMTESTTABLEWRAPPER} WHERE ${C.ID} = {id}")
        .on('id -> id)
        .executeUpdate
      } catch {
        case NonFatal(e) => {
          Logger.error(s"can't delete a ${C.RMTESTTABLEWRAPPER}. ${C.ID} id: $id. Error ${e.toString}")
          throw e
        }
      }
    }
  }

  def get(user: CompanyUser, id: Int): Rmtesttablewrapper = {
    load(user, id) match {
      case Some(p) => p
      case None => emptyRmtesttablewrapper
    }
  }
}


// Router

GET     /api/v1/general/rmtesttablewrapper/:id              controllers.logged.modules.general.Rmtesttablewrappers.get(id: Int)
POST    /api/v1/general/rmtesttablewrapper                  controllers.logged.modules.general.Rmtesttablewrappers.create
PUT     /api/v1/general/rmtesttablewrapper/:id              controllers.logged.modules.general.Rmtesttablewrappers.update(id: Int)
DELETE  /api/v1/general/rmtesttablewrapper/:id              controllers.logged.modules.general.Rmtesttablewrappers.delete(id: Int)




/**/
