public class modRegistrySettings {

  //******************************************************************************
  // Module     : modRegistryAndSettings
  // Description: This standard module contains procedures relating to
  //              retrieving and setting application-related values
  //              from/to the registry and global variables.
  // Procedures :
  //              fnGetDriveFromPath(ByVal strPath As String)
  //              fnIsDBPathValid() As Boolean
  //              fnRegGetAppSettings()
  //              fnRegGetClerkCode() As String
  //              fnRegGetDBName() As String
  //              fnRegGetDBPath() As String
  //              fnRegInitializeForApp()
  //              fnRegSetClerkCode(ByVal strClerkCode As String)
  //              fnRegSetDBName(ByVal strDBName As String)
  //              fnRegSetDBPath(ByVal strDBPath As String)
  // Modified   :
  // 03/26/01 BAW Cleaned with Total Visual CodeTools 2000
  // 01/2002  BAW Optimized per Project Analyzer (removing dead code, adding "$" to Mid/Space, etc.).
  //              Changed all calls to procs in CRegSettings to include a "Root" parameter, indicating
  //              whether HKLM or HKCU should be accessed. This is so any registry writes are done
  //              to HKCU, so they'll be successful on a Win2K PC where the user is non-Administrator.
  // --------------------------------------------------
//Option Explicit
  *Option Compare Binary

  private static final String MCSTRNAME = "modRegistryAndSettings.";

  //-----------------------------------------------------------------------
  // The following are used by procedures interacting with cRegSettings
  // to read/write items in the registry. They should always "jive" with
  // what the Claims Interest setup programs uses!
  //-----------------------------------------------------------------------
  public static final String GCREGCOMPANY = "Sun Life Financial";
  public static final String GCREGAPPNAME = "Claims Interest";
  // Default values of 'missing' registry settings
  public static final String GCEMPTY = "EMPTY";
  *Public Const gcDefaultDBPath = "L:\CLAIMSINTEREST"

  //-----------------------------------------------------------------------
  // The following are used throughout the app when it needs to
  // determine the name of the database or its location (per registry
  // settings for the app). If the registry's entry re: DBName doesn't exist,
  // then the value of gcDefaultDBName will be used as the database name.
  //-----------------------------------------------------------------------
  //Public gstrPath As String
  public static String gstrDBPath = "";
  public static String gstrDBName = "";
  public static String gstrDBPathAndName = "";
  public static String gstrClerkCode = "";
  public static final String GCDEFAULTDBNAME = "CLAIMS.MDB";
  *Public Const gcClaimsManagerClerkCode = "A1GMW"
  //' old one = purple
  *Public Const gcstrPassword = "ireland"

  public static CRegSettings gcReg;

  // 01/31/2002 BAW - Added the following constants to support having default values in HKLM but per-user settings in HKCU
  private static final int HKEY_CURRENT_USER = 0x80000001;
  private static final int HKEY_LOCAL_MACHINE = 0x80000002;



  //////////////////////////////////////////////////////////////////////////////////////////////////
  public static String fnGetDriveFromPath(String strPath) {
    String _rtn = "";
    // Comments  : Returns the drive letter part of the path
    // Parameters: pstrPath - path containing the drive letter
    // Returns   : the drive letter
    // Source    : Total Visual SourceBook 2000
    //
    // Called by : Form_Load( ) in frmSetDatabaseLocation
    //
    int intPos = 0;
    String strTmp = "";
    ":\\"
.equals(Const cstrDelimiter As String);
    "fnGetDriveFromPath"
.equals(Const cstrCurrentProc As String);

    try {

      // Initialize the return value
      strTmp = "";

      // See of the colon and backslash exist
      intPos = strPath.indexOf(cstrDelimiter);

      if (intPos > 0) {
        // They exist, so return the remainder
        strTmp = strPath.substring(0, intPos);
      } 
      else {
        // Look for the colon
        intPos = strPath.indexOf(":");
        if (intPos > 0) {
          // It exists so return the remainder
          strTmp = strPath.substring(0, intPos);
        } 
        else {
          // No drive letter information, so return a zero-length string
          strTmp = "";
        }
      }

      _rtn = strTmp;
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
  public static boolean fnIsDBPathValid() {
    boolean _rtn = false;
    //----------------------------------------------------------------------------
    // Procedure   :  Function fnIsDBPathValid
    // Created by  :  BAW on 04-26-2001 11:18
    //
    // Comments    :
    // Called by   : Form_Load( ) in frmSetDatabaseLocation
    //               Main( ) in modStartup
    //
    // Parameters  : None
    //
    // Return value: True if global vars (gstrDBPath, gstrDBName, gstrDBPathAndName)
    //               point to a valid location in which CLAIMS.MDB exists
    // Modified     :
    //----------------------------------------------------------------------------
    try {
      "fnIsDBPathValid"
.equals(Const cstrCurrentProc As String);
      FileSystemObject fso = null;

      _rtn = true;

      if ((gstrDBPath.equals(GCEMPTY)) || (gstrDBName.equals(GCEMPTY))) {
        _rtn = false;
      } 
      else {
        fso = CreateObject("Scripting.FileSystemObject");
        if (!(fso.FileExists(gstrDBPathAndName))) {
          _rtn = false;
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
  public static void fnRegGetAppSettings() {
    // Comments  : Retrieves current app settings from registry, e.g,
    //             DBPath, DBName, ClerkCode and stores them in their
    //             corresponding global variables (gstrDBPath,
    //             gstrDBName and gstrClerkCode)
    //
    // Called by : Sub Main( ) in modStartup
    // Parameters: None
    // Modified  :
    // --------------------------------------------------
    try {
      "fnRegGetAppSettings"
.equals(Const cstrCurrentProc As String);

      fnRegGetDBPath();
      fnRegGetDBName();
      fnRegGetClerkCode();

      gstrDBPathAndName = gstrDBPath+ "\\"+ gstrDBName;
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
  public static void fnRegGetClerkCode() {
    // Comments  : Retrieves the Clerk Code app-related entry from HKCU in the registry:
    //             Sun Life Financial\Claims Interest\System\ClerkCode, building a
    //             default value for that key if it wasn't found.
    // Parameters: None
    // Returns   : N/A
    // Modified  :
    //   01/31/02 BAW  Changed to retrieve the key from HKCU if possible. Otherwise, get
    //                 the default value from HKLM, if present, and save it to HKCU.
    //                 This is so the app no longer writes to registry keys that may
    //                 not be accessible to a non-Administrator user under Win2K.
    // Called by : fnRegGetAppSettings( ) in modRegistrySettings
    // --------------------------------------------------
    try {
      "fnRegGetClerkCode"
.equals(Const cstrCurrentProc As String);

      gstrClerkCode = gcReg.getSetting(HKEY_CURRENT_USER, "System", "ClerkCode", GCEMPTY);

      if (gstrClerkCode.equals(GCEMPTY)) {
        gstrClerkCode = gcReg.getSetting(HKEY_LOCAL_MACHINE, "System", "ClerkCode", GCEMPTY);
        if (gstrClerkCode.equals(GCEMPTY)) {
          // Build default key in HKLM
          fnRegSetClerkCode(gcClaimsManagerClerkCode, HKEY_LOCAL_MACHINE);
        }
        // Build key in HKCU
        fnRegSetClerkCode(gstrClerkCode);
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



//////////////////////////////////////////////////////////////////////////////////////////////////
  public static void fnRegGetDBName() {
    // Comments  : Retrieves the DBName app-related entry from HKCU in the registry:
    //             Sun Life Financial\Claims Interest\System\DBName, building a
    //             default value for that key if it wasn't found.
    // Parameters: None
    // Returns   : a string containing the value of that registry key
    // Modified  :
    //   01/31/02 BAW  Changed to retrieve the key from HKCU if possible. Otherwise, get
    //                 the default value from HKLM, if present, and save it to HKCU.
    //                 This is so the app no longer writes to registry keys that may
    //                 not be accessible to a non-Administrator user under Win2K.
    //
    // Called by : fnRegGetAppSettings( ) in modRegistrySettings
    // --------------------------------------------------
    try {
      "fnRegGetDBName"
.equals(Const cstrCurrentProc As String);

      gstrDBName = gcReg.getSetting(HKEY_CURRENT_USER, "System", "DBName", GCEMPTY);

      if (gstrDBName.equals(GCEMPTY)) {
        gstrDBName = gcReg.getSetting(HKEY_LOCAL_MACHINE, "System", "DBName", GCEMPTY);
        if (gstrDBName.equals(GCEMPTY)) {
          // Build default key in HKLM
          fnRegSetDBName(GCDEFAULTDBNAME, HKEY_LOCAL_MACHINE);
        }
        // Build key in HKCU
        fnRegSetDBName(gstrDBName);
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



//////////////////////////////////////////////////////////////////////////////////////////////////
  public static void fnRegGetDBPath() {
    // Comments  : Retrieves the DBPath app-related entry from HKCU in the registry:
    //             Sun Life Financial\Claims Interest\System\DBPath, building a
    //             default value for that key if it wasn't found.
    // Parameters: None
    // Modified  :
    //   01/31/02 BAW  Changed to retrieve the key from HKCU if possible. Otherwise, get
    //                 the default value from HKLM, if present, and save it to HKCU.
    //                 This is so the app no longer writes to registry keys that may
    //                 not be accessible to a non-Administrator user under Win2K.
    //
    // Called by : fnRegGetAppSettings( ) in modRegistrySettings
    // --------------------------------------------------
    try {
      "fnRegGetDBPath"
.equals(Const cstrCurrentProc As String);

      gstrDBPath = gcReg.getSetting(HKEY_CURRENT_USER, "System", "DBPath", GCEMPTY);

      if (gstrDBPath.equals(GCEMPTY)) {
        gstrDBPath = gcReg.getSetting(HKEY_LOCAL_MACHINE, "System", "DBPath", GCEMPTY);
        if (gstrDBPath.equals(GCEMPTY)) {
          // Build default key in HKLM
          fnRegSetDBPath(gcDefaultDBPath, HKEY_LOCAL_MACHINE);
        }
        // Store key in HKCU
        fnRegSetDBPath(gstrDBPath);
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



//////////////////////////////////////////////////////////////////////////////////////////////////
  public static void fnRegInitializeForApp() {
    // Comments  : Initializes the CRegSettings object with
    //             app-specific values for Company Name and App Name.
    // Called by : Sub Main() in modStartup
    // Parameters: None
    // Modified  :
    // Called by : fnRegInitializeForApp( ) in modRegistrySettings
    // --------------------------------------------------
    try {
      "fnRegInitializeForApp"
.equals(Const cstrCurrentProc As String);

      // Establish and initialize a global object pointer to an instance of a
      // CRegSettings class whose methods will be used to read/write
      // registry settings.
      gcReg = new CRegSettings();
      gcReg.setCompany(GCREGCOMPANY);
      gcReg.setAppName(GCREGAPPNAME);
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
  public static void fnRegSetClerkCode(String strClerkCode, int root) {
    // Comments  : Stores the Clerk Code in the appropriate app-related
    //             entry in the registry:
    //             Sun Life Financial\Claims Interest\System\ClerkCode
    // Called by : cmdUpdate_Click( ) in frmInsured
    //             fnRegGetAppSettings( ) in modRegistrySettings
    // Parameters: strClerkCode, the value to store
    // Modified  :
    // --------------------------------------------------
    try {
      "fnRegSetClerkCode"
.equals(Const cstrCurrentProc As String);

      gcReg.saveSetting(root, "System", "ClerkCode", strClerkCode.toUpperCase());
      gstrClerkCode = strClerkCode;
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
  public static void fnRegSetDBName(String strDBName, int root) {
    // Comments  : Stores the DBName in the appropriate app-related
    //             entry in the registry:
    //             Sun Life Financial\Claims Interest\System\DBName
    //             and updates related global variables to ensure
    //             they're always kept in synch.
    // Called by : cmdApply_Click( ) in frmSetDatabaseLocation
    //             fnRegGetAppSettings( ) in modRegistrySettings
    //
    // Parameters: strDBName, the value to store
    // Modified  :
    // --------------------------------------------------
    try {
      "fnRegSetDBName"
.equals(Const cstrCurrentProc As String);

      strDBName = strDBName.toUpperCase();
      gcReg.saveSetting(root, "System", "DBName", strDBName);
      gstrDBName = strDBName;
      gstrDBPathAndName = gstrDBPath+ "\\"+ gstrDBName;
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
  public static void fnRegSetDBPath(String strDBPath, int root) {
    // Comments  : Stores the DBPath in the appropriate app-related
    //             entry in the registry:
    //             Sun Life Financial\Claims Interest\System\DBPath
    //             NOTE: A trailing slash will be deleted if it exists.
    //
    // Called by : cmdApply_Click( ) in frmSetDatabaseLocation
    //             fnRegGetAppSettings( ) in modRegistrySettings
    //
    // Parameters: strDBPath, the value to store
    // Modified  :
    // --------------------------------------------------
    try {
      "fnRegSetDBPath"
.equals(Const cstrCurrentProc As String);

      if ("\\"
.equals(strDBPath.substring(strDBPath.length() - 1))) {
        strDBPath = strDBPath.substring(0, strDBPath.length() - 1);
      }

      strDBPath = strDBPath.toUpperCase();
      gcReg.saveSetting(root, "System", "DBPath", strDBPath);
      gstrDBPath = strDBPath;
      gstrDBPathAndName = gstrDBPath+ "\\"+ gstrDBName;
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


case class OdregistrysettingsData(
              id: Option[Int],

              )

object Odregistrysettingss extends Controller with ProvidesUser {

  val odregistrysettingsForm = Form(
    mapping(
      "id" -> optional(number),

  )(OdregistrysettingsData.apply)(OdregistrysettingsData.unapply))

  implicit val odregistrysettingsWrites = new Writes[Odregistrysettings] {
    def writes(odregistrysettings: Odregistrysettings) = Json.obj(
      "id" -> Json.toJson(odregistrysettings.id),
      C.ID -> Json.toJson(odregistrysettings.id),

    )
  }

  def get(id: Int) = GetAction { implicit request =>
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.LIST_ODREGISTRYSETTINGS), { user =>
      Ok(Json.toJson(Odregistrysettings.get(user, id)))
    })
  }

  def update(id: Int) = PostAction { implicit request =>
    Logger.debug("in odregistrysettingss.update")
    odregistrysettingsForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      odregistrysettings => {
        Logger.debug(s"form: ${odregistrysettings.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.EDIT_ODREGISTRYSETTINGS), { user =>
          Ok(
            Json.toJson(
              Odregistrysettings.update(user,
                Odregistrysettings(
                       id,

                ))))
        })
      }
    )
  }

  def create = PostAction { implicit request =>
    Logger.debug("in odregistrysettingss.create")
    odregistrysettingsForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      odregistrysettings => {
        Logger.debug(s"form: ${odregistrysettings.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.NEW_ODREGISTRYSETTINGS), { user =>
          Ok(
            Json.toJson(
              Odregistrysettings.create(user,
                Odregistrysettings(

                ))))
        })
      }
    )
  }

  def delete(id: Int) = PostAction { implicit request =>
    Logger.debug("in odregistrysettingss.delete")
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.DELETE_ODREGISTRYSETTINGS), { user =>
      Odregistrysettings.delete(user, id)
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

case class Odregistrysettings(
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

object Odregistrysettings {

  lazy val emptyOdregistrysettings = Odregistrysettings(
)

  def apply(
      id: Int,
) = {

    new Odregistrysettings(
      id,
)
  }

  def apply(
) = {

    new Odregistrysettings(
)
  }

  private val odregistrysettingsParser: RowParser[Odregistrysettings] = {
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
        Odregistrysettings(
              id,
,
              createdAt,
              updatedAt,
              updatedBy)
    }
  }

  def create(user: CompanyUser, odregistrysettings: Odregistrysettings): Odregistrysettings = {
    save(user, odregistrysettings, true)
  }

  def update(user: CompanyUser, odregistrysettings: Odregistrysettings): Odregistrysettings = {
    save(user, odregistrysettings, false)
  }

  private def save(user: CompanyUser, odregistrysettings: Odregistrysettings, isNew: Boolean): Odregistrysettings = {
    def getFields = {
      List(

      )
    }
    def throwException = {
      throw new RuntimeException(s"Error when saving ${C.ODREGISTRYSETTINGS}")
    }

    DBHelper.saveEx(
      user,
      Register(
        C.ODREGISTRYSETTINGS,
        C.ID,
        odregistrysettings.id,
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

  def load(user: CompanyUser, id: Int): Option[Odregistrysettings] = {
    loadWhere(user, s"${C.ID} = {id}", 'id -> id)
  }

  def loadWhere(user: CompanyUser, where: String, args : scala.Tuple2[scala.Any, anorm.ParameterValue[_]]*) = {
    DB.withConnection(user.database.database) { implicit connection =>
      SQL(s"SELECT t1.*, t2.${C.FK_NAME} FROM ${C.ODREGISTRYSETTINGS} t1 INNER JOIN ${C.???} t2 ON t1.${C.FK_ID} = t2.${C.FK_ID} WHERE $where")
        .on(args: _*)
        .as(odregistrysettingsParser.singleOpt)
    }
  }

  def delete(user: CompanyUser, id: Int) = {
    DB.withConnection(user.database.database) { implicit connection =>
      try {
        SQL(s"DELETE FROM ${C.ODREGISTRYSETTINGS} WHERE ${C.ID} = {id}")
        .on('id -> id)
        .executeUpdate
      } catch {
        case NonFatal(e) => {
          Logger.error(s"can't delete a ${C.ODREGISTRYSETTINGS}. ${C.ID} id: $id. Error ${e.toString}")
          throw e
        }
      }
    }
  }

  def get(user: CompanyUser, id: Int): Odregistrysettings = {
    load(user, id) match {
      case Some(p) => p
      case None => emptyOdregistrysettings
    }
  }
}


// Router

GET     /api/v1/general/odregistrysettings/:id              controllers.logged.modules.general.Odregistrysettingss.get(id: Int)
POST    /api/v1/general/odregistrysettings                  controllers.logged.modules.general.Odregistrysettingss.create
PUT     /api/v1/general/odregistrysettings/:id              controllers.logged.modules.general.Odregistrysettingss.update(id: Int)
DELETE  /api/v1/general/odregistrysettings/:id              controllers.logged.modules.general.Odregistrysettingss.delete(id: Int)




/**/
