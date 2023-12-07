public class CRegSettings {

  // --------------------       Modification History       --------------------
  //
  //  01/2002 BAW - Added "$" to some functions like Trim to optimize it, plus
  //                other minor optimizations per Project Analyzer
  //                Changed all calls to procs in CRegSettings to include a "Root"
  //                parameter, indicating whether HKLM or HKCU should be accessed.
  //                This is so any registry writes are done to HKCU, so they'll
  //                be successful on a Win2K PC where the user is non-Administrator.
  // --------------------------------------------------------------------------


  // *********************************************************************
  //  Copyright ©1997-99 Karl E. Peterson, All Rights Reserved
  // *********************************************************************
  //  You are free to use this code within your own applications, but you
  //  are expressly forbidden from selling or otherwise distributing this
  //  source code without prior written consent.
  // *********************************************************************
//Option Explicit
  *Option Compare Binary

  private static final String MCSTRNAME = "cRegSettings.";

  //
  // Win32 Registry functions
  //
*TODO: API Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
*TODO: API Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
*TODO: API Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
*TODO: API Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As Any) As Long
*TODO: API Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
*TODO: API Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
*TODO: API Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
*TODO: API Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
*TODO: API Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
*TODO: API Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
  //
  // Constants for Windows 32-bit Registry API
  //
  private static final int HKEY_CLASSES_ROOT = 0x80000000;
  private static final int HKEY_CURRENT_USER = 0x80000001;
  private static final int HKEY_LOCAL_MACHINE = 0x80000002;
  private static final int HKEY_USERS = 0x80000003;
  private static final int HKEY_PERFORMANCE_DATA = 0x80000004;
  private static final int HKEY_CURRENT_CONFIG = 0x80000005;
  private static final int HKEY_DYN_DATA = 0x80000006;
  //
  // Reg result codes
  //
  //' New Registry Key created
  private static final int REG_CREATED_NEW_KEY = 0x1;
  //' Existing Key opened
  private static final int REG_OPENED_EXISTING_KEY = 0x2;
  //
  // Reg Create Type Values...
  //
  //' Parameter is reserved
  private static final int REG_OPTION_RESERVED = 0;
  //' Key is preserved when system is rebooted
  private static final int REG_OPTION_NON_VOLATILE = 0;
  //' Key is not preserved when system is rebooted
  private static final int REG_OPTION_VOLATILE = 1;
  //' Created key is a symbolic link
  private static final int REG_OPTION_CREATE_LINK = 2;
  //' open for backup or restore
  private static final int REG_OPTION_BACKUP_RESTORE = 4;
  //
  // Reg Key Security Options
  //
  private static final int DELETE = 0x10000;
  private static final int READ_CONTROL = 0x20000;
  private static final int WRITE_DAC = 0x40000;
  private static final int WRITE_OWNER = 0x80000;
  private static final int SYNCHRONIZE = 0x100000;
  *TODO:** (the data type can't be found for the value [(READ_CONTROL)])Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
  *TODO:** (the data type can't be found for the value [(READ_CONTROL)])Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
  *TODO:** (the data type can't be found for the value [(READ_CONTROL)])Private Const STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
  private static final int STANDARD_RIGHTS_REQUIRED = 0xF0000;
  private static final int STANDARD_RIGHTS_ALL = 0x1F0000;
  private static final int SPECIFIC_RIGHTS_ALL = 0xFFFF;
  private static final int KEY_QUERY_VALUE = 0x1;
  private static final int KEY_SET_VALUE = 0x2;
  private static final int KEY_CREATE_SUB_KEY = 0x4;
  private static final int KEY_ENUMERATE_SUB_KEYS = 0x8;
  private static final int KEY_NOTIFY = 0x10;
  private static final int KEY_CREATE_LINK = 0x20;
  private static final ((STANDARD_RIGHTS_READ KEY_READ = KEY_QUERY_VALUE;
  private static final ((STANDARD_RIGHTS_WRITE KEY_WRITE = KEY_SET_VALUE;
  private static final ((STANDARD_RIGHTS_ALL KEY_ALL_ACCESS = KEY_QUERY_VALUE;
  private static final ((KEY_READ) KEY_EXECUTE = (Not;

  private static final int ERROR_SUCCESS = 0&;
  private static final int ERROR_MORE_DATA = 234;
  private static final int ERROR_NO_MORE_ITEMS = 259;

  //' Unicode nul terminated string
  private static final int REG_SZ = 1;
  //
  // Private member variables
  //
  private String m_company = "";
  private String m_appName = "";
  //
  // Private class constants
  //
  private static final String DEFCOMPANY = "VB and VBA Program Settings";

  // ********************************************
  //  Initialize and Terminate
  // ********************************************
  private void class_Initialize() {
    m_company = DEFCOMPANY;
    m_appName = App.ProductName;
  }

  // ********************************************
  //  Public Properties
  // ********************************************
  public void setCompany(String newVal) {
    // Called by : fnRegInitializeForApp( ) in modRegistrySettings
    if (newVal.length()) {
      m_company = newVal.trim();
    } 
    else {
      m_company = DEFCOMPANY;
    }
  }




  public void setAppName(String newVal) {
    // Called by : fnRegGetAppSettings( ) in modRegistrySettings
    if (newVal.length()) {
      m_appName = newVal.trim();
    } 
    else {
      m_appName = App.ProductName;
    }
  }




  // ********************************************
  //  Public Methods
  // ********************************************



  ////////////////////////////////////////////////////////////////////////////////
  public String getSetting(int root, String section, String key, String default) {
    String _rtn = "";
    // Section   Required. String expression containing the name of the section where the key setting is found.
    //           If omitted, key setting is assumed to be in default subkey.
    // Key       Required. String expression containing the name of the key setting to return.
    // Default   Optional. Expression containing the value to return if no value is set in the key setting.
    //           If omitted, default is assumed to be a zero-length string ("").
    // Called by fnRegGetClerkCode( ) in modRegistrySettings
    //           fnRegGetDBName( ) in modRegistrySettings
    //           fnRegGetDBPath( ) in modRegistrySettings
    "GetSettings"
.equals(Const cstrCurrentProc As String);
    try {

      int nRet = 0;
      int hKey = 0;
      int nType = 0;
      int nBytes = 0;
      String buffer = "";

      // Assume failure and set return to Default
      _rtn = default;

      // Open key
      nRet = RegOpenKeyEx(root, subKey(section), 0&, KEY_ALL_ACCESS, hKey);
      if (nRet == ERROR_SUCCESS) {
        // Set appropriate value for default query
        if (key.equals("*")) {
          key = "";
        }

        // Determine how large the buffer needs to be
        nRet = RegQueryValueEx(hKey, key, 0&, nType, ByVal buffer, nBytes);
        if (nRet == ERROR_SUCCESS) {
          // Build buffer and get data
          if (nBytes > 0) {
            buffer = Space$(nBytes);
            nRet = RegQueryValueEx(hKey, key, 0&, nType, ByVal buffer, buffer.length());
            if (nRet == ERROR_SUCCESS) {
              // Trim NULL and return successful query!
              _rtn = buffer.substring(0, nBytes - 1);
            }
          }
          RegCloseKey(hKey);
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



////////////////////////////////////////////////////////////////////////////////
  public boolean saveSetting(int root, String section, String key, String setting) {
    boolean _rtn = false;
    // Section   Required. String expression containing the name of the section where the key setting is being saved.
    // Key       Required. String expression containing the name of the key setting being saved.
    // Setting   Required. Expression containing the value that key is being set to.
    // Called by fnRegSetClerkCode( ) in modRegistrySettings
    //           fnRegSetDBName( ) in modRegistrySettings
    //           fnRegSetDBPath( ) in modRegistrySettings
    "SaveSetting"
.equals(Const cstrCurrentProc As String);
    try {

      int nRet = 0;
      int hKey = 0;
      int nResult = 0;

      // Open (or create and open) key
      nRet = RegCreateKeyEx(root, subKey(section), 0&, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, hKey, nResult);
      if (nRet == ERROR_SUCCESS) {
        // Set appropriate value for default query
        if (key.equals("*")) {
          key = "";
        }
        // Null-terminate setting, in case it's empty.
        // Strange mirroring can occur otherwise.
        setting = setting+ vbNullChar;
        // Write new value to registry
        nRet = RegSetValueEx(hKey, key, 0&, REG_SZ, ByVal setting, setting.length());
        RegCloseKey(hKey);
      }
      _rtn = (nRet == ERROR_SUCCESS);
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



// ********************************************
//  Private Methods
// ********************************************

////////////////////////////////////////////////////////////////////////////////
  private String subKey(String section) {
    String _rtn = "";
    // Build SubKey from known values
    // Called by  DeleteSetting( ) in CRegSettings
    //            GetAllSettings( ) in CRegSettings
    //            GetSetting( ) in CRegSettings
    //            SaveSetting( ) in CRegSettings
    "Property Get Initialized"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = "Software\\"+ m_company+ "\\"+ m_appName;
      if (section.length()) {
        _rtn = subKey()+ "\\"+ section;
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


case class RegsettingsData(
              id: Option[Int],

              )

object Regsettingss extends Controller with ProvidesUser {

  val regsettingsForm = Form(
    mapping(
      "id" -> optional(number),

  )(RegsettingsData.apply)(RegsettingsData.unapply))

  implicit val regsettingsWrites = new Writes[Regsettings] {
    def writes(regsettings: Regsettings) = Json.obj(
      "id" -> Json.toJson(regsettings.id),
      C.ID -> Json.toJson(regsettings.id),

    )
  }

  def get(id: Int) = GetAction { implicit request =>
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.LIST_REGSETTINGS), { user =>
      Ok(Json.toJson(Regsettings.get(user, id)))
    })
  }

  def update(id: Int) = PostAction { implicit request =>
    Logger.debug("in regsettingss.update")
    regsettingsForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      regsettings => {
        Logger.debug(s"form: ${regsettings.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.EDIT_REGSETTINGS), { user =>
          Ok(
            Json.toJson(
              Regsettings.update(user,
                Regsettings(
                       id,

                ))))
        })
      }
    )
  }

  def create = PostAction { implicit request =>
    Logger.debug("in regsettingss.create")
    regsettingsForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      regsettings => {
        Logger.debug(s"form: ${regsettings.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.NEW_REGSETTINGS), { user =>
          Ok(
            Json.toJson(
              Regsettings.create(user,
                Regsettings(

                ))))
        })
      }
    )
  }

  def delete(id: Int) = PostAction { implicit request =>
    Logger.debug("in regsettingss.delete")
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.DELETE_REGSETTINGS), { user =>
      Regsettings.delete(user, id)
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

case class Regsettings(
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

object Regsettings {

  lazy val emptyRegsettings = Regsettings(
)

  def apply(
      id: Int,
) = {

    new Regsettings(
      id,
)
  }

  def apply(
) = {

    new Regsettings(
)
  }

  private val regsettingsParser: RowParser[Regsettings] = {
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
        Regsettings(
              id,
,
              createdAt,
              updatedAt,
              updatedBy)
    }
  }

  def create(user: CompanyUser, regsettings: Regsettings): Regsettings = {
    save(user, regsettings, true)
  }

  def update(user: CompanyUser, regsettings: Regsettings): Regsettings = {
    save(user, regsettings, false)
  }

  private def save(user: CompanyUser, regsettings: Regsettings, isNew: Boolean): Regsettings = {
    def getFields = {
      List(

      )
    }
    def throwException = {
      throw new RuntimeException(s"Error when saving ${C.REGSETTINGS}")
    }

    DBHelper.saveEx(
      user,
      Register(
        C.REGSETTINGS,
        C.ID,
        regsettings.id,
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

  def load(user: CompanyUser, id: Int): Option[Regsettings] = {
    loadWhere(user, s"${C.ID} = {id}", 'id -> id)
  }

  def loadWhere(user: CompanyUser, where: String, args : scala.Tuple2[scala.Any, anorm.ParameterValue[_]]*) = {
    DB.withConnection(user.database.database) { implicit connection =>
      SQL(s"SELECT t1.*, t2.${C.FK_NAME} FROM ${C.REGSETTINGS} t1 INNER JOIN ${C.???} t2 ON t1.${C.FK_ID} = t2.${C.FK_ID} WHERE $where")
        .on(args: _*)
        .as(regsettingsParser.singleOpt)
    }
  }

  def delete(user: CompanyUser, id: Int) = {
    DB.withConnection(user.database.database) { implicit connection =>
      try {
        SQL(s"DELETE FROM ${C.REGSETTINGS} WHERE ${C.ID} = {id}")
        .on('id -> id)
        .executeUpdate
      } catch {
        case NonFatal(e) => {
          Logger.error(s"can't delete a ${C.REGSETTINGS}. ${C.ID} id: $id. Error ${e.toString}")
          throw e
        }
      }
    }
  }

  def get(user: CompanyUser, id: Int): Regsettings = {
    load(user, id) match {
      case Some(p) => p
      case None => emptyRegsettings
    }
  }
}


// Router

GET     /api/v1/general/regsettings/:id              controllers.logged.modules.general.Regsettingss.get(id: Int)
POST    /api/v1/general/regsettings                  controllers.logged.modules.general.Regsettingss.create
PUT     /api/v1/general/regsettings/:id              controllers.logged.modules.general.Regsettingss.update(id: Int)
DELETE  /api/v1/general/regsettings/:id              controllers.logged.modules.general.Regsettingss.delete(id: Int)




/**/
