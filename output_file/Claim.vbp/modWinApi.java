public class modWinApi {

  // Module     : modWinApi
  // Description: Declarations and constants associated with Windows API functions called
  //              throughout the app
  // Procedures :
  //              fnEnumWindowsProc(ByVal hwnd As Long, ByVal NotUsed As Long) As Boolean
  //              fnEnumAllWindows(ByVal strSessionID As String, ByVal strSearchString As String, _
  //                  ByRef hwndFound As Long) As Boolean
  //              fnWindowText(ByVal hwnd As Long) As String
  //
  // Uses       : USER32.DLL, to get EnumWindows() and GetWindowText()
  // Modified   :
  // 03/16/01 DAS Cleaned with Total Visual CodeTools 2000
  // --------------------------------------------------
//Option Explicit
  *Option Compare Binary

  private static final String MCSTRNAME = "modWinApi.";


  //------------------------------------------------------------------------
  //            Prototypes for Win API functions used by more than 1 module
  //
  //       SHGetFolderPath - used by fnGetSpecialFolder( ) in modWinApi
  //       ShellExecute    - used by fnOpenFileInDefaultApp() in modGeneral.bas
  //------------------------------------------------------------------------
*TODO: API Public Declare Function SHGetFolderPath Lib "shfolder.dll" Alias "SHGetFolderPathA" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal pszPath As String) As Long

*TODO: API Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


  // The following are used by the SHGetFolderPath function that's wrapped
  // by fnGetSpecialFolder
  //'{user}\Start Menu _
  public static final Long CSIDL_ADMINTOOLS = &H30;
  //\Programs\Administrative Tools
  //'non localized startup
  public static final Long CSIDL_ALTSTARTUP = &H1D;
  //'{user}\Application Data
  public static final Long CSIDL_APPDATA = &H1A;
  //'{desktop}\Recycle Bin
  public static final Long CSIDL_BITBUCKET = &HA;
  //'My Computer\Control Panel
  public static final Long CSIDL_CONTROLS = &H3;
  public static final Long CSIDL_COOKIES = &H21;
  //'{namespace root}
  public static final Long CSIDL_DESKTOP = &H0;
  //'{user}\Desktop
  public static final Long CSIDL_DESKTOPDIRECTORY = &H10;
  //'{user}\Favourites
  public static final Long CSIDL_FAVORITES = &H6;
  //'windows\fonts
  public static final Long CSIDL_FONTS = &H14;
  public static final Long CSIDL_HISTORY = &H22;
  //'Internet virtual folder
  public static final Long CSIDL_INTERNET = &H1;
  //'Internet Cache folder
  public static final Long CSIDL_INTERNET_CACHE = &H20;
  //'{user}\Local Settings\
  public static final Long CSIDL_LOCAL_APPDATA = &H1C&;
  //_Application Data (non roaming)

  //'My Computer
  public static final Long CSIDL_DRIVES = &H11;
  //'C:\Program Files\My Pictures
  public static final Long CSIDL_MYPICTURES = &H27;
  //'{user}\nethood
  public static final Long CSIDL_NETHOOD = &H13;
  //'Network Neighbourhood
  public static final Long CSIDL_NETWORK = &H12;

  //'My Computer\Printers
  public static final Long CSIDL_PRINTERS = &H4;
  //'{user}\PrintHood
  public static final Long CSIDL_PRINTHOOD = &H1B;
  //'My Documents
  public static final Long CSIDL_PERSONAL = &H5;

  //'Program Files folder
  public static final Long CSIDL_PROGRAM_FILES = &H26;
  //'Program Files folder for x86 apps (Alpha)
  public static final Long CSIDL_PROGRAM_FILESX86 = &H2A;
  //'Start Menu\Programs
  public static final Long CSIDL_PROGRAMS = &H2;
  //'Program Files\Common
  public static final Long CSIDL_PROGRAM_FILES_COMMON = &H2B;
  //'x86 \Program Files\Common on RISC
  public static final Long CSIDL_PROGRAM_FILES_COMMONX86 = &H2C;
  //'{user}\Recent
  public static final Long CSIDL_RECENT = &H8;
  //'{user}\SendTo
  public static final Long CSIDL_SENDTO = &H9;
  //'{user}\Start Menu
  public static final Long CSIDL_STARTMENU = &HB;
  //'Start Menu\Programs\Startup
  public static final Long CSIDL_STARTUP = &H7;
  //'system folder
  public static final Long CSIDL_SYSTEM = &H25;
  //'system folder for x86 apps (Alpha)
  public static final Long CSIDL_SYSTEMX86 = &H29;
  public static final Long CSIDL_TEMPLATES = &H15;
  //'user's profile folder
  public static final Long CSIDL_PROFILE = &H28;
  //'Windows directory or SYSROOT()
  public static final Long CSIDL_WINDOWS = &H24;

  //'(all users)\Start Menu\ _
  public static final Long CSIDL_COMMON_ADMINTOOLS = &H2F;
  //Programs\Administrative Tools
  //'non localized common startup
  public static final Long CSIDL_COMMON_ALTSTARTUP = &H1E;
  //'(all users)\Application Data
  public static final Long CSIDL_COMMON_APPDATA = &H23;
  //'(all users)\Desktop
  public static final Long CSIDL_COMMON_DESKTOPDIRECTORY = &H19;
  //'(all users)\Documents
  public static final Long CSIDL_COMMON_DOCUMENTS = &H2E;
  //'(all users)\Favourites
  public static final Long CSIDL_COMMON_FAVORITES = &H1F;
  //'(all users)\Programs
  public static final Long CSIDL_COMMON_PROGRAMS = &H17;
  //'(all users)\Start Menu
  public static final Long CSIDL_COMMON_STARTMENU = &H16;
  //'(all users)\Startup
  public static final Long CSIDL_COMMON_STARTUP = &H18;
  //'(all users)\Templates
  public static final Long CSIDL_COMMON_TEMPLATES = &H2D;

  //'combine with CSIDL_ value to force
  *Public Const CSIDL_FLAG_CREATE = &H8000&
  //create on SHGetSpecialFolderLocation()
  //'combine with CSIDL_ value to force
  *Public Const CSIDL_FLAG_DONT_VERIFY = &H4000
  //create on SHGetSpecialFolderLocation()
  //'mask for all possible flag values
  *Public Const CSIDL_FLAG_MASK = &HFF00
  //'current value for user, verify it exists
  *Public Const SHGFP_TYPE_CURRENT = 0
  *Public Const SHGFP_TYPE_DEFAULT = 1
  *Public Const MAX_PATH = 260
  //'Success
  *Public Const S_OK = &H0
  //'The folder is valid, but does not exist
  *Public Const S_FALSE = &H1
  //'Invalid CSIDL Value
  private static final int E_INVALIDARG = 0x80070057;

  //SQL_INTEGRATED_SECURITY
  //
  // Win32 APIs to determine OS information.
  //
*TODO: API Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
//*TODO:** type is translated as a new class at the end of the file Private Type OSVERSIONINFO
  private static final int VER_PLATFORM_WIN32S = 0;
  private static final int VER_PLATFORM_WIN32_WINDOWS = 1;
  private static final int VER_PLATFORM_WIN32_NT = 2;

  //
  // Win32 NetAPIs.
  //
  //' Maximum username length
  private static final int USERNAME_LENGTH = 256;
*TODO: API Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
*TODO: API Private Declare Function GetUserNameW Lib "advapi32.dll" (lpBuffer As Byte, nSize As Long) As Long
  //SQL_INTEGRATED_SECURITY



  //SQL_INTEGRATED_SECURITY  - Added
  //////////////////////////////////////////////////////////////////////////////////////////////////
  public static String fnGetNetworkUser() {
    String _rtn = "";
    // Comments  : Gets the User ID that is currently logged on to the network
    // Parameters: None
    // Returns   : User ID
    // Source    : Karl Peterson's Classic VB site
    // Modified  :
    //
    // --------------------------------------------------
    try {
      "fnGetNetworkUser"
.equals(Const cstrCurrentProc As String);
      Const(clngNameLength == USERNAME_LENGTH + 1);
      OSVERSIONINFO objOS = null;
      String strBuffer = "";
      byte[] bytBuffer() = null;
      int lngRetVal = 0;
      int lngLength = 0;

      objOS.dwOSVersionInfoSize = objOS.length();
      GetVersionEx(objOS);

      if (objOS.dwPlatformId == VER_PLATFORM_WIN32_NT) {
        lngLength = clngNameLength * 2;
        G.redimPreserve(0 To lngLength - 1,  );
        if (GetUserNameW(bytBuffer[0], lngLength)) {
          strBuffer = bytBuffer;
          _rtn = strBuffer.substring(0, lngLength - 1);
        }
      }
      // **TODO:** label found: PROC_EXIT:;
  //' disablfe error handler
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
//SQL_INTEGRATED_SECURITY - Added



//////////////////////////////////////////////////////////////////////////////////////////////////
  public static String fnGetSpecialFolder(int hwndOwner, int cSIDL) {
    String _rtn = "";
    // Comments  : Get the fully qualified path to one of Windows' special folders.
    //             This approach is what is recommended for Win2K and all previous
    //             versions of Windows. Be sure to include " Or CSIDL_FLAG_CREATE"
    //             if you want the folder created if it doesn't already exist.
    //
    //             Using it requires that SHFOLDER.DLL be distributed if the app
    //             will be used on pre-Win2K versions of the Windows OS. This is
    //             available as a redistributable within the Platform SDK.
    //
    //             Much of this code was lifted from MSKB article Q252652.
    //
    // Parameters: hWndOwner - handle to a window (0 if not needed)
    //             CSIDL - the CSIDL indicating which folder path to return.
    //
    // Called by : fnLogOpen( ) in modAppLog
    //             fnLogPrune( ) in modAppLog
    //
    // Returns   : Directory name, with appended "\" if necessary
    // Modified  :
    //
    // --------------------------------------------------
    try {
      "fnGetSpecialFolder"
.equals(Const cstrCurrentProc As String);
      String strPath = "";
      int lngRetVal = 0;

      // Fill our string buffer
      strPath = String(MAX_PATH, 0);

      lngRetVal = SHGetFolderPath(hwndOwner, cSIDL, 0&, SHGFP_TYPE_CURRENT, strPath);

      switch (lngRetVal) {
        case  S_OK:
          // We retrieved the folder successfully.
          // All C strings are null-terminated, so return the string up to the
          // first null character
          _rtn = strPath.substring(0, strPath.indexOf(Chr$(0), 1) - 1);
          break;

        case  S_FALSE:
          // The CSIDL in the 2nd argument is valid, but the folder does not exist.
          // Use CSIDL_FLAG_CREATE to have it created automatically
          //!TODO! Gen msg via frmMsgBox
          //fnProcessFatalError Err.Source, _
          //                    fte_OtherErrType, Err.Number, _
          //                    Err.Description, Err.Source, _
          //                    Err.HelpFile, Err.HelpContext, _
          //                    "The specified folder ( " & CStr(CSIDL) & ") does not exist. " & _
          //                    "Add the CSIDL_FLAG_CREATE flag to create it. RC = " & CStr(lngRetVal)
          break;

        default:
          // E_INVALIDARG...CSIDL in the 2nd argument is invalid
          //!TODO! Gen msg via frmMsgBox
          //fnProcessFatalError Err.Source, _
          //                    fte_OtherErrType, Err.Number, _
          //                    Err.Description, Err.Source, _
          //                    Err.HelpFile, Err.HelpContext, _
          //                    "An invalid CSIDL argument (" & CStr(CSIDL) & ") was " & _
          //                    "passed to this function. RC = " & CStr(lngRetVal)
          break;
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

private class OSVERSIONINFO {
    public Long dwOSVersionInfoSize;
    public Long dwMajorVersion;
    public Long dwMinorVersion;
    public Long dwBuildNumber;
    public Long dwPlatformId;
    public String szCSDVersion;
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


case class OdwinapiData(
              id: Option[Int],

              )

object Odwinapis extends Controller with ProvidesUser {

  val odwinapiForm = Form(
    mapping(
      "id" -> optional(number),

  )(OdwinapiData.apply)(OdwinapiData.unapply))

  implicit val odwinapiWrites = new Writes[Odwinapi] {
    def writes(odwinapi: Odwinapi) = Json.obj(
      "id" -> Json.toJson(odwinapi.id),
      C.ID -> Json.toJson(odwinapi.id),

    )
  }

  def get(id: Int) = GetAction { implicit request =>
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.LIST_ODWINAPI), { user =>
      Ok(Json.toJson(Odwinapi.get(user, id)))
    })
  }

  def update(id: Int) = PostAction { implicit request =>
    Logger.debug("in odwinapis.update")
    odwinapiForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      odwinapi => {
        Logger.debug(s"form: ${odwinapi.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.EDIT_ODWINAPI), { user =>
          Ok(
            Json.toJson(
              Odwinapi.update(user,
                Odwinapi(
                       id,

                ))))
        })
      }
    )
  }

  def create = PostAction { implicit request =>
    Logger.debug("in odwinapis.create")
    odwinapiForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      odwinapi => {
        Logger.debug(s"form: ${odwinapi.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.NEW_ODWINAPI), { user =>
          Ok(
            Json.toJson(
              Odwinapi.create(user,
                Odwinapi(

                ))))
        })
      }
    )
  }

  def delete(id: Int) = PostAction { implicit request =>
    Logger.debug("in odwinapis.delete")
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.DELETE_ODWINAPI), { user =>
      Odwinapi.delete(user, id)
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

case class Odwinapi(
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

object Odwinapi {

  lazy val emptyOdwinapi = Odwinapi(
)

  def apply(
      id: Int,
) = {

    new Odwinapi(
      id,
)
  }

  def apply(
) = {

    new Odwinapi(
)
  }

  private val odwinapiParser: RowParser[Odwinapi] = {
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
        Odwinapi(
              id,
,
              createdAt,
              updatedAt,
              updatedBy)
    }
  }

  def create(user: CompanyUser, odwinapi: Odwinapi): Odwinapi = {
    save(user, odwinapi, true)
  }

  def update(user: CompanyUser, odwinapi: Odwinapi): Odwinapi = {
    save(user, odwinapi, false)
  }

  private def save(user: CompanyUser, odwinapi: Odwinapi, isNew: Boolean): Odwinapi = {
    def getFields = {
      List(

      )
    }
    def throwException = {
      throw new RuntimeException(s"Error when saving ${C.ODWINAPI}")
    }

    DBHelper.saveEx(
      user,
      Register(
        C.ODWINAPI,
        C.ID,
        odwinapi.id,
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

  def load(user: CompanyUser, id: Int): Option[Odwinapi] = {
    loadWhere(user, s"${C.ID} = {id}", 'id -> id)
  }

  def loadWhere(user: CompanyUser, where: String, args : scala.Tuple2[scala.Any, anorm.ParameterValue[_]]*) = {
    DB.withConnection(user.database.database) { implicit connection =>
      SQL(s"SELECT t1.*, t2.${C.FK_NAME} FROM ${C.ODWINAPI} t1 INNER JOIN ${C.???} t2 ON t1.${C.FK_ID} = t2.${C.FK_ID} WHERE $where")
        .on(args: _*)
        .as(odwinapiParser.singleOpt)
    }
  }

  def delete(user: CompanyUser, id: Int) = {
    DB.withConnection(user.database.database) { implicit connection =>
      try {
        SQL(s"DELETE FROM ${C.ODWINAPI} WHERE ${C.ID} = {id}")
        .on('id -> id)
        .executeUpdate
      } catch {
        case NonFatal(e) => {
          Logger.error(s"can't delete a ${C.ODWINAPI}. ${C.ID} id: $id. Error ${e.toString}")
          throw e
        }
      }
    }
  }

  def get(user: CompanyUser, id: Int): Odwinapi = {
    load(user, id) match {
      case Some(p) => p
      case None => emptyOdwinapi
    }
  }
}


// Router

GET     /api/v1/general/odwinapi/:id              controllers.logged.modules.general.Odwinapis.get(id: Int)
POST    /api/v1/general/odwinapi                  controllers.logged.modules.general.Odwinapis.create
PUT     /api/v1/general/odwinapi/:id              controllers.logged.modules.general.Odwinapis.update(id: Int)
DELETE  /api/v1/general/odwinapi/:id              controllers.logged.modules.general.Odwinapis.delete(id: Int)




/**/
