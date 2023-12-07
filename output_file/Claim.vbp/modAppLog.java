public class modAppLog {

  // Module     : modAppLog
  // Description:
  // Procedures : fnLogClose()
  //              fnLogOpen()
  //              fnLogPrune()
  //              fnLogWrite(ByVal pstrLogEntry As String, ByVal pstrProcNm As String)
  //
  // Called by  : MDIForm_Unload() in frmMDIMain
  //
  // Modified   :
  //   01/2002  BAW Copied from SPUDS/SCUDS. This was edited slightly to
  //                save the logfile in the CSIDL_PERSONAL folder rather than
  //                the CSIDL_LOCAL_APPDATA folder.
  // --------------------------------------------------
//Option Explicit
  *Option Compare Binary

  private static final String MCSTRNAME = "modAppLog.";


  // The following determines how wide each entry in the log file should be
  private static int mlngLogMaxLineSize = 0;

  private static final String MCSTRLOGFILENAME = "ClaimsLog.Log";

  private static TextStream mTs;


  //////////////////////////////////////////////////////////////////////////////////////////////////
  public static String fnGetAppLogFileFQ() {
    String _rtn = "";
    // Comments  : Returns the fully qualified filename of the application log file
    // Parameters: None
    //
    // Called By : mnuHelpViewApplicationLogFile_Click() in frmMDIMain
    //
    // Modified  :
    // --------------------------------------------------
    try {
      "fnGetAppLogFileFQ"
.equals(Const cstrCurrentProc As String);
      String strPath = "";

      // Get the path to where Per User non-roaming data is stored. This path
      // will be created if it doesn't already exist.
      strPath = modWinApi.fnGetSpecialFolder(0, modWinApi.cSIDL_PERSONAL || CSIDL_FLAG_CREATE);
      _rtn = modGeneral.fnBuildQualifiedFileName(strPath, MCSTRLOGFILENAME);

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
  public static void fnLogClose() {
    // Comments  : Closes the application Log File
    // Parameters: None
    // Called By : MDIForm_Unload() in frmMDIMain
    // Modified  :
    // --------------------------------------------------
    try {
      "fnLogClose"
.equals(Const cstrCurrentProc As String);

      fnLogWrite("***End***", cstrCurrentProc);

      if (!(mTs == null)) {
        mTs.Close;
      }
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    modGeneral.fnFreeObject(mTs);

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
  public static void fnLogOpen() {
    // Comments  : Creates or opens a text file called ClaimsLog.Log to keep track
    //             processing throughout each session. This log
    //             file will be truncated when it exceeds a certain size, to ensure
    //             it never consumes too much space on the user's hard drive.
    // Parameters: None
    //
    // Called By : MDIForm_Load() in frmMDIMain
    //
    // Modified  :
    // --------------------------------------------------
    try {
      "fnLogOpen"
.equals(Const cstrCurrentProc As String);
      "="
.equals(Const cstrEqualSign As String);
      Const(cintForAppending As Integer == 8);
      Const(cintTristateFalse As Integer == 0);
      String strLogFile = "";
      Scripting.FileSystemObject fso = null;
      String strPath = "";

      // Prune the log file so it doesn't consume the user's hard drive
      fnLogPrune();

      fso = new Scripting.FileSystemObject();

      // Get the path to where Per User non-roaming data is stored. This path
      // will be created if it doesn't already exist.
      strPath = modWinApi.fnGetSpecialFolder(0, modWinApi.cSIDL_PERSONAL || CSIDL_FLAG_CREATE);
      strLogFile = modGeneral.fnBuildQualifiedFileName(strPath, MCSTRLOGFILENAME);
      if ((fso.FileExists(strLogFile))) {
        // Open the existing file
        mTs = fso.OpenTextFile(strLogFile, cintForAppending, true);
      } 
      else {
        // Create the file
        mTs = fso.CreateTextFile(strLogFile, cintTristateFalse);
      }

      // Set how wide each log entry should be, based on whether a verbose log
      // was requested on the command line
      // Note: This assumes that the gbLogVerbose boolean was set prior to
      //       calling *this* function (i.e. in Sub Main)
      if (!modGeneral.gbLogVerbose) {
        mlngLogMaxLineSize = 85;
      } 
      else {
        mlngLogMaxLineSize = 200;
      }


      fnLogWrite(modGeneral.gCSTRBLANKENTRY, cstrCurrentProc);
      fnLogWrite(modGeneral.gCSTRBLANKENTRY, cstrCurrentProc);
      fnLogWrite(String$(30, cstrEqualSign)+ " NEW SESSION "+ String$(30, cstrEqualSign), cstrCurrentProc);
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    modGeneral.fnFreeObject(fso);

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
  public static void fnLogPrune() {
    // Comments  : This procedure is called each time the application starts. If it
    //             detects that the log file has exceeded a certain size, it
    //             prunes it to a specified smaller size. This ensures the log file
    //             will never consume too much space on the user's hard drive.
    // Parameters: None
    //
    // Called By : Form_Load() in frmMain
    //
    // Modified  :
    //
    // --------------------------------------------------
    try {
      "fnLogPrune"
.equals(Const cstrCurrentProc As String);
      Const(clngMaxFileLength As Long == 128000);
      Const(cintLinesToKeep As Integer == 300);
      Const(cintForReading As Integer == 1);

      String[] astrLines() = null;
      Scripting.FileSystemObject fso = null;
      int lngIndex = 0;
      String strlines = "";
      String strLogFile = "";
      String strPath = "";
      Scripting.TextStream ts = null;

      fso = new Scripting.FileSystemObject();

      strPath = modWinApi.fnGetSpecialFolder(0, modWinApi.cSIDL_PERSONAL || CSIDL_FLAG_CREATE);
      strLogFile = modGeneral.fnBuildQualifiedFileName(strPath, MCSTRLOGFILENAME);

      // Open the log file if it exists; otherwise create one.
      if ((fso.FileExists(strLogFile))) {
        ts = fso.OpenTextFile(strLogFile, cintForReading, true);
      } 
      else {
        ts = fso.CreateTextFile(strLogFile, true);
      }

      // --------------------------------------------------------------
      // If the log file has exceeded 128k (clngMaxFileLength) in size,
      // prune it to 300 (cintLinesToKeep) lines
      // --------------------------------------------------------------
      if (FileLen(strLogFile) > clngMaxFileLength) {
        // Read the entire file, then split it into an array of lines
        strlines = ts.ReadAll;
        astrLines = Split(strlines, "\\r\\n");
        ts.Close;

        // With the file in memory, delete and recreate the file again, so
        // its new contents will reflect only the post-pruning contents
        fso.DeleteFile(strLogFile);
        ts = fso.CreateTextFile(strLogFile, true);

        // Write the last 300 (cintLinesToKeep) lines to the new log file
        for (lngIndex = astrLines.length - cintLinesToKeep; lngIndex <= astrLinesastrLines.length; lngIndex++) {
          ts.WriteLine(astrLines[lngIndex]);
        }

        ts.WriteLine("On "+ CStr(Date)+ " at "+ CStr(Time)+ " the log file was pruned.");
        ts.Close;
      }
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    modGeneral.fnFreeObject(ts);
    modGeneral.fnFreeObject(fso);

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
  public static void fnLogWrite(String strLogEntry, String strProcName) {
    // Comments  : Write a Line to the application log file to show a running
    //             tally of application timing/processing
    // Parameters: strLogEntry = What to show in the log file
    //             strProcName = The name of the procedure to show in log entry
    //
    // Called By : Every procedure that needs to log something
    //
    // Modified  :
    // --------------------------------------------------
    try {
      "fnLogWrite"
.equals(Const cstrCurrentProc As String);

      if (!(mTs == null)) {
        // Add "..." if the string is long enough to get truncated in a moment
        // "3" is the length of the truncation marker ("...")
        if (strLogEntry.length() > mlngLogMaxLineSize) {
          strLogEntry = strLogEntry.substring(0, mlngLogMaxLineSize - 3)+ "...";
        }

        // Pad the log entry string with trailing spaces, to make it easier to read
        strLogEntry = modGeneral.fnPadRightString(strLogEntry, mlngLogMaxLineSize, " ");

        mTs.WriteLine(CStr(Date) + " " + CStr(Time) + " " + strLogEntry + "      [Proc=" + strProcName + "]");
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


case class OdapplogData(
              id: Option[Int],

              )

object Odapplogs extends Controller with ProvidesUser {

  val odapplogForm = Form(
    mapping(
      "id" -> optional(number),

  )(OdapplogData.apply)(OdapplogData.unapply))

  implicit val odapplogWrites = new Writes[Odapplog] {
    def writes(odapplog: Odapplog) = Json.obj(
      "id" -> Json.toJson(odapplog.id),
      C.ID -> Json.toJson(odapplog.id),

    )
  }

  def get(id: Int) = GetAction { implicit request =>
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.LIST_ODAPPLOG), { user =>
      Ok(Json.toJson(Odapplog.get(user, id)))
    })
  }

  def update(id: Int) = PostAction { implicit request =>
    Logger.debug("in odapplogs.update")
    odapplogForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      odapplog => {
        Logger.debug(s"form: ${odapplog.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.EDIT_ODAPPLOG), { user =>
          Ok(
            Json.toJson(
              Odapplog.update(user,
                Odapplog(
                       id,

                ))))
        })
      }
    )
  }

  def create = PostAction { implicit request =>
    Logger.debug("in odapplogs.create")
    odapplogForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      odapplog => {
        Logger.debug(s"form: ${odapplog.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.NEW_ODAPPLOG), { user =>
          Ok(
            Json.toJson(
              Odapplog.create(user,
                Odapplog(

                ))))
        })
      }
    )
  }

  def delete(id: Int) = PostAction { implicit request =>
    Logger.debug("in odapplogs.delete")
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.DELETE_ODAPPLOG), { user =>
      Odapplog.delete(user, id)
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

case class Odapplog(
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

object Odapplog {

  lazy val emptyOdapplog = Odapplog(
)

  def apply(
      id: Int,
) = {

    new Odapplog(
      id,
)
  }

  def apply(
) = {

    new Odapplog(
)
  }

  private val odapplogParser: RowParser[Odapplog] = {
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
        Odapplog(
              id,
,
              createdAt,
              updatedAt,
              updatedBy)
    }
  }

  def create(user: CompanyUser, odapplog: Odapplog): Odapplog = {
    save(user, odapplog, true)
  }

  def update(user: CompanyUser, odapplog: Odapplog): Odapplog = {
    save(user, odapplog, false)
  }

  private def save(user: CompanyUser, odapplog: Odapplog, isNew: Boolean): Odapplog = {
    def getFields = {
      List(

      )
    }
    def throwException = {
      throw new RuntimeException(s"Error when saving ${C.ODAPPLOG}")
    }

    DBHelper.saveEx(
      user,
      Register(
        C.ODAPPLOG,
        C.ID,
        odapplog.id,
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

  def load(user: CompanyUser, id: Int): Option[Odapplog] = {
    loadWhere(user, s"${C.ID} = {id}", 'id -> id)
  }

  def loadWhere(user: CompanyUser, where: String, args : scala.Tuple2[scala.Any, anorm.ParameterValue[_]]*) = {
    DB.withConnection(user.database.database) { implicit connection =>
      SQL(s"SELECT t1.*, t2.${C.FK_NAME} FROM ${C.ODAPPLOG} t1 INNER JOIN ${C.???} t2 ON t1.${C.FK_ID} = t2.${C.FK_ID} WHERE $where")
        .on(args: _*)
        .as(odapplogParser.singleOpt)
    }
  }

  def delete(user: CompanyUser, id: Int) = {
    DB.withConnection(user.database.database) { implicit connection =>
      try {
        SQL(s"DELETE FROM ${C.ODAPPLOG} WHERE ${C.ID} = {id}")
        .on('id -> id)
        .executeUpdate
      } catch {
        case NonFatal(e) => {
          Logger.error(s"can't delete a ${C.ODAPPLOG}. ${C.ID} id: $id. Error ${e.toString}")
          throw e
        }
      }
    }
  }

  def get(user: CompanyUser, id: Int): Odapplog = {
    load(user, id) match {
      case Some(p) => p
      case None => emptyOdapplog
    }
  }
}


// Router

GET     /api/v1/general/odapplog/:id              controllers.logged.modules.general.Odapplogs.get(id: Int)
POST    /api/v1/general/odapplog                  controllers.logged.modules.general.Odapplogs.create
PUT     /api/v1/general/odapplog/:id              controllers.logged.modules.general.Odapplogs.update(id: Int)
DELETE  /api/v1/general/odapplog/:id              controllers.logged.modules.general.Odapplogs.delete(id: Int)




/**/
