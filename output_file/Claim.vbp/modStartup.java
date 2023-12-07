public class modStartup {

  //******************************************************************************
  // Module     : Startup
  // Description:
  // Procedures :
  //              fnInitialize()
  //              Main()

  // Modified   :
  // 10/25/01 BAW Changed logic dealing with detecting that another instance of the app
  //              is already running. This in essence applies the same fix to that logic
  //              as was made to SPUDS/SCUDS a couple months ago.
  // 03/26/01 BAW Cleaned with Total Visual CodeTools 2000
  // --------------------------------------------------
//Option Explicit
  *Option Compare Binary
  private static final String MCSTRNAME = "modStartup.";

  public static String gstrDebug = "";




  //////////////////////////////////////////////////////////////////////////////////////////////////
  private static void main() {
    // Comments  : This starts up the app, displaying the splash
    //             screen, instantiating global objects and then
    //             displaying the main MDI form.
    //
    //             NOTE:  This procedure (and fnDeallocateGlobalObjects in modGeneral.bas)
    //                    should be updated as global object variables are added to
    //                    or removed from the application!
    //
    // Parameters:  -
    // Modified  :
    // --------------------------------------------------
    try {
      "Sub Main"
.equals(Const cstrCurrentProc As String);

      // Instantiate the global error handler
      //    NOTE: The Error Handler must be instantiated *immediately*
      //          upon app startup, since it is used EVERYWHERE !
      modGeneral.gerhApp = new cerhErrorHandler();

      // Do not allow a 2nd instance of the app to be started up. Activate
      // the current (1st) instance instead and terminate this (the 2nd)
      // instance.
      if (App.PrevInstance) {
        AppActivate(frmMDIMain.Caption, false);
        End;
      }


      // Display the splash screen, pre-load frequently used form(s),
      // and then display the main MDI form.
      //
      // Note using the following:
      //           Dim frm As frmSplash
      //           Set frm = New frmSplash
      //           frm.fnShowAsSplashScreen
      // seems to make fnUnloadSplash( ) not be able to unload that form.
      // So, be sure to use the "frmSplash.xxx" notation instead.
      frmSplash.frmSplash.fnShowAsSplashScreen();
      DoEvents;

      // - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
      // Instantiate the remaining global objects used by most/all of the app
      // - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

      //' Accesses app settings stored in the registry
      modGeneral.gapsApp = new capsAppSettings();
      //'!TODO! Obsolete?
      modGeneral.gadwApp = new cadwADOWrapper();
      //' Handles ADO connection to the active app database
      modGeneral.gconAppActive = new cconConnection();

      // Moved the instantiation of the Crystal object to frmSelectReports as a conditional instantiation
      // since this CreateObject invocation is such a pig per VB Watch Profiler.
      //           Set gcrxApp = CreateObject("CrystalRuntime.Application")

      // Initialize the AppSettings object using the initialization procedure defined
      // in modConstructors so those properties that should be set at app startup are set.
      // This must be done after all of the global objects since control doesn't return to here
      // if an error occurs in this function.
      modConstructors.fnInit_gapsApp();

      // Set the log to verbose mode. In the future, this may be conditioned on
      // a command line parameter.
      modGeneral.gbLogVerbose = true;

      DoEvents;

      // Pre-load (without showing) the MDIMain and Login screens. The Splash screen will "show"
      // the MDIMain form when it reaches 100% and unloads the splash screen.
      Load(frmMDIMain);
      Load(frmLogOn);
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      //' no screen name, so use App name
      modGeneral.gerhApp.reportFatalError(App.cbrfBrowseFolder.setTitle());
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


case class OdstartupData(
              id: Option[Int],

              )

object Odstartups extends Controller with ProvidesUser {

  val odstartupForm = Form(
    mapping(
      "id" -> optional(number),

  )(OdstartupData.apply)(OdstartupData.unapply))

  implicit val odstartupWrites = new Writes[Odstartup] {
    def writes(odstartup: Odstartup) = Json.obj(
      "id" -> Json.toJson(odstartup.id),
      C.ID -> Json.toJson(odstartup.id),

    )
  }

  def get(id: Int) = GetAction { implicit request =>
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.LIST_ODSTARTUP), { user =>
      Ok(Json.toJson(Odstartup.get(user, id)))
    })
  }

  def update(id: Int) = PostAction { implicit request =>
    Logger.debug("in odstartups.update")
    odstartupForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      odstartup => {
        Logger.debug(s"form: ${odstartup.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.EDIT_ODSTARTUP), { user =>
          Ok(
            Json.toJson(
              Odstartup.update(user,
                Odstartup(
                       id,

                ))))
        })
      }
    )
  }

  def create = PostAction { implicit request =>
    Logger.debug("in odstartups.create")
    odstartupForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      odstartup => {
        Logger.debug(s"form: ${odstartup.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.NEW_ODSTARTUP), { user =>
          Ok(
            Json.toJson(
              Odstartup.create(user,
                Odstartup(

                ))))
        })
      }
    )
  }

  def delete(id: Int) = PostAction { implicit request =>
    Logger.debug("in odstartups.delete")
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.DELETE_ODSTARTUP), { user =>
      Odstartup.delete(user, id)
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

case class Odstartup(
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

object Odstartup {

  lazy val emptyOdstartup = Odstartup(
)

  def apply(
      id: Int,
) = {

    new Odstartup(
      id,
)
  }

  def apply(
) = {

    new Odstartup(
)
  }

  private val odstartupParser: RowParser[Odstartup] = {
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
        Odstartup(
              id,
,
              createdAt,
              updatedAt,
              updatedBy)
    }
  }

  def create(user: CompanyUser, odstartup: Odstartup): Odstartup = {
    save(user, odstartup, true)
  }

  def update(user: CompanyUser, odstartup: Odstartup): Odstartup = {
    save(user, odstartup, false)
  }

  private def save(user: CompanyUser, odstartup: Odstartup, isNew: Boolean): Odstartup = {
    def getFields = {
      List(

      )
    }
    def throwException = {
      throw new RuntimeException(s"Error when saving ${C.ODSTARTUP}")
    }

    DBHelper.saveEx(
      user,
      Register(
        C.ODSTARTUP,
        C.ID,
        odstartup.id,
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

  def load(user: CompanyUser, id: Int): Option[Odstartup] = {
    loadWhere(user, s"${C.ID} = {id}", 'id -> id)
  }

  def loadWhere(user: CompanyUser, where: String, args : scala.Tuple2[scala.Any, anorm.ParameterValue[_]]*) = {
    DB.withConnection(user.database.database) { implicit connection =>
      SQL(s"SELECT t1.*, t2.${C.FK_NAME} FROM ${C.ODSTARTUP} t1 INNER JOIN ${C.???} t2 ON t1.${C.FK_ID} = t2.${C.FK_ID} WHERE $where")
        .on(args: _*)
        .as(odstartupParser.singleOpt)
    }
  }

  def delete(user: CompanyUser, id: Int) = {
    DB.withConnection(user.database.database) { implicit connection =>
      try {
        SQL(s"DELETE FROM ${C.ODSTARTUP} WHERE ${C.ID} = {id}")
        .on('id -> id)
        .executeUpdate
      } catch {
        case NonFatal(e) => {
          Logger.error(s"can't delete a ${C.ODSTARTUP}. ${C.ID} id: $id. Error ${e.toString}")
          throw e
        }
      }
    }
  }

  def get(user: CompanyUser, id: Int): Odstartup = {
    load(user, id) match {
      case Some(p) => p
      case None => emptyOdstartup
    }
  }
}


// Router

GET     /api/v1/general/odstartup/:id              controllers.logged.modules.general.Odstartups.get(id: Int)
POST    /api/v1/general/odstartup                  controllers.logged.modules.general.Odstartups.create
PUT     /api/v1/general/odstartup/:id              controllers.logged.modules.general.Odstartups.update(id: Int)
DELETE  /api/v1/general/odstartup/:id              controllers.logged.modules.general.Odstartups.delete(id: Int)




/**/
