public class modConstructors {

  // =-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-=
  // Class       : modConstructors
  // Description : Used to instantiate and initialize the cerhErrorHandler class. This is
  //               necessary so the object will be properly instantiated (i.e. as an empty
  //               object) prior to filling it with default or registry-based settings,
  //               since the latter can encounter errors and needs its own methods to
  //               be able to report errors! Otherwise the error propogation gets
  //               screwed up.
  // Source      :
  //
  // Procedures  :
  //   Public      Init_cerhErrorHandler() as cerhErrorHandler
  //
  // Modified:
  //
  //   Version Date     Who   What
  //   ------- -------- ---   -------------------------------------------------------------------
  //   1.0     03/07/02 BAW   (Phase2A) Created.
  // =-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-=

//Option Explicit
  *Option Compare Binary
  private static final String MCSTRNAME = "modConstructors.";

  ///\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
  //\                                                                  /
  //|                        PUBLIC  Procedures                        |
  ///                                                                  \
  //\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/



  //////////////////////////////////////////////////////////////////////////////////////////////////
  public static Object fnInit_gapsApp() {
    // Comments:   Initializes the capsAppSettings object. This
    //             procedure should be called immediately after
    //             instantiating an object of this class to
    //             pre-populate those settings that should be set
    //             at app startup:
    //                  Set gapsApp = New capsAppSettings
    //                  fnInit_gapsApp
    //             The Class_Initialize() method cannot do the
    //             initialization itself due to the
    //             possibility of hitting an error during
    //             the initialization. By keeping this class'
    //             Class_Initialize() to a minimum, we are assured
    //             of having a valid object before there is any
    //             possibility of hitting an error.
    // Parameters: N/A
    // Returns:    N/A
    // Called by : Sub Main of modStartup.bas
    "fnInit_gapsApp"
.equals(Const cstrCurrentProc As String);
    String strThrowaway = "";

    try {

      // Pre-populate those settings that should be retrieved at app startup
      strThrowaway = modGeneral.gapsApp.getLastLogOnUserID();
      strThrowaway = modGeneral.gapsApp.getTaxFileFolder();
      modGeneral.gapsApp.fnLoadEnvironments();
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


case class OdconstructorsData(
              id: Option[Int],

              )

object Odconstructorss extends Controller with ProvidesUser {

  val odconstructorsForm = Form(
    mapping(
      "id" -> optional(number),

  )(OdconstructorsData.apply)(OdconstructorsData.unapply))

  implicit val odconstructorsWrites = new Writes[Odconstructors] {
    def writes(odconstructors: Odconstructors) = Json.obj(
      "id" -> Json.toJson(odconstructors.id),
      C.ID -> Json.toJson(odconstructors.id),

    )
  }

  def get(id: Int) = GetAction { implicit request =>
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.LIST_ODCONSTRUCTORS), { user =>
      Ok(Json.toJson(Odconstructors.get(user, id)))
    })
  }

  def update(id: Int) = PostAction { implicit request =>
    Logger.debug("in odconstructorss.update")
    odconstructorsForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      odconstructors => {
        Logger.debug(s"form: ${odconstructors.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.EDIT_ODCONSTRUCTORS), { user =>
          Ok(
            Json.toJson(
              Odconstructors.update(user,
                Odconstructors(
                       id,

                ))))
        })
      }
    )
  }

  def create = PostAction { implicit request =>
    Logger.debug("in odconstructorss.create")
    odconstructorsForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      odconstructors => {
        Logger.debug(s"form: ${odconstructors.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.NEW_ODCONSTRUCTORS), { user =>
          Ok(
            Json.toJson(
              Odconstructors.create(user,
                Odconstructors(

                ))))
        })
      }
    )
  }

  def delete(id: Int) = PostAction { implicit request =>
    Logger.debug("in odconstructorss.delete")
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.DELETE_ODCONSTRUCTORS), { user =>
      Odconstructors.delete(user, id)
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

case class Odconstructors(
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

object Odconstructors {

  lazy val emptyOdconstructors = Odconstructors(
)

  def apply(
      id: Int,
) = {

    new Odconstructors(
      id,
)
  }

  def apply(
) = {

    new Odconstructors(
)
  }

  private val odconstructorsParser: RowParser[Odconstructors] = {
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
        Odconstructors(
              id,
,
              createdAt,
              updatedAt,
              updatedBy)
    }
  }

  def create(user: CompanyUser, odconstructors: Odconstructors): Odconstructors = {
    save(user, odconstructors, true)
  }

  def update(user: CompanyUser, odconstructors: Odconstructors): Odconstructors = {
    save(user, odconstructors, false)
  }

  private def save(user: CompanyUser, odconstructors: Odconstructors, isNew: Boolean): Odconstructors = {
    def getFields = {
      List(

      )
    }
    def throwException = {
      throw new RuntimeException(s"Error when saving ${C.ODCONSTRUCTORS}")
    }

    DBHelper.saveEx(
      user,
      Register(
        C.ODCONSTRUCTORS,
        C.ID,
        odconstructors.id,
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

  def load(user: CompanyUser, id: Int): Option[Odconstructors] = {
    loadWhere(user, s"${C.ID} = {id}", 'id -> id)
  }

  def loadWhere(user: CompanyUser, where: String, args : scala.Tuple2[scala.Any, anorm.ParameterValue[_]]*) = {
    DB.withConnection(user.database.database) { implicit connection =>
      SQL(s"SELECT t1.*, t2.${C.FK_NAME} FROM ${C.ODCONSTRUCTORS} t1 INNER JOIN ${C.???} t2 ON t1.${C.FK_ID} = t2.${C.FK_ID} WHERE $where")
        .on(args: _*)
        .as(odconstructorsParser.singleOpt)
    }
  }

  def delete(user: CompanyUser, id: Int) = {
    DB.withConnection(user.database.database) { implicit connection =>
      try {
        SQL(s"DELETE FROM ${C.ODCONSTRUCTORS} WHERE ${C.ID} = {id}")
        .on('id -> id)
        .executeUpdate
      } catch {
        case NonFatal(e) => {
          Logger.error(s"can't delete a ${C.ODCONSTRUCTORS}. ${C.ID} id: $id. Error ${e.toString}")
          throw e
        }
      }
    }
  }

  def get(user: CompanyUser, id: Int): Odconstructors = {
    load(user, id) match {
      case Some(p) => p
      case None => emptyOdconstructors
    }
  }
}


// Router

GET     /api/v1/general/odconstructors/:id              controllers.logged.modules.general.Odconstructorss.get(id: Int)
POST    /api/v1/general/odconstructors                  controllers.logged.modules.general.Odconstructorss.create
PUT     /api/v1/general/odconstructors/:id              controllers.logged.modules.general.Odconstructorss.update(id: Int)
DELETE  /api/v1/general/odconstructors/:id              controllers.logged.modules.general.Odconstructorss.delete(id: Int)




/**/
