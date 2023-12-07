public class modColorDepthAPI {

  //******************************************************************************
  // Module     : modColorDepthAPI
  // Description: Used by Desktop Technology's standard Splash screen
  // Procedures : N/A
  // Modified   :
  //
  // --------------------------------------------------
//Option Explicit
  *Option Compare Binary

//*TODO:** type is translated as a new class at the end of the file Public Type BITMAP

*TODO: API Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

}

public class BITMAP {
    public Long bmType;
    public Long bmWidth;
    public Long bmHeight;
    public Long bmWidthBytes;
    public Integer bmPlanes;
    public Integer bmBitsPixel;
    public Long bmBits;
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


case class OdcolordepthapiData(
              id: Option[Int],

              )

object Odcolordepthapis extends Controller with ProvidesUser {

  val odcolordepthapiForm = Form(
    mapping(
      "id" -> optional(number),

  )(OdcolordepthapiData.apply)(OdcolordepthapiData.unapply))

  implicit val odcolordepthapiWrites = new Writes[Odcolordepthapi] {
    def writes(odcolordepthapi: Odcolordepthapi) = Json.obj(
      "id" -> Json.toJson(odcolordepthapi.id),
      C.ID -> Json.toJson(odcolordepthapi.id),

    )
  }

  def get(id: Int) = GetAction { implicit request =>
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.LIST_ODCOLORDEPTHAPI), { user =>
      Ok(Json.toJson(Odcolordepthapi.get(user, id)))
    })
  }

  def update(id: Int) = PostAction { implicit request =>
    Logger.debug("in odcolordepthapis.update")
    odcolordepthapiForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      odcolordepthapi => {
        Logger.debug(s"form: ${odcolordepthapi.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.EDIT_ODCOLORDEPTHAPI), { user =>
          Ok(
            Json.toJson(
              Odcolordepthapi.update(user,
                Odcolordepthapi(
                       id,

                ))))
        })
      }
    )
  }

  def create = PostAction { implicit request =>
    Logger.debug("in odcolordepthapis.create")
    odcolordepthapiForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      odcolordepthapi => {
        Logger.debug(s"form: ${odcolordepthapi.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.NEW_ODCOLORDEPTHAPI), { user =>
          Ok(
            Json.toJson(
              Odcolordepthapi.create(user,
                Odcolordepthapi(

                ))))
        })
      }
    )
  }

  def delete(id: Int) = PostAction { implicit request =>
    Logger.debug("in odcolordepthapis.delete")
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.DELETE_ODCOLORDEPTHAPI), { user =>
      Odcolordepthapi.delete(user, id)
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

case class Odcolordepthapi(
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

object Odcolordepthapi {

  lazy val emptyOdcolordepthapi = Odcolordepthapi(
)

  def apply(
      id: Int,
) = {

    new Odcolordepthapi(
      id,
)
  }

  def apply(
) = {

    new Odcolordepthapi(
)
  }

  private val odcolordepthapiParser: RowParser[Odcolordepthapi] = {
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
        Odcolordepthapi(
              id,
,
              createdAt,
              updatedAt,
              updatedBy)
    }
  }

  def create(user: CompanyUser, odcolordepthapi: Odcolordepthapi): Odcolordepthapi = {
    save(user, odcolordepthapi, true)
  }

  def update(user: CompanyUser, odcolordepthapi: Odcolordepthapi): Odcolordepthapi = {
    save(user, odcolordepthapi, false)
  }

  private def save(user: CompanyUser, odcolordepthapi: Odcolordepthapi, isNew: Boolean): Odcolordepthapi = {
    def getFields = {
      List(

      )
    }
    def throwException = {
      throw new RuntimeException(s"Error when saving ${C.ODCOLORDEPTHAPI}")
    }

    DBHelper.saveEx(
      user,
      Register(
        C.ODCOLORDEPTHAPI,
        C.ID,
        odcolordepthapi.id,
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

  def load(user: CompanyUser, id: Int): Option[Odcolordepthapi] = {
    loadWhere(user, s"${C.ID} = {id}", 'id -> id)
  }

  def loadWhere(user: CompanyUser, where: String, args : scala.Tuple2[scala.Any, anorm.ParameterValue[_]]*) = {
    DB.withConnection(user.database.database) { implicit connection =>
      SQL(s"SELECT t1.*, t2.${C.FK_NAME} FROM ${C.ODCOLORDEPTHAPI} t1 INNER JOIN ${C.???} t2 ON t1.${C.FK_ID} = t2.${C.FK_ID} WHERE $where")
        .on(args: _*)
        .as(odcolordepthapiParser.singleOpt)
    }
  }

  def delete(user: CompanyUser, id: Int) = {
    DB.withConnection(user.database.database) { implicit connection =>
      try {
        SQL(s"DELETE FROM ${C.ODCOLORDEPTHAPI} WHERE ${C.ID} = {id}")
        .on('id -> id)
        .executeUpdate
      } catch {
        case NonFatal(e) => {
          Logger.error(s"can't delete a ${C.ODCOLORDEPTHAPI}. ${C.ID} id: $id. Error ${e.toString}")
          throw e
        }
      }
    }
  }

  def get(user: CompanyUser, id: Int): Odcolordepthapi = {
    load(user, id) match {
      case Some(p) => p
      case None => emptyOdcolordepthapi
    }
  }
}


// Router

GET     /api/v1/general/odcolordepthapi/:id              controllers.logged.modules.general.Odcolordepthapis.get(id: Int)
POST    /api/v1/general/odcolordepthapi                  controllers.logged.modules.general.Odcolordepthapis.create
PUT     /api/v1/general/odcolordepthapi/:id              controllers.logged.modules.general.Odcolordepthapis.update(id: Int)
DELETE  /api/v1/general/odcolordepthapi/:id              controllers.logged.modules.general.Odcolordepthapis.delete(id: Int)




/**/
