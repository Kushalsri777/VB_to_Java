public class chrgHourglass {

  //******************************************************************************
  // Module     : chrgHourglass
  // Description: This class implements an alternative way to show an hourglass
  // Procedures :
  //              Property Get Value() - public
  //              Property Let Value() - public
  //              Class_Terminate()
  // Source     : Total Visual SourceBook 2000
  // Modified   :
  // 03/03/02 BAW (Phase2A) Added support for new global error handler'
  // -------------------------------------------------------------------------------
//Option Explicit
  *Option Compare Binary
  private static final String MCSTRNAME = "chrgHourglass.";

  //------------------------------------------
  //            MEMBER VARIABLES
  //------------------------------------------
  //'local copy
  private boolean m_bValue = false;


  //------------------------------------------
  //           PROPERTY GET / LET
  //------------------------------------------
  public boolean getValue() {
    boolean _rtn = false;
    // Comments  : Returns True if the cursor is shown as
    //             an hourglass; False otherwise
    // Parameters: None
    // Modified  :
    // Source    : Total Visual SourceBook 2000
    // --------------------------------------------------
    try {
      "Property Get Value"
.equals(Const cstrCurrentProc As String);

      _rtn = m_bValue;
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
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



  public void setValue(boolean bValue) {
    boolean _rtn = null;
    // Comments  : Changes the cursor to/from an hourglass
    // Parameters: bValue (in) - True to turn the cursor into
    //                  an hourglass; False to set it to the
    //                  default
    // Modified  :
    // Source    : Total Visual SourceBook 2000
    // --------------------------------------------------
    try {
      "Property Let Value"
.equals(Const cstrCurrentProc As String);

      m_bValue = bValue;
      if (m_bValue) {
        Screen.MousePointer = vbHourglass;
      } 
      else {
        Screen.MousePointer = vbDefault;
      }
      // **TODO:** label found: PROC_EXIT:;
      // BAW 03/25/2002 - We might get here IN THE COURSE OF reporting an error from
      //                  an event handler, e.g., as a long-winded process has ended
      //                  due to an error. So, ignore all errors so we don't
      //                  get into a sort of loop trying to propagate an error back
      //                  to the Event Handler that called us. This is why
      //                  On Error Resume Next is used instead of On Error GoTo 0.
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    Exit Property;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
}

  return _rtn;
}



// ********************************************
//  Initialize and Terminate
// ********************************************
  private void class_Terminate() {
    // Comments  : Reset the cursor to the default
    // Parameters: None
    // Modified  :
    // Source    : Total Visual SourceBook 2000
    // --------------------------------------------------
    try {
      "Class_Terminate"
.equals(Const cstrCurrentProc As String);

      Screen.MousePointer = vbDefault;
      // **TODO:** label found: PROC_EXIT:;
      // BAW 03/25/2002 - We might get here IN THE COURSE OF reporting an error from
      //                  an event handler, e.g., as a long-winded process has ended
      //                  due to an error. So, ignore all errors so we don't
      //                  get into a sort of loop trying to propagate an error back
      //                  to the Event Handler that called us. This is why
      //                  On Error Resume Next is used instead of On Error GoTo 0.
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


case class HrghourglassData(
              id: Option[Int],

              )

object Hrghourglasss extends Controller with ProvidesUser {

  val hrghourglassForm = Form(
    mapping(
      "id" -> optional(number),

  )(HrghourglassData.apply)(HrghourglassData.unapply))

  implicit val hrghourglassWrites = new Writes[Hrghourglass] {
    def writes(hrghourglass: Hrghourglass) = Json.obj(
      "id" -> Json.toJson(hrghourglass.id),
      C.ID -> Json.toJson(hrghourglass.id),

    )
  }

  def get(id: Int) = GetAction { implicit request =>
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.LIST_HRGHOURGLASS), { user =>
      Ok(Json.toJson(Hrghourglass.get(user, id)))
    })
  }

  def update(id: Int) = PostAction { implicit request =>
    Logger.debug("in hrghourglasss.update")
    hrghourglassForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      hrghourglass => {
        Logger.debug(s"form: ${hrghourglass.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.EDIT_HRGHOURGLASS), { user =>
          Ok(
            Json.toJson(
              Hrghourglass.update(user,
                Hrghourglass(
                       id,

                ))))
        })
      }
    )
  }

  def create = PostAction { implicit request =>
    Logger.debug("in hrghourglasss.create")
    hrghourglassForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      hrghourglass => {
        Logger.debug(s"form: ${hrghourglass.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.NEW_HRGHOURGLASS), { user =>
          Ok(
            Json.toJson(
              Hrghourglass.create(user,
                Hrghourglass(

                ))))
        })
      }
    )
  }

  def delete(id: Int) = PostAction { implicit request =>
    Logger.debug("in hrghourglasss.delete")
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.DELETE_HRGHOURGLASS), { user =>
      Hrghourglass.delete(user, id)
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

case class Hrghourglass(
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

object Hrghourglass {

  lazy val emptyHrghourglass = Hrghourglass(
)

  def apply(
      id: Int,
) = {

    new Hrghourglass(
      id,
)
  }

  def apply(
) = {

    new Hrghourglass(
)
  }

  private val hrghourglassParser: RowParser[Hrghourglass] = {
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
        Hrghourglass(
              id,
,
              createdAt,
              updatedAt,
              updatedBy)
    }
  }

  def create(user: CompanyUser, hrghourglass: Hrghourglass): Hrghourglass = {
    save(user, hrghourglass, true)
  }

  def update(user: CompanyUser, hrghourglass: Hrghourglass): Hrghourglass = {
    save(user, hrghourglass, false)
  }

  private def save(user: CompanyUser, hrghourglass: Hrghourglass, isNew: Boolean): Hrghourglass = {
    def getFields = {
      List(

      )
    }
    def throwException = {
      throw new RuntimeException(s"Error when saving ${C.HRGHOURGLASS}")
    }

    DBHelper.saveEx(
      user,
      Register(
        C.HRGHOURGLASS,
        C.ID,
        hrghourglass.id,
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

  def load(user: CompanyUser, id: Int): Option[Hrghourglass] = {
    loadWhere(user, s"${C.ID} = {id}", 'id -> id)
  }

  def loadWhere(user: CompanyUser, where: String, args : scala.Tuple2[scala.Any, anorm.ParameterValue[_]]*) = {
    DB.withConnection(user.database.database) { implicit connection =>
      SQL(s"SELECT t1.*, t2.${C.FK_NAME} FROM ${C.HRGHOURGLASS} t1 INNER JOIN ${C.???} t2 ON t1.${C.FK_ID} = t2.${C.FK_ID} WHERE $where")
        .on(args: _*)
        .as(hrghourglassParser.singleOpt)
    }
  }

  def delete(user: CompanyUser, id: Int) = {
    DB.withConnection(user.database.database) { implicit connection =>
      try {
        SQL(s"DELETE FROM ${C.HRGHOURGLASS} WHERE ${C.ID} = {id}")
        .on('id -> id)
        .executeUpdate
      } catch {
        case NonFatal(e) => {
          Logger.error(s"can't delete a ${C.HRGHOURGLASS}. ${C.ID} id: $id. Error ${e.toString}")
          throw e
        }
      }
    }
  }

  def get(user: CompanyUser, id: Int): Hrghourglass = {
    load(user, id) match {
      case Some(p) => p
      case None => emptyHrghourglass
    }
  }
}


// Router

GET     /api/v1/general/hrghourglass/:id              controllers.logged.modules.general.Hrghourglasss.get(id: Int)
POST    /api/v1/general/hrghourglass                  controllers.logged.modules.general.Hrghourglasss.create
PUT     /api/v1/general/hrghourglass/:id              controllers.logged.modules.general.Hrghourglasss.update(id: Int)
DELETE  /api/v1/general/hrghourglass/:id              controllers.logged.modules.general.Hrghourglasss.delete(id: Int)




/**/
