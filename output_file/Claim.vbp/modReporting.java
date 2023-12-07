public class modReporting {

  //******************************************************************************
  // Module     : modReporting
  // Description:
  // Procedures :
  //
  //
  // Modified   :
  // 04/30/02 BAW Made the Crystal Application object a global variable:
  //              defined in modReporting; instantiated in modStartup; deallocated in
  //              fnDeallocateGlobalObjects. This avoids "Out of memory" errors
  //              when the frmReportViewer screen is displayed.
  // --------------------------------------------------
//Option Explicit
  *Option Compare Binary
  private static final String MCSTRNAME = "modReporting.";


  public static CRAXDRT.Report gcReportToPrint;
  public static CRAXDRT.Application gcrxApp;


  //////////////////////////////////////////////////////////////////////////////////////////////////
  public static boolean fnSetFormulaField(String strFormulaName, String strFormulaText) {
    boolean _rtn = false;
    // Comments  : Sets the value of the named Crystal .RPT formula field. Derived from
    //             p592 of George Peck's "Crystal Reports 8: The Complete Reference" book.
    //
    //             NOTE: Assumes the caller set gcReportToPrint to point to the
    //                   correct .RPT file prior to calling this procedure.
    //
    //             NOTE 2: The formulae names are CASE-SENSITIVE ! ! !
    //
    // Parameters: strFormulaName (in) = the name of the formula field (without a "@")
    //             strFormulaText (in) = the text of the formula, in Crystal syntax
    //
    // Returns   : True if named formula was found and updated; False otherwise
    //
    // Called by : fnPrepare_xxx( ) of frmPrintReports
    //             cmdOK_Click( ) of frmInsured
    //
    // Modified  :
    //
    // --------------------------------------------------
    try {
      "fnSetFormulaField"
.equals(Const cstrCurrentProc As String);
      int intCounter = 0;

      for (intCounter = 1; intCounter <= gcReportToPrint.FormulaFields.LinkedMap.size(); intCounter++) {
        if (gcReportToPrint.FormulaFields(intCounter).FormulaFieldName == strFormulaName) {
          gcReportToPrint.FormulaFields(intCounter).Text = modGeneral.fnQuoted(strFormulaText);
          _rtn = true;
          // **TODO:** goto found: GoTo PROC_EXIT;
        }
      }
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
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
        // Save Err object data, if not already saved
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}





//////////////////////////////////////////////////////////////////////////////////////////////////
  public static void fnViewReport() {
    // Comments  : Displays the report in the Report Viewer window.
    //             It assumes that the all of the Report's properties
    //             were set appropriately (e.g., RecordSelectionFormula,
    //             SetDataSource, etc.) before this routine was called.
    // Parameters: N/A
    //
    // Called by : cmdOK_Click( ) in frmSelectReports
    //
    // Returns   : N/A
    //
    // Modified  :
    // --------------------------------------------------
    try {
      "fnViewReport"
.equals(Const cstrCurrentProc As String);
      Form frmChild = null;

      // Report Viewer Form's Load event will automatically display report.
      frmChild = new frmReportViewer();
      frmChild.Show(vbModal);
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    // Terminate the Crystal Report Viewer window, removing it from the Forms collection
    modGeneral.fnFreeObject(frmChild);

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


case class OdreportingData(
              id: Option[Int],

              )

object Odreportings extends Controller with ProvidesUser {

  val odreportingForm = Form(
    mapping(
      "id" -> optional(number),

  )(OdreportingData.apply)(OdreportingData.unapply))

  implicit val odreportingWrites = new Writes[Odreporting] {
    def writes(odreporting: Odreporting) = Json.obj(
      "id" -> Json.toJson(odreporting.id),
      C.ID -> Json.toJson(odreporting.id),

    )
  }

  def get(id: Int) = GetAction { implicit request =>
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.LIST_ODREPORTING), { user =>
      Ok(Json.toJson(Odreporting.get(user, id)))
    })
  }

  def update(id: Int) = PostAction { implicit request =>
    Logger.debug("in odreportings.update")
    odreportingForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      odreporting => {
        Logger.debug(s"form: ${odreporting.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.EDIT_ODREPORTING), { user =>
          Ok(
            Json.toJson(
              Odreporting.update(user,
                Odreporting(
                       id,

                ))))
        })
      }
    )
  }

  def create = PostAction { implicit request =>
    Logger.debug("in odreportings.create")
    odreportingForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      odreporting => {
        Logger.debug(s"form: ${odreporting.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.NEW_ODREPORTING), { user =>
          Ok(
            Json.toJson(
              Odreporting.create(user,
                Odreporting(

                ))))
        })
      }
    )
  }

  def delete(id: Int) = PostAction { implicit request =>
    Logger.debug("in odreportings.delete")
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.DELETE_ODREPORTING), { user =>
      Odreporting.delete(user, id)
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

case class Odreporting(
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

object Odreporting {

  lazy val emptyOdreporting = Odreporting(
)

  def apply(
      id: Int,
) = {

    new Odreporting(
      id,
)
  }

  def apply(
) = {

    new Odreporting(
)
  }

  private val odreportingParser: RowParser[Odreporting] = {
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
        Odreporting(
              id,
,
              createdAt,
              updatedAt,
              updatedBy)
    }
  }

  def create(user: CompanyUser, odreporting: Odreporting): Odreporting = {
    save(user, odreporting, true)
  }

  def update(user: CompanyUser, odreporting: Odreporting): Odreporting = {
    save(user, odreporting, false)
  }

  private def save(user: CompanyUser, odreporting: Odreporting, isNew: Boolean): Odreporting = {
    def getFields = {
      List(

      )
    }
    def throwException = {
      throw new RuntimeException(s"Error when saving ${C.ODREPORTING}")
    }

    DBHelper.saveEx(
      user,
      Register(
        C.ODREPORTING,
        C.ID,
        odreporting.id,
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

  def load(user: CompanyUser, id: Int): Option[Odreporting] = {
    loadWhere(user, s"${C.ID} = {id}", 'id -> id)
  }

  def loadWhere(user: CompanyUser, where: String, args : scala.Tuple2[scala.Any, anorm.ParameterValue[_]]*) = {
    DB.withConnection(user.database.database) { implicit connection =>
      SQL(s"SELECT t1.*, t2.${C.FK_NAME} FROM ${C.ODREPORTING} t1 INNER JOIN ${C.???} t2 ON t1.${C.FK_ID} = t2.${C.FK_ID} WHERE $where")
        .on(args: _*)
        .as(odreportingParser.singleOpt)
    }
  }

  def delete(user: CompanyUser, id: Int) = {
    DB.withConnection(user.database.database) { implicit connection =>
      try {
        SQL(s"DELETE FROM ${C.ODREPORTING} WHERE ${C.ID} = {id}")
        .on('id -> id)
        .executeUpdate
      } catch {
        case NonFatal(e) => {
          Logger.error(s"can't delete a ${C.ODREPORTING}. ${C.ID} id: $id. Error ${e.toString}")
          throw e
        }
      }
    }
  }

  def get(user: CompanyUser, id: Int): Odreporting = {
    load(user, id) match {
      case Some(p) => p
      case None => emptyOdreporting
    }
  }
}


// Router

GET     /api/v1/general/odreporting/:id              controllers.logged.modules.general.Odreportings.get(id: Int)
POST    /api/v1/general/odreporting                  controllers.logged.modules.general.Odreportings.create
PUT     /api/v1/general/odreporting/:id              controllers.logged.modules.general.Odreportings.update(id: Int)
DELETE  /api/v1/general/odreporting/:id              controllers.logged.modules.general.Odreportings.delete(id: Int)




/**/
