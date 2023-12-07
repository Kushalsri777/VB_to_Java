public class frmReportViewer {

  //******************************************************************************
  // Module     : frmReportViewer
  // Description:
  // Procedures:
  //              Form_Load)
  //              Form_Resize()
  //              Form_Unload(ByRef pintCancel As Integer)
  //
  // Modified   :
  //
  // --------------------------------------------------
//Option Explicit
  *Option Compare Binary

  private String mstrScreenName = "";
  private static final Long MCLNGMINFORMWIDTH = 10905;
  private static final Long MCLNGMINFORMHEIGHT = 10530;




  // ' member variable for ReportToPrint property
  // Private m_ReportToPrint As Object

  // ' Used by other forms (such as frmInsured and frmPrintReport) to print a Crystal Report .RPT file
  // Property Get ReportToPrint() As Object
  //     Set ReportToPrint = m_ReportToPrint
  // End Property
  // Property Set ReportToPrint(ByVal newValue As Object)
  //     Set m_ReportToPrint = newValue
  // End Property


  private void form_Load() {
    // Comments  : Open the requested report in a modal
    //             Crystal Report 8 viewer window.
    // Parameters: None
    // Modified  :
    // --------------------------------------------------
    try {
      "Form_Load"
.equals(Const cstrCurrentProc As String);
      chrgHourglass hrgHourglass = null;

      // Set the screen name that will be used to form the Title on message boxes
      mstrScreenName = Me.Caption;

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      Me.Icon = LoadResPicture(modResConstants.gCRES_ICON_MAINAPP, vbResIcon);

      // If the user has ever opened this form before, restore its size & placement.
      // If the restore would result in the form being off-screen, just center it instead.
      if (modGeneral.gapsApp.restoreForm(this) == false) {
        //fnSetFormSize
        //*TODO:** can't found type for with block
        //*With this
        __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = this;
        w___TYPE_NOT_FOUND.Width = MCLNGMINFORMWIDTH;
        w___TYPE_NOT_FOUND.Height = MCLNGMINFORMHEIGHT;
        //fnCenterFormOnScreen Me
        modGeneral.fnCenterFormOnMDI(frmMDIMain, this);
      }

      hrgHourglass = new chrgHourglass();
      hrgHourglass.setValue(true);

      // Set ReportSource using a Property Get on frmPrintReports2
      crxViewer.ReportSource = modReporting.gcReportToPrint;
      crxViewer.EnableCloseButton = true;

      // View the report
      crxViewer.ViewReport;
      // Or, print it without viewing...
      //       crxViewer.PrintReport

      hrgHourglass.setValue(false);
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
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



  private void form_Resize() {
    // Comments  : Resize the viewer control so it fills
    //             the form window
    // Parameters: None
    // Modified  :
    // --------------------------------------------------
    try {
      "Form_Resize"
.equals(Const cstrCurrentProc As String);

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      //With Me
      //    If .Width < mclngMinFormWidth Then
      //        .Width = mclngMinFormWidth
      //    End If
      //    If .Height < mclngMinFormHeight Then
      //        .Height = mclngMinFormHeight
      //    End If
      //End With

      crxViewer.Top = 0;
      crxViewer.Left = 0;
      crxViewer.Height = ScaleHeight;
      crxViewer.Width = ScaleWidth;
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
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



  private void form_Unload(int pintCancel) { // TODO: Use of ByRef founded Private Sub Form_Unload(ByRef pintCancel As Integer)
    // Comments  : Unloads the form
    // Parameters: None
    // Modified  :
    // --------------------------------------------------
    try {
      "Form_Unload"
.equals(Const cstrCurrentProc As String);

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      modGeneral.gapsApp.saveForm(this);
      Unload(this);
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
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


case class RmreportviewerData(
              id: Option[Int],

              )

object Rmreportviewers extends Controller with ProvidesUser {

  val rmreportviewerForm = Form(
    mapping(
      "id" -> optional(number),

  )(RmreportviewerData.apply)(RmreportviewerData.unapply))

  implicit val rmreportviewerWrites = new Writes[Rmreportviewer] {
    def writes(rmreportviewer: Rmreportviewer) = Json.obj(
      "id" -> Json.toJson(rmreportviewer.id),
      C.ID -> Json.toJson(rmreportviewer.id),

    )
  }

  def get(id: Int) = GetAction { implicit request =>
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.LIST_RMREPORTVIEWER), { user =>
      Ok(Json.toJson(Rmreportviewer.get(user, id)))
    })
  }

  def update(id: Int) = PostAction { implicit request =>
    Logger.debug("in rmreportviewers.update")
    rmreportviewerForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      rmreportviewer => {
        Logger.debug(s"form: ${rmreportviewer.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.EDIT_RMREPORTVIEWER), { user =>
          Ok(
            Json.toJson(
              Rmreportviewer.update(user,
                Rmreportviewer(
                       id,

                ))))
        })
      }
    )
  }

  def create = PostAction { implicit request =>
    Logger.debug("in rmreportviewers.create")
    rmreportviewerForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      rmreportviewer => {
        Logger.debug(s"form: ${rmreportviewer.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.NEW_RMREPORTVIEWER), { user =>
          Ok(
            Json.toJson(
              Rmreportviewer.create(user,
                Rmreportviewer(

                ))))
        })
      }
    )
  }

  def delete(id: Int) = PostAction { implicit request =>
    Logger.debug("in rmreportviewers.delete")
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.DELETE_RMREPORTVIEWER), { user =>
      Rmreportviewer.delete(user, id)
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

case class Rmreportviewer(
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

object Rmreportviewer {

  lazy val emptyRmreportviewer = Rmreportviewer(
)

  def apply(
      id: Int,
) = {

    new Rmreportviewer(
      id,
)
  }

  def apply(
) = {

    new Rmreportviewer(
)
  }

  private val rmreportviewerParser: RowParser[Rmreportviewer] = {
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
        Rmreportviewer(
              id,
,
              createdAt,
              updatedAt,
              updatedBy)
    }
  }

  def create(user: CompanyUser, rmreportviewer: Rmreportviewer): Rmreportviewer = {
    save(user, rmreportviewer, true)
  }

  def update(user: CompanyUser, rmreportviewer: Rmreportviewer): Rmreportviewer = {
    save(user, rmreportviewer, false)
  }

  private def save(user: CompanyUser, rmreportviewer: Rmreportviewer, isNew: Boolean): Rmreportviewer = {
    def getFields = {
      List(

      )
    }
    def throwException = {
      throw new RuntimeException(s"Error when saving ${C.RMREPORTVIEWER}")
    }

    DBHelper.saveEx(
      user,
      Register(
        C.RMREPORTVIEWER,
        C.ID,
        rmreportviewer.id,
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

  def load(user: CompanyUser, id: Int): Option[Rmreportviewer] = {
    loadWhere(user, s"${C.ID} = {id}", 'id -> id)
  }

  def loadWhere(user: CompanyUser, where: String, args : scala.Tuple2[scala.Any, anorm.ParameterValue[_]]*) = {
    DB.withConnection(user.database.database) { implicit connection =>
      SQL(s"SELECT t1.*, t2.${C.FK_NAME} FROM ${C.RMREPORTVIEWER} t1 INNER JOIN ${C.???} t2 ON t1.${C.FK_ID} = t2.${C.FK_ID} WHERE $where")
        .on(args: _*)
        .as(rmreportviewerParser.singleOpt)
    }
  }

  def delete(user: CompanyUser, id: Int) = {
    DB.withConnection(user.database.database) { implicit connection =>
      try {
        SQL(s"DELETE FROM ${C.RMREPORTVIEWER} WHERE ${C.ID} = {id}")
        .on('id -> id)
        .executeUpdate
      } catch {
        case NonFatal(e) => {
          Logger.error(s"can't delete a ${C.RMREPORTVIEWER}. ${C.ID} id: $id. Error ${e.toString}")
          throw e
        }
      }
    }
  }

  def get(user: CompanyUser, id: Int): Rmreportviewer = {
    load(user, id) match {
      case Some(p) => p
      case None => emptyRmreportviewer
    }
  }
}


// Router

GET     /api/v1/general/rmreportviewer/:id              controllers.logged.modules.general.Rmreportviewers.get(id: Int)
POST    /api/v1/general/rmreportviewer                  controllers.logged.modules.general.Rmreportviewers.create
PUT     /api/v1/general/rmreportviewer/:id              controllers.logged.modules.general.Rmreportviewers.update(id: Int)
DELETE  /api/v1/general/rmreportviewer/:id              controllers.logged.modules.general.Rmreportviewers.delete(id: Int)




/**/
