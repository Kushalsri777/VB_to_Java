public class frmMsgBox {

  //******************************************************************************
  // Module     : frmMsgBox
  // Description: This form is opened when cerhErrorHandler class needs to report a
  //              message of any type (info, warning, alert, or error)
  // Procedures :
  //    Public    Property Get ButtonClicked() As Long
  //    Public    Property Let ErrorCode(ByVal lngValue As Long)
  //    Public    Property Let ErrorContext(ByVal strValue As String)
  //    Public    Property Let MsgText(ByVal strValue As String)
  //    Public    Property Let ScreenName(ByVal strValue As String)
  //    Private   cmdButton1_Click()
  //    Private   cmdButton2_Click()
  //    Private   cmdDetailsCollapse_Click()
  //    Private   cmdDetailsExpand_Click()
  //    Private   fnSetButtons(ByVal lngValue As eMsgButtons)
  //    Private   fnSetupForDefaultDisplay()
  //    Private   Form_Load()
  //    Private   Form_Unload(ByRef pintCancel As Integer)

  // Modified   :
  // 03/19/02 BAW (Phase2A) Created form
  //
  // --------------------------------------------------
//Option Explicit
  *Option Compare Binary
  private static final Long MCLNGMINFORMWIDTH = 7740;
  private static final Long MCLNGMINFORMHEIGHT_COLLAPSED = 2985;
  private static final Long MCLNGMINFORMHEIGHT_EXPANDED = 4815;
  private static final String MCSTROK = "OK";
  private static final String MCSTRCANCEL = "Cancel";
  private static final String MCSTRYES = "Yes";
  private static final String MCSTRNO = "No";


  // Private variables to hold public properties
  private int m_lngButtonClicked = 0;
  private String m_strMsgText = "";
  private String m_lngErrorCode = "";
  private String m_strErrorContext = "";
  private String m_strScreenName = "";

//*TODO:** enum is translated as a new class at the end of the file Public Enum eMsgButtons


  ///\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
  //\                                                                      /
  //|                         Public Properties                            |
  ///                                                                      \
  ///\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
  public int getButtonClicked() {
    // Returns the ButtonClicked property setting
    return m_lngButtonClicked;
  }

  public void setErrorCode(int lngValue) {
    // Sets the ErrorCode property to the value specified by lngValue.
    // Note that lngValue is a *translated* error code. For an
    // application-specific error, this means that vbObjectError
    // has been subtracted from it already.

    m_lngErrorCode = lngValue;
    lblErrorCode.Caption = CStr(lngValue);

    switch (lngValue) {
      case  modResConstants.gCRES_INFO_START To modResConstants.gCRES_INFO_END:
        fnSetButtons(eMsgButtons.eBTNOKONLY);
        Me.Caption = m_strScreenName+ " - Information";
        imgIcon.Picture = LoadResPicture(modResConstants.gCRES_ICON_INFO, vbResIcon);
        break;

      case  modResConstants.gCRES_WARN_START To modResConstants.gCRES_WARN_END:
        fnSetButtons(eMsgButtons.eBTNOKONLY);
        Me.Caption = m_strScreenName+ " - Warning";
        imgIcon.Picture = LoadResPicture(modResConstants.gCRES_ICON_WARN, vbResIcon);
        break;

      case  modResConstants.gCRES_ALRT_START To modResConstants.gCRES_ALRT_END:
        fnSetButtons(eMsgButtons.eBTNYESNO);
        Me.Caption = m_strScreenName+ " - Alert";
        imgIcon.Picture = LoadResPicture(modResConstants.gCRES_ICON_ALRT, vbResIcon);
        break;

      case  modResConstants.gCRES_NERR_START To modResConstants.gCRES_NERR_END:
        fnSetButtons(eMsgButtons.eBTNOKONLY);
        Me.Caption = m_strScreenName+ " - Error";
        imgIcon.Picture = LoadResPicture(modResConstants.gCRES_ICON_ERR, vbResIcon);
        break;

      default:
        // Intended to be 9000 to 9999 ... or a VB or ADO error code
        fnSetButtons(eMsgButtons.eBTNOKONLY);
        Me.Caption = m_strScreenName+ " - Fatal Error";
        imgIcon.Picture = LoadResPicture(modResConstants.gCRES_ICON_ERR, vbResIcon);
        break;
    }
  }

  public void setErrorContext(String strValue) {
    // Sets the ErrorContext property to the value specified by strValue
    m_strErrorContext = strValue;
    lblErrorContext.Caption = strValue;
  }

  public void setMsgText(String strValue) {
    // Sets the MsgText property to the value specified by strValue
    m_strMsgText = strValue;
    txtMsgText.Text = strValue;
  }

  public void setScreenName(String strValue) {
    // Sets the ScreenName property to the value specified by strValue
    m_strScreenName = strValue;
  }



  ///\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
  //\                                                                      /
  //|                           Public Methods                             |
  ///                                                                      \
  ///\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
  //None


  ///\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
  //\                                                                      /
  //|                            Private Methods                           |
  ///                                                                      \
  ///\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\

  ///////////////////////////////////////////////////////////////////////////////////
  private void cmdButton1_Click() {
    switch (cmdButton1.Caption) {
      case  MCSTROK:
        m_lngButtonClicked = vbOK;
        break;

      case  MCSTRYES:
        m_lngButtonClicked = vbYes;
        break;

      case  MCSTRCANCEL:
        m_lngButtonClicked = vbCancel;
      //' mcstrNO
        break;

      default:
        m_lngButtonClicked = vbNo;
        break;
    }

    Me.Hide;
  }



  ///////////////////////////////////////////////////////////////////////////////////
  private void cmdButton2_Click() {
    //--------------------------------------------------------------------
    // Procedure : cmdButton2_Click
    // Comments  : Set form property that the caller can interrogate, if desired.
    //             Note: the caller should unload this form after that interrogation.
    //
    // Called by : N/A
    // Parameters: N/A
    // Modified  :
    //--------------------------------------------------------------------
    switch (cmdButton2.Caption) {
      case  MCSTROK:
        m_lngButtonClicked = vbOK;
        break;

      case  MCSTRYES:
        m_lngButtonClicked = vbYes;
        break;

      case  MCSTRCANCEL:
        m_lngButtonClicked = vbCancel;
      //' mcstrNO
        break;

      default:
        m_lngButtonClicked = vbNo;
        break;
    }

    Me.Hide;
  }



  ///////////////////////////////////////////////////////////////////////////////////
  private void cmdDetailsCollapse_Click() {
    //--------------------------------------------------------------------
    // Procedure : cmdDetailsCollapse_Click
    // Comments  : Set form and control properties based on the user
    //             indicating they want a "normal" display
    //
    // Called by : N/A
    // Parameters: N/A
    // Modified  :
    //--------------------------------------------------------------------
    fnSetupForDefaultDisplay();
  }



  ///////////////////////////////////////////////////////////////////////////////////
  private void cmdDetailsExpand_Click() {
    //--------------------------------------------------------------------
    // Procedure : cmdDetailsExpand_Click
    // Comments  : Set form and control properties based on the user
    //             indicating they want a "Details" display
    //
    // Called by : N/A
    // Parameters: N/A
    // Modified  :
    //--------------------------------------------------------------------
    fnWindowLock(Me.cbrfBrowseFolder.setHWnd());
    Me.Height = MCLNGMINFORMHEIGHT_EXPANDED;
    Me.Width = MCLNGMINFORMWIDTH;

    fraDetails.Visible = true;
    cmdDetailsExpand.Visible = false;
    cmdDetailsExpand.Enabled = false;
    cmdDetailsCollapse.Visible = true;
    cmdDetailsCollapse.Enabled = true;
    fnWindowUnlock;
  }



  ///////////////////////////////////////////////////////////////////////////////////
  public void fnSetButtons(eMsgButtons lngValue) {
    // Sets the placement and caption of the form's main command buttons.
    Const(lngButtonSpacer As Long == 60);

    //fnWindowLock Me.hWnd
    switch (lngValue) {
      case  eMsgButtons.eBTNYESNO:
        //*TODO:** can't found type for with block
        //*With cmdButton2
        __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = cmdButton2;
        w___TYPE_NOT_FOUND.Visible = true;
        w___TYPE_NOT_FOUND.Default = false;
        w___TYPE_NOT_FOUND.Caption = MCSTRNO;
        w___TYPE_NOT_FOUND.Left = fraMainButtons.Width - w___TYPE_NOT_FOUND.Width - lngButtonSpacer;
        //*TODO:** can't found type for with block
        //*With cmdButton1
        __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = cmdButton1;
        w___TYPE_NOT_FOUND.Visible = true;
        w___TYPE_NOT_FOUND.Caption = MCSTRYES;
        w___TYPE_NOT_FOUND.Default = true;
        w___TYPE_NOT_FOUND.Left = lngButtonSpacer;

        break;

      case  eMsgButtons.eBTNOKCANCEL:
        //*TODO:** can't found type for with block
        //*With cmdButton2
        __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = cmdButton2;
        w___TYPE_NOT_FOUND.Visible = true;
        w___TYPE_NOT_FOUND.Default = false;
        w___TYPE_NOT_FOUND.Caption = MCSTRCANCEL;
        w___TYPE_NOT_FOUND.Left = fraMainButtons.Left + fraMainButtons.Width - w___TYPE_NOT_FOUND.Width - lngButtonSpacer;
        //*TODO:** can't found type for with block
        //*With cmdButton1
        __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = cmdButton1;
        w___TYPE_NOT_FOUND.Visible = true;
        w___TYPE_NOT_FOUND.Caption = MCSTROK;
        w___TYPE_NOT_FOUND.Default = true;
        w___TYPE_NOT_FOUND.Left = lngButtonSpacer;

      //' ebtnOKOnly
        break;

      default:
        //*TODO:** can't found type for with block
        //*With cmdButton2
        __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = cmdButton2;
        w___TYPE_NOT_FOUND.Visible = false;
        w___TYPE_NOT_FOUND.Default = false;
        w___TYPE_NOT_FOUND.Caption = MCSTROK;
        w___TYPE_NOT_FOUND.Left = fraMainButtons.Left + fraMainButtons.Width - w___TYPE_NOT_FOUND.Width - lngButtonSpacer;
        //*TODO:** can't found type for with block
        //*With cmdButton1
        __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = cmdButton1;
        w___TYPE_NOT_FOUND.Visible = true;
        w___TYPE_NOT_FOUND.Default = true;
        w___TYPE_NOT_FOUND.Caption = MCSTROK;
        w___TYPE_NOT_FOUND.Left = (fraMainButtons.Left - w___TYPE_NOT_FOUND.Width) / 2;
        break;
    }
    //fnWindowUnlock
  }



  ///////////////////////////////////////////////////////////////////////////////////
  private void fnSetupForDefaultDisplay() {
    //--------------------------------------------------------------------
    // Procedure : fnSetupForDefaultDisplay
    // Comments  : Set default form and control properties
    //
    // Called by : cmdDetailCollapse_Click( )
    //             Form_Load( )
    // Parameters: N/A
    // Modified  :
    //--------------------------------------------------------------------
    fnWindowLock(Me.cbrfBrowseFolder.setHWnd());
    Me.Height = MCLNGMINFORMHEIGHT_COLLAPSED;
    Me.Width = MCLNGMINFORMWIDTH;

    fraDetails.Visible = true;
    cmdDetailsExpand.Visible = true;
    cmdDetailsExpand.Enabled = true;
    cmdDetailsCollapse.Visible = false;
    cmdDetailsCollapse.Enabled = false;
    fnWindowUnlock;
  }



  ///////////////////////////////////////////////////////////////////////////////////
  private void form_Load() {
    //--------------------------------------------------------------------
    // Procedure : Form_Load
    // Comments  : Loads the form
    //
    // Called by : N/A
    // Parameters: N/A
    // Modified  :
    //--------------------------------------------------------------------
    Me.Icon = LoadResPicture(modResConstants.gCRES_ICON_MAINAPP, vbResIcon);

    // Make sure Cursor reverts back to normal, in case it was left in an hourglass
    Screen.MousePointer = vbDefault;

    fnSetupForDefaultDisplay();
    // Don't center on MDI, since MDI may not have been loaded yet if an
    // error was generated during app startup.
  }



  ///////////////////////////////////////////////////////////////////////////////////
  private void form_Unload(int cancel) {
    //--------------------------------------------------------------------
    // Procedure : Form_Unload
    // Comments  : Unloads the form
    //
    // Called by : N/A
    // Parameters: N/A
    // Modified  :
    //--------------------------------------------------------------------
    Unload(this);
    modGeneral.fnFreeObject(frmMsgBox);
  }
}

public class eMsgButtons {
    public static final int EBTNYESNO = 0;
    public static final int EBTNOKCANCEL = 1;
    public static final int EBTNOKONLY = 2;
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


case class RmmsgboxData(
              id: Option[Int],

              )

object Rmmsgboxs extends Controller with ProvidesUser {

  val rmmsgboxForm = Form(
    mapping(
      "id" -> optional(number),

  )(RmmsgboxData.apply)(RmmsgboxData.unapply))

  implicit val rmmsgboxWrites = new Writes[Rmmsgbox] {
    def writes(rmmsgbox: Rmmsgbox) = Json.obj(
      "id" -> Json.toJson(rmmsgbox.id),
      C.ID -> Json.toJson(rmmsgbox.id),

    )
  }

  def get(id: Int) = GetAction { implicit request =>
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.LIST_RMMSGBOX), { user =>
      Ok(Json.toJson(Rmmsgbox.get(user, id)))
    })
  }

  def update(id: Int) = PostAction { implicit request =>
    Logger.debug("in rmmsgboxs.update")
    rmmsgboxForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      rmmsgbox => {
        Logger.debug(s"form: ${rmmsgbox.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.EDIT_RMMSGBOX), { user =>
          Ok(
            Json.toJson(
              Rmmsgbox.update(user,
                Rmmsgbox(
                       id,

                ))))
        })
      }
    )
  }

  def create = PostAction { implicit request =>
    Logger.debug("in rmmsgboxs.create")
    rmmsgboxForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      rmmsgbox => {
        Logger.debug(s"form: ${rmmsgbox.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.NEW_RMMSGBOX), { user =>
          Ok(
            Json.toJson(
              Rmmsgbox.create(user,
                Rmmsgbox(

                ))))
        })
      }
    )
  }

  def delete(id: Int) = PostAction { implicit request =>
    Logger.debug("in rmmsgboxs.delete")
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.DELETE_RMMSGBOX), { user =>
      Rmmsgbox.delete(user, id)
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

case class Rmmsgbox(
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

object Rmmsgbox {

  lazy val emptyRmmsgbox = Rmmsgbox(
)

  def apply(
      id: Int,
) = {

    new Rmmsgbox(
      id,
)
  }

  def apply(
) = {

    new Rmmsgbox(
)
  }

  private val rmmsgboxParser: RowParser[Rmmsgbox] = {
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
        Rmmsgbox(
              id,
,
              createdAt,
              updatedAt,
              updatedBy)
    }
  }

  def create(user: CompanyUser, rmmsgbox: Rmmsgbox): Rmmsgbox = {
    save(user, rmmsgbox, true)
  }

  def update(user: CompanyUser, rmmsgbox: Rmmsgbox): Rmmsgbox = {
    save(user, rmmsgbox, false)
  }

  private def save(user: CompanyUser, rmmsgbox: Rmmsgbox, isNew: Boolean): Rmmsgbox = {
    def getFields = {
      List(

      )
    }
    def throwException = {
      throw new RuntimeException(s"Error when saving ${C.RMMSGBOX}")
    }

    DBHelper.saveEx(
      user,
      Register(
        C.RMMSGBOX,
        C.ID,
        rmmsgbox.id,
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

  def load(user: CompanyUser, id: Int): Option[Rmmsgbox] = {
    loadWhere(user, s"${C.ID} = {id}", 'id -> id)
  }

  def loadWhere(user: CompanyUser, where: String, args : scala.Tuple2[scala.Any, anorm.ParameterValue[_]]*) = {
    DB.withConnection(user.database.database) { implicit connection =>
      SQL(s"SELECT t1.*, t2.${C.FK_NAME} FROM ${C.RMMSGBOX} t1 INNER JOIN ${C.???} t2 ON t1.${C.FK_ID} = t2.${C.FK_ID} WHERE $where")
        .on(args: _*)
        .as(rmmsgboxParser.singleOpt)
    }
  }

  def delete(user: CompanyUser, id: Int) = {
    DB.withConnection(user.database.database) { implicit connection =>
      try {
        SQL(s"DELETE FROM ${C.RMMSGBOX} WHERE ${C.ID} = {id}")
        .on('id -> id)
        .executeUpdate
      } catch {
        case NonFatal(e) => {
          Logger.error(s"can't delete a ${C.RMMSGBOX}. ${C.ID} id: $id. Error ${e.toString}")
          throw e
        }
      }
    }
  }

  def get(user: CompanyUser, id: Int): Rmmsgbox = {
    load(user, id) match {
      case Some(p) => p
      case None => emptyRmmsgbox
    }
  }
}


// Router

GET     /api/v1/general/rmmsgbox/:id              controllers.logged.modules.general.Rmmsgboxs.get(id: Int)
POST    /api/v1/general/rmmsgbox                  controllers.logged.modules.general.Rmmsgboxs.create
PUT     /api/v1/general/rmmsgbox/:id              controllers.logged.modules.general.Rmmsgboxs.update(id: Int)
DELETE  /api/v1/general/rmmsgbox/:id              controllers.logged.modules.general.Rmmsgboxs.delete(id: Int)




/**/
