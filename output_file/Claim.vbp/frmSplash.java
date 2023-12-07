public class frmSplash {

  //******************************************************************************
  // Module     : frmSplash
  // Description:
  // Procedures :
  //   Private    cmdOK_Click()
  //   Private    fnResetColors()
  //   Public     fnShowAsAboutBox()
  //   Public     fnShowAsSplashScreen()
  //   Private    Form_KeyPress(ByRef pintKeyAscii As Integer)
  //   Private    Form_Load()
  //   Private    Form_Unload(ByRef pintCancel As Integer)
  //   Private    tmrTimer_Timer()

  //
  // Modified   :
  // 03/03/02 BAW (Phase2A) Added support for new global error handler
  // 08/31/01 BAW (Phase2A) Added standardized error handlers
  // 09/25/00 JG  (Phase2A) Cleaned with Total Visual CodeTools 2000
  //
  // --------------------------------------------------
//Option Explicit
  *Option Compare Binary
  private String mstrScreenName = "";

*TODO: API Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
  private static final int MCSWP_NOMOVE = 0x2;
  private static final int MCSWP_NOREDRAW = 0x8;
  private static final int MCSWP_NOSIZE = 0x1;
  private static final int MCHWND_TOPMOST = -1;
  private static final int MCHWND_NOTOPMOST = -2;

  //Define the colors for the statusbar
  private static final int MCSLFBLUE = 4865792;
  private static final int MCSLFORANGE = 1026784;
  private static final int MCSLFWHITE = 16777215;
  private static final int MCSLFBLACK = 0;

  private float msng_EstProgress = 0;



  //////////////////////////////////////////////////////////////////////////////////////////////////
  private void cmdOK_Click() {
    // Comments  : When shown as a Help | About screen, the OK button is visible
    //             and unloads this screen when clicked.
    // Parameters:  -
    // Modified  :
    //
    // --------------------------------------------------
    try {
      "cmdOK_Click"
.equals(Const cstrCurrentProc As String);

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

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



//////////////////////////////////////////////////////////////////////////////////////////////////
  private void fnResetColors() {
    // Comments  : Adjusts the colors of the progress meter per the
    //             user's current screen resolution
    // Parameters:  -
    // Modified  :
    //
    // --------------------------------------------------
    try {
      "fnResetColors"
.equals(Const cstrCurrentProc As String);

      BITMAP bm = null;

      AutoRedraw = true;
      GetObject(Image, bm.length(), bm);

      if (bm.bmBitsPixel == 8) {
        cpbProgressBar.Color1 = MCSLFWHITE;
        cpbProgressBar.Color2 = MCSLFWHITE;
        // Following 3 lines added by BAW to ensure text is visible
        // on top of mcSLFBlue background
        lblApplicationName.ForeColor = MCSLFWHITE;
        lblApplicationDeveloper.ForeColor = MCSLFWHITE;
        lblVersion.ForeColor = MCSLFWHITE;
      } 
      else {
        cpbProgressBar.Color1 = MCSLFWHITE;
        cpbProgressBar.Color2 = MCSLFBLUE;
        // Following 3 lines added by BAW to ensure text is visible
        // on top of mcSLFBlue background
        lblApplicationName.ForeColor = MCSLFWHITE;
        lblApplicationDeveloper.ForeColor = MCSLFWHITE;
        lblVersion.ForeColor = MCSLFWHITE;
      }

      cpbProgressBar.BackColor = MCSLFBLACK;
      txtProgressBorder.BackColor = MCSLFBLACK;
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
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



//////////////////////////////////////////////////////////////////////////////////////////////////
  public void fnShowAsAboutBox() {
    // Comments  : This function is called from the Help | About menu choice
    //             and displays this form with an OK button.
    // Parameters:  -
    // Modified  :
    //
    // --------------------------------------------------
    try {
      "fnShowAsAboutBox"
.equals(Const cstrCurrentProc As String);

      // The following line triggers the Form_Load event if the form has not yet
      // been loaded
      lblProgress.Visible = false;
      cpbProgressBar.Visible = false;
      txtProgressBorder.Visible = false;

      cmdOK.Visible = true;
      Show(vbModal);
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
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



//////////////////////////////////////////////////////////////////////////////////////////////////
  public void fnShowAsSplashScreen() {
    // Comments  : Called by sub Main, this function displays this
    //             form as a splash screen
    // Parameters:  -
    // Modified  :
    //
    // --------------------------------------------------
    try {
      "fnShowAsSplashScreen"
.equals(Const cstrCurrentProc As String);

      cpbProgressBar.Width = 6195;
      cpbProgressBar.chrgHourglass.setValue(0);
      msng_EstProgress = 0;
      tmrTimer.Enabled = true;
      // CMP set the interval to be the number of seconds to display the splashscreen * 100
      // 10000 milliseconds = 10 seconds
      //' was 150
      tmrTimer.Interval = 55;
      Me.Show(vbModeless);

      // Put this form "on top"
      SetWindowPos(Me.cbrfBrowseFolder.setHWnd(), MCHWND_TOPMOST, 0, 0, 0, 0, MCSWP_NOMOVE || MCSWP_NOSIZE);
      // Now let it float behind other windows (like our own app's message boxes!) if warranted.
      SetWindowPos(Me.cbrfBrowseFolder.setHWnd(), MCHWND_NOTOPMOST, 0, 0, 0, 0, MCSWP_NOMOVE || MCSWP_NOSIZE);
      // **TODO:** label found: PROC_EXIT:;
      // Disable the error handler so errors hit here won't be handled by PROC_ERR
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
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



//////////////////////////////////////////////////////////////////////////////////////////////////
  private void form_KeyPress(int pintKeyAscii) { // TODO: Use of ByRef founded Private Sub Form_KeyPress(ByRef pintKeyAscii As Integer)
    // Comments  : If the user presses any key, unload this form.
    // Parameters: pintKeyAscii -
    // Modified  :
    //
    // --------------------------------------------------
    try {
      "Form_KeyPress"
.equals(Const cstrCurrentProc As String);

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

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



//////////////////////////////////////////////////////////////////////////////////////////////////
  private void form_Load() {
    // Comments  : Displays the form with the current version information. The
    //             colors of the progress meter reflect the user's selected
    //             screen resolution.
    // Parameters:  -
    // Modified  :
    //
    // --------------------------------------------------
    try {
      "Form_Load"
.equals(Const cstrCurrentProc As String);

      mstrScreenName = Me.Caption;

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      Me.Icon = LoadResPicture(modResConstants.gCRES_ICON_MAINAPP, vbResIcon);

      // Don't restore previously displayed form size & position

      // Adjust the colors per the user's current screen resolution
      fnResetColors();

      lblApplicationName = App.ProductName;
      lblVersion.Caption = "Application Version "+ App.Major+ "."+ App.Minor+ "."+ App.Revision;

      cpbProgressBar.Width = 6195;
      cpbProgressBar.chrgHourglass.setValue(0);
      cpbProgressBar.Refresh;
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



//////////////////////////////////////////////////////////////////////////////////////////////////
  private void form_Unload(int pintCancel) { // TODO: Use of ByRef founded Private Sub Form_Unload(ByRef pintCancel As Integer)
    // Comments  :
    // Parameters: N/A
    // Modified  :
    //
    // --------------------------------------------------
    try {
      "Form_Unload"
.equals(Const cstrCurrentProc As String);

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      if (modGeneral.bDEBUGAPPTERMINATION) {
        Debug.Print("Entering "+ mstrScreenName+ modResConstants.gCSTRDOT+ cstrCurrentProc);
      }

      // Don't save current form size & position
      modGeneral.fnFreeObject(frmSplash);
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



//////////////////////////////////////////////////////////////////////////////////////////////////
  private void tmrTimer_Timer() {
    // Comments  : This event is triggered each time the Timer's interval
    //             has elapsed. When it detects the progress meter has
    //             gotten to a certain point (currently >= 98%), it
    //             unloads the splash screen.
    // Parameters:  -
    // Modified  :
    // --------------------------------------------------
    try {
      "tmrTimer_Timer"
.equals(Const cstrCurrentProc As String);

      // Set screen name in case errors are reported here or
      // in procedures called by this Event Handler
      modGeneral.gerhApp.setScreenName(mstrScreenName);

      msng_EstProgress = msng_EstProgress + 5;

      //cpbProgressBar.Width = 6195 * (msng_EstProgress / 100)
      cpbProgressBar.chrgHourglass.setValue(msng_EstProgress);
      cpbProgressBar.Refresh;
      if (cpbProgressBar.Visible) {
        cpbProgressBar.SetFocus;
      }
      DoEvents;

      if (msng_EstProgress >= 98) {
        Me.Hide;
        frmMDIMain.Show;
        // Display the Log On screen, which forces the user to either log on to the
        // application or, alternatively, EXIT the application via a call to fnTerminateTheApp().
        frmLogOn.Show(vbModal, frmMDIMain);

        // Uncomment out the next line if you want to test a table wrapper.
        //frmMDIMain.fnTestTableWrapper

        // Don't try showing the Insured form since it's hard to tell that the user successfully
        // logged on vs. clicked the Exit Application button (unless yet another global boolean
        // is set.)  Per Michelle Wilkosky, it's fine to just skip automatically displaying
        // the Insured screen. Hence, the following lines were commented out.
        //       ' If we get here, the user successfully logged on, so now display
        //       ' the Insured screen.
        //       frmMDIMain.fnShowInsuredForm
        Unload(this);
      }
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
      //' Invalid Procedure Call or Argument
      case  5     :
        // Don't know why it's occurring but it's harmless so ignore it.
        /**TODO:** resume found: Resume(Next)*/;
      //' Must close or hide topmost modal form first
        break;

      case  402   :
        // If we detect an invalid DB path upon starting up the app, the frmSetDatabaseLocation
        // form is opened modally so the user can set a valid path. However, we can't open
        // that form modally until the splash screen has been unloaded.
        Unload(this);
        /**TODO:** resume found: Resume(PROC_EXIT)*/;
        break;

      default:
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


case class RmsplashData(
              id: Option[Int],

              )

object Rmsplashs extends Controller with ProvidesUser {

  val rmsplashForm = Form(
    mapping(
      "id" -> optional(number),

  )(RmsplashData.apply)(RmsplashData.unapply))

  implicit val rmsplashWrites = new Writes[Rmsplash] {
    def writes(rmsplash: Rmsplash) = Json.obj(
      "id" -> Json.toJson(rmsplash.id),
      C.ID -> Json.toJson(rmsplash.id),

    )
  }

  def get(id: Int) = GetAction { implicit request =>
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.LIST_RMSPLASH), { user =>
      Ok(Json.toJson(Rmsplash.get(user, id)))
    })
  }

  def update(id: Int) = PostAction { implicit request =>
    Logger.debug("in rmsplashs.update")
    rmsplashForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      rmsplash => {
        Logger.debug(s"form: ${rmsplash.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.EDIT_RMSPLASH), { user =>
          Ok(
            Json.toJson(
              Rmsplash.update(user,
                Rmsplash(
                       id,

                ))))
        })
      }
    )
  }

  def create = PostAction { implicit request =>
    Logger.debug("in rmsplashs.create")
    rmsplashForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      rmsplash => {
        Logger.debug(s"form: ${rmsplash.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.NEW_RMSPLASH), { user =>
          Ok(
            Json.toJson(
              Rmsplash.create(user,
                Rmsplash(

                ))))
        })
      }
    )
  }

  def delete(id: Int) = PostAction { implicit request =>
    Logger.debug("in rmsplashs.delete")
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.DELETE_RMSPLASH), { user =>
      Rmsplash.delete(user, id)
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

case class Rmsplash(
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

object Rmsplash {

  lazy val emptyRmsplash = Rmsplash(
)

  def apply(
      id: Int,
) = {

    new Rmsplash(
      id,
)
  }

  def apply(
) = {

    new Rmsplash(
)
  }

  private val rmsplashParser: RowParser[Rmsplash] = {
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
        Rmsplash(
              id,
,
              createdAt,
              updatedAt,
              updatedBy)
    }
  }

  def create(user: CompanyUser, rmsplash: Rmsplash): Rmsplash = {
    save(user, rmsplash, true)
  }

  def update(user: CompanyUser, rmsplash: Rmsplash): Rmsplash = {
    save(user, rmsplash, false)
  }

  private def save(user: CompanyUser, rmsplash: Rmsplash, isNew: Boolean): Rmsplash = {
    def getFields = {
      List(

      )
    }
    def throwException = {
      throw new RuntimeException(s"Error when saving ${C.RMSPLASH}")
    }

    DBHelper.saveEx(
      user,
      Register(
        C.RMSPLASH,
        C.ID,
        rmsplash.id,
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

  def load(user: CompanyUser, id: Int): Option[Rmsplash] = {
    loadWhere(user, s"${C.ID} = {id}", 'id -> id)
  }

  def loadWhere(user: CompanyUser, where: String, args : scala.Tuple2[scala.Any, anorm.ParameterValue[_]]*) = {
    DB.withConnection(user.database.database) { implicit connection =>
      SQL(s"SELECT t1.*, t2.${C.FK_NAME} FROM ${C.RMSPLASH} t1 INNER JOIN ${C.???} t2 ON t1.${C.FK_ID} = t2.${C.FK_ID} WHERE $where")
        .on(args: _*)
        .as(rmsplashParser.singleOpt)
    }
  }

  def delete(user: CompanyUser, id: Int) = {
    DB.withConnection(user.database.database) { implicit connection =>
      try {
        SQL(s"DELETE FROM ${C.RMSPLASH} WHERE ${C.ID} = {id}")
        .on('id -> id)
        .executeUpdate
      } catch {
        case NonFatal(e) => {
          Logger.error(s"can't delete a ${C.RMSPLASH}. ${C.ID} id: $id. Error ${e.toString}")
          throw e
        }
      }
    }
  }

  def get(user: CompanyUser, id: Int): Rmsplash = {
    load(user, id) match {
      case Some(p) => p
      case None => emptyRmsplash
    }
  }
}


// Router

GET     /api/v1/general/rmsplash/:id              controllers.logged.modules.general.Rmsplashs.get(id: Int)
POST    /api/v1/general/rmsplash                  controllers.logged.modules.general.Rmsplashs.create
PUT     /api/v1/general/rmsplash/:id              controllers.logged.modules.general.Rmsplashs.update(id: Int)
DELETE  /api/v1/general/rmsplash/:id              controllers.logged.modules.general.Rmsplashs.delete(id: Int)




/**/
