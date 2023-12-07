public class cautAuthenticate {

  //!TODO! Add support for determining whether user has access to restricted areas of app
  //       (e.g., is a member of Support or UserAdmin roles). Make this a property that
  //       frmLogOn will query and use to set a global variable that other forms
  //       can reference.

  //--------------------------------------------------------------------------
  // Module     : cautAuthenticate
  // Description: Instantiated by frmLogon to determine if the user
  //              is authenticated to each possible environment's
  //              application SQL Server database
  //
  // Procedures :
  //    Public    AuthenticateAll(ByVal strUserID As String) As String()
  //    Private   fnIsAuthorized(ByVal strUserID, ByVal strEnv As String) As Boolean
  //
  // Revision History: 1.0 BAW 05/16/02 Initial Creation
  //--------------------------------------------------------------------------
//Option Explicit
  *Option Compare Binary

  *#Const DEBUG_ERH = False
  *#Const DEBUG_RST = False

  private static final String MCSTRNAME = "cautAuthenticate.";



  ///\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
  //\                                                                  /
  //|                        PUBLIC  Procedures                        |
  ///                                                                  \
  //\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/

  //////////////////////////////////////////////////////////////////////////////////////////////////
  public String() authenticateEnvironments(String strUserID) {
    String() _rtn = null;
    //--------------------------------------------------------------------------
    // Procedure:   AuthenticateEnvironments
    // Description: Will call the SP_USER stored procedure for each SQL server/database
    //              listed in the Environments section of the registry. It will return
    //              an array of Environment Names that the Log On screen can use
    //              to populate the Environments combo box (so the user is only
    //              presented with a list of Environments for which she/he is
    //              authorized.
    //
    // Called By:   cmdOK_Click of frmLogon
    //
    // Params:
    //    strUserID (in) - the User ID under which user is known by SQL Server
    //
    // Returns:     An array of strings representing the Environments for which
    //              the user is authorized.
    //-----------------------------------------------------------------------------
    "AuthenticateEnvironments"
.equals(Const cstrCurrentProc As String);
    String[] astrAllEnvironments() = null;
    String[] astrAuthorizedEnvironments() = null;
    int intAllIndex = 0;
    int intAuthIndex = 0;

    try {

      // An error is generated during app startup if there are no
      // environment names defined, hence we should always have at least one.
      astrAllEnvironments = modGeneral.gapsApp.getEnvironmentNames();

      // Resize our authorized Environments array based on the max number
      // of possibly authorized environments. Shouldn't be much waste.
      G.redim(astrAuthorizedEnvironments, astrAllEnvironments.length + 1);

      for (intAllIndex = LBound(astrAllEnvironments); intAllIndex <= astrAllEnvironments.length; intAllIndex++) {
        //SQL_INTEGRATED_SECURITY If fnIsAuthorized(strUserID, astrAllEnvironments(intAllIndex)) Then
        astrAuthorizedEnvironments[intAuthIndex] = astrAllEnvironments[intAllIndex];
        intAuthIndex = intAuthIndex + 1;
        //SQL_INTEGRATED_SECURITY End If
      }

      _rtn = astrAuthorizedEnvironments;
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
  public boolean authenticateUser(String strEnvironIn, String strUserIDIn, String strPasswordIn, cconConnection pconIn, Object bActiveDBIn) { // TODO: Use of ByRef founded Public Function AuthenticateUser(ByVal strEnvironIn As String, ByVal strUserIDIn As String, ByVal strPasswordIn As String, ByRef pconIn As cconConnection, Optional ByVal bActiveDBIn = True) As Boolean
    //--------------------------------------------------------------------------
    // Procedure:   AuthenticateUser
    // Description: Logs the user on to the specified Environment, using the
    //              specified Connection object. From the frmLogon screen,
    //              this will be the global connection object (gconAppActive).
    //              From other screens, such as those that need to verify a
    //              record can be deleted without impacting ArchiveDB records,
    //              it is done with a local Connection object that is intended
    //              to point to the Archive DB.
    // Params:      strUserID   (input)        - the UserID of the user
    //              strPassword (input)        - the Password of the user
    //              pconIn      (input/output) - a pointer to the cconConnection object
    //                                           to use for the logging on
    //              strEnviron  (input)        - the name of the Environment
    //                                           to log on. (This may come from the
    //                                           Environments combobox on the Log On
    //                                           screen or it may be the .LoggedOnEnviron
    //                                           property of the cconConnection object)
    //              bArchive    (input)        - True to log onto the Archive DB; False
    //                                           to log onto the Active DB
    //
    // Called By:   cmdOK_Click of frmLogon, to log on to the Active DB
    //              Other screens, when validations or processing must be
    //                     done against the Archive DB
    //
    // Returns:     True if the logon was successful; False otherwise
    //-----------------------------------------------------------------------------
    "AuthenticateUser"
.equals(Const cstrCurrentProc As String);
    "AppRoleClaims"
.equals(Const cstrAppRoleClaims_UserId As String);
    "claims"
.equals(Const cstrAppRoleClaims_Password As String);

    try {

      pconIn.connect(strEnviron:=strEnvironIn, strUserID:=strUserIDIn, strPassword:=strPasswordIn, bActiveDB:=bActiveDBIn);

      // NOTE: All sprocs must be DBO-owned if the application role is put into effect. This means
      //       that the following line may need to be commented out at times while application
      //       development is going on.
      pconIn.setAppRole(strRoleName:=cstrAppRoleClaims_UserId, strRolePassword:=cstrAppRoleClaims_Password);
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



///\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
//\                                                                  /
//|                       PRIVATE  Procedures                        |
///                                                                  \
//\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/


//////////////////////////////////////////////////////////////////////////////////////////////////
  private boolean fnIsAuthorized(String strUserID, String strEnvironment) {
    boolean _rtn = false;
    //--------------------------------------------------------------------------
    // Procedure:   fnIsAuthorized
    // Description: This method uses the "sp_helpuser" system stored
    //              procedure to determine whether the specified
    //              UserID is an authorized user of the database
    //              referenced by the specified Environment.
    //
    // Params:
    //    strUserID      (in) - the User ID under which user is known by SQL Server
    //    strEnvironment (in) - the Environment Name (as established by
    //                          the gapsApp AppSettings object.
    //
    // Returns:     True if the user is authorized; false otherwise.
    //-----------------------------------------------------------------------------
    "fnIsAuthorized"
.equals(Const cstrCurrentProc As String);

    // NOTE: The following Dummy App ID/Password must be set up on every server
    //       on which the application will run
    //' "claimapp"
    "CLAIMAPP"
.equals(Const cstrDummyApp_UserId As String);
    //' "claimapp"
    "CLAIMAPP"
.equals(Const cstrDummyApp_Password As String);

    //' # of input or output params sproc expects
    Const(clngSprocParamCount As Long == 2);
    //' Stored procedure to execute
    "sp_helpuser"
.equals(Const cstrSproc As String);
    ADODB.Parameter prmReturnValue = null;
    ADODB.Parameter prmName_in_DB = null;
    DBRecordSet rstTemp = null;
    New adwTemp = null; cadwADOWrapper
    New conTemp = null; cconConnection

    try {

      // Connect to the specified environment using the Dummy App ID,
      // then execute the sp_helpuser sproc.
      if (conTemp.Connect(strEnvironment, cstrDummyApp_UserId, cstrDummyApp_Password)) {

        if (!(adwTemp.CommandSetSproc(cstrSproc, conTemp))) {
          // **TODO:** goto found: GoTo PROC_EXIT;
        }

        //*TODO:** can't found type for with block
        //*With adwTemp.ADOCommand
        __TYPE_NOT_FOUND w___TYPE_NOT_FOUND = adwTemp.ADOCommand;
        // ---Parameter #1---
        // Define the return value that represents the error code (i.e. reason) why
        // the stored procedure failed.
        prmReturnValue = w___TYPE_NOT_FOUND.CreateParameter(Name:="@return_value", Type:=adInteger, Direction:=adParamReturnValue, .value:=Null);
        w___TYPE_NOT_FOUND.Parameters.Append(prmReturnValue);

        // ---Parameter #2---
        // Define the name_in_db input parameter, which represents the User ID being checked
        prmName_in_DB = w___TYPE_NOT_FOUND.CreateParameter(Name:="@name_in_db", Type:=adVarChar, Direction:=adParamInput, Size:=255, .value:=fnNullIfZLS(varIn:=strUserID, bHandleEmbeddedQuotes:=True));
        w___TYPE_NOT_FOUND.Parameters.Append(prmName_in_DB);

        rstTemp = w___TYPE_NOT_FOUND.Execute();

        if (rstTemp.RecordCount > 0) {
          _rtn = true;
        }
      }
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    modGeneral.fnFreeObject(prmReturnValue);
    modGeneral.fnFreeObject(prmName_in_DB);
    modGeneral.fnFreeRecordset(rstTemp);
    if (modGeneral.fnIsObject(conTemp)) {
      conTemp.Disconnect;
      modGeneral.fnFreeObject(conTemp);
    }
    modGeneral.fnFreeObject(adwTemp);

    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
      //' 4013
      case  vbObjectError + modResConstants.gCRES_NERR_LOGON_FAILURE:
        // This environment will be considered "not authorized"
        VBA.ex.Clear;
        modGeneral.gerhApp.clear();
        // **TODO:** goto found: GoTo PROC_EXIT;
        //Resume Next
      //' The name supplied (xxx) is not a user, role or aliased login
        break;

      case  -2147217900:
        // This environment will be considered "not authorized"
        VBA.ex.Clear;
        modGeneral.gerhApp.clear();
        // **TODO:** goto found: GoTo PROC_EXIT;
        break;

      default:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
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


case class AutauthenticateData(
              id: Option[Int],

              )

object Autauthenticates extends Controller with ProvidesUser {

  val autauthenticateForm = Form(
    mapping(
      "id" -> optional(number),

  )(AutauthenticateData.apply)(AutauthenticateData.unapply))

  implicit val autauthenticateWrites = new Writes[Autauthenticate] {
    def writes(autauthenticate: Autauthenticate) = Json.obj(
      "id" -> Json.toJson(autauthenticate.id),
      C.ID -> Json.toJson(autauthenticate.id),

    )
  }

  def get(id: Int) = GetAction { implicit request =>
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.LIST_AUTAUTHENTICATE), { user =>
      Ok(Json.toJson(Autauthenticate.get(user, id)))
    })
  }

  def update(id: Int) = PostAction { implicit request =>
    Logger.debug("in autauthenticates.update")
    autauthenticateForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      autauthenticate => {
        Logger.debug(s"form: ${autauthenticate.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.EDIT_AUTAUTHENTICATE), { user =>
          Ok(
            Json.toJson(
              Autauthenticate.update(user,
                Autauthenticate(
                       id,

                ))))
        })
      }
    )
  }

  def create = PostAction { implicit request =>
    Logger.debug("in autauthenticates.create")
    autauthenticateForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      autauthenticate => {
        Logger.debug(s"form: ${autauthenticate.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.NEW_AUTAUTHENTICATE), { user =>
          Ok(
            Json.toJson(
              Autauthenticate.create(user,
                Autauthenticate(

                ))))
        })
      }
    )
  }

  def delete(id: Int) = PostAction { implicit request =>
    Logger.debug("in autauthenticates.delete")
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.DELETE_AUTAUTHENTICATE), { user =>
      Autauthenticate.delete(user, id)
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

case class Autauthenticate(
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

object Autauthenticate {

  lazy val emptyAutauthenticate = Autauthenticate(
)

  def apply(
      id: Int,
) = {

    new Autauthenticate(
      id,
)
  }

  def apply(
) = {

    new Autauthenticate(
)
  }

  private val autauthenticateParser: RowParser[Autauthenticate] = {
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
        Autauthenticate(
              id,
,
              createdAt,
              updatedAt,
              updatedBy)
    }
  }

  def create(user: CompanyUser, autauthenticate: Autauthenticate): Autauthenticate = {
    save(user, autauthenticate, true)
  }

  def update(user: CompanyUser, autauthenticate: Autauthenticate): Autauthenticate = {
    save(user, autauthenticate, false)
  }

  private def save(user: CompanyUser, autauthenticate: Autauthenticate, isNew: Boolean): Autauthenticate = {
    def getFields = {
      List(

      )
    }
    def throwException = {
      throw new RuntimeException(s"Error when saving ${C.AUTAUTHENTICATE}")
    }

    DBHelper.saveEx(
      user,
      Register(
        C.AUTAUTHENTICATE,
        C.ID,
        autauthenticate.id,
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

  def load(user: CompanyUser, id: Int): Option[Autauthenticate] = {
    loadWhere(user, s"${C.ID} = {id}", 'id -> id)
  }

  def loadWhere(user: CompanyUser, where: String, args : scala.Tuple2[scala.Any, anorm.ParameterValue[_]]*) = {
    DB.withConnection(user.database.database) { implicit connection =>
      SQL(s"SELECT t1.*, t2.${C.FK_NAME} FROM ${C.AUTAUTHENTICATE} t1 INNER JOIN ${C.???} t2 ON t1.${C.FK_ID} = t2.${C.FK_ID} WHERE $where")
        .on(args: _*)
        .as(autauthenticateParser.singleOpt)
    }
  }

  def delete(user: CompanyUser, id: Int) = {
    DB.withConnection(user.database.database) { implicit connection =>
      try {
        SQL(s"DELETE FROM ${C.AUTAUTHENTICATE} WHERE ${C.ID} = {id}")
        .on('id -> id)
        .executeUpdate
      } catch {
        case NonFatal(e) => {
          Logger.error(s"can't delete a ${C.AUTAUTHENTICATE}. ${C.ID} id: $id. Error ${e.toString}")
          throw e
        }
      }
    }
  }

  def get(user: CompanyUser, id: Int): Autauthenticate = {
    load(user, id) match {
      case Some(p) => p
      case None => emptyAutauthenticate
    }
  }
}


// Router

GET     /api/v1/general/autauthenticate/:id              controllers.logged.modules.general.Autauthenticates.get(id: Int)
POST    /api/v1/general/autauthenticate                  controllers.logged.modules.general.Autauthenticates.create
PUT     /api/v1/general/autauthenticate/:id              controllers.logged.modules.general.Autauthenticates.update(id: Int)
DELETE  /api/v1/general/autauthenticate/:id              controllers.logged.modules.general.Autauthenticates.delete(id: Int)




/**/
