public class modResConstants {

//Option Explicit
  *Option Compare Binary

  // The following constant is used to determine whether the cerhErrorHandler class (gerhApp)
  // is empty (i.e. set to its initialized state) or not.
  public static final Long GCLNGERR_NUM_DEFAULT = 999;

  // The following is used to string together a Screen Name with a Proc Name to form
  // the context when a form encounters an error directly (i.e. non-raised)
  public static final String GCSTRDOT = ".";

  //------------------------------------------------------------------------
  //    Public Constants re: Icons, Bitmaps and others items in CLAIM.RES
  //------------------------------------------------------------------------
  public static final String GCRES_ICON_MAINAPP = "_MAINAPP";
  public static final Long GCRES_ICON_INFO = 101;
  public static final Long GCRES_ICON_WARN = 103;
  public static final Long GCRES_ICON_ALRT = 102;
  public static final Long GCRES_ICON_ERR = 104;
  public static final Long GCRES_ICON_BINOCULARS = 105;


  //------------------------------------------------------------------------
  //       Public Constants re: Warning/Info/Alert/Error messages
  //
  //      The following ranges MUST be used and MUST correspond to IDs
  //                     in the CLAIM.RES resource file.
  //------------------------------------------------------------------------
  //' Lower Bounds for CLAIM.RES
  public static final Integer GCRES_LOWEST_APP_ERROR = 1000;
  //' Upper Bounds for CLAIM.RES
  public static final Integer GCRES_HIGHEST_APP_ERROR = 9999;

  // -=-= Informational Messages =-=-
  //' <Informational messages (1000-1999) start here>
  public static final Integer GCRES_INFO_START = 1000;
  //' Unable to open <@@1> in a browser window. Please ensure that the Internet Explorer and the Adobe Acrobat reader software are installed.
  public static final Integer GCRES_INFO_CANT_LAUNCH_URL = 1001;
  //' Another user (@@1) updated this record since you displayed it. Your changes have been discarded.
  public static final Integer GCRES_INFO_ANOTHER_USER_UPDATED_DISCARDED = 1002;
  //' Another user deleted this record since you displayed it. Your changes have not been saved.
  public static final Integer GCRES_INFO_ANOTHER_USER_DELETED_NOT_SAVED = 1003;
  //' @@1 record(s) were written to the @@2 tax file. The total interest (Box 1) amount was @@3 and the
  public static final Integer GCRES_INFO_TAX_FILE_GEND = 1004;
  // total Interest Withheld (Box 4) amount was @@4.@@CRLF
  // @@5 record(s) were written to the @@6 tax file. The total interest (Box 1) amount was @@7 and the
  // total Interest Withheld (Box 4) amount was @@8.
  //' Another user deleted this record since you displayed it.
  public static final Integer GCRES_INFO_ANOTHER_USER_DELETED = 1005;
  //' The @@1 table is empty.
  public static final Integer GCRES_INFO_TABLE_IS_EMPTY = 1006;
  //' The specified @@1 is not authorized for any of the application's @@2s.
  public static final Integer GCRES_INFO_NO_AUTHENTICATED_ENVIRONMENTS = 1007;
  //' The @@1 has changed. This change may affect the calculations for existing Payees. Please review and, if necessary, recalculate each Payee.
  public static final Integer GCRES_INFO_DT_CHG_MAY_AFFECT_PAYEES = 1008;
  //' Your input was truncated to @@1 character(s).
  public static final Integer GCRES_INFO_INPUT_WAS_TRUNCATED = 1009;
  //' Unable to open @@1. The file either does not exist or no application is associated with files of type @@2.
  public static final Integer GCRES_INFO_CANT_OPEN_FILE = 1014;
  //' <Informational messages (1000-1999) end here>
  public static final Integer GCRES_INFO_END = 1999;
  // -=-= Warnings =-=-
  //' <Warning messages (2000-2999) start here>
  public static final Integer GCRES_WARN_START = 2000;
  //' The drop-down list for @@1 is empty. Since you will be unable to make a selection in this field, this screen may behave unpredictibly.
  public static final Integer GCRES_WARN_CBO_IS_EMPTY = 2001;
  //' The list for @@1 is empty. Since you will be unable to make a selection in this field, this screen may behave unpredictibly.
  public static final Integer GCRES_WARN_LST_IS_EMPTY = 2002;
  //' There is no current Insured record. The Payee screen cannot be opened.
  public static final Integer GCRES_WARN_NO_CURR_INSURED = 2003;
  //' This claims requires a certified @@1 to avoid withholding. Make sure you don't pay interest until it has been received.
  public static final Integer GCRES_WARN_GET_TIN_BEFORE_PAYING_INT = 2004;
  //' The @@1 exceeds @@2. Please verify this amount is correct.
  public static final Integer GCRES_WARN_DTHB_PMT_AMT_MAY_BE_TOO_HIGH = 2005;
  //' The Rate supplied or obtained from the STATE_RULE_T table is non-numeric (@@1). Please try again.
  public static final Integer GCRES_WARN_NONNUMERIC_RATE = 2006;
  //' The Rate supplied or derived from the supplied Rate is a negative number (@@1). Please try again.
  public static final Integer GCRES_WARN_RATE_IS_NEGATIVE = 2007;
  //' The Rate supplied cannot have more than @@1 decimal positions specified. Please try again.
  public static final Integer GCRES_WARN_TOO_MANY_DECIMALS = 2008;
  //' An error was encountered while trying to @@1. This may be due to network unavailability or insufficient authorizations. @@2
  public static final Integer GCRES_WARN_POSSIBLE_FILESYS_PERM_PROBLEM = 2010;

  // Case use 2003, 2004, 2005, 2006, 2007
  //' <Warning messages (2000-2999) end here>
  public static final Integer GCRES_WARN_END = 2999;
  // -=-= Alerts =-=-
  //' <Alert messages (3000-3999) start here>
  public static final Integer GCRES_ALRT_START = 3000;
  // Can use 3001
  //' Are you sure you want to delete this record?
  public static final Integer GCRES_ALRT_OK_TO_DELETE_RECORD = 3002;
  // Can use 3003
  //' You have changes pending.  Do you want to lose them?
  public static final Integer GCRES_ALRT_CHANGES_PENDING = 3004;
  // Can use 3005, 3006, 3008
  //?? Public Const gcRES_ALRT_OK_TO_DELETE_INFO As Integer = 3007                ' You are about to delete the information for @@1. Are you sure you want to delete this?
  //' <Alert messages (3000-3999) end here>
  public static final Integer GCRES_ALRT_END = 3999;
  // -=-= Non-Fatal (i.e. Process Fatal) Errors =-=-
  //' <Non-fatal Error messages (4000-4999) start here>
  public static final Integer GCRES_NERR_START = 4000;
  //' Unable to retrieve the value (@@1) of the requested registry key (@@2).
  public static final Integer GCRES_NERR_CANTOPEN_REGKEY = 4001;
  //' Unable to save the value (@@1) to the requested registry key (@@2).
  public static final Integer GCRES_NERR_CANTSAVE_REGKEY = 4002;
  //' The drive or path specified does not exist. Please be sure to specify an existing drive and directory.
  public static final Integer GCRES_NERR_DRIVE_OR_PATH_NOT_FOUND = 4003;
  //' No records were found with @@1.
  public static final Integer GCRES_NERR_NO_RECS_WERE_FOUND = 4004;
  //' The Rate supplied or derived from the supplied Rate is more than @@1%. This is only allowed when the @@2 is Maine. Please try again.
  public static final Integer GCRES_NERR_INTEREST_RATE_TOO_HIGH = 4005;
  //' Neither Group nor Individual rates were found for the state of @@1 as of @@2. The calculations cannot be done.
  public static final Integer GCRES_NERR_STATE_RATES_NOT_FOUND = 4006;
  //' Individual rates were not found for the state of @@1 as of @@2. The calculations cannot be done.
  public static final Integer GCRES_NERR_INDV_STATE_RATES_NOT_FOUND = 4007;
  //' One or more numeric fields are too large to be stored in the database. Your changes cannot be saved.
  public static final Integer GCRES_NERR_NUMERIC_FLD_TOO_LARGE = 4008;
  //' The calculation was halted since you clicked Cancel. Your changes have not been saved.
  public static final Integer GCRES_NERR_CALC_WAS_CANCELLED = 4010;
  //' The report definition file for the selected report (@@1) could not be found.
  public static final Integer GCRES_NERR_RPTFILE_NOT_FOUND = 4011;
  // MME START - WRUS 4999
  //' Invalid record found on table STATE_RULE_TIER_T (4012) - ' for the state of @@1 as of @@2. The calculations cannot be done.
  public static final Integer GCRES_NERR_INVALID_ENTRY_RULE_TIER_T = 4012;
  // MME END  - WRUS 4999
  // Can use 4012, 4014, 4017, 4018, 4019, 4020, 4021, 4022
  //' The logon was unsuccessful. Please verify the correct User ID and Password were specified correctly, with appropriate case, that you have permissions to the database and the server on which it is located, and that Microsoft Data Access Components (MDAC) is installed. (RC=@@1)
  public static final Integer GCRES_NERR_LOGON_FAILURE = 4013;
  //' A connection could not be established to the @@1 environment's database. (State=@@2)
  public static final Integer GCRES_NERR_CONNECTION_FAILURE = 4015;
  //' One or more registry entries that define how to connect to the selected Environment (@@1) are missing. Without all of these entries, the app cannot connect to the database.
  public static final Integer GCRES_NERR_ENV_REG_ENTRIES_MISSING = 4016;
  //?? Public Const gcRES_NERR_TABLE_IS_EMPTY As Integer = 4023                   ' The @@1 table is empty.
  // Can use 4024
  //?? Public Const gcRES_NERR_MEFS_EFF_DT_BEFORE_CASE_EFF_DT As Integer = 4025   ' The @@1 is prior to the @@2 (@@3).
  //' Programmer error:  An unexpected value was encountered in a SELECT CASE statement.
  public static final Integer GCRES_NERR_UNEXPECTED_VAL_SELECT_CASE = 4026;
  //' The specified record was not found in the database (@@1).
  public static final Integer GCRES_NERR_REC_NOT_FOUND = 4027;
  //' An error occurred while attempting to @@1 this record.
  public static final Integer GCRES_NERR_ERR_WHILE_TRYING_TO = 4028;
  //' This @@1 is associated with one or more records on the @@2 table and cannot be deleted until those records themselves are deleted.
  public static final Integer GCRES_NERR_DEPENDENT_RECS_EXIST = 4029;
  //?? Public Const gcRES_NERR_FUND_USED_AS_MKTVAL_FUND As Integer = 4030         ' Fund @@1 cannot be deleted because it is used as another fund's Market Value Fund Cd.
  //' A record with the specified key (@@1) already exists. Please specify a unique key.
  public static final Integer GCRES_NERR_ADD_WITH_NONUNIQUE_KEY = 4031;
  //' The @@1 specified (@@2) is no longer defined in the @@3 table. Please choose a different value.
  public static final Integer GCRES_NERR_KEY_NOT_FOUND = 4032;
  // Can use 4033, 4035, 4039, 4043, 4044
  //' Cross-field validation errors were found. These must be corrected before @@1:@@2
  public static final Integer GCRES_NERR_CROSS_FLD_VALIDATIONS_FAILED = 4034;
  //?? Public Const gcRES_NERR_THIS_FUNCTIONALITY_NOT_AVAIL As Integer = 4036     ' This functionality isn't available yet.
  //' The @@1 is invalid. @@2
  public static final Integer GCRES_NERR_INVALID_DATA = 4037;
  //?? Public Const gcRES_NERR_NO_DATA_WAS_FOUND As Integer = 4038                ' No data was found @@1.
  //' The following required fields must be supplied before your request can be processed:@@CRLF@@1
  public static final Integer GCRES_NERR_REQD_FIELDS_MISSING = 4041;
  //?? Public Const gcRES_NERR_All_MUST_BE_INPUT As Integer = 4042                ' If any of the following fields are input, then all must be:@@1
  //' <Non-fatal Error messages (4000-4999) end here>
  public static final Integer GCRES_NERR_END = 4999;

  // -=-= Fatal (i.e. App Fatal) Errors =-=-
  //' <Fatal Error messages (9000-9999) start here>
  public static final Integer GCRES_FERR_START = 9000;
  //' No Environments have been defined in the registry. Without these entries (built by the install), the app will be unable to connect to the database.
  public static final Integer GCRES_FERR_NO_ENVS = 9001;
  //' The stored procedure "@@1" was not found.
  public static final Integer GCRES_FERR_SPROC_NOT_FOUND = 9002;
  //' The database object referenced in a SQL statement (@@1) was not found or you have insufficient permissions to access it.
  public static final Integer GCRES_FERR_SQL_STMT_OBJECT_NOT_FOUND = 9003;
  //' <Fatal Error messages (9000-9999) end here>
  public static final Integer GCRES_FERR_END = 9999;

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


case class OdresconstantsData(
              id: Option[Int],

              )

object Odresconstantss extends Controller with ProvidesUser {

  val odresconstantsForm = Form(
    mapping(
      "id" -> optional(number),

  )(OdresconstantsData.apply)(OdresconstantsData.unapply))

  implicit val odresconstantsWrites = new Writes[Odresconstants] {
    def writes(odresconstants: Odresconstants) = Json.obj(
      "id" -> Json.toJson(odresconstants.id),
      C.ID -> Json.toJson(odresconstants.id),

    )
  }

  def get(id: Int) = GetAction { implicit request =>
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.LIST_ODRESCONSTANTS), { user =>
      Ok(Json.toJson(Odresconstants.get(user, id)))
    })
  }

  def update(id: Int) = PostAction { implicit request =>
    Logger.debug("in odresconstantss.update")
    odresconstantsForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      odresconstants => {
        Logger.debug(s"form: ${odresconstants.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.EDIT_ODRESCONSTANTS), { user =>
          Ok(
            Json.toJson(
              Odresconstants.update(user,
                Odresconstants(
                       id,

                ))))
        })
      }
    )
  }

  def create = PostAction { implicit request =>
    Logger.debug("in odresconstantss.create")
    odresconstantsForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      odresconstants => {
        Logger.debug(s"form: ${odresconstants.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.NEW_ODRESCONSTANTS), { user =>
          Ok(
            Json.toJson(
              Odresconstants.create(user,
                Odresconstants(

                ))))
        })
      }
    )
  }

  def delete(id: Int) = PostAction { implicit request =>
    Logger.debug("in odresconstantss.delete")
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.DELETE_ODRESCONSTANTS), { user =>
      Odresconstants.delete(user, id)
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

case class Odresconstants(
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

object Odresconstants {

  lazy val emptyOdresconstants = Odresconstants(
)

  def apply(
      id: Int,
) = {

    new Odresconstants(
      id,
)
  }

  def apply(
) = {

    new Odresconstants(
)
  }

  private val odresconstantsParser: RowParser[Odresconstants] = {
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
        Odresconstants(
              id,
,
              createdAt,
              updatedAt,
              updatedBy)
    }
  }

  def create(user: CompanyUser, odresconstants: Odresconstants): Odresconstants = {
    save(user, odresconstants, true)
  }

  def update(user: CompanyUser, odresconstants: Odresconstants): Odresconstants = {
    save(user, odresconstants, false)
  }

  private def save(user: CompanyUser, odresconstants: Odresconstants, isNew: Boolean): Odresconstants = {
    def getFields = {
      List(

      )
    }
    def throwException = {
      throw new RuntimeException(s"Error when saving ${C.ODRESCONSTANTS}")
    }

    DBHelper.saveEx(
      user,
      Register(
        C.ODRESCONSTANTS,
        C.ID,
        odresconstants.id,
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

  def load(user: CompanyUser, id: Int): Option[Odresconstants] = {
    loadWhere(user, s"${C.ID} = {id}", 'id -> id)
  }

  def loadWhere(user: CompanyUser, where: String, args : scala.Tuple2[scala.Any, anorm.ParameterValue[_]]*) = {
    DB.withConnection(user.database.database) { implicit connection =>
      SQL(s"SELECT t1.*, t2.${C.FK_NAME} FROM ${C.ODRESCONSTANTS} t1 INNER JOIN ${C.???} t2 ON t1.${C.FK_ID} = t2.${C.FK_ID} WHERE $where")
        .on(args: _*)
        .as(odresconstantsParser.singleOpt)
    }
  }

  def delete(user: CompanyUser, id: Int) = {
    DB.withConnection(user.database.database) { implicit connection =>
      try {
        SQL(s"DELETE FROM ${C.ODRESCONSTANTS} WHERE ${C.ID} = {id}")
        .on('id -> id)
        .executeUpdate
      } catch {
        case NonFatal(e) => {
          Logger.error(s"can't delete a ${C.ODRESCONSTANTS}. ${C.ID} id: $id. Error ${e.toString}")
          throw e
        }
      }
    }
  }

  def get(user: CompanyUser, id: Int): Odresconstants = {
    load(user, id) match {
      case Some(p) => p
      case None => emptyOdresconstants
    }
  }
}


// Router

GET     /api/v1/general/odresconstants/:id              controllers.logged.modules.general.Odresconstantss.get(id: Int)
POST    /api/v1/general/odresconstants                  controllers.logged.modules.general.Odresconstantss.create
PUT     /api/v1/general/odresconstants/:id              controllers.logged.modules.general.Odresconstantss.update(id: Int)
DELETE  /api/v1/general/odresconstants/:id              controllers.logged.modules.general.Odresconstantss.delete(id: Int)




/**/
