public class ctxsTaxSettings {

  //******************************************************************************
  // Module     : ctxsTaxSettings
  // Description: This class is a wrapper around the Settings table in the
  //              application database. By instantiating an object of type
  //              ctxsTaxSettings and using its Property Get procedures, the
  //              caller can get at each column in the Settings table (i.e. to
  //              limit data-entry to max length per the Settings table),
  //              without (a) using many global variables and (b) knowing that
  //              behind the scenes the data is coming from a table in the DB.
  // Procedures : Class_Initialize
  //              Property Get for each member variable (i.e. column in the Settings table) - public
  //              Property Let for each member variable (i.e. column in the Settings table) - private
  // Modified   :
  //  01/2002 BAW  Added support for new column in the Settings table: CurrentInterestRate.
  //              Also removed Property Let procedures since the app should not be
  //              updating these fields.
  //
  // --------------------------------------------------
//Option Explicit
  *Option Compare Binary

  private static final String MCSTRNAME = "ctxsTaxSettings.";

  //------------------------------------------
  //            MEMBER VARIABLES
  //------------------------------------------
  // member variable for Address1Length
  private int mintAddress1Length = 0;
  // member variable for Address2Length
  private int mintAddress2Length = 0;
  // member variable for AddressLength
  private int mintAddressLength = 0;
  // member variable for AddrTypeLength
  private int mintAddrTypeLength = 0;
  // member variable for ApplicationLength
  private int mintApplicationLength = 0;
  // member variable for CareOfLength
  private int mintCareOfLength = 0;
  // member variable for CityLength
  private int mintCityLength = 0;
  // member variable for ClaimNumber_GroupLength
  private int mintClaimNumber_GroupLength = 0;
  // member variable for ClaimNumber_IndivLength
  private int mintClaimNumber_IndivLength = 0;
  // member variable for CountryLength
  private int mintCountryLength = 0;
  // member variable for CurrentInterestRate
  private double mdblCurrentInterestRate = 0;
  // member variable for FedIDLength
  private int mintFedIDLength = 0;
  // member variable for Filler1Length
  private int mintFiller1Length = 0;
  // member variable for Filler2Length
  private int mintFiller2Length = 0;
  // member variable for Filler3Length
  private int mintFiller3Length = 0;
  // member variable for FormLength
  private int mintFormLength = 0;
  // member variable for InterestLength
  private int mintInterestLength = 0;
  // member variable for NameLength
  private int mintNameLength = 0;
  // member variable for ProductLength
  private int mintProductLength = 0;
  // member variable for RecordLength
  private int mintRecordLength = 0;
  // member variable for ResLength
  private int mintResLength = 0;
  // member variable for SecNameFlagLength
  private int mintSecNameFlagLength = 0;
  // member variable for SecondNameLength
  private int mintSecondNameLength = 0;
  // member variable for StateLength
  private int mintStateLength = 0;
  // member variable for StateResLength
  private int mintStateResLength = 0;
  // member variable for StateWthhld1Length
  private int mintStateWthhld1Length = 0;
  // member variable for SunCode
  private String mstrSunCode = "";
  // member variable for Tin2NoticeLength
  private int mintTIN2NoticeLength = 0;
  // member variable for TINLength
  private int mintTINLength = 0;
  // member variable for TINTypeLength
  private int mintTINTypeLength = 0;
  // member variable for ZipLength
  private int mintZipLength = 0;


  //------------------------------------------
  //           PROPERTY GET
  //------------------------------------------
  // Get Property Procedure for the Address1Length property
  public int getAddress1Length() {
    return mintAddress1Length;
  }


  // Get Property Procedure for the Address2Length property
  public int getAddress2Length() {
    return mintAddress2Length;
  }


  // Get Property Procedure for the AddressLength property
  public int getAddressLength() {
    return mintAddressLength;
  }


  // Get Property Procedure for the AddrTypeLength property
  public int getAddrTypeLength() {
    return mintAddrTypeLength;
  }


  // Get Property Procedure for the ApplicationLength property
  public int getApplicationLength() {
    return mintApplicationLength;
  }


  // Get Property Procedure for the CareOfLength property
  public int getCareOfLength() {
    return mintCareOfLength;
  }


  // Get Property Procedure for the CityLength property
  public int getCityLength() {
    return mintCityLength;
  }


  // Get Property Procedure for the ClaimNumber_GroupLength property
  public int getClaimNumber_GroupLength() {
    return mintClaimNumber_GroupLength;
  }


  // Get Property Procedure for the ClaimNumber_IndivLength property
  public int getClaimNumber_IndivLength() {
    return mintClaimNumber_IndivLength;
  }


  // Get Property Procedure for the CurrentInterestRate property
  public double getCurrentInterestRate() {
    return mdblCurrentInterestRate;
  }


  // Get Property Procedure for the CountryLength property
  public int getCountryLength() {
    return mintCountryLength;
  }


  // Get Property Procedure for the FedIDLength property
  public int getFedIDLength() {
    return mintFedIDLength;
  }


  // Get Property Procedure for the Filler1Length property
  public int getFiller1Length() {
    return mintFiller1Length;
  }


  // Get Property Procedure for the Filler2Length property
  public int getFiller2Length() {
    return mintFiller2Length;
  }


  // Get Property Procedure for the Filler3Length property
  public int getFiller3Length() {
    return mintFiller3Length;
  }


  // Get Property Procedure for the FormLength property
  public int getFormLength() {
    return mintFormLength;
  }


  // Get Property Procedure for the InterestLength property
  public int getInterestLength() {
    return mintInterestLength;
  }


  // Get Property Procedure for the NameLength property
  public int getNameLength() {
    return mintNameLength;
  }


  // Get Property Procedure for the ProductLength property
  public int getProductLength() {
    return mintProductLength;
  }


  // Get Property Procedure for the RecordLength property
  public int getRecordLength() {
    return mintRecordLength;
  }


  // Get Property Procedure for the ResLength property
  public int getResLength() {
    return mintResLength;
  }


  // Get Property Procedure for the SecNameFlagLength property
  public int getSecNameFlagLength() {
    return mintSecNameFlagLength;
  }


  // Get Property Procedure for the SecondNameLength property
  public int getSecondNameLength() {
    return mintSecondNameLength;
  }


  // Get Property Procedure for the StateLength property
  public int getStateLength() {
    return mintStateLength;
  }


  // Get Property Procedure for the StateResLength property
  public int getStateResLength() {
    return mintStateResLength;
  }


  // Get Property Procedure for the StateWthhld1Length property
  public int getStateWthhld1Length() {
    return mintStateWthhld1Length;
  }


  // Get Property Procedure for the SunCode property
  public String getSunCode() {
    return mstrSunCode;
  }


  // Get Property Procedure for the Tin2NoticeLength property
  public int getTin2NoticeLength() {
    return mintTIN2NoticeLength;
  }


  // Get Property Procedure for the TINLength property
  public int getTINLength() {
    return mintTINLength;
  }


  // Get Property Procedure for the TINTypeLength property
  public int getTINTypeLength() {
    return mintTINTypeLength;
  }


  // Get Property Procedure for the ZipLength property
  public int getZipLength() {
    return mintZipLength;
  }


  // ********************************************
  //  Initialize and Terminate
  // ********************************************
  private void class_Initialize() {
    // Comments  : Accesses the Settings table in CLAIMS.MDB
    //             and populates member variables with its values.
    // Called by : fnInitializeAppConnectionObject
    // Parameters: None
    // Modified  :
    // --------------------------------------------------
    "Class_Initialize"
.equals(Const cstrCurrentProc As String);
    try {

      DBRecordSet mrstSettings = null;
      String strSQL = "";

      mrstSettings = new ADODB.Recordset();

      // This SQL statement returns a single row, which happens to be the
      // entire contents of the Settings table
      strSQL = "SELECT [CurrentInterestRate], [RecordLength], "+ "[Form] as FormLength, [Fed-Id] As FedIDLength, "+ "[Application] As ApplicationLength, [Product] as ProductLength, "+ "[Res] As ResLength, [Name] as NameLength, "+ "[Second-Name] As SecondNameLength, [Address1] as Address1Length, "+ "[Address2] As Address2Length, [City] as CityLength, "+ "[Filler1] As Filler1Length, [State-Res] as StateResLength, "+ "[Filler2] As Filler2Length, [Zip] as ZipLength, "+ "[Addr-Type] As AddrTypeLength, [State-Wthhld1] as StateWthhld1Length, "+ "[TIN] As TINLength, [Tin-Type] as TINTypeLength, "+ "[State] As StateLength, [Suncode], "+ "[ClaimNumber_Group] As ClaimNumber_GroupLength, "+ "[ClaimNumber_Indiv] As ClaimNumber_IndivLength, [Country] as CountryLength, "+ "[Tin2-Notice] As TIN2NoticeLength, [Sec-Name-Flag] as SecNameFlagLength, "+ "[Filler3] As Filler3Length, [Interest] as InterestLength, "+ "[address] As AddressLength, [careof] as CareOfLength "+ "FROM [Settings]";

      // CursorType=adOpenKeyset   - Scrolling fwd/bwd permitted, chgs/del by other users visible
      // LockType=adLockReadOnly   - Recordset is read-only
      mrstSettings.Open(Source:=strSQL, ActiveConnection:=gconAppActive, CursorType:=adOpenKeyset, LockType:=adLockReadOnly, Options:=adCmdText);

      if (mrstSettings.RecordCount > 0) {
        mintAddress1Length = !getAddress1Length();
        mintAddress2Length = !getAddress2Length();
        mintAddressLength = !getAddressLength();
        mintAddrTypeLength = !getAddrTypeLength();
        mintApplicationLength = !getApplicationLength();
        mintCareOfLength = !getCareOfLength();
        mintCityLength = !getCityLength();
        mintCountryLength = !getCountryLength();
        mdblCurrentInterestRate = !getCurrentInterestRate();
        mintFedIDLength = !getFedIDLength();
        mintFiller1Length = !getFiller1Length();
        mintFiller2Length = !getFiller2Length();
        mintFiller3Length = !getFiller3Length();
        mintFormLength = !getFormLength();
        mintInterestLength = !getInterestLength();
        mintNameLength = !getNameLength();
        mintClaimNumber_GroupLength = !getClaimNumber_GroupLength();
        mintClaimNumber_IndivLength = !getClaimNumber_IndivLength();
        mintProductLength = !getProductLength();
        mintRecordLength = !getRecordLength();
        mintResLength = !getResLength();
        mintSecNameFlagLength = !getSecNameFlagLength();
        mintSecondNameLength = !getSecondNameLength();
        mintStateLength = !getStateLength();
        mintStateResLength = !getStateResLength();
        mintStateWthhld1Length = !getStateWthhld1Length();
        mstrSunCode = !getSunCode();
        mintTIN2NoticeLength = !getTin2NoticeLength();
        mintTINLength = !getTINLength();
        mintTINTypeLength = !getTINTypeLength();
        mintZipLength = !getZipLength();
      }
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    modGeneral.fnFreeRecordset(mrstSettings);

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


case class TxstaxsettingsData(
              id: Option[Int],

              )

object Txstaxsettingss extends Controller with ProvidesUser {

  val txstaxsettingsForm = Form(
    mapping(
      "id" -> optional(number),

  )(TxstaxsettingsData.apply)(TxstaxsettingsData.unapply))

  implicit val txstaxsettingsWrites = new Writes[Txstaxsettings] {
    def writes(txstaxsettings: Txstaxsettings) = Json.obj(
      "id" -> Json.toJson(txstaxsettings.id),
      C.ID -> Json.toJson(txstaxsettings.id),

    )
  }

  def get(id: Int) = GetAction { implicit request =>
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.LIST_TXSTAXSETTINGS), { user =>
      Ok(Json.toJson(Txstaxsettings.get(user, id)))
    })
  }

  def update(id: Int) = PostAction { implicit request =>
    Logger.debug("in txstaxsettingss.update")
    txstaxsettingsForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      txstaxsettings => {
        Logger.debug(s"form: ${txstaxsettings.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.EDIT_TXSTAXSETTINGS), { user =>
          Ok(
            Json.toJson(
              Txstaxsettings.update(user,
                Txstaxsettings(
                       id,

                ))))
        })
      }
    )
  }

  def create = PostAction { implicit request =>
    Logger.debug("in txstaxsettingss.create")
    txstaxsettingsForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      txstaxsettings => {
        Logger.debug(s"form: ${txstaxsettings.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.NEW_TXSTAXSETTINGS), { user =>
          Ok(
            Json.toJson(
              Txstaxsettings.create(user,
                Txstaxsettings(

                ))))
        })
      }
    )
  }

  def delete(id: Int) = PostAction { implicit request =>
    Logger.debug("in txstaxsettingss.delete")
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.DELETE_TXSTAXSETTINGS), { user =>
      Txstaxsettings.delete(user, id)
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

case class Txstaxsettings(
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

object Txstaxsettings {

  lazy val emptyTxstaxsettings = Txstaxsettings(
)

  def apply(
      id: Int,
) = {

    new Txstaxsettings(
      id,
)
  }

  def apply(
) = {

    new Txstaxsettings(
)
  }

  private val txstaxsettingsParser: RowParser[Txstaxsettings] = {
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
        Txstaxsettings(
              id,
,
              createdAt,
              updatedAt,
              updatedBy)
    }
  }

  def create(user: CompanyUser, txstaxsettings: Txstaxsettings): Txstaxsettings = {
    save(user, txstaxsettings, true)
  }

  def update(user: CompanyUser, txstaxsettings: Txstaxsettings): Txstaxsettings = {
    save(user, txstaxsettings, false)
  }

  private def save(user: CompanyUser, txstaxsettings: Txstaxsettings, isNew: Boolean): Txstaxsettings = {
    def getFields = {
      List(

      )
    }
    def throwException = {
      throw new RuntimeException(s"Error when saving ${C.TXSTAXSETTINGS}")
    }

    DBHelper.saveEx(
      user,
      Register(
        C.TXSTAXSETTINGS,
        C.ID,
        txstaxsettings.id,
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

  def load(user: CompanyUser, id: Int): Option[Txstaxsettings] = {
    loadWhere(user, s"${C.ID} = {id}", 'id -> id)
  }

  def loadWhere(user: CompanyUser, where: String, args : scala.Tuple2[scala.Any, anorm.ParameterValue[_]]*) = {
    DB.withConnection(user.database.database) { implicit connection =>
      SQL(s"SELECT t1.*, t2.${C.FK_NAME} FROM ${C.TXSTAXSETTINGS} t1 INNER JOIN ${C.???} t2 ON t1.${C.FK_ID} = t2.${C.FK_ID} WHERE $where")
        .on(args: _*)
        .as(txstaxsettingsParser.singleOpt)
    }
  }

  def delete(user: CompanyUser, id: Int) = {
    DB.withConnection(user.database.database) { implicit connection =>
      try {
        SQL(s"DELETE FROM ${C.TXSTAXSETTINGS} WHERE ${C.ID} = {id}")
        .on('id -> id)
        .executeUpdate
      } catch {
        case NonFatal(e) => {
          Logger.error(s"can't delete a ${C.TXSTAXSETTINGS}. ${C.ID} id: $id. Error ${e.toString}")
          throw e
        }
      }
    }
  }

  def get(user: CompanyUser, id: Int): Txstaxsettings = {
    load(user, id) match {
      case Some(p) => p
      case None => emptyTxstaxsettings
    }
  }
}


// Router

GET     /api/v1/general/txstaxsettings/:id              controllers.logged.modules.general.Txstaxsettingss.get(id: Int)
POST    /api/v1/general/txstaxsettings                  controllers.logged.modules.general.Txstaxsettingss.create
PUT     /api/v1/general/txstaxsettings/:id              controllers.logged.modules.general.Txstaxsettingss.update(id: Int)
DELETE  /api/v1/general/txstaxsettings/:id              controllers.logged.modules.general.Txstaxsettingss.delete(id: Int)




/**/
