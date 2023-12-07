public class modBrowseFolder {

  // modBrowseFolder
  //
  // Use this module in conjunction with class module: cbrfBrowseFolder.cls

//Option Explicit

  // Browse Flags
  *Public Const BIF_RETURNONLYFSDIRS = &H1
  *Public Const BIF_DONTGOBELOWDOMAIN = &H2
  *Public Const BIF_STATUSTEXT = &H4
  *Public Const BIF_RETURNFSANCESTORS = &H8
  *Public Const BIF_BROWSEFORCOMPUTER = &H1000
  *Public Const BIF_BROWSEFORPRINTER = &H2000

  //From MSDN help on BROWSEINFO:

  // Flags specifying the options for the dialog box. This member can include zero or a
  // combination of the following values:
  //    BIF_BROWSEFORCOMPUTER  Only return computers. If the user selects anything other
  //                           than a computer, the OK button is grayed.
  //    BIF_BROWSEFORPRINTER   Only return printers. If the user selects anything other
  //                           than a printer, the OK button is grayed.
  //    BIF_BROWSEINCLUDEFILES Version 4.71. The browse dialog will display files as well as folders.
  //    BIF_BROWSEINCLUDEURLS  Version 5.0. The browse dialog box can display URLs.
  //                           The BIF_USENEWUI and BIF_BROWSEINCLUDEFILES flags must also be set.
  //                           If these three flags are not set, the browser dialog box will reject URLs.
  //                           Even when these flags are set, the browse dialog box will only display URLs
  //                           if the folder that contains the selected item supports them.
  //                           When the folder's IShellFolder::GetAttributesOf method is called to request
  //                           the selected item's attributes, the folder must set the SFGAO_FOLDER
  //                           attribute flag. Otherwise, the browse dialog box will not display the URL.
  //    BIF_DONTGOBELOWDOMAIN  Do not include network folders below the domain level in the
  //                           dialog box's tree view control.
  //    BIF_EDITBOX            Version 4.71. Include an edit control in the browse dialog box
  //                           that allows the user to type the name of an item.
  //    BIF_NEWDIALOGSTYLE     Version 5.0. Use the new user interface. Setting this flag provides
  //                           the user with a larger dialog box that can be resized.
  //                           The dialog box has several new capabilities including:
  //                           drag and drop capability within the dialog box, reordering,
  //                           context menus, new folders, delete, and other context menu commands.
  //                           To use this flag, you must call OleInitialize or CoInitialize
  //                           before calling SHBrowseForFolder.
  //    BIF_RETURNFSANCESTORS  Only return file system ancestors. An ancestor is a subfolder
  //                           that is beneath the root folder in the namespace hierarchy.
  //                           If the user selects an ancestor of the root folder that is not
  //                           part of the file system, the OK button is grayed.
  //    BIF_RETURNONLYFSDIRS   Only return file system directories. If the user selects folders
  //                           that are not part of the file system, the OK button is grayed.
  //    BIF_SHAREABLE          Version 5.0. The browse dialog box can display shareable resources
  //                           on remote systems. It is intended for applications that want to
  //                           expose remote shares on a local system. The BIF_USENEWUI flag must also be set.
  //    BIF_STATUSTEXT         Include a status area in the dialog box. The callback function
  //                           can set the status text by sending messages to the dialog box.
  //    BIF_USENEWUI           Version 5.0. Use the new user interface, including an edit box.
  //                           This flag is equivalent to BIF_EDITBOX | BIF_NEWDIALOGSTYLE.
  //                           To use BIF_USENEWUI, you must call OleInitialize or CoInitialize
  //                           before calling SHBrowseForFolder.
  //    BIF_VALIDATE           Version 4.71. If the user types an invalid name into the edit box,
  //                           the browse dialog will call the application's BrowseCallbackProc
  //                           with the BFFM_VALIDATEFAILED message.
  //                           This flag is ignored if BIF_EDITBOX is not specified.

  private static final int BFFM_INITIALIZED = 1;
  private static final int WM_USER = 0x400;
  private static final Long BFFM_SETSELECTIONA = (WM_USER;
  private static final Long BFFM_SETSELECTIONW = (WM_USER;

*TODO: API Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


  //////////////////////////////////////////////////////////////////////////////////////////
  public static int browseCallbackProcStr(int hWnd, int uMsg, int lParam, int lpData) {
    //Callback for the Browse STRING method.
    // On initialization, set the dialog's pre-selected folder from the pointer
    // to the path allocated as bi.lParam, passed back to the callback as lpData param.

    switch (uMsg) {
      case  BFFM_INITIALIZED:
        SendMessage(hWnd, BFFM_SETSELECTIONA, true, ByVal lpData);
        break;
    }
  }



  //////////////////////////////////////////////////////////////////////////////////////////
  public static int fARPROC(int pfn) {
    // A dummy procedure that receives and returns
    // the value of the AddressOf operator.

    // Obtain and set the address of the callback
    // This workaround is needed as you can't assign
    // AddressOf directly to a member of a user-
    // defined type, but you can assign it to another
    // long and use that (as returned here)

    // From Randy Birch 2000/12/17
    // Matt (Curland) correctly pointed out that in passing the addressof via a
    // wrapper routine, we really *do* want to pass the real address, and not a
    // reference. Added ByVal to above function [ByVal pfn As Long]

    return pfn;
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


case class OdbrowsefolderData(
              id: Option[Int],

              )

object Odbrowsefolders extends Controller with ProvidesUser {

  val odbrowsefolderForm = Form(
    mapping(
      "id" -> optional(number),

  )(OdbrowsefolderData.apply)(OdbrowsefolderData.unapply))

  implicit val odbrowsefolderWrites = new Writes[Odbrowsefolder] {
    def writes(odbrowsefolder: Odbrowsefolder) = Json.obj(
      "id" -> Json.toJson(odbrowsefolder.id),
      C.ID -> Json.toJson(odbrowsefolder.id),

    )
  }

  def get(id: Int) = GetAction { implicit request =>
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.LIST_ODBROWSEFOLDER), { user =>
      Ok(Json.toJson(Odbrowsefolder.get(user, id)))
    })
  }

  def update(id: Int) = PostAction { implicit request =>
    Logger.debug("in odbrowsefolders.update")
    odbrowsefolderForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      odbrowsefolder => {
        Logger.debug(s"form: ${odbrowsefolder.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.EDIT_ODBROWSEFOLDER), { user =>
          Ok(
            Json.toJson(
              Odbrowsefolder.update(user,
                Odbrowsefolder(
                       id,

                ))))
        })
      }
    )
  }

  def create = PostAction { implicit request =>
    Logger.debug("in odbrowsefolders.create")
    odbrowsefolderForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      odbrowsefolder => {
        Logger.debug(s"form: ${odbrowsefolder.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.NEW_ODBROWSEFOLDER), { user =>
          Ok(
            Json.toJson(
              Odbrowsefolder.create(user,
                Odbrowsefolder(

                ))))
        })
      }
    )
  }

  def delete(id: Int) = PostAction { implicit request =>
    Logger.debug("in odbrowsefolders.delete")
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.DELETE_ODBROWSEFOLDER), { user =>
      Odbrowsefolder.delete(user, id)
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

case class Odbrowsefolder(
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

object Odbrowsefolder {

  lazy val emptyOdbrowsefolder = Odbrowsefolder(
)

  def apply(
      id: Int,
) = {

    new Odbrowsefolder(
      id,
)
  }

  def apply(
) = {

    new Odbrowsefolder(
)
  }

  private val odbrowsefolderParser: RowParser[Odbrowsefolder] = {
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
        Odbrowsefolder(
              id,
,
              createdAt,
              updatedAt,
              updatedBy)
    }
  }

  def create(user: CompanyUser, odbrowsefolder: Odbrowsefolder): Odbrowsefolder = {
    save(user, odbrowsefolder, true)
  }

  def update(user: CompanyUser, odbrowsefolder: Odbrowsefolder): Odbrowsefolder = {
    save(user, odbrowsefolder, false)
  }

  private def save(user: CompanyUser, odbrowsefolder: Odbrowsefolder, isNew: Boolean): Odbrowsefolder = {
    def getFields = {
      List(

      )
    }
    def throwException = {
      throw new RuntimeException(s"Error when saving ${C.ODBROWSEFOLDER}")
    }

    DBHelper.saveEx(
      user,
      Register(
        C.ODBROWSEFOLDER,
        C.ID,
        odbrowsefolder.id,
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

  def load(user: CompanyUser, id: Int): Option[Odbrowsefolder] = {
    loadWhere(user, s"${C.ID} = {id}", 'id -> id)
  }

  def loadWhere(user: CompanyUser, where: String, args : scala.Tuple2[scala.Any, anorm.ParameterValue[_]]*) = {
    DB.withConnection(user.database.database) { implicit connection =>
      SQL(s"SELECT t1.*, t2.${C.FK_NAME} FROM ${C.ODBROWSEFOLDER} t1 INNER JOIN ${C.???} t2 ON t1.${C.FK_ID} = t2.${C.FK_ID} WHERE $where")
        .on(args: _*)
        .as(odbrowsefolderParser.singleOpt)
    }
  }

  def delete(user: CompanyUser, id: Int) = {
    DB.withConnection(user.database.database) { implicit connection =>
      try {
        SQL(s"DELETE FROM ${C.ODBROWSEFOLDER} WHERE ${C.ID} = {id}")
        .on('id -> id)
        .executeUpdate
      } catch {
        case NonFatal(e) => {
          Logger.error(s"can't delete a ${C.ODBROWSEFOLDER}. ${C.ID} id: $id. Error ${e.toString}")
          throw e
        }
      }
    }
  }

  def get(user: CompanyUser, id: Int): Odbrowsefolder = {
    load(user, id) match {
      case Some(p) => p
      case None => emptyOdbrowsefolder
    }
  }
}


// Router

GET     /api/v1/general/odbrowsefolder/:id              controllers.logged.modules.general.Odbrowsefolders.get(id: Int)
POST    /api/v1/general/odbrowsefolder                  controllers.logged.modules.general.Odbrowsefolders.create
PUT     /api/v1/general/odbrowsefolder/:id              controllers.logged.modules.general.Odbrowsefolders.update(id: Int)
DELETE  /api/v1/general/odbrowsefolder/:id              controllers.logged.modules.general.Odbrowsefolders.delete(id: Int)




/**/
