public class cbrfBrowseFolder {

  // cbrfBrowseFolder
  // 2000/12/17 Copyright © 2000, Larry Rebich, using the VAIO
  // 2000/12/17 larry@buygold.net, www.buygold.net, 760.771.4730
  //            Some parts of this code from Randy Birch and others.
  //
  // Use this class module in conjunction with module: modBrowseFolder.bas

//Option Explicit
  *DefLng A-Z

  //'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  // Some parts copyright ©1996-2000 VBnet, Randy Birch, All Rights Reserved.
  // http://www.mvps.org/vbnet/index.html
  //'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  private static final String MCSTRNAME = "cbrfBrowseFolder.";

//*TODO:** type is translated as a new class at the end of the file Private Type BrowseInfo

  //'allocate and name storage for the structure
  private BrowseInfo BrowseInfo;

*TODO: API Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

*TODO: API Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BrowseInfo) As Long

*TODO: API Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

*TODO: API Private Declare Function SHSimpleIDListFromPath Lib "shell32" Alias "#162" (ByVal szPath As String) As Long

*TODO: API Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long

*TODO: API Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long

*TODO: API Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

  private static final int LMEM_FIXED = 0x0;
  private static final int LMEM_ZEROINIT = 0x40;
  private static final (LMEM_FIXED LPTR = LMEM_ZEROINIT);

  private int m_lngHwnd = 0;
  private String m_strTitle = "";
  private String m_strFolder = "";
  private boolean m_bCancelled = false;
  private int m_lngFlags = 0;


  //////////////////////////////////////////////////////////////////////////////////////////
  public void setFlags(int flags) {
    m_lngFlags = flags;
  }



  //////////////////////////////////////////////////////////////////////////////////////////
  public boolean getCancelled() {
    return m_bCancelled;
  }



  //////////////////////////////////////////////////////////////////////////////////////////
  public void setHWnd(int lnghWnd) {
    m_lngHwnd = lnghWnd;
  }



  //////////////////////////////////////////////////////////////////////////////////////////
  public void setTitle(String strValue) {
    m_strTitle = strValue;
  }



  //////////////////////////////////////////////////////////////////////////////////////////
  public void setFolder(String strValue) {
    m_strFolder = strValue;
  }



  //////////////////////////////////////////////////////////////////////////////////////////
  public String showBrowse(Object varHwnd, Object varTitle, Object varFolder) {
    String _rtn = "";
    //--------------------------------------------------------------------------
    // Procedure:   ShowBrowse
    // Description: Opens a Browse for Folder dialog box so the user can
    //              select an existing folder.
    // Returns:     Folder name that was selected
    // Params:      varHwnd - handle to the calling form
    //              varTitle - title to show on the Browse for Folder screen
    //-----------------------------------------------------------------------------
    "ShowBrowse"
.equals(Const cstrCurrentProc As String);
    int lngSelPath = 0;
    int lngRtn = 0;
    int lngPIDL = 0;
    int intPosition = 0;
    String strFolder = ""; * MAX_PATH

    try {

      // Allow user to specify params in the call
      if (!IsMissing(varHwnd)) {
        m_lngHwnd = varHwnd;
      }
      if (!IsMissing(varTitle)) {
        m_strTitle = varTitle;
      }
      if (!IsMissing(varFolder)) {
        m_strFolder = varFolder;
      }

      //'owner's hWnd
      browseInfo.hOwner = m_lngHwnd;
      //'flags, default is BIF_RETURNONLYFSDIRS
      browseInfo.uFlags = m_lngFlags;
      //'dialog's title
      browseInfo.lpszTitle = m_strTitle;
      //'set to pass an address into a structure
      browseInfo.lpfn = modBrowseFolder.fARPROC(AddressOf modBrowseFolder.browseCallbackProcStr());
      lngSelPath = LocalAlloc(LPTR, m_strFolder.length());
      MoveMemory(ByVal lngSelPath, ByVal m_strFolder, m_strFolder.length());
      //'now into structure
      browseInfo.lParam = lngSelPath;

      // Show Browse for Folder window
      lngPIDL = SHBrowseForFolder(browseInfo);

      if (lngPIDL) {
        // If not cancelled, translate the PIDL into a folder name, drop the characters in that
        // folder name that occur after the Null character, and then free memory.
        if (SHGetPathFromIDList(lngPIDL, strFolder)) {
          strFolder = strFolder.substring(0, strFolder.indexOf(vbNullChar) - 1);
        }
        CoTaskMemFree(lngPIDL);
      } 
      else {
        m_bCancelled = true;
      }

      // Free memory allocated for SelPath
      LocalFree(lngSelPath);

      // Update member variable with user's selection
      m_strFolder = strFolder;

      if (m_strFolder.equals("")) {
        // If the user clicked Cancel in the Browse For Folder window, then inform caller of that
        m_bCancelled = true;
      } 
      else {
        // Otherwise, return the name of the new selected folder
        _rtn = m_strFolder;
      }
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



//////////////////////////////////////////////////////////////////////////////////////////
  private void class_Initialize() {
    //--------------------------------------------------------------------------
    // Procedure:   Class_Initialize
    // Description: Instantiates this class, setting default values as appropriate
    //
    // Returns:     N/A
    // Params:      N/A
    //-----------------------------------------------------------------------------
    "Class_Initialize"
.equals(Const cstrCurrentProc As String);

    try {

      //' Set to default: return only real drives, no virtuals
      m_lngFlags = BIF_RETURNONLYFSDIRS;
      m_strTitle = "Select a Folder";
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
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
//*TODO:** the error label 0: couldn't be found
  }
}
}

private class BrowseInfo {
    public Long hOwner;
    public Long pidlRoot;
    public String pszDisplayName;
    public String lpszTitle;
    public Long uFlags;
    public Long lpfn;
    public Long lParam;
    public Long iImage;
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


case class BrfbrowsefolderData(
              id: Option[Int],

              )

object Brfbrowsefolders extends Controller with ProvidesUser {

  val brfbrowsefolderForm = Form(
    mapping(
      "id" -> optional(number),

  )(BrfbrowsefolderData.apply)(BrfbrowsefolderData.unapply))

  implicit val brfbrowsefolderWrites = new Writes[Brfbrowsefolder] {
    def writes(brfbrowsefolder: Brfbrowsefolder) = Json.obj(
      "id" -> Json.toJson(brfbrowsefolder.id),
      C.ID -> Json.toJson(brfbrowsefolder.id),

    )
  }

  def get(id: Int) = GetAction { implicit request =>
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.LIST_BRFBROWSEFOLDER), { user =>
      Ok(Json.toJson(Brfbrowsefolder.get(user, id)))
    })
  }

  def update(id: Int) = PostAction { implicit request =>
    Logger.debug("in brfbrowsefolders.update")
    brfbrowsefolderForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      brfbrowsefolder => {
        Logger.debug(s"form: ${brfbrowsefolder.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.EDIT_BRFBROWSEFOLDER), { user =>
          Ok(
            Json.toJson(
              Brfbrowsefolder.update(user,
                Brfbrowsefolder(
                       id,

                ))))
        })
      }
    )
  }

  def create = PostAction { implicit request =>
    Logger.debug("in brfbrowsefolders.create")
    brfbrowsefolderForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      brfbrowsefolder => {
        Logger.debug(s"form: ${brfbrowsefolder.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.NEW_BRFBROWSEFOLDER), { user =>
          Ok(
            Json.toJson(
              Brfbrowsefolder.create(user,
                Brfbrowsefolder(

                ))))
        })
      }
    )
  }

  def delete(id: Int) = PostAction { implicit request =>
    Logger.debug("in brfbrowsefolders.delete")
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.DELETE_BRFBROWSEFOLDER), { user =>
      Brfbrowsefolder.delete(user, id)
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

case class Brfbrowsefolder(
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

object Brfbrowsefolder {

  lazy val emptyBrfbrowsefolder = Brfbrowsefolder(
)

  def apply(
      id: Int,
) = {

    new Brfbrowsefolder(
      id,
)
  }

  def apply(
) = {

    new Brfbrowsefolder(
)
  }

  private val brfbrowsefolderParser: RowParser[Brfbrowsefolder] = {
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
        Brfbrowsefolder(
              id,
,
              createdAt,
              updatedAt,
              updatedBy)
    }
  }

  def create(user: CompanyUser, brfbrowsefolder: Brfbrowsefolder): Brfbrowsefolder = {
    save(user, brfbrowsefolder, true)
  }

  def update(user: CompanyUser, brfbrowsefolder: Brfbrowsefolder): Brfbrowsefolder = {
    save(user, brfbrowsefolder, false)
  }

  private def save(user: CompanyUser, brfbrowsefolder: Brfbrowsefolder, isNew: Boolean): Brfbrowsefolder = {
    def getFields = {
      List(

      )
    }
    def throwException = {
      throw new RuntimeException(s"Error when saving ${C.BRFBROWSEFOLDER}")
    }

    DBHelper.saveEx(
      user,
      Register(
        C.BRFBROWSEFOLDER,
        C.ID,
        brfbrowsefolder.id,
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

  def load(user: CompanyUser, id: Int): Option[Brfbrowsefolder] = {
    loadWhere(user, s"${C.ID} = {id}", 'id -> id)
  }

  def loadWhere(user: CompanyUser, where: String, args : scala.Tuple2[scala.Any, anorm.ParameterValue[_]]*) = {
    DB.withConnection(user.database.database) { implicit connection =>
      SQL(s"SELECT t1.*, t2.${C.FK_NAME} FROM ${C.BRFBROWSEFOLDER} t1 INNER JOIN ${C.???} t2 ON t1.${C.FK_ID} = t2.${C.FK_ID} WHERE $where")
        .on(args: _*)
        .as(brfbrowsefolderParser.singleOpt)
    }
  }

  def delete(user: CompanyUser, id: Int) = {
    DB.withConnection(user.database.database) { implicit connection =>
      try {
        SQL(s"DELETE FROM ${C.BRFBROWSEFOLDER} WHERE ${C.ID} = {id}")
        .on('id -> id)
        .executeUpdate
      } catch {
        case NonFatal(e) => {
          Logger.error(s"can't delete a ${C.BRFBROWSEFOLDER}. ${C.ID} id: $id. Error ${e.toString}")
          throw e
        }
      }
    }
  }

  def get(user: CompanyUser, id: Int): Brfbrowsefolder = {
    load(user, id) match {
      case Some(p) => p
      case None => emptyBrfbrowsefolder
    }
  }
}


// Router

GET     /api/v1/general/brfbrowsefolder/:id              controllers.logged.modules.general.Brfbrowsefolders.get(id: Int)
POST    /api/v1/general/brfbrowsefolder                  controllers.logged.modules.general.Brfbrowsefolders.create
PUT     /api/v1/general/brfbrowsefolder/:id              controllers.logged.modules.general.Brfbrowsefolders.update(id: Int)
DELETE  /api/v1/general/brfbrowsefolder/:id              controllers.logged.modules.general.Brfbrowsefolders.delete(id: Int)




/**/
