public class modComboBox {

  //******************************************************************************
  // Module      : modComboBox
  // Procedures  :
  //               fnAddItemDataComboBox( cboIn As ComboBox, ByVal strNewString As String, _
  //                   Optional ByVal lngItemData As Long) As Long
  //               fnADORecordSetToComboBox( rstIn As ADODB.Recordset, cboIn As ComboBox, _
  //                   ByVal strDisplayColumn As String, Optional ByVal varItemDataColumn As Variant)
  //               fnADORecordsetToVFG(ByVal rstIn As ADODB.Recordset, _
  //                   ByRef pvfgIn As VSFlexGrid, Optional ByVal intGridCol As Integer = 0, _
  //                   Optional ByVal varFieldNamesIn As Variant)
  //               fnBinarySearchComboBox( cboIn As ComboBox, ByVal strSearch As String) _
  //                   As Integer
  //               fnDAORecordSetToComboBox( rstIn As DAO.Recordset, cboIn As ComboBox, _
  //                   ByVal strDisplayColumn As String, Optional ByVal varItemDataColumn As Variant)
  //               fnDropComboBoxList( cboIn As ComboBox, ByVal fShow As Boolean)
  //               fnFindStringComboBox(ByRef cboIn As ComboBox, ByVal strSearchIn As String, _
  //                   Optional ByVal bDoExactSearch As Boolean = False) As Long
  //               fnIncrementalSearchCombo( cboIn As ComboBox, intKeyAscii As Integer)
  //               fnInitializefpCombo(ByRef lpcIn As LPADOLib.fpCombo, ByVal intNbrOfCols As Integer, _
  //                   Optional ByVal intColToDisplay As Integer = 0)
  //               fnIsComboBoxDropped(cboIn) As Boolean
  //               fnSearchCBOItemData(ByVal cboIn As ComboBox, ByVal strSearchIn As String) _
  //                   As Integer
  //               fnSearchFPCombo(ByRef lpcIn As LPLib.fpCombo, ByVal strSearchText As String, _
  //                   Optional ByVal intSearchCol As Integer = 0, _
  //                   Optional ByVal lngSearchMethod As lplib.SearchMethodConstants = SearchMethodPartialMatch, _
  //                   Optional ByVal bDefaultToFirstRowIfNotFound As Boolean = True)
  //               fnSetComboBoxItemHeight( cboIn As ComboBox, ByVal sngMultipleItemHeight As Single)
  //               fnSetComboBoxListItems( cboIn As ComboBox, ByVal intItems As Integer)
  //               fnSetComboBoxListWidth( cboIn As ComboBox, ByVal sngMultipleListWidth As Single)
  // Description : Routines to extend the functionality of a standard
  //               VB ComboBox control
  // Source      : Total Visual SourceBook 2000
  //
  // Modified:
  //  01/2002 BAW  Updated fnADORecordsetToComboBox to make it more flexible.
  // --------------------------------------------------
//Option Explicit
  *Option Compare Binary
  private static final String MCSTRNAME = "modComboBox.";

  public static final Long GCLNGNOSELECTION = -1;

*TODO: API Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
*TODO: API Private Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As String) As Long
*TODO: API Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
*TODO: API Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
*TODO: API Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

  private static final int CB_ADDSTRING = 0x143;
  private static final int CB_DELETESTRING = 0x144;
  private static final int CB_DIR = 0x145;
  *TODO:** (the data type can't be found for the value [(-1)])Private Const CB_ERR = (-1)
  *TODO:** (the data type can't be found for the value [(-2)])Private Const CB_ERRSPACE = (-2)
  private static final int CB_FINDSTRING = 0x14C;
  private static final int CB_FINDSTRINGEXACT = 0x158;
  private static final int CB_GETCOUNT = 0x146;
  private static final int CB_GETCURSEL = 0x147;
  private static final int CB_GETDROPPEDCONTROLRECT = 0x152;
  private static final int CB_GETDROPPEDSTATE = 0x157;
  private static final int CB_GETEDITSEL = 0x140;
  private static final int CB_GETEXTENDEDUI = 0x156;
  private static final int CB_GETITEMDATA = 0x150;
  private static final int CB_GETITEMHEIGHT = 0x154;
  private static final int CB_GETLBTEXT = 0x148;
  private static final int CB_GETLBTEXTLEN = 0x149;
  private static final int CB_GETLOCALE = 0x15A;
  private static final int CB_INSERTSTRING = 0x14A;
  private static final int CB_LIMITTEXT = 0x141;
  private static final int CB_MSGMAX = 0x15B;
  private static final int CB_OKAY = 0;
  private static final int CB_RESETCONTENT = 0x14B;
  private static final int CB_SELECTSTRING = 0x14D;
  private static final int CB_SETCURSEL = 0x14E;
  private static final int CB_SETEDITSEL = 0x142;
  private static final int CB_SETEXTENDEDUI = 0x155;
  private static final int CB_SETITEMDATA = 0x151;
  private static final int CB_SETITEMHEIGHT = 0x153;
  private static final int CB_SETLOCALE = 0x159;
  private static final int CB_SHOWDROPDOWN = 0x14F;
  private static final int CB_GETHORIZONTALEXTENT = 0x15D;
  private static final int CB_SETHORIZONTALEXTENT = 0x15E;
  private static final int CB_SETDROPPEDWIDTH = 0x160;

  private static final int CBS_AUTOHSCROLL = 0x40&;
  private static final int CBS_DISABLENOSCROLL = 0x800&;

  private static final int WM_SETREDRAW = 0xB;

//*TODO:** type is translated as a new class at the end of the file Private Type POINTAPI

//*TODO:** type is translated as a new class at the end of the file Private Type RECT


  //////////////////////////////////////////////////////////////////////////////////////////////////
  public static int fnAddItemDataComboBox(ComboBox cboIn, String strNewString, int lngItemData) { // TODO: Use of ByRef founded Public Function fnAddItemDataComboBox(ByRef cboIn As ComboBox, ByVal strNewString As String, Optional ByVal lngItemData As Long) As Long
    int _rtn = 0;
    // Comments  : Adds an item to a combobox, and the
    //             itemdata value associated with that item,
    //             in a single function call
    // Parameters: cboIn - ComboBox control to add to
    //             strNewString - string to add to list box
    //             lngItemData - value to add to associated
    //             ItemData() entry
    // Returns   : the ListIndex value of the newly-added item
    // Source    : Total Visual SourceBook 2000
    //
    "fnAddItemDataComboBox"
.equals(Const cstrCurrentProc As String);
    try {

      cboIn.AddItem(strNewString);
      cboIn.ItemData(cboIn.NewIndex) = lngItemData;
      _rtn = cboIn.NewIndex;
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
  public static void fnADORecordSetToComboBox(DBRecordSet rstIn, ComboBox cboIn, String strDisplayColumn, Object varItemDataColumn, boolean bClear, boolean bAddAllEntry, boolean bAddNullEntry, boolean bAddBlankEntry) { // TODO: Use of ByRef founded Public Sub fnADORecordSetToComboBox(ByRef rstIn As ADODB.Recordset, ByRef cboIn As ComboBox, ByVal strDisplayColumn As String, Optional ByVal varItemDataColumn As Variant, Optional ByVal bClear As Boolean = True, Optional ByVal bAddAllEntry As Boolean = False, Optional ByVal bAddNullEntry As Boolean = False, Optional ByVal bAddBlankEntry As Boolean = False)
    // Comments  : Displays the contents of an ADO recordset in
    //             a standard unbound combo box
    // Parameters:
    //    rstIn             - Recordset to read. Caller must create
    //    cboIn             - Combo box to load
    //    strDisplayColumn  - name of the column in rstIn to display
    //                        in the combo box
    //    varItemDataColumn - name of the column in rstIn to load into the
    //                        ItemData property of the combo box. The data in
    //                        this column MUST be storable in a long integer, and there
    //                        must be no 'null' values. Generally this field will be a
    //                        long integer Primary Key value associated with the value
    //                        to be displayed in the list
    //    bAddBlankEntry    - add a blank entry first?
    //    bAddAllEntry      - add an "--All--" item first?
    //    bAddNullEntry     - add an "<NULL> item next?
    // Returns   : Nothing
    // Source    : Total Visual SourceBook 2000
    // Modified  :
    //  01/2002 BAW Added bClear optional parameter, so this procedure could be
    //  used even by those comboboxes that need a blank or "**" entry added at
    //  the beginning (before the recordset is loaded).
    //
    //  05/10/02 CMP - added bAddAllItem in.
    //
    "fnADORecordSetToComboBox"
.equals(Const cstrCurrentProc As String);
    String strIDC = "";
    boolean bSplashLoaded = false;

    try {

      // if a column name is supplied in the varItemDataColumn parameter,
      // use this as the field name in the recordset to use to supply values
      // as the ItemData property of the list array
      if (!IsMissing(varItemDataColumn)) {
        strIDC = CStr(varItemDataColumn);
      } 
      else {
        strIDC = "";
      }

      if (bClear) {
        cboIn.Clear;
      }

      if (bAddBlankEntry) {
        cboIn.AddItem(modGeneral.gCSTRBLANKENTRY);
      }

      if (bAddAllEntry) {
        cboIn.AddItem(modGeneral.gCSTRALLENTRY);
      }

      if (bAddNullEntry) {
        cboIn.AddItem(modGeneral.gCSTRNULLENTRY);
      }

      bSplashLoaded = modGeneral.fnIsFormLoaded("frmSplash");

      // Load the named column in each row of the recordset into the combo box
      // If specified, load the value from the selected column into the
      // ItemData property. Use the original sort order, and all existing
      // rows of the recordset that is passed.

      if (!(rstIn.BOF && rstIn.EOF)) {
        rstIn.MoveFirst;
        do Until .EOF          //'ECG 5/16/2002 added this for nullable CBOs inwhich all returned DB records are NULL
          if (!(rstIn(strDisplayColumn) == null)) {
            cboIn.AddItem(rstIn(strDisplayColumn));
            if (strIDC != "") {
              cboIn.ItemData(cboIn.NewIndex) = rstIn(strIDC);
            }
            if (bSplashLoaded) {
              //' Allow progress meter to get updated!
              DoEvents;
            }
          //'ECG 5/16/2002 added this for nullable CBOs inwhich all returned DB records are NULL
          }
          rstIn.MoveNext;
        }
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



  *#If False Then
  //////////////////////////////////////////////////////////////////////////////////////////////////
  public static void fnADORecordsetToVFG(DBRecordSet rstIn, VSFlexGrid pvfgIn, int intGridCol, Object varFieldNamesIn) { // TODO: Use of ByRef founded Public Sub fnADORecordsetToVFG(ByVal rstIn As ADODB.Recordset, ByRef pvfgIn As VSFlexGrid, Optional ByVal intGridCol As Integer = 0, Optional ByVal varFieldNamesIn As Variant)
    // Comments  : Populates a VSFlexgrid column with the complete contents of an ADO recordset.
    // Parameters: rstIn  (in)          - Recordset to read. Caller must create
    //             pvfgIn (in/out)      - VSFlexGrid to load
    //             intGridCol (in)      - the column to hide
    //             varFieldNamesIn (in) - when only selective columns of the recordset
    //                                    should be used to populate the grid. Caller must set.
    // Returns   : Nothing
    // Modified  :
    //
    "fnADORecordsetToVFG"
.equals(Const cstrCurrentProc As String);
    String strComboBoxString = "";
    String strFieldNamesString = "";
    DBField fldName = null;

    try {

      if (IsMissing(varFieldNamesIn)) {
        // Collect all of our field names - the BuildComboList needs them.
        strFieldNamesString = "";
        for (int _i = 0; _i < rstIn.Fields.size(); _i++) {
          fldName = rstIn.Fields.item(_i);
          strFieldNamesString = strFieldNamesString+ fldName.Name+ ",";
        }

        // Chop the last "," off
        strFieldNamesString = strFieldNamesString.substring(0, strFieldNamesString.length() - 1);
      } 
      else {
        strFieldNamesString = varFieldNamesIn;
      }

      // Build the string
      strComboBoxString = pvfgIn.BuildComboList(rstIn, strFieldNamesString);

      pvfgIn.Select(0, 0);

      // Populate the ComboBox and indicate that the dropdown arrow should always be visible.
      pvfgIn.ColComboList(intGridCol) = strComboBoxString;
      pvfgIn.ShowComboButton = flexSBAlways;
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {

    // Clean-up statements go here
    modGeneral.fnFreeObject(fldName);

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
  *#End If




  //////////////////////////////////////////////////////////////////////////////////////////////////
  public static int fnBinarySearchComboBox(ComboBox cboIn, String strSearch) {
    int _rtn = 0;
    // Comments  : Use a binary search to find a value in a combo box
    // Parameters: cboIn - combo box to search. combo box must have
    //             the .Sorted property set to true
    //             strSearch - string to search for.
    // Returns   : A value of 0 or greater indicates the line on
    //             which the string was found in the combo box. A
    //             negative value indicates that the line was not found
    // Source    : Total Visual SourceBook 2000
    //
    "fnBinarySearchComboBox"
.equals(Const cstrCurrentProc As String);
    int intLineNum = 0;
    int intNumRecs = 0;
    boolean fFound = false;
    int intLowNum = 0;
    int intHighNum = 0;
    int intMidNum = 0;

    try {

      // Negative number indicates string not found
      intLineNum = -1;

      intNumRecs = cboIn.ListCount;
      fFound = false;
      intLowNum = 0;
      intHighNum = intNumRecs - 1;

      do      intMidNum = (intLowNum + intHighNum) \ 2;
      // check first half of list. If found search the bottom half
      if (strSearch.toUpperCase() < cboIn.List(intMidNum).toUpperCase()) {
        intHighNum = intMidNum - 1;
        // check the last half of the list. If found search the top half
      } 
      else if (strSearch.toUpperCase() > cboIn.List(intMidNum).toUpperCase()) {
        intLowNum = intMidNum + 1;
      } 
      else {
        // value found
        fFound = true;
        intLineNum = intMidNum;
      }
    } while (fFound || (intHighNum < intLowNum)) {

    // return line found, or -1 if not found
    _rtn = intLineNum;
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
// NOTE: The following requires a reference to DAO in order to compile cleanly
//
// Public Sub fnDAORecordSetToComboBox( rstIn As DAO.Recordset, cboIn As ComboBox, _
//     ByVal strDisplayColumn As String, Optional ByVal varItemDataColumn As Variant)
//     ' Comments  : Displays the contents of a recordset in
//     '             a standard unbound combo box
//     ' Parameters: rstIn - recordset to read. Caller must create
//     '             cboIn - combo box to load
//     '             strDisplayColumn - name of the column in rstIn to display
//     '             in the combo box
//     '             varItemDataColumn - name of the column in rstIn to load
//     '             into the ItemData property of the combo box. The data in
//     '             this column MUST be storable in a long integer, and there
//     '             must be no 'null' values. Generally this field will be a
//     '             long integer Primary Key value associated with the value
//     '             to be displayed in the list
//     ' Returns   : Nothing
//     ' Source    : Total Visual SourceBook 2000
//     '
//     Const cstrCurrentProc As String = "fnDAORecordSetToComboBox"
//     Dim strIDC As String
//
//     On Error GoTo PROC_ERR
//
//     ' if a column name is supplied in the varItemDataColumn parameter,
//     ' use this as the field name in the recordset to use to supply values
//     ' as the ItemData property of the list array
//     If Not IsMissing(varItemDataColumn) Then
//         strIDC = CStr(varItemDataColumn)
//     Else
//         strIDC = vbNullString
//     End If
//
//     cboIn.Clear
//
//     ' load the named column in each row of the recordset into the combo box
//     ' If specified, load the value from the selected column into the
//     ' ItemData property. Use the original sort order, and all existing
//     ' rows of the recordset that is passed.
//     With rstIn
//         If .RecordCount <> 0 Then
//             .MoveFirst
//             Do Until .EOF
//                 cboIn.AddItem rstIn(strDisplayColumn)
//                 If strIDC <> vbNullString Then
//                     cboIn.ItemData(cboIn.NewIndex) = rstIn(strIDC)
//                 End If
//                 .MoveNext
//             Loop
//         End If
//
//     End With
//PROC_EXIT:
//    On Error GoTo 0     ' disable error handler
//    ' Clean-up statements go here
//    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
//        gerhApp.PropagateError mcstrName & cstrCurrentProc
//    End If
//    Exit Sub
//PROC_ERR:
//    Select Case Err.Number
//        'Case statements for expected errors go here
//        Case Else
//            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
//    End Select
//    Resume PROC_EXIT
// End Sub
//



//////////////////////////////////////////////////////////////////////////////////////////////////
  public static void fnDropComboBoxList(ComboBox cboIn, boolean fShow) { // TODO: Use of ByRef founded Public Sub fnDropComboBoxList(ByRef cboIn As ComboBox, ByVal fShow As Boolean)
    // Comments  : Causes a combo box to show or hide the list portion
    // Parameters: cboIn - the combo box to modify
    //             fShow - True to show the list, false to hide the list
    // Returns   : Nothing
    // Source    : Total Visual SourceBook 2000
    //
    "fnDropComboBoxList"
.equals(Const cstrCurrentProc As String);
    int lngReturn = 0;

    try {

      lngReturn = SendMessage(cboIn.hWnd, CB_SHOWDROPDOWN, fShow, ByVal 0&);
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



//////////////////////////////////////////////////////////////////////////////////////////////////
  public static int fnFindStringComboBox(ComboBox cboIn, String strSearchIn, boolean bDoExactSearch) { // TODO: Use of ByRef founded Public Function fnFindStringComboBox(ByRef cboIn As ComboBox, ByVal strSearchIn As String, Optional ByVal bDoExactSearch As Boolean = False) As Long
    int _rtn = 0;
    // Comments  : Finds the line on a combo box containing the search string
    // Parameters: cboIn          (in/out) - The combo box to search
    //             strSearchIn    (in)     - The search value
    //             bDoExactSearch (in)     - True to find exact matches only;
    //                                       False to find partial matching prefixes
    //
    // Returns   : -1 if string not found; otherwise the index of the line
    //             containing the string
    // Source    : Total Visual SourceBook 2000
    //
    "fnFindStringComboBox"
.equals(Const cstrCurrentProc As String);
    int lngFound = 0;

    try {

      if (bDoExactSearch) {
        lngFound = SendMessage(cboIn.hWnd, CB_FINDSTRINGEXACT, -1, ByVal strSearchIn);

      } 
      else {
        lngFound = SendMessage(cboIn.hWnd, CB_FINDSTRING, -1, ByVal strSearchIn);
      }

      _rtn = lngFound;
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
  public static void fnIncrementalSearchCombo(ComboBox cboIn, int intKeyAscii) { // TODO: Use of ByRef founded Public Sub fnIncrementalSearchCombo(ByRef cboIn As ComboBox, ByRef intKeyAscii As Integer)
    // Comments  : Allows incremental searching of a combo box.
    //             Call this proc from the KeyUp event of the combo
    // Parameters: cboIn - the combo box to search
    //             intKeyAscii - the ASCII key value passed to the KeyPress event
    //             of the combo box
    // Returns   : Nothing
    // Source    : Total Visual SourceBook 2000
    //
    try {

      "fnIncrementalSearchCombo"
.equals(Const cstrCurrentProc As String);
      int cb = 0;
      String fnFindString = "";
      int intLoop = 0;

      if (intKeyAscii < 32 || intKeyAscii > 127) {
        // **TODO:** goto found: GoTo PROC_EXIT;
      }

      if (cboIn.SelLength == 0) {
        Debug.Print("SelLength = "+ cboIn.SelLength);
        fnFindString = cboIn.Text+ Chr$(intKeyAscii);
      } 
      else {
        fnFindString = cboIn.Text.substring(0, cboIn.SelStart)+ Chr$(intKeyAscii);
      }

      Debug.Print("Looking for: ["+ fnFindString+ "]");
      cb = SendMessageStr(cboIn.hWnd, CB_FINDSTRING, 1, ByVal fnFindString);

      if (cb != CB_ERR) {
        Debug.Print("Found it...");
        cboIn.Text = cboIn.List(cb);

        for (intLoop = 0; intLoop <= (cboIn.ListCount - 1); intLoop++) {
          if (cboIn.Text == cboIn.List(intLoop)) {
            Debug.Print("setting .ListIndex to "+ ((Integer) intLoop).toString());
            cboIn.ListIndex = intLoop;
          }
        }

        cboIn.SelStart = fnFindString.length();
        cboIn.SelLength = cboIn.Text.length() - cboIn.SelStart;
        intKeyAscii = 0;
      } 
      else {
        Debug.Print("didn't find it...");
        Debug.Print("ListIndex left at "+ cboIn.ListIndex);
      }
      Debug.Print(" - - - - ");
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



////////////////////////////////////////////////////////////////////////////////
  public static void fnInitializefpCombo(LPLib.fpCombo lpcIn, boolean bShowColHeaders, boolean bSortable, int lngNbrOfCols, int lngEditCol, int lngNbrOfRowsInDropdown) { // TODO: Use of ByRef founded Public Sub fnInitializefpCombo(ByRef lpcIn As LPLib.fpCombo, Optional ByVal bShowColHeaders As Boolean = False, Optional ByVal bSortable As Boolean = True, Optional ByVal lngNbrOfCols As Long = 1, Optional ByVal lngEditCol As Long = 0, Optional ByVal lngNbrOfRowsInDropdown As Long = 8)
    // Comments  : Sets default properties of an fpCombo control used as
    //             a multi-column combobox.
    // Parameters:
    //     lpcIn (in/out)              - fpCombo box to configure
    //     bShowColHeaders (in)        - If True, sets up column headers
    //     bSortable (in)              - If True, configures for case-insensitive sort
    //     lngNbrOfCols (in)           - Number of columns to have in the combobox
    //     lngEditCol (in)             - Col to display in edit box (0-based)
    //     lngNbrOfRowsInDropdown (in) - Nbr of rows to show in dropdown
    //
    // Returns   : N/A
    //
    "fnInitializefpCombo"
.equals(Const cstrCurrentProc As String);

    try {

      // Protect against an intColToDisplay that references a non-existent column
      if (lngEditCol > lngNbrOfCols) {
        Debug.Print("Programmer Error: Invalid intColToDisplay passed to "+ MCSTRNAME+ cstrCurrentProc);
        lngEditCol = 0;
      }

      lpcIn.Clear;
      lpcIn.Row = GCLNGNOSELECTION;
      lpcIn.Style = StyleDropDownList;
      lpcIn.Columns = lngNbrOfCols;
      lpcIn.ColumnEdit = lngEditCol;
      // Set this so width of columns can be controlled when control is initialized
      lpcIn.ColumnWidthScale = ColumnWidthScaleAvgCharWidth;

      // Colors
      lpcIn.BackColor = vbWindowBackground;
      lpcIn.ForeColor = vbWindowText;

      // Appearance
      lpcIn.ListApplyTo = ListApplyToAllCols;
      lpcIn.LineStyle = LineStyleNone;
      lpcIn.LineApplyTo = LineApplyToCols;
      lpcIn.Appearance = Appearance3D;
      lpcIn.MaxDrop = lngNbrOfRowsInDropdown;
      lpcIn.ListWidth = GCLNGNOSELECTION;
      //' DataAutoSizeColsBestGuess DataAutoSizeColsMaxColWidth
      lpcIn.DataAutoSizeCols = DataAutoSizeColsBestGuess;
      //' Don't resize to show an entire row at bottom
      lpcIn.NoIntegralHeight = true;
      //' Show horizontal scrollbar only when needed
      lpcIn.ScrollBarH = ScrollBarHShowWhenNeeded;
      //' Show vertical scrollbar only when needed
      lpcIn.ScrollBarV = ScrollBarVShowWhenNeeded;

      // Search behavior
      lpcIn.AutoSearch = AutoSearchMultipleChar;
      lpcIn.AutoSearchFill = true;
      //' Delay in milliseconds (default = 500)
      lpcIn.AutoSearchFillDelay = 200;
      lpcIn.SearchIgnoreCase = true;

      // If requested, set up column headers to be displayed
      if (bShowColHeaders) {
        lpcIn.ListApplyTo = ListApplyToColHeaders;
        lpcIn.ColumnHeaderShow = true;
        lpcIn.AlignH = AlignHCenter;
        lpcIn.LineStyle = LineStyleLoweredwLine;
        lpcIn.BackColor = vbButtonFace;
      }

      // If requested, configure for case-insensitive ascending sort
      if (bSortable) {
        lpcIn.ColSortDataType = ColSortDataTypeTextNoCase;
        lpcIn.Sorted = SortedAscending;
        lpcIn.SortState = SortStateActiveReSort;
      }

      // Column definitions
      lpcIn.ListApplyTo = ListApplyToIndividual;
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


//////////////////////////////////////////////////////////////////////////////////////////////////
  public static boolean fnIsComboBoxDropped(ComboBox cboIn) { // TODO: Use of ByRef founded Public Function fnIsComboBoxDropped(ByRef cboIn As ComboBox) As Boolean
    boolean _rtn = false;
    // Comments  : Determines if the list portion of a combo box is
    //             is currently visible
    // Parameters: cboIn - combo box to check
    // Returns   : True if list portion is visible; false if not
    // Source    : Total Visual SourceBook 2000
    //
    "fnIsComboBoxDropped"
.equals(Const cstrCurrentProc As String);
    int lngResult = 0;

    try {

      lngResult = SendMessage(cboIn.hWnd, CB_GETDROPPEDSTATE, 0&, ByVal 0&);
      _rtn = (lngResult != 0);
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
  public static int fnSearchCBOItemData(ComboBox cboIn, String strSearchIn) {
    int _rtn = 0;
    // Comments  : Does a case-insensitive search for a certain value
    //             in the ItemData property of a ComboBox
    // Parameters: cboIn       (in) - Combo box to search.
    //             strSearchIn (in) - String to search for.
    // Returns   : A value of 0 or greater indicates the line on
    //             which the string was found in the combo box. A
    //             negative value indicates that the line was not found
    //
    "fnSearchCBOItemData"
.equals(Const cstrCurrentProc As String);
    Const(clngNotFound As Long == -1);
    int intFoundEntry = 0;
    boolean bFound = false;
    int intCurrentEntry = 0;
    int intLastEntry = 0;

    try {

      intFoundEntry = clngNotFound;

      bFound = false;
      intCurrentEntry = 0;
      intLastEntry = cboIn.ListCount - 1;

      strSearchIn = strSearchIn.toUpperCase();

      while ((Not bFound) && (intCurrentEntry <= intLastEntry)) {
        if (strSearchIn.equals(cboIn.ItemData(intCurrentEntry).toUpperCase())) {
          bFound = true;
          intFoundEntry = intCurrentEntry;
        }
        intCurrentEntry = intCurrentEntry + 1;
      }

      // Return line found, or -1 if not found
      _rtn = intFoundEntry;
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



////////////////////////////////////////////////////////////////////////////////
  public static void fnSearchFPCombo(LPLib.fpCombo lpcIn, String strSearchText, int intSearchCol, LPLib.SearchMethodConstants lngSearchMethod, boolean bDefaultToFirstRowIfNotFound) { // TODO: Use of ByRef founded Public Sub fnSearchFPCombo(ByRef lpcIn As LPLib.fpCombo, ByVal strSearchText As String, Optional ByVal intSearchCol As Integer = 0, Optional ByVal lngSearchMethod As LPLib.SearchMethodConstants = SearchMethodPartialMatch, Optional ByVal bDefaultToFirstRowIfNotFound As Boolean = True)
    // Comments  : Searches an fpCombo control for the specified value in the specified
    //             column, using the specified search method.
    // Parameters: lpcIn (in/out)                    - fpCombo box to configure
    //             strSearchText (in)                - The text to search for
    //             intSearchCol (in)                 - The column in which to search for strSearchText
    //             lngSearchMethod (in)              - How to search, e.g., exact vs. partial match
    //                                                 (Partial works best!)
    //             bDefaultToFirstRowIfNotFound (in) - Only used if the search was unsuccessful.
    //                                                 * If True, selects the first row (Use this
    //                                                   if the fpCombo control has a "blank" entry as its
    //                                                   first row to denote "no selection"
    //                                                 * If False, it returns -1 (no selection).
    // Returns:    N/A
    "fnSearchFPCombo"
.equals(Const cstrCurrentProc As String);
    Const(clngFirstRow As Long == 0);
    try {

      // Clear the current selection
      lpcIn.Row = GCLNGNOSELECTION;
      // Get the search string
      lpcIn.SearchText = strSearchText;
      // Search the specified column, for partial matches.
      lpcIn.ColumnSearch = intSearchCol;
      lpcIn.SearchMethod = SearchMethodPartialMatch;
      // Set the SearchIndex to reflect a "unsuccessful search" default value
      lpcIn.SearchIndex = GCLNGNOSELECTION;
      lpcIn.Action = ActionSearch;

      // If a match is found, scroll to and select the item; Otherwise scroll to and
      // select the first row (the blank entry)
      if (lpcIn.SearchIndex != GCLNGNOSELECTION) {
        lpcIn.Row = lpcIn.SearchIndex;
        lpcIn.ListIndex = lpcIn.SearchIndex;
      } 
      else {
        lpcIn.Action = ActionClearSearchBuffer;
        if (bDefaultToFirstRowIfNotFound) {
          lpcIn.Row = clngFirstRow;
          lpcIn.ListIndex = clngFirstRow;
        } 
        else {
          lpcIn.Row = GCLNGNOSELECTION;
          lpcIn.ListIndex = GCLNGNOSELECTION;
        }
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


//////////////////////////////////////////////////////////////////////////////////////////////////
  public static void fnSetComboBoxItemHeight(ComboBox cboIn, float sngMultipleItemHeight) {
    // Comments  : Set the height of items in a combo box
    // Parameters: cboIn - ComboBox to modify
    //             sngMultipleItemHeight - multiple of the current height of
    //             an item in the combo box. For example, a value of 2 would
    //             double the height of combo box items
    // Returns   : Nothing
    // Source    : Total Visual SourceBook 2000
    //
    "fnSetComboBoxItemHeight"
.equals(Const cstrCurrentProc As String);
    int lngReturn = 0;
    int lngCurHeight = 0;
    int lngPixels = 0;

    try {

      // Get the current height of a standard item in the combo box
      lngCurHeight = SendMessage(cboIn.hWnd, CB_GETITEMHEIGHT, 0&, ByVal 0&);

      // Multiply this by the new value multiplier
      lngPixels = (lngCurHeight * sngMultipleItemHeight);

      // Tell Windows to change the item height to the new value
      lngReturn = SendMessage(cboIn.hWnd, CB_SETITEMHEIGHT, 0&, ByVal lngPixels);

      // Repaint the combo to show the new values
      cboIn.Refresh;
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



//////////////////////////////////////////////////////////////////////////////////////////////////
  public static void fnSetComboBoxListItems(ComboBox cboIn, int intItems) { // TODO: Use of ByRef founded Public Sub fnSetComboBoxListItems(ByRef cboIn As ComboBox, ByVal intItems As Integer)
    // Comments  : Sets the number of items in the combo box list
    // Parameters: cboIn -  the combo box to modify
    //             intItems -  the number of items
    // Returns   : Nothing
    // Source    : Total Visual SourceBook 2000
    //
    "fnSetComboBoxListItems"
.equals(Const cstrCurrentProc As String);
    POINTAPI ptCoords = null;
    RECT rectScreenCoords = null;
    int lngNewHeight = 0;
    int lngCurHeight = 0;
    int intTPPX = 0;
    int intTPPY = 0;
    int intParentScaleMode = 0;

    try {

      //Saves the ScaleMode of Parent Object
      intParentScaleMode = cboIn.Parent.ScaleMode;
      cboIn.Parent.ScaleMode = vbTwips;

      intTPPX = Screen.TwipsPerPixelX;
      intTPPY = Screen.TwipsPerPixelY;

      // get current item height
      lngCurHeight = SendMessage(cboIn.hWnd, CB_GETITEMHEIGHT, 0&, ByVal 0&);

      // calculate new height
      lngNewHeight = (lngCurHeight + 1) * (intItems + 1);

      // get the coordinates of the combo box on the screen
      GetWindowRect(cboIn.hWnd, rectScreenCoords);

      // fill pt struct
      ptCoords.x = rectScreenCoords.left;
      ptCoords.y = rectScreenCoords.top;

      // get the coordinates of the combo box on the form
      ScreenToClient(cboIn.Parent.cbrfBrowseFolder.setHWnd(), ptCoords);

      // resize the combo box
      MoveWindow(cboIn.hWnd, ptCoords.x, ptCoords.y, cboIn.Width \ intTPPX, lngNewHeight, -1);

      // Resets the Parent object's ScaleMode
      cboIn.Parent.ScaleMode = intParentScaleMode;
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




//////////////////////////////////////////////////////////////////////////////////////////////////
  public static void fnSetComboBoxListWidth(ComboBox cboIn, float sngMultipleListWidth) { // TODO: Use of ByRef founded Public Sub fnSetComboBoxListWidth(ByRef cboIn As ComboBox, ByVal sngMultipleListWidth As Single)
    // Comments  : Set the width of the drop-down list portion of a combo box
    // Parameters: cboIn - the combo box to modify
    //             sngMultipleListWidth - a multiple of the current with of
    //             the actual combo box
    // Returns   : Nothing
    // Source    : Total Visual SourceBook 2000
    //
    "fnSetComboBoxListWidth"
.equals(Const cstrCurrentProc As String);
    int lngReturn = 0;
    int lngCurWidth = 0;
    int lngPixels = 0;

    try {

      // Get the current width of the combo box list
      lngCurWidth = SendMessage(cboIn.hWnd, CB_GETHORIZONTALEXTENT, 0&, ByVal 0&);

      // Calculate the new width
      if (lngCurWidth == 0) {
        lngPixels = (cboIn.Width \ Screen.TwipsPerPixelX) * sngMultipleListWidth;
      } 
      else {
        lngPixels = (lngCurWidth * sngMultipleListWidth);
      }

      // Tell windows the new width of the combo box list
      lngReturn = SendMessage(cboIn.hWnd, CB_SETDROPPEDWIDTH, ByVal lngPixels, ByVal 0&);

      cboIn.Refresh;
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

private class POINTAPI {
    public Long x;
    public Long y;
}


private class RECT {
    public Long left;
    public Long top;
    public Long right;
    public Long bottom;
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


case class OdcomboboxData(
              id: Option[Int],

              )

object Odcomboboxs extends Controller with ProvidesUser {

  val odcomboboxForm = Form(
    mapping(
      "id" -> optional(number),

  )(OdcomboboxData.apply)(OdcomboboxData.unapply))

  implicit val odcomboboxWrites = new Writes[Odcombobox] {
    def writes(odcombobox: Odcombobox) = Json.obj(
      "id" -> Json.toJson(odcombobox.id),
      C.ID -> Json.toJson(odcombobox.id),

    )
  }

  def get(id: Int) = GetAction { implicit request =>
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.LIST_ODCOMBOBOX), { user =>
      Ok(Json.toJson(Odcombobox.get(user, id)))
    })
  }

  def update(id: Int) = PostAction { implicit request =>
    Logger.debug("in odcomboboxs.update")
    odcomboboxForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      odcombobox => {
        Logger.debug(s"form: ${odcombobox.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.EDIT_ODCOMBOBOX), { user =>
          Ok(
            Json.toJson(
              Odcombobox.update(user,
                Odcombobox(
                       id,

                ))))
        })
      }
    )
  }

  def create = PostAction { implicit request =>
    Logger.debug("in odcomboboxs.create")
    odcomboboxForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      odcombobox => {
        Logger.debug(s"form: ${odcombobox.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.NEW_ODCOMBOBOX), { user =>
          Ok(
            Json.toJson(
              Odcombobox.create(user,
                Odcombobox(

                ))))
        })
      }
    )
  }

  def delete(id: Int) = PostAction { implicit request =>
    Logger.debug("in odcomboboxs.delete")
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.DELETE_ODCOMBOBOX), { user =>
      Odcombobox.delete(user, id)
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

case class Odcombobox(
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

object Odcombobox {

  lazy val emptyOdcombobox = Odcombobox(
)

  def apply(
      id: Int,
) = {

    new Odcombobox(
      id,
)
  }

  def apply(
) = {

    new Odcombobox(
)
  }

  private val odcomboboxParser: RowParser[Odcombobox] = {
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
        Odcombobox(
              id,
,
              createdAt,
              updatedAt,
              updatedBy)
    }
  }

  def create(user: CompanyUser, odcombobox: Odcombobox): Odcombobox = {
    save(user, odcombobox, true)
  }

  def update(user: CompanyUser, odcombobox: Odcombobox): Odcombobox = {
    save(user, odcombobox, false)
  }

  private def save(user: CompanyUser, odcombobox: Odcombobox, isNew: Boolean): Odcombobox = {
    def getFields = {
      List(

      )
    }
    def throwException = {
      throw new RuntimeException(s"Error when saving ${C.ODCOMBOBOX}")
    }

    DBHelper.saveEx(
      user,
      Register(
        C.ODCOMBOBOX,
        C.ID,
        odcombobox.id,
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

  def load(user: CompanyUser, id: Int): Option[Odcombobox] = {
    loadWhere(user, s"${C.ID} = {id}", 'id -> id)
  }

  def loadWhere(user: CompanyUser, where: String, args : scala.Tuple2[scala.Any, anorm.ParameterValue[_]]*) = {
    DB.withConnection(user.database.database) { implicit connection =>
      SQL(s"SELECT t1.*, t2.${C.FK_NAME} FROM ${C.ODCOMBOBOX} t1 INNER JOIN ${C.???} t2 ON t1.${C.FK_ID} = t2.${C.FK_ID} WHERE $where")
        .on(args: _*)
        .as(odcomboboxParser.singleOpt)
    }
  }

  def delete(user: CompanyUser, id: Int) = {
    DB.withConnection(user.database.database) { implicit connection =>
      try {
        SQL(s"DELETE FROM ${C.ODCOMBOBOX} WHERE ${C.ID} = {id}")
        .on('id -> id)
        .executeUpdate
      } catch {
        case NonFatal(e) => {
          Logger.error(s"can't delete a ${C.ODCOMBOBOX}. ${C.ID} id: $id. Error ${e.toString}")
          throw e
        }
      }
    }
  }

  def get(user: CompanyUser, id: Int): Odcombobox = {
    load(user, id) match {
      case Some(p) => p
      case None => emptyOdcombobox
    }
  }
}


// Router

GET     /api/v1/general/odcombobox/:id              controllers.logged.modules.general.Odcomboboxs.get(id: Int)
POST    /api/v1/general/odcombobox                  controllers.logged.modules.general.Odcomboboxs.create
PUT     /api/v1/general/odcombobox/:id              controllers.logged.modules.general.Odcomboboxs.update(id: Int)
DELETE  /api/v1/general/odcombobox/:id              controllers.logged.modules.general.Odcomboboxs.delete(id: Int)




/**/
