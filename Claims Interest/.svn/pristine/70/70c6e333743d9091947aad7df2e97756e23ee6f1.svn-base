Attribute VB_Name = "modComboBox"
'******************************************************************************
' Module      : modComboBox
' Procedures  :
'               fnAddItemDataComboBox( cboIn As ComboBox, ByVal strNewString As String, _
'                   Optional ByVal lngItemData As Long) As Long
'               fnADORecordSetToComboBox( rstIn As ADODB.Recordset, cboIn As ComboBox, _
'                   ByVal strDisplayColumn As String, Optional ByVal varItemDataColumn As Variant)
'               fnADORecordsetToVFG(ByVal rstIn As ADODB.Recordset, _
'                   ByRef pvfgIn As VSFlexGrid, Optional ByVal intGridCol As Integer = 0, _
'                   Optional ByVal varFieldNamesIn As Variant)
'               fnBinarySearchComboBox( cboIn As ComboBox, ByVal strSearch As String) _
'                   As Integer
'               fnDAORecordSetToComboBox( rstIn As DAO.Recordset, cboIn As ComboBox, _
'                   ByVal strDisplayColumn As String, Optional ByVal varItemDataColumn As Variant)
'               fnDropComboBoxList( cboIn As ComboBox, ByVal fShow As Boolean)
'               fnFindStringComboBox(ByRef cboIn As ComboBox, ByVal strSearchIn As String, _
'                   Optional ByVal bDoExactSearch As Boolean = False) As Long
'               fnIncrementalSearchCombo( cboIn As ComboBox, intKeyAscii As Integer)
'               fnInitializefpCombo(ByRef lpcIn As LPADOLib.fpCombo, ByVal intNbrOfCols As Integer, _
'                   Optional ByVal intColToDisplay As Integer = 0)
'               fnIsComboBoxDropped(cboIn) As Boolean
'               fnSearchCBOItemData(ByVal cboIn As ComboBox, ByVal strSearchIn As String) _
'                   As Integer
'               fnSearchFPCombo(ByRef lpcIn As LPLib.fpCombo, ByVal strSearchText As String, _
'                   Optional ByVal intSearchCol As Integer = 0, _
'                   Optional ByVal lngSearchMethod As lplib.SearchMethodConstants = SearchMethodPartialMatch, _
'                   Optional ByVal bDefaultToFirstRowIfNotFound As Boolean = True)
'               fnSetComboBoxItemHeight( cboIn As ComboBox, ByVal sngMultipleItemHeight As Single)
'               fnSetComboBoxListItems( cboIn As ComboBox, ByVal intItems As Integer)
'               fnSetComboBoxListWidth( cboIn As ComboBox, ByVal sngMultipleListWidth As Single)
' Description : Routines to extend the functionality of a standard
'               VB ComboBox control
' Source      : Total Visual SourceBook 2000
'
' Modified:
'  01/2002 BAW  Updated fnADORecordsetToComboBox to make it more flexible.
' --------------------------------------------------
Option Explicit
Option Compare Binary
Private Const mcstrName As String = "modComboBox."

Public Const gclngNoSelection As Long = -1

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) _
    As Long
Private Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As String) _
    As Long
Private Declare Function GetWindowRect Lib "user32" _
    (ByVal hWnd As Long, lpRect As RECT) _
    As Long
Private Declare Function ScreenToClient Lib "user32" _
    (ByVal hWnd As Long, lpPoint As POINTAPI) _
    As Long
Private Declare Function MoveWindow Lib "user32" _
    (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, _
    ByVal nHeight As Long, ByVal bRepaint As Long) _
    As Long

Private Const CB_ADDSTRING = &H143
Private Const CB_DELETESTRING = &H144
Private Const CB_DIR = &H145
Private Const CB_ERR = (-1)
Private Const CB_ERRSPACE = (-2)
Private Const CB_FINDSTRING = &H14C
Private Const CB_FINDSTRINGEXACT = &H158
Private Const CB_GETCOUNT = &H146
Private Const CB_GETCURSEL = &H147
Private Const CB_GETDROPPEDCONTROLRECT = &H152
Private Const CB_GETDROPPEDSTATE = &H157
Private Const CB_GETEDITSEL = &H140
Private Const CB_GETEXTENDEDUI = &H156
Private Const CB_GETITEMDATA = &H150
Private Const CB_GETITEMHEIGHT = &H154
Private Const CB_GETLBTEXT = &H148
Private Const CB_GETLBTEXTLEN = &H149
Private Const CB_GETLOCALE = &H15A
Private Const CB_INSERTSTRING = &H14A
Private Const CB_LIMITTEXT = &H141
Private Const CB_MSGMAX = &H15B
Private Const CB_OKAY = 0
Private Const CB_RESETCONTENT = &H14B
Private Const CB_SELECTSTRING = &H14D
Private Const CB_SETCURSEL = &H14E
Private Const CB_SETEDITSEL = &H142
Private Const CB_SETEXTENDEDUI = &H155
Private Const CB_SETITEMDATA = &H151
Private Const CB_SETITEMHEIGHT = &H153
Private Const CB_SETLOCALE = &H159
Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_GETHORIZONTALEXTENT = &H15D
Private Const CB_SETHORIZONTALEXTENT = &H15E
Private Const CB_SETDROPPEDWIDTH = &H160

Private Const CBS_AUTOHSCROLL = &H40&
Private Const CBS_DISABLENOSCROLL = &H800&

Private Const WM_SETREDRAW = &HB

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function fnAddItemDataComboBox(ByRef cboIn As ComboBox, ByVal strNewString As String, _
    Optional ByVal lngItemData As Long) As Long
    ' Comments  : Adds an item to a combobox, and the
    '             itemdata value associated with that item,
    '             in a single function call
    ' Parameters: cboIn - ComboBox control to add to
    '             strNewString - string to add to list box
    '             lngItemData - value to add to associated
    '             ItemData() entry
    ' Returns   : the ListIndex value of the newly-added item
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "fnAddItemDataComboBox"
    On Error GoTo PROC_ERR

    With cboIn
        .AddItem strNewString
        .ItemData(.NewIndex) = lngItemData
        fnAddItemDataComboBox = .NewIndex
    End With
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub fnADORecordSetToComboBox(ByRef rstIn As ADODB.Recordset, ByRef cboIn As ComboBox, _
    ByVal strDisplayColumn As String, Optional ByVal varItemDataColumn As Variant, _
    Optional ByVal bClear As Boolean = True, Optional ByVal bAddAllEntry As Boolean = False, _
    Optional ByVal bAddNullEntry As Boolean = False, Optional ByVal bAddBlankEntry As Boolean = False)
    ' Comments  : Displays the contents of an ADO recordset in
    '             a standard unbound combo box
    ' Parameters:
    '    rstIn             - Recordset to read. Caller must create
    '    cboIn             - Combo box to load
    '    strDisplayColumn  - name of the column in rstIn to display
    '                        in the combo box
    '    varItemDataColumn - name of the column in rstIn to load into the
    '                        ItemData property of the combo box. The data in
    '                        this column MUST be storable in a long integer, and there
    '                        must be no 'null' values. Generally this field will be a
    '                        long integer Primary Key value associated with the value
    '                        to be displayed in the list
    '    bAddBlankEntry    - add a blank entry first?
    '    bAddAllEntry      - add an "--All--" item first?
    '    bAddNullEntry     - add an "<NULL> item next?
    ' Returns   : Nothing
    ' Source    : Total Visual SourceBook 2000
    ' Modified  :
    '  01/2002 BAW Added bClear optional parameter, so this procedure could be
    '  used even by those comboboxes that need a blank or "**" entry added at
    '  the beginning (before the recordset is loaded).
    '
    '  05/10/02 CMP - added bAddAllItem in.
    '
    Const cstrCurrentProc As String = "fnADORecordSetToComboBox"
    Dim strIDC As String
    Dim bSplashLoaded As Boolean

    On Error GoTo PROC_ERR

    ' if a column name is supplied in the varItemDataColumn parameter,
    ' use this as the field name in the recordset to use to supply values
    ' as the ItemData property of the list array
    If Not IsMissing(varItemDataColumn) Then
        strIDC = CStr(varItemDataColumn)
    Else
        strIDC = vbNullString
    End If

    If bClear Then
        cboIn.Clear
    End If

    If bAddBlankEntry Then
       cboIn.AddItem gcstrBlankEntry
    End If

    If bAddAllEntry Then
       cboIn.AddItem gcstrAllEntry
    End If
    
    If bAddNullEntry Then
       cboIn.AddItem gcstrNullEntry
    End If
    
    bSplashLoaded = fnIsFormLoaded("frmSplash")

    ' Load the named column in each row of the recordset into the combo box
    ' If specified, load the value from the selected column into the
    ' ItemData property. Use the original sort order, and all existing
    ' rows of the recordset that is passed.
        
    With rstIn
        If Not (.BOF And .EOF) Then
            .MoveFirst
            Do Until .EOF
                If Not (IsNull(rstIn(strDisplayColumn))) Then   'ECG 5/16/2002 added this for nullable CBOs inwhich all returned DB records are NULL
                    cboIn.AddItem rstIn(strDisplayColumn)
                    If strIDC <> vbNullString Then
                        cboIn.ItemData(cboIn.NewIndex) = rstIn(strIDC)
                    End If
                    If bSplashLoaded Then
                         DoEvents   ' Allow progress meter to get updated!
                    End If
                End If                                          'ECG 5/16/2002 added this for nullable CBOs inwhich all returned DB records are NULL
                .MoveNext
            Loop
        End If
    End With
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



#If False Then
'////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub fnADORecordsetToVFG(ByVal rstIn As ADODB.Recordset, _
                               ByRef pvfgIn As VSFlexGrid, _
                               Optional ByVal intGridCol As Integer = 0, _
                               Optional ByVal varFieldNamesIn As Variant)
    ' Comments  : Populates a VSFlexgrid column with the complete contents of an ADO recordset.
    ' Parameters: rstIn  (in)          - Recordset to read. Caller must create
    '             pvfgIn (in/out)      - VSFlexGrid to load
    '             intGridCol (in)      - the column to hide
    '             varFieldNamesIn (in) - when only selective columns of the recordset
    '                                    should be used to populate the grid. Caller must set.
    ' Returns   : Nothing
    ' Modified  :
    '
    Const cstrCurrentProc   As String = "fnADORecordsetToVFG"
    Dim strComboBoxString   As String
    Dim strFieldNamesString As String
    Dim fldName             As Field
    
    On Error GoTo PROC_ERR

    If IsMissing(varFieldNamesIn) Then
        ' Collect all of our field names - the BuildComboList needs them.
        strFieldNamesString = vbNullString
        For Each fldName In rstIn.Fields
            strFieldNamesString = strFieldNamesString & fldName.Name & ","
        Next fldName
    
        ' Chop the last "," off
        strFieldNamesString = Left$(strFieldNamesString, Len(strFieldNamesString) - 1)
    Else
        strFieldNamesString = varFieldNamesIn
    End If
    
    ' Build the string
    strComboBoxString = pvfgIn.BuildComboList(rstIn, strFieldNamesString)
    
    pvfgIn.Select 0, 0
    
    ' Populate the ComboBox and indicate that the dropdown arrow should always be visible.
    pvfgIn.ColComboList(intGridCol) = strComboBoxString
    pvfgIn.ShowComboButton = flexSBAlways
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    
    ' Clean-up statements go here
    fnFreeObject fldName
    
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub
#End If




'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function fnBinarySearchComboBox(cboIn As ComboBox, ByVal strSearch As String) _
    As Integer
    ' Comments  : Use a binary search to find a value in a combo box
    ' Parameters: cboIn - combo box to search. combo box must have
    '             the .Sorted property set to true
    '             strSearch - string to search for.
    ' Returns   : A value of 0 or greater indicates the line on
    '             which the string was found in the combo box. A
    '             negative value indicates that the line was not found
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "fnBinarySearchComboBox"
    Dim intLineNum As Integer
    Dim intNumRecs As Integer
    Dim fFound As Boolean
    Dim intLowNum As Integer
    Dim intHighNum As Integer
    Dim intMidNum As Integer

    On Error GoTo PROC_ERR

    ' Negative number indicates string not found
    intLineNum = -1

    intNumRecs = cboIn.ListCount
    fFound = False
    intLowNum = 0
    intHighNum = intNumRecs - 1

    Do
        intMidNum = (intLowNum + intHighNum) \ 2
        ' check first half of list. If found search the bottom half
        If UCase$(strSearch) < UCase$(cboIn.List(intMidNum)) Then
            intHighNum = intMidNum - 1
        ' check the last half of the list. If found search the top half
        ElseIf UCase$(strSearch) > UCase$(cboIn.List(intMidNum)) Then
            intLowNum = intMidNum + 1
        Else
            ' value found
            fFound = True
            intLineNum = intMidNum
        End If
    Loop Until fFound Or (intHighNum < intLowNum)

    ' return line found, or -1 if not found
    fnBinarySearchComboBox = intLineNum
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
' NOTE: The following requires a reference to DAO in order to compile cleanly
'
' Public Sub fnDAORecordSetToComboBox( rstIn As DAO.Recordset, cboIn As ComboBox, _
'     ByVal strDisplayColumn As String, Optional ByVal varItemDataColumn As Variant)
'     ' Comments  : Displays the contents of a recordset in
'     '             a standard unbound combo box
'     ' Parameters: rstIn - recordset to read. Caller must create
'     '             cboIn - combo box to load
'     '             strDisplayColumn - name of the column in rstIn to display
'     '             in the combo box
'     '             varItemDataColumn - name of the column in rstIn to load
'     '             into the ItemData property of the combo box. The data in
'     '             this column MUST be storable in a long integer, and there
'     '             must be no 'null' values. Generally this field will be a
'     '             long integer Primary Key value associated with the value
'     '             to be displayed in the list
'     ' Returns   : Nothing
'     ' Source    : Total Visual SourceBook 2000
'     '
'     Const cstrCurrentProc As String = "fnDAORecordSetToComboBox"
'     Dim strIDC As String
'
'     On Error GoTo PROC_ERR
'
'     ' if a column name is supplied in the varItemDataColumn parameter,
'     ' use this as the field name in the recordset to use to supply values
'     ' as the ItemData property of the list array
'     If Not IsMissing(varItemDataColumn) Then
'         strIDC = CStr(varItemDataColumn)
'     Else
'         strIDC = vbNullString
'     End If
'
'     cboIn.Clear
'
'     ' load the named column in each row of the recordset into the combo box
'     ' If specified, load the value from the selected column into the
'     ' ItemData property. Use the original sort order, and all existing
'     ' rows of the recordset that is passed.
'     With rstIn
'         If .RecordCount <> 0 Then
'             .MoveFirst
'             Do Until .EOF
'                 cboIn.AddItem rstIn(strDisplayColumn)
'                 If strIDC <> vbNullString Then
'                     cboIn.ItemData(cboIn.NewIndex) = rstIn(strIDC)
'                 End If
'                 .MoveNext
'             Loop
'         End If
'
'     End With
'PROC_EXIT:
'    On Error GoTo 0     ' disable error handler
'    ' Clean-up statements go here
'    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
'        gerhApp.PropagateError mcstrName & cstrCurrentProc
'    End If
'    Exit Sub
'PROC_ERR:
'    Select Case Err.Number
'        'Case statements for expected errors go here
'        Case Else
'            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
'    End Select
'    Resume PROC_EXIT
' End Sub
'



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub fnDropComboBoxList(ByRef cboIn As ComboBox, ByVal fShow As Boolean)
    ' Comments  : Causes a combo box to show or hide the list portion
    ' Parameters: cboIn - the combo box to modify
    '             fShow - True to show the list, false to hide the list
    ' Returns   : Nothing
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "fnDropComboBoxList"
    Dim lngReturn As Long

    On Error GoTo PROC_ERR

    lngReturn = SendMessage(cboIn.hWnd, _
        CB_SHOWDROPDOWN, fShow, ByVal 0&)
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function fnFindStringComboBox(ByRef cboIn As ComboBox, ByVal strSearchIn As String, _
    Optional ByVal bDoExactSearch As Boolean = False) As Long
    ' Comments  : Finds the line on a combo box containing the search string
    ' Parameters: cboIn          (in/out) - The combo box to search
    '             strSearchIn    (in)     - The search value
    '             bDoExactSearch (in)     - True to find exact matches only;
    '                                       False to find partial matching prefixes
    '
    ' Returns   : -1 if string not found; otherwise the index of the line
    '             containing the string
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "fnFindStringComboBox"
    Dim lngFound As Long

    On Error GoTo PROC_ERR

    If bDoExactSearch Then
        lngFound = SendMessage(cboIn.hWnd, _
            CB_FINDSTRINGEXACT, -1, ByVal strSearchIn)

    Else
        lngFound = SendMessage(cboIn.hWnd, _
            CB_FINDSTRING, -1, ByVal strSearchIn)
    End If

    fnFindStringComboBox = lngFound
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub fnIncrementalSearchCombo(ByRef cboIn As ComboBox, ByRef intKeyAscii As Integer)
    ' Comments  : Allows incremental searching of a combo box.
    '             Call this proc from the KeyUp event of the combo
    ' Parameters: cboIn - the combo box to search
    '             intKeyAscii - the ASCII key value passed to the KeyPress event
    '             of the combo box
    ' Returns   : Nothing
    ' Source    : Total Visual SourceBook 2000
    '
    On Error GoTo PROC_ERR

    Const cstrCurrentProc   As String = "fnIncrementalSearchCombo"
    Dim cb                  As Long
    Dim fnFindString        As String
    Dim intLoop             As Integer

    If intKeyAscii < 32 Or intKeyAscii > 127 Then
        GoTo PROC_EXIT
    End If

    If cboIn.SelLength = 0 Then
        Debug.Print "SelLength = " & cboIn.SelLength
        fnFindString = cboIn.Text & Chr$(intKeyAscii)
    Else
        fnFindString = Left$(cboIn.Text, cboIn.SelStart) & Chr$(intKeyAscii)
    End If
    
    Debug.Print "Looking for: [" & fnFindString & "]"
    cb = SendMessageStr(cboIn.hWnd, CB_FINDSTRING, 1, ByVal fnFindString)

    If cb <> CB_ERR Then
        Debug.Print "Found it..."
        cboIn.Text = cboIn.List(cb)

        For intLoop = 0 To (cboIn.ListCount - 1)
            If cboIn.Text = cboIn.List(intLoop) Then
                Debug.Print "setting .ListIndex to " & intLoop
                cboIn.ListIndex = intLoop
            End If
        Next intLoop

        cboIn.SelStart = Len(fnFindString)
        cboIn.SelLength = Len(cboIn.Text) - cboIn.SelStart
        intKeyAscii = 0
    Else
        Debug.Print "didn't find it..."
        Debug.Print "ListIndex left at " & cboIn.ListIndex
    End If
    Debug.Print " - - - - "
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'//////////////////////////////////////////////////////////////////////////////
Public Sub fnInitializefpCombo(ByRef lpcIn As LPLib.fpCombo, _
    Optional ByVal bShowColHeaders As Boolean = False, _
    Optional ByVal bSortable As Boolean = True, _
    Optional ByVal lngNbrOfCols As Long = 1, _
    Optional ByVal lngEditCol As Long = 0, _
    Optional ByVal lngNbrOfRowsInDropdown As Long = 8)
    ' Comments  : Sets default properties of an fpCombo control used as
    '             a multi-column combobox.
    ' Parameters:
    '     lpcIn (in/out)              - fpCombo box to configure
    '     bShowColHeaders (in)        - If True, sets up column headers
    '     bSortable (in)              - If True, configures for case-insensitive sort
    '     lngNbrOfCols (in)           - Number of columns to have in the combobox
    '     lngEditCol (in)             - Col to display in edit box (0-based)
    '     lngNbrOfRowsInDropdown (in) - Nbr of rows to show in dropdown
    '
    ' Returns   : N/A
    '
    Const cstrCurrentProc As String = "fnInitializefpCombo"
 
    On Error GoTo PROC_ERR

    ' Protect against an intColToDisplay that references a non-existent column
    If lngEditCol > lngNbrOfCols Then
        Debug.Print "Programmer Error: Invalid intColToDisplay passed to " & mcstrName & cstrCurrentProc
        lngEditCol = 0
    End If

    With lpcIn
        .Clear
        .Row = gclngNoSelection
        .Style = StyleDropDownList
        .Columns = lngNbrOfCols
        .ColumnEdit = lngEditCol
        ' Set this so width of columns can be controlled when control is initialized
        .ColumnWidthScale = ColumnWidthScaleAvgCharWidth
        
        ' Colors
        .BackColor = vbWindowBackground
        .ForeColor = vbWindowText
        
        ' Appearance
        .ListApplyTo = ListApplyToAllCols
        .LineStyle = LineStyleNone
        .LineApplyTo = LineApplyToCols
        .Appearance = Appearance3D
        .MaxDrop = lngNbrOfRowsInDropdown
        .ListWidth = gclngNoSelection
        .DataAutoSizeCols = DataAutoSizeColsBestGuess ' DataAutoSizeColsBestGuess DataAutoSizeColsMaxColWidth
        .NoIntegralHeight = True                        ' Don't resize to show an entire row at bottom
        .ScrollBarH = ScrollBarHShowWhenNeeded          ' Show horizontal scrollbar only when needed
        .ScrollBarV = ScrollBarVShowWhenNeeded          ' Show vertical scrollbar only when needed
        
        ' Search behavior
        .AutoSearch = AutoSearchMultipleChar
        .AutoSearchFill = True
        .AutoSearchFillDelay = 200                      ' Delay in milliseconds (default = 500)
        .SearchIgnoreCase = True
        
        ' If requested, set up column headers to be displayed
        If bShowColHeaders Then
            .ListApplyTo = ListApplyToColHeaders
            .ColumnHeaderShow = True
            .AlignH = AlignHCenter
            .LineStyle = LineStyleLoweredwLine
            .BackColor = vbButtonFace
        End If
        
        ' If requested, configure for case-insensitive ascending sort
        If bSortable Then
            .ColSortDataType = ColSortDataTypeTextNoCase
            .Sorted = SortedAscending
            .SortState = SortStateActiveReSort
        End If
            
        ' Column definitions
        .ListApplyTo = ListApplyToIndividual
    End With
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub


'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function fnIsComboBoxDropped(ByRef cboIn As ComboBox) As Boolean
    ' Comments  : Determines if the list portion of a combo box is
    '             is currently visible
    ' Parameters: cboIn - combo box to check
    ' Returns   : True if list portion is visible; false if not
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "fnIsComboBoxDropped"
    Dim lngResult As Long

    On Error GoTo PROC_ERR

    lngResult = SendMessage(cboIn.hWnd, CB_GETDROPPEDSTATE, 0&, ByVal 0&)
    fnIsComboBoxDropped = (lngResult <> 0)
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function fnSearchCBOItemData(ByVal cboIn As ComboBox, ByVal strSearchIn As String) _
    As Integer
    ' Comments  : Does a case-insensitive search for a certain value
    '             in the ItemData property of a ComboBox
    ' Parameters: cboIn       (in) - Combo box to search.
    '             strSearchIn (in) - String to search for.
    ' Returns   : A value of 0 or greater indicates the line on
    '             which the string was found in the combo box. A
    '             negative value indicates that the line was not found
    '
    Const cstrCurrentProc As String = "fnSearchCBOItemData"
    Const clngNotFound    As Long = -1
    Dim intFoundEntry     As Integer
    Dim bFound            As Boolean
    Dim intCurrentEntry   As Integer
    Dim intLastEntry      As Integer

    On Error GoTo PROC_ERR

    intFoundEntry = clngNotFound

    bFound = False
    intCurrentEntry = 0
    intLastEntry = cboIn.ListCount - 1

    strSearchIn = UCase$(strSearchIn)

    Do While (Not bFound) And (intCurrentEntry <= intLastEntry)
        If strSearchIn = UCase$(cboIn.ItemData(intCurrentEntry)) Then
            bFound = True
            intFoundEntry = intCurrentEntry
        End If
        intCurrentEntry = intCurrentEntry + 1
    Loop

    ' Return line found, or -1 if not found
    fnSearchCBOItemData = intFoundEntry
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'//////////////////////////////////////////////////////////////////////////////
Public Sub fnSearchFPCombo(ByRef lpcIn As LPLib.fpCombo, _
    ByVal strSearchText As String, _
    Optional ByVal intSearchCol As Integer = 0, _
    Optional ByVal lngSearchMethod As LPLib.SearchMethodConstants = SearchMethodPartialMatch, _
    Optional ByVal bDefaultToFirstRowIfNotFound As Boolean = True)
    ' Comments  : Searches an fpCombo control for the specified value in the specified
    '             column, using the specified search method.
    ' Parameters: lpcIn (in/out)                    - fpCombo box to configure
    '             strSearchText (in)                - The text to search for
    '             intSearchCol (in)                 - The column in which to search for strSearchText
    '             lngSearchMethod (in)              - How to search, e.g., exact vs. partial match
    '                                                 (Partial works best!)
    '             bDefaultToFirstRowIfNotFound (in) - Only used if the search was unsuccessful.
    '                                                 * If True, selects the first row (Use this
    '                                                   if the fpCombo control has a "blank" entry as its
    '                                                   first row to denote "no selection"
    '                                                 * If False, it returns -1 (no selection).
    ' Returns:    N/A
    Const cstrCurrentProc   As String = "fnSearchFPCombo"
    Const clngFirstRow      As Long = 0
    On Error GoTo PROC_ERR
 
    With lpcIn
        ' Clear the current selection
        .Row = gclngNoSelection
        ' Get the search string
        .SearchText = strSearchText
        ' Search the specified column, for partial matches.
        .ColumnSearch = intSearchCol
        .SearchMethod = SearchMethodPartialMatch
        ' Set the SearchIndex to reflect a "unsuccessful search" default value
        .SearchIndex = gclngNoSelection
        .Action = ActionSearch
        
        ' If a match is found, scroll to and select the item; Otherwise scroll to and
        ' select the first row (the blank entry)
        If .SearchIndex <> gclngNoSelection Then
            .Row = .SearchIndex
            .ListIndex = .SearchIndex
        Else
            .Action = ActionClearSearchBuffer
            If bDefaultToFirstRowIfNotFound Then
                .Row = clngFirstRow
                .ListIndex = clngFirstRow
            Else
                .Row = gclngNoSelection
                .ListIndex = gclngNoSelection
            End If
        End If
    End With
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub


'////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub fnSetComboBoxItemHeight(cboIn As ComboBox, ByVal sngMultipleItemHeight As Single)
    ' Comments  : Set the height of items in a combo box
    ' Parameters: cboIn - ComboBox to modify
    '             sngMultipleItemHeight - multiple of the current height of
    '             an item in the combo box. For example, a value of 2 would
    '             double the height of combo box items
    ' Returns   : Nothing
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "fnSetComboBoxItemHeight"
    Dim lngReturn As Long
    Dim lngCurHeight As Long
    Dim lngPixels As Long

    On Error GoTo PROC_ERR

    ' Get the current height of a standard item in the combo box
    lngCurHeight = SendMessage(cboIn.hWnd, _
        CB_GETITEMHEIGHT, 0&, ByVal 0&)

    ' Multiply this by the new value multiplier
    lngPixels = (lngCurHeight * sngMultipleItemHeight)

    ' Tell Windows to change the item height to the new value
    lngReturn = SendMessage(cboIn.hWnd, _
        CB_SETITEMHEIGHT, 0&, ByVal lngPixels)

    ' Repaint the combo to show the new values
    cboIn.Refresh
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub fnSetComboBoxListItems(ByRef cboIn As ComboBox, ByVal intItems As Integer)
    ' Comments  : Sets the number of items in the combo box list
    ' Parameters: cboIn -  the combo box to modify
    '             intItems -  the number of items
    ' Returns   : Nothing
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "fnSetComboBoxListItems"
    Dim ptCoords As POINTAPI
    Dim rectScreenCoords As RECT
    Dim lngNewHeight As Long
    Dim lngCurHeight As Long
    Dim intTPPX As Integer
    Dim intTPPY As Integer
    Dim intParentScaleMode As Integer

    On Error GoTo PROC_ERR

    'Saves the ScaleMode of Parent Object
    intParentScaleMode = cboIn.Parent.ScaleMode
    cboIn.Parent.ScaleMode = vbTwips

    intTPPX = Screen.TwipsPerPixelX
    intTPPY = Screen.TwipsPerPixelY

    ' get current item height
    lngCurHeight = SendMessage(cboIn.hWnd, _
        CB_GETITEMHEIGHT, 0&, ByVal 0&)

    ' calculate new height
    lngNewHeight = (lngCurHeight + 1) * (intItems + 1)

    ' get the coordinates of the combo box on the screen
    GetWindowRect cboIn.hWnd, rectScreenCoords

    ' fill pt struct
    ptCoords.X = rectScreenCoords.Left
    ptCoords.Y = rectScreenCoords.Top

    ' get the coordinates of the combo box on the form
    ScreenToClient cboIn.Parent.hWnd, ptCoords

    ' resize the combo box
    MoveWindow cboIn.hWnd, _
        ptCoords.X, _
        ptCoords.Y, _
        cboIn.Width \ intTPPX, _
        lngNewHeight, _
        -1

    ' Resets the Parent object's ScaleMode
    cboIn.Parent.ScaleMode = intParentScaleMode
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub




'////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub fnSetComboBoxListWidth(ByRef cboIn As ComboBox, ByVal sngMultipleListWidth As Single)
    ' Comments  : Set the width of the drop-down list portion of a combo box
    ' Parameters: cboIn - the combo box to modify
    '             sngMultipleListWidth - a multiple of the current with of
    '             the actual combo box
    ' Returns   : Nothing
    ' Source    : Total Visual SourceBook 2000
    '
    Const cstrCurrentProc As String = "fnSetComboBoxListWidth"
    Dim lngReturn As Long
    Dim lngCurWidth As Long
    Dim lngPixels As Long

    On Error GoTo PROC_ERR

    ' Get the current width of the combo box list
    lngCurWidth = SendMessage(cboIn.hWnd, _
        CB_GETHORIZONTALEXTENT, 0&, ByVal 0&)

    ' Calculate the new width
    If lngCurWidth = 0 Then
        lngPixels = _
            (cboIn.Width \ Screen.TwipsPerPixelX) * sngMultipleListWidth
    Else
        lngPixels = (lngCurWidth * sngMultipleListWidth)
    End If

    ' Tell windows the new width of the combo box list
    lngReturn = SendMessage(cboIn.hWnd, _
        CB_SETDROPPEDWIDTH, ByVal lngPixels, ByVal 0&)

    cboIn.Refresh
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub
