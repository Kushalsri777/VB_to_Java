VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ctcrtCurrentRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------
' Procedure:   ctcrtCurrentRate
' Description: Provides properties and methods to support the current_rate_t table values
'
'              Do a Find on "!CUSTOMIZE!" to locate places where the table
'              wrapper must be changed to work for a different table.
'
'              NOTE: PUBLIC PROPERTIES corresponding to table columns should be
'                    named so that they reflect the table column name with
'                    underscores eliminated and each word beginning with a
'                    capital letter.
'
'                    The naming convention for MEMBER VARIABLES that store
'                    public properties should be follow this standard:
'                    "m_dddPPPP" where "m_" identifies it as a member variable,
'                    "ddd" indicates the data type (i.e. lng, str, dte) and PPPP
'                    is the name of the public property to which it corresponds.
'
'                 Examples:
'
'                    TableCol         PublicProperty  MemberVariable     Constant
'                    ---------------  --------------  -----------------  ----------------
'                    current rate_MGR_PRV_CD  CurrentRateMgrPrvCd    m_strCurrentRateMgrPrvCd  mcstrCurrentRateMgrPrvCd
'                    current rate_SVS_IND     CurrentRateSvsInd      m_bCurrentRateSvsInd      mcstrCurrentRateSvsInd
'
'
'                 NOTE also that navigation should be done via the **ADO Wrapper's**
'                 navigation methods instead of directly referencing the navigation
'                 methods on a ADODB.Recordset object!
'
' Revisions:   1.0 ECG 01/02/2002 Initial creation.
'              2.0 ECG 01/29/2002 Added Error handling used in Betsy's Prototype.
'              3.0 BAW 05/02/2002 Added additional comments including instructions on
'                                 where this class should be changed to work for a
'                                 different table.
'
'
' Procedures  :
'   Private     Class_Initialize()
'   Private     Class_Terminate()
'   Public      Property Get AllowableCharacters(ByVal strTagIn As String) As String
'   Public      Property Get CurrentLookupRecordNumber() As Long
'   Public      Property Get DecimalPositions(ByVal strTagIn As String) As Integer
'   Public      Property Get DefaultValue(ByVal strTagIn As String) As Variant
'   Public      Property Get DollarPositions(ByVal strTagIn As String) As Integer
'   Public      Property Get Format(ByVal strTagIn As String) As String
'   Public      Property Get CurrentRateCd() As String
'   Public      Property Let CurrentRateCd(ByVal strValue As String)
'   Public      Property Get CurrentRateMgr() As String
'   Public      Property Let CurrentRateMgr(ByVal strValue As String)
'   Public      Property Get CurrentRateMgrPrvCd() As String
'   Public      Property Let CurrentRateMgrPrvCd(ByVal strValue As String)
'   Public      Property Get CurrentRateNm() As String
'   Public      Property Let CurrentRateNm(ByVal strValue As String)
'   Public      Property Get CurrentRateSvsInd() As Boolean
'   Public      Property Let CurrentRateSvsInd(ByVal bValue As Boolean)
'   Public      Property Get IsKey(ByVal strTagIn As String) As Boolean
'   Public      Property Get IsNullable(ByVal strTagIn As String) As Boolean
'   Public      Property Get LookupData() As ADODB.Recordset
'   Public      Property Get LookupIsAtBOF() As Boolean
'   Public      Property Get LookupIsAtEOF() As Boolean
'   Public      Property Get LookupRecordCount() As Long
'   Public      Property Get LstUpdDtm() As Date
'   Public      Property Let LstUpdDtm(ByVal NewValue As String)
'   Public      Property Get LstUpdUserId() As String
'   Public      Property Let LstUpdUserId(ByVal strValue As String)
'   Public      Property Get Mask(ByVal strTagIn As String) As String
'   Public      Property Get MaxCharacters(ByVal strTagIn As String) As Long
'   Public      Property Get MktvalCurrentRateCd() As String
'   Public      Property Let MktvalCurrentRateCd(ByVal strValue As String)
'   Public      Property Get ShouldForceToUppercase(ByVal strTagIn As String) As Boolean
'   Public      AddRecord() as Boolean
'   Public      CheckForAnotherUsersChanges(ByVal lngWhatOperation As enumWhatOperationIsBeingAttempted, _
'                   ByRef strACF2 As String) As Long
'   Public      DeleteRecord() As Boolean
'   Public      GetLookupData() As Boolean
'   Public      GetSingleRecord(ByVal strKey1 As String) As Boolean
'   Public      GoToFirstRecord()
'   Public      GoToLastRecord()
'   Public      GoToNextRecord()
'   Public      GoToPreviousRecord()
'   Public      UpdateRecord() As Boolean
'   Private     fnGetColMetaData(ByRef pudtCol As udtColumn, _
'                   ByRef prstIn As ADODB.Recordset) As Boolean
'   Private     fnGetProperty(ByVal strTagIn As String) As udtColumn
'   Private     GetRelativeRecord(ByVal strKey1 As String, _
'                   ByVal lngPositionDirection As enumPositionDirection) As Boolean
'   Private     fnLoadColMetaData() As Boolean
'   Private     fnSelectRecord(ByVal strKey1 As String) As ADODB.Recordset
'
'-----------------------------------------------------------------------------
Option Explicit
Option Compare Binary

'!CUSTOMIZE! Change both the filename and class name to represent the main table.
'!CUSTOMIZE! Change mcstrName to reflect the class name, followed by a period.
Private Const mcstrName As String = "ctcrtCurrentRate."

'...............................................................................................
'!CUSTOMIZE!
' These are the private variables corresponding to PUBLIC properties.
' There should be one (of type udtColumn) for each column in the table that this class accesses.
'...............................................................................................
Private m_strCurrentRateCd                   As udtColumn
Private m_strCurrentRateNm                   As udtColumn
Private m_dteLstUpdDtm                       As udtColumn
Private m_strLstUpdUserId                    As udtColumn
Private m_bCurrentRateSVSInd                 As udtColumn
Private m_strMktValCurrentRateCd             As udtColumn
Private m_strCurrentRateMgrPrvCd             As udtColumn
Private m_strCurrentRateMgrCurrentRateCd     As udtColumn

'...............................................................................................
'!CUSTOMIZE!
' Create one Const for each column in the table, defining the table column to which it refers.
'...............................................................................................
Private Const mcstrCurrentRateCd             As String = "current rate_CD"
Private Const mcstrCurrentRateNm             As String = "current rate_NM"
Private Const mcstrLstUpdDtm                 As String = "LST_UPD_DTM"
Private Const mcstrLstUpdUserId              As String = "LST_UPD_USER_ID"
Private Const mcstrCurrentRateSvsInd         As String = "current rate_SVS_IND"
Private Const mcstrMktvalCurrentRateCd       As String = "MKTVAL_current rate_CD"
Private Const mcstrCurrentRateMgrPrvCd       As String = "current rate_MGR_PRV_CD"
Private Const mcstrCurrentRateMgrFundCd      As String = "current rate_MGR_current rate_CD"


'...............................................................................................
' Other private variables that do NOT correspond to PUBLIC properties.
'...............................................................................................
' m_adwADO is a private instantiation of the ADO Wrapper, used to do ADO things like
' navigation, executing a stored procedure, etc.
Private m_adwADO                        As cadwADOWrapper

' The next 2 vars (m_dteLstUpdDtm_Original and m_strLstUpdUserId_Original) are used by
' the CheckForAnotherUsersChanges method to determine if another user affected the
' record since *this* user originally retrieved the record.
Private m_dteLstUpdDtm_Original         As Date
Private m_strLstUpdUserId_Original      As String

' m_rstLookup contains selected columns for each row in the table and is used by the form
' to populate its Lookup VSFlexGrid control that the user uses to hop directly to a desired record.
' m_rstLookup should be PRIVATE! If anyone besides this class needs to reference properties of this
' Recordset, then those properties should be exposed as public properties of this class.
Private m_rstLookup                     As ADODB.Recordset


'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'\                                                                  /
'|          CLASS_INITIALIZE / CLASS_TERMINATE    Procedures        |
'/                                                                  \
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/

'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub Class_Initialize()
    ' **************************************************************************
    ' Function  : Class_Initialize
    ' Purpose   : Starting Point for Object
    '             >GetLookupData (Recordset of KEy Columns for every row in table)
    '             >Populate Object Field Properties with Table's First Record
    
    ' Parameters: N/A
    ' Returns   : Boolean
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Class_Initialize"
    On Error GoTo PROC_ERR

    Set m_adwADO = New cadwADOWrapper

    ' Refresh lookup RST and set LookupRecordCount / CurrentLookupRecNbr properties
    GetLookupData
    
    ' Get all columns for the 1st record in the Lookup RST and load to member vars.
    ' If there are no records (m_rstLookup is Nothing), then initialize the
    ' properties that correspond to table columns. (Caller must take action if
    ' m_rstLookup Is Nothing!!!)
    If m_rstLookup.RecordCount <> 0 Then
        '!CUSTOMIZE! The constant referenced below should refer to the key field.
        GetSingleRecord m_rstLookup.Fields(mcstrCurrentRateCd).value
    Else
        fnClearPropertyValues
    End If
        
    ' Obtain meta data about each table column from the DBMS and load it to the
    ' properties that correspond to those table columns
    fnLoadColMetaData
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
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
Private Sub Class_Terminate()
    ' **************************************************************************
    ' Function  : Class_Terminate
    ' Purpose   : Closes the private recordset variable, then frees members
    '             associated with internal objects
    ' Parameters: N/A
    ' Returns   : Int
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Class_Terminate"
    On Error GoTo PROC_ERR

    fnFreeRecordset m_rstLookup
    fnFreeObject m_adwADO
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
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





'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'\                                                                  /
'|                PROPERTY GET/LET    Procedures                    |
'/                                                                  \
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/

'!CUSTOMIZE! so that there is a Property Get and Let for each table column.

'////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Get AllowableCharacters(ByVal strTagIn As String) As String
    ' **************************************************************************
    ' Function  : Property Get AllowableCharacters
    ' Purpose   : Retrieves the default Format Mask (i.e. #####.###) from property
    ' Parameters: N/A
    ' Returns   : String
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get AllowableCharacters"
    On Error GoTo PROC_ERR

    AllowableCharacters = fnGetProperty(strTagIn).AllowableCharacters
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Property
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Property



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Get CurrentLookupRecordNumber() As Long
    '--------------------------------------------------------------------------
    ' Procedure:   Property Get CurrentLookupRecordNumber
    ' Description: Retrieve the record number of the record currently in context
    ' Returns:     record position as Long
    '-----------------------------------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc       As String = "Property Get CurrentLookupRecordNumber"

    CurrentLookupRecordNumber = m_rstLookup.AbsolutePosition
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Property
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Property



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Get DecimalPositions(ByVal strTagIn As String) As Integer
    ' **************************************************************************
    ' Function  : Property Get DecimalPositions
    ' Purpose   : Retrieves the max number of decimal positions from the
    '             named property
    ' Parameters: N/A
    ' Returns   : Integer
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get DecimalPositions"
    On Error GoTo PROC_ERR

    DecimalPositions = fnGetProperty(strTagIn).DecimalPositions
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Property
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Property


'////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Get DefaultValue(ByVal strTagIn As String) As Variant
    ' **************************************************************************
    ' Function  : Property Get DefaultValue
    ' Purpose   : Retrieves the Default Value from the
    '             named property
    '
    '             It's up to the CALLER to see if DefaultValue = Empty and,
    '             if so, not to use the return value. It's also the
    '             caller's responsibility to do any data type conversion
    '             that might be necessary, such as turning a
    '
    ' Parameters: N/A
    ' Returns   : Variant
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get DefaultValue"
    On Error GoTo PROC_ERR
    Dim udtTemp As udtColumn
    Dim varTemp As Variant

    udtTemp = fnGetProperty(strTagIn)    ' for efficiency

    '!TODO! Consider whether another variation of this procedure should
    '       be created...where it accepts a parameter of type Control
    '       which it can then set if the column has a default value
    '       and do nothing to if the column has no default value.

    If udtTemp.HasDefault Then
        varTemp = udtTemp.DefaultValue

        ' Strip leading single quote
        If Left(varTemp, 1) = "'" And Len(varTemp) > 1 Then
            varTemp = Right(varTemp, Len(varTemp) - 1)
        End If
        ' Strip trailing single quote
        If Right(varTemp, 1) = "'" And Len(varTemp) > 1 Then
            varTemp = Left(varTemp, Len(varTemp) - 1)
        End If

        DefaultValue = varTemp
        GoTo PROC_EXIT
    Else
        DefaultValue = Empty
    End If
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Property
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Property



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Get DollarPositions(ByVal strTagIn As String) As Integer
    ' **************************************************************************
    ' Function  : Property Get DollarPositions
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : Integer
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get DollarPositions"
    On Error GoTo PROC_ERR

    DollarPositions = fnGetProperty(strTagIn).DollarPositions
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Property
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Property



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Get Format(ByVal strTagIn As String) As String
    ' **************************************************************************
    ' Function  : Property Get Format
    ' Purpose   : Retrieves the default Format Mask (i.e. #####.###) from property
    ' Parameters: N/A
    ' Returns   : String
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get Format"
    On Error GoTo PROC_ERR

    Format = fnGetProperty(strTagIn).Format
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Property
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Property



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Get CurrentRateCd() As String
    ' **************************************************************************
    ' Function  : Property Get CurrentRateCd
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : string
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get CurrentRateCd"
    On Error GoTo PROC_ERR

    CurrentRateCd = CStr(m_strCurrentRateCd.value)
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Property
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Property



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Let CurrentRateCd(ByVal strValue As String)
    ' **************************************************************************
    ' Function  : Property Let CurrentRateCd
    ' Purpose   : Assigns new Value to property
    ' Parameters: ByVal NewValue As string
    ' Returns   :
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let CurrentRateCd"
    On Error GoTo PROC_ERR

    m_strCurrentRateCd.value = strValue
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Property
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Property



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Get CurrentRateMgr() As String
    ' **************************************************************************
    ' Function  : Property Get CurrentRateMgr
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : Int
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get CurrentRateMgr"
    On Error GoTo PROC_ERR

    CurrentRateMgr = CStr(m_strCurrentRateMgrCurrentRateCd.value)
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Property
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Property



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Let CurrentRateMgr(ByVal strValue As String)
    ' **************************************************************************
    ' Function  : Property Let CurrentRateMgr
    ' Purpose   : Assigns new Value to property
    ' Parameters: ByVal NewValue As Integer
    ' Returns   :
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let CurrentRateMgr"
    On Error GoTo PROC_ERR

    m_strCurrentRateMgrCurrentRateCd.value = strValue
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Property
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Property



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Get CurrentRateMgrPrvCd() As String
    ' **************************************************************************
    ' Function  : Property Get CurrentRateMgrPrvCd
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : string
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get CurrentRateMgrPrvCd"
    On Error GoTo PROC_ERR

    CurrentRateMgrPrvCd = CStr(m_strCurrentRateMgrPrvCd.value)
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Property
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Property



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Let CurrentRateMgrPrvCd(ByVal strValue As String)
    ' **************************************************************************
    ' Function  : Property Let CurrentRateMgrPrvCd
    ' Purpose   : Assigns new Value to property
    ' Parameters: ByVal NewValue As Integer
    ' Returns   :
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let CurrentRateMgrPrvCd"
    On Error GoTo PROC_ERR

    m_strCurrentRateMgrPrvCd.value = strValue
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Property
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Property



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Get CurrentRateNm() As String
    ' **************************************************************************
    ' Function  : Property Get CurrentRateNm
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : Int
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get CurrentRateNm"
    On Error GoTo PROC_ERR

    CurrentRateNm = CStr(m_strCurrentRateNm.value)
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Property
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Property



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Let CurrentRateNm(ByVal strValue As String)
    ' **************************************************************************
    ' Function  : Property Let CurrentRateNm
    ' Purpose   : Assigns new Value to property
    ' Parameters: ByVal NewValue As Integer
    ' Returns   :
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let CurrentRateNm"
    On Error GoTo PROC_ERR

    m_strCurrentRateNm.value = strValue
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Property
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Property



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Get CurrentRateSvsInd() As Boolean
    ' **************************************************************************
    ' Function  : Property Get CurrentRateSvsInd
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : Int
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get CurrentRateSvsInd"
    On Error GoTo PROC_ERR

    ' Note that this field is stored as a Y or N.

    CurrentRateSvsInd = CBool(m_bCurrentRateSVSInd.value)
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Property
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Property



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Let CurrentRateSvsInd(ByVal bValue As Boolean)
    ' **************************************************************************
    ' Function  : Property Let CurrentRateSvsInd
    ' Purpose   : Assigns new Value to property
    ' Parameters: ByVal NewValue As Integer
    ' Returns   :
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let CurrentRateSvsInd"
    On Error GoTo PROC_ERR

    ' Note that this field is stored as a Y or N.

    m_bCurrentRateSVSInd.value = bValue
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Property
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Property



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Get IsKey(ByVal strTagIn As String) As Boolean
    ' **************************************************************************
    ' Function  : Property Get IsKey
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : Int
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get IsKey"
    On Error GoTo PROC_ERR

    IsKey = fnGetProperty(strTagIn).IsKey
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Property
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Property



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Get IsNullable(ByVal strTagIn As String) As Boolean
    ' **************************************************************************
    ' Function  : Property Get IsNullable
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : Int
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get IsNullable"
    On Error GoTo PROC_ERR

    IsNullable = fnGetProperty(strTagIn).IsNullable
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Property
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Property



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Get LookupData() As ADODB.Recordset
    '--------------------------------------------------------------------------
    ' Procedure:   Get_LookupData
    ' Description: Get a copy of the objects Lookup Recordset
    ' Returns:     ADODB.Recordset
    '-----------------------------------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc       As String = "Get_LookupData"

    Set LookupData = m_rstLookup
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Property
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Property



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Get LookupIsAtBOF() As Boolean
    '--------------------------------------------------------------------------
    ' Procedure:   Property Get LookupIsAtBOF
    ' Description: Indicates whether the Lookup recordset is at BOF
    '              (i.e. prior to the last record in the recordset).
    '
    '              Both LookupIsAtBOF() and LookupIsAtEOF() will return True
    '              if there are no records in the m_rstLookup recordset.
    '
    ' Returns:     True if it is at BOF; False otherwise
    '-----------------------------------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc       As String = "Property Get LookupIsAtBOF"

    If Not (m_rstLookup Is Nothing) Then
        LookupIsAtBOF = m_rstLookup.BOF
    Else
        LookupIsAtBOF = True
    End If
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Property
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Property



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Get LookupIsAtEOF() As Boolean
    '--------------------------------------------------------------------------
    ' Procedure:   Property Get LookupIsAtEOF
    ' Description: Indicates whether the Lookup recordset is at EOF
    '              (i.e. beyond the last record in the recordset)
    '
    '              Both LookupIsAtBOF() and LookupIsAtEOF() will return True
    '              if there are no records in the m_rstLookup recordset.
    '
    ' Returns:     True if it is at EOF; False otherwise
    '-----------------------------------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc       As String = "Property Get LookupIsAtEOF"

    If Not (m_rstLookup Is Nothing) Then
        LookupIsAtEOF = m_rstLookup.EOF
    Else
        LookupIsAtEOF = True
    End If
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Property
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Property



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Get LookupRecordCount() As Long
    ' **************************************************************************
    ' Function  : Property Get LookupRecordCount
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : string
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get LookupRecordCount"
    On Error GoTo PROC_ERR

    If Not (m_rstLookup Is Nothing) Then
        LookupRecordCount = m_rstLookup.RecordCount
    Else
        LookupRecordCount = 0
    End If
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Property
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Property



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Get LstUpdDtm() As Date
    ' **************************************************************************
    ' Function  : Property Get LstUpdDtm
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : Int
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get LstUpdDtm"
    On Error GoTo PROC_ERR

    LstUpdDtm = CDate(m_dteLstUpdDtm.value)
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Property
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Property



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Let LstUpdDtm(ByVal dteValue As Date)
    ' **************************************************************************
    ' Function  : Property Let LstUpdDtm
    ' Purpose   : Assigns new Value to property
    ' Parameters: ByVal NewValue As Integer
    ' Returns   :
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let LstUpdDtm"
    On Error GoTo PROC_ERR

    m_dteLstUpdDtm.value = dteValue
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Property
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Property



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Get LstUpdUserId() As String
    ' **************************************************************************
    ' Function  : Property Get LstUpdUserId
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : Int
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get LstUpdUserId"
    On Error GoTo PROC_ERR

    LstUpdUserId = CStr(m_strLstUpdUserId.value)
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Property
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Property



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Let LstUpdUserId(ByVal strValue As String)
    ' **************************************************************************
    ' Function  : Property Let LstUpdUserId
    ' Purpose   : Assigns new Value to property
    ' Parameters: ByVal NewValue As Integer
    ' Returns   :
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let LstUpdUserId"
    On Error GoTo PROC_ERR

    m_strLstUpdUserId.value = strValue
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Property
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Property



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Get Mask(ByVal strTagIn As String) As String
    ' **************************************************************************
    ' Function  : Property Get Mask
    ' Purpose   : Retrieves the default Mask (i.e. #####.###) from property
    ' Parameters: N/A
    ' Returns   : String
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get Mask"
    On Error GoTo PROC_ERR

    Mask = fnGetProperty(strTagIn).Mask
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Property
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Property



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Get MaxCharacters(ByVal strTagIn As String) As Long
    ' **************************************************************************
    ' Function  : Property Get MaxCharacters
    ' Purpose   : Retrieves the number of allowable characters from property
    ' Parameters: N/A
    ' Returns   : Long
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get MaxCharacters"
    On Error GoTo PROC_ERR

    MaxCharacters = fnGetProperty(strTagIn).MaxCharacters
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Property
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Property



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Get MktValCurrentRateCd() As String
    ' **************************************************************************
    ' Function  : Property Get MktvalCurrentRateCd
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : Int
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get MktvalCurrentRateCd"
    On Error GoTo PROC_ERR

    MktValCurrentRateCd = CStr(m_strMktValCurrentRateCd.value)
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Property
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Property



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Let MktValCurrentRateCd(ByVal strValue As String)
    ' **************************************************************************
    ' Function  : Property Let MktvalCurrentRateCd
    ' Purpose   : Assigns new Value to property
    ' Parameters: ByVal NewValue As Integer
    ' Returns   :
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let MktvalCurrentRateCd"
    On Error GoTo PROC_ERR

    m_strMktValCurrentRateCd.value = strValue
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Property
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Property



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Get ShouldForceToUppercase(ByVal strTagIn As String) As Boolean
    ' **************************************************************************
    ' Function  : Property Get ShouldForceToUppercase
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : Int
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get ShouldForceToUppercase"
    On Error GoTo PROC_ERR

    ShouldForceToUppercase = fnGetProperty(strTagIn).ShouldForceToUppercase
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Property
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Property





'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'\                                                                  /
'|                        PUBLIC  Procedures                        |
'/                                                                  \
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/

'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function AddRecord() As Boolean
    '--------------------------------------------------------------------------
    ' Procedure:   AddRecord
    ' Description: Adds a single record based on key value
    '              selection.
    ' Returns:     boolean
    ' Params:      Not necessary, they will be derived from properties the form
    '              should have already set
    ' Date:        04/11/2002
    '-----------------------------------------------------------------------------
    
    '!CUSTOMIZE!  Customize the name of the stored procedure, the number and names
    '             of the parameters and, perhaps, the return values trapped in the
    '             error handler.
    
    Const cstrCurrentProc          As String = "AddRecord"
    Const cstrSproc                As String = "dbo.proc_current rate_insert"  ' Stored procedure to execute
    Dim rstTemp                    As ADODB.Recordset
    Dim bSuccessful                As Boolean
    Dim prmReturnValue             As ADODB.Parameter
    Dim prmCurrentRateCd                  As ADODB.Parameter
    Dim prmCurrentRateNm                  As ADODB.Parameter
    Dim prmCurrentRateSvsInd              As ADODB.Parameter
    Dim prmMktvalCurrentRateCd            As ADODB.Parameter
    Dim prmCurrentRateMgrPrvCd            As ADODB.Parameter
    Dim prmCurrentRateMgr           As ADODB.Parameter
    Dim prmInvalid_Key             As ADODB.Parameter

    On Error GoTo PROC_ERR
    
    '...........................................................................
    ' No need to check to see if another user updated or deleted this record
    ' since we're doing an Add.
    '...........................................................................

    If Not (m_adwADO.CommandSetSproc(cstrSproc)) Then
        GoTo PROC_EXIT
    End If

    ' For Char/VarChar fields,
    '     * Use fnNullIfZLS to ensure Nulls are appropriately handled.
    '     * Do *not* set the optional 2nd parameter to fnNullIfZLS to True.
    ' For numeric fields,
    '     * Use fnNullIfZero to ensure Nulls are appropriately handled.
    ' For Y/N fields,
    '     * Use fnBoolToYN to ensure True/False is appropriately translated.

    With m_adwADO.ADOCommand
        ' ---Parameter #1---
        ' Define the return value that represents the error code (i.e. reason) why
        ' the stored procedure failed.
        Set prmReturnValue = .CreateParameter(Name:="@return_value", _
                                              Type:=adInteger, _
                                              Direction:=adParamReturnValue, _
                                              value:=Null)
        .Parameters.Append prmReturnValue

        ' ---Parameter #2---
        ' Define the current rate_CD parameter
        Set prmCurrentRateCd = .CreateParameter(Name:="current rate_cd", _
                                         Type:=adChar, _
                                         Direction:=adParamInput, _
                                         Size:=8, _
                                         value:=fnNullIfZLS(varIn:=CurrentRateCd, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmCurrentRateCd

        ' ---Parameter #3---
        ' Define the current rate_NM parameter
        Set prmCurrentRateNm = .CreateParameter(Name:="current rate_nm", _
                                         Type:=adVarChar, _
                                         Direction:=adParamInput, _
                                         Size:=50, _
                                         value:=fnNullIfZLS(varIn:=CurrentRateNm, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmCurrentRateNm

        ' ---Parameter #4---
        ' Define the current rate_SVS_IND parameter
        Set prmCurrentRateSvsInd = .CreateParameter(Name:="current rate_svs_ind", _
                                         Type:=adChar, _
                                         Direction:=adParamInput, _
                                         Size:=1, _
                                         value:=fnBoolToYN(CurrentRateSvsInd))
        .Parameters.Append prmCurrentRateSvsInd

        ' ---Parameter #5---
        ' Define the MKTVAL_current rate_CD parameter
        Set prmMktvalCurrentRateCd = .CreateParameter(Name:="mktval_current rate_cd", _
                                         Type:=adChar, _
                                         Direction:=adParamInput, _
                                         Size:=8, _
                                         value:=fnNullIfZLS(varIn:=MktValCurrentRateCd, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmMktvalCurrentRateCd

        ' ---Parameter #6---
        ' Define the current rate_MGR_PRV_CD parameter
        Set prmCurrentRateMgrPrvCd = .CreateParameter(Name:="current rate_mgr_prv_cd", _
                                         Type:=adChar, _
                                         Direction:=adParamInput, _
                                         Size:=8, _
                                         value:=fnNullIfZLS(varIn:=CurrentRateMgrPrvCd, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmCurrentRateMgrPrvCd

        ' ---Parameter #7---
        ' Define the current rate_MGR_current rate_CD parameter
        Set prmCurrentRateMgr = .CreateParameter(Name:="current rate_mgr_current rate_cd", _
                                         Type:=adChar, _
                                         Direction:=adParamInput, _
                                         Size:=8, _
                                         value:=fnNullIfZLS(varIn:=CurrentRateMgr, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmCurrentRateMgr

        ' ---Parameter #8---
        ' Define the Invalid_Key output parameter, which reflects *which* foreign
        ' key violation was encountered.
        Set prmInvalid_Key = .CreateParameter(Name:="@Invalid_Key", _
                                                  Type:=adVarChar, _
                                                  Size:=255, _
                                                  Direction:=adParamOutput)
        .Parameters.Append prmInvalid_Key

        ' Do the Add
        .Execute
    End With

    '...........................................................................
    ' Refresh the Lookup recordset, re-retrieve the just-added record so that
    ' record is *still* the current record, and load its data to the
    ' table wrapper's class properties so all table columns (including
    ' those set by the DBMS like identity and Last Updated columns) are
    ' up-to-date.
    '...........................................................................
    bSuccessful = GetRelativeRecord(CurrentRateCd, epdSameRecord)

    AddRecord = True
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
    
    ' Clean-up statements go here
    fnFreeRecordset rstTemp
    fnFreeObject prmReturnValue
    fnFreeObject prmCurrentRateCd
    fnFreeObject prmCurrentRateNm
    fnFreeObject prmCurrentRateSvsInd
    fnFreeObject prmMktvalCurrentRateCd
    fnFreeObject prmCurrentRateMgrPrvCd
    fnFreeObject prmCurrentRateMgr
    fnFreeObject prmInvalid_Key
    
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case prmReturnValue
        Case gcRES_NERR_ERR_WHILE_TRYING_TO   ' 4028
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_ERR_WHILE_TRYING_TO, _
                                       mcstrName & cstrCurrentProc, _
                                       "add"
            Resume PROC_EXIT
        Case gcRES_NERR_ADD_WITH_NONUNIQUE_KEY   ' 4031
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_ADD_WITH_NONUNIQUE_KEY, _
                                       mcstrName & cstrCurrentProc, _
                                       RTrim$(CurrentRateCd)
            Resume PROC_EXIT
        Case gcRES_NERR_KEY_NOT_FOUND ' 4032
            If UCase$(prmInvalid_Key) = "MKTVAL_current rate_CD" Then
                gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_KEY_NOT_FOUND, _
                                           mcstrName & cstrCurrentProc, _
                                           "Market Value current rate Cd", RTrim$(MktValCurrentRateCd), "current rate"
                Resume PROC_EXIT
            Else
                ' current rate_MGR_PRV_CD
                gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_KEY_NOT_FOUND, _
                                           mcstrName & cstrCurrentProc, _
                                           "current rate Mgr", RTrim$(CurrentRateMgrPrvCd), "Provider"
                Resume PROC_EXIT
            End If
    End Select

    ' If any other errors exist, i.e. in Err object, then let it fall through into default error handling.

    Select Case Err.Number
        Case -2147217900 ' Object not found
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_FERR_SPROC_NOT_FOUND, _
                                       mcstrName & cstrCurrentProc, _
                                       cstrSproc
            Resume PROC_EXIT
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function CheckForAnotherUsersChanges(ByVal lngWhatOperation As enumWhatOperationIsBeingAttempted, _
    ByRef strACF2 As String) As Long
    '--------------------------------------------------------------------------
    ' Procedure:   CheckForAnotherUsersChanges
    ' Description: Check to see if another user has altered or deleted
    '              the record that is about to be operated upon. This is called
    '              (directly or indirectly) by each form's cmdDelete_Click and
    '              cmdUpdate_Click event handlers.
    '
    '              NOTE: The caller must check for every possible return value
    '                    that a given lngWhatOperation value could hit!
    '
    ' Returns:     A return code indicating what has occured, so the form
    '              can determine what to do about it. A zero value means the form
    '              doesn't have to do anything.
    '
    ' Params:
    '    lngWhatOperation (in) - indicates whether an Update or Delete is being attempted
    '    strACF2 (out)         - for some errors, reflects the ACF2 id of the user who updated
    '                            the record
    '
    ' Date:        04/27/2002
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc       As String = "CheckForAnotherUsersChanges"
    Const clngNoError           As Long = 0
    ' rstSingleRecord_Fresh contains all columns of a single row in this class' underlying table. It reflects
    ' the now-current contents of the record that is about to be updated or deleted, so the CheckForAnotherUsersChanges
    ' process can determine if another user updated the record since it was originally retrieved.
    Dim rstSingleRecord_Fresh       As ADODB.Recordset

    On Error GoTo PROC_ERR

    '...........................................................................
    ' See if another user deleted or updated the record since we last retrieved it...
    '...........................................................................
    
    ' The following statement will raise a 4027 if the specified record isn't found. PROC_ERR does
    ' a Resume Next so the first validation (to see if another user deleted the record) needs
    ' to check for both .RecordCount=0 --or-- rst=Nothing; otherwise a runtime error 91
    ' (Object variable or With block not set) is raised.
    '!CUSTOMIZE! fnSelectRecord call should pass the key column(s)
    Set rstSingleRecord_Fresh = fnSelectRecord(CurrentRateCd)
    
    ' Disconnect the recordset so we can edit the data, if desired, for testing purposes
    fnFreeRecordset rstSingleRecord_Fresh.ActiveConnection

    If (rstSingleRecord_Fresh Is Nothing) Or (rstSingleRecord_Fresh.RecordCount = 0) Then
        Select Case lngWhatOperation
            Case ewoUpdate
                CheckForAnotherUsersChanges = vbObjectError + gcRES_INFO_ANOTHER_USER_DELETED_NOT_SAVED
            Case Else       ' ewoDelete
                CheckForAnotherUsersChanges = vbObjectError + gcRES_INFO_ANOTHER_USER_DELETED
        End Select
        GoTo PROC_EXIT
    End If
    
    ' Note: A "<>" comparison on the date/time values reports a false positive. Use DateDiff( ) instead.
    ' Convert dates to string using CStr( ) to avoid ADO's millisecond rounding which could result in a false positive.
    With rstSingleRecord_Fresh
        If lngWhatOperation = ewoUpdate Then
            If (DateDiff("s", CStr(m_dteLstUpdDtm_Original), CStr(!lst_upd_dtm)) <> 0) _
               Or _
               (!lst_upd_user_id <> m_strLstUpdUserId_Original) Then
                    strACF2 = !lst_upd_user_id
                    CheckForAnotherUsersChanges = vbObjectError + gcRES_INFO_ANOTHER_USER_UPDATED_DISCARDED
            End If
        End If
    End With
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler

    ' Clean-up statements go here
    fnFreeRecordset rstSingleRecord_Fresh
    
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        Case 91     ' Object variable or With block variable not set
            ' This error will be encountered if the call to .SelectRecord didn't find the specified
            ' record, e.g., another user deleted it. Ignore it so the logic that generates the
            ' desired "transformed" error code will be hit.
            Resume Next
        Case vbObjectError + gcRES_NERR_REC_NOT_FOUND   ' 4027
            ' If we got record not found from the call to SelectRecord(), then wipe out traces
            ' of that error and do a Resume Next. This will allow this proc to
            ' transform *that* error into the one we really want:
            ' gcRES_INFO_ANOTHER_USER_DELETED_NOT_SAVED vs. gcRES_INFO_ANOTHER_USER_DELETED
            gerhApp.Clear
            Resume Next
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function DeleteRecord() As Boolean
    '--------------------------------------------------------------------------
    ' Procedure:   DeleteRecord
    ' Description: Deletes a single record based on the value(s) in the
    '              properties that correspond to the table's key(s)
    ' Returns:     True if successful, False otherwise
    '-----------------------------------------------------------------------------
    
    '!CUSTOMIZE!  Customize the name of the stored procedure, the number and names
    '             of the parameters and, perhaps, the return values trapped in the
    '             error handler.
    
    Const cstrCurrentProc       As String = "DeleteRecord"
    Const cstrSproc                As String = "dbo.proc_current rate_delete"  ' Stored procedure to execute
    Dim bSuccessful                As Boolean
    Dim rstTemp                    As ADODB.Recordset
    Dim prmReturnValue             As ADODB.Parameter
    Dim prmCurrentRateCd           As ADODB.Parameter
    Dim prmDependent_Table         As ADODB.Parameter

    On Error GoTo PROC_ERR

    If Not (m_adwADO.CommandSetSproc(cstrSproc)) Then
        GoTo PROC_EXIT
    End If

    ' For Char/VarChar fields,
    '     * Use fnNullIfZLS to ensure Nulls are appropriately handled.
    '     * Do *not* set the optional 2nd parameter to fnNullIfZLS to True.
    ' For numeric fields,
    '     * Use fnNullIfZero to ensure Nulls are appropriately handled.
    ' For Y/N fields,
    '     * Use fnBoolToYN to ensure True/False is appropriately translated.

    With m_adwADO.ADOCommand
        ' ---Parameter #1---
        ' Define the return value that represents the error code (i.e. reason) why
        ' the stored procedure failed.
        Set prmReturnValue = .CreateParameter(Name:="@return_value", _
                                              Type:=adInteger, _
                                              Direction:=adParamReturnValue, _
                                              value:=Null)
        .Parameters.Append prmReturnValue

        ' ---Parameter #2---
        ' Define the current rate_CD parameter
        Set prmCurrentRateCd = .CreateParameter(Name:="current rate_cd", _
                                         Type:=adChar, _
                                         Direction:=adParamInput, _
                                         Size:=8, _
                                         value:=fnNullIfZLS(varIn:=CurrentRateCd, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmCurrentRateCd

        ' ---Parameter #3---
        ' Define the Dependent_Table output parameter, which reflects *which* dependent table
        ' contains rows with a foreign key equal to the key being deleted.
        Set prmDependent_Table = .CreateParameter(Name:="@Dependent_Table", _
                                                  Type:=adVarChar, _
                                                  Size:=255, _
                                                  Direction:=adParamOutput)
        .Parameters.Append prmDependent_Table

        ' Do the Delete
        .Execute
    End With
    
    '...........................................................................
    ' Refresh the Lookup recordset, reposition the Lookup data so the record
    ' prior to the one just deleted is now the current record. Load that
    ' record's data to the table wrapper's class properties.
    '...........................................................................
    bSuccessful = GetRelativeRecord(CurrentRateCd, epdPreviousRecord)
    
    DeleteRecord = True
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
    
    ' Clean-up statements go here
    fnFreeRecordset rstTemp
    fnFreeObject prmReturnValue
    fnFreeObject prmCurrentRateCd
    fnFreeObject prmDependent_Table
    
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case prmReturnValue
        Case gcRES_NERR_REC_NOT_FOUND       ' 4027
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_REC_NOT_FOUND, _
                                       mcstrName & cstrCurrentProc, _
                                       "current rate Code " & RTrim$(CurrentRateCd)
            Resume PROC_EXIT
        Case gcRES_NERR_ERR_WHILE_TRYING_TO   ' 4028
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_ERR_WHILE_TRYING_TO, _
                                       mcstrName & cstrCurrentProc, _
                                       "add"
            Resume PROC_EXIT
        Case gcRES_NERR_DEPENDENT_RECS_EXIST   ' 4029
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_DEPENDENT_RECS_EXIST, _
                                       mcstrName & cstrCurrentProc, _
                                       "current rate Cd (" & RTrim$(CurrentRateCd) & ")", UCase$(prmDependent_Table)
            Resume PROC_EXIT
    End Select

    ' If any other errors exist, i.e. in Err object, then let it fall through into default error handling.

    Select Case Err.Number
        Case -2147217900 ' Object not found
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_FERR_SPROC_NOT_FOUND, _
                                       mcstrName & cstrCurrentProc, _
                                       cstrSproc
            Resume PROC_EXIT
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function GetLookupData() As Boolean
    '--------------------------------------------------------------------------
    ' Procedure:   GetLookupData
    ' Description: Get all rows but only particular columns
    ' Returns:     True if successful; False otherwise
    ' Params:      None
    ' Date:        01/07/2002
    '-----------------------------------------------------------------------------
    
    '!CUSTOMIZE!  Customize the name of the stored procedure, the number and names
    '             of the parameters and, perhaps, the return values trapped in the
    '             error handler.
    
    Const cstrCurrentProc       As String = "GetLookupData"
    Const cstrSproc                As String = "dbo.proc_current rate_lu_select"  ' Stored procedure to execute
    Dim prmReturnValue             As ADODB.Parameter

    On Error GoTo PROC_ERR

    If Not (m_adwADO.CommandSetSproc(cstrSproc)) Then
        GoTo PROC_EXIT
    End If

    With m_adwADO.ADOCommand
        ' ---Parameter #1---
        ' Define the return value that represents the error code (i.e. reason) why
        ' the stored procedure failed.
        Set prmReturnValue = .CreateParameter(Name:="@return_value", _
                                              Type:=adInteger, _
                                              Direction:=adParamReturnValue, _
                                              value:=Null)
        .Parameters.Append prmReturnValue

        Set m_rstLookup = .Execute()
    End With
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler

    ' Clean-up statements go here
    ' Have to check for "Not (m_rstLookup Is Nothing)" to avoid a
    ' "91 - Object variable or With block variable not set" runtime error
    If Not (m_rstLookup Is Nothing) Then
        fnFreeObject m_rstLookup.ActiveConnection           ' Disconnect the Recordset
    End If
    fnFreeObject prmReturnValue

    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case prmReturnValue
'        Case gcRES_NERR_REC_NOT_FOUND       ' 4027
'            ' Wipe out any trace of this error, but return False so the caller
'            ' knows to go into Add mode if desired. NOTE: The caller can also
'            ' identify this by looking at the LookupRecordCount public property.
'            Err.Clear
'            GetLookupData = False
'            Resume PROC_EXIT
        Case gcRES_NERR_ERR_WHILE_TRYING_TO   ' 4028
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_ERR_WHILE_TRYING_TO, _
                                       mcstrName & cstrCurrentProc, _
                                       "locate"
            Resume PROC_EXIT
    End Select

    ' If any other errors exist, i.e. in Err object, then let it fall through into default error handling.

    Select Case Err.Number
        Case -2147217900 ' Object not found
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_FERR_SPROC_NOT_FOUND, _
                                       mcstrName & cstrCurrentProc, _
                                       cstrSproc
            Resume PROC_EXIT
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function GetRelativeRecord(ByVal strKey1 As String, _
    ByVal lngPositionDirection As enumPositionDirection) As Boolean
    '--------------------------------------------------------------------------
    ' Procedure:   GetRelativeRecord
    ' Description: Refreshes the Lookup recordset and repositions it to
    '              the record relative to the specified key value. Then,
    '              it resets each of the class properties that correspond
    '              to columns in the underlying table so the form is able
    '              to load that newly-positioned-to record's data to
    '              itrs on-screen controls.
    '
    '              NOTE: For each table key, there should be an input parameter
    '                    and a local var (i.e. strKey1ForNewRec) of the
    '                    appropriate data type! Also, the setting of
    '                    the .Filter property below must reflect each
    '                    table key.
    '
    ' Params:
    '     strKey1              (in) = CurrentRateCd value from which to do the relative
    '                                 repositioning
    '     lngPositionDirection (in) = Indicates to which relative record the
    '                                 recordset should be positioned (relative
    '                                 to the strKey1 parameter value).
    '
    '
    ' Called By:   cmdDelete_Click( ) of frmcurrent rate.frm
    '              cmdUpdate_Click( ) of frmcurrent rate.frm
    '              cmdNavigate_Click( ) of frmcurrent rate.frm
    '              fnAddRecord( ) of frmcurrent rate.frm
    '
    ' Returns:     True if successful; False otherwise
    '-----------------------------------------------------------------------------
    
    '!CUSTOMIZE!  This proc and all calls to it must be customized to reflect
    '             one parameter for each key column of the table. Make sure the
    '             parameter is defined to be of the right data type. Also,
    '             the way the recordset's .Find property is set must be changed
    '             to reflect each key column so the right record will be located.
    '             Also make sure that the substitution values passed to
    '             SaveAppSpecificError are correct and TRIM'd if appropriate.

    Const cstrCurrentProc       As String = "GetRelativeRecord"
    Const cintNoRecords         As Integer = 0
    Dim rstTemp                 As ADODB.Recordset
    Dim strKey1ForNewRec        As String

    On Error GoTo PROC_ERR

    '...........................................................................
    ' Refresh the lookup data (m_rstLookupData) so other's changes
    ' --and our own-- are now reflected in it. This resets the Lookup data,
    ' record count, and current record number, and leaves the Lookup recordset
    ' positioned to the first record (if there are records) or BOF (if there are
    ' no records).
    '...........................................................................
    GetLookupData

    With m_rstLookup
        Select Case lngPositionDirection
            Case epdPreviousRecord
                ' Make visible only those rows with keys prior to the specified key
                .Filter = "current rate_cd < '" & strKey1 & "'"
                If .RecordCount <> 0 Then
                    ' Grab the last record. The one with the highest key less than the
                    ' specified key is the one we want.
                    m_adwADO.MoveLast m_rstLookup
                    '!CUSTOMIZE! The constant referenced below should refer to the key field.
                    strKey1ForNewRec = .Fields(mcstrCurrentRateCd).value
                    ' Okay, we got it. Now make all records in the Lookup recordset
                    ' visible again, and then reposition to the new record.
                    .Filter = adFilterNone
                    m_adwADO.MoveFirst m_rstLookup
                    .Find "current rate_cd = '" & strKey1ForNewRec & "'"
                    ' If the new record wasn't found, generate an error. We should
                    ' never hit this error, except due to bad program logic, since
                    ' it means the new record whose key value was just identified
                    ' could not be found.
                    If .EOF Then
                        ' Should never hit this code. It means the new record whose key
                        ' value was just identified could not be found.
                        gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_ERR_WHILE_TRYING_TO, _
                                                   mcstrName & cstrCurrentProc, _
                                                   "locate"
                        Resume PROC_EXIT
                    End If
                Else
                    ' No records meet the criteria.
                    ' If there are any records, show the first one.
                    ' If there are no records, the caller (the form) should go into Add mode
                    ' upon seeing that the m_rstLookup.LookupRecordCount = 0.
                    .Filter = adFilterNone
                    If m_rstLookup.RecordCount <> 0 Then
                        m_adwADO.MoveFirst m_rstLookup
                    End If
                End If
            
            
            Case epdNextRecord
                ' Make visible only those rows with keys prior to the specified key
                .Filter = "current rate_cd > '" & strKey1 & "'"
                If .RecordCount <> 0 Then
                    ' Grab the first record. The one with the lowest key higher than the
                    ' specified key is the one we want.
                    m_adwADO.MoveFirst m_rstLookup
                    '!CUSTOMIZE! The constant referenced below should refer to the key field.
                    strKey1ForNewRec = .Fields(mcstrCurrentRateCd).value
                    ' Okay, we got it. Now make all records in the Lookup recordset
                    ' visible again, and then reposition to the new record.
                    .Filter = adFilterNone
                    m_adwADO.MoveFirst m_rstLookup
                    .Find "current rate_cd = '" & strKey1ForNewRec & "'"
                    ' If the new record wasn't found, generate an error. We should
                    ' never hit this error, except due to bad program logic, since
                    ' it means the new record whose key value was just identified
                    ' could not be found.
                    If .EOF Then
                        gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_ERR_WHILE_TRYING_TO, _
                                                   mcstrName & cstrCurrentProc, _
                                                   "locate"
                        Resume PROC_EXIT
                    End If
                Else
                    ' No records meet the criteria.
                    ' If there are any records, show the last one.
                    ' If there are no records, the caller (the form) should go into Add mode
                    ' upon seeing that the m_rstLookup.LookupRecordCount = 0.
                    .Filter = adFilterNone
                    If m_rstLookup.RecordCount <> 0 Then
                        m_adwADO.MoveLast m_rstLookup
                    End If
                End If
            
            
            
            Case epdSameRecord
                ' This is used by Update processing, where we just want to
                ' stay on the just-updated record but make sure its current
                ' data (from the DBMS) is loaded to the class properties and
                ' other's changes to any record in the Lookup recordset are
                ' visible.
                .Find "current rate_cd = '" & strKey1 & "'"
                ' If the record wasn't found, generate an error. We should
                ' never hit this error, except due to bad program logic, since
                ' it means the current record could not be found.
                If .EOF Then
                    gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_ERR_WHILE_TRYING_TO, _
                                               mcstrName & cstrCurrentProc, _
                                               "locate"
                    Resume PROC_EXIT
                End If
            
            
            Case epdFirstRecord
                ' This operation is used by the form's cmdNavigate_Click.
                ' It ignores the passed-in key parameter(s).
                ' Do *not* generate an error if we hit BOF since that event
                ' handler will use see that the Lookup recordset's position
                ' is at BOF and throw the form into Add mode.
                m_adwADO.MoveFirst m_rstLookup
            
            
            Case epdLastRecord
                ' This operation is used by the form's cmdNavigate_Click.
                ' It ignores the passed-in key parameter(s).
                ' Do *not* generate an error if we hit BOF since that event
                ' handler will use see that the Lookup recordset's position
                ' is at EOF and throw the form into Add mode.
                m_adwADO.MoveLast m_rstLookup
            
            
            Case Else
                gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_UNEXPECTED_VAL_SELECT_CASE, _
                    mcstrName & cstrCurrentProc
                GoTo PROC_EXIT
        End Select
    
        '...........................................................................
        ' Get the column data for the just-repositioned-to record and load it to the
        ' class properties corresponding to those columns.
        '...........................................................................
        If .BOF And .EOF Then
            fnClearPropertyValues
        Else
            '!CUSTOMIZE! The constant referenced below should refer to the key field.
            GetSingleRecord .Fields(mcstrCurrentRateCd).value
        End If
    End With
PROC_EXIT:
    On Error GoTo 0     ' disable error handler

    ' Clean-up statements go here
    fnFreeRecordset rstTemp

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
Public Function GetSingleRecord(ByVal strKey1 As String, _
    Optional ByVal bSynchLookupRST As Boolean = False) As Boolean
    '--------------------------------------------------------------------------
    ' Procedure:   GetSingleRecord
    ' Description: Obtains data from the database for the specified key(s).
    '              It then loads its columnar values to the class properties
    '              that correspond to those columns. It also saves the
    '              Last Updated info to separate member variables so it can
    '              be used (when/if the user tries to update or delete
    '              the record) to determine if another user affected this
    '              record since this function retrieved it.
    '
    '              This proc should not *refresh* the Lookup recordset. It
    '              is merely retrieving all of the table columns for the
    '              specified record. The Lookup recordset only contains
    '              a subset of the columns for that key.
    '
    '
    '              NOTE: For each table key, there should be an input parameter
    '                    of the appropriate data type!
    '
    ' Returns:     Boolean
    ' Params:
    '    strKey1         (in) = represents the primary key for the table (current rate_cd)
    '    bSynchLookupRST (in) = indicates whether the Lookup recordset should be
    '                           repositioned to the record this function just
    '                           retrieved. This would be set to True by the
    '                           form's vfgLookup_ChangeEdit event handler to ensure
    '                           the "record x of y" will be set appropriately when
    '                           it calls fnLoadControls.
    '-----------------------------------------------------------------------------
    
    '!CUSTOMIZE!  This proc and all calls to it must be customized to reflect
    '             one parameter for each key column of the table. Make sure the
    '             parameter is defined to be of the right data type. The
    '             "With Me" block must be updated to reflect the current set of
    '             wrapper properties and table column names. Also, the way
    '             the recordset's .Find property is set must be changed
    '             to reflect each key column so the right record will be located.
    '             Also make sure that the substitution values passed to
    '             SaveAppSpecificError are correct and TRIM'd if appropriate.

    Const cstrCurrentProc       As String = "GetSingleRecord"
    Dim rstTemp                 As ADODB.Recordset

    On Error GoTo PROC_ERR

    GetSingleRecord = False

    '!CUSTOMIZE! fnSelectRecord call should pass the key column(s)
    Set rstTemp = fnSelectRecord(strKey1)

    With Me
        ' For Char/VarChar fields,
        '     * Use fnZLSIfNull to ensure Nulls are appropriately translated.
        ' For Numeric fields,
        '     * Use fnZeroIfNull to ensure Nulls are appropriately translated.
        ' For Boolean fields,
        '     * Use fnYNToBool to ensure True/False is appropriately translated.
        .CurrentRateCd = fnZLSIfNull(rstTemp!current_rate_cd)
        .CurrentRateNm = fnZLSIfNull(rstTemp!current_rate_nm)
        .LstUpdDtm = fnZLSIfNull(rstTemp!lst_upd_dtm)
        .LstUpdUserId = fnZLSIfNull(rstTemp!lst_upd_user_id)
        .CurrentRateSvsInd = fnYNToBool(rstTemp!current_rate_svs_ind)
        .MktValCurrentRateCd = fnZLSIfNull(rstTemp!mktval_current_rate_cd)
        .CurrentRateMgrPrvCd = fnZLSIfNull(rstTemp!current_rate_mgr_prv_cd)
        .CurrentRateMgr = fnZLSIfNull(rstTemp!current_rate_mgr_current_rate_cd)
    End With

    ' Save original Last Updated info, to be used during UpdateRecord( ) and DeleteRecord( )
    ' to determine if another user updated the record since it was retrieved.
    m_dteLstUpdDtm_Original = fnZLSIfNull(rstTemp!lst_upd_dtm)
    m_strLstUpdUserId_Original = fnZLSIfNull(rstTemp!lst_upd_user_id)
    
    If bSynchLookupRST Then
        With m_rstLookup
            m_adwADO.MoveFirst m_rstLookup
            .Find "current rate_cd = '" & CurrentRateCd & "'"
            If .EOF Then
                gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_ERR_WHILE_TRYING_TO, _
                                           mcstrName & cstrCurrentProc, _
                                           "locate"
                Resume PROC_EXIT
            End If
        End With
    End If

    GetSingleRecord = True
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
    
    ' Clean-up statements go here
    fnFreeRecordset rstTemp
    
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    GetSingleRecord = False

    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub GoToFirstRecord()
    ' **************************************************************************
    ' Function  : GoToFirstRecord
    ' Purpose   : Moves to the First record in the table
    ' Parameters: N/A
    ' Returns   : Int
    ' **************************************************************************
    Const cstrCurrentProc   As String = "GoToFirstRecord"
    On Error GoTo PROC_ERR

    If m_rstLookup.RecordCount > 0 Then
        m_adwADO.MoveFirst m_rstLookup
        ' Get the requested record and reposition the Lookup recordset to that record
        '!CUSTOMIZE! The constant referenced below should refer to the key field.
        GetSingleRecord m_rstLookup.Fields(mcstrCurrentRateCd).value
    End If
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
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
Public Sub GoToLastRecord()
    ' **************************************************************************
    ' Function  : GoToLastRecord
    ' Purpose   : Moves to the Last record in the table
    ' Parameters: N/A
    ' Returns   : Int
    ' **************************************************************************
    Const cstrCurrentProc   As String = "GoToLastRecord"
    On Error GoTo PROC_ERR

    If m_rstLookup.RecordCount > 0 Then
        m_adwADO.MoveLast m_rstLookup
        ' Get the requested record and reposition the Lookup recordset to that record
        '!CUSTOMIZE! The constant referenced below should refer to the key field.
        GetSingleRecord m_rstLookup.Fields(mcstrCurrentRateCd).value
    End If
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
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
Public Sub GoToNextRecord()
    ' **************************************************************************
    ' Function  : GoToNextRecord
    ' Purpose   : Moves to the next record in the table
    ' Parameters: N/A
    ' Returns   : Int
    ' **************************************************************************
    Const cstrCurrentProc   As String = "GoToNextRecord"
    On Error GoTo PROC_ERR

    With m_rstLookup
        If .RecordCount > 0 Then
            m_adwADO.MoveNext m_rstLookup
            If LookupIsAtBOF Or LookupIsAtEOF Then
                GetLookupData
                GetRelativeRecord CurrentRateCd, epdNextRecord
            Else
                ' Get the requested record and reposition the Lookup recordset to that record
                '!CUSTOMIZE! The constant referenced below should refer to the key field.
                GetSingleRecord m_rstLookup.Fields(mcstrCurrentRateCd).value
            End If
        End If
    End With
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
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
Public Sub GoToPreviousRecord()
    ' **************************************************************************
    ' Function  : GoToPreviousRecord
    ' Purpose   : Moves to the previous record in the table
    ' Parameters: N/A
    ' Returns   : Int
    ' **************************************************************************
    Const cstrCurrentProc   As String = "GoToPreviousRecord"
    On Error GoTo PROC_ERR

    With m_rstLookup
        If .RecordCount > 0 Then
            m_adwADO.MovePrev m_rstLookup
            If LookupIsAtBOF Or LookupIsAtEOF Then
                GetLookupData
                GetRelativeRecord CurrentRateCd, epdPreviousRecord
            Else
                ' Get the requested record and reposition the Lookup recordset to that record
                '!CUSTOMIZE! The constant referenced below should refer to the key field.
                GetSingleRecord m_rstLookup.Fields(mcstrCurrentRateCd).value
            End If
        End If
    End With
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
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
Public Function HaveDependents(ByVal strKey1 As String, ByRef strDependentTable As String) As Boolean
    ' Comments  : Determines whether the current record can be deleted without
    '             hitting a referential integrity violation due to either:
    '             a. row(s) existing in other tables that use the current key value
    '                as a foreign key
    '             b. (for current_rate_t table only, I think) row(s) existing in the same table
    '                which has a circular reference to to the current key value.
    '             The calling form should look at the return value. If True, then
    '             the form's Delete button should be disabled.
    '
    ' Parameters:
    '   strKey1
    '
    ' Called by : fnSetCommandButtons( ) in each maintenance screen
    '
    ' Returns   : True if there are children or other dependencies; False otherwise
    ' Modified  :
    ' --------------------------------------------------
    
    '!CUSTOMIZE!  This proc must be customized to return True unconditionally, if
    '             the table has no dependencies to any other tables.
    '
    '             Otherwise, it must be customized to have an input parameter of
    '             the correct data type for each key to the table, call the
    '             correct stored procedure with the correct number and type of
    '             paraemters, and interpret its return values correctly.
    Const cstrCurrentProc       As String = "HaveDependents"
    Const cstrSproc            As String = "dbo.proc_current rate_verify_dependents"  ' Stored procedure to execute
    Dim adwTemp                As cadwADOWrapper
    Dim rstTemp                As ADODB.Recordset
    Dim prmReturnValue         As ADODB.Parameter
    Dim prmCurrentRateCd       As ADODB.Parameter
    Dim prmDependent_Table     As ADODB.Parameter

    On Error GoTo PROC_ERR

    Set adwTemp = New cadwADOWrapper
    adwTemp.CommandSetSproc strSprocName:=cstrSproc

    With adwTemp.ADOCommand
        ' ---Parameter #1---
        ' Define the return value that represents the error code (i.e. reason) why
        ' the stored procedure failed.
        Set prmReturnValue = .CreateParameter(Name:="@return_value", _
                                              Type:=adInteger, _
                                              Direction:=adParamReturnValue, _
                                              value:=Null)
        .Parameters.Append prmReturnValue

        ' ---Parameter #2---
        ' Define the input parameter that represents the key value to see
        ' if it exists as a foreign key on dependent tables
        Set prmCurrentRateCd = .CreateParameter(Name:="@current rate_cd", _
                                         Type:=adChar, _
                                         Direction:=adParamInput, _
                                         Size:=8, _
                                         value:=strKey1)
        .Parameters.Append prmCurrentRateCd

        ' ---Parameter #3---
        ' Define the output parameter that indicates whether **any** dependent table
        ' has children. If True, we need to look at prm2 and report a 4029 error.
        Set prmDependent_Table = .CreateParameter(Name:="@Dependent_Table", _
                                                  Type:=adVarChar, _
                                                  Size:=255, _
                                                  Direction:=adParamOutput)
        .Parameters.Append prmDependent_Table

        ' Now execute the sproc...and you get access to those output parameters
        ' as well as, if applicable, the recordset/resultset it returns
        Set rstTemp = .Execute
    End With
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
    
    ' Clean-up statements go here
    fnFreeObject adwTemp
    fnFreeRecordset rstTemp
    fnFreeObject prmReturnValue
    fnFreeObject prmCurrentRateCd
    fnFreeObject prmDependent_Table
    
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case prmReturnValue
        ' The 4027 return code should only occur if a multi-user situation has occurred,
        ' such as User A going into Update mode on a record that another user just
        ' deleted. For this reason, we'll remove any trace that this error occurred
        ' and return True. When and if the user clicks Update, then they'll get a
        ' message that another user deleted the record.
        Case gcRES_NERR_REC_NOT_FOUND           ' 4027
            HaveDependents = True
            strDependentTable = "Unknown"   ' This is actually ignored by the caller
            ' Remove any trace that this error occurred since we're not going to report it as an error.
            Err.Clear
            Resume PROC_EXIT
        Case gcRES_NERR_DEPENDENT_RECS_EXIST    ' 4029
            HaveDependents = True
            strDependentTable = prmDependent_Table
            ' Remove any trace that this error occurred since we're not going to report it as an error.
            Err.Clear
            Resume PROC_EXIT
    End Select

    ' If any other errors exist, i.e. in Err object, then let it fall through into default error handling.

    Select Case Err.Number
        Case -2147217900 ' Object not found
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_FERR_SPROC_NOT_FOUND, _
                                       mcstrName & cstrCurrentProc, _
                                       cstrSproc
            Resume PROC_EXIT
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function UpdateRecord() As Boolean
    '--------------------------------------------------------------------------
    ' Procedure:   UpdateRecord
    ' Description: Updates a single record based on current values stored in
    '              the class properties corresponding to table columns.
    '
    ' Returns:     True if successful; False otherwise
    ' Params:      N/A
    ' Date:        01/07/2002
    '-----------------------------------------------------------------------------

    '!CUSTOMIZE!  Customize the name of the stored procedure, the number and names
    '             of the parameters and, perhaps, the return values trapped in the
    '             error handler.

    Const cstrCurrentProc          As String = "UpdateRecord"
    Const cstrSproc                As String = "dbo.proc_current rate_update"  ' Stored procedure to execute
    Dim rstTemp                    As ADODB.Recordset
    Dim bSuccessful                As Boolean
    Dim prmReturnValue             As ADODB.Parameter
    Dim prmCurrentRateCd           As ADODB.Parameter
    Dim prmCurrentRateNm           As ADODB.Parameter
    Dim prmCurrentRateSvsInd       As ADODB.Parameter
    Dim prmMktvalCurrentRateCd     As ADODB.Parameter
    Dim prmCurrentRateMgrPrvCd     As ADODB.Parameter
    Dim prmCurrentRateMgr           As ADODB.Parameter
    Dim prmInvalid_Key             As ADODB.Parameter
    
    On Error GoTo PROC_ERR
    
    If Not (m_adwADO.CommandSetSproc(cstrSproc)) Then
        GoTo PROC_EXIT
    End If
    
    ' For Char/VarChar fields,
    '     * Use fnNullIfZLS to ensure Nulls are appropriately handled.
    '     * Do *not* set the optional 2nd parameter to fnNullIfZLS to True.
    ' For numeric fields,
    '     * Use fnNullIfZero to ensure Nulls are appropriately handled.
    ' For Y/N fields,
    '     * Use fnBoolToYN to ensure True/False is appropriately translated.
    

    With m_adwADO.ADOCommand
        ' ---Parameter #1---
        ' Define the return value that represents the error code (i.e. reason) why
        ' the stored procedure failed.
        Set prmReturnValue = .CreateParameter(Name:="@return_value", _
                                              Type:=adInteger, _
                                              Direction:=adParamReturnValue, _
                                              value:=Null)
        .Parameters.Append prmReturnValue

        ' ---Parameter #2---
        ' Define the current rate_CD parameter
        Set prmCurrentRateCd = .CreateParameter(Name:="current rate_cd", _
                                         Type:=adChar, _
                                         Direction:=adParamInput, _
                                         Size:=8, _
                                         value:=fnNullIfZLS(varIn:=CurrentRateCd, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmCurrentRateCd

        ' ---Parameter #3---
        ' Define the current rate_NM parameter
        Set prmCurrentRateNm = .CreateParameter(Name:="current rate_nm", _
                                         Type:=adVarChar, _
                                         Direction:=adParamInput, _
                                         Size:=50, _
                                         value:=fnNullIfZLS(varIn:=CurrentRateNm, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmCurrentRateNm

        ' ---Parameter #4---
        ' Define the current rate_SVS_IND parameter
        Set prmCurrentRateSvsInd = .CreateParameter(Name:="current rate_svs_ind", _
                                             Type:=adChar, _
                                             Direction:=adParamInput, _
                                             Size:=1, _
                                             value:=fnBoolToYN(CurrentRateSvsInd))
        .Parameters.Append prmCurrentRateSvsInd

        ' ---Parameter #5---
        ' Define the MKTVAL_current rate_CD parameter
        Set prmMktvalCurrentRateCd = .CreateParameter(Name:="mktval_current rate_cd", _
                                         Type:=adChar, _
                                         Direction:=adParamInput, _
                                         Size:=8, _
                                         value:=fnNullIfZLS(varIn:=MktValCurrentRateCd, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmMktvalCurrentRateCd

        ' ---Parameter #6---
        ' Define the current rate_MGR_PRV_CD parameter
        Set prmCurrentRateMgrPrvCd = .CreateParameter(Name:="current rate_mgr_prv_cd", _
                                         Type:=adChar, _
                                         Direction:=adParamInput, _
                                         Size:=8, _
                                         value:=fnNullIfZLS(varIn:=CurrentRateMgrPrvCd, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmCurrentRateMgrPrvCd

        ' ---Parameter #7---
        ' Define the current rate_MGR_current rate_CD parameter
        Set prmCurrentRateMgr = .CreateParameter(Name:="current rate_mgr_current rate_cd", _
                                         Type:=adChar, _
                                         Direction:=adParamInput, _
                                         Size:=8, _
                                         value:=fnNullIfZLS(varIn:=CurrentRateMgr, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmCurrentRateMgr

        ' ---Parameter #8---
        ' Define the Invalid_Key output parameter, which reflects *which* foreign
        ' key violation was encountered.
        Set prmInvalid_Key = .CreateParameter(Name:="@Invalid_Key", _
                                                  Type:=adVarChar, _
                                                  Size:=255, _
                                                  Direction:=adParamOutput)
        .Parameters.Append prmInvalid_Key

        ' Do the Update
        .Execute
    End With

    '...........................................................................
    ' Refresh the Lookup recordset, re-retrieve the just-updated record so that
    ' record is *still* the current record, and load its data to the table
    ' wrapper's class properties so all table columns (including those set by
    ' the DBMS like identity columns and Last Updated columns) are up-to-date.
    '...........................................................................
    bSuccessful = GetRelativeRecord(CurrentRateCd, epdSameRecord)

    UpdateRecord = True
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
    
    ' Clean-up statements go here
    fnFreeRecordset rstTemp
    fnFreeObject prmReturnValue
    fnFreeObject prmCurrentRateCd
    fnFreeObject prmCurrentRateNm
    fnFreeObject prmCurrentRateSvsInd
    fnFreeObject prmMktvalCurrentRateCd
    fnFreeObject prmCurrentRateMgrPrvCd
    fnFreeObject prmCurrentRateMgr
    fnFreeObject prmInvalid_Key
    
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case prmReturnValue
        Case gcRES_NERR_REC_NOT_FOUND       ' 4027
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_REC_NOT_FOUND, _
                                       mcstrName & cstrCurrentProc, _
                                       "current rate Code " & RTrim$(CurrentRateCd)
            Resume PROC_EXIT
        Case gcRES_NERR_ERR_WHILE_TRYING_TO   ' 4028
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_ERR_WHILE_TRYING_TO, _
                                       mcstrName & cstrCurrentProc, _
                                       "update"
            Resume PROC_EXIT
        Case gcRES_NERR_KEY_NOT_FOUND ' 4032
            If UCase$(prmInvalid_Key) = "MKTVAL_current rate_CD" Then
                gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_KEY_NOT_FOUND, _
                                           mcstrName & cstrCurrentProc, _
                                           "Market Value current rate Cd", RTrim$(MktValCurrentRateCd), "current rate"
                Resume PROC_EXIT
            Else
                ' current rate_MGR_PRV_CD
                gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_KEY_NOT_FOUND, _
                                           mcstrName & cstrCurrentProc, _
                                           "current rate Mgr", RTrim$(CurrentRateMgrPrvCd), "Provider"
                Resume PROC_EXIT
            End If
    End Select

    ' If any other errors exist, i.e. in Err object, then let it fall through into default error handling.

    Select Case Err.Number
        Case -2147217900 ' Object not found
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_FERR_SPROC_NOT_FOUND, _
                                       mcstrName & cstrCurrentProc, _
                                       cstrSproc
            Resume PROC_EXIT
            
        ' Added -2147217887 check to fix Claims Interest bug 2454.
        Case -2147217887 ' Invalid Character Value for Cast Specification
            ' (Internally manifested in sproc as "arithmetic overflow error converting numeric to data type numeric")
            ' gcRES_NERR_NUMERIC_FLD_TOO_LARGE as Integer (4008) = One or more numeric fields are too large to be stored in the database. Your changes cannot be saved.
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_NUMERIC_FLD_TOO_LARGE, _
                                       mcstrName & cstrCurrentProc
            Resume PROC_EXIT
            
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function





'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'\                                                                  /
'|                        PRIVATE  Procedures                       |
'/                                                                  \
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/

'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub fnClearPropertyValues()
    '--------------------------------------------------------------------------
    ' Procedure:   fnClearPropertyValues
    ' Description: Sets the value of each class property that corresponds to a table column
    '              so it is wiped out. The remaining attributes of that UDTColumn structure
    '              are left intact. This proc is used when a navigation or refreshing
    '              of the Lookup recordset resulted in having no records (.BOF and/or .EOF
    '              is True) or the AbsolutePosition is invalid). Without calling this
    '              proc when those situations occur, we'd still have the previous record's
    '              value displayed.
    '
    ' Returns:     N/A
    '-----------------------------------------------------------------------------
    
    '!CUSTOMIZE!  This proc should set each class property (of type UDTColumn) that
    '             corresponds to a table column. What it is set to depends on its
    '             data type: Strings => vbNullString
    '                        Numeric => 0
    '                        Booleans => False
    '                        Dates => Now
    
    Const cstrCurrentProc          As String = "fnClearPropertyValues"

    On Error GoTo PROC_ERR

    With Me
        .CurrentRateCd = vbNullString
        .CurrentRateNm = vbNullString
        .LstUpdDtm = Now
        .LstUpdUserId = vbNullString
        .CurrentRateSvsInd = False
        .MktValCurrentRateCd = vbNullString
        .CurrentRateMgrPrvCd = vbNullString
        .CurrentRateMgr = vbNullString
    End With

    ' Also reset the saved "original" values for the Last Updated info
    m_dteLstUpdDtm_Original = Now
    m_strLstUpdUserId_Original = vbNullString
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
    
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
Private Function fnGetColMetaData(ByRef pudtCol As udtColumn, ByRef prstIn As ADODB.Recordset) As Boolean
    '--------------------------------------------------------------------------
    ' Procedure:   fnGetColMetaData
    ' Description: Given a variable that represents a table column,
    '              load its meta data to its udtColumn-defined properties,
    '              setting default values based on the data type of that
    '              table column.
    '
    '              NOTE: If the default properties aren't right for a
    '                    particular column, then the table wrapper should
    '                    have override values coded in its fnLoadColMetaData
    '                    method.
    '
    '                    **THIS** fnGETColMetaData method should be identical
    '                    in all table wrappers!
    '
    ' Returns:     True if successful; False otherwise
    '
    ' Params:      pudtCol  (in/out)  - mbr var associated with a table column
    '              prstIn   (in/out)  - recordset containing meta data, positioned
    '                                   to the row that corresponds to the
    '                                   specified column (pudtCol)
    ' Date:        04/03/2002
    '-----------------------------------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc       As String = "fnGetColMetaData"
    Const cstrTrue              As String = "TRUE"
    Const cstrNumericChar       As String = "#"
    Const cstrZeroChar          As String = "0"
    Const cstrDecimalChar       As String = "."
    Const cstrAnyCharChar       As String = "&"
    Dim strDomainNameToParse    As String
    Dim strDefaultValueToParse  As String
    Dim strEditedDefaultValue   As String

    ' Do NOT reposition the prstIn recordset as this will mess up the caller (fnGetMetaData)
    ' who is calling *this* proc for each table column (i.e. row in that recordset).


    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    ' The following data types are used within the TRS tables:
    '       Identity fields   = adInteger
    '       Char fields       = adChar
    '       Varchar fields    = adChar
    '       Dates/lst_upd_dtm = adDBTimeStamp   (e.g., lst_upd_dtm or eff_dt)
    '       Area codes        = adNumeric
    '       Indicators        = adChar          (e.g., Yes/No or other values)
    '       Percents/Factors  = adNumeric
    '       Monetary amounts  = adNumeric

    ' Certain TYPE_NAMEs are typically used to control default values or indicate
    ' certain usage, such as:
    '    a. dom_id_key........for identity columns
    '    b. dom_pct...........for percentages
    '    c. dom_ind...........for indicator fields (bound to rule_ind, this is a Y, N
    '       dom_indyn.........for indicator fields (bound to rul_indyn, this is a Y, N or Null
    '       dom_char1_ind.....for indicator fields (bound to rule_char1_ind, this is a Y or N (*not* null)
    '    d. dom_dt_nn.........for dates (not nullable)
    '       dom_dt-null.......for dates (nullable)
    '       dom_lst_upd_dtm...for dates, this sets the system date (getdate()) as the default value on an Insert
    '    e. dom_lst_upd_id....to set the user's ACF2 (suser_sname()) as the default value on an Insert
    '
    ' The following meta data appears to be available for these data types:
    '
    ' Property:  (ignored) (HasDefault) (Default (Is       (Dollar    (Decimal    (Max
    '                                    Value)  Nullable)  Positions) Positions)  Characters)
    '
    '                                                                             CHARACTER_
    '            DATETIME_  COLUMN_     COLUMN_  IS_       NUMERIC_   NUMERIC_    MAXIMUM_     DOMAIN_
    ' DATA_TYPE  PRECISION  HASDEFAULT  DEFAULT  NULLABLE  PRECISION  SCALE       LENGTH       NAME
    ' ---------  ---------  ----------  -------  --------  ---------  --------    ----------   -------
    ' adInteger     No         Yes       Yes-1       Yes       Yes       No          No         Yes-2
    ' adChar        No         Yes       Yes-1       Yes       No        No          Yes        Yes-2
    ' adNumeric     No         Yes       Yes-1       Yes       Yes       Yes         No         Yes-2
    ' adDBTimeStamp Yes        Yes       Yes-1       Yes       No        No          No         Yes-2
    '
    ' Legend:
    '   Yes-1 - COLUMN_DEF is present only when COLUMN_HASDEFAULT is present and is set to TRUE
    '   Yes-2 - TYPE_NAME is present only when a domain name has been assigned. It appears to be
    '          able to be present on any data type.
    '
    ' DATETIME_PRECISION appears to be meaningless in the Sun Life environment. Danny Khoury thinks
    ' it refers to "smalldatetime" versus <regular> "datetime", and we only use the latter. Hence,
    ' we won't bother collecting this piece of meta data for dates.
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    With pudtCol
        .ColName = UCase$(prstIn("COLUMN_NAME").value)
        .DataType = prstIn("DATA_TYPE").value

        strDefaultValueToParse = UCase$(fnZLSIfNull(prstIn("COLUMN_DEF").value))
        strDomainNameToParse = UCase$(fnZLSIfNull(prstIn("TYPE_NAME").value))

        ' ~~~~~~~~~~~~~ Set default value, if applicable ~~~~~~~~~~~~
        If IsNull(prstIn("COLUMN_DEF").value) = False Then
            .HasDefault = True
        Else
            .HasDefault = False
        End If
        ' If column has a default value, grab it. The COLUMN_DEF column is absent if
        ' there is no default
        If .HasDefault Then
            ' Char (not VarChar) fields may have a default value that is narrower than the column's width.
            ' This is okay in a SQL environment, because SQL will pad it with trailing spaces
            ' when doing an Inserts. Our GUI, however, doesn't like having a default value
            ' narrower than the column width if the column in question is used as the selectable
            ' column in a ComboBox. (The fnClearControls function fails with the equivalent
            ' of the user selecting a value that's not in the list.) Avoid this by manually
            ' adding trailing spaces here.
            If .DataType = adChar Then
                strEditedDefaultValue = prstIn("COLUMN_DEF").value
                ' Strip leading single quote
                If Left(strEditedDefaultValue, 1) = "'" And Len(strEditedDefaultValue) > 1 Then
                    strEditedDefaultValue = Right(strEditedDefaultValue, Len(strEditedDefaultValue) - 1)
                End If
                ' Strip trailing single quote
                If Right(strEditedDefaultValue, 1) = "'" And Len(strEditedDefaultValue) > 1 Then
                    strEditedDefaultValue = Left(strEditedDefaultValue, Len(strEditedDefaultValue) - 1)
                End If
                .DefaultValue = fnPadRightString(strEditedDefaultValue, _
                                            CInt(prstIn("LENGTH").value))
            Else
                .DefaultValue = prstIn("COLUMN_DEF").value
            End If
        Else
            ' "default" DefaultValue value (cute, huh?)...may be overriden in next code chunks
            .DefaultValue = Empty
        End If

        ' In some cases, the COLUMN_DEF value is not a value, per se,
        ' but SQL text that indicates which Rule or Default should be applied.
        ' If this is the case, then override the actual COLUMN_DEF value
        ' with an interpreted value equivalent to what the DBMS would have set.

        ' DEF_LST_UPD_ID is typically used for the Lst_Upd_User_Id column,
        ' indicating to set it to the logged on user.
        If InStr(1, strDefaultValueToParse, "DEF_LST_UPD_ID") > 0 Then
            .DefaultValue = gconAppActive.LastLogOnUserID
        End If

        ' DEF_LST_UPD_DTM is typically used for the Lst_Upd_Dtm column,
        ' indicating to set it to the System Date.
        '!TODO! I think this default value is meaningless but harmless.
        '       The app shouldn't even reference this field on an INSERT to ensure that the
        '       DBMS sets it itself based on the exact date/time that the INSERT occurs.
        '       On an UPDATE statement, the app *should* (and *must*) reference this
        '       column to ensure it is updated, but it should be set by the form immediately
        '       prior to issuing the UPDATE.  If my thoughts are correct, maybe the
        '       following IF should be deleted.
        If InStr(1, strDefaultValueToParse, "DEF_LST_UPD_DTM") > 0 Then
            .DefaultValue = Date
        End If

        ' All TRS tables use only the "DOM_IND" domain name for indicator columns.
        ' This domain name indicates the column must be valued Y or N. Columns
        ' bound to this domain name have a default constraint set so its default
        ' value is "N".
        ' For indicator columns, transform its default value from a literal "N" or "Y"
        ' to its corresponding Boolean value since it will typically be represented
        ' on forms as a checkbox.
        If (InStr(1, strDomainNameToParse, "DOM_IND") > 0) Then
            If .DefaultValue = "Y" Then
                .DefaultValue = True
            Else
                .DefaultValue = False
            End If
        End If

        If prstIn("NULLABLE").value Then
            .IsNullable = True
        Else
            .IsNullable = False
        End If

        ' If the data model indicates, for instance, that a column's Numeric Scale
        ' is (9,7), that means there are 9 numeric positions --excluding the decimal point--
        ' of which 7 are decimal positions....i.e. "99.9999999"

        Select Case .DataType
            Case dbDecimal
                ' Save original values (will be used to code sproc parameters).
                .NumericScale = CByte(prstIn("SCALE").value)
                .Precision = CByte(prstIn("PRECISION").value)
                ' Save interpreted equivalents. These may be overriden in fnLoadColMetaData( ).
                .DecimalPositions = CInt(prstIn("SCALE").value)
                .DollarPositions = CInt(prstIn("PRECISION").value) - .DecimalPositions
                .MaxCharacters = 0
                If Right(.ColName, 7) = "AREA_CD" Then
                    ' Area Codes require 3 numeric positions if input.
                    ' They should be displayed via a TextBox control.
                    .Format = "###"
                    .Mask = vbNullString
                    .MaxCharacters = 3
                ElseIf Right(.ColName, 4) = "_AMT" Then
                    ' These are assumed to be currency with up to 14 dollar/2 decimal positions.
                    ' These should typically  be displayed via a TextBox control.
                    ' For a Grid, .Format should be "($##,###,###,###,##0.00)"
                    ' For any other control, .Format should be "$##,###,###,###,##0.00;($##,###,###,###,##0.00)"
                    .Format = "$##,###,###,###,##0.00"
                    .Mask = vbNullString
                ElseIf Right(.ColName, 8) = "UNIT_QTY" Then
                    ' These should typically be displayed via a TextBox control.
                    ' These have 11 dollar/6 decimal positions and can be negative!
                    ' For a Grid, .Format should be "(##,###,###,##0.000000)"
                    ' For any other control, .Format should be "##,###,###,##0.000000;(##,###,###,##0.000000)"
                    .Format = "##,###,###,##0.000000;(##,###,###,##0.000000)"
                    .Mask = vbNullString
                Else
                    ' These should be displayed via a TextBox control.
                    '
                    ' NOTE: Other items like SHARES and PERCENTAGES will need to overridden
                    ' on a table-specific basis in fnLoadColMetaData( ) since there is no
                    ' easy way to recognize and process these fields.
                    '
                    ' Example: Though Percentages should end in "_PCT", the number
                    '          of decimal positions vary.
                    .Format = String(.DollarPositions, cstrNumericChar)
                    If .DecimalPositions > 0 Then
                        .Format = .Format & "." & String(.DecimalPositions, cstrNumericChar)
                    End If
                    .Mask = vbNullString
                End If
                .AllowableCharacters = ".0123456789"
                .ShouldForceToUppercase = False

            Case dbInteger
                .DecimalPositions = 0
                .DollarPositions = CInt(prstIn("PRECISION").value) - .DecimalPositions
                .MaxCharacters = 0
                .Format = String(.DollarPositions, cstrNumericChar)
                .Mask = vbNullString
                .AllowableCharacters = "0123456789"
                .ShouldForceToUppercase = False

            Case dbChar, dbVarChar
                .DecimalPositions = 0
                .DollarPositions = 0
                .MaxCharacters = CInt(prstIn("LENGTH").value)
                If (Right(.ColName, 8) = "PHON_NUM") Or (Right(.ColName, 7) = "FAX_NUM") Then
                    ' Phone and Fax Numbers should be displayed via a MaskEdBox control.
                    .Format = vbNullString
                    .Mask = "###-####"
                Else
                    ' All other adChar fields should be displayed via a TextBox control.
                    .Format = String(.MaxCharacters, cstrAnyCharChar)
                    .Mask = vbNullString
                End If
                .AllowableCharacters = _
                    "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789@!#$%^&*()-+_=~:;.,<>\|/?' "
                .ShouldForceToUppercase = False

            Case dbDateTime
                .DecimalPositions = 0
                .DollarPositions = 0
                .MaxCharacters = 0
                ' These should typically be displayed via  DTPicker control.
                ' For a Grid,     .Format should be "MM/DD/YYYY"
                ' For a DTPicker, .Format should be "MM/dd/yyy"
                ' For a TextBox, .Format should be "mm/dd/yyyy"
                .Format = "MM/dd/yyy"
                .Mask = vbNullString
                .AllowableCharacters = "0123456789/-"
                .ShouldForceToUppercase = False
            Case Else
                gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_UNEXPECTED_VAL_SELECT_CASE, _
                    mcstrName & cstrCurrentProc
                GoTo PROC_EXIT
         End Select
    End With

    fnGetColMetaData = True
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler

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
Private Function fnGetProperty(ByVal strTagIn As String) As udtColumn
    '--------------------------------------------------------------------------
    ' Procedure:   fnGetTableProperty
    ' Description: Given the tag name, it returns a pointer to the specified
    '              table class's public property.
    '
    ' Params:      N/A
    '    strTagIn  (in)     A string containing the Property Name, typically
    '                       from the form control's Tag property
    '
    ' Returns:     A pointer to the public property
    '-----------------------------------------------------------------------------
    
    '!CUSTOMIZE!  There should be one Case statement for each table column.
    '             Each Case statement should reference a class property that
    '             corresponds to a table column, and the fnGetProperty return
    '             value should be set to the private variable (of type UDTColumn)
    '             that corresponds to that table column/class property.
    
    Const cstrCurrentProc       As String = "fnGetProperty"
    
    On Error GoTo PROC_ERR

    strTagIn = UCase$(strTagIn)
    Select Case strTagIn
        Case "CurrentRateCd"
            fnGetProperty = m_strCurrentRateCd
        Case "CurrentRateNm"
            fnGetProperty = m_strCurrentRateNm
        Case "LSTUPDDTM"
            fnGetProperty = m_dteLstUpdDtm
        Case "LSTUPDUSERID"
            fnGetProperty = m_strLstUpdUserId
        Case "CurrentRateSvsInd"
            fnGetProperty = m_bCurrentRateSVSInd
        Case "MKTVALCurrentRateCd"
            fnGetProperty = m_strMktValCurrentRateCd
        Case "CurrentRateMgrPrvCd"
            fnGetProperty = m_strCurrentRateMgrPrvCd
        Case "CurrentRateMgr"
            fnGetProperty = m_strCurrentRateMgrCurrentRateCd
        Case Else
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_UNEXPECTED_VAL_SELECT_CASE, _
                mcstrName & cstrCurrentProc
            GoTo PROC_EXIT
    End Select
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
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
Private Function fnLoadColMetaData() As Boolean
    '--------------------------------------------------------------------------
    ' Procedure:   fnLoadColMetaData
    ' Description: For each table column, collect its meta data
    '              and load it to its corresponding UDT.
    '
    '              NOTE: Here is where you would override that meta data if it's
    '                    warranted for a given column. For instance, a column
    '                    that holds a numeric column that is allowed to have a
    '                    negative value should have its default
    '                    a. .AllowableCharacters property overriden to allow a "-" sign
    '                    b. .Format property overriden to specify that a negative
    '                       value should be enclosed within parentheses.
    '
    ' See fnGetColMetaData( ) in this module to see what the defaults are set to
    ' for a given data type.

    '
    ' Returns:     True if successful; False otherwise
    ' Params:      None
    ' Date:        01/07/2002
    '-----------------------------------------------------------------------------
    
    '!CUSTOMIZE!  The call to GetMetaData_Column should pass in the correct table name.
    '
    '             In the first Select Case block,
    '             * there should be one Case statement for each table column.
    '             * Each Case statement should reference the constant that
    '               corresponds to the table column, and it should pass the private
    '               variable (of type UDTColumn) that corresponds to that
    '               table column/class property to fnGetColMetaData.
    '
    '             In the second Select Case block (getting info re: primary keys),
    '             * there should be one Case statement for each table column.
    '             * Each Case statement should reference the constant that
    '               corresponds to the table column, and it should call the IsKey
    '               method of the priva variable (of type UDTColumn) that
    '               corresponds to that table column/class property.


    Const cstrCurrentProc       As String = "fnLoadColMetaData"
    Dim rstMetaData As ADODB.Recordset

    On Error GoTo PROC_ERR

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '       Get meta data, like nullability, data type, default value, etc.
    '       and set override values
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    m_adwADO.GetMetaData_Columns "current_rate_t", rstMetaData

    ' NOTES:
    '    a. Phone Numbers are Char fields, but we have to override the meta data to ensure
    '       only numbers can be input to them.
    '    b. If the GUI is imposing a default value where the DBMS does not have one,
    '       you must set .HasDefault to True and *then* set the .DefaultValue value.
    '
    '       WARNING: Char (not VarChar) fields may have a default value that is narrower than the column's width.
    '                This is okay in a SQL environment, because SQL will pad it with trailing spaces
    '                when doing an Inserts. Our GUI, however, doesn't like having a default value
    '                narrower than the column width if the column in question is used as the selectable
    '                column in a ComboBox. (The fnClearControls function fails with the equivalent
    '                of the user selecting a value that's not in the list.) Avoid this by making sure
    '                any override default value you specify here is as wide as the column to which
    '                it should be applied.

    With rstMetaData
        Do Until .EOF
            ' Make sure SELECT CASE lists all table columns, including the LST_UPD_xxx!
            Select Case UCase$(rstMetaData("COLUMN_NAME").value)
                Case mcstrCurrentRateCd
                    fnGetColMetaData m_strCurrentRateCd, rstMetaData
                    ' Per the screen spec, the following overrides
                    ' the DBMS max width of 8 until such time
                    ' that the Trades import files can all handle
                    ' a full-width/8-character current rate code.
                    m_strCurrentRateCd.MaxCharacters = 4
                    ' Force to be uppercase, per screen spec
                    m_strCurrentRateCd.ShouldForceToUppercase = True
                Case mcstrCurrentRateNm
                    fnGetColMetaData m_strCurrentRateNm, rstMetaData
                Case mcstrLstUpdDtm
                    fnGetColMetaData m_dteLstUpdDtm, rstMetaData
                Case mcstrLstUpdUserId
                    fnGetColMetaData m_strLstUpdUserId, rstMetaData
                Case mcstrCurrentRateSvsInd
                    fnGetColMetaData m_bCurrentRateSVSInd, rstMetaData
                Case mcstrMktvalCurrentRateCd
                    fnGetColMetaData m_strMktValCurrentRateCd, rstMetaData
                Case mcstrCurrentRateMgrPrvCd
                    fnGetColMetaData m_strCurrentRateMgrPrvCd, rstMetaData
                Case mcstrCurrentRateMgrFundCd
                    fnGetColMetaData m_strCurrentRateMgrCurrentRateCd, rstMetaData
                    ' Force to be uppercase, per screen spec
                    m_strCurrentRateMgrCurrentRateCd.ShouldForceToUppercase = True
                Case mcstrLstUpdDtm
                    fnGetColMetaData m_dteLstUpdDtm, rstMetaData
                Case mcstrLstUpdUserId
                    fnGetColMetaData m_strLstUpdUserId, rstMetaData
                Case Else
                    gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_UNEXPECTED_VAL_SELECT_CASE, _
                        mcstrName & cstrCurrentProc
                    GoTo PROC_EXIT
            End Select
            m_adwADO.MoveNext rstMetaData
        Loop
        .Close        ' Close now so the recordset can be reused
    End With

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '             Get meta data concerning which columns are key fields
    '
    '     If a given COLUMN_NAME is returned in the recordset, it is a primary key.
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    m_adwADO.GetMetaData_PrimaryKeys "current_rate_t", rstMetaData

    With rstMetaData
        Do Until .EOF
            ' The SELECT CASE should list all table columns
            ' (though you could skip the LST_UPD_xxx columns if you change the Case Else,
            '  since these would never be a key)
            Select Case UCase$(rstMetaData("COLUMN_NAME").value)
                Case mcstrCurrentRateCd
                    m_strCurrentRateCd.IsKey = True
                Case mcstrCurrentRateNm
                    m_strCurrentRateNm.IsKey = True
                Case mcstrLstUpdDtm
                    m_dteLstUpdDtm.IsKey = True
                Case mcstrLstUpdUserId
                    m_strLstUpdUserId.IsKey = True
                Case mcstrCurrentRateSvsInd
                    m_bCurrentRateSVSInd.IsKey = True
                Case mcstrMktvalCurrentRateCd
                    m_strMktValCurrentRateCd.IsKey = True
                Case mcstrCurrentRateMgrPrvCd
                    m_strCurrentRateMgrPrvCd.IsKey = True
                Case mcstrCurrentRateMgrFundCd
                    m_strCurrentRateMgrCurrentRateCd.IsKey = True
                Case Else
                    gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_UNEXPECTED_VAL_SELECT_CASE, _
                        mcstrName & cstrCurrentProc
                    GoTo PROC_EXIT
            End Select
            m_adwADO.MoveNext rstMetaData
        Loop
        .Close        ' Close now so the recordset can be reused
    End With

    fnLoadColMetaData = True
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler

    ' Clean-up statements go here
    fnFreeRecordset rstMetaData

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
Private Function fnSelectRecord(ByVal strKey1 As String) As ADODB.Recordset
    '--------------------------------------------------------------------------
    ' Procedure:   fnSelectRecord
    ' Description: Selects a single record based on the value(s) in the
    '              properties that correspond to the table's key(s)
    '
    '              NOTE: For each table key, there should be a parameter
    '                    of the appropriate data type!
    '
    ' Parameters:
    '     strKey1 (in) - the key to the table that should be retrieved
    '
    ' Returns:     A disconnected ADODB.Recordset containing all table columns
    '              for the specified key
    '-----------------------------------------------------------------------------
    
    '!CUSTOMIZE!  This proc and all calls to it must be customized to reflect
    '             one parameter for each key column of the table. Make sure the
    '             parameter is defined to be of the right data type. Also,
    '             the way the recordset's .Find property is set must be changed
    '             to reflect each key column so the right record will be located.
    '             Also make sure that the substitution values passed to
    '             SaveAppSpecificError are correct and TRIM'd if appropriate.
    
    Const cstrCurrentProc          As String = "fnSelectRecord"
    Const cstrSproc                As String = "dbo.proc_current rate_select"  ' Stored procedure to execute
    Dim rstTemp                    As ADODB.Recordset
    Dim prmReturnValue             As ADODB.Parameter
    Dim prmCurrentRateCd           As ADODB.Parameter

    On Error GoTo PROC_ERR

    If Not (m_adwADO.CommandSetSproc(cstrSproc)) Then
        GoTo PROC_EXIT
    End If

    ' For Char/VarChar fields,
    '     * Use fnNullIfZLS to ensure Nulls are appropriately handled.
    '     * Do *not* set the optional 2nd parameter to fnNullIfZLS to True.
    ' For numeric fields,
    '     * Use fnNullIfZero to ensure Nulls are appropriately handled.
    ' For Y/N fields,
    '     * Use fnBoolToYN to ensure True/False is appropriately translated.

    With m_adwADO.ADOCommand
        ' ---Parameter #1---
        ' Define the return value that represents the error code (i.e. reason) why
        ' the stored procedure failed.
        Set prmReturnValue = .CreateParameter(Name:="@return_value", _
                                              Type:=adInteger, _
                                              Direction:=adParamReturnValue, _
                                              value:=Null)
        .Parameters.Append prmReturnValue

        ' ---Parameter #2---
        ' Define the current rate_CD parameter
        Set prmCurrentRateCd = .CreateParameter(Name:="current rate_cd", _
                                         Type:=adChar, _
                                         Direction:=adParamInput, _
                                         Size:=8, _
                                         value:=fnNullIfZLS(varIn:=strKey1, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmCurrentRateCd

        Set rstTemp = .Execute()
    End With
    
    rstTemp.ActiveConnection = Nothing
    Set fnSelectRecord = rstTemp
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
    
    ' Clean-up statements go here
    
    ' Do *not* do "fnFreeRecordset rstTemp" since this will cause the recordset returned
    ' by this function to be wiped out as well!
    fnFreeObject prmReturnValue
    fnFreeObject prmCurrentRateCd
    
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case prmReturnValue
        Case gcRES_NERR_REC_NOT_FOUND       ' 4027
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_REC_NOT_FOUND, _
                                       mcstrName & cstrCurrentProc, _
                                       "current rate Code " & RTrim$(strKey1)
            Resume PROC_EXIT
        Case gcRES_NERR_ERR_WHILE_TRYING_TO   ' 4028
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_ERR_WHILE_TRYING_TO, _
                                       mcstrName & cstrCurrentProc, _
                                       "locate"
            Resume PROC_EXIT
    End Select

    ' If any other errors exist, i.e. in Err object, then let it fall through into default error handling.

    Select Case Err.Number
        Case -2147217900 ' Object not found
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_FERR_SPROC_NOT_FOUND, _
                                       mcstrName & cstrCurrentProc, _
                                       cstrSproc
            Resume PROC_EXIT
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function
