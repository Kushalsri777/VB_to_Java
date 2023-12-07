VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ctclmClaim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'!TODO! Add methods (or screen code??) to load m_rstLookup to appropriate comboboxes
'       and translate selected cbo entry to corresponding m_rstLookup row.
'!TODO! GUI Code - make SSN required if LOB = group; optional otherwise
'!TODO! Add code to set/get clm_num, for group, based on clm_pol_num and clm_insd_ssn_num;
'       otherwise set it to clm_pol_num only.  Changes to Property Set/Get methods??

'--------------------------------------------------------------------------
' Procedure:   ctclmClaim
' Description: Provides properties and methods to support the claim_t table values
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
'                    clm_num          ClmNum          m_strClmNum        mcstrClmNum
'                    clm_proof_dt     ClmProofDt      m_dteClmProofDt    mcstrClmProofDt
'
'
'                 NOTE also that navigation should be done via the **ADO Wrapper's**
'                 navigation methods instead of directly referencing the navigation
'                 methods on a ADODB.Recordset object!
'
' Revisions:
'   Modified:     Berry Kropiwka - Added Compact Calc - 2019-09-27
'
'
' Procedures  :
'   Private     Class_Initialize()
'   Private     Class_Terminate()
'   Public      Property Get AdmnSystCd() As String
'   Public      Property Let AdmnSystCd(ByVal strValue As String)
'   Public      Property Get AllowableCharacters(ByVal strTagIn As String) As String
'   Public      Property Get ClmCompactClcnInd() As Boolean
'   Public      Property Let ClmCompactClcnInd(ByVal bValue As Boolean)
'   Public      Property Get ClmForResDthInd() As Boolean
'   Public      Property Let ClmForResDthInd(ByVal bValue As Boolean)
'   Public      Property Get ClmId() As Long
'   Public      Property Let ClmId(ByVal lngValue As Long)
'   Public      Property Get ClmInsdDthDt() As Date
'   Public      Property Let ClmInsdDthDt(ByVal dteValue As Date)
'   Public      Property Get ClmInsdFirstNm() As String
'   Public      Property Let ClmInsdFirstNm(ByVal strValue As String)
'   Public      Property Get ClmInsdLastNm() As String
'   Public      Property Let ClmInsdLastNm(ByVal strValue As String)
'   Public      Property Get ClmInsdSsnNum() As String
'   Public      Property Let ClmInsdSsnNum(ByVal strValue As String)
'   Public      Property Get ClmNum() As String
'   Public      Property Let ClmNum(ByVal strValue As String)
'   Public      Property Get ClmPolNum() As String
'   Public      Property Let ClmPolNum(ByVal strValue As String)
'   Public      Property Get ClmProofDt() As Date
'   Public      Property Let ClmProofDt(ByVal dteValue As Date)
'   Public      Property Get ClmTotClmPdAmt() As Double
'   Public      Property Let ClmTotClmPdAmt(ByVal dblValue As Double)
'   Public      Property Get ClmTotDthbPmtAmt() As Double
'   Public      Property Let ClmTotDthbPmtAmt(ByVal dblValue As Double)
'   Public      Property Get ClmTotIntAmt() As Double
'   Public      Property Let ClmTotIntAmt(ByVal dblValue As Double)
'   Public      Property Get ClmTotWthldAmt() As Double
'   Public      Property Let ClmTotWthldAmt(ByVal dblValue As Double)
'   Public      Property Get CurrentLookupRecordNumber() As Long
'   Public      Property Get DecimalPositions(ByVal strTagIn As String) As Integer
'   Public      Property Get DefaultValue(ByVal strTagIn As String) As Variant
'   Public      Property Get DollarPositions(ByVal strTagIn As String) As Integer
'   Public      Property Get Format(ByVal strTagIn As String) As String
'   Public      Property Get InsdDthResStCd() As String
'   Public      Property Let InsdDthResStCd(ByVal strValue As String)
'   Public      Property Get IsKey(ByVal strTagIn As String) As Boolean
'   Public      Property Get IsNullable(ByVal strTagIn As String) As Boolean
'   Public      Property Get IssStCd() As String
'   Public      Property Let IssStCd(ByVal strValue As String)
'   Public      Property Get LookupData() As ADODB.Recordset
'   Public      Property Get LookupData_Claim() As ADODB.Recordset
'   Public      Property Get LookupData_Name() As ADODB.Recordset
'   Public      Property Get LookupData_SSN() As ADODB.Recordset
'   Public      Property Get LookupIsAtBOF() As Boolean
'   Public      Property Get LookupIsAtEOF() As Boolean
'   Public      Property Get LookupRecordCount() As Long
'   Public      Property Get LstUpdtDtm() As Date
'   Public      Property Let LstUpdtDtm(ByVal NewValue As String)
'   Public      Property Get LstUpdtUserId() As String
'   Public      Property Let LstUpdtUserId(ByVal strValue As String)
'   Public      Property Get Mask(ByVal strTagIn As String) As String
'   Public      Property Get MaxCharacters(ByVal strTagIn As String) As Long
'   Public      Property Get PycoTypCd() As String
'   Public      Property Let PycoTypCd(ByVal strValue As String)
'   Public      Property Get ShouldForceToUppercase(ByVal strTagIn As String) As Boolean
'   Public      AddRecord() as Boolean
'   Public      CheckForAnotherUsersChanges(ByVal lngWhatOperation As enumWhatOperationIsBeingAttempted, _
'                   ByRef strACF2 As String) As Long
'   Public      DeleteRecord() As Boolean
'   Public      GetClmIdFromClmNum(ByVal strClmNum As String) As Integer
'   Public      GetClmNumFromClmID(ByVal lngClmID As Long) As Variant
'   Public      GetLobCdFromAdmnSystCd(ByVal strAdmnSystCd As String) As String
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
Private Const mcstrName As String = "ctclmClaim."

'...............................................................................................
'!CUSTOMIZE!
' These are the private variables corresponding to PUBLIC properties.
' There should be one (of type udtColumn) for each column in the table that this class accesses.
'...............................................................................................
Private m_strAdmnSystCd                 As udtColumn
Private m_lngClmId                      As udtColumn
Private m_dteClmInsdDthDt               As udtColumn
Private m_strClmInsdFirstNm             As udtColumn
Private m_strClmInsdLastNm              As udtColumn
Private m_strClmInsdSsnNum              As udtColumn
Private m_strClmNum                     As udtColumn
Private m_strClmPolNum                  As udtColumn
Private m_dteClmProofDt                 As udtColumn
Private m_varClmTotClmPdAmt             As udtColumn
Private m_varClmTotDthbPmtAmt           As udtColumn
Private m_varClmTotIntAmt               As udtColumn
Private m_varClmTotWthldAmt             As udtColumn
Private m_strInsdDthResStCd             As udtColumn
Private m_strIssStCd                    As udtColumn
Private m_dteLstUpdtDtm                 As udtColumn
Private m_strLstUpdtUserId              As udtColumn
Private m_strPycoTypCd                  As udtColumn
Private m_bClmForResDthInd              As udtColumn
Private m_bClmCompactClcnInd                 As udtColumn

'...............................................................................................
'!CUSTOMIZE!
' Create one Const for each column in the table, defining the table column to which it refers.
'...............................................................................................
Private Const mcstrAdmnSystCd           As String = "ADMN_SYST_CD"
Private Const mcstrClmId                As String = "CLM_ID"
Private Const mcstrClmInsdDthDt         As String = "CLM_INSD_DTH_DT"
Private Const mcstrClmInsdFirstNm       As String = "CLM_INSD_FIRST_NM"
Private Const mcstrClmInsdLastNm        As String = "CLM_INSD_LAST_NM"
Private Const mcstrClmInsdSsnNum        As String = "CLM_INSD_SSN_NUM"
Private Const mcstrClmNum               As String = "CLM_NUM"
Private Const mcstrClmPolNum            As String = "CLM_POL_NUM"
Private Const mcstrClmProofDt           As String = "CLM_PROOF_DT"
Private Const mcstrClmTotClmPdAmt       As String = "CLM_TOT_CLM_PD_AMT"
Private Const mcstrClmTotDthbPmtAmt     As String = "CLM_TOT_DTHB_PMT_AMT"
Private Const mcstrClmTotIntAmt         As String = "CLM_TOT_INT_AMT"
Private Const mcstrClmTotWthldAmt       As String = "CLM_TOT_WTHLD_AMT"
Private Const mcstrInsdDthResStCd       As String = "INSD_DTH_RES_ST_CD"
Private Const mcstrIssStCd              As String = "ISS_ST_CD"
Private Const mcstrLstUpdtDtm           As String = "LST_UPDT_DTM"
Private Const mcstrLstUpdtUserId        As String = "LST_UPDT_USER_ID"
Private Const mcstrPycoTypCd            As String = "PYCO_TYP_CD"
Private Const mcstrClmForResDthInd      As String = "CLM_FOR_RES_DTH_IND"
Private Const mcstrClmCompactClcnInd         As String = "CLM_COMPACT_CLCN_IND"

'...............................................................................................
' Other private variables that do NOT correspond to PUBLIC properties.
'...............................................................................................
' m_adwADO is a private instantiation of the ADO Wrapper, used to do ADO things like
' navigation, executing a stored procedure, etc.
Private m_adwADO                        As cadwADOWrapper

' The next 2 vars (m_dteLstUpdtDtm_Original and m_strLstUpdtUserId_Original) are used by
' the CheckForAnotherUsersChanges method to determine if another user affected the
' record since *this* user originally retrieved the record.
Private m_dteLstUpdtDtm_Original         As Date
Private m_strLstUpdtUserId_Original      As String

' m_rstLookup contains selected columns for each row in the table and is used by the form
' to populate its 3 Lookup VSFlexGrid controls that the user uses to hop directly to a desired record.
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
        GetSingleRecord m_rstLookup.Fields(mcstrClmId).value
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
Public Property Get AdmnSystCd() As String
    ' **************************************************************************
    ' Function  : Property Get AdmnSystCd
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : string
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get AdmnSystCd"
    On Error GoTo PROC_ERR

    AdmnSystCd = CStr(m_strAdmnSystCd.value)
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
Public Property Let AdmnSystCd(ByVal strValue As String)
    ' **************************************************************************
    ' Function  : Property Let AdmnSystCd
    ' Purpose   : Assigns new Value to property
    ' Parameters: ByVal NewValue As String
    ' Returns   :
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let AdmnSystCd"
    On Error GoTo PROC_ERR

    m_strAdmnSystCd.value = strValue
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
Public Property Get ClmForResDthInd() As Boolean
    ' **************************************************************************
    ' Function  : Property Get ClmForResDthInd
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : Boolean
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get ClmForResDthInd"
    On Error GoTo PROC_ERR

    ' Note that this field is stored as a Y or N.
    ClmForResDthInd = CBool(m_bClmForResDthInd.value)
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
Public Property Let ClmForResDthInd(ByVal bValue As Boolean)
    ' **************************************************************************
    ' Function  : Property Let ClmForResDthInd
    ' Purpose   : Assigns new Value to property
    ' Parameters: ByVal bValue As Boolean
    ' Returns   :
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let ClmForResDthInd"
    On Error GoTo PROC_ERR

    ' Note that this field is stored as a Y or N.
    m_bClmForResDthInd.value = bValue
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
Public Property Get ClmCompactClcnInd() As Boolean
    ' **************************************************************************
    ' Function  : Property Get ClmCompactClcnInd
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : Boolean
    ' Created   : Berry Kropiwka 2019-09-27
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get ClmCompactClcnInd"
    On Error GoTo PROC_ERR

    ' Note that this field is stored as a T or F.
    ClmCompactClcnInd = CBool(m_bClmCompactClcnInd.value)
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
Public Property Let ClmCompactClcnInd(ByVal bValue As Boolean)
    ' **************************************************************************
    ' Function  : Property Let ClmCompactClcnInd
    ' Purpose   : Assigns new Value to property
    ' Parameters: ByVal bValue As Boolean
    ' Returns   :
    ' Created   : Berry Kropiwka 2019-09-27
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let ClmCompactClcnInd"
    On Error GoTo PROC_ERR

    ' Note that this field is stored as a T or F.
    m_bClmCompactClcnInd.value = bValue
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
Public Property Get ClmId() As Long
    ' **************************************************************************
    ' Function  : Property Get ClmId
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : double
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get ClmId"
    On Error GoTo PROC_ERR

    ClmId = CLng(m_lngClmId.value)
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
Public Property Let ClmId(ByVal lngValue As Long)
    ' **************************************************************************
    ' Function  : Property Let ClmId
    ' Purpose   : Assigns new Value to property
    ' Parameters: ByVal lngValue As Long
    ' Returns   :
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let ClmId"
    On Error GoTo PROC_ERR

    m_lngClmId.value = lngValue
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
Public Property Get ClmInsdDthDt() As Date
    ' **************************************************************************
    ' Function  : Property Get ClmInsdDthDt
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : string
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get ClmInsdDthDt"
    On Error GoTo PROC_ERR

    ClmInsdDthDt = CDate(m_dteClmInsdDthDt.value)
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
Public Property Let ClmInsdDthDt(ByVal dteValue As Date)
    ' **************************************************************************
    ' Function  : Property Let ClmInsdDthDt
    ' Purpose   : Assigns new Value to property
    ' Parameters: ByVal NewValue As String
    ' Returns   :
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let ClmInsdDthDt"
    On Error GoTo PROC_ERR

    m_dteClmInsdDthDt.value = dteValue
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
Public Property Get ClmInsdFirstNm() As String
    ' **************************************************************************
    ' Function  : Property Get ClmInsdFirstNm
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : string
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get ClmInsdFirstNm"
    On Error GoTo PROC_ERR

    ClmInsdFirstNm = CStr(m_strClmInsdFirstNm.value)
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
Public Property Let ClmInsdFirstNm(ByVal strValue As String)
    ' **************************************************************************
    ' Function  : Property Let ClmInsdFirstNm
    ' Purpose   : Assigns new Value to property
    ' Parameters: ByVal NewValue As String
    ' Returns   :
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let ClmInsdFirstNm"
    On Error GoTo PROC_ERR

    m_strClmInsdFirstNm.value = strValue
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
Public Property Get ClmInsdLastNm() As String
    ' **************************************************************************
    ' Function  : Property Get ClmInsdLastNm
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : string
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get ClmInsdLastNm"
    On Error GoTo PROC_ERR

    ClmInsdLastNm = CStr(m_strClmInsdLastNm.value)
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
Public Property Let ClmInsdLastNm(ByVal strValue As String)
    ' **************************************************************************
    ' Function  : Property Let ClmInsdLastNm
    ' Purpose   : Assigns new Value to property
    ' Parameters: ByVal NewValue As String
    ' Returns   :
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let ClmInsdLastNm"
    On Error GoTo PROC_ERR

    m_strClmInsdLastNm.value = strValue
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
Public Property Get ClmInsdSsnNum() As String
    ' **************************************************************************
    ' Function  : Property Get ClmInsdSsnNum
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : string
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get ClmInsdSsnNum"
    On Error GoTo PROC_ERR

    ClmInsdSsnNum = CStr(m_strClmInsdSsnNum.value)
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
Public Property Let ClmInsdSsnNum(ByVal strValue As String)
    ' **************************************************************************
    ' Function  : Property Let ClmInsdSsnNum
    ' Purpose   : Assigns new Value to property
    ' Parameters: ByVal NewValue As String
    ' Returns   :
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let ClmInsdSsnNum"
    On Error GoTo PROC_ERR

    m_strClmInsdSsnNum.value = strValue
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
Public Property Get ClmNum() As String
    ' **************************************************************************
    ' Function  : Property Get ClmNum
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : string
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get ClmNum"
    On Error GoTo PROC_ERR

    ClmNum = CStr(m_strClmNum.value)
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
Public Property Let ClmNum(ByVal strValue As String)
    ' **************************************************************************
    ' Function  : Property Let ClmNum
    ' Purpose   : Assigns new Value to property
    ' Parameters: ByVal NewValue As String
    ' Returns   :
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let ClmNum"
    On Error GoTo PROC_ERR

    m_strClmNum.value = strValue
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
Public Property Get ClmPolNum() As String
    ' **************************************************************************
    ' Function  : Property Get ClmPolNum
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : string
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get ClmPolNum"
    On Error GoTo PROC_ERR

    ClmPolNum = CStr(m_strClmPolNum.value)
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
Public Property Let ClmPolNum(ByVal strValue As String)
    ' **************************************************************************
    ' Function  : Property Let ClmPolNum
    ' Purpose   : Assigns new Value to property
    ' Parameters: ByVal NewValue As String
    ' Returns   :
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let ClmPolNum"
    On Error GoTo PROC_ERR

    m_strClmPolNum.value = strValue
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
Public Property Get ClmProofDt() As Date
    ' **************************************************************************
    ' Function  : Property Get ClmProofDt
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : string
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get ClmProofDt"
    On Error GoTo PROC_ERR

    ClmProofDt = CDate(m_dteClmProofDt.value)
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
Public Property Let ClmProofDt(ByVal dteValue As Date)
    ' **************************************************************************
    ' Function  : Property Let ClmProofDt
    ' Purpose   : Assigns new Value to property
    ' Parameters: ByVal NewValue As String
    ' Returns   :
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let ClmProofDt"
    On Error GoTo PROC_ERR

    m_dteClmProofDt.value = dteValue
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
Public Property Get ClmTotClmPdAmt() As Variant
    ' **************************************************************************
    ' Function  : Property Get ClmTotClmPdAmt
    ' Purpose   : Retrieves current value (which could be NULL) from property
    ' Parameters: N/A
    ' Returns   : double
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get ClmTotClmPdAmt"
    On Error GoTo PROC_ERR

    ClmTotClmPdAmt = m_varClmTotClmPdAmt.value
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
Public Property Let ClmTotClmPdAmt(ByVal varValue As Variant)
    ' **************************************************************************
    ' Function  : Property Let ClmTotClmPdAmt
    ' Purpose   : Assigns new Value (which could be NULL) to property
    ' Parameters: ByVal varValue
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let ClmTotClmPdAmt"
    On Error GoTo PROC_ERR

    m_varClmTotClmPdAmt.value = varValue
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
Public Property Get ClmTotDthbPmtAmt() As Variant
    ' **************************************************************************
    ' Function  : Property Get ClmTotDthbPmtAmt
    ' Purpose   : Retrieves current value (which could be NULL) from property
    ' Parameters: N/A
    ' Returns   : double
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get ClmTotDthbPmtAmt"
    On Error GoTo PROC_ERR

    ClmTotDthbPmtAmt = m_varClmTotDthbPmtAmt.value
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
Public Property Let ClmTotDthbPmtAmt(ByVal varValue As Variant)
    ' **************************************************************************
    ' Function  : Property Let ClmTotDthbPmtAmt
    ' Purpose   : Assigns new Value (which could be NULL) to property
    ' Parameters: ByVal varValue
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let ClmTotDthbPmtAmt"
    On Error GoTo PROC_ERR

    m_varClmTotDthbPmtAmt.value = varValue
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
Public Property Get ClmTotIntAmt() As Variant
    ' **************************************************************************
    ' Function  : Property Get ClmTotIntAmt
    ' Purpose   : Retrieves current value (which could be NULL) from property
    ' Parameters: N/A
    ' Returns   : double
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get ClmTotIntAmt"
    On Error GoTo PROC_ERR

    ClmTotIntAmt = m_varClmTotIntAmt.value
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
Public Property Let ClmTotIntAmt(ByVal varValue As Variant)
    ' **************************************************************************
    ' Function  : Property Let ClmTotIntAmt
    ' Purpose   : Assigns new Value (which could be NULL) to property
    ' Parameters: ByVal varValue
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let ClmTotIntAmt"
    On Error GoTo PROC_ERR

    m_varClmTotIntAmt.value = varValue
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
Public Property Get ClmTotWthldAmt() As Variant
    ' **************************************************************************
    ' Function  : Property Get ClmTotWthldAmt
    ' Purpose   : Retrieves current value (which could be NULL) from property
    ' Parameters: N/A
    ' Returns   : double
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get ClmTotWthldAmt"
    On Error GoTo PROC_ERR

    ClmTotWthldAmt = m_varClmTotWthldAmt.value
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
Public Property Let ClmTotWthldAmt(ByVal varValue As Variant)
    ' **************************************************************************
    ' Function  : Property Let ClmTotWthldAmt
    ' Purpose   : Assigns new Value (which could be NULL) to property
    ' Parameters: ByVal varValue
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let ClmTotWthldAmt"
    On Error GoTo PROC_ERR

    m_varClmTotWthldAmt.value = varValue
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
    '             that might be necessary.
    '
    ' Parameters: N/A
    ' Returns   : Variant
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get DefaultValue"
    On Error GoTo PROC_ERR
    Dim udtTemp As udtColumn
    Dim varTemp As Variant

    udtTemp = fnGetProperty(strTagIn)    ' for efficiency

    If udtTemp.HasDefault Then
        varTemp = udtTemp.DefaultValue

        ' Strip leading single quote
        If Left$(varTemp, 1) = "'" And Len(varTemp) > 1 Then
            varTemp = Right$(varTemp, Len(varTemp) - 1)
        End If
        ' Strip trailing single quote
        If Right$(varTemp, 1) = "'" And Len(varTemp) > 1 Then
            varTemp = Left$(varTemp, Len(varTemp) - 1)
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
Public Property Get InsdDthResStCd() As String
    ' **************************************************************************
    ' Function  : Property Get InsdDthResStCd
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : string
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get InsdDthResStCd"
    On Error GoTo PROC_ERR

    InsdDthResStCd = CStr(m_strInsdDthResStCd.value)
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
Public Property Let InsdDthResStCd(ByVal strValue As String)
    ' **************************************************************************
    ' Function  : Property Let InsdDthResStCd
    ' Purpose   : Assigns new Value to property
    ' Parameters: ByVal NewValue As String
    ' Returns   :
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let InsdDthResStCd"
    On Error GoTo PROC_ERR

    m_strInsdDthResStCd.value = strValue
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
Public Property Get IssStCd() As String
    ' **************************************************************************
    ' Function  : Property Get IssStCd
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : string
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get IssStCd"
    On Error GoTo PROC_ERR

    IssStCd = CStr(m_strIssStCd.value)
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
Public Property Let IssStCd(ByVal strValue As String)
    ' **************************************************************************
    ' Function  : Property Let IssStCd
    ' Purpose   : Assigns new Value to property
    ' Parameters: ByVal NewValue As String
    ' Returns   :
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let IssStCd"
    On Error GoTo PROC_ERR

    m_strIssStCd.value = strValue
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
Public Property Get LookupData_Claim() As Variant
    '--------------------------------------------------------------------------
    ' Procedure:   Get LookupData_Claim
    ' Description: Return an array containing just the desired columns
    '              that should be populated in the Lookup fpCombo control
    ' Returns:     variant array
    '-----------------------------------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc       As String = "Get_LookupData_Claim"
    Dim aRows()                 As Variant
    Dim strOriginalClmNum       As String
    
    ' The .GetRows method changes the positioning within the m_rstLookup, which
    ' causes the "record x of y" label to incorrectly get set to the
    ' .RecordCount value. So, save the key of the current record, issue the
    ' .GetRows and then navigate back to the original position within the rst.
    
    '!TODO! 06/28/03 BAW - The following code needs to be tightened up.
    ' If the m_rstLookup recordset is empty, the first line within the WITH
    ' block gets runtime error 3021 "Either BOF or EOF is True, or the current
    ' record has been deleted. Requested operation requires a current record."
    ' so the logic should be improved to return an initialized array (or one
    ' with a pseudo blank row) if the m_rstLookup is at BOF or EOF
    ' so the caller (fnLoadLpcLookup) will work.
    With m_rstLookup
        strOriginalClmNum = .Fields(mcstrClmNum).value
    
        ' m_rstLookup is already sorted in the order we want: clm_num
        aRows = .GetRows(Rows:=adGetRowsRest, Start:=adBookmarkFirst)
    End With
        
    GetRelativeRecord strOriginalClmNum, epdSameRecord
            
    LookupData_Claim = aRows
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
    
    ' Clean-up statements go here
    Erase aRows
    
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
Public Property Get LookupData_Name() As Variant
    '--------------------------------------------------------------------------
    ' Procedure:   Get LookupData_Name
    ' Description: Retrieves a sorted ADO Recordset, loads it to an array
    '              containing for fast population to the Name Lookup
    '              fpCombo control on the Insured screen
    ' Returns:     variant array
    '-----------------------------------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc       As String = "Get_LookupData_Name"
    Const cstrSproc             As String = "dbo.proc_claim_lu_select2"
    Dim aRows()                 As Variant
    Dim rstTemp                 As ADODB.Recordset
    Dim prmReturnValue          As ADODB.Parameter

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

        Set rstTemp = .Execute()
    End With
    
    aRows = rstTemp.GetRows(Rows:=adGetRowsRest, Start:=adBookmarkFirst)
    
    LookupData_Name = aRows
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
    
    ' Clean-up statements go here
    Erase aRows
    fnFreeObject rstTemp
    fnFreeObject prmReturnValue
    
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
Public Property Get LookupData_SSN() As Variant
    '--------------------------------------------------------------------------
    ' Procedure:   Get LookupData_SSN
    ' Description: Retrieves a sorted ADO Recordset, loads it to an array
    '              containing for fast population to the SSN Lookup
    '              fpCombo control on the Insured screen
    ' Returns:     variant array
    '-----------------------------------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc       As String = "Get_LookupData_SSN"
    Const cstrSproc             As String = "dbo.proc_claim_lu_select3"
    Dim aRows()                 As Variant
    Dim rstTemp                 As ADODB.Recordset
    Dim prmReturnValue          As ADODB.Parameter
    
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

        Set rstTemp = .Execute()
    End With
    
    aRows = rstTemp.GetRows(Rows:=adGetRowsRest, Start:=adBookmarkFirst)
    
    LookupData_SSN = aRows
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
    
    ' Clean-up statements go here
    Erase aRows
    fnFreeObject rstTemp
    fnFreeObject prmReturnValue
   
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
Public Property Get LstUpdtDtm() As Date
    ' **************************************************************************
    ' Function  : Property Get LstUpdtDtm
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : Int
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get LstUpdtDtm"
    On Error GoTo PROC_ERR

    LstUpdtDtm = CDate(m_dteLstUpdtDtm.value)
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
Public Property Let LstUpdtDtm(ByVal dteValue As Date)
    ' **************************************************************************
    ' Function  : Property Let LstUpdtDtm
    ' Purpose   : Assigns new Value to property
    ' Parameters: ByVal NewValue As Integer
    ' Returns   :
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let LstUpdtDtm"
    On Error GoTo PROC_ERR

    m_dteLstUpdtDtm.value = dteValue
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
Public Property Get LstUpdtUserId() As String
    ' **************************************************************************
    ' Function  : Property Get LstUpdtUserId
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : Int
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get LstUpdtUserId"
    On Error GoTo PROC_ERR

    LstUpdtUserId = CStr(m_strLstUpdtUserId.value)
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
Public Property Let LstUpdtUserId(ByVal strValue As String)
    ' **************************************************************************
    ' Function  : Property Let LstUpdtUserId
    ' Purpose   : Assigns new Value to property
    ' Parameters: ByVal NewValue As Integer
    ' Returns   :
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let LstUpdtUserId"
    On Error GoTo PROC_ERR

    m_strLstUpdtUserId.value = strValue
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
Public Property Get PycoTypCd() As String
    ' **************************************************************************
    ' Function  : Property Get PycoTypCd
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : string
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get PycoTypCd"
    On Error GoTo PROC_ERR

    PycoTypCd = CStr(m_strPycoTypCd.value)
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
Public Property Let PycoTypCd(ByVal strValue As String)
    ' **************************************************************************
    ' Function  : Property Let PycoTypCd
    ' Purpose   : Assigns new Value to property
    ' Parameters: ByVal NewValue As String
    ' Returns   :
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let PycoTypCd"
    On Error GoTo PROC_ERR

    m_strPycoTypCd.value = strValue
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
    ' Modified:     Berry Kropiwka - Added Compact Calc - 2019-09-27
    
    '!CUSTOMIZE!  Customize the name of the stored procedure, the number and names
    '             of the parameters and, perhaps, the return values trapped in the
    '             error handler.
    
    Const cstrCurrentProc           As String = "AddRecord"
    Const cstrSproc                 As String = "dbo.proc_claim_insert"  ' Stored procedure to execute
    Dim rstTemp                     As ADODB.Recordset
    Dim bSuccessful                 As Boolean
    Dim prmReturnValue              As ADODB.Parameter
    Dim prmAdmnSystCd               As ADODB.Parameter
    Dim prmClmInsdDthDt             As ADODB.Parameter
    Dim prmClmInsdFirstNm           As ADODB.Parameter
    Dim prmClmInsdLastNm            As ADODB.Parameter
    Dim prmClmInsdSsnNum            As ADODB.Parameter
    Dim prmClmNum                   As ADODB.Parameter
    Dim prmClmPolNum                As ADODB.Parameter
    Dim prmClmProofDt               As ADODB.Parameter
    Dim prmClmTotClmPdAmt           As ADODB.Parameter
    Dim prmClmTotDthbPmtAmt         As ADODB.Parameter
    Dim prmClmTotIntAmt             As ADODB.Parameter
    Dim prmClmTotWthldAmt           As ADODB.Parameter
    Dim prmInsdDthResStCd           As ADODB.Parameter
    Dim prmIssStCd                  As ADODB.Parameter
    Dim prmPycoTypCd                As ADODB.Parameter
    Dim prmClmForResDthInd          As ADODB.Parameter
    Dim prmClmCompactClcnInd        As ADODB.Parameter
    Dim prmInvalid_Key              As ADODB.Parameter
    Dim prmNew_Id                   As ADODB.Parameter

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
        ' Define the ADMN_SYST_CD parameter
        Set prmAdmnSystCd = .CreateParameter(Name:="@admn_syst_cd", _
                                         Type:=adChar, _
                                         Direction:=adParamInput, _
                                         Size:=2, _
                                         value:=fnNullIfZLS(varIn:=AdmnSystCd, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmAdmnSystCd

        ' ---Parameter #3---
        ' Define the CLM_INSD_DTH_DT parameter
        Set prmClmInsdDthDt = .CreateParameter(Name:="@clm_insd_dth_dt", _
                                         Type:=adDBTimeStamp, _
                                         Direction:=adParamInput, _
                                         Size:=16, _
                                         value:=ClmInsdDthDt)
        .Parameters.Append prmClmInsdDthDt

        ' ---Parameter #4---
        ' Define the CLM_INSD_FIRST_NM parameter
        Set prmClmInsdFirstNm = .CreateParameter(Name:="@clm_insd_first_nm", _
                                         Type:=adVarChar, _
                                         Direction:=adParamInput, _
                                         Size:=50, _
                                         value:=fnNullIfZLS(varIn:=ClmInsdFirstNm, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmClmInsdFirstNm

        ' ---Parameter #5---
        ' Define the CLM_INSD_LAST_NM parameter
        Set prmClmInsdLastNm = .CreateParameter(Name:="@clm_insd_last_nm", _
                                         Type:=adVarChar, _
                                         Direction:=adParamInput, _
                                         Size:=50, _
                                         value:=fnNullIfZLS(varIn:=ClmInsdLastNm, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmClmInsdLastNm

        ' ---Parameter #6---
        ' Define the CLM_INSD_SSN_NUM parameter
        Set prmClmInsdSsnNum = .CreateParameter(Name:="@clm_insd_ssn_num", _
                                         Type:=adChar, _
                                         Direction:=adParamInput, _
                                         Size:=9, _
                                         value:=fnNullIfZLS(varIn:=ClmInsdSsnNum, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmClmInsdSsnNum

        ' ---Parameter #7---
        ' Define the CLM_NUM parameter
        Set prmClmNum = .CreateParameter(Name:="@clm_num", _
                                         Type:=adVarChar, _
                                         Direction:=adParamInput, _
                                         Size:=20, _
                                         value:=fnNullIfZLS(varIn:=ClmNum, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmClmNum

        ' ---Parameter #8---
        ' Define the CLM_POL_NUM parameter
        Set prmClmPolNum = .CreateParameter(Name:="@clm_pol_num", _
                                         Type:=adChar, _
                                         Direction:=adParamInput, _
                                         Size:=15, _
                                         value:=fnNullIfZLS(varIn:=ClmPolNum, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmClmPolNum

        ' ---Parameter #9---
        ' Define the CLM_PROOF_DT parameter
        Set prmClmProofDt = .CreateParameter(Name:="@clm_proof_dt", _
                                         Type:=adDBTimeStamp, _
                                         Direction:=adParamInput, _
                                         Size:=16, _
                                         value:=ClmProofDt)
        .Parameters.Append prmClmProofDt
        
        ' ---Parameter #10---
        ' Define the CLM_TOT_CLM_PD_AMT parameter (nullable)
        Set prmClmTotClmPdAmt = .CreateParameter(Name:="@clm_tot_clm_pd_amt", _
                                          Type:=adNumeric, _
                                          Direction:=adParamInput, _
                                          value:=ClmTotClmPdAmt)
        With m_varClmTotClmPdAmt
            prmClmTotClmPdAmt.Precision = .Precision
            prmClmTotClmPdAmt.NumericScale = .NumericScale
        End With
        .Parameters.Append prmClmTotClmPdAmt

        ' ---Parameter #11---
        ' Define the CLM_TOT_DTHB_PMT_AMT parameter (nullable)
        Set prmClmTotDthbPmtAmt = .CreateParameter(Name:="@clm_tot_dthb_pmt_amt", _
                                          Type:=adNumeric, _
                                          Direction:=adParamInput, _
                                          value:=ClmTotDthbPmtAmt)
        With m_varClmTotDthbPmtAmt
            prmClmTotDthbPmtAmt.Precision = .Precision
            prmClmTotDthbPmtAmt.NumericScale = .NumericScale
        End With
        .Parameters.Append prmClmTotDthbPmtAmt

        ' ---Parameter #12---
        ' Define the CLM_TOT_INT_AMT parameter (nullable)
        Set prmClmTotIntAmt = .CreateParameter(Name:="@clm_tot_int_amt", _
                                          Type:=adNumeric, _
                                          Direction:=adParamInput, _
                                          value:=ClmTotIntAmt)
        With m_varClmTotIntAmt
            prmClmTotIntAmt.Precision = .Precision
            prmClmTotIntAmt.NumericScale = .NumericScale
        End With
        .Parameters.Append prmClmTotIntAmt

        ' ---Parameter #13---
        ' Define the CLM_TOT_WTHLD_AMT parameter (nullable)
        Set prmClmTotWthldAmt = .CreateParameter(Name:="@clm_tot_wthld_amt", _
                                          Type:=adNumeric, _
                                          Direction:=adParamInput, _
                                          value:=ClmTotWthldAmt)
        With m_varClmTotWthldAmt
            prmClmTotWthldAmt.Precision = .Precision
            prmClmTotWthldAmt.NumericScale = .NumericScale
        End With
        .Parameters.Append prmClmTotWthldAmt
        
        ' ---Parameter #14---
        ' Define the INSD_DTH_RES_ST_CD parameter
        Set prmInsdDthResStCd = .CreateParameter(Name:="@insd_dth_res_st_cd", _
                                         Type:=adChar, _
                                         Direction:=adParamInput, _
                                         Size:=2, _
                                         value:=fnNullIfZLS(varIn:=InsdDthResStCd, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmInsdDthResStCd
        
        ' ---Parameter #15---
        ' Define the ISS_ST_CD parameter
        Set prmIssStCd = .CreateParameter(Name:="@iss_st_cd", _
                                         Type:=adChar, _
                                         Direction:=adParamInput, _
                                         Size:=2, _
                                         value:=fnNullIfZLS(varIn:=IssStCd, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmIssStCd
        
        ' ---Parameter #16---
        ' Define the PYCO_TYP_CD parameter
        Set prmPycoTypCd = .CreateParameter(Name:="@pyco_typ_cd", _
                                         Type:=adChar, _
                                         Direction:=adParamInput, _
                                         Size:=8, _
                                         value:=fnNullIfZLS(varIn:=PycoTypCd, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmPycoTypCd
        
        ' ---Parameter #17---
        ' Define the clm_for_res_dth_ind parameter
        Set prmClmForResDthInd = .CreateParameter(Name:="@clm_for_res_dth_ind", _
                                         Type:=adChar, _
                                         Direction:=adParamInput, _
                                         Size:=1, _
                                         value:=fnBoolToYN(bIn:=ClmForResDthInd))
        .Parameters.Append prmClmForResDthInd
        
        ' ---Parameter #18---
        ' Define the clm_compact_clcn_ind parameter
        Set prmClmCompactClcnInd = .CreateParameter(Name:="@clm_compact_clcn_ind", _
                                         Type:=adChar, _
                                         Direction:=adParamInput, _
                                         Size:=1, _
                                         value:=fnBoolToTF(bIn:=ClmCompactClcnInd))
        .Parameters.Append prmClmCompactClcnInd
        
        ' ---Parameter #19---
        ' Define the Invalid_Key output parameter, which reflects *which* foreign
        ' key violation was encountered.
        Set prmInvalid_Key = .CreateParameter(Name:="@Invalid_Key", _
                                                  Type:=adVarChar, _
                                                  Size:=255, _
                                                  Direction:=adParamOutput)
        .Parameters.Append prmInvalid_Key

        ' ---Parameter #20---
        ' Define the output parameter that represents the CLM_ID value that
        ' was assigned to the record we're trying to insert
        Set prmNew_Id = .CreateParameter(Name:="@New_Id", _
                                              Type:=adInteger, _
                                              Direction:=adParamOutput, _
                                              value:=Null)
        .Parameters.Append prmNew_Id

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
    bSuccessful = GetRelativeRecord(GetClmNumFromClmID(prmNew_Id.value), epdSameRecord)

    AddRecord = True
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
    
    ' Clean-up statements go here
    fnFreeRecordset rstTemp
    fnFreeObject prmReturnValue
    fnFreeObject prmAdmnSystCd
    fnFreeObject prmClmInsdDthDt
    fnFreeObject prmClmInsdFirstNm
    fnFreeObject prmClmInsdLastNm
    fnFreeObject prmClmInsdSsnNum
    fnFreeObject prmClmNum
    fnFreeObject prmClmPolNum
    fnFreeObject prmClmProofDt
    fnFreeObject prmClmTotClmPdAmt
    fnFreeObject prmClmTotDthbPmtAmt
    fnFreeObject prmClmTotIntAmt
    fnFreeObject prmClmTotWthldAmt
    fnFreeObject prmInsdDthResStCd
    fnFreeObject prmIssStCd
    fnFreeObject prmPycoTypCd
    fnFreeObject prmClmForResDthInd
    fnFreeObject prmClmCompactClcnInd
    fnFreeObject prmInvalid_Key
    fnFreeObject prmNew_Id
    
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case prmReturnValue
        Case gcRES_NERR_ERR_WHILE_TRYING_TO
            ' 4028 = An error occurred while attempting to @@1 this record.
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_ERR_WHILE_TRYING_TO, _
                                       mcstrName & cstrCurrentProc, _
                                       "add"
            Resume PROC_EXIT
        Case gcRES_NERR_ADD_WITH_NONUNIQUE_KEY
            ' 4031 = A record with the specified key (@@1) already exists. Please specify a unique key.
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_ADD_WITH_NONUNIQUE_KEY, _
                                       mcstrName & cstrCurrentProc, _
                                       "Claim Number " & RTrim$(ClmNum)
            Resume PROC_EXIT
        Case gcRES_NERR_KEY_NOT_FOUND
            ' 4032 = The @@1 specified (@@2) is no longer defined in the @@3 table. Please choose a different value.
            If UCase$(prmInvalid_Key) = "ADMN_SYST_CD" Then
                gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_KEY_NOT_FOUND, _
                                           mcstrName & cstrCurrentProc, _
                                           "Admin System", RTrim$(AdmnSystCd), "ADMIN_SYSTEM_T"
                Resume PROC_EXIT
            ElseIf UCase$(prmInvalid_Key) = "INSD_DTH_RES_ST_CD" Then
                gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_KEY_NOT_FOUND, _
                                           mcstrName & cstrCurrentProc, _
                                           "Residence State", RTrim$(InsdDthResStCd), "STATE_T"
                Resume PROC_EXIT
            ElseIf UCase$(prmInvalid_Key) = "ISS_ST_CD" Then
                gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_KEY_NOT_FOUND, _
                                           mcstrName & cstrCurrentProc, _
                                           "Issue State", RTrim$(IssStCd), "STATE_T"
                Resume PROC_EXIT
            Else
                ' PYCO_TYP_CD
                gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_KEY_NOT_FOUND, _
                                           mcstrName & cstrCurrentProc, _
                                           "Company Type", RTrim$(PycoTypCd), "PAYOR_COMPANY_TYPE_T"
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
    Set rstSingleRecord_Fresh = fnSelectRecord(ClmId)
    
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
            If (DateDiff("s", CStr(m_dteLstUpdtDtm_Original), CStr(!lst_updt_dtm)) <> 0) _
               Or _
               (!lst_updt_user_id <> m_strLstUpdtUserId_Original) Then
                    strACF2 = !lst_updt_user_id
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
    
    Const cstrCurrentProc          As String = "DeleteRecord"
    Const cstrSproc                As String = "dbo.proc_claim_delete"  ' Stored procedure to execute
    Dim bSuccessful                As Boolean
    Dim adwTemp                    As cadwADOWrapper
    Dim rstTemp                    As ADODB.Recordset
    Dim prmReturnValue             As ADODB.Parameter
    Dim prmClmId                   As ADODB.Parameter
    Dim prmDependent_Table         As ADODB.Parameter
    Dim strClmNum                  As String

    On Error GoTo PROC_ERR

    Set adwTemp = New cadwADOWrapper
    adwTemp.CommandSetSproc strSprocName:=cstrSproc

    ' Save the Claim Number associated with the Claim that's going to be deleted.
    ' We'll need it afterwards to position to the previous record.
    strClmNum = GetClmNumFromClmID(ClmId)
    
    ' For Char/VarChar fields,
    '     * Use fnNullIfZLS to ensure Nulls are appropriately handled.
    '     * Do *not* set the optional 2nd parameter to fnNullIfZLS to True.
    ' For numeric fields,
    '     * Use fnNullIfZero to ensure Nulls are appropriately handled.
    ' For Y/N fields,
    '     * Use fnBoolToYN to ensure True/False is appropriately translated.

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
        ' Define the CLM_ID parameter
        Set prmClmId = .CreateParameter(Name:="@clm_id", _
                                        Type:=adInteger, _
                                        Direction:=adParamInput, _
                                        value:=ClmId)
        .Parameters.Append prmClmId

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
    bSuccessful = GetRelativeRecord(strClmNum, epdPreviousRecord)
    
    DeleteRecord = True
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
    
    ' Clean-up statements go here
    fnFreeRecordset rstTemp
    fnFreeObject adwTemp
    fnFreeObject prmReturnValue
    fnFreeObject prmClmId
    fnFreeObject prmDependent_Table
    
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case prmReturnValue
        Case gcRES_NERR_REC_NOT_FOUND
            ' 4027 = The specified record was not found in the database (@@1).
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_REC_NOT_FOUND, _
                                       mcstrName & cstrCurrentProc, _
                                       "Claim ID " & RTrim$(ClmId)
            Resume PROC_EXIT
        Case gcRES_NERR_ERR_WHILE_TRYING_TO
            ' 4028 = An error occurred while attempting to @@1 this record.
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_ERR_WHILE_TRYING_TO, _
                                       mcstrName & cstrCurrentProc, _
                                       "delete"
            Resume PROC_EXIT
        Case gcRES_NERR_DEPENDENT_RECS_EXIST
            ' 4029 = This @@1 is associated with one or more records on the @@2 table and cannot be deleted until those records themselves are deleted.
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_DEPENDENT_RECS_EXIST, _
                                       mcstrName & cstrCurrentProc, _
                                       "Claim ID (" & RTrim$(ClmId) & ")", UCase$(prmDependent_Table)
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
Public Function GetClmIdFromClmNum(ByVal strClmNum As String) As Variant
    '--------------------------------------------------------------------------
    ' Procedure:   GetClmIdFromClmNum
    ' Description: Query database for the Claim ID for a specified Claim Number
    ' Params:
    '               strClmNum  (in)  The Claim Number to translate.
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc               As String = "GetClmIdFromClmNum"
    Const cstrSproc                     As String = "dbo.proc_clm_id_lu_select"
    Dim rstTemp                         As ADODB.Recordset
    Dim prmReturnValue                  As ADODB.Parameter
    Dim prmClmNum                       As ADODB.Parameter
    Dim prmClmId                        As ADODB.Parameter
    
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

        ' ---Parameter #2---
        Set prmClmNum = .CreateParameter(Name:="clm_num", _
                                         Type:=adVarChar, _
                                         Direction:=adParamInput, _
                                         Size:=20, _
                                         value:=strClmNum)
        .Parameters.Append prmClmNum

        ' ---Parameter #2---
        Set prmClmId = .CreateParameter(Name:="clm_id", _
                                         Type:=adInteger, _
                                         Direction:=adParamInputOutput, _
                                         value:=Null)
        .Parameters.Append prmClmId

        Set rstTemp = .Execute()
    End With

    GetClmIdFromClmNum = prmClmId.value
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    fnFreeObject prmReturnValue
    fnFreeObject prmClmNum
    fnFreeObject prmClmId
    
    fnFreeRecordset rstTemp
    
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
Public Function GetClmNumFromClmID(ByVal lngClmID As Long) As Variant
    '--------------------------------------------------------------------------
    ' Procedure:   GetClmIdFromClmNum
    ' Description: Query database for the Claim ID for a specified Claim Number
    ' Params:
    '               strClmID  (in)  The Claim ID to translate.
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc               As String = "GetClmNumFromClmID"
    Const cstrSproc                     As String = "dbo.proc_clm_num_lu_select"
    Dim rstTemp                         As ADODB.Recordset
    Dim prmReturnValue                  As ADODB.Parameter
    Dim prmClmNum                       As ADODB.Parameter
    Dim prmClmId                        As ADODB.Parameter
    
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

        ' ---Parameter #2---
        Set prmClmId = .CreateParameter(Name:="clm_id", _
                                         Type:=adInteger, _
                                         Direction:=adParamInput, _
                                         value:=lngClmID)
        .Parameters.Append prmClmId

        ' ---Parameter #3---
        Set prmClmNum = .CreateParameter(Name:="clm_num", _
                                         Type:=adVarChar, _
                                         Direction:=adParamInputOutput, _
                                         Size:=20, _
                                         value:=Null)
        .Parameters.Append prmClmNum


        Set rstTemp = .Execute()
    End With

    GetClmNumFromClmID = prmClmNum.value
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    fnFreeObject prmReturnValue
    fnFreeObject prmClmId
    fnFreeObject prmClmNum
    
    fnFreeRecordset rstTemp
    
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
Public Function GetLobCdFromAdmnSystCd(ByVal strAdmnSystCd As String) As String
    '--------------------------------------------------------------------------
    ' Procedure:   GetLobCdFromAdmnSystCd
    ' Description: Query database for the Line-of-Business Code for a specified
    '              Admin System Code
    ' Params:
    '              strAdmnSystCd  (in)  The Admin System Code to translate.
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc               As String = "GetLobCdFromAdmnSystCd"
    Const cstrSproc                     As String = "dbo.proc_lob_cd_lu_select"
    Dim rstTemp                         As ADODB.Recordset
    Dim prmReturnValue                  As ADODB.Parameter
    Dim prmAdmnSystCd                   As ADODB.Parameter
    Dim prmLobCd                        As ADODB.Parameter
    
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

        ' ---Parameter #2---
        Set prmAdmnSystCd = .CreateParameter(Name:="admn_syst_cd", _
                                         Type:=adChar, _
                                         Direction:=adParamInput, _
                                         Size:=2, _
                                         value:=strAdmnSystCd)
        .Parameters.Append prmAdmnSystCd

        ' ---Parameter #2---
        Set prmLobCd = .CreateParameter(Name:="lob_cd", _
                                         Type:=adChar, _
                                         Direction:=adParamOutput, _
                                         Size:=1, _
                                         value:=Null)
        .Parameters.Append prmLobCd

        Set rstTemp = .Execute()
    End With

    GetLobCdFromAdmnSystCd = prmLobCd.value
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    fnFreeObject prmReturnValue
    fnFreeObject prmAdmnSystCd
    fnFreeObject prmLobCd
    
    fnFreeRecordset rstTemp
    
    On Error GoTo 0
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case prmReturnValue
        Case gcRES_NERR_REC_NOT_FOUND
            ' 4027 = The specified record was not found in the database (@@1).
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_REC_NOT_FOUND, _
                                       mcstrName & cstrCurrentProc, _
                                       "Claim ID " & RTrim$(ClmId)
            Resume PROC_EXIT
        Case gcRES_NERR_ERR_WHILE_TRYING_TO
            ' 4028 = An error occurred while attempting to @@1 this record.
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
    Const cstrSproc             As String = "dbo.proc_claim_lu_select"
    Dim prmReturnValue          As ADODB.Parameter

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
        Case gcRES_NERR_ERR_WHILE_TRYING_TO
            ' 4028 = An error occurred while attempting to @@1 this record.
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
    '              the record relative to the specified **logical** key value.
    '              Then, it resets each of the class properties that correspond
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
    '                    This table's true key is CLM_ID, but the Insured screen
    '                    shows claims in CLM_NUM sequence. Hence, the key this
    '                    proc uses is CLM_NUM.
    '
    ' Params:
    '     strKey1              (in) = ClmNum value from which to do the relative
    '                                 repositioning
    '     lngPositionDirection (in) = Indicates to which relative record the
    '                                 recordset should be positioned (relative
    '                                 to the strKey1 parameter value).
    '
    '
    ' Called By:   cmdDelete_Click( ) of frmclaim.frm
    '              cmdUpdate_Click( ) of frmclaim.frm
    '              cmdNavigate_Click( ) of frmclaim.frm
    '              fnAddRecord( ) of frmclaim.frm
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
                .Filter = "clm_num < '" & strKey1 & "'"
                If .RecordCount <> 0 Then
                    ' Grab the last record. The one with the highest key less than the
                    ' specified key is the one we want.
                    m_adwADO.MoveLast m_rstLookup
                    '!CUSTOMIZE! The constant referenced below should refer to the key field.
                    strKey1ForNewRec = .Fields(mcstrClmNum).value
                    ' Okay, we got it. Now make all records in the Lookup recordset
                    ' visible again, and then reposition to the new record.
                    .Filter = adFilterNone
                    m_adwADO.MoveFirst m_rstLookup
                    .Find "clm_num = '" & strKey1ForNewRec & "'"
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
                .Filter = "clm_num > '" & strKey1 & "'"
                If .RecordCount <> 0 Then
                    ' Grab the first record. The one with the lowest key higher than the
                    ' specified key is the one we want.
                    m_adwADO.MoveFirst m_rstLookup
                    '!CUSTOMIZE! The constant referenced below should refer to the key field.
                    strKey1ForNewRec = .Fields(mcstrClmNum).value
                    ' Okay, we got it. Now make all records in the Lookup recordset
                    ' visible again, and then reposition to the new record.
                    .Filter = adFilterNone
                    m_adwADO.MoveFirst m_rstLookup
                    .Find "clm_num = '" & strKey1ForNewRec & "'"
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
                .Find "clm_num = '" & strKey1 & "'"
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
            GetSingleRecord lngKey1:=.Fields(mcstrClmId).value, bSynchLookupRST:=True
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
Public Function GetSingleRecord(ByVal lngKey1 As Long, _
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
    '    lngKey1         (in) = represents the primary key for the table (clm_id)
    '    bSynchLookupRST (in) = indicates whether the Lookup recordset should be
    '                           repositioned to the record this function just
    '                           retrieved. This would be set to True by the
    '                           form's vfgLookup_ChangeEdit event handler to ensure
    '                           the "record x of y" will be set appropriately when
    '                           it calls fnLoadControls.
    '-----------------------------------------------------------------------------
    ' Modified:     Berry Kropiwka - Added Compact Calc - 2019-09-27
    
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
    Set rstTemp = fnSelectRecord(lngKey1)

    With Me
        ' To ensure Nulls and/or True/False are appropriately translated, use:
        '     adChar        (Boolean, regardless of nullability)    = fnYNToBool
        '     adChar        (Other, regardless of nullability)      = fnZLSIfNull
        '     adDBTimeStamp (regardless of nullability)             = N/A
        '     adInteger     (not nullable)                          = fnZeroIfNull
        '     adInteger     (nullable)                              = N/A
        '     adNumeric     (not nullable)                          = fnZeroIfNull
        '     adNumeric     (nullable)                              = N/A
        '     adVarChar     (regardless of nullability)             = fnZLSIfNull
        .AdmnSystCd = fnZLSIfNull(rstTemp!admn_syst_cd)
        .ClmId = fnZeroIfNull(rstTemp!clm_id)
        .ClmInsdDthDt = rstTemp!clm_insd_dth_dt
        .ClmInsdFirstNm = fnZLSIfNull(rstTemp!clm_insd_first_nm)
        .ClmInsdLastNm = fnZLSIfNull(rstTemp!clm_insd_last_nm)
        .ClmInsdSsnNum = fnZLSIfNull(rstTemp!clm_insd_ssn_num)
        .ClmNum = fnZLSIfNull(rstTemp!clm_num)
        .ClmPolNum = fnZLSIfNull(rstTemp!clm_pol_num)
        .ClmProofDt = rstTemp!clm_proof_dt
        .ClmTotClmPdAmt = rstTemp!clm_tot_clm_pd_amt
        .ClmTotDthbPmtAmt = rstTemp!clm_tot_dthb_pmt_amt
        .ClmTotIntAmt = rstTemp!clm_tot_int_amt
        .ClmTotWthldAmt = rstTemp!clm_tot_wthld_amt
        .InsdDthResStCd = fnZLSIfNull(rstTemp!insd_dth_res_st_cd)
        .IssStCd = fnZLSIfNull(rstTemp!iss_st_cd)
        .LstUpdtDtm = rstTemp!lst_updt_dtm
        .LstUpdtUserId = fnZLSIfNull(rstTemp!lst_updt_user_id)
        .PycoTypCd = fnZLSIfNull(rstTemp!pyco_typ_cd)
        .ClmForResDthInd = fnYNToBool(rstTemp!clm_for_res_dth_ind)
        .ClmCompactClcnInd = fnTFToBool(rstTemp!clm_compact_clcn_ind)
    End With

    ' Save original Last Updated info, to be used during UpdateRecord( ) and DeleteRecord( )
    ' to determine if another user updated the record since it was retrieved.
    m_dteLstUpdtDtm_Original = fnZLSIfNull(rstTemp!lst_updt_dtm)
    m_strLstUpdtUserId_Original = fnZLSIfNull(rstTemp!lst_updt_user_id)
    
    If bSynchLookupRST Then
        With m_rstLookup
            m_adwADO.MoveFirst m_rstLookup
            .Find "clm_id = " & ClmId
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
        GetSingleRecord m_rstLookup.Fields(mcstrClmId).value
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
        GetSingleRecord m_rstLookup.Fields(mcstrClmId).value
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
                GetRelativeRecord ClmNum, epdNextRecord
            Else
                ' Get the requested record and reposition the Lookup recordset to that record
                '!CUSTOMIZE! The constant referenced below should refer to the key field.
                GetSingleRecord m_rstLookup.Fields(mcstrClmId).value
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
                GetRelativeRecord ClmNum, epdPreviousRecord
            Else
                ' Get the requested record and reposition the Lookup recordset to that record
                '!CUSTOMIZE! The constant referenced below should refer to the key field.
                GetSingleRecord m_rstLookup.Fields(mcstrClmId).value
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
Public Function HaveDependents(ByVal lngKey1 As Long, ByRef strDependentTable As String) As Boolean
    ' Comments  : Determines whether the current record can be deleted without
    '             hitting a referential integrity violation due to either:
    '             a. row(s) existing in other tables that use the current key value
    '                as a foreign key
    '             b. (for tables with circular references only) row(s) existing in the same table
    '                which has a circular reference to to the current key value.
    '             The calling form should look at the return value. If True, then
    '             the form's Delete button should be disabled.
    '
    ' Parameters:
    '   lngKey1
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
    Const cstrCurrentProc      As String = "HaveDependents"
    Const cstrSproc            As String = "dbo.proc_claim_verify_dependents"  ' Stored procedure to execute
    Dim adwTemp                As cadwADOWrapper
    Dim rstTemp                As ADODB.Recordset
    Dim prmReturnValue         As ADODB.Parameter
    Dim prmClmId               As ADODB.Parameter
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
        Set prmClmId = .CreateParameter(Name:="@clm_id", _
                                         Type:=adInteger, _
                                         Direction:=adParamInput, _
                                         value:=lngKey1)
        .Parameters.Append prmClmId

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
    fnFreeObject prmClmId
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
        Case gcRES_NERR_REC_NOT_FOUND
            ' 4027 = The specified record was not found in the database (@@1).
            HaveDependents = True
            strDependentTable = "Unknown"   ' This is actually ignored by the caller
            ' Remove any trace that this error occurred since we're not going to report it as an error.
            Err.Clear
            Resume PROC_EXIT
        Case gcRES_NERR_DEPENDENT_RECS_EXIST
            ' 4029 = This @@1 is associated with one or more records on the @@2 table and cannot be deleted until those records themselves are deleted.
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
    ' Modified:     Berry Kropiwka - Added Compact Calc - 2019-09-27

    '!CUSTOMIZE!  Customize the name of the stored procedure, the number and names
    '             of the parameters and, perhaps, the return values trapped in the
    '             error handler.

    Const cstrCurrentProc           As String = "UpdateRecord"
    Const cstrSproc                 As String = "dbo.proc_claim_update"  ' Stored procedure to execute
    Dim rstTemp                     As ADODB.Recordset
    Dim bSuccessful                 As Boolean
    Dim prmReturnValue              As ADODB.Parameter
    Dim prmClmId                    As ADODB.Parameter
    Dim prmAdmnSystCd               As ADODB.Parameter
    Dim prmClmInsdDthDt             As ADODB.Parameter
    Dim prmClmInsdFirstNm           As ADODB.Parameter
    Dim prmClmInsdLastNm            As ADODB.Parameter
    Dim prmClmInsdSsnNum            As ADODB.Parameter
    Dim prmClmNum                   As ADODB.Parameter
    Dim prmClmPolNum                As ADODB.Parameter
    Dim prmClmProofDt               As ADODB.Parameter
    Dim prmClmTotClmPdAmt           As ADODB.Parameter
    Dim prmClmTotDthbPmtAmt         As ADODB.Parameter
    Dim prmClmTotIntAmt             As ADODB.Parameter
    Dim prmClmTotWthldAmt           As ADODB.Parameter
    Dim prmInsdDthResStCd           As ADODB.Parameter
    Dim prmIssStCd                  As ADODB.Parameter
    Dim prmPycoTypCd                As ADODB.Parameter
    Dim prmClmForResDthInd          As ADODB.Parameter
    Dim prmClmCompactClcnInd        As ADODB.Parameter
    Dim prmInvalid_Key              As ADODB.Parameter
    
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
        ' Define the CLM_ID parameter
        Set prmClmId = .CreateParameter(Name:="@clm_id", _
                                        Type:=adInteger, _
                                        Direction:=adParamInput, _
                                        value:=ClmId)
        .Parameters.Append prmClmId

        ' ---Parameter #3---
        ' Define the ADMN_SYST_CD parameter
        Set prmAdmnSystCd = .CreateParameter(Name:="@admn_syst_cd", _
                                         Type:=adChar, _
                                         Direction:=adParamInput, _
                                         Size:=2, _
                                         value:=fnNullIfZLS(varIn:=AdmnSystCd, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmAdmnSystCd

        ' ---Parameter #4---
        ' Define the CLM_INSD_DTH_DT parameter
        Set prmClmInsdDthDt = .CreateParameter(Name:="@clm_insd_dth_dt", _
                                         Type:=adDBTimeStamp, _
                                         Direction:=adParamInput, _
                                         Size:=16, _
                                         value:=ClmInsdDthDt)
        .Parameters.Append prmClmInsdDthDt

        ' ---Parameter #5---
        ' Define the CLM_INSD_FIRST_NM parameter
        Set prmClmInsdFirstNm = .CreateParameter(Name:="@clm_insd_first_nm", _
                                         Type:=adVarChar, _
                                         Direction:=adParamInput, _
                                         Size:=50, _
                                         value:=fnNullIfZLS(varIn:=ClmInsdFirstNm, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmClmInsdFirstNm

        ' ---Parameter #6---
        ' Define the CLM_INSD_LAST_NM parameter
        Set prmClmInsdLastNm = .CreateParameter(Name:="@clm_insd_last_nm", _
                                         Type:=adVarChar, _
                                         Direction:=adParamInput, _
                                         Size:=50, _
                                         value:=fnNullIfZLS(varIn:=ClmInsdLastNm, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmClmInsdLastNm

        ' ---Parameter #7---
        ' Define the CLM_INSD_SSN_NUM parameter
        Set prmClmInsdSsnNum = .CreateParameter(Name:="@clm_insd_ssn_num", _
                                         Type:=adChar, _
                                         Direction:=adParamInput, _
                                         Size:=9, _
                                         value:=fnNullIfZLS(varIn:=ClmInsdSsnNum, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmClmInsdSsnNum

        ' ---Parameter #8---
        ' Define the CLM_NUM parameter
        Set prmClmNum = .CreateParameter(Name:="@clm_num", _
                                         Type:=adVarChar, _
                                         Direction:=adParamInput, _
                                         Size:=20, _
                                         value:=fnNullIfZLS(varIn:=ClmNum, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmClmNum

        ' ---Parameter #9---
        ' Define the CLM_POL_NUM parameter
        Set prmClmPolNum = .CreateParameter(Name:="@clm_pol_num", _
                                         Type:=adChar, _
                                         Direction:=adParamInput, _
                                         Size:=15, _
                                         value:=fnNullIfZLS(varIn:=ClmPolNum, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmClmPolNum

        ' ---Parameter #10---
        ' Define the CLM_PROOF_DT parameter
        Set prmClmProofDt = .CreateParameter(Name:="@clm_proof_dt", _
                                         Type:=adDBTimeStamp, _
                                         Direction:=adParamInput, _
                                         Size:=16, _
                                         value:=ClmProofDt)
        .Parameters.Append prmClmProofDt
        
        ' ---Parameter #11---
        ' Define the CLM_TOT_CLM_PD_AMT parameter
        Set prmClmTotClmPdAmt = .CreateParameter(Name:="@clm_tot_clm_pd_amt", _
                                          Type:=adNumeric, _
                                          Direction:=adParamInput, _
                                          value:=ClmTotClmPdAmt)
        With m_varClmTotClmPdAmt
            prmClmTotClmPdAmt.Precision = .Precision
            prmClmTotClmPdAmt.NumericScale = .NumericScale
        End With
        .Parameters.Append prmClmTotClmPdAmt

        ' ---Parameter #12---
        ' Define the CLM_TOT_DTHB_PMT_AMT parameter
        Set prmClmTotDthbPmtAmt = .CreateParameter(Name:="@clm_tot_dthb_pmt_amt", _
                                          Type:=adNumeric, _
                                          Direction:=adParamInput, _
                                          value:=ClmTotDthbPmtAmt)
        With m_varClmTotDthbPmtAmt
            prmClmTotDthbPmtAmt.Precision = .Precision
            prmClmTotDthbPmtAmt.NumericScale = .NumericScale
        End With
        .Parameters.Append prmClmTotDthbPmtAmt

        ' ---Parameter #13---
        ' Define the CLM_TOT_INT_AMT parameter
        Set prmClmTotIntAmt = .CreateParameter(Name:="@clm_tot_int_amt", _
                                          Type:=adNumeric, _
                                          Direction:=adParamInput, _
                                          value:=ClmTotIntAmt)
        With m_varClmTotIntAmt
            prmClmTotIntAmt.Precision = .Precision
            prmClmTotIntAmt.NumericScale = .NumericScale
        End With
        .Parameters.Append prmClmTotIntAmt

        ' ---Parameter #14---
        ' Define the CLM_TOT_WTHLD_AMT parameter
        Set prmClmTotWthldAmt = .CreateParameter(Name:="@clm_tot_wthld_amt", _
                                          Type:=adNumeric, _
                                          Direction:=adParamInput, _
                                          value:=ClmTotWthldAmt)
        With m_varClmTotWthldAmt
            prmClmTotWthldAmt.Precision = .Precision
            prmClmTotWthldAmt.NumericScale = .NumericScale
        End With
        .Parameters.Append prmClmTotWthldAmt
        
        ' ---Parameter #15---
        ' Define the INSD_DTH_RES_ST_CD parameter
        Set prmInsdDthResStCd = .CreateParameter(Name:="@insd_dth_res_st_cd", _
                                         Type:=adChar, _
                                         Direction:=adParamInput, _
                                         Size:=2, _
                                         value:=fnNullIfZLS(varIn:=InsdDthResStCd, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmInsdDthResStCd
        
        ' ---Parameter #16---
        ' Define the ISS_ST_CD parameter
        Set prmIssStCd = .CreateParameter(Name:="@iss_st_cd", _
                                         Type:=adChar, _
                                         Direction:=adParamInput, _
                                         Size:=2, _
                                         value:=fnNullIfZLS(varIn:=IssStCd, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmIssStCd
        
        ' ---Parameter #17---
        ' Define the PYCO_TYP_CD parameter
        Set prmPycoTypCd = .CreateParameter(Name:="@pyco_typ_cd", _
                                         Type:=adChar, _
                                         Direction:=adParamInput, _
                                         Size:=8, _
                                         value:=fnNullIfZLS(varIn:=PycoTypCd, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmPycoTypCd

        ' ---Parameter #18---
        ' Define the CLM_FOR_RES_DTH_IND parameter
        Set prmClmForResDthInd = .CreateParameter(Name:="clm_for_res_dth_ind", _
                                             Type:=adChar, _
                                             Direction:=adParamInput, _
                                             Size:=1, _
                                             value:=fnBoolToYN(ClmForResDthInd))
        .Parameters.Append prmClmForResDthInd

        ' ---Parameter #19---
        ' Define the clm_compact_clcn_ind parameter
        Set prmClmCompactClcnInd = .CreateParameter(Name:="clm_compact_clcn_ind", _
                                             Type:=adChar, _
                                             Direction:=adParamInput, _
                                             Size:=1, _
                                             value:=fnBoolToTF(ClmCompactClcnInd))
        .Parameters.Append prmClmCompactClcnInd

        ' ---Parameter #20---
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
    bSuccessful = GetRelativeRecord(ClmNum, epdSameRecord)

    UpdateRecord = True
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
    
    ' Clean-up statements go here
    fnFreeRecordset rstTemp
    fnFreeObject prmReturnValue
    fnFreeObject prmClmId
    fnFreeObject prmAdmnSystCd
    fnFreeObject prmClmInsdDthDt
    fnFreeObject prmClmInsdFirstNm
    fnFreeObject prmClmInsdLastNm
    fnFreeObject prmClmInsdSsnNum
    fnFreeObject prmClmNum
    fnFreeObject prmClmPolNum
    fnFreeObject prmClmProofDt
    fnFreeObject prmClmTotClmPdAmt
    fnFreeObject prmClmTotDthbPmtAmt
    fnFreeObject prmClmTotIntAmt
    fnFreeObject prmClmTotWthldAmt
    fnFreeObject prmInsdDthResStCd
    fnFreeObject prmIssStCd
    fnFreeObject prmPycoTypCd
    fnFreeObject prmClmForResDthInd
    fnFreeObject prmClmCompactClcnInd
    fnFreeObject prmInvalid_Key
    
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case prmReturnValue
        Case gcRES_NERR_REC_NOT_FOUND
            ' 4027 = The specified record was not found in the database (@@1).
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_REC_NOT_FOUND, _
                                       mcstrName & cstrCurrentProc, _
                                       "Claim Number " & RTrim$(ClmNum)
            Resume PROC_EXIT
        Case gcRES_NERR_ERR_WHILE_TRYING_TO
            ' 4028 = An error occurred while attempting to @@1 this record.
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_ERR_WHILE_TRYING_TO, _
                                       mcstrName & cstrCurrentProc, _
                                       "update"
            Resume PROC_EXIT
        Case gcRES_NERR_KEY_NOT_FOUND
            ' 4032 = The @@1 specified (@@2) is no longer defined in the @@3 table. Please choose a different value.
            If UCase$(prmInvalid_Key) = "ADMN_SYST_CD" Then
                gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_KEY_NOT_FOUND, _
                                           mcstrName & cstrCurrentProc, _
                                           "Admin System", RTrim$(AdmnSystCd), "ADMIN_SYSTEM_T"
                Resume PROC_EXIT
            ElseIf UCase$(prmInvalid_Key) = "INSD_DTH_RES_ST_CD" Then
                gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_KEY_NOT_FOUND, _
                                           mcstrName & cstrCurrentProc, _
                                           "Residence State", RTrim$(InsdDthResStCd), "STATE_T"
                Resume PROC_EXIT
            ElseIf UCase$(prmInvalid_Key) = "ISS_ST_CD" Then
                gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_KEY_NOT_FOUND, _
                                           mcstrName & cstrCurrentProc, _
                                           "Issue State", RTrim$(IssStCd), "STATE_T"
                Resume PROC_EXIT
            Else
                ' PYCO_TYP_CD
                gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_KEY_NOT_FOUND, _
                                           mcstrName & cstrCurrentProc, _
                                           "Company Type", RTrim$(PycoTypCd), "PAYOR_COMPANY_TYPE_T"
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
            
        Case -2147217873 ' Cannot insert duplicate key row in object 'xxx' with unique index 'yyy'
            ' 4031 = A record with the specified key (@@1) already exists. Please specify a unique key.
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_ADD_WITH_NONUNIQUE_KEY, _
                                       mcstrName & cstrCurrentProc, _
                                       "Claim Number " & RTrim$(ClmNum)
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
    ' Modified:     Berry Kropiwka - Added Compact Calc - 2019-09-27
    
    '!CUSTOMIZE!  This proc should set each class property (of type UDTColumn) that
    '             corresponds to a table column. What it is set to depends on its
    '             data type: Strings => vbNullString
    '                        Numeric => 0
    '                        Nullable numerics => Null
    '                        Booleans => False
    '                        Dates => Now
    
    Const cstrCurrentProc          As String = "fnClearPropertyValues"
    Const clngZero                 As Long = 0

    On Error GoTo PROC_ERR

    With Me
        .AdmnSystCd = vbNullString
        .ClmId = clngZero
        .ClmInsdDthDt = Now
        .ClmInsdFirstNm = vbNullString
        .ClmInsdLastNm = vbNullString
        .ClmInsdSsnNum = vbNullString
        .ClmNum = vbNullString
        .ClmPolNum = vbNullString
        .ClmProofDt = Now
        .ClmTotClmPdAmt = Null
        .ClmTotDthbPmtAmt = Null
        .ClmTotIntAmt = Null
        .ClmTotWthldAmt = Null
        .InsdDthResStCd = vbNullString
        .IssStCd = vbNullString
        .LstUpdtDtm = Now
        .LstUpdtUserId = vbNullString
        .PycoTypCd = vbNullString
        .ClmForResDthInd = False
        .ClmCompactClcnInd = False
    End With

    ' Also reset the saved "original" values for the Last Updated info
    m_dteLstUpdtDtm_Original = Now
    m_strLstUpdtUserId_Original = vbNullString
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
'UPDATED FOR SQL 2008
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
    '       Dates/lst_updt_dtm= adDBTimeStamp   (e.g., lst_updt_dtm or eff_dt)
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
    '       dom_lst_updt_dtm..for dates, this sets the system date (getdate()) as the default value on an Insert
    '    e. dom_lst_updt_id...to set the user's ACF2 (suser_sname()) as the default value on an Insert
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
                If Left$(strEditedDefaultValue, 1) = "'" And Len(strEditedDefaultValue) > 1 Then
                    strEditedDefaultValue = Right$(strEditedDefaultValue, Len(strEditedDefaultValue) - 1)
                End If
                ' Strip trailing single quote
                If Right$(strEditedDefaultValue, 1) = "'" And Len(strEditedDefaultValue) > 1 Then
                    strEditedDefaultValue = Left$(strEditedDefaultValue, Len(strEditedDefaultValue) - 1)
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

        ' DEF_LST_UPD_USER is typically used for the Lst_Updt_User_Id column,
        ' indicating to set it to the logged on user.
        If InStr(1, strDefaultValueToParse, "DEF_LST_UPD_ID") > 0 Then
            .DefaultValue = gconAppActive.LastLogOnUserID
        End If

        ' DEF_LST_UPD_DTM is typically used for the Lst_Updt_Dtm column,
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

        ' All Claims Interest tables use only the "DOM_IND" domain name for indicator columns.
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
        ' Debug.Print .ColName
        Select Case .DataType
            Case dbDecimal
                ' Save original values (will be used to code sproc parameters).
                .NumericScale = CByte(prstIn("SCALE").value)
                .Precision = CByte(prstIn("PRECISION").value)
                ' Save interpreted equivalents. These may be overriden in fnLoadColMetaData( ).
                .DecimalPositions = CInt(prstIn("SCALE").value)
                .DollarPositions = CInt(prstIn("PRECISION").value) - .DecimalPositions
                .MaxCharacters = 0
                If Right$(.ColName, 7) = "ZIP_CD" Then
                    ' Area Codes require 5 numeric positions if input.
                    ' They should be displayed via a TextBox control.
                    .Format = "&&&&&"
                    .Mask = vbNullString
                    .MaxCharacters = 5
                ElseIf Right$(.ColName, 7) = "ZIP4_CD" Then
                    ' Area Codes require 4 numeric positions if input.
                    ' They should be displayed via a TextBox control.
                    .Format = "&&&&"
                    .Mask = vbNullString
                    .MaxCharacters = 4
                Else
                    ' These should be displayed via a TextBox control.
                    '
                    ' NOTE: Other numeric items will need to overridden
                    ' on a table-specific basis in fnLoadColMetaData( ) since there is no
                    ' easy way to recognize and process these fields. For instance,
                    ' amounts (ending in _AMT) could have a varying number of dollar or
                    ' decimal positions.
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
                If Right$(.ColName, 7) = "SSN_NUM" Then
                    ' Social Security Numbers should be displayed via a MaskEdBox control.
                    .Format = vbNullString
                    .Mask = "###-##-####"
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
    ' Modified:     Berry Kropiwka - Added Compact Calc - 2019-09-27
    
    '!CUSTOMIZE!  There should be one Case statement for each table column.
    '             Each Case statement should reference a class property that
    '             corresponds to a table column, and the fnGetProperty return
    '             value should be set to the private variable (of type UDTColumn)
    '             that corresponds to that table column/class property.
    
    Const cstrCurrentProc       As String = "fnGetProperty"
    
    On Error GoTo PROC_ERR

    strTagIn = UCase$(strTagIn)
    Select Case strTagIn
        Case "ADMNSYSTCD"
            fnGetProperty = m_strAdmnSystCd
        Case "CLMID"
            fnGetProperty = m_lngClmId
        Case "CLMINSDDTHDT"
            fnGetProperty = m_dteClmInsdDthDt
        Case "CLMINSDFIRSTNM"
            fnGetProperty = m_strClmInsdFirstNm
        Case "CLMINSDLASTNM"
            fnGetProperty = m_strClmInsdLastNm
        Case "CLMINSDSSNNUM"
            fnGetProperty = m_strClmInsdSsnNum
        Case "CLMNUM"
            fnGetProperty = m_strClmNum
        Case "CLMPOLNUM"
            fnGetProperty = m_strClmPolNum
        Case "CLMPROOFDT"
            fnGetProperty = m_dteClmProofDt
        Case "CLMTOTCLMPDAMT"
            fnGetProperty = m_varClmTotClmPdAmt
        Case "CLMTOTDTHBPMTAMT"
            fnGetProperty = m_varClmTotDthbPmtAmt
        Case "CLMTOTINTAMT"
            fnGetProperty = m_varClmTotIntAmt
        Case "CLMTOTWTHLDAMT"
            fnGetProperty = m_varClmTotWthldAmt
        Case "INSDDTHRESSTCD"
            fnGetProperty = m_strInsdDthResStCd
        Case "ISSSTCD"
            fnGetProperty = m_strIssStCd
        Case "LSTUPDTDTM"
            fnGetProperty = m_dteLstUpdtDtm
        Case "LSTUPDTUSERID"
            fnGetProperty = m_strLstUpdtUserId
        Case "PYCOTYPCD"
            fnGetProperty = m_strPycoTypCd
        Case "CLMFORRESDTHIND"
            fnGetProperty = m_bClmForResDthInd
        Case "CLMCOMPACTCLCNIND"
            fnGetProperty = m_bClmCompactClcnInd
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
    ' Modified:     Berry Kropiwka - Added Compact Calc - 2019-09-27
    
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
    m_adwADO.GetMetaData_Columns "claim_t", rstMetaData

    ' NOTES:
    '    a. Phone Numbers and SSNs are Char fields, but we have to override the meta data to
    '       ensure only numbers can be input to them.
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
            ' Make sure SELECT CASE lists all table columns, including the LST_UPDT_xxx!
            'Debug.Print UCase$(rstMetaData("COLUMN_NAME").value)
            Select Case UCase$(rstMetaData("COLUMN_NAME").value)
                Case mcstrAdmnSystCd
                    fnGetColMetaData m_strAdmnSystCd, rstMetaData
                Case mcstrClmId
                    fnGetColMetaData m_lngClmId, rstMetaData
                Case mcstrClmInsdDthDt
                    fnGetColMetaData m_dteClmInsdDthDt, rstMetaData
                Case mcstrClmInsdFirstNm
                    fnGetColMetaData m_strClmInsdFirstNm, rstMetaData
                    m_strClmInsdFirstNm.ShouldForceToUppercase = True
                Case mcstrClmInsdLastNm
                    fnGetColMetaData m_strClmInsdLastNm, rstMetaData
                    m_strClmInsdLastNm.ShouldForceToUppercase = True
                Case mcstrClmInsdSsnNum
                    fnGetColMetaData m_strClmInsdSsnNum, rstMetaData
                    m_strClmInsdSsnNum.AllowableCharacters = "0123456789"
                Case mcstrClmNum
                    fnGetColMetaData m_strClmNum, rstMetaData
                    m_strClmNum.ShouldForceToUppercase = True
                Case mcstrClmPolNum
                    fnGetColMetaData m_strClmPolNum, rstMetaData
                    m_strClmPolNum.ShouldForceToUppercase = True
                Case mcstrClmProofDt
                    fnGetColMetaData m_dteClmProofDt, rstMetaData
                Case mcstrClmTotClmPdAmt
                    fnGetColMetaData m_varClmTotClmPdAmt, rstMetaData
                Case mcstrClmTotDthbPmtAmt
                    fnGetColMetaData m_varClmTotDthbPmtAmt, rstMetaData
                Case mcstrClmTotIntAmt
                    fnGetColMetaData m_varClmTotIntAmt, rstMetaData
                Case mcstrClmTotWthldAmt
                    fnGetColMetaData m_varClmTotWthldAmt, rstMetaData
                Case mcstrInsdDthResStCd
                    fnGetColMetaData m_strInsdDthResStCd, rstMetaData
                Case mcstrIssStCd
                    fnGetColMetaData m_strIssStCd, rstMetaData
                Case mcstrLstUpdtDtm
                    fnGetColMetaData m_dteLstUpdtDtm, rstMetaData
                Case mcstrLstUpdtUserId
                    fnGetColMetaData m_strLstUpdtUserId, rstMetaData
                Case mcstrPycoTypCd
                    fnGetColMetaData m_strPycoTypCd, rstMetaData
                Case mcstrClmForResDthInd
                    fnGetColMetaData m_bClmForResDthInd, rstMetaData
                Case mcstrClmCompactClcnInd
                    fnGetColMetaData m_bClmCompactClcnInd, rstMetaData
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
    m_adwADO.GetMetaData_PrimaryKeys "claim_t", rstMetaData

    With rstMetaData
        Do Until .EOF
            ' The SELECT CASE should list all table columns
            ' (though you could skip the LST_UPDT_xxx columns if you change the Case Else,
            '  since these would never be a key)
            Select Case UCase$(rstMetaData("COLUMN_NAME").value)
                Case mcstrAdmnSystCd
                    m_strAdmnSystCd.IsKey = True
                Case mcstrClmId
                    m_lngClmId.IsKey = True
                Case mcstrClmInsdDthDt
                    m_dteClmInsdDthDt.IsKey = True
                Case mcstrClmInsdFirstNm
                    m_strClmInsdFirstNm.IsKey = True
                Case mcstrClmInsdLastNm
                    m_strClmInsdLastNm.IsKey = True
                Case mcstrClmInsdSsnNum
                    m_strClmInsdSsnNum.IsKey = True
                Case mcstrClmNum
                    m_strClmNum.IsKey = True
                Case mcstrClmPolNum
                    m_strClmPolNum.IsKey = True
                Case mcstrClmProofDt
                    m_dteClmProofDt.IsKey = True
                Case mcstrClmTotClmPdAmt
                    m_varClmTotClmPdAmt.IsKey = True
                Case mcstrClmTotDthbPmtAmt
                    m_varClmTotDthbPmtAmt.IsKey = True
                Case mcstrClmTotIntAmt
                    m_varClmTotIntAmt.IsKey = True
                Case mcstrClmTotWthldAmt
                    m_varClmTotWthldAmt.IsKey = True
                Case mcstrInsdDthResStCd
                    m_strInsdDthResStCd.IsKey = True
                Case mcstrIssStCd
                    m_strIssStCd.IsKey = True
                Case mcstrLstUpdtDtm
                    m_dteLstUpdtDtm.IsKey = True
                Case mcstrLstUpdtUserId
                    m_strLstUpdtUserId.IsKey = True
                Case mcstrPycoTypCd
                    m_strPycoTypCd.IsKey = True
                Case mcstrClmForResDthInd
                    m_bClmForResDthInd.IsKey = True
                Case mcstrClmCompactClcnInd
                    m_bClmCompactClcnInd.IsKey = True
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
Private Function fnSelectRecord(ByVal lngKey1 As Long) As ADODB.Recordset
    '--------------------------------------------------------------------------
    ' Procedure:   fnSelectRecord
    ' Description: Selects a single record based on the value(s) in the
    '              properties that correspond to the table's key(s)
    '
    '              NOTE: For each table key, there should be a parameter
    '                    of the appropriate data type!
    '
    ' Parameters:
    '     lngKey1 (in) - the key to the table that should be retrieved
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
    Const cstrSproc                As String = "dbo.proc_claim_select"  ' Stored procedure to execute
    Dim rstTemp                    As ADODB.Recordset
    Dim prmReturnValue             As ADODB.Parameter
    Dim prmClmId                   As ADODB.Parameter

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
        ' Define the CLM_ID parameter
        Set prmClmId = .CreateParameter(Name:="@clm_id", _
                                         Type:=adInteger, _
                                         Direction:=adParamInput, _
                                         value:=fnNullIfZero(lngKey1))
        .Parameters.Append prmClmId

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
    fnFreeObject prmClmId
    
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case prmReturnValue
        Case gcRES_NERR_REC_NOT_FOUND
            ' 4027 = The specified record was not found in the database (@@1).
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_REC_NOT_FOUND, _
                                       mcstrName & cstrCurrentProc, _
                                       "Claim ID " & RTrim$(lngKey1)
            Resume PROC_EXIT
        Case gcRES_NERR_ERR_WHILE_TRYING_TO
            ' 4028 = An error occurred while attempting to @@1 this record.
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
