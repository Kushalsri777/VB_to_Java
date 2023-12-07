VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ctpyePayee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------
' Procedure:   ctpyePayee
' Description: Provides properties and methods to support the payee_t table values
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
'
'
' Procedures  :
'   Private     Class_Initialize()
'   Private     Class_Terminate()
'   Public      Property Get AllowableCharacters(ByVal strTagIn As String) As String
'   Public      Property Get CalcStCd() As String
'   Public      Property Let CalcStCd(ByVal strValue As String)
'   Public      Property Get ClmId() As Long
'   Public      Property Let ClmId(ByVal lngValue As Long)
'   Public      Property Get CurrentLookupRecordNumber() As Long
'   Public      Property Get DecimalPositions(ByVal strTagIn As String) As Integer
'   Public      Property Get DefaultValue(ByVal strTagIn As String) As Variant
'   Public      Property Get DollarPositions(ByVal strTagIn As String) As Integer
'   Public      Property Get Format(ByVal strTagIn As String) As String
'   Public      InitPayee(lngClmIdIn As Long)
'   Public      Property Get IsKey(ByVal strTagIn As String) As Boolean
'   Public      Property Get IsNullable(ByVal strTagIn As String) As Boolean
'   Public      Property Get LookupData() As ADODB.Recordset
'   Public      Property Get LookupData_Name() As Variant
'   Public      Property Get LookupIsAtBOF() As Boolean
'   Public      Property Get LookupIsAtEOF() As Boolean
'   Public      Property Get LookupRecordCount() As Long
'   Public      Property Get LstUpdtDtm() As Date
'   Public      Property Let LstUpdtDtm(ByVal NewValue As String)
'   Public      Property Get LstUpdtUserId() As String
'   Public      Property Let LstUpdtUserId(ByVal strValue As String)
'   Public      Property Get Mask(ByVal strTagIn As String) As String
'   Public      Property Get MaxCharacters(ByVal strTagIn As String) As Long
'   Public      Property Get PayeAddrLn1Txt() As String
'   Public      Property Let PayeAddrLn1Txt(ByVal strValue As String)
'   Public      Property Get PayeAddrLn2Txt() As String
'   Public      Property Let PayeAddrLn2Txt(ByVal strValue As String)
'   Public      Property Get PayeCareOfTxt() As String
'   Public      Property Let PayeCareOfTxt(ByVal strValue As String)
'   Public      Property Get PayeCityNmTxt() As String
'   Public      Property Let PayeCityNmTxt(ByVal strValue As String)
'   Public      Property Get PayeClmIntAmt() As Double
'   Public      Property Let PayeClmIntAmt(ByVal dblValue As Double)
'   Public      Property Get PayeClmIntRt() As Double
'   Public      Property Let PayeClmIntRt(ByVal dblValue As Double)
'   Public      Property Get PayeClmPdAmt() As Double
'   Public      Property Let PayeClmPdAmt(ByVal dblValue As Double)
'   Public      Property Get PayeDfltOvrdInd() As Boolean
'   Public      Property Let PayeDfltOvrdInd(ByVal bValue As Boolean)
'   Public      Property Get PayeDthbPmtAmt() As Double
'   Public      Property Let PayeDthbPmtAmt(ByVal dblValue As Double)
'   Public      Property Get PayeFullNm() As String
'   Public      Property Let PayeFullNm(ByVal strValue As String)
'   Public      Property Get PayeId() As Long
'   Public      Property Let PayeId(ByVal lngValue As Long)
'   Public      Property Get PayeIntDaysPdNum() As Integer
'   Public      Property Let PayeIntDaysPdNum(ByVal intValue As Integer)
'   Public      Property Get PayePmtDt() As Date
'   Public      Property Let PayePmtDt(ByVal dteValue As Date)
'   Public      Property Get PayeSsnTinNum() As String
'   Public      Property Let PayeSsnTinNum(ByVal strValue As String)
'   Public      Property Get PayeSsnTinTypCd() As String
'   Public      Property Let PayeSsnTinTypCd(ByVal strValue As String)
'   Public      Property Get PayeStCd() As String
'   Public      Property Let PayeStCd(ByVal strValue As String)
'   Public      Property Get PayeWthldAmt() As Double
'   Public      Property Let PayeWthldAmt(ByVal dblValue As Double)
'   Public      Property Get PayeWthldRt() As Double
'   Public      Property Let PayeWthldRt(ByVal dblValue As Double)
'   Public      Property Get PayeZip4Cd() As String
'   Public      Property Let PayeZip4Cd(ByVal strValue As String)
'   Public      Property Get PayeZipCd() As String
'   Public      Property Let PayeZipCd(ByVal strValue As String)
'   Public      Property Get ShouldForceToUppercase(ByVal strTagIn As String) As Boolean
'   Public      AddRecord() as Boolean
'   Public      CheckForAnotherUsersChanges(ByVal lngWhatOperation As enumWhatOperationIsBeingAttempted, _
'                   ByRef strACF2 As String) As Long
'   Public      DeleteRecord() As Boolean
'   Public      GetClmNumFromClmID(ByVal lngClmID As Long) As Variant
'   Public      GetLookupData() As Boolean
'   Public      GetPayeFullNmFromPayeID(ByVal lngPayeID As Long) As Variant
'   Public      GetPayeesForClaim(ByVal lngClmId As Long) As ADODB.Recordset
'   Public      GetClaimForPayeeClaim(ByVal lngClmId As Long) As ADODB.Recordset
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
Private Const mcstrName As String = "ctpyePayee."

'...............................................................................................
'!CUSTOMIZE!
' These are the private variables corresponding to PUBLIC properties.
' There should be one (of type udtColumn) for each column in the table that this class accesses.
'...............................................................................................
Private m_strCalcStCd                   As udtColumn
Private m_lngClmId                      As udtColumn
Private m_dteLstUpdtDtm                 As udtColumn
Private m_strLstUpdtUserId              As udtColumn
Private m_strPayeAddrLn1Txt             As udtColumn
Private m_strPayeAddrLn2Txt             As udtColumn
Private m_strPayeCareOfTxt              As udtColumn
Private m_strPayeCityNmTxt              As udtColumn
Private m_dblPayeClmIntAmt              As udtColumn
Private m_dblPayeClmIntRt               As udtColumn
Private m_dblPayeClmPdAmt               As udtColumn
Private m_bPayeDfltOvrdInd              As udtColumn
Private m_bstrPaye1099INTInd             As udtColumn  '' BZ4999 October 2013 Non US payee - SXS
Private m_dblPayeDthbPmtAmt             As udtColumn
Private m_strPayeFullNm                 As udtColumn
Private m_lngPayeId                     As udtColumn
Private m_intPayeIntDaysPdNum           As udtColumn
Private m_dtePayePmtDt                  As udtColumn
Private m_strPayeSsnTinNum              As udtColumn
Private m_strPayeSsnTinTypCd            As udtColumn
Private m_strPayeStCd                   As udtColumn
Private m_dblPayeWthldAmt               As udtColumn
Private m_dblPayeWthldRt                As udtColumn
Private m_strPayeZip4Cd                 As udtColumn
Private m_strPayeZipCd                  As udtColumn


'...............................................................................................
'!CUSTOMIZE!
' Create one Const for each column in the table, defining the table column to which it refers.
'...............................................................................................
Private Const mcstrCalcStCd             As String = "CALC_ST_CD"
Private Const mcstrClmId                As String = "CLM_ID"
Private Const mcstrLstUpdtDtm           As String = "LST_UPDT_DTM"
Private Const mcstrLstUpdtUserId        As String = "LST_UPDT_USER_ID"
Private Const mcstrPayeAddrLn1Txt       As String = "PAYE_ADDR_LN1_TXT"
Private Const mcstrPayeAddrLn2Txt       As String = "PAYE_ADDR_LN2_TXT"
Private Const mcstrPayeCareOfTxt        As String = "PAYE_CARE_OF_TXT"
Private Const mcstrPayeCityNmTxt        As String = "PAYE_CITY_NM_TXT"
Private Const mcstrPayeClmIntAmt        As String = "PAYE_CLM_INT_AMT"
Private Const mcstrPayeClmIntRt         As String = "PAYE_CLM_INT_RT"
Private Const mcstrPayeClmPdAmt         As String = "PAYE_CLM_PD_AMT"
Private Const mcstrPayeDfltOvrdInd      As String = "PAYE_DFLT_OVRD_IND"
Private Const mcstrPaye1099INTInd       As String = "PAYE_1099INT_IND"  '' BZ4999 October 2013 Non US payee - SXS
Private Const mcstrPayeDthbPmtAmt       As String = "PAYE_DTHB_PMT_AMT"
Private Const mcstrPayeFullNm           As String = "PAYE_FULL_NM"
Private Const mcstrPayeId               As String = "PAYE_ID"
Private Const mcstrPayeIntDaysPdNum     As String = "PAYE_INT_DAYS_PD_NUM"
Private Const mcstrPayePmtDt            As String = "PAYE_PMT_DT"
Private Const mcstrPayeSsnTinNum        As String = "PAYE_SSN_TIN_NUM"
Private Const mcstrPayeSsnTinTypCd      As String = "PAYE_SSN_TIN_TYP_CD"
Private Const mcstrPayeStCd             As String = "PAYE_ST_CD"
Private Const mcstrPayeWthldAmt         As String = "PAYE_WTHLD_AMT"
Private Const mcstrPayeWthldRt          As String = "PAYE_WTHLD_RT"
Private Const mcstrPayeZip4Cd           As String = "PAYE_ZIP4_CD"
Private Const mcstrPayeZipCd            As String = "PAYE_ZIP_CD"


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
Public Property Get CalcStCd() As String
    ' **************************************************************************
    ' Function  : Property Get CalcStCd
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : string
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get CalcStCd"
    On Error GoTo PROC_ERR

    CalcStCd = CStr(m_strCalcStCd.value)
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
Public Property Let CalcStCd(ByVal strValue As String)
    ' **************************************************************************
    ' Function  : Property Let CalcStCd
    ' Purpose   : Assigns new Value to property
    ' Parameters: ByVal strValue As String
    ' Returns   :
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let CalcStCd"
    On Error GoTo PROC_ERR

    m_strCalcStCd.value = strValue
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
Public Property Get ClmId() As Long
    ' **************************************************************************
    ' Function  : Property Get ClmId
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : Long
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
Public Property Get LookupData_Name() As Variant
    '--------------------------------------------------------------------------
    ' Procedure:   Get LookupData_Name
    ' Description: Return an array containing just the desired columns
    '              that should be populated in the Lookup fpCombo control
    ' Returns:     variant array
    '-----------------------------------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc       As String = "Get_LookupData_Name"
    Dim aRows()                 As Variant
    Dim strOriginalPayeFullNum  As String
    
    ' The .GetRows method changes the positioning within the m_rstLookup, which
    ' causes the "record x of y" label to incorrectly get set to the
    ' .RecordCount value. So, save the key of the current record, issue the
    ' .GetRows and then navigate back to the original position within the rst.
    With m_rstLookup
        strOriginalPayeFullNum = .Fields(mcstrPayeFullNm).value

        ' m_rstLookup is already sorted in the order we want: clm_num
        aRows = .GetRows(Rows:=adGetRowsRest, Start:=adBookmarkFirst)
    End With
    
    GetRelativeRecord strOriginalPayeFullNum, epdSameRecord
    LookupData_Name = aRows
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
    ' Returns   : Date
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
    ' Parameters: ByVal dteValue As Date
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
    ' Returns   : String
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
    ' Parameters: ByVal strValue As String
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
Public Property Get PayeAddrLn1Txt() As String
    ' **************************************************************************
    ' Function  : Property Get PayeAddrLn1Txt
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : string
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get PayeAddrLn1Txt"
    On Error GoTo PROC_ERR

    PayeAddrLn1Txt = CStr(m_strPayeAddrLn1Txt.value)
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
Public Property Let PayeAddrLn1Txt(ByVal strValue As String)
    ' **************************************************************************
    ' Function  : Property Let PayeAddrLn1Txt
    ' Purpose   : Assigns new Value to property
    ' Parameters: ByVal strValue As String
    ' Returns   :
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let PayeAddrLn1Txt"
    On Error GoTo PROC_ERR

    m_strPayeAddrLn1Txt.value = strValue
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
Public Property Get PayeAddrLn2Txt() As String
    ' **************************************************************************
    ' Function  : Property Get PayeAddrLn2Txt
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : string
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get PayeAddrLn2Txt"
    On Error GoTo PROC_ERR

    PayeAddrLn2Txt = CStr(m_strPayeAddrLn2Txt.value)
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
Public Property Let PayeAddrLn2Txt(ByVal strValue As String)
    ' **************************************************************************
    ' Function  : Property Let PayeAddrLn2Txt
    ' Purpose   : Assigns new Value to property
    ' Parameters: ByVal strValue As String
    ' Returns   :
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let PayeAddrLn2Txt"
    On Error GoTo PROC_ERR

    m_strPayeAddrLn2Txt.value = strValue
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
Public Property Get PayeCareOfTxt() As String
    ' **************************************************************************
    ' Function  : Property Get PayeCareOfTxt
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : string
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get PayeCareOfTxt"
    On Error GoTo PROC_ERR

    PayeCareOfTxt = CStr(m_strPayeCareOfTxt.value)
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
Public Property Let PayeCareOfTxt(ByVal strValue As String)
    ' **************************************************************************
    ' Function  : Property Let PayeCareOfTxt
    ' Purpose   : Assigns new Value to property
    ' Parameters: ByVal strValue As String
    ' Returns   :
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let PayeCareOfTxt"
    On Error GoTo PROC_ERR

    m_strPayeCareOfTxt.value = strValue
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
Public Property Get PayeCityNmTxt() As String
    ' **************************************************************************
    ' Function  : Property Get PayeCityNmTxt
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : string
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get PayeCityNmTxt"
    On Error GoTo PROC_ERR

    PayeCityNmTxt = CStr(m_strPayeCityNmTxt.value)
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
Public Property Let PayeCityNmTxt(ByVal strValue As String)
    ' **************************************************************************
    ' Function  : Property Let PayeCityNmTxt
    ' Purpose   : Assigns new Value to property
    ' Parameters: ByVal strValue As String
    ' Returns   :
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let PayeCityNmTxt"
    On Error GoTo PROC_ERR

    m_strPayeCityNmTxt.value = strValue
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
Public Property Get PayeClmIntAmt() As Double
    ' **************************************************************************
    ' Function  : Property Get PayeClmIntAmt
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : double
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get PayeClmIntAmt"
    On Error GoTo PROC_ERR

    PayeClmIntAmt = m_dblPayeClmIntAmt.value
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
Public Property Let PayeClmIntAmt(ByVal dblValue As Double)
    ' **************************************************************************
    ' Function  : Property Let PayeClmIntAmt
    ' Purpose   : Assigns new Value to property
    ' Parameters: ByVal dblValue
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let PayeClmIntAmt"
    On Error GoTo PROC_ERR

    m_dblPayeClmIntAmt.value = dblValue
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
Public Property Get PayeClmIntRt() As Double
    ' **************************************************************************
    ' Function  : Property Get PayeClmIntRt
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : double
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get PayeClmIntRt"
    On Error GoTo PROC_ERR

    PayeClmIntRt = m_dblPayeClmIntRt.value
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
Public Property Let PayeClmIntRt(ByVal dblValue As Double)
    ' **************************************************************************
    ' Function  : Property Let PayeClmIntRt
    ' Purpose   : Assigns new Value to property
    ' Parameters: ByVal dblValue
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let PayeClmIntRt"
    On Error GoTo PROC_ERR

    m_dblPayeClmIntRt.value = dblValue
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
Public Property Get PayeClmPdAmt() As Double
    ' **************************************************************************
    ' Function  : Property Get PayeClmPdAmt
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : double
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get PayeClmPdAmt"
    On Error GoTo PROC_ERR

    PayeClmPdAmt = m_dblPayeClmPdAmt.value
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
Public Property Let PayeClmPdAmt(ByVal dblValue As Double)
    ' **************************************************************************
    ' Function  : Property Let PayeClmPdAmt
    ' Purpose   : Assigns new Value to property
    ' Parameters: ByVal dblValue
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let PayeClmPdAmt"
    On Error GoTo PROC_ERR

    m_dblPayeClmPdAmt.value = dblValue
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
'''''''''''''''''''''''''''''''''''' '' BZ4999 October 2013 Non US payee - SXS
'////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Get Paye1099INTInd() As Boolean
    ' **************************************************************************
    ' Function  : Property Get PayeDfltOvrdInd
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : Boolean
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get Paye1099INTInd"
    On Error GoTo PROC_ERR

    ' Note that this field is stored as a Y or N.
    Paye1099INTInd = CBool(m_bstrPaye1099INTInd.value)
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
'' BZ4999 October 2013 Non US payee - SXS
Public Property Let Paye1099INTInd(ByVal bValue As Boolean)
    ' **************************************************************************
    ' Function  : Property Let Paye1099INTInd
    ' Purpose   : Assigns new Value to property
    ' Parameters: ByVal bValue As Boolean
    ' Returns   :
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let Paye1099INTInd"
    On Error GoTo PROC_ERR

    ' Note that this field is stored as a Y or N.
    m_bstrPaye1099INTInd.value = bValue
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
''''''''''''''''''''''''''''''''''''


'////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Get PayeDfltOvrdInd() As Boolean
    ' **************************************************************************
    ' Function  : Property Get PayeDfltOvrdInd
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : Boolean
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get PayeDfltOvrdInd"
    On Error GoTo PROC_ERR

    ' Note that this field is stored as a Y or N.
    PayeDfltOvrdInd = CBool(m_bPayeDfltOvrdInd.value)
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
Public Property Let PayeDfltOvrdInd(ByVal bValue As Boolean)
    ' **************************************************************************
    ' Function  : Property Let PayeDfltOvrdInd
    ' Purpose   : Assigns new Value to property
    ' Parameters: ByVal bValue As Boolean
    ' Returns   :
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let PayeDfltOvrdInd"
    On Error GoTo PROC_ERR

    ' Note that this field is stored as a Y or N.
    m_bPayeDfltOvrdInd.value = bValue
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
Public Property Get PayeDthbPmtAmt() As Double
    ' **************************************************************************
    ' Function  : Property Get PayeDthbPmtAmt
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : double
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get PayeDthbPmtAmt"
    On Error GoTo PROC_ERR

    PayeDthbPmtAmt = CDbl(m_dblPayeDthbPmtAmt.value)
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
Public Property Let PayeDthbPmtAmt(ByVal dblValue As Double)
    ' **************************************************************************
    ' Function  : Property Let PayeDthbPmtAmt
    ' Purpose   : Assigns new Value to property
    ' Parameters: ByVal dblValue As Double
    ' Returns   :
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let PayeDthbPmtAmt"
    On Error GoTo PROC_ERR

    m_dblPayeDthbPmtAmt.value = dblValue
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
Public Property Get PayeFullNm() As String
    ' **************************************************************************
    ' Function  : Property Get PayeFullNm
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : string
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get PayeFullNm"
    On Error GoTo PROC_ERR

    PayeFullNm = CStr(m_strPayeFullNm.value)
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
Public Property Let PayeFullNm(ByVal strValue As String)
    ' **************************************************************************
    ' Function  : Property Let PayeFullNm
    ' Purpose   : Assigns new Value to property
    ' Parameters: ByVal strValue As String
    ' Returns   :
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let PayeFullNm"
    On Error GoTo PROC_ERR

    m_strPayeFullNm.value = strValue
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
Public Property Get PayeId() As Long
    ' **************************************************************************
    ' Function  : Property Get PayeId
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : Long
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get PayeId"
    On Error GoTo PROC_ERR

    PayeId = CLng(m_lngPayeId.value)
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
Public Property Let PayeId(ByVal lngValue As Long)
    ' **************************************************************************
    ' Function  : Property Let PayeId
    ' Purpose   : Assigns new Value to property
    ' Parameters: ByVal lngValue As Long
    ' Returns   :
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let PayeId"
    On Error GoTo PROC_ERR

    m_lngPayeId.value = lngValue
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
Public Property Get PayeIntDaysPdNum() As Integer
    ' **************************************************************************
    ' Function  : Property Get PayeIntDaysPdNum
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : Integer
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get PayeIntDaysPdNum"
    On Error GoTo PROC_ERR

    PayeIntDaysPdNum = m_intPayeIntDaysPdNum.value
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
Public Property Let PayeIntDaysPdNum(ByVal intValue As Integer)
    ' **************************************************************************
    ' Function  : Property Let PayeIntDaysPdNum
    ' Purpose   : Assigns new Value (which could be NULL) to property
    ' Parameters: ByVal intValue
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let PayeIntDaysPdNum"
    On Error GoTo PROC_ERR

    m_intPayeIntDaysPdNum.value = intValue
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
Public Property Get PayePmtDt() As Date
    ' **************************************************************************
    ' Function  : Property Get PayePmtDt
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : Date
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get PayePmtDt"
    On Error GoTo PROC_ERR

    PayePmtDt = CDate(m_dtePayePmtDt.value)
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
Public Property Let PayePmtDt(ByVal dteValue As Date)
    ' **************************************************************************
    ' Function  : Property Let PayePmtDt
    ' Purpose   : Assigns new Value to property
    ' Parameters: ByVal dteValue As Date
    ' Returns   :
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let PayePmtDt"
    On Error GoTo PROC_ERR

    m_dtePayePmtDt.value = dteValue
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
Public Property Get PayeSsnTinNum() As String
    ' **************************************************************************
    ' Function  : Property Get PayeSsnTinNum
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : string
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get PayeSsnTinNum"
    On Error GoTo PROC_ERR

    PayeSsnTinNum = CStr(m_strPayeSsnTinNum.value)
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
Public Property Let PayeSsnTinNum(ByVal strValue As String)
    ' **************************************************************************
    ' Function  : Property Let PayeSsnTinNum
    ' Purpose   : Assigns new Value to property
    ' Parameters: ByVal strValue As String
    ' Returns   :
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let PayeSsnTinNum"
    On Error GoTo PROC_ERR

    m_strPayeSsnTinNum.value = strValue
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
Public Property Get PayeSsnTinTypCd() As String
    ' **************************************************************************
    ' Function  : Property Get PayeSsnTinTypCd
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : string
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get PayeSsnTinTypCd"
    On Error GoTo PROC_ERR

    PayeSsnTinTypCd = CStr(m_strPayeSsnTinTypCd.value)
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
Public Property Let PayeSsnTinTypCd(ByVal strValue As String)
    ' **************************************************************************
    ' Function  : Property Let PayeSsnTinTypCd
    ' Purpose   : Assigns new Value to property
    ' Parameters: ByVal strValue As String
    ' Returns   :
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let PayeSsnTinTypCd"
    On Error GoTo PROC_ERR

    m_strPayeSsnTinTypCd.value = strValue
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
Public Property Get PayeStCd() As String
    ' **************************************************************************
    ' Function  : Property Get PayeStCd
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : string
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get PayeStCd"
    On Error GoTo PROC_ERR

    PayeStCd = CStr(m_strPayeStCd.value)
    
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
Public Property Let PayeStCd(ByVal strValue As String)
    ' **************************************************************************
    ' Function  : Property Let PayeStCd
    ' Purpose   : Assigns new Value to property
    ' Parameters: ByVal strValue As String
    ' Returns   :
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let PayeStCd"
    On Error GoTo PROC_ERR

    m_strPayeStCd.value = strValue
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
Public Property Get PayeWthldAmt() As Double
    ' **************************************************************************
    ' Function  : Property Get PayeWthldAmt
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : Double
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get PayeWthldAmt"
    On Error GoTo PROC_ERR

    PayeWthldAmt = m_dblPayeWthldAmt.value
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
Public Property Let PayeWthldAmt(ByVal dblValue As Double)
    ' **************************************************************************
    ' Function  : Property Let PayeWthldAmt
    ' Purpose   : Assigns new Value to property
    ' Parameters: ByVal dblValue
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let PayeWthldAmt"
    On Error GoTo PROC_ERR

    m_dblPayeWthldAmt.value = dblValue
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
Public Property Get PayeWthldRt() As Double
    ' **************************************************************************
    ' Function  : Property Get PayeWthldRt
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : Double
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get PayeWthldRt"
    On Error GoTo PROC_ERR

    PayeWthldRt = m_dblPayeWthldRt.value
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
Public Property Let PayeWthldRt(ByVal dblValue As Double)
    ' **************************************************************************
    ' Function  : Property Let PayeWthldRt
    ' Purpose   : Assigns new Value to property
    ' Parameters: ByVal dblValue
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let PayeWthldRt"
    On Error GoTo PROC_ERR

    m_dblPayeWthldRt.value = dblValue
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
Public Property Get PayeZip4Cd() As String
    ' **************************************************************************
    ' Function  : Property Get PayeZip4Cd
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : string
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get PayeZip4Cd"
    On Error GoTo PROC_ERR

    PayeZip4Cd = CStr(m_strPayeZip4Cd.value)
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
Public Property Let PayeZip4Cd(ByVal strValue As String)
    ' **************************************************************************
    ' Function  : Property Let PayeZip4Cd
    ' Purpose   : Assigns new Value to property
    ' Parameters: ByVal strValue As String
    ' Returns   :
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let PayeZip4Cd"
    On Error GoTo PROC_ERR

    m_strPayeZip4Cd.value = strValue
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
Public Property Get PayeZipCd() As String
    ' **************************************************************************
    ' Function  : Property Get PayeZipCd
    ' Purpose   : Retrieves current value from property
    ' Parameters: N/A
    ' Returns   : string
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Get PayeZipCd"
    On Error GoTo PROC_ERR

    PayeZipCd = CStr(m_strPayeZipCd.value)
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
Public Property Let PayeZipCd(ByVal strValue As String)
    ' **************************************************************************
    ' Function  : Property Let PayeZipCd
    ' Purpose   : Assigns new Value to property
    ' Parameters: ByVal strValue As String
    ' Returns   :
    ' **************************************************************************
    Const cstrCurrentProc   As String = "Property Let PayeZipCd"
    On Error GoTo PROC_ERR

    m_strPayeZipCd.value = strValue
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
    ' Returns   : Boolean
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
    
    Const cstrCurrentProc           As String = "AddRecord"
    Const cstrSproc                 As String = "dbo.proc_payee_insert"  ' Stored procedure to execute
    Dim rstTemp                     As ADODB.Recordset
    Dim bSuccessful                 As Boolean
    Dim prmReturnValue              As ADODB.Parameter
    Dim prmInvalid_Key              As ADODB.Parameter
    Dim prmCalcStCd                 As ADODB.Parameter
    Dim prmClmId                    As ADODB.Parameter
    Dim prmPayeAddrLn1Txt           As ADODB.Parameter
    Dim prmPayeAddrLn2Txt           As ADODB.Parameter
    Dim prmPayeCareOfTxt            As ADODB.Parameter
    Dim prmPayeCityNmTxt            As ADODB.Parameter
    Dim prmPayeClmIntAmt            As ADODB.Parameter
    Dim prmPayeClmIntRt             As ADODB.Parameter
    Dim prmPayeClmPdAmt             As ADODB.Parameter
    Dim prmPayeDfltOvrdInd          As ADODB.Parameter
    Dim prmPayeDthbPmtAmt           As ADODB.Parameter
    Dim prmPayeFullNm               As ADODB.Parameter
    Dim prmPayeIntDaysPdNum         As ADODB.Parameter
    Dim prmPayePmtDt                As ADODB.Parameter
    Dim prmPayeSsnTinNum            As ADODB.Parameter
    Dim prmPayeSsnTinTypCd          As ADODB.Parameter
    Dim prmPayeStCd                 As ADODB.Parameter
    Dim prmPayeWthldAmt             As ADODB.Parameter
    Dim prmPayeWthldRt              As ADODB.Parameter
    Dim prmPayeZip4Cd               As ADODB.Parameter
    Dim prmPayeZipCd                As ADODB.Parameter
    Dim prmNew_Id                   As ADODB.Parameter
    Dim prmPaye1099INTInd           As ADODB.Parameter   '' BZ4999 October 2013 Non US payee - SXS

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
        ' Define the CALC_ST_CD parameter
        Set prmCalcStCd = .CreateParameter(Name:="@calc_st_cd", _
                                         Type:=adChar, _
                                         Direction:=adParamInput, _
                                         Size:=2, _
                                         value:=fnNullIfZLS(varIn:=CalcStCd, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmCalcStCd

        ' ---Parameter #3---
        ' Define the CLM_ID parameter
        Set prmClmId = .CreateParameter(Name:="@clm_id", _
                                        Type:=adInteger, _
                                        Direction:=adParamInput, _
                                        value:=ClmId)
        .Parameters.Append prmClmId

        ' ---Parameter #4---
        ' Define the PAYE_ADR_LN1_TXT parameter
        Set prmPayeAddrLn1Txt = .CreateParameter(Name:="@paye_addr_ln1_txt", _
                                         Type:=adVarChar, _
                                         Direction:=adParamInput, _
                                         Size:=40, _
                                         value:=fnNullIfZLS(varIn:=PayeAddrLn1Txt, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmPayeAddrLn1Txt

        ' ---Parameter #5---
        ' Define the PAYE_ADR_LN2_TXT parameter
        Set prmPayeAddrLn2Txt = .CreateParameter(Name:="@paye_addr_ln2_txt", _
                                         Type:=adVarChar, _
                                         Direction:=adParamInput, _
                                         Size:=40, _
                                         value:=fnNullIfZLS(varIn:=PayeAddrLn2Txt, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmPayeAddrLn2Txt

        ' ---Parameter #6---
        ' Define the PAYE_ADR_LN1_TXT parameter
        Set prmPayeCareOfTxt = .CreateParameter(Name:="@paye_care_of_txt", _
                                         Type:=adVarChar, _
                                         Direction:=adParamInput, _
                                         Size:=40, _
                                         value:=fnNullIfZLS(varIn:=PayeCareOfTxt, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmPayeCareOfTxt

        ' ---Parameter #7---
        ' Define the PAYE_CITY_NM_TXT parameter
        Set prmPayeCityNmTxt = .CreateParameter(Name:="@paye_city_nm_txt", _
                                         Type:=adVarChar, _
                                         Direction:=adParamInput, _
                                         Size:=25, _
                                         value:=fnNullIfZLS(varIn:=PayeCityNmTxt, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmPayeCityNmTxt

        ' ---Parameter #8---
        ' Define the PAYE_CLM_INT_AMT parameter
        Set prmPayeClmIntAmt = .CreateParameter(Name:="@paye_clm_int_amt", _
                                          Type:=adNumeric, _
                                          Direction:=adParamInput, _
                                          value:=PayeClmIntAmt)
        With m_dblPayeClmIntAmt
            prmPayeClmIntAmt.Precision = .Precision
            prmPayeClmIntAmt.NumericScale = .NumericScale
        End With
        .Parameters.Append prmPayeClmIntAmt

        ' ---Parameter #9---
        ' Define the PAYE_CLM_INT_RT parameter
        Set prmPayeClmIntRt = .CreateParameter(Name:="@paye_clm_int_rt", _
                                          Type:=adNumeric, _
                                          Direction:=adParamInput, _
                                          value:=PayeClmIntRt)
        With m_dblPayeClmIntRt
            prmPayeClmIntRt.Precision = .Precision
            prmPayeClmIntRt.NumericScale = .NumericScale
        End With
        .Parameters.Append prmPayeClmIntRt

        ' ---Parameter #10---
        ' Define the PAYE_CLM_PD_AMT parameter
        Set prmPayeClmPdAmt = .CreateParameter(Name:="@paye_clm_pd_amt", _
                                          Type:=adNumeric, _
                                          Direction:=adParamInput, _
                                          value:=PayeClmPdAmt)
        With m_dblPayeClmPdAmt
            prmPayeClmPdAmt.Precision = .Precision
            prmPayeClmPdAmt.NumericScale = .NumericScale
        End With
        .Parameters.Append prmPayeClmPdAmt

        ' ---Parameter #11---
        ' Define the PAYE_DFLT_OVRD_IND parameter
        Set prmPayeDfltOvrdInd = .CreateParameter(Name:="@paye_dflt_ovrd_ind", _
                                         Type:=adChar, _
                                         Direction:=adParamInput, _
                                         Size:=1, _
                                         value:=fnBoolToYN(bIn:=PayeDfltOvrdInd))
        .Parameters.Append prmPayeDfltOvrdInd

        ' ---Parameter #12---
        ' Define the PAYE_DTHB_PMT_AMT parameter
        Set prmPayeDthbPmtAmt = .CreateParameter(Name:="@paye_dthb_pmt_amt", _
                                          Type:=adNumeric, _
                                          Direction:=adParamInput, _
                                          value:=PayeDthbPmtAmt)
        With m_dblPayeDthbPmtAmt
            prmPayeDthbPmtAmt.Precision = .Precision
            prmPayeDthbPmtAmt.NumericScale = .NumericScale
        End With
        .Parameters.Append prmPayeDthbPmtAmt

        ' ---Parameter #13---
        ' Define the PAYE_FULL_NM parameter
        Set prmPayeFullNm = .CreateParameter(Name:="@paye_full_nm", _
                                         Type:=adVarChar, _
                                         Direction:=adParamInput, _
                                         Size:=40, _
                                         value:=fnNullIfZLS(varIn:=PayeFullNm, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmPayeFullNm

        ' ---Parameter #14---
        ' Define the PAYE_INT_DAYS_PD_NUM parameter
        Set prmPayeIntDaysPdNum = .CreateParameter(Name:="@paye_int_days_pd_num", _
                                        Type:=adInteger, _
                                        Direction:=adParamInput, _
                                        value:=PayeIntDaysPdNum)
        .Parameters.Append prmPayeIntDaysPdNum

        ' ---Parameter #15---
        ' Define the PAYE_PMT_DT parameter
        Set prmPayePmtDt = .CreateParameter(Name:="@paye_pmt_dt", _
                                         Type:=adDBTimeStamp, _
                                         Direction:=adParamInput, _
                                         Size:=16, _
                                         value:=PayePmtDt)
        .Parameters.Append prmPayePmtDt
        
        ' ---Parameter #16---
        ' Define the PAYE_SSN_TIN_NUM parameter
        Set prmPayeSsnTinNum = .CreateParameter(Name:="@paye_ssn_tin_num", _
                                         Type:=adChar, _
                                         Direction:=adParamInput, _
                                         Size:=9, _
                                         value:=fnNullIfZLS(varIn:=PayeSsnTinNum, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmPayeSsnTinNum
        
        ' ---Parameter #17---
        ' Define the PAYE_SSN_TIN_TYP_CD parameter
        Set prmPayeSsnTinTypCd = .CreateParameter(Name:="@paye_ssn_tin_typ_cd", _
                                         Type:=adChar, _
                                         Direction:=adParamInput, _
                                         Size:=1, _
                                         value:=fnNullIfZLS(varIn:=PayeSsnTinTypCd, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmPayeSsnTinTypCd
        
        ' ---Parameter #18---
        ' Define the PAYE_ST_CD parameter
        Set prmPayeStCd = .CreateParameter(Name:="@paye_st_cd", _
                                         Type:=adChar, _
                                         Direction:=adParamInput, _
                                         Size:=2, _
                                         value:=fnNullIfZLS(varIn:=PayeStCd, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmPayeStCd
        
        ' ---Parameter #19---
        ' Define the PAY_WTHLD_AMT parameter
        Set prmPayeWthldAmt = .CreateParameter(Name:="@paye_wthld_amt", _
                                          Type:=adNumeric, _
                                          Direction:=adParamInput, _
                                          value:=PayeWthldAmt)
        With m_dblPayeWthldAmt
            prmPayeWthldAmt.Precision = .Precision
            prmPayeWthldAmt.NumericScale = .NumericScale
        End With
        .Parameters.Append prmPayeWthldAmt

        ' ---Parameter #20---
        ' Define the PAY_WTHLD_RT parameter
        Set prmPayeWthldRt = .CreateParameter(Name:="@paye_wthld_rt", _
                                          Type:=adNumeric, _
                                          Direction:=adParamInput, _
                                          value:=PayeWthldRt)
        With m_dblPayeWthldRt
            prmPayeWthldRt.Precision = .Precision
            prmPayeWthldRt.NumericScale = .NumericScale
        End With
        .Parameters.Append prmPayeWthldRt

       
        ' ---Parameter #21---
        ' Define the PAYE_ZIP4_cd parameter
        Set prmPayeZip4Cd = .CreateParameter(Name:="@paye_zip4_cd", _
                                         Type:=adChar, _
                                         Direction:=adParamInput, _
                                         Size:=4, _
                                         value:=fnNullIfZLS(varIn:=PayeZip4Cd, bHandleEmbeddedQuotes:=True))
        '' BZ4999 October 2013 Non US payee - SXS
        If prmPayeStCd = "ZZ" Then
           prmPayeZip4Cd = " "
        End If
        .Parameters.Append prmPayeZip4Cd
        
        ' ---Parameter #22---
        ' Define the PAYE_ZIP_CD parameter
        Set prmPayeZipCd = .CreateParameter(Name:="@paye_zip_cd", _
                                         Type:=adChar, _
                                         Direction:=adParamInput, _
                                         Size:=5, _
                                         value:=fnNullIfZLS(varIn:=PayeZipCd, bHandleEmbeddedQuotes:=True))
             '' BZ4999 October 2013 Non US payee - SXS
        If prmPayeStCd = "ZZ" Then
           prmPayeZipCd = " "
        End If
        .Parameters.Append prmPayeZipCd
        
        ' ---Parameter #23---  '' BZ4999 October 2013 Non US payee - SXS
        ' Define the PAYE_DFLT_OVRD_IND parameter
        Set prmPaye1099INTInd = .CreateParameter(Name:="@Paye_1099INT_Ind", _
                                         Type:=adChar, _
                                         Direction:=adParamInput, _
                                         Size:=1, _
                                         value:=fnBoolToYN(bIn:=Paye1099INTInd))
        .Parameters.Append prmPaye1099INTInd

        
        ' ---Parameter #24---
        ' Define the Invalid_Key output parameter, which reflects *which* foreign
        ' key violation was encountered.
        Set prmInvalid_Key = .CreateParameter(Name:="@Invalid_Key", _
                                                  Type:=adVarChar, _
                                                  Size:=255, _
                                                  Direction:=adParamOutput)
        .Parameters.Append prmInvalid_Key

        ' ---Parameter #25---
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
    bSuccessful = GetRelativeRecord(GetPayeFullNmFromPayeID(prmNew_Id.value), epdSameRecord)

    AddRecord = True
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
    
    ' Clean-up statements go here
    fnFreeRecordset rstTemp
    fnFreeObject prmReturnValue
    fnFreeObject prmInvalid_Key
    fnFreeObject prmCalcStCd
    fnFreeObject prmClmId
    fnFreeObject prmPayeAddrLn1Txt
    fnFreeObject prmPayeAddrLn2Txt
    fnFreeObject prmPayeCareOfTxt
    fnFreeObject prmPayeCityNmTxt
    fnFreeObject prmPayeClmIntAmt
    fnFreeObject prmPayeClmIntRt
    fnFreeObject prmPayeClmPdAmt
    fnFreeObject prmPayeDfltOvrdInd
    fnFreeObject prmPaye1099INTInd  '' BZ4999 October 2013 Non US payee - SXS
    fnFreeObject prmPayeDthbPmtAmt
    fnFreeObject prmPayeFullNm
    fnFreeObject prmPayeIntDaysPdNum
    fnFreeObject prmPayePmtDt
    fnFreeObject prmPayeSsnTinNum
    fnFreeObject prmPayeSsnTinTypCd
    fnFreeObject prmPayeStCd
    fnFreeObject prmPayeWthldAmt
    fnFreeObject prmPayeWthldRt
    fnFreeObject prmPayeZip4Cd
    fnFreeObject prmPayeZipCd
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
                                       "Payee Name " & RTrim$(PayeFullNm) & "/Claim Number " & GetClmNumFromClmID(ClmId)
            Resume PROC_EXIT
        Case gcRES_NERR_KEY_NOT_FOUND
            ' 4032 = The @@1 specified (@@2) is no longer defined in the @@3 table. Please choose a different value.
            If UCase$(prmInvalid_Key) = "CALC_ST_CD" Then
                gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_KEY_NOT_FOUND, _
                                           mcstrName & cstrCurrentProc, _
                                           "Calc State", RTrim$(CalcStCd), "STATE_T"
                Resume PROC_EXIT
            ElseIf UCase$(prmInvalid_Key) = "PAYE_ST_CD" Then
                ' PAYE_ST_CD
                gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_KEY_NOT_FOUND, _
                                           mcstrName & cstrCurrentProc, _
                                           "Payee State ", RTrim$(PayeStCd), "STATE_T"
                Resume PROC_EXIT
            Else
                ' CLM_ID
                gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_KEY_NOT_FOUND, _
                                           mcstrName & cstrCurrentProc, _
                                           "Claim ID ", CStr(ClmId), "CLAIM_T"
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
    Set rstSingleRecord_Fresh = fnSelectRecord(PayeId)
    
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
    Const cstrSproc                As String = "dbo.proc_payee_delete"  ' Stored procedure to execute
    Dim bSuccessful                As Boolean
    Dim adwTemp                    As cadwADOWrapper
    Dim rstTemp                    As ADODB.Recordset
    Dim prmReturnValue             As ADODB.Parameter
    Dim prmPayeId                  As ADODB.Parameter
    Dim prmDependent_Table         As ADODB.Parameter
    Dim strPayeFullNm              As String

    On Error GoTo PROC_ERR


    Set adwTemp = New cadwADOWrapper
    adwTemp.CommandSetSproc strSprocName:=cstrSproc

    ' Save the Payee Full name associated with the Payee that's going to be deleted.
    ' We'll need it afterwards to position to the previous record.
    strPayeFullNm = GetPayeFullNmFromPayeID(PayeId)

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
        ' Define the paye_id parameter
        Set prmPayeId = .CreateParameter(Name:="@paye_id", _
                                        Type:=adInteger, _
                                        Direction:=adParamInput, _
                                        value:=PayeId)
        .Parameters.Append prmPayeId

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
    bSuccessful = GetRelativeRecord(strPayeFullNm, epdPreviousRecord)
    
    DeleteRecord = True
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
    
    ' Clean-up statements go here
    fnFreeRecordset rstTemp
    fnFreeObject adwTemp
    fnFreeObject prmReturnValue
    fnFreeObject prmPayeId
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
                                       "Payee ID " & RTrim$(PayeId)
            Resume PROC_EXIT
        Case gcRES_NERR_ERR_WHILE_TRYING_TO
            ' 4028 = An error occurred while attempting to @@1 this record.
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_ERR_WHILE_TRYING_TO, _
                                       mcstrName & cstrCurrentProc, _
                                       "delete"
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
                                         Type:=adChar, _
                                         Direction:=adParamInputOutput, _
                                         Size:=15, _
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
Public Function GetLookupData(ByVal lngClmIdIn As Long) As Boolean
    '--------------------------------------------------------------------------
    ' Procedure:   fnGetLookupData_ClaimNum
    ' Description: Get all rows with the specified Claim ID but only particular columns
    ' Returns:     True if successful; False otherwise
    ' Params:      None
    ' Date:        01/07/2002
    '-----------------------------------------------------------------------------
    
    '!CUSTOMIZE!  Customize the name of the stored procedure, the number and names
    '             of the parameters and, perhaps, the return values trapped in the
    '             error handler.
    
    Const cstrCurrentProc       As String = "fnGetLookupData_ClaimNum"
    Const cstrSproc             As String = "dbo.proc_payee_lu_select"
    Dim prmReturnValue          As ADODB.Parameter
    Dim prmClmId                As ADODB.Parameter

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
        ' Define the CLM_ID parameter
        Set prmClmId = .CreateParameter(Name:="@clm_id", _
                                        Type:=adInteger, _
                                        Direction:=adParamInput, _
                                        value:=lngClmIdIn)
        .Parameters.Append prmClmId

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
        Case gcRES_NERR_REC_NOT_FOUND
            ' 4027 = The specified record was not found in the database (@@1).
            ' Wipe out any trace of this error, but return False so the caller
            ' knows to go into Add mode if desired. NOTE: The caller can also
            ' identify this by looking at the LookupRecordCount public property.
            Err.Clear
            GetLookupData = False
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
Public Function GetPayeFullNmFromPayeID(ByVal lngPayeID As Long) As Variant
    '--------------------------------------------------------------------------
    ' Procedure:   GetPayeFullNmFromPayeID
    ' Description: Query database for the Payee Full Name for a specified Payee ID
    ' Params:
    '               lngPayeID  (in)  The Payee ID to translate.
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc               As String = "GetPayeFullNmFromPayeID"
    Const cstrSproc                     As String = "dbo.proc_paye_full_nm_lu_select"
    Dim rstTemp                         As ADODB.Recordset
    Dim prmReturnValue                  As ADODB.Parameter
    Dim prmPayeFullNm                   As ADODB.Parameter
    Dim prmPayeId                       As ADODB.Parameter
    
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
        Set prmPayeId = .CreateParameter(Name:="paye_id", _
                                         Type:=adInteger, _
                                         Direction:=adParamInput, _
                                         value:=lngPayeID)
        .Parameters.Append prmPayeId

        ' ---Parameter #3---
        Set prmPayeFullNm = .CreateParameter(Name:="paye_full_nm", _
                                         Type:=adVarChar, _
                                         Direction:=adParamInputOutput, _
                                         Size:=40, _
                                         value:=Null)
        .Parameters.Append prmPayeFullNm


        Set rstTemp = .Execute()
    End With

    GetPayeFullNmFromPayeID = prmPayeFullNm.value
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    fnFreeObject prmReturnValue
    fnFreeObject prmPayeId
    fnFreeObject prmPayeFullNm
    
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
Public Function GetPayeesForClaim(ByVal lngClmID As Long) As ADODB.Recordset
    '--------------------------------------------------------------------------
    ' Procedure:   GetPayeesForClaim
    ' Description: Returns a recordset containing selected data on each Payee
    '              associated with the specified CLM_ID
    '
    ' Parameters:
    '     lngClmId (in) - the CLM_ID of the desired Claim
    '
    ' Returns:     A disconnected ADODB.Recordset containing selected table
    '              columns for the specified key
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc          As String = "GetPayeesForClaim"
    Const cstrSproc                As String = "dbo.proc_payee_select4"  ' Stored procedure to execute
    Dim rstTemp                    As ADODB.Recordset
    Dim prmReturnValue             As ADODB.Parameter
    Dim prmClmId                  As ADODB.Parameter

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
        ' Define the paye_id parameter
        Set prmClmId = .CreateParameter(Name:="@clm_id", _
                                         Type:=adInteger, _
                                         Direction:=adParamInput, _
                                         value:=fnNullIfZero(lngClmID))
        .Parameters.Append prmClmId

        Set rstTemp = .Execute()
    End With
    
    rstTemp.ActiveConnection = Nothing
    Set GetPayeesForClaim = rstTemp
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
                                       "Claim ID " & RTrim$(lngClmID) & "/Claim Number " & GetClmNumFromClmID(lngClmID)
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

''////////////////////////////////////////////////////////////////////////////////////////////////
'Public Function GetClaimForPayeeClaim(ByVal lngClmID As Long) As ADODB.Recordset
'    '--------------------------------------------------------------------------
'    ' Procedure:   GetClaimForPayeeClaim
'    ' Description: Returns a recordset containing selected data on the Claim for the Payee
'    '              associated with the specified CLM_ID
'    '
'    ' Parameters:
'    '     lngClmId (in) - the CLM_ID of the desired Claim
'    '
'    ' Returns:     A disconnected ADODB.Recordset containing selected table
'    '              columns for the specified key
'    '
'    ' Created:     Berry Kropiwka 2019-09-27
'    '-----------------------------------------------------------------------------
'    Const cstrCurrentProc          As String = "GetClaimForPayeeClaim"
'    Const cstrSproc                As String = "dbo.proc_claim_select"  ' Stored procedure to execute
'    Dim rstTemp                    As ADODB.Recordset
'    Dim prmReturnValue             As ADODB.Parameter
'    Dim prmClmId                   As ADODB.Parameter
'
'    On Error GoTo PROC_ERR
'
'    If Not (m_adwADO.CommandSetSproc(cstrSproc)) Then
'        GoTo PROC_EXIT
'    End If
'
'    ' For Char/VarChar fields,
'    '     * Use fnNullIfZLS to ensure Nulls are appropriately handled.
'    '     * Do *not* set the optional 2nd parameter to fnNullIfZLS to True.
'    ' For numeric fields,
'    '     * Use fnNullIfZero to ensure Nulls are appropriately handled.
'    ' For Y/N fields,
'    '     * Use fnBoolToYN to ensure True/False is appropriately translated.
'
'    With m_adwADO.ADOCommand
'        ' ---Parameter #1---
'        ' Define the return value that represents the error code (i.e. reason) why
'        ' the stored procedure failed.
'        Set prmReturnValue = .CreateParameter(Name:="@return_value", _
'                                              Type:=adInteger, _
'                                              Direction:=adParamReturnValue, _
'                                              value:=Null)
'        .Parameters.Append prmReturnValue
'
'        ' ---Parameter #2---
'        ' Define the paye_id parameter
'        Set prmClmId = .CreateParameter(Name:="@clm_id", _
'                                         Type:=adInteger, _
'                                         Direction:=adParamInput, _
'                                         value:=fnNullIfZero(lngClmID))
'        .Parameters.Append prmClmId
'
'        Set rstTemp = .Execute()
'    End With
'
'    rstTemp.ActiveConnection = Nothing
'    Set GetClaimForPayeeClaim = rstTemp
'PROC_EXIT:
'    On Error GoTo 0     ' Disable error handler
'
'    ' Clean-up statements go here
'
'    ' Do *not* do "fnFreeRecordset rstTemp" since this will cause the recordset returned
'    ' by this function to be wiped out as well!
'    fnFreeObject prmReturnValue
'    fnFreeObject prmClmId
'
'    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
'        gerhApp.PropagateError mcstrName & cstrCurrentProc
'    End If
'    Exit Function
'PROC_ERR:
'    Select Case prmReturnValue
'        Case gcRES_NERR_REC_NOT_FOUND
'            ' 4027 = The specified record was not found in the database (@@1).
'            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_REC_NOT_FOUND, _
'                                       mcstrName & cstrCurrentProc, _
'                                       "Claim ID " & RTrim$(lngClmID) & "/Claim Number " & GetClmNumFromClmID(lngClmID)
'            Resume PROC_EXIT
'        Case gcRES_NERR_ERR_WHILE_TRYING_TO
'            ' 4028 = An error occurred while attempting to @@1 this record.
'            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_ERR_WHILE_TRYING_TO, _
'                                       mcstrName & cstrCurrentProc, _
'                                       "locate"
'            Resume PROC_EXIT
'    End Select
'
'    ' If any other errors exist, i.e. in Err object, then let it fall through into default error handling.
'
'    Select Case Err.Number
'        Case -2147217900 ' Object not found
'            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_FERR_SPROC_NOT_FOUND, _
'                                       mcstrName & cstrCurrentProc, _
'                                       cstrSproc
'            Resume PROC_EXIT
'        Case Else
'            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
'    End Select
'    Resume PROC_EXIT
'End Function

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
    '                    This table's true key is PAYE_ID, but the Insured screen
    '                    shows claims in PAYE_FULL_NM sequence. Hence, the key this
    '                    proc uses is PAYE_FULL_NM.
    '
    ' Params:
    '     strKey1              (in) = PayeFullNm value from which to do the relative
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

    ' In case the Payee Name contains an embedded single quote (e.g. O'Dell),
    ' change it to 2 single quotes so SQL will interpret it correctly
    strKey1 = Replace(strKey1, "'", "''")

    '...........................................................................
    ' Refresh the lookup data (m_rstLookupData) so other's changes
    ' --and our own-- are now reflected in it. This resets the Lookup data,
    ' record count, and current record number, and leaves the Lookup recordset
    ' positioned to the first record (if there are records) or BOF (if there are
    ' no records).
    '...........................................................................
    GetLookupData ClmId

    With m_rstLookup
        Select Case lngPositionDirection
            Case epdPreviousRecord
                ' Make visible only those rows with keys prior to the specified key
                .Filter = "paye_full_nm < '" & strKey1 & "'"
                If .RecordCount <> 0 Then
                    ' Grab the last record. The one with the highest key less than the
                    ' specified key is the one we want.
                    m_adwADO.MoveLast m_rstLookup
                    
                    '!CUSTOMIZE! The constant referenced below should refer to the key field.
                    strKey1ForNewRec = .Fields(mcstrPayeFullNm).value
                    ' Accommodate embedded single quotes (e.g. O'Dell) by changing it to 2 single quotes
                    strKey1ForNewRec = Replace(strKey1ForNewRec, "'", "''")
                    
                    ' Okay, we got it. Now make all records in the Lookup recordset
                    ' visible again, and then reposition to the new record.
                    .Filter = adFilterNone
                    m_adwADO.MoveFirst m_rstLookup
                    .Find "paye_full_nm = '" & strKey1ForNewRec & "'"
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
                .Filter = "paye_full_nm > '" & strKey1 & "'"
                If .RecordCount <> 0 Then
                    ' Grab the first record. The one with the lowest key higher than the
                    ' specified key is the one we want.
                    m_adwADO.MoveFirst m_rstLookup
                    
                    '!CUSTOMIZE! The constant referenced below should refer to the key field.
                    strKey1ForNewRec = .Fields(mcstrPayeFullNm).value
                    ' Accommodate embedded single quotes (e.g. O'Dell) by changing it to 2 single quotes
                    strKey1ForNewRec = Replace(strKey1ForNewRec, "'", "''")
                    
                    ' Okay, we got it. Now make all records in the Lookup recordset
                    ' visible again, and then reposition to the new record.
                    .Filter = adFilterNone
                    m_adwADO.MoveFirst m_rstLookup
                    .Find "paye_full_nm = '" & strKey1ForNewRec & "'"
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
                .Find "paye_full_nm = '" & strKey1 & "'"
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
            GetSingleRecord lngKey1:=.Fields(mcstrPayeId).value, bSynchLookupRST:=True
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
    '    lngKey1         (in) = represents the primary key for the table (paye_id)
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
        .CalcStCd = fnZLSIfNull(rstTemp!calc_st_cd)
        .ClmId = fnZeroIfNull(rstTemp!clm_id)
        .LstUpdtDtm = rstTemp!lst_updt_dtm
        .LstUpdtUserId = fnZLSIfNull(rstTemp!lst_updt_user_id)
        .PayeAddrLn1Txt = fnZLSIfNull(rstTemp!paye_addr_ln1_txt)
        .PayeAddrLn2Txt = fnZLSIfNull(rstTemp!paye_addr_ln2_txt)
        .PayeCareOfTxt = fnZLSIfNull(rstTemp!paye_care_of_txt)
        .PayeCityNmTxt = fnZLSIfNull(rstTemp!paye_city_nm_txt)
        .PayeClmIntAmt = fnZeroIfNull(rstTemp!paye_clm_int_amt)
        .PayeClmIntRt = fnZeroIfNull(rstTemp!paye_clm_int_rt)
        .PayeClmPdAmt = fnZeroIfNull(rstTemp!paye_clm_pd_amt)
        .PayeDfltOvrdInd = fnYNToBool(rstTemp!paye_dflt_ovrd_ind)
        .PayeDthbPmtAmt = fnZeroIfNull(rstTemp!paye_dthb_pmt_amt)
        .PayeFullNm = fnZLSIfNull(rstTemp!paye_full_nm)
        .PayeId = fnZeroIfNull(rstTemp!paye_id)
        .PayeIntDaysPdNum = fnZeroIfNull(rstTemp!paye_int_days_pd_num)
        .PayePmtDt = rstTemp!paye_pmt_dt
        .PayeSsnTinNum = fnZLSIfNull(rstTemp!paye_ssn_tin_num)
        .PayeSsnTinTypCd = fnZLSIfNull(rstTemp!paye_ssn_tin_typ_cd)
        .PayeStCd = fnZLSIfNull(rstTemp!paye_St_cd)
        .PayeWthldAmt = rstTemp!paye_wthld_amt
        .PayeWthldRt = rstTemp!paye_wthld_rt
        .PayeZip4Cd = fnZLSIfNull(rstTemp!paye_zip4_cd)
        .PayeZipCd = fnZLSIfNull(rstTemp!paye_zip_cd)
        .Paye1099INTInd = fnYNToBool(rstTemp!paye_1099int_ind)   '' BZ4999 October 2013 Non US payee - SXS
    End With

    ' Save original Last Updated info, to be used during UpdateRecord( ) and DeleteRecord( )
    ' to determine if another user updated the record since it was retrieved.
    m_dteLstUpdtDtm_Original = fnZLSIfNull(rstTemp!lst_updt_dtm)
    m_strLstUpdtUserId_Original = fnZLSIfNull(rstTemp!lst_updt_user_id)
    
    If bSynchLookupRST Then
        With m_rstLookup
            m_adwADO.MoveFirst m_rstLookup
            .Find "paye_id = " & PayeId
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
        GetSingleRecord m_rstLookup.Fields(mcstrPayeId).value
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
        GetSingleRecord m_rstLookup.Fields(mcstrPayeId).value
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
                GetLookupData ClmId
                GetRelativeRecord PayeFullNm, epdNextRecord
            Else
                ' Get the requested record and reposition the Lookup recordset to that record
                '!CUSTOMIZE! The constant referenced below should refer to the key field.
                GetSingleRecord m_rstLookup.Fields(mcstrPayeId).value
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
                GetLookupData ClmId
                GetRelativeRecord PayeFullNm, epdPreviousRecord
            Else
                ' Get the requested record and reposition the Lookup recordset to that record
                '!CUSTOMIZE! The constant referenced below should refer to the key field.
                GetSingleRecord m_rstLookup.Fields(mcstrPayeId).value
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
    '             b. (for payee_t table only, I think) row(s) existing in the same table
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
    Const cstrSproc            As String = "dbo.proc_payee_verify_dependents"  ' Stored procedure to execute
    Dim adwTemp                As cadwADOWrapper
    Dim rstTemp                As ADODB.Recordset
    Dim prmReturnValue         As ADODB.Parameter
    Dim prmPayeId              As ADODB.Parameter
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
        Set prmPayeId = .CreateParameter(Name:="@paye_id", _
                                         Type:=adInteger, _
                                         Direction:=adParamInput, _
                                         value:=lngKey1)
        .Parameters.Append prmPayeId

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
    fnFreeObject prmPayeId
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
            HaveDependents = False
            strDependentTable = "Unknown"   ' This is actually ignored by the caller
            ' Remove any trace that this error occurred since we're not going to report it as an error.
            Err.Clear
            Resume PROC_EXIT
        Case gcRES_NERR_DEPENDENT_RECS_EXIST
            ' 4029 = This @@1 is associated with one or more records on the @@2 table and cannot be deleted until those records themselves are deleted.
            ' NOTE: Currently this table has no dependencies and hence the sproc doesn't ever return a 4029.
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
Public Sub InitPayee(ByVal lngClmIdIn As Long)
    ' Comments:   Completes the initialization of the Payee object that
    '             was started in the Class_Initialize procedure. This
    '             procedure should be called immediately after
    '             instantiating an object of this class to
    '             load lookup data associated with the
    '             specified Claim ID.
    '                  Set mtWrapper = New Payee
    '                  InitPayee
    '             The Class_Initialize() method cannot do the
    '             initialization itself since it requires a parameter.
    ' Parameters: lngClmId = the Claim ID with which the desired payees are associated
    ' Returns:    N/A
    ' Called by :
    Const cstrCurrentProc As String = "InitPayee"

    On Error GoTo PROC_ERR

    ' Store the Claim ID!!!
    ClmId = lngClmIdIn
    
    ' Refresh lookup RST and set LookupRecordCount / CurrentLookupRecNbr properties
    GetLookupData lngClmIdIn
    
    ' Get all columns for the 1st record in the Lookup RST and load to member vars.
    ' If there are no records (m_rstLookup is Nothing), then initialize the
    ' properties that correspond to table columns. (Caller must take action if
    ' m_rstLookup Is Nothing!!!)
    If m_rstLookup.RecordCount <> 0 Then
        '!CUSTOMIZE! The constant referenced below should refer to the key field.
        GetSingleRecord m_rstLookup.Fields(mcstrPayeId).value
    Else
        fnClearPropertyValues
    End If
        
    ' Obtain meta data about each table column from the DBMS and load it to the
    ' properties that correspond to those table columns
    fnLoadColMetaData
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

    Const cstrCurrentProc           As String = "UpdateRecord"
    Const cstrSproc                 As String = "dbo.proc_payee_update"  ' Stored procedure to execute
    Dim rstTemp                     As ADODB.Recordset
    Dim bSuccessful                 As Boolean
    Dim prmReturnValue              As ADODB.Parameter
    Dim prmInvalid_Key              As ADODB.Parameter
    Dim prmPayeId                   As ADODB.Parameter
    Dim prmCalcStCd                 As ADODB.Parameter
    Dim prmClmId                    As ADODB.Parameter
    Dim prmPayeAddrLn1Txt           As ADODB.Parameter
    Dim prmPayeAddrLn2Txt           As ADODB.Parameter
    Dim prmPayeCareOfTxt            As ADODB.Parameter
    Dim prmPayeCityNmTxt            As ADODB.Parameter
    Dim prmPayeClmIntAmt            As ADODB.Parameter
    Dim prmPayeClmIntRt             As ADODB.Parameter
    Dim prmPayeClmPdAmt             As ADODB.Parameter
    Dim prmPayeDfltOvrdInd          As ADODB.Parameter
    Dim prmPayeDthbPmtAmt           As ADODB.Parameter
    Dim prmPayeFullNm               As ADODB.Parameter
    Dim prmPayeIntDaysPdNum         As ADODB.Parameter
    Dim prmPayePmtDt                As ADODB.Parameter
    Dim prmPayeSsnTinNum            As ADODB.Parameter
    Dim prmPayeSsnTinTypCd          As ADODB.Parameter
    Dim prmPayeStCd                 As ADODB.Parameter
    Dim prmPayeWthldAmt             As ADODB.Parameter
    Dim prmPayeWthldRt              As ADODB.Parameter
    Dim prmPayeZip4Cd               As ADODB.Parameter
    Dim prmPayeZipCd                As ADODB.Parameter
    Dim prmPaye1099INTInd          As ADODB.Parameter  '' BZ4999 October 2013 Non US payee - SXS
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
        ' Define the PAYE_ID parameter
        Set prmPayeId = .CreateParameter(Name:="@paye_id", _
                                        Type:=adInteger, _
                                        Direction:=adParamInput, _
                                        value:=PayeId)
        .Parameters.Append prmPayeId

        ' ---Parameter #3---
        ' Define the CALC_ST_CD parameter
        Set prmCalcStCd = .CreateParameter(Name:="@calc_st_cd", _
                                         Type:=adChar, _
                                         Direction:=adParamInput, _
                                         Size:=2, _
                                         value:=fnNullIfZLS(varIn:=CalcStCd, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmCalcStCd

        ' ---Parameter #4---
        ' Define the CLM_ID parameter
        Set prmClmId = .CreateParameter(Name:="@clm_id", _
                                        Type:=adInteger, _
                                        Direction:=adParamInput, _
                                        value:=ClmId)
        .Parameters.Append prmClmId

        ' ---Parameter #5---
        ' Define the PAYE_ADR_LN1_TXT parameter
        Set prmPayeAddrLn1Txt = .CreateParameter(Name:="@paye_addr_ln1_txt", _
                                         Type:=adVarChar, _
                                         Direction:=adParamInput, _
                                         Size:=40, _
                                         value:=fnNullIfZLS(varIn:=PayeAddrLn1Txt, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmPayeAddrLn1Txt

        ' ---Parameter #6---
        ' Define the PAYE_ADR_LN2_TXT parameter
        Set prmPayeAddrLn2Txt = .CreateParameter(Name:="@paye_addr_ln2_txt", _
                                         Type:=adVarChar, _
                                         Direction:=adParamInput, _
                                         Size:=40, _
                                         value:=fnNullIfZLS(varIn:=PayeAddrLn2Txt, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmPayeAddrLn2Txt

        ' ---Parameter #7---
        ' Define the PAYE_CARE_OF_TXT parameter
        Set prmPayeCareOfTxt = .CreateParameter(Name:="@paye_care_of_txt", _
                                         Type:=adVarChar, _
                                         Direction:=adParamInput, _
                                         Size:=40, _
                                         value:=fnNullIfZLS(varIn:=PayeCareOfTxt, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmPayeCareOfTxt

        ' ---Parameter #8---
        ' Define the PAYE_CITY_NM_TXT parameter
        Set prmPayeCityNmTxt = .CreateParameter(Name:="@paye_city_nm_txt", _
                                         Type:=adVarChar, _
                                         Direction:=adParamInput, _
                                         Size:=25, _
                                         value:=fnNullIfZLS(varIn:=PayeCityNmTxt, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmPayeCityNmTxt

        ' ---Parameter #9---
        ' Define the PAYE_CLM_INT_AMT parameter
        Set prmPayeClmIntAmt = .CreateParameter(Name:="@paye_clm_int_amt", _
                                          Type:=adNumeric, _
                                          Direction:=adParamInput, _
                                          value:=PayeClmIntAmt)
        With m_dblPayeClmIntAmt
            prmPayeClmIntAmt.Precision = .Precision
            prmPayeClmIntAmt.NumericScale = .NumericScale
        End With
        .Parameters.Append prmPayeClmIntAmt

        ' ---Parameter #10---
        ' Define the PAYE_CLM_INT_RT parameter
        Set prmPayeClmIntRt = .CreateParameter(Name:="@paye_clm_int_rt", _
                                          Type:=adNumeric, _
                                          Direction:=adParamInput, _
                                          value:=PayeClmIntRt)
        With m_dblPayeClmIntRt
            prmPayeClmIntRt.Precision = .Precision
            prmPayeClmIntRt.NumericScale = .NumericScale
        End With
        .Parameters.Append prmPayeClmIntRt

        ' ---Parameter #11---
        ' Define the PAYE_CLM_PD_AMT parameter
        Set prmPayeClmPdAmt = .CreateParameter(Name:="@paye_clm_pd_amt", _
                                          Type:=adNumeric, _
                                          Direction:=adParamInput, _
                                          value:=PayeClmPdAmt)
        With m_dblPayeClmPdAmt
            prmPayeClmPdAmt.Precision = .Precision
            prmPayeClmPdAmt.NumericScale = .NumericScale
        End With
        .Parameters.Append prmPayeClmPdAmt

        ' ---Parameter #12---
        ' Define the PAYE_DFLT_OVRD_IND parameter
        Set prmPayeDfltOvrdInd = .CreateParameter(Name:="@paye_dflt_ovrd_ind", _
                                         Type:=adChar, _
                                         Direction:=adParamInput, _
                                         Size:=1, _
                                         value:=fnBoolToYN(bIn:=PayeDfltOvrdInd))
        .Parameters.Append prmPayeDfltOvrdInd

        ' ---Parameter #13---
        ' Define the PAYE_DTHB_PMT_AMT parameter
        Set prmPayeDthbPmtAmt = .CreateParameter(Name:="@paye_dthb_pmt_amt", _
                                          Type:=adNumeric, _
                                          Direction:=adParamInput, _
                                          value:=PayeDthbPmtAmt)
        With m_dblPayeDthbPmtAmt
            prmPayeDthbPmtAmt.Precision = .Precision
            prmPayeDthbPmtAmt.NumericScale = .NumericScale
        End With
        .Parameters.Append prmPayeDthbPmtAmt

        ' ---Parameter #14---
        ' Define the PAYE_FULL_NM parameter
        Set prmPayeFullNm = .CreateParameter(Name:="@paye_full_nm", _
                                         Type:=adVarChar, _
                                         Direction:=adParamInput, _
                                         Size:=40, _
                                         value:=fnNullIfZLS(varIn:=PayeFullNm, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmPayeFullNm

        ' ---Parameter #15---
        ' Define the PAYE_INT_DAYS_PD_NUM parameter
        Set prmPayeIntDaysPdNum = .CreateParameter(Name:="@paye_int_days_pd_num", _
                                        Type:=adInteger, _
                                        Direction:=adParamInput, _
                                        value:=PayeIntDaysPdNum)
        .Parameters.Append prmPayeIntDaysPdNum

        ' ---Parameter #16---
        ' Define the PAYE_PMT_DT parameter
        Set prmPayePmtDt = .CreateParameter(Name:="@paye_pmt_dt", _
                                         Type:=adDBTimeStamp, _
                                         Direction:=adParamInput, _
                                         Size:=16, _
                                         value:=PayePmtDt)
        .Parameters.Append prmPayePmtDt
        
        ' ---Parameter #17---
        ' Define the PAYE_SSN_TIN_NUM parameter
        Set prmPayeSsnTinNum = .CreateParameter(Name:="@paye_ssn_tin_num", _
                                         Type:=adChar, _
                                         Direction:=adParamInput, _
                                         Size:=9, _
                                         value:=fnNullIfZLS(varIn:=PayeSsnTinNum, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmPayeSsnTinNum
        
        ' ---Parameter #18---
        ' Define the PAYE_SSN_TIN_TYP_CD parameter
        Set prmPayeSsnTinTypCd = .CreateParameter(Name:="@paye_ssn_tin_typ_cd", _
                                         Type:=adChar, _
                                         Direction:=adParamInput, _
                                         Size:=1, _
                                         value:=fnNullIfZLS(varIn:=PayeSsnTinTypCd, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmPayeSsnTinTypCd
        
        ' ---Parameter #19---
        ' Define the PAYE_ST_CD parameter
        Set prmPayeStCd = .CreateParameter(Name:="@paye_st_cd", _
                                         Type:=adChar, _
                                         Direction:=adParamInput, _
                                         Size:=2, _
                                         value:=fnNullIfZLS(varIn:=PayeStCd, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmPayeStCd
        
        ' ---Parameter #20---
        ' Define the PAY_WTHLD_AMT parameter
        Set prmPayeWthldAmt = .CreateParameter(Name:="@paye_wthld_amt", _
                                          Type:=adNumeric, _
                                          Direction:=adParamInput, _
                                          value:=PayeWthldAmt)
        With m_dblPayeWthldAmt
            prmPayeWthldAmt.Precision = .Precision
            prmPayeWthldAmt.NumericScale = .NumericScale
        End With
        .Parameters.Append prmPayeWthldAmt

        ' ---Parameter #21---
        ' Define the PAY_WTHLD_RT parameter
        Set prmPayeWthldRt = .CreateParameter(Name:="@paye_wthld_rt", _
                                          Type:=adNumeric, _
                                          Direction:=adParamInput, _
                                          value:=PayeWthldRt)
        With m_dblPayeWthldRt
            prmPayeWthldRt.Precision = .Precision
            prmPayeWthldRt.NumericScale = .NumericScale
        End With
        .Parameters.Append prmPayeWthldRt

       
        ' ---Parameter #22---
        ' Define the PAYE_ZIP4_cd parameter
        Set prmPayeZip4Cd = .CreateParameter(Name:="@paye_zip4_cd", _
                                         Type:=adChar, _
                                         Direction:=adParamInput, _
                                         Size:=4, _
                                         value:=fnNullIfZLS(varIn:=PayeZip4Cd, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmPayeZip4Cd
        
        ' ---Parameter #23---
        ' Define the PAYE_ZIP_CD parameter
        Set prmPayeZipCd = .CreateParameter(Name:="@paye_zip_cd", _
                                         Type:=adChar, _
                                         Direction:=adParamInput, _
                                         Size:=5, _
                                         value:=fnNullIfZLS(varIn:=PayeZipCd, bHandleEmbeddedQuotes:=True))
        .Parameters.Append prmPayeZipCd
  
        ' ---Parameter #24---  '' BZ4999 October 2013 Non US payee - SXS
        ' Define the Paye1099INTInd parameter
        Set prmPaye1099INTInd = .CreateParameter(Name:="@Paye_1099INT_Ind", _
                                         Type:=adChar, _
                                         Direction:=adParamInput, _
                                         Size:=1, _
                                         value:=fnBoolToYN(bIn:=Paye1099INTInd))
        .Parameters.Append prmPaye1099INTInd

        ' ---Parameter #25---
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
    bSuccessful = GetRelativeRecord(PayeFullNm, epdSameRecord)

    UpdateRecord = True
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
    
    ' Clean-up statements go here
    fnFreeRecordset rstTemp
    fnFreeObject prmReturnValue
    fnFreeObject prmInvalid_Key
    fnFreeObject prmPayeId
    fnFreeObject prmCalcStCd
    fnFreeObject prmClmId
    fnFreeObject prmPayeAddrLn1Txt
    fnFreeObject prmPayeAddrLn2Txt
    fnFreeObject prmPayeCareOfTxt
    fnFreeObject prmPayeCityNmTxt
    fnFreeObject prmPayeClmIntAmt
    fnFreeObject prmPayeClmIntRt
    fnFreeObject prmPayeClmPdAmt
    fnFreeObject prmPayeDfltOvrdInd
    fnFreeObject prmPaye1099INTInd   '' BZ4999 October 2013 Non US payee - SXS
    fnFreeObject prmPayeDthbPmtAmt
    fnFreeObject prmPayeFullNm
    fnFreeObject prmPayeIntDaysPdNum
    fnFreeObject prmPayePmtDt
    fnFreeObject prmPayeSsnTinNum
    fnFreeObject prmPayeSsnTinTypCd
    fnFreeObject prmPayeStCd
    fnFreeObject prmPayeWthldAmt
    fnFreeObject prmPayeWthldRt
    fnFreeObject prmPayeZip4Cd
    fnFreeObject prmPayeZipCd
    
    
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    ' Can get -2147217887 "Invalid character value for cast specification" if too
    ' many characters are defined for a numeric field, e.g., 1111111.22222 to store to
    ' a field defined as 99.99999.

    Select Case prmReturnValue
        Case gcRES_NERR_REC_NOT_FOUND
            ' 4027 = The specified record was not found in the database (@@1).
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_REC_NOT_FOUND, _
                                       mcstrName & cstrCurrentProc, _
                                       "Payee ID " & RTrim$(PayeId)
            Resume PROC_EXIT
        Case gcRES_NERR_ERR_WHILE_TRYING_TO
            ' 4028 = An error occurred while attempting to @@1 this record.
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_ERR_WHILE_TRYING_TO, _
                                       mcstrName & cstrCurrentProc, _
                                       "update"
            Resume PROC_EXIT
        Case gcRES_NERR_KEY_NOT_FOUND
            ' 4032 = The @@1 specified (@@2) is no longer defined in the @@3 table. Please choose a different value.
            If UCase$(prmInvalid_Key) = "CLM_ID" Then
                gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_KEY_NOT_FOUND, _
                                           mcstrName & cstrCurrentProc, _
                                           "Claim ID", RTrim$(ClmId), "CLAIM_T"
                Resume PROC_EXIT
            ElseIf UCase$(prmInvalid_Key) = "CALC_ST_CD" Then
                gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_KEY_NOT_FOUND, _
                                           mcstrName & cstrCurrentProc, _
                                           "Calc State", RTrim$(CalcStCd), "STATE_T"
                Resume PROC_EXIT
            Else
                ' PAYE_ST_CD
                gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_KEY_NOT_FOUND, _
                                           mcstrName & cstrCurrentProc, _
                                           "Payee State", RTrim$(PayeStCd), "STATE_T"
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
                                       "Payee Name " & RTrim$(PayeFullNm) & "/Claim Number " & GetClmNumFromClmID(ClmId)
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
    '                        Nullable numerics => Null
    '                        Booleans => False
    '                        Dates => Now
    
    Const cstrCurrentProc          As String = "fnClearPropertyValues"
    Const clngZero                 As Long = 0

    On Error GoTo PROC_ERR

    With Me
        .CalcStCd = vbNullString
        .LstUpdtDtm = Now
        .LstUpdtUserId = vbNullString
        .PayeAddrLn1Txt = vbNullString
        .PayeAddrLn2Txt = vbNullString
        .PayeCareOfTxt = vbNullString
        .PayeCityNmTxt = vbNullString
        .PayeClmIntAmt = clngZero
        .PayeClmIntRt = clngZero
        .PayeClmPdAmt = clngZero
        .PayeDfltOvrdInd = False
        .Paye1099INTInd = True  '' BZ4999 October 2013 Non US payee - SXS
        .PayeDthbPmtAmt = clngZero
        .PayeFullNm = vbNullString
        .PayeId = clngZero
        .PayeIntDaysPdNum = clngZero
        .PayePmtDt = Now
        .PayeSsnTinNum = vbNullString
        .PayeSsnTinTypCd = vbNullString
        .PayeStCd = vbNullString
        .PayeWthldAmt = clngZero
        .PayeWthldRt = clngZero
        .PayeZip4Cd = vbNullString
        .PayeZipCd = vbNullString
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
    ' The following data types are used within the Claims Interest tables:
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
        If InStr(1, strDefaultValueToParse, "DEF_LST_UPD_USER") > 0 Then
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

        Select Case .DataType
            Case dbDecimal
                ' Save original values (will be used to code sproc parameters).
                .NumericScale = CByte(prstIn("SCALE").value)
                .Precision = CByte(prstIn("PRECISION").value)
                ' Save interpreted equivalents. These may be overriden in fnLoadColMetaData( ).
                .DecimalPositions = CInt(prstIn("SCALE").value)
                .DollarPositions = CInt(prstIn("PRECISION").value) - .DecimalPositions
                .MaxCharacters = 0
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
                ' Establish default values that will usually be in effect if a
                ' TextBox control is used
                .MaxCharacters = CInt(prstIn("LENGTH").value)
                .AllowableCharacters = _
                    "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789@!#$%^&*()-+_=~:;.,<>\|/?' "
                .Format = String(.MaxCharacters, cstrAnyCharChar)
                .Mask = vbNullString
                .ShouldForceToUppercase = False
                
                If Right$(.ColName, 8) = "SSN_NUM" Then
                    ' Social Security Numbers should be displayed via a MaskEdBox control.
                    .Format = vbNullString
                    .Mask = "###-##-####"
                    .AllowableCharacters = "0123456789"
                End If
                If Right$(.ColName, 6) = "ZIP_CD" Then
                    ' Zip Codes require 5 numeric positions if input.
                    ' They should be displayed via a TextBox control.
                    .Mask = vbNullString
                    .MaxCharacters = 5
                    .Format = String(.MaxCharacters, cstrNumericChar)
                    .AllowableCharacters = "0123456789"
                End If
                If Right$(.ColName, 7) = "ZIP4_CD" Then
                    ' Zip Codes require 4 numeric positions if input.
                    ' They should be displayed via a TextBox control.
                    .Mask = vbNullString
                    .MaxCharacters = 4
                    .Format = String(.MaxCharacters, cstrNumericChar)
                    .AllowableCharacters = "0123456789"
                End If

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
        Case "CALCSTCD"
            fnGetProperty = m_strCalcStCd
        Case "CLMID"
            fnGetProperty = m_lngClmId
        Case "LSTUPDTDTM"
            fnGetProperty = m_dteLstUpdtDtm
        Case "LSTUPDTUSERID"
            fnGetProperty = m_strLstUpdtUserId
        Case "PAYEADDRLN1TXT"
            fnGetProperty = m_strPayeAddrLn1Txt
        Case "PAYEADDRLN2TXT"
            fnGetProperty = m_strPayeAddrLn2Txt
        Case "PAYECAREOFTXT"
            fnGetProperty = m_strPayeCareOfTxt
        Case "PAYECITYNMTXT"
            fnGetProperty = m_strPayeCityNmTxt
        Case "PAYECLMINTAMT"
            fnGetProperty = m_dblPayeClmIntAmt
        Case "PAYECLMINTRT"
            fnGetProperty = m_dblPayeClmIntRt
        Case "PAYECLMPDAMT"
            fnGetProperty = m_dblPayeClmPdAmt
        Case "PAYEDFLTOVRDIND"
            fnGetProperty = m_bPayeDfltOvrdInd
        Case "PAYE1099IND"   '' BZ4999 October 2013 Non US payee - SXS
            fnGetProperty = m_bstrPaye1099INTInd
        Case "PAYEDTHBPMTAMT"
            fnGetProperty = m_dblPayeDthbPmtAmt
        Case "PAYEFULLNM"
            fnGetProperty = m_strPayeFullNm
        Case "PAYEID"
            fnGetProperty = m_lngPayeId
        Case "PAYEINTDAYSPDNUM"
            fnGetProperty = m_intPayeIntDaysPdNum
        Case "PAYEPMTDT"
            fnGetProperty = m_dtePayePmtDt
        Case "PAYESSNTINNUM"
            fnGetProperty = m_strPayeSsnTinNum
        Case "PAYESSNTINTYPCD"
            fnGetProperty = m_strPayeSsnTinTypCd
        Case "PAYESTCD"
            fnGetProperty = m_strPayeStCd
        Case "PAYEWTHLDAMT"
            fnGetProperty = m_dblPayeWthldAmt
        Case "PAYEWTHLDRT"
            fnGetProperty = m_dblPayeWthldRt
        Case "PAYEZIP4CD"
            fnGetProperty = m_strPayeZip4Cd
        Case "PAYEZIPCD"
            fnGetProperty = m_strPayeZipCd
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
    m_adwADO.GetMetaData_Columns "payee_t", rstMetaData

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
            Select Case UCase$(rstMetaData("COLUMN_NAME").value)
                Case mcstrCalcStCd
                    fnGetColMetaData m_strCalcStCd, rstMetaData
                Case mcstrClmId
                    fnGetColMetaData m_lngClmId, rstMetaData
                Case mcstrLstUpdtDtm
                    fnGetColMetaData m_dteLstUpdtDtm, rstMetaData
                Case mcstrLstUpdtUserId
                    fnGetColMetaData m_strLstUpdtUserId, rstMetaData
                Case mcstrPayeAddrLn1Txt
                    fnGetColMetaData m_strPayeAddrLn1Txt, rstMetaData
                    m_strPayeAddrLn1Txt.ShouldForceToUppercase = True
                Case mcstrPayeAddrLn2Txt
                    fnGetColMetaData m_strPayeAddrLn2Txt, rstMetaData
                    m_strPayeAddrLn2Txt.ShouldForceToUppercase = True
                Case mcstrPayeCareOfTxt
                    fnGetColMetaData m_strPayeCareOfTxt, rstMetaData
                    m_strPayeCareOfTxt.ShouldForceToUppercase = True
                Case mcstrPayeCityNmTxt
                    fnGetColMetaData m_strPayeCityNmTxt, rstMetaData
                    m_strPayeCityNmTxt.ShouldForceToUppercase = True
                Case mcstrPayeClmIntAmt
                    fnGetColMetaData m_dblPayeClmIntAmt, rstMetaData
                Case mcstrPayeClmIntRt
                    fnGetColMetaData m_dblPayeClmIntRt, rstMetaData
                Case mcstrPayeClmPdAmt
                    fnGetColMetaData m_dblPayeClmPdAmt, rstMetaData
                Case mcstrPayeDfltOvrdInd
                    fnGetColMetaData m_bPayeDfltOvrdInd, rstMetaData
                Case mcstrPayeDthbPmtAmt
                    fnGetColMetaData m_dblPayeDthbPmtAmt, rstMetaData
                    ' GUI enforces this default value of zero, not DBMS
                    m_dblPayeDthbPmtAmt.DefaultValue = 0
                Case mcstrPayeFullNm
                    fnGetColMetaData m_strPayeFullNm, rstMetaData
                    m_strPayeFullNm.ShouldForceToUppercase = True
                Case mcstrPayeId
                    fnGetColMetaData m_lngPayeId, rstMetaData
                Case mcstrPayeIntDaysPdNum
                    fnGetColMetaData m_intPayeIntDaysPdNum, rstMetaData
                Case mcstrPayePmtDt
                    fnGetColMetaData m_dtePayePmtDt, rstMetaData
                Case mcstrPayeSsnTinNum
                    fnGetColMetaData m_strPayeSsnTinNum, rstMetaData
                    m_strPayeSsnTinNum.AllowableCharacters = "0123456789"
                Case mcstrPayeSsnTinTypCd
                    fnGetColMetaData m_strPayeSsnTinTypCd, rstMetaData
                    m_strPayeSsnTinTypCd.DefaultValue = gcstrBlankEntry
                Case mcstrPayeStCd
                    fnGetColMetaData m_strPayeStCd, rstMetaData
                Case mcstrPayeWthldAmt
                    fnGetColMetaData m_dblPayeWthldAmt, rstMetaData
                Case mcstrPayeWthldRt
                    fnGetColMetaData m_dblPayeWthldRt, rstMetaData
                Case mcstrPayeZip4Cd
                    fnGetColMetaData m_strPayeZip4Cd, rstMetaData
                Case mcstrPayeZipCd
                    fnGetColMetaData m_strPayeZipCd, rstMetaData
                Case mcstrPaye1099INTInd
                    fnGetColMetaData m_bstrPaye1099INTInd, rstMetaData
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
    m_adwADO.GetMetaData_PrimaryKeys "payee_t", rstMetaData

    With rstMetaData
        Do Until .EOF
            ' The SELECT CASE should list all table columns
            ' (though you could skip the LST_UPDT_xxx columns if you change the Case Else,
            '  since these would never be a key)
            Select Case UCase$(rstMetaData("COLUMN_NAME").value)
                Case mcstrCalcStCd
                    m_strCalcStCd.IsKey = True
                Case mcstrClmId
                    m_lngClmId.IsKey = True
                Case mcstrLstUpdtDtm
                    m_dteLstUpdtDtm.IsKey = True
                Case mcstrLstUpdtUserId
                    m_strLstUpdtUserId.IsKey = True
                Case mcstrPayeAddrLn1Txt
                    m_strPayeAddrLn1Txt.IsKey = True
                Case mcstrPayeAddrLn2Txt
                    m_strPayeAddrLn2Txt.IsKey = True
                Case mcstrPayeCareOfTxt
                    m_strPayeCareOfTxt.IsKey = True
                Case mcstrPayeCityNmTxt
                    m_strPayeCityNmTxt.IsKey = True
                Case mcstrPayeClmIntAmt
                    m_dblPayeClmIntAmt.IsKey = True
                Case mcstrPayeClmIntRt
                    m_dblPayeClmIntRt.IsKey = True
                Case mcstrPayeClmPdAmt
                    m_dblPayeClmPdAmt.IsKey = True
                Case mcstrPayeDfltOvrdInd
                    m_bPayeDfltOvrdInd.IsKey = True
                Case mcstrPayeDthbPmtAmt
                    m_dblPayeDthbPmtAmt.IsKey = True
                Case mcstrPayeFullNm
                    m_strPayeFullNm.IsKey = True
                Case mcstrPayeId
                    m_lngPayeId.IsKey = True
                Case mcstrPayeIntDaysPdNum
                    m_intPayeIntDaysPdNum.IsKey = True
                Case mcstrPayePmtDt
                    m_dtePayePmtDt.IsKey = True
                Case mcstrPayeSsnTinNum
                    m_strPayeSsnTinNum.IsKey = True
                Case mcstrPayeSsnTinTypCd
                    m_strPayeSsnTinTypCd.IsKey = True
                Case mcstrPayeStCd
                    m_strPayeStCd.IsKey = True
                Case mcstrPayeWthldAmt
                    m_dblPayeWthldAmt.IsKey = True
                Case mcstrPayeWthldRt
                    m_dblPayeWthldRt.IsKey = True
                Case mcstrPayeZip4Cd
                    m_strPayeZip4Cd.IsKey = True
                Case mcstrPayeZipCd
                    m_strPayeZipCd.IsKey = True
                Case mcstrPaye1099INTInd
                    m_bstrPaye1099INTInd.IsKey = True
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
    Const cstrSproc                As String = "dbo.proc_payee_select"  ' Stored procedure to execute
    Dim rstTemp                    As ADODB.Recordset
    Dim prmReturnValue             As ADODB.Parameter
    Dim prmPayeId                  As ADODB.Parameter

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
        ' Define the paye_id parameter
        Set prmPayeId = .CreateParameter(Name:="@paye_id", _
                                         Type:=adInteger, _
                                         Direction:=adParamInput, _
                                         value:=fnNullIfZero(lngKey1))
        .Parameters.Append prmPayeId

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
    fnFreeObject prmPayeId
    
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
                                       "Payee ID " & RTrim$(lngKey1)
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
