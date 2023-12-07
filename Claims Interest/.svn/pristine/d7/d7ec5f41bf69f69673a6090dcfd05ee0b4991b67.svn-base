VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ctxsTaxSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************
' Module     : ctxsTaxSettings
' Description: This class is a wrapper around the Settings table in the
'              application database. By instantiating an object of type
'              ctxsTaxSettings and using its Property Get procedures, the
'              caller can get at each column in the Settings table (i.e. to
'              limit data-entry to max length per the Settings table),
'              without (a) using many global variables and (b) knowing that
'              behind the scenes the data is coming from a table in the DB.
' Procedures : Class_Initialize
'              Property Get for each member variable (i.e. column in the Settings table) - public
'              Property Let for each member variable (i.e. column in the Settings table) - private
' Modified   :
'  01/2002 BAW  Added support for new column in the Settings table: CurrentInterestRate.
'              Also removed Property Let procedures since the app should not be
'              updating these fields.
'
' --------------------------------------------------
Option Explicit
Option Compare Binary

Private Const mcstrName             As String = "ctxsTaxSettings."

'------------------------------------------
'            MEMBER VARIABLES
'------------------------------------------
' member variable for Address1Length
 Private mintAddress1Length As Integer
' member variable for Address2Length
 Private mintAddress2Length As Integer
' member variable for AddressLength
 Private mintAddressLength As Integer
' member variable for AddrTypeLength
 Private mintAddrTypeLength As Integer
' member variable for ApplicationLength
 Private mintApplicationLength As Integer
' member variable for CareOfLength
 Private mintCareOfLength As Integer
' member variable for CityLength
 Private mintCityLength As Integer
' member variable for ClaimNumber_GroupLength
 Private mintClaimNumber_GroupLength  As Integer
 ' member variable for ClaimNumber_IndivLength
Private mintClaimNumber_IndivLength  As Integer
' member variable for CountryLength
 Private mintCountryLength As Integer
 ' member variable for CurrentInterestRate
 Private mdblCurrentInterestRate As Double
' member variable for FedIDLength
 Private mintFedIDLength As Integer
' member variable for Filler1Length
 Private mintFiller1Length As Integer
' member variable for Filler2Length
Private mintFiller2Length  As Integer
' member variable for Filler3Length
Private mintFiller3Length  As Integer
' member variable for FormLength
Private mintFormLength  As Integer
' member variable for InterestLength
Private mintInterestLength  As Integer
' member variable for NameLength
Private mintNameLength  As Integer
' member variable for ProductLength
Private mintProductLength  As Integer
' member variable for RecordLength
Private mintRecordLength  As Integer
' member variable for ResLength
Private mintResLength  As Integer
' member variable for SecNameFlagLength
Private mintSecNameFlagLength  As Integer
' member variable for SecondNameLength
Private mintSecondNameLength  As Integer
' member variable for StateLength
Private mintStateLength  As Integer
' member variable for StateResLength
Private mintStateResLength  As Integer
' member variable for StateWthhld1Length
Private mintStateWthhld1Length  As Integer
' member variable for SunCode
Private mstrSunCode  As String
' member variable for Tin2NoticeLength
Private mintTIN2NoticeLength  As Integer
' member variable for TINLength
Private mintTINLength  As Integer
' member variable for TINTypeLength
Private mintTINTypeLength  As Integer
' member variable for ZipLength
Private mintZipLength  As Integer


'------------------------------------------
'           PROPERTY GET
'------------------------------------------
' Get Property Procedure for the Address1Length property
Public Property Get Address1Length() As Integer
    Address1Length = mintAddress1Length
End Property


' Get Property Procedure for the Address2Length property
Public Property Get Address2Length() As Integer
    Address2Length = mintAddress2Length
End Property


' Get Property Procedure for the AddressLength property
Public Property Get AddressLength() As Integer
    AddressLength = mintAddressLength
End Property


' Get Property Procedure for the AddrTypeLength property
Public Property Get AddrTypeLength() As Integer
    AddrTypeLength = mintAddrTypeLength
End Property


' Get Property Procedure for the ApplicationLength property
Public Property Get ApplicationLength() As Integer
    ApplicationLength = mintApplicationLength
End Property


' Get Property Procedure for the CareOfLength property
Public Property Get CareOfLength() As Integer
    CareOfLength = mintCareOfLength
End Property


' Get Property Procedure for the CityLength property
Public Property Get CityLength() As Integer
    CityLength = mintCityLength
End Property


' Get Property Procedure for the ClaimNumber_GroupLength property
Public Property Get ClaimNumber_GroupLength() As Integer
    ClaimNumber_GroupLength = mintClaimNumber_GroupLength
End Property


' Get Property Procedure for the ClaimNumber_IndivLength property
Public Property Get ClaimNumber_IndivLength() As Integer
    ClaimNumber_IndivLength = mintClaimNumber_IndivLength
End Property


' Get Property Procedure for the CurrentInterestRate property
Public Property Get CurrentInterestRate() As Double
    CurrentInterestRate = mdblCurrentInterestRate
End Property


' Get Property Procedure for the CountryLength property
Public Property Get CountryLength() As Integer
    CountryLength = mintCountryLength
End Property


' Get Property Procedure for the FedIDLength property
Public Property Get FedIDLength() As Integer
    FedIDLength = mintFedIDLength
End Property


' Get Property Procedure for the Filler1Length property
Public Property Get Filler1Length() As Integer
    Filler1Length = mintFiller1Length
End Property


' Get Property Procedure for the Filler2Length property
Public Property Get Filler2Length() As Integer
    Filler2Length = mintFiller2Length
End Property


' Get Property Procedure for the Filler3Length property
Public Property Get Filler3Length() As Integer
    Filler3Length = mintFiller3Length
End Property


' Get Property Procedure for the FormLength property
Public Property Get FormLength() As Integer
    FormLength = mintFormLength
End Property


' Get Property Procedure for the InterestLength property
Public Property Get InterestLength() As Integer
    InterestLength = mintInterestLength
End Property


' Get Property Procedure for the NameLength property
Public Property Get NameLength() As Integer
    NameLength = mintNameLength
End Property


' Get Property Procedure for the ProductLength property
Public Property Get ProductLength() As Integer
    ProductLength = mintProductLength
End Property


' Get Property Procedure for the RecordLength property
Public Property Get RecordLength() As Integer
    RecordLength = mintRecordLength
End Property


' Get Property Procedure for the ResLength property
Public Property Get ResLength() As Integer
    ResLength = mintResLength
End Property


' Get Property Procedure for the SecNameFlagLength property
Public Property Get SecNameFlagLength() As Integer
    SecNameFlagLength = mintSecNameFlagLength
End Property


' Get Property Procedure for the SecondNameLength property
Public Property Get SecondNameLength() As Integer
    SecondNameLength = mintSecondNameLength
End Property


' Get Property Procedure for the StateLength property
Public Property Get StateLength() As Integer
    StateLength = mintStateLength
End Property


' Get Property Procedure for the StateResLength property
Public Property Get StateResLength() As Integer
    StateResLength = mintStateResLength
End Property


' Get Property Procedure for the StateWthhld1Length property
Public Property Get StateWthhld1Length() As Integer
    StateWthhld1Length = mintStateWthhld1Length
End Property


' Get Property Procedure for the SunCode property
Public Property Get SunCode() As String
    SunCode = mstrSunCode
End Property


' Get Property Procedure for the Tin2NoticeLength property
Public Property Get Tin2NoticeLength() As Integer
    Tin2NoticeLength = mintTIN2NoticeLength
End Property


' Get Property Procedure for the TINLength property
Public Property Get TINLength() As Integer
    TINLength = mintTINLength
End Property


' Get Property Procedure for the TINTypeLength property
Public Property Get TINTypeLength() As Integer
    TINTypeLength = mintTINTypeLength
End Property


' Get Property Procedure for the ZipLength property
Public Property Get ZipLength() As Integer
    ZipLength = mintZipLength
End Property


' ********************************************
'  Initialize and Terminate
' ********************************************
Private Sub Class_Initialize()
    ' Comments  : Accesses the Settings table in CLAIMS.MDB
    '             and populates member variables with its values.
    ' Called by : fnInitializeAppConnectionObject
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    Const cstrCurrentProc As String = "Class_Initialize"
    On Error GoTo PROC_ERR

    Dim mrstSettings As ADODB.Recordset
    Dim strSQL As String
    
    Set mrstSettings = New ADODB.Recordset
    
    ' This SQL statement returns a single row, which happens to be the
    ' entire contents of the Settings table
    strSQL = "SELECT [CurrentInterestRate], [RecordLength], " & _
        "[Form] as FormLength, [Fed-Id] As FedIDLength, " & _
        "[Application] As ApplicationLength, [Product] as ProductLength, " & _
        "[Res] As ResLength, [Name] as NameLength, " & _
        "[Second-Name] As SecondNameLength, [Address1] as Address1Length, " & _
        "[Address2] As Address2Length, [City] as CityLength, " & _
        "[Filler1] As Filler1Length, [State-Res] as StateResLength, " & _
        "[Filler2] As Filler2Length, [Zip] as ZipLength, " & _
        "[Addr-Type] As AddrTypeLength, [State-Wthhld1] as StateWthhld1Length, " & _
        "[TIN] As TINLength, [Tin-Type] as TINTypeLength, " & _
        "[State] As StateLength, [Suncode], " & _
        "[ClaimNumber_Group] As ClaimNumber_GroupLength, " & _
        "[ClaimNumber_Indiv] As ClaimNumber_IndivLength, [Country] as CountryLength, " & _
        "[Tin2-Notice] As TIN2NoticeLength, [Sec-Name-Flag] as SecNameFlagLength, " & _
        "[Filler3] As Filler3Length, [Interest] as InterestLength, " & _
        "[address] As AddressLength, [careof] as CareOfLength " & _
        "FROM [Settings]"
    
    ' CursorType=adOpenKeyset   - Scrolling fwd/bwd permitted, chgs/del by other users visible
    ' LockType=adLockReadOnly   - Recordset is read-only
    mrstSettings.Open Source:=strSQL, _
                      ActiveConnection:=gconAppActive, _
                      CursorType:=adOpenKeyset, _
                      LockType:=adLockReadOnly, _
                      Options:=adCmdText
            
    If mrstSettings.RecordCount > 0 Then
        With mrstSettings
            mintAddress1Length = !Address1Length
            mintAddress2Length = !Address2Length
            mintAddressLength = !AddressLength
            mintAddrTypeLength = !AddrTypeLength
            mintApplicationLength = !ApplicationLength
            mintCareOfLength = !CareOfLength
            mintCityLength = !CityLength
            mintCountryLength = !CountryLength
            mdblCurrentInterestRate = !CurrentInterestRate
            mintFedIDLength = !FedIDLength
            mintFiller1Length = !Filler1Length
            mintFiller2Length = !Filler2Length
            mintFiller3Length = !Filler3Length
            mintFormLength = !FormLength
            mintInterestLength = !InterestLength
            mintNameLength = !NameLength
            mintClaimNumber_GroupLength = !ClaimNumber_GroupLength
            mintClaimNumber_IndivLength = !ClaimNumber_IndivLength
            mintProductLength = !ProductLength
            mintRecordLength = !RecordLength
            mintResLength = !ResLength
            mintSecNameFlagLength = !SecNameFlagLength
            mintSecondNameLength = !SecondNameLength
            mintStateLength = !StateLength
            mintStateResLength = !StateResLength
            mintStateWthhld1Length = !StateWthhld1Length
            mstrSunCode = !SunCode
            mintTIN2NoticeLength = !Tin2NoticeLength
            mintTINLength = !TINLength
            mintTINTypeLength = !TINTypeLength
            mintZipLength = !ZipLength
        End With
    End If
PROC_EXIT:
    On Error GoTo 0     ' disable error handler

    ' Clean-up statements go here
    fnFreeRecordset mrstSettings

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

