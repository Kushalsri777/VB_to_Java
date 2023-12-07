VERSION 5.00
Begin VB.Form frmTestTableWrapper 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test Table Wrapper"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   465
      Left            =   1725
      TabIndex        =   0
      Top             =   1350
      Width           =   1215
   End
End
Attribute VB_Name = "frmTestTableWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % %
'#If False Then


'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
' PAYEE_T WRAPPER
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\


'--------------------------------------------------------------------------------------------
' Customize for the table wrapper you want to test, by doing:
'   a. An Edit/Replace All on "ctclmClaim"   (the name of the table wrapper class)
'   b. Edit the fnDisplayRecord proc so it reflects each table column
'   c. Look for !CUSTOMIZE! tags for other places where editing is needed.
'--------------------------------------------------------------------------------------------





Option Explicit
Option Compare Binary

'--------------------------------------------------------------------------------------------
' Just need declarations from some maintenance screen
'--------------------------------------------------------------------------------------------
Private mstrScreenName As String

Private Const mclngMinFormWidth As Long = 4800
Private Const mclngMinFormHeight As Long = 3600

Private Const mclngClmIdToUse As Long = 279742      ' Make sure this CLM_ID value exists!

Private m_tsOutput As Scripting.TextStream

Private mctlFirstUpdateableField_Add As Control
Private mctlFirstUpdateableField_Upd As Control


' mtWrapper is an instance of the table wrapper corresponding to the main table maintained by this form.
Private mtWrapper               As ctpyePayee




'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub cmdClose_Click()
    Unload Me
End Sub




'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    Const cstrCurrentProc       As String = "Form_Load"
    Const cstrOutputFile        As String = "C:\WrapperTest.txt"
    Dim rst                     As ADODB.Recordset
    Dim fld                     As ADODB.Field
        Dim dteToday                As Date
    Dim fso                     As Scripting.FileSystemObject
    
    On Error GoTo PROC_ERR

    '--------------------------------------------------------------------------------------------
    ' Do not change
    '--------------------------------------------------------------------------------------------
    mstrScreenName = Me.Caption
    gerhApp.ScreenName = mstrScreenName
    Me.Icon = LoadResPicture(gcRES_ICON_MAINAPP, vbResIcon)
    With Me
        .Width = mclngMinFormWidth
        .Height = mclngMinFormHeight
    End With
    fnCenterFormOnMDI frmMDIMain, Me
    Set mctlFirstUpdateableField_Add = cmdClose
    Set mctlFirstUpdateableField_Upd = cmdClose
    Set fso = New Scripting.FileSystemObject
    Set m_tsOutput = fso.CreateTextFile(cstrOutputFile, True)
    

    '--------------------------------------------------------------------------------------------
    ' 1.  Verify the table wrapper was instantiated successfully.
    '--------------------------------------------------------------------------------------------
    Set mtWrapper = New ctpyePayee
    mtWrapper.InitPayee mclngClmIdToUse      ' -- using 279742 as the CLM_ID
    
    '!CUSTOMIZE! Include a call to fnDisplayMetaData for each property associated with a table column
    With mtWrapper
        m_tsOutput.WriteLine "----------------------------------------------------"
        m_tsOutput.WriteLine " 1. Meta Data"
        m_tsOutput.WriteLine "----------------------------------------------------"
        fnDisplayMetaData "CalcStCd"
        fnDisplayMetaData "ClmId"
        fnDisplayMetaData "LstUpdtDtm"
        fnDisplayMetaData "LstUpdtUserId"
        fnDisplayMetaData "PayeAddrLn1Txt"
        fnDisplayMetaData "PayeAddrLn2Txt"
        fnDisplayMetaData "PayeCareOfTxt"
        fnDisplayMetaData "PayeCityNmTxt"
        fnDisplayMetaData "PayeClmIntAmt"
        fnDisplayMetaData "PayeClmIntRt"
        fnDisplayMetaData "PayeDfltOvrdInd"
        fnDisplayMetaData "PayeDthbPmtAmt"
        fnDisplayMetaData "PayeFullNm"
        fnDisplayMetaData "PayeId"
        fnDisplayMetaData "PayeIntDaysPdNum"
        fnDisplayMetaData "PayePmtDt"
        fnDisplayMetaData "PayeSsnTinNum"
        fnDisplayMetaData "PayeSsnTinTypCd"
        fnDisplayMetaData "PayeWthldAmt"
        fnDisplayMetaData "PayeWthldRt"
        fnDisplayMetaData "PayeZip4Cd"
        fnDisplayMetaData "PayeZipCd"
    End With
    
    '--------------------------------------------------------------------------------------------
    ' 2.  Verify that the Lookup recordset was successfully populated (right columns, and that the
    '     number of rows retrieved is appropriate)
    '--------------------------------------------------------------------------------------------
    m_tsOutput.WriteLine vbCrLf
    m_tsOutput.WriteLine "----------------------------------------------------"
    m_tsOutput.WriteLine " 2. Lookup Recordset Population"
    m_tsOutput.WriteLine "----------------------------------------------------"
    m_tsOutput.WriteLine "The m_rstLookup recordset was populated as follows: "
    m_tsOutput.WriteLine "   Record Count = " & CStr(mtWrapper.LookupRecordCount)
    m_tsOutput.WriteLine "   Current Rec# = " & CStr(mtWrapper.CurrentLookupRecordNumber)
    
    Set rst = mtWrapper.LookupData
     
    For Each fld In rst.Fields
        m_tsOutput.WriteLine "   For the first record in the m_rstLookup recordset:"
        m_tsOutput.WriteLine "     column = " & fld.Name & ", value = [" & CStr(fnZLSIfNull(fld.value)) & "]"
    Next fld
    
     '--------------------------------------------------------------------------------------------
    ' 3. Test navigation methods
    '--------------------------------------------------------------------------------------------
    m_tsOutput.WriteLine vbCrLf
    m_tsOutput.WriteLine "----------------------------------------------------"
    m_tsOutput.WriteLine " 3. Navigation"
    m_tsOutput.WriteLine "----------------------------------------------------"
    With mtWrapper
        If .LookupRecordCount > 0 Then
            mtWrapper.GoToLastRecord
            m_tsOutput.WriteLine "Navigated to last record, rec #" & .CurrentLookupRecordNumber
            mtWrapper.GoToFirstRecord
            m_tsOutput.WriteLine "Navigated to first record, rec #" & .CurrentLookupRecordNumber
        End If
        If .LookupRecordCount > 1 Then
            mtWrapper.GoToNextRecord
            m_tsOutput.WriteLine "Navigated to next record, rec #" & .CurrentLookupRecordNumber
            mtWrapper.GoToPreviousRecord
            m_tsOutput.WriteLine "Navigated to previous record, rec #" & .CurrentLookupRecordNumber
        End If
    End With

    '--------------------------------------------------------------------------------------------
    ' 4. Test ADD capability
    '--------------------------------------------------------------------------------------------
    m_tsOutput.WriteLine vbCrLf
    m_tsOutput.WriteLine "----------------------------------------------------"
    m_tsOutput.WriteLine " 4. Add"
    m_tsOutput.WriteLine "----------------------------------------------------"
    With mtWrapper
        '!CUSTOMIZE!
        ' Set properties, as if the user input them. Be sure to choose values that
        ' should result in a successful Add.
        dteToday = Now
        .ClmId = mclngClmIdToUse
        .CalcStCd = "FL"
        .LstUpdtDtm = Now
        .LstUpdtUserId = "TEST    "
        .PayeAddrLn1Txt = "ADDRESS LINE 1"
        .PayeAddrLn2Txt = "ADDRESS LINE 2"
        .PayeCareOfTxt = "CARE OF INFO"
        .PayeCityNmTxt = "JACKSONVILLE"
        .PayeClmIntAmt = 17.75
        .PayeClmIntRt = 3.5
        .PayeClmPdAmt = 100017.75
        .PayeDfltOvrdInd = False
        .PayeDthbPmtAmt = 10000
        .PayeFullNm = "BETSY TESTCASE"
        '-- identity .PayeId =
        .PayeIntDaysPdNum = 7
        .PayePmtDt = CDate("01/03/2003")
        .PayeSsnTinNum = 123445555
        .PayeSsnTinTypCd = "P"
        .PayeStCd = "FL"
        .PayeWthldAmt = 0
        .PayeWthldRt = 0
        .PayeZip4Cd = "4444"
        .PayeZipCd = "12345"
      ' Call the Add method
        .AddRecord
        ' Display the values of the current just-added record
        m_tsOutput.WriteLine "Just added this record: "
        fnDisplayRecord
    End With
    
    '--------------------------------------------------------------------------------------------
    ' 5. Test UPDATE capability
    '--------------------------------------------------------------------------------------------
    m_tsOutput.WriteLine vbCrLf
    m_tsOutput.WriteLine "----------------------------------------------------"
    m_tsOutput.WriteLine " 5. Update"
    m_tsOutput.WriteLine "----------------------------------------------------"
    With mtWrapper
        '!CUSTOMIZE!
        ' Set properties, to simulate the user editing the record just added. Be sure to choose
        ' values that should result in a successful Update.  Be sure to NOT change key fields!
        .PayeIntDaysPdNum = 12
        .LstUpdtDtm = Now
        .LstUpdtUserId = "K000    "
        ' Call the Update method
        .UpdateRecord
        ' Display the values of the current just-updated record
        m_tsOutput.WriteLine "After updating, this record now looks like this: "
        fnDisplayRecord
    End With
    
    '--------------------------------------------------------------------------------------------
    ' 6. Test DELETE capability
    '--------------------------------------------------------------------------------------------
    m_tsOutput.WriteLine vbCrLf
    m_tsOutput.WriteLine "----------------------------------------------------"
    m_tsOutput.WriteLine " 6. Delete"
    m_tsOutput.WriteLine "----------------------------------------------------"
    With mtWrapper
        ' Using properties already set by the Add and Update tests, call the
        ' Delete method to delete the record.
        .DeleteRecord
        ' Display the values of the record previous to the one just deleted
        m_tsOutput.WriteLine "After deleting, the new current record is now this: "
        fnDisplayRecord
    End With
    
    m_tsOutput.Close
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here
    fnFreeObject m_tsOutput
    fnFreeObject fso

    ' Report the error, since this is an event handler
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.ReportFatalError mstrScreenName
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            ' Save Err object data, if not already saved
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub


'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub fnDisplayMetaData(ByVal strTagIn As String)
    Const cstrCurrentProc       As String = "fnDisplayMetaData"
    On Error GoTo PROC_EXIT
    
    With mtWrapper
        m_tsOutput.WriteLine "The " & strTagIn & " property: "
        m_tsOutput.WriteLine "  AllowableCharacters=" & .AllowableCharacters(strTagIn)
        m_tsOutput.WriteLine "  DefaultValue=[" & CStr(.DefaultValue(strTagIn)) & "]"
        m_tsOutput.WriteLine "  DollarPositions=" & CStr(.DollarPositions(strTagIn))
        m_tsOutput.WriteLine "  DecimalPositions=" & CStr(.DecimalPositions(strTagIn))
        m_tsOutput.WriteLine "  IsNullable=" & .IsNullable(strTagIn)
        m_tsOutput.WriteLine "  Format=" & .Format(strTagIn)
        m_tsOutput.WriteLine "  Mask=" & .Mask(strTagIn)
        m_tsOutput.WriteLine "  IsKey=" & .IsKey(strTagIn)
        m_tsOutput.WriteLine "  ShouldForceToUppercase=" & .ShouldForceToUppercase(strTagIn)
        m_tsOutput.WriteLine "  MaxCharacters=" & CStr(.MaxCharacters(strTagIn))
    End With
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here

    ' Report the error, since this is an event handler
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.ReportFatalError mstrScreenName
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            ' Save Err object data, if not already saved
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub




'!CUSTOMIZE! Edit this procedure so it lists each column of your table.
Private Sub fnDisplayRecord()
    Const cstrCurrentProc       As String = "fnDisplayRecord"
    On Error GoTo PROC_EXIT
    
    With mtWrapper
        m_tsOutput.WriteLine "PayeId = [" & .PayeId & "]"
        m_tsOutput.WriteLine "CalcStCd = [" & .CalcStCd & "]"
        m_tsOutput.WriteLine "LstUpdtDtm = [" & .LstUpdtDtm & "]"
        m_tsOutput.WriteLine "LstUpdtUserId = [" & .LstUpdtUserId & "]"
        m_tsOutput.WriteLine "PayeAddrLn1Txt = [" & .PayeAddrLn1Txt & "]"
        m_tsOutput.WriteLine "PayeAddrLn2Txt = [" & .PayeAddrLn2Txt & "]"
        m_tsOutput.WriteLine "PayeCareOfTxt = [" & .PayeCareOfTxt & "]"
        m_tsOutput.WriteLine "PayeCityNmTxt = [" & .PayeCityNmTxt & "]"
        m_tsOutput.WriteLine "PayeClmIntAmt = [" & .PayeClmIntAmt & "]"
        m_tsOutput.WriteLine "PayeClmIntRt = [" & .PayeClmIntRt & "]"
        m_tsOutput.WriteLine "PayeClmPdAmt = [" & .PayeClmPdAmt & "]"
        m_tsOutput.WriteLine "PayeDfltOvrdInd = [" & .PayeDfltOvrdInd & "]"
        m_tsOutput.WriteLine "PayeDthbPmtAmt = [" & .PayeDthbPmtAmt & "]"
        m_tsOutput.WriteLine "PayeFullNm = [" & .PayeFullNm & "]"
        m_tsOutput.WriteLine '-- identity .PayeId =
        m_tsOutput.WriteLine "PayeIntDaysPdNum = [" & .PayeIntDaysPdNum & "]"
        m_tsOutput.WriteLine "PayePmtDt = [" & .PayePmtDt & "]"
        m_tsOutput.WriteLine "PayeSsnTinNum = [" & .PayeSsnTinNum & "]"
        m_tsOutput.WriteLine "PayeSsnTinTypCd = [" & .PayeSsnTinTypCd & "]"
        m_tsOutput.WriteLine "PayeStCd = [" & .PayeStCd & "]"
        m_tsOutput.WriteLine "PayeWthldAmt = [" & .PayeWthldAmt & "]"
        m_tsOutput.WriteLine "PayeWthldRt = [" & .PayeWthldRt & "]"
        m_tsOutput.WriteLine "PayeZip4Cd = [" & .PayeZip4Cd & "]"
        m_tsOutput.WriteLine "PayeZipCd = [" & .PayeZipCd & "]"
    End With
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here

    ' Report the error, since this is an event handler
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.ReportFatalError mstrScreenName
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            ' Save Err object data, if not already saved
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub


'#End If
' % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % %






' % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % %
#If False Then


'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
' CLAIM_T WRAPPER
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\


'--------------------------------------------------------------------------------------------
' Customize for the table wrapper you want to test, by doing:
'   a. An Edit/Replace All on "ctclmClaim"   (the name of the table wrapper class)
'   b. Edit the fnDisplayRecord proc so it reflects each table column
'   c. Look for !CUSTOMIZE! tags for other places where editing is needed.
'--------------------------------------------------------------------------------------------





Option Explicit
Option Compare Binary

'--------------------------------------------------------------------------------------------
' Just need declarations from some maintenance screen
'--------------------------------------------------------------------------------------------
Private mstrScreenName As String

Private Const mclngMinFormWidth As Long = 4800
Private Const mclngMinFormHeight As Long = 3600

Private m_tsOutput As Scripting.TextStream

Private mctlFirstUpdateableField_Add As Control
Private mctlFirstUpdateableField_Upd As Control
' mbInLookupMode determines whether the user is in the process of doing a search using the Lookup ComboBox
Private mbInLookupMode          As Boolean

' mbInAddMode determines whether the user has begun the process of adding a new record to the table.
' Note that Add mode is independent of Update mode
Private mbInAddMode             As Boolean

' mtWrapper is an instance of the table wrapper corresponding to the main table maintained by this form.
Private mtWrapper               As ctclmClaim

' m_bIsDirty corresponds to the public property called IsDirty.
' All maintenance screens should have this field and that property! When True, it indicates
' that the user has made --but not yet saved-- changes to a record. The MDI form will query
' this property if the user opens the File menu, since the Exit option should be disabled if
' any form has outstanding changes.
' Be sure to use this variable's corresponding Property Let to change its value.
' Do **NOT** set m_bIsDirty itself, as this will using the Property Let proc will
' ensure the Close button caption is always synchronized with the value of the property.
Private m_bIsDirty              As Boolean




'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub cmdClose_Click()
    Unload Me
End Sub




'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    Const cstrCurrentProc       As String = "Form_Load"
    Const cintNoRowsInTable     As Integer = 0
    Const cstrOutputFile        As String = "C:\WrapperTest.txt"
    Dim rst                     As ADODB.Recordset
    Dim fld                     As ADODB.Field
    Dim strProperty1            As String
    Dim strProperty2            As String
    Dim strProperty3            As String
    Dim dteToday                As Date
    Dim fso                     As Scripting.FileSystemObject
    
    On Error GoTo PROC_ERR

    '--------------------------------------------------------------------------------------------
    ' Do not change
    '--------------------------------------------------------------------------------------------
    mstrScreenName = Me.Caption
    gerhApp.ScreenName = mstrScreenName
    Me.Icon = LoadResPicture(gcRES_ICON_MAINAPP, vbResIcon)
    With Me
        .Width = mclngMinFormWidth
        .Height = mclngMinFormHeight
    End With
    fnCenterFormOnMDI frmMDIMain, Me
    Set mctlFirstUpdateableField_Add = cmdClose
    Set mctlFirstUpdateableField_Upd = cmdClose
    Set fso = New Scripting.FileSystemObject
    Set m_tsOutput = fso.CreateTextFile(cstrOutputFile, True)
    

    '--------------------------------------------------------------------------------------------
    ' 1.  Verify the table wrapper was instantiated successfully.
    '--------------------------------------------------------------------------------------------
    Set mtWrapper = New ctclmClaim
    
    '!CUSTOMIZE! Include a call to fnDisplayMetaData for each property associated with a table column
    With mtWrapper
        m_tsOutput.WriteLine "----------------------------------------------------"
        m_tsOutput.WriteLine " 1. Meta Data"
        m_tsOutput.WriteLine "----------------------------------------------------"
        fnDisplayMetaData "AdmnSystCd"
        fnDisplayMetaData "ClmId"
        fnDisplayMetaData "ClmInsdDthDt"
        fnDisplayMetaData "ClmInsdFirstNm"
        fnDisplayMetaData "ClmInsdLastNm"
        fnDisplayMetaData "ClmInsdSsnNum"
        fnDisplayMetaData "ClmNum"
        fnDisplayMetaData "ClmPolNum"
        fnDisplayMetaData "ClmProofDt"
        fnDisplayMetaData "ClmTotClmPdAmt"
        fnDisplayMetaData "ClmTotDthbPmtAmt"
        fnDisplayMetaData "ClmTotIntAmt"
        fnDisplayMetaData "ClmTotWthldAmt"
        fnDisplayMetaData "InsdDthResStCd"
        fnDisplayMetaData "IssStCd"
        fnDisplayMetaData "LstUpdtDtm"
        fnDisplayMetaData "LstUpdtUserId"
        fnDisplayMetaData "PycoTypCd"
    End With
    
    '--------------------------------------------------------------------------------------------
    ' 2.  Verify that the Lookup recordset was successfully populated (right columns, and that the
    '     number of rows retrieved is appropriate)
    '--------------------------------------------------------------------------------------------
    m_tsOutput.WriteLine vbCrLf
    m_tsOutput.WriteLine "----------------------------------------------------"
    m_tsOutput.WriteLine " 2. Lookup Recordset Population"
    m_tsOutput.WriteLine "----------------------------------------------------"
    m_tsOutput.WriteLine "The m_rstLookup recordset was populated as follows: "
    m_tsOutput.WriteLine "   Record Count = " & CStr(mtWrapper.LookupRecordCount)
    m_tsOutput.WriteLine "   Current Rec# = " & CStr(mtWrapper.CurrentLookupRecordNumber)
    
    Set rst = mtWrapper.LookupData
     
    For Each fld In rst.Fields
        m_tsOutput.WriteLine "   For the first record in the m_rstLookup recordset:"
        m_tsOutput.WriteLine "     column = " & fld.Name & ", value = [" & CStr(fnZLSIfNull(fld.value)) & "]"
    Next fld
    
     '--------------------------------------------------------------------------------------------
    ' 3. Test navigation methods
    '--------------------------------------------------------------------------------------------
    m_tsOutput.WriteLine vbCrLf
    m_tsOutput.WriteLine "----------------------------------------------------"
    m_tsOutput.WriteLine " 3. Navigation"
    m_tsOutput.WriteLine "----------------------------------------------------"
    With mtWrapper
        If .LookupRecordCount > 0 Then
            mtWrapper.GoToLastRecord
            m_tsOutput.WriteLine "Navigated to last record, rec #" & .CurrentLookupRecordNumber
            mtWrapper.GoToFirstRecord
            m_tsOutput.WriteLine "Navigated to first record, rec #" & .CurrentLookupRecordNumber
        End If
        If .LookupRecordCount > 1 Then
            mtWrapper.GoToNextRecord
            m_tsOutput.WriteLine "Navigated to next record, rec #" & .CurrentLookupRecordNumber
            mtWrapper.GoToPreviousRecord
            m_tsOutput.WriteLine "Navigated to previous record, rec #" & .CurrentLookupRecordNumber
        End If
    End With

    '--------------------------------------------------------------------------------------------
    ' 4. Test ADD capability
    '--------------------------------------------------------------------------------------------
    m_tsOutput.WriteLine vbCrLf
    m_tsOutput.WriteLine "----------------------------------------------------"
    m_tsOutput.WriteLine " 4. Add"
    m_tsOutput.WriteLine "----------------------------------------------------"
    With mtWrapper
        '!CUSTOMIZE!
        ' Set properties, as if the user input them. Be sure to choose values that
        ' should result in a successful Add.
        dteToday = Now
        .AdmnSystCd = "02"
        .ClmInsdDthDt = CDate("01/03/2003")
        .ClmInsdFirstNm = "BETSY"
        .ClmInsdLastNm = "TESTCASE"
        .ClmInsdSsnNum = "265802222"
        .ClmNum = "1234567"
        .ClmPolNum = "1234567"
        .ClmProofDt = CDate("01/03/2003")
        .ClmTotClmPdAmt = 0
        .ClmTotDthbPmtAmt = 0
        .ClmTotIntAmt = 0
        .ClmTotWthldAmt = 0
        .InsdDthResStCd = "FL"
        .IssStCd = "FL"
        .LstUpdtDtm = Now
        .LstUpdtUserId = "TEST    "
        .PycoTypCd = "I1"
      ' Call the Add method
        .AddRecord
        ' Display the values of the current just-added record
        m_tsOutput.WriteLine "Just added this record: "
        fnDisplayRecord
    End With
    
    '--------------------------------------------------------------------------------------------
    ' 5. Test UPDATE capability
    '--------------------------------------------------------------------------------------------
    m_tsOutput.WriteLine vbCrLf
    m_tsOutput.WriteLine "----------------------------------------------------"
    m_tsOutput.WriteLine " 5. Update"
    m_tsOutput.WriteLine "----------------------------------------------------"
    With mtWrapper
        '!CUSTOMIZE!
        ' Set properties, to simulate the user editing the record just added. Be sure to choose
        ' values that should result in a successful Update.  Be sure to NOT change key fields!
        .InsdDthResStCd = "MA"
        .IssStCd = "MA"
        .LstUpdtDtm = Now
        .LstUpdtUserId = "K000    "
        ' Call the Update method
        .UpdateRecord
        ' Display the values of the current just-updated record
        m_tsOutput.WriteLine "After updating, this record now looks like this: "
        fnDisplayRecord
    End With
    
    '--------------------------------------------------------------------------------------------
    ' 6. Test DELETE capability
    '--------------------------------------------------------------------------------------------
    m_tsOutput.WriteLine vbCrLf
    m_tsOutput.WriteLine "----------------------------------------------------"
    m_tsOutput.WriteLine " 6. Delete"
    m_tsOutput.WriteLine "----------------------------------------------------"
    With mtWrapper
        ' Using properties already set by the Add and Update tests, call the
        ' Delete method to delete the record.
        .DeleteRecord
        ' Display the values of the record previous to the one just deleted
        m_tsOutput.WriteLine "After deleting, the new current record is now this: "
        fnDisplayRecord
    End With
    
    m_tsOutput.Close
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here
    fnFreeObject m_tsOutput
    fnFreeObject fso

    ' Report the error, since this is an event handler
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.ReportFatalError mstrScreenName
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            ' Save Err object data, if not already saved
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub


'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub fnDisplayMetaData(ByVal strTagIn As String)
    Const cstrCurrentProc       As String = "fnDisplayMetaData"
    On Error GoTo PROC_EXIT
    
    With mtWrapper
        m_tsOutput.WriteLine "The " & strTagIn & " property: "
        m_tsOutput.WriteLine "  AllowableCharacters=" & .AllowableCharacters(strTagIn)
        m_tsOutput.WriteLine "  DefaultValue=[" & CStr(.DefaultValue(strTagIn)) & "]"
        m_tsOutput.WriteLine "  DollarPositions=" & CStr(.DollarPositions(strTagIn))
        m_tsOutput.WriteLine "  DecimalPositions=" & CStr(.DecimalPositions(strTagIn))
        m_tsOutput.WriteLine "  IsNullable=" & .IsNullable(strTagIn)
        m_tsOutput.WriteLine "  Format=" & .Format(strTagIn)
        m_tsOutput.WriteLine "  Mask=" & .Mask(strTagIn)
        m_tsOutput.WriteLine "  IsKey=" & .IsKey(strTagIn)
        m_tsOutput.WriteLine "  ShouldForceToUppercase=" & .ShouldForceToUppercase(strTagIn)
        m_tsOutput.WriteLine "  MaxCharacters=" & CStr(.MaxCharacters(strTagIn))
    End With
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here

    ' Report the error, since this is an event handler
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.ReportFatalError mstrScreenName
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            ' Save Err object data, if not already saved
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub




'!CUSTOMIZE! Edit this procedure so it lists each column of your table.
Private Sub fnDisplayRecord()
    Const cstrCurrentProc       As String = "fnDisplayRecord"
    On Error GoTo PROC_EXIT
    
    With mtWrapper
        m_tsOutput.WriteLine "ClmId = [" & .ClmId & "]"
        m_tsOutput.WriteLine "AdmnSystCd = [" & .AdmnSystCd & "]"
        m_tsOutput.WriteLine "ClmInsdDthDt = [" & .ClmInsdDthDt & "]"
        m_tsOutput.WriteLine "ClmInsdFirstNm = [" & .ClmInsdFirstNm & "]"
        m_tsOutput.WriteLine "ClmInsdLastNm = [" & .ClmInsdLastNm & "]"
        m_tsOutput.WriteLine "ClmInsdSsnNum = [" & .ClmInsdSsnNum & "]"
        m_tsOutput.WriteLine "ClmNum = [" & .ClmNum & "]"
        m_tsOutput.WriteLine "ClmPolNum = [" & .ClmPolNum & "]"
        m_tsOutput.WriteLine "ClmProofDt = [" & .ClmProofDt & "]"
        m_tsOutput.WriteLine "ClmTotClmPdAmt = [" & .ClmTotClmPdAmt & "]"
        m_tsOutput.WriteLine "ClmTotDthbPmtAmt = [" & .ClmTotDthbPmtAmt & "]"
        m_tsOutput.WriteLine "ClmTotIntAmt = [" & .ClmTotIntAmt & "]"
        m_tsOutput.WriteLine "ClmTotWthldAmt = [" & .ClmTotWthldAmt & "]"
        m_tsOutput.WriteLine "InsdDthResStCd = [" & .InsdDthResStCd & "]"
        m_tsOutput.WriteLine "IssStCd = [" & .IssStCd & "]"
        m_tsOutput.WriteLine "LstUpdtDtm = [" & .LstUpdtDtm & "]"
        m_tsOutput.WriteLine "LstUpdtUserId = [" & .LstUpdtUserId & "]"
        m_tsOutput.WriteLine "PycoTypCd = [" & .PycoTypCd & "]"
    End With
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here

    ' Report the error, since this is an event handler
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.ReportFatalError mstrScreenName
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            ' Save Err object data, if not already saved
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub


#End If
' % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % %
