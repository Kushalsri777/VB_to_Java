VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmStateRule 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "State Rule"
   ClientHeight    =   5580
   ClientLeft      =   570
   ClientTop       =   1395
   ClientWidth     =   10320
   FillColor       =   &H80000001&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000005&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   10320
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboLobCd 
      Height          =   315
      ItemData        =   "frmStateRule.frx":0000
      Left            =   3075
      List            =   "frmStateRule.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   5
      ToolTipText     =   "The line of business"
      Top             =   735
      Width           =   540
   End
   Begin VB.TextBox txtStrlIntRptgFlrAmt 
      Height          =   315
      Left            =   8925
      TabIndex        =   28
      ToolTipText     =   "The minimum amount of claim interest for which the IRS requires 1099 reporting"
      Top             =   4155
      Width           =   1290
   End
   Begin VB.TextBox txtIruleDsc 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H80000012&
      Height          =   315
      Left            =   1725
      TabIndex        =   14
      Text            =   "The greater of Current Rate and Current Loan Rate"
      Top             =   2715
      Width           =   5115
   End
   Begin VB.TextBox txtStrlSpclInstrTxt 
      Height          =   1065
      Left            =   1725
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      ToolTipText     =   "Instructions concerning how the claim should be paid and/or calculated"
      Top             =   3450
      Width           =   5130
   End
   Begin VB.Frame fraRequiredFrom 
      Caption         =   "Required From"
      Height          =   1290
      Left            =   7140
      TabIndex        =   17
      Top             =   1305
      Width           =   3090
      Begin VB.TextBox txtStrlIntReqdOfstNum 
         Height          =   315
         Left            =   1575
         TabIndex        =   21
         ToolTipText     =   "Number of days past the specified date after which interest must be paid"
         Top             =   720
         Width           =   555
      End
      Begin VB.ComboBox cboReqdIdtypCd 
         Height          =   315
         ItemData        =   "frmStateRule.frx":0004
         Left            =   1575
         List            =   "frmStateRule.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   19
         ToolTipText     =   "Indicates the date used in determining whether interest must be paid"
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblStrlIntReqdOfstNum 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nbr &of Days Offset :"
         Height          =   195
         Left            =   150
         TabIndex        =   20
         Top             =   735
         Width           =   1470
      End
      Begin VB.Label lblReqdIdtypCd 
         Caption         =   "Da&te Type:"
         Height          =   285
         Left            =   150
         TabIndex        =   18
         Top             =   375
         Width           =   1065
      End
   End
   Begin VB.Frame fraCalculatedFrom 
      Caption         =   "Calculated From"
      Height          =   1290
      Left            =   7140
      TabIndex        =   22
      Top             =   2760
      Width           =   3090
      Begin VB.ComboBox cboCalcIdtypCd 
         Height          =   315
         ItemData        =   "frmStateRule.frx":0008
         Left            =   1575
         List            =   "frmStateRule.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   24
         ToolTipText     =   "Indicates the date used in determine how to calculate claims interest"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtStrlIntCalcOfstNum 
         Height          =   315
         Left            =   1575
         TabIndex        =   26
         ToolTipText     =   "Number of days past the specified date after which interest should be calculated"
         Top             =   720
         Width           =   555
      End
      Begin VB.Label lblCalcIdtypCd 
         Caption         =   "Date T&ype:"
         Height          =   285
         Left            =   150
         TabIndex        =   23
         Top             =   375
         Width           =   1065
      End
      Begin VB.Label lbltxtStrlIntCalcOfstNum 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nbr of Days Offs&et:"
         Height          =   195
         Left            =   150
         TabIndex        =   25
         Top             =   735
         Width           =   1425
      End
   End
   Begin VB.ComboBox cboStCd 
      Height          =   315
      Left            =   600
      Style           =   2  'Dropdown List
      TabIndex        =   3
      ToolTipText     =   "The state for which this rule applies"
      Top             =   735
      Width           =   795
   End
   Begin VB.TextBox txtStrlIntRuleAmt 
      Height          =   315
      Left            =   4530
      TabIndex        =   13
      ToolTipText     =   "The amount used, in conjunction with the Rule Code, to calculate claims interest"
      Top             =   2340
      Width           =   1455
   End
   Begin VB.CommandButton cmdNavigate 
      Caption         =   "<<"
      Height          =   375
      Index           =   0
      Left            =   60
      TabIndex        =   30
      ToolTipText     =   "Go to first record"
      Top             =   5025
      Width           =   435
   End
   Begin VB.CommandButton cmdNavigate 
      Caption         =   "<"
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   31
      ToolTipText     =   "Go to previous record"
      Top             =   5025
      Width           =   435
   End
   Begin VB.CommandButton cmdNavigate 
      Caption         =   ">"
      Height          =   375
      Index           =   2
      Left            =   900
      TabIndex        =   32
      ToolTipText     =   "Go to next record"
      Top             =   5025
      Width           =   435
   End
   Begin VB.CommandButton cmdNavigate 
      Caption         =   ">>"
      Height          =   375
      Index           =   3
      Left            =   1320
      TabIndex        =   33
      ToolTipText     =   "Go to last record"
      Top             =   5025
      Width           =   435
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   2708
      TabIndex        =   35
      ToolTipText     =   "Add a new Insured"
      Top             =   5025
      Width           =   1215
   End
   Begin VB.ComboBox cboLookup 
      Height          =   315
      Left            =   840
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Type or select a Claim Number to view"
      Top             =   180
      Width           =   630
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   6488
      TabIndex        =   37
      ToolTipText     =   "Cancel your changes or close this screen"
      Top             =   5025
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   5228
      TabIndex        =   36
      ToolTipText     =   "Delete this Insured"
      Top             =   5025
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   375
      Left            =   3968
      TabIndex        =   34
      ToolTipText     =   "Save your changes"
      Top             =   5025
      Width           =   1215
   End
   Begin VB.ComboBox cboIruleCd 
      Height          =   315
      ItemData        =   "frmStateRule.frx":000C
      Left            =   1725
      List            =   "frmStateRule.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   11
      ToolTipText     =   "Indicates which Sun Life company is responsible for paying this claim"
      Top             =   2325
      Width           =   1515
   End
   Begin MSComCtl2.DTPicker dtpStrlEffDt 
      Height          =   315
      Left            =   1725
      TabIndex        =   7
      ToolTipText     =   "The date on which the rule becomes effective"
      Top             =   1305
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "MM/dd/yyy"
      Format          =   55508995
      CurrentDate     =   37013
      MinDate         =   21916
   End
   Begin MSComCtl2.DTPicker dtpStrlEndDt 
      Height          =   315
      Left            =   1725
      TabIndex        =   9
      ToolTipText     =   "The date on which the rule becomes inactive"
      Top             =   1680
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      CheckBox        =   -1  'True
      CustomFormat    =   "MM/dd/yyy"
      Format          =   55508995
      CurrentDate     =   37013
      MinDate         =   21916
   End
   Begin VSFlex7LCtl.VSFlexGrid vfgLookup 
      Height          =   315
      Left            =   1920
      TabIndex        =   38
      Top             =   120
      Width           =   1215
      _cx             =   2143
      _cy             =   556
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   0
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   1
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
      ShowComboButton =   2
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   1
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Label lblLobCd 
      AutoSize        =   -1  'True
      Caption         =   "&Line Of Business:"
      Height          =   195
      Left            =   1725
      TabIndex        =   4
      Top             =   750
      Width           =   1230
   End
   Begin VB.Label lblStrlIntRptgFlrAmt 
      AutoSize        =   -1  'True
      Caption         =   "Interest Reporting &Floor:"
      Height          =   195
      Left            =   7140
      TabIndex        =   27
      Top             =   4170
      Width           =   1800
   End
   Begin VB.Label lblRecordPosition 
      BackStyle       =   0  'Transparent
      Caption         =   "Record x of y"
      Height          =   285
      Left            =   75
      TabIndex        =   29
      Top             =   4710
      Width           =   2625
   End
   Begin VB.Label lblStrlEndDt 
      AutoSize        =   -1  'True
      Caption         =   "E&nd Date:"
      Height          =   195
      Left            =   600
      TabIndex        =   8
      Top             =   1710
      Width           =   720
   End
   Begin VB.Label lblStrlEffDt 
      AutoSize        =   -1  'True
      Caption         =   "Effecti&ve Date:"
      Height          =   195
      Left            =   600
      TabIndex        =   6
      Top             =   1335
      Width           =   1095
   End
   Begin VB.Label lblStrlSpclInstrTxt 
      AutoSize        =   -1  'True
      Caption         =   "S&pecial Instructions:"
      Height          =   390
      Left            =   600
      TabIndex        =   15
      Top             =   3465
      Width           =   990
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblStCd 
      AutoSize        =   -1  'True
      Caption         =   "&State:"
      Height          =   195
      Left            =   75
      TabIndex        =   2
      Top             =   795
      Width           =   450
   End
   Begin VB.Label lblLookup 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lookup:"
      ForeColor       =   &H80000013&
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   570
   End
   Begin VB.Shape shpLookup 
      BackColor       =   &H80000003&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   510
      Left            =   60
      Top             =   75
      Width           =   10155
   End
   Begin VB.Label lblIruleCd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Rule Code:"
      Height          =   195
      Left            =   600
      TabIndex        =   10
      Top             =   2385
      Width           =   795
   End
   Begin VB.Label lblStrlIntRuleAmt 
      AutoSize        =   -1  'True
      Caption         =   "Rule A&mount:"
      Height          =   195
      Left            =   3480
      TabIndex        =   12
      Top             =   2385
      Width           =   975
   End
End
Attribute VB_Name = "frmStateRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'!TODO! Customize for this form
#If False Then
' - - - - - - - - - - - - - - - - - - - - - - - - - - -
'
' To get around bug 2C14 (beep upon entering the
' Fund screen) the Tab Order for this form was changed.
' Instead of the vfgLookup's label & vfg control being
' at the top of the tab order, the txtFundMgrPrvCd label and
' text box control are.
'
' - - - - - - - - - - - - - - - - - - - - - - - - - - -

'******************************************************************************
' Module     : frmStateRule
' Description:
' Procedures:
'

'
' Modified   :
' --------------------------------------------------
Option Explicit
Option Compare Binary

Private mstrScreenName As String

Private Const mclngMinFormWidth As Long = 10410
Private Const mclngMinFormHeight As Long = 5955

' mtWrapper is an instance of the table wrapper corresponding to the main table maintained by this form.
Private mtWrapper               As ctsrlStateRule

Dim mstrSearchForClaimNumber As String  '!TODO! - remove once code is cleaned-up
'!TODO! Review thoroughly and trash stuff unnecessarily carried over from frmInsured

Private Const gcstrBlankEntry As String = " "
' Define a constant for each field that may get an error or warning. This
' should match the text of that control's associated Label control.
Private Const mcstrTxtClaimNumberLabel As String = "Claim Number"
Private Const mcstrCboSystemLabel As String = "System"
Private Const mcstrCboCompanyLabel As String = "Company"
Private Const mcstrTxtInsuredLabel As String = "Insured"
Private Const mcstrcboStCdLabel As String = "State"
Private Const mcstrDtpDateOfDeathLabel As String = "Date Of Death"
Private Const mcstrDtpDateOfProofLabel As String = "Date Of Proof"
Private Const mcstrTxtClerkCodeLabel As String = "Clerk Code"

Dim mrstLookup As ADODB.Recordset
Dim mrstPayees As ADODB.Recordset
Dim mrstInsureds As ADODB.Recordset

Dim mfRecordEdited As Boolean

' mbInLookupMode determines whether the user is in the process of doing a search using the Lookup ComboBox
Private mbInLookupMode          As Boolean

' mbInAddMode determines whether the user has begun the process of adding a new record to the table.
' Note that Add mode is independent of Update mode
Private mbInAddMode             As Boolean

Dim mstrLOB As String

Private mctlFirstUpdateableField_Add As Control
Private mctlFirstUpdateableField_Upd As Control

' m_bIsDirty corresponds to the public property called IsDirty.
' All maintenance screens should have this field and that property! When True, it indicates
' that the user has made --but not yet saved-- changes to a record. The MDI form will query
' this property if the user opens the File menu, since the Exit option should be disabled if
' any form has outstanding changes.
' Be sure to use this variable's corresponding Property Let to change its value.
' Do **NOT** set m_bIsDirty itself, as this will using the Property Let proc will
' ensure the Close button caption is always synchronized with the value of the property.
Private m_bIsDirty              As Boolean

'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'\                                                                  /
'|                PROPERTY GET/LET    Procedures                    |
'/                                                                  \
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/


'////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Get IsDirty() As Boolean
    ' Returns True if the record displayed in the form has been
    ' edited; False otherwise.
    Const cstrCurrentProc As String = "Property Get IsDirty"
    On Error GoTo PROC_ERR

    IsDirty = m_bIsDirty
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    Exit Property
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Property



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Property Let IsDirty(ByVal bValue As Boolean)
    ' Sets the value of the IsDirty property. This should ONLY be set by this form itself.
    '
    ' Be sure to use this Property Let to change the value of the m_bIsDirty variable.
    ' Do **NOT** set m_bIsDirty itself, since using the Property Let proc will ensure
    ' that the Close button caption is always synchronized with the value of this property.
    Const cstrCurrentProc As String = "Let IsDirty"
    On Error GoTo PROC_ERR

    m_bIsDirty = bValue

    ' Adjust Close button caption accordingly
    If bValue Then
        cmdClose.Caption = "&Cancel"
    Else
        cmdClose.Caption = "&Close"
    End If
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    Exit Property
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Property

'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'\                                                                  /
'|                Procedures and Event Handlers                     |
'/                                                                  \
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub cboIruleCd_Click()
    ' Comments  : Flag the field as having changed
    ' Parameters:  -
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "cboCompany_Click"

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    fnInitializeEditMode
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



' OBSOLETE '////////////////////////////////////////////////////////////////////////////////////////////////
' OBSOLETE Private Sub cboLookup_Click()
' OBSOLETE     ' Comments  :
' OBSOLETE     ' Parameters:  -
' OBSOLETE     ' Modified  :
' OBSOLETE     ' --------------------------------------------------
' OBSOLETE     On Error GoTo PROC_ERR
' OBSOLETE     Const cstrCurrentProc As String = "cboLookup_Click"
' OBSOLETE     Dim intCurrentRecordPosition As Integer
' OBSOLETE     Dim hrgHourglass As chrgHourglass
' OBSOLETE
' OBSOLETE     ' Set screen name in case errors are reported here or
' OBSOLETE     ' in procedures called by this Event Handler
' OBSOLETE     gerhApp.ScreenName = mstrScreenName
' OBSOLETE
' OBSOLETE     ' CMP - not sure where to put the hourglass, so start at the top...
' OBSOLETE     Set hrgHourglass = New chrgHourglass
' OBSOLETE     hrgHourglass.value = True
' OBSOLETE
' OBSOLETE     If mbInAddMode Then
' OBSOLETE         mbInAddMode = False
' OBSOLETE         fnEnableKeyFields False
' OBSOLETE     End If
' OBSOLETE
' OBSOLETE     ' Skip further processing if there are no Insureds, if the empty entry was selected,
' OBSOLETE     ' or if the user specified nothing (i.e. a Null string) in the
' OBSOLETE     ' "Search for Claim Number" box.
' OBSOLETE     If (mrstLookup.RecordCount) = 0 Or ((cboLookup.Text = gcstrBlankEntry) And (mstrSearchForClaimNumber = vbNullString)) Then
' OBSOLETE         GoTo PROC_EXIT
' OBSOLETE     End If
' OBSOLETE
' OBSOLETE     ' If the user just hit Escape, then don't process the return value.
' OBSOLETE     If mstrSearchForClaimNumber = "blankval" Then
' OBSOLETE         mstrSearchForClaimNumber = vbNullString
' OBSOLETE         GoTo PROC_EXIT
' OBSOLETE     End If
' OBSOLETE
' OBSOLETE     ' If we're processing an actual keystroke response, then find the first record
' OBSOLETE     ' that matches the full/partial Claim Number specified
' OBSOLETE     If mstrSearchForClaimNumber <> vbNullString Then
' OBSOLETE        intCurrentRecordPosition = mrstInsureds.AbsolutePosition
' OBSOLETE        mrstInsureds.Find "ClaimNumber like '" & mstrSearchForClaimNumber & "*'", , , adBookmarkFirst
' OBSOLETE        ' If we are at EOF, then the search text wasn't found, so revert back to the
' OBSOLETE        ' original entry.
' OBSOLETE        mstrSearchForClaimNumber = vbNullString
' OBSOLETE        If mrstInsureds.EOF Then
' OBSOLETE           'cboLookup = gcstrBlankEntry
' OBSOLETE           mrstInsureds.AbsolutePosition = intCurrentRecordPosition
' OBSOLETE           GoTo PROC_EXIT
' OBSOLETE        End If
' OBSOLETE     Else
' OBSOLETE        ' Otherwise, they've navigated the list, and chosen an entry, so go there.
' OBSOLETE        mrstInsureds.Find "ClaimNumber = '" & cboLookup.Text & "'", , , adBookmarkFirst
' OBSOLETE     End If
' OBSOLETE
' OBSOLETE     lblRecordPosition = fnShowRecordPosition(mrstInsureds)
' OBSOLETE
' OBSOLETE     fnLoadControls
' OBSOLETE     fnGetChildren
' OBSOLETE     Me.Refresh  ' This is needed to avoid corruption on the display
' OBSOLETE     If mrstInsureds.RecordCount > 1 Then
' OBSOLETE         fnSetNavigationButtons True
' OBSOLETE     Else
' OBSOLETE         ' There is only 1 record in the recordset. Cannot navigate forward/backward
' OBSOLETE         fnSetNavigationButtons False
' OBSOLETE     End If
' OBSOLETE     mfRecordEdited = False
' OBSOLETE     fnSetCommandButtons True
' OBSOLETE     fnInitializeMenuItems
' OBSOLETE
' OBSOLETE     ' "Empty out" the Lookup box by settings its value to the first (blank) entry
' OBSOLETE     cboLookup = gcstrBlankEntry
' OBSOLETE PROC_EXIT:
' OBSOLETE     ' Disable the error handler so errors hit here won't be handled by PROC_ERR
' OBSOLETE     On Error GoTo 0
' OBSOLETE
' OBSOLETE     ' Clean-up statements go here
' OBSOLETE     hrgHourglass.value = False
' OBSOLETE
' OBSOLETE     ' Report the error, since this is an event handler
' OBSOLETE     If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
' OBSOLETE         gerhApp.ReportFatalError mstrScreenName
' OBSOLETE     End If
' OBSOLETE     Exit Sub
' OBSOLETE PROC_ERR:
' OBSOLETE     Select Case Err.Number
' OBSOLETE         'Case statements for expected errors go here
' OBSOLETE         Case Else
' OBSOLETE             ' Save Err object data, if not already saved
' OBSOLETE             gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
' OBSOLETE     End Select
' OBSOLETE     Resume PROC_EXIT
' OBSOLETE End Sub
' OBSOLETE
' OBSOLETE
' OBSOLETE '////////////////////////////////////////////////////////////////////////////////////////////////
' OBSOLETE Private Sub cboLookup_GotFocus()
' OBSOLETE     ' Comments  : Turns on the Lookup flag
' OBSOLETE     ' Parameters: N/A
' OBSOLETE     ' Modified  :
' OBSOLETE     ' --------------------------------------------------
' OBSOLETE     On Error GoTo PROC_ERR
' OBSOLETE     Const cstrCurrentProc As String = "cboLookup_GotFocus"
' OBSOLETE
' OBSOLETE     ' Set screen name in case errors are reported here or
' OBSOLETE     ' in procedures called by this Event Handler
' OBSOLETE     gerhApp.ScreenName = mstrScreenName
' OBSOLETE
' OBSOLETE     mbInLookupMode = True
' OBSOLETE PROC_EXIT:
' OBSOLETE     ' Disable the error handler so errors hit here won't be handled by PROC_ERR
' OBSOLETE     On Error GoTo 0
' OBSOLETE     ' Clean-up statements go here
' OBSOLETE
' OBSOLETE     ' Report the error, since this is an event handler
' OBSOLETE     If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
' OBSOLETE         gerhApp.ReportFatalError mstrScreenName
' OBSOLETE     End If
' OBSOLETE     Exit Sub
' OBSOLETE PROC_ERR:
' OBSOLETE     Select Case Err.Number
' OBSOLETE         'Case statements for expected errors go here
' OBSOLETE         Case Else
' OBSOLETE             ' Save Err object data, if not already saved
' OBSOLETE             gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
' OBSOLETE     End Select
' OBSOLETE     Resume PROC_EXIT
' OBSOLETE End Sub
' OBSOLETE
' OBSOLETE
' OBSOLETE '////////////////////////////////////////////////////////////////////////////////////////////////
' OBSOLETE Private Sub cboLookup_KeyPress(ByRef pintKeyAscii As Integer)
' OBSOLETE     ' Comments  : Allow incremental searching by displaying a
' OBSOLETE     '             "Search for Policy Number" dialog box
' OBSOLETE     ' Parameters: N/A
' OBSOLETE     ' Modified  :
' OBSOLETE     ' --------------------------------------------------
' OBSOLETE     On Error GoTo PROC_ERR
' OBSOLETE     Const cstrCurrentProc As String = "cboLookup_KeyPress"
' OBSOLETE
' OBSOLETE     ' Set screen name in case errors are reported here or
' OBSOLETE     ' in procedures called by this Event Handler
' OBSOLETE     gerhApp.ScreenName = mstrScreenName
' OBSOLETE
' OBSOLETE    ' Our user wants to type some information... Only respond if the key pressed is valid.
' OBSOLETE    ' 48-57 = digits 0-9    65-90 = characters A-Z    97-122 = characters a-z
' OBSOLETE     If (pintKeyAscii > 47 And pintKeyAscii < 58) Or (pintKeyAscii > 64 And pintKeyAscii < 91) Or (pintKeyAscii > 96 And pintKeyAscii < 123) Then
' OBSOLETE        frmSearchForClaimNumber.txtClmNum = Chr$(pintKeyAscii)
' OBSOLETE        frmSearchForClaimNumber.txtClmNum.SelStart = Len(frmSearchForClaimNumber.txtClmNum.Text)
' OBSOLETE        frmSearchForClaimNumber.Show vbModal
' OBSOLETE        ' We're using a frmSearchForClaimNumber as a modal - once we're done, get the text, and unload the form.
' OBSOLETE        mstrSearchForClaimNumber = frmSearchForClaimNumber.txtClmNum.Text
' OBSOLETE        Me.Refresh
' OBSOLETE        Unload frmSearchForClaimNumber
' OBSOLETE        ' Since we intercepted the keystroke, and passed it on to frmSearchForClaimNumber, we don't want to pass the same key onto
' OBSOLETE        ' cbolookup..so, return a space. This will cause cbolookup_click to occur.
' OBSOLETE        pintKeyAscii = 32
' OBSOLETE     End If
' OBSOLETE PROC_EXIT:
' OBSOLETE     ' Disable the error handler so errors hit here won't be handled by PROC_ERR
' OBSOLETE     On Error GoTo 0
' OBSOLETE     ' Clean-up statements go here
' OBSOLETE
' OBSOLETE     ' Report the error, since this is an event handler
' OBSOLETE     If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
' OBSOLETE         gerhApp.ReportFatalError mstrScreenName
' OBSOLETE     End If
' OBSOLETE     Exit Sub
' OBSOLETE PROC_ERR:
' OBSOLETE     Select Case Err.Number
' OBSOLETE         'Case statements for expected errors go here
' OBSOLETE         Case Else
' OBSOLETE             ' Save Err object data, if not already saved
' OBSOLETE             gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
' OBSOLETE     End Select
' OBSOLETE     Resume PROC_EXIT
' OBSOLETE End Sub
' OBSOLETE
' OBSOLETE
' OBSOLETE '////////////////////////////////////////////////////////////////////////////////////////////////
' OBSOLETE Private Sub cboLookup_LostFocus()
' OBSOLETE     ' Comments  : Turns on the Lookup flag
' OBSOLETE     ' Parameters: N/A
' OBSOLETE     ' Modified  :
' OBSOLETE     ' --------------------------------------------------
' OBSOLETE     On Error GoTo PROC_ERR
' OBSOLETE     Const cstrCurrentProc As String = "cboLookup_LostFocus"
' OBSOLETE
' OBSOLETE     ' Set screen name in case errors are reported here or
' OBSOLETE     ' in procedures called by this Event Handler
' OBSOLETE     gerhApp.ScreenName = mstrScreenName
' OBSOLETE
' OBSOLETE     mbInLookupMode = False
' OBSOLETE PROC_EXIT:
' OBSOLETE     ' Disable the error handler so errors hit here won't be handled by PROC_ERR
' OBSOLETE     On Error GoTo 0
' OBSOLETE     ' Clean-up statements go here
' OBSOLETE
' OBSOLETE     ' Report the error, since this is an event handler
' OBSOLETE     If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
' OBSOLETE         gerhApp.ReportFatalError mstrScreenName
' OBSOLETE     End If
' OBSOLETE     Exit Sub
' OBSOLETE PROC_ERR:
' OBSOLETE     Select Case Err.Number
' OBSOLETE         'Case statements for expected errors go here
' OBSOLETE         Case Else
' OBSOLETE             ' Save Err object data, if not already saved
' OBSOLETE             gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
' OBSOLETE     End Select
' OBSOLETE     Resume PROC_EXIT
' OBSOLETE End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub cmdAdd_Click()
    ' Comments  : Handles the adding of a new record.
    ' Parameters:  -
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "cmdAdd_Click"

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    fnAddRecord
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


'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub cmdClose_Click()
    ' Comments  : If the user clicked the Close button, see if
    '             there are outstanding data changes that have not been saved.
    '             If so, instruct the user how to proceed depending on whether
    '             they want to save or lose their changes.
    '
    '             NOTE: The logic in this function should closely resemble that
    '                   in the Form_QueryUnload event handler!
    ' Parameters: N/A
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "cmdClose_Click"
    Dim strMsg As String

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    Unload Me
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


'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub cmdDelete_Click()
    ' Comments  : Deletes the current record. Note: This button
    '             will be disabled if any children to this
    '             record (i.e. Payees to this Insured) exist,
    '             forcing the user to first delete those children
    '             and then delete the parent.
    ' Parameters: N/A
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "cmdDelete_Click"
    Dim intButtonClicked As Integer
    Dim hrgHourglass As chrgHourglass

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    ' .......................................................................
    ' Make sure the user really, really, really wants to delete this record.
    ' .......................................................................
    ' 3002 = Are you sure you want to delete this record?
    intButtonClicked = gerhApp.ReportNonFatal(vbObjectError + gcRES_ALRT_OK_TO_DELETE_RECORD, _
                           mstrScreenName & gcstrDOT & cstrCurrentProc)

    If (intButtonClicked = vbNo) Or (intButtonClicked = gcintClickedCloseButton) Then
        GoTo PROC_EXIT
    End If

    Set hrgHourglass = New chrgHourglass
    hrgHourglass.value = True

    mrstInsureds.DELETE
    fnRequeryAndRepositionAfterDelete
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0

    ' Clean-up statements go here
    hrgHourglass.value = False

    ' Report the error, since this is an event handler
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.ReportFatalError mstrScreenName
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case -2147217864
            ' ADO error: "Row cannot be located for updating. Some values may have been
            ' changed since it was last read."
            ' ...basically, another user deleted or updated the record since THIS user retrieved it
'!TODO! Gen msg via frmMsgBox
            'MsgBox "Another user changed or deleted this record since you began viewing it. " & _
            '       "Your request to delete the record cannot be done. " & vbCrLf & vbCrLf & _
            '       "If the record is still displayed once you click OK, then just try the " & _
            '       "Delete again. " & vbCrLf & vbCrLf & _
            '       "If the record is no longer displayed once you click on OK, " & _
            '       "then the record was deleted by some other user.", vbOKOnly + vbExclamation, _
            '       mcstrDialogTitle
            mrstInsureds.CancelUpdate   ' Discard pending row changes
            fnRequeryAndRepositionAfterDelete
        Case -2147467259
            ' ADO error: "The record cannot be deleted or changed because table 'PAYEE' includes
            ' related records"
            ' ...basically, another user added the first PAYEE since THIS user retrieved this Insured
'!TODO! Gen msg via frmMsgBox
            'MsgBox "Another user added one or more Payees since you began viewing this Insured. " & _
            '       "Your request to delete the record cannot be done. ", _
            '       vbOKOnly + vbExclamation, mcstrDialogTitle
            mrstInsureds.CancelUpdate   ' Discard pending row changes
            fnRequeryAndRepositionAfterDelete
        Case Else
            ' Save Err object data, if not already saved
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub cmdNavigate_Click(ByRef pintIndex As Integer)
    ' Comments  : Enables/Disables the navigation buttons
    '             which is a control array:
    '             (0) = go to first record
    '             (1) = go to prev  record
    '             (2) = go to next  record
    '             (3) = go to last  record
    ' Parameters:  -
    ' Modified  :
    ' --------------------------------------------------
    Const cstrCurrentProc As String = "cmdNavigate_Click"
    On Error GoTo PROC_ERR

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

#If False Then
    With mrstInsureds
        Select Case pintIndex
            Case 0
                .MoveFirst
            Case 1
                .MovePrevious
                If .BOF Then
                    .MoveFirst
                End If
            Case 2
                .MoveNext
                If .EOF Then
                    .MoveLast
                End If
            Case 3
                .MoveLast
        End Select
        lblRecordPosition = fnShowRecordPosition(mrstInsureds)
    End With

    fnLoadControls                  ' Populate controls with the "new" current record
    fnGetChildren                   ' Get data from subordinate table(s)...Payees in this case
    fnSetNavigationButtons True     ' Enable navigation buttons
    mfRecordEdited = False
    fnSetCommandButtons True        ' Enable command buttons

    fnSetFocusToFirstUpdateableField
#End If

    With mtWrapper
        Select Case pintIndex
            Case navFirst
                .GoToFirstRecord
            Case navPrev
                .GoToPreviousRecord
            Case navNext
                .GoToNextRecord
            Case Else   ' Go to Last
                .GoToLastRecord
        End Select

        IsDirty = False
    
        If (.CurrentLookupRecordNumber = adPosBOF) Or _
        (.CurrentLookupRecordNumber = adPosEOF) Or _
        (.CurrentLookupRecordNumber = adPosUnknown) Then
            gerhApp.ReportNonFatal vbObjectError + gcRES_INFO_TABLE_IS_EMPTY, _
                                   mstrScreenName & gcstrDOT & cstrCurrentProc
            fnAddRecord
        Else
            ' Note that the Lookup VSFlexGrid control's selection is no longer synchronized
            ' with the table wrapper's CurrentLookupRecordNumber. In other words,
            ' the CurrentLookupRecordNumber may indicate we're on the 5th record and,
            ' by virtue of fnLoadControls being called following each navigation, that should
            ' the same record that is currently displayed on-screen. However, the Lookup
            ' VSFlexGrid is not necessarily *itself* positioned to the 5th record.
            ' The total number of entries in that control, however, should jive with the
            ' table wrapper's LookupRecordCount property.

            ' Clear the Lookup VSFlexGrid control's selection
            vfgLookup.Select 0, 0

            ' Load current record's properties to form's controls, reset navigation buttons and set "rec x of y" label
            fnLoadControls
            IsDirty = False
            fnSetCommandButtons True
        End If
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



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub cmdUpdate_Click()
    ' Comments  : Validates and applies changes to an
    '             existing record, as well as saves the
    '             data associated with a new record.
    ' Parameters:  -
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "cmdUpdate_Click"
    Dim hrgHourglass As chrgHourglass

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    If fnValidData Then
        Set hrgHourglass = New chrgHourglass
        hrgHourglass.value = True

        If mbInAddMode Then
            mrstInsureds.AddNew
            'mintLastRecord = mintLastRecord + 1
        End If

        ' Load screen fields to the current record in the Recordset
        fnLoadRecord
        ' Save the changes to the current record in the Recordset
        mrstInsureds.Update

        fnRequeryAndRepositionAfterUpdate
    End If
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0

    ' Clean-up statements go here
    hrgHourglass.value = False
    
    ' Report the error, since this is an event handler
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.ReportFatalError mstrScreenName
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case -2147217864
            ' ADO error: "Row cannot be located for updating. Some values may have been
            ' changed since it was last read."
            ' ...basically, another user deleted or updated the record since THIS user retrieved it
'!TODO! Gen msg via frmMsgBox
            'MsgBox "Another user changed or deleted this record since you began viewing it. " & _
            '       "Your changes cannot be saved. ", vbOKOnly + vbExclamation, _
            '       mcstrDialogTitle
            mrstInsureds.CancelUpdate   ' Discard pending row changes
            fnRequeryAndRepositionAfterUpdate
        Case -2147467259
            ' Jet Engine Error: "The changes you requested to the table were not successful
            ' because they would create duplicate values in the index, primary key,
            ' or relationship. Change the data in the field or fields that contain duplicate
            ' data, remove the index, or redefine the index to permit duplicate entries
            ' and try again."
            ' ... basically, user is trying to add a record with a key that already exists
'!TODO! Gen msg via frmMsgBox
            'MsgBox "The Claim Number you have specified (" & txtClaimNumber & ") already exists " & _
            '       "in the database. After clicking OK, please change the Claim Number or " & _
            '       "press Escape to abandon the Add of this record.", vbOKOnly + vbExclamation, _
            '       mcstrDialogTitle
            ' .CancelUpdate added for bug 0038, which states that once the user gets the above
            ' "duplicate key" error, it doesn't go away.
            mrstInsureds.CancelUpdate   ' Discard row inserted by .AddNew
        Case Else
            ' Save Err object data, if not already saved
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub fnAddRecord()
    ' Comments  : This function handles adding a new record. It is called
    '             by cmdAdd_Click (when the user clicks the Add button)
    '             and by cmdDelete_Click (when the last record in the
    '             recordset is deleted)
    ' Parameters:  -
    ' Returns   :  -
    ' Modified  :
    ' --------------------------------------------------
    Const cstrCurrentProc As String = "fnAddRecord"

    On Error GoTo PROC_ERR

#If False Then
    ' All we do here is display an empty record. The cmdUpdate_Click event
    ' handler actually does the add when it sees that mbInAddMode=True.
    ' Adds and Updates are treated very nearly the same in that event handler!

    mbInAddMode = True
    ' Display empty (initialized) values for on-screen controls
    fnClearControls

    ' Call fnGetChildren, which result in an mrstPayees Recordset
    ' with 0 records since there should be none with a key =
    ' txtClaimNumber (which is set to a null string by fnClearControls).
    ' Calling fnGetChildren also serves the purpose of initializing
    ' the msgPayees grid control to reflect 0 Payees.
    fnGetChildren

    ' Enable and set focus to the Claim Number (the key field
    ' to the record associated with mrstInsureds so the user can
    ' specify a value.
    fnEnableKeyFields True

    ' Restrike "Record x of y" to reflect pending Add
    lblRecordPosition = fnShowRecordPosition(mrstInsureds)

    mfRecordEdited = False
    fnSetCommandButtons False
    fnSetNavigationButtons False
#End If

    ' All we do here is display an empty record. The cmdUpdate_Click event
    ' handler actually does the add when it sees that mbInAddMode=True.
    ' Adds and Updates are treated very nearly the same in that event handler!

    mbInAddMode = True

    ' Display empty or initialized values for on-screen controls
    fnClearControls
    IsDirty = False

    ' No need to call fnGetChildren/fnHaveDependents here (unlike Claims Interest).

    ' Enable and set focus to key field(s) so the user can specify a value.
    ' This **must** be done as the user goes into Add mode, so they can specify
    ' the key(s) for the record they're adding.
    fnSetAvailabilityOfKeyFields

    ' Restrike "Record x of y" to reflect pending Add. Can't call fnShowRecordPosition
    ' since it is based on a recordset's AbsolutePosition which, in unbound /disconnected mode,
    ' isn't set appropriately.
    lblRecordPosition = "Record ? of " & mtWrapper.LookupRecordCount

    IsDirty = False
    fnSetCommandButtons False

    fnSetNavigationButtons bUnconditionalDisable:=True
    
    ' Make sure first field gets the focus. Note, when Add mode is triggered
    ' from Form_Load, this statement accomplishes nothing: the control isn't yet visible,
    ' so it can't receive the focus. This is why Form_Activate must also call this function.
    fnSetFocusToFirstUpdateableField
    '!TODO! Once the issue re: the cboAdmnSystCd dropping down without our wanting it to upon
    '       initial display of this screen, then the mctlFirstUpdateableField_Add should be
    '       put back to cboAdmnSystCd and this next IF statement removed.
    If cboAdmnSystCd.Visible Then
        cboAdmnSystCd.SetFocus
    End If
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub


'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub fnClearControls()
    ' Comments  : Initializes screen controls in order to add a new record
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnClearControls"


    ' cboStCd                   loaded
    ' cboLobCd                  loaded
    ' dtpStrlfEffDt
    ' dtpStrlEndDt
    ' cboIruleCd                loaded
    ' txtStrlIntRuleAmt
    ' txtIruleDsc
    ' txtStrlSpclInstrTxt
    ' cboReqdIdtypCd            loaded
    ' txtStrlIntReqdOfstNum
    ' cboCalcIdtypCd            loaded
    ' txtStrlIntCalcOfstNum
    ' txtStrlIntRptgFlrAmt
    
'!TODO! Customize for this form
#If False Then
    txtClmNum = vbNullString

    ' For System and Company combo boxes, select 1st entry by default.
    ' Can't set to Null since it's "limit to list"
    cboSystem.ListIndex = 0         ' "SOLAR"
    cboIruleCd.ListIndex = 0        ' "PROOF"
    cboStCd.ListIndex = 0   ' " Other"



    ' DateTimePicker controls (dtpDateOfDeath and dtpDateOfProof) will
    ' automatically be set to today's date. Cannot set them to Null
    ' unless their CheckBox property is set to True.
    dtpStrlEffDt.value = Date
    
        ' For DTPicker controls that correspond to nullable columns whose current value
    ' is Null, we want it to appear with its Checkbox deselected (indicating there is
    ' no date) but with the current date as its value in case the user selects
    ' the Checkbox to specify a Freeze Dt. When the current value is moved to the DTPicker
    ' control, the Checkbox will become deselected if the current value is Null.
    '
    ' NOTE: The Checkbox property just indicates whether a Checkbox should be displayed
    '       on the control. It does **not** indicate whether there a date has or hasn't
    '       been set.
    dtpStrlEndDt.CheckBox = False
    dtpStrlEndDt.value = Date          ' Set value that will still display even after...
    dtpStrlEndDt.CheckBox = True
    dtpStrlEndDt.value = Null          ' ..set to Nulls
#End If
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub fnEnableKeyFields(fEnable As Boolean)
    ' Comments  : Enables/Disables entry to the txtClaimNumber field.
    '             It should only be enabled if in Add mode.
    ' Called By : fnAddRecord, cboLookup_Click and cmdUpdate_Click
    ' Parameters: fEnable=True to enable/unlock it; False otherwise
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnEnableKeyFields"

    With cboStCd
        If fEnable Then
            .Locked = False
            .TabStop = True
            .BackColor = vbWindowBackground
            .ForeColor = vbWindowText
            If .Visible Then
                .SetFocus
            End If
        Else
            .Locked = True
            .TabStop = False
            .BackColor = vbButtonFace
            .ForeColor = vbButtonText
            ' Move focus off this field (to the first editable field) now that it's locked again
            fnSetFocusToFirstUpdateableField
        End If
    End With
    
    With cboLobCd
        If fEnable Then
            .Locked = False
            .TabStop = True
            .BackColor = vbWindowBackground
            .ForeColor = vbWindowText
            If .Visible Then
                .SetFocus
            End If
        Else
            .Locked = True
            .TabStop = False
            .BackColor = vbButtonFace
            .ForeColor = vbButtonText
            ' Move focus off this field (to the first editable field) now that it's locked again
            fnSetFocusToFirstUpdateableField
        End If
    End With
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub fnGetChildren()
    ' Comments  : Loads data associated from tables that are
    '             subordinate (i.e. children) to the table
    '             supplying the main data for this form
    ' Parameters:  -
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnGetChildren"
    Dim strSQL As String

    ' --- Build the Recordset object for Payee data (mrstPayees) ---
    '     that's associated with the current Insured.

    ' Close recordset, if open. This is to avoid an ADO error 3705 "Operation not
    ' allowed when object is open" which can occur the 2nd time you're in here
    ' trying to change the properties of and open the mrstPayees Recordset object.
    mrstPayees.Close

    strSQL = "Select [ClaimNumber], " & _
             "[PAYEE], [ADDRESS1], [ADDRESS2], [STATE], [TIN], " & _
             "[INTEREST] As InterestAmt, [TOTAL] As TotalAmt, " & _
             "[Payment] As PaymentAmt, [Rate], [Withholdingrate] As WithholdingPercent, " & _
             "[Withheld] As WithheldAmt, [Date of Payment] As DateOfPayment " & _
             " FROM [PAYEE] WHERE [ClaimNumber] = '" & _
             "010001001" & "' ORDER BY [PAYEE]"     'TEMP - 010001001 was txtClaimNumber
    ' CursorType=adOpenKeyset - Scrolling fwd/bwd permitted, chgs/del by other users visible
    ' LockType=adLockReadOnly - Recordset is read-only
    mrstPayees.Open Source:=strSQL, _
                    ActiveConnection:=gconAppActive, _
                    CursorType:=adOpenKeyset, _
                    LockType:=adLockReadOnly, _
                    Options:=adCmdText

PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case 3704 ' ADO: Operation not allowed when object is closed
            ' This can occur when we try to close the mrstPayees Recordset object on the
            ' very first time through this procedures, when that object is not yet open.
            Resume Next
        Case Else
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub fnInitializeEditMode()
    ' Comments  : Sets up the environment for editing a record.
    ' Parameters: N/A
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnInitializeEditMode"

    If mfRecordEdited = False Then
        mfRecordEdited = True
        fnSetCommandButtons False
        fnSetNavigationButtons False
        fnInitializeMenuItems
    End If
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub fnLoadCboIruleCd()
    ' Comments  : Populates cboIruleCd combo box
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnLoadCboIruleCd"

    '!TODO! - Load from INTEREST_RATE_RULE_T table
    cboIruleCd.AddItem "CURRT"
    cboIruleCd.AddItem "CURRT+X"
    cboIruleCd.AddItem "CURRT-X"
    cboIruleCd.AddItem "GTCRT&LN"
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub fnLoadCboCalcIdtypCd()
    ' Comments  : Populates cboCalcIdtypCd combo box
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnLoadCboCalcIdtypCd"

    '!TODO! - Load from INTEREST_DATE_TYPE_T table
    cboCalcIdtypCd.AddItem "NONE"
    cboCalcIdtypCd.AddItem "DEATH"
    cboCalcIdtypCd.AddItem "PROOF"
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub fnLoadCboLobCd()
    ' Comments  : Populates cboLobCd combo box
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnLoadCboLobCd"

    '!TODO! - Load from LINE_OF_BUSINESS_T table
    cboLobCd.AddItem "I"
    cboLobCd.AddItem "G"
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub


'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub fnLoadCboLookup()
    ' Comments  : Populates cboLookup combo box
    ' Parameters: None
    ' Modified  :
    '  01/2002 BAW Changed the cbo population to use the
    '              fnADORecordSetToComboBox( ) function to
    '              improve performance. Also added a DoEvents
    '              if the Splash screen is loaded.
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnLoadCboLookup"

    ' Empty out existing contents, if any
    cboLookup.Clear
    
    ' Add " " row as the first entry. This will be used to initialize the
    ' Lookup combo box when a lookup action has been completed. Then
    cboLookup.AddItem gcstrBlankEntry
    fnADORecordSetToComboBox mrstLookup, cboLookup, "ClaimNumber", , False
    
    ' Allow the progress meter on the splash screen to get updated
    If fnIsFormLoaded("frmSplash") Then
        DoEvents
    End If
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case 3021 ' ADO: Either BOF or EOF is True, or the current record has been deleted. Requested operation requires a current record.
            ' This will occur when there are (as yet) no Payees for the current Claim Number
            Resume Next
        Case Else
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub fnLoadCboReqdIdtypCd()
    ' Comments  : Populates cboReqdIdtypCd combo box
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnLoadCboReqdIdtypCd"

    '!TODO! - Load from INTEREST_DATE_TYPE_T table
    cboReqdIdtypCd.AddItem "NONE"
    cboReqdIdtypCd.AddItem "DEATH"
    cboReqdIdtypCd.AddItem "PROOF"
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub fnLoadCboStCd()
    ' Comments  : Populates cboStCd combo box
    ' Parameters: None
    ' Modified  :
    '  01/2002 BAW Changed the cbo population to use the
    '              fnADORecordSetToComboBox( ) function to
    '              improve performance.
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnLoadCboStCd"
    Dim rstStates As ADODB.Recordset
    Dim strSQL As String

    ' Build recordset that drives Insured State combo box
    ' "DISTINCT" used to make sure only unique entries are listed. Without this,
    ' states with multiple rows for different lines-of-business (LOB = G or I) would
    ' be listed twice.
    Set rstStates = New ADODB.Recordset
    strSQL = "SELECT DISTINCT [StateAbbr] FROM [State98] ORDER BY StateAbbr"

    With rstStates
        ' CursorType=adOpenStatic   - Scrolling fwd/bwd permitted, add/chg/del by other users not visible
        ' LockType=adLockReadOnly   - Read-only; Modifications are not permitted
        .Open Source:=strSQL, ActiveConnection:=gconAppActive, _
                    CursorType:=adOpenStatic, LockType:=adLockReadOnly, Options:=adCmdText
    End With
    ' Add " " (blank) row to make it more obvious to the user that this should be filled in.
    ' This will be the default value when doing an Add.
    cboStCd.AddItem gcstrBlankEntry
    ' Populate the combobox using the recordset
    fnADORecordSetToComboBox rstStates, cboStCd, "StateAbbr", , False
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0

    ' Clean-up statements go here
    fnFreeRecordset rstStates
    
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub




'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub fnLoadControls()
    ' Comments  : Populates screen controls with data from recordset
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnLoadControls"

#If False Then
    With mrstInsureds
        ' the following will trigger txtClaimNumber_Change( ) which sets the line-of-business (LOB)
        ' and then repopulates the System combo box based on the LOB
        txtClaimNumber = !ClaimNumber

        If IsNull(!System) Then
            ' Select 1st entry, by default. Can't set to Null since it's "limit to list"
            cboSystem.ListIndex = 0
        Else
            cboSystem = !System
        End If

        If IsNull(!Company) Then
            ' Select 1st entry, by default. Can't set to Null since it's "limit to list"
            cboIruleCd.ListIndex = 0
        Else
            cboIruleCd = !Company
        End If

        txtInsured = fnIfNull(!Insured)

        If IsNull(!State) Then
            ' Select 1st entry, by default. Can't set to Null since it's "limit to list"
            cboStCd.ListIndex = 0
        Else
            cboStCd = !State
        End If

        dtpDateOfDeath.value = !DateOfDeath
        dtpDateOfProof.value = !DateOfProof

        ' Save the original value of these 2 dates fields. If they change and Payees
        ' exist at that time, a warning should be issued to indicate the change may
        ' necessitate a recalculation of the Payee's values.
        mstrOrigDateOfDeath = !DateOfDeath
        mstrOrigDateOfProof = !DateOfProof

        txtStrlIntRuleAmt = fnIfNull(!ClerkCode)
        txtTotalPayments = !TotalPayments

        txtTotalWithheld = !TotalWithheld
        mcurTotalWithheld = !TotalWithheld  ' the unformatted version of txtTotalWithheld

        txtTotalWithInterest = !TotalWithInterest
        txtTotalClaimInterest = !TotalInterest
    End With
#End If


    cboStCd = "TX"
    cboLobCd = "I"
    dtpStrlEffDt = #8/15/2001#
    
    
    ' For DTPicker controls that correspond to nullable columns whose current value
    ' is Null, we want it to appear with its Checkbox deselected (indicating there is
    ' no date) but with the current date as its value in case the user selects
    ' the Checkbox to specify a Freeze Dt. When the current value is moved to the DTPicker
    ' control, the Checkbox will become deselected if the current value is Null.
    '
    ' NOTE: The Checkbox property just indicates whether a Checkbox should be displayed
    '       on the control. It does **not** indicate whether there a date has or hasn't
    '       been set.
    dtpStrlEndDt.CheckBox = False
    dtpStrlEndDt.value = Date          ' Set value that will still display even after...
    dtpStrlEndDt.CheckBox = True
    dtpStrlEndDt.value = Null          ' ..set to Nulls
    
    
    cboIruleCd.ListIndex = 0
    txtStrlIntRuleAmt = vbNullString
    txtStrlSpclInstrTxt = "REFER TO ACLI REGS ON 'DELAYED PAYMENTS'"
    '!TODO! Add logic to turn double-quotes within special instructions field to single-quotes
    cboReqdIdtypCd = "PROOF"
    txtStrlIntReqdOfstNum = 0
    cboCalcIdtypCd = "PROOF"
    txtStrlIntCalcOfstNum = 0
    txtStrlIntRptgFlrAmt = "$600.00"
    
    
    ' Set to False to show there are no pending changes. Loading data to controls above
    ' could trigger fnInitializeEditMode to falsely think there is a pending change.
    mfRecordEdited = False
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub fnLoadRecord()
    ' Comments  : Populates DB record with data from screen controls
    '             in anticipation of saving it as a new or updated record.
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnLoadRecord"

'!TODO! Customize for this form
#If False Then
    With mrstInsureds
        ' Only set ClaimNumber if in Add mode. This is to avoid a spurious
        ' ADODB -2147467259 "The record cannot be deleted or changed because table
        ' "Payees"..." error caused by changing the key field (even overlaying it
        ' with the same value) of table that has dependent tables.
        If mbInAddMode Then
            !ClaimNumber = txtClaimNumber
        End If

        !System = cboSystem
        !Company = cboIruleCd
        !Insured = txtInsured
        !State = cboStCd

        !DateOfDeath = dtpDateOfDeath.value
        !DateOfProof = dtpDateOfProof.value

        !ClerkCode = txtStrlIntRuleAmt

        fnLoadRecordWithCalculatedControls
    End With
#End If
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub fnLoadRecordWithCalculatedControls()
    ' Comments  : Populates DB record with data from screen controls
    '             that are calculated
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnLoadRecordWithCalculatedControls"

#If False Then
    If Not mbInAddMode Then
        ' Extra precaution...to always calc totals across Payees
        ' before doing a save. This will also ensure the totals
        ' are 0 for an Add.  Can't call this prodedure on an Add
        ' since there is no current record and it will get a
        ' ADO 3021 error: "Either BOF or EOF is true or the current
        ' record has been deleted. Requested operation requires a
        ' current record."
        fnCalcTotalsForAllPayees
    End If

    With mrstInsureds
        ' The following fields cannot be edited by the user but are calculated
        ' by the program
        !TotalPayments = txtTotalPayments
        !TotalWithheld = mcurTotalWithheld  ' the unformatted version of txtTotalWithheld
        !TotalWithInterest = txtTotalWithInterest
        !TotalInterest = txtTotalClaimInterest
    End With
#End If
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub fnLoadVfgLookup()
    '--------------------------------------------------------------------------
    ' Procedure:   fnLoadVfgLookup
    ' Description: Populates the ComboBox for the Lookup control
    ' Params:      N/A
    ' Returns:     N/A
    ' Date:        04/12/2002
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc          As String = "fnLoadVfgLookup"
    Const cstrSproc                As String = "dbo.proc_fund_lu_select"     ' Stored procedure to execute
    Dim rstTemp                    As ADODB.Recordset
    Dim prmReturnValue             As ADODB.Parameter
    Dim adwTemp                    As cadwADOWrapper

    On Error GoTo PROC_ERR

    Set adwTemp = New cadwADOWrapper

    vfgLookup.Clear

    If Not (adwTemp.CommandSetSproc(cstrSproc)) Then
        GoTo PROC_EXIT
    End If

    With adwTemp.ADOCommand
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

    If rstTemp.RecordCount <> 0 Then
        fnADORecordsetToVFG rstIn:=rstTemp, _
                        pvfgIn:=vfgLookup
    Else
        ' Add a single empty row so there will be a drop-down arrow and code
        ' that does a selection to force the vfgLookup to work the way TRS needs it to
        ' won't have issues when there are no records in the recordset that is
        ' supposed to populate the control.
        vfgLookup.ColComboList(-1) = " ; "
    End If
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0

    ' Clean-up statements go here
    fnFreeRecordset rstTemp
    fnFreeObject adwTemp

    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case prmReturnValue
        Case gcRES_NERR_ERR_WHILE_TRYING_TO   ' 4028
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_ERR_WHILE_TRYING_TO, _
                                       mstrScreenName & gcstrDOT & cstrCurrentProc, _
                                       "locate"
            Resume PROC_EXIT
    End Select

    ' If any other errors exist, i.e. in Err object, then let it fall through into default error handling.

    Select Case Err.Number
        Case -2147217900 ' Object not found
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_FERR_SPROC_NOT_FOUND, _
                                       mstrScreenName & gcstrDOT & cstrCurrentProc, _
                                       cstrSproc
            Resume PROC_EXIT
        Case Else
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub fnRefreshAllCombos()
    '--------------------------------------------------------------------------
    ' Procedure:   fnRefreshAllCombos
    ' Description: Repopulates each ComboBox or VSFlexGrid control
    '              so they reflect this and other users' changes. This proc
    '              should be called after each Add, Update or Delete.
    '
    ' Params:      N/A
    ' Called by:   cmdUpdate_Click() of frmFund
    '              cmdDelete_Click() of frmFund
    '              Form_Load() of frmFund
    '
    ' Returns:     N/A
    '-----------------------------------------------------------------------------
    '!CUSTOMIZE!    This should call a function to load each ComboBox or
    '               VSFlexGrid control on the form. This will ensure that
    '               when one is refreshed (i.e. to make this and other
    '               user's changes visible), *all* will be.
    Const cstrCurrentProc       As String = "fnRefreshAllCombos"
    On Error GoTo PROC_ERR

    fnLoadVfgLookup         ' #1 = Lookup (FUND_CD)
    'fnLoadCboMktvalFundCd   ' #2 = Market Value Fund Code (MKTVAL_FUND_CD)
    'fnLoadCboFundMgrPrvCd   ' #3 = Fund Mgr (PRV_CD)
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub fnRequeryAndRepositionAfterDelete()
    '----------------------------------------------------------------------------
    ' Procedure :  Sub fnRequeryAndRepositionAfterDelete
    ' Created by:  BAW on 04-26-2001 08:55
    '
    ' Comments  : Requeries recordsets and repositions them. This procedure is called
    '             after a Delete is successfully performed, or one is
    '             attempted but gets a "another user has changed or deleted this
    '             record..." sort of multi-user error.
    ' Called by : cmdDelete_Click
    ' Parameters: None
    ' Modified  :
    '----------------------------------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnRequeryAndRepositionAfterDelete"

    ' Make sure Recordsets reflect record(s) added/deleted by this and
    ' other users. For just-added records, this will ensure they appear in
    ' sorted sequence in the Recordset, rather than appearing at the end.
    mrstInsureds.Requery
    mrstLookup.Requery
    fnLoadCboLookup

'!TODO! Customize for this form
#If False Then
    If mrstInsureds.RecordCount > 0 Then
        ' Try redisplaying the record the user just tried to delete (in case an update
        ' done by another user made *this* user's delete fail. If it is found, terrific.
        '
        ' Otherwise (EOF=True), the conflict must have been caused by another user deleting the
        ' record *this* user was trying to delete, so look for the first record found
        ' with a key *higher* than the one the user just tried to delete. If that fails
        ' (i.e. EOF is true), then just show the last record in the recordset.
        mrstInsureds.MoveFirst
        mrstInsureds.Find Criteria:="[ClaimNumber] = '" & txtClaimNumber & "'", _
                          SkipRecords:=0, SearchDirection:=adSearchForward
        If mrstInsureds.AbsolutePosition = adPosEOF Then
            mrstInsureds.MoveFirst
            mrstInsureds.Find Criteria:="[ClaimNumber] > '" & txtClaimNumber & "'", _
                          SkipRecords:=0, SearchDirection:=adSearchForward
            If mrstInsureds.AbsolutePosition = adPosEOF Then
                mrstInsureds.MoveLast
            End If
        End If

        ' Restrike "Record x of y" to reflect current position (needed if adds/deletes
        ' were done by this or other users
        lblRecordPosition = fnShowRecordPosition(mrstInsureds)

        fnLoadControls
        fnGetChildren

        mfRecordEdited = False
        fnSetFocusToFirstUpdateableField

        If mrstInsureds.RecordCount > 1 Then
            fnSetNavigationButtons True
        Else
            ' There is only 1 record in the recordset
            fnSetNavigationButtons False
        End If

        fnSetCommandButtons True
        fnInitializeMenuItems
    Else
        If mrstInsureds.RecordCount = 0 Then
            ' Requery is necessary to avoid a -2147217885 error (Row handle
            ' referred to a deleted row or a row marked for deletion)
            mrstInsureds.Requery
            fnAddRecord
        End If
    End If
#End If
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub fnRequeryAndRepositionAfterUpdate()
    '----------------------------------------------------------------------------
    ' Procedure :  Sub fnRequeryAndRepositionAfterUpdate
    ' Created by:  BAW on 04-26-2001 08:55
    '
    ' Comments  : Requeries recordsets and repositions them. This should be called
    '             after an Update is successfully performed, or one is
    '             attempted but gets a "another user has changed or deleted this
    '             record..." sort of multi-user error.
    ' Called by : cmdUpdate_Click
    ' Parameters: None
    ' Modified  :
    '----------------------------------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnRequeryAndRepositionAfterUpdate"

'!TODO! Customize for this form
#If False Then
    ' Make sure Recordsets reflect record(s) added/deleted by this and
    ' other users. For just-added records, this will ensure they appear in
    ' sorted sequence in the Recordset, rather than appearing at the end.
    mrstInsureds.Requery
    mrstLookup.Requery
    fnLoadCboLookup

    If mrstInsureds.RecordCount > 0 Then
        ' Reposition the Recordset back to the record just updated/added. If
        ' this record isn't found, the Recordste is positioned to the
        ' end of the Recordset.
        mrstInsureds.MoveFirst
        mrstInsureds.Find Criteria:="[ClaimNumber] = '" & txtClaimNumber & "'", _
                          SkipRecords:=0, SearchDirection:=adSearchForward

        ' If the record has been deleted, it won't be found and the Find method will
        ' leave the Recordset positioned on the last record but with EOF = True.
        ' To avoid an ADO error ("Either EOF or BOF is True...") in fnLoadControls,
        ' we must get rid of the EOF condition.So, move to the first record found
        ' with a key higher than the record just updated.
        If mrstInsureds.AbsolutePosition = adPosEOF Then
            mrstInsureds.MoveFirst
            mrstInsureds.Find Criteria:="[ClaimNumber] > '" & txtClaimNumber & "'", _
                          SkipRecords:=0, SearchDirection:=adSearchForward
            If mrstInsureds.AbsolutePosition = adPosEOF Then
                mrstInsureds.MoveLast
            End If
        End If

        ' Restrike "Record x of y" to reflect current position (needed if adds/deletes
        ' were done by this or other users
        lblRecordPosition = fnShowRecordPosition(mrstInsureds)

        If mbInAddMode Then
            mbInAddMode = False
            fnEnableKeyFields False
        End If

        fnLoadControls
        fnGetChildren

        mfRecordEdited = False
        fnSetFocusToFirstUpdateableField
        fnSetNavigationButtons True
        fnSetCommandButtons True
        fnInitializeMenuItems
    Else
        If mrstInsureds.RecordCount = 0 Then
            fnAddRecord
        End If
    End If
#End If
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub fnSetAvailabilityOfKeyFields()
    '--------------------------------------------------------------------------
    ' Procedure:   fnSetAvailabilityOfKeyFields
    ' Description: Determines whether a control representing a key field
    '              should be display-only.
    '
    ' Params:      N/A
    ' Called by:   fnLoadControls()
    '              cmdAdd_Click()
    ' Returns:     N/A
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc       As String = "fnSetAvailabilityOfKeyFields"
    Dim ctl As Control
    On Error GoTo PROC_ERR

    For Each ctl In Me.Controls
        With ctl
            ' if control corresponds to a SQL Server table column, then try
            ' to set its default properties. The Tag property contains
            ' the name of its property within the table class.
            If Len(.Tag) > 0 Then
                ' If it's a key, disable it unless we're in Add mode
                If mtWrapper.IsKey(.Tag) Then
                    If mbInAddMode Then
                        .Locked = False
                        .TabStop = True
                        .BackColor = vbWindowBackground
                        .ForeColor = vbWindowText
                        .Enabled = True
                    Else
                        .Locked = True
                        .TabStop = False
                        .BackColor = vbButtonFace
                        .ForeColor = vbButtonText
                        .Enabled = False
                    End If
                End If
            End If
        End With
    Next ctl
    
    fnSetFocusToFirstUpdateableField
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub fnSetCommandButtons(ByVal bEnable As Boolean)
    '----------------------------------------------------------------------------
    ' Procedure :  Sub fnSetCommandButtons
    ' Created by:  BAW on 04-26-2001 08:55
    '
    ' Comments  : Enables/Disables the command buttons, per boolean parameter
    '             Here's how the button enabling should work. Note it assumes
    '             that mfRecordEdited and mbInAddMode have been set prior to
    '             calling this function, e.g., they accurately reflect whether
    '             or not there are edits outstanding and/or the user is in
    '             Add mode, respectively.
    '             Remember, though: mbInAddMode and mfRecordEdited are
    '             independent of one another!
    '
    '     State          ADD btn  UPD btn  DEL btn  CLOSE btn PAYEE btn PRTRPT btn
    '    --------------  -------- -------- -------- --------- --------- ----------
    '    Add mode       disabled  enabled  disabled enabled   disabled  disabled
    '    (no edits yet)
    '
    '    Edits o/s      disabled  enabled  disabled enabled   disabled  disabled
    '
    '    No edits o/s   enabled   disabled enabled  enabled   enabled   enabled
    '    & #Payees = 0
    '
    '    No edits o/s   enabled   disabled disabled enabled   enabled   enabled
    '    & #Payees > 0
    '
    ' Called by : fnAddRecord and fnInitializeEditMode, with bEnable = False
    '
    '             cboLookup_Click, cmdDelete_Click, cmdNavigate_Click, cmdUpdate_Click
    '             (when updating existing record) and Form_Load, with
    '             bEnable = True
    '
    ' Parameters: bEnable - indicates whether Add/Update buttons should be enabled
    '                       or disabled
    '
    ' Modified  :
    '----------------------------------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnSetCommandButtons"

    cmdAdd.Enabled = bEnable
    cmdUpdate.Enabled = Not bEnable
    cmdUpdate.Default = Not bEnable


    ' Can only delete an Insured/Claim when (a) there are no Payees and
    ' (b) when you're not in the middle of an Add or Update
    If (mfRecordEdited = False) And (mrstPayees.RecordCount <= 0) _
       And (Not mbInAddMode) Then
            cmdDelete.Enabled = True
    Else
            cmdDelete.Enabled = False
    End If
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case 3704   ' Operation is not allowed when the object is closed
            ' Trying to access mrstPayees (i.e. from Form_Load) when the recordset
            ' has not yet been opened.  Just ignore...a subsequent call, after it
            ' HAS been opened, should set things straight
            Resume Next
        Case Else
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub fnSetFocusToFirstUpdateableField()
    '--------------------------------------------------------------------------
    ' Procedure:   fnSetFocusToFirstUpdateableField
    ' Description: Moves the focus to the first editable (i.e. updateable) field on the screen
    '
    ' Params:      N/A
    ' Called by:
    '
    ' Returns:     N/A
    '-----------------------------------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnSetFocusToFirstUpdateableField"

    ' Set focus to first editable field, by default
    If mbInAddMode Then
        If mctlFirstUpdateableField_Add.Visible Then
            mctlFirstUpdateableField_Add.SetFocus
        End If
    Else
        If mctlFirstUpdateableField_Upd.Visible Then
            mctlFirstUpdateableField_Upd.SetFocus
        End If
    End If
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub fnSetNavigationButtons(Optional ByVal bUnconditionalDisable As Boolean = False)
    '----------------------------------------------------------------------------
    ' Procedure  : fnSetNavigationButtons
    ' Description: Enables/Disables the control array of navigation buttons, based
    '              on the bEnable input parameter
    '
    ' Parameters:  bUnconditionalDisable (in) - indicates whether buttons should be disabled
    '                  regardless of where the current record position is in the recordset.
    '                  This will generally be set to True only via the
    '                  fnAddRecords( ) and fnInitializeEditMode( ) procs.
    '
    ' Called by :
    '              cmdDelete_Click( )
    '              cndNavigate_Click( )
    '              fnAddRecord( )
    '              fnInitializeEditMode( )
    '              Form_Load( )
    '              vfgLookup_Click( )
    '
    ' Returns   :  N/A
    ' Modified  :
    '  04/23/02 BAW  Blended fnRefreshNavigationButtons( ) and fnSetNavigationButtons
    '                so that disabling is always done enmasse, or enabling/disabling
    '                is based on the current record position within the Lookup recordset.
    '----------------------------------------------------------------------------
    Const cstrCurrentProc As String = "fnSetNavigationButtons"
    Dim cmd          As CommandButton
    Dim bHaveRecords As Boolean

    On Error GoTo PROC_ERR

    If bUnconditionalDisable Then
        For Each cmd In cmdNavigate
            cmd.Enabled = False
        Next
        GoTo PROC_EXIT
    End If
    
    
    '...........................................................
    ' Enable navigation buttons based on where we're currently
    ' positioned in the Lookup recordset
    '...........................................................
    
    ' Default to all buttons enabled if there are records in the Lookup recordset; Otherwise, disable them all.
    bHaveRecords = (mtWrapper.LookupRecordCount <> 0)
    For Each cmd In cmdNavigate
        cmd.Enabled = bHaveRecords
    Next

    ' Now selectively disable if our current record position causes certain navigation to be unavailable/illogical.
    If bHaveRecords Then
        If mtWrapper.CurrentLookupRecordNumber = 1 Then
            cmdNavigate(navFirst).Enabled = False
            cmdNavigate(navPrev).Enabled = False
        End If

        If mtWrapper.CurrentLookupRecordNumber = mtWrapper.LookupRecordCount Then
            cmdNavigate(navNext).Enabled = False
            cmdNavigate(navLast).Enabled = False
        End If
    End If
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub fnSetPropertiesForPayeeScreen(bSendEmptyName As Boolean)
    '----------------------------------------------------------------------------
    ' Procedure :  Sub fnSetPropertiesForPayeeScreen
    ' Created by:  BAW on 04-26-2001 08:55
    '
    ' Comments  : Sets member variables so they can be accessed from/by Payee screen
    ' Called by : msgPayees_DblClick and cmdAddPayee_Click
    ' Parameters: N/A
    '
    ' Modified  :
    '----------------------------------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnSetPropertiesForPayeeScreen"

    ' Note: If there are no Payees, msgPayees.Row will be set to 0 (the header row)
    msgPayees.Col = 1   ' Payee Name column (2nd column, current row)

    If bSendEmptyName Then
        InsuredCurrentPayeeName = vbNullString
    Else
        InsuredCurrentPayeeName = msgPayees.Text
    End If

    InsuredClmNum = txtClaimNumber
    InsuredClmInsdDthDt = dtpDateOfDeath.value
    InsuredClmProofDt = dtpDateOfProof.value
    InsuredLobCd = mstrLOB
    InsuredState = cboStCd
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Function fnValidData() As Boolean
    ' Comments  : Determines if all data is valid, including
    '             whether all required fields have been input.
    '             This function is called by cmdUpdate_Click.
    '             If a data error is found, it returns False
    '             which directs the caller to stop processing.
    '             It also generates warnings, by calling
    '             WarningData(), but only if no errors were
    '             found up to that point.
    ' Parameters: N/A
    ' Returns   : True if all data is valid; False otherwise
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnValidData"
    Dim bErrorFound As Boolean
    Dim ctlFirstToFail As Control
    Dim intFailures As Integer
    Dim strFieldList As String
    Dim strMsgText As String
    Dim strSQL As String
    Dim rstTempPayees As ADODB.Recordset
    Dim intLength As Integer

    fnValidData = True

    ' Check the fields in a left-to-right, top-to-bottom screen sequence.
    '     1. txtClaimNumber      5. cboStCd
    '     2. cboSystem           6. dtpDatefDeath
    '     3. cboIruleCd          7. dtpDatefProof
    '     4. txtInsured          8. txtStrlIntRuleAmt

    ' ------------- First, verify required fields are missing --------------
    '   If the column definition in the Access table has "Required=Yes" or
    '   is a Key field, or if the business area has indicated a field
    '   is required, then include processing for that column here.

    ' The only time the following might ever get hit is when mbInAddMode=True
    ' since that is the only time the Claim Number field is editable.
    If IsNull(txtClaimNumber) Or txtClaimNumber = vbNullString Then
        If intFailures = 0 Then
            strFieldList = vbCrLf & mcstrTxtClaimNumberLabel
            Set ctlFirstToFail = txtClaimNumber
        Else
            strFieldList = strFieldList & vbCrLf & mcstrTxtClaimNumberLabel
        End If
        intFailures = intFailures + 1
    End If

    If IsNull(cboSystem) Or cboSystem = vbNullString Then
        If intFailures = 0 Then
            strFieldList = vbCrLf & mcstrCboSystemLabel
            Set ctlFirstToFail = cboSystem
        Else
            strFieldList = strFieldList & vbCrLf & mcstrCboSystemLabel
        End If
        intFailures = intFailures + 1
    End If

    If IsNull(cboIruleCd) Or cboIruleCd = vbNullString Then
        If intFailures = 0 Then
            strFieldList = vbCrLf & mcstrCboCompanyLabel
            Set ctlFirstToFail = cboIruleCd
        Else
            strFieldList = strFieldList & vbCrLf & mcstrCboCompanyLabel
        End If
        intFailures = intFailures + 1
    End If

    If IsNull(txtInsured) Or txtInsured = vbNullString Then
        If intFailures = 0 Then
            strFieldList = vbCrLf & mcstrTxtInsuredLabel
            Set ctlFirstToFail = txtInsured
        Else
            strFieldList = strFieldList & vbCrLf & mcstrTxtInsuredLabel
        End If
        intFailures = intFailures + 1
    End If

    ' Check for empty values or a space value (the latter is the default value which must be
    ' changed by the user)
    If IsNull(cboStCd) Or cboStCd = vbNullString Or cboStCd = " " Then
        If intFailures = 0 Then
            strFieldList = vbCrLf & mcstrcboStCdLabel
            Set ctlFirstToFail = cboStCd
        Else
            strFieldList = strFieldList & vbCrLf & mcstrcboStCdLabel
        End If
        intFailures = intFailures + 1
    End If

    If IsNull(txtStrlIntRuleAmt) Or txtStrlIntRuleAmt = vbNullString Then
        If intFailures = 0 Then
            strFieldList = vbCrLf & mcstrTxtClerkCodeLabel
            Set ctlFirstToFail = txtStrlIntRuleAmt
        Else
            strFieldList = strFieldList & vbCrLf & mcstrTxtClerkCodeLabel
        End If
        intFailures = intFailures + 1
    End If

    If intFailures <> 0 Then
        bErrorFound = True
        fnValidData = False
        If ctlFirstToFail.Visible Then
            ctlFirstToFail.SetFocus
        End If
        gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_REQD_FIELDS_MISSING, _
                                   mstrScreenName & gcstrDOT & cstrCurrentProc, _
                                   strFieldList
        GoTo PROC_EXIT
    End If



    ' ------------------- Now, do cross-field validations --------------------


    intFailures = 0     ' Reset for this section of error validations

    ' If the 5th position of the Claim Number = G (Group), then the
    ' length must be 15.
    intLength = Len(txtClaimNumber)
    If InStr(1, txtClaimNumber, "G", vbTextCompare) = 6 Then
        If intLength <> 15 Then
            intFailures = intFailures + 1
            Set ctlFirstToFail = txtClaimNumber
            strMsgText = strMsgText & vbCrLf & _
                         "For Group, the Claim Number must be 15 characters long."
        End If
    ' If it's an Individual claim number, then it must between 7 and 9 characters long.
    ElseIf intLength < 7 Or intLength > 9 Then
        intFailures = intFailures + 1
        Set ctlFirstToFail = txtClaimNumber
        strMsgText = strMsgText & vbCrLf & _
                    "For Individual, the Claim Number must be 7, 8 or 9 characters long."
    End If

    ' Disallow a future-dated Date of Death
    If DateValue(dtpDateOfDeath.value) > Date Then
        intFailures = intFailures + 1
        Set ctlFirstToFail = dtpDateOfDeath
        strMsgText = strMsgText & vbCrLf & _
                     "The " & mcstrDtpDateOfDeathLabel & " (" & dtpDateOfDeath.value & _
                     ") cannot be in the future."
    End If

    ' Disallow a future-dated Date of Proof
    If DateValue(dtpDateOfProof.value) > Date Then
        intFailures = intFailures + 1
        Set ctlFirstToFail = dtpDateOfProof
        strMsgText = strMsgText & vbCrLf & _
                     "The " & mcstrDtpDateOfProofLabel & " (" & dtpDateOfProof.value & _
                     ") cannot be in the future."
    End If

    If DateValue(dtpDateOfProof.value) < DateValue(dtpDateOfDeath.value) Then
        intFailures = intFailures + 1
        Set ctlFirstToFail = dtpDateOfProof
        strMsgText = strMsgText & vbCrLf & _
                     "The " & mcstrDtpDateOfProofLabel & " (" & dtpDateOfProof.value & _
                     ") must be on or after the " & mcstrDtpDateOfDeathLabel & " (" & _
                     dtpDateOfDeath.value & ")."
    End If

    ' Determine whether any Payees exist with a Date Of Payment earlier than the
    ' Insured's Date of PROOF. The date value must be surrounded by '#' so it's
    ' correctly interpreted in the SQL command.
    ' Note that this query should return 1 row, whether or not there
    ' actually ARE any Payees that match the criteria.
    Set rstTempPayees = New ADODB.Recordset
    strSQL = "Select COUNT([ClaimNumber]) As CntOfBadPayees FROM [PAYEE] " & _
             "WHERE ([ClaimNumber] = '" & txtClaimNumber & "' ) " & _
             "AND ([Date of Payment] < #" & dtpDateOfProof.value & "#)"
    With rstTempPayees
        ' CursorType=adOpenStatic   - Scrolling fwd/bwd permitted, add/chg/del by other users not visible
        ' LockType=adLockReadOnly   - Read-only; Modifications are not permitted
        .Open Source:=strSQL, ActiveConnection:=gconAppActive, _
                    CursorType:=adOpenStatic, LockType:=adLockReadOnly, Options:=adCmdText
        If .RecordCount > 0 Then
            If !CntOfBadPayees > 0 Then
                intFailures = intFailures + 1
                Set ctlFirstToFail = dtpDateOfProof
                strMsgText = strMsgText & vbCrLf & _
                             "One or more Payees exist with a Date Of Payment " & _
                             "earlier than the " & mcstrDtpDateOfProofLabel & "."
            End If
        End If
    End With


    ' Determine whether any Payees exist with a Date Of Payment earlier than the
    ' Insured's Date of DEATH. The date value must be surrounded by '#' so it's
    ' correctly interpreted in the SQL command.
    ' Note that this query should return 1 row, whether or not there
    ' actually ARE any Payees that match the criteria.

    ' Close recordset so it can be reopened with different properties, e.g., Source:=strSQL
    rstTempPayees.Close
    strSQL = "Select COUNT([ClaimNumber]) As CntOfBadPayees FROM [PAYEE] " & _
             "WHERE ([ClaimNumber] = '" & txtClaimNumber & "' ) " & _
             "AND ([Date of Payment] < #" & dtpDateOfDeath.value & "#)"
    With rstTempPayees
        ' CursorType=adOpenStatic   - Scrolling fwd/bwd permitted, add/chg/del by other users not visible
        ' LockType=adLockReadOnly   - Read-only; Modifications are not permitted
        .Open Source:=strSQL, ActiveConnection:=gconAppActive, _
                    CursorType:=adOpenStatic, LockType:=adLockReadOnly, Options:=adCmdText
        If .RecordCount > 0 Then
            If !CntOfBadPayees > 0 Then
                intFailures = intFailures + 1
                Set ctlFirstToFail = dtpDateOfDeath
                strMsgText = strMsgText & vbCrLf & _
                             "One or more Payees exist with a Date Of Payment " & _
                             "earlier than the " & mcstrDtpDateOfDeathLabel & "."
            End If
        End If
    End With

    If intFailures <> 0 Then
        bErrorFound = True
        fnValidData = False
        If ctlFirstToFail.Visible Then
            ctlFirstToFail.SetFocus
        End If
        gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_CROSS_FLD_VALIDATIONS_FAILED, _
                                   mstrScreenName & gcstrDOT & cstrCurrentProc, _
                                   "this record can be updated", strMsgText
        GoTo PROC_EXIT
    End If

    ' If no errors found, continue with checking for warnings
    If Not bErrorFound Then
        fnWarningData
    End If
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0

    ' Clean-up statements go here
    fnFreeRecordset rstTempPayees
    fnFreeObject ctlFirstToFail

    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub fnWarningData()
    ' Comments  : Validates fields, generating warnings if appropriate.
    '             It should NOT cause ValidData (this procedure's caller)
    '             to return False, since we want updates to proceed.
    ' Parameters: N/A
    ' Returns   : N/A
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnWarningData"

    ' Add warnings, if necessary

PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
    ' Comments  :
    ' Parameters:  -
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "Form_Activate"

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    ' Since this form is hidden which the Payee form is visible, clicking on the Payee
    ' form can trigger the frmInsured's Form_Activate event. Therefore, the bulk
    ' of the processing in this event is conditioned on whether it (frmInsured)
    ' is visible or not. If not visible, we don't want to mess up the Payee-related
    ' values that could mess up the processing in the Payee form.
    If Me.Visible Then
        fnSetFocusToFirstUpdateableField
    End If
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
        Case 5      ' Invalid procedure call or argument
            ' Caused by setting focus to a field that's not yet visible
            Resume Next
        Case Else
            ' Save Err object data, if not already saved
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub Form_Initialize()
    ' Comments  : Sets up ADODB.Recordsets used throughout the form
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "Form_Initialize"
    Dim strSQL As String

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    ' ----- Build the recordset of Insured data for the Lookup box (mrstLookup) -----
    Set mrstLookup = New ADODB.Recordset
    strSQL = "SELECT [ClaimNumber] FROM [Policy Information] ORDER BY [ClaimNumber]"
    ' CursorType=adOpenKeyset   - Scrolling fwd/bwd permitted, chgs/del by other users visible
    ' LockType=adLockOptimistic - Modifications to data are cached until UpdateBatch method called
    mrstLookup.Open Source:=strSQL, _
                    ActiveConnection:=gconAppActive, _
                    CursorType:=adOpenKeyset, _
                    LockType:=adLockOptimistic, _
                    Options:=adCmdText

    ' Allow the progress meter on the splash screen to get updated
    If fnIsFormLoaded("frmSplash") Then
        DoEvents
    End If
    
    ' ------------- Build the recordset of Insured data (mrstInsureds) -------------
    Set mrstInsureds = New ADODB.Recordset
    strSQL = "SELECT [ClaimNumber], [SYSTEM], " & _
        "[Company], [Insured], [State], [DATE OF death] As DateOfDeath, " & _
        "[Date of Proof] As DateOfProof, [Clerk] As ClerkCode, " & _
        "[Deposit Amount] As TotalPayments, [Withholding (Interest)] As TotalWithheld, " & _
        "[TOTAL] As TotalWithInterest, [INTEREST] As TotalInterest " & _
        "FROM [Policy Information] ORDER BY [ClaimNumber]"
    ' CursorType=adOpenKeyset   - Scrolling fwd/bwd permitted, chgs/del by other users visible
    ' LockType=adLockOptimistic - Modifications to data are cached until UpdateBatch method called
    mrstInsureds.Open Source:=strSQL, _
                      ActiveConnection:=gconAppActive, _
                      CursorType:=adOpenKeyset, _
                      LockType:=adLockOptimistic, _
                      Options:=adCmdText


    ' ------------- Initialize the Recordset Object for Payee data (mrstPayees) -------------
    '    It will be populated in fnGetChildren and loaded to the MSFlexGrid
    '    control in fnFillPayeeGrid( ). (The latter is called by fnGetChildren().)
    Set mrstPayees = New ADODB.Recordset
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
        Case 5      ' Invalid procedure call or argument
            ' Caused by setting focus to a field that's not yet visible
            Resume Next
        Case Else
            ' Save Err object data, if not already saved
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    ' Comments  :
    ' Parameters:  -
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "Form_Load"

    ' Set the screen name that will be used to form the Title on message boxes
    mstrScreenName = Me.Caption

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    ' Identify the icons that will be used for the form and the picture next to the Lookup ComboBox
    Me.Icon = LoadResPicture(gcRES_ICON_MAINAPP, vbResIcon)

    ' If the user has ever opened this form before, restore its size & placement.
    ' If the restore would result in the form being off-screen, just center it instead.
    If gapsApp.RestoreForm(Me) = False Then
        With Me
            .Width = mclngMinFormWidth
            .Height = mclngMinFormHeight
        End With
        fnCenterFormOnMDI frmMDIMain, Me
    End If

    ' Set our ComboBox (FlexGrid) control setting
    With vfgLookup
       .Rows = 1
       .Cols = 2    ' FUND_CD, FUND_NM
       .FixedRows = 0
       .FixedCols = 0

       .ScrollBars = flexScrollBarNone
       .GridLines = flexGridNone
       .ColKey(0) = "KeyValues"
       .ColWidth(1) = 0
       .FocusRect = flexFocusLight
       .ShowComboButton = flexSBAlways
       .ColWidth(.ColIndex("KeyValues")) = .Width - 50
       .RowHeight(0) = .Height - 50
       .Select 0, 0
    End With

    ' Set the control to receive the focus after errors (the first editable field
    ' on the screen), dependent upon whether we're in Add Mode or not. If in Add mode,
    ' this control would typically be the first control that corresponds to a Key field.
    ' If not in Add mode, this control would typically be the topmost/leftmost
    ' "always updateable" control on the screen (excepting the Lookup ComboBox).
    Set mctlFirstUpdateableField_Add = cboStCd
    Set mctlFirstUpdateableField_Upd = dtpStrlEndDt

    ' Instantiate and initialize a table wrapper object for the appropriate table(s).
    Set mtWrapper = New ctsrlStateRule

    ' Set the control to receive the focus after errors (the first editable field
    ' on the screen)
    Set mctlFirstEditableField = cboLookup

    ' Make editable alphanumeric textboxes be forced to uppercase
    'fnSetTextBoxCase txtClaimNumber, gcintForceUpperCase
    'fnSetTextBoxCase txtInsured, gcintForceUpperCase
    'fnSetTextBoxCase txtStrlIntRuleAmt, gcintForceUpperCase

    ' Set on-screen label that shows "Record x of y"
    lblRecordPosition = "Record 1 of 50"   'TEMP fnShowRecordPosition(mrstInsureds)


    ' Populate combo boxes, using mrst___ objects established
    ' in Form_Initialize
    ' Note: The MSFlexGrid containing Payee data is populated in fnGetChildren( ),
    '       via its call to fnFillPayeeGrid( ).
    fnLoadCboLookup
    fnLoadCboStCd
    fnLoadCboLobCd
    fnLoadCboIruleCd
    fnLoadCboReqdIdtypCd
    fnLoadCboCalcIdtypCd
    
    ' Allow the progress meter on the splash screen to get updated
    If fnIsFormLoaded("frmSplash") Then
        DoEvents
    End If

    If mrstInsureds.RecordCount > 0 Then
        mrstInsureds.MoveFirst
        fnLoadControls
        fnGetChildren
        If mrstInsureds.RecordCount > 1 Then
            fnSetNavigationButtons True
        Else
            ' There is only 1 record in the recordset
            fnSetNavigationButtons False
        End If
        fnSetCommandButtons True
        fnEnableKeyFields False
    Else
        fnAddRecord
    End If

    mbInLookupMode = False
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



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub Form_QueryUnload(ByRef pintCancel As Integer, ByRef pintUnloadMode As Integer)
    ' Comments  :
    ' Parameters:
    '    pintCancel     (in/out) - if set to True, refuses to honor the unload request.
    '    pintUnloadMode (in/out) - Identifies what triggered the unload request
    '
    ' --------------------------------------------------------------------------------------------
    Dim intButtonClicked                As Integer
    Const cstrCurrentProc               As String = "Form_QueryUnload"
    On Error GoTo PROC_ERR

    ' Set screen name in case errors are reported here or in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    If gbAmProcessingAnAppFatalError Then
        ' ALWAYS let the form be unloaded, with no prompts to the user, if shutting
        ' down the app due to an application fatal error having been hit.
        GoTo PROC_EXIT
    End If
    
    If (Not mbInAddMode) And (Not IsDirty) Then
        ' Let the form be closed if the user is in neither Add nor Update mode.
        GoTo PROC_EXIT
    End If

    ' Since Update (IsDirty) mode can be True while in Add mode, we must check for Add mode first.
    ' Otherwise, Adds where the user has started typing (thus setting IsDirty to True) will be
    ' treated like an Update, when it should be treated like an Add.
    If mbInAddMode Then
        intButtonClicked = gerhApp.ReportNonFatal(vbObjectError + gcRES_ALRT_CHANGES_PENDING, _
                           mstrScreenName & gcstrDOT & cstrCurrentProc)
        If intButtonClicked = vbYes Then
            ' If they want to abandon an Add before they started data entry, let them!
            ' Redisplay the form with the *first* record now showing
            mtWrapper.GetLookupData
            If mtWrapper.LookupIsAtBOF And mtWrapper.LookupIsAtEOF Then
                ' There are no records in the table, so let the form close (If we went into Add
                ' mode, the user would never be able to exit the screen!)
            Else
                pintCancel = True
                mtWrapper.GoToFirstRecord
                '!TODO!: Have to code for the situation where the user is abandoning the
                '        Add of the table's first record...e.g., go into Add mode.
                ' Load current record's properties to form's controls, reset
                ' navigation buttons and set "rec x of y" label
                fnLoadControls
                mbInAddMode = False
                fnSetCommandButtons True
                ' This **must** be done as the user leaves Add mode, so that the key fields
                ' will now be protected to prevent the user from being able to edit them.
                ' Editing a key field is allowed only when in Add mode.
                fnSetAvailabilityOfKeyFields
            End If
            mbInLookupMode = False
        Else
            ' User doesn't want to abandon the Add that's still in progress, so ignore the request
            ' to close the form and redisplay the form with the same data and with the user's Add
            ' still in progress.
            pintCancel = True
        End If
    Else    ' IsDirty (a.k.a. in Update mode)
        intButtonClicked = gerhApp.ReportNonFatal(vbObjectError + gcRES_ALRT_CHANGES_PENDING, _
                           mstrScreenName & gcstrDOT & cstrCurrentProc)
        If intButtonClicked = vbYes Then
            ' Abandon their pending changes and redisplay the same record as it *now* appears in
            ' the database
            pintCancel = True
            mtWrapper.GetRelativeRecord mtWrapper.FundCd, epdSameRecord
            '!TODO!: Have to code for the situation where another user deleted the record whose
            '        edits *this* user is abandoning....e.g., go into Add mode
            fnLoadControls
            IsDirty = False
            fnSetCommandButtons True
        Else
            ' User wants to keep pending changes, so ignore the request to close the form and redisplay
            ' the form with the same record showing and with the user's pending changes still pending.
            pintCancel = True
        End If
    End If
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



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub Form_Unload(ByRef pintCancel As Integer)
    ' Comments  : Close the form
    ' Parameters: pintCancel (in/out), if set to True
    '             the unload is aborted
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "Form_Unload"

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    gapsApp.SaveForm Me

    IsDirty = False

    fnFreeRecordset mrstLookup
    fnFreeRecordset mrstPayees
    fnFreeRecordset mrstInsureds
    fnFreeObject mtWrapper
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




'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub txtStrlIntRuleAmt_Change()
    ' Comments  :
    ' Parameters:  -
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "txtClerkCode_Change"

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    fnInitializeEditMode
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



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub txtStrlIntRuleAmt_GotFocus()
    ' Comments  : Select control's contents to facilitate editing
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "txtClerkCode_GotFocus"

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    fnHighlightText txtStrlIntRuleAmt
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


'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub vfgLookup_ChangeEdit()
    ' Comments  : New Lookup functionality.  Implementation of VS Flex Grid
    ' Parameters: N/A
    ' Modified  : CMP 4/27/2002
    '
    ' --------------------------------------------------
    Const cstrCurrentProc               As String = "vfgLookup_ChangeEdit"
    Dim hrgHourglass                    As chrgHourglass
    Dim strRecordKeytoRetrieve          As String

    On Error GoTo PROC_ERR

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    vfgLookup.SetFocus
    strRecordKeytoRetrieve = vfgLookup.EditText

    ' Turn on hourglass, in case the lookup is slow
    Set hrgHourglass = New chrgHourglass
    hrgHourglass.value = True

    ' If there are no records in the main table maintained by this form,
    ' if the blank entry was selected, or if the user typed in nothing
    ' (i.e. a blank entry in the Lookup box), then skip further processing.
    ' There's nothing to do a lookup on!
    ' If the LookupRecordCount = 0 then we should already be in Add mode
    ' and thus should just stay as we are.
    If (mtWrapper.LookupRecordCount = 0) Or _
        (vfgLookup.EditText = gcstrBlankEntry) Then
            GoTo PROC_EXIT
    End If

    ' If the user is in Add mode, interpret a lookup request to mean they want
    ' to exit Add mode and lose any outstanding changes. Retrieve and display
    ' the first record in the table.
    If mbInAddMode Then
        mbInAddMode = False
        mtWrapper.GoToFirstRecord
        fnSetAvailabilityOfKeyFields
        fnLoadControls
        fnSetCommandButtons True
    Else
        ' If the user has selected something, retrieve the appropriate record
        ' and update the table wrapper's properties accordingly.
        mtWrapper.GetSingleRecord strKey1:=strRecordKeytoRetrieve, bSynchLookupRST:=True
        Me.Refresh
        ' A Lookup request (or navigation request, for that matter) is interpreted
        ' to mean the user wants to discard pending changes, if any, so turn off
        ' the IsDirty flag.
        IsDirty = False
        ' Load current record's properties to form's controls, reset navigation buttons
        ' and set "rec x of y" label
        fnLoadControls
        fnSetCommandButtons True
        Me.Refresh
    End If


    ' Set the Lookup control's displayed selected text to a null string, so the
    ' user doesn't get confused. Without this code, the Lookup box continues to display
    ' the value last selected for lookup purposes, even when the user has since positioned
    ' to a different record by virtue of doing a Delete or Add or using the navigation buttons.
    ' NOTE: Can't do "vfgLookup.Select 0,0" without "breaking" the user's ability to do
    '       typeahead searches with the keyboard!
    vfgLookup.Text = vbNullString
    vfgLookup.TextMatrix(0, 0) = vbNullString
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
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub vfgLookup_GotFocus()
    '-----------------------------------------------------------------------------
    ' Function     vfgLookup_GotFocus
    ' Purpose      Turn on Lookup Mode now that the user has entered this control.
    ' Parameters   N/A
    ' Returns      N/A
    ' Date:        12/19/2001
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc          As String = "vfgLookup_GotFocus"
    On Error GoTo PROC_ERR

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    ' CMP - FORCE the drop down to occur.
    vfgLookup.Select 0, 0
    SendKeys "{ENTER}"

    mbInLookupMode = True
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



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub vfgLookup_LostFocus()
    '-----------------------------------------------------------------------------
    ' Function     vfgLookup_LostFocus
    ' Purpose      Turn off Lookup Mode now that the user has left that control.
    ' Parameters   N/A
    ' Returns      N/A
    ' Date:        12/19/2001
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc          As String = "vfgLookup_LostFocus"
    On Error GoTo PROC_ERR

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    ' Select the "dummy" column.
    vfgLookup.Select 0, 1

    ' Set the Lookup control's displayed selected text to a null string, so the
    ' user doesn't get confused. Without this code, the Lookup box continues to display
    ' the value last selected for lookup purposes, even when the user has since positioned
    ' to a different record by virtue of doing a Delete or Add or using the navigation buttons.
    ' NOTE: Can't do "vfgLookup.Select 0,0" without "breaking" the user's ability to do
    '       typeahead searches with the keyboard!
    vfgLookup.Text = vbNullString
    vfgLookup.TextMatrix(0, 0) = vbNullString
    vfgLookup.Refresh

    mbInLookupMode = False
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
