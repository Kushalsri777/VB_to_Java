VERSION 5.00
Begin VB.Form frmMsgBox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Caption Is Set Programmatically"
   ClientHeight    =   4380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMsgText 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   1365
      Left            =   75
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   585
      Width           =   7365
   End
   Begin VB.Frame fraDetails 
      Caption         =   "Details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1290
      Left            =   75
      TabIndex        =   5
      Top             =   2580
      Width           =   7365
      Begin VB.Label lblContextLabel 
         Caption         =   "Context:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   300
         TabIndex        =   10
         Top             =   615
         Width           =   625
      End
      Begin VB.Label lblCodeLabel 
         Caption         =   "Code:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   300
         TabIndex        =   9
         Top             =   330
         Width           =   625
      End
      Begin VB.Label lblErrorContext 
         Caption         =   "Error Context is set programmatically"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   1005
         TabIndex        =   7
         Top             =   615
         Width           =   6240
      End
      Begin VB.Label lblErrorCode 
         Caption         =   "Error Code is set programmatically"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1005
         TabIndex        =   6
         Top             =   330
         Width           =   6165
      End
   End
   Begin VB.Frame fraMainButtons 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2490
      TabIndex        =   0
      Top             =   2025
      Width           =   2610
      Begin VB.CommandButton cmdButton2 
         Caption         =   "Button2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1335
         TabIndex        =   2
         Top             =   45
         Width           =   1215
      End
      Begin VB.CommandButton cmdButton1 
         Caption         =   "Button1"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   60
         TabIndex        =   1
         Top             =   45
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdDetailsCollapse 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7050
      TabIndex        =   4
      Top             =   3930
      Width           =   390
   End
   Begin VB.CommandButton cmdDetailsExpand 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7050
      TabIndex        =   3
      Top             =   2070
      Width           =   390
   End
   Begin VB.Image imgIcon 
      Height          =   520
      Left            =   75
      Top             =   30
      Width           =   615
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
' Module     : frmMsgBox
' Description: This form is opened when cerhErrorHandler class needs to report a
'              message of any type (info, warning, alert, or error)
' Procedures :
'    Public    Property Get ButtonClicked() As Long
'    Public    Property Let ErrorCode(ByVal lngValue As Long)
'    Public    Property Let ErrorContext(ByVal strValue As String)
'    Public    Property Let MsgText(ByVal strValue As String)
'    Public    Property Let ScreenName(ByVal strValue As String)
'    Private   cmdButton1_Click()
'    Private   cmdButton2_Click()
'    Private   cmdDetailsCollapse_Click()
'    Private   cmdDetailsExpand_Click()
'    Private   fnSetButtons(ByVal lngValue As eMsgButtons)
'    Private   fnSetupForDefaultDisplay()
'    Private   Form_Load()
'    Private   Form_Unload(ByRef pintCancel As Integer)

' Modified   :
' 03/19/02 BAW (Phase2A) Created form
'
' --------------------------------------------------
Option Explicit
Option Compare Binary
Private Const mclngMinFormWidth As Long = 7740
Private Const mclngMinFormHeight_Collapsed As Long = 2985
Private Const mclngMinFormHeight_Expanded As Long = 4815
Private Const mcstrOK     As String = "OK"
Private Const mcstrCANCEL As String = "Cancel"
Private Const mcstrYES    As String = "Yes"
Private Const mcstrNO     As String = "No"


' Private variables to hold public properties
Private m_lngButtonClicked As Long
Private m_strMsgText       As String
Private m_lngErrorCode     As String
Private m_strErrorContext  As String
Private m_strScreenName    As String

Public Enum eMsgButtons
    ebtnYesNo = 0
    ebtnOKCancel = 1
    ebtnOKOnly = 2
End Enum


'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'\                                                                      /
'|                         Public Properties                            |
'/                                                                      \
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
Public Property Get ButtonClicked() As Long
    ' Returns the ButtonClicked property setting
    ButtonClicked = m_lngButtonClicked
End Property

Public Property Let ErrorCode(ByVal lngValue As Long)
    ' Sets the ErrorCode property to the value specified by lngValue.
    ' Note that lngValue is a *translated* error code. For an
    ' application-specific error, this means that vbObjectError
    ' has been subtracted from it already.

    m_lngErrorCode = lngValue
    lblErrorCode.Caption = CStr(lngValue)

    Select Case lngValue
        Case gcRES_INFO_START To gcRES_INFO_END
            fnSetButtons ebtnOKOnly
            Me.Caption = m_strScreenName & " - Information"
            imgIcon.Picture = LoadResPicture(gcRES_ICON_INFO, vbResIcon)
        Case gcRES_WARN_START To gcRES_WARN_END
            fnSetButtons ebtnOKOnly
            Me.Caption = m_strScreenName & " - Warning"
            imgIcon.Picture = LoadResPicture(gcRES_ICON_WARN, vbResIcon)
        Case gcRES_ALRT_START To gcRES_ALRT_END
            fnSetButtons ebtnYesNo
            Me.Caption = m_strScreenName & " - Alert"
            imgIcon.Picture = LoadResPicture(gcRES_ICON_ALRT, vbResIcon)
        Case gcRES_NERR_START To gcRES_NERR_END
            fnSetButtons ebtnOKOnly
            Me.Caption = m_strScreenName & " - Error"
            imgIcon.Picture = LoadResPicture(gcRES_ICON_ERR, vbResIcon)
        Case Else
            ' Intended to be 9000 to 9999 ... or a VB or ADO error code
            fnSetButtons ebtnOKOnly
            Me.Caption = m_strScreenName & " - Fatal Error"
            imgIcon.Picture = LoadResPicture(gcRES_ICON_ERR, vbResIcon)
    End Select
End Property

Public Property Let ErrorContext(ByVal strValue As String)
    ' Sets the ErrorContext property to the value specified by strValue
    m_strErrorContext = strValue
    lblErrorContext.Caption = strValue
End Property

Public Property Let MsgText(ByVal strValue As String)
    ' Sets the MsgText property to the value specified by strValue
    m_strMsgText = strValue
    txtMsgText.Text = strValue
End Property

Public Property Let ScreenName(ByVal strValue As String)
    ' Sets the ScreenName property to the value specified by strValue
    m_strScreenName = strValue
End Property



'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'\                                                                      /
'|                           Public Methods                             |
'/                                                                      \
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'None


'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'\                                                                      /
'|                            Private Methods                           |
'/                                                                      \
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\

'/////////////////////////////////////////////////////////////////////////////////
Private Sub cmdButton1_Click()
    Select Case cmdButton1.Caption
        Case mcstrOK
            m_lngButtonClicked = vbOK
        Case mcstrYES
            m_lngButtonClicked = vbYes
        Case mcstrCANCEL
            m_lngButtonClicked = vbCancel
        Case Else           ' mcstrNO
            m_lngButtonClicked = vbNo
    End Select

    Me.Hide
End Sub



'/////////////////////////////////////////////////////////////////////////////////
Private Sub cmdButton2_Click()
    '--------------------------------------------------------------------
    ' Procedure : cmdButton2_Click
    ' Comments  : Set form property that the caller can interrogate, if desired.
    '             Note: the caller should unload this form after that interrogation.
    '
    ' Called by : N/A
    ' Parameters: N/A
    ' Modified  :
    '--------------------------------------------------------------------
    Select Case cmdButton2.Caption
        Case mcstrOK
            m_lngButtonClicked = vbOK
        Case mcstrYES
            m_lngButtonClicked = vbYes
        Case mcstrCANCEL
            m_lngButtonClicked = vbCancel
        Case Else           ' mcstrNO
            m_lngButtonClicked = vbNo
    End Select

    Me.Hide
End Sub



'/////////////////////////////////////////////////////////////////////////////////
Private Sub cmdDetailsCollapse_Click()
    '--------------------------------------------------------------------
    ' Procedure : cmdDetailsCollapse_Click
    ' Comments  : Set form and control properties based on the user
    '             indicating they want a "normal" display
    '
    ' Called by : N/A
    ' Parameters: N/A
    ' Modified  :
    '--------------------------------------------------------------------
    fnSetupForDefaultDisplay
End Sub



'/////////////////////////////////////////////////////////////////////////////////
Private Sub cmdDetailsExpand_Click()
    '--------------------------------------------------------------------
    ' Procedure : cmdDetailsExpand_Click
    ' Comments  : Set form and control properties based on the user
    '             indicating they want a "Details" display
    '
    ' Called by : N/A
    ' Parameters: N/A
    ' Modified  :
    '--------------------------------------------------------------------
    fnWindowLock Me.hWnd
    Me.Height = mclngMinFormHeight_Expanded
    Me.Width = mclngMinFormWidth

    fraDetails.Visible = True
    cmdDetailsExpand.Visible = False
    cmdDetailsExpand.Enabled = False
    cmdDetailsCollapse.Visible = True
    cmdDetailsCollapse.Enabled = True
    fnWindowUnlock
End Sub



'/////////////////////////////////////////////////////////////////////////////////
Public Sub fnSetButtons(ByVal lngValue As eMsgButtons)
    ' Sets the placement and caption of the form's main command buttons.
    Const lngButtonSpacer As Long = 60

    'fnWindowLock Me.hWnd
    Select Case lngValue
        Case ebtnYesNo
            With cmdButton2
                .Visible = True
                .Default = False
                .Caption = mcstrNO
                .Left = fraMainButtons.Width - .Width - lngButtonSpacer
            End With
            With cmdButton1
                .Visible = True
                .Caption = mcstrYES
                .Default = True
                .Left = lngButtonSpacer
            End With

        Case ebtnOKCancel
            With cmdButton2
                .Visible = True
                .Default = False
                .Caption = mcstrCANCEL
                .Left = fraMainButtons.Left + fraMainButtons.Width - .Width - lngButtonSpacer
            End With
            With cmdButton1
                .Visible = True
                .Caption = mcstrOK
                .Default = True
                .Left = lngButtonSpacer
            End With

        Case Else           ' ebtnOKOnly
            With cmdButton2
                .Visible = False
                .Default = False
                .Caption = mcstrOK
                .Left = fraMainButtons.Left + fraMainButtons.Width - .Width - lngButtonSpacer
            End With
            With cmdButton1
                .Visible = True
                .Default = True
                .Caption = mcstrOK
                .Left = (fraMainButtons.Left - .Width) / 2
            End With
    End Select
    'fnWindowUnlock
End Sub



'/////////////////////////////////////////////////////////////////////////////////
Private Sub fnSetupForDefaultDisplay()
    '--------------------------------------------------------------------
    ' Procedure : fnSetupForDefaultDisplay
    ' Comments  : Set default form and control properties
    '
    ' Called by : cmdDetailCollapse_Click( )
    '             Form_Load( )
    ' Parameters: N/A
    ' Modified  :
    '--------------------------------------------------------------------
    fnWindowLock Me.hWnd
    Me.Height = mclngMinFormHeight_Collapsed
    Me.Width = mclngMinFormWidth

    fraDetails.Visible = True
    cmdDetailsExpand.Visible = True
    cmdDetailsExpand.Enabled = True
    cmdDetailsCollapse.Visible = False
    cmdDetailsCollapse.Enabled = False
    fnWindowUnlock
End Sub



'/////////////////////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    '--------------------------------------------------------------------
    ' Procedure : Form_Load
    ' Comments  : Loads the form
    '
    ' Called by : N/A
    ' Parameters: N/A
    ' Modified  :
    '--------------------------------------------------------------------
    Me.Icon = LoadResPicture(gcRES_ICON_MAINAPP, vbResIcon)

    ' Make sure Cursor reverts back to normal, in case it was left in an hourglass
    Screen.MousePointer = vbDefault

    fnSetupForDefaultDisplay
    ' Don't center on MDI, since MDI may not have been loaded yet if an
    ' error was generated during app startup.
End Sub



'/////////////////////////////////////////////////////////////////////////////////
Private Sub Form_Unload(Cancel As Integer)
    '--------------------------------------------------------------------
    ' Procedure : Form_Unload
    ' Comments  : Unloads the form
    '
    ' Called by : N/A
    ' Parameters: N/A
    ' Modified  :
    '--------------------------------------------------------------------
    Unload Me
    fnFreeObject frmMsgBox 
End Sub
