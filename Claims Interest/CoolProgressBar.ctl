VERSION 5.00
Begin VB.UserControl CoolProgressBar 
   ClientHeight    =   240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   ScaleHeight     =   16
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   256
   ToolboxBitmap   =   "CoolProgressBar.ctx":0000
   Begin VB.PictureBox picBackbuffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   3840
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3840
   End
End
Attribute VB_Name = "CoolProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Module     : CoolProgressBar
' Description:
' Procedures : Refresh()
'              Get BackColor()
'              Let BackColor(ByVal pNew_BackColor As OLE_COLOR)
'              Get BorderStyle()
'              Let BorderStyle(ByVal pintNew_BorderStyle As Integer)
'              Get Color1()
'              Let Color1(ByVal pNew_Color1 As OLE_COLOR)
'              Get Color2()
'              Let Color2(ByVal pNew_Color2 As OLE_COLOR)
'              Get Enabled()
'              Let Enabled(ByVal pfNew_Enabled As Boolean)
'              Get Max()
'              Let Max(ByVal pintNew_Max As Integer)
'              Get Min()
'              Let Min(ByVal pintNew_Min As Integer)
'              Get Orientation()
'              Let Orientation(ByVal pintNew_Orientation As Integer)
'              Get Value()
'              Let Value(ByVal pintNew_Value As Integer)
'              BlendColors()
'              ConvertToRGB()
'              DrawGrad()
'              UserControl_Click()
'              UserControl_DblClick()
'              UserControl_Initialize()
'              UserControl_InitProperties()
'              UserControl_MouseDown(pintButton As Integer, pintShift As Integer, psngX As Single, psngY As Single)
'              UserControl_MouseMove(pintButton As Integer, pintShift As Integer, psngX As Single, psngY As Single)
'              UserControl_MouseUp(pintButton As Integer, pintShift As Integer, psngX As Single, psngY As Single)
'              UserControl_Paint()
'              UserControl_ReadProperties(PropBag As PropertyBag)
'              UserControl_Resize()
'              UserControl_WriteProperties(PropBag As PropertyBag)

' Modified   :
' 03/27/01 s Cleaned with Total Visual CodeTools 2000
'
' --------------------------------------------------

'
'(Cool Progress Bar by Jotaf98) ____________________
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' Read "Readme.txt" for more details.
'
'
'(Contact - E-mail: jotaf98@hotmail.com) ___________
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'

Option Explicit

Const mcstrCurrentModule As String = "CoolProgressBar Control"
Const mcstrDialogTitle As String = "CoolProgressBar Control"

'Default Property Values:
Const mc_def_Orientation = 0
Const mc_def_Value = 100
Const mc_def_Min = 0
Const mc_def_Max = 100
Const mc_def_Color2 = 16777215
Const mc_def_Color1 = 16711680
'Property Variables:
Dim mint_Orientation As Integer
Dim mint_Value As Integer
Dim mint_Min As Integer
Dim mint_Max As Integer
Dim mlng_Color2 As Long
Dim mlng_Color1 As Long
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp

'APIs

'Draws a pixel (used to draw the grad effect)
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

'After drawing the grad to the backbuffer, StretchBlt
'will stretch it to fit the control (still in the
'backbuffer)
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

'Then, when it's needed, copy it to the control using
'BitBlt
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long


'If the grad has already been drawn or not
Dim mfGradDone As Boolean

'The Alpha - the translucency level between the
'first color and the second color.
Dim mintAlpha As Integer

'First color's RGB values
Dim mbytBc_Red1 As Byte
Dim mbytBc_Green1 As Byte
Dim mbytBc_Blue1 As Byte

'Second color's RGB values
Dim mbytBc_Red2 As Byte
Dim mbytBc_Green2 As Byte
Dim mbytBc_Blue2 As Byte

'Final RGB values
Dim mintBc_RedF As Integer
Dim mintBc_GreenF As Integer
Dim mintBc_BlueF As Integer

'When figuring out the bar's width/height based on
'Max, Min and Value, it will be stored here.
Dim mintTempSize As Integer

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
    ' Comments  :
    ' Parameters:  -
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "Refresh"
    
    UserControl.Refresh
PROC_EXIT:
    On Error Resume Next
    Exit Sub
PROC_ERR:
    Select Case Err.Number
    'Case statements for expected errors go here
    Case Else
        ' Display msgbox re: fatal error and terminate the app
        Call fnProcessFatalError(mcstrCurrentModule & "." & cstrCurrentProc, _
                                 fte_DefaultErrType, Err.Number, _
                                 Err.Description, Err.Source, _
                                 Err.HelpFile, Err.HelpContext)
    End Select
    Resume PROC_EXIT
End Sub



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
    ' Comments  :
    ' Parameters:  -
    ' Returns   : OLE_COLOR
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "BackColor Get"
   
    BackColor = UserControl.BackColor
PROC_EXIT:
    On Error Resume Next
    Exit Property
PROC_ERR:
    Select Case Err.Number
    'Case statements for expected errors go here
    Case Else
        ' Display msgbox re: fatal error and terminate the app
        Call fnProcessFatalError(mcstrCurrentModule & "." & cstrCurrentProc, _
                                 fte_DefaultErrType, Err.Number, _
                                 Err.Description, Err.Source, _
                                 Err.HelpFile, Err.HelpContext)
    End Select
    Resume PROC_EXIT
End Property

Public Property Let BackColor(ByVal pNew_BackColor As OLE_COLOR)
    ' Comments  :
    ' Parameters: pNew_BackColor -
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "BackColor Let"
   
    UserControl.BackColor() = pNew_BackColor
   
    'Also set the backbuffer's color
    picBackbuffer.BackColor() = pNew_BackColor
   
    'Redraw the grad effect
    DrawGrad
   
    'Repaint the control
    UserControl_Paint
   
    PropertyChanged "BackColor"
PROC_EXIT:
    On Error Resume Next
    Exit Property
PROC_ERR:
    Select Case Err.Number
    'Case statements for expected errors go here
    Case Else
        ' Display msgbox re: fatal error and terminate the app
        Call fnProcessFatalError(mcstrCurrentModule & "." & cstrCurrentProc, _
                                 fte_DefaultErrType, Err.Number, _
                                 Err.Description, Err.Source, _
                                 Err.HelpFile, Err.HelpContext)
    End Select
    Resume PROC_EXIT
End Property



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
    ' Comments  :
    ' Parameters:  -
    ' Returns   : Integer
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "BorderStyle Get"
   
    BorderStyle = UserControl.BorderStyle
PROC_EXIT:
    On Error Resume Next
    Exit Property
PROC_ERR:
    Select Case Err.Number
    'Case statements for expected errors go here
    Case Else
        ' Display msgbox re: fatal error and terminate the app
        Call fnProcessFatalError(mcstrCurrentModule & "." & cstrCurrentProc, _
                                 fte_DefaultErrType, Err.Number, _
                                 Err.Description, Err.Source, _
                                 Err.HelpFile, Err.HelpContext)
    End Select
    Resume PROC_EXIT
End Property

Public Property Let BorderStyle(ByVal pintNew_BorderStyle As Integer)
    ' Comments  :
    ' Parameters: pintNew_BorderStyle -
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "BorderStyle Let"
   
    UserControl.BorderStyle() = pintNew_BorderStyle
    PropertyChanged "BorderStyle"
PROC_EXIT:
    On Error Resume Next
    Exit Property
PROC_ERR:
    Select Case Err.Number
    'Case statements for expected errors go here
    Case Else
        ' Display msgbox re: fatal error and terminate the app
        Call fnProcessFatalError(mcstrCurrentModule & "." & cstrCurrentProc, _
                                 fte_DefaultErrType, Err.Number, _
                                 Err.Description, Err.Source, _
                                 Err.HelpFile, Err.HelpContext)
    End Select
    Resume PROC_EXIT
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,16711680
Public Property Get Color1() As OLE_COLOR
   ' Comments  :
   ' Parameters:  -
   ' Returns   : OLE_COLOR
   ' Modified  :
   '
   ' --------------------------------------------------
   On Error GoTo PROC_ERR
   Const cstrCurrentProc As String = "Color1 Get"
   
   Color1 = mlng_Color1
PROC_EXIT:
    On Error Resume Next
    Exit Property
PROC_ERR:
    Select Case Err.Number
    'Case statements for expected errors go here
    Case Else
        ' Display msgbox re: fatal error and terminate the app
        Call fnProcessFatalError(mcstrCurrentModule & "." & cstrCurrentProc, _
                                 fte_DefaultErrType, Err.Number, _
                                 Err.Description, Err.Source, _
                                 Err.HelpFile, Err.HelpContext)
    End Select
    Resume PROC_EXIT
End Property

Public Property Let Color1(ByVal pNew_Color1 As OLE_COLOR)
    ' Comments  :
    ' Parameters: pNew_Color1 -
    ' Modified  :
    '
    ' --------------------------------------------------
    'TVCodeTools ErrorEnablerStart
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "Color1 Let"
   
    mlng_Color1 = pNew_Color1
   
    'Convert the new color to RGB
    ConvertToRGB
   
    'Redraw the grad effect
    DrawGrad
   
    'Repaint the control
    UserControl_Paint
   
    PropertyChanged "Color1"
PROC_EXIT:
    On Error Resume Next
    Exit Property
PROC_ERR:
    Select Case Err.Number
    'Case statements for expected errors go here
    Case Else
        ' Display msgbox re: fatal error and terminate the app
        Call fnProcessFatalError(mcstrCurrentModule & "." & cstrCurrentProc, _
                                 fte_DefaultErrType, Err.Number, _
                                 Err.Description, Err.Source, _
                                 Err.HelpFile, Err.HelpContext)
    End Select
    Resume PROC_EXIT
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,16777215
Public Property Get Color2() As OLE_COLOR
    ' Comments  :
    ' Parameters:  -
    ' Returns   : OLE_COLOR
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "Color2 Get"
   
    Color2 = mlng_Color2
PROC_EXIT:
    On Error Resume Next
    Exit Property
PROC_ERR:
    Select Case Err.Number
    'Case statements for expected errors go here
    Case Else
        ' Display msgbox re: fatal error and terminate the app
        Call fnProcessFatalError(mcstrCurrentModule & "." & cstrCurrentProc, _
                                 fte_DefaultErrType, Err.Number, _
                                 Err.Description, Err.Source, _
                                 Err.HelpFile, Err.HelpContext)
    End Select
    Resume PROC_EXIT
End Property

Public Property Let Color2(ByVal pNew_Color2 As OLE_COLOR)
    ' Comments  :
    ' Parameters: pNew_Color2 -
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "Color2 Let"
   
    mlng_Color2 = pNew_Color2
   
    'Convert the new color to RGB
    ConvertToRGB
   
    'Redraw the grad effect
    DrawGrad
   
    'Repaint the control
    UserControl_Paint
   
    PropertyChanged "Color2"
PROC_EXIT:
    On Error Resume Next
    Exit Property
PROC_ERR:
    Select Case Err.Number
    'Case statements for expected errors go here
    Case Else
        ' Display msgbox re: fatal error and terminate the app
        Call fnProcessFatalError(mcstrCurrentModule & "." & cstrCurrentProc, _
                                 fte_DefaultErrType, Err.Number, _
                                 Err.Description, Err.Source, _
                                 Err.HelpFile, Err.HelpContext)
    End Select
    Resume PROC_EXIT
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    ' Comments  :
    ' Parameters:  -
    ' Returns   : Boolean
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "Enabled Get"
   
    Enabled = UserControl.Enabled
PROC_EXIT:
    On Error Resume Next
    Exit Property
PROC_ERR:
    Select Case Err.Number
    'Case statements for expected errors go here
    Case Else
        ' Display msgbox re: fatal error and terminate the app
        Call fnProcessFatalError(mcstrCurrentModule & "." & cstrCurrentProc, _
                                 fte_DefaultErrType, Err.Number, _
                                 Err.Description, Err.Source, _
                                 Err.HelpFile, Err.HelpContext)
    End Select
    Resume PROC_EXIT
End Property

Public Property Let Enabled(ByVal pfNew_Enabled As Boolean)
    ' Comments  :
    ' Parameters: pfNew_Enabled -
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "Enabled Let"
   
    UserControl.Enabled() = pfNew_Enabled
    PropertyChanged "Enabled"
PROC_EXIT:
    On Error Resume Next
    Exit Property
PROC_ERR:
    Select Case Err.Number
    'Case statements for expected errors go here
    Case Else
        ' Display msgbox re: fatal error and terminate the app
        Call fnProcessFatalError(mcstrCurrentModule & "." & cstrCurrentProc, _
                                 fte_DefaultErrType, Err.Number, _
                                 Err.Description, Err.Source, _
                                 Err.HelpFile, Err.HelpContext)
    End Select
    Resume PROC_EXIT
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,100
Public Property Get Max() As Integer
    ' Comments  :
    ' Parameters:  -
    ' Returns   : Integer
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "Max Get"
   
    Max = mint_Max
PROC_EXIT:
    On Error Resume Next
    Exit Property
PROC_ERR:
    Select Case Err.Number
    'Case statements for expected errors go here
    Case Else
        ' Display msgbox re: fatal error and terminate the app
        Call fnProcessFatalError(mcstrCurrentModule & "." & cstrCurrentProc, _
                                 fte_DefaultErrType, Err.Number, _
                                 Err.Description, Err.Source, _
                                 Err.HelpFile, Err.HelpContext)
    End Select
    Resume PROC_EXIT
End Property

Public Property Let Max(ByVal pintNew_Max As Integer)
    ' Comments  :
    ' Parameters: pintNew_Max -
    ' Modified  :
    '
    ' --------------------------------------------------
   On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "Max Let"
   
    mint_Max = pintNew_Max
   
    'Can't use a Max smaller or equal to Min!
    If mint_Max <= mint_Min Then
        MsgBox """Max"" must be greater than ""Min"".", , mcstrDialogTitle
        mint_Max = mint_Min + 1
    End If
   
    'If the Value is greater than Max, set it to Max
    If mint_Value > mint_Max Then
        mint_Value = mint_Max
    End If
   
    'Redraw the grad effect
    DrawGrad
   
    'Repaint
    UserControl_Paint
   
    PropertyChanged "Max"
PROC_EXIT:
    On Error Resume Next
    Exit Property
PROC_ERR:
    Select Case Err.Number
    'Case statements for expected errors go here
    Case Else
        ' Display msgbox re: fatal error and terminate the app
        Call fnProcessFatalError(mcstrCurrentModule & "." & cstrCurrentProc, _
                                 fte_DefaultErrType, Err.Number, _
                                 Err.Description, Err.Source, _
                                 Err.HelpFile, Err.HelpContext)
    End Select
    Resume PROC_EXIT
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get Min() As Integer
    ' Comments  :
    ' Parameters:  -
    ' Returns   : Integer
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "Min Get"
   
    Min = mint_Min
PROC_EXIT:
    On Error Resume Next
    Exit Property
PROC_ERR:
    Select Case Err.Number
    'Case statements for expected errors go here
    Case Else
        ' Display msgbox re: fatal error and terminate the app
        Call fnProcessFatalError(mcstrCurrentModule & "." & cstrCurrentProc, _
                                 fte_DefaultErrType, Err.Number, _
                                 Err.Description, Err.Source, _
                                 Err.HelpFile, Err.HelpContext)
    End Select
    Resume PROC_EXIT
End Property

Public Property Let Min(ByVal pintNew_Min As Integer)
    ' Comments  :
    ' Parameters: pintNew_Min -
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "Min Let"
   
    mint_Min = pintNew_Min
   
    'Can't use a Min greater or equal to Max!
    If mint_Min >= mint_Max Then
        MsgBox """Min"" must be smaller than ""Max"".", , mcstrDialogTitle
        mint_Min = mint_Max - 1
    End If
   
    'If the Value is smaller than Min, set it to Min
    If mint_Value < mint_Min Then
        mint_Value = mint_Min
    End If
   
    'Redraw the grad effect
    DrawGrad
   
    'Repaint
    UserControl_Paint
   
    PropertyChanged "Min"
PROC_EXIT:
    On Error Resume Next
    Exit Property
PROC_ERR:
    Select Case Err.Number
    'Case statements for expected errors go here
    Case Else
        ' Display msgbox re: fatal error and terminate the app
        Call fnProcessFatalError(mcstrCurrentModule & "." & cstrCurrentProc, _
                                 fte_DefaultErrType, Err.Number, _
                                 Err.Description, Err.Source, _
                                 Err.HelpFile, Err.HelpContext)
    End Select
    Resume PROC_EXIT
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get Orientation() As Integer
    ' Comments  :
    ' Parameters:  -
    ' Returns   : Integer
    ' Modified  :
    '
    ' --------------------------------------------------
   On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "Orientation Get"
    
   Orientation = mint_Orientation
PROC_EXIT:
    On Error Resume Next
    Exit Property
PROC_ERR:
    Select Case Err.Number
    'Case statements for expected errors go here
    Case Else
        ' Display msgbox re: fatal error and terminate the app
        Call fnProcessFatalError(mcstrCurrentModule & "." & cstrCurrentProc, _
                                 fte_DefaultErrType, Err.Number, _
                                 Err.Description, Err.Source, _
                                 Err.HelpFile, Err.HelpContext)
    End Select
    Resume PROC_EXIT
End Property

Public Property Let Orientation(ByVal pintNew_Orientation As Integer)
    ' Comments  :
    ' Parameters: pintNew_Orientation -
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "Orientation Let"
   
    mint_Orientation = pintNew_Orientation
   
    'Only accept 0 (Horizontal) or 1 (Vertical)
    If mint_Orientation <> 0 And mint_Orientation <> 1 Then
        mint_Orientation = 0
    End If
   
    'Redraw the grad effect
    DrawGrad
   
    'Repaint the control
    UserControl_Paint
   
    PropertyChanged "Orientation"
PROC_EXIT:
    On Error Resume Next
    Exit Property
PROC_ERR:
    Select Case Err.Number
    'Case statements for expected errors go here
    Case Else
        ' Display msgbox re: fatal error and terminate the app
        Call fnProcessFatalError(mcstrCurrentModule & "." & cstrCurrentProc, _
                                 fte_DefaultErrType, Err.Number, _
                                 Err.Description, Err.Source, _
                                 Err.HelpFile, Err.HelpContext)
    End Select
    Resume PROC_EXIT
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,100
Public Property Get value() As Integer
    ' Comments  :
    ' Parameters:  -
    ' Returns   : Integer
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
   Const cstrCurrentProc As String = "Value Get"
   
    value = mint_Value
PROC_EXIT:
    On Error Resume Next
    Exit Property
PROC_ERR:
    Select Case Err.Number
    'Case statements for expected errors go here
    Case Else
        ' Display msgbox re: fatal error and terminate the app
        Call fnProcessFatalError(mcstrCurrentModule & "." & cstrCurrentProc, _
                                 fte_DefaultErrType, Err.Number, _
                                 Err.Description, Err.Source, _
                                 Err.HelpFile, Err.HelpContext)
    End Select
    Resume PROC_EXIT
End Property

Public Property Let value(ByVal pintNew_Value As Integer)
    ' Comments  :
    ' Parameters: pintNew_Value -
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "Value Let"
   
    mint_Value = pintNew_Value
   
    'Can't have a value greater than max or smaller than min
    If mint_Value > mint_Max Then
        mint_Value = mint_Max
    End If
    If mint_Value < mint_Min Then
        mint_Value = mint_Min
    End If
   
    PropertyChanged "Value"
   
    'Repaint
    UserControl_Paint
PROC_EXIT:
    On Error Resume Next
    Exit Property
PROC_ERR:
    Select Case Err.Number
    'Case statements for expected errors go here
    Case Else
        ' Display msgbox re: fatal error and terminate the app
        Call fnProcessFatalError(mcstrCurrentModule & "." & cstrCurrentProc, _
                                 fte_DefaultErrType, Err.Number, _
                                 Err.Description, Err.Source, _
                                 Err.HelpFile, Err.HelpContext)
    End Select
    Resume PROC_EXIT
End Property

'
'(Functions for the grad effect) ___________________
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'

Private Sub BlendColors()
    ' Comments  : Gets a grad between 2 colors (the results now are in
    '             "bc_[Red/Green/Blue]F")
    ' Parameters:  -
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "BlendColors"
   
    ' Find the difference between Red 1 and Red 2.
    If mbytBc_Red2 > mbytBc_Red1 Then
        mintBc_RedF = mbytBc_Red2 - mbytBc_Red1
    Else
        mintBc_RedF = mbytBc_Red1 - mbytBc_Red2
        mintBc_RedF = -mintBc_RedF
    End If
   
    'This is the core function for Red
    mintBc_RedF = mintBc_RedF / 256 * mintAlpha + mbytBc_Red1
   
    'Find the difference between Green 1 and Green 2.
    If mbytBc_Green2 > mbytBc_Green1 Then
        mintBc_GreenF = mbytBc_Green2 - mbytBc_Green1
    Else
        mintBc_GreenF = mbytBc_Green1 - mbytBc_Green2
        mintBc_GreenF = -mintBc_GreenF
    End If
   
    'This is the core function for Green
    mintBc_GreenF = mintBc_GreenF / 256 * mintAlpha + mbytBc_Green1
   
    'Find the difference between Blue 1 and Blue 2.
    If mbytBc_Blue2 > mbytBc_Blue1 Then
        mintBc_BlueF = mbytBc_Blue2 - mbytBc_Blue1
    Else
        mintBc_BlueF = mbytBc_Blue1 - mbytBc_Blue2
        mintBc_BlueF = -mintBc_BlueF
    End If
   
    'This is the core function for Blue
    mintBc_BlueF = mintBc_BlueF / 256 * mintAlpha + mbytBc_Blue1
PROC_EXIT:
    On Error Resume Next
    Exit Sub
PROC_ERR:
    Select Case Err.Number
    'Case statements for expected errors go here
    Case Else
        ' Display msgbox re: fatal error and terminate the app
        Call fnProcessFatalError(mcstrCurrentModule & "." & cstrCurrentProc, _
                                 fte_DefaultErrType, Err.Number, _
                                 Err.Description, Err.Source, _
                                 Err.HelpFile, Err.HelpContext)
    End Select
    Resume PROC_EXIT
End Sub

'Converts the Long colors to RGB values
Private Sub ConvertToRGB()
    ' Comments  :
    ' Parameters:  -
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "ConvertToRGB"
   
    mbytBc_Red1 = mlng_Color1 And 255
    mbytBc_Green1 = (mlng_Color1 And 65280) \ 256
    mbytBc_Blue1 = (mlng_Color1 And 16711680) \ 65535
   
    mbytBc_Red2 = mlng_Color2 And 255
    mbytBc_Green2 = (mlng_Color2 And 65280) \ 256
    mbytBc_Blue2 = (mlng_Color2 And 16711680) \ 65535
PROC_EXIT:
    On Error Resume Next
    Exit Sub
PROC_ERR:
    Select Case Err.Number
    'Case statements for expected errors go here
    Case Else
        ' Display msgbox re: fatal error and terminate the app
        Call fnProcessFatalError(mcstrCurrentModule & "." & cstrCurrentProc, _
                                 fte_DefaultErrType, Err.Number, _
                                 Err.Description, Err.Source, _
                                 Err.HelpFile, Err.HelpContext)
    End Select
    Resume PROC_EXIT
End Sub

'Will draw the grad effect to picBackbuffer so whenever
'we need to draw the bar, we just copy the part we need
'from it
Private Sub DrawGrad()
    ' Comments  :
    ' Parameters:  -
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "DrawGrad"
   
    'Counter
    Dim intI As Integer
   
    'Convert the colors to RGB
    ConvertToRGB
   
    'Resize the backbuffer to make sure the bar fits in it
    picBackbuffer.Move 0, 0, ScaleWidth, ScaleHeight
   
    'Draw horizontally or vertically, depending on the
    'Orientation property
    If mint_Orientation = 0 Then '<Horizontal>
        'Loop trough all possible grad colors
        For intI = 0 To 255
            'Set the Alpha to the current counter value
            mintAlpha = intI
         
            'Blend the colors
            BlendColors
         
            'Draw the new pixel
            SetPixelV picBackbuffer.hdc, intI, 0, RGB(mintBc_RedF, mintBc_GreenF, mintBc_BlueF)
        Next intI
      
        'Stretch the tiny line we have drawn to fit the
        'control
        StretchBlt picBackbuffer.hdc, 0, 0, ScaleWidth, ScaleHeight, picBackbuffer.hdc, 0, 0, 255, 1, vbSrcCopy
    ElseIf mint_Orientation = 1 Then '<Vertical>
        'Loop trough all possible grad colors
        For intI = 0 To 255
            'Set the Alpha to [255 - the current counter
            'value] (this will make it so Color1 is at
            'the bottom and Color2 at the top)
            mintAlpha = 255 - intI
         
            'Blend the colors
            BlendColors
         
            'Draw the new pixel
            SetPixelV picBackbuffer.hdc, 0, intI, RGB(mintBc_RedF, mintBc_GreenF, mintBc_BlueF)
        Next intI
      
        'Stretch the tiny line we have drawn to fit the
        'control
        StretchBlt picBackbuffer.hdc, 0, 0, ScaleWidth, ScaleHeight, picBackbuffer.hdc, 0, 0, 1, 255, vbSrcCopy
    End If
   
    'Grad done
    mfGradDone = True
PROC_EXIT:
    On Error Resume Next
    Exit Sub
PROC_ERR:
    Select Case Err.Number
    'Case statements for expected errors go here
    Case Else
        ' Display msgbox re: fatal error and terminate the app
        Call fnProcessFatalError(mcstrCurrentModule & "." & cstrCurrentProc, _
                                 fte_DefaultErrType, Err.Number, _
                                 Err.Description, Err.Source, _
                                 Err.HelpFile, Err.HelpContext)
    End Select
    Resume PROC_EXIT
End Sub

Private Sub UserControl_Click()
    ' Comments  :
    ' Parameters:  -
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "UserControl_Click"
   
    RaiseEvent Click
PROC_EXIT:
    On Error Resume Next
    Exit Sub
PROC_ERR:
    Select Case Err.Number
    'Case statements for expected errors go here
    Case Else
        ' Display msgbox re: fatal error and terminate the app
        Call fnProcessFatalError(mcstrCurrentModule & "." & cstrCurrentProc, _
                                 fte_DefaultErrType, Err.Number, _
                                 Err.Description, Err.Source, _
                                 Err.HelpFile, Err.HelpContext)
    End Select
    Resume PROC_EXIT
End Sub

Private Sub UserControl_DblClick()
    ' Comments  :
    ' Parameters:  -
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "UserControl_DblClick"
   
    RaiseEvent DblClick
PROC_EXIT:
    On Error Resume Next
    Exit Sub
PROC_ERR:
    Select Case Err.Number
    'Case statements for expected errors go here
    Case Else
        ' Display msgbox re: fatal error and terminate the app
        Call fnProcessFatalError(mcstrCurrentModule & "." & cstrCurrentProc, _
                                 fte_DefaultErrType, Err.Number, _
                                 Err.Description, Err.Source, _
                                 Err.HelpFile, Err.HelpContext)
    End Select
    Resume PROC_EXIT
End Sub

Private Sub UserControl_Initialize()
    ' Comments  :
    ' Parameters:  -
    ' Modified  :
    '
    ' --------------------------------------------------
    'This will fix a really weird bug - try
    'commenting these lines and creating a
    'progress bar by double-clicking its
    'icon to see what I mean!
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "UserControl_Initialize"
   
    UserControl.Width = UserControl.Width + 60
    UserControl.Height = 255
PROC_EXIT:
    On Error Resume Next
    Exit Sub
PROC_ERR:
    Select Case Err.Number
    'Case statements for expected errors go here
    Case Else
        ' Display msgbox re: fatal error and terminate the app
        Call fnProcessFatalError(mcstrCurrentModule & "." & cstrCurrentProc, _
                                 fte_DefaultErrType, Err.Number, _
                                 Err.Description, Err.Source, _
                                 Err.HelpFile, Err.HelpContext)
    End Select
    Resume PROC_EXIT
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    ' Comments  :
    ' Parameters:  -
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "UserControl_InitProperties"
   
    mlng_Color2 = mc_def_Color2
    mlng_Color1 = mc_def_Color1
    mint_Min = mc_def_Min
    mint_Max = mc_def_Max
    mint_Orientation = mc_def_Orientation
    mint_Value = mc_def_Value
PROC_EXIT:
    On Error Resume Next
    Exit Sub
PROC_ERR:
    Select Case Err.Number
    'Case statements for expected errors go here
    Case Else
        ' Display msgbox re: fatal error and terminate the app
        Call fnProcessFatalError(mcstrCurrentModule & "." & cstrCurrentProc, _
                                 fte_DefaultErrType, Err.Number, _
                                 Err.Description, Err.Source, _
                                 Err.HelpFile, Err.HelpContext)
    End Select
    Resume PROC_EXIT
End Sub

Private Sub UserControl_MouseDown(pintButton As Integer, pintShift As Integer, psngX As Single, psngY As Single)
    ' Comments  :
    ' Parameters: pintButton
    '             pintShift
    '             psngX
    '             psngY -
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "UserControl_MouseDown"
   
    RaiseEvent MouseDown(pintButton, pintShift, psngX, psngY)
PROC_EXIT:
    On Error Resume Next
    Exit Sub
PROC_ERR:
    Select Case Err.Number
    'Case statements for expected errors go here
    Case Else
        ' Display msgbox re: fatal error and terminate the app
        Call fnProcessFatalError(mcstrCurrentModule & "." & cstrCurrentProc, _
                                 fte_DefaultErrType, Err.Number, _
                                 Err.Description, Err.Source, _
                                 Err.HelpFile, Err.HelpContext)
    End Select
    Resume PROC_EXIT
End Sub

Private Sub UserControl_MouseMove(pintButton As Integer, pintShift As Integer, psngX As Single, psngY As Single)
    ' Comments  :
    ' Parameters: pintButton
    '             pintShift
    '             psngX
    '             psngY -
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "UserControl_MouseMove"
   
    RaiseEvent MouseMove(pintButton, pintShift, psngX, psngY)
PROC_EXIT:
    On Error Resume Next
    Exit Sub
PROC_ERR:
    Select Case Err.Number
    'Case statements for expected errors go here
    Case Else
        ' Display msgbox re: fatal error and terminate the app
        Call fnProcessFatalError(mcstrCurrentModule & "." & cstrCurrentProc, _
                                 fte_DefaultErrType, Err.Number, _
                                 Err.Description, Err.Source, _
                                 Err.HelpFile, Err.HelpContext)
    End Select
    Resume PROC_EXIT
End Sub

Private Sub UserControl_MouseUp(pintButton As Integer, pintShift As Integer, psngX As Single, psngY As Single)
    ' Comments  :
    ' Parameters: pintButton
    '             pintShift
    '             psngX
    '             psngY -
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "UserControl_MouseUp"
   
    RaiseEvent MouseUp(pintButton, pintShift, psngX, psngY)
PROC_EXIT:
    On Error Resume Next
    Exit Sub
PROC_ERR:
    Select Case Err.Number
    'Case statements for expected errors go here
    Case Else
        ' Display msgbox re: fatal error and terminate the app
        Call fnProcessFatalError(mcstrCurrentModule & "." & cstrCurrentProc, _
                                 fte_DefaultErrType, Err.Number, _
                                 Err.Description, Err.Source, _
                                 Err.HelpFile, Err.HelpContext)
    End Select
    Resume PROC_EXIT
End Sub

'This is the most important sub - it draws the control
Private Sub UserControl_Paint()
    ' Comments  :
    ' Parameters:  -
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "UserControl_Paint"
   
    'If the grad is not already drawn, draw it
    If mfGradDone = False Then
        DrawGrad
    End If
   
    'Calculate TempSize horizontally or vertically, depending
    'on the Orientation property
    If mint_Orientation = 0 Then '<Horizontal>
        'This will get the width of the bar based on the value
        mintTempSize = ScaleWidth / (mint_Max - mint_Min) * mint_Value
      
        'Clear the control and copy the grad to it, according to
        'the new width of the bar
        Cls
        BitBlt hdc, 0, 0, mintTempSize, ScaleHeight, picBackbuffer.hdc, 0, 0, vbSrcCopy
    ElseIf mint_Orientation = 1 Then '<Vertical>
        'This will get the height of the bar based on the value
        mintTempSize = ScaleHeight / (mint_Max - mint_Min) * (mint_Max - mint_Value)
      
        'Clear the control and copy the grad to it, according to
        'the new width of the bar
        Cls
        BitBlt hdc, 0, mintTempSize, ScaleWidth, ScaleHeight - mintTempSize, picBackbuffer.hdc, 0, mintTempSize, vbSrcCopy
    End If
PROC_EXIT:
    On Error Resume Next
    Exit Sub
PROC_ERR:
    Select Case Err.Number
    'Case statements for expected errors go here
    Case Else
        ' Display msgbox re: fatal error and terminate the app
        Call fnProcessFatalError(mcstrCurrentModule & "." & cstrCurrentProc, _
                                 fte_DefaultErrType, Err.Number, _
                                 Err.Description, Err.Source, _
                                 Err.HelpFile, Err.HelpContext)
    End Select
    Resume PROC_EXIT
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    ' Comments  :
    ' Parameters: PropBag -
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "UserControl_ReadProperties"
   
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    mlng_Color2 = PropBag.ReadProperty("Color2", mc_def_Color2)
    mlng_Color1 = PropBag.ReadProperty("Color1", mc_def_Color1)
    mint_Min = PropBag.ReadProperty("Min", mc_def_Min)
    mint_Max = PropBag.ReadProperty("Max", mc_def_Max)
    mint_Orientation = PropBag.ReadProperty("Orientation", mc_def_Orientation)
    mint_Value = PropBag.ReadProperty("Value", mc_def_Value)
PROC_EXIT:
    On Error Resume Next
    Exit Sub
PROC_ERR:
    Select Case Err.Number
    'Case statements for expected errors go here
    Case Else
        ' Display msgbox re: fatal error and terminate the app
        Call fnProcessFatalError(mcstrCurrentModule & "." & cstrCurrentProc, _
                                 fte_DefaultErrType, Err.Number, _
                                 Err.Description, Err.Source, _
                                 Err.HelpFile, Err.HelpContext)
    End Select
    Resume PROC_EXIT
End Sub

Private Sub UserControl_Resize()
    ' Comments  :
    ' Parameters:  -
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "UserControl_Resize"
   
    If Not mfGradDone Then
        Exit Sub
    End If
   
    'Redraw the grad effect
    DrawGrad
   
    'Repaint the control
    UserControl_Paint
PROC_EXIT:
    On Error Resume Next
    Exit Sub
PROC_ERR:
    Select Case Err.Number
    'Case statements for expected errors go here
    Case Else
        ' Display msgbox re: fatal error and terminate the app
        Call fnProcessFatalError(mcstrCurrentModule & "." & cstrCurrentProc, _
                                 fte_DefaultErrType, Err.Number, _
                                 Err.Description, Err.Source, _
                                 Err.HelpFile, Err.HelpContext)
    End Select
    Resume PROC_EXIT
End Sub


'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    ' Comments  :
    ' Parameters: PropBag -
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "UserControl_WriteProperties"
   
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("Color2", mlng_Color2, mc_def_Color2)
    Call PropBag.WriteProperty("Color1", mlng_Color1, mc_def_Color1)
    Call PropBag.WriteProperty("Min", mint_Min, mc_def_Min)
    Call PropBag.WriteProperty("Max", mint_Max, mc_def_Max)
    Call PropBag.WriteProperty("Orientation", mint_Orientation, mc_def_Orientation)
    Call PropBag.WriteProperty("Value", mint_Value, mc_def_Value)
PROC_EXIT:
    On Error Resume Next
    Exit Sub
PROC_ERR:
    Select Case Err.Number
    'Case statements for expected errors go here
    Case Else
        ' Display msgbox re: fatal error and terminate the app
        Call fnProcessFatalError(mcstrCurrentModule & "." & cstrCurrentProc, _
                                 fte_DefaultErrType, Err.Number, _
                                 Err.Description, Err.Source, _
                                 Err.HelpFile, Err.HelpContext)
    End Select
    Resume PROC_EXIT
End Sub

