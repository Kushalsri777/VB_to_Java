VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPayee 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Payee"
   ClientHeight    =   8476
   ClientLeft      =   1716
   ClientTop       =   780
   ClientWidth     =   10179
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.15
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPayee.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8476
   ScaleWidth      =   10179
   ShowInTaskbar   =   0   'False
   Begin LpLib.fpCombo lpcLookupName 
      Height          =   286
      Left            =   1326
      TabIndex        =   1
      Top             =   117
      Width           =   4563
      _Version        =   196608
      _ExtentX        =   8049
      _ExtentY        =   504
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.1509
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
      Columns         =   0
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      WrapList        =   0   'False
      WrapWidth       =   0
      AutoSearch      =   1
      SearchMethod    =   0
      VirtualMode     =   0   'False
      VRowCount       =   0
      DataSync        =   3
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483627
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ScrollHScale    =   2
      ScrollHInc      =   0
      ColsFrozen      =   0
      ScrollBarV      =   1
      NoIntegralHeight=   0   'False
      HighestPrecedence=   0
      AllowColResize  =   0
      AllowColDragDrop=   0
      ReadOnly        =   0   'False
      VScrollSpecial  =   0   'False
      VScrollSpecialType=   0
      EnableKeyEvents =   -1  'True
      EnableTopChangeEvent=   -1  'True
      DataAutoHeadings=   -1  'True
      DataAutoSizeCols=   2
      SearchIgnoreCase=   -1  'True
      ScrollBarH      =   1
      DataFieldList   =   ""
      ColumnEdit      =   -1
      ColumnBound     =   -1
      Style           =   0
      MaxDrop         =   8
      ListWidth       =   -1
      EditHeight      =   -1
      GrayAreaColor   =   -2147483633
      ListLeftOffset  =   0
      ComboGap        =   -2
      MaxEditLen      =   150
      VirtualPageSize =   0
      VirtualPagesAhead=   0
      ExtendCol       =   0
      ColumnLevels    =   1
      ListGrayAreaColor=   -2147483637
      GroupHeaderHeight=   -1
      GroupHeaderShow =   0   'False
      AllowGrpResize  =   0
      AllowGrpDragDrop=   0
      MergeAdjustView =   0   'False
      ColumnHeaderShow=   0   'False
      ColumnHeaderHeight=   -1
      GrpsFrozen      =   0
      BorderGrayAreaColor=   -2147483637
      ExtendRow       =   0
      ListPosition    =   0
      ButtonThreeDAppearance=   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      Redraw          =   -1  'True
      AutoSearchFill  =   0   'False
      AutoSearchFillDelay=   500
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      AllowAnimate    =   -1  'True
      ColDesigner     =   "frmPayee.frx":030A
   End
   Begin VB.Frame fraStatesToCalculate 
      Caption         =   "States Used in Automatic Calculation"
      Height          =   3555
      Left            =   5325
      TabIndex        =   26
      Top             =   720
      Width           =   4755
      Begin VB.CheckBox chkClmForResDthInd_UsedInAutoCalc 
         Caption         =   "Foreign Res. at Death?"
         Enabled         =   0   'False
         Height          =   280
         Left            =   240
         TabIndex        =   73
         Top             =   1350
         Width           =   1995
      End
      Begin VB.TextBox txtIssStCd_UsedInAutoCalc 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         ToolTipText     =   "The Insured's state of residence, at time of issue"
         Top             =   2400
         Width           =   420
      End
      Begin VB.TextBox txtIssStCdSpecialInstructions_UsedInAutoCalc 
         BackColor       =   &H8000000F&
         ForeColor       =   &H80000012&
         Height          =   675
         Left            =   180
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   35
         TabStop         =   0   'False
         ToolTipText     =   "Special instructions associated with the Issue State"
         Top             =   2760
         Width           =   4455
      End
      Begin VB.TextBox txtInsdDthResStCdSpecialInstructions_UsedInAutoCalc 
         BackColor       =   &H8000000F&
         ForeColor       =   &H80000012&
         Height          =   675
         Left            =   180
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         TabStop         =   0   'False
         ToolTipText     =   "Special instructions associated with the Insured Residence State"
         Top             =   1680
         Width           =   4455
      End
      Begin VB.TextBox txtInsdDthResStCd_UsedInAutoCalc 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   4195
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         ToolTipText     =   "The Insured's state of residence, at time of death"
         Top             =   1320
         Width           =   420
      End
      Begin VB.TextBox txtPayeStCd_UsedInAutoCalc 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   4215
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "The net proceeds payable to this Payee"
         Top             =   240
         Width           =   420
      End
      Begin VB.TextBox txtPayeStCdSpecialInstructions_UsedInAutoCalc 
         BackColor       =   &H8000000F&
         ForeColor       =   &H80000012&
         Height          =   675
         Left            =   180
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   "Special Instructions associated with the Payee Residence State"
         Top             =   600
         Width           =   4455
      End
      Begin VB.Label lblIssStCd_UsedInAutoCalc 
         BackStyle       =   0  'Transparent
         Caption         =   "Issue State:"
         Height          =   285
         Left            =   3240
         TabIndex        =   33
         Top             =   2460
         Width           =   1005
      End
      Begin VB.Label lblInsdDthResStCd_UsedInAutoCalc 
         BackStyle       =   0  'Transparent
         Caption         =   "Insured Residence State:"
         Height          =   280
         Left            =   2340
         TabIndex        =   30
         Top             =   1380
         Width           =   1830
      End
      Begin VB.Label lblPayeStCd_UsedInAutoCalc 
         BackStyle       =   0  'Transparent
         Caption         =   "Payee Residence State:"
         Height          =   285
         Left            =   2400
         TabIndex        =   27
         Top             =   300
         Width           =   1830
      End
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   375
      Left            =   3627
      TabIndex        =   69
      ToolTipText     =   "Save your changes"
      Top             =   7980
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   4905
      TabIndex        =   70
      ToolTipText     =   "Delete this Payee"
      Top             =   7980
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   6165
      TabIndex        =   71
      ToolTipText     =   "Cancel your changes or close this screen"
      Top             =   7980
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   2385
      TabIndex        =   68
      ToolTipText     =   "Add a new Payee"
      Top             =   7980
      Width           =   1215
   End
   Begin VB.CommandButton cmdNavigate 
      Caption         =   ">>"
      Height          =   375
      Index           =   3
      Left            =   1380
      TabIndex        =   67
      ToolTipText     =   "Go to last record"
      Top             =   7980
      Width           =   435
   End
   Begin VB.CommandButton cmdNavigate 
      Caption         =   ">"
      Height          =   375
      Index           =   2
      Left            =   960
      TabIndex        =   66
      ToolTipText     =   "Go to next record"
      Top             =   7980
      Width           =   435
   End
   Begin VB.CommandButton cmdNavigate 
      Caption         =   "<"
      Height          =   375
      Index           =   1
      Left            =   540
      TabIndex        =   65
      ToolTipText     =   "Go to previous record"
      Top             =   7980
      Width           =   435
   End
   Begin VB.CommandButton cmdNavigate 
      Caption         =   "<<"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   64
      ToolTipText     =   "Go to first record"
      Top             =   7980
      Width           =   435
   End
   Begin VB.Frame fraClaimInfo 
      Caption         =   "Claim Information"
      Height          =   3270
      Left            =   120
      TabIndex        =   36
      Top             =   4365
      Width           =   9945
      Begin VB.Frame fraCalculation 
         Caption         =   "Calculation"
         Height          =   2325
         Left            =   225
         TabIndex        =   37
         Top             =   240
         Width           =   5670
         Begin VB.ComboBox cboCalcStCd 
            BackColor       =   &H8000000F&
            Height          =   315
            ItemData        =   "frmPayee.frx":05C9
            Left            =   1725
            List            =   "frmPayee.frx":05CB
            Style           =   2  'Dropdown List
            TabIndex        =   43
            ToolTipText     =   "State upon which the final calculation was based"
            Top             =   1050
            Width           =   795
         End
         Begin VB.CheckBox chkPayeDfltOvrdInd 
            Caption         =   "&Override:"
            Height          =   315
            Left            =   240
            TabIndex        =   41
            ToolTipText     =   "Indicates whether the default rule or state used in calculations should be overridden"
            Top             =   705
            Width           =   1065
         End
         Begin VB.TextBox txtCalcStCdSpecialInstructions_UsedInAutoCalc 
            BackColor       =   &H8000000F&
            ForeColor       =   &H80000012&
            Height          =   675
            Left            =   1725
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   47
            TabStop         =   0   'False
            ToolTipText     =   "Special Instructions associated with the state used in the calculation"
            Top             =   1440
            Width           =   3810
         End
         Begin EditLib.fpDoubleSingle ipdPayeWthldRt 
            Height          =   315
            Left            =   1725
            TabIndex        =   39
            Top             =   285
            Width           =   945
            _Version        =   196608
            _ExtentX        =   1667
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.1509
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483633
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            ButtonDisable   =   0   'False
            ButtonHide      =   0   'False
            ButtonIncrement =   1
            ButtonMin       =   0
            ButtonMax       =   100
            ButtonStyle     =   0
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483633
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "0"
            DecimalPlaces   =   -1
            DecimalPoint    =   ""
            FixedPoint      =   0   'False
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "-9000000000"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
            IncDec          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483633
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDoubleSingle ipdPayeClmIntRt 
            Height          =   315
            Left            =   3900
            TabIndex        =   45
            Top             =   1050
            Width           =   945
            _Version        =   196608
            _ExtentX        =   1667
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.1509
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483633
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            ButtonDisable   =   0   'False
            ButtonHide      =   0   'False
            ButtonIncrement =   1
            ButtonMin       =   0
            ButtonMax       =   100
            ButtonStyle     =   0
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483633
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "0"
            DecimalPlaces   =   -1
            DecimalPoint    =   ""
            FixedPoint      =   0   'False
            LeadZero        =   0
            MaxValue        =   "100"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
            IncDec          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483633
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label lblWarningAboutOverride 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WARNING: The previous calculation was overridden!"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.15
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   765
            Left            =   2895
            TabIndex        =   40
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label lblCalcStCdSpecialInstructions 
            Caption         =   "Special Instructions:"
            Height          =   435
            Left            =   750
            TabIndex        =   46
            Top             =   1440
            Width           =   900
         End
         Begin VB.Label lblPayeWthldRt 
            BackStyle       =   0  'Transparent
            Caption         =   "&Withholding Rate:"
            Height          =   285
            Left            =   225
            TabIndex        =   38
            Top             =   300
            Width           =   1395
         End
         Begin VB.Label lblPayeClmIntRt 
            BackStyle       =   0  'Transparent
            Caption         =   "Inte&rest Rate:"
            Height          =   285
            Left            =   2790
            TabIndex        =   44
            Top             =   1080
            Width           =   1170
         End
         Begin VB.Label lblCalcStCd 
            BackStyle       =   0  'Transparent
            Caption         =   "Ca&lc State:"
            Height          =   285
            Left            =   750
            TabIndex        =   42
            Top             =   1080
            Width           =   975
         End
      End
      Begin MSComCtl2.DTPicker dtpPayePmtDt 
         Height          =   315
         Left            =   7845
         TabIndex        =   49
         ToolTipText     =   "Date on which the Payee will be paid the proceeds from this claim"
         Top             =   240
         Width           =   1635
         _ExtentX        =   2875
         _ExtentY        =   551
         _Version        =   393216
         Format          =   129040385
         CurrentDate     =   37013
         MinDate         =   21916
      End
      Begin EditLib.fpCurrency ipcPayeDthbPmtAmt 
         Height          =   315
         Left            =   7845
         TabIndex        =   53
         Top             =   975
         Width           =   2000
         _Version        =   196608
         _ExtentX        =   3528
         _ExtentY        =   556
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.1509
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         ThreeDInsideStyle=   1
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   1
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   1
         ThreeDFrameWidth=   0
         BorderStyle     =   0
         BorderColor     =   -2147483642
         BorderWidth     =   1
         ButtonDisable   =   0   'False
         ButtonHide      =   0   'False
         ButtonIncrement =   1
         ButtonMin       =   0
         ButtonMax       =   100
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   -2147483637
         InvalidOption   =   0
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   "$0.00"
         CurrencyDecimalPlaces=   -1
         CurrencyNegFormat=   0
         CurrencyPlacement=   0
         CurrencySymbol  =   ""
         DecimalPoint    =   ""
         FixedPoint      =   -1  'True
         LeadZero        =   0
         MaxValue        =   "9000000000"
         MinValue        =   "-9000000000"
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
         IncInt          =   1
         IncDec          =   1
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpCurrency ipcPayeWthldAmt 
         Height          =   315
         Left            =   7845
         TabIndex        =   59
         Top             =   1680
         Width           =   2000
         _Version        =   196608
         _ExtentX        =   3528
         _ExtentY        =   556
         Enabled         =   0   'False
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.1509
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ThreeDInsideStyle=   1
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   1
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   1
         ThreeDFrameWidth=   0
         BorderStyle     =   0
         BorderColor     =   -2147483642
         BorderWidth     =   1
         ButtonDisable   =   0   'False
         ButtonHide      =   0   'False
         ButtonIncrement =   1
         ButtonMin       =   0
         ButtonMax       =   100
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   -2147483637
         InvalidOption   =   0
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   "$0.00"
         CurrencyDecimalPlaces=   -1
         CurrencyNegFormat=   0
         CurrencyPlacement=   0
         CurrencySymbol  =   ""
         DecimalPoint    =   ""
         FixedPoint      =   -1  'True
         LeadZero        =   0
         MaxValue        =   "9000000000"
         MinValue        =   "-9000000000"
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
         IncInt          =   1
         IncDec          =   1
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpCurrency ipcPayeClmIntAmt 
         Height          =   315
         Left            =   7845
         TabIndex        =   56
         Top             =   1320
         Width           =   2000
         _Version        =   196608
         _ExtentX        =   3528
         _ExtentY        =   556
         Enabled         =   0   'False
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.1509
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ThreeDInsideStyle=   1
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   1
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   1
         ThreeDFrameWidth=   0
         BorderStyle     =   0
         BorderColor     =   -2147483642
         BorderWidth     =   1
         ButtonDisable   =   0   'False
         ButtonHide      =   0   'False
         ButtonIncrement =   1
         ButtonMin       =   0
         ButtonMax       =   100
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   -2147483637
         InvalidOption   =   0
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   "$0.00"
         CurrencyDecimalPlaces=   -1
         CurrencyNegFormat=   0
         CurrencyPlacement=   0
         CurrencySymbol  =   ""
         DecimalPoint    =   ""
         FixedPoint      =   -1  'True
         LeadZero        =   0
         MaxValue        =   "9000000000"
         MinValue        =   "-9000000000"
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
         IncInt          =   1
         IncDec          =   1
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpCurrency ipcPayeClmPdAmt 
         Height          =   315
         Left            =   7830
         TabIndex        =   61
         Top             =   2235
         Width           =   2000
         _Version        =   196608
         _ExtentX        =   3528
         _ExtentY        =   556
         Enabled         =   0   'False
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.1509
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ThreeDInsideStyle=   1
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   1
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   1
         ThreeDFrameWidth=   0
         BorderStyle     =   0
         BorderColor     =   -2147483642
         BorderWidth     =   1
         ButtonDisable   =   0   'False
         ButtonHide      =   0   'False
         ButtonIncrement =   1
         ButtonMin       =   0
         ButtonMax       =   100
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   -2147483637
         InvalidOption   =   0
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   "$0.00"
         CurrencyDecimalPlaces=   -1
         CurrencyNegFormat=   0
         CurrencyPlacement=   0
         CurrencySymbol  =   ""
         DecimalPoint    =   ""
         FixedPoint      =   -1  'True
         LeadZero        =   0
         MaxValue        =   "9000000000"
         MinValue        =   "-9000000000"
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
         IncInt          =   1
         IncDec          =   1
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpDoubleSingle ipdPayeIntDaysPdNum 
         Height          =   315
         Left            =   7845
         TabIndex        =   51
         Top             =   615
         Width           =   735
         _Version        =   196608
         _ExtentX        =   1296
         _ExtentY        =   556
         Enabled         =   0   'False
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.1509
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ThreeDInsideStyle=   1
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   1
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   1
         ThreeDFrameWidth=   0
         BorderStyle     =   0
         BorderColor     =   -2147483642
         BorderWidth     =   1
         ButtonDisable   =   0   'False
         ButtonHide      =   0   'False
         ButtonIncrement =   1
         ButtonMin       =   0
         ButtonMax       =   100
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   -2147483637
         InvalidOption   =   0
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   "0"
         DecimalPlaces   =   -1
         DecimalPoint    =   ""
         FixedPoint      =   0   'False
         LeadZero        =   0
         MaxValue        =   "9000000000"
         MinValue        =   "-9000000000"
         NegFormat       =   1
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
         IncInt          =   1
         IncDec          =   1
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.Line linTotals 
         BorderWidth     =   2
         X1              =   6705
         X2              =   9840
         Y1              =   2100
         Y2              =   2100
      End
      Begin VB.Label lblMinus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7590
         TabIndex        =   58
         Top             =   1725
         Width           =   210
      End
      Begin VB.Label lblPlus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7590
         TabIndex        =   55
         Top             =   1365
         Width           =   210
      End
      Begin VB.Label lblPayeIntDaysPdNum 
         BackStyle       =   0  'Transparent
         Caption         =   "Days of Interest Paid:"
         Height          =   285
         Left            =   6180
         TabIndex        =   50
         Top             =   630
         Width           =   1605
      End
      Begin VB.Label lblCalculationInfo 
         Alignment       =   2  'Center
         Height          =   495
         Left            =   225
         TabIndex        =   62
         Top             =   2640
         Width           =   9495
      End
      Begin VB.Label lblPayeDthbPmtAmt 
         BackStyle       =   0  'Transparent
         Caption         =   "DB Pay&ment:"
         Height          =   285
         Left            =   6180
         TabIndex        =   52
         Top             =   990
         Width           =   1260
      End
      Begin VB.Label lblPayeWthldAmt 
         BackStyle       =   0  'Transparent
         Caption         =   "Interest Withheld:"
         Height          =   285
         Left            =   6180
         TabIndex        =   57
         Top             =   1695
         Width           =   1365
      End
      Begin VB.Label lblPayePmtDt 
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Pa&yment:"
         Height          =   285
         Left            =   6180
         TabIndex        =   48
         Top             =   255
         Width           =   1560
      End
      Begin VB.Label lblPayeClmPdAmt 
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         Height          =   285
         Left            =   7005
         TabIndex        =   60
         Top             =   2250
         Width           =   645
      End
      Begin VB.Label lblPayeClmIntAmt 
         BackStyle       =   0  'Transparent
         Caption         =   "Claim Interest:"
         Height          =   285
         Left            =   6180
         TabIndex        =   54
         Top             =   1335
         Width           =   1260
      End
   End
   Begin VB.Frame fraPayeeInfo 
      Caption         =   "Payee"
      Height          =   3555
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   5115
      Begin VB.CheckBox ChkPaye1099Ind 
         Caption         =   "1099INT"
         Height          =   255
         Left            =   3240
         TabIndex        =   74
         Top             =   2400
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.ComboBox cboPayeSsnTinTypCd 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   24
         ToolTipText     =   "The state in which the Insured resided at time of issue"
         Top             =   2760
         Width           =   540
      End
      Begin VB.ComboBox cboPayeStCd 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   16
         ToolTipText     =   "The state in which the Insured resided at time of issue"
         Top             =   2040
         Width           =   795
      End
      Begin VB.CommandButton cmdCloneThisPayee 
         Caption         =   "Clone This &Payee"
         Height          =   375
         Left            =   3300
         TabIndex        =   25
         ToolTipText     =   "Close this screen"
         Top             =   3000
         Width           =   1590
      End
      Begin EditLib.fpText iptPayeFullNm 
         Height          =   315
         Left            =   1080
         TabIndex        =   6
         Top             =   240
         Width           =   3855
         _Version        =   196608
         _ExtentX        =   6800
         _ExtentY        =   556
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.1509
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         ThreeDInsideStyle=   1
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   1
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   1
         ThreeDFrameWidth=   0
         BorderStyle     =   0
         BorderColor     =   -2147483642
         BorderWidth     =   1
         ButtonDisable   =   0   'False
         ButtonHide      =   0   'False
         ButtonIncrement =   1
         ButtonMin       =   0
         ButtonMax       =   100
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         AutoCase        =   0
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   255
         InvalidOption   =   0
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   ""
         CharValidationText=   ""
         MaxLength       =   255
         MultiLine       =   0   'False
         PasswordChar    =   ""
         IncHoriz        =   0.25
         BorderGrayAreaColor=   -2147483637
         NoPrefix        =   0   'False
         ScrollV         =   0   'False
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpText iptPayeCareOfTxt 
         Height          =   315
         Left            =   1080
         TabIndex        =   8
         Top             =   600
         Width           =   3855
         _Version        =   196608
         _ExtentX        =   6800
         _ExtentY        =   556
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.1509
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         ThreeDInsideStyle=   1
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   1
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   1
         ThreeDFrameWidth=   0
         BorderStyle     =   0
         BorderColor     =   -2147483642
         BorderWidth     =   1
         ButtonDisable   =   0   'False
         ButtonHide      =   0   'False
         ButtonIncrement =   1
         ButtonMin       =   0
         ButtonMax       =   100
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         AutoCase        =   0
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   255
         InvalidOption   =   0
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   ""
         CharValidationText=   ""
         MaxLength       =   255
         MultiLine       =   0   'False
         PasswordChar    =   ""
         IncHoriz        =   0.25
         BorderGrayAreaColor=   -2147483637
         NoPrefix        =   0   'False
         ScrollV         =   0   'False
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpText iptPayeAddrLn1Txt 
         Height          =   315
         Left            =   1080
         TabIndex        =   10
         Top             =   960
         Width           =   3855
         _Version        =   196608
         _ExtentX        =   6800
         _ExtentY        =   556
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.1509
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         ThreeDInsideStyle=   1
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   1
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   1
         ThreeDFrameWidth=   0
         BorderStyle     =   0
         BorderColor     =   -2147483642
         BorderWidth     =   1
         ButtonDisable   =   0   'False
         ButtonHide      =   0   'False
         ButtonIncrement =   1
         ButtonMin       =   0
         ButtonMax       =   100
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         AutoCase        =   0
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   255
         InvalidOption   =   0
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   ""
         CharValidationText=   ""
         MaxLength       =   255
         MultiLine       =   0   'False
         PasswordChar    =   ""
         IncHoriz        =   0.25
         BorderGrayAreaColor=   -2147483637
         NoPrefix        =   0   'False
         ScrollV         =   0   'False
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpText iptPayeAddrLn2Txt 
         Height          =   315
         Left            =   1080
         TabIndex        =   12
         Top             =   1320
         Width           =   3855
         _Version        =   196608
         _ExtentX        =   6800
         _ExtentY        =   556
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.1509
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         ThreeDInsideStyle=   1
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   1
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   1
         ThreeDFrameWidth=   0
         BorderStyle     =   0
         BorderColor     =   -2147483642
         BorderWidth     =   1
         ButtonDisable   =   0   'False
         ButtonHide      =   0   'False
         ButtonIncrement =   1
         ButtonMin       =   0
         ButtonMax       =   100
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         AutoCase        =   0
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   255
         InvalidOption   =   0
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   ""
         CharValidationText=   ""
         MaxLength       =   255
         MultiLine       =   0   'False
         PasswordChar    =   ""
         IncHoriz        =   0.25
         BorderGrayAreaColor=   -2147483637
         NoPrefix        =   0   'False
         ScrollV         =   0   'False
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpText iptPayeCityNmTxt 
         Height          =   315
         Left            =   1080
         TabIndex        =   14
         Top             =   1680
         Width           =   3855
         _Version        =   196608
         _ExtentX        =   6800
         _ExtentY        =   556
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.1509
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         ThreeDInsideStyle=   1
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   1
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   1
         ThreeDFrameWidth=   0
         BorderStyle     =   0
         BorderColor     =   -2147483642
         BorderWidth     =   1
         ButtonDisable   =   0   'False
         ButtonHide      =   0   'False
         ButtonIncrement =   1
         ButtonMin       =   0
         ButtonMax       =   100
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         AutoCase        =   0
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   255
         InvalidOption   =   0
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   ""
         CharValidationText=   ""
         MaxLength       =   255
         MultiLine       =   0   'False
         PasswordChar    =   ""
         IncHoriz        =   0.25
         BorderGrayAreaColor=   -2147483637
         NoPrefix        =   0   'False
         ScrollV         =   0   'False
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpText iptPayeZipCd 
         Height          =   315
         Left            =   2400
         TabIndex        =   18
         Top             =   2040
         Width           =   855
         _Version        =   196608
         _ExtentX        =   1508
         _ExtentY        =   556
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.1509
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         ThreeDInsideStyle=   1
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   1
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   1
         ThreeDFrameWidth=   0
         BorderStyle     =   0
         BorderColor     =   -2147483642
         BorderWidth     =   1
         ButtonDisable   =   0   'False
         ButtonHide      =   0   'False
         ButtonIncrement =   1
         ButtonMin       =   0
         ButtonMax       =   100
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         AutoCase        =   0
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   255
         InvalidOption   =   0
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   ""
         CharValidationText=   ""
         MaxLength       =   255
         MultiLine       =   0   'False
         PasswordChar    =   ""
         IncHoriz        =   0.25
         BorderGrayAreaColor=   -2147483637
         NoPrefix        =   0   'False
         ScrollV         =   0   'False
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpText iptPayeZip4Cd 
         Height          =   315
         Left            =   3360
         TabIndex        =   20
         Top             =   2040
         Width           =   735
         _Version        =   196608
         _ExtentX        =   1296
         _ExtentY        =   556
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.1509
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         ThreeDInsideStyle=   1
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   1
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   1
         ThreeDFrameWidth=   0
         BorderStyle     =   0
         BorderColor     =   -2147483642
         BorderWidth     =   1
         ButtonDisable   =   0   'False
         ButtonHide      =   0   'False
         ButtonIncrement =   1
         ButtonMin       =   0
         ButtonMax       =   100
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         AutoCase        =   0
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   255
         InvalidOption   =   0
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   ""
         CharValidationText=   ""
         MaxLength       =   255
         MultiLine       =   0   'False
         PasswordChar    =   ""
         IncHoriz        =   0.25
         BorderGrayAreaColor=   -2147483637
         NoPrefix        =   0   'False
         ScrollV         =   0   'False
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpMask ipmPayeSsnTinNum 
         Height          =   315
         Left            =   1080
         TabIndex        =   22
         Top             =   2400
         Width           =   1875
         _Version        =   196608
         _ExtentX        =   3307
         _ExtentY        =   556
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.1509
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         ThreeDInsideStyle=   1
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   1
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   1
         ThreeDFrameWidth=   0
         BorderStyle     =   0
         BorderColor     =   -2147483642
         BorderWidth     =   1
         ButtonDisable   =   0   'False
         ButtonHide      =   0   'False
         ButtonIncrement =   1
         ButtonMin       =   0
         ButtonMax       =   100
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   255
         InvalidOption   =   0
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         AllowOverflow   =   0   'False
         BestFit         =   0   'False
         ClipMode        =   0
         DataFormatEx    =   0
         Mask            =   ""
         PromptChar      =   "_"
         PromptInclude   =   0   'False
         RequireFill     =   0   'False
         BorderGrayAreaColor=   -2147483637
         NoPrefix        =   0   'False
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         AutoTab         =   0   'False
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.Label lblDash 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.11
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   3270
         TabIndex        =   19
         Top             =   2150
         Width           =   75
      End
      Begin VB.Label lblPayeSsnTinNum 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&TIN/SSN:"
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   2460
         Width           =   660
      End
      Begin VB.Label lblPayeFullNm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Name:"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   300
         Width           =   465
      End
      Begin VB.Label lblPayeAddrLn1Txt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address&1:"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   1020
         Width           =   735
      End
      Begin VB.Label lblPayZipCd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Zip:"
         Height          =   195
         Left            =   2040
         TabIndex        =   17
         Top             =   2100
         Width           =   270
      End
      Begin VB.Label lblPayeCityNmTxt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C&ity:"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1740
         Width           =   345
      End
      Begin VB.Label lblPayeStCd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&State:"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   2100
         Width           =   450
      End
      Begin VB.Label lblPayeSsnTinTypCd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TIN Typ&e:"
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   2820
         Width           =   720
      End
      Begin VB.Label lblPayeAddrLn2Txt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address&2:"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   1380
         Width           =   735
      End
      Begin VB.Label lblPayeCareOfTxt 
         AutoSize        =   -1  'True
         Caption         =   "Care O&f:"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   660
         Width           =   630
      End
   End
   Begin VB.Label lblHowToSetNull 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Press F2 to clear a field."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7750
      TabIndex        =   72
      Top             =   7680
      Width           =   2295
   End
   Begin VB.Label lblClmNum_label 
      BackStyle       =   0  'Transparent
      Caption         =   "Claim #"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3075
      TabIndex        =   2
      Top             =   540
      Width           =   855
   End
   Begin VB.Label lblClmNum 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.51
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3975
      TabIndex        =   3
      Top             =   510
      Width           =   3675
   End
   Begin VB.Label lblLookup 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name Lookup:"
      ForeColor       =   &H80000013&
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   195
      Width           =   1020
   End
   Begin VB.Shape shpLookup 
      BackColor       =   &H80000003&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   435
      Left            =   120
      Top             =   60
      Width           =   9955
   End
   Begin VB.Label lblRecordPosition 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Record x of y"
      Height          =   195
      Left            =   120
      TabIndex        =   63
      Top             =   7680
      Width           =   975
   End
End
Attribute VB_Name = "frmPayee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'!TODO! Add override capability
'******************************************************************************
' Module     : frmPayee
' Description:
' Procedures :
'              cboCalcStCd_Click()
'              cboPayeSsnTinTypCd_Click()
'              cboPayeStCd_Click()
'              chkPayeDfltOvrdInd_Click()
'              cmdAdd_Click()
'              cmdCloneThisPayee_Click()
'              cmdClose_Click()
'              cmdDelete_Click()
'              cmdNavigate_Click(ByRef pintIndex As Integer)
'              cmdUpdate_Click()
'              dtpPayePmtDt_Change()
'              fnAddRecord()
'              fnBindControlsToTableWrapper()
'              fnCalcAndLogMSIInfo(ByRef msiIn As StateInfo, strDesc As String)
'              fnCalcClaimForState(ByRef msiStateIn As StateInfo) As Boolean
'              fnCalcClaimInterest() As Boolean
'              fnClearControls()
'              fnGetCurrentIntRate(ByVal dtePayePmtDt As Date) As Double
'              fnGetFieldLabel(ByVal strControlName As String) As String
'              fnGetInterestRate(ByRef siRatesIn As StateInfo) As Currency
'              fnGetListOfStates() As ADODB.Recordset
'              fnGetStateInfo_InsdDthResStCd()
'              fnGetStateInfo_IssStCd()
'              fnGetStateInfo_Override()
'              fnGetStateInfo_PayeStCd()
'              fnInitializeCalcInfo(ByRef msiIn As StateInfo)
'              fnInitializeEditMode()
'              fnLoadCboPayeSsnTinTypCd()
'              fnLoadCbosForStates()
'              fnLoadControls()
'              fnLoadLpcLookup()
'              fnPerformLookup(ByRef lpcIn As LPLib.fpCombo)
'              fnPromptForRate(ByVal strPromptText As String) As Double
'              fnRefreshAllCombos()
'              fnResetStateRules()
'              fnSetAvailabilityOfControls(Optional ByVal bChangeFocus = True)
'              fnSetCommandButtons(ByVal bEnable As Boolean)
'              fnSetDefaultControlProperties()
'              fnSetFocusToFirstUpdateableField()
'              fnSetNavigationButtons(Optional ByVal bUnconditionalDisable As Boolean = False)
'              fnSetupScreenControls()
'              fnValidData() As Boolean
'              fnWarningData()
'              Form_Activate()
'              Form_Initialize()
'              Form_Load()
'              Form_QueryUnload(ByRef pintCancel As Integer, ByRef pintUnloadMode As Integer)
'              Form_Unload(ByRef pintCancel As Integer)
'              ipcPayeDthbPmtAmt_Change()
'              ipdPayeClmIntRt_Change()
'              ipdPayeWthldRt_Change()
'              ipmPayeSsnTinNum_Change()
'              iptPayeAddrLn1Txt_Change()
'              iptPayeAddrLn2Txt_Change()
'              iptPayeCareOfTxt_Change()
'              iptPayeCityNmTxt_Change()
'              iptPayeFullNm_Change()
'              iptPayeZipCd_Change()
'              lpcLookupName_Click()
'              lpcLookupName_GotFocus()
'              lpcLookupName_KeyDown(ByRef intKeyCode As Integer, ByRef intShift As Integer)
'              lpcLookupName_LostFocus()
'              TestStub1()
'              TestStub1Sub(siIn As StateInfo, curRate As Currency)
' Modified   :
' 03/26/01 BAW Cleaned with Total Visual CodeTools 2000
' 06/18/01 BAW Updated to avoid ADO Error 3001 when doing a Find on a Payee name
'              that contains an embedded single quote, e.g., O'Dell
' 10/11/01 BAW Additional changes to accommodate single quotes in Payee Name
' 01/2002  BAW Updated calcs to ignore scope, and calc interest 3 ways, using the method that
'              resulted in the highest interest as the "final way." This involved added a
'              Contract Issue State to the Payee screen as well. Also, removed
'              "#If gcfLOOKUP" stuff since we definitely want Lookup capability. (At one
'              time before v2.2 was released, we thought the performance might be too bad to keep it.)
'              Also, optimized per Project Analyzer (removing dead code, adding "$" to Mid/Space, etc.)
' Modified  : Berry Kropiwka 2019-09-27, added code from compact calc
' --------------------------------------------------
Option Explicit
Option Compare Binary

Private mstrScreenName As String

Private Const mclngMinFormWidth As Long = 10275
Private Const mclngMinFormHeight As Long = 8955

' The following constants identify, for fpCombo controls used as Lookups,
' which column is displayed in the Edit portion of the control (index = mcintDisplayCol_xxxx,
' where xxxx is the fpCombo control's name).
Private Const mcintDisplayCol_lpcLookupName     As Integer = 0

' These constants define the valid entries in the cboPayeSsnTinTypCd combobox.
Private Const mcstrPayeeIsABusiness             As String = "B"
Private Const mcstrPayeeIsAPerson               As String = "P"

' These constants define the masks for the ipmPayeSsnTinNum control
Private Const mcstrTINMask                      As String = "##-#######"
Private Const mcstrSSNMask                      As String = "###-##-####"
Private Const mcstrUnknownTinTypeMask           As String = "#########"

' These constants define the columns within the Lookup/Multi-column combo boxes.
' These are used to give a name to a given column of the fpCombo control so
' it can be referenced by name, not by number.
Private Const mcstrDisplayCol                   As String = "DISPLAY_COL"
Private Const mcstrPayeId                       As String = "PAYE_ID"
Private Const mcstrPayeFullNm                   As String = "PAYE_FULL_NM"

' mtWrapper is an instance of the table wrapper corresponding to the main table maintained by this form.
Private mtWrapper               As ctpyePayee

' Define a constant for each field that may get an error. This should match
' the text of that control's associated Label control.
Private Const mcstrIptPayeFullNmLabel       As String = "Name"                ' Editable only upon an Add
Private Const mcstrIptPayeCareOfTxtLabel    As String = "Care Of"
Private Const mcstrIptPayeAddrLn1TxtLabel   As String = "Address1"
Private Const mcstrIptPayeAddrLn2TxtLabel   As String = "Address2"
Private Const mcstrIptPayeCityNmTxtLabel    As String = "City"
Private Const mcstrCboPayeStCdLabel         As String = "State"                 ' Payee's Residence State at Insured's Death (short name)
Private Const mcstrIptPayeZipCdLabel        As String = "Zip"
Private Const mcstrIptPayeZip4CdLabel       As String = "Zip"
Private Const mcstrIpmPayeSsnTinNumLabel    As String = "TIN/SSN"
Private Const mcstrCboPayeSsnTinTypCdLabel  As String = "TIN Type"
Private Const mcstrChkPaye1099IndLabel      As String = "1099INT"      '' BZ4999 October 2013 Non US payee - SXS
'OBSOLETE Private Const mcstrCboContractIssueStateLabel As String = "Contract Issue State"    ' Insured's Residence State at Issue (short name)
Private Const mcstrIpdPayeWthldRtLabel      As String = "Withholding Rate"
Private Const mcstrChkPayeDfltOvrdIndLabel  As String = "Override"
Private Const mcstrCboCalcStCdLabel         As String = "Calc State"
Private Const mcstrIpdPayeClmIntRtLabel     As String = "Interest Rate"
Private Const mcstrDtpPayePmtDtLabel        As String = "Date of Payment"
Private Const mcstrIpdPayeIntDaysPdNumLabel As String = "Days of Interest Paid"
Private Const mcstrIpcPayeDthbPmtAmtLabel   As String = "DB Payment"
Private Const mcstrIpcPayeClmIntAmtLabel    As String = "Claim Interest"
Private Const mcstrIpcPayeWthldAmtLabel     As String = "Interest Withheld"
Private Const mcstrIpcPayeClmPdAmtLabel     As String = "Total"

Private Const mcstrTxtInsdDthResStCd_UsedInAutoCalcLabel As String = "Issue State"
' Labels from Insured screen
Private Const mcstrDtpClmInsdDthDtLabel                 As String = "Date Of Death"
Private Const mcstrDtpClmProofDtLabel                   As String = "Date Of Proof"

Private Const mcstrTxtPayeIDLabel                       As String = "Payee ID"

'Dim mrstLookup As ADODB.Recordset
'Dim mrstPayee As ADODB.Recordset
Dim mfrmMyInsuredForm           As Form

' mbInLookupMode determines whether the user is in the process of doing a search using the Lookup ComboBox
Private mbInLookupMode                  As Boolean

' mbInAddMode determines whether the user has begun the process of adding a new record to the table.
' Note that Add mode is independent of Update mode
Private mbInAddMode                     As Boolean

Private mctlFirstUpdateableField_Add    As Control
Private mctlFirstUpdateableField_Upd    As Control

' The following field (mcurTotalWithheld) is a "cousin" to ipcPayeWthldAmt
' that appears on-screen. ipcPayeWthldAmt is formatted with the Format( )
' function to display as ($$$.$$) since it reduces the total amount paid
' for a claim. However, mcurTotalWithheld is the unformatted equivalent,
' unformatted so that the value -- with its sign preserved -- can be
' stored.  When the Format( ) function adds "(" and ")" around a string and
' that string is stored, it's regarded as a negative number. Yech!
Dim mcurTotalWithheld                   As Currency

Dim msiInsdDthResStCd                   As StateInfo
Dim msiPayeStCd                         As StateInfo
Dim msiIssStCd                          As StateInfo
Dim msiOverride                         As StateInfo
Dim msiCalcStCd                         As StateInfo
Dim msiCompactCalc                      As StateInfo


' m_bIsDirty corresponds to the public property called IsDirty.
' All maintenance screens should have this field and that property! When True, it indicates
' that the user has made --but not yet saved-- changes to a record. The MDI form will query
' this property if the user opens the File menu, since the Exit option should be disabled if
' any form has outstanding changes.
' Be sure to use this variable's corresponding Property Let to change its value.
' Do **NOT** set m_bIsDirty itself, as this will using the Property Let proc will
' ensure the Close button caption is always synchronized with the value of the property.
Private m_bIsDirty                      As Boolean

' The following UDT is used to suport the "Clone This Payee" functionality
Private Type udtPayeeClone
    PayeFullNm                                          As String
    PayeCareOfTxt                                       As String
    PayeAddrLn1Txt                                      As String
    PayeAddrLn2Txt                                      As String
    PayeCityNmTxt                                       As String
    PayeStCd                                            As String
    PayeZipCd                                           As String
    PayeZip4Cd                                          As String
    PayeSsnTinNum                                       As String
    PayeSsnTinTypCd                                     As String
    PayePmtDt                                           As Date
    PayeDthbPmtAmt                                      As Double
    ClmId                                               As Long
    paye_1099int_ind                                    As String   '' BZ4999 October 2013 Non US payee - SXS
    ' Calc-related fields
    PayeStCd_UsedInAutoCalc                             As String
    PayeStCdSpecialInstructions_UsedInAutoCalc          As String
    bClmForResDthInd_UsedInAutoCalc                     As Boolean
    InsdDthResStCd_UsedInAutoCalc                       As String
    InsdDthResStCdSpecialInstructions_UsedInAutoCalc    As String
    IssStCd_UsedInAutoCalc                              As String
    IssStCdSpecialInstructions_UsedInAutoCalc           As String
    bClmForCompactCalc_UsedInAutoCalc                   As Boolean
End Type

'MME START WRUS 4999
Public DblScreenDBPaymentValue As Double
'MME END WRUS 4999

Private m_AdmPolicySystem        As String
Private m_upcPayeeClone          As udtPayeeClone
Private Const m_GROUP_ADMIN_SYS                   As String = "GROUP"

'Private Const variable for the compact filling state code.  When setting the state to this variable in fnCalcClaimInterest when we calcuate the msiCompactCalc it
'   will using the state rule for compact filling.
Private Const cstCompactFilling As String = "YY"


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
    Const cstrCancel        As String = "&Cancel"
    Const cstrClose         As String = "&Close"
    
    On Error GoTo PROC_ERR

    m_bIsDirty = bValue

    ' Adjust Close button caption accordingly. Do it conditionally, to avoid
    ' flickering when the user does a lot of quick scrolling.
    If bValue Then
        If cmdClose.Caption <> cstrCancel Then
            cmdClose.Caption = cstrCancel
        End If
    Else
        If cmdClose.Caption <> cstrClose Then
            cmdClose.Caption = cstrClose
        End If
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
Private Sub cboCalcStCd_Click()
    ' Comments  : Sets a flag to indicate the current record has been
    '             edited, and thus Update button becomes enabled
    ' Parameters: N/A
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "cboCalcStCd_Click"
 

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    fnInitializeEditMode
    
    ' Populate msiOverride structure with data from the STATE_RULE_T
    ' row that matches the Calc State.
    fnGetStateInfo_Override
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
Private Sub cboPayeSsnTinTypCd_Click()
    ' Comments  :
    ' Parameters:  -
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "cboPayeSsnTinTypCd_Click"
    Dim strSavedSSN       As String

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    fnInitializeEditMode
    
    ' Reset the mask and format the value accordingly
    strSavedSSN = ipmPayeSsnTinNum.UnFmtText
    ipmPayeSsnTinNum.Mask = vbNullString
    If cboPayeSsnTinTypCd.Text = mcstrPayeeIsABusiness Then
        ipmPayeSsnTinNum.Mask = mcstrTINMask
        ipmPayeSsnTinNum.UnFmtText = fnSSNTIN_AddDash(strIn:=strSavedSSN, bIsTin:=True)
    ElseIf cboPayeSsnTinTypCd.Text = mcstrPayeeIsAPerson Then
        ipmPayeSsnTinNum.Mask = mcstrSSNMask
        ipmPayeSsnTinNum.UnFmtText = fnSSNTIN_AddDash(strIn:=strSavedSSN, bIsTin:=False)
    Else
        ipmPayeSsnTinNum.Mask = mcstrUnknownTinTypeMask
        ipmPayeSsnTinNum.UnFmtText = strSavedSSN
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
Private Sub cboPayeStCd_Click()
    ' Comments  : Sets a flag to indicate the current record has been
    '             edited, and thus Update button becomes enabled
    ' Parameters: N/A
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "cboPayeStCd_Click"
 

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    fnInitializeEditMode
    
    txtPayeStCd_UsedInAutoCalc.Text = cboPayeStCd.Text
    ' Populate msiPayeStCd structure with data from the STATE_RULE_T
    ' row that matches the Payee State.
    '' BZ4999 October 2013 Non US payee - SXS
    ChkPaye1099Ind.value = "1"
    If cboPayeStCd.Text = "ZZ" Then
       ChkPaye1099Ind.value = "0"
       ChkPaye1099Ind.ForeColor = -2147483632  ''Button Shadow
       iptPayeZip4Cd.Text = "     "
       iptPayeZipCd.Text = "    "
       ChkPaye1099Ind.Enabled = False
    Else
       ChkPaye1099Ind.ForeColor = -2147483630 '' Button Text
       fnGetStateInfo_PayeStCd
        ChkPaye1099Ind.Enabled = True
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
'''''''''''''''''''''''''''''  '' BZ4999 October 2013 Non US payee - SXS
'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub ChkPaye1099Ind_Click()

    ' Comments  : Sets a flag to indicate the current record has been
    '             edited, and thus Update button becomes enabled
    ' Parameters: N/A
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "ChkPaye_1099INd_Click"
 
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


''''''''''''''''''''''''''''''''''''''
'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub chkPayeDfltOvrdInd_Click()
    ' Comments  : Sets a flag to indicate the current record has been
    '             edited, and thus Update button becomes enabled
    ' Parameters: N/A
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "chkPayeDfltOvrdInd_Click"
 
    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    fnInitializeEditMode
    
     If chkPayeDfltOvrdInd.value = vbChecked Then
        fnEnableDisableControl ctlIn:=cboCalcStCd, bEnable:=True
        fnEnableDisableControl ctlIn:=ipdPayeClmIntRt, bEnable:=True
     Else
        fnEnableDisableControl ctlIn:=cboCalcStCd, bEnable:=False
        fnEnableDisableControl ctlIn:=ipdPayeClmIntRt, bEnable:=False
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
Private Sub cmdCloneThisPayee_Click()
    ' Comments  : If the user clicked this button, save current fields
    '             to a udt, simulate an cmdAdd_Click event, and use the
    '             udt to pre-fill the new record.
    '
    ' Parameters: N/A
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "cmdCloneThisPayee_Click"

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    ' Save current values
    With m_upcPayeeClone
        .PayeFullNm = iptPayeFullNm.Text
        .PayeCareOfTxt = iptPayeCareOfTxt.Text
        .PayeAddrLn1Txt = iptPayeAddrLn1Txt.Text
        .PayeAddrLn2Txt = iptPayeAddrLn2Txt.Text
        .PayeCityNmTxt = iptPayeCityNmTxt.Text
        .PayeStCd = cboPayeStCd.Text
        .PayeZipCd = iptPayeZipCd.Text
        .PayeZip4Cd = iptPayeZip4Cd.Text
        .PayeSsnTinNum = ipmPayeSsnTinNum.UnFmtText
        .PayeSsnTinTypCd = cboPayeSsnTinTypCd.Text
        .paye_1099int_ind = (ChkPaye1099Ind.value = vbChecked)  '' BZ4999 October 2013 Non US payee - SXS
        .PayePmtDt = dtpPayePmtDt.value
        .PayeDthbPmtAmt = ipcPayeDthbPmtAmt.value
        .ClmId = mtWrapper.ClmId
        ' Calc-related fields
        .PayeStCd_UsedInAutoCalc = txtPayeStCd_UsedInAutoCalc.Text
        .PayeStCdSpecialInstructions_UsedInAutoCalc = txtPayeStCdSpecialInstructions_UsedInAutoCalc.Text
        .bClmForResDthInd_UsedInAutoCalc = (chkClmForResDthInd_UsedInAutoCalc.value = vbChecked)
        .InsdDthResStCd_UsedInAutoCalc = txtInsdDthResStCd_UsedInAutoCalc.Text
        .InsdDthResStCdSpecialInstructions_UsedInAutoCalc = txtInsdDthResStCdSpecialInstructions_UsedInAutoCalc.Text
        .IssStCd_UsedInAutoCalc = txtIssStCd_UsedInAutoCalc.Text
        .IssStCdSpecialInstructions_UsedInAutoCalc = txtIssStCdSpecialInstructions_UsedInAutoCalc.Text
    End With
    
    ' Hide updates to the window until we're done. This avoids ugly screen flickering
    fnWindowLock Me.hWnd
    
    fnAddRecord
    
    With m_upcPayeeClone
        iptPayeFullNm.Text = .PayeFullNm
        iptPayeCareOfTxt.Text = .PayeCareOfTxt
        iptPayeAddrLn1Txt.Text = .PayeAddrLn1Txt
        iptPayeAddrLn2Txt.Text = .PayeAddrLn2Txt
        iptPayeCityNmTxt.Text = .PayeCityNmTxt
        cboPayeStCd.Text = .PayeStCd
        iptPayeZipCd.Text = .PayeZipCd
        iptPayeZip4Cd.Text = .PayeZip4Cd
        If .paye_1099int_ind Then
           ChkPaye1099Ind.value = vbChecked
        Else
           ChkPaye1099Ind.value = vbUnchecked
        End If
        ipmPayeSsnTinNum.UnFmtText = .PayeSsnTinNum
        cboPayeSsnTinTypCd.Text = .PayeSsnTinTypCd
        dtpPayePmtDt.value = .PayePmtDt
        ipcPayeDthbPmtAmt.value = .PayeDthbPmtAmt
        mtWrapper.ClmId = .ClmId
        ' Calc-related fields
        txtPayeStCd_UsedInAutoCalc.Text = .PayeStCd_UsedInAutoCalc
        txtPayeStCdSpecialInstructions_UsedInAutoCalc.Text = .PayeStCdSpecialInstructions_UsedInAutoCalc
        If .bClmForResDthInd_UsedInAutoCalc Then
            chkClmForResDthInd_UsedInAutoCalc.value = vbChecked
        Else
            chkClmForResDthInd_UsedInAutoCalc.value = vbUnchecked
        End If
        txtInsdDthResStCd_UsedInAutoCalc.Text = .InsdDthResStCd_UsedInAutoCalc
        txtInsdDthResStCdSpecialInstructions_UsedInAutoCalc.Text = .InsdDthResStCdSpecialInstructions_UsedInAutoCalc
        txtIssStCd_UsedInAutoCalc.Text = .IssStCd_UsedInAutoCalc
        txtIssStCdSpecialInstructions_UsedInAutoCalc.Text = .IssStCdSpecialInstructions_UsedInAutoCalc
    End With
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    
    ' Clean-up statements go here
    fnWindowUnlock

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
    Dim intButtonClicked             As Integer
    Dim lngReturnValue               As Long
    Dim strACF2                      As String
    Dim hrgHourglass                 As chrgHourglass

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    ' .......................................................................
    ' Make sure the user really, really, really wants to delete this record.
    ' .......................................................................
    ' 3002 = Are you sure you want to delete this record?
    intButtonClicked = gerhApp.ReportNonFatal(vbObjectError + gcRES_ALRT_OK_TO_DELETE_RECORD, _
                           mstrScreenName & gcstrDOT & cstrCurrentProc)

    Me.Refresh
    
    If (intButtonClicked = vbNo) Or (intButtonClicked = gcintClickedCloseButton) Then
        GoTo PROC_EXIT
    End If
    
    ' .......................................................................
    ' Proceed with the Delete.
    ' *  If another user has updated the record, we don't care. No message
    '    should be generated, and the delete should proceed.
    ' *  If another user has deleted the record, display a message to
    '    that effect and then show the record whose key value immediately
    '    preceeds the record we wanted to delete (or if there are now
    '    no other records in the table, go into Add mode).
    ' *  If no other user did anything with this record, then just
    '    delete it and then show the record whose key value immediately
    '    preceeds the record we wanted to delete (or if there are now
    '    no other records in the table, go into Add mode).
    '
    ' Note that .GetRelativeRecord( ) can be called directly but is also
    ' called via the .DeleteRecord( ) method. In both cases, it
    ' refreshes the Lookup recordset (m_rstLookup) before positioning
    ' to the desired relative record.
    '
    ' Anytime the Lookup recordset is refreshed, we need to reload the
    ' vfgLookup VSFlexGrid control.
    ' .......................................................................
    Set hrgHourglass = New chrgHourglass
    hrgHourglass.value = True
     
    ' Hide updates to the window until we're done. This avoids ugly screen flickering
    fnWindowLock Me.hWnd
     
    With mtWrapper
        lngReturnValue = .CheckForAnotherUsersChanges(ewoDelete, strACF2)
        If lngReturnValue = vbObjectError + gcRES_INFO_ANOTHER_USER_DELETED Then
            ' Another user has deleted the record that *this* user is trying to
            ' delete. So, display a message to that effect, refresh the Lookup
            ' recordset and then show the record whose key value immediately
            ' preceeds the record this user wanted to delete.
            gerhApp.ReportNonFatal vbObjectError + gcRES_INFO_ANOTHER_USER_DELETED, _
                                   mstrScreenName & gcstrDOT & cstrCurrentProc
            ' Can't use the GetPayeFullNmFromPayeID( ) method since the PAYEE_T row
            ' doesn't exist and hence a "-2147217900" (Payee ID does not exist) error
            ' throws things off.
            .GetRelativeRecord .PayeFullNm, epdPreviousRecord
        Else
            ' If another user updated the record *this* user is trying to delete,
            ' we don't care. No message should be generated and the delete should
            ' proceed as if no other user did anything to this record.
            '
            ' If no other user did anything with this record, then delete it,
            ' refresh the Lookup recordset, and then show the record whose
            ' key value immediately preceeds the record this user wanted to
            ' delete.
            .DeleteRecord
        End If

        ' Repopulate the all Lookup and ComboBox controls so
        ' they reflects this and other users' changes.
        fnRefreshAllCombos


        ' If there are no records now in the table (based on this user's or
        ' another user's actions), then go into Add mode. Otherwise, display
        ' the now-current record. fnLoadControls will set the navigation buttons
        ' and "record x of y" label as appropriate.
        If .LookupRecordCount > 0 Then
            fnLoadControls
            fnSetCommandButtons True
        Else
            fnAddRecord
        End If
    End With
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0

    ' Clean-up statements go here
    If Not (hrgHourglass Is Nothing) Then
        hrgHourglass.value = False
    End If
    fnFreeObject hrgHourglass
    fnWindowUnlock

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

        If bDebugAppTermination Then
            Debug.Print "   Turning off Update mode (#1) in " & mstrScreenName & gcstrDOT & cstrCurrentProc
        End If
        IsDirty = False
    
        If (.CurrentLookupRecordNumber = adPosBOF) Or _
        (.CurrentLookupRecordNumber = adPosEOF) Or _
        (.CurrentLookupRecordNumber = adPosUnknown) Then
            gerhApp.ReportNonFatal vbObjectError + gcRES_INFO_TABLE_IS_EMPTY, _
                                   mstrScreenName & gcstrDOT & cstrCurrentProc
            fnAddRecord
        Else
            ' Note that the Lookup controls' selection is no longer synchronized
            ' with the table wrapper's CurrentLookupRecordNumber. In other words,
            ' the CurrentLookupRecordNumber may indicate we're on the 5th record and,
            ' by virtue of fnLoadControls being called following each navigation, that should
            ' the same record that is currently displayed on-screen. However, the Lookup
            ' controls themselves are not necessarily *itself* positioned to the 5th record.
            ' The total number of entries in that control, however, should jive with the
            ' table wrapper's LookupRecordCount property.

            ' Load current record's properties to form's controls, reset navigation buttons and set "rec x of y" label
            fnLoadControls
            If bDebugAppTermination Then
                Debug.Print "   Turning off Update mode (#2) in " & mstrScreenName & gcstrDOT & cstrCurrentProc
            End If
            
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
    ' Comments:    This function handles updating an existing record or, if in Add mode,
    '              the adding of a new record. It is called when the user clicks the
    '              Update button, as well as by Form_QueryUnload when the user
    '              attempts to close the form while edits are outstanding.
    ' Parameters:  -
    ' Modified  :
    ' --------------------------------------------------
    Const cstrCurrentProc           As String = "cmdUpdate_Click"
    Dim lngReturnValue              As Long
    Dim strACF2                     As String
    Dim hrgHourglass                As chrgHourglass

    On Error GoTo PROC_ERR

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    Set hrgHourglass = New chrgHourglass
    hrgHourglass.value = True

    If Not (fnValidData()) Then
        GoTo PROC_EXIT
    End If

    ' Update the Table wrapper's properties with the screen values
    With mtWrapper
        .PayeFullNm = iptPayeFullNm.Text
        .PayeCareOfTxt = iptPayeCareOfTxt.Text
        .PayeAddrLn1Txt = iptPayeAddrLn1Txt.Text
        .PayeAddrLn2Txt = iptPayeAddrLn2Txt.Text
        .PayeCityNmTxt = iptPayeCityNmTxt.Text
        .PayeStCd = cboPayeStCd.Text
        .PayeZipCd = iptPayeZipCd.Text
        .PayeZip4Cd = iptPayeZip4Cd.Text
        
        .PayeSsnTinNum = ipmPayeSsnTinNum.UnFmtText      ' Use .UnFmtText to get rid of mask characters in fpMask control
        ' cboPayeSsnTinTypCd corresponds to a Nullable field, so accommodate Nulls
        If cboPayeSsnTinTypCd.Text = gcstrBlankEntry Then
            .PayeSsnTinTypCd = vbNullString
        Else
            .PayeSsnTinTypCd = cboPayeSsnTinTypCd.Text
        End If
         '' BZ4999 October 2013 Non US payee - SXS
         .Paye1099INTInd = (ChkPaye1099Ind.value = vbChecked)
        .PayeWthldRt = ipdPayeWthldRt.value

        .PayeDfltOvrdInd = (chkPayeDfltOvrdInd.value = vbChecked)
        .PayePmtDt = dtpPayePmtDt.value
        .PayeDthbPmtAmt = ipcPayeDthbPmtAmt.value
        
        .LstUpdtUserId = gconAppActive.LastLogOnUserID
        .LstUpdtDtm = Now
    End With
    
    ' These will propagate back an error if the Insert/Update failed.
    If mbInAddMode Then
        fnCalcClaimInterest
        
        ' Update wrapper with values calculated or possibly affected by fnCalcClaimInterest( )
        With mtWrapper
            .PayeIntDaysPdNum = ipdPayeIntDaysPdNum.value
            .PayeClmIntAmt = ipcPayeClmIntAmt.value
            .PayeWthldAmt = ipcPayeWthldAmt.value
            .PayeClmPdAmt = ipcPayeClmPdAmt.value
            .CalcStCd = cboCalcStCd.Text
            .PayeClmIntRt = ipdPayeClmIntRt.value
        End With

        ' Add the record, refresh the lookup recordset and reposition
        ' to the record just added
        mtWrapper.AddRecord
        ' Turn off Add mode since the Add was successful
        mbInAddMode = False
        ' Repopulate the all Lookup and ComboBox controls so
        ' they reflects this and other users' changes.
        fnRefreshAllCombos
        ' This **must** be done as the user leaves Add mode, so that the key fields
        ' will now be protected to prevent the user from being able to edit them.
        ' Editing a key field is allowed only when in Add mode.
        fnSetAvailabilityOfControls
    Else
        With mtWrapper
            ' Determine whether another user updated or deleted the record about to be updated.
            ' Note: this multi-user checking is performed on an Update but not an Add.
            lngReturnValue = .CheckForAnotherUsersChanges(ewoUpdate, strACF2)

            If lngReturnValue = vbObjectError + gcRES_INFO_ANOTHER_USER_DELETED_NOT_SAVED Then
                gerhApp.ReportNonFatal vbObjectError + gcRES_INFO_ANOTHER_USER_DELETED_NOT_SAVED, _
                                       mstrScreenName & gcstrDOT & cstrCurrentProc
                ' Discard *this* user's pending changes and show the previous record.
                ' Can't use the GetPayeFullNmFromPayeID( ) method since the PAYEE_T row
                ' doesn't exist and hence a "-2147217900" (Payee ID does not exist) error
                ' throws things off.
                .GetRelativeRecord .PayeFullNm, epdPreviousRecord
                
            ElseIf lngReturnValue = vbObjectError + gcRES_INFO_ANOTHER_USER_UPDATED_DISCARDED Then
                gerhApp.ReportNonFatal vbObjectError + gcRES_INFO_ANOTHER_USER_UPDATED_DISCARDED, _
                                       mstrScreenName & gcstrDOT & cstrCurrentProc, _
                                       Trim$(strACF2)
                ' Discard *this* user's pending changes by re-retrieving the current record
                ' as it currently looks on the database and refreshing the lookup recordset
                ' Can't use the GetPayeFullNmFromPayeID( ) method since the PAYEE_T row
                ' doesn't exist and hence a "-2147217900" (Payee ID does not exist) error
                ' throws things off.
                .GetRelativeRecord .PayeFullNm, epdSameRecord
            Else
                fnCalcClaimInterest
                
                ' Update wrapper with values calculated or possibly affected by fnCalcClaimInterest( )
                With mtWrapper
                    .PayeIntDaysPdNum = ipdPayeIntDaysPdNum.value
                    .PayeClmIntAmt = ipcPayeClmIntAmt.value
                    .PayeWthldAmt = ipcPayeWthldAmt.value
                    .PayeClmPdAmt = ipcPayeClmPdAmt.value
                    .CalcStCd = cboCalcStCd.Text
                    .PayeClmIntRt = ipdPayeClmIntRt.value
                End With
                
                ' Update the record with this user's pending changes, refresh the lookup
                ' recordset and reposition to the record just updated
                mtWrapper.UpdateRecord
            End If

            ' Turn off Update mode since the Update was either successful or abandoned
            If bDebugAppTermination Then
                Debug.Print "   Turning off Update mode in " & mstrScreenName & gcstrDOT & cstrCurrentProc
            End If
            IsDirty = False

            ' Repopulate the all Lookup and ComboBox controls so
            ' they reflects this and other users' changes.
            fnRefreshAllCombos
        End With
    End If

    ' Do an immediate repaint. This allows the Insured screen to be redrawn BEFORE all
    ' the work of requerying and repainting is started. When the requerying/repainting is done,
    ' only small parts of the screen (not the whole screen) will need to be repainted. This
    ' eliminates the user seeing a very slow repainting.
    Me.Refresh

    If mtWrapper.LookupRecordCount > 0 Then
        ' Ensure the on-screen controls reflect the record just added/updated, in case the
        ' DBMS altered it in some way, e.g., determining an Identity column value and
        ' getting the most up-to-date Last Updated info. This also sets the navigation
        ' buttons and updates the "record x of y" label
        fnLoadControls
        
        ' Display the message indicating the calc was overriden if the Override
        ' checkbox is still selected.
        lblWarningAboutOverride.Visible = (chkPayeDfltOvrdInd.value = vbChecked)
        
        fnSetCommandButtons True
        
        If ((Len(ipmPayeSsnTinNum.UnFmtText) = 0) Or (ipmPayeSsnTinNum.UnFmtText = "000000000")) And _
           ((CDbl(ipdPayeWthldRt.value) = 0) Or Len(ipdPayeWthldRt.value) = 0) And _
            (CDbl(ipcPayeClmIntAmt.UnFmtText) >= msiCalcStCd.StrlIntRptgFlrAmt) Then
                ' gcRES_WARN_GET_TIN_BEFORE_PAYING_INT (2004) = This claims requires a certified @@1 to avoid withholding.
                '                                               Make sure you don't pay interest until it has been received.
                ' Per Michelle Wilkosky, this warning should gen'd if the calculated interest equals or exceeds the
                '      state reporting floor AND (either the TIN was not supplied or was set to all zeroes) AND
                '      (either the Withholding Rate was not supplied or was set to 0)
                gerhApp.ReportNonFatal vbObjectError + gcRES_WARN_GET_TIN_BEFORE_PAYING_INT, _
                                       mstrScreenName & gcstrDOT & cstrCurrentProc, _
                                       mcstrIpmPayeSsnTinNumLabel
        End If
    Else
        fnAddRecord
    End If
    
    Me.Refresh
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0

    ' Clean-up statements go here
    If Not (hrgHourglass Is Nothing) Then
        hrgHourglass.value = False
    End If
    fnFreeObject hrgHourglass
    
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
Private Sub dtpPayePmtDt_Change()
    ' Comments  : Since this field was just changed, reset
    '             Enabled property on command and navigation
    '             buttons as appropriate given that the user
    '             is in the middle of updating a record.
    ' Parameters:  -
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "dtpPayePmtDt_Change"

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    fnInitializeEditMode
    
    fnResetStateRules
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

    ' All we do here is display an empty record. The cmdUpdate_Click event
    ' handler actually does the add when it sees that mbInAddMode=True.
    ' Adds and Updates are treated very nearly the same in that event handler!

    mbInAddMode = True

    ' Display empty or initialized values for on-screen controls
    fnClearControls
                            
    ' We can't populate the StateInfo structures associated with the various State Codes
    ' used in the calculations since, with the 06/2003 release of the system, the state
    ' rules now can vary by Date of Payment. So, the dtpPayePmtDt_Change event handler
    ' populates them. Here, however, we can set the State Codes themselves and empty out
    ' their associated Special Instructions.
    ' 1. Insured State of Residence at time of death (carried over from Insured screen)
    txtInsdDthResStCd_UsedInAutoCalc.Text = mfrmMyInsuredForm.InsuredInsdDthResStCd
    txtInsdDthResStCdSpecialInstructions_UsedInAutoCalc.Text = vbNullString
    If mfrmMyInsuredForm.InsuredClmForResDthInd Then
        Me.chkClmForResDthInd_UsedInAutoCalc = vbChecked
    Else
        Me.chkClmForResDthInd_UsedInAutoCalc = vbUnchecked
    End If
    ' 2. Contract Issue State (carried over from Insured screen)
    txtIssStCd_UsedInAutoCalc.Text = mfrmMyInsuredForm.InsuredIssStCd
    txtIssStCdSpecialInstructions_UsedInAutoCalc.Text = vbNullString
    ' 3. Payee Residence State at time of death
    txtPayeStCd_UsedInAutoCalc.Text = cboPayeStCd.Text
    txtPayeStCdSpecialInstructions_UsedInAutoCalc.Text = vbNullString
                            
    If bDebugAppTermination Then
        Debug.Print "   Turning off Update mode (#1) in " & mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    IsDirty = False

    ' Enable and set focus to key field(s) so the user can specify a value.
    ' This **must** be done as the user goes into Add mode, so they can specify
    ' the key(s) for the record they're adding.
    fnSetAvailabilityOfControls

    ' Restrike "Record x of y" to reflect pending Add. Can't call fnShowRecordPosition
    ' since it is based on a recordset's AbsolutePosition which, in unbound /disconnected mode,
    ' isn't set appropriately.
    lblRecordPosition = "Record ? of " & mtWrapper.LookupRecordCount

    If bDebugAppTermination Then
        Debug.Print "   Turning off Update mode (#2) in " & mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    IsDirty = False
    fnSetCommandButtons False

    fnSetNavigationButtons bUnconditionalDisable:=True
    
    ' Make sure first field gets the focus. Note, when Add mode is triggered
    ' from Form_Load, this statement accomplishes nothing: the control isn't yet visible,
    ' so it can't receive the focus. This is why Form_Activate must also call this function.
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
Private Sub fnBindControlsToTableWrapper()
'--------------------------------------------------------------------------
    ' Procedure:   fnBindControls
    ' Description: Binds the on-screen controls to the table wrapper class
    '              properties with which they are associated. This is done so
    '              various control properties can be set based on meta data
    '              gathered by the table wrapper class.
    '
    ' Params:      N/A
    ' Returns:     N/A
    ' Date:        04/04/2002
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc          As String = "fnBindControlsToTableWrapper"
    On Error GoTo PROC_ERR
 
 
    iptPayeFullNm.Tag = "PayeFullNm"
    iptPayeCareOfTxt.Tag = "PayeCareOfTxt"
    iptPayeAddrLn1Txt.Tag = "PayeAddrLn1Txt"
    iptPayeAddrLn2Txt.Tag = "PayeAddrLn2Txt"
    iptPayeCityNmTxt.Tag = "PayeCityNmTxt"
    cboPayeStCd.Tag = "PayeStCd"
    iptPayeZipCd.Tag = "PayeZipCd"
    iptPayeZip4Cd.Tag = "PayeZip4Cd"
    ChkPaye1099Ind.Tag = "Paye1099Ind"
    ipmPayeSsnTinNum.UnFmtText = "PayeSsnTinNum"
    cboPayeSsnTinTypCd.Tag = "PayeSsnTinTypCd"
    ipdPayeWthldRt.Tag = "PayeWthldRt"
    ipdPayeClmIntRt.Tag = "PayeClmIntRt"
    chkPayeDfltOvrdInd.Tag = "PayeDfltOvrdInd"
    cboCalcStCd.Tag = "CalcStCd"
    dtpPayePmtDt.Tag = "PayePmtDt"
    ipdPayeIntDaysPdNum.Tag = "PayeIntDaysPdNum"
    ipcPayeDthbPmtAmt.Tag = "PayeDthbPmtAmt"
    ipcPayeClmIntAmt.Tag = "PayeClmIntAmt"
    ipcPayeWthldAmt.Tag = "PayeWthldAmt"
    ipcPayeClmPdAmt.Tag = "PayeClmPdAmt"
    
'!TODO! - ClmId too?
'!TODO! - PayeId too?
    
    ' LstUpdDtm     isn't shown on-screen
    ' LstUpdUserId  isn't shown on-screen
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
Private Sub fnCalcAndLogMSIInfo(ByRef msiIn As StateInfo, ByVal strDesc As String, _
    Optional ByVal bUseSuppliedRate As Boolean = False)
    ' Comments  : This function will calculate and log info from the specified State Info structure
    '             to the application log file
    ' Parameters: None
    ' Returns   : True, if successful; False otherwise
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc   As String = "fnCalcAndLogMSIInfo"
    Const cstrDec11_5       As String = "#####0.00000"
    Const cstrCurrency      As String = "Currency"

    With msiIn
        .CalculationInfo = "The claim interest was calculated based on rates " & _
            "for the " & strDesc & ". "
        ' The next statement does the calc, and updates the msiIn structure with the results of the calc
        fnCalcClaimForState msiIn, strDesc, bUseSuppliedRate
        fnLogWrite " ", cstrCurrentProc
        fnLogWrite strDesc & ":", cstrCurrentProc
        fnLogWrite "  State:                          " & .StCd, cstrCurrentProc
        fnLogWrite "  Line-of-business:               " & .LobCd, cstrCurrentProc
        fnLogWrite "  Rule Effective Date:            " & fnZLSIfNull(.StrlEffDt), cstrCurrentProc
        fnLogWrite "  Rule End Date:                  " & fnZLSIfNull(.StrlEndDt), cstrCurrentProc
        fnLogWrite "  Interest Required Date Type:    " & .ReqdIdtypCd, cstrCurrentProc
        fnLogWrite "  Interest Required Offset:       " & .StrlIntReqdOfstNum, cstrCurrentProc
        fnLogWrite "  Interest Calculation Date Type: " & .CalcIdtypCd, cstrCurrentProc
        fnLogWrite "  Interest Calculation Offset:    " & .StrlIntCalcOfstNum, cstrCurrentProc
        fnLogWrite "  Interest Rule Code:             " & .IruleCd, cstrCurrentProc
        fnLogWrite "  Interest Rule Amount:           " & Format$(fnZeroIfNull(.StrlIntRuleAmt), cstrDec11_5), cstrCurrentProc
        fnLogWrite "  Reporting Floor Amount:         " & Format$(.StrlIntRptgFlrAmt, "Currency"), cstrCurrentProc
        fnLogWrite gcstrBlankEntry, cstrCurrentProc
        fnLogWrite "  SpecialInstructions:            " & .StrlSpclInstrTxt, cstrCurrentProc
        fnLogWrite "  Figured From Date:              " & CStr(DateValue(.FiguredFromDate)), cstrCurrentProc
        fnLogWrite "  PayablePeriodEndDate:           " & CStr(DateValue(.PayablePeriodEndDate)), cstrCurrentProc
        fnLogWrite "  InterestRateToUse:              " & Format$(.InterestRateToUse, cstrDec11_5), cstrCurrentProc
        fnLogWrite gcstrBlankEntry, cstrCurrentProc
        fnLogWrite "  NbrOfDaysToPayInterest:         " & .NbrOfDaysToPayInterest, cstrCurrentProc
        fnLogWrite "  ClaimInterest:                  " & Format$(.ClaimInterestAmt, cstrCurrency), cstrCurrentProc
        fnLogWrite "  Withheld:                       " & Format$(.WithheldAmt, cstrCurrency), cstrCurrentProc
        fnLogWrite "  TotalForThisPayee:              " & Format$(.TotalForThisPayee, cstrCurrency), cstrCurrentProc
        fnLogWrite "  CalculationInfo:                " & .CalculationInfo, cstrCurrentProc
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
Private Function fnCalcClaimInterest() As Boolean
    ' Comments  : This function will calculate Claim Interest
    '             for this Payee
    ' Parameters: None
    ' Returns   : True, if successful; False otherwise
    ' Modified  :
    ' Modified  : Berry Kropiwka 2019-09-27, added code for compact calc
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnCalcClaimInterest"

    fnLogWrite gcstrBlankEntry, cstrCurrentProc
    fnLogWrite "---New Calc---", cstrCurrentProc
    fnLogWrite "Inputs:", cstrCurrentProc
    fnLogWrite "  Claim Number:                   " & mfrmMyInsuredForm.InsuredClmNum, cstrCurrentProc
    fnLogWrite "  Payee:                          " & iptPayeFullNm.Text, cstrCurrentProc
    fnLogWrite "  Insured Foreign Res. at Death:  " & (chkClmForResDthInd_UsedInAutoCalc.value = vbChecked), cstrCurrentProc
    fnLogWrite "  Insured Residence State:        " & mfrmMyInsuredForm.InsuredInsdDthResStCd, cstrCurrentProc
    fnLogWrite "  Contract Issue State:           " & mfrmMyInsuredForm.InsuredIssStCd, cstrCurrentProc
    fnLogWrite "  Date of Proof:                  " & CStr(DateValue(mfrmMyInsuredForm.InsuredClmProofDt)), cstrCurrentProc
    fnLogWrite "  Date of Death:                  " & CStr(DateValue(mfrmMyInsuredForm.InsuredClmInsdDthDt)), cstrCurrentProc
    fnLogWrite "  Date of Payment:                " & CStr(DateValue(dtpPayePmtDt.value)), cstrCurrentProc
    fnLogWrite "  Payment:                        " & ipcPayeDthbPmtAmt.Text, cstrCurrentProc
    fnLogWrite "  Withholding Percent:            " & ipdPayeWthldRt.Text, cstrCurrentProc
    fnLogWrite "  Calculation Override:           " & (chkPayeDfltOvrdInd.value = vbChecked) & _
                                                  "    CalcState=" & cboCalcStCd.Text & "    InterestRate=" & CStr(ipdPayeClmIntRt.value), cstrCurrentProc

    ' Initialize the calculated amounts in each StateInfo structure
    fnInitializeCalcInfo msiCalcStCd
    fnInitializeCalcInfo msiInsdDthResStCd
    fnInitializeCalcInfo msiPayeStCd
    fnInitializeCalcInfo msiIssStCd
    If mfrmMyInsuredForm.chkClmCmpCalInd.value = vbChecked Then
        ' This is Compact Calcatution
        fnInitializeCalcInfo msiCompactCalc
    End If
    
    'Debug.Print "msiCalcStCd.StrlEffDt, msiInsdDthResStCd.StrlEffDt, msiPayeStCd.StrlEffDt, msiIssStCd.StrlEffDt, msiOverride.StrlEffDt:"
    'Debug.Print msiCalcStCd.StrlEffDt, msiInsdDthResStCd.StrlEffDt, msiPayeStCd.StrlEffDt, msiIssStCd.StrlEffDt, msiOverride.StrlEffDt
    
    ' If a Calculation Override is being done, then ONLY do a calc using the Calc St/Interest Rate. (1-way)
    If chkPayeDfltOvrdInd.value = vbChecked Then
        msiOverride.StCd = cboCalcStCd.Text
        msiOverride.InterestRateToUse = ipdPayeClmIntRt.value
        fnCalcAndLogMSIInfo msiOverride, "Calc State", True
        msiCalcStCd = msiOverride
    Else
        ' Only do a calc using Insured Residence State At Time Of Death if the Insured lived within the
        ' United States and its territories (3-way)
        If chkClmForResDthInd_UsedInAutoCalc.value = vbUnchecked Then
            fnCalcAndLogMSIInfo msiInsdDthResStCd, "Insured's Residence State"
           '' BZ4999 October 2013 Non US payee sxs
            If cboPayeStCd <> "ZZ" Then
                fnCalcAndLogMSIInfo msiPayeStCd, "Payee's Residence State"
            End If
            fnCalcAndLogMSIInfo msiIssStCd, "Contract Issue State"
        Else
            ' Otherwise do a 2-way calc
            '' BZ4999 October 2013 Non US payee sxs
            If cboPayeStCd <> "ZZ" Then
                fnCalcAndLogMSIInfo msiPayeStCd, "Payee's Residence State"
            End If
            fnCalcAndLogMSIInfo msiIssStCd, "Contract Issue State"
        End If
        If mfrmMyInsuredForm.chkClmCmpCalInd.value = vbChecked Then
            ' This is Compact Calcatution
                'if death to payment is less than 31 days then use current rate
                'if over pay from proof to payment, with an 10% rate
            msiCompactCalc.StCd = cstCompactFilling
            fnCalcAndLogMSIInfo msiCompactCalc, "Compact Calculation"
        End If
        ' Pick the one that calculated the highest Claim Interest
        ' (Be sure not to look at msiInsdDthResStCd first since it would have an empty StCd value
        '  if the Foreign Residence at Death checkbox is selected and hence the Update could fail
        '  due to "blank" not being defined on STATE_T.)
        '' BZ4999 October 2013 Non US payee - SXS
            If cboPayeStCd = "ZZ" Then
                msiCalcStCd = msiIssStCd
            Else
                msiCalcStCd = msiPayeStCd
            End If
        'Y027 07-Nov-2012
        'If its a GROUP Policy and any of these five states are involved then ignore that state
        If m_AdmPolicySystem = m_GROUP_ADMIN_SYS Then
            'Assign it to a state which is not to be ignored
            '' BZ4999 October 2013 Non US payee - SXS
            If fnAnomolyState(msiPayeStCd.StCd) = False And cboPayeStCd <> "ZZ" Then
                msiCalcStCd = msiPayeStCd
            ElseIf fnAnomolyState(msiInsdDthResStCd.StCd) = False Then
                msiCalcStCd = msiInsdDthResStCd
            Else
                msiCalcStCd = msiIssStCd
            End If
            If msiIssStCd.ClaimInterestAmt > msiCalcStCd.ClaimInterestAmt And _
                fnAnomolyState(msiIssStCd.StCd) = False Then
                msiCalcStCd = msiIssStCd
            End If
            If msiInsdDthResStCd.ClaimInterestAmt > msiCalcStCd.ClaimInterestAmt And _
                fnAnomolyState(msiInsdDthResStCd.StCd) = False Then
                msiCalcStCd = msiInsdDthResStCd
            End If
            If msiPayeStCd.ClaimInterestAmt > msiCalcStCd.ClaimInterestAmt And _
                fnAnomolyState(msiPayeStCd.StCd) = False Then
                msiCalcStCd = msiPayeStCd
            End If
            If mfrmMyInsuredForm.chkClmCmpCalInd.value = vbChecked Then
                ' This is Compact Calcatution
                If msiCompactCalc.ClaimInterestAmt > msiCalcStCd.ClaimInterestAmt And _
                    fnAnomolyState(msiCompactCalc.StCd) = False Then
                    msiCalcStCd = msiCompactCalc
                End If
            End If
            'At this point we have the highest interest rate of a non anomolous state or
            'we have an anomolous state
        Else
            'Non group policy
            If msiInsdDthResStCd.ClaimInterestAmt > msiCalcStCd.ClaimInterestAmt Then
                msiCalcStCd = msiInsdDthResStCd
            End If
            If msiIssStCd.ClaimInterestAmt > msiCalcStCd.ClaimInterestAmt Then
                msiCalcStCd = msiIssStCd
            End If
            If mfrmMyInsuredForm.chkClmCmpCalInd.value = vbChecked Then
                ' This is Compact Calcatution
                If msiCompactCalc.ClaimInterestAmt > msiCalcStCd.ClaimInterestAmt Then
                    msiCalcStCd = msiCompactCalc
                End If
            End If
        End If
    End If

    fnLogWrite gcstrBlankEntry, cstrCurrentProc
    fnLogWrite "The state selected was " & msiCalcStCd.StCd & ".", cstrCurrentProc
    
    cboCalcStCd.Text = msiCalcStCd.StCd
    txtCalcStCdSpecialInstructions_UsedInAutoCalc.Text = msiCalcStCd.StrlSpclInstrTxt

    If fnAnomolyState(msiCalcStCd.StCd) = True And m_AdmPolicySystem = m_GROUP_ADMIN_SYS Then
    
        lblCalculationInfo = "Note: This is a group policy - an interest rate of 0% applies."
        ipcPayeClmIntAmt.Text = 0
        ipcPayeWthldAmt.Text = 0
        mcurTotalWithheld = 0
        ipdPayeIntDaysPdNum.Text = 0
        ipdPayeClmIntRt.UnFmtText = 0
        ipcPayeClmPdAmt.Text = ipcPayeDthbPmtAmt.Text
    
    Else
    
        ' Initialize on-screen label that shows some info about how the
        ' calculation was done.
        lblCalculationInfo = msiCalcStCd.CalculationInfo
        ipcPayeClmIntAmt.Text = msiCalcStCd.ClaimInterestAmt
        ipcPayeWthldAmt.Text = msiCalcStCd.WithheldAmt
        mcurTotalWithheld = msiCalcStCd.WithheldAmt
        ipdPayeIntDaysPdNum.Text = msiCalcStCd.NbrOfDaysToPayInterest
        ipcPayeClmPdAmt.Text = msiCalcStCd.TotalForThisPayee
        
        ipdPayeClmIntRt.UnFmtText = msiCalcStCd.InterestRateToUse
        

    End If
            

    
    fnCalcClaimInterest = True
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function


Private Function fnAnomolyState(ByRef strStateCode As String) As Boolean
    If strStateCode <> "AR" And strStateCode <> "FL" And _
        strStateCode <> "IN" And strStateCode <> "IL" And _
                   strStateCode <> "MA" And strStateCode <> " " Then
        fnAnomolyState = False
    Else
        fnAnomolyState = True
    End If
End Function

'////////////////////////////////////////////////////////////////////////////////////////////////
Private Function fnCalcClaimForState(ByRef msiStateIn As StateInfo, ByVal strDesc As String, _
    ByVal bUseSuppliedRate As Boolean) As Boolean
    ' Comments  : This function will calculate Claim Interest
    '             for the specified State.
    ' Parameters:
    '             msiStateIn (in/out)   Pointer to a StateInfo structure, containing state to calculate
    '             strDesc (in)          Descriptive text to appear in prompts to explain why rate is needed
    '             bUseSuppliedRate (in) If True, indicates to use supplied rate (as if the IRULE_CD = "SPECAMT")
    '                                   and to suppress all rate-related prompts for that state.
    ' Returns   : True, if successful; False otherwise
    ' Modified  :
    '  07/17/03 K758  For bug 2455, modified the calc of # of Days Of Interest To Be Paid
    '                 to floor it at zero, so negative numbers won't come through and
    '                 adversely affect the Claims Interest Amount's calculation.
    ' --------------------------------------------------
    On Error GoTo PROC_ERR

    Const cstrCurrentProc As String = "fnCalcClaimForState"
    Const cdblFloorOfZero As Double = 0
    Dim dteDateOfPayment As Date

    dteDateOfPayment = DateValue(dtpPayePmtDt.value)
    
    '       **************************************************
    '       **************************************************
    '         The steps listed below are described in the
    '         ClaimsInterest_HowToDoManualCalc.Doc document
    '         on \\500ip03\Vol2\DesktopTechnology\Deploy\Docs
    '       **************************************************
    '       **************************************************

    ' ................................................................
    '    Step 1. No interest will be paid if CalcIdtypCd=None
    ' ................................................................
    If UCase$(msiStateIn.CalcIdtypCd) = "NONE    " Then
        msiStateIn.CalculationInfo = msiStateIn.CalculationInfo & "No interest was paid due " & _
            "to that state's Calculation Interest Date Type Code specification. "
        msiStateIn.NbrOfDaysToPayInterest = 0
        GoTo STEP8
    End If

    ' ................................................................
    '    Step 2. No interest will be paid if ReqdIdtypCd=None
    ' ................................................................
    If UCase$(msiStateIn.ReqdIdtypCd) = "NONE    " Then
        msiStateIn.CalculationInfo = msiStateIn.CalculationInfo & "No interest was paid due " & _
            "to that state's Required Interest Date Type Code specification. "
        msiStateIn.NbrOfDaysToPayInterest = 0
        GoTo STEP8
    End If

    ' ................................................................
    '    Step 3. Calculate the Payable Period End Date. If the
    '            claim is being paid ON or BEFORE that date,
    '            we do NOT pay interest on the claim. If the
    '            claim is being paid AFTER that date, we may have
    '            to pay interest on the claim (unless Step 7 says
    '            otherwise).
    ' ................................................................
    If UCase$(msiStateIn.ReqdIdtypCd) = "PROOF   " Then
        msiStateIn.PayablePeriodEndDate = DateValue(mfrmMyInsuredForm.InsuredClmProofDt) _
                                  + msiStateIn.StrlIntReqdOfstNum
    Else
        msiStateIn.PayablePeriodEndDate = DateValue(mfrmMyInsuredForm.InsuredClmInsdDthDt) _
                                  + msiStateIn.StrlIntReqdOfstNum
    End If
    
    ' ................................................................
    '    Step 4. No interest will be paid if the claim is being
    '            paid on or before the Payable Period End Date.
    ' ................................................................
    If dteDateOfPayment <= msiStateIn.PayablePeriodEndDate Then
        msiStateIn.CalculationInfo = msiStateIn.CalculationInfo & "No interest was paid due " & _
            "to the Date Of Payment being within the Payable Period. "
        msiStateIn.NbrOfDaysToPayInterest = 0
        GoTo STEP8
    End If

    ' ................................................................
    '    Step 5. Calculate the Figured From Date. This will be
    '            used to calculate the number of days of interest
    '            to pay.
    ' ................................................................
    If UCase$(msiStateIn.CalcIdtypCd) = "PROOF   " Then
        msiStateIn.FiguredFromDate = DateValue(mfrmMyInsuredForm.InsuredClmProofDt) _
                             + msiStateIn.StrlIntCalcOfstNum
    Else
        msiStateIn.FiguredFromDate = DateValue(mfrmMyInsuredForm.InsuredClmInsdDthDt) _
                             + msiStateIn.StrlIntCalcOfstNum
    End If

    ' ................................................................
    '    Step 6. Calculate the number of days to pay interest on
    '            the claim.
    '            NOTE: If this is set to 0, then no interest or
    '            withholding will be paid when 0 is plugged into
    '            the formulaes in Steps 8 and 9.
    ' ................................................................
    msiStateIn.NbrOfDaysToPayInterest = DateDiff("d", msiStateIn.FiguredFromDate, dteDateOfPayment)
    
    ' Per bug 2455, floor the NbrOfDaysToPayInterest so negative numbers are turned into 0.
    msiStateIn.NbrOfDaysToPayInterest = fnAtLeast(msiStateIn.NbrOfDaysToPayInterest, cdblFloorOfZero)
    
    msiStateIn.CalculationInfo = msiStateIn.CalculationInfo & "The # of days (" & _
                         msiStateIn.NbrOfDaysToPayInterest & _
                         ") was based on " & CStr(msiStateIn.FiguredFromDate) & " to " & _
                         CStr(dteDateOfPayment) & ". "

    ' ................................................................
    '    Step 7. Determine the interest rate to use to calculate
    '            Claims Interest
    ' ................................................................
    If bUseSuppliedRate Then
        ' Do nothing...the InterestRateToUse was previously set
    Else
        msiStateIn.InterestRateToUse = fnGetInterestRate(msiStateIn, strDesc)
    End If


STEP8:
    ' ................................................................
    '    Step 8. Calculate the Claims Interest Amount, rounded to
    '            2 decimal positions
    ' ................................................................
    msiStateIn.ClaimInterestAmt = (CCur(ipcPayeDthbPmtAmt.Text) * (msiStateIn.InterestRateToUse / 100))
    msiStateIn.ClaimInterestAmt = msiStateIn.ClaimInterestAmt * (msiStateIn.NbrOfDaysToPayInterest / 365)
    msiStateIn.ClaimInterestAmt = Round(msiStateIn.ClaimInterestAmt, 2)

    ' ................................................................
    '    Step 9. Calculate the Withheld Amount, rounded to
    '            2 decimal positions. If the Claim Interest Amount
    '             is zero, then the Withheld Amount will be zero.
    ' ................................................................
    msiStateIn.WithheldAmt = msiStateIn.ClaimInterestAmt * (CCur(ipdPayeWthldRt.Text) / 100)
    msiStateIn.WithheldAmt = Round(msiStateIn.WithheldAmt, 2)

    ' ................................................................
    '    Step 10. Calculate the Total Amount to be paid for this Payee.
    ' ................................................................
    msiStateIn.TotalForThisPayee = CCur(ipcPayeDthbPmtAmt.Text) + msiStateIn.ClaimInterestAmt - msiStateIn.WithheldAmt

    fnCalcClaimForState = True
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function


'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub fnClearControls()
    ' Comments  : Initializes screen controls in order to add a new record
    ' Parameters: None
    ' Called by : fnAddRecord of frmPayee
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc       As String = "fnClearControls"
    Const cintZero              As Integer = 0
    Dim ctl                     As Control
    Dim varDefaultValue         As Variant
    Dim strSavedMask            As String

    ' Hide updates to the window until we're done. This avoids ugly screen flickering
    fnWindowLock Me.hWnd

    iptPayeFullNm.Text = vbNullString
    iptPayeCareOfTxt.Text = vbNullString
    iptPayeAddrLn1Txt.Text = vbNullString
    iptPayeAddrLn2Txt.Text = vbNullString
    iptPayeCityNmTxt.Text = vbNullString
    
    If cboCalcStCd.ListCount > 0 Then
        cboCalcStCd.ListIndex = 0   ' Select first (blank) entry
    Else
        gerhApp.ReportNonFatal vbObjectError + gcRES_WARN_CBO_IS_EMPTY, _
                               mstrScreenName & gcstrDOT & cstrCurrentProc, _
                               mcstrCboCalcStCdLabel
    End If
        
    iptPayeZipCd.Text = vbNullString
    iptPayeZip4Cd.Text = vbNullString
    
    ' NOTE: For MaskEdBox and fpMask controls, have to remove mask before clearing out the control
    '       since the vbNullString value doesn't match the mask specification.
    strSavedMask = ipmPayeSsnTinNum.Mask
    ipmPayeSsnTinNum.Mask = vbNullString
    ipmPayeSsnTinNum.Text = vbNullString
    ipmPayeSsnTinNum.Mask = strSavedMask
    
    If cboPayeStCd.ListCount > 0 Then
        cboPayeStCd.ListIndex = 0   ' Select first (blank) entry
    Else
        gerhApp.ReportNonFatal vbObjectError + gcRES_WARN_CBO_IS_EMPTY, _
                               mstrScreenName & gcstrDOT & cstrCurrentProc, _
                               mcstrCboPayeStCdLabel
    End If

    ipdPayeWthldRt.value = cintZero
    
    
    ' ------------------------------------------------------------------------
    '    The checkPayeDfltOvrdInd, cboCalcStCd and iptPayeClmIntRt and
    '   lblWarningAboutOverride controls are all tied to one another with
    '   regard to their availability and initialization.
    ' ------------------------------------------------------------------------
    ' Set their values
    chkPayeDfltOvrdInd.value = vbUnchecked

    If cboCalcStCd.ListCount > 0 Then
        cboCalcStCd.ListIndex = 0   ' Select first (blank) entry
    Else
        gerhApp.ReportNonFatal vbObjectError + gcRES_WARN_CBO_IS_EMPTY, _
                               mstrScreenName & gcstrDOT & cstrCurrentProc, _
                               mcstrCboCalcStCdLabel
    End If

    ipdPayeClmIntRt.value = cintZero
    
    ' Set their availability
    fnEnableDisableControl ctlIn:=ChkPaye1099Ind, bEnable:=True  '' BZ4999 October 2013 Non US payee - SXS
    fnEnableDisableControl ctlIn:=chkPayeDfltOvrdInd, bEnable:=False
    fnEnableDisableControl ctlIn:=cboCalcStCd, bEnable:=False
    fnEnableDisableControl ctlIn:=ipdPayeClmIntRt, bEnable:=False
    lblWarningAboutOverride.Visible = False
    ' ------------------------------------------------------------------------
    
    ' DateTimePicker controls (dtpPayePmtDt) will
    ' automatically be set to today's date. Cannot set them to Null
    ' unless their CheckBox property is set to True.
    dtpPayePmtDt.value = Date
    fnResetStateRules
    
    ipdPayeIntDaysPdNum.value = cintZero
    ipcPayeDthbPmtAmt.value = cintZero
    ipcPayeClmIntAmt.value = cintZero
    ipcPayeWthldAmt.value = cintZero
    ipcPayeClmPdAmt.value = cintZero
    
    ' intitialize the label that describes how the calculation was done.
    lblCalculationInfo = vbNullString
    
    ' Initialize fields that will be set when calculation is done
    txtCalcStCdSpecialInstructions_UsedInAutoCalc = vbNullString
    mcurTotalWithheld = 0   ' non-displayed version of ipcPayeWthldAmt
    
    ' Skip initialization of Insured Residence State and Insured Residence State's
    ' Special Instructions since these should not change from payee to payee as they
    ' are based on the Insured screen to which all payees belong.
        
    ' Skip initialization of Payee Residence State and Payee Residence State's
    ' Special Instructions since they are set when the cboPayeStCd is changed
    ' (as occurred when its ListIndex was set to 0 above).
    '       txtPayeStCd_UsedInAutoCalc = vbNullString
    '       txtPayeStCdSpecialInstructions_UsedInAutoCalc = vbNullString
        
    ' Skip initialization of Contract Issue State's Special Instructions since
    ' it is set when the Contract Issue State changed (as occurred when its
    ' ListIndex was set to 0 above) .
    '       txtIssStCdSpecialInstructions_UsedInAutoCalc = vbNullString

    
    ' Next, set each control to its default value per the meta data, if available.
    For Each ctl In Me.Controls
        With ctl
            ' Debug.Print ctl.Name & vbTab & ctl.Tag
            ' If control corresponds to a SQL Server table column, then try
            ' to set its default properties. The Tag property contains
            ' the name of its property within the table class.
            If Len(.Tag) > 0 Then
                ' If there's a default value, use it
                varDefaultValue = mtWrapper.DefaultValue(.Tag)
                If Not (IsEmpty(varDefaultValue)) Then
                    If (TypeOf ctl Is TextBox) Or (TypeOf ctl Is fpText) Or (TypeOf ctl Is ComboBox) Or (TypeOf ctl Is ListBox) Then
                        .Text = varDefaultValue
                    ElseIf (TypeOf ctl Is fpCurrency) Then
                        .value = varDefaultValue
                    'ElseIf (TypeOf ctl Is MaskEdBox) Then
                    '    .SelText = varDefaultValue
                    ElseIf (TypeOf ctl Is fpMask) Or (TypeOf ctl Is fpDoubleSingle) Then
                        .UnFmtText = varDefaultValue
                    ElseIf TypeOf ctl Is CheckBox Then
                        ' Bug thinks the default value is "Y" or "N" when really it's True or False
                        If UCase$(varDefaultValue) = True Then
                            .value = vbChecked
                        Else
                            .value = vbUnchecked
                        End If
                    ElseIf TypeOf ctl Is Label Then
                        .Caption = varDefaultValue
                    ElseIf TypeOf ctl Is fpCombo Then
                        fnSearchFPCombo lpcIn:=ctl, strSearchText:=varDefaultValue, intSearchCol:=1
                    End If
                End If
            End If
        End With
    Next ctl
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    
    ' Clean-up statements go here
    fnWindowUnlock
    
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
Private Function fnGetListOfStates() As ADODB.Recordset
    '--------------------------------------------------------------------------
    ' Procedure:   fnGetListOfStates
    ' Description: Builds an ADODB.Recordset containing state codes in STATE_T
    ' Params:      N/A
    ' Returns:     N/A
    ' Modified:
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc          As String = "fnGetListOfStates"
    Const cstrSproc                As String = "dbo.proc_state_lu_select" ' Stored procedure to execute
    Dim prmReturnValue             As ADODB.Parameter
    Dim adwTemp                    As cadwADOWrapper
 
    On Error GoTo PROC_ERR

    Set adwTemp = New cadwADOWrapper

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

        Set fnGetListOfStates = .Execute()
        ' Do not close this recordset.
    End With
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0

    ' Clean-up statements go here
    ' Do not free the fnGetListOfStates recordset!
    fnFreeObject adwTemp
    fnFreeObject prmReturnValue

    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    Exit Function
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
End Function


'////////////////////////////////////////////////////////////////////////////////////////////////
Private Function fnGetCurrentIntRate(ByVal dtePayePmtDt As Date) As Double
    '--------------------------------------------------------------------------
    ' Procedure:   fnGetCurrentIntRate
    ' Description: retrieves the Current Rate in effect on the Date of Payment
    ' Params:      N/A
    ' Returns:     double, representing the Current Rate
    ' Modified:
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc          As String = "fnGetCurrentIntRate"
    Const cstrSproc                As String = "dbo.proc_current_rate_select" ' Stored procedure to execute
    Dim prmReturnValue             As ADODB.Parameter
    Dim prmPayePmtDt               As ADODB.Parameter
    Dim prmCurrIntRt               As ADODB.Parameter
    Dim adwTemp                    As cadwADOWrapper
    Dim rstTemp                    As ADODB.Recordset
 
    On Error GoTo PROC_ERR

    Set adwTemp = New cadwADOWrapper

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

        ' ---Parameter #2---
        Set prmPayePmtDt = .CreateParameter(Name:="@paye_pmt_dt", _
                                              Type:=adDBTimeStamp, _
                                              Direction:=adParamInput, _
                                              Size:=16, _
                                              value:=dtePayePmtDt)
        .Parameters.Append prmPayePmtDt

        ' ---Parameter #3---
        Set prmCurrIntRt = .CreateParameter(Name:="@curr_int_rt", _
                                              Type:=adNumeric, _
                                              Direction:=adParamOutput, _
                                              value:=Null)
        .Parameters.Append prmCurrIntRt
        ' Have to hard-code the precision/scale since we have no meta data for this table
        With prmCurrIntRt
            .Precision = 11
            .NumericScale = 5
        End With
        
        Set rstTemp = .Execute()
        
        If IsNull(prmCurrIntRt.value) Then
            fnGetCurrentIntRate = gclngNoSelection             ' -1
        Else
            fnGetCurrentIntRate = prmCurrIntRt.value
        End If
    End With
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0

    ' Clean-up statements go here
    fnFreeObject adwTemp
    fnFreeRecordset rstTemp
    fnFreeObject prmReturnValue
    fnFreeObject prmPayePmtDt
    fnFreeObject prmCurrIntRt

    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case prmReturnValue
        Case gcRES_NERR_REC_NOT_FOUND
            ' 4027 = The specified record was not found in the database (@@1).
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_REC_NOT_FOUND, _
                                       mstrScreenName & gcstrDOT & cstrCurrentProc, _
                                       "Current Rate effective on [" & FormatDateTime(dtePayePmtDt, vbShortDate) & "]"
            Resume PROC_EXIT
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
End Function


'////////////////////////////////////////////////////////////////////////////////////////////////
Private Function fnGetFieldLabel(ByVal strControlName As String) As String
    '--------------------------------------------------------------------------
    ' Procedure:   fnGetFieldLabel
    ' Description: Given a control name, return the value of the control's label
    '
    ' Params:      N/A
    '    strControlName  (in) A string containing the control's name
    '
    ' Returns:     A string containing the controls' label.
    '-----------------------------------------------------------------------------
    
    '!CUSTOMIZE!  There should be one Case statement for each control that
    '             corresponds to a table column. Each Case statement should
    '             reference a Const literal that indicates how the control is
    '             labelled on-screen.

    Const cstrCurrentProc       As String = "fnGetFieldLabel"
 
    On Error GoTo PROC_ERR

    Select Case strControlName
        Case "iptPayeFullNm"
            fnGetFieldLabel = mcstrIptPayeFullNmLabel
        Case "iptPayeCareOfTxt"
            fnGetFieldLabel = mcstrIptPayeCareOfTxtLabel
        Case "iptPayeAddrLn1Txt"
            fnGetFieldLabel = mcstrIptPayeAddrLn1TxtLabel
        Case "iptPayeAddrLn2Txt"
            fnGetFieldLabel = mcstrIptPayeAddrLn2TxtLabel
        Case "iptPayeCityNmTxt"
            fnGetFieldLabel = mcstrIptPayeCityNmTxtLabel
        Case "cboPayeStCd"
            fnGetFieldLabel = mcstrCboPayeStCdLabel
        Case "iptPayeZipCd"
            fnGetFieldLabel = mcstrIptPayeZipCdLabel
        Case "iptPayeZip4Cd"
            fnGetFieldLabel = mcstrIptPayeZip4CdLabel
        Case "ChkPaye1099Ind"
            fnGetFieldLabel = mcstrChkPaye1099IndLabel   '' BZ4999 October 2013 Non US payee - SXS
        Case "ipmPayeSsnTinNum"
            fnGetFieldLabel = mcstrIpmPayeSsnTinNumLabel
        Case "cboPayeSsnTinTypCd"
            fnGetFieldLabel = mcstrCboPayeSsnTinTypCdLabel
        Case "ipdPayeWthldRt"
            fnGetFieldLabel = mcstrIpdPayeWthldRtLabel
        Case "ipdPayeClmIntRt"
            fnGetFieldLabel = mcstrIpdPayeClmIntRtLabel
        Case "chkPayeDfltOvrdInd"
            fnGetFieldLabel = mcstrChkPayeDfltOvrdIndLabel
        Case "cboCalcStCd"
            fnGetFieldLabel = mcstrCboCalcStCdLabel
        Case "dtpPayePmtDt"
            fnGetFieldLabel = mcstrDtpPayePmtDtLabel
        Case "ipdPayeIntDaysPdNum"
            fnGetFieldLabel = mcstrIpdPayeIntDaysPdNumLabel
        Case "ipcPayeDthbPmtAmt"
            fnGetFieldLabel = mcstrIpcPayeDthbPmtAmtLabel
        Case "ipcPayeClmIntAmt"
            fnGetFieldLabel = mcstrIpcPayeClmIntAmtLabel
        Case "ipcPayeWthldAmt"
            fnGetFieldLabel = mcstrIpcPayeWthldAmtLabel
        Case "ipcPayeClmPdAmt"
            fnGetFieldLabel = mcstrIpcPayeClmPdAmtLabel
        Case Else
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_UNEXPECTED_VAL_SELECT_CASE, _
                mstrScreenName & gcstrDOT & cstrCurrentProc
            GoTo PROC_EXIT
    End Select
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function


'////////////////////////////////////////////////////////////////////////////////////////////////
Private Function fnGetInterestRate(ByRef siRatesIn As StateInfo, ByVal strDesc As String) As Double
    ' Comments  : This function will return the interest
    '             rate. This rate is taken from the State98
    '             table directly if it is numeric. Otherwise
    '             the user is prompted for it.
    ' Parameters: siRatesIn (in) - a StateInfo structure that
    '                              contains pertinent info from a row in the
    '                              STATE_RULE_T table.
    '             strDesc (in)   - Descriptive text that indicates why a rate is needed (to appear in prompts)
    ' Returns   : The interest rate as a Double
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc               As String = "fnGetInterestRate"
    Const cintMaxInterestRate           As Integer = 12
    Const cstrPleaseSpecify             As String = "Please specify the "
    Const cstrInEffectOn                As String = " in effect on "
    Const cstrForThe                    As String = " for the "
    Const cstrOf                        As String = " of "
    Const cstrCurrLoanRt                As String = "Current Loan Rate"
    Const cstrDerivedRateBasedOn        As String = "The derived Rate based on "
    Const cstrCurrRtInEffectOn          As String = "The Current Rate in effect on "
    Const cstrIsNotDefdSupplyARate      As String = " is not defined. Please supply the Rate to use "
    Const cstrIsNegNbrSupplyRateToUse   As String = " is a negative number. Please supply the Rate to use "
    Const cstrCurrLoanRtX               As String = "the Current Loan Rate("
    Const cstrCurrRtX                   As String = "the Current Rate("
    Const cstrCloseParenMinus           As String = ") - "
    Const cstrCloseParenPlus            As String = ") + "
    Const cstrStateRuleAmtX             As String = "the State Rule Amount("
    Const cstrCloseParen                As String = ") "
    Const cstrTheMaxOf                  As String = "the maximum of "
    Const cstrTheMinOf                  As String = "the minimum of "
    Const cstrTheGreaterOf              As String = "the greater of "
    Const cstrAnd                       As String = " and "
    Dim dblCurrIntRt                    As Double
    Dim dblCurrLoanRt                   As Double
    Dim strPayePmtDt                    As String
    Dim dblRateToUse                    As Double

    strPayePmtDt = FormatDateTime(dtpPayePmtDt.value, vbShortDate)
    
DETERMINE_RATE:
    Select Case siRatesIn.IruleCd
        Case "CLNW/MAX"
            dblCurrLoanRt = fnPromptForRate(cstrPleaseSpecify & cstrCurrLoanRt & cstrInEffectOn & strPayePmtDt _
                & cstrForThe & strDesc & cstrOf & siRatesIn.StCd)
            If dblCurrLoanRt > siRatesIn.StrlIntRuleAmt Then
                dblRateToUse = siRatesIn.StrlIntRuleAmt
            Else
                dblRateToUse = dblCurrLoanRt
            End If
            If dblRateToUse < 0 Then
                ' If the derived rate is negative, ask the user to supply it
                dblRateToUse = fnPromptForRate(cstrDerivedRateBasedOn & cstrTheMaxOf & cstrCurrLoanRtX & CStr(dblCurrLoanRt) & _
                               cstrCloseParen & cstrAnd & cstrStateRuleAmtX & CStr(siRatesIn.StrlIntRuleAmt) & _
                               cstrCloseParen & cstrIsNegNbrSupplyRateToUse & _
                               cstrForThe & strDesc & cstrOf & siRatesIn.StCd)
            End If
            
        Case "CLNW/MIN"
            dblCurrLoanRt = fnPromptForRate(cstrPleaseSpecify & cstrCurrLoanRt & cstrInEffectOn & strPayePmtDt _
                & cstrForThe & strDesc & cstrOf & siRatesIn.StCd)
            If dblCurrLoanRt < siRatesIn.StrlIntRuleAmt Then
                dblRateToUse = siRatesIn.StrlIntRuleAmt
            Else
                dblRateToUse = dblCurrLoanRt
            End If
            If dblRateToUse < 0 Then
                ' If the derived rate is negative, ask the user to supply it
                dblRateToUse = fnPromptForRate(cstrDerivedRateBasedOn & cstrTheMinOf & cstrCurrLoanRtX & CStr(dblCurrLoanRt) & _
                               cstrCloseParen & cstrAnd & cstrStateRuleAmtX & CStr(siRatesIn.StrlIntRuleAmt) & _
                               cstrCloseParen & cstrIsNegNbrSupplyRateToUse & _
                               cstrForThe & strDesc & cstrOf & siRatesIn.StCd)
            End If
            
        Case "CRTW/MAX"
            dblCurrIntRt = fnGetCurrentIntRate(dtpPayePmtDt.value)
            If dblCurrIntRt = gclngNoSelection Then
                ' If Current Rate was not found, ask the user to supply it
                dblCurrIntRt = fnPromptForRate(cstrCurrRtInEffectOn & strPayePmtDt & cstrIsNotDefdSupplyARate _
                               & cstrForThe & strDesc & cstrOf & siRatesIn.StCd)
            End If
            If dblCurrIntRt > siRatesIn.StrlIntRuleAmt Then
                dblRateToUse = siRatesIn.StrlIntRuleAmt
            Else
                dblRateToUse = dblCurrIntRt
            End If
            If dblRateToUse < 0 Then
                ' If the derived rate is negative, ask the user to supply it
                dblRateToUse = fnPromptForRate(cstrDerivedRateBasedOn & cstrTheMaxOf & cstrCurrRtX & CStr(dblCurrIntRt) & _
                               cstrCloseParen & cstrAnd & cstrStateRuleAmtX & CStr(siRatesIn.StrlIntRuleAmt) & _
                               cstrCloseParen & cstrIsNegNbrSupplyRateToUse & _
                               cstrForThe & strDesc & cstrOf & siRatesIn.StCd)
            End If
            
        Case "CRTW/MIN"
            dblCurrIntRt = fnGetCurrentIntRate(dtpPayePmtDt.value)
            ' If Current Rate was not found, ask the user to supply it
            If dblCurrIntRt = gclngNoSelection Then
                dblCurrIntRt = fnPromptForRate(cstrCurrRtInEffectOn & strPayePmtDt & cstrIsNotDefdSupplyARate _
                               & cstrForThe & strDesc & cstrOf & siRatesIn.StCd)
            End If
            If dblCurrIntRt < siRatesIn.StrlIntRuleAmt Then
                dblRateToUse = siRatesIn.StrlIntRuleAmt
            Else
                dblRateToUse = dblCurrIntRt
            End If
            If dblRateToUse < 0 Then
                ' If the derived rate is negative, ask the user to supply it
                dblRateToUse = fnPromptForRate(cstrDerivedRateBasedOn & cstrTheMinOf & cstrCurrRtX & CStr(dblCurrIntRt) & _
                               cstrCloseParen & cstrAnd & cstrStateRuleAmtX & CStr(siRatesIn.StrlIntRuleAmt) & _
                               cstrCloseParen & cstrIsNegNbrSupplyRateToUse & _
                               cstrForThe & strDesc & cstrOf & siRatesIn.StCd)
            End If
            
        Case "CURLN   "
            dblRateToUse = fnPromptForRate(cstrPleaseSpecify & cstrCurrLoanRt & cstrInEffectOn & strPayePmtDt _
                & cstrForThe & strDesc & cstrOf & siRatesIn.StCd)
        
        Case "CURLN+X "
            dblCurrLoanRt = fnPromptForRate(cstrPleaseSpecify & cstrCurrLoanRt & cstrInEffectOn & strPayePmtDt _
                & cstrForThe & strDesc & cstrOf & siRatesIn.StCd)
            dblRateToUse = dblCurrLoanRt + siRatesIn.StrlIntRuleAmt
            If dblRateToUse < 0 Then
                ' If the derived rate is negative, ask the user to supply it
                dblRateToUse = fnPromptForRate(cstrDerivedRateBasedOn & cstrCurrLoanRtX & CStr(dblCurrIntRt) & _
                               cstrCloseParenPlus & cstrStateRuleAmtX & CStr(siRatesIn.StrlIntRuleAmt) & _
                               cstrCloseParen & cstrIsNegNbrSupplyRateToUse & _
                               cstrForThe & strDesc & cstrOf & siRatesIn.StCd)
            End If
        
        Case "CURLN-X "
            dblCurrLoanRt = fnPromptForRate(cstrPleaseSpecify & cstrCurrLoanRt & cstrInEffectOn & strPayePmtDt _
                & cstrForThe & strDesc & cstrOf & siRatesIn.StCd)
            dblRateToUse = dblCurrLoanRt - siRatesIn.StrlIntRuleAmt
            If dblRateToUse < 0 Then
                ' If the derived rate is negative, ask the user to supply it
                dblRateToUse = fnPromptForRate(cstrDerivedRateBasedOn & cstrCurrLoanRtX & CStr(dblCurrLoanRt) & _
                               cstrCloseParenMinus & cstrStateRuleAmtX & CStr(siRatesIn.StrlIntRuleAmt) & _
                               cstrCloseParen & cstrIsNegNbrSupplyRateToUse & _
                               cstrForThe & strDesc & cstrOf & siRatesIn.StCd)
            End If
            
        Case "CURRT   "
            dblCurrIntRt = fnGetCurrentIntRate(dtpPayePmtDt.value)
            If dblCurrIntRt = gclngNoSelection Then
                dblCurrIntRt = fnPromptForRate(cstrCurrRtInEffectOn & strPayePmtDt & cstrIsNotDefdSupplyARate _
                               & cstrForThe & strDesc & cstrOf & siRatesIn.StCd)
            ElseIf dblCurrIntRt < 0 Then
                ' If Current Rate was found but is negative, ask the user to supply it
                dblCurrIntRt = fnPromptForRate(cstrCurrRtInEffectOn & strPayePmtDt & cstrIsNegNbrSupplyRateToUse _
                               & cstrForThe & strDesc & cstrOf & siRatesIn.StCd)
            End If
            dblRateToUse = dblCurrIntRt
            
        Case "CURRT+X "
            dblCurrIntRt = fnGetCurrentIntRate(dtpPayePmtDt.value)
            If dblCurrIntRt = gclngNoSelection Then
                dblCurrIntRt = fnPromptForRate(cstrCurrRtInEffectOn & strPayePmtDt & cstrIsNotDefdSupplyARate _
                               & cstrForThe & strDesc & cstrOf & siRatesIn.StCd)
            End If
            dblRateToUse = dblCurrIntRt + siRatesIn.StrlIntRuleAmt
            If dblRateToUse < 0 Then
                ' If the derived rate is negative, ask the user to supply it
                dblRateToUse = fnPromptForRate(cstrDerivedRateBasedOn & cstrCurrRtX & CStr(dblCurrIntRt) & _
                               cstrCloseParenPlus & cstrStateRuleAmtX & CStr(siRatesIn.StrlIntRuleAmt) & _
                               cstrCloseParen & cstrIsNegNbrSupplyRateToUse & _
                               cstrForThe & strDesc & cstrOf & siRatesIn.StCd)
            End If
        
        Case "CURRT-X "
            dblCurrIntRt = fnGetCurrentIntRate(dtpPayePmtDt.value)
            If dblCurrIntRt = gclngNoSelection Then
                dblCurrIntRt = fnPromptForRate(cstrCurrRtInEffectOn & strPayePmtDt & cstrIsNotDefdSupplyARate _
                               & cstrForThe & strDesc & cstrOf & siRatesIn.StCd)
            End If
            dblRateToUse = dblCurrIntRt - siRatesIn.StrlIntRuleAmt
            If dblRateToUse < 0 Then
                ' If the derived rate is negative, ask the user to supply it
                dblRateToUse = fnPromptForRate(cstrDerivedRateBasedOn & cstrCurrRtX & CStr(dblCurrIntRt) & _
                               cstrCloseParenMinus & cstrStateRuleAmtX & CStr(siRatesIn.StrlIntRuleAmt) & _
                               cstrCloseParen & cstrIsNegNbrSupplyRateToUse & _
                               cstrForThe & strDesc & cstrOf & siRatesIn.StCd)
            End If
        
        Case "GTCLN&X "
            dblCurrLoanRt = fnPromptForRate(cstrPleaseSpecify & cstrCurrLoanRt & cstrInEffectOn & strPayePmtDt _
                & cstrForThe & strDesc & cstrOf & siRatesIn.StCd)
            If dblCurrLoanRt > siRatesIn.StrlIntRuleAmt Then
                dblRateToUse = dblCurrLoanRt
            Else
                dblRateToUse = siRatesIn.StrlIntRuleAmt
            End If
            If dblRateToUse < 0 Then
                ' If the derived rate is negative, ask the user to supply it
                dblRateToUse = fnPromptForRate(cstrDerivedRateBasedOn & cstrTheGreaterOf & cstrCurrLoanRtX & CStr(dblCurrLoanRt) & _
                               cstrCloseParen & cstrAnd & cstrStateRuleAmtX & CStr(siRatesIn.StrlIntRuleAmt) & _
                               cstrCloseParen & cstrIsNegNbrSupplyRateToUse & _
                               cstrForThe & strDesc & cstrOf & siRatesIn.StCd)
            End If
            
        Case "GTCRT&LN"
            dblCurrIntRt = fnGetCurrentIntRate(dtpPayePmtDt.value)
            If dblCurrIntRt = gclngNoSelection Then
                dblCurrIntRt = fnPromptForRate(cstrCurrRtInEffectOn & strPayePmtDt & cstrIsNotDefdSupplyARate _
                               & cstrForThe & strDesc & cstrOf & siRatesIn.StCd)
            End If
            dblCurrLoanRt = fnPromptForRate(cstrPleaseSpecify & cstrCurrLoanRt & cstrInEffectOn & strPayePmtDt _
                & cstrForThe & strDesc & cstrOf & siRatesIn.StCd)
            If dblCurrIntRt > dblCurrLoanRt Then
                dblRateToUse = dblCurrIntRt
            Else
                dblRateToUse = dblCurrLoanRt
            End If
            If dblRateToUse < 0 Then
                ' If the derived rate is negative, ask the user to supply it
                dblRateToUse = fnPromptForRate(cstrDerivedRateBasedOn & cstrTheGreaterOf & cstrCurrRtX & CStr(dblCurrIntRt) & _
                               cstrCloseParen & cstrAnd & cstrCurrLoanRtX & CStr(dblCurrLoanRt) & _
                               cstrCloseParen & cstrIsNegNbrSupplyRateToUse & _
                               cstrForThe & strDesc & cstrOf & siRatesIn.StCd)
            End If
        
        Case "GTCRT&X "
            dblCurrIntRt = fnGetCurrentIntRate(dtpPayePmtDt.value)
            If dblCurrIntRt = gclngNoSelection Then
                dblCurrIntRt = fnPromptForRate(cstrCurrRtInEffectOn & strPayePmtDt & cstrIsNotDefdSupplyARate _
                               & cstrForThe & strDesc & cstrOf & siRatesIn.StCd)
            End If
            If dblCurrIntRt > siRatesIn.StrlIntRuleAmt Then
                dblRateToUse = dblCurrIntRt
            Else
                dblRateToUse = siRatesIn.StrlIntRuleAmt
            End If
            If dblRateToUse < 0 Then
                ' If the derived rate is negative, ask the user to supply it
                dblRateToUse = fnPromptForRate(cstrDerivedRateBasedOn & cstrTheGreaterOf & cstrCurrRtX & CStr(dblCurrIntRt) & _
                               cstrCloseParen & cstrAnd & cstrStateRuleAmtX & CStr(siRatesIn.StrlIntRuleAmt) & _
                               cstrCloseParen & cstrIsNegNbrSupplyRateToUse & _
                               cstrForThe & strDesc & cstrOf & siRatesIn.StCd)
            End If
        
        Case "PROMPT  "
            dblRateToUse = fnPromptForRate(cstrPleaseSpecify & "Interest Rate to use" _
                & cstrForThe & strDesc & cstrOf & siRatesIn.StCd)
                
        Case "SPECAMT "
            dblRateToUse = siRatesIn.StrlIntRuleAmt
            If dblRateToUse < 0 Then
                ' If the derived rate is negative, ask the user to supply it
                dblRateToUse = fnPromptForRate(cstrDerivedRateBasedOn & cstrStateRuleAmtX & CStr(siRatesIn.StrlIntRuleAmt) & _
                               cstrCloseParen & cstrIsNegNbrSupplyRateToUse & _
                               cstrForThe & strDesc & cstrOf & siRatesIn.StCd)
            End If
            
        Case Else
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_UNEXPECTED_VAL_SELECT_CASE, _
                mstrScreenName & gcstrDOT & cstrCurrentProc
            GoTo PROC_EXIT
    End Select
        
    ' Per Michelle Wilkosky, the following check should be performed against all interest rates, whether user-supplied,
    ' calculated, or obtained from the STATE_RULE_T table, and whether it represents an interest rate or loan rate.
    If dblRateToUse < 0 Then
        ' (gcRES_WARN_RATE_IS_NEGATIVE (2007) = The Rate supplied or derived from the supplied Rate
        '                                       is a negative number (@@1). Please try again.
        gerhApp.ReportNonFatal vbObjectError + gcRES_WARN_RATE_IS_NEGATIVE, _
                               mstrScreenName & gcstrDOT & cstrCurrentProc, _
                               CStr(dblRateToUse)
        GoTo DETERMINE_RATE
    End If

   ' MME WRUS 4999 - Per Dave O'Connor - the max interest rate check is no longer needed
    ' Per Michelle Wilkosky, the following check should be performed against all interest rates, whether user-supplied,
    ' calculated, or obtained from the STATE_RULE_T table, and whether it represents an interest rate or loan rate.
    If IsNumeric(dblRateToUse) Then
        'If Val(dblRateToUse) > cintMaxInterestRate And siRatesIn.StCd <> "ME" Then
        '    ' 4005 = The interest rate supplied is more than @@1%. This is only allowed when the @@2 is Maine. Please try again.
        '    gerhApp.ReportNonFatal vbObjectError + gcRES_NERR_INTEREST_RATE_TOO_HIGH, _
        '                           mstrScreenName & gcstrDOT & cstrCurrentProc, _
        '                           CStr(cintMaxInterestRate), mcstrCboCalcStCdLabel
        '    GoTo DETERMINE_RATE
        'End If
    Else
        ' gcRES_WARN_NONNUMERIC_RATE (2006) = The Rate supplied or obtained from the STATE_RULE_T table is non-numeric (@@1). Please try again.
        gerhApp.ReportNonFatal vbObjectError + gcRES_WARN_NONNUMERIC_RATE, _
                               mstrScreenName & gcstrDOT & cstrCurrentProc, _
                               CStr(dblRateToUse)
        GoTo DETERMINE_RATE
    End If

    fnGetInterestRate = dblRateToUse
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function


'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub fnGetStateInfo_Override()
    ' Comments  : This function retrieves the calculation rule info
    '             for the Overriden Calc State and/or Interest Rate
    ' Parameters:  -
    ' Returns   :  -
    ' Modified  :
    ' --------------------------------------------------
    Const cstrCurrentProc As String = "fnGetStateInfo_Override"

    On Error GoTo PROC_ERR

    If cboCalcStCd.ListIndex > 0 Then
    
    ' MME START WRUS 4999 - ADDED EXTRA PARAMATERS
    
        fnGetStateInfo cboCalcStCd.Text, _
                       mfrmMyInsuredForm.InsuredLobCd, _
                       DateValue(dtpPayePmtDt.value), _
                       mfrmMyInsuredForm.InsuredClmID, _
                       DblScreenDBPaymentValue, _
                       msiOverride
    Else
        fnInitializeStateInfo msiOverride
    End If
    ' Store Interest Rate in StateInfo structure too
    If IsNumeric(ipdPayeClmIntRt.UnFmtText) Then
        msiOverride.InterestRateToUse = ipdPayeClmIntRt.UnFmtText
    Else
        ' If the user cleared the contents, set the StateInfo structure and
        ' screen field to 0
        msiOverride.InterestRateToUse = 0
        ipdPayeClmIntRt.UnFmtText = 0
    End If
    txtCalcStCdSpecialInstructions_UsedInAutoCalc.Text = msiOverride.StrlSpclInstrTxt
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
Private Sub fnGetStateInfo_InsdDthResStCd()
    ' Comments  : This function retrieves the calculation rule info
    '             for the Insured Residence State at time of death
    ' Parameters:  -
    ' Returns   :  -
    ' Modified  :
    ' --------------------------------------------------
    Const cstrCurrentProc As String = "fnGetStateInfo_InsdDthResStCd"

    On Error GoTo PROC_ERR

    ' If the Foreign Residence At Death checkbox is selected, there's no InsdDthResStCd so bypass
    ' the call to fnGetStateInfo( ) and just initalize the StateInfo structure
    If mfrmMyInsuredForm.InsuredClmForResDthInd Then
        chkClmForResDthInd_UsedInAutoCalc.value = vbChecked
        fnInitializeStateInfo msiInsdDthResStCd
    Else
        chkClmForResDthInd_UsedInAutoCalc.value = vbUnchecked
        
     ' MME START WRUS 4999 - ADDED EXTRA PARAMATERS
     
        fnGetStateInfo mfrmMyInsuredForm.InsuredInsdDthResStCd, _
                       mfrmMyInsuredForm.InsuredLobCd, _
                       DateValue(dtpPayePmtDt.value), _
                       mfrmMyInsuredForm.InsuredClmID, _
                       DblScreenDBPaymentValue, _
                       msiInsdDthResStCd
    End If
    
    With msiInsdDthResStCd
        txtInsdDthResStCd_UsedInAutoCalc.Text = .StCd
        txtInsdDthResStCdSpecialInstructions_UsedInAutoCalc.Text = .StrlSpclInstrTxt
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
Private Sub fnGetStateInfo_IssStCd()
    ' Comments  : This function retrieves the calculation rule info
    '             for the Contract Issue State
    ' Parameters:  -
    ' Returns   :  -
    ' Modified  :
    ' --------------------------------------------------
    Const cstrCurrentProc As String = "fnGetStateInfo_IssStCd"

    On Error GoTo PROC_ERR

' MME START WRUS 4999 - ADDED EXTRA PARAMATERS

    fnGetStateInfo mfrmMyInsuredForm.InsuredIssStCd, _
                   mfrmMyInsuredForm.InsuredLobCd, _
                   DateValue(dtpPayePmtDt.value), _
                   mfrmMyInsuredForm.InsuredClmID, _
                   DblScreenDBPaymentValue, _
                   msiIssStCd
    With msiIssStCd
        txtIssStCd_UsedInAutoCalc.Text = .StCd
        txtIssStCdSpecialInstructions_UsedInAutoCalc.Text = .StrlSpclInstrTxt
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
Private Sub fnGetStateInfo_PayeStCd()
    ' Comments  : This function retrieves the calculation rule info
    '             for the Payee Residence State at time of death
    ' Parameters:  -
    ' Returns   :  -
    ' Modified  :
    ' --------------------------------------------------
    Const cstrCurrentProc As String = "fnGetStateInfo_PayeStCd"

    On Error GoTo PROC_ERR

    If cboPayeStCd.ListIndex > 0 Then
    
    ' MME START WRUS 4999 - ADDED EXTRA PARAMATERS
    '''' '' BZ4999 October 2013 Non US payee - SXS
        If cboPayeStCd.Text <> "ZZ" Then
         fnGetStateInfo cboPayeStCd.Text, _
                       mfrmMyInsuredForm.InsuredLobCd, _
                       DateValue(dtpPayePmtDt.value), _
                       mfrmMyInsuredForm.InsuredClmID, _
                       DblScreenDBPaymentValue, _
                       msiPayeStCd
        End If
    Else
        fnInitializeStateInfo msiPayeStCd
    End If
    txtPayeStCdSpecialInstructions_UsedInAutoCalc.Text = msiPayeStCd.StrlSpclInstrTxt
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
Private Sub fnGetStateInfo_Compact()
    ' Comments  : This function retrieves the calculation rule info
    '             for the Payee Residence State at time of death
    ' Parameters:  -
    ' Returns   :  -
    ' Modified  : Berry Kropiwka -11-06-2019 - New for Compact Filling Calcuation
    ' --------------------------------------------------
    Const cstrCurrentProc As String = "fnGetStateInfo_Compact"

    On Error GoTo PROC_ERR

    If cboPayeStCd.ListIndex > 0 Then
    
    ' MME START WRUS 4999 - ADDED EXTRA PARAMATERS
    '''' '' BZ4999 October 2013 Non US payee - SXS
        If cboPayeStCd.Text <> "ZZ" Then
         fnGetStateInfo cstCompactFilling, _
                       mfrmMyInsuredForm.InsuredLobCd, _
                       DateValue(dtpPayePmtDt.value), _
                       mfrmMyInsuredForm.InsuredClmID, _
                       DblScreenDBPaymentValue, _
                       msiCompactCalc
        End If
    Else
        fnInitializeStateInfo msiCompactCalc
    End If
    txtPayeStCdSpecialInstructions_UsedInAutoCalc.Text = msiCompactCalc.StrlSpclInstrTxt
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
Private Sub fnInitializeCalcInfo(ByRef msiIn As StateInfo)
    ' Comments  : This function will initialize the calculated amount fields in
    '             the specified State Info structure
    ' Parameters: None
    ' Returns   : N/A
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc       As String = "fnInitializeCalcInfo"
    Const cintZero              As Integer = 0

    With msiIn
        .NbrOfDaysToPayInterest = cintZero
        .InterestRateToUse = cintZero
        .ClaimInterestAmt = cintZero
        .WithheldAmt = cintZero
        .TotalForThisPayee = cintZero
        .CalculationInfo = vbNullString
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
Private Sub fnInitializeEditMode()
    ' Description: Enabled/disables command and navigation buttons, as well
    '             as flips on a flag to indicate the record has been edited.
    ' Parameters : N/A
    ' Returns    : N/A
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnInitializeEditMode"
 
    If IsDirty = False Then
        IsDirty = True
        fnSetCommandButtons False
        fnSetNavigationButtons bUnconditionalDisable:=True
        ' Clear the Calculation Override checkbox if the user is editing the record. They must re-select that
        ' checkbox if they want to override it again.
        ' If the checkbox is selected but the corresponding wrapper property is the equivalent of "unchecked"
        ' then assume the user just made it go from Unchecked to Checked and thus don't change anything.
        ' Otherwise, assume the previous calc had been overriden and thus deselect the indicator so the *next*
        ' calc, by default, will not be overriden.
        If (Not mbInAddMode) And (chkPayeDfltOvrdInd.value = vbChecked) And (mtWrapper.PayeDfltOvrdInd) Then
            chkPayeDfltOvrdInd.value = vbUnchecked
            lblWarningAboutOverride.Visible = False
            fnEnableDisableControl ctlIn:=ChkPaye1099Ind, bEnable:=True  '' BZ4999 October 2013 Non US payee - SXS
            fnEnableDisableControl ctlIn:=chkPayeDfltOvrdInd, bEnable:=True
            fnEnableDisableControl ctlIn:=cboCalcStCd, bEnable:=False
            fnEnableDisableControl ctlIn:=ipdPayeClmIntRt, bEnable:=False
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
Private Sub fnLoadCbosForStates()
    '--------------------------------------------------------------------------
    ' Procedure:   fnLoadCbosForStates
    ' Description: Populates the Calc State and [Payee Residence] State
    '              comboboxes
    ' Params:      N/A
    ' Returns:     N/A
    ' Modified:
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc          As String = "fnLoadCbosForStates"
    Dim rstStates                  As ADODB.Recordset
 
    On Error GoTo PROC_ERR

    Set rstStates = fnGetListOfStates()
    
    With cboCalcStCd
        .Clear
        
        ' Add a blank entry as the first entry of the combobox. This will force the user to select
        ' an entry (no default selection) since fnValidData will generate an error if the blank
        ' entry is still selected when the user clicks Update.
        .AddItem gcstrBlankEntry

        fnADORecordSetToComboBox rstIn:=rstStates, _
                                 cboIn:=cboCalcStCd, _
                                 strDisplayColumn:="st_cd", _
                                 bClear:=False
    End With

    With cboPayeStCd
        .Clear
        
        ' Add a blank entry as the first entry of the combobox. This will force the user to select
        ' an entry (no default selection) since fnValidData will generate an error if the blank
        ' entry is still selected when the user clicks Update.
        .AddItem gcstrBlankEntry

        fnADORecordSetToComboBox rstIn:=rstStates, _
                                 cboIn:=cboPayeStCd, _
                                 strDisplayColumn:="st_cd", _
                                 bClear:=False
    End With
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
Private Sub fnLoadCboPayeSsnTinTypCd()
    ' Comments  : Populates CboPayeSsnTinTypCd combo box
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnLoadCboPayeSsnTinTypCd"

    With cboPayeSsnTinTypCd
        .Clear
        .AddItem gcstrBlankEntry          ' blank (default)
        .AddItem mcstrPayeeIsAPerson      ' P = Person
        .AddItem mcstrPayeeIsABusiness    ' B = Business
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



Private Sub fnLoadControls()
    ' Comments  : Populates screen controls with data from recordset
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc       As String = "fnLoadControls"
    Const cstrPayeeIsABusiness  As String = "B"

    With mtWrapper
        ' Deliberately load Date of Payment before other State Codes, so the StateInfo structures will be
        ' set up correctly, with the fewest number of event-driven iterations
        dtpPayePmtDt.value = .PayePmtDt
        
        iptPayeFullNm.Text = .PayeFullNm
        iptPayeCareOfTxt.Text = fnZLSIfNull(.PayeCareOfTxt)
        iptPayeAddrLn1Txt.Text = .PayeAddrLn1Txt
        iptPayeAddrLn2Txt.Text = fnZLSIfNull(.PayeAddrLn2Txt)
        iptPayeCityNmTxt.Text = .PayeCityNmTxt
        cboPayeStCd.Text = .PayeStCd
        iptPayeZipCd.Text = .PayeZipCd
        iptPayeZip4Cd.Text = fnZLSIfNull(.PayeZip4Cd)
        '' BZ4999 October 2013 Non US payee - SXS
        If .Paye1099INTInd Then
           ChkPaye1099Ind = vbChecked
        Else
           ChkPaye1099Ind.value = vbUnchecked
        End If
        
        ' Set the TIN Type before we set the SSN Tin Num. This makes
        ' sure the ipmPayeSsnTinNum control's mask gets set appropriately
        ' so the next code chunk (setting the value of the
        ' ipvPayeSsnTinNum control) will work correctly.
        '
        ' Note: Since this is a nullable field, the cbo must have
        ' a blank entry in it, so the fnZLSIfNull( ) call will work.
        If LenB(.PayeSsnTinTypCd) = 0 Then
            cboPayeSsnTinTypCd.Text = gcstrBlankEntry
        Else
            cboPayeSsnTinTypCd.Text = fnZLSIfNull(.PayeSsnTinTypCd)
        End If
        
        ' NOTE: For MaskEdBox or Input Pro Mask controls, have to do special processing based on
        '       whether or not the field is empty, to avoid a 380 "invalid property value" runtime error.
        '       * If it's empty, temporarily delete the mask, set the value, and then restore
        '         the mask.
        '       * If it's not empty, format the value so it will be "valid" per the .Mask
        '         (for phone numbers, this means inserting a dash between characters 3 and 4).
        If LenB(.PayeSsnTinNum) = 0 Then
            ipmPayeSsnTinNum.Mask = vbNullString
            ipmPayeSsnTinNum.Text = vbNullString
            ipmPayeSsnTinNum.Mask = mcstrUnknownTinTypeMask
        Else
            If .PayeSsnTinTypCd = cstrPayeeIsABusiness Then
                ipmPayeSsnTinNum.Text = fnSSNTIN_AddDash(strIn:=.PayeSsnTinNum, bIsTin:=True)
            Else
                ipmPayeSsnTinNum.Text = fnSSNTIN_AddDash(strIn:=.PayeSsnTinNum, bIsTin:=False)
            End If
        End If
        
        ipdPayeWthldRt.Text = .PayeWthldRt
        
        ' ------------------------------------------------------------------------
        '    The checkPayeDfltOvrdInd, cboCalcStCd and iptPayeClmIntRt and
        '   lblWarningAboutOverride controls are all tied to one another with
        '   regard to their availability and initialization.
        ' ------------------------------------------------------------------------
        cboCalcStCd.Text = .CalcStCd
        ipdPayeClmIntRt.Text = .PayeClmIntRt
        
        If .PayeDfltOvrdInd Then
            chkPayeDfltOvrdInd.Enabled = True
            chkPayeDfltOvrdInd.value = vbChecked
            lblWarningAboutOverride.Visible = True
            fnEnableDisableControl ctlIn:=cboCalcStCd, bEnable:=True
            fnEnableDisableControl ctlIn:=ipdPayeClmIntRt, bEnable:=True
        Else
            chkPayeDfltOvrdInd.Enabled = True
            chkPayeDfltOvrdInd.value = vbUnchecked
            lblWarningAboutOverride.Visible = False
            fnEnableDisableControl ctlIn:=cboCalcStCd, bEnable:=False
            fnEnableDisableControl ctlIn:=ipdPayeClmIntRt, bEnable:=False
        End If
        ' ------------------------------------------------------------------------

        
        ipdPayeIntDaysPdNum.Text = .PayeIntDaysPdNum
        ipcPayeDthbPmtAmt.Text = .PayeDthbPmtAmt
        ipcPayeClmIntAmt.Text = .PayeClmIntAmt
        ipcPayeWthldAmt.Text = .PayeWthldAmt
        ipcPayeClmPdAmt.Text = .PayeClmPdAmt
        
        ' ClmId         isn't shown on-screen
        ' PayeId        isn't shown on-screen
        ' LstUpdDtm     isn't shown on-screen
        ' LstUpdUserId  isn't shown on-screen
    End With
        
    ' No need to load txtPayeStCd_UsedInAutoCalc and txtPayeStCdSpecialInstructions_UsedInAutoCalc,
    ' since they is set when the Payee State (cboPayeStCd) is set.
    
    ' No need to load txtInsdDthResStCd_UsedInAutoCalc and txtInsdDthResStCdSpecialInstructions_UsedInAutoCalc,
    ' since they are set during Form_Load and don't need to be changed.

    ' No need to set txtCalcStCdSpecialInstructions_UsedInAutoCalc; it is set whenever
    ' the txtCalculationState changes.

    ' Make sure Navigation buttons are enabled/disabled based on current record position in the Lookup recordset
    fnSetNavigationButtons bUnconditionalDisable:=False

    ' Update the "record x of y" label
    lblRecordPosition = fnShowRecordPosition(mtWrapper.LookupData)

    If bDebugAppTermination Then
        Debug.Print "   Turning off Update mode in " & mstrScreenName & gcstrDOT & cstrCurrentProc
    End If

    ' Set to False to show there are no pending changes. Loading data to controls above
    ' could trigger fnInitializeEditMode to falsely think there is a pending change.
    IsDirty = False
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
Private Sub fnLoadLpcLookup()
    '--------------------------------------------------------------------------
    ' Procedure:   fnLoadLpcLookup
    ' Description: Populates the specified fpCombo Lookup control using
    '              the mtWrapper's lookup recordset
    ' Params:      N/A
    ' Returns:     N/A
    ' Modified:
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc       As String = "fnLoadLpcLookup"
    Const cintRowDimension      As Integer = 2
    Dim aRows()                 As Variant
    Dim lngRow                  As Long
    On Error GoTo PROC_ERR
    
  
    With lpcLookupName
        .Clear
        .SortState = SortStateSuspend
    
        If mtWrapper.LookupRecordCount <> 0 Then
            aRows = mtWrapper.LookupData_Name()
        End If

        ' Add a blank entry as the first entry of the combobox. This will force the user to select
        ' an entry (no default selection) since fnValidData will generate an error if the blank
        ' entry is still selected when the user clicks Update.
        
        ' Set .Row to -1 so insertion works okay whether or not the fpCombo is sorted
        .Row = gclngNoSelection
        ' 3 columns: PAYE_FULL_NM, PAYE_ID and CLM_ID
        .InsertRow = gcstrBlankEntry & vbTab & gcstrBlankEntry & vbTab & gcstrBlankEntry
        ' Next statement gets a run-time error 9 (subscript out of range) if the aRows array is empty
        For lngRow = 0 To UBound(aRows, cintRowDimension)
            ' There are 3 columns in the array and fpCombo control (indexed 0 thru 2).
            .InsertRow = aRows(0, lngRow) & vbTab & aRows(1, lngRow) & vbTab & aRows(2, lngRow)
        Next
        '.SortState = SortStateActiveReSort
    End With
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0

    ' Clean-up statements go here
    Erase aRows

    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case 9  ' subscript out of range
            Resume Next
        Case Else
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub fnPerformLookup(ByRef lpcIn As LPLib.fpCombo)
    ' Comments  : Retrieves selected record
    ' Parameters: N/A
    ' Called by : lpcLookupClaim_Click
    '             lpcLookupClaim_KeyDown (if Enter was pressed)
    ' Modified  :
    '
    ' --------------------------------------------------
    Const cstrCurrentProc               As String = "fnPerformLookup"
    Dim hrgHourglass                    As chrgHourglass
    Dim lngRecordKeyToRetrieve          As Long

    On Error GoTo PROC_ERR

    With lpcIn
        .Col = 0
    
        '.SetFocus
        .ColFromName = mcstrPayeId
        
        ' If there are no records in the main table maintained by this form,
        ' if the blank entry was selected, or if the user typed in nothing
        ' (i.e. a blank entry in the Lookup box), then skip further processing.
        ' There's nothing to do a lookup on!
        ' If the LookupRecordCount = 0 then we should already be in Add mode
        ' and thus should just stay as we are.
        If (mtWrapper.LookupRecordCount = 0) Or _
            (.ColText = gcstrBlankEntry) Or _
            (.ColText = vbNullString) Then
                GoTo PROC_EXIT
        End If
        
        ' Above GoTo avoids a run-time error 13 (type mismatch) on the next
        ' statement if .ColText = gcstrBlankEntry
        lngRecordKeyToRetrieve = .ColText
        
        ' Restore focus back to the display column
        .ColFromName = mcstrDisplayCol
    End With

    ' Turn on hourglass, in case the lookup is slow
    Set hrgHourglass = New chrgHourglass
    hrgHourglass.value = True

    ' If the user issues a lookup request while in Add mode or while there are
    ' pending changes, then it is interpreted to mean that all pending changes
    ' should be discarded. Hence, turn off Add mode and the IsDirty flag and then
    ' retrieve the selected record.
    If bDebugAppTermination Then
        Debug.Print "   Turning off Add mode in " & mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    mbInAddMode = False
    If bDebugAppTermination Then
        Debug.Print "   Turning off Update mode in " & mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    IsDirty = False
    fnSetAvailabilityOfControls bChangeFocus:=False
    mtWrapper.GetSingleRecord lngKey1:=lngRecordKeyToRetrieve, bSynchLookupRST:=True
    Me.Refresh
    ' Load current record's properties to form's controls, reset navigation buttons
    ' and set "rec x of y" label
    fnLoadControls
    fnSetCommandButtons True
    Me.Refresh
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    
    ' Clean-up statements go here
    If Not (hrgHourglass Is Nothing) Then
        hrgHourglass.value = False
    End If
    fnFreeObject hrgHourglass

    ' Report the error, since this is an event handler
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.ReportFatalError mstrScreenName & gcstrDOT & cstrCurrentProc
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
Private Function fnPromptForRate(ByVal strPromptText As String) As Double
    '--------------------------------------------------------------------------
    ' Procedure:   fnPromptForRate
    ' Description: Prompts the user to supply the Current Loan rate effective on
    '              the Date of Payment
    ' Params:      N/A
    ' Returns:     double, representing the Current Loan Rate
    ' Modified:
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc          As String = "fnPromptForRate"
    Dim strTemp                    As String
    Dim dblRate                    As Double
 
    On Error GoTo PROC_ERR

PROMPT_FOR_RATE:
    strTemp = InputBox(strPromptText)
    If strTemp = vbNullString Then
        ' vbNullString is returned if the user clicked Cancel in the InputBox. Call SaveAppSpecificError so the
        ' error is reported upstream, effecting the cancellation of the Update process.
        ' gcRES_NERR_CALC_WAS_CANCELLED (4010) = The calculation was halted since you clicked Cancel. Your changes have not been saved.
        gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_CALC_WAS_CANCELLED, _
                               mstrScreenName & gcstrDOT & cstrCurrentProc
        GoTo PROC_EXIT
    End If
    
    If IsNumeric(strTemp) Then
        dblRate = CDbl(strTemp)
        If dblRate < 0 Then
            ' gcRES_WARN_RATE_IS_NEGATIVE (2007) = The Rate supplied or derived from the supplied Rate is a negative number (@@1). Please try again.
            gerhApp.ReportNonFatal vbObjectError + gcRES_WARN_RATE_IS_NEGATIVE, _
                                   mstrScreenName & gcstrDOT & cstrCurrentProc, _
                                   CStr(dblRate)
            GoTo PROMPT_FOR_RATE
        End If
        ' Silently round up to 5 decimals, to ensure input can be successfully stored in DB
        dblRate = Round(dblRate, 5)
        'intLengthOfRate = Len(strTemp)
        'varPositionOfDecimal = InStr(1, strTemp, ".", vbTextCompare)
        'If IsNull(varPositionOfDecimal) Then
        '    varPositionOfDecimal = 0
        'End If
        'intMaxDecimalsAllowed = mtWrapper.DecimalPositions(ipdPayeClmIntRt.Tag)
        'If (intLengthOfRate - varPositionOfDecimal) > intMaxDecimalsAllowed Then
        '    ' gcRES_WARN_TOO_MANY_DECIMALS (2008) = The Rate supplied cannot have more than @@1 decimal positions specified. Please try again.
        '    gerhApp.ReportNonFatal vbObjectError + gcRES_WARN_TOO_MANY_DECIMALS, _
        '                           mstrScreenName & gcstrDOT & cstrCurrentProc, _
        '                           intMaxDecimalsAllowed
        '    GoTo PROMPT_FOR_RATE
        'End If
    Else
        ' gcRES_WARN_NONNUMERIC_RATE (2006) = The Rate supplied or obtained from the STATE_RULE_T table is non-numeric (@@1). Please try again.
        gerhApp.ReportNonFatal vbObjectError + gcRES_WARN_NONNUMERIC_RATE, _
                               mstrScreenName & gcstrDOT & cstrCurrentProc, _
                               strTemp
        GoTo PROMPT_FOR_RATE
    End If
    
    fnPromptForRate = dblRate
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    
    ' Clean-up statements go here
    
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function


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

    fnLoadLpcLookup             ' Payee Full Name (PAYE_FULL_NM, PAYE_ID, CLM_ID)
    fnLoadCboPayeSsnTinTypCd    ' SSN Tin Type (single column: P or B or blank)
    fnLoadCbosForStates         ' Calc State / Payee State (ST_CD)
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
Private Sub fnResetStateRules()
    '--------------------------------------------------------------------------
    ' Procedure:   fnResetStateRules
    ' Description: Resets all rules. This function should be called if the
    '              Payee's Date of Payment has been changed.
    '
    ' Params:      N/A
    ' Called by:   fnClearControls
    '              dtpPayePmgDt_Change() of frmPayee
    '
    ' Returns:     N/A
    ' Modifed:     Berry Kropiwka - 11-06-2019 - Add fngetstateinfo_compact for compact filling
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc       As String = "fnResetStateRules"
    On Error GoTo PROC_ERR

    ' Populate StateInfo structures with data from the STATE_RULE_T
    ' row that matches the various State Codes.
    fnGetStateInfo_InsdDthResStCd
    fnGetStateInfo_IssStCd
    fnGetStateInfo_PayeStCd
    fnGetStateInfo_Override
    If mfrmMyInsuredForm.chkClmCmpCalInd.value = vbChecked Then
        ' This is Compact Calcatution
        fnGetStateInfo_Compact
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
Private Sub fnSetAvailabilityOfControls(Optional ByVal bChangeFocus As Boolean = True)
    '--------------------------------------------------------------------------
    ' Procedure:   fnSetAvailabilityOfControls
    ' Description: Determines whether a control representing a lookup
    '              or a key field should be display-only.
    '
    ' Params:      bChangeFocus - If True, moves the focus to the first updateable field.
    '
    ' Called by:   cmdUpdate_Click
    '              fnAddRecord
    '              Form_Load
    '              Form_QueryUnload
    ' Returns:     N/A
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc       As String = "fnSetAvailabilityOfControls"
    Dim ctl                     As Control
 
    On Error GoTo PROC_ERR

    For Each ctl In Me.Controls
        With ctl
            ' Debug.Print ctl.Name & vbTab & ctl.Tag
            
            ' If the control corresponds to a SQL Server table column that's a key field, then
            ' only enable it if in Add mode.
            If Len(.Tag) > 0 Then
                ' If it's a key, disable it unless we're in Add mode
                If mtWrapper.IsKey(.Tag) Then
                    'Debug.Print .Tag & " is a key field, per meta data"
                    If mbInAddMode Then
                        fnEnableDisableControl ctlIn:=ctl, bEnable:=True
                    Else
                        fnEnableDisableControl ctlIn:=ctl, bEnable:=False
                    End If
                End If
            End If
        End With
    Next ctl
    
    If bChangeFocus Then
        fnSetFocusToFirstUpdateableField
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
Private Sub fnSetCommandButtons(ByVal bEnable As Boolean)
    '----------------------------------------------------------------------------
    ' Procedure : fnSetCommandButtons
    '
    ' Comments  : Enables/Disables the command buttons, per boolean parameter
    '             Here's how the button enabling should work. Note it assumes
    '             that IsDirty and mbInAddMode have been set prior to
    '             calling this routine, e.g., they accurately reflect whether
    '             or not there are edits outstanding and/or the user is in
    '             Add mode, respectively.
    '             Remember, though: mbInAddMode and IsDirty are
    '             independent of one another!
    '
    '     State          ADD btn  UPD btn  DEL btn  CLOSE btn
    '    --------------  -------- -------- -------- ---------
    '    Add mode       disabled  enabled  disabled enabled
    '    (no edits yet)
    '
    '    Edits o/s      disabled  enabled  disabled enabled
    '
    '    No edits o/s   enabled   disabled enabled  enabled

    '
    ' Called by : fnAddRecord and fnInitializeEditMode, with bEnable = False
    '
    '             lpcLookupName_Click, cmdDelete_Click, cmdNavigate_Click, cmdUpdate_Click
    '             (when updating existing record) and Form_Load, with
    '             bEnable = True
    '
    ' Parameters: bEnable - indicates whether Add/Update buttons should be enabled
    '                       or disabled
    '
    ' Modified  :
    '----------------------------------------------------------------------------
    Const cstrCurrentProc    As String = "fnSetCommandButtons"
    Dim strDependent_Table   As String
    Dim bHaveDependents      As Boolean
    
    On Error GoTo PROC_ERR
    
    ' Hide updates to the window until we're done. This avoids ugly screen flickering
    fnWindowLock Me.hWnd

    cmdAdd.Enabled = bEnable
    cmdUpdate.Enabled = Not bEnable

    If mbInAddMode Then
        bHaveDependents = False
    Else
        bHaveDependents = mtWrapper.HaveDependents(mtWrapper.ClmId, strDependent_Table)
    End If

    ' Can only delete a record when (a) when you're not in the middle of an Add or Update
    ' and (b) there are no rows in dependent tables (i.e. children).
    If (IsDirty Or mbInAddMode) Then
        cmdDelete.Enabled = False
    Else
        If (bHaveDependents) Then
            cmdDelete.Enabled = False
        Else
            cmdDelete.Enabled = True
        End If
    End If

    ' Can only use to the Clone This Payee button when you're NOT in the middle of
    ' an Add or Update.
    If (Not IsDirty) And (Not mbInAddMode) Then
        cmdCloneThisPayee.Enabled = True
    Else
        cmdCloneThisPayee.Enabled = False
    End If
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    
    ' Clean-up statements go here
    fnWindowUnlock
    
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
Private Sub fnSetDefaultControlProperties()
    '--------------------------------------------------------------------------
    ' Procedure:   fnSetDefaultControlProperties
    ' Description: Sets default properties of controls bound to table columns
    '              in the table wrapper class, using the meta data that class
    '              gathered.
    '
    '              These defaults are initially based on the data type
    '              of the column (see the table wrapper's fnGetColMetaData method)
    '              but then overriden, if desired, in the table wrapper's
    '              fnLoadColMetaData method.
    '
    '              NOTE: Tags should only be present if the control
    '                    is bound to a property of the table wrapper class.
    '                    Also, the entire contents of the Tag should be the
    '                    name of the public property in that class.
    '
    ' Params:      N/A
    ' Called by:   Form_Load
    ' Returns:     N/A
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc       As String = "fnSetDefaultControlProperties"
    Dim ctl                     As Control
    Dim bSavedIsDirty           As Boolean
 
    On Error GoTo PROC_ERR

    ' This procedure can be called, among other places, by Form_Load before
    ' the screen controls have been loaded with values. Setting some of
    ' those controls' properties can trigger their Change event which
    ' causes fnSetCommandButtons and ultimately the table wrapper's
    ' HaveDependents procs to be called. The latter can fail with a
    ' spurious error if the key hasn't been set. Since we really don't
    ' care about this processing because it'll be hit again after
    ' data *has* been loaded to the controls, let's just fake it out
    ' here by making any Change event hit by this proc's code
    ' think that "IsDirty" processing has already been done. We'll restore
    ' the IsDirty flag when we're done.

    ' Start of "fake out"
    If bDebugAppTermination Then
        Debug.Print "   Saving then faking out Update mode in " & mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    
    bSavedIsDirty = IsDirty
    IsDirty = True

    For Each ctl In Me.Controls
        With ctl
             ' Debug.Print ctl.Name
              ' If ctl.Name = "fraPayeeInfo" Then
              '    Debug.Print "yes"
              ' End If
            ' If control corresponds to a SQL Server table column, then try
            ' to set its default properties. The Tag property contains
            ' the name of its property within the table class.
            If Len(.Tag) > 0 Then
                ' If it's a key, disable it unless we're in Add mode
                If mtWrapper.IsKey(.Tag) Then
                    .Enabled = mbInAddMode
                End If

                If TypeOf ctl Is TextBox Then
                    .MaxLength = mtWrapper.MaxCharacters(.Tag)
                End If
                
                If TypeOf ctl Is fpText Then
                    .MaxLength = mtWrapper.MaxCharacters(.Tag)
                    ' Convert VB True (-1) to 1 (AutoCaseUpper) and vb False (0) to 0 (AutoCaseNone)
                    .AutoCase = Abs(mtWrapper.ShouldForceToUppercase(.Tag))
                    .CharValidationText = mtWrapper.AllowableCharacters(.Tag)
                End If
                
                If TypeOf ctl Is fpDateTime Then
                    .CalGrayAreaStyle 2             ' Show days from previous/next month if possible
                    .CalGrayAreaAllowScroll True    ' Allow user to scroll by clicking in gray area of calendar
                    .InvalidOption = ShowData       ' Show invalid date w/ diff bkgrd color; don't auto-correct it
                    .PopUpType = PopCalendar
                    .UserEntry = UserEntryFormatted
                    .ButtonStyle = ButtonStyleDropDown
                    .DateTimeFormat = IntlShortDate
                End If
                
                'If TypeOf ctl Is MaskEdBox Then
                '    .MaxLength = mtWrapper.MaxCharacters(.Tag)
                '    .Mask = mtWrapper.Mask(.Tag)
                'End If
                
                If TypeOf ctl Is fpMask Then
                    .MaxLength = mtWrapper.MaxCharacters(.Tag)
                    .Mask = mtWrapper.Mask(.Tag)
                End If
            End If
            ' Make all fpCurrency controls have the same formatting  to start with.
            If TypeOf ctl Is fpCurrency Then
                .AlignTextH = AlignTextHRight
                .AllowNull = True
                .NullColor = vbRed
                .BackColor = vbWindowBackground
                .ForeColor = vbWindowText
                .InvalidColor = vbWindowText
                .CurrencyNegFormat = pd1p    ' ($1)
                .LeadZero = NoLeadingZero
                .UseSeparator = True
                .OnFocusNoSelect = False
                .OnFocusAlignH = OnFocusAlignHRight
                .NegToggle = False          ' Cannot use "-" on numeric keypad to toggle btwn pos/neg
                .MinValue = 0               ' Negatives not allowed
                .NoSpecialKeys = AllKeysEnabled
                ' Next line is needed to avoid -2147217887 (Invalid character value for cast specification) error
                ' when storing too big a value into a SQL column.
                .MaxValue = fnTranslateToMaxValue(mtWrapper.DollarPositions(.Tag), mtWrapper.DecimalPositions(.Tag))
            End If
            ' Make all fpDoubleSingle controls have the same formatting  to start with.
            If TypeOf ctl Is fpDoubleSingle Then
                .AlignTextH = AlignTextHRight
                .AllowNull = True
                .NullColor = vbRed
                .BackColor = vbWindowBackground
                .ForeColor = vbWindowText
                .InvalidColor = vbWindowText
                .LeadZero = NoLeadingZero
                .UseSeparator = True
                .OnFocusNoSelect = False
                .OnFocusAlignH = OnFocusAlignHRight
                .DecimalPlaces = mtWrapper.DecimalPositions(.Tag)
                .FixedPoint = True
                .NegFormat = n1      ' (1)
                .NegToggle = False   ' Cannot use "-" on numeric keypad to toggle btwn pos/neg
                .MinValue = 0        ' Negatives not allowed
                .NoSpecialKeys = AllKeysEnabled
                ' Next line is needed to avoid -2147217887 (Invalid character value for cast specification) error
                ' when storing too big a value into a SQL column.
                .MaxValue = fnTranslateToMaxValue(mtWrapper.DollarPositions(.Tag), mtWrapper.DecimalPositions(.Tag))
            End If
            ' Make all fpMask controls have the same formatting to start with.
            If TypeOf ctl Is fpMask Then
                .AlignTextH = AlignTextHLeft
                .AllowNull = False
                .NullColor = vbRed
                .BackColor = vbWindowBackground
                .ForeColor = vbWindowText
                .InvalidColor = vbWindowText
                .HideSelection = True
                .OnFocusNoSelect = False
                .OnFocusAlignH = OnFocusAlignHLeft
                .PromptChar = "_"            ' Char to display in unfilled positions?
                .PromptInclude = False       ' Include Prompt char when saving bound value to DB?
                .RequireFill = 0             ' Trigger event if all prompt chars not supplied when ctl loses focus?
            End If
            ' Make all fpText controls have the same formatting to start with.
            ' NOTE: Be sure not to un-do any settings done earlier in this procedure, for fpText
            '       for controls bound to a table column!!
            If TypeOf ctl Is fpText Then
                .AlignTextH = AlignTextHLeft ' How to align horizontally?
                .AllowNull = False           ' Is Null a valid value? User can press Ctrl-N or F2 to insert a Null value.
                .NullColor = vbRed           ' Color of contents when value is Null
                .BackColor = vbWindowBackground
                .ForeColor = vbWindowText
                .InvalidColor = vbWindowText
                .HideSelection = True        ' contents selected when ctl loses focus?
                .OnFocusNoSelect = False     ' don't select contents when ctl receives focus?
                .MultiLine = False
            End If
        End With
    Next ctl
    
    ' Per Michelle, the Date of Payment can be future-dated up to 5 days.
    ''''''''''''''''dtpPayePmtDt.MaxDate = DateAdd("d", 5, Now)  '''''' BZ 6495 SXS
    dtpPayePmtDt.MaxDate = DateAdd("d", 30, Now)
    
    ipdPayeWthldRt.MaxValue = 99#        ' 99.00000
    
    ' End of "fake out"
    If bDebugAppTermination Then
        Debug.Print "   Restoring saved Update mode afer fake out in " & mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    
    IsDirty = bSavedIsDirty
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
        Case 5
            ' Invalid Procedure Call or Argument  (See MSKB Article Q242347)
            Resume Next
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
    '              lpcLookupName_Click( )
    '
    ' Returns   :  N/A
    ' Modified  :
    '----------------------------------------------------------------------------
    Const cstrCurrentProc As String = "fnSetNavigationButtons"
    Dim cmd               As CommandButton
    Dim bHaveRecords      As Boolean

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
Private Sub fnSetupScreenControls()
    '--------------------------------------------------------------------------
    ' Procedure:   fnSetupScreenControls
    ' Description: This procedure:
    '              * Binds the on-screen controls to the table wrapper class
    '                properties with which they are associated.
    '              * Sets default settings for those controls' properties
    '              * Binds editable TextBoxes controls to the Extended TextBox
    '                class so they will behave appropriately and in a consistent
    '                manner.
    '
    ' Params:      N/A
    ' Returns:     N/A
    ' Date:        04/04/2002
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc          As String = "fnSetupScreenControls"
    On Error GoTo PROC_ERR
 
    ' Set each control's Tag property to identify the table class property to which it corresponds,
    ' set its defaults attributes per the DBMS' meta data. In addition, for those columns that correspond
    ' to editable TextBox controls, set properties of its associated ExtendedTextBox variable, so
    ' it will behave appropriately and in a standard manner.
    fnBindControlsToTableWrapper
    
    ' Set default attributes for those controls, per the DBMS' meta data
    fnSetDefaultControlProperties
    
    ' Disable controls that are always "display-only"
    fnEnableDisableControl ctlIn:=txtPayeStCd_UsedInAutoCalc, bEnable:=False
    fnEnableDisableControl ctlIn:=txtPayeStCdSpecialInstructions_UsedInAutoCalc, bEnable:=False
    fnEnableDisableControl ctlIn:=chkClmForResDthInd_UsedInAutoCalc, bEnable:=False
    fnEnableDisableControl ctlIn:=txtInsdDthResStCd_UsedInAutoCalc, bEnable:=False
    fnEnableDisableControl ctlIn:=txtInsdDthResStCdSpecialInstructions_UsedInAutoCalc, bEnable:=False
    fnEnableDisableControl ctlIn:=txtIssStCd_UsedInAutoCalc, bEnable:=False
    fnEnableDisableControl ctlIn:=txtIssStCdSpecialInstructions_UsedInAutoCalc, bEnable:=False
    fnEnableDisableControl ctlIn:=txtCalcStCdSpecialInstructions_UsedInAutoCalc, bEnable:=False
    fnEnableDisableControl ctlIn:=ipdPayeIntDaysPdNum, bEnable:=False
    fnEnableDisableControl ctlIn:=ipcPayeClmIntAmt, bEnable:=False
    fnEnableDisableControl ctlIn:=ipcPayeWthldAmt, bEnable:=False
    fnEnableDisableControl ctlIn:=ipcPayeClmPdAmt, bEnable:=False
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
    Const cstrCurrentProc                   As String = "fnValidData"
    Const cintMinZipLength                  As Integer = 5
    Const cintMaxZipLength                  As Integer = 9
    Const cintTinLengthIfInput              As Integer = 9
    Dim bErrorFound                         As Boolean
    Dim ctl                                 As Control
    Dim ctlFirstToFail                      As Control
    Dim intFailures                         As Integer
    Dim strFieldList                        As String
    Dim strMsgText                          As String
    Dim intLengthToTest                     As Integer
    

    fnValidData = True

    ' Check the fields in a left-to-right, top-to-bottom screen sequence.
    '     1. iptPayeFullNm        9. ipmPayeSsnTinNum
    '     2. iptPayeCareOfTxt    10. cboPayeSsnTinTypCd
    '     3. iptPayeAddrLn1Txt   11. ipdPayeWthldRt
    '     4. iptPayeAddrLn2Txt   12. chkPayeDftOvrdInd
    '     5. iptPayeCityNmTxt    13. cboCalcStCd
    '     6. cboPayeStCd         14. ipdPayeClmIntRt
    '     7. iptPayeZipCd        15. dtpPayePmtDt
    '     8. iptPayeZip4Cd       16. ipcPayeDthbPmtAmt
    '                            17. chkpaye1099ind
    ' ------------- 1.  Verify required fields are missing --------------
    ' Check key fields too, although they should be absent only if
    ' the user is in Add mode and neglected to specify their values.
 
    ' Using Metadata, verify that all fields are populated if required to (i.e. IsNullable() is True)
    For Each ctl In Me.Controls
        With ctl
            ' Debug.Print ctl.Name
            
            ' If the control corresponds to a SQL Server table column, then determine if its
            ' Not Nullable (i.e. required). The Tag property contains the name of its
            ' property within the table class. ' If it's Not Nullable and not input...then
            ' generate an error
            ' Skip over the control that is bound to PAYE_ID since this is a hidden
            ' field and thus the user shouldn't be informed if it hasn't been set yet.
            ' Skip over the control that is bound to CALC_ST_CD since this, unless overriden,
            ' is a disabled field that is set automatically during the Update processing.
            If Len(.Tag) > 0 And (.Tag <> "PayeId") And (.Tag <> "CalcStCd") Then
                If Not (mtWrapper.IsNullable(.Tag)) Then
                    If TypeOf ctl Is fpCombo Then
                        ' Special handling for fpCombo since its default property
                        ' isn't the one that must be checked
                        If (Len(ctl.ColText) = 0) Or (ctl.ColText = gcstrBlankEntry) Then
                            If intFailures = 0 Then
                                strFieldList = vbCrLf & fnGetFieldLabel(ctl.Name)
                                Set ctlFirstToFail = ctl
                            Else
                                strFieldList = strFieldList & vbCrLf & fnGetFieldLabel(ctl.Name)
                            End If
                            intFailures = intFailures + 1
                        End If
                    Else
                         '' BZ4999 October 2013 Non US payee - SXS
                        If fnGetFieldLabel(ctl.Name) = "Zip" And ChkPaye1099Ind = 0 Then
                        Else
                          If (Len(ctl) = 0) Or (ctl = gcstrBlankEntry) Or _
                                          ((IsNull(ctl)) And fnGetFieldLabel(ctl.Name) = "Zip") Then
                            If intFailures = 0 Then
                                strFieldList = vbCrLf & fnGetFieldLabel(ctl.Name)
                                     
                                Set ctlFirstToFail = ctl
                            Else
                                strFieldList = strFieldList & vbCrLf & fnGetFieldLabel(ctl.Name)
                            End If
                            intFailures = intFailures + 1
                          End If
                        End If
                    End If
                 End If
            End If
        End With
    Next ctl

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



    ' ------------- 2.  Verify other characteristics are valid --------------

    ' Reset for this section of error validations
    strMsgText = vbNullString
    intFailures = 0

    ' The Zip Code must be input, either as a 5-digit or 9-digit (zip+4) number
        intLengthToTest = Len(ipmPayeSsnTinNum.UnFmtText)
       If (intLengthToTest <> 0) And (intLengthToTest = cintMinZipLength) Or (intLengthToTest = cintMaxZipLength) Then
        ' Do Nothing
       Else
        intFailures = intFailures + 1
        Set ctlFirstToFail = iptPayeZipCd
        strMsgText = strMsgText & vbCrLf & _
                     "The " & mcstrIptPayeZipCdLabel & " must be either " & cintMinZipLength & " or " & _
                     cintMaxZipLength & " digits."
       End If
    ' End If

    ' If input, the SSN/Tin must be 9 digits
    intLengthToTest = Len(ipmPayeSsnTinNum.UnFmtText)
    If (intLengthToTest <> 0) And (intLengthToTest <> cintTinLengthIfInput) Then
        intFailures = intFailures + 1
        Set ctlFirstToFail = ipmPayeSsnTinNum
        strMsgText = strMsgText & vbCrLf & _
                     "The " & mcstrIpmPayeSsnTinNumLabel & " must be " & cintTinLengthIfInput & " digits."
    End If

    ' If either the SSN/TIN or SSN/TIN Type is input, then both must be input
    If ((Len(ipmPayeSsnTinNum.UnFmtText) <> 0) And (cboPayeSsnTinTypCd.Text <> gcstrBlankEntry)) Or _
        ((Len(ipmPayeSsnTinNum.UnFmtText) = 0) And (cboPayeSsnTinTypCd.Text = gcstrBlankEntry)) Then
            ' Okay
    Else
        intFailures = intFailures + 1
        Set ctlFirstToFail = ipmPayeSsnTinNum
        strMsgText = strMsgText & vbCrLf & _
                     "If either the " & mcstrIpmPayeSsnTinNumLabel & " or " & mcstrCboPayeSsnTinTypCdLabel & _
                     " is input, then both must be input."
    End If

    ' The Date of Payment must be on or after the Insured's Date of Death
    If DateValue(dtpPayePmtDt.value) < DateValue(mfrmMyInsuredForm.InsuredClmInsdDthDt) Then
        intFailures = intFailures + 1
        Set ctlFirstToFail = dtpPayePmtDt
        strMsgText = strMsgText & vbCrLf & _
                     "The " & mcstrDtpPayePmtDtLabel & " (" & dtpPayePmtDt.value & _
                     ") must be on or after the Insured's " & mcstrDtpClmInsdDthDtLabel & _
                     " (" & mfrmMyInsuredForm.InsuredClmInsdDthDt & ")."
    End If

    ' The Date of Payment must be on or after the Date of Proof
    If DateValue(dtpPayePmtDt.value) < DateValue(mfrmMyInsuredForm.InsuredClmProofDt) Then
        intFailures = intFailures + 1
        Set ctlFirstToFail = dtpPayePmtDt
        strMsgText = strMsgText & vbCrLf & _
                     "The " & mcstrDtpPayePmtDtLabel & " (" & dtpPayePmtDt.value & _
                     ") must be on or after the Insured's " & mcstrDtpClmProofDtLabel & _
                     " (" & mfrmMyInsuredForm.InsuredClmProofDt & ")."
    End If

     
    ' If the user is overriding the Calc State or Interest Rate, then one but only one
    ' of those fields can be input.
    If chkPayeDfltOvrdInd.value = vbChecked Then
        If (cboCalcStCd.Text = gcstrBlankEntry) And (Len(ipdPayeClmIntRt.UnFmtText) = 0) Then
            intFailures = intFailures + 1
            Set ctlFirstToFail = cboCalcStCd
            strMsgText = strMsgText & vbCrLf & _
                         "If the " & mcstrChkPayeDfltOvrdIndLabel & " is selected, then both the " & _
                         mcstrCboCalcStCdLabel & " and the " & mcstrIpdPayeClmIntRtLabel & " must be input."
        End If
    End If
     
     
    ' DB Payment must be a positive non-zero amount.
    If CDbl(ipcPayeDthbPmtAmt.UnFmtText) <= 0 Then
        intFailures = intFailures + 1
        Set ctlFirstToFail = ipcPayeDthbPmtAmt
        strMsgText = strMsgText & vbCrLf & _
                     "The " & mcstrIpcPayeDthbPmtAmtLabel & " must be supplied as a positive non-zero amount."
    End If

'!TODO! This can go away if the control enforces this upon data entry (and for pasting)
    If Not (IsNumeric(ipdPayeClmIntRt.Text)) Then
        intFailures = intFailures + 1
        Set ctlFirstToFail = ipdPayeClmIntRt
        strMsgText = strMsgText & vbCrLf & _
                     "The " & mcstrIpdPayeClmIntRtLabel & " must be numeric."
    End If

   
    ' Validation for Payee Interest Rate (> 12 and CalcState <> "ME") relocated to
    ' fnGetInterestRate since this is a protected field and thus needs to be validated
    ' when the user supplies it...from fnGetInterestRate.
    
    If intFailures <> 0 Then
        bErrorFound = True
        fnValidData = False
        If ctlFirstToFail.Visible Then
            ctlFirstToFail.SetFocus
        End If
        gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_CROSS_FLD_VALIDATIONS_FAILED, _
                                   mstrScreenName & gcstrDOT & cstrCurrentProc, _
                                   "your request can be processed", strMsgText
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
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case 5
            ' Invalid Procedure Call or Argument  (See MSKB Article Q242347)
            Resume Next
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
    Const cstrCurrentProc   As String = "fnWarningData"
    Const cdbl99Million     As Double = 99999999#

    If ipcPayeDthbPmtAmt.value > cdbl99Million Then
        'gcRES_WARN_DTHB_PMT_AMT_MAY_BE_TOO_HIGH (2005) = The @@1 exceeds @@2. Please verify this amount is correct.
        gerhApp.ReportNonFatal vbObjectError + gcRES_WARN_DTHB_PMT_AMT_MAY_BE_TOO_HIGH, _
                               mstrScreenName & gcstrDOT & cstrCurrentProc, _
                               mcstrIpcPayeDthbPmtAmtLabel, "$99m"
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
Private Sub Form_Activate()
    ' Comments  : I think this is an unnecessary event handler! (Betsy 05/14/2001)
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "Form_Activate"

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

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
        Case Else
            ' Save Err object data, if not already saved
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub


'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub Form_Initialize()
    ' Comments  : Intializes the form.
    ' Parameters: None
    ' Modified  :
    ' 01/2002 BAW - Populate the new Insured Residence State text box.
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "Form_Initialize"

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    ' Once the Payee screen becomes the active form, we lose the ability to reference
    ' the Insured form using "frmInsured" (Remember, in an MDI environment, there
    ' could be multiple instances of an Insured form loaded). So, while it still is
    ' the active form, set a reference to it.
    Set mfrmMyInsuredForm = frmMDIMain.ActiveForm
    
    ' Set "Claim#" caption on the form
    lblClmNum = mfrmMyInsuredForm.InsuredClmNum
    
    'Y027 07-11-2012
    m_AdmPolicySystem = mfrmMyInsuredForm.AdminSystemCode
    
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
    ' Parameters: None
    ' Modified  :
    '   01/2002 BAW - Populate the new Insured Residence State Special Instructions
    '                 field, based on the state set on the Insured screen.
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

    '...............................................................................
    ' Set our fpCombo control settings, for those used as Lookups. Since these
    ' contain lots of rows (5000 or so, currently), they are loaded with sorted data
    ' rather than having the control itself sort its contents. This GREATLY improves
    ' the time it takes to display the form and refresh the control!
    '...............................................................................
    ' 1. Name Lookup
    With lpcLookupName
        fnInitializefpCombo lpcIn:=lpcLookupName, bShowColHeaders:=False, bSortable:=False, _
            lngNbrOfCols:=2, lngEditCol:=mcintDisplayCol_lpcLookupName, lngNbrOfRowsInDropdown:=8
        ' Column definitions
        .Col = 0                                            ' First column, Primary sort
            .ColHeaderText = mcstrIptPayeFullNmLabel
            .ColName = mcstrDisplayCol
        .Col = 1                                            ' Second column
            .ColHeaderText = mcstrTxtPayeIDLabel
            .ColName = mcstrPayeId
            .ColHide = True
        .ColumnSearch = mcintDisplayCol_lpcLookupName
    End With


    ' Set the control to receive the focus after errors (the first editable field
    ' on the screen), dependent upon whether we're in Add Mode or not. If in Add mode,
    ' this control would typically be the first control that corresponds to a Key field.
    ' If not in Add mode, this control would typically be the topmost/leftmost
    ' "always updateable" control on the screen (excepting the Lookup ComboBox).
    Set mctlFirstUpdateableField_Add = iptPayeFullNm
    Set mctlFirstUpdateableField_Upd = iptPayeFullNm

    ' Instantiate and initialize a table wrapper object for the appropriate table(s).
    Set mtWrapper = New ctpyePayee
    mtWrapper.InitPayee mfrmMyInsuredForm.InsuredClmID

    ' Populate the Insured's state of residence (at time of death) and its corresponding
    ' Special Instructions.
    txtInsdDthResStCd_UsedInAutoCalc.Text = mfrmMyInsuredForm.InsuredInsdDthResStCd
    txtInsdDthResStCdSpecialInstructions_UsedInAutoCalc.Text = msiInsdDthResStCd.StrlSpclInstrTxt


    ' Bind the on-screen controls to the table wrapper class properties with which they
    ' are associated. set default settings for those controls' properties, and
    ' bind editable TextBoxes controls to the Extended TextBox class so they will
    ' behave appropriately and in a consistent manner.
    fnSetupScreenControls
    
    ' Populate all ComboBoxes and ListPro controls
    fnRefreshAllCombos

    ' "And" condition added 05/22/01 for bug 0033, so this screen can be invoked
    ' through the Insured screen's msgPayees grid control (to edit an existing Payee)
    ' as identified by a non-empty InsuredCurrentPayeeName field or to add a new Payee
    ' by the user clicking the Add Payees button on the Insured screen (identified
    ' by an *empty* InsuredCurrentPayeeName field).
    With mtWrapper
        If .LookupRecordCount > 0 And mfrmMyInsuredForm.InsuredCurrentPayeeName <> vbNullString Then
            ' Pull up the Payee on whose name they double-clicked in the Payee grid of the Insured screen
            .GoToFirstRecord
            .GetSingleRecord lngKey1:=mfrmMyInsuredForm.InsuredCurrentPayeeID, bSynchLookupRST:=True
            fnLoadControls
            
            ' Populate StateInfo structures with data from the STATE_RULE_T
            ' row that matches the State Codes as set on the Insured screen.
            fnGetStateInfo_InsdDthResStCd
            fnGetStateInfo_IssStCd
            
            fnSetCommandButtons True
        Else
            ' Populate StateInfo structures with data from the STATE_RULE_T
            ' row that matches the State Codes as set on the Insured screen.
            fnGetStateInfo_InsdDthResStCd
            fnGetStateInfo_IssStCd
        
            ' Go into "Add" mode. (The user clicked the Add Payee button on the Insured screen)
            fnAddRecord
        End If
    End With

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

    If bDebugAppTermination Then
        Debug.Print "Entering " & mstrScreenName & gcstrDOT & cstrCurrentProc
    End If

    If gbAmProcessingAnAppFatalError Then
        ' ALWAYS let the form be unloaded, with no prompts to the user, if shutting
        ' down the app due to an application fatal error having been hit.
        If bDebugAppTermination Then
            Debug.Print "   Early exit from " & mstrScreenName & gcstrDOT & cstrCurrentProc & " since processing a fatal error."
        End If
        GoTo PROC_EXIT
    End If
    
    If (Not mbInAddMode) And (Not IsDirty) Then
        ' Let the form be closed if the user is in neither Add nor Update mode.
        ' Let the form be closed if the user is in neither Add nor Update mode.
        If bDebugAppTermination Then
            Debug.Print "   Early exit from " & mstrScreenName & gcstrDOT & cstrCurrentProc & " since not in Add or Update mode."
        End If
        GoTo PROC_EXIT
    End If

    ' Since Update (IsDirty) mode can be True while in Add mode, we must check for Add mode first.
    ' Otherwise, Adds where the user has started typing (thus setting IsDirty to True) will be
    ' treated like an Update, when it should be treated like an Add.
    If mbInAddMode Then
        If IsDirty Then
            If bDebugAppTermination Then
                Debug.Print "   Add/Update mode: Prompt the user re: okay to discard pending changes, in " & mstrScreenName & gcstrDOT & cstrCurrentProc
            End If
            intButtonClicked = gerhApp.ReportNonFatal(vbObjectError + gcRES_ALRT_CHANGES_PENDING, _
                           mstrScreenName & gcstrDOT & cstrCurrentProc)
        Else
            If bDebugAppTermination Then
                Debug.Print "   Add mode only (not Update): Do not prompt the user re: okay to discard pending changes, in " & mstrScreenName & gcstrDOT & cstrCurrentProc
            End If
            intButtonClicked = vbYes
        End If
        If intButtonClicked = vbYes Then
            If bDebugAppTermination Then
                Debug.Print "      User opted to discard pending changes in " & mstrScreenName & gcstrDOT & cstrCurrentProc
            End If
        
            ' If they want to abandon an Add before they started data entry, let them!
            ' Redisplay the form with the *first* record now showing
            mtWrapper.GetLookupData mfrmMyInsuredForm.InsuredClmID
            If mtWrapper.LookupIsAtBOF And mtWrapper.LookupIsAtEOF Then
                ' There are no records in the table, so let the form close (If we went into Add
                ' mode, the user would never be able to exit the screen!)
            Else
                If Not gbAmTryingToTerminateTheApp Then
                    pintCancel = True
                    mtWrapper.GoToFirstRecord
                    '!TODO!: Have to code for the situation where the user is abandoning the
                    '        Add of the table's first record...e.g., go into Add mode.
                    ' Load current record's properties to form's controls, reset
                    ' navigation buttons and set "rec x of y" label
                    fnLoadControls
                    If bDebugAppTermination Then
                        Debug.Print "         Turn off Add mode in " & mstrScreenName & gcstrDOT & cstrCurrentProc
                    End If
                
                    mbInAddMode = False
                    fnSetCommandButtons True
                    ' This **must** be done as the user leaves Add mode, so that the key fields
                    ' will now be protected to prevent the user from being able to edit them.
                    ' Editing a key field is allowed only when in Add mode.
                    fnSetAvailabilityOfControls
                End If
            End If
            mbInLookupMode = False
        Else
            If bDebugAppTermination Then
                Debug.Print "      User opted NOT to discard pending changes in " & mstrScreenName & gcstrDOT & cstrCurrentProc
            End If
            
            ' User doesn't want to abandon the Add that's still in progress, so ignore the request
            ' to close the form and redisplay the form with the same data and with the user's Add
            ' still in progress.
            pintCancel = True
        End If
    Else    ' IsDirty (a.k.a. in Update mode)
        If bDebugAppTermination Then
            Debug.Print "   Update mode only (not Add): Prompt the user re: okay to discard pending changes, in " & mstrScreenName & gcstrDOT & cstrCurrentProc
        End If
    
        intButtonClicked = gerhApp.ReportNonFatal(vbObjectError + gcRES_ALRT_CHANGES_PENDING, _
                           mstrScreenName & gcstrDOT & cstrCurrentProc)
        If intButtonClicked = vbYes Then
            If bDebugAppTermination Then
                Debug.Print "      User opted to discard pending changes in " & mstrScreenName & gcstrDOT & cstrCurrentProc
            End If
                    
            If Not gbAmTryingToTerminateTheApp Then
                ' Abandon their pending changes and redisplay the same record as it *now* appears in
                ' the database
                pintCancel = True
                mtWrapper.GetRelativeRecord mtWrapper.PayeFullNm, epdSameRecord
                '!TODO!: Have to code for the situation where another user deleted the record whose
                '        edits *this* user is abandoning....e.g., go into Add mode
                fnLoadControls
                If bDebugAppTermination Then
                    Debug.Print "         Turn off Update mode in " & mstrScreenName & gcstrDOT & cstrCurrentProc
                End If
                IsDirty = False
                
                fnSetCommandButtons True
            End If
        Else
            If bDebugAppTermination Then
                Debug.Print "      User opted NOT to discard pending changes in " & mstrScreenName & gcstrDOT & cstrCurrentProc
            End If
        
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
' No resize event handler. This is a non-resizable form.
'////////////////////////////////////////////////////////////////////////////////////////////////


'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub Form_Unload(ByRef pintCancel As Integer)
    ' Comments  : Close the form
    ' Parameters: pvarLastRow
    '             pintLastCol -
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "Form_Unload"

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    If bDebugAppTermination Then
        Debug.Print "Entering " & mstrScreenName & gcstrDOT & cstrCurrentProc
    End If

    gapsApp.SaveForm Me

    IsDirty = False

    fnFreeObject mfrmMyInsuredForm
    fnFreeObject mtWrapper
    DoEvents
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


Private Sub ipdPayeClmIntRt_Change()
    ' Comments  : Sets a flag to indicate the current record has been
    '             edited, and thus Update button becomes enabled
    ' Parameters: N/A
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "ipdPayeClmIntRt_Change"
 

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    fnInitializeEditMode
    
    ' Populate msiOverride structure with data from the STATE_RULE_T
    ' row that matches the Calc State and Interest Rate
    fnGetStateInfo_Override
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
Private Sub ipmPayeSsnTinNum_Change()
    ' Comments  :
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "ipmPayeSsnTinNum_Change"

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
Private Sub lpcLookupName_Click()
    ' Comments  : Retrieve selected record
    ' Parameters: N/A
    '
    ' --------------------------------------------------
    Const cstrCurrentProc               As String = "lpcLookupName_Click"

    On Error GoTo PROC_ERR

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    fnPerformLookup lpcLookupName
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
Private Sub lpcLookupName_GotFocus()
    '-----------------------------------------------------------------------------
    ' Function     lpcLookupName_GotFocus
    ' Purpose      Turn on Lookup Mode and drop down the list now that the user has entered this control.
    ' Parameters   N/A
    ' Returns      N/A
    ' Date:        12/19/2001
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc          As String = "lpcLookupName_GotFocus"
    On Error GoTo PROC_ERR

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    'lpcLookupName.ListDown = True

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
Private Sub lpcLookupName_KeyDown(ByRef intKeyCode As Integer, ByRef intShift As Integer)
    '-----------------------------------------------------------------------------
    ' Function     lpcLookupName_KeyDown
    ' Purpose      If the user presses Enter, make it do just what the Click event does
    '              (i.e. display the selected record)
    ' Parameters   intKeyCode - ASCII code of key that was pressed
    '              intShift - indicates whether the Shift key was pressed
    ' Returns      N/A
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc          As String = "lpcLookupName_KeyDown"
    On Error GoTo PROC_ERR

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    If intKeyCode = vbKeyReturn Then
        fnPerformLookup lpcLookupName
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
Private Sub lpcLookupName_LostFocus()
    '-----------------------------------------------------------------------------
    ' Function     lpcLookupName_LostFocus
    ' Purpose      Turn off Lookup Mode now that the user has left that control.
    ' Parameters   N/A
    ' Returns      N/A
    ' Date:        12/19/2001
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc          As String = "lpcLookupName_LostFocus"
    Const clngFirstRow             As Long = 0
    On Error GoTo PROC_ERR

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    ' Display the first (blank) entry in the Lookup control so the
    ' user doesn't get confused. Without this code, the Lookup box continues to display
    ' the value last selected for lookup purposes, even when the user has since positioned
    ' to a different record by virtue of doing a Delete or Add or using the navigation buttons.
    With lpcLookupName
        .Row = clngFirstRow
        .ListIndex = clngFirstRow
        .Action = ActionClearSearchBuffer
    End With

    'fnSearchFPCombo lpcLookupName, gcstrBlankEntry, mcintDisplayCol_lpcLookupName
    lpcLookupName.Refresh

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
Private Sub iptPayeAddrLn1Txt_Change()
    ' Comments  : Limits the number of characters input to that able to
    '             be stored on the CheckFree file
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "iptPayeAddrLn1Txt_Change"

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
Private Sub iptPayeAddrLn2Txt_Change()
    ' Comments  : Limits the number of characters input to that able to
    '             be stored on the CheckFree file
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "iptPayeAddrLn2Txt_Change"

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
Private Sub iptPayeCareOfTxt_Change()
    ' Comments  : Limits the number of characters input to that able to
    '             be stored on the CheckFree file
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "iptPayeCareOfTxt_Change"

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
Private Sub iptPayeCityNmTxt_Change()
    ' Comments  : Limits the number of characters input to that able to
    '             be stored on the CheckFree file
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "iptPayeCityNmTxt_Change"

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
Private Sub iptPayeFullNm_Change()
    ' Comments  : Limits the number of characters input to that able to
    '             be stored on the CheckFree file
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "iptPayeFullNm_Change"

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
Private Sub ipcPayeDthbPmtAmt_Change()
    ' Comments  : Set a flag indicating some change has been made.
    '             The formatting of this numeric field won't be
    '             done until the LostFocus event. If done here,
    '             the user's input moves from right to left.
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "ipcPayeDthbPmtAmt_Change"

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    fnInitializeEditMode
    
    'MME START WRUS 4999 - Added the following to ensure that tier 2 (state_rule_t) entries are used if needed.
    
    If Not IsNumeric(ipcPayeDthbPmtAmt.UnFmtText) Then
       DblScreenDBPaymentValue = 0
    Else
       DblScreenDBPaymentValue = ipcPayeDthbPmtAmt.UnFmtText
    End If
    
    fnResetStateRules
    
    'MME END WRUS 4999
    
    If Not IsNumeric(ipcPayeDthbPmtAmt.UnFmtText) Then
        ' If the user cleared the contents, set the screen field to 0
        ipcPayeDthbPmtAmt.UnFmtText = 0
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
Private Sub ipdPayeWthldRt_Change()
    ' Comments  :
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "ipdPayeWthldRt_Change"

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    fnInitializeEditMode
    
    If Not IsNumeric(ipdPayeWthldRt.UnFmtText) Then
        ' If the user cleared the contents, set the screen field to 0
        ipdPayeWthldRt.UnFmtText = 0
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
Private Sub iptPayeZipCd_Change()
    ' Comments  : Limits the number of characters input to that able to
    '             be stored on the CheckFree file
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "txtPayZipCd_Change"

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

'''''''''''''''
'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub chkpaye1099ind_Change()
    ' Comments  : Limits the number of characters input to that able to
    '             be stored on the CheckFree file
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "chkpaye1099ind_Change"

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






' %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
' %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'
'   The following procedures exist only to facilitate testing. They should
'   ONLY be called from the Immediate window and not from other procedures
'   in this form or project.
'
'
'   To use these, set a breakpoint at the top of the Form_Initialize event
'   handler. Then, once you've stopped at the breakpoint, type the routine
'   name in the Immediate window.
'       Correct:   TestStub2         Incorrect:  ? TestStub2
'                                                TestStub2()
'
' %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
' %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Private Sub TestStub1()
    On Error GoTo PROC_ERR
    Dim siTemp          As StateInfo
    Dim curRate         As Currency
    
    siTemp.StrlIntRuleAmt = 0
    curRate = fnGetInterestRate(siTemp, "test")
    TestStub1Sub siTemp, curRate
    
    siTemp.StrlIntRuleAmt = "8"
    curRate = fnGetInterestRate(siTemp, "test")
    TestStub1Sub siTemp, curRate

    siTemp.StrlIntRuleAmt = "LOAN RATE - just hit Enter"
    curRate = fnGetInterestRate(siTemp, "test")
    TestStub1Sub siTemp, curRate
    
    siTemp.StrlIntRuleAmt = "Current Rate - enter a numeric rate"
    curRate = fnGetInterestRate(siTemp, "test")
    TestStub1Sub siTemp, curRate
    
    siTemp.StrlIntRuleAmt = "  > of current or 6%"
    curRate = fnGetInterestRate(siTemp, "test")
    TestStub1Sub siTemp, curRate
    
    siTemp.StrlIntRuleAmt = "Rate Condition"
    curRate = fnGetInterestRate(siTemp, "test")
    TestStub1Sub siTemp, curRate
PROC_EXIT:
    On Error Resume Next
    Exit Sub
PROC_ERR:
    Debug.Print "Error at line " & Erl
    Debug.Print "Error " & Err.Number & ": " & Err.Description
    Debug.Assert False
    Resume PROC_EXIT
End Sub



Private Sub TestStub1Sub(siIn As StateInfo, curRate As Currency)
    On Error GoTo PROC_ERR
    Dim strScope As String
   
    Debug.Print "RateIn=[" & siIn.StrlIntRuleAmt & "]    RateUsed=[" & curRate & "]"
PROC_EXIT:
    On Error Resume Next
    Exit Sub
PROC_ERR:
    Debug.Print "Error at line " & Erl
    Debug.Print "Error " & Err.Number & ": " & Err.Description
    Debug.Assert False
    Resume PROC_EXIT
End Sub

