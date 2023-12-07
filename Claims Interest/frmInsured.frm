VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmInsured 
   AutoRedraw      =   -1  'True
   Caption         =   "Insured"
   ClientHeight    =   6994
   ClientLeft      =   585
   ClientTop       =   1404
   ClientWidth     =   12350
   FillColor       =   &H80000001&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.15
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000005&
   Icon            =   "frmInsured.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   6994
   ScaleWidth      =   12350
   Begin LpLib.fpCombo lpcPycoTypCd 
      Height          =   286
      Left            =   9373
      TabIndex        =   11
      Top             =   689
      Width           =   1573
      _Version        =   196608
      _ExtentX        =   2899
      _ExtentY        =   527
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
      ColDesigner     =   "frmInsured.frx":030A
   End
   Begin LpLib.fpCombo lpcAdmnSystCd 
      Height          =   286
      Left            =   3003
      TabIndex        =   7
      Top             =   689
      Width           =   1638
      _Version        =   196608
      _ExtentX        =   3019
      _ExtentY        =   527
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
      ColDesigner     =   "frmInsured.frx":05C9
   End
   Begin LpLib.fpCombo lpcLookupSSN 
      Height          =   286
      Left            =   8905
      TabIndex        =   5
      Top             =   182
      Width           =   3367
      _Version        =   196608
      _ExtentX        =   6206
      _ExtentY        =   527
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
      ColDesigner     =   "frmInsured.frx":0888
   End
   Begin LpLib.fpCombo lpcLookupName 
      Height          =   286
      Left            =   3380
      TabIndex        =   3
      Top             =   182
      Width           =   4888
      _Version        =   196608
      _ExtentX        =   9010
      _ExtentY        =   527
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
      ColDesigner     =   "frmInsured.frx":0B47
   End
   Begin LpLib.fpCombo lpcLookupClaim 
      Height          =   286
      Left            =   715
      TabIndex        =   1
      Top             =   182
      Width           =   2002
      _Version        =   196608
      _ExtentX        =   3690
      _ExtentY        =   527
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
      ColDesigner     =   "frmInsured.frx":0E06
   End
   Begin EditLib.fpText iptClmPolNum 
      Height          =   315
      Left            =   5955
      TabIndex        =   9
      Top             =   690
      Width           =   1980
      _Version        =   196608
      _ExtentX        =   3492
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
   Begin VB.TextBox txtClmNum 
      BackColor       =   &H80000001&
      ForeColor       =   &H80000012&
      Height          =   315
      Left            =   10020
      Locked          =   -1  'True
      TabIndex        =   53
      TabStop         =   0   'False
      ToolTipText     =   "Policy Number"
      Top             =   3720
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.CommandButton cmdNavigate 
      Caption         =   "<<"
      Height          =   375
      Index           =   0
      Left            =   60
      TabIndex        =   45
      ToolTipText     =   "Go to first record"
      Top             =   4080
      Width           =   435
   End
   Begin VB.CommandButton cmdNavigate 
      Caption         =   "<"
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   46
      ToolTipText     =   "Go to previous record"
      Top             =   4080
      Width           =   435
   End
   Begin VB.CommandButton cmdNavigate 
      Caption         =   ">"
      Height          =   375
      Index           =   2
      Left            =   900
      TabIndex        =   47
      ToolTipText     =   "Go to next record"
      Top             =   4080
      Width           =   435
   End
   Begin VB.CommandButton cmdNavigate 
      Caption         =   ">>"
      Height          =   375
      Index           =   3
      Left            =   1320
      TabIndex        =   48
      ToolTipText     =   "Go to last record"
      Top             =   4080
      Width           =   435
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   3675
      TabIndex        =   40
      ToolTipText     =   "Add a new Insured"
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   7455
      TabIndex        =   43
      ToolTipText     =   "Cancel your changes or close this screen"
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   6195
      TabIndex        =   42
      ToolTipText     =   "Delete this Insured"
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   375
      Left            =   4935
      TabIndex        =   41
      ToolTipText     =   "Save your changes"
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Frame fraPayees 
      Caption         =   "Payees"
      Height          =   1980
      Left            =   75
      TabIndex        =   49
      Top             =   4485
      Width           =   12135
      Begin MSFlexGridLib.MSFlexGrid msgPayees 
         Height          =   1515
         Left            =   90
         TabIndex        =   51
         ToolTipText     =   "Payees on this claim"
         Top             =   450
         Width           =   11925
         _ExtentX        =   21039
         _ExtentY        =   2684
         _Version        =   393216
         Cols            =   14
         BackColorSel    =   -2147483637
         BackColorBkg    =   -2147483633
         ScrollTrack     =   -1  'True
         HighLight       =   2
         AllowUserResizing=   1
         FormatString    =   $"frmInsured.frx":10C5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.1509
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblGridInstructions 
         Alignment       =   1  'Right Justify
         Caption         =   "Double-click on an existing Payee to edit that Payee."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7320
         TabIndex        =   50
         ToolTipText     =   "Double-click on an existing Payee to edit that Payee."
         Top             =   180
         Width           =   4575
      End
   End
   Begin VB.Frame fraClaim 
      Caption         =   "Claim Information Across All Payees"
      Height          =   2535
      Left            =   7020
      TabIndex        =   28
      Top             =   1170
      Width           =   5175
      Begin EditLib.fpCurrency ipcClmTotDthbPmtAmt 
         Height          =   315
         Left            =   2385
         TabIndex        =   30
         Top             =   375
         Width           =   2535
         _Version        =   196608
         _ExtentX        =   4471
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
      Begin EditLib.fpCurrency ipcClmTotIntAmt 
         Height          =   315
         Left            =   2385
         TabIndex        =   32
         Top             =   720
         Width           =   2535
         _Version        =   196608
         _ExtentX        =   4471
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
      Begin EditLib.fpCurrency ipcClmTotWthldAmt 
         Height          =   315
         Left            =   2385
         TabIndex        =   35
         Top             =   1095
         Width           =   2535
         _Version        =   196608
         _ExtentX        =   4471
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
      Begin EditLib.fpCurrency ipcClmTotClmPdAmt 
         Height          =   315
         Left            =   2385
         TabIndex        =   38
         Top             =   1605
         Width           =   2535
         _Version        =   196608
         _ExtentX        =   4471
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
      Begin VB.Label lblMinus 
         Alignment       =   1  'Right Justify
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
         Height          =   195
         Left            =   2160
         TabIndex        =   36
         Top             =   1155
         Width           =   165
      End
      Begin VB.Label lblPlus 
         Alignment       =   1  'Right Justify
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
         Height          =   195
         Left            =   2175
         TabIndex        =   33
         Top             =   780
         Width           =   165
      End
      Begin VB.Line linTotals 
         BorderWidth     =   2
         X1              =   240
         X2              =   4920
         Y1              =   1500
         Y2              =   1500
      End
      Begin VB.Label lblClmTotDthbPmtAmt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total DB Payment:"
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   435
         Width           =   1335
      End
      Begin VB.Label lblClmTotWthldAmt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Interest Withheld:"
         Height          =   195
         Left            =   240
         TabIndex        =   34
         Top             =   1155
         Width           =   1725
      End
      Begin VB.Label lblClmTotClmPdAmt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         Height          =   195
         Left            =   1740
         TabIndex        =   37
         Top             =   1665
         Width           =   420
      End
      Begin VB.Label lbClmTotIntAmt 
         AutoSize        =   -1  'True
         Caption         =   "Total Claim Interest:"
         Height          =   195
         Left            =   240
         TabIndex        =   31
         Top             =   780
         Width           =   1470
      End
   End
   Begin VB.Frame fraInsured 
      Caption         =   "Insured"
      Height          =   2535
      Left            =   468
      TabIndex        =   12
      Top             =   1170
      Width           =   5880
      Begin VB.CheckBox chkClmCmpCalInd 
         Caption         =   "Compact Filling"
         Height          =   247
         Left            =   3393
         TabIndex        =   54
         Top             =   1170
         Width           =   1651
      End
      Begin VB.ComboBox cboInsdDthResStCd 
         Height          =   286
         Left            =   4665
         Style           =   2  'Dropdown List
         TabIndex        =   21
         ToolTipText     =   "The state in which the Insured resided at time of death"
         Top             =   1440
         Width           =   690
      End
      Begin VB.CheckBox chkClmForResDthInd 
         Caption         =   "Forei&gn Residence at Death?"
         Height          =   315
         Left            =   585
         TabIndex        =   19
         Top             =   1426
         Width           =   2535
      End
      Begin VB.ComboBox cboIssStCd 
         Height          =   286
         Left            =   1905
         Style           =   2  'Dropdown List
         TabIndex        =   18
         ToolTipText     =   "The state in which the Insured resided at time of issue"
         Top             =   1150
         Width           =   690
      End
      Begin MSComCtl2.DTPicker dtpClmProofDt 
         Height          =   312
         Left            =   1911
         TabIndex        =   25
         ToolTipText     =   "The date on which the death certificate was received in the Home Office"
         Top             =   2106
         Width           =   1326
         _ExtentX        =   2348
         _ExtentY        =   551
         _Version        =   393216
         Format          =   129040385
         CurrentDate     =   37013
         MinDate         =   21916
      End
      Begin MSComCtl2.DTPicker dtpClmInsdDthDt 
         Height          =   315
         Left            =   1905
         TabIndex        =   23
         ToolTipText     =   "The date on which the Insured died"
         Top             =   1755
         Width           =   1335
         _ExtentX        =   2348
         _ExtentY        =   551
         _Version        =   393216
         Format          =   129040385
         CurrentDate     =   37013
         MinDate         =   21916
      End
      Begin EditLib.fpMask ipmClmInsdSsnNum 
         Height          =   315
         Left            =   3900
         TabIndex        =   27
         Top             =   2100
         Width           =   1740
         _Version        =   196608
         _ExtentX        =   3069
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
      Begin EditLib.fpText iptClmInsdFirstNm 
         Height          =   315
         Left            =   1080
         TabIndex        =   14
         Top             =   375
         Width           =   4590
         _Version        =   196608
         _ExtentX        =   8096
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
      Begin EditLib.fpText iptClmInsdLastNm 
         Height          =   315
         Left            =   1080
         TabIndex        =   16
         Top             =   720
         Width           =   4590
         _Version        =   196608
         _ExtentX        =   8096
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
      Begin VB.Label lblClmInsdDthDt 
         AutoSize        =   -1  'True
         Caption         =   "Dat&e Of Death:"
         Height          =   195
         Left            =   585
         TabIndex        =   22
         Top             =   1815
         Width           =   1110
      End
      Begin VB.Label lblInsdDthResStCd 
         AutoSize        =   -1  'True
         Caption         =   "Reside&nce State:"
         Height          =   195
         Left            =   3393
         TabIndex        =   20
         Top             =   1486
         Width           =   1235
      End
      Begin VB.Label lblClmInsdSsnNum 
         Caption         =   "&SSN:"
         Height          =   195
         Left            =   3393
         TabIndex        =   26
         Top             =   2160
         Width           =   377
      End
      Begin VB.Label lblIssStCd 
         AutoSize        =   -1  'True
         Caption         =   "&Issue State:"
         Height          =   195
         Left            =   585
         TabIndex        =   17
         Top             =   1196
         Width           =   885
      End
      Begin VB.Label lblClmInsdFirstNm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "First Na&me:"
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   435
         Width           =   825
      End
      Begin VB.Label lblClmProofDt 
         AutoSize        =   -1  'True
         Caption         =   "Date of &Proof:"
         Height          =   195
         Left            =   585
         TabIndex        =   24
         Top             =   2160
         Width           =   1035
      End
      Begin VB.Label lblClmInsdLastNm 
         Caption         =   "&Last Name:"
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   780
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdAddPayee 
      Caption         =   "Add Payee..."
      Height          =   375
      Left            =   5565
      TabIndex        =   52
      ToolTipText     =   "Add a new Payee"
      Top             =   6560
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrintReport 
      Caption         =   "Print Rep&ort"
      Height          =   375
      Left            =   10980
      TabIndex        =   44
      ToolTipText     =   "Print a claim report for this Insured"
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label lblPycoTypCd 
      BackStyle       =   0  'Transparent
      Caption         =   "Company &Type:"
      Height          =   195
      Left            =   8175
      TabIndex        =   10
      Top             =   750
      Width           =   1155
   End
   Begin VB.Label lblAdmnSystCd 
      BackStyle       =   0  'Transparent
      Caption         =   "Admin S&ystem:"
      Height          =   195
      Left            =   1935
      TabIndex        =   6
      Top             =   750
      Width           =   1095
   End
   Begin VB.Label lblLookupSSN 
      BackStyle       =   0  'Transparent
      Caption         =   "SSN Lookup:"
      ForeColor       =   &H80000013&
      Height          =   435
      Left            =   8340
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblLookupName 
      BackStyle       =   0  'Transparent
      Caption         =   "Name Lookup:"
      ForeColor       =   &H80000013&
      Height          =   435
      Left            =   2775
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblLookupClaim 
      BackStyle       =   0  'Transparent
      Caption         =   "Claim Lookup:"
      ForeColor       =   &H80000013&
      Height          =   435
      Left            =   135
      TabIndex        =   0
      Top             =   120
      Width           =   600
   End
   Begin VB.Label lblRecordPosition 
      BackStyle       =   0  'Transparent
      Caption         =   "Record x of y"
      Height          =   195
      Left            =   75
      TabIndex        =   39
      Top             =   3765
      Width           =   2625
   End
   Begin VB.Shape shpLookup 
      BackColor       =   &H80000003&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   510
      Left            =   75
      Top             =   75
      Width           =   12255
   End
   Begin VB.Label lblClmPolNum 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Policy Num&ber:"
      Height          =   195
      Left            =   4861
      TabIndex        =   8
      Top             =   750
      Width           =   1065
   End
End
Attribute VB_Name = "frmInsured"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
' Module     : frmInsured
' Description:
' Procedures:
'              ClmCmpCalInd_Click()
'              cboInsdDthResStCd_Click()
'              cboIssStCd_Click()
'              chkClmForResDthInd_Click()
'              cmdAdd_Click()
'              cmdAddPayee_Click()
'              cmdClose_Click()
'              cmdDelete_Click()
'              cmdNavigate_Click(ByRef pintIndex As Integer)
'              cmdPrintReport_Click()
'              cmdUpdate_Click()
'              dtpClmInsdDthDt_Change()
'              dtpClmProofDt_Change()
'              fnAddRecord()
'              fnBindControlsToTableWrapper()
'              fnCalcTotalsForAllPayees()
'              fnCalcTotalsForAllPayees(ByVal lngClmID As Long) As ADODB.Recordset
'              fnClearControls()
'              fnFillPayeeGrid()
'              fnGetChildren()
'              fnGetData_IndividualReport() As ADODB.Recordset
'              fnGetFieldLabel(ByVal strControlName As String) As String
'              fnGetLobCd() As String
'              fnGetPayeesNeedingRecalcDueToDeath(lngClmIdIn As Long, dteClmInsdDthDtIn As Date) As Long
'              fnGetPayeesNeedingRecalcDueToProof(lngClmIdIn As Long, dteClmProofDtIn As Date) As Long
'              fnGetReportFile() As String
'              fnInitializeEditMode()
'              fnGetDefaultPayorCompany(ByVal strClmPolNum As String, ByVal strAdmnSystCd As String) As String
'              fnLoadCboInsdDthResStCd()
'              fnLoadCboIssStCd()
'              fnLoadControls()
'              fnLoadLpcAdmnSystCd()
'              fnLoadLpcLookup(ByRef lpcIn As LPLib.fpCombo, ByVal lngLookupType As EnumLookupType)
'              fnLoadLpcPycoTypCd()
'              fnLoadRecordWithCalculatedControls()
'              fnPerformLookup(ByRef lpcIn As LPLib.fpCombo)
'              fnRefreshAllCombos()
'              fnSetAvailabilityOfControls(Optional ByVal bChangeFocus = True)
'              fnSetCommandButtons(ByVal bEnable As Boolean)
'              fnSetDefaultControlProperties()
'              fnSetFocusToFirstUpdateableField()
'              fnSetInsdDthResStCdAvailability()
'              fnSetNavigationButtons(Optional ByVal bUnconditionalDisable As Boolean = False)
'              fnSetPropertiesForPayeeScreen(bSendEmptyName As Boolean)
'              fnSetTxtClmNum()
'              fnSetupScreenControls()
'              fnValidData() As Boolean
'              fnWarningData()
'              Form_Activate()
'              Form_Load()
'              Form_QueryUnload(ByRef pintCancel As Integer, ByRef pintUnloadMode As Integer)
'              Form_Resize()
'              Form_Unload(ByRef pintCancel As Integer)
'              ipmClmInsdSsnNum_Change()
'              iptClmInsdFirstNm_Change()
'              iptClmInsdLastNm_Change()
'              iptClmPolNum_Change()
'              lpcAdmnSystCd_Change()
'              lpcAdmnSystCd_GotFocus()
'              lpcLookupClaim_Click()
'              lpcLookupClaim_GotFocus()
'              lpcLookupClaim_KeyDown(ByRef intKeyCode As Integer, ByRef intShift As Integer)
'              lpcLookupClaim_LostFocus()
'              lpcLookupName_Click()
'              lpcLookupName_GotFocus()
'              lpcLookupName_KeyDown(ByRef intKeyCode As Integer, ByRef intShift As Integer)
'              lpcLookupName_LostFocus()
'              lpcLookupSSN_Click()
'              lpcLookupSSN_GotFocus()
'              lpcLookupSSN_KeyDown(ByRef intKeyCode As Integer, ByRef intShift As Integer)
'              lpcLookupSSN_LostFocus()
'              lpcPycoTypCd_Change()
'              lpcPycoTypCd_GotFocus()
'              msgPayees_DblClick()
'              ClmCmpCalInd_Click()
'               fnSetCompactFillingCheckBox()
'
' Modified   :
' 03/26/01 BAW Cleaned with Total Visual CodeTools 2000
' 01/2002  BAW Removed "#If gcfLOOKUP" stuff since we definitely want Lookup capability. (At one
'              time before v2.2 was released, we thought the performance might be too bad to keep it.)
'              Also optimized per Project Analyzer (removing dead code, adding "$" to Mid/Space, etc.).
'              Also updated the cboLoadCboInsdDthResStCd and fnLoadCboLookupClaim to improve performance.
' Modified:     Berry Kropiwka - Added Compact Calc - 2019-09-27
' --------------------------------------------------
Option Explicit
Option Compare Binary

Private mstrScreenName As String

Private Const mclngMinFormWidth                 As Long = 12465
Private Const mclngMinFormHeight                As Long = 7500
' The following constants identify, for fpCombo controls used as multi-column comboboxes,
' which column is displayed in the Edit portion of the control (index = mcintDisplayCol_xxxx)
' and which is saved to a column in its corresponding SQL table (index = mcintStoreCol_xxxx),
' where xxxx is the fpCombo control's name.
Private Const mcintDisplayCol_lpcAdmnSystCd     As Integer = 0
Private Const mcintStoreCol_lpcAdmnSystCd       As Integer = 1
Private Const mcintDisplayCol_lpcPycoTypCd      As Integer = 0
Private Const mcintStoreCol_lpcPycoTypCd        As Integer = 1

Private Const mcstrGroupLOB                     As String = "G"
Private Const mcstrIndividualLOB                As String = "I"

Private Const mcstrPyco_Subsidiary              As String = "Subsidiary"
Private Const mcstrPyco_Parent                  As String = "Parent"

  ' Leverage/Claimbuilder Project - K723 - 07/05/2014
Private Const mcstrPyco_SLHIC                   As String = "SLHIC"
Private Const mcstrAdmnSystSOLAR                As String = "24"

' The following constants identify, for fpCombo controls used as Lookups,
' which column is displayed in the Edit portion of the control (index = mcintDisplayCol_xxxx,
' where xxxx is the fpCombo control's name).
Private Const mcintDisplayCol_lpcLookupClaim    As Integer = 0
Private Const mcintDisplayCol_lpcLookupSSN      As Integer = 0
Private Const mcintDisplayCol_lpcLookupName     As Integer = 0

' These constants define the columns within the Lookup/Multi-column combo boxes.
' These are used to give a name to a given column of the fpCombo control so
' it can be referenced by name, not by number.
Private Const mcstrDisplayCol                   As String = "DISPLAY_COL"
Private Const mcstrClmId                        As String = "CLM_ID"
Private Const mcstrClmInsdDthDt                 As String = "CLM_INSD_DTH_DT"
Private Const mcstrClmInsdFirstNm               As String = "CLM_INSD_FIRST_NM"
Private Const mcstrClmInsdLastNm                As String = "CLM_INSD_LAST_NM"
Private Const mcstrClmInsdSsnNum                As String = "CLM_INSD_SSN_NUM"
Private Const mcstrClmNum                       As String = "CLM_NUM"
Private Const mcstrAdmnSystDsc                  As String = "ADMN_SYST_DSC"
Private Const mcstrAdmnSystCd                   As String = "ADMN_SYST_CD"
Private Const mcstrPycoTypDsc                   As String = "PYCO_TYP_DSC"
Private Const mcstrPycoTypCd                    As String = "PYCO_TYP_CD"

'-----------------------------------------------------------------------
' The following Enum is used by fnLoadLpcLookup and denotes which
' lookup is being populated.
'-----------------------------------------------------------------------
Public Enum EnumLookupType
    elt_Claim = 0
    elt_Name = 1
    elt_SSN = 2
End Enum


' mtWrapper is an instance of the table wrapper corresponding to the main table maintained by this form.
Private mtWrapper               As ctclmClaim
' mtPayee is an instance of the table wrapper corresponding to the table that is SUBORDINATE to the
' main table maintained by this form
Private mtPayee                 As ctpyePayee


' Define a constant for each field that may get an error or warning. This
' should match the text of that control's associated Label control.
Private Const mcstrLpcAdmnSystCdLabel As String = "Admin System"
Private Const mcstrIptClmPolNumLabel As String = "Policy Number"
Private Const mcstrLpcPycoTypCdLabel As String = "Company Type"
Private Const mcstrIptClmInsdFirstNmLabel As String = "First Name"
Private Const mcstrIptClmInsdLastNmLabel As String = "Last Name"
Private Const mcstrCboIssStCdLabel As String = "Issue State"
Private Const mcstrCboInsdDthResStCdLabel As String = "Residence State"
Private Const mcstrChkClmCmpCalIndLabel As String = "Compact Filling"
Private Const mcstrChkClmForResDthIndLabel As String = "Foreign Residence at Death"
Private Const mcstrDtpClmInsdDthDtLabel As String = "Date of Death"
Private Const mcstrDtpClmProofDtLabel As String = "Date of Proof"
Private Const mcstrIpmClmInsdSsnNumLabel As String = "SSN"
Private Const mcstrIpcClmTotDthbPmtAmtLabel As String = "Total DB Payment"
Private Const mcstrIpcClmTotIntAmtLabel As String = "Total Claim Interest"
Private Const mcstrIpcClmTotWthldAmtLabel As String = "Total Interest Withheld"
Private Const mcstrIpcClmTotClmPdAmtLabel As String = "Total"

Private Const mcstrTxtClmNumLabel As String = "Claim Number"
Private Const mcstrTxtClmIDLabel As String = "Claim ID"


Dim mrstPayees  As ADODB.Recordset

' mbInLookupMode determines whether the user is in the process of doing a search using the Lookup ComboBox
Private mbInLookupMode          As Boolean

' mbInAddMode determines whether the user has begun the process of adding a new record to the table.
' Note that Add mode is independent of Update mode
Private mbInAddMode             As Boolean

Private mctlFirstUpdateableField_Add As Control
Private mctlFirstUpdateableField_Upd As Control

Private mstrOrigDateOfDeath As String
Private mstrOrigDateOfProof As String

Private mintAdmnSyst_MinPolNumLength As Integer
Private mintAdmnSyst_MaxPolNumLength As Integer
Private mstrAdmnSyst_DfltPycoTypDsc  As String
Private mstrAdmnSyst_TaxRptgInd      As String


'------------------------------------------
'            MEMBER VARIABLES
'
' These are used by the Payee screen.
'------------------------------------------
' member variable for InsuredClmForResDthInd property
Private m_bInsuredClmForResDthInd As Boolean
' member variable for InsuredClmID property
Private m_lngInsuredClmID As Long
' member variable for InsuredClmNum property
Private m_strInsuredClmNum As String
' member variable for InsuredCurrentPayeeName property
Private m_strInsuredCurrentPayeeName As String
' member variable for InsuredCurrentPayeeID property
Private m_lngInsuredCurrentPayeeID As Long
' member variable for InsuredClmInsdDthDt property
Private m_dteInsuredClmInsdDthDt As Date
' member variable for InsuredClmProofDt property
Private m_dteInsuredClmProofDt As Date
' member variable for InsuredLobCd property
Private m_strInsuredLobCd As String
' member variable for InsuredInsdDthResStCd property
Private m_strInsuredInsdDthResStCd As String
' member variable for InsuredIssStCd property
Private m_strInsuredIssStCd As String

' m_bIsDirty corresponds to the public property called IsDirty.
' All maintenance screens should have this field and that property! When True, it indicates
' that the user has made --but not yet saved-- changes to a record. The MDI form will query
' this property if the user opens the File menu, since the Exit option should be disabled if
' any form has outstanding changes.
' Be sure to use this variable's corresponding Property Let to change its value.
' Do **NOT** set m_bIsDirty itself, as this will using the Property Let proc will
' ensure the Close button caption is always synchronized with the value of the property.
Private m_bIsDirty              As Boolean

'Private Const variable for the compact filling state code. Used when filter the Issued State and Residence State.
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
    Const cstrCurrentProc   As String = "Let IsDirty"
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


'////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Get InsuredClmForResDthInd() As Boolean
    Const cstrCurrentProc As String = "Property Get InsuredClmForResDthInd"
    On Error GoTo PROC_ERR

    InsuredClmForResDthInd = m_bInsuredClmForResDthInd
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
Private Property Let InsuredClmForResDthInd(ByVal bValue As Boolean)
    Const cstrCurrentProc As String = "Property Let InsuredClmForResDthInd"
    On Error GoTo PROC_ERR

    m_bInsuredClmForResDthInd = bValue
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
Public Property Get InsuredClmID() As Long
    Const cstrCurrentProc As String = "Property Get InsuredClmID"
    On Error GoTo PROC_ERR

    InsuredClmID = m_lngInsuredClmID
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
Private Property Let InsuredClmID(ByVal lngValue As Long)
    Const cstrCurrentProc As String = "Property Let InsuredClmID"
    On Error GoTo PROC_ERR

    m_lngInsuredClmID = lngValue
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

'Y027 07-11-2012
'////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Get AdminSystemCode() As String
    Const cstrCurrentProc As String = "Property Get InsuredClmNum"
    On Error GoTo PROC_ERR

    AdminSystemCode = lpcAdmnSystCd.Text
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
Public Property Get InsuredClmNum() As String
    Const cstrCurrentProc As String = "Property Get InsuredClmNum"
    On Error GoTo PROC_ERR

    InsuredClmNum = m_strInsuredClmNum
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
Private Property Let InsuredClmNum(ByVal strValue As String)
    Const cstrCurrentProc As String = "Property Let InsuredClmNum"
    On Error GoTo PROC_ERR

    m_strInsuredClmNum = strValue
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
Public Property Get InsuredClmProofDt() As Date
    Const cstrCurrentProc As String = "Property Get InsuredClmProofDt"
    On Error GoTo PROC_ERR

    InsuredClmProofDt = m_dteInsuredClmProofDt
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
Private Property Let InsuredClmProofDt(ByVal dteValue As Date)
    Const cstrCurrentProc As String = "Property Let InsuredClmProofDt"
    On Error GoTo PROC_ERR

    m_dteInsuredClmProofDt = dteValue
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
Public Property Get InsuredCurrentPayeeID() As Long
    Const cstrCurrentProc As String = "Property Get InsuredCurrentPayeeID"
    On Error GoTo PROC_ERR

    InsuredCurrentPayeeID = m_lngInsuredCurrentPayeeID
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
Private Property Let InsuredCurrentPayeeID(ByVal lngValue As Long)
    Const cstrCurrentProc As String = "Property Let InsuredCurrentPayeeID"
    On Error GoTo PROC_ERR

    m_lngInsuredCurrentPayeeID = lngValue
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
Public Property Get InsuredCurrentPayeeName() As String
    Const cstrCurrentProc As String = "Property Get InsuredCurrentPayeeName"
    On Error GoTo PROC_ERR

    InsuredCurrentPayeeName = m_strInsuredCurrentPayeeName
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
Private Property Let InsuredClmInsdDthDt(ByVal dteValue As Date)
    Const cstrCurrentProc As String = "Property Let InsuredClmInsdDthDt"
    On Error GoTo PROC_ERR

    m_dteInsuredClmInsdDthDt = dteValue
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
Public Property Get InsuredClmInsdDthDt() As Date
    Const cstrCurrentProc As String = "Property Get InsuredClmInsdDthDt"
    On Error GoTo PROC_ERR

    InsuredClmInsdDthDt = m_dteInsuredClmInsdDthDt
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
Private Property Let InsuredCurrentPayeeName(ByVal strValue As String)
    Const cstrCurrentProc As String = "Property Let InsuredCurrentPayeeName"
    On Error GoTo PROC_ERR

    m_strInsuredCurrentPayeeName = strValue
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
Public Property Get InsuredInsdDthResStCd() As String
    Const cstrCurrentProc As String = "Property Get InsuredInsdDthResStCd"
    On Error GoTo PROC_ERR

    InsuredInsdDthResStCd = m_strInsuredInsdDthResStCd
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
Private Property Let InsuredInsdDthResStCd(ByVal strValue As String)
    Const cstrCurrentProc As String = "Property Let InsuredInsdDthResStCd"
    On Error GoTo PROC_ERR

    m_strInsuredInsdDthResStCd = strValue
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
Public Property Get InsuredIssStCd() As String
    Const cstrCurrentProc As String = "Property Get InsuredIssStCd"
    On Error GoTo PROC_ERR

    InsuredIssStCd = m_strInsuredIssStCd
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
Private Property Let InsuredIssStCd(ByVal strValue As String)
    Const cstrCurrentProc As String = "Property Let InsuredIssStCd"
    On Error GoTo PROC_ERR

    m_strInsuredIssStCd = strValue
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
Public Property Get InsuredLobCd() As String
    Const cstrCurrentProc As String = "Property Get InsuredLobCd"
    On Error GoTo PROC_ERR

    InsuredLobCd = m_strInsuredLobCd
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
Private Property Let InsuredLobCd(ByVal strValue As String)
    Const cstrCurrentProc As String = "Property Let InsuredLobCd"
    On Error GoTo PROC_ERR

    m_strInsuredLobCd = strValue
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
'|                      PRIVATE    Procedures                       |
'/                                                                  \
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/


'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub cboInsdDthResStCd_Click()
    ' Comments  : Sets a flag to indicate the current record has been
    '             edited, and thus Update button becomes enabled
    ' Parameters: N/A
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "cboInsdDthResStCd_Click"

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
Private Sub cboIssStCd_Click()
    ' Comments  : Sets a flag to indicate the current record has been
    '             edited, and thus Update button becomes enabled
    ' Parameters: N/A
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "cboIssStCd_Click"
 

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    fnInitializeEditMode
    
    ' Enable or Disable the Compact Filling check box based on Admin System
    fnSetCompactFillingCheckBox Me.lpcAdmnSystCd.Text, Me.cboIssStCd.Text
    
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
Private Sub chkClmCmpCalInd_Click()
    ' Comments  : Sets a flag to indicate the current record has been
    '             edited, and thus Update button becomes enabled
    ' Parameters: N/A
    ' Modified:     Berry Kropiwka - Added Compact Calc - 2019-09-27
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "chkClmCmpCalInd_Click"
 
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
Private Sub chkClmForResDthInd_Click()
    ' Comments  : Sets a flag to indicate the current record has been
    '             edited, and thus Update button becomes enabled
    ' Parameters: N/A
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "chkClmForResDthInd_Click"
 
    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    fnInitializeEditMode
    
    ' Disable the Residence State combobox if this checkbox is selected; otherwise enable it.
    fnSetInsdDthResStCdAvailability
    
    ' Blank out the Residence State's selection if that control is now disabled.
    If chkClmForResDthInd.value = vbUnchecked Then
        cboInsdDthResStCd.Text = gcstrBlankEntry
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
Private Sub cmdAddPayee_Click()
    ' Comments  : Opens the Payee maintenance screen
    ' Parameters:  -
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc   As String = "cmdAddPayee_Click"
    Dim frmChild            As Form
    Dim strSaveClaimNumber  As String
    Dim hrgHourglass        As chrgHourglass
    Dim lngReturnValue      As Long
    Dim strACF2             As String
 
    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName


    If mtWrapper.LookupRecordCount > 0 Then
        fnSetPropertiesForPayeeScreen bSendEmptyName:=True
        ' Following statement triggers the Form_Initialize & Form_Load events in frmPayee
        Set frmChild = New frmPayee
        ' Following statement triggers the Form_Activate event in frmPayee
        frmChild.Show vbModal

        Set hrgHourglass = New chrgHourglass
        hrgHourglass.value = True

        ' Note: You *must* requery the Insured and Payee recordsets to accomodate the possibility
        '       that another user (a) add/changed/deleted one more Payees for the
        '       current Insured and (b) returned to the Insured screen which triggered an update
        '       to the Insured record for the claim-wide totals it carries. If you don't do the
        '       requeries then a -2147217864 "row cannot be located for updating..." error could
        '       occur. So, we'll do the requerying automatically with no visible indication to the
        '       user that it occured unless the requerying revealed that another user deleted the
        '       current claim number and hence the Insured with the next higher claim number will
        '       be displayed (otherwise the same claim remains being displayed).

        strSaveClaimNumber = iptClmPolNum.Text

        ' Do an immediate repaint. This allows the Insured screen to be redrawn BEFORE all
        ' the work of requerying and repainting is started. When the requerying/repainting is done,
        ' only small parts of the screen (not the whole screen) will need to be repainted. This
        ' eliminates the user seeing a very slow repainting.
        Me.Refresh

        '!TODO! The following looks like unnecessary (i.e. dead) code
        'If txtClmNum <> strSaveClaimNumber Then
            'MsgBox "Another user has deleted the Claim Number (" & strSaveClaimNumber & ") you were viewing.", _
            '       vbOKOnly + vbInformation, mcstrDialogTitle
        'End If
        
        hrgHourglass.value = True
        
        fnGetChildren
        
        ' 01/31/2001 BAW - Add another Refresh to speed up repainting
        Me.Refresh
        
        ' Totals may have changed. Update the Insured record just in case.
        fnLoadRecordWithCalculatedControls
        
        ' 01/31/2001 BAW - Add another Refresh to speed up repainting
        Me.Refresh
        
        ' 01/31/2001 BAW - Add another Refresh to speed up repainting
        Me.Refresh
        
        With mtWrapper
            ' Determine whether another user updated or deleted the record about to be updated.
            ' Note: this multi-user checking is performed on an Update but not an Add.
            lngReturnValue = .CheckForAnotherUsersChanges(ewoUpdate, strACF2)

            If lngReturnValue = vbObjectError + gcRES_INFO_ANOTHER_USER_DELETED_NOT_SAVED Then
                gerhApp.ReportNonFatal vbObjectError + gcRES_INFO_ANOTHER_USER_DELETED_NOT_SAVED, _
                                           mstrScreenName & gcstrDOT & cstrCurrentProc
                ' Discard *this* user's pending changes and show the previous record.
                ' Can't use the GetClmNumFromClmId( ) method since the CLAIM_T row
                ' doesn't exist and hence a "-2147217900" (Claim ID does not exist) error
                ' throws things off.
                .GetRelativeRecord .ClmNum, epdPreviousRecord
            ' Do NOT bother to check for another UPDATING the record, since all we're doing is
            ' updating the total fields. Let the totals update go through.
            '   ElseIf lngReturnValue = vbObjectError + gcRES_INFO_ANOTHER_USER_UPDATED_DISCARDED Then
            '       gerhApp.ReportNonFatal vbObjectError + gcRES_INFO_ANOTHER_USER_UPDATED_DISCARDED, _
            '                           mstrScreenName & gcstrDOT & cstrCurrentProc, _
            '                           Trim$(strACF2)
            '       ' Discard *this* user's pending changes by re-retrieving the current record
            '       ' as it currently looks on the database and refreshing the lookup recordset.
            '       ' Can't use the GetClmNumFromClmId( ) method since the CLAIM_T row
            '       ' doesn't exist and hence a "-2147217900" (Claim ID does not exist) error
            '       ' throws things off.
            '       .GetRelativeRecord .ClmNum, epdSameRecord
            Else
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
            ' they reflects this and other users' changes. Then call fnLoadControls
            ' to make sure comboboxes' selection is reset as appropriate
            fnRefreshAllCombos
        End With
        
        ' Have to call fnLoadControls here, like in cmdAdd_Click and cmdDelete_Click and cmdUpdate_Click,
        ' to ensure refreshed comboboxes have their previous value still selected.
        If mtWrapper.LookupRecordCount > 0 Then
            ' Ensure the on-screen controls reflect the record just added/updated, in case the
            ' DBMS altered it in some way, e.g., determining an Identity column value and
            ' getting the most up-to-date Last Updated info. This also sets the navigation
            ' buttons and updates the "record x of y" label
            fnLoadControls
            fnSetCommandButtons True
        Else
            fnAddRecord
        End If
    Else
        ' 2003 = There is no current Insured record. The Payee screen cannot be opened.
        gerhApp.ReportNonFatal vbObjectError + gcRES_WARN_NO_CURR_INSURED, _
                               mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0

    ' Clean-up statements go here
    If Not (hrgHourglass Is Nothing) Then
        hrgHourglass.value = False
    End If
    fnFreeObject hrgHourglass
    ' Terminate the Payee form, removing it from the Forms collection
    fnFreeObject frmChild

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
    ' Purpose     Will close the screen
    '
    '             NOTE: The logic in this function should closely resemble that
    '                   in the Form_QueryUnload event handler!
    ' Parameters: N/A
    ' Returns:    N/A
    ' Modified:
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
    Const cstrCurrentProc            As String = "cmdDelete_Click"
    Dim intButtonClicked             As Integer
    Dim lngReturnValue               As Long
    Dim strACF2                      As String
    Dim hrgHourglass                 As chrgHourglass
    On Error GoTo PROC_ERR
    
    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    ' .......................................................................
    ' Make sure the user really, really, really wants to delete this record.
    ' .......................................................................
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
            ' Can't use the GetClmNumFromClmId( ) method since the CLAIM_T row
            ' doesn't exist and hence a "-2147217900" (Claim ID does not exist) error
            ' throws things off.
            .GetRelativeRecord .ClmNum, epdPreviousRecord
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
Private Sub cmdPrintReport_Click()
    ' Comments  :
    ' Parameters:  -
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc   As String = "cmdPrintReport_Click"
    Dim crDB                As CRAXDRT.Database
    Dim hrgHourglass        As chrgHourglass
    Dim rstReportData       As New ADODB.Recordset
 
    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    Set hrgHourglass = New chrgHourglass
    hrgHourglass.value = True

    ' Moved the instantiation of the Crystal object here (from modStartup) as a conditional instantiation
    ' since this CreateObject invocation is such a pig per VB Watch Profiler.
    If (gcrxApp Is Nothing) Then
        Set gcrxApp = CreateObject("CrystalRuntime.Application")
    End If

    Set gcReportToPrint = gcrxApp.OpenReport(fnGetReportFile())
    Set crDB = gcReportToPrint.Database

    ' Build an ADODB.Recordset containing the info to appear on the report
    Set rstReportData = fnGetData_IndividualReport()

    ' Tell the report the where its data is coming from, e.g., the
    ' ADODB.Recordset just created
    gcReportToPrint.Database.SetDataSource Data:=rstReportData, dataTag:=3, tableNumber:=1

    ' ...............................................................................
    ' Set formula field(s) in the report that supply additional info that
    ' is not in the recordset (typically singularly-occuring data)
    ' ...............................................................................
    fnSetFormulaField "formulaReportName", "Individual Report"
    fnSetFormulaField "formulaReportPeriodDescript", vbNullString   ' No criteria for this report

    ' ...............................................................................
    ' Tell the report where the data is coming from (overriding whatever might
    ' have been set at design-time). All of the following is necessary since
    ' the location and Connect string set within the .RPT itself may not be
    ' accurate in a production environment (or even on another developer's PC)
    ' ...............................................................................
    With crDB
        .SetDataSource rstReportData
    End With
    'With crDB.Tables.Item(1)
    '    .SetLogOnInfo pServerName:=strDBPath, pDatabaseName:=vbNullString, _
    '                       pUserID:=vbNullString, pPassword:=vbNullString
    'End With

    If Not (hrgHourglass Is Nothing) Then
        hrgHourglass.value = False
    End If

    ' Print report to modal Viewer window
    fnViewReport

    ' Make sure this window is shown on top of all other windows in the app
    ' after the Viewer window is closedlkj  lkj klj
    fnSetTopmostWindow Me, bTopmost:=True
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    
    ' Clean-up statements go here
    fnFreeObject hrgHourglass
    fnFreeObject crDB

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
    ' Params:      N/A
    ' Returns:     N/A
    ' Modified  :
    ' Modified:     Berry Kropiwka - Added Compact Calc - 2019-09-27
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
        .ClmNum = txtClmNum
        
        ' Just get the ADMN_SYST_CD (column 1) of the selected row in lpcAdmnSystCd
        lpcAdmnSystCd.Col = mcintStoreCol_lpcAdmnSystCd
        .AdmnSystCd = lpcAdmnSystCd.ColText
        
        .ClmPolNum = iptClmPolNum.Text
        .ClmForResDthInd = (chkClmForResDthInd.value = vbChecked)
        .ClmCompactClcnInd = (chkClmCmpCalInd.value = vbChecked)
        
        ' Just get the PYCO_TYP_CD (column 1) of the selected row in lpcPycoTypCd
        lpcPycoTypCd.Col = mcintStoreCol_lpcPycoTypCd
        .PycoTypCd = lpcPycoTypCd.ColText
        
        .ClmInsdFirstNm = iptClmInsdFirstNm.Text
        .ClmInsdLastNm = iptClmInsdLastNm.Text
        
        ' cboIssStCd corresponds to a Nullable field, so accommodate Nulls
        If cboIssStCd.Text = gcstrBlankEntry Then
            .IssStCd = vbNullString
        Else
            .IssStCd = cboIssStCd.Text
        End If
        ' cboInsdDthResStCd corresponds to a Nullable field, so accommodate Nulls
        If cboInsdDthResStCd.Text = gcstrBlankEntry Then
            .InsdDthResStCd = vbNullString
        Else
            .InsdDthResStCd = cboInsdDthResStCd.Text
        End If
        
        .ClmInsdDthDt = dtpClmInsdDthDt.value
        .ClmProofDt = dtpClmProofDt.value
        .ClmInsdSsnNum = ipmClmInsdSsnNum.UnFmtText      ' Use .UnFmtText to get rid of mask characters in fpMask control
        
        .ClmTotDthbPmtAmt = ipcClmTotDthbPmtAmt.value    ' Use .Value to get unformatted value of fpCurrency control
        .ClmTotIntAmt = ipcClmTotIntAmt.value
        .ClmTotWthldAmt = ipcClmTotWthldAmt.value
        .ClmTotClmPdAmt = ipcClmTotClmPdAmt.value

        .LstUpdtUserId = gconAppActive.LastLogOnUserID
        .LstUpdtDtm = Now
    End With
    
    ' These will propagate back an error if the Insert/Update failed.
    If mbInAddMode Then
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
                ' Can't use the GetClmNumFromClmId( ) method since the CLAIM_T row
                ' doesn't exist and hence a "-2147217900" (Claim ID does not exist) error
                ' throws things off.
                .GetRelativeRecord .ClmNum, epdPreviousRecord
            ElseIf lngReturnValue = vbObjectError + gcRES_INFO_ANOTHER_USER_UPDATED_DISCARDED Then
                gerhApp.ReportNonFatal vbObjectError + gcRES_INFO_ANOTHER_USER_UPDATED_DISCARDED, _
                                       mstrScreenName & gcstrDOT & cstrCurrentProc, _
                                       Trim$(strACF2)
                ' Discard *this* user's pending changes by re-retrieving the current record
                ' as it currently looks on the database and refreshing the lookup recordset
                ' Can't use the GetClmNumFromClmId( ) method since the CLAIM_T row
                ' doesn't exist and hence a "-2147217900" (Claim ID does not exist) error
                ' throws things off.
                .GetRelativeRecord .ClmNum, epdSameRecord
            Else
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
        fnSetCommandButtons True
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
Private Sub dtpClmInsdDthDt_Change()
    ' Comments  : Since this field was just changed, reset
    '             Enabled property on command and navigation
    '             buttons as appropriate given that the user
    '             is in the middle of updating a record.
    ' Parameters:  -
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "dtpClmInsdDthDt_Change"
 
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
Private Sub dtpClmProofDt_Change()
    ' Comments  : Since this field was just changed, reset
    '             Enabled property on command and navigation
    '             buttons as appropriate given that the user
    '             is in the middle of updating a record.
    ' Parameters:  -
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "dtpClmProofDt_Change"
 
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
Private Sub fnAddRecord()
    ' Comments  : This function handles adding a new record. It is called
    '             by cmdAdd_Click (when the user clicks the Add button)
    '             and by cmdDelete_Click (when the last record in the
    '             recordset is deleted)
    ' Parameters: N/A
    ' Returns   : N/A
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
                            
    If bDebugAppTermination Then
        Debug.Print "   Turning off Update mode (#1) in " & mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    IsDirty = False

    ' Initialize Payee recordset to avoid run-time error 91 - object variable or with block variable not set
    Set mrstPayees = New ADODB.Recordset
    ' Only show the 1st row (column headers) in Payee Grid
    msgPayees.Rows = 1

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
    ' Modified:     Berry Kropiwka - Added Compact Calc - 2019-09-27
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc          As String = "fnBindControlsToTableWrapper"
    On Error GoTo PROC_ERR
 
    lpcAdmnSystCd.Tag = "AdmnSystCd"
    iptClmPolNum.Tag = "ClmPolNum"
    lpcPycoTypCd.Tag = "PycoTypCd"
    iptClmInsdFirstNm.Tag = "ClmInsdFirstNm"
    iptClmInsdLastNm.Tag = "ClmInsdLastNm"
    cboIssStCd.Tag = "IssStCd"
    cboInsdDthResStCd.Tag = "InsdDthResStCd"
    dtpClmInsdDthDt.Tag = "ClmInsdDthDt"
    dtpClmProofDt.Tag = "ClmProofDt"
    ipmClmInsdSsnNum.Tag = "ClmInsdSsnNum"
    ipcClmTotDthbPmtAmt.Tag = "ClmTotDthbPmtAmt"
    ipcClmTotIntAmt.Tag = "ClmTotIntAmt"
    ipcClmTotWthldAmt.Tag = "ClmTotWthldAmt"
    ipcClmTotClmPdAmt.Tag = "ClmTotClmPdAmt"
    txtClmNum.Tag = "ClmNum"
    chkClmForResDthInd.Tag = "ClmForResDthInd"
    chkClmCmpCalInd.Tag = "ClmCompactClcnInd"
    
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
Private Function fnCalcTotalsForAllPayees(ByVal lngClmID As Long) As ADODB.Recordset
    ' Comments  : This function will add up all of the Payee for each
    '             policy/claim to produce totals
    ' Parameters:
    '     lngClmId (in) - the CLM_ID of the desired Claim
    ' Returns:     A disconnected ADODB.Recordset containing calculated
    '              columns for the specified key
    ' Modified  :
    ' --------------------------------------------------
    Const cstrCurrentProc       As String = "fnCalcTotalsForAllPayees"
    Const cstrSproc                As String = "dbo.proc_payee_totals_for_claim" ' Stored procedure to execute
    Dim rstTemp                    As ADODB.Recordset
    Dim prmReturnValue             As ADODB.Parameter
    Dim prmClmId                   As ADODB.Parameter
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

        ' ---Parameter #2---
        ' Define the CLM_ID parameter
        Set prmClmId = .CreateParameter(Name:="@clm_id", _
                                        Type:=adInteger, _
                                        Direction:=adParamInput, _
                                        value:=lngClmID)
        .Parameters.Append prmClmId

        Set rstTemp = .Execute()
    End With

     rstTemp.ActiveConnection = Nothing
    Set fnCalcTotalsForAllPayees = rstTemp

    With rstTemp
        ' Use fnZeroIfNull to accommodate Nulls in case there are no Payees yet defined for this claim
        ipcClmTotDthbPmtAmt.Text = fnZeroIfNull(!ClmTotDthbPmtAmt)
        ipcClmTotWthldAmt.Text = fnZeroIfNull(!ClmTotWthldAmt)
        ipcClmTotIntAmt.Text = fnZeroIfNull(!ClmTotIntAmt)
        ipcClmTotClmPdAmt.Text = fnZeroIfNull(!ClmTotClmPdAmt)
    End With
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    
    ' Clean-up statements go here
    fnFreeRecordset rstTemp
    fnFreeObject prmReturnValue
    fnFreeObject prmClmId
    fnFreeObject adwTemp
    
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
                                       "Claim ID " & RTrim$(lngClmID) & "/Claim Number " & mtWrapper.GetClmNumFromClmID(lngClmID)
            Resume PROC_EXIT
        Case gcRES_NERR_ERR_WHILE_TRYING_TO
            ' 4028 = An error occurred while attempting to @@1 this record.
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
Private Sub fnClearControls()
    ' Comments:   Initializes screen controls in order to add a new record
    ' Parameters: N/A
    ' Returns:    N/A
    ' Modified:
    ' Modified:     Berry Kropiwka - Added Compact Calc - 2019-09-27
    ' --------------------------------------------------
    Const cstrCurrentProc       As String = "fnClearControls"
    Const cintZero              As Integer = 0
    Const clngFirstEntry        As Long = 0
    Dim ctl                     As Control
    Dim varDefaultValue         As Variant
    Dim strSavedMask            As String
 
    On Error GoTo PROC_ERR
    
    ' Hide updates to the window until we're done. This avoids ugly screen flickering
    fnWindowLock Me.hWnd

    chkClmForResDthInd.value = vbUnchecked
    chkClmCmpCalInd.value = vbUnchecked
    
    ' Enable or Disable the Compact Filling check box based on Admin System
    fnSetCompactFillingCheckBox Me.lpcAdmnSystCd.Text, Me.cboIssStCd.Text
    
    ' Select the first entry (the blank row) in the Admin System's fpCombo control.
    fnSearchFPCombo lpcAdmnSystCd, gcstrBlankEntry, mcintStoreCol_lpcAdmnSystCd
    
    iptClmPolNum.Text = vbNullString
    
    ' Select the first entry (the blank row) in the Company Type's fpCombo control.
    fnSearchFPCombo lpcPycoTypCd, gcstrBlankEntry, mcintStoreCol_lpcPycoTypCd
    
    iptClmInsdFirstNm.Text = vbNullString
    iptClmInsdLastNm.Text = vbNullString

    If cboIssStCd.ListCount > 0 Then
        cboIssStCd.ListIndex = 0      ' Select first (blank) entry
    Else
        gerhApp.ReportNonFatal vbObjectError + gcRES_WARN_CBO_IS_EMPTY, _
                               mstrScreenName & gcstrDOT & cstrCurrentProc, _
                               mcstrCboIssStCdLabel
    End If

    If cboInsdDthResStCd.ListCount > cintZero Then
        cboInsdDthResStCd.ListIndex = cintZero      ' Select first (blank) entry
    Else
        gerhApp.ReportNonFatal vbObjectError + gcRES_WARN_CBO_IS_EMPTY, _
                               mstrScreenName & gcstrDOT & cstrCurrentProc, _
                               mcstrCboInsdDthResStCdLabel
    End If

    ' DateTimePicker controls (dtpClmInsdDthDt and dtpClmProofDt) will
    ' automatically be set to today's date. Cannot set them to Null
    ' unless their CheckBox property is set to True.
    dtpClmInsdDthDt.value = Date
    dtpClmProofDt.value = Date


    ' NOTE: For MaskEdBox controls, have to remove mask before clearing out the control
    '       since the vbNullString value doesn't match the mask specification.
    strSavedMask = ipmClmInsdSsnNum.Mask
    ipmClmInsdSsnNum.Mask = vbNullString
    ipmClmInsdSsnNum.Text = vbNullString
    ipmClmInsdSsnNum.Mask = strSavedMask

    ' ' Select the "VUL" entry in the Product Family ComboBox, if present, otherwise select
    ' ' the first entry. If the ComboBox is empty, display a message to the user to warn them of
    ' ' unpredictible behavior.
    ' If cboPfamCd.ListCount > 0 Then
    '     lngEntryFoundSlot = fnFindStringComboBox(cboIn:=cboPfamCd, strSearchIn:="VUL     ", bDoExactSearch:=True)
    '     If lngEntryFoundSlot = clngNotFound Then
    '         cboPfamCd.ListIndex = clngFirstEntry
    '     Else
    '         cboPfamCd.ListIndex = lngEntryFoundSlot
    '     End If
    ' Else
    '     gerhApp.ReportNonFatal vbObjectError + gcRES_WARN_CBO_IS_EMPTY, _
    '                            mstrScreenName & gcstrDOT & cstrCurrentProc, _
    '                            mcstrCboPfamCdLabel
    ' End If

    ipcClmTotDthbPmtAmt.Text = cintZero
    ipcClmTotIntAmt.Text = cintZero
    ipcClmTotWthldAmt.Text = cintZero
    ipcClmTotClmPdAmt.Text = cintZero
    
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
                    ElseIf (TypeOf ctl Is fpMask) Then
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
Private Sub fnFillPayeeGrid()
    ' Comments  : Loads the MSFlexGrid control with
    '             Payee data for the current Insured
    ' Called By : fnGetChildren()
    '
    ' Parameters:  -
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnFillPayeeGrid"
    Dim intRecordCounter As Integer
 

'!TODO! Change to use vsflexgrid
    With msgPayees
        ' Set Rows to reflect # of records +1 (for header row)
        .Rows = mrstPayees.RecordCount + 1
        ' Fill in columns of grid per current recordset row
        For intRecordCounter = 1 To mrstPayees.RecordCount
            .Row = intRecordCounter

            ' Column 1 - Counter
            .Col = 0
            .Text = intRecordCounter
            ' Column 2 - Payee Full Name
            .Col = 1
            .Text = mrstPayees!paye_full_nm
            ' Column 3 - Address Line 1
            .Col = 2
            .Text = fnZLSIfNull(mrstPayees!paye_addr_ln1_txt)
            ' Column 4 - Address Line2
            .Col = 3
            .Text = fnZLSIfNull(mrstPayees!paye_addr_ln2_txt)
            ' Column 5 - Payee Residence State
            .Col = 4
            .Text = fnZLSIfNull(mrstPayees!calc_st_cd)
            ' Column 6 - Date Of Payment
            .Col = 5
            .Text = mrstPayees!paye_pmt_dt
            ' Column 7 - TIN/SSN
            .Col = 6
            .Text = fnZLSIfNull(mrstPayees!paye_ssn_tin_num)
'!TODO! Change to use meta data for formatting!
            ' Column 8 - Interest Amt
            .Col = 7
            .Text = Format$(mrstPayees!paye_clm_int_amt, "###,###,##0.00")
            ' Column 9 - Total Claim Amt for Payee
            .Col = 8
            .Text = Format$(mrstPayees!paye_clm_pd_amt, "###,###,##0.00")
            ' Column 10 - DB Payment
            .Col = 9
            .Text = Format$(mrstPayees!paye_dthb_pmt_amt, "###,###,##0.00")
            ' Column 11 - Interest Rate
            .Col = 10
            .Text = Format$(mrstPayees!paye_clm_int_rt, "###,##0.00000")
            ' Column 12 - Withholding Rate
            .Col = 11
            .Text = Format$(mrstPayees!paye_wthld_rt, "###,##0.00000")
            ' Column 13 - Interest Withheld
            ' TotalAmt is reduced by the Withheld Amt, so show Withheld
            ' Amt as a negative number. (It is stored as a positive number.)
            .Col = 12
            .Text = Format$(mrstPayees!paye_wthld_amt, "(###,###,##0.00)")
            ' Column 14 - Payee ID
            ' This is needed so the Insured screen can tell the Payee
            ' screen which Payee to display
            .Col = 13
            .Text = mrstPayees!paye_id
            ' Make the width=0 to effectively hide it
            .ColWidth(13) = 0

            ' Read next record in recordset and loop
            mrstPayees.MoveNext
         Next intRecordCounter
         
         .Row = 1   ' 1st (non-column header) row
         .Col = 1   ' 2nd column - Payee name
    End With

    fnCalcTotalsForAllPayees mtWrapper.ClmId
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
Private Sub fnGetAdminSysMetadata()
    ' Comments  : This function looks up admin system metadata based on the current
    '             value selected in the Admin System combobox.
    ' Parameters: N/A
    ' Returns:     N/A. However it sets some module-level variables.
    ' Modified  :
    ' --------------------------------------------------
    Const cstrCurrentProc       As String = "fnGetAdminSysMetadata"
    Const cstrSproc                As String = "dbo.proc_admin_system_select2" ' Stored procedure to execute
    Dim prmReturnValue             As ADODB.Parameter
    Dim prmAdmnSystCd              As ADODB.Parameter
    Dim prmMinLength               As ADODB.Parameter
    Dim prmMaxLength               As ADODB.Parameter
    Dim prmDfltPycoTypDsc          As ADODB.Parameter
    Dim prmTaxRptgInd              As ADODB.Parameter
    Dim adwTemp                    As cadwADOWrapper
   
    On Error GoTo PROC_ERR
    
    ' Set default values in case of error
    mintAdmnSyst_MinPolNumLength = 1
    mintAdmnSyst_MaxPolNumLength = mtWrapper.MaxCharacters(iptClmPolNum.Tag)
    mstrAdmnSyst_TaxRptgInd = vbNullString
    mstrAdmnSyst_DfltPycoTypDsc = vbNullString
    
    ' Just get the ADMN_SYST_CD (column 1) of the selected row in lpcAdmnSystCd.
    ' If it hasn't been input yet (i.e. it is still blank) then just accept default
    ' values set above and bypass the sproc call. This avoids a SQL error hit when
    ' passing an invalid value into the @admn_syst_cd parameter.
    lpcAdmnSystCd.Col = mcintStoreCol_lpcAdmnSystCd
    If (lpcAdmnSystCd.ColText = vbNullString) Or (lpcAdmnSystCd.ColText = gcstrBlankEntry) Then
        GoTo PROC_EXIT
    End If
    
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
        ' Define the ADMN_SYST_CD input parameter
        Set prmAdmnSystCd = .CreateParameter(Name:="@admn_syst_cd", _
                                        Type:=adChar, _
                                        Direction:=adParamInput, _
                                        value:=lpcAdmnSystCd.ColText, _
                                        Size:=2)
        .Parameters.Append prmAdmnSystCd
        
          ' ---Parameter #3---
        ' Define the MIN_LENGTH output parameter
        Set prmMinLength = .CreateParameter(Name:="@MinLength", _
                                         Type:=adSmallInt, _
                                         Direction:=adParamOutput, _
                                         Size:=2)
        .Parameters.Append prmMinLength

          ' ---Parameter #4---
        ' Define the MAX_LENGTH output parameter
        Set prmMaxLength = .CreateParameter(Name:="@MaxLength", _
                                         Type:=adSmallInt, _
                                         Direction:=adParamOutput, _
                                         Size:=2)
        .Parameters.Append prmMaxLength
        
        ' ---Parameter #5---
        ' Define the DFLT_PYCO_TYP_DSC output parameter
        Set prmDfltPycoTypDsc = .CreateParameter(Name:="@DfltPycoTypDsc", _
                                        Type:=adVarChar, _
                                        Direction:=adParamOutput, _
                                        Size:=60)
        .Parameters.Append prmDfltPycoTypDsc
    
        ' ---Parameter #6---
        ' Define the TAX_RPTG_IND output parameter
        Set prmTaxRptgInd = .CreateParameter(Name:="@TaxRptgInd", _
                                        Type:=adChar, _
                                        Direction:=adParamOutput, _
                                        Size:=1)
        .Parameters.Append prmTaxRptgInd
        
        .Execute
    End With

    
    mintAdmnSyst_MinPolNumLength = prmMinLength.value
    mintAdmnSyst_MaxPolNumLength = prmMaxLength.value
    mstrAdmnSyst_TaxRptgInd = prmTaxRptgInd.value
    mstrAdmnSyst_DfltPycoTypDsc = prmDfltPycoTypDsc.value
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    
    ' Clean-up statements go here
    fnFreeObject prmReturnValue
    fnFreeObject prmAdmnSystCd
    fnFreeObject prmMinLength
    fnFreeObject prmMaxLength
    fnFreeObject prmDfltPycoTypDsc
    fnFreeObject prmTaxRptgInd
    
    fnFreeObject adwTemp
    
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case prmReturnValue
        Case gcRES_NERR_REC_NOT_FOUND
            ' 4027 = The specified record was not found in the database (@@1).
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_REC_NOT_FOUND, _
                                       mstrScreenName & gcstrDOT & cstrCurrentProc, _
                                       mcstrLpcAdmnSystCdLabel & ": " & lpcAdmnSystCd.ColText
            Resume PROC_EXIT
        Case gcRES_NERR_ERR_WHILE_TRYING_TO
            ' 4028 = An error occurred while attempting to @@1 this record.
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_ERR_WHILE_TRYING_TO, _
                                       mstrScreenName & gcstrDOT & cstrCurrentProc, _
                                       "locate the " & mcstrLpcAdmnSystCdLabel & " metadata for"
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
Private Sub fnGetChildren()
    ' Comments  : Loads data associated from tables that are
    '             subordinate (i.e. children) to the table
    '             supplying the main data for this form
    ' Parameters:  -
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnGetChildren"

 
    ' --- Build the Recordset object for Payee data (mrstPayees) ---
    '     that's associated with the current Insured.

    Set mrstPayees = mtPayee.GetPayeesForClaim(mtWrapper.ClmId)

    ' Load MSFlexGrid with Payee records, if any. Disallow Delete
    ' of Insured/Claim record if there are Payee records. (The user
    ' must delete Payees before attempting to delete the Insured/Claim.)
    If mrstPayees.RecordCount > 0 Then
        fnFillPayeeGrid
    Else
        ' Only show the 1st row (column headers)
        msgPayees.Rows = 1
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
Private Function fnGetData_IndividualReport() As ADODB.Recordset
    '--------------------------------------------------------------------------
    ' Procedure:   fnGetData_IndividualReport
    ' Description: Builds a recordset containing data needed to send
    '              to the .RPT file associated with a report.
    '
    '
    ' Parameters:  N/A
    '
    ' Returns:     A disconnected ADODB.Recordset
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc     As String = "fnGetData_IndividualReport"
    Const cstrSQLView         As String = "dbo.IndividualReport_v"
    Dim strSQL                As String
    Dim strWhereClmID         As String
    Dim strOrderBy            As String
    Dim rstTemp               As ADODB.Recordset
 
    On Error GoTo PROC_ERR

    strWhereClmID = " WHERE clm_id = " & mtWrapper.ClmId & vbCr
    strOrderBy = " ORDER BY clm_id, paye_full_nm"

    strSQL = "SELECT * from " & cstrSQLView & strWhereClmID & strOrderBy
    
    Set rstTemp = gadwApp.Execute_SQL_AsRST(gconAppActive, strSQL)
    #If DEBUG_RST Then
        Debug.Print "In " & cstrCurrentProc & ", " & CStr(rstTemp.RecordCount) & " records were retrieved in the rst."
        Debug.Print "SQL statement is: " & vbCr & strSQL
    #End If
    
    ' Disconnect the recordset
    rstTemp.ActiveConnection = Nothing
    
    Set fnGetData_IndividualReport = rstTemp
PROC_EXIT:
    On Error GoTo 0     ' Disable error handler
    
    ' Clean-up statements go here
    
    ' DO NOT do "fnFreeRecordset rstTemp" since this will cause the recordset
    ' returned by this function to be wiped out as well!
    
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            ' Save Err object data, if not already saved
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
    ' Modified:     Berry Kropiwka - Added Compact Calc - 2019-09-27
    
    '!CUSTOMIZE!  There should be one Case statement for each control that
    '             corresponds to a table column. Each Case statement should
    '             reference a Const literal that indicates how the control is
    '             labelled on-screen.

    Const cstrCurrentProc       As String = "fnGetFieldLabel"
 
    On Error GoTo PROC_ERR

    Select Case strControlName
        Case "lpcAdmnSystCd"
            fnGetFieldLabel = mcstrLpcAdmnSystCdLabel
        Case "iptClmPolNum"
            fnGetFieldLabel = mcstrIptClmPolNumLabel
        Case "lpcPycoTypCd"
            fnGetFieldLabel = mcstrLpcPycoTypCdLabel
        Case "iptClmInsdFirstNm"
            fnGetFieldLabel = mcstrIptClmInsdFirstNmLabel
        Case "iptClmInsdLastNm"
            fnGetFieldLabel = mcstrIptClmInsdLastNmLabel
        Case "cboIssStCd"
            fnGetFieldLabel = mcstrCboIssStCdLabel
        Case "cboInsdDthResStCd"
            fnGetFieldLabel = mcstrCboInsdDthResStCdLabel
        Case "dtpClmInsdDthDt"
            fnGetFieldLabel = mcstrDtpClmInsdDthDtLabel
        Case "dtpClmProofDt"
            fnGetFieldLabel = mcstrDtpClmProofDtLabel
        Case "ipmClmInsdSsnNum"
            fnGetFieldLabel = mcstrIpmClmInsdSsnNumLabel
        Case "ipcClmTotDthbPmtAmt"
            fnGetFieldLabel = mcstrIpcClmTotDthbPmtAmtLabel
        Case "ipcClmTotIntAmt"
            fnGetFieldLabel = mcstrIpcClmTotIntAmtLabel
        Case "ipcClmTotWthldAmt"
            fnGetFieldLabel = mcstrIpcClmTotWthldAmtLabel
        Case "ipcClmTotClmPdAmt"
            fnGetFieldLabel = mcstrIpcClmTotClmPdAmtLabel
        Case "txtClmNum"
            fnGetFieldLabel = mcstrTxtClmNumLabel
        Case "chkClmForResDthInd"
            fnGetFieldLabel = mcstrChkClmForResDthIndLabel
        Case "chkClmCmpCalInd"
            fnGetFieldLabel = mcstrChkClmCmpCalIndLabel
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
Private Function fnGetLobCd() As String
    '--------------------------------------------------------------------------
    ' Procedure:   fnGetLobCd
    ' Description: This procedure gets the Line-of-business, given what the
    '              Admin System fpCombo box is currently set to. It defaults to
    '              "I" (for Individual).
    '
    ' Params:      N/A
    ' Returns:     "G" if a Group-based Admin System is selected; "I" otherwise
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc           As String = "fnGetLobCd"
    On Error GoTo PROC_ERR
 
    If (lpcAdmnSystCd.ColText <> vbNullString) And (lpcAdmnSystCd.ColText <> gcstrBlankEntry) Then
        fnGetLobCd = mtWrapper.GetLobCdFromAdmnSystCd(lpcAdmnSystCd.ColText)
    Else
        fnGetLobCd = mcstrIndividualLOB
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
        Case Else
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Function fnGetPayeesNeedingRecalcDueToDeath(ByVal lngClmIdIn As Long, ByVal dteClmInsdDthDtIn As Date) As Long
    '--------------------------------------------------------------------------
    ' Procedure:   fnGetPayeesNeedingRecalcDueToDeath
    ' Description: Returns the number of Payees for the claim that have a
    '              Date of Payment prior to the Date of Death
    ' Params:      N/A
    ' Returns:     N/A
    ' Date:        04/12/2002
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc          As String = "fnGetPayeesNeedingRecalcDueToDeath"
    Const cstrSproc                As String = "dbo.proc_payee_select3" ' Stored procedure to execute
    Dim rstTemp                    As ADODB.Recordset
    Dim prmReturnValue             As ADODB.Parameter
    Dim prmClmId                   As ADODB.Parameter
    Dim prmClmInsdDthDt            As ADODB.Parameter
    Dim prmNbrOfPayees             As ADODB.Parameter
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

        ' ---Parameter #2---
        ' Define the CLM_ID parameter
        Set prmClmId = .CreateParameter(Name:="@clm_id", _
                                        Type:=adInteger, _
                                        Direction:=adParamInput, _
                                        value:=lngClmIdIn)
        .Parameters.Append prmClmId

        ' ---Parameter #3---
        ' Define the CLM_PROOF_DT parameter
        Set prmClmInsdDthDt = .CreateParameter(Name:="@clm_insd_dth_dt", _
                                         Type:=adDBTimeStamp, _
                                         Direction:=adParamInput, _
                                         Size:=16, _
                                         value:=dteClmInsdDthDtIn)
        .Parameters.Append prmClmInsdDthDt

        ' ---Parameter #4---
        ' Define the NBR_OF_PAYEES parameter
        Set prmNbrOfPayees = .CreateParameter(Name:="@nbr_of_payees", _
                                        Type:=adInteger, _
                                        Direction:=adParamInputOutput, _
                                        value:=Null)
        .Parameters.Append prmNbrOfPayees

        Set rstTemp = .Execute()
    End With

    fnGetPayeesNeedingRecalcDueToDeath = prmNbrOfPayees.value
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0

    ' Clean-up statements go here
    fnFreeRecordset rstTemp
    fnFreeObject adwTemp
    fnFreeObject prmClmId
    fnFreeObject prmClmInsdDthDt
    fnFreeObject prmNbrOfPayees

    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case prmReturnValue
        Case gcRES_NERR_REC_NOT_FOUND       ' 4027
            ' Note that this error is presented as a 4027 rather than a 4037!
            ' 4037 = The @@1 is invalid. @@2
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_INVALID_DATA, _
                                       mstrScreenName & gcstrDOT & cstrCurrentProc, _
                                       "Claim ID or Date of Proof", _
                                       "The need for any Payee recalculation cannot be determined " & _
                                       "when any of these fields are NULL."
            Resume PROC_EXIT
        Case gcRES_NERR_ERR_WHILE_TRYING_TO
            ' 4028 = An error occurred while attempting to @@1 this record.
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
Private Function fnGetPayeesNeedingRecalcDueToProof(ByVal lngClmIdIn As Long, ByVal dteClmProofDtIn As Date) As Long
    '--------------------------------------------------------------------------
    ' Procedure:   fnGetPayeesNeedingRecalcDueToProof
    ' Description: Returns the number of Payees for the claim that have a
    '              Date of Payment prior to the Date of Proof
    ' Params:      N/A
    ' Returns:     N/A
    ' Date:        04/12/2002
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc          As String = "fnGetPayeesNeedingRecalcDueToProof"
    Const cstrSproc                As String = "dbo.proc_payee_select2" ' Stored procedure to execute
    Dim rstTemp                    As ADODB.Recordset
    Dim prmReturnValue             As ADODB.Parameter
    Dim prmClmId                   As ADODB.Parameter
    Dim prmClmProofDt              As ADODB.Parameter
    Dim prmNbrOfPayees             As ADODB.Parameter
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

        ' ---Parameter #2---
        ' Define the CLM_ID parameter
        Set prmClmId = .CreateParameter(Name:="@clm_id", _
                                        Type:=adInteger, _
                                        Direction:=adParamInput, _
                                        value:=lngClmIdIn)
        .Parameters.Append prmClmId

        ' ---Parameter #3---
        ' Define the CLM_PROOF_DT parameter
        Set prmClmProofDt = .CreateParameter(Name:="@clm_proof_dt", _
                                         Type:=adDBTimeStamp, _
                                         Direction:=adParamInput, _
                                         Size:=16, _
                                         value:=dteClmProofDtIn)
        .Parameters.Append prmClmProofDt

        ' ---Parameter #4---
        ' Define the NBR_OF_PAYEES parameter
        Set prmNbrOfPayees = .CreateParameter(Name:="@nbr_of_payees", _
                                        Type:=adInteger, _
                                        Direction:=adParamInputOutput, _
                                        value:=Null)
        .Parameters.Append prmNbrOfPayees

        Set rstTemp = .Execute()
    End With

    fnGetPayeesNeedingRecalcDueToProof = prmNbrOfPayees.value
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0

    ' Clean-up statements go here
    fnFreeRecordset rstTemp
    fnFreeObject adwTemp
    fnFreeObject prmClmId
    fnFreeObject prmClmProofDt
    fnFreeObject prmNbrOfPayees

    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case prmReturnValue
        Case gcRES_NERR_REC_NOT_FOUND       ' 4027
            ' Note that this error is presented as a 4027 rather than a 4037!
            ' 4037 = The @@1 is invalid. @@2
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_INVALID_DATA, _
                                       mstrScreenName & gcstrDOT & cstrCurrentProc, _
                                       "Claim ID or Date of Proof", _
                                       "The need for any Payee recalculation cannot be determined " & _
                                       "when any of these fields are NULL."
            Resume PROC_EXIT
        Case gcRES_NERR_ERR_WHILE_TRYING_TO
            ' 4028 = An error occurred while attempting to @@1 this record.
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
Private Function fnGetReportFile() As String
    ' Comments  : Using the selection in the Select A Report ListBox,
    '             this proc retrieves the corresponding .RPT's filename
    '             from the Report Meta Data array.
    ' Parameters: N/A
    ' Returns   : String - the name of the .RPT file for that report
    ' Modified  :
    '
    ' --------------------------------------------------
    Const cstrCurrentProc As String = "fnGetReportFile"
    Dim fso               As Scripting.FileSystemObject
 
    On Error GoTo PROC_ERR
    
    Set fso = New Scripting.FileSystemObject

    fnGetReportFile = fso.BuildPath(App.Path, "Individual_CR8.rpt")

    ' Non-fatal error if .RPT doesn't exist or if we couldn't determine the .RPT filename
    If Not (fso.FileExists(fnGetReportFile)) Then
        gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_RPTFILE_NOT_FOUND, _
                                   mstrScreenName & gcstrDOT & cstrCurrentProc, _
                                   fnGetReportFile
        GoTo PROC_EXIT
    End If
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0
    ' Clean-up statements go here
    fnFreeObject fso
    
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mstrScreenName & gcstrDOT & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            ' Save Err object data, if not already saved
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function

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
Private Function fnGetDefaultPayorCompany(ByVal strClmPolNum As String, ByVal strAdmnSystCd As String) As String
    ' Description: Parses supplied Policy Number and/or uses metadata to determine whether it represents a
    '              a special policy - one whose Parent Company should be
    '              set to a particular value. This includes:
    '                  * AdmnSystCd = 02 (ALIS)  -- all policies
    '                  *              22 (CYBER) -- all policies
    '                  *              37 (VPAS)  -- all policies
    '                  *              SOLAR policies beginning with 'UL'
    ' Parameters : N/A
    ' Returns    : Default Payor Company value
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc   As String = "fnGetDefaultPayorCompany"
    Dim strChars1Thru2      As String
    
    strChars1Thru2 = UCase$(Left$(strClmPolNum, 2))
    
    fnGetDefaultPayorCompany = mstrAdmnSyst_DfltPycoTypDsc
    
    ' Override admin system default, if policy # is "special" SOLAR range
    If (strChars1Thru2 = "UL") Then
        fnGetDefaultPayorCompany = mcstrPyco_Subsidiary
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
        Case Else
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function


'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub fnLoadCboInsdDthResStCd()
    '--------------------------------------------------------------------------
    ' Procedure:   fnLoadcboInsdDthResStCd
    ' Description: Populates the Residence State combobox using a sproc
    ' Params:      N/A
    ' Returns:     N/A
    ' Modified:
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc          As String = "fnLoadCboInsdDthResStCd"
    Const cstrSproc                As String = "dbo.proc_state_lu_select" ' Stored procedure to execute
    Dim rstTemp                    As ADODB.Recordset
    Dim prmReturnValue             As ADODB.Parameter
    Dim adwTemp                    As cadwADOWrapper
 
    On Error GoTo PROC_ERR

    Set adwTemp = New cadwADOWrapper

    cboInsdDthResStCd.Clear

    ' Add a blank entry as the first entry of the combobox. This will force the user to select
    ' an entry (no default selection) since fnValidData will generate an error if the blank
    ' entry is still selected when the user clicks Update.
    cboInsdDthResStCd.AddItem gcstrBlankEntry
    
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

    'Filter out Compact Filling State
    rstTemp.Filter = "st_cd <> '" & cstCompactFilling & "'"

    ' Add the following, if the combobox contains a hidden ID column associated with the
    ' column that *is* displayed:      varItemDataColumn:="co_id",
    fnADORecordSetToComboBox rstIn:=rstTemp, _
                             cboIn:=cboInsdDthResStCd, _
                             strDisplayColumn:="st_cd", _
                             bClear:=False
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0

    ' Clean-up statements go here
    rstTemp.Close
    fnFreeRecordset rstTemp
    fnFreeObject adwTemp
    fnFreeObject prmReturnValue

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
Private Sub fnLoadCboIssStCd()
    '--------------------------------------------------------------------------
    ' Procedure:   fnLoadcboIssStCd
    ' Description: Populates the Issue State combobox using a sproc
    ' Params:      N/A
    ' Returns:     N/A
    ' Modified:
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc          As String = "fnLoadCboIssStCd"
    Const cstrSproc                As String = "dbo.proc_state_lu_select" ' Stored procedure to execute
    Dim rstTemp                    As ADODB.Recordset
    Dim prmReturnValue             As ADODB.Parameter
    Dim adwTemp                    As cadwADOWrapper
 
    On Error GoTo PROC_ERR

    Set adwTemp = New cadwADOWrapper

    cboIssStCd.Clear

    ' Add a blank entry as the first entry of the combobox. This will force the user to select
    ' an entry (no default selection) since fnValidData will generate an error if the blank
    ' entry is still selected when the user clicks Update.
    cboIssStCd.AddItem gcstrBlankEntry
    
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
    
    'Filter out Compact Filling State
    rstTemp.Filter = "st_cd <> '" & cstCompactFilling & "'"

    ' Add the following, if the combobox contains a hidden ID column associated with the
    ' column that *is* displayed:      varItemDataColumn:="co_id",
    fnADORecordSetToComboBox rstIn:=rstTemp, _
                             cboIn:=cboIssStCd, _
                             strDisplayColumn:="st_cd", _
                             bClear:=False
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0

    ' Clean-up statements go here
    rstTemp.Close
    fnFreeRecordset rstTemp
    fnFreeObject adwTemp
    fnFreeObject prmReturnValue

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
Private Sub fnLoadControls()
    ' Description: Will take applicable value from the recordset and put them into
    '              the screen controls
    ' Params:      N/A
    ' Returns:     N/A
    ' Modified  :
    ' --------------------------------------------------
    ' Modified:     Berry Kropiwka - Added Compact Calc - 2019-09-27
    Const cstrCurrentProc       As String = "fnLoadControls"
    Dim strSavedMask            As String
 
    On Error GoTo PROC_ERR
    
    With mtWrapper
        ' The AdmnSystCd is displayed via a fpCombo control since it must display a
        ' Description but store a Code. Search the control's contents for the Code,
        ' so that the row with that Code's corresponding Description will be selected
        fnSearchFPCombo lpcAdmnSystCd, .AdmnSystCd, mcintStoreCol_lpcAdmnSystCd
        
        ' The following will trigger iptClmPolNum_Change( ) which sets the line-of-business (LOB)
        ' and then repopulates the Admin System combo box based on the LOB
        iptClmPolNum.Text = Trim$(.ClmPolNum)
        
        If .ClmForResDthInd Then
            chkClmForResDthInd.value = vbChecked
        Else
            chkClmForResDthInd.value = vbUnchecked
        End If
        
        If .ClmCompactClcnInd Then
            chkClmCmpCalInd.value = vbChecked
        Else
            chkClmCmpCalInd.value = vbUnchecked
        End If

        ' Set the availability of the InsdDthResStCd based on the Foreign Residence at Death checkbox selection.
        fnSetInsdDthResStCdAvailability
        
        ' The PycoTypCd is displayed via a fpCombo control since it must display a
        ' Description but store a Code. Search the control's contents for the Code,
        ' so that the row with that Code's corresponding Description will be selected
        fnSearchFPCombo lpcPycoTypCd, .PycoTypCd, mcintStoreCol_lpcPycoTypCd

        iptClmInsdFirstNm = .ClmInsdFirstNm
        iptClmInsdLastNm = .ClmInsdLastNm
        
        ' cboIssStCd corresponds to a Nullable field, so accommodate Nulls
        If .IssStCd = vbNullString Then
            cboIssStCd.Text = gcstrBlankEntry
        Else
            cboIssStCd.Text = .IssStCd
        End If
        
        ' cboInsdDthResStCd corresponds to a Nullable field, so accommodate Nulls
        If .InsdDthResStCd = vbNullString Then
            cboInsdDthResStCd.Text = gcstrBlankEntry
        Else
            cboInsdDthResStCd.Text = .InsdDthResStCd
        End If
        
        dtpClmInsdDthDt.value = .ClmInsdDthDt
        dtpClmProofDt.value = .ClmProofDt
        
'!TODO! Should these Original dates be of type Date?
        ' Save the original value of these 2 dates fields. If they change and Payees
        ' exist at that time, a warning should be issued to indicate the change may
        ' necessitate a recalculation of the Payee's values.
        mstrOrigDateOfDeath = .ClmInsdDthDt
        mstrOrigDateOfProof = .ClmProofDt

        ' NOTE: For MaskEdBox controls, have to do special processing based on whether or not the
        '       field is empty, to avoid a 380 "invalid property value" runtime error.
        '       * If it's empty, temporarily delete the mask, set the value, and then restore
        '         the mask.
        '       * If it's not empty, format the value so it will be "valid" per the .Mask
        '         (for phone numbers, this means inserting a dash between characters 3 and 4).
        If LenB(.ClmInsdSsnNum) = 0 Then
            strSavedMask = ipmClmInsdSsnNum.Mask
            ipmClmInsdSsnNum.Mask = vbNullString
            ipmClmInsdSsnNum.Text = vbNullString
            ipmClmInsdSsnNum.Mask = strSavedMask
        Else
            ipmClmInsdSsnNum.Text = fnSSNTIN_AddDash(strIn:=.ClmInsdSsnNum, bIsTin:=False)
        End If

        ipcClmTotDthbPmtAmt.Text = .ClmTotDthbPmtAmt
        ipcClmTotIntAmt.Text = .ClmTotIntAmt
        ipcClmTotWthldAmt.Text = .ClmTotWthldAmt
        ipcClmTotClmPdAmt.Text = .ClmTotClmPdAmt
        txtClmNum.Text = .ClmNum
        
        ' ClmId         isn't shown on-screen
        ' LstUpdDtm     isn't shown on-screen
        ' LstUpdUserId  isn't shown on-screen
    End With

    ' Get the Payees associated with the claim and populate the Payees grid
    fnGetChildren

    ' Make sure Navigation buttons are enabled/disabled based on current record position in the Lookup recordset
    fnSetNavigationButtons bUnconditionalDisable:=False

    ' Update the "record x of y" label
    lblRecordPosition = fnShowRecordPosition(mtWrapper.LookupData)
    
    ' Enable or Disable the Compact Filling check box based on Admin System
    fnSetCompactFillingCheckBox Me.lpcAdmnSystCd.Text, Me.cboIssStCd.Text

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
Private Sub fnLoadLpcAdmnSystCd()
    '--------------------------------------------------------------------------
    ' Procedure:   fnLoadLpcAdmnSystCd
    ' Description: Populates the Admin System fpCombo control using a sproc
    ' Params:      N/A
    ' Returns:     N/A
    ' Modified:
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc          As String = "fnLoadLpcAdmnSystCd"
    Const cstrSproc                As String = "dbo.proc_admin_system_lu_select" ' Stored procedure to execute
    Dim rstTemp                    As ADODB.Recordset
    Dim prmReturnValue             As ADODB.Parameter
    Dim adwTemp                    As cadwADOWrapper
 
    On Error GoTo PROC_ERR

    Set adwTemp = New cadwADOWrapper

    With lpcAdmnSystCd
        .Clear
        ' Add a blank entry as the first entry of the combobox. This will force the user to select
        ' an entry (no default selection) since fnValidData will generate an error if the blank
        ' entry is still selected when the user clicks Update.
        
        ' Set .Row to -1 so insertion works okay whether or not the fpCombo is sorted
        .Row = gclngNoSelection
        .InsertRow = gcstrBlankEntry & vbTab & gcstrBlankEntry
    End With
    
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

    With rstTemp
        If .RecordCount <> 0 Then
            .MoveFirst
            Do Until .EOF
                ' Set .Row to -1 so insertion works okay whether or not the fpCombo is sorted
                lpcAdmnSystCd.Row = gclngNoSelection
                lpcAdmnSystCd.InsertRow = .Fields(mcstrAdmnSystDsc).value & vbTab & .Fields(mcstrAdmnSystCd).value
                .MoveNext
            Loop
        Else
            gerhApp.ReportNonFatal vbObjectError + gcRES_WARN_CBO_IS_EMPTY, _
                                   mstrScreenName & gcstrDOT & cstrCurrentProc, _
                                   mcstrLpcAdmnSystCdLabel
        End If
    End With
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0

    ' Clean-up statements go here
    fnFreeRecordset rstTemp
    fnFreeObject adwTemp
    fnFreeObject prmReturnValue

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
Private Sub fnLoadLpcLookup(ByRef lpcIn As LPLib.fpCombo, ByVal lngLookupType As EnumLookupType)
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
    
  
    With lpcIn
        .Clear
        .Row = gclngNoSelection
        .SortState = SortStateSuspend
    End With
    
    Select Case lngLookupType
        Case elt_Claim
            aRows = mtWrapper.LookupData_Claim()
            With lpcIn
                .InsertRow = gcstrBlankEntry & vbTab & gcstrBlankEntry
                For lngRow = 0 To UBound(aRows, cintRowDimension)
                    ' There are 2 columns in the array and fpCombo control (indexed 0 thru 1).
                    .InsertRow = aRows(0, lngRow) & vbTab & aRows(1, lngRow)
                Next
                ' Reset property to ensure whole width displays
                .DataAutoSizeCols = DataAutoSizeColsMaxColWidth
            End With
        Case elt_Name
            aRows = mtWrapper.LookupData_Name()
            With lpcIn
                .InsertRow = gcstrBlankEntry & vbTab & gcstrBlankEntry & vbTab & gcstrBlankEntry & vbTab & gcstrBlankEntry
                For lngRow = 0 To UBound(aRows, cintRowDimension)
                    ' There are 4 columns in the array and fpCombo control (indexed 0 thru 3)
                    .InsertRow = aRows(0, lngRow) & vbTab & aRows(1, lngRow) & vbTab & aRows(2, lngRow) & vbTab & aRows(3, lngRow)
                Next
                ' Reset property to ensure whole width displays
                .DataAutoSizeCols = DataAutoSizeColsMaxColWidth
            End With
        Case elt_SSN
            aRows = mtWrapper.LookupData_SSN()
            With lpcIn
                .InsertRow = gcstrBlankEntry & vbTab & gcstrBlankEntry & vbTab & gcstrBlankEntry
                For lngRow = 0 To UBound(aRows, cintRowDimension)
                    ' There are 3 columns in the array and fpCombo control (indexed 0 thru 2)
                    .InsertRow = aRows(0, lngRow) & vbTab & aRows(1, lngRow) & vbTab & aRows(2, lngRow)
                Next
                ' Reset property to ensure whole width displays
                .DataAutoSizeCols = DataAutoSizeColsMaxColWidth
            End With
        Case Else
            gerhApp.SaveAppSpecificErr vbObjectError + gcRES_NERR_UNEXPECTED_VAL_SELECT_CASE, _
                mstrScreenName & gcstrDOT & cstrCurrentProc
            GoTo PROC_EXIT
    End Select
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
        Case Else
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub fnLoadLpcPycoTypCd()
    ' Comments  : Populates Company Type fpCombo control using a sproc
    ' Parameters: N/A
    ' Returns   : N/A
    ' Modified  :
    ' --------------------------------------------------
    Const cstrCurrentProc As String = "fnLoadLpcPycoTypCd"
    Const cstrSproc                As String = "dbo.proc_payor_company_type_lu_select" ' Stored procedure to execute
    Dim rstTemp                    As ADODB.Recordset
    Dim prmReturnValue             As ADODB.Parameter
    Dim adwTemp                    As cadwADOWrapper
 
    On Error GoTo PROC_ERR

    Set adwTemp = New cadwADOWrapper

    With lpcPycoTypCd
        .Clear
        ' Add a blank entry as the first entry of the combobox. This will be the default entry
        ' until the user specifies a Policy Number, since the true default is based on
        ' whether the first digits of the Policy Number begin with UL, UV, UZ or 222.
        ' Note too that fnValidData will generate an error if the blank
        ' entry is still selected when the user clicks Update.
        
        ' Set .Row to -1 so insertion works okay whether or not the fpCombo is sorted
        .Row = gclngNoSelection
        .InsertRow = gcstrBlankEntry & vbTab & gcstrBlankEntry
    End With

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

    With rstTemp
        If .RecordCount <> 0 Then
            .MoveFirst
            Do Until .EOF
                ' Set .Row to -1 so insertion works okay whether or not the fpCombo is sorted
                lpcPycoTypCd.Row = gclngNoSelection
                lpcPycoTypCd.InsertRow = .Fields(mcstrPycoTypDsc).value & vbTab & .Fields(mcstrPycoTypCd).value
                .MoveNext
            Loop
        Else
            gerhApp.ReportNonFatal vbObjectError + gcRES_WARN_CBO_IS_EMPTY, _
                                   mstrScreenName & gcstrDOT & cstrCurrentProc, _
                                   mcstrLpcPycoTypCdLabel
        
        End If
    End With
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0

    ' Clean-up statements go here
    fnFreeRecordset rstTemp
    fnFreeObject adwTemp
    fnFreeObject prmReturnValue

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
Private Sub fnLoadRecordWithCalculatedControls()
    ' Comments  : Populates DB record with data from screen controls
    '             that are calculated
    ' Parameters: None
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnLoadRecordWithCalculatedControls"
 
    If Not mbInAddMode Then
        ' Extra precaution...to always calc totals across Payees
        ' before doing a save. This will also ensure the totals
        ' are 0 for an Add.  Can't call this prodedure on an Add
        ' since there is no current record and it will get a
        ' ADO 3021 error: "Either BOF or EOF is true or the current
        ' record has been deleted. Requested operation requires a
        ' current record."
        fnCalcTotalsForAllPayees mtWrapper.ClmId
    End If

    With mtWrapper
        ' The following fields cannot be edited by the user but are calculated
        ' by the program
        .ClmTotDthbPmtAmt = ipcClmTotDthbPmtAmt.UnFmtText
        .ClmTotIntAmt = ipcClmTotIntAmt.UnFmtText
        .ClmTotWthldAmt = ipcClmTotWthldAmt.UnFmtText
        .ClmTotClmPdAmt = ipcClmTotClmPdAmt.UnFmtText
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
Private Sub fnSetCompactFillingCheckBox(strAdminSystem As String, strState As String)
    ' Comments  : Retrieves the state record for the current claim record
    ' Parameters: N/A
    ' Called by : lpcAdmnSystCd_Change
    '             fnLoadControls
    ' Modified  : Berry Kropiwka 11-06-2019
    '
    ' --------------------------------------------------
    Dim strSQL                As String
    Dim rstTemp               As ADODB.Recordset
    Dim cstrCurrentProc       As String
    cstrCurrentProc = "fnSetCompactFillingCheckBox"
    On Error GoTo PROC_ERR
    If strAdminSystem = "SOLAR" Then 'Or strAdminSystem = "LEVERAGE" Then
        'Now check to make sure the claim is in a state that allows Compact Filling
        strSQL = "SELECT st_compact_clcn_allow_ind from dbo.state_t WHERE st_cd = '" & strState & "'"
        Set rstTemp = gadwApp.Execute_SQL_AsRST(gconAppActive, strSQL)
        If rstTemp.RecordCount > 0 Then
            If rstTemp!st_compact_clcn_allow_ind = "T" Then
                Me.chkClmCmpCalInd.Enabled = True
            Else
                Me.chkClmCmpCalInd.Enabled = False
                Me.chkClmCmpCalInd.value = vbUnchecked
            End If
        Else
            Me.chkClmCmpCalInd.Enabled = False
            Me.chkClmCmpCalInd.value = vbUnchecked
        End If
        rstTemp.ActiveConnection = Nothing
    Else
        If strAdminSystem = "" And strState = "" Then
            Me.chkClmCmpCalInd.Enabled = True
        Else
            Me.chkClmCmpCalInd.Enabled = False
            Me.chkClmCmpCalInd.value = vbUnchecked
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
        .ColFromName = mcstrClmId
        
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
Private Sub fnRefreshAllCombos()
    '--------------------------------------------------------------------------
    ' Procedure:   fnRefreshAllCombos
    ' Description: Repopulates each ComboBox or VSFlexGrid control
    '              so they reflect this and other users' changes. This proc
    '              should be called after each Add, Update or Delete.
    '
    ' Params:      N/A
    ' Called by:   cmdUpdate_Click of frmFund
    '              cmdDelete_Click of frmFund
    '              Form_Load of frmFund
    '
    ' Returns:     N/A
    '-----------------------------------------------------------------------------
    '!CUSTOMIZE!    This should call a function to load each ComboBox or
    '               VSFlexGrid control on the form. This will ensure that
    '               when one is refreshed (i.e. to make this and other
    '               user's changes visible), *all* will be.
    Const cstrCurrentProc       As String = "fnRefreshAllCombos"
    On Error GoTo PROC_ERR

    fnLoadLpcLookup lpcLookupClaim, elt_Claim   ' #1 = Claim Number (CLM_NUM, CLM_ID)
    fnLoadLpcLookup lpcLookupName, elt_Name     ' #2 = Insured Name (CLM_INSD_LAST_NM, CLM_INSD_FIRST_NM, CLM_NUM, CLM_ID)
    fnLoadLpcLookup lpcLookupSSN, elt_SSN       ' #3 = Insured SSN (CLM_INSD_SSN_NUM, CLM_ID)
    fnLoadLpcAdmnSystCd                         ' #4 = Admin System (ADMN_SYST_DSC, ADMN_SYST_CD)
    fnLoadLpcPycoTypCd                          ' #5 = Payor Company Type (PYCO_TYP_DSC, PYCO_TYP_CD)
    fnLoadCboIssStCd                            ' #6 = Issue State (ISS_ST_CD)
    fnLoadCboInsdDthResStCd                     ' #7 = Insured State of Residence at time of Death (INSD_DTH_RES_ST_CD)
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
    '             calling this function, e.g., they accurately reflect whether
    '             or not there are edits outstanding and/or the user is in
    '             Add mode, respectively.
    '             Remember, though: mbInAddMode and IsDirty are
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
    '    & #Children = 0
    '
    '    No edits o/s   enabled   disabled disabled enabled   enabled   enabled
    '    & #Children > 0
    '
    ' Called by : fnAddRecord and fnInitializeEditMode, with bEnable = False
    '
    '             lpcLookupClaim_Click, lpcLookupName_Click, lpcLookupSSN_Click,
    '             cmdDelete_Click, cmdNavigate_Click, cmdUpdate_Click
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

    ' Can only go to the Payees or Print an Individual Report when you're NOT in the middle of
    ' an Add or Update. It doesn't matter whether you have Payees though!
    If (Not IsDirty) And (Not mbInAddMode) Then
        cmdAddPayee.Enabled = True
        cmdPrintReport.Enabled = True
    Else
        cmdAddPayee.Enabled = False
        cmdPrintReport.Enabled = False
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
    
    ' Disallow future-dated Date of Death or Date of Proof
    dtpClmInsdDthDt.MaxDate = Now
    dtpClmProofDt.MaxDate = Now
    
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
        Case Else
            gerhApp.SaveErrObjectData mstrScreenName & gcstrDOT & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub


'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub fnSetInsdDthResStCdAvailability()
    '--------------------------------------------------------------------------
    ' Procedure:   fnSetInsdDthResStCdAvailability
    ' Description: Sets the availability of the InsdDthResStCd based on
    '              whether the Foreign Residence at Death checkbox is selected.
    '
    ' Params:      n/a
    ' Called by:   Form_Load of frmInsured
    '              fnLoadControls of frmInsured
    ' Returns:     n/a
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc       As String = "fnSetInsdDthResStCdAvailability"
    On Error GoTo PROC_ERR
 
    If chkClmForResDthInd.value = vbChecked Then
        cboInsdDthResStCd.Text = gcstrBlankEntry
        lblInsdDthResStCd.Enabled = False
        cboInsdDthResStCd.Enabled = False
    Else
        lblInsdDthResStCd.Enabled = True
        cboInsdDthResStCd.Enabled = True
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
    '              lpcLookupClaim_Click( )
    '              lpcLookupName_Click( )
    '              lpcLookupSSN_Click( )
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
Private Sub fnSetPropertiesForPayeeScreen(ByVal bSendEmptyName As Boolean)
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
 
    With msgPayees
        ' Note: If there are no Payees, msgPayees.Row will be set to 0 (the header row)
        .Col = 1   ' Payee Name column (2nd column, current row)

        If bSendEmptyName Then
            InsuredCurrentPayeeName = vbNullString
            InsuredCurrentPayeeID = 0
        Else
            InsuredCurrentPayeeName = .Text
            ' Get Payee ID from same row, different column
            .Col = 13
            InsuredCurrentPayeeID = .Text
        End If
    End With

    With mtWrapper
        InsuredClmID = .ClmId
        InsuredClmForResDthInd = .ClmForResDthInd
        InsuredClmInsdDthDt = .ClmInsdDthDt
        InsuredClmNum = .ClmNum
        InsuredClmProofDt = .ClmProofDt
        InsuredInsdDthResStCd = .InsdDthResStCd
        InsuredIssStCd = .IssStCd
        InsuredLobCd = fnGetLobCd()
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
Private Sub fnSetTxtClmNum()
    '--------------------------------------------------------------------------
    ' Procedure:   fnSetTxtClmNum
    ' Description: This procedure sets the hidden txtClmNum control, based
    '              on AdmnSystCd, ClmNum and, for Group, ClmInsdSSNNum
    '
    ' Params:      N/A
    ' Returns:     N/A
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc           As String = "fnSetTxtClmNum"
    On Error GoTo PROC_ERR
 
    With txtClmNum
        If fnGetLobCd() = mcstrGroupLOB Then
            .Text = iptClmPolNum.Text & mcstrGroupLOB & ipmClmInsdSsnNum.UnFmtText
        Else
            .Text = iptClmPolNum.Text
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
    fnEnableDisableControl ctlIn:=ipcClmTotDthbPmtAmt, bEnable:=False
    fnEnableDisableControl ctlIn:=ipcClmTotIntAmt, bEnable:=False
    fnEnableDisableControl ctlIn:=ipcClmTotWthldAmt, bEnable:=False
    fnEnableDisableControl ctlIn:=ipcClmTotClmPdAmt, bEnable:=False
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
    Const cintClmInsdSsnNumMinLgth          As Integer = 9
    Dim bErrorFound                         As Boolean
    Dim ctl                                 As Control
    Dim ctlFirstToFail                      As Control
    Dim intFailures                         As Integer
    Dim strFieldList                        As String
    Dim strMsgText                          As String
    Dim intLengthToTest                     As Integer
    Dim strLOB                              As String
    Dim strDefaultPycoTypDsc                As String
 
    fnValidData = True

    ' Check the fields in a left-to-right, top-to-bottom screen sequence.
    '     1. cboAdmnSystCd         7. chkClmForResDthInd
    '     2. iptClmPolNum          8. cboInsdDthResStCd
    '     3. lpcPycoTypCd          9. dtpInsdDthDt
    '     4. iptClmInsdFirstNm    10. dtpClmProofDt
    '     5. iptClmInsdLastNm     11. ipmClmInsdSsnNum
    '     6. cboIssStCd

    ' ------------- First, verify required fields are missing --------------
    
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
            ' Skip over the control that is bound to CLM_NUM since this is a hidden
            ' field and thus the user shouldn't be informed if it hasn't been set yet.
            If Len(.Tag) > 0 And (.Tag <> "ClmNum") Then
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
                        If (Len(ctl) = 0) Or (ctl = gcstrBlankEntry) Then
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
        End With
    Next ctl

    ' Check the Issue State, which is a nullable column. This may not have been input
    ' if the claim was entered prior to 2002 or so. However, if the user modifies the
    ' claim once the backend has been ported to SQL Server, or on new claims, then they **will** have to supply it.
    If (LenB(cboIssStCd.Text) = 0) Or (cboIssStCd.Text = gcstrBlankEntry) Then
        If intFailures = 0 Then
            strFieldList = vbCrLf & mcstrCboIssStCdLabel
            Set ctlFirstToFail = cboIssStCd
        Else
            strFieldList = strFieldList & vbCrLf & mcstrCboIssStCdLabel
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

    intLengthToTest = Len(iptClmPolNum.Text)
    ' Min/Max lengths were set upon each change to the Admin System control.
    If (intLengthToTest < mintAdmnSyst_MinPolNumLength) Or (intLengthToTest > mintAdmnSyst_MaxPolNumLength) Then
        intFailures = intFailures + 1
        Set ctlFirstToFail = iptClmPolNum
        strMsgText = strMsgText & vbCrLf & _
            "For the selected Admin System, the " & mcstrIptClmPolNumLabel & " must be between " & mintAdmnSyst_MinPolNumLength & _
            " and " & mintAdmnSyst_MaxPolNumLength & " characters long, inclusive."
    End If

    ' If the CLM_NUM (the logical key to this table) has changed, verify it has not been changed to
    ' one that already exists.
    'If txtClmNum.Text <> mtWrapper.ClmNum Then
        

    ' Verify the Payor Company Type is set to an appropriate default value
    strDefaultPycoTypDsc = fnGetDefaultPayorCompany(iptClmPolNum.Text, lpcAdmnSystCd.ColText)
    
    If InStr(gapsApp.LastLogonEnvironment, "Sun") = 1 Then
        If lpcPycoTypCd.Text <> strDefaultPycoTypDsc Then
            If Not (lpcAdmnSystCd.ColText = mcstrAdmnSystSOLAR And lpcPycoTypCd.Text = mcstrPyco_SLHIC And strDefaultPycoTypDsc <> mcstrPyco_Subsidiary) Then
            'report the error
                     intFailures = intFailures + 1
                     Set ctlFirstToFail = lpcPycoTypCd
                     strMsgText = strMsgText & vbCrLf & _
                                  "The selected " & mcstrLpcPycoTypCdLabel & " is invalid for this " & _
                                  mcstrLpcAdmnSystCdLabel & " or " & mcstrIptClmPolNumLabel & "."
            End If
        End If
    Else
        If lpcPycoTypCd.Text <> strDefaultPycoTypDsc Then
             intFailures = intFailures + 1
             Set ctlFirstToFail = lpcPycoTypCd
             strMsgText = strMsgText & vbCrLf & _
                          "The selected " & mcstrLpcPycoTypCdLabel & " is invalid for this " & _
                          mcstrLpcAdmnSystCdLabel & " or " & mcstrIptClmPolNumLabel & "."
        End If
    End If
    
    ' Make sure the user selected a non-blank entry in the Residence State & Issue State controls.
    If chkClmForResDthInd.value = vbUnchecked Then
        If (LenB(cboInsdDthResStCd.Text) = 0) Or (cboInsdDthResStCd.Text = gcstrBlankEntry) Then
            intFailures = intFailures + 1
            Set ctlFirstToFail = cboInsdDthResStCd
            strMsgText = strMsgText & vbCrLf & _
                         "Unless " & mcstrChkClmForResDthIndLabel & " is selected, the " & mcstrCboInsdDthResStCdLabel & _
                         " must be supplied."
        End If
    End If

    ' Verify the Date of Proof is on or after the Date of Death
    If DateValue(dtpClmProofDt.value) < DateValue(dtpClmInsdDthDt.value) Then
        intFailures = intFailures + 1
        Set ctlFirstToFail = dtpClmProofDt
        strMsgText = strMsgText & vbCrLf & _
                     "The " & mcstrDtpClmProofDtLabel & " (" & dtpClmProofDt.value & _
                     ") must be on or after the " & mcstrDtpClmInsdDthDtLabel & " (" & _
                     dtpClmInsdDthDt.value & ")."
    End If

    ' Determine whether any Payees exist with a Date Of Payment earlier than the
    ' Insured's Date of PROOF.  Skip if in Add mode, since there would be no Payees and
    ' the ClmId would be invalid.
    If Not mbInAddMode Then
        If fnGetPayeesNeedingRecalcDueToProof(mtWrapper.ClmId, mtWrapper.ClmProofDt) > 0 Then
            intFailures = intFailures + 1
            Set ctlFirstToFail = dtpClmProofDt
            strMsgText = strMsgText & vbCrLf & _
                         "One or more Payees exist with a Date Of Payment " & _
                         "earlier than the " & mcstrDtpClmProofDtLabel & "."
        End If

        ' Determine whether any Payees exist with a Date Of Payment earlier than the
        ' Insured's Date of DEATH.
        If fnGetPayeesNeedingRecalcDueToDeath(mtWrapper.ClmId, mtWrapper.ClmInsdDthDt) > 0 Then
            intFailures = intFailures + 1
            Set ctlFirstToFail = dtpClmInsdDthDt
            strMsgText = strMsgText & vbCrLf & _
                        "One or more Payees exist with a Date Of Payment " & _
                        "earlier than the " & mcstrDtpClmInsdDthDtLabel & "."
        End If
    End If

    ' Verify that a 9-character SSN was input, if anything was input to that field
    intLengthToTest = Len(ipmClmInsdSsnNum.UnFmtText)
    If intLengthToTest <> 0 And intLengthToTest <> cintClmInsdSsnNumMinLgth Then
        intFailures = intFailures + 1
        Set ctlFirstToFail = iptClmPolNum
        strMsgText = strMsgText & vbCrLf & _
                     "If input, the " & mcstrIpmClmInsdSsnNumLabel & " must be " & CStr(cintClmInsdSsnNumMinLgth) & " characters long."
    End If

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
    fnFreeObject ctlFirstToFail

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
 
    If (Not mbInAddMode) Then
        If (mrstPayees.RecordCount > 0) Then
            If DateValue(mstrOrigDateOfDeath) <> DateValue(dtpClmInsdDthDt.value) Then
                ' 1008 = The @@1 has changed. This change may affect the calculations for existing Payees. Please review and, if necessary, recalculate each Payee.
                gerhApp.ReportNonFatal vbObjectError + gcRES_INFO_DT_CHG_MAY_AFFECT_PAYEES, _
                    mstrScreenName & gcstrDOT & cstrCurrentProc, mcstrDtpClmInsdDthDtLabel
            End If
            If DateValue(mstrOrigDateOfProof) <> DateValue(dtpClmProofDt.value) Then
                ' 1008 = The @@1 has changed. This change may affect the calculations for existing Payees. Please review and, if necessary, recalculate each Payee.
                gerhApp.ReportNonFatal vbObjectError + gcRES_INFO_DT_CHG_MAY_AFFECT_PAYEES, _
                    mstrScreenName & gcstrDOT & cstrCurrentProc, mcstrDtpClmProofDtLabel
            End If
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
        Case 5  ' Invalid procedure call or argument
            ' Caused by setting the focus to a field that's not yet visible
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
    
    '...............................................................................
    ' Set our fpCombo control settings, for those used as Lookups. Since these
    ' contain lots of rows (5000 or so, currently), they are loaded with sorted data
    ' rather than having the control itself sort its contents. This GREATLY improves
    ' the time it takes to display the form and refresh the control!
    '...............................................................................
    ' 1. Claim Lookup
    With lpcLookupClaim
        fnInitializefpCombo lpcIn:=lpcLookupClaim, bShowColHeaders:=False, bSortable:=False, _
            lngNbrOfCols:=2, lngEditCol:=mcintDisplayCol_lpcLookupClaim, lngNbrOfRowsInDropdown:=8
        ' Column definitions
        .Col = 0                                            ' First column, Primary sort
            .ColHeaderText = mcstrTxtClmNumLabel
            .ColName = mcstrDisplayCol
            .ColWidth = 20
        .Col = 1                                            ' Second column
            .ColHeaderText = mcstrTxtClmIDLabel
            .ColName = mcstrClmId
            .ColHide = True
        .ColumnSearch = mcintDisplayCol_lpcLookupClaim
    End With
    ' 2. Name Lookup
    With lpcLookupName
        fnInitializefpCombo lpcIn:=lpcLookupName, bShowColHeaders:=False, bSortable:=False, _
            lngNbrOfCols:=4, lngEditCol:=mcintDisplayCol_lpcLookupName, lngNbrOfRowsInDropdown:=8
        ' Since there are multiple visible columns, show lines on this one
        .ListApplyTo = ListApplyToAllCols
        .LineStyle = LineStyleLowered
        ' Column definitions
        .Col = 0                                            ' First column, Primary sort
            .ColHeaderText = mcstrIptClmInsdLastNmLabel
            .ColName = mcstrDisplayCol
        .Col = 1                                            ' Second column, first Secondary sort
            .ColHeaderText = mcstrIptClmInsdFirstNmLabel
            .ColName = mcstrClmInsdFirstNm
        .Col = 2                                            ' Third column, second Secondary sort
            .ColHeaderText = mcstrTxtClmNumLabel
            .ColName = mcstrClmNum
            .ColWidth = 20
        .Col = 3                                            ' Fourth column
            .ColHeaderText = mcstrTxtClmIDLabel
            .ColName = mcstrClmId
            .ColHide = True
        .ColumnSearch = mcintDisplayCol_lpcLookupName
    End With
    ' 3. SSN Lookup
    With lpcLookupSSN
        fnInitializefpCombo lpcIn:=lpcLookupSSN, bShowColHeaders:=False, bSortable:=False, _
            lngNbrOfCols:=3, lngEditCol:=mcintDisplayCol_lpcLookupSSN, lngNbrOfRowsInDropdown:=8
        ' Since there are multiple visible columns, show lines on this one
        .ListApplyTo = ListApplyToAllCols
        .LineStyle = LineStyleLowered
         ' Column definitions
        .Col = 0                                            ' First column, Primary sort
            .ColHeaderText = mcstrIpmClmInsdSsnNumLabel
            .ColName = mcstrDisplayCol
        .Col = 1                                            ' Second column, second Secondary sort
            .ColHeaderText = mcstrTxtClmNumLabel
            .ColName = mcstrClmNum
            .ColWidth = 20
        .Col = 2                                            ' Third column
            .ColHeaderText = mcstrTxtClmIDLabel
            .ColName = mcstrClmId
            .ColHide = True
        .ColumnSearch = mcintDisplayCol_lpcLookupSSN
    End With

    '...............................................................................
    ' Set our fpCombo control settings, for those used as multi-column comboboxes.
    '...............................................................................
    ' 1. ADMN_SYST_CD
    With lpcAdmnSystCd
        fnInitializefpCombo lpcIn:=lpcAdmnSystCd, bShowColHeaders:=False, bSortable:=True, _
            lngNbrOfCols:=2, lngEditCol:=mcintDisplayCol_lpcAdmnSystCd, lngNbrOfRowsInDropdown:=8
        ' Column definitions
        .Col = mcintDisplayCol_lpcAdmnSystCd               ' 1st column (description), Primary sort
            .ColName = mcstrAdmnSystDsc
            .ColSortSeq = 0
            .ColSorted = SortedAscending
        .Col = mcintStoreCol_lpcAdmnSystCd                 ' 2nd column (code)
            .ColName = mcstrAdmnSystCd
            .ColHide = True
    End With
    ' 2. PYCO_TYP_CD
    With lpcPycoTypCd
        fnInitializefpCombo lpcIn:=lpcPycoTypCd, bShowColHeaders:=False, bSortable:=True, _
            lngNbrOfCols:=2, lngEditCol:=mcintDisplayCol_lpcPycoTypCd, lngNbrOfRowsInDropdown:=8
        ' Column definitions
        .Col = mcintDisplayCol_lpcPycoTypCd               ' 1st column (description), Primary sort
            .ColName = mcstrPycoTypDsc
            .ColSortSeq = 0
            .ColSorted = SortedAscending
        .Col = mcintStoreCol_lpcPycoTypCd                 ' 2nd column (code)
            .ColName = mcstrPycoTypCd
            .ColHide = True
    End With
    
    ' Set the control to receive the focus after errors (the first editable field
    ' on the screen), dependent upon whether we're in Add Mode or not. If in Add mode,
    ' this control would typically be the first control that corresponds to a Key field.
    ' If not in Add mode, this control would typically be the topmost/leftmost
    ' "always updateable" control on the screen (excepting the Lookup ComboBox).
    Set mctlFirstUpdateableField_Add = lpcAdmnSystCd
    Set mctlFirstUpdateableField_Upd = lpcAdmnSystCd

    ' NOTE: The next IF block probably isn't necessary now that the Insured
    ' screen is no longer automatically displayed after the user initially
    ' logs on.
    ' Allow the progress meter on the splash screen to get updated
    If fnIsFormLoaded("frmSplash") Then
        DoEvents
    End If

    ' Instantiate and initialize a table wrapper object for the appropriate table(s).
    Set mtWrapper = New ctclmClaim
    ' Instantiate and initialize a table wrapper object for the Payee table. This will be used
    ' to get data associated with the current claim.
    Set mtPayee = New ctpyePayee
    
    ' Bind the on-screen controls to the table wrapper class properties with which they
    ' are associated. set default settings for those controls' properties, and
    ' bind editable TextBoxes controls to the Extended TextBox class so they will
    ' behave appropriately and in a consistent manner.
    fnSetupScreenControls
    
    ' Populate all ComboBoxes and ListPro controls
    fnRefreshAllCombos

    ' Always go into Add mode, per the user, to ensure they don't inadvertently start
    ' editing that first record.
    fnAddRecord

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
            mtWrapper.GetLookupData
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
                mtWrapper.GetRelativeRecord mtWrapper.ClmNum, epdSameRecord
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
Private Sub Form_Resize()
    ' Comments  : Resize the form
    ' Parameters: N/A
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc                               As String = "Form_Resize"
    Const cintCmdAddPayees_Orig_FormHeightLessBtnTop    As Integer = 940  '855
    Const cintFraPayees_Orig_FormWidthLessFrameWidth    As Integer = 330
    Const cintFraPayees_Orig_Height                     As Integer = 1980
    Const cintSpacerBorderAroundAllEdgesOfForm          As Integer = 15
    Const cintMsgPayees_Orig_Width                      As Integer = 11925
    Const cintMsgPayees_Orig_Height                     As Integer = 1515
    Const cintLblGridInstructions_Orig_Left             As Integer = 7320

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName
    
    If Me.WindowState = vbNormal And Me.Visible Then
        ' Bypass if vbMinimized or vbMaximized, to avoid run-time error 384
        ' which says" "a form can't be moved or sized while minimized or maximized"
        If (Me.Height < mclngMinFormHeight) Then
            Me.Height = mclngMinFormHeight
        End If
        If (Me.Width < mclngMinFormWidth) Then
            Me.Width = mclngMinFormWidth
        End If

        cmdAddPayee.Left = (Me.Width - cmdAddPayee.Width) / 2
        cmdAddPayee.Top = Me.Height - cintCmdAddPayees_Orig_FormHeightLessBtnTop
        fraPayees.Width = Me.Width - cintFraPayees_Orig_FormWidthLessFrameWidth
        fraPayees.Height = cintFraPayees_Orig_Height + _
            Me.Height - (mclngMinFormHeight + cintSpacerBorderAroundAllEdgesOfForm)
        msgPayees.Width = cintMsgPayees_Orig_Width + _
            Me.Width - (mclngMinFormWidth + (cintSpacerBorderAroundAllEdgesOfForm * 2))
        msgPayees.Height = cintMsgPayees_Orig_Height + _
            Me.Height - (mclngMinFormHeight + cintSpacerBorderAroundAllEdgesOfForm)
        lblGridInstructions.Left = Me.Width - _
            (cintLblGridInstructions_Orig_Left - (cintSpacerBorderAroundAllEdgesOfForm * 2))
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

    If bDebugAppTermination Then
        Debug.Print "Entering " & mstrScreenName & gcstrDOT & cstrCurrentProc
    End If

    gapsApp.SaveForm Me

    IsDirty = False

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
Private Sub ipmClmInsdSsnNum_Change()
    ' Comments  : Sets a flag to indicate the current record has been
    '             edited, and thus Update button becomes enabled
    ' Parameters: N/A
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fpmClmInsdSsnNum_Change"
 
    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    ' Ensure availability of navigation & command buttons is set appropriately
    fnInitializeEditMode
    
    ' Set the hidden Claim Number field.
    fnSetTxtClmNum
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
Private Sub iptClmInsdFirstNm_Change()
    ' Comments  : Sets a flag to indicate the current record has been
    '             edited, and thus Update button becomes enabled
    ' Parameters: N/A
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fptClmInsdFirstNm_Change"
 
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
Private Sub iptClmInsdLastNm_Change()
    ' Comments  : Sets a flag to indicate the current record has been
    '             edited, and thus Update button becomes enabled
    ' Parameters: N/A
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fptClmInsdFirstNm_Change"
 
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
Private Sub iptClmPolNum_Change()
    ' Comments  : Sets a flag to indicate the current record has been
    '             edited, and thus Update button becomes enabled
    ' Parameters: N/A
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc       As String = "fptClmPolNum_Change"
    Dim strDefaultPycoTypDsc    As String
 
    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    ' Ensure availability of navigation & command buttons is set appropriately
    fnInitializeEditMode
    
    ' Use the store column (ADMN_SYST_CD) column of the Admin System combobox,
    ' then determine if the Company Type should change based on the new input
    lpcAdmnSystCd.Col = mcintStoreCol_lpcAdmnSystCd
    
    strDefaultPycoTypDsc = fnGetDefaultPayorCompany(iptClmPolNum.Text, lpcAdmnSystCd.ColText)
    fnSearchFPCombo lpcPycoTypCd, strDefaultPycoTypDsc, mcintDisplayCol_lpcPycoTypCd
    
    ' Set the hidden Claim Number field.
    fnSetTxtClmNum
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
Private Sub lpcAdmnSystCd_Change()
    ' Comments  : Sets a flag to indicate the current record has been
    '             edited, and thus Update button becomes enabled. It also
    '             dictates how long the Policy Number (ClmPolNum) can be, given
    '             the Admin System chosen.
    ' Parameters: N/A
    ' Modified  : Berry Kropiwka - 2019-11-06 - Added code to enable or disable the Compact Filling check box bases on admin system and state
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc       As String = "lpcAdmnSystCd_Change"
    Dim strDefaultPycoTypDsc    As String
 
    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    fnInitializeEditMode
    
    ' Get metadata (min/max policy # allowed, tax rptg ind, default pyco type code) and store in module-level variables
    fnGetAdminSysMetadata
    
    With iptClmPolNum
        ' There is no .MinLength property on this control  :-)
        .MaxLength = mintAdmnSyst_MaxPolNumLength
    End With
    
    ' Use the store column (ADMN_SYST_CD) column of the Admin System combobox,
    ' then determine if the Company Type should change based on the new input
    lpcAdmnSystCd.Col = mcintStoreCol_lpcAdmnSystCd
    
    strDefaultPycoTypDsc = fnGetDefaultPayorCompany(iptClmPolNum.Text, lpcAdmnSystCd.ColText)
    fnSearchFPCombo lpcPycoTypCd, strDefaultPycoTypDsc, mcintDisplayCol_lpcPycoTypCd
    
    ' Set the hidden Claim Number field.
    fnSetTxtClmNum
    
    ' enable or disable the Compact Filling check box based on Admin System
    fnSetCompactFillingCheckBox Me.lpcAdmnSystCd.Text, Me.cboIssStCd.Text

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
Private Sub lpcAdmnSystCd_GotFocus()
    '-----------------------------------------------------------------------------
    ' Function     lpcAdmnSystCd_GotFocus
    ' Purpose      Display the drop down list now that the user has entered this control.
    ' Parameters   N/A
    ' Returns      N/A
    ' Date:        12/19/2001
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc          As String = "lpcAdmnSystCd_GotFocus"
    On Error GoTo PROC_ERR

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    'lpcAdmnSystCd.ListDown = True
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
Private Sub lpcLookupClaim_Click()
    ' Comments  : Retrieve selected record
    ' Parameters: N/A
    '
    ' --------------------------------------------------
    Const cstrCurrentProc               As String = "lpcLookupClaim_Click"

    On Error GoTo PROC_ERR

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    fnPerformLookup lpcLookupClaim
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
Private Sub lpcLookupClaim_GotFocus()
    '-----------------------------------------------------------------------------
    ' Function     lpcLookupClaim_GotFocus
    ' Purpose      Turn on Lookup Mode and drop down the list now that the user has entered this control.
    ' Parameters   N/A
    ' Returns      N/A
    ' Date:        12/19/2001
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc          As String = "lpcLookupClaim_GotFocus"
    On Error GoTo PROC_ERR

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    'lpcLookupClaim.ListDown = True

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
Private Sub lpcLookupClaim_KeyDown(ByRef intKeyCode As Integer, ByRef intShift As Integer)
    '-----------------------------------------------------------------------------
    ' Function     lpcLookupClaim_KeyDown
    ' Purpose      If the user presses Enter, make it do just what the Click event does
    '              (i.e. display the selected record)
    ' Parameters   intKeyCode - ASCII code of key that was pressed
    '              intShift - indicates whether the Shift key was pressed
    ' Returns      N/A
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc          As String = "lpcLookupClaim_KeyDown"
    On Error GoTo PROC_ERR

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    If intKeyCode = vbKeyReturn Then
        fnPerformLookup lpcLookupClaim
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
Private Sub lpcLookupClaim_LostFocus()
    '-----------------------------------------------------------------------------
    ' Function     lpcLookupClaim_LostFocus
    ' Purpose      Turn off Lookup Mode now that the user has left that control.
    ' Parameters   N/A
    ' Returns      N/A
    ' Date:        12/19/2001
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc          As String = "lpcLookupClaim_LostFocus"
    Const clngFirstRow             As Long = 0
    On Error GoTo PROC_ERR

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    ' Display the first (blank) entry in the Lookup control so the
    ' user doesn't get confused. Without this code, the Lookup box continues to display
    ' the value last selected for lookup purposes, even when the user has since positioned
    ' to a different record by virtue of doing a Delete or Add or using the navigation buttons.
    With lpcLookupClaim
        .Row = clngFirstRow
        .ListIndex = clngFirstRow
        .Action = ActionClearSearchBuffer
    End With

    'fnSearchFPCombo lpcLookupClaim, gcstrBlankEntry, mcintDisplayCol_lpcLookupClaim
    lpcLookupClaim.Refresh

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
Private Sub lpcLookupName_Click()
    ' Comments  : Retrieve selected record
    ' Parameters: N/A
    ' Modified  : CMP 4/27/2002
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
    Const cstrCurrentProc          As String = "lpcLookupClaim_KeyDown"
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
Private Sub lpcLookupSSN_Click()
    ' Comments  : Retrieve selected record
    ' Parameters: N/A
    ' Modified  :
    '
    ' --------------------------------------------------
    Const cstrCurrentProc               As String = "lpcLookupSSN_Click"

    On Error GoTo PROC_ERR

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    fnPerformLookup lpcLookupSSN
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
Private Sub lpcLookupSSN_GotFocus()
    '-----------------------------------------------------------------------------
    ' Function     lpcLookupSSN_GotFocus
    ' Purpose      Turn on Lookup Mode and drop down the list now that the user has entered this control.
    ' Parameters   N/A
    ' Returns      N/A
    ' Date:        12/19/2001
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc          As String = "lpcLookupSSN_GotFocus"
    On Error GoTo PROC_ERR

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    'lpcLookupSSN.ListDown = True
    
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
Private Sub lpcLookupSSN_KeyDown(ByRef intKeyCode As Integer, ByRef intShift As Integer)
    '-----------------------------------------------------------------------------
    ' Function     lpcLookupSSN_KeyDown
    ' Purpose      If the user presses Enter, make it do just what the Click event does
    '              (i.e. display the selected record)
    ' Parameters   intKeyCode - ASCII code of key that was pressed
    '              intShift - indicates whether the Shift key was pressed
    ' Returns      N/A
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc          As String = "lpcLookupSSN_KeyDown"
    On Error GoTo PROC_ERR

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    If intKeyCode = vbKeyReturn Then
        fnPerformLookup lpcLookupSSN
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
Private Sub lpcLookupSSN_LostFocus()
    '-----------------------------------------------------------------------------
    ' Function     lpcLookupSSN_LostFocus
    ' Purpose      Turn off Lookup Mode now that the user has left that control.
    ' Parameters   N/A
    ' Returns      N/A
    ' Date:        12/19/2001
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc          As String = "lpcLookupSSN_LostFocus"
    Const clngFirstRow             As Long = 0
    On Error GoTo PROC_ERR

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    ' Display the first (blank) entry in the Lookup control so the
    ' user doesn't get confused. Without this code, the Lookup box continues to display
    ' the value last selected for lookup purposes, even when the user has since positioned
    ' to a different record by virtue of doing a Delete or Add or using the navigation buttons.
    With lpcLookupSSN
        .Row = clngFirstRow
        .ListIndex = clngFirstRow
        .Action = ActionClearSearchBuffer
    End With
    'fnSearchFPCombo lpcLookupSSN, gcstrBlankEntry, mcintDisplayCol_lpcLookupSSN
    lpcLookupSSN.Refresh

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
Private Sub lpcPycoTypCd_Change()
    ' Comments  : Sets a flag to indicate the current record has been
    '             edited, and thus Update button becomes enabled
    ' Parameters: N/A
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "lpcPycoTypCd_Change"
 
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
Private Sub lpcPycoTypCd_GotFocus()
    '-----------------------------------------------------------------------------
    ' Function     lpcPycoTypCd_GotFocus
    ' Purpose      Display the drop down list now that the user has entered this control.
    ' Parameters   N/A
    ' Returns      N/A
    ' Date:        12/19/2001
    '-----------------------------------------------------------------------------
    Const cstrCurrentProc          As String = "lpcPycoTypCd_GotFocus"
    On Error GoTo PROC_ERR

    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    'lpcPycoTypCd.ListDown = True
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
Private Sub msgPayees_DblClick()
        ' Comments  : This event handler is triggered when the user double-clicks in
        '             the Payee grid to indicate they want to edit that Payee
        ' Parameters:  -
        ' Modified  :
        ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc   As String = "msgPayee_DblClick"
    Dim frmChild            As Form
    Dim strSaveClaimNumber  As String
    Dim hrgHourglass        As chrgHourglass
    Dim lngReturnValue      As Long
    Dim strACF2             As String
    
    ' Set screen name in case errors are reported here or
    ' in procedures called by this Event Handler
    gerhApp.ScreenName = mstrScreenName

    ' Set the current column to 1 (the Payee Name)
    fnSetPropertiesForPayeeScreen bSendEmptyName:=False
    ' Following statement triggers the Form_Initialize & Form_Load events in frmPayee
    Set frmChild = New frmPayee
    ' Following statement triggers the Form_Activate event in frmPayee
    frmChild.Show vbModal
    
    ' Do DoEvents to allow the Payee screen to fully disappear.
    DoEvents
    
    fnWindowLock Me.hWnd

    Set hrgHourglass = New chrgHourglass
    hrgHourglass.value = True

    ' Update the Payees Recordset to reflect any Payees just added, changed or deleted
    ' when the Payees screen was open. Then, update the msgPayees grid and recalculate
    ' totals across all Payees.
    ' Note: You *must* requery the Insured and Payee recordsets to accomodate the possibility
    '       that another user (a) add/changed/deleted one more Payees for the
    '       current Insured and (b) returned to the Insured screen which triggered an update
    '       to the Insured record for the claim-wide totals it carries. If you don't do the
    '       requeries then a -2147217864 "row cannot be located for updating..." error could
    '       occur. So, we'll do the requerying automatically with no visible indication to the
    '       user that it occured unless the requerying revealed that another user deleted the
    '       current claim number and hence the Insured with the next higher claim number will
    '       be displayed (otherwise the same claim remains being displayed).
    strSaveClaimNumber = iptClmPolNum.Text
    hrgHourglass.value = False

'!TODO! The following looks like unnecessary (i.e. dead) code
'   If iptClmPolNum <> strSaveClaimNumber Then
'   '!TODO! Gen msg via frmMsgBox
'        'MsgBox "Another user has deleted the Claim Number (" & strSaveClaimNumber & ") you were viewing.", _
'        '       vbOKOnly + vbInformation, mcstrDialogTitle
'    End If
    
    hrgHourglass.value = True

    fnGetChildren
    
    ' 01/31/2001 BAW - Add another Refresh to speed up repainting
    Me.Refresh
    
    ' Totals may have changed. Update the Insured record just in case.
    fnLoadRecordWithCalculatedControls
    
    ' 01/31/2001 BAW - Add another Refresh to speed up repainting
    Me.Refresh
    
    With mtWrapper
        ' Determine whether another user updated or deleted the record about to be updated.
        ' Note: this multi-user checking is performed on an Update but not an Add.
        lngReturnValue = .CheckForAnotherUsersChanges(ewoUpdate, strACF2)

        If lngReturnValue = vbObjectError + gcRES_INFO_ANOTHER_USER_DELETED_NOT_SAVED Then
            gerhApp.ReportNonFatal vbObjectError + gcRES_INFO_ANOTHER_USER_DELETED_NOT_SAVED, _
                                       mstrScreenName & gcstrDOT & cstrCurrentProc
            ' Discard *this* user's pending changes and show the previous record.
            ' Can't use the GetClmNumFromClmId( ) method since the CLAIM_T row
            ' doesn't exist and hence a "-2147217900" (Claim ID does not exist) error
            ' throws things off.
            .GetRelativeRecord .ClmNum, epdPreviousRecord
        ' Do NOT bother to check for another UPDATING the record, since all we're doing is
        ' updating the total fields. Let the totals update go through.
        '   ElseIf lngReturnValue = vbObjectError + gcRES_INFO_ANOTHER_USER_UPDATED_DISCARDED Then
        '           gerhApp.ReportNonFatal vbObjectError + gcRES_INFO_ANOTHER_USER_UPDATED_DISCARDED, _
        '                                   mstrScreenName & gcstrDOT & cstrCurrentProc, _
        '                                   Trim$(strACF2)
        '       ' Discard *this* user's pending changes by re-retrieving the current record
        '       ' as it currently looks on the database and refreshing the lookup recordset.
        '       ' Can't use the GetClmNumFromClmId( ) method since the CLAIM_T row
        '       ' doesn't exist and hence a "-2147217900" (Claim ID does not exist) error
        '       ' throws things off.
        '       .GetRelativeRecord .ClmNum, epdSameRecord
        Else
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

    ' Have to call fnLoadControls here, like in cmdAdd_Click and cmdDelete_Click and cmdUpdate_Click,
    ' to ensure refreshed comboboxes have their previous value still selected.
    If mtWrapper.LookupRecordCount > 0 Then
        ' Ensure the on-screen controls reflect the record just added/updated, in case the
        ' DBMS altered it in some way, e.g., determining an Identity column value and
        ' getting the most up-to-date Last Updated info. This also sets the navigation
        ' buttons and updates the "record x of y" label
        fnLoadControls
        fnSetCommandButtons True
    Else
        fnAddRecord
    End If
PROC_EXIT:
    ' Disable the error handler so errors hit here won't be handled by PROC_ERR
    On Error GoTo 0

    ' Clean-up statements go here
    If Not (hrgHourglass Is Nothing) Then
        hrgHourglass.value = False
    End If
    fnFreeObject hrgHourglass
    ' Terminate the Payee form, removing it from the Forms collection
    fnFreeObject frmChild
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

