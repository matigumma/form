VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "edt32x30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "spr32x60.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "flp32a30.ocx"
Object = "{1A07B6A1-0856-4327-8539-6F46C49D49FA}#4.1#0"; "UC_BarraDeComandos.ocx"
Object = "{CA75FFD7-EB66-447F-9507-14B66B46F9A5}#3.3#0"; "UC_SeleccionEnLista.OCX"
Object = "{48932A52-981F-101B-A7FB-4A79242FD97B}#3.1#0"; "tab32x30.ocx"
Object = "{A18D4668-91EF-101C-84A6-BA990A365A4E}#3.0#0"; "mem32x30.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmPedidos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pedidos"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9870
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   9870
   Begin TabproLib.vaTabPro tabPedidos 
      Height          =   6255
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   8295
      _Version        =   196609
      _ExtentX        =   14631
      _ExtentY        =   11033
      _StockProps     =   100
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabsPerRow      =   7
      TabCount        =   7
      AlignTextH      =   1
      AlignTextV      =   1
      ActiveTabBold   =   0   'False
      AlignPictureV   =   1
      OffsetFromClientTop=   -1  'True
      DataFormat      =   ""
      BookCornerGuardWidth=   105
      BookCornerGuardLength=   390
      DataField       =   ""
      TabCaption      =   "Pedidos.frx":0000
      PageEarMarkPictureNext=   "Pedidos.frx":0500
      PageEarMarkPicturePrev=   "Pedidos.frx":051C
      EarMarkPictureNext=   "Pedidos.frx":0538
      EarMarkPicturePrev=   "Pedidos.frx":0554
      Begin TabproLib.vaTabPro tabDivision 
         Height          =   5295
         Left            =   -23175
         TabIndex        =   144
         Top             =   -20895
         Width           =   8055
         _Version        =   196609
         _ExtentX        =   14208
         _ExtentY        =   9340
         _StockProps     =   100
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabsPerRow      =   2
         TabCount        =   2
         AlignTextH      =   1
         OffsetFromClientTop=   -1  'True
         DataFormat      =   ""
         BookCornerGuardWidth=   105
         BookCornerGuardLength=   390
         DataField       =   ""
         TabCaption      =   "Pedidos.frx":0570
         PageEarMarkPictureNext=   "Pedidos.frx":083C
         PageEarMarkPicturePrev=   "Pedidos.frx":0858
         EarMarkPictureNext=   "Pedidos.frx":0874
         EarMarkPicturePrev=   "Pedidos.frx":0890
         Begin FPSpread.vaSpread sprDivisionPedido 
            Height          =   3615
            Left            =   240
            TabIndex        =   145
            Top             =   1320
            Width           =   7590
            _Version        =   393216
            _ExtentX        =   13388
            _ExtentY        =   6376
            _StockProps     =   64
            DisplayRowHeaders=   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            RetainSelBlock  =   0   'False
            ScrollBars      =   2
            SelectBlockOptions=   0
            SpreadDesigner  =   "Pedidos.frx":08AC
            UserResize      =   1
         End
         Begin FPSpread.vaSpread sprDivisionRemitos 
            Height          =   3615
            Left            =   -22830
            TabIndex        =   146
            Top             =   -19935
            Width           =   7590
            _Version        =   393216
            _ExtentX        =   13388
            _ExtentY        =   6376
            _StockProps     =   64
            Enabled         =   0   'False
            DisplayRowHeaders=   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            RetainSelBlock  =   0   'False
            ScrollBars      =   2
            SelectBlockOptions=   0
            SpreadDesigner  =   "Pedidos.frx":0B60
            UserResize      =   1
         End
         Begin UC_SeleccionEnLista.SeleccionEnLista selCuentasDivision 
            Height          =   315
            Left            =   240
            TabIndex        =   147
            Top             =   780
            Width           =   3285
            _ExtentX        =   5794
            _ExtentY        =   556
            PideAlias       =   -1  'True
            AnchoDelAlias   =   700
            ListIndex       =   0
         End
         Begin EditLib.fpDoubleSingle txtCantidadDivision 
            Height          =   315
            Left            =   3720
            TabIndex        =   148
            Top             =   780
            Width           =   1215
            _Version        =   196608
            _ExtentX        =   2143
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483637
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
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   2
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
            Text            =   "0,0000"
            DecimalPlaces   =   4
            DecimalPoint    =   ""
            FixedPoint      =   -1  'True
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
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin UC_BarraDeComandos.BarraDeComandos bdcDivision 
            Height          =   375
            Left            =   5280
            TabIndex        =   149
            Top             =   780
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   661
            BackColor       =   0
            BotoneraReducida=   -1  'True
            Enabled         =   -1  'True
            CaptionBoton(0) =   ""
            ImagenBoton(0)  =   "Pedidos.frx":0E14
            EstiloBoton(0)  =   -1  'True
            BotonDefault(0) =   0   'False
            BotonCancel(0)  =   0   'False
            CaptionBoton(1) =   ""
            ImagenBoton(1)  =   "Pedidos.frx":1256
            EstiloBoton(1)  =   -1  'True
            BotonDefault(1) =   0   'False
            BotonCancel(1)  =   0   'False
            CaptionBoton(2) =   ""
            ImagenBoton(2)  =   "Pedidos.frx":1698
            EstiloBoton(2)  =   -1  'True
            BotonDefault(2) =   0   'False
            BotonCancel(2)  =   0   'False
            CaptionBoton(3) =   ""
            ImagenBoton(3)  =   "Pedidos.frx":1ADA
            EstiloBoton(3)  =   -1  'True
            BotonDefault(3) =   0   'False
            BotonCancel(3)  =   0   'False
            CaptionBoton(4) =   ""
            EstiloBoton(4)  =   0   'False
            BotonDefault(4) =   0   'False
            BotonCancel(4)  =   0   'False
            CaptionBoton(5) =   ""
            EstiloBoton(5)  =   0   'False
            BotonDefault(5) =   0   'False
            BotonCancel(5)  =   0   'False
            CaptionBoton(6) =   ""
            EstiloBoton(6)  =   0   'False
            BotonDefault(6) =   0   'False
            BotonCancel(6)  =   0   'False
            CaptionBoton(7) =   ""
            EstiloBoton(7)  =   0   'False
            BotonDefault(7) =   0   'False
            BotonCancel(7)  =   0   'False
            CaptionBoton(8) =   ""
            EstiloBoton(8)  =   0   'False
            BotonDefault(8) =   0   'False
            BotonCancel(8)  =   0   'False
         End
         Begin VB.Label Label42 
            Caption         =   "&Cantidad"
            Height          =   255
            Left            =   3720
            TabIndex        =   151
            Top             =   465
            Width           =   735
         End
         Begin VB.Label Label43 
            Caption         =   "Cuenta"
            Height          =   255
            Left            =   240
            TabIndex        =   150
            Top             =   465
            Width           =   615
         End
      End
      Begin LpLib.fpList lstBusqueda 
         Height          =   3960
         Left            =   -15255
         TabIndex        =   167
         Top             =   -18960
         Visible         =   0   'False
         Width           =   255
         _Version        =   196608
         _ExtentX        =   450
         _ExtentY        =   6985
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         MousePointer    =   0
         Object.TabStop         =   0   'False
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Columns         =   0
         Sorted          =   0
         LineWidth       =   1
         SelDrawFocusRect=   -1  'True
         ColumnSeparatorChar=   9
         ColumnSearch    =   -1
         ColumnWidthScale=   2
         RowHeight       =   -1
         MultiSelect     =   0
         WrapList        =   0   'False
         WrapWidth       =   0
         SelMax          =   -1
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
         DataField       =   ""
         OLEDragMode     =   0
         OLEDropMode     =   0
         Redraw          =   -1  'True
         ResizeRowToFont =   0   'False
         TextTipMultiLine=   0
         ColDesigner     =   "Pedidos.frx":245C
      End
      Begin FPSpread.vaSpread sprDetalle 
         Height          =   2415
         Left            =   -22950
         TabIndex        =   27
         Top             =   -18015
         Width           =   7590
         _Version        =   393216
         _ExtentX        =   13388
         _ExtentY        =   4260
         _StockProps     =   64
         Enabled         =   0   'False
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         SpreadDesigner  =   "Pedidos.frx":26B0
         UserResize      =   1
      End
      Begin VB.Frame fraDetalle 
         Enabled         =   0   'False
         Height          =   2520
         Left            =   -22980
         TabIndex        =   28
         Top             =   -20460
         Width           =   7620
         Begin EditLib.fpText txtBusqueda 
            Height          =   315
            Left            =   1200
            TabIndex        =   168
            Top             =   225
            Visible         =   0   'False
            Width           =   5175
            _Version        =   196608
            _ExtentX        =   9128
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            Text            =   ""
            CharValidationText=   ""
            MaxLength       =   255
            MultiLine       =   0   'False
            PasswordChar    =   ""
            IncHoriz        =   0,25
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
         Begin VB.CommandButton cmdBusqueda 
            Height          =   315
            Left            =   7110
            Picture         =   "Pedidos.frx":2964
            Style           =   1  'Graphical
            TabIndex        =   166
            Top             =   210
            Width           =   375
         End
         Begin RichTextLib.RichTextBox txtDescripcion 
            Height          =   315
            Left            =   1185
            TabIndex        =   31
            Top             =   610
            Width           =   6300
            _ExtentX        =   11113
            _ExtentY        =   556
            _Version        =   393217
            BorderStyle     =   0
            MultiLine       =   0   'False
            ScrollBars      =   1
            MaxLength       =   500
            TextRTF         =   $"Pedidos.frx":2EEE
         End
         Begin EditLib.fpDoubleSingle txtPrecio 
            Height          =   315
            Left            =   4215
            TabIndex        =   34
            Top             =   1470
            Width           =   1335
            _Version        =   196608
            _ExtentX        =   2355
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483637
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
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   2
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
            Text            =   "0,0000"
            DecimalPlaces   =   4
            DecimalPoint    =   ""
            FixedPoint      =   -1  'True
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
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin UC_SeleccionEnLista.SeleccionEnLista selArticulos 
            Height          =   315
            Left            =   1185
            TabIndex        =   29
            Top             =   210
            Width           =   5880
            _ExtentX        =   10372
            _ExtentY        =   556
            PideAlias       =   -1  'True
            MuestraBoton    =   -1  'True
            EtiquetaDeAlias =   ""
            EtiquetaDeLista =   "Item"
            ValorInicial    =   "Seleccione Producto"
            IniciaConValor  =   0   'False
            ListIndex       =   0
         End
         Begin EditLib.fpDoubleSingle txtImporte 
            Height          =   315
            Left            =   6315
            TabIndex        =   35
            Top             =   1470
            Width           =   1170
            _Version        =   196608
            _ExtentX        =   2064
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483637
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
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   2
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
            Text            =   "0,00"
            DecimalPlaces   =   2
            DecimalPoint    =   ""
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "-9000000000"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   -1  'True
            IncInt          =   1
            IncDec          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin UC_SeleccionEnLista.SeleccionEnLista selDestinos 
            Height          =   315
            Left            =   5535
            TabIndex        =   36
            Tag             =   "o"
            Top             =   2040
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            AnchoDelAlias   =   195
            ListIndex       =   0
         End
         Begin EditLib.fpDoubleSingle txtCantidad 
            Height          =   315
            Left            =   1185
            TabIndex        =   33
            Top             =   1440
            Width           =   1485
            _Version        =   196608
            _ExtentX        =   2619
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483637
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
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   2
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
            Text            =   "0,0000"
            DecimalPlaces   =   4
            DecimalPoint    =   ""
            FixedPoint      =   -1  'True
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
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDoubleSingle txtPrecioRep 
            Height          =   315
            Left            =   1185
            TabIndex        =   30
            Top             =   1020
            Width           =   1485
            _Version        =   196608
            _ExtentX        =   2619
            _ExtentY        =   556
            Enabled         =   0   'False
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483624
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483637
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
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   2
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
            ControlType     =   1
            Text            =   "0,0000"
            DecimalPlaces   =   4
            DecimalPoint    =   ""
            FixedPoint      =   -1  'True
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
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDoubleSingle txtPrecioReferencial 
            Height          =   315
            Left            =   4215
            TabIndex        =   32
            Top             =   1020
            Width           =   1335
            _Version        =   196608
            _ExtentX        =   2355
            _ExtentY        =   556
            Enabled         =   0   'False
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483637
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
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   2
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
            Text            =   "0,0000"
            DecimalPlaces   =   4
            DecimalPoint    =   ""
            FixedPoint      =   -1  'True
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
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label lblMoneda1 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   180
            Left            =   2640
            TabIndex        =   162
            Top             =   1080
            Width           =   675
         End
         Begin VB.Label lblMoneda 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   195
            Left            =   6585
            TabIndex        =   161
            Top             =   705
            Width           =   675
         End
         Begin VB.Label lblUltimaVenta 
            Caption         =   "P. Ult. Vta."
            Height          =   255
            Left            =   3360
            TabIndex        =   155
            Top             =   1095
            Width           =   1005
         End
         Begin VB.Label lblPrecioArticulo 
            Caption         =   "P. Art�culo"
            Height          =   255
            Left            =   240
            TabIndex        =   142
            Top             =   1095
            Width           =   1095
         End
         Begin VB.Label lblEsquemaImpositivo 
            Caption         =   "Esquema impositivo"
            Height          =   255
            Left            =   3900
            TabIndex        =   77
            Top             =   2085
            Width           =   1575
         End
         Begin VB.Label lblItem 
            Caption         =   "&Art�culo"
            Height          =   255
            Left            =   240
            TabIndex        =   83
            Top             =   255
            Width           =   855
         End
         Begin VB.Label lblCantidad 
            Caption         =   "&Cantidad"
            Height          =   255
            Left            =   240
            TabIndex        =   82
            Top             =   1515
            Width           =   735
         End
         Begin VB.Label Label10 
            Caption         =   "&Precio"
            Height          =   255
            Left            =   3360
            TabIndex        =   81
            Top             =   1515
            Width           =   735
         End
         Begin VB.Label Label13 
            Caption         =   "I&mporte"
            Height          =   255
            Left            =   5640
            TabIndex        =   80
            Top             =   1515
            Width           =   615
         End
         Begin VB.Label Label14 
            Caption         =   "&Descripci�n"
            Height          =   255
            Left            =   240
            TabIndex        =   79
            Top             =   660
            Width           =   975
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000010&
            X1              =   0
            X2              =   7560
            Y1              =   1890
            Y2              =   1890
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000005&
            X1              =   0
            X2              =   7560
            Y1              =   1905
            Y2              =   1905
         End
      End
      Begin UC_BarraDeComandos.BarraDeComandos bdcCanje 
         Height          =   795
         Left            =   -23190
         TabIndex        =   103
         Top             =   -20610
         Width           =   2550
         _ExtentX        =   4498
         _ExtentY        =   1402
         BackColor       =   0
         CantidadDeBotones=   2
         Enabled         =   0   'False
         CaptionBoton(0) =   "&Grabar"
         ImagenBoton(0)  =   "Pedidos.frx":2F70
         EstiloBoton(0)  =   -1  'True
         BotonDefault(0) =   0   'False
         BotonCancel(0)  =   0   'False
         CaptionBoton(1) =   "&Imprimir"
         ImagenBoton(1)  =   "Pedidos.frx":33B2
         EstiloBoton(1)  =   -1  'True
         BotonDefault(1) =   0   'False
         BotonCancel(1)  =   0   'False
         CaptionBoton(2) =   "&Cerrar"
         ImagenBoton(2)  =   "Pedidos.frx":37F4
         EstiloBoton(2)  =   -1  'True
         BotonDefault(2) =   0   'False
         BotonCancel(2)  =   0   'False
         CaptionBoton(3) =   ""
         EstiloBoton(3)  =   0   'False
         BotonDefault(3) =   0   'False
         BotonCancel(3)  =   0   'False
         CaptionBoton(4) =   ""
         EstiloBoton(4)  =   0   'False
         BotonDefault(4) =   0   'False
         BotonCancel(4)  =   0   'False
         CaptionBoton(5) =   ""
         EstiloBoton(5)  =   0   'False
         BotonDefault(5) =   0   'False
         BotonCancel(5)  =   0   'False
         CaptionBoton(6) =   ""
         EstiloBoton(6)  =   0   'False
         BotonDefault(6) =   0   'False
         BotonCancel(6)  =   0   'False
         CaptionBoton(7) =   ""
         EstiloBoton(7)  =   0   'False
         BotonDefault(7) =   0   'False
         BotonCancel(7)  =   0   'False
         CaptionBoton(8) =   ""
         EstiloBoton(8)  =   0   'False
         BotonDefault(8) =   0   'False
         BotonCancel(8)  =   0   'False
      End
      Begin VB.Frame Frame2 
         Caption         =   "Datos de la Orden de Compra"
         Enabled         =   0   'False
         Height          =   2655
         Left            =   -23055
         TabIndex        =   119
         Top             =   -19815
         Width           =   7815
         Begin VB.CommandButton cmdUltimoO 
            Height          =   260
            Left            =   2880
            MaskColor       =   &H00FF00FF&
            Picture         =   "Pedidos.frx":3C36
            Style           =   1  'Graphical
            TabIndex        =   160
            TabStop         =   0   'False
            ToolTipText     =   "Pr�ximo n�mero"
            Top             =   360
            UseMaskColor    =   -1  'True
            Width           =   260
         End
         Begin UC_SeleccionEnLista.SeleccionEnLista selPuertos 
            Height          =   315
            Left            =   1125
            TabIndex        =   120
            Top             =   1440
            Width           =   3000
            _ExtentX        =   5292
            _ExtentY        =   556
            EtiquetaDeLista =   "Cosecha"
            AnchoDelAlias   =   0
            ListIndex       =   0
         End
         Begin EditLib.fpDateTime txtFechaContratoAcreditacion 
            Height          =   315
            Left            =   6045
            TabIndex        =   121
            Top             =   1440
            Width           =   1455
            _Version        =   196608
            _ExtentX        =   2566
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483637
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   16777215
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
            ButtonStyle     =   2
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   1
            AlignTextV      =   1
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
            Text            =   "24/07/2001"
            DateCalcMethod  =   4
            DateTimeFormat  =   5
            UserDefinedFormat=   "dd/mm/yyyy"
            DateMax         =   "00000000"
            DateMin         =   "00000000"
            TimeMax         =   "000000"
            TimeMin         =   "000000"
            TimeString1159  =   ""
            TimeString2359  =   ""
            DateDefault     =   "00000000"
            TimeDefault     =   "000000"
            TimeStyle       =   0
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            PopUpType       =   1
            DateCalcY2KSplit=   60
            CaretPosition   =   0
            IncYear         =   1
            IncMonth        =   1
            IncDay          =   1
            IncHour         =   1
            IncMinute       =   1
            IncSecond       =   1
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDateTime txtFComprobanteO 
            Height          =   315
            Left            =   6045
            TabIndex        =   122
            Top             =   360
            Width           =   1455
            _Version        =   196608
            _ExtentX        =   2566
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483637
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   16777215
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
            ButtonStyle     =   2
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   1
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483643
            InvalidOption   =   2
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483643
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "24/03/2004"
            DateCalcMethod  =   4
            DateTimeFormat  =   5
            UserDefinedFormat=   "dd/mm/yyyy"
            DateMax         =   "00000000"
            DateMin         =   "00000000"
            TimeMax         =   "000000"
            TimeMin         =   "000000"
            TimeString1159  =   ""
            TimeString2359  =   ""
            DateDefault     =   "00000000"
            TimeDefault     =   "000000"
            TimeStyle       =   0
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            PopUpType       =   1
            DateCalcY2KSplit=   60
            CaretPosition   =   0
            IncYear         =   1
            IncMonth        =   1
            IncDay          =   1
            IncHour         =   1
            IncMinute       =   1
            IncSecond       =   1
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDateTime txtFechaContratoDesde 
            Height          =   315
            Left            =   6045
            TabIndex        =   123
            Top             =   720
            Width           =   1455
            _Version        =   196608
            _ExtentX        =   2566
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483637
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   16777215
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
            ButtonStyle     =   2
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   1
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483643
            InvalidOption   =   2
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483643
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "24/03/2004"
            DateCalcMethod  =   4
            DateTimeFormat  =   5
            UserDefinedFormat=   "dd/mm/yyyy"
            DateMax         =   "00000000"
            DateMin         =   "00000000"
            TimeMax         =   "000000"
            TimeMin         =   "000000"
            TimeString1159  =   ""
            TimeString2359  =   ""
            DateDefault     =   "00000000"
            TimeDefault     =   "000000"
            TimeStyle       =   0
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            PopUpType       =   1
            DateCalcY2KSplit=   60
            CaretPosition   =   0
            IncYear         =   1
            IncMonth        =   1
            IncDay          =   1
            IncHour         =   1
            IncMinute       =   1
            IncSecond       =   1
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDateTime txtFechaContratoHasta 
            Height          =   315
            Left            =   6045
            TabIndex        =   124
            Top             =   1080
            Width           =   1455
            _Version        =   196608
            _ExtentX        =   2566
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483637
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   16777215
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
            ButtonStyle     =   2
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   1
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483643
            InvalidOption   =   2
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483643
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "24/03/2004"
            DateCalcMethod  =   4
            DateTimeFormat  =   5
            UserDefinedFormat=   "dd/mm/yyyy"
            DateMax         =   "00000000"
            DateMin         =   "00000000"
            TimeMax         =   "000000"
            TimeMin         =   "000000"
            TimeString1159  =   ""
            TimeString2359  =   ""
            DateDefault     =   "00000000"
            TimeDefault     =   "000000"
            TimeStyle       =   0
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            PopUpType       =   1
            DateCalcY2KSplit=   60
            CaretPosition   =   0
            IncYear         =   1
            IncMonth        =   1
            IncDay          =   1
            IncHour         =   1
            IncMinute       =   1
            IncSecond       =   1
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin UC_SeleccionEnLista.SeleccionEnLista selComisionistasO 
            Height          =   315
            Left            =   1125
            TabIndex        =   125
            Top             =   2190
            Width           =   3000
            _ExtentX        =   5292
            _ExtentY        =   556
            EtiquetaDeLista =   "Cosecha"
            AnchoDelAlias   =   0
            ListIndex       =   0
         End
         Begin UC_SeleccionEnLista.SeleccionEnLista selCorredores 
            Height          =   315
            Left            =   1125
            TabIndex        =   126
            Top             =   1830
            Width           =   3000
            _ExtentX        =   5292
            _ExtentY        =   556
            EtiquetaDeLista =   "Cosecha"
            AnchoDelAlias   =   0
            ListIndex       =   0
         End
         Begin EditLib.fpDoubleSingle txtComisionCorredor 
            Height          =   315
            Left            =   6045
            TabIndex        =   127
            Top             =   1800
            Width           =   1455
            _Version        =   196608
            _ExtentX        =   2566
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            AlignTextH      =   2
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
            Text            =   "0,00"
            DecimalPlaces   =   2
            DecimalPoint    =   ""
            FixedPoint      =   -1  'True
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
         Begin EditLib.fpDoubleSingle txtComisionVendedor 
            Height          =   315
            Left            =   6045
            TabIndex        =   128
            Top             =   2190
            Width           =   1455
            _Version        =   196608
            _ExtentX        =   2566
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            AlignTextH      =   2
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
            Text            =   "0,00"
            DecimalPlaces   =   2
            DecimalPoint    =   ""
            FixedPoint      =   -1  'True
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
         Begin EditLib.fpBoolean chkGeneraCaratula 
            Height          =   315
            Left            =   120
            TabIndex        =   129
            Top             =   720
            Visible         =   0   'False
            Width           =   1575
            _Version        =   196608
            _ExtentX        =   2778
            _ExtentY        =   556
            Enabled         =   0   'False
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ThreeDInsideStyle=   0
            ThreeDInsideHighlightColor=   -2147483637
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   0
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            AutoToggle      =   -1  'True
            BooleanStyle    =   0
            ToggleFalse     =   ""
            TextFalse       =   "Genera car�tula"
            BooleanPicture  =   2
            AlignPictureH   =   4
            AlignPictureV   =   1
            GroupId         =   0
            GroupTag        =   0
            GroupSelect     =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            MultiLine       =   -1  'True
            AlignTextH      =   0
            AlignTextV      =   1
            ToggleTrue      =   ""
            TextTrue        =   "Genera Caratula"
            Value           =   1
            BooleanMode     =   0
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            BorderGrayAreaColor=   -2147483637
            ToggleGrayed    =   ""
            TextGrayed      =   ""
            AllowMnemonic   =   -1  'True
            BackColor       =   -2147483633
            ForeColor       =   -2147483640
            ThreeDOnFocusInvert=   0   'False
            Caption         =   "Genera car�tula"
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            BooleanDataType =   0
            OLEDropMode     =   0
         End
         Begin UC_SeleccionEnLista.SeleccionEnLista selTiposDeCaratulas 
            Height          =   315
            Left            =   1125
            TabIndex        =   130
            Top             =   1080
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            EtiquetaDeAlias =   "Alias"
            EtiquetaDeLista =   "Cuenta Corriente"
            AnchoDelAlias   =   195
            ListIndex       =   0
         End
         Begin EditLib.fpText txtSucursalO 
            Height          =   315
            Left            =   1125
            TabIndex        =   131
            Top             =   360
            Width           =   495
            _Version        =   196608
            _ExtentX        =   873
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            Text            =   ""
            CharValidationText=   "0123456789"
            MaxLength       =   4
            MultiLine       =   0   'False
            PasswordChar    =   ""
            IncHoriz        =   0,25
            BorderGrayAreaColor=   -2147483637
            NoPrefix        =   -1  'True
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
         Begin EditLib.fpText txtNumeroO 
            Height          =   315
            Left            =   1635
            TabIndex        =   132
            Top             =   360
            Width           =   1155
            _Version        =   196608
            _ExtentX        =   2037
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            Text            =   ""
            CharValidationText=   "0123456789"
            MaxLength       =   8
            MultiLine       =   0   'False
            PasswordChar    =   ""
            IncHoriz        =   0,25
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
         Begin VB.Label lblTipo 
            Caption         =   "Tipo:"
            Height          =   255
            Left            =   165
            TabIndex        =   104
            Top             =   1125
            Width           =   855
         End
         Begin VB.Label Label28 
            Caption         =   "N�mero:"
            Height          =   255
            Left            =   165
            TabIndex        =   105
            Top             =   405
            Width           =   735
         End
         Begin VB.Label lblPuertos 
            Caption         =   "Puerto:"
            Height          =   255
            Left            =   165
            TabIndex        =   141
            Top             =   1515
            Width           =   615
         End
         Begin VB.Label Label29 
            Caption         =   "Plazo hasta:"
            Height          =   255
            Left            =   4965
            TabIndex        =   140
            Top             =   1120
            Width           =   1215
         End
         Begin VB.Label Label30 
            Caption         =   "Plazo desde:"
            Height          =   255
            Left            =   4965
            TabIndex        =   139
            Top             =   760
            Width           =   1215
         End
         Begin VB.Label Label31 
            Caption         =   "Registraci�n:"
            Height          =   255
            Left            =   4965
            TabIndex        =   138
            Top             =   400
            Width           =   975
         End
         Begin VB.Label Label32 
            Caption         =   "Acreditaci�n:"
            Height          =   255
            Left            =   4965
            TabIndex        =   137
            Top             =   1480
            Width           =   975
         End
         Begin VB.Label lblComisionVendedor 
            Caption         =   "Comisi�n (%):"
            Height          =   255
            Left            =   4965
            TabIndex        =   136
            Top             =   2235
            Width           =   975
         End
         Begin VB.Label lblComisionCorredor 
            Caption         =   "Comisi�n (%):"
            Height          =   255
            Left            =   4965
            TabIndex        =   135
            Top             =   1840
            Width           =   975
         End
         Begin VB.Label lblCorredores 
            Caption         =   "Corredor:"
            Height          =   255
            Left            =   165
            TabIndex        =   134
            Top             =   1875
            Width           =   975
         End
         Begin VB.Label Label35 
            Caption         =   "Comisionista:"
            Height          =   255
            Left            =   165
            TabIndex        =   133
            Top             =   2235
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Datos del canje"
         Enabled         =   0   'False
         Height          =   1575
         Left            =   -23055
         TabIndex        =   106
         Top             =   -17055
         Width           =   7815
         Begin UC_SeleccionEnLista.SeleccionEnLista selMercaderias 
            Height          =   315
            Left            =   1245
            TabIndex        =   107
            Top             =   360
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   556
            EtiquetaDeLista =   "Nombre"
            AnchoDelAlias   =   500
            ListIndex       =   0
         End
         Begin UC_SeleccionEnLista.SeleccionEnLista selDestinosO 
            Height          =   315
            Left            =   1245
            TabIndex        =   108
            Top             =   720
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   556
            EtiquetaDeLista =   "Destino"
            AnchoDelAlias   =   315
            ListIndex       =   0
         End
         Begin UC_SeleccionEnLista.SeleccionEnLista selModalidadesO 
            Height          =   315
            Left            =   1245
            TabIndex        =   109
            Top             =   1080
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   556
            EtiquetaDeLista =   "Destino"
            AnchoDelAlias   =   315
            ListIndex       =   0
         End
         Begin EditLib.fpDoubleSingle txtPrecioO 
            Height          =   315
            Left            =   5880
            TabIndex        =   113
            Top             =   360
            Width           =   1575
            _Version        =   196608
            _ExtentX        =   2778
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            AlignTextH      =   2
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
            Text            =   "0,0000"
            DecimalPlaces   =   4
            DecimalPoint    =   ""
            FixedPoint      =   -1  'True
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
         Begin EditLib.fpLongInteger txtKilos 
            Height          =   315
            Left            =   5880
            TabIndex        =   114
            TabStop         =   0   'False
            Top             =   720
            Width           =   1575
            _Version        =   196608
            _ExtentX        =   2778
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            AlignTextH      =   2
            AlignTextV      =   1
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
            MaxValue        =   "2147483647"
            MinValue        =   "-2147483648"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483633
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDoubleSingle txtImporteO 
            Height          =   315
            Left            =   5880
            TabIndex        =   117
            Top             =   1080
            Width           =   1575
            _Version        =   196608
            _ExtentX        =   2778
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            AlignTextH      =   2
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
            Text            =   "0,00"
            DecimalPlaces   =   2
            DecimalPoint    =   ""
            FixedPoint      =   -1  'True
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
         Begin VB.Label Label38 
            Caption         =   "Importe:"
            Height          =   255
            Left            =   5160
            TabIndex        =   118
            Top             =   1120
            Width           =   735
         End
         Begin VB.Label lblPrecio 
            Caption         =   "Precio:"
            Height          =   255
            Left            =   5160
            TabIndex        =   116
            Top             =   405
            Width           =   975
         End
         Begin VB.Label lblKilos 
            Caption         =   "Kilos:"
            Height          =   255
            Left            =   5160
            TabIndex        =   115
            Top             =   765
            Width           =   615
         End
         Begin VB.Label Label37 
            Caption         =   "Mercader�a:"
            Height          =   255
            Left            =   240
            TabIndex        =   112
            Top             =   390
            Width           =   1215
         End
         Begin VB.Label lblDestino 
            Caption         =   "Destino:"
            Height          =   255
            Left            =   240
            TabIndex        =   111
            Top             =   765
            Width           =   855
         End
         Begin VB.Label lblModalidad 
            Caption         =   "Modalidad:"
            Height          =   255
            Left            =   240
            TabIndex        =   110
            Top             =   1120
            Width           =   975
         End
      End
      Begin UC_SeleccionEnLista.SeleccionEnLista selCondicionesDeFlete 
         Height          =   720
         Left            =   -22275
         TabIndex        =   40
         Top             =   -16320
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   1270
         MuestraFrames   =   -1  'True
         MuestraBoton    =   -1  'True
         EtiquetaDeLista =   "Condici�n de flete"
         AnchoDelAlias   =   0
         CabezalCombo    =   -1  'True
         ListIndex       =   0
      End
      Begin VB.Frame fraExistencias 
         Enabled         =   0   'False
         Height          =   450
         Left            =   -21015
         TabIndex        =   38
         Top             =   -20880
         Width           =   5655
         Begin VB.Label lblExistenciaComercial 
            Caption         =   "0"
            Height          =   255
            Left            =   4320
            TabIndex        =   86
            Top             =   160
            Width           =   1215
         End
         Begin VB.Label lblExistenciaFisica 
            Caption         =   "0"
            Height          =   255
            Left            =   1560
            TabIndex        =   85
            Top             =   165
            Width           =   1095
         End
         Begin VB.Label Label20 
            Caption         =   "Existencia comercial:"
            Height          =   255
            Left            =   2760
            TabIndex        =   76
            Top             =   165
            Width           =   1575
         End
         Begin VB.Label Label19 
            Caption         =   "Existencia f�sica:"
            Height          =   255
            Left            =   240
            TabIndex        =   75
            Top             =   160
            Width           =   1215
         End
      End
      Begin VB.PictureBox picCompras 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   3615
         Left            =   -23175
         ScaleHeight     =   3615
         ScaleWidth      =   7935
         TabIndex        =   23
         Top             =   -21015
         Width           =   7935
         Begin FPSpread.vaSpread sprPedidosC 
            Height          =   1455
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   7575
            _Version        =   393216
            _ExtentX        =   13361
            _ExtentY        =   2566
            _StockProps     =   64
            DisplayRowHeaders=   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            RetainSelBlock  =   0   'False
            ScrollBars      =   2
            SelectBlockOptions=   0
            SpreadDesigner  =   "Pedidos.frx":4068
            UserResize      =   1
         End
         Begin FPSpread.vaSpread sprDetalleC 
            Height          =   1455
            Left            =   120
            TabIndex        =   25
            Top             =   1680
            Width           =   7575
            _Version        =   393216
            _ExtentX        =   13361
            _ExtentY        =   2566
            _StockProps     =   64
            DisplayRowHeaders=   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            RetainSelBlock  =   0   'False
            ScrollBars      =   2
            SelectBlockOptions=   0
            SpreadDesigner  =   "Pedidos.frx":431C
            UserResize      =   1
         End
         Begin UC_BarraDeComandos.BarraDeComandos bdcCopiar 
            Height          =   375
            Left            =   7080
            TabIndex        =   26
            Top             =   3210
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   661
            BackColor       =   0
            CantidadDeBotones=   1
            BotoneraReducida=   -1  'True
            Enabled         =   -1  'True
            CaptionBoton(0) =   ""
            ImagenBoton(0)  =   "Pedidos.frx":45D0
            EstiloBoton(0)  =   -1  'True
            BotonDefault(0) =   0   'False
            BotonCancel(0)  =   0   'False
            CaptionBoton(1) =   ""
            ImagenBoton(1)  =   "Pedidos.frx":4F52
            EstiloBoton(1)  =   -1  'True
            BotonDefault(1) =   0   'False
            BotonCancel(1)  =   0   'False
            CaptionBoton(2) =   ""
            ImagenBoton(2)  =   "Pedidos.frx":5394
            EstiloBoton(2)  =   -1  'True
            BotonDefault(2) =   0   'False
            BotonCancel(2)  =   0   'False
            CaptionBoton(3) =   ""
            EstiloBoton(3)  =   0   'False
            BotonDefault(3) =   0   'False
            BotonCancel(3)  =   0   'False
            CaptionBoton(4) =   ""
            EstiloBoton(4)  =   0   'False
            BotonDefault(4) =   0   'False
            BotonCancel(4)  =   0   'False
            CaptionBoton(5) =   ""
            EstiloBoton(5)  =   0   'False
            BotonDefault(5) =   0   'False
            BotonCancel(5)  =   0   'False
            CaptionBoton(6) =   ""
            EstiloBoton(6)  =   0   'False
            BotonDefault(6) =   0   'False
            BotonCancel(6)  =   0   'False
            CaptionBoton(7) =   ""
            EstiloBoton(7)  =   0   'False
            BotonDefault(7) =   0   'False
            BotonCancel(7)  =   0   'False
            CaptionBoton(8) =   ""
            EstiloBoton(8)  =   0   'False
            BotonDefault(8) =   0   'False
            BotonCancel(8)  =   0   'False
         End
         Begin VB.Label Label15 
            Caption         =   "Compras"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   78
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.Frame fraFletes 
         Enabled         =   0   'False
         Height          =   3735
         Left            =   -22215
         TabIndex        =   39
         Top             =   -20175
         Width           =   6135
         Begin UC_SeleccionEnLista.SeleccionEnLista selTransportistas 
            Height          =   315
            Left            =   1440
            TabIndex        =   41
            Top             =   360
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   556
            ListIndex       =   0
         End
         Begin UC_SeleccionEnLista.SeleccionEnLista selChoferes 
            Height          =   315
            Left            =   1440
            TabIndex        =   42
            Top             =   840
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   556
            ListIndex       =   0
         End
         Begin UC_SeleccionEnLista.SeleccionEnLista selCamiones 
            Height          =   315
            Left            =   1440
            TabIndex        =   43
            Top             =   1320
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   556
            ListIndex       =   0
         End
         Begin EditLib.fpLongInteger txtKmTierra 
            Height          =   315
            Left            =   1440
            TabIndex        =   44
            Top             =   2160
            Width           =   1215
            _Version        =   196608
            _ExtentX        =   2143
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            AlignTextH      =   2
            AlignTextV      =   1
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
            MaxValue        =   "2147483647"
            MinValue        =   "-2147483648"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483633
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpLongInteger txtKmAsfalto 
            Height          =   315
            Left            =   1440
            TabIndex        =   47
            Top             =   2640
            Width           =   1215
            _Version        =   196608
            _ExtentX        =   2143
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            AlignTextH      =   2
            AlignTextV      =   1
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
            MaxValue        =   "2147483647"
            MinValue        =   "-2147483648"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483633
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDoubleSingle txtTarifaTierra 
            Height          =   315
            Left            =   2760
            TabIndex        =   45
            Top             =   2160
            Width           =   1455
            _Version        =   196608
            _ExtentX        =   2566
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            AlignTextH      =   2
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
            Text            =   "0,00"
            DecimalPlaces   =   2
            DecimalPoint    =   ""
            FixedPoint      =   -1  'True
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
         Begin EditLib.fpDoubleSingle txtTarifaAsfalto 
            Height          =   315
            Left            =   2760
            TabIndex        =   48
            Top             =   2640
            Width           =   1455
            _Version        =   196608
            _ExtentX        =   2566
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            AlignTextH      =   2
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
            Text            =   "0,00"
            DecimalPlaces   =   2
            DecimalPoint    =   ""
            FixedPoint      =   -1  'True
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
         Begin EditLib.fpDoubleSingle txtImporteTierra 
            Height          =   315
            Left            =   4320
            TabIndex        =   46
            Top             =   2160
            Width           =   1455
            _Version        =   196608
            _ExtentX        =   2566
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            AlignTextH      =   2
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
            Text            =   "0,00"
            DecimalPlaces   =   2
            DecimalPoint    =   ""
            FixedPoint      =   -1  'True
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
         Begin EditLib.fpDoubleSingle txtImporteAsfalto 
            Height          =   315
            Left            =   4320
            TabIndex        =   49
            Top             =   2640
            Width           =   1455
            _Version        =   196608
            _ExtentX        =   2566
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            AlignTextH      =   2
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
            Text            =   "0,00"
            DecimalPlaces   =   2
            DecimalPoint    =   ""
            FixedPoint      =   -1  'True
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
         Begin EditLib.fpLongInteger txtKmTotal 
            Height          =   315
            Left            =   1440
            TabIndex        =   50
            Top             =   3240
            Width           =   1215
            _Version        =   196608
            _ExtentX        =   2143
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            AlignTextH      =   2
            AlignTextV      =   1
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
            MaxValue        =   "2147483647"
            MinValue        =   "-2147483648"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483633
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDoubleSingle txtImporteTotal 
            Height          =   315
            Left            =   4320
            TabIndex        =   51
            Top             =   3240
            Width           =   1455
            _Version        =   196608
            _ExtentX        =   2566
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            AlignTextH      =   2
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
            Text            =   "0,00"
            DecimalPlaces   =   2
            DecimalPoint    =   ""
            FixedPoint      =   -1  'True
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
         Begin VB.Line Line14 
            BorderColor     =   &H80000001&
            BorderWidth     =   2
            X1              =   4320
            X2              =   5760
            Y1              =   3120
            Y2              =   3120
         End
         Begin VB.Line Line13 
            BorderColor     =   &H80000001&
            BorderWidth     =   2
            X1              =   1440
            X2              =   2640
            Y1              =   3120
            Y2              =   3120
         End
         Begin VB.Label Label33 
            Caption         =   "Empresa:"
            Height          =   375
            Left            =   480
            TabIndex        =   73
            Top             =   405
            Width           =   1695
         End
         Begin VB.Label Label34 
            Caption         =   "Chofer:"
            Height          =   255
            Left            =   480
            TabIndex        =   72
            Top             =   870
            Width           =   735
         End
         Begin VB.Label Label11 
            Caption         =   "Cami�n:"
            Height          =   255
            Left            =   480
            TabIndex        =   71
            Top             =   1350
            Width           =   735
         End
         Begin VB.Label Label23 
            Alignment       =   2  'Center
            Caption         =   "Kil�metros"
            Height          =   255
            Left            =   1440
            TabIndex        =   70
            Top             =   1920
            Width           =   1215
         End
         Begin VB.Label Label40 
            Alignment       =   2  'Center
            Caption         =   "Importe"
            Height          =   255
            Left            =   4320
            TabIndex        =   69
            Top             =   1920
            Width           =   1455
         End
         Begin VB.Label Label36 
            Alignment       =   2  'Center
            Caption         =   "Tarifa"
            Height          =   255
            Left            =   2640
            TabIndex        =   68
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Label Label24 
            Caption         =   "Tierra"
            Height          =   255
            Left            =   480
            TabIndex        =   67
            Top             =   2205
            Width           =   975
         End
         Begin VB.Label Label25 
            Caption         =   "Asfalto"
            Height          =   255
            Left            =   480
            TabIndex        =   66
            Top             =   2685
            Width           =   1095
         End
         Begin VB.Line Line5 
            BorderColor     =   &H80000003&
            X1              =   0
            X2              =   6120
            Y1              =   1800
            Y2              =   1800
         End
         Begin VB.Line Line6 
            BorderColor     =   &H80000005&
            X1              =   0
            X2              =   6120
            Y1              =   1815
            Y2              =   1815
         End
         Begin VB.Label Label26 
            Caption         =   "Total"
            Height          =   255
            Left            =   480
            TabIndex        =   65
            Top             =   3285
            Width           =   1095
         End
      End
      Begin VB.Frame fraGeneral 
         Height          =   4815
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   7575
         Begin UC_SeleccionEnLista.SeleccionEnLista selCuentas 
            Height          =   315
            Left            =   1200
            TabIndex        =   7
            Top             =   360
            Width           =   3285
            _ExtentX        =   5794
            _ExtentY        =   556
            PideAlias       =   -1  'True
            AnchoDelAlias   =   700
            ListIndex       =   0
         End
         Begin UC_SeleccionEnLista.SeleccionEnLista selSubCuentas 
            Height          =   315
            Left            =   5640
            TabIndex        =   8
            Top             =   360
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   556
            EtiquetaDeLista =   "Sub Cuenta"
            AnchoDelAlias   =   225
            ListIndex       =   0
         End
         Begin UC_SeleccionEnLista.SeleccionEnLista selModalidades 
            Height          =   315
            Left            =   1200
            TabIndex        =   9
            Top             =   960
            Width           =   3275
            _ExtentX        =   5768
            _ExtentY        =   556
            AnchoDelAlias   =   315
            ListIndex       =   0
         End
         Begin UC_SeleccionEnLista.SeleccionEnLista selProveedores 
            Height          =   315
            Left            =   1200
            TabIndex        =   11
            Top             =   1560
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   556
            PideAlias       =   -1  'True
            AnchoDelAlias   =   700
            ListIndex       =   0
         End
         Begin EditLib.fpDateTime txtFVencimiento 
            Height          =   315
            Left            =   1200
            TabIndex        =   21
            Tag             =   "o"
            Top             =   4080
            Width           =   1455
            _Version        =   196608
            _ExtentX        =   2566
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483637
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   16777215
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
            ButtonStyle     =   2
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   2
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
            Text            =   "09/04/2002"
            DateCalcMethod  =   4
            DateTimeFormat  =   5
            UserDefinedFormat=   "dd/mm/yyyy"
            DateMax         =   "00000000"
            DateMin         =   "00000000"
            TimeMax         =   "000000"
            TimeMin         =   "000000"
            TimeString1159  =   ""
            TimeString2359  =   ""
            DateDefault     =   "00000000"
            TimeDefault     =   "000000"
            TimeStyle       =   0
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            PopUpType       =   0
            DateCalcY2KSplit=   60
            CaretPosition   =   0
            IncYear         =   1
            IncMonth        =   1
            IncDay          =   1
            IncHour         =   1
            IncMinute       =   1
            IncSecond       =   1
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin UC_SeleccionEnLista.SeleccionEnLista selMonedas 
            Height          =   315
            Left            =   1200
            TabIndex        =   18
            Top             =   3480
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            AnchoDelAlias   =   0
            ListIndex       =   0
         End
         Begin EditLib.fpDoubleSingle txtCotizacion 
            Height          =   315
            Left            =   3720
            TabIndex        =   19
            Top             =   3480
            Width           =   1095
            _Version        =   196608
            _ExtentX        =   1931
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483637
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
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   2
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
            Text            =   "0,0000"
            DecimalPlaces   =   4
            DecimalPoint    =   ""
            FixedPoint      =   -1  'True
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
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin UC_SeleccionEnLista.SeleccionEnLista selExpresadoEn 
            Height          =   315
            Left            =   6120
            TabIndex        =   20
            Top             =   3480
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            AnchoDelAlias   =   0
            ListIndex       =   0
         End
         Begin UC_SeleccionEnLista.SeleccionEnLista selLugaresDeRecepcion 
            Height          =   315
            Left            =   1200
            TabIndex        =   13
            Top             =   2160
            Width           =   3345
            _ExtentX        =   5900
            _ExtentY        =   556
            MuestraBoton    =   -1  'True
            AnchoDelAlias   =   435
            ListIndex       =   0
         End
         Begin UC_SeleccionEnLista.SeleccionEnLista selCampanias 
            Height          =   315
            Left            =   5640
            TabIndex        =   10
            Top             =   960
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   556
            EtiquetaDeLista =   "Sub Cuenta"
            AnchoDelAlias   =   225
            ListIndex       =   0
         End
         Begin UC_SeleccionEnLista.SeleccionEnLista selTiposDePedidos 
            Height          =   315
            Left            =   5640
            TabIndex        =   12
            Top             =   1560
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   556
            AnchoDelAlias   =   225
            ListIndex       =   0
         End
         Begin UC_SeleccionEnLista.SeleccionEnLista selUnidadesDeNegocio 
            Height          =   315
            Left            =   5640
            TabIndex        =   14
            Top             =   2160
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   556
            EtiquetaDeLista =   "Unidad de Negocios"
            AnchoDelAlias   =   195
            ListIndex       =   0
         End
         Begin EditLib.fpBoolean chkCotizacionFija 
            Height          =   225
            Left            =   2835
            TabIndex        =   164
            Top             =   3840
            Width           =   1995
            _Version        =   196608
            _ExtentX        =   3528
            _ExtentY        =   397
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ThreeDInsideStyle=   0
            ThreeDInsideHighlightColor=   -2147483637
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   0
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            AutoToggle      =   -1  'True
            BooleanStyle    =   0
            ToggleFalse     =   ""
            TextFalse       =   "Negocio con TC fijo?"
            BooleanPicture  =   2
            AlignPictureH   =   4
            AlignPictureV   =   1
            GroupId         =   0
            GroupTag        =   0
            GroupSelect     =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   0
            MultiLine       =   -1  'True
            AlignTextH      =   0
            AlignTextV      =   1
            ToggleTrue      =   ""
            TextTrue        =   ""
            Value           =   0
            BooleanMode     =   0
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            BorderGrayAreaColor=   -2147483637
            ToggleGrayed    =   ""
            TextGrayed      =   ""
            AllowMnemonic   =   -1  'True
            BackColor       =   -2147483633
            ForeColor       =   -2147483640
            ThreeDOnFocusInvert=   0   'False
            Caption         =   "Negocio con TC fijo?"
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            BooleanDataType =   0
            OLEDropMode     =   0
         End
         Begin UC_SeleccionEnLista.SeleccionEnLista selListaPrecios 
            Height          =   315
            Left            =   1200
            TabIndex        =   15
            Top             =   2760
            Width           =   3275
            _ExtentX        =   5768
            _ExtentY        =   556
            AnchoDelAlias   =   435
            ListIndex       =   0
         End
         Begin UC_SeleccionEnLista.SeleccionEnLista selDepositos 
            Height          =   315
            Left            =   5640
            TabIndex        =   16
            Top             =   2760
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   556
            AnchoDelAlias   =   180
            ListIndex       =   0
         End
         Begin VB.Label Label17 
            Caption         =   "Dep�sito"
            Height          =   255
            Left            =   4680
            TabIndex        =   17
            Top             =   2805
            Width           =   975
         End
         Begin VB.Label lblListaPrecios 
            Caption         =   "Lista Precios"
            Height          =   255
            Left            =   240
            TabIndex        =   165
            Top             =   2805
            Width           =   1455
         End
         Begin VB.Label Label41 
            Caption         =   "Unidad"
            Height          =   255
            Left            =   4680
            TabIndex        =   143
            Top             =   2205
            Width           =   615
         End
         Begin VB.Label Label27 
            Caption         =   "F. de pago"
            Height          =   255
            Left            =   4680
            TabIndex        =   64
            Top             =   1605
            Width           =   855
         End
         Begin VB.Label Label21 
            Caption         =   "Campa�a"
            Height          =   255
            Left            =   4680
            TabIndex        =   74
            Top             =   1005
            Width           =   975
         End
         Begin VB.Label lblLugarRecepcion 
            Caption         =   "Recepci�n"
            Height          =   255
            Left            =   120
            TabIndex        =   84
            Top             =   2205
            Width           =   1455
         End
         Begin VB.Label lblExpresadoEn 
            Caption         =   "Expresado en"
            Height          =   255
            Left            =   5040
            TabIndex        =   94
            Top             =   3525
            Width           =   1095
         End
         Begin VB.Label lblCotizacion 
            Caption         =   "Cotizaci�n:"
            Height          =   255
            Left            =   2880
            TabIndex        =   93
            Top             =   3525
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "Moneda"
            Height          =   255
            Left            =   240
            TabIndex        =   92
            Top             =   3525
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "Vencimiento"
            Height          =   270
            Left            =   240
            TabIndex        =   91
            Top             =   4125
            Width           =   975
         End
         Begin VB.Line Line9 
            BorderColor     =   &H80000005&
            X1              =   0
            X2              =   7560
            Y1              =   3255
            Y2              =   3255
         End
         Begin VB.Line Line10 
            BorderColor     =   &H80000010&
            X1              =   0
            X2              =   7560
            Y1              =   3240
            Y2              =   3240
         End
         Begin VB.Label lblCuentas 
            Caption         =   "Cuenta"
            Height          =   255
            Left            =   120
            TabIndex        =   90
            Top             =   405
            Width           =   1575
         End
         Begin VB.Label lblSubCuentas 
            Caption         =   "Sub cuenta"
            Height          =   255
            Left            =   4680
            TabIndex        =   89
            Top             =   400
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Facturaci�n"
            Height          =   255
            Left            =   120
            TabIndex        =   88
            Top             =   1005
            Width           =   1695
         End
         Begin VB.Label lblProveedores 
            Caption         =   "Proveedor"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   87
            Top             =   1605
            Width           =   855
         End
      End
      Begin VB.Frame fraOtros 
         Enabled         =   0   'False
         Height          =   5415
         Left            =   -22695
         TabIndex        =   52
         Top             =   -20775
         Width           =   7215
         Begin MemoLib.fpMemo txtCondiciones 
            Height          =   855
            Left            =   1200
            TabIndex        =   57
            Top             =   2160
            Width           =   5535
            _Version        =   196608
            _ExtentX        =   9763
            _ExtentY        =   1508
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            HideSelection   =   -1  'True
            NullColor       =   -2147483637
            OnFocusNoSelect =   -1  'True
            OnFocusPosition =   3
            ControlType     =   0
            Text            =   ""
            WordWrap        =   -1  'True
            ShowEOL         =   0   'False
            SelMode         =   0
            LineLimit       =   20
            ScrollBars      =   2
            PageWidth       =   0
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483633
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ProcessTab      =   0   'False
            TabLength       =   0
            AutoMenu        =   0   'False
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin UC_SeleccionEnLista.SeleccionEnLista selCondicionesComerciales 
            Height          =   315
            Left            =   1200
            TabIndex        =   55
            Top             =   1680
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            ListIndex       =   0
         End
         Begin EditLib.fpDateTime txtFCondicion 
            Height          =   315
            Left            =   5400
            TabIndex        =   56
            Tag             =   "o"
            Top             =   1680
            Width           =   1335
            _Version        =   196608
            _ExtentX        =   2355
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483637
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   16777215
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
            ButtonStyle     =   2
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   2
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
            Text            =   "09/04/2002"
            DateCalcMethod  =   4
            DateTimeFormat  =   5
            UserDefinedFormat=   "dd/mm/yyyy"
            DateMax         =   "00000000"
            DateMin         =   "00000000"
            TimeMax         =   "000000"
            TimeMin         =   "000000"
            TimeString1159  =   ""
            TimeString2359  =   ""
            DateDefault     =   "00000000"
            TimeDefault     =   "000000"
            TimeStyle       =   0
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            PopUpType       =   0
            DateCalcY2KSplit=   60
            CaretPosition   =   0
            IncYear         =   1
            IncMonth        =   1
            IncDay          =   1
            IncHour         =   1
            IncMinute       =   1
            IncSecond       =   1
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin UC_SeleccionEnLista.SeleccionEnLista selComisionistas 
            Height          =   315
            Left            =   1200
            TabIndex        =   53
            Top             =   360
            Width           =   2910
            _ExtentX        =   5133
            _ExtentY        =   556
            PideAlias       =   -1  'True
            AnchoDelAlias   =   700
            ListIndex       =   0
         End
         Begin EditLib.fpDoubleSingle txtComision 
            Height          =   315
            Left            =   5850
            TabIndex        =   54
            Top             =   360
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            AlignTextH      =   2
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
            Text            =   "0,00"
            DecimalPlaces   =   2
            DecimalPoint    =   ""
            FixedPoint      =   -1  'True
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
         Begin MemoLib.fpMemo txtComentarioViejo 
            Height          =   1335
            Left            =   1200
            TabIndex        =   60
            Top             =   3840
            Visible         =   0   'False
            Width           =   5655
            _Version        =   196608
            _ExtentX        =   9975
            _ExtentY        =   2355
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            HideSelection   =   -1  'True
            NullColor       =   -2147483637
            OnFocusNoSelect =   -1  'True
            OnFocusPosition =   3
            ControlType     =   0
            Text            =   ""
            WordWrap        =   -1  'True
            ShowEOL         =   0   'False
            SelMode         =   0
            LineLimit       =   20
            ScrollBars      =   2
            PageWidth       =   0
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483633
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ProcessTab      =   0   'False
            TabLength       =   0
            AutoMenu        =   0   'False
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin UC_SeleccionEnLista.SeleccionEnLista selComentarios 
            Height          =   315
            Left            =   1200
            TabIndex        =   58
            Top             =   3360
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   556
            ListIndex       =   0
         End
         Begin EditLib.fpBoolean chkAgregar 
            Height          =   375
            Left            =   5760
            TabIndex        =   59
            Top             =   3360
            Width           =   975
            _Version        =   196608
            _ExtentX        =   1720
            _ExtentY        =   661
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ThreeDInsideStyle=   0
            ThreeDInsideHighlightColor=   -2147483637
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   0
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            AutoToggle      =   -1  'True
            BooleanStyle    =   0
            ToggleFalse     =   ""
            TextFalse       =   "Agregar"
            BooleanPicture  =   2
            AlignPictureH   =   3
            AlignPictureV   =   1
            GroupId         =   0
            GroupTag        =   0
            GroupSelect     =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            MultiLine       =   -1  'True
            AlignTextH      =   0
            AlignTextV      =   1
            ToggleTrue      =   ""
            TextTrue        =   ""
            Value           =   0
            BooleanMode     =   0
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            BorderGrayAreaColor=   -2147483637
            ToggleGrayed    =   ""
            TextGrayed      =   ""
            AllowMnemonic   =   -1  'True
            BackColor       =   -2147483633
            ForeColor       =   -2147483640
            ThreeDOnFocusInvert=   0   'False
            Caption         =   "Agregar"
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            BooleanDataType =   0
            OLEDropMode     =   0
         End
         Begin UC_SeleccionEnLista.SeleccionEnLista selCobradores 
            Height          =   315
            Left            =   1200
            TabIndex        =   156
            Top             =   840
            Width           =   2910
            _ExtentX        =   5133
            _ExtentY        =   556
            PideAlias       =   -1  'True
            AnchoDelAlias   =   700
            ListIndex       =   0
         End
         Begin EditLib.fpDoubleSingle txtComisionCobrador 
            Height          =   315
            Left            =   5850
            TabIndex        =   157
            Top             =   840
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            AlignTextH      =   2
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
            Text            =   "0,00"
            DecimalPlaces   =   2
            DecimalPoint    =   ""
            FixedPoint      =   -1  'True
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
         Begin RichTextLib.RichTextBox txtComentario 
            Height          =   1335
            Left            =   1200
            TabIndex        =   163
            Top             =   3840
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   2355
            _Version        =   393217
            Enabled         =   -1  'True
            TextRTF         =   $"Pedidos.frx":57D6
         End
         Begin VB.Image imgMail 
            Height          =   480
            Left            =   4215
            Picture         =   "Pedidos.frx":5858
            Top             =   315
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label lblComisionCobrador 
            Caption         =   "Comisi�n (%)"
            Height          =   255
            Left            =   4845
            TabIndex        =   159
            Top             =   885
            Width           =   975
         End
         Begin VB.Label Label45 
            Caption         =   "Cobrador"
            Height          =   255
            Left            =   240
            TabIndex        =   158
            Top             =   885
            Width           =   1575
         End
         Begin VB.Line Line12 
            BorderColor     =   &H80000005&
            X1              =   0
            X2              =   7200
            Y1              =   1335
            Y2              =   1335
         End
         Begin VB.Line Line11 
            BorderColor     =   &H80000000&
            X1              =   0
            X2              =   7200
            Y1              =   1320
            Y2              =   1320
         End
         Begin VB.Line Line3 
            BorderColor     =   &H80000000&
            X1              =   0
            X2              =   7200
            Y1              =   3165
            Y2              =   3165
         End
         Begin VB.Line Line4 
            BorderColor     =   &H80000005&
            X1              =   0
            X2              =   7200
            Y1              =   3180
            Y2              =   3180
         End
         Begin VB.Label Label3 
            Caption         =   "Condici�n"
            Height          =   255
            Left            =   240
            TabIndex        =   100
            Top             =   1725
            Width           =   855
         End
         Begin VB.Label Label9 
            Caption         =   "Vencimiento"
            Height          =   270
            Left            =   4440
            TabIndex        =   99
            Top             =   1725
            Width           =   1095
         End
         Begin VB.Label Label16 
            Caption         =   "Comisionista"
            Height          =   255
            Left            =   240
            TabIndex        =   98
            Top             =   405
            Width           =   1575
         End
         Begin VB.Label lblComision 
            Caption         =   "Comisi�n (%)"
            Height          =   255
            Left            =   4845
            TabIndex        =   97
            Top             =   405
            Width           =   975
         End
         Begin VB.Label lblComentario 
            Caption         =   "Comentario"
            Height          =   975
            Left            =   240
            TabIndex        =   96
            Top             =   3840
            Width           =   975
         End
         Begin VB.Label Label6 
            Caption         =   "Leyenda"
            Height          =   255
            Left            =   240
            TabIndex        =   95
            Top             =   3405
            Width           =   855
         End
      End
      Begin FPSpread.vaSpread sprPedidos 
         Height          =   1695
         Left            =   -22935
         TabIndex        =   22
         Top             =   -17295
         Width           =   7575
         _Version        =   393216
         _ExtentX        =   13361
         _ExtentY        =   2990
         _StockProps     =   64
         Enabled         =   0   'False
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         SpreadDesigner  =   "Pedidos.frx":5B62
         UserResize      =   1
      End
      Begin UC_BarraDeComandos.BarraDeComandos bdcDetalle 
         Height          =   375
         Left            =   -22965
         TabIndex        =   37
         Top             =   -20895
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   661
         BackColor       =   0
         CantidadDeBotones=   3
         BotoneraReducida=   -1  'True
         Enabled         =   0   'False
         CaptionBoton(0) =   ""
         ImagenBoton(0)  =   "Pedidos.frx":5E16
         EstiloBoton(0)  =   -1  'True
         BotonDefault(0) =   0   'False
         BotonCancel(0)  =   0   'False
         CaptionBoton(1) =   ""
         ImagenBoton(1)  =   "Pedidos.frx":6258
         EstiloBoton(1)  =   -1  'True
         BotonDefault(1) =   0   'False
         BotonCancel(1)  =   0   'False
         CaptionBoton(2) =   ""
         ImagenBoton(2)  =   "Pedidos.frx":669A
         EstiloBoton(2)  =   -1  'True
         BotonDefault(2) =   0   'False
         BotonCancel(2)  =   0   'False
         CaptionBoton(3) =   ""
         EstiloBoton(3)  =   0   'False
         BotonDefault(3) =   0   'False
         BotonCancel(3)  =   0   'False
         CaptionBoton(4) =   ""
         EstiloBoton(4)  =   0   'False
         BotonDefault(4) =   0   'False
         BotonCancel(4)  =   0   'False
         CaptionBoton(5) =   ""
         EstiloBoton(5)  =   0   'False
         BotonDefault(5) =   0   'False
         BotonCancel(5)  =   0   'False
         CaptionBoton(6) =   ""
         EstiloBoton(6)  =   0   'False
         BotonDefault(6) =   0   'False
         BotonCancel(6)  =   0   'False
         CaptionBoton(7) =   ""
         EstiloBoton(7)  =   0   'False
         BotonDefault(7) =   0   'False
         BotonCancel(7)  =   0   'False
         CaptionBoton(8) =   ""
         EstiloBoton(8)  =   0   'False
         BotonDefault(8) =   0   'False
         BotonCancel(8)  =   0   'False
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1020
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      Begin VB.PictureBox Picture1 
         Height          =   310
         Left            =   1800
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   152
         Top             =   480
         Width           =   315
         Begin VB.CommandButton cmdUltimo 
            Height          =   260
            Left            =   0
            MaskColor       =   &H00FF00FF&
            Picture         =   "Pedidos.frx":6ADC
            Style           =   1  'Graphical
            TabIndex        =   153
            TabStop         =   0   'False
            ToolTipText     =   "Pr�ximo n�mero"
            Top             =   0
            UseMaskColor    =   -1  'True
            Width           =   260
         End
      End
      Begin EditLib.fpMask txtNumero 
         Height          =   315
         Left            =   840
         TabIndex        =   2
         Top             =   480
         Width           =   885
         _Version        =   196608
         _ExtentX        =   1561
         _ExtentY        =   556
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         ThreeDInsideStyle=   1
         ThreeDInsideHighlightColor=   -2147483637
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   1
         ThreeDOutsideHighlightColor=   16777215
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
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   2
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
         AllowOverflow   =   0   'False
         BestFit         =   0   'False
         ClipMode        =   0
         DataFormatEx    =   0
         Mask            =   "########"
         PromptChar      =   "_"
         PromptInclude   =   0   'False
         RequireFill     =   0   'False
         BorderGrayAreaColor=   -2147483637
         NoPrefix        =   0   'False
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483637
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         AutoTab         =   0   'False
         ButtonColor     =   -2147483637
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpMask txtSucursal 
         Height          =   315
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   510
         _Version        =   196608
         _ExtentX        =   900
         _ExtentY        =   556
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         ThreeDInsideStyle=   1
         ThreeDInsideHighlightColor=   -2147483637
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   1
         ThreeDOutsideHighlightColor=   16777215
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
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   2
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
         AllowOverflow   =   0   'False
         BestFit         =   0   'False
         ClipMode        =   0
         DataFormatEx    =   0
         Mask            =   "####"
         PromptChar      =   "_"
         PromptInclude   =   0   'False
         RequireFill     =   0   'False
         BorderGrayAreaColor=   -2147483637
         NoPrefix        =   0   'False
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483637
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         AutoTab         =   0   'False
         ButtonColor     =   -2147483637
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpDateTime txtFComprobante 
         Height          =   315
         Left            =   6480
         TabIndex        =   4
         Tag             =   "o"
         Top             =   480
         Width           =   1335
         _Version        =   196608
         _ExtentX        =   2355
         _ExtentY        =   556
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         ThreeDInsideStyle=   1
         ThreeDInsideHighlightColor=   -2147483637
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   1
         ThreeDOutsideHighlightColor=   16777215
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
         ButtonStyle     =   2
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   2
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
         Text            =   "09/04/2002"
         DateCalcMethod  =   4
         DateTimeFormat  =   5
         UserDefinedFormat=   "dd/mm/yyyy"
         DateMax         =   "00000000"
         DateMin         =   "00000000"
         TimeMax         =   "000000"
         TimeMin         =   "000000"
         TimeString1159  =   ""
         TimeString2359  =   ""
         DateDefault     =   "00000000"
         TimeDefault     =   "000000"
         TimeStyle       =   0
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483637
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         PopUpType       =   0
         DateCalcY2KSplit=   60
         CaretPosition   =   0
         IncYear         =   1
         IncMonth        =   1
         IncDay          =   1
         IncHour         =   1
         IncMinute       =   1
         IncSecond       =   1
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpText txtNroInterno 
         Height          =   315
         Left            =   4800
         TabIndex        =   3
         Top             =   480
         Width           =   975
         _Version        =   196608
         _ExtentX        =   1720
         _ExtentY        =   556
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         ThreeDInsideStyle=   1
         ThreeDInsideHighlightColor=   -2147483637
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
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   2
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
         Text            =   ""
         CharValidationText=   ""
         MaxLength       =   255
         MultiLine       =   0   'False
         PasswordChar    =   ""
         IncHoriz        =   0,25
         BorderGrayAreaColor=   -2147483637
         NoPrefix        =   0   'False
         ScrollV         =   0   'False
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483637
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483637
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin UC_SeleccionEnLista.SeleccionEnLista selComprobantes 
         Height          =   315
         Left            =   2400
         TabIndex        =   101
         Top             =   480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         AnchoDelAlias   =   675
         ListIndex       =   0
      End
      Begin VB.Label lblSelCpte 
         Caption         =   "Seleccionar Cpte"
         Height          =   255
         Left            =   2400
         TabIndex        =   102
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "N� interno"
         Height          =   270
         Left            =   4800
         TabIndex        =   63
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Fecha"
         Height          =   270
         Left            =   6480
         TabIndex        =   62
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "N�mero"
         Height          =   255
         Left            =   360
         TabIndex        =   61
         Top             =   240
         Width           =   735
      End
   End
   Begin UC_BarraDeComandos.BarraDeComandos bdcAbm 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   7455
      Left            =   8525
      TabIndex        =   154
      Top             =   0
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   13150
      CantidadDeBotones=   5
      DistribucionVertical=   -1  'True
      Dockeable       =   -1  'True
      Enabled         =   -1  'True
      CaptionBoton(0) =   "&Agregar"
      ImagenBoton(0)  =   "Pedidos.frx":6F0E
      EstiloBoton(0)  =   -1  'True
      BotonDefault(0) =   0   'False
      BotonCancel(0)  =   0   'False
      CaptionBoton(1) =   "&Modificar"
      ImagenBoton(1)  =   "Pedidos.frx":7350
      EstiloBoton(1)  =   -1  'True
      BotonDefault(1) =   0   'False
      BotonCancel(1)  =   0   'False
      CaptionBoton(2) =   "&Eliminar"
      ImagenBoton(2)  =   "Pedidos.frx":7792
      EstiloBoton(2)  =   -1  'True
      BotonDefault(2) =   0   'False
      BotonCancel(2)  =   0   'False
      CaptionBoton(3) =   "&Imprimir"
      ImagenBoton(3)  =   "Pedidos.frx":7BD4
      EstiloBoton(3)  =   -1  'True
      BotonDefault(3) =   0   'False
      BotonCancel(3)  =   0   'False
      CaptionBoton(4) =   "&Cerrar"
      ImagenBoton(4)  =   "Pedidos.frx":8016
      EstiloBoton(4)  =   -1  'True
      BotonDefault(4) =   0   'False
      BotonCancel(4)  =   0   'False
      CaptionBoton(5) =   "A&plicar"
      ImagenBoton(5)  =   "Pedidos.frx":8458
      EstiloBoton(5)  =   -1  'True
      BotonDefault(5) =   0   'False
      BotonCancel(5)  =   0   'False
      CaptionBoton(6) =   "E&tiqueta"
      ImagenBoton(6)  =   "Pedidos.frx":889A
      EstiloBoton(6)  =   -1  'True
      BotonDefault(6) =   0   'False
      BotonCancel(6)  =   0   'False
      CaptionBoton(7) =   "Imp&rimir CP"
      ImagenBoton(7)  =   "Pedidos.frx":8D9C
      EstiloBoton(7)  =   -1  'True
      BotonDefault(7) =   0   'False
      BotonCancel(7)  =   -1  'True
      CaptionBoton(8) =   ""
      EstiloBoton(8)  =   0   'False
      BotonDefault(8) =   0   'False
      BotonCancel(8)  =   0   'False
   End
End
Attribute VB_Name = "frmPedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DBHelper As New DBHelper
Dim Pedidos As New M6_Pedidos
Dim Empresa As New EmpresaCOM
Dim CuentasCorrientes As New M0_CuentasCorrientes
Dim SubCuentasCorrientes As New M0_SubCuentasCorrientes
Dim LugaresDeRecepcion As New M0_CtasCtesLugDeRecepcion
Dim Depositos As New M0_Depositos
Dim TiposDeMovimientos As New M6_TiposDeMovimientos
Dim Monedas As New M0_Monedas
Dim Leyendas As New M0_Leyendas
Dim CondicionesComerciales As New M6_CondicionesComerciales
Dim Destinos As New M0_Destinos
Dim Fletes As New M4_FletesBUS
Dim TarifasFletes As New M4_TarifasFletes
Dim Transportistas As New M4_Transportistas
Dim Choferes As New M4_Choferes
Dim Camiones As New M4_Camiones
Dim Listas As New M0_Listas
Dim Comisionistas As New M0_Comisionistas
Dim Cobradores As New M0_Cobradores
    Dim Articulos As New M6_Articulos
Dim DerivadorFAC As New DerivadorFAC
Dim FormModulo6 As New M6_FormulariosFAC
Dim Formularios As New M0_Formularios
Dim Campanias As New M6_Campanias
Dim CondicionesDeFlete As New M6_CondicionesDeFlete
Dim ListasM6 As New M6_Listas
Dim Mercaderias As New M3_Mercaderias
Dim Operaciones As New M3_Operaciones
Dim Corredores As New M3_Corredores
Dim Puertos As New M3_Puertos
Dim Caratulas As New M3_Caratulas
Dim CaratulasTipos As New M3_CaratulasTipos
Dim UnidadesDeNegocio As New M0_UnidadesDeNegocio
Dim ListaPrecios As New M6_ListaPrecios

Public TipoDeOperacion As EnumTiposDeOperacion
Public TipoDePedido As EnumTiposDePedido

Public IDPedidos As Long, IDFletes As Long, NoEntrar As Boolean

Dim CambioAnchoColumnas As Boolean, IDOperaciones As Long

Private Enum colP
    Numero = 1
    Fecha = 2
    vencimiento = 3
    Condiciones = 4
    tipo = 5
    Comentario = 6
    IDPedido = 7
    Asociada = 8
End Enum

Private Enum colC
    articulo = 1
    Descripcion = 2
    Cantidad = 3
    Precio = 4
    Importe = 5
    Deposito = 6
    Destino = 7
    IDArticulos = 8
    IDDepositos = 9
    IDDestinos = 10
    Id = 11
    TipoC = 12
    PrecioReferencial = 13
    CantidadAutorizada = 14
    Remitido = 15
    Facturado = 16
End Enum

Private Enum colD
    articulo = 1
    CantidadTotal = 2
    Cuenta = 3
    Cantidad = 4
    Id = 5
    IDCuenta = 6
    IDRemitos = 7
End Enum

Private Enum CtrlTab
    tabGeneral = 0
    tabPedidos = 1
    tabDetalle = 2
    tabFletes = 3
    tabOtros = 4
    tabcanje = 5
    tabDivision = 6
End Enum

Private ParametroFCO As Boolean         'NJE 11/12/2017

Private WithEvents oMail As clsCDOmail
Attribute oMail.VB_VarHelpID = -1
Dim UsaListaPrecios As Boolean
Dim ParametroWRH As Boolean
Dim EsExtranjero As Boolean


Private Sub Form_Load()

    CentraForm Me
    Screen.MousePointer = vbArrowHourglass
    
    Pedidos.CadenaDeConexion = CadenaDeConexion
    CuentasCorrientes.CadenaDeConexion = CadenaDeConexion
    SubCuentasCorrientes.CadenaDeConexion = CadenaDeConexion
    LugaresDeRecepcion.CadenaDeConexion = CadenaDeConexion
    Depositos.CadenaDeConexion = CadenaDeConexion
    TiposDeMovimientos.CadenaDeConexion = CadenaDeConexion
    Monedas.CadenaDeConexion = CadenaDeConexion
    Leyendas.CadenaDeConexion = CadenaDeConexion
    CondicionesComerciales.CadenaDeConexion = CadenaDeConexion
    Destinos.CadenaDeConexion = CadenaDeConexion
    Fletes.CadenaDeConexion = CadenaDeConexion
    TarifasFletes.CadenaDeConexion = CadenaDeConexion
    Transportistas.CadenaDeConexion = CadenaDeConexion
    Choferes.CadenaDeConexion = CadenaDeConexion
    Camiones.CadenaDeConexion = CadenaDeConexion
    Listas.CadenaDeConexion = CadenaDeConexion
    Comisionistas.CadenaDeConexion = CadenaDeConexion
    Cobradores.CadenaDeConexion = CadenaDeConexion
                Articulos.CadenaDeConexion = CadenaDeConexion
    DerivadorFAC.CadenaDeConexion = CadenaDeConexion
    FormModulo6.CadenaDeConexion = CadenaDeConexion
    Formularios.CadenaDeConexion = CadenaDeConexion
    Campanias.CadenaDeConexion = CadenaDeConexion
    CondicionesDeFlete.CadenaDeConexion = CadenaDeConexion
    ListasM6.CadenaDeConexion = CadenaDeConexion
    Mercaderias.CadenaDeConexion = CadenaDeConexion
    Operaciones.CadenaDeConexion = CadenaDeConexion
    Corredores.CadenaDeConexion = CadenaDeConexion
    Puertos.CadenaDeConexion = CadenaDeConexion
    Caratulas.CadenaDeConexion = CadenaDeConexion
    CaratulasTipos.CadenaDeConexion = CadenaDeConexion
    UnidadesDeNegocio.CadenaDeConexion = CadenaDeConexion
    ListaPrecios.CadenaDeConexion = CadenaDeConexion
    
    'TiposDeMovimientos.TomaUno Tipo
    'Me.Caption = TiposDeMovimientos.Descripcion
  
    AplicaPermisosAFormulario Me, bdcAbm, "34420"
    AplicaPermisosAFormulario Me, bdcCanje, "33"
  
  
'    NoEntrar = True
  
    'PDM 13/10/2017 15:31 //T.94805 Solicitaba ocultar esta solapa. Agrega condiciones para la cantidad de solapas.
    Me.tabPedidos.Tab = CtrlTab.tabDivision
    Me.tabPedidos.TabState = 1
    
    Me.tabPedidos.TabsPerRow = 5
    
    If TipoDeOperacion = m6Compra Then
        'PDM 13/11/2017 09:58 //T.18463 Se oculta el tab de canje solo para pedidos de compra.
        Me.tabPedidos.Tab = CtrlTab.tabcanje
        Me.tabPedidos.TabState = 1
    End If
    
    If TipoDeOperacion = m6Compra And TipoDePedido = m6Ampliacion Then
        Me.tabPedidos.TabsPerRow = 7
    End If
    
    If TipoDeOperacion = m6Venta And TipoDePedido = m6Ampliacion Then
        Me.tabPedidos.TabsPerRow = 7
    End If
    
    
    
    ArmaGrillaDivision sprDivisionPedido
    ArmaGrillaDivision sprDivisionRemitos
    
    If TipoDeOperacion = m6Compra Or TipoDePedido = m6Ampliacion Then
        selProveedores.Visible = False
        lblProveedores.Visible = False
        selProveedores.Id = 0
        If TipoDePedido = m6Ampliacion Then
            Me.tabPedidos.Tab = CtrlTab.tabcanje
            Me.tabPedidos.TabState = 1
            Me.tabPedidos.TabsPerRow = Me.tabPedidos.TabsPerRow - 1
        End If
        Me.tabPedidos.Tab = CtrlTab.tabDivision
        Me.tabPedidos.TabState = 1
        Me.tabPedidos.TabsPerRow = Me.tabPedidos.TabsPerRow - 1
    End If
    
    If TipoDeOperacion = m6Venta And TraeParametro("UDM", mbInsumos) = 1 Then
        Me.tabPedidos.Tab = CtrlTab.tabcanje
        If Me.tabPedidos.TabState = 0 Then
            Me.tabPedidos.TabState = 1
            Me.tabPedidos.TabsPerRow = Me.tabPedidos.TabsPerRow - 1
        End If
        Me.tabPedidos.Tab = CtrlTab.tabDivision
        If Me.tabPedidos.TabState = 0 Then
            Me.tabPedidos.TabState = 1
            Me.tabPedidos.TabsPerRow = Me.tabPedidos.TabsPerRow - 1
        End If
    End If
    ' AAB - Parametro ' Pedidos: Toma Precio de Listas de Precios
    UsaListaPrecios = TraeParametro("LPP", mbInsumos)
    
    If TipoDeOperacion = m6Compra Or TipoDePedido = m6Ampliacion Then
        sprPedidos.Height = 4695
        HabilitarControles picCompras, False
    End If
    
    Me.selListaPrecios.Visible = False
    Me.lblListaPrecios.Visible = False
    If TipoDeOperacion = m6Venta Then
        If TraeParametro("HLP", mbInsumos) = 1 Then ' Parametro: Habilita uso de Listas de Precios
            If TraeParametro("LPP", mbInsumos) = 1 Then  ' Parametro: Pedidos: Toma Precio de Listas de Precios
                Me.lblListaPrecios.Visible = True
                Me.selListaPrecios.Visible = True
            End If
        End If
    End If
    
    
    If TipoDePedido = m6Original Then
        If TipoDeOperacion = m6Compra Then
            Me.Caption = "Pedido de Compra de Mercader�a"
        ElseIf TipoDeOperacion = m6Venta Then
            Me.Caption = "Pedido de Venta de Mercader�a"
        End If
    Else
        If TipoDeOperacion = m6Compra Then
            Me.Caption = "Ampliaci�n/reducci�n de pedidos de compra"
        ElseIf TipoDeOperacion = m6Venta Then
            Me.Caption = "Ampliaci�n/reducci�n de pedidos de venta"
        End If
        HabilitarControles fraGeneral, False
        HabilitarControles fraFletes, False
        HabilitarControles fraOtros, False
        HabilitarControles fraDetalle, False
        lblCuentas.Enabled = True
        selCuentas.Enabled = True
        lblCantidad.Enabled = True
        txtCantidad.Enabled = True
        Me.Picture1.Visible = False
    End If
    bdcAbm.Estado = "00001"
    
    If TipoDeOperacion = m6Compra Then
        chkGeneraCaratula.Value = ValueTrue
    Else
        chkGeneraCaratula.Value = ValueFalse
    End If
    txtFComprobante.Text = date
    txtFCondicion.Text = date
    txtFVencimiento.Text = date
    txtFComprobanteO.Text = date
    txtFechaContratoDesde.Text = date
    txtFechaContratoHasta.Text = date
    txtFechaContratoAcreditacion.Text = date
    
    With selUnidadesDeNegocio
        .CamposVisibles = "Descripcion"
        .LlenaLista UnidadesDeNegocio.Lista
        .IDDeDefault = Usuarios.IDUnidadesDeNegocio
    End With
    
    With selProveedores
        .CamposVisibles = "Nombre"
        .CampoDelAlias = "Alias"
        '.ClaveEnRegistro = Me.Caption + " - 2 - " + .Name
        .LlenaLista rsCuentasCorrientesEscritura.Clone
        .AnchoDeLista = 4000
        selProveedores_Click
    End With
    
    With selCuentas
        .CamposVisibles = "Nombre"
        .CampoDelAlias = "Alias"
        .ClaveEnRegistro = Me.Caption + " - 1 - " + .Name
        .CantidadDeFilas = 30
        .AnchoDeLista = 4000
        .LlenaLista rsCuentasCorrientesEscritura.Clone
    End With
    
    With selSubCuentas
        .CamposVisibles = "Descripcion"
    End With
    
    With selTiposDePedidos
        .CamposVisibles = "Descripcion"
        .ClaveEnRegistro = Me.Caption + " - 3 - " + .Name
        .CantidadDeFilas = 20
        .LlenaLista ListasM6.ListaTiposDePedidos
    End With
    
    With selLugaresDeRecepcion
        .CamposVisibles = "Descripcion"
    End With
    
    With selCondicionesDeFlete
        .CamposVisibles = "Descripcion"
        .LlenaLista CondicionesDeFlete.Lista
    End With
    selCuentas_LostFocus
    
    With selModalidades
        .CamposVisibles = "Descripcion"
        .LlenaLista Listas.ListaModalidades(mbInsumos, IIf(TipoDeOperacion = m6Compra, -EnumModalidades.m6Directa, 9999))
        .CantidadDeFilas = 30
        selModalidades_Click
    End With
    
    With Me.selCuentasDivision
        .CamposVisibles = "Nombre"
        .CampoDelAlias = "Alias"
        .CantidadDeFilas = 30
        .AnchoDeLista = 4000
        .LlenaLista rsCuentasCorrientesEscritura.Clone
    End With
        
    
    NoEntrar = True
    
    With selArticulos
        .PideAlias = True
        .CamposVisibles = "NombreP"
        .CampoValorInicial = "NombreP"
        .CampoDelAlias = "Alias"
        .CantidadDeFilas = 30
        .AnchoDeLista = 5000
        .LlenaLista rsArticulos.Clone
    End With
    
        
    With selDepositos
        .CamposVisibles = "Descripcion"
        .CampoValorInicial = "Descripcion"
        .ValorInicial = "sin asignar"
        .LlenaLista Depositos.Lista(mbInsumos)
    End With
    
    'PDM 04/09/2017 13:11 //T.17784 Se deshabilita el campo, ya se habia modificado anteriormente para que no cargue esta lista. Se vuelde a deshabilitar en sprDetalle_Click
    'JAM 14/01/2021 //T.31149 Se habilita el campo, un cliente lo pidi�, se agrega el valor inicial "sin asignar" para que no moleste a los demas.
    'selDepositos.Enabled = False
    
    NoEntrar = False

    With selDestinos
        .CamposVisibles = "Descripcion"
        .LlenaLista Destinos.Lista(mbInsumos)
    End With
    
    'PDM 12/03/2019 16:49 //T.15331 En algunas PC no funciona .ClaveEnRegistro por eso se agrego que guarde la moneda en el registro.
    Dim Moneda As String
    With selMonedas
        .CamposVisibles = "Descripcion"
        '.ClaveEnRegistro = .Name + str(Modulo)
        '.IDDeDefault = 2 ' Dolar
        .LlenaLista Monedas.Lista
        Moneda = LeerRegistry("Software\AS\PedidosMoneda", "Moneda")
        If Moneda <> "" Then
            .Id = CInt(Moneda)
            .IDDeDefault = CInt(Moneda)
        Else
            .Id = 1
            .IDDeDefault = 1
        End If
        selMonedas_LostFocus
    End With

    'PDM 24/10/2019 16:28 //T.20773
    Dim MonedaExpresado As String
    With selExpresadoEn
        .CamposVisibles = "Descripcion"
        .LlenaLista Monedas.Lista
        If selMonedas.Id <> 1 Then
            MonedaExpresado = LeerRegistry("Software\AS\PedidosMonedaExpresado", "MonedaExpresado")
        Else
            MonedaExpresado = 1
        End If
        If MonedaExpresado <> "" Then
            .Id = CInt(MonedaExpresado)
            .IDDeDefault = CInt(MonedaExpresado)
        Else
            .Id = 1
            .IDDeDefault = 1
        End If
        selExpresadoEn_LostFocus
    End With
    
    
'    With selExpresadoEn
'        .CamposVisibles = "Descripcion"
'        .LlenaLista Monedas.Lista
'    End With
    
    With selComentarios
        .CamposVisibles = "Descripcion"
        .LlenaLista Leyendas.Lista '(mbInsumos)
    End With
    
    ParametroFCO = CBool(CCur2(TraeParametro("FCO", mbInsumos)))
    
    With selCondicionesComerciales
        .CamposVisibles = "Descripcion"
        .CampoValorInicial = "Descripcion"  'NJE 11/12/2017 - TCK 52389: Agrego la opci�n ninguna para el caso de no utilizar este campo
        .ValorInicial = "ninguna"           'NJE 11/12/2017 - TCK 52389: Agrego la opci�n ninguna para el caso de no utilizar este campo
        .LlenaLista CondicionesComerciales.Lista
        .tag = IIf(ParametroFCO And TipoDeOperacion <> m6Compra, "o", "") 'NJE 11/12/2017 - TCK 52389: Es obligatorio o no de acuerdo al par�metro
        If CBool(CCur2(TraeParametro("PCC", mbInsumos))) = True And TipoDeOperacion <> m6Compra Then
            CondicionesComerciales.TomaUnoPorPredeterminado
            .IDDeDefault = CondicionesComerciales.Id
            .Id = CondicionesComerciales.Id
        End If
    End With
    
    With selComisionistas
        .CamposVisibles = "Descripcion"
        .CampoValorInicial = "Descripcion"
        .ValorInicial = "ninguno"
        .LlenaLista Comisionistas.Lista
    End With
    
    
    With selCobradores
        .CamposVisibles = "Descripcion"
        .CampoValorInicial = "Descripcion"
        .ValorInicial = "ninguno"
        If TipoDeOperacion = m6Venta And TraeParametro("UVC", mbBase) = True Then
            .LlenaLista Comisionistas.Lista
        Else
            .LlenaLista Cobradores.Lista
        End If
    End With
        
    With selTransportistas
        .CamposVisibles = "Descripcion"
        .CampoValorInicial = "Descripcion"
        .ValorInicial = "flete contratado por el " & IIf(TipoDeOperacion = m6Compra, "productor", "proveedor")
        .LlenaLista Transportistas.Lista
    End With
    selTransportistas_Click
    
    With selCampanias
        .CantidadDeFilas = 30
        .CamposVisibles = "Descripcion"
        .IDDeDefault = Campanias.TomaPredeterminada
        .Id = .IDDeDefault
        .LlenaLista Campanias.Lista
    End With
    
    With selMercaderias
        .CamposVisibles = "Nombre"
        .CantidadDeFilas = 100
        .LlenaLista Mercaderias.Lista()
    End With
    
    With selDestinosO
        .CamposVisibles = "Descripcion"
        .LlenaLista Destinos.Lista(mbCereales, 2, EnumTiposDeDestino.m3Canje)
    End With
    
    With selModalidadesO
        .CamposVisibles = "Descripcion"
        .LlenaLista Listas.ListaModalidades(mbCereales)
    End With
    
    With selTiposDeCaratulas
        .CamposVisibles = "Descripcion"
        selTiposDeCaratulas.LlenaLista CaratulasTipos.Lista(IIf(TipoDeOperacion = m6Compra, m3ContratoDeVenta, m3ContratoDeCompra))
    End With
    
    With selCorredores
        .CamposVisibles = "Descripcion"
        .CampoValorInicial = "Descripcion"
        .ValorInicial = "directo"
        .LlenaLista Corredores.Lista
    End With
    
    With selComisionistasO
        .CamposVisibles = "Descripcion"
        .CampoValorInicial = "Descripcion"
        .ValorInicial = "directo"
        .LlenaLista Comisionistas.Lista
    End With
    
    With selPuertos
        .CamposVisibles = "Descripcion"
        .ClaveEnRegistro = .Name + str(Modulo)
        .LlenaLista Puertos.Lista
    End With
    
    txtSucursalO.Text = "0000"
    txtNumeroO.Text = FormatoRG(Pedidos.UltimoNumero(m6Compra) + 1, 8, True, False)
    '    txtNumeroO.Text = FormatoRG(Operaciones.UltimoNumero(m6CompraOrden) + 1, 8, True, False)
    
    ArmaGrillaPedidos sprPedidos
    ArmaGrillaCuerpo sprDetalle
    
    
    ArmaGrillaPedidos sprPedidosC
    ArmaGrillaCuerpo sprDetalleC
    
    'LeeFormatoGrillas Usuarios.ID, Me, sprPedidos, TipoDeOperacion
    'LeeFormatoGrillas Usuarios.ID, Me, sprPedidosC, TipoDeOperacion
    'LeeFormatoGrillas Usuarios.ID, Me, sprDetalle, TipoDeOperacion
    'LeeFormatoGrillas Usuarios.ID, Me, sprDetalleC, TipoDeOperacion

    txtSucursal.Text = "0001"

    If Formularios.TomaUnoPorUsuario(IIf(TipoDeOperacion = m6Compra, m6PedidoCompra, m6PedidoVenta), Usuarios.Id, m6SinModalidad) = 0 Then
        bdcAbm.Habilitar , , , False
    Else
        bdcAbm.Habilitar False, , , False
    End If
    
    'PDM 26/09/2017 14:51 //T.48123
    If TipoDeOperacion = m6Compra Then
        lblLugarRecepcion.Caption = "Origen"
    ElseIf TipoDeOperacion = m6Venta Then
        lblLugarRecepcion.Caption = "Recepci�n"
    End If
    
    ColoreaObligatorios Me
    
    Screen.MousePointer = vbDefault
        
    If IDPedidos > 0 Then
        Pedidos.TomaUno IDPedidos
        Me.txtSucursal.Text = Pedidos.Sucursal
        Me.txtNumero.Text = Pedidos.Numero
        txtNumero_LostFocus
    End If
    
    'PDM 29/09/2017 15:45 //T.38089 Oculta la solapa de pedidos, cuando se carga un pedido nuevo.
    If TipoDePedido = m6Original Then
        Me.tabPedidos.Tab = CtrlTab.tabPedidos
        Me.tabPedidos.TabState = 1
    End If

    'JAM 9/12/2017 en ventas fija la moneda para CSD carga solo en dolares
    If TraeParametro("CSD", mbTodosModulos) = True Then
        'PDM 10/04/2019 12:03 //T.16225 Se agrego para que deshabilite el selector de moneda.
        Me.selExpresadoEn.IDDeDefault = selMonedas.Id
        selMonedas.IDDeDefault = 2
        selMonedas.BackColorCombo = mbObligatorio
        selMonedas.Enabled = False
        txtCotizacion.BackColor = mbObligatorio
        selMonedas.tag = "o"
        selMonedas_LostFocus
        
        Me.selExpresadoEn.Enabled = False
        Me.lblExpresadoEn.Enabled = False
        
    End If
    
    If TipoDeOperacion = m6Venta Then
        With Me.selListaPrecios
            .CamposVisibles = "Descripcion"
            '.CampoValorInicial = "Descripcion"
            '.ValorInicial = "ninguna"
            .LlenaLista ListaPrecios.Lista
            ListaPrecios.TomaUnoPorPredeterminado
            .IDDeDefault = ListaPrecios.Id
            .Id = ListaPrecios.Id
        End With
    End If
    ParametroWRH = TraeParametro("WRH", mbInsumos)
    If ParametroWRH = True And TipoDeOperacion = m6Venta Then
        With selProveedores
            .CamposVisibles = "Nombre"
            .CampoDelAlias = "Alias"
            .CampoValorInicial = "Nombre"
            .ValorInicial = "empresa"
            .IDDeDefault = 0
            '.ClaveEnRegistro = Me.Caption + " - 2 - " + .Name
            .LlenaLista rsCuentasCorrientesEscritura.Clone
            
            .AnchoDeLista = 4000
'            selProveedores_Click
        End With
'        NoEntrar = True
        selProveedores.Enabled = True
        lblProveedores.Caption = "Titular"
        lblProveedores.Enabled = True
        lblProveedores.Visible = True
    End If
    
    
    'PDM 28/07/2021 13:06 //T.34588
    If (TraeParametro("UFP", mbInsumos) = 1 And TipoDeOperacion = m6Venta) Then 'Or (TraeParametro("UFP", mbInsumos) = 1 And TipoDeOperacion = m6Venta And IDPedidos > 0) Then
        txtFComprobante.Enabled = False
    Else
        txtFComprobante.Enabled = True
    End If
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    GrabaFormatoGrillas Usuarios.Id, Me, sprPedidos, TipoDeOperacion
    GrabaFormatoGrillas Usuarios.Id, Me, sprPedidosC, TipoDeOperacion
    GrabaFormatoGrillas Usuarios.Id, Me, sprDetalle, TipoDeOperacion
    GrabaFormatoGrillas Usuarios.Id, Me, sprDetalleC, TipoDeOperacion

    CargarMemo 0
    
    Set UnidadesDeNegocio = Nothing

End Sub

    Private Sub bdcAbm_Click()
        
        Dim rs As Recordset, TipoDeImpresion As EnumContextoImpresion
        Dim b As Integer
        Dim OkPedido As Long
        Dim a As Integer
        
        Select Case bdcAbm.Valor
                
            Case "agregar", "modificar", "imprimir"
                    
                    If Not (ValidarDatos(Me) And ValidarFechas) Then    'NJE 01/03/2017: Agregado validarfechas
                        Exit Sub
                    End If
                    
                    'PDM 06/08/2018 13:34 //T.10597 10608
                    If ErroTipoConsumo = True Then
                        MsgBox "Los pedidos de consumos solo pueden utilizarse con una cuenta corriente con el CUIT de la empresa." & Chr(13) & "Pedido no guardado.", vbInformation
                        Exit Sub
                    End If

                    
                    If TipoDeOperacion = m6Venta Then
                        If bdcAbm.Valor = "agregar" Or bdcAbm.Valor = "imprimir" Then
                            CuentasCorrientes.TomaUno selCuentas.Id, Usuarios.Id
                            If CuentasCorrientes.Bloqueada = True Then
                                MsgBox "Esta cuenta se encuentra bloqueada para Pedidos.", vbInformation
                                Exit Sub
                            End If
                        End If
                    End If
                                    
                    CompletarDatos
                    
                    If Pedidos.rsCuerpo.RecordCount = 0 Then
                        MsgBox "El pedido no tiene art�culos cargados", vbExclamation
                        Exit Sub
                    End If
                    
                    If TipoDePedido = m6Ampliacion Then
                        If IDPedidos > 0 Then
                            If TraeParametro("UDM", mbInsumos) = True Then
                                Empresa.LeerDatos
                                If Empresa.NroIdentificadorFiscal = "30-70952682-7" Or Empresa.NroIdentificadorFiscal = "30-50076581-6" Then  ' AG
                                    For a = 1 To sprDetalle.MaxRows
                                        ' JAM 21/09/2021 TK 36644 Al control de que est� o no en un despacho se le agrega el articulo.
                                        sprDetalle.Col = colC.articulo
                                        sprDetalle.Row = a
                                        If sprDetalle.ForeColor <> mbNegro Then
                                            If Val(sprTexto(sprDetalle, colC.IDArticulos, a)) > 0 And (IDPedidos > 0 Or (IDPedidos = 0 And CCur2(sprTexto(sprDetalle, colC.Cantidad, a)) <> 0)) Then
                                                If Pedidos.ListaPorSQL("SELECT IDDespachos=ISNULL((SELECT MAX(DC.IDDespachos) FROM M6_DespachosCuerpo DC JOIN M6_Despachos D ON D.ID=DC.IDDespachos WHERE D.Activa=1 AND DC.IDPedidos=" & IDPedidos & " AND DC.IDItems=" & sprTexto(sprDetalle, colC.IDArticulos, a) & "),0)")!IDDespachos > 0 Then
                                                    If Pedidos.ListaPorSQL("SELECT Usado=ISNULL((SELECT Usado=ISNULL((SELECT MAX(1) FROM M6_DespachosCuerpo DC JOIN M6_Despachos D ON D.ID=DC.IDDespachos WHERE D.Activa=1 AND DC.IDPedidos=" & IDPedidos & " AND (DC.Cantidad>DC.CantidadARemitir OR DC.Cantidad>DC.CantidadAFacturar AND DC.IDItems=" & sprTexto(sprDetalle, colC.IDArticulos, a) & ")),0)),0)")!Usado = 1 Then
                                                        If Me.txtCantidad > 0 Or (Me.txtCantidad < 0 And Pedidos.ListaPorSQL("SELECT Usado=ISNULL((SELECT Usado=ISNULL((SELECT MAX(DC.Cantidad-DC.CantidadARemitir) FROM M6_DespachosCuerpo DC JOIN M6_Despachos D ON D.ID=DC.IDDespachos WHERE D.Activa=1 AND DC.IDPedidos=" & IDPedidos & " AND (DC.Cantidad>DC.CantidadARemitir OR DC.CantidadARemitir=DC.CantidadAFacturar AND DC.IDItems=" & sprTexto(sprDetalle, colC.IDArticulos, a) & ")),0)),0)")!Usado < -Me.txtCantidad.Text) Then
                                                            MsgBox "No es posible modificar un pedido que fue incluido en un despacho.", vbInformation, "Despachos"
                                                            Exit Sub
                                                        ElseIf Me.txtCantidad.Text < 0 Then  'Reduccion -- JAM 3/11/2021 TK. 37650  ac� se reduce el despacho.
                                                            Pedidos.EjecutaSQL "UPDATE M6_DespachosCuerpo SET Cantidad=Cantidad" & Replace(Me.txtCantidad.Text, ",", ".") & ", CantidadARemitir=CantidadARemitir" & Replace(Me.txtCantidad.Text, ",", ".") & ", CantidadAFacturar=CantidadAFacturar" & Replace(Me.txtCantidad.Text, ",", ".") & " WHERE IDPedidos=" & IDPedidos & " AND IDItems=" & sprTexto(sprDetalle, colC.IDArticulos, a) & " "
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Next a
                                End If
                            End If
                        End If
                    End If
                    
                    If bdcAbm.Valor = "agregar" Or bdcAbm.Valor = "imprimir" Then
                        If IDPedidos = 0 Then
                            TipoDeImpresion = mbImpresionOriginal
                            Dim Pedidos1 As New M6_Pedidos
                            Pedidos1.CadenaDeConexion = CadenaDeConexion
                            Pedidos1.TomaUnoPorNumero Pedidos.Propio, Pedidos.Sucursal, Pedidos.Numero
                            If Pedidos1.Id > 0 Then
                                MsgBox "Numero de pedido existente", vbCritical
                            Else
                                IDPedidos = Pedidos.Grabar(Usuarios.Id, mbAlta, TipoDePedido)
                                
                                'P 20140611 17:03 //Pregunta si se quiere enviar un mail al comercial.
                                'P 20150513 09:54 // Envia mail para los pedidos de Compra
                                If TipoDeOperacion = m6Compra Then
                                    If TraeParametro("ETC", mbInsumos) = "1" Then
                                        If IDPedidos > 0 Then
                                            If MsgBox("Desea enviar un mail con los datos del pedido a todos los Comerciales ?", vbYesNo, Me.Caption) = vbYes Then
                                                EnviaMailTodos "Alta"
                                            End If
                                        End If
                                    Else
                                        If IDPedidos > 0 And selComisionistas.Id > 0 Then
                                            Comisionistas.TomaUno selComisionistas.Id
                                            'P 20140612 11:48 //Verifica que el comisionista tenga mail.
                                            If Trim(Comisionistas.mail) <> "" Then
                                                If MsgBox("Desea enviar un mail con los datos del pedido al Comercial", vbYesNo, Me.Caption) = vbYes Then
                                                    EnviaMail "Alta"
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                                
                                'P 20150513 09:54 //Envia Mail para los Pedidos de Venta
                                If TipoDeOperacion = m6Venta Then
                                    If TraeParametro("ETV", mbInsumos) = "1" Then
                                        If IDPedidos > 0 Then
                                            If MsgBox("Desea enviar un mail con los datos del pedido a todos los Comerciales ?", vbYesNo, Me.Caption) = vbYes Then
                                                EnviaMailTodos "Alta"
                                            End If
                                        End If
                                    Else
                                        If IDPedidos > 0 And selComisionistas.Id > 0 Then
                                            Comisionistas.TomaUno selComisionistas.Id
                                            'P 20140612 11:48 //Verifica que el comisionista tenga mail.
                                            If Trim(Comisionistas.mail) <> "" Then
                                                If MsgBox("Desea enviar un mail con los datos del pedido al Comercial", vbYesNo, Me.Caption) = vbYes Then
                                                    EnviaMail "Alta"
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                            Set Pedidos1 = Nothing
                        Else
                            TipoDeImpresion = mbReimpresion
                        End If
                        
                        If IDPedidos > 0 Then
                            If bdcAbm.Valor = "imprimir" Then
                            
                                If Mid(Me.bdcAbm.Estado, 2, 1) = "1" Then
                                    DerivadorFAC.ImprimirComprobante IDPedidos, IIf(TipoDeOperacion = m6Compra, m6PedidoCompra, m6PedidoVenta), selModalidades.Id, mbInsumos, txtFComprobante.Text, Me.txtSucursal.Text, Me.txtNumero.Text, TipoDeImpresion, True
                                Else
                                    frmImprimirCpte.Imprimir IDPedidos, Pedidos.IDFormularios, txtFComprobante.Text, TipoDeImpresion, IIf(TipoDeOperacion = m6Compra, m6PedidoCompra, m6PedidoVenta), selModalidades.Id, mbInsumos
                                    If frmImprimirCpte.ImprimioCorrectamente = False Then
                                    Pedidos.Id = IDPedidos
                                    IDPedidos = Pedidos.Grabar(Usuarios.Id, mbBaja, TipoDePedido)
                                    IDPedidos = 0
                                    Exit Sub
                                    End If
                                End If
                            End If
                            txtCantidad.MaxValue = 99999999
                            If IDPedidos > 0 And IDOperaciones > 0 Then
                                If TipoDeOperacion = m6Compra Then
                                    OkPedido = Pedidos.GrabaAplicaciones(IDPedidos, 0, IDOperaciones)
                                Else
                                    OkPedido = Pedidos.GrabaAplicaciones(0, IDPedidos, IDOperaciones)
                                End If
                                If OkPedido > 0 Then
                                    IDOperaciones = 0   ' reseteo la operacion
                                Else
                                    MsgBox "El pedido no pudo aplicarse a la Orden", vbExclamation
                                End If
                            End If
                        Else
                            MsgBox "El pedido no pudo grabarse", vbExclamation
                        End If
                        Me.tabPedidos.ActiveTab = CtrlTab.tabGeneral
                        txtSucursal.SetFocus
                        bdcAbm.Estado = "00001"
                    Else
                        'modificacion
                        IDPedidos = Pedidos.Grabar(Usuarios.Id, mbModifica, TipoDePedido)
                        If IDPedidos <= 0 Then
                            MsgBox "El pedido no pudo modificarse", vbExclamation
                        Else
                            If IDPedidos > 0 And IDOperaciones > 0 Then
                            If TipoDeOperacion = m6Compra Then
                                OkPedido = Pedidos.GrabaAplicaciones(IDPedidos, 0, IDOperaciones)
                            Else
                                OkPedido = Pedidos.GrabaAplicaciones(0, IDPedidos, IDOperaciones)
                            End If
                            If OkPedido > 0 Then
                                IDOperaciones = 0   ' reseteo la operacion
                            Else
                                MsgBox "El pedido no pudo aplicarse a la Orden", vbExclamation
                            End If
                            End If
                            
                            If IDPedidos > 0 And TipoDePedido = m6Ampliacion Then
                                If TraeParametro("UDM", mbInsumos) = "1" Then
                                    Empresa.LeerDatos
                                    If Empresa.NroIdentificadorFiscal = "30-70952682-7" Or Empresa.NroIdentificadorFiscal = "30-50076581-6" Then
                                        Pedidos.CorrigeDespachos IDPedidos
                                    End If
                                End If
                            End If
                            
                            If TraeParametro("ETC", mbInsumos) = "1" Then
                                If IDPedidos > 0 And TipoDeOperacion = m6Compra Then
                                    If MsgBox("Desea enviar un mail con los datos del pedido a todos los Comerciales ?", vbYesNo, Me.Caption) = vbYes Then
                                        EnviaMailTodos "Modificacion"
                                    End If
                                End If
                            Else
                            'P 20140612 09:56 // Pregunta si se quiere enviar un mail al comercial.
                                If IDPedidos > 0 And selComisionistas.Id > 0 Then
                                    Comisionistas.TomaUno selComisionistas.Id
                                    'P 20140612 11:48 //Verifica que el comisionista tenga mail.
                                    If Trim(Comisionistas.mail) <> "" Then
                                        If MsgBox("Desea enviar un mail con los datos del pedido al Comercial", vbYesNo, Me.Caption) = vbYes Then
                                            EnviaMail "Modificacion"
                                        End If
                                    End If
                                End If
                            End If
                            
                        End If
                        txtCantidad.MaxValue = 99999999
                        Me.tabPedidos.ActiveTab = CtrlTab.tabGeneral
                        txtSucursal.SetFocus
                        bdcAbm.Estado = "00001"
                    End If
                    
                    If IDPedidos = -100 Then
                    'If IDPedidos > 0 Then
                        If TipoDePedido = m6Original And sprDivisionPedido.MaxRows > 0 And sprDivisionPedido.MaxRows < 500 Then
                            For b = 1 To sprDivisionPedido.MaxRows
                                If selCuentas.Id <> sprTexto(sprDivisionPedido, colD.IDCuenta, b) And sprTexto(sprDivisionPedido, colD.articulo, b) <> "" Then
                                    Pedidos.GrabaCuerpo IDPedidos, sprTexto(sprDivisionPedido, colD.Id, b), Me.selDestinos.Id, 0, "Divisi�n", sprTexto(sprDivisionPedido, colD.Cantidad, b), 0, m6CuerpoOriginal, 0, sprTexto(sprDivisionPedido, colD.IDCuenta, b)
                                End If
                            Next b
                        End If
                        
                        If TipoDePedido = m6Original And sprDivisionRemitos.MaxRows > 0 And sprDivisionRemitos.MaxRows < 500 Then
                            Dim Remitos As New M6_Remitos
                            Remitos.CadenaDeConexion = CadenaDeConexion
                            
                            For b = 1 To sprDivisionRemitos.MaxRows
                                If Left(sprTexto(sprDivisionRemitos, colD.articulo, b), 7) = "Remito " And CCur2(Mid(sprTexto(sprDivisionRemitos, colD.articulo, b), 7 + 1)) > 0 Then
                                    Remitos.BorraDivididos IDPedidos, Me.selCuentas.Id, Mid(sprTexto(sprDivisionRemitos, colD.articulo, b), 7 + 1)
                                End If
                            Next b
                            Dim IDRemitos As Long
                            For b = 1 To sprDivisionRemitos.MaxRows
                                If sprTexto(sprDivisionRemitos, colD.IDCuenta, b) = Me.selCuentas.Id Then
                                    IDRemitos = sprTexto(sprDivisionRemitos, colD.IDRemitos, b)
                                End If
                                If sprTexto(sprDivisionRemitos, colD.IDCuenta, b) <> "" And sprTexto(sprDivisionRemitos, colD.articulo, b) <> "" Then
                                    Remitos.GrabaCuerpoDeDivididos IDPedidos, IDRemitos, sprTexto(sprDivisionRemitos, colD.Id, b), sprTexto(sprDivisionRemitos, colD.IDCuenta, b), sprTexto(sprDivisionRemitos, colD.Cantidad, b)
                                End If
                            Next b
                            Set Remitos = Nothing
                            
                        End If
                        InicializaOrden     ' vacia los campos del formulario
                    End If
        
            Case "eliminar"
                    If IDPedidos > 0 Then
                        If TraeParametro("UDM", mbInsumos) = True Then
                            Empresa.LeerDatos
                            If Empresa.NroIdentificadorFiscal = "30-70952682-7" Or Empresa.NroIdentificadorFiscal = "30-50076581-6" Then  ' AG
                                If Pedidos.ListaPorSQL("SELECT IDDespachos=ISNULL((SELECT MAX(DC.IDDespachos) FROM M6_DespachosCuerpo DC JOIN M6_Despachos D ON D.ID=DC.IDDespachos WHERE D.Activa=1 AND DC.IDPedidos=" & IDPedidos & "),0)")!IDDespachos > 0 Then
                                    If Pedidos.ListaPorSQL("SELECT Usado=ISNULL((SELECT 1 FROM M6_DespachosCuerpo DC JOIN M6_Despachos D ON D.ID=DC.IDDespachos WHERE D.Activa=1 AND DC.IDPedidos=" & IDPedidos & " AND DC.Cantidad>0 AND (DC.Cantidad>DC.CantidadARemitir OR DC.Cantidad>DC.CantidadAFacturar)),0)")!Usado = 1 Then
                                        MsgBox "No es posible eliminar un pedido que fue incluido en un despacho.", vbInformation, "Despachos"
                                        Exit Sub
                                    End If
                                End If
                            End If
                        End If
                    End If
                    If MsgBox("Confirma la eliminaci�n del pedido ?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
                        IDPedidos = Pedidos.Grabar(Usuarios.Id, mbBaja, TipoDePedido)
                        If IDPedidos > 0 Then
                            Me.tabPedidos.ActiveTab = CtrlTab.tabGeneral
                            bdcAbm.Estado = "00001"
                            txtSucursal.SetFocus
                            txtCantidad.MaxValue = 99999999
                        Else
                            MsgBox "El pedido no pudo eliminarse, posee aplicadaciones de remitos o facturas", vbExclamation
                        End If
                        InicializaOrden
                    End If
            
            Case "cerrar"
                    Unload Me
        End Select

    End Sub

Private Sub bdcDetalle_Click()

    Dim a As Integer
    Dim b As Integer

    If bdcDetalle.Boton = 0 Or bdcDetalle.Boton = 1 Then
        If CCur2(Me.txtPrecio.Text) <= 0 And TraeParametro("POP", mbInsumos) = True Then
            MsgBox "El precio NO puede ser cero.", vbCritical, "Renglon de pedido"
            Exit Sub
        End If
    End If
        
    Select Case bdcDetalle.Boton
           Case 0
                For a = 1 To sprDetalle.MaxRows
                    
                    If CCur2(sprTexto(sprDetalle, colC.IDArticulos, a)) = selArticulos.Id Then
                        If TipoDePedido = m6Original Then
                            a = sprDetalle.MaxRows
                            MsgBox "No se permite agregar dos renglones con el mismo art�culo. Debe modificar el existente", vbExclamation, Me.Caption
                        Else
                            If Me.txtPrecio.Text <> CCur2(sprTexto(sprDetalle, colC.Precio, a)) Then
                                Me.txtPrecio.Text = CCur2(sprTexto(sprDetalle, colC.Precio, a))
                            End If
                        End If
                    
                    End If
                    
                    If Val(sprTexto(sprDetalle, colC.IDArticulos, a)) = 0 Then Exit For
                    
                Next a
                If a < sprDetalle.MaxRows Then
                    If TipoDeOperacion = m6Venta Then
                        If selModalidades.Id = EnumModalidades.m6CtaYOrden 
                        Or selModalidades.Id = EnumModalidades.m6CtaYOrdenCanje 
                        Or selModalidades.Id = EnumModalidades.m6CtaYOrdenCanjeFuturo Then
                            Dim rs1 As Recordset
                            Set rs1 = Pedidos.ListaPosicionComercial(0, selModalidades.Id, 0, Me.selArticulos.Id, "01/01/1900", date, 0, 0, m6ConSaldoaFacturar, False)

                            If Not (rs1.BOF And rs1.EOF) Then
                                If rs1!PendFactura = 0 Then
                                    If MsgBox("No existen compras de cuenta y orden con pendiente de liquidaci�n con este art�culo. Agrega de todas formas?", vbCritical + vbYesNo, "Renglon de pedido") = vbNo Then
                                        Exit Sub
                                    End If
                                End If
                            Else
                                If MsgBox("No existen compras de cuenta y orden con pendiente de liquidaci�n con este art�culo. Agrega de todas formas?", vbCritical + vbYesNo, "Renglon de pedido") = vbNo Then
                                    Exit Sub
                                End If
                            End If

                            Set rs1 = Nothing
                        ElseIf selModalidades.Id <> EnumModalidades.m6Consumo Then
                            If (selModalidades.Id = EnumModalidades.m6CanjeFuturo And selDestinos.Id = 25) 
                            Or (selModalidades.Id <> EnumModalidades.m6CanjeFuturo And selDestinos.Id = 27) Then
                                If MsgBox("No coincide la Modalidad seleccionada '" & selModalidades.Text & "' con el Esquema Impositivo '" & selDestinos.Text & "'. Agrega de todas formas?", vbCritical + vbYesNo, "Renglon de pedido") = vbNo Then
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                    
                    If TipoDePedido = m6Ampliacion Then
                        For b = 1 To sprDetalle.MaxRows
                            If Me.txtCantidad.Text < 0 And CCur2(sprTexto(sprDetalle, colC.IDArticulos, b)) = Me.selArticulos.Id Then
                                If CCur2(sprTexto(sprDetalle, colC.Cantidad, b)) + Me.txtCantidad.Text < CCur2(sprTexto(sprDetalle, 15, b)) Then
                                    MsgBox "No puede reducir a " & CCur2(sprTexto(sprDetalle, colC.Cantidad, b)) + Me.txtCantidad.Text & " este articulo porque la reduccion es menor a lo remitido (" & CCur2(sprTexto(sprDetalle, 15, b)) & ")", vbInformation, "Pedidos"
                                    Exit Sub
                                End If
                                If CCur2(sprTexto(sprDetalle, colC.Cantidad, b)) + Me.txtCantidad.Text < CCur2(sprTexto(sprDetalle, 16, b)) Then
                                    MsgBox "No puede reducir a " & CCur2(sprTexto(sprDetalle, colC.Cantidad, b)) + Me.txtCantidad.Text & " este articulo porque la reduccion es menor a lo facturado (" & CCur2(sprTexto(sprDetalle, 16, b)) & ")", vbInformation, "Pedidos"
                                    Exit Sub
                                End If
                            End If
                        Next b
                    End If
                    LLenaCuerpo a
                End If
                If TipoDePedido = m6Ampliacion Then
                    sprAsignaTexto sprDetalle, colC.articulo, colC.articulo, a + 1, a + 1, "click aqu� para ampliar o reducir", , , , , , mbBloqueoAzul
                End If
                
           Case 1
                a = sprDetalle.ActiveRow
                If (selModalidades.Id = EnumModalidades.m6CanjeFuturo And selDestinos.Id = 25) Or (selModalidades.Id <> EnumModalidades.m6CanjeFuturo And selDestinos.Id = 27) Then
                    If MsgBox("No coincide la Modalidad seleccionada '" & selModalidades.Text & "' con el Esquema Impositivo '" & selDestinos.Text & "'. Agrega de todas formas?", vbCritical + vbYesNo, "Renglon de pedido") = vbNo Then
                        Exit Sub
                    End If
                End If
                LLenaCuerpo a
                
           Case 2
                If IDPedidos > 0 Then
                    Dim rs As Recordset
                    Set rs = Pedidos.ListaCuerpo(IDPedidos, True, True, Val(sprTexto(sprDetalle, colC.IDArticulos, sprDetalle.ActiveRow)))
                    If Not (rs.BOF And rs.EOF) Then
                        If rs!Facturado <> 0 Or rs!Remitido <> 0 Then
                            MsgBox "No se puede eliminar. Este art�culo tiene " & IIf(rs!Facturado <> 0, rs!Facturado & " unidades facturadas", "") & IIf(rs!Facturado <> 0 And rs!Remitido <> 0, " y ", "") & IIf(rs!Remitido <> 0, rs!Remitido & " unidades remitidas", ""), vbCritical, "Renglon de pedido"
                            Exit Sub
                        End If
                    End If
                    Set rs = Nothing
                End If
                sprDetalle.DeleteRows sprDetalle.ActiveRow, 1
    
    End Select

    If TipoDePedido = m6Ampliacion Then
        If txtCantidad.Text >= 0 Then
            sprFormatoCeldas sprDetalle, colC.articulo, colC.Id, a, a, , , , , mbStandardOk
        Else
            sprFormatoCeldas sprDetalle, colC.articulo, colC.Id, a, a, , , , , mbStandardFalla
        End If
    End If

End Sub

Private Sub cmdUltimoO_Click()
    If TipoDeOperacion = m6Compra Then
        Operaciones.IDTiposDeOperaciones = 7
    Else
        Operaciones.IDTiposDeOperaciones = 1
    End If
    txtNumeroO.Text = Operaciones.UltimoNumero(Operaciones.IDTiposDeOperaciones, CCur2(txtSucursalO.Text)) + 1
    txtNumeroO.Text = FormatoRG(txtNumeroO.Text, 8, True, False)
    
    ' txtNumeroO.SetFocus

End Sub

Private Sub bdcCopiar_Click()
    
    Dim a As Integer, b As Integer
    
    sprDetalle.ClearRange -1, -1, -1, -1, True
    For a = 1 To sprDetalleC.MaxRows
        For b = 1 To sprDetalleC.MaxCols
            sprDetalleC.Row = a
            sprDetalle.Row = a
            sprDetalleC.Col = b
            sprDetalle.Col = b
            'If b <> colC.ID Then
                sprDetalle.Text = sprDetalleC.Text
            'End If
        Next b
    Next a
    HabilitarControles fraDetalle, False
    bdcDetalle.Estado = "000"
    txtCantidad.Enabled = True

End Sub

Private Sub bdcCanje_Click()
    Dim Accion As Integer
    Dim OkPedido As Long
    
    If IDOperaciones > 0 Then
        Accion = mbModifica
    Else
        Accion = mbAlta
    End If
    ' grabar orden despues del pedido.
    If txtNumeroO.Text > "00000000" Then
        GenerarOrden Accion
        If IDOperaciones > 0 Then
            If bdcCanje.Valor = "grabar" Then
                MsgBox IIf(TipoDeOperacion = m6Compra, "Orden", "Autorizaci�n") & " de venta n� " & txtSucursalO.Text & "-" & txtNumeroO.Text & " generada", vbExclamation
            Else
                Operaciones.TomaUno IDOperaciones
                DerivadorFAC.ImprimirComprobante Operaciones.IDMovimientos, m3Ordendeventa, m3SinModalidad, mbCereales, Operaciones.Fecha, Operaciones.SucursalImpreso, Operaciones.NumeroImpreso, mbImpresionOriginal, True
            End If
            If IDPedidos > 0 And IDOperaciones > 0 Then
               If TipoDeOperacion = m6Compra Then
                   OkPedido = Pedidos.GrabaAplicaciones(IDPedidos, 0, IDOperaciones)
               Else
                   OkPedido = Pedidos.GrabaAplicaciones(0, IDPedidos, IDOperaciones)
               End If
               If OkPedido > 0 Then
                   IDOperaciones = 0   ' reseteo la operacion
               Else
                   MsgBox "La orden no pudo aplicarse al pedido", vbExclamation
               End If
            End If
        Else
            MsgBox "La " & IIf(TipoDeOperacion = m6Compra, "Orden", "Autorizaci�n") & " de venta no pudo ser generada", vbExclamation
        End If
    Else
        MsgBox "La " & IIf(TipoDeOperacion = m6Compra, "Orden", "Autorizaci�n") & " de venta no pudo ser generada. No tiene Nro", vbExclamation
    End If
End Sub

Private Sub bdcDivision_Click()

    Dim a As Integer, b As Integer

    Dim spr As vaSpread

    If Me.tabDivision.ActiveTab = 0 Then
        Set spr = sprDivisionPedido
    Else
        Set spr = sprDivisionRemitos
    End If
    
    With spr 'DivisionPedido
        
        If bdcDivision.Boton <> 3 And sprTexto(spr, colD.articulo, .ActiveRow) = "" Then Exit Sub
        
        Select Case bdcDivision.Boton
                
               Case 0 'agregar
                    If selCuentas.Id <> selCuentasDivision.Id And Me.txtCantidadDivision.Text > 0 And CCur2(sprTexto(spr, colD.Cantidad, .ActiveRow)) > CCur2(txtCantidadDivision.Text) Then
                        .MaxRows = .MaxRows + 1
                        .InsertRows .ActiveRow + 1, 1
                        .SetText colD.articulo, .ActiveRow + 1, sprTexto(spr, colD.articulo, .ActiveRow)
                        .SetText colD.CantidadTotal, .ActiveRow + 1, sprTexto(spr, colD.CantidadTotal, .ActiveRow)
                        .SetText colD.Cuenta, .ActiveRow + 1, selCuentasDivision.Text
                        .SetText colD.Cantidad, .ActiveRow + 1, txtCantidadDivision.Text
                        .SetText colD.Id, .ActiveRow + 1, sprTexto(spr, colD.Id, .ActiveRow)
                        .SetText colD.IDCuenta, .ActiveRow + 1, selCuentasDivision.Id
                        .SetText colD.IDRemitos, .ActiveRow + 1, sprTexto(spr, colD.IDRemitos, .ActiveRow)
                        '
                        .SetText colD.Cantidad, .ActiveRow, CCur2(sprTexto(spr, colD.Cantidad, .ActiveRow)) - CCur2(txtCantidadDivision.Text)
                    End If
               Case 1 'modificar
                    If selCuentas.Text <> sprTexto(spr, colD.Cuenta, .ActiveRow) Then
                        Dim cantidadanterior As Currency
                        cantidadanterior = sprTexto(spr, colD.Cantidad, .ActiveRow)
                        .SetText colD.Cuenta, .ActiveRow, selCuentasDivision.Text
                        .SetText colD.Cantidad, .ActiveRow, txtCantidadDivision.Text
                        For a = .ActiveRow To 1 Step -1
                            If selCuentas.Text = sprTexto(spr, colD.Cuenta, a) Then
                                .SetText colD.Cantidad, a, CCur2(sprTexto(spr, colD.Cantidad, a)) + cantidadanterior - CCur2(txtCantidadDivision.Text)
                            ElseIf sprTexto(spr, colD.Cuenta, a) = "" Then
                                Exit For
                            End If
                        Next a
                    End If
                    
               Case 2 'eliminar
                    If selCuentas.Text <> sprTexto(spr, colD.Cuenta, .ActiveRow) Then
                        For a = .ActiveRow To 1 Step -1
                            If selCuentas.Text = sprTexto(spr, colD.Cuenta, a) Then
                                .SetText colD.Cantidad, a, CCur2(sprTexto(spr, colD.Cantidad, a)) + CCur2(sprTexto(spr, colD.Cantidad, .ActiveRow))
                            ElseIf sprTexto(spr, colD.Cuenta, a) = "" Then
                                Exit For
                            End If
                        Next a
                        .DeleteRows .ActiveRow, 1
                        .MaxRows = .MaxRows - 1
                    End If
               Case 3 'copiar a remitos
                    Dim Porcentaje As Double
                    
                    For b = 1 To sprDivisionRemitos.MaxRows
                        If selCuentas.Text = sprTexto(sprDivisionRemitos, colD.Cuenta, b) Then
                            sprDivisionRemitos.SetText colD.Cantidad, b, CCur2(sprTexto(sprDivisionRemitos, colD.CantidadTotal, b))
                        ElseIf sprTexto(sprDivisionRemitos, colD.Cuenta, b) <> "" Then
                            sprDivisionRemitos.DeleteRows b, 1
                            sprDivisionRemitos.MaxRows = sprDivisionRemitos.MaxRows - 1
                        End If
                    Next b
                    sprDivisionRemitos.MaxRows = sprDivisionRemitos.MaxRows + 1
                    For a = 1 To sprDivisionPedido.MaxRows
                        If CCur2(sprTexto(sprDivisionPedido, colD.Id, a)) <> 0 Then
                            Porcentaje = CCur2(sprTexto(sprDivisionPedido, colD.Cantidad, a)) / CCur2(sprTexto(sprDivisionPedido, colD.CantidadTotal, a))
                            For b = 1 To sprDivisionRemitos.MaxRows
                                If CCur2(sprTexto(sprDivisionRemitos, colD.Id, b)) = CCur2(sprTexto(sprDivisionPedido, colD.Id, a)) Then
                                    With sprDivisionRemitos
                                        If CCur2(sprTexto(sprDivisionRemitos, colD.IDCuenta, b)) = CCur2(sprTexto(sprDivisionPedido, colD.IDCuenta, a)) Then
                                            .SetText colD.Cantidad, b, CCur2(sprTexto(sprDivisionRemitos, colD.CantidadTotal, b)) * Porcentaje
                                        Else
                                            .SetActiveCell 1, b
                                            .MaxRows = .MaxRows + 1
                                            .InsertRows .ActiveRow + 1, 1
                                            .SetText colD.articulo, .ActiveRow + 1, sprTexto(sprDivisionPedido, colD.articulo, a)
                                            .SetText colD.CantidadTotal, .ActiveRow + 1, sprTexto(sprDivisionRemitos, colD.CantidadTotal, .ActiveRow)
                                            .SetText colD.Cuenta, .ActiveRow + 1, selCuentasDivision.Text
                                            .SetText colD.Cantidad, .ActiveRow + 1, CCur2(sprTexto(sprDivisionRemitos, colD.CantidadTotal, .ActiveRow + 1)) * Porcentaje
                                            .SetText colD.Id, .ActiveRow + 1, sprTexto(sprDivisionPedido, colD.Id, a)
                                            .SetText colD.IDCuenta, .ActiveRow + 1, selCuentasDivision.Id
                                            .SetText colD.IDRemitos, .ActiveRow + 1, sprTexto(sprDivisionPedido, colD.IDRemitos, a)
                                            b = b + 1
                                        End If
                                    End With
                                End If
                            Next b
                        End If
                    Next a
                    sprDivisionRemitos.MaxRows = sprDivisionRemitos.MaxRows - 1
                    Me.tabDivision.ActiveTab = 1
                    Me.bdcDivision.Habilitar , , , 0
        End Select
        
    End With

End Sub

Private Sub cmdUltimo_Click()
    
    NoEntrar = True
    
    txtNumero.Text = Pedidos.UltimoNumero(txtSucursal.Text, IIf(TipoDeOperacion = m6Compra, False, True)) + 1
    txtNumero.Text = FormatoRG(txtNumero.Text, 8, True, False)
    
    NoEntrar = False
    
    txtNumero.SetFocus

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If

End Sub

Private Sub selComentarios_Click()
    
    If chkAgregar.Value = ValueFalse Then
        txtComentario.Text = selComentarios.Text
    Else
        txtComentario.Text = txtComentario.Text & vbLf & selComentarios.Text
    End If

End Sub

Private Sub selComisionistas_Click()

    If selComisionistas.Id = 0 Then
        txtcomision.Text = 0
        txtcomision.Enabled = False
        lblComision.Enabled = False
        imgMail.Visible = False
    Else
        Comisionistas.TomaUno selComisionistas.Id
        txtcomision.Text = Comisionistas.Comision
        txtcomision.Enabled = True
        lblComision.Enabled = True
        
'P 20140612 12:10 //Muestra la imagen indicando que este comisionista tiene mail.
        If Comisionistas.mail <> "" Then
            imgMail.Visible = True
        Else
            imgMail.Visible = False
        End If
    End If
        
End Sub

Private Sub selCobradores_Click()

    If selCobradores.Id = 0 Then
        txtComisionCobrador.Text = 0
        txtComisionCobrador.Enabled = False
        lblComisionCobrador.Enabled = False
    Else
        If TipoDeOperacion = m6Venta And TraeParametro("UVC", mbBase) = True Then
            Comisionistas.TomaUno selCobradores.Id
            txtComisionCobrador.Text = Comisionistas.Comision
        Else
            Cobradores.TomaUno selCobradores.Id
            txtComisionCobrador.Text = Cobradores.Comision
        End If
        txtComisionCobrador.Enabled = True
        lblComisionCobrador.Enabled = True
    End If
        
End Sub

Private Sub selCondicionesComerciales_Click()

    Dim rs As New Recordset

    If Not ParametroFCO Then txtCondiciones.Text = selCondicionesComerciales.Text   'NJE 11/12/2017 - TCK 52389: Agrego el comentario como antes si no estoy manejando la condici�n comercial con este par�metro
    
    'PDM 02/07/2018 12:38 //T.9810 Se agrego que le sume a la Fecha la primera cantidad de dias que la condicion Comercial tiene en detalle.
    Set rs = CondicionesComerciales.ListaDetalle(selCondicionesComerciales.Id)
    If Not (rs.EOF Or rs.BOF) Then
        rs.MoveFirst
        txtFCondicion.Text = CDate(txtFCondicion.Text) + Val(rs.Fields("Dias"))
        'PDM 17/07/2018 12:29 //T.9703
        txtFVencimiento.Text = txtFCondicion.Text
    End If
    
End Sub

Private Sub selCondicionesDeFlete_GotFocus()

    If selCondicionesDeFlete.tag = "1" Then
        selCondicionesDeFlete.LlenaLista CondicionesDeFlete.Lista
        selCondicionesDeFlete.tag = ""
    End If

End Sub

Private Sub selCuentas_Click()
    'PDM 20/09/2017 09:59 //T.31735 Cuando la cuenta esta seleccionada no carga los datos.
    'AlSeleccionarCuenta
    'CargarMemo selCuentas.ID

End Sub

Private Sub selCuentas_LostFocus()
    'PDM 20/09/2017 09:59 //T.31735 Cuando la cuenta esta seleccionada no carga los datos.
    AlSeleccionarCuenta
    CargarMemo selCuentas.Id
    
End Sub


Private Sub AlSeleccionarCuenta()
    
    If NoEntrar = True Then Exit Sub
    
    Dim rs As New Recordset
    Dim a As Integer
    
    EsExtranjero = False
    CuentasCorrientes.TomaUno selCuentas.Id
    
    lblLugarRecepcion.Caption = "Recepci�n"
    CuentasCorrientes.TomaUno selCuentas.Id
'    lblLugarRecepcion.Left = 240
'    If Left(TomaPartePlus(selCuentas.ValorDevuelto, Chr(9), mbUltima), 1) = "5" Then
    If Left(CuentasCorrientes.NroIdentificadorFiscal, 1) = "5" Then
        EsExtranjero = True
'        lblLugarRecepcion.Left = 130
        lblLugarRecepcion.Caption = "Despachante"
        selLugaresDeRecepcion.tag = "o"
    End If
    
    selLugaresDeRecepcion.LlenaLista LugaresDeRecepcion.Lista(selCuentas.Id, mbTodosModulos)
    If selLugaresDeRecepcion.Filas = 0 Then
'        CuentasCorrientes.TomaUno selCuentas.Id
        With rs
            .Fields.Append "ID", adInteger, 4
            .Fields.Append "Descripcion", adVarChar, 50
            .Open
            .AddNew
            !Id = 0
            !Descripcion = CuentasCorrientes.DireccionE
            .Update
            
        End With
        selLugaresDeRecepcion.tag = ""
        selLugaresDeRecepcion.LlenaLista rs
    End If
    selLugaresDeRecepcion_LostFocus
    selSubCuentas.LlenaLista SubCuentasCorrientes.ListaPorCtaCte(selCuentas.Id)
    If selSubCuentas.CantidadDeFilas = -1 Then
        Me.lblSubCuentas.Enabled = False
        selSubCuentas.Enabled = False
    Else
        Me.lblSubCuentas.Enabled = True
        selSubCuentas.Enabled = True
    End If
    Set rs = Nothing
    
    With selComisionistas
        Comisionistas.TomaUnoPorCuentaCorriente selCuentas.Id
        If IDPedidos = 0 Then
            .IDDeDefault = Comisionistas.Id
            .Id = Comisionistas.Id
        End If
        selComisionistas_Click
        If TipoDeOperacion = m6Venta Then If TraeParametro("FCP", mbInsumos) = True Then selComisionistas.Enabled = False 'IFB 31/07/2020 - TCK 26618 - Fija al comisionista en el Pedido
    End With

    With selCobradores
        If TipoDeOperacion = m6Venta And TraeParametro("UVC", mbBase) = True Then
            Comisionistas.TomaUnoPorCuentaCorriente selCuentas.Id
            If IDPedidos = 0 Then
                .IDDeDefault = Comisionistas.Id
                .Id = Comisionistas.Id
            End If
        Else
            Cobradores.TomaUnoPorCuentaCorriente selCuentas.Id
            If IDPedidos = 0 Then
                .IDDeDefault = Cobradores.Id
                .Id = Cobradores.Id
            End If
        End If
        selCobradores_Click
    End With
    'IFB 05/08/2020 - TCK 26618 - Fija al comisionista en el Pedido | Si est� tambi�n el par�metro UVC se tiene que deshabilitar el selector de cobradores
    If TipoDeOperacion = m6Venta Then If TraeParametro("FCP", mbInsumos) = True And TraeParametro("UVC", mbBase) = True Then selCobradores.Enabled = False
    
    If TipoDeOperacion = m6Venta Then
        If TraeParametro("HLP", mbInsumos) = True Then ' Habilita uso de Listas de Precios
            If UsaListaPrecios = True Then   ' Pedidos: Toma Precio de Listas de Precios
                CuentasCorrientes.TomaUno selCuentas.Id
                If CuentasCorrientes.IDTiposDeCuentasCorrientes > 0 Then
                    ListaPrecios.TomaUnoPorTiposDeCtasCtes CuentasCorrientes.IDTiposDeCuentasCorrientes
                    If ListaPrecios.Id > 0 Then
                        Me.selListaPrecios.IDDeDefault = ListaPrecios.Id
                    End If
                End If
            End If
        End If
    End If
    
    LlenaControl sprPedidos, Pedidos.Lista(IIf(TipoDeOperacion = m6Compra, False, True), selCuentas.Id, m6SaldoTodos), "27192025302600"
    sprBloqueaCeldas sprPedidos, colP.Numero, colP.IDPedido, 1, sprPedidos.MaxRows, True, mbBloqueoAmarillo
    
    If Me.selModalidades.Id = 0 And EsExtranjero = True Then
        Me.selModalidades.IDDeDefault = EnumModalidades.m6Exportacion
        Me.selModalidades.Id = EnumModalidades.m6Exportacion
    ElseIf selModalidades.Id > 0 Then
        selModalidades_Click
    End If

End Sub

Private Sub selArticulos_ButtonClick()
    
    frmArticulos.Show
    selArticulos.LlenaLista rsArticulos.Clone
    selArticulos.tag = "1"

End Sub

Private Sub selArticulos_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

    If selArticulos.tag = "1" Then
        selArticulos.LlenaLista rsArticulos.Clone
        selArticulos.tag = ""
    End If

End Sub

' cuando selecciono un articulo para agregar:
Private Sub selArticulos_Click()

    Dim rs As Recordset
    
    If NoEntrar = True Then Exit Sub
        If Val(sprTexto(sprDetalle, colC.Id, sprDetalle.ActiveRow)) = 0 Then
        If TipoDePedido = m6Original And IDPedidos = 0 Then
            bdcDetalle.Estado = "100"
            txtCantidad.Enabled = True
        End If
    End If
    
    'Articulos.TomaUno selArticulos.ID
    'txtPrecioRep.Text = IIf(TraeParametro("PUV", mbInsumos) = False, Articulos.Precio, Articulos.PrecioUV)
    'txtPrecioRep.Text = IIf(TraeParametro("PUV", mbInsumos) = False, Articulos.Precio, Articulos.Precio)
        
    Set rs = Articulos.ListaSaldos(
        selArticulos.Id, 
        IIf(selDepositos.Id >= 0, selDepositos.Id, 0), 
        date, 
        True, 
        False, 
        m3ValorizadoSinValorizar, 
        mbDepositoTodos
        )
    If Not (rs.BOF And rs.EOF) Then
        lblExistenciaFisica.Caption = rs!Fisico
        If rs!Fisico >= 0 Then
            lblExistenciaFisica.ForeColor = mbNegro
        Else
            lblExistenciaFisica.ForeColor = mbStandardFalla
        End If
        lblExistenciaComercial.Caption = rs!Posicion
        If rs!Posicion >= 0 Then
            lblExistenciaComercial.ForeColor = mbNegro
        Else
            lblExistenciaComercial.ForeColor = mbStandardFalla
        End If
    End If

'P 20140822 17:53 //Cargo el precio de Referencia buscando en M0_MovimientosCuerpo el ultimo movimiento para el articulo y la moneda.

    If selMonedas.Id = 1 Then
        lblMoneda.Caption = "Pesos"
    ElseIf selMonedas.Id = 2 Then
        lblMoneda.Caption = "Dolares"
    End If

    Articulos.TomaUnoUltimoPrecio selArticulos.Id, selMonedas.Id
    
'P 20140827 13:02 //Muestra los precios de referencia en los pedidos.

    
    lblUltimaVenta.Caption = "P. Ult. Vta."
    If TipoDeOperacion = m6Compra Or selModalidades.Id = EnumModalidades.m6Consumo Then
        lblMoneda1.Caption = ""
        lblPrecioArticulo.Caption = "P. Art�culo"
        txtPrecioRep.Text = Articulos.Precio
        lblUltimaVenta.Caption = "Ult. Comp."
        txtPrecioReferencial.Text = Articulos.PrecioUC
        If selModalidades.Id = EnumModalidades.m6Consumo Then
            Me.txtPrecio.Text = Articulos.PrecioUC
        Else
            Me.txtPrecio.Text = 0
        End If
        If Articulos.PrecioUC = 0 And Articulos.Precio <> 0 Then
            txtPrecioReferencial.Text = Articulos.Precio
        End If
    ElseIf TipoDeOperacion = m6Venta Then
        
        txtPrecioReferencial.Text = Articulos.PrecioUV
        
        If TraeParametro("PUV", mbInsumos) = False Then
            txtPrecioRep.Text = Articulos.Precio
            lblPrecioArticulo.Caption = "P. Art�culo"
            lblMoneda1.Caption = ""
        Else
            txtPrecioRep.Text = Articulos.PrecioUC
            'lblPrecioArticulo.Caption = "P. Ult. Comp." 'IFB 26/03/2021 - TCK 32577
            lblPrecioArticulo.Caption = "Precio UC"
            lblMoneda1.Caption = lblMoneda.Caption
            If Articulos.PrecioUC = 0 And Articulos.Precio <> 0 Then
                txtPrecioRep.Text = Articulos.Precio
                lblPrecioArticulo.Caption = "P. Art�culo"
                lblMoneda1.Caption = ""
            End If
        End If
    End If
    
    If TipoDeOperacion = m6Compra Or selModalidades.Id = EnumModalidades.m6Consumo Then
        txtPrecioReferencial.Text = Articulos.PrecioUC
    Else
    If TraeParametro("HLP", mbInsumos) = True Then ' Habilita uso de Listas de Precios
        If UsaListaPrecios = True Then   ' Pedidos: Toma Precio de Listas de Precios
            If TipoDeOperacion = m6Venta Then
                If selListaPrecios.Id > 0 And selArticulos.Id > 0 Then
                    Dim ListaPreciosCuerpo As New M6_ListaPreciosCuerpo
                    ListaPreciosCuerpo.CadenaDeConexion = CadenaDeConexion
                    txtPrecio.Text = ListaPreciosCuerpo.TomaPrecioFinal(selListaPrecios.Id, selArticulos.Id)
                End If
            End If
        End If
    Else
        txtPrecio.Text = IIf(TraeParametro("CPR", mbInsumos) = True, txtPrecioReferencial.Text, 0)
    End If
    End If

End Sub


Private Sub selDepositos_Click()
    
    Dim rs As Recordset
    
    If NoEntrar = True Or selArticulos.Id = 0 Then Exit Sub
        
    Set rs = Articulos.ListaSaldos(selArticulos.Id, IIf(selDepositos.Id >= 0, selDepositos.Id, 0), date, True, False, m3ValorizadoSinValorizar, mbDepositoTodos)
    If Not (rs.BOF And rs.EOF) Then
        lblExistenciaFisica.Caption = rs!Fisico
        If rs!Fisico >= 0 Then
            lblExistenciaFisica.ForeColor = mbNegro
        Else
            lblExistenciaFisica.ForeColor = mbStandardFalla
        End If
        lblExistenciaComercial.Caption = rs!Posicion
        If rs!Posicion >= 0 Then
            lblExistenciaComercial.ForeColor = mbNegro
        Else
            lblExistenciaComercial.ForeColor = mbStandardFalla
        End If
    End If


End Sub

Private Sub selExpresadoEn_LostFocus()

    GrabarRegistry "Software\AS\PedidosMonedaExpresado", "MonedaExpresado", selExpresadoEn.Id

End Sub

Private Sub selLugaresDeRecepcion_ButtonClick()

    frmCtasCtesLugDeRecepcion.selCuentasCorrientes.IDDeDefault = selCuentas.Id
    frmCtasCtesLugDeRecepcion.CargaFrm mbInsumos
    
    selLugaresDeRecepcion.tag = "1"

End Sub

Private Sub selLugaresDeRecepcion_Click()

    If selLugaresDeRecepcion.Id > 0 Then
        'LugaresDeRecepcion.TomaUno selLugaresDeRecepcion.ID
        'If txtKmAsfalto.Text = 0 Then txtKmAsfalto.Text = LugaresDeRecepcion.KmAsfalto
        'If txtKmTierra.Text = 0 Then txtKmTierra.Text = LugaresDeRecepcion.KmTierra
    End If

End Sub

Private Sub selLugaresDeRecepcion_LostFocus()
    'PDM 19/08/2021 14:13 //T.35104
    If EsExtranjero Then
        ' El LugaresDeRecepcion corresponde al despachante de aduana. el pais corresponde al Importador
        LugaresDeRecepcion.TomaUno Me.selLugaresDeRecepcion.Id
        If LugaresDeRecepcion.CodigoOficial = "" And ParametroWRH Then
            MsgBox "El codigo de Aduana est� vacio debe actualizar el dato", vbCritical, "ERROR DE ADUANA"
            selLugaresDeRecepcion_ButtonClick
        End If
    End If
End Sub

Private Sub selLugaresDeRecepcion_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

    If selLugaresDeRecepcion.tag = "1" Then
        selLugaresDeRecepcion.LlenaLista LugaresDeRecepcion.Lista(selCuentas.Id, mbTodosModulos)
        selLugaresDeRecepcion.tag = ""
    End If

End Sub

Private Sub selMonedas_LostFocus()
   
    If selMonedas.Id > 0 Then
        Monedas.TomaUno selMonedas.Id
        If Monedas.Predeterminada = True Then
            txtCotizacion.Text = 1
            lblCotizacion.Enabled = False
            txtCotizacion.Enabled = False
            selExpresadoEn.Enabled = False
            'PDM 05/11/2019 10:18 //T.20773
            If selMonedas.Id = 1 Then
                selExpresadoEn.IDDeDefault = 1
                selExpresadoEn.Id = 1
            End If
        Else
            lblCotizacion.Enabled = True
            txtCotizacion.Enabled = True
            txtCotizacion.Text = Monedas.TomaUnaCotizacion(selMonedas.Id, Me.txtFComprobante.Text)
            selExpresadoEn.Enabled = True
        End If
        
        GrabarRegistry "Software\AS\PedidosMoneda", "Moneda", selMonedas.Id
                
    End If

End Sub

Private Sub selProveedores_Click()

    If NoEntrar = True Then Exit Sub
    
    If TipoDePedido = m6Ampliacion Or TipoDeOperacion = m6Venta And selProveedores.Enabled = True Then
        LlenaControl sprPedidosC, Pedidos.Lista(False, selProveedores.Id, m6SaldoTodos), "27192025302600"
        sprBloqueaCeldas sprPedidosC, colP.Numero, colP.IDPedido, 1, sprPedidosC.MaxRows, True, mbBloqueoAmarillo
        sprDetalleC.ClearRange -1, -1, -1, -1, True
        ArmaGrillaCuerpo sprDetalleC
    End If
    
End Sub

Private Sub selModalidades_Click()
    
    If selModalidades.Id = EnumModalidades.m6Directa Or 
       selModalidades.Id = EnumModalidades.m6CtaYOrden Or 
       selModalidades.Id = EnumModalidades.m6CtaYOrdenCanje Or 
       selModalidades.Id = EnumModalidades.m6CtaYOrdenCanjeFuturo Then
        selProveedores.Enabled = True
        lblProveedores.Enabled = True
    ElseIf lblProveedores.Caption <> "Titular" Then
        selProveedores.Enabled = False
        lblProveedores.Enabled = False
        selProveedores.Id = 0
    End If
    
    If TipoDeOperacion = m6Venta And IDPedidos > 0 Then
        If Pedidos.IDModalidades <> selModalidades.Id And TraeParametro("MND", mbInsumos) = False Then
            If Pedidos.IDModalidades = m6Directa Or selModalidades.Id = m6Directa Then
                Dim rs As Recordset
                Set rs = Pedidos.ListaPorSQL("SELECT ID=COUNT(M.ID) FROM M0_Movimientos M JOIN M6_ArticulosAplicaciones AA ON AA.IDMovimientos=M.ID AND AA.IDPedidos=" & Pedidos.Id & " AND M.IDModalidades" & IIf(Pedidos.IDModalidades = m6Directa, "=", "<>") & "6003 WHERE M.Activa=1")
                If Not (rs.BOF And rs.EOF) Then
                    If rs.RecordCount > 0 Then
                        If rs!Id > 0 Then
                            MsgBox "No es posible cambiar la modalidad de facturaci�n, porque ya existen facturan con la modalidad original del pedido"
                            selModalidades.IDDeDefault = Pedidos.IDModalidades
                        End If
                    End If
                End If
                Set rs = Nothing
            End If
        End If
    End If
    
    lblEsquemaImpositivo.Caption = "Esquema impositivo" ' JAM 07/05/2121 T. 33791 Maneja el label cuando es consumo
    sprAsignaTexto Me.sprDetalle, colC.Destino, colC.Destino, SpreadHeader, SpreadHeader, "Esquema impositivo"
                                                        '// 6010
    If TipoDeOperacion = m6Venta And selModalidades.Id = EnumModalidades.m6Consumo Then
        If ErroTipoConsumo = True Then
             MsgBox "Los pedidos de consumos solo pueden utilizarse con una cuenta corriente con el CUIT de la empresa", vbInformation
        End If
        
        With selDestinos
            .CamposVisibles = "Descripcion"
            Dim DestinosNoComerciales As New M6_DestinosNoComerciales
            DestinosNoComerciales.CadenaDeConexion = CadenaDeConexion
            .LlenaLista DestinosNoComerciales.Lista
            Set DestinosNoComerciales = Nothing
            
            lblEsquemaImpositivo.Caption = "Destino no comercial"  ' JAM 07/05/2121 T. 33791 Maneja el label cuando es consumo
            sprAsignaTexto Me.sprDetalle, colC.Destino, colC.Destino, SpreadHeader, SpreadHeader, "Destino no comercial"
            lblUltimaVenta.Caption = "Ult. Comp."
        End With
    Else
        With selDestinos
            .CamposVisibles = "Descripcion"
            .LlenaLista Destinos.Lista(mbInsumos)
        End With
    End If
    
End Sub

Private Function ErroTipoConsumo() As Boolean
    'PDM 06/08/2018 13:34 //T.10597 10608
    If Empresa.NroIdentificadorFiscal = "" Then
        Empresa.LeerDatos
    End If
    CuentasCorrientes.TomaUno Me.selCuentas.Id
    If CuentasCorrientes.NroIdentificadorFiscal <> Empresa.NroIdentificadorFiscal And selModalidades.Id = EnumModalidades.m6Consumo Then
        ErroTipoConsumo = True
    Else
        ErroTipoConsumo = False
    End If
End Function

Private Sub selTransportistas_Click()
    
    If selTransportistas.Id > 0 Then
        With selChoferes
            .Clear
            .CamposVisibles = "Descripcion"
            .LlenaLista Choferes.ListaPorTransportista(selTransportistas.Id)
        End With
    
        With selCamiones
            .Clear
            .CamposVisibles = "Descripcion"
            .LlenaLista Camiones.ListaPorTransportista(selTransportistas.Id)
        End With
    End If

End Sub

Private Sub selCondicionesDeFlete_ButtonClick()
    
    frmCondicionesDeFlete.Show
    selCondicionesDeFlete.LlenaLista CondicionesDeFlete.Lista
    selCondicionesDeFlete.tag = "1"

End Sub

Private Sub sprDetalle_Click(ByVal Col As Long, ByVal Row As Long)

    Screen.MousePointer = vbHourglass
    DoEvents

    HabilitarControles fraDetalle, True
    
    'PDM 04/09/2017 13:31 //T.17784 Ya estba comentada la carga del list en el load form.
    'JAM 14/01/2021 //T.31149 Se habilita el campo, un cliente lo pidi�, se agrega el valor inicial "sin asignar" para que no moleste a los demas.
    'If IDPedidos = 0 Then
    '    selDepositos.Enabled = False
    'End If

    If Not (TipoDePedido = m6Original And IDPedidos > 0) Then
        
        If Val(sprTexto(sprDetalle, colC.Id, Row)) = 0 Then
            If TipoDePedido = m6Original Then
                bdcDetalle.Estado = "100"
                txtCantidad.Enabled = True
            Else
                bdcDetalle.Estado = "000"
                txtCantidad.Enabled = True
            End If
        Else
            If IDPedidos = 0 Then
                bdcDetalle.Estado = "011"
            Else
                bdcDetalle.Estado = "001"
            End If
            txtCantidad.Enabled = True
        End If
    
        If Val(sprTexto(sprDetalle, colC.IDArticulos, Row)) > 0 Then
            If Pedidos.IDTiposDePedidos <> m6Consignacion And IDPedidos > 0 Then
                txtCantidad.MaxValue = CCur(sprTexto(sprDetalle, colC.Cantidad, Row))
            Else
                txtCantidad.MaxValue = 99999999
            End If
            If TipoDePedido = m6Original Or sprDetalle.ForeColor <> 0 Then
                If IDPedidos = 0 Then
                    txtCantidad.Enabled = True
                    bdcDetalle.Estado = "011"
                Else
                    txtCantidad.Enabled = False
                    bdcDetalle.Estado = "011"
                End If
            Else
                bdcDetalle.Estado = "000"
                txtCantidad.Enabled = False
            End If
        Else
            If Pedidos.IDTiposDePedidos <> m6Consignacion And TipoDePedido = m6Original Then
                bdcDetalle.Estado = "000"
                txtCantidad.Enabled = False
            Else
                HabilitarControles fraDetalle, True
                bdcDetalle.Estado = "100"
                txtCantidad.MaxValue = 99999999
                txtCantidad.Enabled = True
            End If
        End If
    
    Else
        If TipoDePedido = m6Original And IDPedidos > 0 Then
            If Val(sprTexto(sprDetalle, colC.IDArticulos, Row)) > 0 Then
                bdcDetalle.Estado = "010"
                txtCantidad.Enabled = False
                selArticulos.Enabled = False
            Else
                bdcDetalle.Estado = "000"
            End If
        End If
    End If

'If Val(sprTexto(sprDetalle, colC.ID, Row)) <> 0 Then
    DoEvents
    LLenaDetalle Row
'End If
    
    Screen.MousePointer = vbDefault
    DoEvents

End Sub

Private Sub sprPedidos_Click(ByVal Col As Long, ByVal Row As Long)
            
    Dim rs As Recordset, a As Integer
        
    IDPedidos = Val(sprTexto(sprPedidos, colP.IDPedido, Row))
    Pedidos.TomaUno IDPedidos
    NoEntrar = True
    CompletarCampos
    NoEntrar = False
    Set rs = Pedidos.ListaAplicaciones(0, IDPedidos, 0)
    If Not (rs.BOF And rs.EOF) Then
        Do While Not rs.EOF
            For a = 1 To sprPedidosC.MaxRows
                If Val(sprTexto(sprPedidosC, colP.IDPedido, a)) = rs!IDPedidosCompra Then
                    sprAsignaTexto sprPedidosC, colP.Asociada, colP.Asociada, a, a, 1
                    LlenaControl sprDetalleC, Pedidos.ListaCuerpo(Val(sprTexto(sprPedidosC, colP.IDPedido, a)) * IIf(selModalidades.Id = EnumModalidades.m6Consumo, -1, 1))
                    sprBloqueaCeldas sprDetalleC, colC.articulo, colC.IDDestinos, 1, sprDetalleC.MaxRows, True, mbBloqueoAmarillo
                    Exit Do
                End If
            Next a
            rs.MoveNext
        Loop
    End If

End Sub

Private Sub sprPedidosC_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    
    If NoEntrar = False Then
        NoEntrar = True
        With sprPedidosC
            If ButtonDown = 1 Then
                .Col = colP.Asociada
                .Row = -1
                .Value = 0
                .Row = Row
                .Value = 1
            End If
        End With
        NoEntrar = False
    End If

End Sub

Private Sub sprPedidosC_Click(ByVal Col As Long, ByVal Row As Long)

    LlenaControl sprDetalleC, Pedidos.ListaCuerpo(Val(sprTexto(sprPedidosC, colP.IDPedido, Row))), "000120030405060708091011"
    sprBloqueaCeldas sprDetalleC, colC.articulo, colC.IDDestinos, 1, sprDetalleC.MaxRows, True, mbBloqueoAmarillo
    
    If IDPedidos = 0 Then
        bdcCopiar_Click
        'LlenaControl sprDetalle, Pedidos.ListaCuerpo(Val(sprTexto(sprPedidosC, colP.IDPedido, Row))), "000120030405060708091011"
        'ArmaGrillaCuerpo sprDetalle
    End If
    
End Sub

Private Sub tabDivision_TabActivate(TabToActivate As Integer)

    If TabToActivate = 0 Then
        Me.bdcDivision.Habilitar , , , 1
    Else
        Me.bdcDivision.Habilitar , , , 0
    End If
    
End Sub

Private Sub tabPedidos_TabActivate(TabToActivate As Integer)

    Dim a As Integer
    Dim Importe As Currency
    
    If TabToActivate = CtrlTab.tabcanje Then
        If selModalidades.Id = EnumModalidades.m6CtaYOrdenCanje Or selModalidades.Id = EnumModalidades.m6CtaYOrdenCanjeFuturo Or _
            selModalidades.Id = EnumModalidades.m6Canje Or selModalidades.Id = EnumModalidades.m6CanjeFuturo Then
            ' obtengo el importe total del pedidos
            For a = 1 To sprDetalle.MaxRows
                If CCur2(sprTexto(sprDetalle, colC.Cantidad, a)) <> 0 Then
                    Importe = Importe + (sprTexto(sprDetalle, colC.Cantidad, a) * sprTexto(sprDetalle, colC.Precio, a))
                End If
            Next a
            txtImporteO.Text = Format(Importe, "######,0#")
        End If
    End If
    
    
End Sub

Private Sub tabPedidos_TabPageShown(ActiveTab As Integer, ActivePage As Integer)
' If TipoDeOperacion = m6Venta And (.IDModalidades = EnumModalidades.m6Directa Or .IDModalidades = EnumModalidades.m6CtaYOrden Or .IDModalidades = EnumModalidades.m6CtaYOrdenCanje Or .IDModalidades = EnumModalidades.m6CtaYOrdenCanjeFuturo) Then
    If ActiveTab = 5 Then
        If selModalidades.Id = EnumModalidades.m6Canje Or selModalidades.Id = EnumModalidades.m6CanjeFuturo Or selModalidades.Id = EnumModalidades.m6CtaYOrdenCanje Or selModalidades.Id = EnumModalidades.m6CtaYOrdenCanjeFuturo Then
            ' activar botones grabar orden
            bdcCanje.Habilitar 1, 1
        Else
            ' desactiva botones grabar orden
            bdcCanje.Habilitar 0, 0
        End If
        If TipoDeOperacion = m6Compra Then
            Me.Frame2.Caption = "Datos de la Orden de Venta"
        Else
            Me.Frame2.Caption = "Datos de la Orden de Compra"
        End If
    End If
End Sub

Private Sub txtCantidad_LostFocus()

    txtImporte.Text = txtCantidad.Value * txtPrecio.Value

End Sub

Private Sub txtComentario_Change()

    If Len(txtComentario.Text) > 3000 Then
        MsgBox "El comentario (" & Len(txtComentario.Text) & ") no puede tener mas de 3000 caracteres de largo"
        Me.txtComentario.SetFocus
    Else
        lblComentario.Caption = "Comentario" & vbLf & vbLf & "(restan: " & 3000 - Len(txtComentario.Text) & " caracteres)"
    End If

End Sub

Private Sub txtFCondicion_LostFocus()
    'PDM 04/11/2019 12:39 //T.21149
    If TraeParametro("MFV", mbInsumos) = True Then
        txtFVencimiento.Text = txtFCondicion.Text
    End If
    
End Sub

Private Sub txtFVencimiento_LostFocus()

    'Me.tabPedidos.ActiveTab = CtrlTab.tabDetalle
    'Me.selArticulos.SetFocus

End Sub

Private Sub txtImporteTotal_LostFocus()

    Me.tabPedidos.ActiveTab = CtrlTab.tabOtros
    Me.selComisionistas.SetFocus

End Sub

Private Sub txtKilos_LostFocus()
    If txtImporteO.Text = 0 Then
        txtImporteO.Text = txtKilos.Text * txtPrecioO.Text / 100
    ElseIf txtKilos.Text = 0 And txtPrecioO.Text > 0 Then
        txtKilos.Text = Format(txtImporteO.Text / (txtPrecioO.Text / 100), "######.0#")
    End If
End Sub

Private Sub txtKmAsfalto_LostFocus()

    txtTarifaAsfalto.Text = TarifasFletes.TomaTarifa(0, txtKmAsfalto.Text)
    txtImporteAsfalto.Text = txtKmAsfalto.Text * txtTarifaAsfalto.Text
    txtKmTotal.Text = Val(txtKmAsfalto.Text) + Val(txtKmTierra.Text)
    txtImporteTotal.Text = CCur(txtImporteTierra.Text) + CCur(txtImporteAsfalto.Text)

End Sub

Private Sub txtKmTierra_LostFocus()
    
    txtTarifaTierra.Text = TarifasFletes.TomaTarifa(0, txtKmTierra.Text, True)
    txtImporteTierra.Text = txtKmTierra.Text * txtTarifaTierra.Text
    txtKmTotal.Text = Val(txtKmAsfalto.Text) + Val(txtKmTierra.Text)
    txtImporteTotal.Text = CCur(txtImporteTierra.Text) + CCur(txtImporteAsfalto.Text)

End Sub

Private Sub txtNumeroO_LostFocus()
    Dim IDOperacionAnterior As Long
    Dim IDPedidoAplicado As Long
    Dim ErrorOrden As Boolean
    
    txtNumeroO.Text = FormatoRG(txtNumeroO.Text, 8, True, False)
    IDOperacionAnterior = IDOperaciones
    ErrorOrden = False
    ' verificar si existe
    IDOperaciones = Operaciones.TomaUnoPorNumeroImpreso(txtSucursalO.Text, txtNumeroO.Text, IIf(TipoDeOperacion = m6Compra, EnumTiposDeOperaciones.m3VentaOrden, EnumTiposDeOperaciones.m3CompraOrden))
    ' busco si esta aplicada
    If IDOperaciones > 0 Then
        If Me.selCuentas.Id = Operaciones.IDCuentasCorrientes Then
            If IDOperaciones <> IDOperacionAnterior And IDOperaciones > 0 Then
                Destinos.TomaUno Operaciones.IDDestinos
                If Destinos.IDTiposDestino <> EnumTiposDeDestino.m3Canje And Destinos.IDTiposDestino <> EnumTiposDeDestino.m3PagoEnEspecies Then
                    MsgBox "La Orden no es de canje", vbCritical, "Orden para relacionar con canje"
                    ErrorOrden = True
                Else
                    IDPedidoAplicado = Pedidos.VerificaAplicaciones(IDOperaciones) ' Aplicado a otro pedido  PedidoCompras  -PedidoVenta
                    If IDPedidoAplicado <> 0 And IDPedidos = 0 Then
                        MsgBox "La Orden ya est� aplicada a otro pedido", vbCritical, "Orden para relacionar con canje"
                        ErrorOrden = True
                    ElseIf Abs(IDPedidoAplicado) <> IDPedidos And IDPedidos <> 0 And IDPedidoAplicado <> 0 Then
                        MsgBox "La Orden ya est� aplicada", vbCritical, "Error en Nro de orden"
                        ErrorOrden = True
                    End If
                End If
            End If
        Else
            MsgBox "La Orden NO es del cliente: " & Me.selCuentas.Text, vbCritical, "Error en la Cuenta Corriente"
            ErrorOrden = True
        End If
        If ErrorOrden = False Then
            CompletarOrden
        Else
            IDOperaciones = 0
            txtNumeroO.Text = "00000000"
        End If
    Else
        MsgBox "No se encontr� La Orden", vbCritical, "Orden para relacionar con canje"
    End If
End Sub

Private Sub txtPrecioO_LostFocus()
    
    If txtImporteO.Text = 0 Then
        txtImporteO.Text = txtKilos.Text * txtPrecioO.Text / 100
    ElseIf txtPrecioO.Text > 0 Then
        txtKilos.Text = Format(txtImporteO.Text / (txtPrecioO.Text / 100), "######.0#")
    End If

End Sub

Private Sub txtTarifaAsfalto_LostFocus()
    
    txtImporteAsfalto.Text = txtKmAsfalto.Text * txtTarifaAsfalto.Text
    txtKmTotal.Text = Val(txtKmAsfalto.Text) + Val(txtKmTierra.Text)
    txtImporteTotal.Text = CCur(txtImporteTierra.Text) + CCur(txtImporteAsfalto.Text)

End Sub

Private Sub txtTarifaTierra_LostFocus()
    
    txtImporteTierra.Text = txtKmTierra.Text * txtTarifaTierra.Text
    txtKmTotal.Text = Val(txtKmAsfalto.Text) + Val(txtKmTierra.Text)
    txtImporteTotal.Text = CCur(txtImporteTierra.Text) + CCur(txtImporteAsfalto.Text)

End Sub

Private Sub txtPrecio_LostFocus()

    txtImporte.Text = txtCantidad.Value * txtPrecio.Value

End Sub

Private Sub txtSucursal_LostFocus()
    
    txtSucursal.Text = FormatoRG(txtSucursal.Text, 4, True, False)
   ' If txtSucursal.Text = "0000" Then
   '     txtSucursal.Text = "0001"
   ' End If
    
End Sub

Private Sub txtNumero_LostFocus()
    
    If NoEntrar Then Exit Sub
    
    txtNumero.Text = FormatoRG(txtNumero.Text, 8, True, False)
    
    IDPedidos = Pedidos.TomaUnoPorNumero(IIf(TipoDeOperacion = m6Compra, False, True), txtSucursal.Text, txtNumero.Text)
    
    If TipoDePedido = m6Ampliacion Then
        If IDPedidos = 0 Then
            'P 20150827 11:04 // Se comento lo de abajo porque hacia un bucle y se colgaba cuando habia otro formulario abierto. T.56530
            'P 20150827 11:04 // Las cuatro lineas de abajo se agregaron para que cuando ponga como numero cero le habilite las solapas para cargar otra ampliacion, de lo contrario tenia que salir y volver a entrar. T.56530
            'txtNumero.SetFocus
            'Exit Sub
            Me.tabPedidos.Tab = CtrlTab.tabGeneral
            Me.tabPedidos.TabState = 0
            Me.tabPedidos.Tab = CtrlTab.tabPedidos
            Me.tabPedidos.TabState = 0
            
        Else
            Me.tabPedidos.Tab = CtrlTab.tabGeneral
            Me.tabPedidos.TabState = 2
            Me.tabPedidos.Tab = CtrlTab.tabFletes
            Me.tabPedidos.TabState = 2
            Me.tabPedidos.Tab = CtrlTab.tabOtros
            Me.tabPedidos.TabState = 2
            Me.tabPedidos.Tab = CtrlTab.tabPedidos
            Me.tabPedidos.TabState = 2
            Me.tabPedidos.Tab = CtrlTab.tabDetalle
            Me.tabPedidos.TabState = 0
            Me.tabPedidos.ActiveTab = CtrlTab.tabDetalle
            Me.tabPedidos.TabsPerRow = 5
            
        End If
    End If
    
    CompletarCampos
    
    If TipoDeOperacion = m6Venta And IDPedidos > 0 Then
        IDOperaciones = Pedidos.TomaAplicaciones(0, IDPedidos)   ' compras
    ElseIf TipoDeOperacion = m6Compra And IDPedidos > 0 Then
        IDOperaciones = Pedidos.TomaAplicaciones(IDPedidos, 0)  ' ventas
    End If
    If IDOperaciones > 0 Then
        Operaciones.TomaUno (IDOperaciones)
        CompletarOrden
    Else
        InicializaOrden
    End If
    
    'PDM 23/10/2019 10:15 //T.20665 Se agrego que para ventas y compras no habilite el ingreso de cantidades.
    If TipoDePedido = m6Original And IDPedidos > 0 And (TipoDeOperacion = m6Venta Or TipoDeOperacion = m6Compra) Then
        txtCantidad.Enabled = False
    End If
    
    
End Sub

Private Sub txtNroInterno_LostFocus()

    If Val(txtNumero.Text) = 0 And Val(txtNroInterno.Text) > 0 Then
        IDPedidos = Pedidos.TomaUnoPorNroInterno(IIf(TipoDeOperacion = m6Compra, False, True), txtNroInterno.Text)
        CompletarCampos
    End If
    
End Sub

                Public Sub CompletarDatos()

                    Dim a As Integer, b As Integer
                    
                    ControlaSeparadorDecimal

                    With Pedidos
                        If IDPedidos = 0 Then
                            .IDFletes = 0
                        End If
                        .IDCuentasCorrientes = selCuentas.Id
                        .IDSubCuentasCorrientes = IIf(selSubCuentas.Id > 0, selSubCuentas.Id, 0)
                        .IDCtasCtesLugaresDeRecepcion = selLugaresDeRecepcion.Id
                        .IDCampanias = selCampanias.Id
                        .IDComprobantes = IIf(TipoDeOperacion = m6Compra, m6PedidoCompra, m6PedidoVenta)
                        .IDModalidades = selModalidades.Id
                        
                        If Not selModalidades.Id = EnumModalidades.m6Consumo Then   'NJE 18/09/2018 - TCK 11535: Cuando se hizo un pedido de consumo, que no guarde esa opci�n para que de entrada no pida ni valide todas las reglas de un pedido de consumo
                            GrabarRegistry "Software\AS\M6IDModalidades", "", selModalidades.Id
                        End If
                        
                        If TipoDeOperacion = m6Venta And (.IDModalidades = EnumModalidades.m6Directa Or .IDModalidades = EnumModalidades.m6CtaYOrden Or .IDModalidades = EnumModalidades.m6CtaYOrdenCanje Or .IDModalidades = EnumModalidades.m6CtaYOrdenCanjeFuturo) Or ParametroWRH = True Then
                            .IDCtaCteProveedor = selProveedores.Id
                            If selProveedores.Enabled = True Then GrabarRegistry "Software\AS\M6IDCtaCteProveedor", "", selProveedores.Id
                        Else
                            .IDCtaCteProveedor = 0
                        End If
                        .IDFormularios = Formularios.TomaUnoPorUsuario(IIf(TipoDeOperacion = m6Compra, m6PedidoCompra, m6PedidoVenta), Usuarios.Id, m6SinModalidad)
                        .IDUnidadesDeNegocio = selUnidadesDeNegocio.Id ' Usuarios.IDUnidadesDeNegocio
                        .IDMonedas = selMonedas.Id
                        .IDExpresadoEn = selExpresadoEn.Id
                        .IDComisionistas = selComisionistas.Id
                        .IDCobradores = selCobradores.Id
                        .Cotizacion = txtCotizacion.Text
                        .Comision = txtcomision.Text
                        .ComisionCobrador = txtComisionCobrador.Text
                        .NroInterno = txtNroInterno.Text
                        
                        'PDM 28/06/2022 12:59 //T.42921
                        '.Fecha = date
                        '.FechaDeAlta = txtFComprobante.Text
                        
                        .Fecha = txtFComprobante.Text
                        If .FechaDeAlta = "00:00:00" Then
                            .FechaDeAlta = date
                        End If
                        
                        'If .FechaDeAlta = "00:00:00" Then
                        '    .FechaDeAlta = date
                        'End If
                        
                        .FechaVencimiento = txtFVencimiento.Text
                        .FechaCondiciones = txtFCondicion.Text
                        .TipoVenta = True
                        .Propio = IIf(TipoDeOperacion = m6Compra, False, True)
                        
                        .Sucursal = txtSucursal.Text
                        
                        If IDPedidos = 0 Then 'JAM 13/8/18 verifico que el pedido numero de pedido no se haya cargado justo desde otra maquina. Horrible pero poco probable que suceda
                            Dim Pedidos1 As New M6_Pedidos
                            Pedidos1.CadenaDeConexion = CadenaDeConexion
                            Pedidos1.TomaUnoPorNumero .Propio, .Sucursal, txtNumero.Text
                            If Pedidos1.Id > 0 Then
                                txtNumero.Text = Pedidos.UltimoNumero(txtSucursal.Text, IIf(TipoDeOperacion = m6Compra, False, True)) + 1
                                txtNumero.Text = FormatoRG(txtNumero.Text, 8, True, False)
                                Pedidos1.TomaUnoPorNumero .Propio, .Sucursal, txtNumero.Text
                                If Pedidos1.Id > 0 Then
                                    MsgBox "No se puede agregar el pedido, el numero este ya fue ingresado."
                                End If
                            End If
                            Set Pedidos1 = Nothing
                        End If
                        
                        .Numero = txtNumero.Text
                        
                        .LugarDeRecepcion = selLugaresDeRecepcion.Text
                        .Condiciones = txtCondiciones.Text
                        .Comentario = txtComentario.Text
                        
                        .IDCondicionesComerciales = selCondicionesComerciales.Id
                        
                        If .Status = "" Then
                            .Status = "00000000000000000000"
                        End If
                        If chkCotizacionFija.Value = ValueTrue Then
                            .Status = Left(.Status, 3) & "1" & Mid(.Status, 5)
                        End If
                        If IDPedidos = 0 Then
                            .IDPedidoAsociado = 0
                        End If
                        For a = 1 To sprPedidosC.MaxRows
                            If Val(sprTexto(sprPedidosC, colP.Asociada, a)) = 1 Then
                                .IDPedidoAsociado = Val(sprTexto(sprPedidosC, colP.IDPedido, a))
                                'Exit Sub
                                Exit For
                            End If
                        Next a
                        .IDTiposDePedidos = selTiposDePedidos.Id
                        .IDCondicionesDeFlete = selCondicionesDeFlete.Id
                        
                        .IDListaPrecios = selListaPrecios.Id
                        
                    End With
                    
                    With Pedidos.rsCuerpo
                        If .State > 0 And .RecordCount > 0 Then
                            .MoveFirst
                            Do While .EOF = False
                                .Delete adAffectCurrent
                                .MoveNext
                            Loop
                        End If
                        For a = 1 To sprDetalle.MaxRows
                            If Val(sprTexto(sprDetalle, colC.IDArticulos, a)) > 0 
                            And (
                                    IDPedidos > 0 Or 
                                    (IDPedidos = 0 And CCur2(sprTexto(sprDetalle, colC.Cantidad, a)) <> 0)
                                ) 
                            Then
                                .AddNew
                                !Descripcion = sprTexto(sprDetalle, colC.Descripcion, a)
                                
                                'If TipoDePedido = m6Original And sprDivisionPedido.MaxRows > 0 And sprDivisionPedido.MaxRows < 500 Then
                                '    For b = 1 To sprDivisionPedido.MaxRows
                                '        If selCuentas.ID = sprTexto(sprDivisionPedido, colD.IDCuenta, b) And Val(sprTexto(sprDetalle, colC.IDArticulos, a)) = Val(sprTexto(sprDivisionPedido, colD.ID, b)) Then
                                '            !Cantidad = CCur2(sprTexto(sprDivisionPedido, colD.Cantidad, b))
                                '        End If
                                '    Next b
                                'Else
                                    !Cantidad = CCur2(sprTexto(sprDetalle, colC.Cantidad, a))
                                'End If
                                
                                !Precio = CCur2(sprTexto(sprDetalle, colC.Precio, a))
                                !IDArticulos = sprTexto(sprDetalle, colC.IDArticulos, a)
                                !IDDepositos = CCur2(sprTexto(sprDetalle, colC.IDDepositos, a))
                                !IDDestinos = sprTexto(sprDetalle, colC.IDDestinos, a)
                                !TipoC = Val(sprTexto(sprDetalle, colC.TipoC, a))
                                !IDAnterior = CCur2(sprTexto(sprDetalle, colC.Id, a))
                                !PrecioReferencial = CCur2(sprTexto(sprDetalle, colC.PrecioReferencial, a))
                                !CantidadAutorizada = CCur2(sprTexto(sprDetalle, colC.CantidadAutorizada, a))
                                
                                .Update
                            End If
                        Next a
                    End With
                    
                    With Pedidos.Fletes
                        .IDComprobantes = IIf(TipoDeOperacion = m6Compra, m6PedidoCompra, m6PedidoVenta)
                        .IDTransportistas = selTransportistas.Id
                        .IDChoferes = selChoferes.Id
                        .IDCamiones = selCamiones.Id
                        .KmTierra = txtKmTierra.Text
                        .TarifaTierra = txtTarifaTierra.Text
                        .KMAsfalto = txtKmAsfalto.Text
                        .TarifaAsfalto = txtTarifaAsfalto.Text
                        .Modulo = mbInsumos
                        .FacturaAlCliente = True
                        .FacturaDelTransportista = True
                        If selTransportistas.Id <= 0 Then
                            .FacturaAlCliente = False
                            .FacturaDelTransportista = False
                            .IDMovimientosEmisor = 0
                            .IDMovimientosReceptor = 0
                        End If
                        If IDPedidos = 0 Then
                            .IDMovimientosEmisor = 0
                            .IDMovimientosReceptor = 0
                        End If
                    End With
                    
                End Sub








                Public Sub CompletarCampos()

                Dim Comisionistas As New M0_Comisionistas

                Comisionistas.CadenaDeConexion = CadenaDeConexion

                    Dim a  As Integer
                    With Pedidos
                        
                        selCuentas.Enabled = True
                        sprDivisionPedido.MaxRows = 0
                        sprDivisionRemitos.MaxRows = 0
                        chkCotizacionFija.Value = ValueFalse
                        
                        If IDPedidos > 0 Then
                            txtNroInterno.Text = IIf(.NroInterno = "", txtNroInterno.Text, .NroInterno)
                            txtSucursal.Text = IIf(.Sucursal = "", txtSucursal.Text, .Sucursal)
                            txtNumero.Text = IIf(.Numero = "", txtNumero.Text, .Numero)
                            selCuentas.IDDeDefault = .IDCuentasCorrientes
                            If Pedidos.TieneMovimientos = True Then
                                selCuentas.Enabled = False
                            End If
                            AlSeleccionarCuenta
                            selSubCuentas.IDDeDefault = .IDSubCuentasCorrientes
                            selLugaresDeRecepcion.IDDeDefault = .IDCtasCtesLugaresDeRecepcion
                            selCampanias.IDDeDefault = .IDCampanias
                            selMonedas.IDDeDefault = .IDMonedas
                            selExpresadoEn.IDDeDefault = .IDExpresadoEn
                            txtCotizacion.Text = .Cotizacion
                            selComisionistas.IDDeDefault = .IDComisionistas
                            'PDM 13/11/2017 16:03 //T.16754 Se modifico para que muestre la condicionn comercial.
                            selCondicionesComerciales.IDDeDefault = IIf(IsNull(.IDCondicionesComerciales), 0, .IDCondicionesComerciales)
                            
                            'P 20140612 14:32 //Muestra la imagen si tiene mail.
                            Comisionistas.TomaUno .IDComisionistas
                            If Comisionistas.mail <> "" Then
                                imgMail.Visible = True
                            Else
                                imgMail.Visible = False
                            End If
                            
                            If Mid(.Status, m6StatusTCFijo, 1) = "1" Then
                                chkCotizacionFija.Value = ValueTrue
                            Else
                                chkCotizacionFija.Value = ValueFalse
                            End If
                            
                            selCobradores.IDDeDefault = .IDCobradores
                            selComisionistas.Id = .IDComisionistas
                            selCobradores_Click 'IFB 22/02/2021 - TCK 31871
                            txtcomision.Text = .Comision
                            txtComisionCobrador.Text = .ComisionCobrador
                            'PDM 29/07/2022 11:09 //T.42921
                            txtFComprobante.Text = .Fecha
                            'txtFComprobante.Text = IIf(.Fecha = 0, date, .Fecha)
                            'txtFComprobante.Text = IIf(.FechaDeAlta = 0, date, .FechaDeAlta)
                            
                            txtFVencimiento.Text = IIf(.FechaVencimiento = 0, date, .FechaVencimiento)
                            txtFCondicion.Text = IIf(.FechaCondiciones = 0, date, .FechaCondiciones)
                            txtCondiciones.Text = .Condiciones
                            txtComentario.Text = .Comentario
                            selTiposDePedidos.IDDeDefault = .IDTiposDePedidos
                            selCondicionesDeFlete.IDDeDefault = .IDCondicionesDeFlete
                            
                            selListaPrecios.IDDeDefault = IIf(IsNull(.IDListaPrecios), 0, .IDListaPrecios)
                            
                            If TipoDePedido = m6Ampliacion Then
                                bdcDetalle.Enabled = True
                                txtCantidad.Enabled = True
                            Else
                                bdcDetalle.Estado = "000"
                                txtCantidad.Enabled = False
                            End If
                        Else
                            txtFComprobante.Text = date
                            txtFVencimiento.Text = date
                            txtFCondicion.Text = date
                        End If
                        
                        
                        If IDPedidos > 0 Then
                            ' Division del pedido
                            Dim rs As Recordset, rs1 As Recordset
                            Set rs = .ListaCuerpo(IDPedidos, True, False, 0, False)
                            LlenaControl sprDetalle, rs, "00010203040506070809101124251415"
                                                        ' 1 2 3 4 5 6 7 8 910111213141516
                                                                    
                            'For a = sprDetalle.MaxRows To 1 Step -1
                            '    If sprTexto(sprDetalle, colC.IDDepositos, a) > 0 Then
                            '        selDepositos.IDDeDefault = Val(sprTexto(sprDetalle, colC.IDDepositos, a))
                            '        selDepositos.ID = selDepositos.ID
                            '        Exit For
                            '    End If
                            'Next a
                                                                                
                                                                    
                            'PDM 28/02/2018 13:28 //T.7004 Se agrego para que traiga el precio del primer articulo de la grilla.
                            If Not (rs.EOF And rs.BOF) Then
                                rs.MoveFirst
                                    'If rs.MaxRows = 0 Then
                                    Articulos.TomaUnoUltimoPrecio rs!IDItems, selMonedas.Id
                                    If TipoDeOperacion = m6Compra Then
                                        lblPrecioArticulo.Caption = "P. Art�culo"
                                        txtPrecioRep.Text = Articulos.Precio
                                    ElseIf TipoDeOperacion = m6Venta Then
                                        If TraeParametro("PUV", mbInsumos) = False Then
                                            txtPrecioRep.Text = Articulos.Precio
                                            lblPrecioArticulo.Caption = "P. Art�culo"
                                            lblMoneda1.Caption = ""
                                        Else
                                            txtPrecioRep.Text = Articulos.PrecioUC
                                            If Articulos.PrecioUC = 0 And Articulos.Precio <> 0 Then
                                                txtPrecioRep.Text = Articulos.Precio
                                            End If
                                        End If
                                    End If
                            'End If
                            End If


                            Set rs1 = .NextRecordset1
                            If Not (rs1.EOF And rs1.BOF) Then
                                rs1.MoveFirst
                                
                                With sprDivisionPedido
                                    .MaxRows = 0
                                    Do While Not rs1.EOF
                                        
                                        .MaxRows = .MaxRows + 1
                                        .SetText colD.articulo, .MaxRows, rs1!ItemsNombre
                                        .SetText colD.CantidadTotal, .MaxRows, rs1!Total
                                        .SetText colD.Cuenta, .MaxRows, rs1!CuentaNombre
                                        .SetText colD.Cantidad, .MaxRows, rs1!Cantidad
                                        .SetText colD.Id, .MaxRows, rs1!IDItems
                                        .SetText colD.IDCuenta, .MaxRows, rs1!IDCuentasCorrientes
                                        rs1.MoveNext
                                        If Not rs1.EOF Then
                                            If sprTexto(sprDivisionPedido, colD.Id, .MaxRows) <> rs1!IDItems Then
                                                .MaxRows = .MaxRows + 1
                                                .Row = .MaxRows
                                                .Col = -1
                                                .BackColor = EnumColores.mbBloqueoGris
                                            End If
                                        End If
                                    Loop
                                End With
                            End If
                            ''''''
                            ' Division de los remitos
                            If 1 = 2 Then
                                Set rs = .ListaRemitosCuerpoAplicados(IDPedidos)
                                LlenaControl sprDivisionRemitos, rs, ""
                                If Not (rs.EOF And rs.BOF) Then
                                    rs.MoveFirst
                                    With sprDivisionRemitos
                                        Dim Ultimo As String
                                        .MaxRows = 0
                                        Do While Not rs.EOF
                                            
                                            If Ultimo <> rs!RemitoNumero Then
                                                .MaxRows = .MaxRows + 1
                                                .SetText colD.articulo, .MaxRows, "Remito " & rs!RemitoNumero
                                                .Row = .MaxRows
                                                .Col = -1
                                                .BackColor = EnumColores.mbAmarilloClaro
                                                Ultimo = rs!RemitoNumero
                                            End If
                                            .MaxRows = .MaxRows + 1
                                            
                                            .SetText colD.articulo, .MaxRows, rs!ItemsNombre
                                            .SetText colD.CantidadTotal, .MaxRows, rs!Total
                                            .SetText colD.Cuenta, .MaxRows, rs!CuentaNombre
                                            .SetText colD.Cantidad, .MaxRows, rs!Cantidad
                                            .SetText colD.Id, .MaxRows, rs!IDItems
                                            .SetText colD.IDCuenta, .MaxRows, rs!IDCuentasCorrientes
                                            .SetText colD.IDRemitos, .MaxRows, rs!IDRemitos
                                            rs.MoveNext
                                            If Not rs.EOF Then
                                                If sprTexto(sprDivisionRemitos, colD.Id, .MaxRows) <> rs!IDItems Then
                                                    .MaxRows = .MaxRows + 1
                                                    .Row = .MaxRows
                                                    .Col = -1
                                                    .BackColor = EnumColores.mbBloqueoGris
                                                End If
                                            End If
                                        Loop
                                        If .MaxRows > 1 Then
                                            .SetSelection -1, 2, -1, 2
                                        End If
                                    End With
                                End If
                                ''''''
                            End If
                        Else
                            Me.sprDetalle.ClearRange -1, -1, -1, -1, True
                        End If
                        
                        ArmaGrillaCuerpo sprDetalle
                        
                        sprPedidosC.ClearRange -1, -1, -1, -1, True
                        sprDetalleC.ClearRange -1, -1, -1, -1, True
                        ArmaGrillaPedidos sprPedidosC
                        ArmaGrillaCuerpo sprDetalleC
                        
                        
                        If TipoDePedido = m6Ampliacion Then
                            For a = 1 To sprDetalle.MaxRows
                                If sprTexto(sprDetalle, colC.articulo, a) = "" Then
                                    sprAsignaTexto sprDetalle, colC.articulo, colC.articulo, a, a, "click aqu� para ampliar o reducir", , , , , , mbBloqueoAzul
                                    Exit For
                                End If
                            Next a
                        End If
                        
                        selModalidades.IDDeDefault = IIf(IDPedidos = 0, Val(LeerRegistry("Software\AS\M6IDModalidades", "")), .IDModalidades)
                        selModalidades_Click
                        If ParametroWRH = True And TipoDeOperacion = m6Venta Then
                            selProveedores.IDDeDefault = .IDCtaCteProveedor
                            selProveedores.Id = .IDCtaCteProveedor
                        Else
                            selProveedores.IDDeDefault = IIf(IDPedidos = 0, Val(LeerRegistry("Software\AS\M6IDCtaCteProveedor", "")), .IDCtaCteProveedor)
                        End If
                        selProveedores_Click
                        
                        'If Pedidos.IDPedidoAsociado > 0 Then
                            'With sprPedidosC
                            '    For a = 1 To .MaxRows
                            '        If sprTexto(sprPedidosC, colP.IDPedido, a) = Pedidos.IDPedidoAsociado Then
                            '            .Col = colP.Asociada
                            '            .Row = a
                            '            .CellType = CellTypeCheckBox
                            '            .TypeCheckCenter = True
                            '            .Text = 1
                            '        Else
                            '            .Text = ""
                            '        End If
                            '    Next a
                            'End With
                        'End If
                        
                        sprDetalle_Click 1, 1
                        
                    End With
                    
                    With Fletes
                        If Pedidos.IDFletes > 0 Then
                            .TomaUno Pedidos.IDFletes
                            selTransportistas.IDDeDefault = .IDTransportistas
                            selChoferes.IDDeDefault = .IDChoferes
                            selCamiones.IDDeDefault = .IDCamiones
                            txtKmTierra.Text = .KmTierra
                            txtTarifaTierra.Text = .TarifaTierra
                            txtKmAsfalto.Text = .KMAsfalto
                            txtTarifaAsfalto.Text = .TarifaAsfalto
                        End If
                        selTransportistas_Click
                    End With
                    
                    If IDPedidos > 0 And TipoDePedido = m6Original Then
                        bdcAbm.Estado = "01111"
                        bdcDetalle.Estado = "010"
                    Else
                        If txtNumero.Text = "00000000" Then
                            bdcAbm.Estado = "00011"
                        Else
                            If TipoDePedido = m6Original Then
                                bdcAbm.Estado = "10001"
                            Else
                                bdcAbm.Estado = "01001"
                            End If
                        End If
                        If IDPedidos > 0 Then
                            bdcDetalle.Estado = "000"
                        Else
                            bdcDetalle.Estado = "111"
                        End If
                    End If
                    txtCantidad.Enabled = True
                    
                    selDepositos_Click

                            
                End Sub

Public Sub CompletarOrden()
'    Dim CampaniasC As New M3_Campanias
'    CampaniasC.CadenaDeConexion = CadenaDeConexion
    
    With Operaciones
        ' Completa datos de la orden
        selMercaderias.Id = .IDMercaderias
        selMercaderias.IDDeDefault = .IDMercaderias
        selDestinosO.Id = .IDDestinos
        selDestinosO.IDDeDefault = .IDDestinos
        selModalidadesO.Id = .IDModalidades
        selModalidadesO.IDDeDefault = .IDModalidades
        selTiposDeCaratulas.Id = .IDTiposDeCaratulas
        selTiposDeCaratulas.IDDeDefault = .IDTiposDeCaratulas
        selPuertos.Id = .IDPuertos
        selCorredores.Id = .IDCorredores
        selComisionistasO.Id = .IDComisionistas
        
        txtSucursalO.Text = .SucursalImpreso
        txtNumeroO.Text = .NumeroImpreso
        txtPrecioO.Text = .Precio
        txtKilos.Text = .Kilos
        txtImporteO = .Precio * .Kilos / 100
        txtFComprobanteO = .Fecha
        txtFechaContratoDesde = .FechaContratoDesde
        txtFechaContratoHasta = .FechaContratoHasta
        txtFechaContratoAcreditacion = .FechaContratoAcreditacion
        txtComisionCorredor = .Comision
      
    End With
    
End Sub

Public Sub ArmaGrillaPedidos(ByVal spr As vaSpread)
    
    With spr
        
        .Redraw = False
        .ProcessTab = True
        .SetRefStyle (2)
        .ArrowsExitEditMode = True
        .EditEnterAction = EditEnterActionRight
        .EditModeReplace = True
        .ColHeaderRows = 1
        .RowHeaderCols = 0
        .UnitType = UnitTypeTwips
        .ScrollBars = ScrollBarsVertical
        .MaxCols = colP.Asociada
        .MaxRows = IIf(.MaxRows = 0, 100, .MaxRows)
        
        .ColWidth(colP.Numero) = 1200
        .ColWidth(colP.Fecha) = 970
        .ColWidth(colP.vencimiento) = 970
        .ColWidth(colP.Condiciones) = 1000
        .ColWidth(colP.tipo) = 1000
        .ColWidth(colP.Comentario) = 800
        .ColWidth(colP.IDPedido) = 0
        
        sprAsignaTexto spr, colP.Numero, colP.Numero, SpreadHeader, SpreadHeader, "N�mero"
        sprAsignaTexto spr, colP.Fecha, colP.Fecha, SpreadHeader, SpreadHeader, "Fecha"
        sprAsignaTexto spr, colP.vencimiento, colP.vencimiento, SpreadHeader, SpreadHeader, "Vencimiento"
        sprAsignaTexto spr, colP.Condiciones, colP.Condiciones, SpreadHeader, SpreadHeader, "Condiciones"
        sprAsignaTexto spr, colP.Comentario, colP.Comentario, SpreadHeader, SpreadHeader, "Comentario"
        sprAsignaTexto spr, colP.tipo, colP.tipo, SpreadHeader, SpreadHeader, "Tipo"
        sprAsignaTexto spr, colP.Asociada, colP.Asociada, SpreadHeader, SpreadHeader, "Asoc."
        
        sprBloqueaCeldas spr, colP.Numero, colP.IDPedido, 1, .MaxRows, True, mbBloqueoAmarillo
        
        If spr.Name = "sprPedidosC" Then
            .ColWidth(colP.Comentario) = 1000
            .ColWidth(colP.Asociada) = 500
            .Col = colP.Asociada
            .Row = -1
            .CellType = CellTypeCheckBox
            .TypeHAlign = TypeHAlignCenter
        Else
            .ColWidth(colP.Comentario) = 1500
            .ColWidth(colP.Asociada) = 0
        End If
        
        .Redraw = True
        
    End With
    
End Sub

Public Sub ArmaGrillaCuerpo(ByVal spr As vaSpread)
    
    Dim a As Integer
    
        With spr
            .Redraw = False
            .ProcessTab = True
            .SetRefStyle (2)
            .ArrowsExitEditMode = True
            .EditEnterAction = EditEnterActionRight
            .EditModeReplace = True
            .ColHeaderRows = 1
            .RowHeaderCols = 0
            .UnitType = UnitTypeTwips
            .ScrollBars = ScrollBarsBoth
            .MaxCols = colC.CantidadAutorizada + 2
            .MaxRows = 100
            'If a = 1 Then .OperationMode = OperationModeExtended
            
            .ColWidth(colC.articulo) = 2950
            .ColWidth(colC.Descripcion) = 700
            .ColWidth(colC.Cantidad) = 1200
            .ColWidth(colC.Precio) = 1200
            .ColWidth(colC.Importe) = 1200
            .ColWidth(colC.Deposito) = 2000
            .ColWidth(colC.Destino) = 2000
            .ColWidth(colC.IDArticulos) = 0
            .ColWidth(colC.IDDepositos) = 0
            .ColWidth(colC.IDDestinos) = 0
            .ColWidth(colC.Id) = 0
            .ColWidth(colC.TipoC) = 0
            .ColWidth(colC.PrecioReferencial) = 0
            .ColWidth(colC.CantidadAutorizada) = 0
            
            sprAsignaTexto spr, colC.articulo, colC.articulo, SpreadHeader, SpreadHeader, "Art�culo"
            sprAsignaTexto spr, colC.Descripcion, colC.Descripcion, SpreadHeader, SpreadHeader, "Descripci�n"
            
            sprAsignaTexto spr, colC.Cantidad, colC.Cantidad, SpreadHeader, SpreadHeader, IIf(spr.Name = "sprDetalle", "Cantidad", "Disponible")
            sprAsignaTexto spr, colC.Precio, colC.Precio, SpreadHeader, SpreadHeader, "Precio"
            sprAsignaTexto spr, colC.Importe, colC.Importe, SpreadHeader, SpreadHeader, "Importe"
            sprAsignaTexto spr, colC.Deposito, colC.Deposito, SpreadHeader, SpreadHeader, "Dep�sito"
            sprAsignaTexto spr, colC.Destino, colC.Destino, SpreadHeader, SpreadHeader, "Esquema impositivo"
            
            sprFormatoNumero spr, colC.Cantidad, colC.Precio, -1, -1, "0", "0", False, True, 4, -1
            sprFormatoNumero spr, colC.Importe, colC.Importe, -1, -1, "0", "0", False, True, 2, -1
            
            sprBloqueaCeldas spr, colC.articulo, colC.IDDestinos, 1, .MaxRows, True, mbBloqueoAmarillo
            
            For a = 1 To .MaxRows
                If Val(sprTexto(spr, colC.TipoC, a)) = 1 Then
                    sprFormatoCeldas spr, colC.articulo, colC.Id, a, a, , , , , mbStandardOk
                ElseIf Val(sprTexto(spr, colC.TipoC, a)) = -1 Then
                    sprFormatoCeldas spr, colC.articulo, colC.Id, a, a, , , , , mbStandardFalla
                End If
            Next a
            
            'NJE 27/08/2018 - TCK 10882: Dejo que el m�ximo sea 500 caracteres
            .Row = 1
            .Row2 = .MaxRows
            .Col = colC.Descripcion
            .Col2 = colC.Descripcion
            .CellType = CellTypeEdit
            .TypeMaxEditLen = 500
            
            .Redraw = True
        End With
    
End Sub

Public Sub ArmaGrillaDivision(spr As vaSpread)
     
    With spr
        .Redraw = False
        .ProcessTab = True
        .SetRefStyle (2)
        .ArrowsExitEditMode = True
        .EditEnterAction = EditEnterActionRight
        .EditModeReplace = True
        .ColHeaderRows = 1
        .RowHeaderCols = 0
        .UnitType = UnitTypeTwips
        .ScrollBars = ScrollBarsVertical
        .MaxCols = colD.IDRemitos
        '.MaxRows = 0
        .OperationMode = OperationModeRow
        .SelBackColor = mbDiaActual

        .ColWidth(colD.articulo) = 2450
        .ColWidth(colD.CantidadTotal) = 1100
        .ColWidth(colD.Cuenta) = 2600
        .ColWidth(colD.Cantidad) = 1100
        .ColWidth(colD.Id) = 0
        .ColWidth(colD.IDCuenta) = 0
        .ColWidth(colD.IDRemitos) = 0
        
        sprAsignaTexto spr, colD.articulo, colD.articulo, SpreadHeader, SpreadHeader, "Art�culo"
        sprAsignaTexto spr, colD.CantidadTotal, colD.CantidadTotal, SpreadHeader, SpreadHeader, "Total"
        sprAsignaTexto spr, colD.Cuenta, colD.Cuenta, SpreadHeader, SpreadHeader, "Cuenta"
        sprAsignaTexto spr, colD.Cantidad, colD.Cantidad, SpreadHeader, SpreadHeader, "Cantidad"
        
        sprBloqueaCeldas spr, colD.articulo, colD.Cantidad, 1, .MaxRows, True, mbBloqueoAmarillo
        
        sprFormatoNumero spr, colD.CantidadTotal, colD.CantidadTotal, -1, -1, "0", "0", False, True, 4, -1
        sprFormatoNumero spr, colD.Cantidad, colD.Cantidad, -1, -1, "0", "0", False, True, 4, -1
        
        .Col = colD.articulo
        .ColMerge = MergeAlways
        .Col = colD.CantidadTotal
        .ColMerge = MergeAlways
        
        .Redraw = True
        
    End With
    
End Sub

Private Sub LLenaCuerpo(ByVal Fila As Integer)
        
    If txtImporte.Text = 0 Then
        txtImporte.Text = CCur(txtCantidad.Text) * CDbl(txtPrecio.Text)
    End If
    
    Articulos.TomaUno selArticulos.Id
    sprAsignaTexto sprDetalle, colC.articulo, colC.articulo, Fila, Fila, Articulos.Nombre
    sprAsignaTexto sprDetalle, colC.Descripcion, colC.Descripcion, Fila, Fila, txtDescripcion.Text
    sprAsignaTexto sprDetalle, colC.Cantidad, colC.Cantidad, Fila, Fila, txtCantidad.Text
    sprAsignaTexto sprDetalle, colC.Precio, colC.Precio, Fila, Fila, txtPrecio.Text
    sprAsignaTexto sprDetalle, colC.Importe, colC.Importe, Fila, Fila, txtImporte.Text
    sprAsignaTexto sprDetalle, colC.Deposito, colC.Deposito, Fila, Fila, selDepositos.Text
    sprAsignaTexto sprDetalle, colC.Destino, colC.Destino, Fila, Fila, selDestinos.Text
    sprAsignaTexto sprDetalle, colC.IDArticulos, colC.IDArticulos, Fila, Fila, Articulos.Id
    sprAsignaTexto sprDetalle, colC.IDDepositos, colC.IDDepositos, Fila, Fila, selDepositos.Id
    sprAsignaTexto sprDetalle, colC.IDDestinos, colC.IDDestinos, Fila, Fila, selDestinos.Id
    sprAsignaTexto sprDetalle, colC.Id, colC.Id, Fila, Fila, 0
    sprAsignaTexto sprDetalle, colC.TipoC, colC.TipoC, Fila, Fila, 0
    sprAsignaTexto sprDetalle, colC.PrecioReferencial, colC.PrecioReferencial, Fila, Fila, txtPrecioReferencial.Text
        
End Sub

'JA 20151106 //Se agrega funci�n para limpiar el detalle. Esto es una de las majoras que podr�a evitar el error de los tikets: / T.96681 - 38552
Private Sub LimpiarDetalle()
    
    DoEvents
   
    txtDescripcion.Text = ""
    txtDescripcion.TextRTF = ""
    txtCantidad.Text = ""
    txtPrecio.Text = ""
    txtImporte.Text = ""
    selArticulos.IDDeDefault = -1
    selArticulos.Text = ""
    selArticulos.ValorInicial = "Ninguno Seleccionado"
    selDepositos.IDDeDefault = 0 'Val(sprTexto(sprDetalle, colC.IDDepositos, Fila))
    selDestinos.IDDeDefault = 0 'Val(sprTexto(sprDetalle, colC.IDDestinos, Fila))
    txtPrecioReferencial.Text = "" 'sprTexto(sprDetalle, colC.PrecioReferencial, Fila)
    
End Sub

Private Sub LLenaDetalle(ByVal Fila As Integer)
    Dim articulo As String

    articulo = sprTexto(sprDetalle, colC.articulo, Fila)
    
    DoEvents
    
    If ((articulo <> "click aqu� para ampliar o reducir") And (articulo <> "")) Then
    'JA 20151106 // Si se seleccion� un articulo en la grilla.  Parte de las mejoras que podr�a evitar el error de los tikets: / T.96681 - 38552
        txtDescripcion.Text = sprTexto(sprDetalle, colC.Descripcion, Fila)
        txtDescripcion.TextRTF = sprTexto(sprDetalle, colC.Descripcion, Fila)
        txtCantidad.Text = sprTexto(sprDetalle, colC.Cantidad, Fila)
        txtPrecio.Text = sprTexto(sprDetalle, colC.Precio, Fila)
        txtImporte.Text = sprTexto(sprDetalle, colC.Importe, Fila)
        selArticulos.IDDeDefault = Val(sprTexto(sprDetalle, colC.IDArticulos, Fila))
        selDepositos.IDDeDefault = Val(sprTexto(sprDetalle, colC.IDDepositos, Fila))
        selDestinos.IDDeDefault = Val(sprTexto(sprDetalle, colC.IDDestinos, Fila))
        txtPrecioReferencial.Text = sprTexto(sprDetalle, colC.PrecioReferencial, Fila)
    Else
    'JA 20151106 // Si NO se seleccion� un articulo en la grilla. Parte de las mejoras que podr�a evitar el error de los tikets: / T.96681 - 38552
        LimpiarDetalle
    End If
        
End Sub

Private Sub GenerarOrden(ByVal Accion As Integer)

    Dim IDcaratulas As Long
    Dim CampaniasM3 As New M3_Campanias
    
    CampaniasM3.CadenaDeConexion = CadenaDeConexion
    Screen.MousePointer = vbHourglass
    
    With Operaciones
        .Numero = txtNumeroO.Text
        .SucursalImpreso = txtSucursalO.Text
        .NumeroImpreso = txtNumeroO.Text
        .Fecha = txtFComprobanteO.Text
        .IDTiposDeOperaciones = IIf(TipoDeOperacion = m6Compra, m3VentaOrden, m3CompraOrden)
        .IDCuentasCorrientes = selCuentas.Id
        .IDMercaderias = selMercaderias.Id
        .IDCampanias = CampaniasM3.TomaPredeterminada
        CampaniasM3.TomaUno .IDCampanias
        If Me.selCampanias.Text <> CampaniasM3.Descripcion Then
            .IDCampanias = CampaniasM3.TomaPorDescripcion(Trim(Me.selCampanias.Text))
        End If
        .IDDestinos = selDestinosO.Id
        .IDCorredores = selCorredores.Id
        .Comision = txtComisionCorredor.Text
        .IDComisionistas = selComisionistasO.Id
        .PorcentajeComisionCom = txtComisionVendedor.Text
        .IDPuertos = selPuertos.Id
        .Precio = txtPrecioO.Text
        .Kilos = txtKilos.Text
        .Porcentaje = 100
        .Factor = 100
        .IDUnidadesDeNegocio = 0
        .IDModalidades = selModalidadesO.Id
        .FechaContratoDesde = txtFechaContratoDesde.Text
        .FechaContratoHasta = txtFechaContratoHasta.Text
        .FechaContratoAcreditacion = txtFechaContratoAcreditacion.Text
        .IDTiposDeCaratulas = Me.selTiposDeCaratulas.Id
        
        GrabarMovimientos
    End With
    
    If chkGeneraCaratula.Value = ValueTrue Or TipoDeOperacion = m6Compra Then
        With Caratulas
            .Numero = txtNumeroO.Text
            .Campo = ""
            .IDUnidadesDeNegocio = 0
            .FechaAlta = txtFComprobanteO.Text
            .FechaFacturacion = txtFechaContratoAcreditacion.Text
            .IDCuentasCorrientes = selCuentas.Id
            .IDSubCuentasCorrientes = 0
            .IDLugaresDeRecepcion = 0
            .IDCampanias = CampaniasM3.TomaPredeterminada
            .IDCorredores = selCorredores.Id
            .IDMercaderias = selMercaderias.Id
            .KilosPactados = txtKilos.Text
            .IDTiposDeCaratulas = selTiposDeCaratulas.Id
            .Status = "A---------"
            '.ContratoComprador = txtContratoComprador.Text
            '.ContratoCorredor = txtContratoCorredor.Text
            IDcaratulas = .ABM(Usuarios.Id, Accion)
        End With
    
        If IDcaratulas > 0 Then
            With Operaciones.rsCaratulas
                .AddNew
                !IDcaratulas = IDcaratulas
                !Kilos = txtKilos.Text
                .Update
            End With
        End If
    End If
    
    IDOperaciones = Operaciones.Grabar(Usuarios.Id, Accion, True)
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub GrabarMovimientos()

    With Operaciones.Movimientos
            
        .CadenaDeConexion = CadenaDeConexion
                
        .LlevaRenglonACero
                
        .MovFechaDelComprobante = txtFComprobanteO.Text
        .MovFechaDeRegistro = date
        
        .MovIDComprobantes = EnumComprobantes.m3Ordendeventa
        .MovIDCuentasCorrientes = selCuentas.Id
        .MovIDSubCuentasCorrientes = 0
        .MovIDModalidades = selModalidades.Id
        .MovDescripcion = ""
        .MovTipoVenta = False
        .MovIDFormularios = Formularios.TomaUnoPorUsuario(m3Ordendeventa, Usuarios.Id, m3SinModalidad)
        .MovStatus = CorrigeStatus(.MovStatus, mbStatusImpresion, mbStatusPendiente)
        .MovFechaDelComprobante = txtFComprobanteO.Text
        
        .MovSucursal = txtSucursalO.Text
        .MovNumero = txtNumeroO.Text
        .MovNumeroInterno = txtNumeroO.Text
        
        '.MovCAI = txtCAI.Text
        '.MovCAIFechaVencimiento = txtFechaVencimientoCAI.Text
        .MovIDUnidadesDeNegocio = Usuarios.IDUnidadesDeNegocio
        .MovUsuario = Usuarios.Id
        .MovComentario = ""
        .MovNroRemito = " "
        .MovIDComisionistas = selComisionistasO.Id
        .MovPorcentajeComisionCom = txtComisionVendedor.Text
        .MovModulo = mbCereales
    
        .LimpiaRenglon
        .CuerpoIDMonedas = selMonedas.Id
        .CuerpoCotizacion = CCur2("0" & txtCotizacion.Text)
        .CuerpoFechaDeVencimiento = txtFVencimiento.Text
        .AgregaRenglon
        
    End With
    
End Sub

Private Sub InicializaOrden()
        
        ' Completa datos de la orden
        selMercaderias.Id = 0
        selMercaderias.IDDeDefault = 0
        selDestinosO.Id = 0
        selDestinosO.IDDeDefault = 0
        selModalidadesO.Id = 0
        selModalidadesO.IDDeDefault = 0
        selTiposDeCaratulas.Id = 0
        selTiposDeCaratulas.IDDeDefault = 0
        selPuertos.Id = 0
        selCorredores.Id = 0
        selComisionistasO.Id = 0
        
        txtSucursalO.Text = "0000"
        txtNumeroO.Text = "00000000"
        txtPrecioO.Text = Format(0, "#####,0#")
        txtKilos.Text = Format(0, "#####0")
        txtImporteO = Format(0, "#####,0#")
        txtFComprobanteO.Text = date
        txtFechaContratoDesde.Text = date
        txtFechaContratoHasta.Text = date
        txtFechaContratoAcreditacion.Text = date
        txtComisionCorredor = 0
        
End Sub

Private Sub EnviaMailTodos(Accion As String)
    
    Dim Usuario As M0_UsuariosCuentasDeMail
    
        
    'obtiene datos del usuario logueado
    Dim UsuariosMail As New M0_UsuariosCuentasDeMail
    
    UsuariosMail.CadenaDeConexion = CadenaDeConexion
    UsuariosMail.TomaUno Usuarios.Id
    Dim Imagen As String
    Imagen = UsuariosMail.Imagen
    If UsuariosMail.LinkImagen <> "" Then
        Imagen = UsuariosMail.LinkImagen
    End If
    
    'obtiene lista de comisionistas
    Dim m As New M0_Comisionistas
    Dim rs As New Recordset
    m.CadenaDeConexion = CadenaDeConexion
    Set rs = m.Lista
    
    'setea datos comunes de todos los emails a enviar
    Set oMail = New clsCDOmail
    Dim Mensaje As String
    With oMail
        .Servidor = UsuariosMail.Servidor
        .Puerto = UsuariosMail.Puerto
        .UseAuntentificacion = UsuariosMail.UseAuntentificacion
        .SSL = UsuariosMail.SSL
        .Usuario = UsuariosMail.Usuario
        .Password = UsuariosMail.Contrase�a
        .Adjunto = ""
        If Accion = "Alta" Then
            .Asunto = Me.Caption & " -- " & Me.txtNumero.Text
        ElseIf Accion = "Modificacion" Then
            .Asunto = "Modificaci�n " & Me.Caption & " -- " & Me.txtNumero.Text
        End If
        .de = UsuariosMail.mail
        
        'itera por los usuarios comerciales/comisionistas y envia emails
        Do While Not rs.EOF
            'usuario.usuario = rs.fields("descripcion").value
            .para = rs.Fields("mail").Value
            
            Mensaje = "<HTML><HEAD><TITLE></TITLE></HEAD><BODY>"
            Mensaje = Mensaje & "<P><DIV ALIGN=center><img src=" & Chr(34) & Imagen & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & " height=" & Chr(34) & UsuariosMail.Alto & Chr(34) & " width=" & Chr(34) & UsuariosMail.Ancho & Chr(34) & "/></DIV>"
            Mensaje = Mensaje & "<HR align=" & Chr(34) & "CENTER" & Chr(34) & " size=" & Chr(34) & "2" & Chr(34) & " width=" & Chr(34) & UsuariosMail.Ancho & Chr(34) & " noshade>"
            Mensaje = Mensaje & "<BR><font size=4><DIV ALIGN=center><STRONG>" & .Asunto & "</STRONG></font><BR><BR></DIV>"
            Mensaje = Mensaje & "<STRONG>Se�or/ar: " & Comisionistas.Descripcion & "</STRONG><BR>"
            Mensaje = Mensaje & "<BR>Por medio de la presente mail, se envian los detalles de la operaci�n realizada por " & "<STRONG>" & Usuarios.Descripcion & "</STRONG><BR><BR>"   ' Por favor responder a este mismo e-mail, dando conformidad del negocio dentro de los 10 d�as, ya que en caso de no recibir observaciones del mismo, se tomara como v�lido y de conformidad para las partes.</font>"
            Mensaje = Mensaje & "<BR><STRONG>Pedido Nro&nbsp;&nbsp;&nbsp;&nbsp; </STRONG>" & Me.txtNumero.Text & "<BR>"
            Mensaje = Mensaje & "<STRONG>Nro. Interno&nbsp;&nbsp;&nbsp; </STRONG>" & txtNroInterno.Text & "<BR>"
            Mensaje = Mensaje & "<STRONG>Cuenta&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </STRONG>" & Left(Me.selCuentas.Text, 45) & "<BR>"
            Mensaje = Mensaje & "<STRONG>Fecha&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </STRONG>" & Me.txtFComprobante.Text & " <BR>"
            Mensaje = Mensaje & "<STRONG>Vencimiento&nbsp;&nbsp; </STRONG>" & Me.txtFVencimiento.Text & "<BR>"
            Mensaje = Mensaje & "<BR>-----------------------------------------------------------<BR><BR>"
            
            Dim Fila As Integer
            'P 20140612 11:46 //Articulos del Pedido.
            For Fila = 1 To sprDetalle.MaxRows
                If sprTexto(sprDetalle, colC.articulo, Fila) <> "" Then
                    Mensaje = Mensaje & "<STRONG>Articulo&nbsp;&nbsp;&nbsp;&nbsp;:&nbsp; </STRONG>" & sprTexto(sprDetalle, colC.articulo, Fila) & "<BR>"
                    Mensaje = Mensaje & "<STRONG>Cantidad&nbsp;&nbsp;:&nbsp; </STRONG>" & sprTexto(sprDetalle, colC.Cantidad, Fila) & "<BR>"
                    Mensaje = Mensaje & "<STRONG>Precio&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;:&nbsp; </STRONG>" & sprTexto(sprDetalle, colC.Precio, Fila) & "<BR>"
                    Mensaje = Mensaje & "<STRONG>Importe&nbsp;&nbsp;&nbsp;&nbsp;:&nbsp; </STRONG>" & sprTexto(sprDetalle, colC.Importe, Fila) & "<BR><BR>"
                End If
            Next

            Mensaje = Mensaje & "<BR><BR><BR><BR><font size=1>Cl�usula de Confidencialidad: Este mensaje, y en su caso, cualquier archivo anexo al mismo, puede contener informaci�n confidencial o legalmente protegida (LOPD 15/1999 de 13 de Diciembre), siendo para uso exclusivo del destinatario. No hay renuncia a la confidencialidad o secreto profesional por cualquier transmisi�n defectuosa o err�nea, y queda expresamente prohibida su divulgaci�n, copia o distribuci�n a terceros sin la autorizaci�n expresa del remitente. Si ha recibido este mensaje por error, se ruega lo notifique al remitente enviando un mensaje al correo electr�nico " & UsuariosMail.mail & " y proceda inmediatamente al borrado del mensaje original y de todas sus copias. Gracias por su colaboraci�n."
            Mensaje = Mensaje & "</BODY></HTML>"
            .Mensaje = Mensaje
            .Enviar_Backup Replace(Imagen, "cid:", ""), True
            
            'siguiente usuario
            rs.MoveNext
        Loop
    
    End With
    
Fin:
    Set oMail = Nothing
    Set UsuariosMail = Nothing
    
    
End Sub


Private Sub EnviaMail(ByVal Accion As String)

    Dim UsuariosMail As New M0_UsuariosCuentasDeMail
    Dim NumeroALetras As New NumeroALetras
    Dim Comisionistas As New M0_Comisionistas
    Dim Fila As Integer
    
    UsuariosMail.CadenaDeConexion = CadenaDeConexion
    Comisionistas.CadenaDeConexion = CadenaDeConexion
    
    UsuariosMail.TomaUno Usuarios.Id
    
    
    Set oMail = New clsCDOmail
    With oMail
         'datos para enviar
         
        .Servidor = UsuariosMail.Servidor
        .Puerto = UsuariosMail.Puerto
        .UseAuntentificacion = UsuariosMail.UseAuntentificacion
        .SSL = UsuariosMail.SSL
        .Usuario = UsuariosMail.Usuario
        .Password = UsuariosMail.Contrase�a
        
        .Adjunto = ""
        
        'P 20140612 14:37 //Si es una Modificaci�n y se envia mail avisa en al asunto y en el Titulo que es una modificacion.
        If Accion = "Alta" Then
            .Asunto = Me.Caption & " -- " & Me.txtNumero.Text
        ElseIf Accion = "Modificacion" Then
            .Asunto = "Modificaci�n " & Me.Caption & " " & Me.txtNumero.Text
        End If
        
        .de = UsuariosMail.mail
        Comisionistas.TomaUno selComisionistas.Id
        
        'P 20140612 11:48 //Verifica que el comisionista tenga mail.
        ' Chequeado antes de llamar a EnviaMail
        '        If Comisionistas.mail = "" Then
        '            MsgBox "No se puede enviar el mail a este comisionista, porque no tiene una direcci�n de mail asignada.", vbCritical
        '            GoTo Fin
        '        End If
        
        .para = Comisionistas.mail   ' "pablo.muller@agrosistemas.com.ar"  'CuentasCorrientes.eMail
        
        Dim Mensaje As String
        Dim Imagen As String
        Imagen = UsuariosMail.Imagen
        If UsuariosMail.LinkImagen <> "" Then
            Imagen = UsuariosMail.LinkImagen
        End If

'P 20140612 11:47 //Titulo del mail.
        Mensaje = "<HTML><HEAD><TITLE></TITLE></HEAD><BODY>"
        Mensaje = Mensaje & "<P><DIV ALIGN=center><img src=" & Chr(34) & Imagen & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & " height=" & Chr(34) & UsuariosMail.Alto & Chr(34) & " width=" & Chr(34) & UsuariosMail.Ancho & Chr(34) & "/></DIV>"
        Mensaje = Mensaje & "<HR align=" & Chr(34) & "CENTER" & Chr(34) & " size=" & Chr(34) & "2" & Chr(34) & " width=" & Chr(34) & UsuariosMail.Ancho & Chr(34) & " noshade>"
        Mensaje = Mensaje & "<BR><font size=4><DIV ALIGN=center><STRONG>" & .Asunto & "</STRONG></font><BR><BR></DIV>"
        
'P 20140612 11:45 //Texto del Asunto.
        Mensaje = Mensaje & "<STRONG>Se�or/ar: " & Comisionistas.Descripcion & "</STRONG><BR>"
        Mensaje = Mensaje & "<BR>Por medio de la presente mail, se envian los detalles de la operaci�n realizada por " & "<STRONG>" & Usuarios.Descripcion & "</STRONG><BR><BR>"   ' Por favor responder a este mismo e-mail, dando conformidad del negocio dentro de los 10 d�as, ya que en caso de no recibir observaciones del mismo, se tomara como v�lido y de conformidad para las partes.</font>"

'P 20140612 11:46 //Detalle del Pedido.
        Mensaje = Mensaje & "<BR><STRONG>Pedido Nro&nbsp;&nbsp;&nbsp;&nbsp; </STRONG>" & Me.txtNumero.Text & "<BR>"
        Mensaje = Mensaje & "<STRONG>Nro. Interno&nbsp;&nbsp;&nbsp; </STRONG>" & txtNroInterno.Text & "<BR>"
        Mensaje = Mensaje & "<STRONG>Cuenta&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </STRONG>" & Left(Me.selCuentas.Text, 45) & "<BR>"
        Mensaje = Mensaje & "<STRONG>Fecha&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </STRONG>" & Me.txtFComprobante.Text & " <BR>"
        Mensaje = Mensaje & "<STRONG>Vencimiento&nbsp;&nbsp; </STRONG>" & Me.txtFVencimiento.Text & "<BR>"
        'P 20150430 10:20 //Se agregaron estos datos al mail. T.30323
        Mensaje = Mensaje & "<BR>"
        Mensaje = Mensaje & "<STRONG>Condici�n Comercial&nbsp;&nbsp;&nbsp; </STRONG>" & selCondicionesComerciales.Text & "<BR>"
        Mensaje = Mensaje & "<STRONG>Comentario Condici�n&nbsp;&nbsp; </STRONG>" & txtCondiciones.Text & "<BR>"
        Mensaje = Mensaje & "<BR>"
        Mensaje = Mensaje & "<STRONG>Leyenda&nbsp;&nbsp; </STRONG>" & selComentarios.Text & "<BR>"
        Mensaje = Mensaje & "<STRONG>Comentario Leyenda&nbsp;&nbsp; </STRONG>" & txtComentario.Text & "<BR>"
        
        
        Mensaje = Mensaje & "<BR>-----------------------------------------------------------<BR><BR>"

'P 20140612 11:46 //Articulos del Pedido.
        For Fila = 1 To sprDetalle.MaxRows
            If sprTexto(sprDetalle, colC.articulo, Fila) <> "" Then
                'Mensaje = Mensaje & sprTexto(sprDetalle, colC.Articulo, Fila) & sprTexto(sprDetalle, colC.Cantidad, Fila) & "-- " & sprTexto(sprDetalle, colC.Precio, Fila) & "-- " & sprTexto(sprDetalle, colC.Importe, Fila) & "<BR>"
                Mensaje = Mensaje & "<STRONG>Articulo&nbsp;&nbsp;&nbsp;&nbsp;:&nbsp; </STRONG>" & sprTexto(sprDetalle, colC.articulo, Fila) & "<BR>"
                Mensaje = Mensaje & "<STRONG>Cantidad&nbsp;&nbsp;:&nbsp; </STRONG>" & sprTexto(sprDetalle, colC.Cantidad, Fila) & "<BR>"
                Mensaje = Mensaje & "<STRONG>Precio&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;:&nbsp; </STRONG>" & sprTexto(sprDetalle, colC.Precio, Fila) & "<BR>"
                Mensaje = Mensaje & "<STRONG>Importe&nbsp;&nbsp;&nbsp;&nbsp;:&nbsp; </STRONG>" & sprTexto(sprDetalle, colC.Importe, Fila) & "<BR><BR>"
            End If
        Next

'P 20140612 11:46 //Usuario que cargo el pedido.
        'Mensaje = Mensaje & "<STRONG>Solicitado por   : </STRONG>" & Usuarios.Descripcion & "<BR><BR>"

'P 20140612 11:47 //Pie del mail.
        Mensaje = Mensaje & "<BR><BR><BR><BR><font size=1>Cl�usula de Confidencialidad: Este mensaje, y en su caso, cualquier archivo anexo al mismo, puede contener informaci�n confidencial o legalmente protegida (LOPD 15/1999 de 13 de Diciembre), siendo para uso exclusivo del destinatario. No hay renuncia a la confidencialidad o secreto profesional por cualquier transmisi�n defectuosa o err�nea, y queda expresamente prohibida su divulgaci�n, copia o distribuci�n a terceros sin la autorizaci�n expresa del remitente. Si ha recibido este mensaje por error, se ruega lo notifique al remitente enviando un mensaje al correo electr�nico " & UsuariosMail.mail & " y proceda inmediatamente al borrado del mensaje original y de todas sus copias. Gracias por su colaboraci�n."
        Mensaje = Mensaje & "</BODY></HTML>"
        
        .Mensaje = Mensaje
        
'P 20140612 11:47 //Envia el mail.
        .Enviar_Backup Replace(Imagen, "cid:", ""), True
    
    End With
    
Fin:
    Set oMail = Nothing
    Set UsuariosMail = Nothing
    Set NumeroALetras = Nothing
    
End Sub
' envio completo
Private Sub oMail_EnvioCompleto()
    MsgBox "Mensaje enviado", vbInformation, "Operaciones"
End Sub
' error al enviar
Private Sub oMail_Error(Descripcion As String, Numero As Variant)
    MsgBox Descripcion, vbCritical, "Operaciones" & Numero
End Sub

Private Function ValidarFechas() As Boolean
'NJE 01/03/2017: Creaci�n
Dim Msj As String

    Msj = ""
    ValidarFechas = True

    If Not IsDate(txtFComprobante.Text) Then
        Msj = Msj & vbCrLf & "La fecha del comprobante es inv�lida"
    Else
        If CDate(txtFComprobante.Text) < CDate("01/01/1900") Then
            Msj = Msj & vbCrLf & "La fecha del comprobante es inv�lida"
        End If
    End If
    
    If Not IsDate(txtFCondicion.Text) Then
        Msj = Msj & vbCrLf & "La fecha de vencimiento de la condici�n es inv�lida"
    Else
        If CDate(txtFCondicion.Text) < CDate("01/01/1900") Then
            Msj = Msj & vbCrLf & "La fecha de vencimiento de la condici�n es inv�lida"
        End If
    End If
    
    If Not IsDate(txtFVencimiento.Text) Then
        Msj = Msj & vbCrLf & "La fecha de vencimiento del comprobante es inv�lida"
    Else
        If CDate(txtFVencimiento.Text) < CDate("01/01/1900") Then
            Msj = Msj & vbCrLf & "La fecha de vencimiento del comprobante es inv�lida"
        End If
    End If
    
    If Not IsDate(txtFComprobanteO.Text) Then
        Msj = Msj & vbCrLf & "La fecha de registraci�n del comprobante de origen es inv�lida"
    Else
        If CDate(txtFComprobanteO.Text) < CDate("01/01/1900") Then
            Msj = Msj & vbCrLf & "La fecha de registraci�n del comprobante de origen es inv�lida"
        End If
    End If
    
    If Not IsDate(txtFechaContratoDesde.Text) Then
        Msj = Msj & vbCrLf & "La fecha de plazo desde del contrato es inv�lida"
    Else
        If CDate(txtFechaContratoDesde.Text) < CDate("01/01/1900") Then
            Msj = Msj & vbCrLf & "La fecha de plazo desde del contrato es inv�lida"
        End If
    End If
    
    If Not IsDate(txtFechaContratoHasta.Text) Then
        Msj = Msj & vbCrLf & "La fecha de plazo hasta del contrato es inv�lida"
    Else
        If CDate(txtFechaContratoHasta.Text) < CDate("01/01/1900") Then
            Msj = Msj & vbCrLf & "La fecha de plazo hasta del contrato es inv�lida"
        End If
    End If
    
    If Not IsDate(txtFechaContratoAcreditacion.Text) Then
        Msj = Msj & vbCrLf & "La fecha de acreditaci�n del contrato es inv�lida"
    Else
        If CDate(txtFechaContratoAcreditacion.Text) < CDate("01/01/1900") Then
            Msj = Msj & vbCrLf & "La fecha de acreditaci�n del contrato es inv�lida"
        End If
    End If
    
    If Msj <> "" Then
        ValidarFechas = False
        Msj = "Se encontraron los siguientes inconvenientes:" & Msj
        MsgBox Msj, vbExclamation, Me.Caption
    End If
End Function

'IFB 19/03/2021 - TCK 32577
Private Sub cmdBusqueda_Click()

    If txtBusqueda.Visible = False Then
        txtBusqueda.Top = selArticulos.Top
        txtBusqueda.Left = selArticulos.Left
        txtBusqueda.Width = selArticulos.Width
        txtBusqueda.Visible = True
        
        lstBusqueda.Top = txtBusqueda.Top + 3250
        lstBusqueda.Left = txtBusqueda.Left + 350
        lstBusqueda.Width = txtBusqueda.Width
        lstBusqueda.ScrollBarV = ScrollBarVShow
        
        lstBusqueda.Height = 2700
        
        txtBusqueda.ZOrder (0)
        lstBusqueda.ZOrder (0)
        
        lstBusqueda.Visible = True
        txtBusqueda.SetFocus
        txtBusqueda.OnFocusNoSelect = True
        txtBusqueda.Text = ""
        lstBusqueda.Clear

        lstBusqueda.AddItem ""
        lstBusqueda.AddItem "Escriba tres letras para iniciar la b�squeda."
        lstBusqueda.AddItem "No distingue entre may�sculas y min�sculas."
        
    Else
        txtBusqueda.Visible = False
        lstBusqueda.Visible = False
    End If
End Sub

Private Sub lstBusqueda_Click()
    selArticulos.IDDeDefault = lstBusqueda.ItemData(lstBusqueda.ListIndex)
    selArticulos.SetFocus
    txtBusqueda.Visible = False
    lstBusqueda.Visible = False
    If lstBusqueda.ItemData(lstBusqueda.ListIndex) > 0 Then selArticulos_Click 'IFB 08/04/2021 - TCK 33080
End Sub

Private Sub txtBusqueda_Change()
    Dim a As Long
    Dim Encontro As Boolean
    
    lstBusqueda.Clear
    If Len(txtBusqueda.Text) > 2 Then
        rsArticulos.MoveFirst
        Do While Not rsArticulos.EOF
            If (UCase(rsArticulos!NombreP) Like "*" & UCase(txtBusqueda.Text) & "*") Then Encontro = True
            If Encontro = True Then
                lstBusqueda.AddItem rsArticulos!NombreP & IIf(rsArticulos!Alias <> "", "   (" & rsArticulos!NombreP & ")", "")
                lstBusqueda.Row = lstBusqueda.NewIndex
                lstBusqueda.ItemData(lstBusqueda.NewIndex) = rsArticulos!Id
            End If
            rsArticulos.MoveNext
            Encontro = False
        Loop

    Else
        lstBusqueda.AddItem ""
        lstBusqueda.AddItem "Escriba tres letras para iniciar la b�squeda."
        lstBusqueda.AddItem "No distingue entre may�sculas y min�sculas."
    End If

End Sub

Private Sub txtBusqueda_LostFocus()
    If Me.ActiveControl.Name <> "lstBusqueda" Then
        If txtBusqueda.Visible Then
            txtBusqueda.Visible = False
            lstBusqueda.Visible = False
        End If
    End If
End Sub
