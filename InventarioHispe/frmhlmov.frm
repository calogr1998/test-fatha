VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Begin VB.Form frmhlmov 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ayuda de Movimientos"
   ClientHeight    =   7200
   ClientLeft      =   15
   ClientTop       =   1155
   ClientWidth     =   11925
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7200
   ScaleWidth      =   11925
   Begin Threed.SSPanel SSPanel1 
      Height          =   6585
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   11760
      _Version        =   65536
      _ExtentX        =   20743
      _ExtentY        =   11615
      _StockProps     =   15
      BackColor       =   -2147483648
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Begin MSMask.MaskEdBox txtfecha 
         Height          =   285
         Left            =   180
         TabIndex        =   6
         Top             =   6165
         Visible         =   0   'False
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtope 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         TabIndex        =   2
         ToolTipText     =   "Simplifique la consulta digitando el Simbolo (*)"
         Top             =   6165
         Visible         =   0   'False
         Width           =   2445
      End
      Begin VB.Data DataAyuda 
         Appearance      =   0  'Flat
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   315
         Left            =   1800
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2115
         Visible         =   0   'False
         Width           =   1155
      End
      Begin TrueDBGrid70.TDBGrid datagrid 
         Bindings        =   "frmhlmov.frx":0000
         Height          =   5415
         Left            =   90
         TabIndex        =   1
         Top             =   90
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   9551
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Nº Mov."
         Columns(0).DataField=   "F4NUMMOV"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Proveedor"
         Columns(1).DataField=   "F4NOMPRV"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "T/D"
         Columns(2).DataField=   "F4TIPDOC"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Serie"
         Columns(3).DataField=   "F4SERDOC"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Nº Doc."
         Columns(4).DataField=   "F4NUMDOC"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Fecha"
         Columns(5).DataField=   "F4FECHA"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Mon"
         Columns(6).DataField=   "F4MONEDA"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "B.I."
         Columns(7).DataField=   "F4BASIMP"
         Columns(7).NumberFormat=   "Standard"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Exonerado"
         Columns(8).DataField=   "F4MONINA"
         Columns(8).NumberFormat=   "Standard"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "I.G.V"
         Columns(9).DataField=   "F4IGV"
         Columns(9).NumberFormat=   "Standard"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "Total"
         Columns(10).DataField=   "F4TOTAL"
         Columns(10).NumberFormat=   "Standard"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "Concepto"
         Columns(11).DataField=   "F4REFERE"
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   0
         Columns(12)._MaxComboItems=   5
         Columns(12).Caption=   "Codprov"
         Columns(12).DataField=   "F4CODPRV"
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(13)._VlistStyle=   0
         Columns(13)._MaxComboItems=   5
         Columns(13).Caption=   "Usuario"
         Columns(13).DataField=   "F4USUARIOING"
         Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(14)._VlistStyle=   0
         Columns(14)._MaxComboItems=   5
         Columns(14).Caption=   "F4CODIGV"
         Columns(14).DataField=   "F4CODIGV"
         Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   15
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   2
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0).ScrollBars=   3
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=15"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1349"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1270"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=74272"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=4657"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=4577"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).AllowSizing=0"
         Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=74272"
         Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(15)=   "Column(2).Width=661"
         Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=582"
         Splits(0)._ColumnProps(18)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(19)=   "Column(2).AllowSizing=0"
         Splits(0)._ColumnProps(20)=   "Column(2)._ColStyle=74272"
         Splits(0)._ColumnProps(21)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(22)=   "Column(3).Width=900"
         Splits(0)._ColumnProps(23)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(24)=   "Column(3)._WidthInPix=820"
         Splits(0)._ColumnProps(25)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(26)=   "Column(3).AllowSizing=0"
         Splits(0)._ColumnProps(27)=   "Column(3)._ColStyle=74272"
         Splits(0)._ColumnProps(28)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(29)=   "Column(4).Width=1799"
         Splits(0)._ColumnProps(30)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(4)._WidthInPix=1720"
         Splits(0)._ColumnProps(32)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(33)=   "Column(4).AllowSizing=0"
         Splits(0)._ColumnProps(34)=   "Column(4)._ColStyle=74272"
         Splits(0)._ColumnProps(35)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(36)=   "Column(5).Width=1773"
         Splits(0)._ColumnProps(37)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(5)._WidthInPix=1693"
         Splits(0)._ColumnProps(39)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(40)=   "Column(5).AllowSizing=0"
         Splits(0)._ColumnProps(41)=   "Column(5)._ColStyle=74272"
         Splits(0)._ColumnProps(42)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(43)=   "Column(6).Width=661"
         Splits(0)._ColumnProps(44)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(45)=   "Column(6)._WidthInPix=582"
         Splits(0)._ColumnProps(46)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(47)=   "Column(6).AllowSizing=0"
         Splits(0)._ColumnProps(48)=   "Column(6)._ColStyle=74272"
         Splits(0)._ColumnProps(49)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(50)=   "Column(7).Width=1535"
         Splits(0)._ColumnProps(51)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(52)=   "Column(7)._WidthInPix=1455"
         Splits(0)._ColumnProps(53)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(54)=   "Column(7).AllowSizing=0"
         Splits(0)._ColumnProps(55)=   "Column(7)._ColStyle=74274"
         Splits(0)._ColumnProps(56)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(57)=   "Column(8).Width=1799"
         Splits(0)._ColumnProps(58)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(59)=   "Column(8)._WidthInPix=1720"
         Splits(0)._ColumnProps(60)=   "Column(8)._EditAlways=0"
         Splits(0)._ColumnProps(61)=   "Column(8).AllowSizing=0"
         Splits(0)._ColumnProps(62)=   "Column(8)._ColStyle=74274"
         Splits(0)._ColumnProps(63)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(64)=   "Column(9).Width=1588"
         Splits(0)._ColumnProps(65)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(66)=   "Column(9)._WidthInPix=1508"
         Splits(0)._ColumnProps(67)=   "Column(9)._EditAlways=0"
         Splits(0)._ColumnProps(68)=   "Column(9).AllowSizing=0"
         Splits(0)._ColumnProps(69)=   "Column(9)._ColStyle=74274"
         Splits(0)._ColumnProps(70)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(71)=   "Column(10).Width=2011"
         Splits(0)._ColumnProps(72)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(73)=   "Column(10)._WidthInPix=1931"
         Splits(0)._ColumnProps(74)=   "Column(10)._EditAlways=0"
         Splits(0)._ColumnProps(75)=   "Column(10).AllowSizing=0"
         Splits(0)._ColumnProps(76)=   "Column(10)._ColStyle=74274"
         Splits(0)._ColumnProps(77)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(78)=   "Column(11).Width=5609"
         Splits(0)._ColumnProps(79)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(80)=   "Column(11)._WidthInPix=5530"
         Splits(0)._ColumnProps(81)=   "Column(11)._EditAlways=0"
         Splits(0)._ColumnProps(82)=   "Column(11).AllowSizing=0"
         Splits(0)._ColumnProps(83)=   "Column(11)._ColStyle=74272"
         Splits(0)._ColumnProps(84)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(85)=   "Column(12).Width=3228"
         Splits(0)._ColumnProps(86)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(87)=   "Column(12)._WidthInPix=3149"
         Splits(0)._ColumnProps(88)=   "Column(12)._EditAlways=0"
         Splits(0)._ColumnProps(89)=   "Column(12).AllowSizing=0"
         Splits(0)._ColumnProps(90)=   "Column(12)._ColStyle=74272"
         Splits(0)._ColumnProps(91)=   "Column(12).Visible=0"
         Splits(0)._ColumnProps(92)=   "Column(12).Order=13"
         Splits(0)._ColumnProps(93)=   "Column(13).Width=3228"
         Splits(0)._ColumnProps(94)=   "Column(13).DividerColor=0"
         Splits(0)._ColumnProps(95)=   "Column(13)._WidthInPix=3149"
         Splits(0)._ColumnProps(96)=   "Column(13)._EditAlways=0"
         Splits(0)._ColumnProps(97)=   "Column(13).AllowSizing=0"
         Splits(0)._ColumnProps(98)=   "Column(13)._ColStyle=74272"
         Splits(0)._ColumnProps(99)=   "Column(13).Visible=0"
         Splits(0)._ColumnProps(100)=   "Column(13).Order=14"
         Splits(0)._ColumnProps(101)=   "Column(14).Width=2699"
         Splits(0)._ColumnProps(102)=   "Column(14).DividerColor=0"
         Splits(0)._ColumnProps(103)=   "Column(14)._WidthInPix=2619"
         Splits(0)._ColumnProps(104)=   "Column(14)._EditAlways=0"
         Splits(0)._ColumnProps(105)=   "Column(14)._ColStyle=65808"
         Splits(0)._ColumnProps(106)=   "Column(14).Visible=0"
         Splits(0)._ColumnProps(107)=   "Column(14).Order=15"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   12632256
         RowDividerColor =   12632256
         RowSubDividerColor=   12632256
         DirectionAfterEnter=   1
         MaxRows         =   250000
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=0,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&H8000000F&"
         _StyleDefs(11)  =   ":id=2,.fgcolor=&H80000012&,.bold=0,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(12)  =   ":id=2,.strikethrough=0,.charset=0"
         _StyleDefs(13)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(15)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(16)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(17)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(18)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(19)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(20)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(21)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(22)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(23)  =   "Splits(0).Style:id=43,.parent=1"
         _StyleDefs(24)  =   "Splits(0).CaptionStyle:id=52,.parent=4"
         _StyleDefs(25)  =   "Splits(0).HeadingStyle:id=44,.parent=2"
         _StyleDefs(26)  =   "Splits(0).FooterStyle:id=45,.parent=3"
         _StyleDefs(27)  =   "Splits(0).InactiveStyle:id=46,.parent=5"
         _StyleDefs(28)  =   "Splits(0).SelectedStyle:id=48,.parent=6"
         _StyleDefs(29)  =   "Splits(0).EditorStyle:id=47,.parent=7"
         _StyleDefs(30)  =   "Splits(0).HighlightRowStyle:id=49,.parent=8"
         _StyleDefs(31)  =   "Splits(0).EvenRowStyle:id=50,.parent=9"
         _StyleDefs(32)  =   "Splits(0).OddRowStyle:id=51,.parent=10"
         _StyleDefs(33)  =   "Splits(0).RecordSelectorStyle:id=53,.parent=11"
         _StyleDefs(34)  =   "Splits(0).FilterBarStyle:id=54,.parent=12"
         _StyleDefs(35)  =   "Splits(0).Columns(0).Style:id=28,.parent=43,.alignment=0,.valignment=1"
         _StyleDefs(36)  =   ":id=28,.locked=-1"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=44,.alignment=2"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=45"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=47"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=43,.alignment=0,.valignment=1"
         _StyleDefs(41)  =   ":id=32,.locked=-1"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=44,.alignment=2"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=45"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=47"
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=58,.parent=43,.alignment=0,.valignment=1"
         _StyleDefs(46)  =   ":id=58,.locked=-1"
         _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=44,.alignment=2"
         _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=45"
         _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=47"
         _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=62,.parent=43,.alignment=0,.valignment=1"
         _StyleDefs(51)  =   ":id=62,.locked=-1"
         _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=44,.alignment=2"
         _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=45"
         _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=47"
         _StyleDefs(55)  =   "Splits(0).Columns(4).Style:id=66,.parent=43,.alignment=0,.valignment=1"
         _StyleDefs(56)  =   ":id=66,.locked=-1"
         _StyleDefs(57)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=44,.alignment=2"
         _StyleDefs(58)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=45"
         _StyleDefs(59)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=47"
         _StyleDefs(60)  =   "Splits(0).Columns(5).Style:id=70,.parent=43,.alignment=0,.valignment=1"
         _StyleDefs(61)  =   ":id=70,.locked=-1"
         _StyleDefs(62)  =   "Splits(0).Columns(5).HeadingStyle:id=67,.parent=44,.alignment=2"
         _StyleDefs(63)  =   "Splits(0).Columns(5).FooterStyle:id=68,.parent=45"
         _StyleDefs(64)  =   "Splits(0).Columns(5).EditorStyle:id=69,.parent=47"
         _StyleDefs(65)  =   "Splits(0).Columns(6).Style:id=74,.parent=43,.alignment=0,.valignment=1"
         _StyleDefs(66)  =   ":id=74,.locked=-1"
         _StyleDefs(67)  =   "Splits(0).Columns(6).HeadingStyle:id=71,.parent=44,.alignment=2"
         _StyleDefs(68)  =   "Splits(0).Columns(6).FooterStyle:id=72,.parent=45"
         _StyleDefs(69)  =   "Splits(0).Columns(6).EditorStyle:id=73,.parent=47"
         _StyleDefs(70)  =   "Splits(0).Columns(7).Style:id=78,.parent=43,.alignment=1,.valignment=1"
         _StyleDefs(71)  =   ":id=78,.locked=-1"
         _StyleDefs(72)  =   "Splits(0).Columns(7).HeadingStyle:id=75,.parent=44,.alignment=2"
         _StyleDefs(73)  =   "Splits(0).Columns(7).FooterStyle:id=76,.parent=45"
         _StyleDefs(74)  =   "Splits(0).Columns(7).EditorStyle:id=77,.parent=47"
         _StyleDefs(75)  =   "Splits(0).Columns(8).Style:id=82,.parent=43,.alignment=1,.valignment=1"
         _StyleDefs(76)  =   ":id=82,.locked=-1"
         _StyleDefs(77)  =   "Splits(0).Columns(8).HeadingStyle:id=79,.parent=44,.alignment=2"
         _StyleDefs(78)  =   "Splits(0).Columns(8).FooterStyle:id=80,.parent=45"
         _StyleDefs(79)  =   "Splits(0).Columns(8).EditorStyle:id=81,.parent=47"
         _StyleDefs(80)  =   "Splits(0).Columns(9).Style:id=86,.parent=43,.alignment=1,.valignment=1"
         _StyleDefs(81)  =   ":id=86,.locked=-1"
         _StyleDefs(82)  =   "Splits(0).Columns(9).HeadingStyle:id=83,.parent=44,.alignment=2"
         _StyleDefs(83)  =   "Splits(0).Columns(9).FooterStyle:id=84,.parent=45"
         _StyleDefs(84)  =   "Splits(0).Columns(9).EditorStyle:id=85,.parent=47"
         _StyleDefs(85)  =   "Splits(0).Columns(10).Style:id=90,.parent=43,.alignment=1,.valignment=1"
         _StyleDefs(86)  =   ":id=90,.locked=-1"
         _StyleDefs(87)  =   "Splits(0).Columns(10).HeadingStyle:id=87,.parent=44,.alignment=2"
         _StyleDefs(88)  =   "Splits(0).Columns(10).FooterStyle:id=88,.parent=45"
         _StyleDefs(89)  =   "Splits(0).Columns(10).EditorStyle:id=89,.parent=47"
         _StyleDefs(90)  =   "Splits(0).Columns(11).Style:id=94,.parent=43,.alignment=0,.valignment=1"
         _StyleDefs(91)  =   ":id=94,.locked=-1"
         _StyleDefs(92)  =   "Splits(0).Columns(11).HeadingStyle:id=91,.parent=44,.alignment=2"
         _StyleDefs(93)  =   "Splits(0).Columns(11).FooterStyle:id=92,.parent=45"
         _StyleDefs(94)  =   "Splits(0).Columns(11).EditorStyle:id=93,.parent=47"
         _StyleDefs(95)  =   "Splits(0).Columns(12).Style:id=98,.parent=43,.alignment=0,.valignment=1"
         _StyleDefs(96)  =   ":id=98,.locked=-1"
         _StyleDefs(97)  =   "Splits(0).Columns(12).HeadingStyle:id=95,.parent=44,.alignment=2"
         _StyleDefs(98)  =   "Splits(0).Columns(12).FooterStyle:id=96,.parent=45"
         _StyleDefs(99)  =   "Splits(0).Columns(12).EditorStyle:id=97,.parent=47"
         _StyleDefs(100) =   "Splits(0).Columns(13).Style:id=102,.parent=43,.alignment=0,.valignment=1"
         _StyleDefs(101) =   ":id=102,.locked=-1"
         _StyleDefs(102) =   "Splits(0).Columns(13).HeadingStyle:id=99,.parent=44,.alignment=2"
         _StyleDefs(103) =   "Splits(0).Columns(13).FooterStyle:id=100,.parent=45"
         _StyleDefs(104) =   "Splits(0).Columns(13).EditorStyle:id=101,.parent=47"
         _StyleDefs(105) =   "Splits(0).Columns(14).Style:id=106,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(106) =   "Splits(0).Columns(14).HeadingStyle:id=103,.parent=44"
         _StyleDefs(107) =   "Splits(0).Columns(14).FooterStyle:id=104,.parent=45"
         _StyleDefs(108) =   "Splits(0).Columns(14).EditorStyle:id=105,.parent=47"
         _StyleDefs(109) =   "Named:id=33:Normal"
         _StyleDefs(110) =   ":id=33,.parent=0"
         _StyleDefs(111) =   "Named:id=34:Heading"
         _StyleDefs(112) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(113) =   ":id=34,.wraptext=-1"
         _StyleDefs(114) =   "Named:id=35:Footing"
         _StyleDefs(115) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(116) =   "Named:id=36:Selected"
         _StyleDefs(117) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(118) =   "Named:id=37:Caption"
         _StyleDefs(119) =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(120) =   "Named:id=38:HighlightRow"
         _StyleDefs(121) =   ":id=38,.parent=33,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
         _StyleDefs(122) =   "Named:id=39:EvenRow"
         _StyleDefs(123) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(124) =   "Named:id=40:OddRow"
         _StyleDefs(125) =   ":id=40,.parent=33"
         _StyleDefs(126) =   "Named:id=41:RecordSelector"
         _StyleDefs(127) =   ":id=41,.parent=34"
         _StyleDefs(128) =   "Named:id=42:FilterBar"
         _StyleDefs(129) =   ":id=42,.parent=33"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "[F3] ---> Filtro Avanzado: Proveedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   135
         TabIndex        =   5
         Top             =   5580
         Width           =   2505
      End
      Begin VB.Label Label1 
         Caption         =   "Ayuda: Hacer Click Derecho en la Columna que desee buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   555
         Left            =   8595
         TabIndex        =   4
         Top             =   5625
         Visible         =   0   'False
         Width           =   2265
      End
      Begin VB.Label lblbusca 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   180
         TabIndex        =   3
         Top             =   5850
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin ActiveToolBars.SSActiveToolBars TlbBarra 
      Left            =   135
      Top             =   6705
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   7
      Tools           =   "frmhlmov.frx":002B
      ToolBars        =   "frmhlmov.frx":58B0
   End
   Begin VB.Menu mnupri 
      Caption         =   ""
      Begin VB.Menu mnufiltro 
         Caption         =   "Filtrar"
      End
      Begin VB.Menu mnufiltroavanz 
         Caption         =   "Filtro Avanzado"
      End
      Begin VB.Menu mnuordasc 
         Caption         =   "Ord. Asc"
      End
      Begin VB.Menu mnuorddesc 
         Caption         =   "Ord. Desc"
      End
      Begin VB.Menu mnutodo 
         Caption         =   "Mostrar Todos"
      End
   End
End
Attribute VB_Name = "frmhlmov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs      As Recordset
Dim xCodigo As Integer
Dim csql    As String
Dim ctitulo As String

Sub ImprimeEnDOS_Cabecera(pPagina As Integer)
Dim strCab As String
 
   strCab = ""
 '                    1234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
 '                             1         2         3         4         5         6         7         8         9         0        11        12        13        14        15        16        17        18        19
   strCab = strCab & wempresa & Space(196 - Len(wempresa)) & "Fecha  : " & Date & " " & Chr(13) & Chr(10)
   strCab = strCab & "Gestión de Compras - Infoplus                                                                                                                          " & Space(45) & "Pagina : " & Format(pPagina, "@@@@@@@@@@") & " " & Chr(13) & Chr(10)
   strCab = strCab & "                                                                                                                                                                          " & Chr(13) & Chr(10)
   strCab = strCab & "                                                                                                  REGISTRO DE COMPRAS                                                                      " & Chr(13) & Chr(10)
   strCab = strCab & Space(98) & ctitulo & " " & Chr(13) & Chr(10)
   strCab = strCab & "                                                                                                                                                                          " & Chr(13) & Chr(10)
   strCab = strCab & "|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|" & Chr(13) & Chr(10)
   strCab = strCab & "|          |   DOCUMENTO   |            |                         |BASE IMPONIBLE |BASE IMPONIBLE |BASE IMPONIBLE |ADQUISICIONES|IGV CON DERECHO|IGV OP.GRAVADAS|IGV SIN DERECHO|  BASE IMP.  |    IGV      |               |" & Chr(13) & Chr(10)
   strCab = strCab & "|  FECHA   |TD SER  NUMERO |    RUC     |     RAZON SOCIAL        |CON DER.A CRED.|OP.GRAV.Y NO GR|SIN DER.A CRED.| NO GRAVADAS |CREDITO FISCAL |Y NO GRAVADAS  |CREDITO FISCAL |    18%      |    18%      |    TOTAL      |" & Chr(13) & Chr(10)
   strCab = strCab & "|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|" & Chr(13) & Chr(10)
 '                    1234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
 '                             1         2         3         4         5         6         7         8         9         0        11        12        13        14        15        16        17        18        19
    Print #1, strCab

End Sub

Private Sub GENERA_TEMP()
Dim dbtempo As DAO.Database
Dim tbtempo As DAO.Recordset
Dim sw_cambio_igv   As Boolean
                    
    If DataAyuda.Recordset.RecordCount > 0 Then
        ctitulo = "Mes de " & wmesregcompras & " de " & wanno
        Set dbtempo = OpenDatabase(wrutatemp & "temp_com.mdb")
        Debug.Print wrutatemp & "\temp_com.mdb"
        dbtempo.Execute ("Delete From temp_gg")
        Set tbtempo = dbtempo.OpenRecordset("temp_gg")
        DataAyuda.Recordset.MoveFirst
        Do While Not DataAyuda.Recordset.EOF
            tbtempo.AddNew
            tbtempo.Fields("temp_mov") = DataAyuda.Recordset.Fields("f4nummov")
            tbtempo.Fields("temp_fecha") = DataAyuda.Recordset.Fields("f4fecha")
            tbtempo.Fields("TEMP_TDOC") = Format(DataAyuda.Recordset.Fields("F4TIPDOC"), "00")
            'If DataAyuda.Recordset.Fields("f4nomprv") = "ALFONSO SANCHES CARRASCO" Then
            '    MsgBox "vsd"
            'End If
            tbtempo.Fields("temp_serie") = Format(DataAyuda.Recordset.Fields("f4serdoc"), "000")
            tbtempo.Fields("temp_docum") = Format(DataAyuda.Recordset.Fields("f4numdoc"), "0000000")
            tbtempo.Fields("temp_prov") = DataAyuda.Recordset.Fields("f4nomprv") & ""
            tbtempo.Fields("temp_detal") = DataAyuda.Recordset.Fields("f4refere") & ""
            If Format(DataAyuda.Recordset.Fields("F4TIPDOC"), "00") = "02" Then
                tbtempo.Fields("TEMP_BIMP") = 0#
                tbtempo.Fields("TEMP_EXON") = Val("" & DataAyuda.Recordset.Fields("F4BASIMP")) + Val("" & DataAyuda.Recordset.Fields("F4MONINA"))
                tbtempo.Fields("temp_totals") = Val("" & DataAyuda.Recordset.Fields("F4BASIMP")) + Val("" & DataAyuda.Recordset.Fields("F4MONINA"))
            Else
                sw_cambio_igv = False
                'If wmesregcompras >= wf1mescambio_igv Then
                '    If Month(dataayuda.Recordset.Fields("f4fecha")) < Val(wf1mescambio_igv) Then
                '        sw_cambio_igv = True
                '    Else
                '        sw_cambio_igv = False
                '    End If
                'Else
                '    sw_cambio_igv = False
                'End If
                
                'If wmesregcompras >= wf1mescambio_igv Then
                '    If Val(DataAyuda.Recordset.Fields("F4PORC_IGV") & "") <> wigv Then
                '        sw_cambio_igv = True
                '    Else
                '        sw_cambio_igv = False
                '    End If
                'Else
                '    sw_cambio_igv = False
                'End If
                
                'If sw_cambio_igv = False Then
                    Select Case DataAyuda.Recordset.Fields("F4CODIGV") & ""
                        Case "001":
                            tbtempo.Fields("TEMP_BIMP") = Val("" & DataAyuda.Recordset.Fields("F4BASIMP"))
                            tbtempo.Fields("temp_igvs") = Val("" & DataAyuda.Recordset.Fields("f4igv"))
                        Case "002":
                            tbtempo.Fields("TEMP_BIMP_GYNG") = Val("" & DataAyuda.Recordset.Fields("F4BASIMP"))
                            tbtempo.Fields("TEMP_IGVS_GYNG") = Val("" & DataAyuda.Recordset.Fields("f4igv"))
                        Case "003":
                            tbtempo.Fields("TEMP_BIMP_SIN") = Val("" & DataAyuda.Recordset.Fields("F4BASIMP"))
                            tbtempo.Fields("TEMP_IGVS_SIN") = Val("" & DataAyuda.Recordset.Fields("f4igv"))
                    End Select
                'Else
                '   tbtempo.Fields("TEMP_BIMP_OTRO") = Val("" & DataAyuda.Recordset.Fields("F4BASIMP"))
                '   tbtempo.Fields("TEMP_IGVS_OTRO") = Val("" & DataAyuda.Recordset.Fields("f4igv"))
                'End If
                
                tbtempo.Fields("TEMP_EXON") = Val("" & DataAyuda.Recordset.Fields("F4OTRIMP")) + Val("" & DataAyuda.Recordset.Fields("F4MONINA")) + Val("" & DataAyuda.Recordset.Fields("F4REDSUMA")) - Val("" & DataAyuda.Recordset.Fields("F4REDRESTA")) - Val("" & DataAyuda.Recordset.Fields("F4DCTO"))
                tbtempo.Fields("temp_totals") = Val("" & DataAyuda.Recordset.Fields("f4total"))
                tbtempo.Fields("temp_totalD") = Val("" & DataAyuda.Recordset.Fields("TOTDOL"))
                tbtempo.Fields("temp_IGVD") = Val("" & DataAyuda.Recordset.Fields("F4TIPCAM"))
            End If

            tbtempo.Fields("TEMP_TIPPROV") = IIf(traerCampo("EF2PROVEEDORES", "F2TIPPROV", "F2NEWRUC", DataAyuda.Recordset.Fields("F4RUCPRV") & "") = "N", "REGISTRO DE COMPRAS NACIONALES", "REGISTRO DE COMPRAS IMPORTACION")
            tbtempo.Fields("empresa") = wf4empresa
            tbtempo.Fields("temp_moneda") = IIf(DataAyuda.Recordset.Fields("f4moneda") & "" = "S", "S/.", "US$")
            tbtempo.Fields("TEMP_RUC") = DataAyuda.Recordset.Fields("F4RUCPRV") & ""
            tbtempo.Fields("MES") = ctitulo
            tbtempo.Fields("TEMP_CODIGV") = DataAyuda.Recordset.Fields("F4CODIGV") & ""
            tbtempo.Fields("TEMP_NUMDEPOSITO") = Trim(DataAyuda.Recordset.Fields("F4NUMDEPOSITO") & "")
            tbtempo.Fields("TEMP_FECHADEPOSITO") = DataAyuda.Recordset.Fields("F4FECHADEPOSITO")
            tbtempo.Update
            DataAyuda.Recordset.MoveNext
        Loop
        tbtempo.Close
        dbtempo.Close
    Else
        MsgBox "No hay registros.", 48, "Atención"
    End If

End Sub

Private Sub datagrid_DblClick()
    
    datagrid_KeyPress 13

End Sub

Private Sub datagrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
       datagrid.col = 1
       mnufiltroavanz_Click
    End If
    
    datagrid_KeyPress KeyCode
End Sub

Private Sub datagrid_KeyPress(KeyAscii As Integer)
  
    Select Case KeyAscii
        Case 13:
            gnummov = "" & datagrid.Columns(0)
            Unload Me
        Case 27:
            gnummov = ""
            Unload Me
    
    End Select

End Sub

Private Sub datagrid_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

    Label1.Visible = True
    
End Sub

Private Sub datagrid_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    
    If txtope.Visible = True Then txtope.Visible = False
    If lblbusca.Visible = True Then lblbusca.Visible = False
    If txtfecha.Visible = True Then txtfecha.Visible = False
    mnufiltro.Caption = "Filtrar [" + datagrid.Columns(datagrid.col).Text + "]"
    Select Case Button
        Case 2
            PopupMenu mnupri
    End Select
 
End Sub

Private Sub Form_Activate()

    datagrid.OddRowStyle.BackColor = &HC0FFFF
    datagrid.EvenRowStyle.BackColor = &HFFFFFF
    datagrid.HighlightRowStyle.BackColor = vbActiveTitleBar
    datagrid.HighlightRowStyle.ForeColor = vbWhite
    datagrid.AlternatingRowStyle = True
    
End Sub

Private Sub Form_Load()

    DataAyuda.DatabaseName = wrutabancos & "\DB_BANCOS.mdb"
    
    csql = "SELECT REGISDOC.F4PORC_IGV, REGISDOC.F4FECHADEPOSITO, REGISDOC.F4NUMDEPOSITO, REGISDOC.F4NUMMOV, REGISDOC.F4NOMPRV, REGISDOC.F4TIPDOC, REGISDOC.F4SERDOC, REGISDOC.F4NUMDOC, REGISDOC.F4FECHA, REGISDOC.F4MONEDA, VAL(FORMAT(IIf([REGISDOC].[F4MONEDA]='S',[REGISDOC].[F4BASIMP],[REGISDOC].[F4BASIMP]*[REGISDOC].[F4TIPCAM]),'0.00')) AS F4BASIMP,  REGISDOC.F4REFERE, REGISDOC.F4CODPRV, REGISDOC.F4USUARIOING, REGISDOC.F4DCTO, REGISDOC.F4RUCPRV,  REGISDOC.F4CODIGV,VAL(FORMAT(IIf([REGISDOC].[F4MONEDA]='S',[REGISDOC].[F4MONINA],[REGISDOC].[F4MONINA]*[REGISDOC].[F4TIPCAM]),'0.00')) AS F4MONINA,VAL(FORMAT(IIf([REGISDOC].[F4MONEDA]='S',[REGISDOC].[F4IGV],[REGISDOC].[F4IGV]*[REGISDOC].[F4TIPCAM]),'0.00')) AS F4IGV,VAL(FORMAT(IIf([REGISDOC].[F4MONEDA]='S',[REGISDOC].[F4TOTAL],[REGISDOC].[F4TOTAL]*[REGISDOC].[F4TIPCAM]),'0.00')) AS F4TOTAL,VAL(FORMAT(IIf([REGISDOC].[F4MONEDA]='S',[REGISDOC].[f4redsuma],[REGISDOC].[f4redsuma]*[REGISDOC].[F4TIPCAM]),'0.00')) AS f4redsuma,VAL(FORMAT(IIf([REGISDOC].[F4MONEDA]='S' " & _
            ",[REGISDOC].[f4redresta],[REGISDOC].[f4redresta]*[REGISDOC].[F4TIPCAM]),'0.00')) AS f4redresta,VAL(FORMAT(IIf([REGISDOC].[F4MONEDA]='S',[REGISDOC].[F4DCTO],[REGISDOC].[F4DCTO]*[REGISDOC].[F4TIPCAM]),'0.00')) AS F4DCTO,VAL(FORMAT(IIf([REGISDOC].[F4MONEDA]='S',[REGISDOC].[F4OTRIMP],[REGISDOC].[F4OTRIMP]*[REGISDOC].[F4TIPCAM]),'0.00')) AS F4OTRIMP,REGISDOC.F4TIPCAM, VAL(FORMAT(IIf([REGISDOC].[F4MONEDA]='S',0,[REGISDOC].[F4TOTAL]),'0.00')) AS TOTDOL " & _
            "FROM REGISDOC WHERE REGISDOC.F4MESMOV = '" & Format(wmesregcompras, "00") & "'" & _
            "ORDER BY F4NUMMOV DESC"
    DataAyuda.RecordSource = csql
    DataAyuda.Refresh
    
End Sub

Private Sub mnufiltro_Click()

    Set rs = DataAyuda.Recordset
    Select Case DataAyuda.Recordset.Fields(datagrid.Columns(datagrid.col).DataField).Type
        Case 10
            rs.Filter = "[" + datagrid.Columns(datagrid.col).DataField + "]" + " = '" + Trim("" & datagrid.Columns(datagrid.col).Text) + "'"
        Case 4
            rs.Filter = "[" + datagrid.Columns(datagrid.col).DataField + "]" + " = " + datagrid.Columns(datagrid.col).Text
        Case 8
            If IsDate(datagrid.Columns(datagrid.col).Text) Then
                rs.Filter = "[" + datagrid.Columns(datagrid.col).DataField + "]" + "=#" + datagrid.Columns(datagrid.col).Text + "#"
            Else
                MsgBox "Ingrese una Fecha Valida..!", 32, "Advertencia"
                Exit Sub
            End If
    End Select
    Rem jcg Set DataAyuda.Recordset = rs.OpenRecordset(rs.Type)
    Set rs = Nothing
    
End Sub

Private Sub mnufiltroavanz_Click()
  
    Select Case datagrid.col
        Case 0:
            lblbusca.Visible = True
            lblbusca.Caption = datagrid.Columns(datagrid.col).Caption
            txtope.Visible = True
            txtope.Text = ""
            txtope.SetFocus
        Case 1
            lblbusca.Visible = True
            lblbusca.Caption = datagrid.Columns(datagrid.col).Caption
            txtope.Visible = True
            txtope.Text = ""
            txtope.SetFocus
        Case 4
            lblbusca.Visible = True
            lblbusca.Caption = datagrid.Columns(datagrid.col).Caption
            txtope.Visible = True
            txtope.Text = ""
            txtope.SetFocus
        Case 5
            lblbusca.Visible = True
            lblbusca.Caption = datagrid.Columns(datagrid.col).Caption
            txtfecha.Visible = True
            txtfecha.Text = Date
            txtfecha.SetFocus
    End Select
    
End Sub

Private Sub mnuordasc_Click()

    Set rs = DataAyuda.Recordset
    rs.Sort = "[" + datagrid.Columns(datagrid.col).DataField + "] Asc"
    Rem jcg Set DataAyuda.Recordset = rs.OpenRecordset(rs.Type)
    Set rs = Nothing
  
End Sub

Private Sub mnuorddesc_Click()

    Set rs = DataAyuda.Recordset
    rs.Sort = "[" + datagrid.Columns(datagrid.col).DataField + "] Desc"
    Rem jcg Set DataAyuda.Recordset = rs.OpenRecordset(rs.Type)
    Set rs = Nothing
    
End Sub

Private Sub MnuTodo_Click()
  
    DataAyuda.DatabaseName = wrutabanco & "\DB_BANCOS.mdb"
    DataAyuda.RecordSource = csql
    DataAyuda.Refresh
    
End Sub

Private Sub SSPanel1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    
    Label1.Visible = False
    
End Sub

Private Sub TlbBarra_ToolClick(ByVal Tool As ActiveToolBars.SSTool)

    Select Case Tool.ID
        Case "ID_Imprimir":
            Me.MousePointer = 11
            'IMPRIMIR_DOS
            Me.MousePointer = 1
        Case "ID_ImprimirenWindows":
            Me.MousePointer = 11
            IMPRIMIR_WINDOWS
            Me.MousePointer = 1
        Case "Salir":
            Unload Me
    End Select
    
End Sub

Private Sub txtfecha_GotFocus()
 
    txtfecha.SelStart = 0: txtfecha.SelLength = Len(txtfecha.Text)
    
End Sub

Private Sub txtfecha_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtope.Text = txtfecha.Text
        'txtope_KeyPress 13
    End If

End Sub

Private Sub IMPRIMIR_WINDOWS()

    GENERA_TEMP
    With acr_regcompras
        .DataControl1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & "\temp_com.MDB;Persist Security Info=False"
        If cons_regcompras.optorden(0).Value = True Then
            .DataControl1.Source = "Select * from temp_gg ORDER BY TEMP_TIPPROV DESC,TEMP_MOV"
            Debug.Print "Select * from temp_gg ORDER BY TEMP_TIPPROV DESC,TEMP_MOV"
        Else
            .DataControl1.Source = "Select * from temp_gg ORDER BY TEMP_TIPPROV DESC,TEMP_FECHA"
        End If
        .fldtitulo.Text = ctitulo & " - Moneda: Soles"
        .fldempresa.Text = wf4empresa
        .fldfecha.Text = Format(Date, "dd/mm/yyyy")
        .Show vbModal
    End With
    
End Sub

Private Sub IMPRIMIR_DOS()
Dim conTmp      As New ADODB.Connection
Dim recTmp      As New ADODB.Recordset
Dim strArc      As String
Dim Pagina      As Integer
Dim Fila        As Integer
Dim csql        As String
Dim nsubbi      As Double
Dim nsubexo     As Double
Dim nsubigv     As Double
Dim nsuttotal   As Double
Dim nsubbi_sin  As Double
Dim nsubigv_sin As Double
        
    GENERA_TEMP
   
    strArc = Trim(gcoduse) & ".TXT"
    conTmp.Provider = "Microsoft.JET.OLEDB.4.0; Data Source=" & wrutatemp & "\temp_com.MDB; Persist Security Info=False"
    conTmp.Open
    
    Open Trim(gcoduse) & ".TXT" For Output As #1
    Pagina = 0
    Fila = 10
    If wordenfecha = "P" Then
        csql = "Select * from temp_gg Order By TEMP_RUC,TEMP_TDOC,temp_serie,temp_docum"
    Else
        csql = "Select * from temp_gg Order By TEMP_FECHA"
    End If

    nsubbi = 0#: nsubexo = 0#: nsubigv = 0#: nsuttotal = 0#: nsubbi_sin = 0#: nsubigv_sin = 0#
    Set recTmp = conTmp.Execute(csql)
    Do While Not recTmp.EOF
        If Fila = 56 Then
            '---------------------------------
            '---------------------------------
            strlin = Space(234)
            strlin = String(220, "-")
            Print #1, strlin
            Fila = Fila + 1
            strlin = Space(234)
            Mid(strlin, 45, 15) = "SubTotal"
            Mid(strlin, 69, 15) = Format$(Format(nsubbi, "#,###,##0.00"), "@@@@@@@@@@@@")
            Mid(strlin, 84, 15) = Format$(Format(nsubbi_gyng, "#,###,##0.00"), "@@@@@@@@@@@@")
            Mid(strlin, 99, 15) = Format$(Format(nsubbi_sin, "#,###,##0.00"), "@@@@@@@@@@@@")
            Mid(strlin, 114, 15) = Format$(Format(nsubexo, "#,###,##0.00"), "@@@@@@@@@@@@")
            Mid(strlin, 129, 15) = Format$(Format(nsubigv, "#,###,##0.00"), "@@@@@@@@@@@@")
            Mid(strlin, 144, 15) = Format$(Format(nsubigv_gyng, "#,###,##0.00"), "@@@@@@@@@@@@")
            Mid(strlin, 159, 15) = Format$(Format(nsubigv_sin, "#,###,##0.00"), "@@@@@@@@@@@@")
            Mid(strlin, 174, 15) = Format$(Format(nsubbi_otro, "#,###,##0.00"), "@@@@@@@@@@@@")
            Mid(strlin, 189, 15) = Format$(Format(nsubigv_otro, "#,###,##0.00"), "@@@@@@@@@@@@")
            Mid(strlin, 207, 15) = Format$(Format(nsuttotal, "#,###,##0.00"), "@@@@@@@@@@@@")
            Print #1, strlin
            Fila = Fila + 1
            strlin = String(220, "-")
            Print #1, strlin
            Fila = Fila + 1
            '---------------------------------
            '---------------------------------
            Fila = 10
        End If
        If Fila = 10 Then
            If Pagina > 0 Then
               Print #1, Chr(12)
            End If
            Pagina = Pagina + 1
            ImprimeEnDOS_Cabecera Pagina
        End If
        
        strlin = Space(234)
        intPos = 2:            Mid(strlin, intPos, intPos + 11) = "" & recTmp("TEMP_FECHA")
        intPos = intPos + 11:  Mid(strlin, intPos + 1, intPos + 3) = "" & recTmp("TEMP_TDOC")
        intPos = intPos + 3:   Mid(strlin, intPos + 2, intPos + 4) = "" & recTmp("TEMP_SERIE")
        intPos = intPos + 4:   Mid(strlin, intPos + 2, intPos + 8) = "" & recTmp("TEMP_DOCUM")
        intPos = intPos + 8:   Mid(strlin, intPos + 2, intPos + 12) = "" & recTmp("TEMP_RUC")
        intPos = intPos + 12:  Mid(strlin, intPos + 2, intPos + 26) = Mid("" & recTmp("temp_prov"), 1, 25)
        intPos = intPos + 26:  Mid(strlin, intPos + 3, intPos + 15) = Format$(Format("" & recTmp("TEMP_BIMP"), "#,###,##0.00"), "@@@@@@@@@@@@")
        intPos = intPos + 15:  Mid(strlin, intPos + 3, intPos + 15) = Format$(Format("" & recTmp("TEMP_BIMP_GYNG"), "#,###,##0.00"), "@@@@@@@@@@@@")
        intPos = intPos + 15:  Mid(strlin, intPos + 3, intPos + 15) = Format$(Format("" & recTmp("TEMP_BIMP_SIN"), "#,###,##0.00"), "@@@@@@@@@@@@")
        intPos = intPos + 15:  Mid(strlin, intPos + 3, intPos + 15) = Format$(Format("" & recTmp("TEMP_EXON"), "#,###,##0.00"), "@@@@@@@@@@@@")
        intPos = intPos + 15:  Mid(strlin, intPos + 3, intPos + 15) = Format$(Format("" & recTmp("TEMP_IGVS"), "#,###,##0.00"), "@@@@@@@@@@@@")
        intPos = intPos + 15:  Mid(strlin, intPos + 3, intPos + 15) = Format$(Format("" & recTmp("TEMP_IGVS_GYNG"), "#,###,##0.00"), "@@@@@@@@@@@@")
        intPos = intPos + 15:  Mid(strlin, intPos + 3, intPos + 15) = Format$(Format("" & recTmp("TEMP_IGVS_SIN"), "#,###,##0.00"), "@@@@@@@@@@@@")
        intPos = intPos + 15:  Mid(strlin, intPos + 3, intPos + 15) = Format$(Format("" & recTmp("TEMP_BIMP_OTRO"), "#,###,##0.00"), "@@@@@@@@@@@@")
        intPos = intPos + 15:  Mid(strlin, intPos + 3, intPos + 15) = Format$(Format("" & recTmp("TEMP_IGVS_OTRO"), "#,###,##0.00"), "@@@@@@@@@@@@")
        intPos = intPos + 18:  Mid(strlin, intPos + 3, intPos + 15) = Format$(Format("" & recTmp("TEMP_TOTALS"), "#,###,##0.00"), "@@@@@@@@@@@@")
        nsubbi = nsubbi + recTmp("TEMP_BIMP")
        nsubexo = nsubexo + recTmp("TEMP_EXON")
        nsubigv = nsubigv + recTmp("TEMP_IGVS")
        nsuttotal = nsuttotal + recTmp("TEMP_TOTALS")
        nsubbi_sin = nsubbi_sin + recTmp("TEMP_BIMP_SIN")
        nsubigv_sin = nsubigv_sin + recTmp("TEMP_IGVS_SIN")
        
        nsubbi_gyng = nsubbi_gyng + recTmp("TEMP_BIMP_GYNG")
        nsubigv_gyng = nsubigv_gyng + recTmp("TEMP_IGVS_GYNG")
        nsubbi_otro = nsubbi_otro + recTmp("TEMP_BIMP_OTRO")
        nsubigv_otro = nsubigv_otro + recTmp("TEMP_IGVS_OTRO")
        
        Print #1, strlin
        Fila = Fila + 1
        recTmp.MoveNext
    Loop
    recTmp.Close
    Set recClose = Nothing
    conTmp.Close
    Set conTmp = Nothing
    
    Print #1, Space(138)
    '---------------------------------
    strlin = Space(234)
    strlin = String(220, "-")
    Print #1, strlin
    Fila = Fila + 1
    strlin = Space(234)
    Mid(strlin, 45, 15) = "Total Final "
    Mid(strlin, 69, 15) = Format$(Format(nsubbi, "#,###,##0.00"), "@@@@@@@@@@@@")
    Mid(strlin, 84, 15) = Format$(Format(nsubbi_gyng, "#,###,##0.00"), "@@@@@@@@@@@@")
    Mid(strlin, 99, 15) = Format$(Format(nsubbi_sin, "#,###,##0.00"), "@@@@@@@@@@@@")
    Mid(strlin, 114, 15) = Format$(Format(nsubexo, "#,###,##0.00"), "@@@@@@@@@@@@")
    Mid(strlin, 129, 15) = Format$(Format(nsubigv, "#,###,##0.00"), "@@@@@@@@@@@@")
    Mid(strlin, 144, 15) = Format$(Format(nsubigv_gyng, "#,###,##0.00"), "@@@@@@@@@@@@")
    Mid(strlin, 159, 15) = Format$(Format(nsubigv_sin, "#,###,##0.00"), "@@@@@@@@@@@@")
    Mid(strlin, 174, 15) = Format$(Format(nsubbi_otro, "#,###,##0.00"), "@@@@@@@@@@@@")
    Mid(strlin, 189, 15) = Format$(Format(nsubigv_otro, "#,###,##0.00"), "@@@@@@@@@@@@")
    Mid(strlin, 207, 15) = Format$(Format(nsuttotal, "#,###,##0.00"), "@@@@@@@@@@@@")
    Print #1, strlin
    Fila = Fila + 1
    strlin = String(220, "-")
    Print #1, strlin
    Fila = Fila + 1
    '---------------------------------
    
    Close #1
    
    frmReporte.NombreArchivo = strArc
    frmReporte.Show 1

End Sub
