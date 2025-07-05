VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form FrmReporte 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gastos por Proveedor"
   ClientHeight    =   7425
   ClientLeft      =   1680
   ClientTop       =   1380
   ClientWidth     =   11655
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
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7425
   ScaleWidth      =   11655
   Begin TrueOleDBGrid70.TDBGrid GRID 
      Bindings        =   "FREPORTE.frx":0000
      Height          =   5100
      Left            =   90
      TabIndex        =   15
      Top             =   1485
      Width           =   11460
      _ExtentX        =   20214
      _ExtentY        =   8996
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Fecha"
      Columns(0).DataField=   "FECHA"
      Columns(0).DataWidth=   19
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Comprob."
      Columns(1).DataField=   "COMPROBANTE"
      Columns(1).DataWidth=   255
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "N°.Docum."
      Columns(2).DataField=   "NUMDOC"
      Columns(2).DataWidth=   255
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Cta. Cont"
      Columns(3).DataField=   "ctacont"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "C.C."
      Columns(4).DataField=   "centro"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Detalle"
      Columns(5).DataField=   "DETALLE"
      Columns(5).DataWidth=   255
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Base Imp."
      Columns(6).DataField=   "PCOMPRA"
      Columns(6).DataWidth=   23
      Columns(6).NumberFormat=   "Standard"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "TC"
      Columns(7).DataField=   "TC"
      Columns(7).DataWidth=   23
      Columns(7).NumberFormat=   "Standard"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   8
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).ScrollBars=   2
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=8"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1667"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1588"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=66082"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=1879"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1799"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=65809"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=1984"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1905"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=65808"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=2090"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=2011"
      Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=66080"
      Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(25)=   "Column(4).Width=2090"
      Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=2011"
      Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=66080"
      Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(31)=   "Column(5).Width=6085"
      Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=6006"
      Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=66080"
      Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(37)=   "Column(6).Width=1826"
      Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=1746"
      Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=66082"
      Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(43)=   "Column(7).Width=1402"
      Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=1323"
      Splits(0)._ColumnProps(46)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(47)=   "Column(7)._ColStyle=66082"
      Splits(0)._ColumnProps(48)=   "Column(7).Order=8"
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
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=0,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=228,.bold=0,.fontsize=780,.italic=0"
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
      _StyleDefs(35)  =   "Splits(0).Columns(0).Style:id=28,.parent=43,.alignment=1,.valignment=1"
      _StyleDefs(36)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=44,.alignment=2"
      _StyleDefs(37)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=45"
      _StyleDefs(38)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=47"
      _StyleDefs(39)  =   "Splits(0).Columns(1).Style:id=32,.parent=43,.alignment=2,.valignment=2"
      _StyleDefs(40)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=44"
      _StyleDefs(41)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=45"
      _StyleDefs(42)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=47"
      _StyleDefs(43)  =   "Splits(0).Columns(2).Style:id=58,.parent=43,.alignment=0,.valignment=2"
      _StyleDefs(44)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=44"
      _StyleDefs(45)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=45"
      _StyleDefs(46)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=47"
      _StyleDefs(47)  =   "Splits(0).Columns(3).Style:id=62,.parent=43,.alignment=0,.valignment=1"
      _StyleDefs(48)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=44,.alignment=2"
      _StyleDefs(49)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=45"
      _StyleDefs(50)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=47"
      _StyleDefs(51)  =   "Splits(0).Columns(4).Style:id=66,.parent=43,.alignment=0,.valignment=1"
      _StyleDefs(52)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=44,.alignment=2"
      _StyleDefs(53)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=45"
      _StyleDefs(54)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=47"
      _StyleDefs(55)  =   "Splits(0).Columns(5).Style:id=70,.parent=43,.alignment=0,.valignment=1"
      _StyleDefs(56)  =   "Splits(0).Columns(5).HeadingStyle:id=67,.parent=44,.alignment=2"
      _StyleDefs(57)  =   "Splits(0).Columns(5).FooterStyle:id=68,.parent=45"
      _StyleDefs(58)  =   "Splits(0).Columns(5).EditorStyle:id=69,.parent=47"
      _StyleDefs(59)  =   "Splits(0).Columns(6).Style:id=74,.parent=43,.alignment=1,.valignment=1"
      _StyleDefs(60)  =   "Splits(0).Columns(6).HeadingStyle:id=71,.parent=44,.alignment=2"
      _StyleDefs(61)  =   "Splits(0).Columns(6).FooterStyle:id=72,.parent=45"
      _StyleDefs(62)  =   "Splits(0).Columns(6).EditorStyle:id=73,.parent=47"
      _StyleDefs(63)  =   "Splits(0).Columns(7).Style:id=78,.parent=43,.alignment=1,.valignment=1"
      _StyleDefs(64)  =   "Splits(0).Columns(7).HeadingStyle:id=75,.parent=44,.alignment=2"
      _StyleDefs(65)  =   "Splits(0).Columns(7).FooterStyle:id=76,.parent=45"
      _StyleDefs(66)  =   "Splits(0).Columns(7).EditorStyle:id=77,.parent=47"
      _StyleDefs(67)  =   "Named:id=33:Normal"
      _StyleDefs(68)  =   ":id=33,.parent=0"
      _StyleDefs(69)  =   "Named:id=34:Heading"
      _StyleDefs(70)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(71)  =   ":id=34,.wraptext=-1"
      _StyleDefs(72)  =   "Named:id=35:Footing"
      _StyleDefs(73)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(74)  =   "Named:id=36:Selected"
      _StyleDefs(75)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(76)  =   "Named:id=37:Caption"
      _StyleDefs(77)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(78)  =   "Named:id=38:HighlightRow"
      _StyleDefs(79)  =   ":id=38,.parent=33,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(80)  =   "Named:id=39:EvenRow"
      _StyleDefs(81)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(82)  =   "Named:id=40:OddRow"
      _StyleDefs(83)  =   ":id=40,.parent=33"
      _StyleDefs(84)  =   "Named:id=41:RecordSelector"
      _StyleDefs(85)  =   ":id=41,.parent=34"
      _StyleDefs(86)  =   "Named:id=42:FilterBar"
      _StyleDefs(87)  =   ":id=42,.parent=33"
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   270
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6660
      Visible         =   0   'False
      Width           =   1452
   End
   Begin TrueOleDBGrid70.TDBGrid Grid1 
      Bindings        =   "FREPORTE.frx":0014
      Height          =   5100
      Left            =   90
      TabIndex        =   14
      Top             =   1485
      Width           =   11460
      _ExtentX        =   20214
      _ExtentY        =   8996
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Fecha"
      Columns(0).DataField=   "FECHA"
      Columns(0).DataWidth=   19
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Comprob."
      Columns(1).DataField=   "COMPROBANTE"
      Columns(1).DataWidth=   255
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "N°.Docum."
      Columns(2).DataField=   "NUMDOC"
      Columns(2).DataWidth=   255
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Detalle"
      Columns(3).DataField=   "detalle"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Base Imp."
      Columns(4).DataField=   "BASIMP"
      Columns(4).DataWidth=   23
      Columns(4).NumberFormat=   "Standard"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "IGV"
      Columns(5).DataField=   "IGV"
      Columns(5).DataWidth=   23
      Columns(5).NumberFormat=   "Standard"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Exon."
      Columns(6).DataField=   "EXON"
      Columns(6).DataWidth=   23
      Columns(6).NumberFormat=   "Standard"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "P.Compra"
      Columns(7).DataField=   "PCOMPRA"
      Columns(7).NumberFormat=   "Standard"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   8
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).ScrollBars=   2
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=8"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1799"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1720"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=66082"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=1958"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1879"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=66081"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=2196"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2117"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=66080"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=5609"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=5530"
      Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=66080"
      Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(25)=   "Column(4).Width=1773"
      Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=1693"
      Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=66082"
      Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(31)=   "Column(5).Width=1429"
      Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=1349"
      Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=66082"
      Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(37)=   "Column(6).Width=1773"
      Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=1693"
      Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=66082"
      Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(43)=   "Column(7).Width=2487"
      Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=2408"
      Splits(0)._ColumnProps(46)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(47)=   "Column(7)._ColStyle=66082"
      Splits(0)._ColumnProps(48)=   "Column(7).Order=8"
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
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=0,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=114,.bold=0,.fontsize=780,.italic=0"
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
      _StyleDefs(35)  =   "Splits(0).Columns(0).Style:id=28,.parent=43,.alignment=1,.valignment=1"
      _StyleDefs(36)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=44,.alignment=2"
      _StyleDefs(37)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=45"
      _StyleDefs(38)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=47"
      _StyleDefs(39)  =   "Splits(0).Columns(1).Style:id=32,.parent=43,.alignment=2,.valignment=1"
      _StyleDefs(40)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=44,.alignment=2"
      _StyleDefs(41)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=45"
      _StyleDefs(42)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=47"
      _StyleDefs(43)  =   "Splits(0).Columns(2).Style:id=58,.parent=43,.alignment=0,.valignment=1"
      _StyleDefs(44)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=44,.alignment=2"
      _StyleDefs(45)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=45"
      _StyleDefs(46)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=47"
      _StyleDefs(47)  =   "Splits(0).Columns(3).Style:id=62,.parent=43,.alignment=0,.valignment=1"
      _StyleDefs(48)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=44,.alignment=2"
      _StyleDefs(49)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=45"
      _StyleDefs(50)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=47"
      _StyleDefs(51)  =   "Splits(0).Columns(4).Style:id=66,.parent=43,.alignment=1,.valignment=1"
      _StyleDefs(52)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=44,.alignment=2"
      _StyleDefs(53)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=45"
      _StyleDefs(54)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=47"
      _StyleDefs(55)  =   "Splits(0).Columns(5).Style:id=70,.parent=43,.alignment=1,.valignment=1"
      _StyleDefs(56)  =   "Splits(0).Columns(5).HeadingStyle:id=67,.parent=44,.alignment=2"
      _StyleDefs(57)  =   "Splits(0).Columns(5).FooterStyle:id=68,.parent=45"
      _StyleDefs(58)  =   "Splits(0).Columns(5).EditorStyle:id=69,.parent=47"
      _StyleDefs(59)  =   "Splits(0).Columns(6).Style:id=74,.parent=43,.alignment=1,.valignment=1"
      _StyleDefs(60)  =   "Splits(0).Columns(6).HeadingStyle:id=71,.parent=44,.alignment=2"
      _StyleDefs(61)  =   "Splits(0).Columns(6).FooterStyle:id=72,.parent=45"
      _StyleDefs(62)  =   "Splits(0).Columns(6).EditorStyle:id=73,.parent=47"
      _StyleDefs(63)  =   "Splits(0).Columns(7).Style:id=78,.parent=43,.alignment=1,.valignment=1"
      _StyleDefs(64)  =   "Splits(0).Columns(7).HeadingStyle:id=75,.parent=44,.alignment=2"
      _StyleDefs(65)  =   "Splits(0).Columns(7).FooterStyle:id=76,.parent=45"
      _StyleDefs(66)  =   "Splits(0).Columns(7).EditorStyle:id=77,.parent=47"
      _StyleDefs(67)  =   "Named:id=33:Normal"
      _StyleDefs(68)  =   ":id=33,.parent=0"
      _StyleDefs(69)  =   "Named:id=34:Heading"
      _StyleDefs(70)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(71)  =   ":id=34,.wraptext=-1"
      _StyleDefs(72)  =   "Named:id=35:Footing"
      _StyleDefs(73)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(74)  =   "Named:id=36:Selected"
      _StyleDefs(75)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(76)  =   "Named:id=37:Caption"
      _StyleDefs(77)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(78)  =   "Named:id=38:HighlightRow"
      _StyleDefs(79)  =   ":id=38,.parent=33,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(80)  =   "Named:id=39:EvenRow"
      _StyleDefs(81)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(82)  =   "Named:id=40:OddRow"
      _StyleDefs(83)  =   ":id=40,.parent=33"
      _StyleDefs(84)  =   "Named:id=41:RecordSelector"
      _StyleDefs(85)  =   ":id=41,.parent=34"
      _StyleDefs(86)  =   "Named:id=42:FilterBar"
      _StyleDefs(87)  =   ":id=42,.parent=33"
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   480
      Left            =   90
      TabIndex        =   8
      Top             =   0
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   847
      _StockProps     =   15
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
      Begin Threed.SSCommand BtnExitO 
         Height          =   330
         Left            =   570
         TabIndex        =   17
         ToolTipText     =   "Salir"
         Top             =   90
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   582
         _StockProps     =   78
         Picture         =   "FREPORTE.frx":0028
      End
      Begin Threed.SSCommand BtnPrint 
         Height          =   330
         Left            =   180
         TabIndex        =   16
         ToolTipText     =   "Imprimir"
         Top             =   90
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   582
         _StockProps     =   78
         Picture         =   "FREPORTE.frx":0182
      End
   End
   Begin Crystal.CrystalReport REPORTE 
      Left            =   1728
      Top             =   48
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.PictureBox Panel3D1 
      BackColor       =   &H8000000A&
      Height          =   600
      Left            =   3105
      ScaleHeight     =   540
      ScaleWidth      =   8370
      TabIndex        =   6
      Top             =   6660
      Width           =   8430
      Begin Threed.SSPanel pnlbi 
         Height          =   390
         Left            =   1005
         TabIndex        =   11
         Top             =   105
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   688
         _StockProps     =   15
         BackColor       =   -2147483638
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         BevelInner      =   1
         Alignment       =   4
      End
      Begin Threed.SSPanel pnligv 
         Height          =   390
         Left            =   3075
         TabIndex        =   12
         Top             =   105
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   688
         _StockProps     =   15
         BackColor       =   -2147483638
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         BevelInner      =   1
         Alignment       =   4
      End
      Begin Threed.SSPanel pnlexo 
         Height          =   390
         Left            =   4995
         TabIndex        =   13
         Top             =   105
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   688
         _StockProps     =   15
         BackColor       =   -2147483638
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         BevelInner      =   1
         Alignment       =   4
      End
      Begin Threed.SSPanel pnltotal 
         Height          =   390
         Left            =   6930
         TabIndex        =   9
         Top             =   105
         Width           =   1260
         _Version        =   65536
         _ExtentX        =   2222
         _ExtentY        =   688
         _StockProps     =   15
         BackColor       =   -2147483638
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         BevelInner      =   1
         Alignment       =   4
      End
      Begin VB.Label lblexo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Exon."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   4500
         TabIndex        =   7
         Top             =   180
         Width           =   450
      End
      Begin VB.Label lbltotal 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   6480
         TabIndex        =   3
         Top             =   180
         Width           =   405
      End
      Begin VB.Label lbligv 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "I.G.V."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   2520
         TabIndex        =   4
         Top             =   180
         Width           =   420
      End
      Begin VB.Label lblbi 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Base Imp."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   180
         TabIndex        =   5
         Top             =   180
         Width           =   840
      End
   End
   Begin VB.PictureBox Frame3D1 
      BackColor       =   &H8000000A&
      Height          =   915
      Left            =   45
      ScaleHeight     =   855
      ScaleWidth      =   11385
      TabIndex        =   0
      Top             =   495
      Width           =   11445
      Begin Threed.SSPanel panel 
         Height          =   345
         Left            =   1200
         TabIndex        =   10
         Top             =   330
         Width           =   10005
         _Version        =   65536
         _ExtentX        =   17648
         _ExtentY        =   609
         _StockProps     =   15
         BackColor       =   -2147483638
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         BevelInner      =   1
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   225
         TabIndex        =   1
         Top             =   405
         Width           =   840
      End
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Label3"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2205
      TabIndex        =   2
      Top             =   90
      Width           =   9240
   End
End
Attribute VB_Name = "FrmReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BtnExitO_Click()

    Unload Me

End Sub


Private Sub BtnPrint_Click()

    Select Case xtipo
        Case "0":
            REPORTE.DataFiles(0) = wrutatemp & "\" & "temp_com.mdb"
            REPORTE.ReportFileName = wrutatemp & "\" & "provgas.rpt"
            REPORTE.Action = 1
        Case "1", "2":
            REPORTE.DataFiles(0) = wrutatemp & "\" & "temp_com.mdb"
            REPORTE.ReportFileName = wrutatemp & "\" & "provgasd.rpt"
            REPORTE.Action = 1
    End Select
    
End Sub


Private Sub Form_Activate()

    GRID.EvenRowStyle.BackColor = &HFFFFFF
    GRID.OddRowStyle.BackColor = &HC0FFFF
    GRID.HighlightRowStyle.BackColor = vbActiveTitleBar
    GRID.HighlightRowStyle.ForeColor = vbWhite
    GRID.AlternatingRowStyle = True

    Grid1.EvenRowStyle.BackColor = &HFFFFFF
    Grid1.OddRowStyle.BackColor = &HC0FFFF
    Grid1.HighlightRowStyle.BackColor = vbActiveTitleBar
    Grid1.HighlightRowStyle.ForeColor = vbWhite
    Grid1.AlternatingRowStyle = True

End Sub

Private Sub Form_Load()
'On Error GoTo error_bd

    Set dbcompras = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set tbregisdoc = dbcompras.OpenRecordset("regisdoc")
    Set tbregismov = dbcompras.OpenRecordset("regismov")

    Llena_Grid
    
    Data1.DatabaseName = (wrutatemp & "\temp_com.mdb")
    
    If xtipo = "0" Then
        Grid1.Visible = True
        GRID.Visible = False
        Data1.RecordSource = "select fecha, comprobante, numdoc,detalle, basimp,igv,exon,pcompra from temgapro"
    Else
        If xtipo = "1" Or xtipo = "2" Then
            Grid1.Visible = False
            GRID.Visible = True
            Data1.RecordSource = "select fecha, comprobante,ctacont, centro,detalle, pcompra, tc from temgapro"
        End If
   End If
   
   Data1.Refresh
   tbregismov.Close
   tbregisdoc.Close
   dbcompras.Close

    Exit Sub

'error_bd:
'    Resume Next

End Sub

Private Sub Frame3D1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

  '  apagar

End Sub

Private Sub Llena_Grid()
Dim CONSULTA    As DAO.Recordset
Dim consulta1   As DAO.Recordset
Dim dbtemp      As DAO.Database
Dim temprogas   As DAO.Recordset
Dim ctitu       As String
Dim nbi         As Double, nigv As Double, ntotal As Double, nexo As Double
Dim xtipdoc     As String

On Error GoTo error_cons

    Set dbtemp = OpenDatabase(wrutatemp & "\temp_com.mdb")
    Set temprogas = dbtemp.OpenRecordset("temgapro")
    dbtemp.Execute ("delete from Temgapro")

    If xmodo = "0" Then
        ctitu = "DEL MES DE " & UCase(dev_mes(xmes))
    Else
        ctitu = "AL MES DE " & UCase(dev_mes(xmes))
    End If
    If xmoneda = "0" Then
        ctitu = ctitu & " - SOLES"
    Else
        ctitu = ctitu & " - DOLARES"
    End If

    lbltitulo.Caption = ctitu
    Select Case xtipo
        Case "0":
            lblbi.Visible = True
            pnlbi.Visible = True
            lbligv.Visible = True
            pnligv.Visible = True
            lblexo.Visible = True
            pnlexo.Visible = True
            lbltotal.Visible = True
            PnlTotal.Visible = True
            nbi = 0#: nigv = 0#: ntotal = 0#: nexo = 0#

            If xmodo = "0" Then
                Set CONSULTA = dbcompras.OpenRecordset("select * from regisdoc where trim(f4codprv)='" + Trim(xpro) + "' and trim(f4mesmov)='" + Trim(xmes) + "' order by f4mesmov,f4nummov")
            Else
                If Len(Trim(xpro)) = 0 Then
                    Set CONSULTA = dbcompras.OpenRecordset("select * from regisdoc where trim(f4mesmov)<='" + Trim(xmes) + "' order by f4mesmov,f4nummov")
                Else
                    Set CONSULTA = dbcompras.OpenRecordset("select * from regisdoc where trim(f4codprv)='" + Trim(xpro) + "' and trim(f4mesmov)<='" + Trim(xmes) + "' order by f4mesmov,f4nummov")
                End If
            End If
            
            Do While Not CONSULTA.EOF
                temprogas.AddNew
                If Len(Trim(xpro)) = 0 Then
                    panel.Caption = "TODOS"
                Else
                    panel.Caption = "" & CONSULTA.Fields("f4codprv") & " - " & CONSULTA.Fields("f4nomprv")
                End If
                 temprogas.Fields("fecha") = "" & CONSULTA.Fields("f4fecha")
                temprogas.Fields("comprobante") = "" & CONSULTA.Fields("f4nummov")
                temprogas.Fields("numdoc") = "" & CONSULTA.Fields("f4serdoc") & "/" & CONSULTA.Fields("f4numdoc")
                temprogas.Fields("detalle") = "" & CONSULTA.Fields("f4refere")
                temprogas.Fields("empresa") = UCase(wnomcia)
                temprogas.Fields("TITULO") = panel.Caption
                temprogas.Fields("TITU") = ctitu
                temprogas.Fields("tc") = Val("" & CONSULTA.Fields("f4tipcam"))
                If Trim(xmoneda) = "0" Then
                    If Trim(CONSULTA.Fields("f4moneda")) = "S" Then
                        temprogas.Fields("BasImp") = "" & Format(CONSULTA.Fields("f4basimp"), "###,###,##0.00")
                        temprogas.Fields("IGV") = "" & Format(CONSULTA.Fields("f4igv"), "###,###,##0.00")
                        temprogas.Fields("Exon") = "" & Format(CONSULTA.Fields("f4monina"), "###,###,##0.00")
                        temprogas.Fields("Pcompra") = "" & Format(CONSULTA.Fields("f4total"), "###,###,##0.00")
                    Else
                        temprogas.Fields("BasImp") = "" & Format(Format(CONSULTA.Fields("f4basimp"), "#0.00") * Format(CONSULTA.Fields("f4tipcam"), "#0.000"), "###,###,##0.00")
                        temprogas.Fields("IGV") = "" & Format(Format(CONSULTA.Fields("f4igv"), "#0.00") * Format(CONSULTA.Fields("f4tipcam"), "#0.000"), "###,###,##0.00")
                        temprogas.Fields("Exon") = "" & Format(Format(CONSULTA.Fields("f4monina"), "#0.00") * Format(CONSULTA.Fields("f4tipcam"), "#0.000"), "###,###,##0.00")
                        temprogas.Fields("Pcompra") = "" & Format(Format(CONSULTA.Fields("f4total"), "#0.00") * Format(CONSULTA.Fields("f4tipcam"), "#0.000"), "###,###,##0.00")
                    End If
                    temprogas.Fields("moneda") = "SOLES"
                Else
                    If Trim(CONSULTA.Fields("f4moneda")) = "D" Then
                        temprogas.Fields("BasImp") = "" & Format(CONSULTA.Fields("f4basimp"), "###,###,##0.00")
                        temprogas.Fields("IGV") = "" & Format(CONSULTA.Fields("f4igv"), "###,###,##0.00")
                        temprogas.Fields("Exon") = "" & Format(CONSULTA.Fields("f4monina"), "###,###,##0.00")
                        temprogas.Fields("Pcompra") = "" & Format(CONSULTA.Fields("f4total"), "###,###,##0.00")
                    Else
                        temprogas.Fields("BasImp") = "" & Format(Format(CONSULTA.Fields("f4basimp"), "#0.00") / Format(CONSULTA.Fields("f4tipcam"), "#0.000"), "###,###,##0.00")
                        temprogas.Fields("IGV") = "" & Format(Format(CONSULTA.Fields("f4igv"), "#0.00") / Format(CONSULTA.Fields("f4tipcam"), "#0.000"), "###,###,##0.00")
                        temprogas.Fields("Exon") = "" & Format(Format(CONSULTA.Fields("f4monina"), "#0.00") / Format(CONSULTA.Fields("f4tipcam"), "#0.000"), "###,###,##0.00")
                        temprogas.Fields("Pcompra") = "" & Format(Format(CONSULTA.Fields("f4total"), "#0.00") / Format(CONSULTA.Fields("f4tipcam"), "#0.000"), "###,###,##0.00")
                    End If
                    temprogas.Fields("moneda") = "DOLARES"
                End If
                nbi = nbi + temprogas.Fields("BasImp")
                nigv = nigv + temprogas.Fields("igv")
                nexo = nexo + temprogas.Fields("exon")
                ntotal = ntotal + temprogas.Fields("Pcompra")

                CONSULTA.MoveNext
                temprogas.Update
                If CONSULTA.EOF Then Exit Do
                'Grid1.Rows = Grid1.Rows + 1
            Loop

            pnlbi.Caption = Format(nbi, "###,###,##0.00")
            pnligv.Caption = Format(nigv, "###,###,##0.00")
            pnlexo.Caption = Format(nexo, "###,###,##0.00")
            PnlTotal.Caption = Format(ntotal, "###,###,##0.00")
                    
        Case "1":
            lblbi.Visible = True
            pnlbi.Visible = True
            lbligv.Visible = False
            pnligv.Visible = False
            lblexo.Visible = False
            pnlexo.Visible = False
            lbltotal.Visible = False
            PnlTotal.Visible = False
            nbi = 0#
            
            If xmodo = "0" Then
                Set CONSULTA = dbcompras.OpenRecordset("select * from regisdoc where trim(f4codprv)='" + Trim(xpro) + "' and trim(f4mesmov)='" + Trim(xmes) + "' order by f4mesmov,f4nummov")
            Else
                Set CONSULTA = dbcompras.OpenRecordset("select * from regisdoc where trim(f4codprv)='" + Trim(xpro) + "' and trim(f4mesmov)<='" + Trim(xmes) + "' order by f4mesmov,f4nummov")
            End If
            Do While Not CONSULTA.EOF
                panel.Caption = "" & CONSULTA.Fields("f4codprv") & " - " & CONSULTA.Fields("f4nomprv")
                Set consulta1 = dbcompras.OpenRecordset("select * from regismov where trim(f4mesmov)='" + Trim(CONSULTA.Fields("f4mesmov")) + "' and trim(f4nummov)='" + Trim(CONSULTA.Fields("f4nummov")) + "' order by f4mesmov,f4nummov")
                Do While Not consulta1.EOF
                    temprogas.AddNew
                    temprogas.Fields("fecha") = "" & CONSULTA.Fields("f4fecha")
                    temprogas.Fields("comprobante") = "" & CONSULTA.Fields("f4nummov")
                    temprogas.Fields("numdoc") = "" & CONSULTA.Fields("f4serdoc") & "/" & CONSULTA.Fields("f4numdoc")
                    temprogas.Fields("ctacont") = "" & Trim(consulta1.Fields("f3ctacon"))
                    temprogas.Fields("centro") = "" & Trim(consulta1.Fields("f3cencos"))
                    temprogas.Fields("detalle") = "" & Trim(consulta1.Fields("f3concepto"))
                    temprogas.Fields("empresa") = UCase(wnomcia)
                    temprogas.Fields("TITULO") = panel.Caption
                    temprogas.Fields("TITU") = ctitu
                    temprogas.Fields("tc") = Val("" & CONSULTA.Fields("f4tipcam"))
            
                    '-------------------------------------------
                    xtipdoc = ""
                    TbDocumento.Index = "IDCODDOC"
                    TbDocumento.Seek "=", CONSULTA.Fields("f4tipdoc") & ""
                    If Not TbDocumento.NoMatch Then
                        xtipdoc = UCase(TbDocumento.Fields("f2abrev") & "")
                    End If
                    '-------------------------------------------
                    If Trim(xmoneda) = "0" Then
                        If xtipdoc = "CRE" Then
                            If Trim(CONSULTA.Fields("f4moneda")) = "S" Then
                              
                                temprogas.Fields("pcompra") = consulta1.Fields("f3importe") * -1
                            Else
                              
                                temprogas.Fields("pcompra") = "" & Format((Format(consulta1.Fields("f3importe"), "#0.00") * Format(CONSULTA.Fields("f4tipcam"), "#0.000") * -1), "###,###,##0.00")
                            End If
                        Else
                            If Trim(CONSULTA.Fields("f4moneda")) = "S" Then
                              
                                temprogas.Fields("pcompra") = consulta1.Fields("f3importe")
                            Else
                              
                                temprogas.Fields("pcompra") = "" & Format(Format(consulta1.Fields("f3importe"), "#0.00") * Format(CONSULTA.Fields("f4tipcam"), "#0.000"), "###,###,##0.00")
                            End If
                        End If
                        temprogas.Fields("moneda") = "SOLES"
                    Else
                        If xtipdoc = "CRE" Then
                            If Trim(CONSULTA.Fields("f4moneda")) = "D" Then
                                temprogas.Fields("pcompra") = consulta1.Fields("f3importe") * -1
                            Else
                                temprogas.Fields("pcompra") = "" & Format((Format(consulta1.Fields("f3importe"), "#0.00") / Format(CONSULTA.Fields("f4tipcam"), "#0.000") * -1), "###,###,##0.00")
                            End If
                        Else
                            If Trim(CONSULTA.Fields("f4moneda")) = "D" Then
                                temprogas.Fields("pcompra") = consulta1.Fields("f3importe")
                            Else
                                temprogas.Fields("pcompra") = "" & Format(Format(consulta1.Fields("f3importe"), "#0.00") / Format(CONSULTA.Fields("f4tipcam"), "#0.000"), "###,###,##0.00")
                            End If
                        End If
                        temprogas.Fields("moneda") = "DOLARES"
                    End If

                    nbi = nbi + temprogas.Fields("pcompra")

                    temprogas.Update
                    consulta1.MoveNext
                    If consulta1.EOF Then Exit Do
                Loop
                CONSULTA.MoveNext
                If CONSULTA.EOF Then Exit Do
            Loop

            pnlbi.Caption = Format(nbi, "###,###,##0.00")

        Case "2":
            lblbi.Visible = True
            pnlbi.Visible = True
            lbligv.Visible = False
            pnligv.Visible = False
            lblexo.Visible = False
            pnlexo.Visible = False
            lbltotal.Visible = False
            PnlTotal.Visible = False
            nbi = 0#
            
            If xmodo = "0" Then
                If Len(Trim(xpro)) = 0 Then
                    Set CONSULTA = dbcompras.OpenRecordset("select * from regisdoc where trim(f4mesmov)='" + Trim(xmes) + "' order by f4mesmov,f4nummov")
                Else
                    Set CONSULTA = dbcompras.OpenRecordset("select * from regisdoc where trim(f4codprv)='" + Trim(xpro) + "' and trim(f4mesmov)='" + Trim(xmes) + "' order by f4mesmov,f4nummov")
                End If
            Else
                If Len(Trim(xpro)) = 0 Then
                    Set CONSULTA = dbcompras.OpenRecordset("select * from regisdoc where trim(f4mesmov) <='" + Trim(xmes) + "' order by f4mesmov,f4nummov")
                Else
                    Set CONSULTA = dbcompras.OpenRecordset("select * from regisdoc where trim(f4codprv)='" + Trim(xpro) + "' and trim(f4mesmov)<='" + Trim(xmes) + "' order by f4mesmov,f4nummov")
                End If
            End If
            Do While Not CONSULTA.EOF
                If Len(Trim(xpro)) <> 0 Then
                    panel.Caption = "" & CONSULTA.Fields("f4codprv") & " - " & CONSULTA.Fields("f4nomprv")
                Else
                    panel.Caption = ""
                End If
                Set consulta1 = dbcompras.OpenRecordset("select * from regismov where trim(f4mesmov)='" + Trim(CONSULTA.Fields("f4mesmov")) + "' and trim(f4nummov)='" + Trim(CONSULTA.Fields("f4nummov")) + "' and trim(f3gasto)='" + Trim(xgasto) + "' order by f4mesmov,f4nummov")
                Do While Not consulta1.EOF
                    temprogas.AddNew
                    temprogas.Fields("fecha") = "" & CONSULTA.Fields("f4fecha")
                    temprogas.Fields("comprobante") = "" & CONSULTA.Fields("f4nummov")
                    temprogas.Fields("numdoc") = "" & CONSULTA.Fields("f4serdoc") & "/" & CONSULTA.Fields("f4numdoc")
                    temprogas.Fields("ctacont") = "" & Trim(consulta1.Fields("f3ctacon"))
                    temprogas.Fields("orden") = "" & Trim(consulta1.Fields("f3cencos"))
                    temprogas.Fields("detalle") = "" & Trim(consulta1.Fields("f3concepto"))
                    temprogas.Fields("empresa") = UCase(wnomcia)
                    temprogas.Fields("TITULO") = panel.Caption
                    temprogas.Fields("TITU") = ctitu
                    temprogas.Fields("tc") = Val("" & CONSULTA.Fields("f4tipcam"))
                    
                    '-------------------------------------------
                    xtipdoc = ""
                    TbDocumento.Index = "IDCODDOC"
                    TbDocumento.Seek "=", CONSULTA.Fields("f4tipdoc") & ""
                    If Not TbDocumento.NoMatch Then
                        xtipdoc = UCase(TbDocumento.Fields("f2abrev") & "")
                    End If
                    '-------------------------------------------
                    If Trim(xmoneda) = "0" Then
                        If xtipdoc = "CRE" Then
                            If Trim(CONSULTA.Fields("f4moneda")) = "S" Then
                                temprogas.Fields("pcompra") = consulta1.Fields("f3importe") * -1
                            Else
                                temprogas.Fields("pcompra") = "" & Format((Format(consulta1.Fields("f3importe"), "#0.00") * Format(CONSULTA.Fields("f4tipcam"), "#0.000") * -1), "###,###,##0.00")
                            End If
                        Else
                            If Trim(CONSULTA.Fields("f4moneda")) = "S" Then
                                temprogas.Fields("pcompra") = consulta1.Fields("f3importe")
                            Else
                                temprogas.Fields("pcompra") = "" & Format(Format(consulta1.Fields("f3importe"), "#0.00") * Format(CONSULTA.Fields("f4tipcam"), "#0.000"), "###,###,##0.00")
                            End If
                        End If
                        temprogas.Fields("moneda") = "SOLES"
                    Else
                        If xtipdoc = "CRE" Then
                            If Trim(CONSULTA.Fields("f4moneda")) = "D" Then
                                temprogas.Fields("pcompra") = consulta1.Fields("f3importe") * -1
                            Else
                                temprogas.Fields("pcompra") = "" & Format((Format(consulta1.Fields("f3importe"), "#0.00") / Format(CONSULTA.Fields("f4tipcam"), "#0.000") * -1), "###,###,##0.00")
                            End If
                        Else
                            If Trim(CONSULTA.Fields("f4moneda")) = "D" Then
                                temprogas.Fields("pcompra") = consulta1.Fields("f3importe")
                            Else
                                temprogas.Fields("pcompra") = "" & Format(Format(consulta1.Fields("f3importe"), "#0.00") / Format(CONSULTA.Fields("f4tipcam"), "#0.000"), "###,###,##0.00")
                            End If
                        End If
                        temprogas.Fields("moneda") = "DOLARES"
                    End If

                    nbi = nbi + temprogas.Fields("pcompra")

                    temprogas.Update
                    consulta1.MoveNext
                    If consulta1.EOF Then Exit Do
                Loop
                CONSULTA.MoveNext
                If CONSULTA.EOF Then Exit Do
            Loop

            pnlbi.Caption = Format(nbi, "###,###,##0.00")

    End Select

    temprogas.Close
    dbtemp.Close

    Exit Sub

error_cons:
    Resume Next
               
End Sub
