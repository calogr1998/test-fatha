VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmayuopera 
   Caption         =   "Ayuda de Operaciones"
   ClientHeight    =   4305
   ClientLeft      =   840
   ClientTop       =   1980
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   9315
   Begin Threed.SSPanel SSPanel1 
      Height          =   4200
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   9195
      _Version        =   65536
      _ExtentX        =   16219
      _ExtentY        =   7408
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Begin VB.TextBox txtope 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   180
         TabIndex        =   2
         ToolTipText     =   "Simplifique la consulta digitando el Simbolo (*)"
         Top             =   3825
         Visible         =   0   'False
         Width           =   2445
      End
      Begin VB.Data dc_opera 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   1080
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1215
         Visible         =   0   'False
         Width           =   1680
      End
      Begin TrueOleDBGrid70.TDBGrid DBGrid1 
         Bindings        =   "frmayuopera.frx":0000
         Height          =   3075
         Left            =   135
         TabIndex        =   1
         Top             =   90
         Width           =   8925
         _ExtentX        =   15743
         _ExtentY        =   5424
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Código"
         Columns(0).DataField=   "CODIGO"
         Columns(0).DataWidth=   255
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "                               Descripcion"
         Columns(1).DataField=   "TIPO"
         Columns(1).DataWidth=   255
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Destino"
         Columns(2).DataField=   "DESTINO"
         Columns(2).DataWidth=   255
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).DataField=   ""
         Columns(3).DataWidth=   255
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "DIAS"
         Columns(4).DataField=   "DIAS"
         Columns(4).DataWidth=   11
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "MARCHE"
         Columns(5).DataField=   "MARCHE"
         Columns(5).DataWidth=   255
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "TIPDOCU"
         Columns(6).DataField=   "TIPDOCU"
         Columns(6).DataWidth=   255
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).DataField=   ""
         Columns(7).DataWidth=   255
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "TIPOPE"
         Columns(8).DataField=   "TIPOPE"
         Columns(8).DataWidth=   255
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "TRANSFER"
         Columns(9).DataField=   "TRANSFER"
         Columns(9).DataWidth=   255
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "CODIGTO"
         Columns(10).DataField=   "CODIGTO"
         Columns(10).DataWidth=   255
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "tipo_operacion"
         Columns(11).DataField=   "tipo_operacion"
         Columns(11).DataWidth=   255
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   0
         Columns(12)._MaxComboItems=   5
         Columns(12).Caption=   "tipo_letra"
         Columns(12).DataField=   "tipo_letra"
         Columns(12).DataWidth=   255
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(13)._VlistStyle=   0
         Columns(13)._MaxComboItems=   5
         Columns(13).Caption=   "DIFERIDO"
         Columns(13).DataField=   "DIFERIDO"
         Columns(13).DataWidth=   255
         Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(14)._VlistStyle=   0
         Columns(14)._MaxComboItems=   5
         Columns(14).Caption=   "CHEQUES_DEV"
         Columns(14).DataField=   "CHEQUES_DEV"
         Columns(14).DataWidth=   255
         Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   15
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   2
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   2
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=15"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2275"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2196"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=74000"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=11695"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=11615"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).AllowSizing=0"
         Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=74000"
         Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(15)=   "Column(2).Width=1323"
         Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=1244"
         Splits(0)._ColumnProps(18)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(19)=   "Column(2).AllowSizing=0"
         Splits(0)._ColumnProps(20)=   "Column(2)._ColStyle=74000"
         Splits(0)._ColumnProps(21)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(22)=   "Column(3).Width=5503"
         Splits(0)._ColumnProps(23)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(24)=   "Column(3)._WidthInPix=5424"
         Splits(0)._ColumnProps(25)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(26)=   "Column(3).AllowSizing=0"
         Splits(0)._ColumnProps(27)=   "Column(3)._ColStyle=74000"
         Splits(0)._ColumnProps(28)=   "Column(3).Visible=0"
         Splits(0)._ColumnProps(29)=   "Column(3).AllowFocus=0"
         Splits(0)._ColumnProps(30)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(31)=   "Column(4).Width=1535"
         Splits(0)._ColumnProps(32)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(4)._WidthInPix=1455"
         Splits(0)._ColumnProps(34)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(35)=   "Column(4).AllowSizing=0"
         Splits(0)._ColumnProps(36)=   "Column(4)._ColStyle=73730"
         Splits(0)._ColumnProps(37)=   "Column(4).Visible=0"
         Splits(0)._ColumnProps(38)=   "Column(4).AllowFocus=0"
         Splits(0)._ColumnProps(39)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(40)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(41)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(42)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(43)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(44)=   "Column(5).AllowSizing=0"
         Splits(0)._ColumnProps(45)=   "Column(5)._ColStyle=74000"
         Splits(0)._ColumnProps(46)=   "Column(5).Visible=0"
         Splits(0)._ColumnProps(47)=   "Column(5).AllowFocus=0"
         Splits(0)._ColumnProps(48)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(49)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(50)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(51)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(52)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(53)=   "Column(6).AllowSizing=0"
         Splits(0)._ColumnProps(54)=   "Column(6)._ColStyle=65808"
         Splits(0)._ColumnProps(55)=   "Column(6).Visible=0"
         Splits(0)._ColumnProps(56)=   "Column(6).AllowFocus=0"
         Splits(0)._ColumnProps(57)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(58)=   "Column(7).Width=2725"
         Splits(0)._ColumnProps(59)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(60)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(61)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(62)=   "Column(7).AllowSizing=0"
         Splits(0)._ColumnProps(63)=   "Column(7)._ColStyle=65808"
         Splits(0)._ColumnProps(64)=   "Column(7).Visible=0"
         Splits(0)._ColumnProps(65)=   "Column(7).AllowFocus=0"
         Splits(0)._ColumnProps(66)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(67)=   "Column(8).Width=2725"
         Splits(0)._ColumnProps(68)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(69)=   "Column(8)._WidthInPix=2646"
         Splits(0)._ColumnProps(70)=   "Column(8)._EditAlways=0"
         Splits(0)._ColumnProps(71)=   "Column(8).AllowSizing=0"
         Splits(0)._ColumnProps(72)=   "Column(8)._ColStyle=74000"
         Splits(0)._ColumnProps(73)=   "Column(8).Visible=0"
         Splits(0)._ColumnProps(74)=   "Column(8).AllowFocus=0"
         Splits(0)._ColumnProps(75)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(76)=   "Column(9).Width=2725"
         Splits(0)._ColumnProps(77)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(78)=   "Column(9)._WidthInPix=2646"
         Splits(0)._ColumnProps(79)=   "Column(9)._EditAlways=0"
         Splits(0)._ColumnProps(80)=   "Column(9).AllowSizing=0"
         Splits(0)._ColumnProps(81)=   "Column(9)._ColStyle=65808"
         Splits(0)._ColumnProps(82)=   "Column(9).Visible=0"
         Splits(0)._ColumnProps(83)=   "Column(9).AllowFocus=0"
         Splits(0)._ColumnProps(84)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(85)=   "Column(10).Width=2725"
         Splits(0)._ColumnProps(86)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(87)=   "Column(10)._WidthInPix=2646"
         Splits(0)._ColumnProps(88)=   "Column(10)._EditAlways=0"
         Splits(0)._ColumnProps(89)=   "Column(10).AllowSizing=0"
         Splits(0)._ColumnProps(90)=   "Column(10)._ColStyle=65808"
         Splits(0)._ColumnProps(91)=   "Column(10).Visible=0"
         Splits(0)._ColumnProps(92)=   "Column(10).AllowFocus=0"
         Splits(0)._ColumnProps(93)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(94)=   "Column(11).Width=2725"
         Splits(0)._ColumnProps(95)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(96)=   "Column(11)._WidthInPix=2646"
         Splits(0)._ColumnProps(97)=   "Column(11)._EditAlways=0"
         Splits(0)._ColumnProps(98)=   "Column(11).AllowSizing=0"
         Splits(0)._ColumnProps(99)=   "Column(11)._ColStyle=74000"
         Splits(0)._ColumnProps(100)=   "Column(11).Visible=0"
         Splits(0)._ColumnProps(101)=   "Column(11).AllowFocus=0"
         Splits(0)._ColumnProps(102)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(103)=   "Column(12).Width=2725"
         Splits(0)._ColumnProps(104)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(105)=   "Column(12)._WidthInPix=2646"
         Splits(0)._ColumnProps(106)=   "Column(12)._EditAlways=0"
         Splits(0)._ColumnProps(107)=   "Column(12).AllowSizing=0"
         Splits(0)._ColumnProps(108)=   "Column(12)._ColStyle=65808"
         Splits(0)._ColumnProps(109)=   "Column(12).Visible=0"
         Splits(0)._ColumnProps(110)=   "Column(12).Order=13"
         Splits(0)._ColumnProps(111)=   "Column(13).Width=2725"
         Splits(0)._ColumnProps(112)=   "Column(13).DividerColor=0"
         Splits(0)._ColumnProps(113)=   "Column(13)._WidthInPix=2646"
         Splits(0)._ColumnProps(114)=   "Column(13)._EditAlways=0"
         Splits(0)._ColumnProps(115)=   "Column(13).AllowSizing=0"
         Splits(0)._ColumnProps(116)=   "Column(13)._ColStyle=65808"
         Splits(0)._ColumnProps(117)=   "Column(13).Visible=0"
         Splits(0)._ColumnProps(118)=   "Column(13).AllowFocus=0"
         Splits(0)._ColumnProps(119)=   "Column(13).Order=14"
         Splits(0)._ColumnProps(120)=   "Column(14).Width=2725"
         Splits(0)._ColumnProps(121)=   "Column(14).DividerColor=0"
         Splits(0)._ColumnProps(122)=   "Column(14)._WidthInPix=2646"
         Splits(0)._ColumnProps(123)=   "Column(14)._EditAlways=0"
         Splits(0)._ColumnProps(124)=   "Column(14).AllowSizing=0"
         Splits(0)._ColumnProps(125)=   "Column(14)._ColStyle=65808"
         Splits(0)._ColumnProps(126)=   "Column(14).Visible=0"
         Splits(0)._ColumnProps(127)=   "Column(14).AllowFocus=0"
         Splits(0)._ColumnProps(128)=   "Column(14).Order=15"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
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
         DeadAreaBackColor=   13160660
         RowDividerColor =   13160660
         RowSubDividerColor=   13160660
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=0,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&H8000000F&"
         _StyleDefs(9)   =   ":id=2,.fgcolor=&H80000012&"
         _StyleDefs(10)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(11)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(12)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(13)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(14)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(15)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(16)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(17)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(18)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(19)  =   "Splits(0).Style:id=43,.parent=1"
         _StyleDefs(20)  =   "Splits(0).CaptionStyle:id=52,.parent=4"
         _StyleDefs(21)  =   "Splits(0).HeadingStyle:id=44,.parent=2"
         _StyleDefs(22)  =   "Splits(0).FooterStyle:id=45,.parent=3"
         _StyleDefs(23)  =   "Splits(0).InactiveStyle:id=46,.parent=5"
         _StyleDefs(24)  =   "Splits(0).SelectedStyle:id=48,.parent=6"
         _StyleDefs(25)  =   "Splits(0).EditorStyle:id=47,.parent=7"
         _StyleDefs(26)  =   "Splits(0).HighlightRowStyle:id=49,.parent=8"
         _StyleDefs(27)  =   "Splits(0).EvenRowStyle:id=50,.parent=9"
         _StyleDefs(28)  =   "Splits(0).OddRowStyle:id=51,.parent=10"
         _StyleDefs(29)  =   "Splits(0).RecordSelectorStyle:id=53,.parent=11"
         _StyleDefs(30)  =   "Splits(0).FilterBarStyle:id=54,.parent=12"
         _StyleDefs(31)  =   "Splits(0).Columns(0).Style:id=28,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(32)  =   ":id=28,.locked=-1"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=44"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=45"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=47"
         _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(37)  =   ":id=32,.locked=-1"
         _StyleDefs(38)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=44"
         _StyleDefs(39)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=45"
         _StyleDefs(40)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=47"
         _StyleDefs(41)  =   "Splits(0).Columns(2).Style:id=58,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(42)  =   ":id=58,.locked=-1"
         _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=44"
         _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=45"
         _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=47"
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=62,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(47)  =   ":id=62,.locked=-1"
         _StyleDefs(48)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=44"
         _StyleDefs(49)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=45"
         _StyleDefs(50)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=47"
         _StyleDefs(51)  =   "Splits(0).Columns(4).Style:id=66,.parent=43,.alignment=1,.locked=-1"
         _StyleDefs(52)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=44,.alignment=3"
         _StyleDefs(53)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=45"
         _StyleDefs(54)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=47"
         _StyleDefs(55)  =   "Splits(0).Columns(5).Style:id=70,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(56)  =   ":id=70,.locked=-1"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=67,.parent=44"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=68,.parent=45"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=69,.parent=47"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=74,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=71,.parent=44"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=72,.parent=45"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=73,.parent=47"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=78,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=75,.parent=44"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=76,.parent=45"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=77,.parent=47"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=82,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(69)  =   ":id=82,.locked=-1"
         _StyleDefs(70)  =   "Splits(0).Columns(8).HeadingStyle:id=79,.parent=44"
         _StyleDefs(71)  =   "Splits(0).Columns(8).FooterStyle:id=80,.parent=45"
         _StyleDefs(72)  =   "Splits(0).Columns(8).EditorStyle:id=81,.parent=47"
         _StyleDefs(73)  =   "Splits(0).Columns(9).Style:id=86,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(74)  =   "Splits(0).Columns(9).HeadingStyle:id=83,.parent=44"
         _StyleDefs(75)  =   "Splits(0).Columns(9).FooterStyle:id=84,.parent=45"
         _StyleDefs(76)  =   "Splits(0).Columns(9).EditorStyle:id=85,.parent=47"
         _StyleDefs(77)  =   "Splits(0).Columns(10).Style:id=90,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(78)  =   "Splits(0).Columns(10).HeadingStyle:id=87,.parent=44"
         _StyleDefs(79)  =   "Splits(0).Columns(10).FooterStyle:id=88,.parent=45"
         _StyleDefs(80)  =   "Splits(0).Columns(10).EditorStyle:id=89,.parent=47"
         _StyleDefs(81)  =   "Splits(0).Columns(11).Style:id=94,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(82)  =   ":id=94,.locked=-1"
         _StyleDefs(83)  =   "Splits(0).Columns(11).HeadingStyle:id=91,.parent=44"
         _StyleDefs(84)  =   "Splits(0).Columns(11).FooterStyle:id=92,.parent=45"
         _StyleDefs(85)  =   "Splits(0).Columns(11).EditorStyle:id=93,.parent=47"
         _StyleDefs(86)  =   "Splits(0).Columns(12).Style:id=98,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(87)  =   "Splits(0).Columns(12).HeadingStyle:id=95,.parent=44"
         _StyleDefs(88)  =   "Splits(0).Columns(12).FooterStyle:id=96,.parent=45"
         _StyleDefs(89)  =   "Splits(0).Columns(12).EditorStyle:id=97,.parent=47"
         _StyleDefs(90)  =   "Splits(0).Columns(13).Style:id=102,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(91)  =   "Splits(0).Columns(13).HeadingStyle:id=99,.parent=44"
         _StyleDefs(92)  =   "Splits(0).Columns(13).FooterStyle:id=100,.parent=45"
         _StyleDefs(93)  =   "Splits(0).Columns(13).EditorStyle:id=101,.parent=47"
         _StyleDefs(94)  =   "Splits(0).Columns(14).Style:id=106,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(95)  =   "Splits(0).Columns(14).HeadingStyle:id=103,.parent=44"
         _StyleDefs(96)  =   "Splits(0).Columns(14).FooterStyle:id=104,.parent=45"
         _StyleDefs(97)  =   "Splits(0).Columns(14).EditorStyle:id=105,.parent=47"
         _StyleDefs(98)  =   "Named:id=33:Normal"
         _StyleDefs(99)  =   ":id=33,.parent=0"
         _StyleDefs(100) =   "Named:id=34:Heading"
         _StyleDefs(101) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(102) =   ":id=34,.wraptext=-1"
         _StyleDefs(103) =   "Named:id=35:Footing"
         _StyleDefs(104) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(105) =   "Named:id=36:Selected"
         _StyleDefs(106) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(107) =   "Named:id=37:Caption"
         _StyleDefs(108) =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(109) =   "Named:id=38:HighlightRow"
         _StyleDefs(110) =   ":id=38,.parent=33,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
         _StyleDefs(111) =   "Named:id=39:EvenRow"
         _StyleDefs(112) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(113) =   "Named:id=40:OddRow"
         _StyleDefs(114) =   ":id=40,.parent=33"
         _StyleDefs(115) =   "Named:id=41:RecordSelector"
         _StyleDefs(116) =   ":id=41,.parent=34"
         _StyleDefs(117) =   "Named:id=42:FilterBar"
         _StyleDefs(118) =   ":id=42,.parent=33"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "[F3] ---> Filtro Avanzado: Descripción"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   135
         TabIndex        =   5
         Top             =   3195
         Width           =   2610
      End
      Begin VB.Label Label1 
         Caption         =   "Ayuda: Hacer Click Derecho en la Columna que desee buscar"
         ForeColor       =   &H00800000&
         Height          =   555
         Left            =   6615
         TabIndex        =   4
         Top             =   3555
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
         Left            =   135
         TabIndex        =   3
         Top             =   3555
         Visible         =   0   'False
         Width           =   60
      End
   End
   Begin VB.Menu MnuPri 
      Caption         =   ""
      Begin VB.Menu mnufiltro 
         Caption         =   "Filtrar"
      End
      Begin VB.Menu mnufiltroavanz 
         Caption         =   "Filtro Avanzado:"
      End
      Begin VB.Menu MnuOrdAsc 
         Caption         =   "Ord. Asc"
      End
      Begin VB.Menu MnuOrdDesc 
         Caption         =   "Ord. Desc"
      End
      Begin VB.Menu MnuTodo 
         Caption         =   "Mostrar Todos"
      End
   End
End
Attribute VB_Name = "frmayuopera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dboperas    As DAO.Database
Dim t_dato      As DAO.Recordset
Dim rs          As DAO.Recordset
Dim xCodigo     As Integer

Private Sub DBGrid1_DblClick()

    DBGrid1_KeyPress 13
    
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 114 Then
     dbgrid1.col = 1
     mnufiltroavanz_Click
    End If
    
    DBGrid1_KeyPress KeyCode

End Sub

Private Sub DBGrid1_KeyPress(KeyAscii As Integer)

   Select Case KeyAscii
      Case 13:
         gcodtip = dbgrid1.Columns(0)
         Unload Me
      Case 27:
         gcodtip = ""
         Unload Me
   End Select

End Sub

Private Sub DBGrid1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Visible = True

End Sub

Private Sub DBGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If txtope.Visible = True Then txtope.Visible = False
    If lblbusca.Visible = True Then lblbusca.Visible = False
  mnufiltro.Caption = "Filtrar [" + dbgrid1.Columns(dbgrid1.col).Text + "]"
   Select Case Button
          Case 2
               PopupMenu mnupri
   End Select
End Sub

Private Sub Form_Activate()
    
    dbgrid1.OddRowStyle.BackColor = &HC0FFFF
    dbgrid1.EvenRowStyle.BackColor = &HFFFFFF
    
    dbgrid1.HighlightRowStyle.BackColor = vbActiveTitleBar
    dbgrid1.HighlightRowStyle.ForeColor = vbWhite
    dbgrid1.AlternatingRowStyle = True
    
End Sub

Private Sub Form_Load()
           
    Set dboperas = OpenDatabase(wrutabancos & "\DB_TABLA.MDB")
    Set t_dato = dboperas.OpenRecordset("BF8TMOV")
    
    dc_opera.DatabaseName = wrutabancos & "\db_tabla.mdb"
    dc_opera.RecordSource = "select * from bf8tmov where destino= '" & "E" & "' and tipope<> '" & "T" & "' order by codigo"
    dc_opera.Refresh
    dbgrid1.Refresh
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    t_dato.Close
    dboperas.Close

End Sub

Private Sub mnufiltro_Click()
  Set rs = dc_opera.Recordset
  Select Case dc_opera.Recordset.Fields(dbgrid1.col).Type
         Case 10
              rs.Filter = "[" + dc_opera.Recordset.Fields(dbgrid1.col).Name + "]" + " = '" + dbgrid1.Columns(dbgrid1.col).Text + "'"
         Case 4
              rs.Filter = "[" + dc_opera.Recordset.Fields(dbgrid1.col).Name + "]" + " = " + dbgrid1.Columns(dbgrid1.col).Text
         Case 8
              If IsDate(dbgrid1.Columns(dbgrid1.col).Text) Then
                 rs.Filter = "[" + dc_opera.Recordset.Fields(dbgrid1.col).Name + "]" + "=#" + dbgrid1.Columns(dbgrid1.col).Text + "#"
              Else
                 MsgBox "Ingrese una Fecha Valida..!", 32, "Advertencia"
                 Exit Sub
              End If
  End Select
  Set dc_opera.Recordset = rs.OpenRecordset(rs.Type)
  Set rs = Nothing
End Sub

Private Sub mnufiltroavanz_Click()
  Select Case dbgrid1.col
        Case 0:
            lblbusca.Visible = True
            lblbusca.Caption = dbgrid1.Columns(dbgrid1.col).Caption
            txtope.Visible = True
            txtope.Text = ""
            txtope.SetFocus
        Case 1
            lblbusca.Visible = True
            lblbusca.Caption = dbgrid1.Columns(dbgrid1.col).Caption
            'mkfecha.Visible = True
            'mkfecha.Text = Date
            txtope.Visible = True
            txtope.Text = ""
            'mkfecha.SetFocus
            txtope.SetFocus
        Case 2
            lblbusca.Visible = True
            lblbusca.Caption = dbgrid1.Columns(dbgrid1.col).Caption
            txtope.Visible = True
            txtope.Text = ""
            txtope.SetFocus
    End Select
End Sub

Private Sub mnuordasc_Click()
Set rs = dc_opera.Recordset
  rs.Sort = "[" + dc_opera.Recordset.Fields(dbgrid1.col).Name + "] Asc"
  Set dc_opera.Recordset = rs.OpenRecordset(rs.Type)
  Set rs = Nothing
End Sub

Private Sub mnuorddesc_Click()
  Set rs = dc_opera.Recordset
  rs.Sort = "[" + dc_opera.Recordset.Fields(dbgrid1.col).Name + "] Desc"
  Set dc_opera.Recordset = rs.OpenRecordset(rs.Type)
  Set rs = Nothing
End Sub

Private Sub MnuTodo_Click()
    
    dc_opera.DatabaseName = wrutabancos & "\db_tabla.mdb"
    dc_opera.RecordSource = "select * from bf8tmov where destino= '" & "E" & "' and tipope<> '" & "T" & "' order by codigo"
    dc_opera.Refresh
    
End Sub
Private Sub SSPanel1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Label1.Visible = False
End Sub

Private Sub txtope_KeyPress(KeyAscii As Integer)
Dim SQL     As String
Dim DATO    As String

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        txtope.Text = "*" & txtope.Text
       If Len(txtope) = 0 Then txtope.SetFocus: Exit Sub
          If InStr(txtope, "*") = Len(Trim(txtope)) Then
             DATO = Left(txtope, Len(Trim(txtope)) - 1)
          Else
              DATO = txtope.Text
          End If
          Set rs = dc_opera.Recordset
          rs.Filter = "[" + dbgrid1.Columns(dbgrid1.col).DataField + "]" + " Like  '" + DATO + "*'"
          If rs.EOF Then txtope.SetFocus: Exit Sub Else txtope.Visible = False: lblbusca.Visible = False
          Set dc_opera.Recordset = rs.OpenRecordset(rs.Type)
          Set rs = Nothing
         dbgrid1.SetFocus
    
    End If
End Sub
