VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmhlpctas 
   Caption         =   "Ayuda de Cuentas Corrientes"
   ClientHeight    =   4365
   ClientLeft      =   450
   ClientTop       =   1980
   ClientWidth     =   9240
   Icon            =   "frmhlpctas.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   9240
   Begin VB.Data Data_Ctas 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1890
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1350
      Visible         =   0   'False
      Width           =   1725
   End
   Begin TrueOleDBGrid70.TDBGrid DBGrid1 
      Bindings        =   "frmhlpctas.frx":030A
      Height          =   3210
      Left            =   135
      TabIndex        =   1
      Top             =   135
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   5662
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Código"
      Columns(0).DataField=   "CODCTA"
      Columns(0).DataWidth=   11
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "                       Banco"
      Columns(1).DataField=   "BANCO"
      Columns(1).DataWidth=   255
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Número de Cuenta"
      Columns(2).DataField=   "NUMCTA"
      Columns(2).DataWidth=   11
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "      Cuenta Contable"
      Columns(3).DataField=   "F5CODCTA"
      Columns(3).DataWidth=   255
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Moneda"
      Columns(4).DataField=   "MONEDA"
      Columns(4).DataWidth=   19
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).DataField=   ""
      Columns(5).DataWidth=   255
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).DataField=   ""
      Columns(6).DataWidth=   255
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Movimiento"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).DataField=   ""
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   16
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   2
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).ScrollBars=   2
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=16"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1667"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1588"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
      Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=73728"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=5556"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=5477"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1).AllowSizing=0"
      Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=74000"
      Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(15)=   "Column(2).Width=4260"
      Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=4180"
      Splits(0)._ColumnProps(18)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(19)=   "Column(2).AllowSizing=0"
      Splits(0)._ColumnProps(20)=   "Column(2)._ColStyle=73728"
      Splits(0)._ColumnProps(21)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(22)=   "Column(3).Width=2619"
      Splits(0)._ColumnProps(23)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(3)._WidthInPix=2540"
      Splits(0)._ColumnProps(25)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(26)=   "Column(3).AllowSizing=0"
      Splits(0)._ColumnProps(27)=   "Column(3)._ColStyle=74000"
      Splits(0)._ColumnProps(28)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(29)=   "Column(4).Width=1244"
      Splits(0)._ColumnProps(30)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(4)._WidthInPix=1164"
      Splits(0)._ColumnProps(32)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(33)=   "Column(4).AllowSizing=0"
      Splits(0)._ColumnProps(34)=   "Column(4)._ColStyle=73728"
      Splits(0)._ColumnProps(35)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(36)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(37)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(39)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(40)=   "Column(5).AllowSizing=0"
      Splits(0)._ColumnProps(41)=   "Column(5)._ColStyle=65808"
      Splits(0)._ColumnProps(42)=   "Column(5).Visible=0"
      Splits(0)._ColumnProps(43)=   "Column(5).AllowFocus=0"
      Splits(0)._ColumnProps(44)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(45)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(46)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(47)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(48)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(49)=   "Column(6).AllowSizing=0"
      Splits(0)._ColumnProps(50)=   "Column(6)._ColStyle=65808"
      Splits(0)._ColumnProps(51)=   "Column(6).Visible=0"
      Splits(0)._ColumnProps(52)=   "Column(6).AllowFocus=0"
      Splits(0)._ColumnProps(53)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(54)=   "Column(7).Width=1535"
      Splits(0)._ColumnProps(55)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(56)=   "Column(7)._WidthInPix=1455"
      Splits(0)._ColumnProps(57)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(58)=   "Column(7).AllowSizing=0"
      Splits(0)._ColumnProps(59)=   "Column(7)._ColStyle=65538"
      Splits(0)._ColumnProps(60)=   "Column(7).Visible=0"
      Splits(0)._ColumnProps(61)=   "Column(7).AllowFocus=0"
      Splits(0)._ColumnProps(62)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(63)=   "Column(8).Width=1667"
      Splits(0)._ColumnProps(64)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(65)=   "Column(8)._WidthInPix=1588"
      Splits(0)._ColumnProps(66)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(67)=   "Column(8).AllowSizing=0"
      Splits(0)._ColumnProps(68)=   "Column(8)._ColStyle=65538"
      Splits(0)._ColumnProps(69)=   "Column(8).Visible=0"
      Splits(0)._ColumnProps(70)=   "Column(8).AllowFocus=0"
      Splits(0)._ColumnProps(71)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(72)=   "Column(9).Width=2725"
      Splits(0)._ColumnProps(73)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(74)=   "Column(9)._WidthInPix=2646"
      Splits(0)._ColumnProps(75)=   "Column(9)._EditAlways=0"
      Splits(0)._ColumnProps(76)=   "Column(9).AllowSizing=0"
      Splits(0)._ColumnProps(77)=   "Column(9)._ColStyle=65808"
      Splits(0)._ColumnProps(78)=   "Column(9).Visible=0"
      Splits(0)._ColumnProps(79)=   "Column(9).AllowFocus=0"
      Splits(0)._ColumnProps(80)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(81)=   "Column(10).Width=2725"
      Splits(0)._ColumnProps(82)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(83)=   "Column(10)._WidthInPix=2646"
      Splits(0)._ColumnProps(84)=   "Column(10)._EditAlways=0"
      Splits(0)._ColumnProps(85)=   "Column(10).AllowSizing=0"
      Splits(0)._ColumnProps(86)=   "Column(10)._ColStyle=65808"
      Splits(0)._ColumnProps(87)=   "Column(10).Visible=0"
      Splits(0)._ColumnProps(88)=   "Column(10).AllowFocus=0"
      Splits(0)._ColumnProps(89)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(90)=   "Column(11).Width=2725"
      Splits(0)._ColumnProps(91)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(92)=   "Column(11)._WidthInPix=2646"
      Splits(0)._ColumnProps(93)=   "Column(11)._EditAlways=0"
      Splits(0)._ColumnProps(94)=   "Column(11).AllowSizing=0"
      Splits(0)._ColumnProps(95)=   "Column(11)._ColStyle=65808"
      Splits(0)._ColumnProps(96)=   "Column(11).Visible=0"
      Splits(0)._ColumnProps(97)=   "Column(11).AllowFocus=0"
      Splits(0)._ColumnProps(98)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(99)=   "Column(12).Width=2725"
      Splits(0)._ColumnProps(100)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(101)=   "Column(12)._WidthInPix=2646"
      Splits(0)._ColumnProps(102)=   "Column(12)._EditAlways=0"
      Splits(0)._ColumnProps(103)=   "Column(12).AllowSizing=0"
      Splits(0)._ColumnProps(104)=   "Column(12)._ColStyle=65808"
      Splits(0)._ColumnProps(105)=   "Column(12).Visible=0"
      Splits(0)._ColumnProps(106)=   "Column(12).AllowFocus=0"
      Splits(0)._ColumnProps(107)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(108)=   "Column(13).Width=2725"
      Splits(0)._ColumnProps(109)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(110)=   "Column(13)._WidthInPix=2646"
      Splits(0)._ColumnProps(111)=   "Column(13)._EditAlways=0"
      Splits(0)._ColumnProps(112)=   "Column(13).AllowSizing=0"
      Splits(0)._ColumnProps(113)=   "Column(13)._ColStyle=65808"
      Splits(0)._ColumnProps(114)=   "Column(13).Visible=0"
      Splits(0)._ColumnProps(115)=   "Column(13).AllowFocus=0"
      Splits(0)._ColumnProps(116)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(117)=   "Column(14).Width=2725"
      Splits(0)._ColumnProps(118)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(119)=   "Column(14)._WidthInPix=2646"
      Splits(0)._ColumnProps(120)=   "Column(14)._EditAlways=0"
      Splits(0)._ColumnProps(121)=   "Column(14).AllowSizing=0"
      Splits(0)._ColumnProps(122)=   "Column(14)._ColStyle=65808"
      Splits(0)._ColumnProps(123)=   "Column(14).Visible=0"
      Splits(0)._ColumnProps(124)=   "Column(14).AllowFocus=0"
      Splits(0)._ColumnProps(125)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(126)=   "Column(15).Width=2725"
      Splits(0)._ColumnProps(127)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(128)=   "Column(15)._WidthInPix=2646"
      Splits(0)._ColumnProps(129)=   "Column(15)._EditAlways=0"
      Splits(0)._ColumnProps(130)=   "Column(15).AllowSizing=0"
      Splits(0)._ColumnProps(131)=   "Column(15)._ColStyle=65538"
      Splits(0)._ColumnProps(132)=   "Column(15).Visible=0"
      Splits(0)._ColumnProps(133)=   "Column(15).AllowFocus=0"
      Splits(0)._ColumnProps(134)=   "Column(15).Order=16"
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=13,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(31)  =   "Splits(0).Columns(0).Style:id=28,.parent=43,.alignment=0,.locked=-1"
      _StyleDefs(32)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=44,.alignment=3"
      _StyleDefs(33)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=45"
      _StyleDefs(34)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=47"
      _StyleDefs(35)  =   "Splits(0).Columns(1).Style:id=32,.parent=43,.alignment=0,.valignment=2"
      _StyleDefs(36)  =   ":id=32,.locked=-1"
      _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=44"
      _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=45"
      _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=47"
      _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=58,.parent=43,.alignment=0,.locked=-1"
      _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=44,.alignment=3"
      _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=45"
      _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=47"
      _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=62,.parent=43,.alignment=0,.valignment=2"
      _StyleDefs(45)  =   ":id=62,.locked=-1"
      _StyleDefs(46)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=44"
      _StyleDefs(47)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=45"
      _StyleDefs(48)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=47"
      _StyleDefs(49)  =   "Splits(0).Columns(4).Style:id=66,.parent=43,.alignment=0,.locked=-1"
      _StyleDefs(50)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=44,.alignment=3"
      _StyleDefs(51)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=45"
      _StyleDefs(52)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=47"
      _StyleDefs(53)  =   "Splits(0).Columns(5).Style:id=70,.parent=43,.alignment=0,.valignment=2"
      _StyleDefs(54)  =   "Splits(0).Columns(5).HeadingStyle:id=67,.parent=44"
      _StyleDefs(55)  =   "Splits(0).Columns(5).FooterStyle:id=68,.parent=45"
      _StyleDefs(56)  =   "Splits(0).Columns(5).EditorStyle:id=69,.parent=47"
      _StyleDefs(57)  =   "Splits(0).Columns(6).Style:id=74,.parent=43,.alignment=0,.valignment=2"
      _StyleDefs(58)  =   "Splits(0).Columns(6).HeadingStyle:id=71,.parent=44"
      _StyleDefs(59)  =   "Splits(0).Columns(6).FooterStyle:id=72,.parent=45"
      _StyleDefs(60)  =   "Splits(0).Columns(6).EditorStyle:id=73,.parent=47"
      _StyleDefs(61)  =   "Splits(0).Columns(7).Style:id=78,.parent=43,.alignment=1"
      _StyleDefs(62)  =   "Splits(0).Columns(7).HeadingStyle:id=75,.parent=44,.alignment=3"
      _StyleDefs(63)  =   "Splits(0).Columns(7).FooterStyle:id=76,.parent=45"
      _StyleDefs(64)  =   "Splits(0).Columns(7).EditorStyle:id=77,.parent=47"
      _StyleDefs(65)  =   "Splits(0).Columns(8).Style:id=82,.parent=43,.alignment=1"
      _StyleDefs(66)  =   "Splits(0).Columns(8).HeadingStyle:id=79,.parent=44,.alignment=3"
      _StyleDefs(67)  =   "Splits(0).Columns(8).FooterStyle:id=80,.parent=45"
      _StyleDefs(68)  =   "Splits(0).Columns(8).EditorStyle:id=81,.parent=47"
      _StyleDefs(69)  =   "Splits(0).Columns(9).Style:id=86,.parent=43,.alignment=0,.valignment=2"
      _StyleDefs(70)  =   "Splits(0).Columns(9).HeadingStyle:id=83,.parent=44"
      _StyleDefs(71)  =   "Splits(0).Columns(9).FooterStyle:id=84,.parent=45"
      _StyleDefs(72)  =   "Splits(0).Columns(9).EditorStyle:id=85,.parent=47"
      _StyleDefs(73)  =   "Splits(0).Columns(10).Style:id=90,.parent=43,.alignment=0,.valignment=2"
      _StyleDefs(74)  =   "Splits(0).Columns(10).HeadingStyle:id=87,.parent=44"
      _StyleDefs(75)  =   "Splits(0).Columns(10).FooterStyle:id=88,.parent=45"
      _StyleDefs(76)  =   "Splits(0).Columns(10).EditorStyle:id=89,.parent=47"
      _StyleDefs(77)  =   "Splits(0).Columns(11).Style:id=94,.parent=43,.alignment=0,.valignment=2"
      _StyleDefs(78)  =   "Splits(0).Columns(11).HeadingStyle:id=91,.parent=44"
      _StyleDefs(79)  =   "Splits(0).Columns(11).FooterStyle:id=92,.parent=45"
      _StyleDefs(80)  =   "Splits(0).Columns(11).EditorStyle:id=93,.parent=47"
      _StyleDefs(81)  =   "Splits(0).Columns(12).Style:id=98,.parent=43,.alignment=0,.valignment=2"
      _StyleDefs(82)  =   "Splits(0).Columns(12).HeadingStyle:id=95,.parent=44"
      _StyleDefs(83)  =   "Splits(0).Columns(12).FooterStyle:id=96,.parent=45"
      _StyleDefs(84)  =   "Splits(0).Columns(12).EditorStyle:id=97,.parent=47"
      _StyleDefs(85)  =   "Splits(0).Columns(13).Style:id=102,.parent=43,.alignment=0,.valignment=2"
      _StyleDefs(86)  =   "Splits(0).Columns(13).HeadingStyle:id=99,.parent=44"
      _StyleDefs(87)  =   "Splits(0).Columns(13).FooterStyle:id=100,.parent=45"
      _StyleDefs(88)  =   "Splits(0).Columns(13).EditorStyle:id=101,.parent=47"
      _StyleDefs(89)  =   "Splits(0).Columns(14).Style:id=106,.parent=43,.alignment=0,.valignment=2"
      _StyleDefs(90)  =   "Splits(0).Columns(14).HeadingStyle:id=103,.parent=44"
      _StyleDefs(91)  =   "Splits(0).Columns(14).FooterStyle:id=104,.parent=45"
      _StyleDefs(92)  =   "Splits(0).Columns(14).EditorStyle:id=105,.parent=47"
      _StyleDefs(93)  =   "Splits(0).Columns(15).Style:id=110,.parent=43,.alignment=1"
      _StyleDefs(94)  =   "Splits(0).Columns(15).HeadingStyle:id=107,.parent=44,.alignment=3"
      _StyleDefs(95)  =   "Splits(0).Columns(15).FooterStyle:id=108,.parent=45"
      _StyleDefs(96)  =   "Splits(0).Columns(15).EditorStyle:id=109,.parent=47"
      _StyleDefs(97)  =   "Named:id=33:Normal"
      _StyleDefs(98)  =   ":id=33,.parent=0"
      _StyleDefs(99)  =   "Named:id=34:Heading"
      _StyleDefs(100) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(101) =   ":id=34,.wraptext=-1"
      _StyleDefs(102) =   "Named:id=35:Footing"
      _StyleDefs(103) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(104) =   "Named:id=36:Selected"
      _StyleDefs(105) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(106) =   "Named:id=37:Caption"
      _StyleDefs(107) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(108) =   "Named:id=38:HighlightRow"
      _StyleDefs(109) =   ":id=38,.parent=33,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(110) =   "Named:id=39:EvenRow"
      _StyleDefs(111) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(112) =   "Named:id=40:OddRow"
      _StyleDefs(113) =   ":id=40,.parent=33"
      _StyleDefs(114) =   "Named:id=41:RecordSelector"
      _StyleDefs(115) =   ":id=41,.parent=34"
      _StyleDefs(116) =   "Named:id=42:FilterBar"
      _StyleDefs(117) =   ":id=42,.parent=33"
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9210
      _Version        =   65536
      _ExtentX        =   16245
      _ExtentY        =   7646
      _StockProps     =   15
      BackColor       =   -2147483648
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
      Begin VB.TextBox txtope 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   135
         TabIndex        =   2
         ToolTipText     =   "Simplifique la consulta digitando el Simbolo (*)"
         Top             =   3915
         Visible         =   0   'False
         Width           =   2445
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "[F3] ---> Filtro Avanzado: Descripción"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   135
         TabIndex        =   5
         Top             =   3420
         Width           =   2610
      End
      Begin VB.Label Label1 
         Caption         =   "Ayuda: Hacer Click Derecho en la Descripcion que desee buscar"
         ForeColor       =   &H00800000&
         Height          =   465
         Left            =   6570
         TabIndex        =   4
         Top             =   3780
         Visible         =   0   'False
         Width           =   2490
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
         Top             =   3645
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
Attribute VB_Name = "frmhlpctas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs      As DAO.Recordset
Dim xCodigo As Integer

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
         gcodcta = Data_Ctas.Recordset.Fields("codcta")
         Unload Me
      Case 27:
         gcodcta = 0
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
    dbgrid1.EvenRowStyle.BackColor = &HFFFFFF
    dbgrid1.OddRowStyle.BackColor = &HC0FFFF
    dbgrid1.HighlightRowStyle.BackColor = vbActiveTitleBar
    dbgrid1.HighlightRowStyle.ForeColor = vbWhite
    dbgrid1.AlternatingRowStyle = True
    'Dim aa As String
    'aa = "asc"
    'Set rs = Data_Ctas.Recordset
    ''rs.Sort = "[" + DBGrid1.Columns(2).DataField + "] " + aa
    'Set Data_Ctas.Recordset = rs.OpenRecordset(rs.Type)
    'Set rs = Nothing

End Sub

Private Sub Form_Load()
    Data_Ctas.DatabaseName = wrutabancos & "\db_tabla.mdb"
    Data_Ctas.RecordSource = "Select BANCOS.BANCO,BF5PLA.CODCTA,BF5PLA.NUMCTA,BF5PLA.F5CODCTA,BF5PLA.Moneda ,bf5pla.f5saldo99,bf5pla.titular,bf5pla.ultmov From BANCOS,BF5PLA Where BANCOS.CODIGO=BF5PLA.CODBAN Order By BF5PLA.CODCTA"
    Data_Ctas.Refresh
    dbgrid1.Refresh
  
End Sub

Private Sub Form_Unload(Cancel As Integer)

   If Len(Trim(gcodcta)) = 0 Then
      gcodcta = ""
   End If

End Sub

Private Sub mnufiltro_Click()
  Set rs = Data_Ctas.Recordset
  Select Case Data_Ctas.Recordset.Fields(dbgrid1.Columns(dbgrid1.col).DataField).Type
         Case 10
              rs.Filter = "[" + dbgrid1.Columns(dbgrid1.col).DataField + "]" + " = '" + Trim("" & dbgrid1.Columns(dbgrid1.col).Text) + "'"
         Case 4
              rs.Filter = "[" + dbgrid1.Columns(dbgrid1.col).DataField + "]" + " = " + dbgrid1.Columns(dbgrid1.col).Text
         Case 8
              If IsDate(dbgrid1.Columns(dbgrid1.col).Text) Then
                 rs.Filter = "[" + dbgrid1.Columns(dbgrid1.col).DataField + "]" + "=#" + dbgrid1.Columns(dbgrid1.col).Text + "#"
              Else
                 MsgBox "Ingrese una Fecha Valida..!", 32, "Advertencia"
                 Exit Sub
              End If
  End Select
  Set Data_Ctas.Recordset = rs.OpenRecordset(rs.Type)
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
            txtope.Visible = True
            txtope.Text = ""
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
Set rs = Data_Ctas.Recordset
  rs.Sort = "[" + Data_Ctas.Recordset.Fields(dbgrid1.col).Name + "] Asc"
  Set Data_Ctas.Recordset = rs.OpenRecordset(rs.Type)
  Set rs = Nothing
End Sub

Private Sub mnuorddesc_Click()
  Set rs = Data_Ctas.Recordset
  rs.Sort = "[" + Data_Ctas.Recordset.Fields(dbgrid1.col).Name + "] Desc"
  Set Data_Ctas.Recordset = rs.OpenRecordset(rs.Type)
  Set rs = Nothing
End Sub

Private Sub MnuTodo_Click()
    Data_Ctas.DatabaseName = wrutabancos & "\db_tabla.mdb"
    Data_Ctas.RecordSource = "Select BANCOS.BANCO,BF5PLA.CODCTA,BF5PLA.NUMCTA,BF5PLA.F5CODCTA,BF5PLA.Moneda ,bf5pla.f5saldo99,bf5pla.titular,bf5pla.ultmov From BANCOS,BF5PLA Where BANCOS.CODIGO=BF5PLA.CODBAN Order By BF5PLA.CODCTA"
    Data_Ctas.Refresh
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
          Set rs = Data_Ctas.Recordset
          rs.Filter = "[" + dbgrid1.Columns(dbgrid1.col).DataField + "]" + " Like  '" + DATO + "*'"
          If rs.EOF Then txtope.SetFocus: Exit Sub Else txtope.Visible = False: lblbusca.Visible = False
          Set Data_Ctas.Recordset = rs.OpenRecordset(rs.Type)
          Set rs = Nothing
          dbgrid1.SetFocus
    
    End If
End Sub
