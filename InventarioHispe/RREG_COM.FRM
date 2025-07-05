VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmRepReg 
   Appearance      =   0  'Flat
   Caption         =   "Consulta Registro de Compras"
   ClientHeight    =   6615
   ClientLeft      =   1605
   ClientTop       =   1365
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
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6615
   ScaleWidth      =   11925
   Begin Threed.SSPanel SSPanel2 
      Height          =   600
      Left            =   3375
      TabIndex        =   6
      Top             =   5940
      Width           =   8475
      _Version        =   65536
      _ExtentX        =   14949
      _ExtentY        =   1058
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
      Begin Threed.SSPanel txtbase 
         Height          =   285
         Left            =   2250
         TabIndex        =   8
         Top             =   225
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "0.00"
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
         Alignment       =   4
      End
      Begin Threed.SSPanel txtigv 
         Height          =   285
         Left            =   3420
         TabIndex        =   9
         Top             =   225
         Width           =   825
         _Version        =   65536
         _ExtentX        =   1455
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "0.00"
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
         Alignment       =   4
      End
      Begin Threed.SSPanel txtexonerado 
         Height          =   285
         Left            =   4320
         TabIndex        =   10
         Top             =   225
         Width           =   780
         _Version        =   65536
         _ExtentX        =   1376
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "0.00"
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
         Alignment       =   4
      End
      Begin Threed.SSPanel txtotrimp 
         Height          =   285
         Left            =   5220
         TabIndex        =   11
         Top             =   225
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "0.00"
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
         Alignment       =   4
      End
      Begin Threed.SSPanel txttotal 
         Height          =   285
         Left            =   6030
         TabIndex        =   12
         Top             =   225
         Width           =   1050
         _Version        =   65536
         _ExtentX        =   1852
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "0.00"
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
         Alignment       =   4
      End
      Begin Threed.SSPanel txttotald 
         Height          =   285
         Left            =   7110
         TabIndex        =   13
         Top             =   225
         Width           =   1005
         _Version        =   65536
         _ExtentX        =   1773
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "0.00"
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
         Alignment       =   4
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TOTALES =====>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   90
         TabIndex        =   7
         Top             =   240
         Width           =   1830
      End
   End
   Begin VB.Data DATAAYUDA 
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1620
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2745
      Visible         =   0   'False
      Width           =   1836
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Bindings        =   "RREG_COM.frx":0000
      Height          =   4875
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   8599
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "No.Mov"
      Columns(0).DataField=   "F4NUMMOV"
      Columns(0).DataWidth=   255
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Ser"
      Columns(1).DataField=   "F4SERDOC"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "No.Doc."
      Columns(2).DataField=   "F4NUMDOC"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "RUC"
      Columns(3).DataField=   "f4rucprv"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Proveedor"
      Columns(4).DataField=   "f4nomprv"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Detalle"
      Columns(5).DataField=   "F4REFERE"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Base Imp."
      Columns(6).DataField=   "F4BASIMP"
      Columns(6).DataWidth=   23
      Columns(6).NumberFormat=   "Standard"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "I.G.V."
      Columns(7).DataField=   "F4IGV"
      Columns(7).NumberFormat=   "Standard"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Exonerado"
      Columns(8).DataField=   "F4MONINA"
      Columns(8).DataWidth=   23
      Columns(8).NumberFormat=   "Standard"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Otr Imp."
      Columns(9).DataField=   "F4OTRIMP"
      Columns(9).DataWidth=   23
      Columns(9).NumberFormat=   "Standard"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Total S/."
      Columns(10).DataField=   "F4TOTAL"
      Columns(10).DataWidth=   23
      Columns(10).NumberFormat=   "Standard"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "Total US$"
      Columns(11).DataField=   "F4TOTALD"
      Columns(11).NumberFormat=   "Standard"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "T.C."
      Columns(12).DataField=   "F4TIPCAM"
      Columns(12).DataWidth=   23
      Columns(12).NumberFormat=   "0.000"
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   13
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).ScrollBars=   2
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=13"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1296"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1217"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=65808"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=714"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=635"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=65808"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=1508"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1429"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=65808"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=1614"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1535"
      Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=65808"
      Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(25)=   "Column(4).Width=2990"
      Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=2910"
      Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=65808"
      Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(31)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=65808"
      Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(37)=   "Column(6).Width=1614"
      Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=1535"
      Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=66082"
      Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(43)=   "Column(7).Width=1032"
      Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=953"
      Splits(0)._ColumnProps(46)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(47)=   "Column(7)._ColStyle=66082"
      Splits(0)._ColumnProps(48)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(49)=   "Column(8).Width=1693"
      Splits(0)._ColumnProps(50)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(51)=   "Column(8)._WidthInPix=1614"
      Splits(0)._ColumnProps(52)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(53)=   "Column(8)._ColStyle=66082"
      Splits(0)._ColumnProps(54)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(55)=   "Column(9).Width=1270"
      Splits(0)._ColumnProps(56)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(57)=   "Column(9)._WidthInPix=1191"
      Splits(0)._ColumnProps(58)=   "Column(9)._EditAlways=0"
      Splits(0)._ColumnProps(59)=   "Column(9)._ColStyle=66082"
      Splits(0)._ColumnProps(60)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(61)=   "Column(10).Width=1455"
      Splits(0)._ColumnProps(62)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(63)=   "Column(10)._WidthInPix=1376"
      Splits(0)._ColumnProps(64)=   "Column(10)._EditAlways=0"
      Splits(0)._ColumnProps(65)=   "Column(10)._ColStyle=66082"
      Splits(0)._ColumnProps(66)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(67)=   "Column(11).Width=1455"
      Splits(0)._ColumnProps(68)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(69)=   "Column(11)._WidthInPix=1376"
      Splits(0)._ColumnProps(70)=   "Column(11)._EditAlways=0"
      Splits(0)._ColumnProps(71)=   "Column(11)._ColStyle=66082"
      Splits(0)._ColumnProps(72)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(73)=   "Column(12).Width=767"
      Splits(0)._ColumnProps(74)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(75)=   "Column(12)._WidthInPix=688"
      Splits(0)._ColumnProps(76)=   "Column(12)._EditAlways=0"
      Splits(0)._ColumnProps(77)=   "Column(12)._ColStyle=66082"
      Splits(0)._ColumnProps(78)=   "Column(12).Order=13"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=780,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=780,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&H8000000F&"
      _StyleDefs(11)  =   ":id=2,.fgcolor=&H80000012&,.bold=0,.fontsize=780,.italic=0,.underline=0"
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
      _StyleDefs(35)  =   "Splits(0).Columns(0).Style:id=28,.parent=43,.alignment=0,.valignment=2"
      _StyleDefs(36)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=44"
      _StyleDefs(37)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=45"
      _StyleDefs(38)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=47"
      _StyleDefs(39)  =   "Splits(0).Columns(1).Style:id=32,.parent=43,.alignment=0,.valignment=2"
      _StyleDefs(40)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=44"
      _StyleDefs(41)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=45"
      _StyleDefs(42)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=47"
      _StyleDefs(43)  =   "Splits(0).Columns(2).Style:id=58,.parent=43,.alignment=0,.valignment=2"
      _StyleDefs(44)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=44"
      _StyleDefs(45)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=45"
      _StyleDefs(46)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=47"
      _StyleDefs(47)  =   "Splits(0).Columns(3).Style:id=62,.parent=43,.alignment=0,.valignment=2"
      _StyleDefs(48)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=44"
      _StyleDefs(49)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=45"
      _StyleDefs(50)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=47"
      _StyleDefs(51)  =   "Splits(0).Columns(4).Style:id=66,.parent=43,.alignment=0,.valignment=2"
      _StyleDefs(52)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=44"
      _StyleDefs(53)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=45"
      _StyleDefs(54)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=47"
      _StyleDefs(55)  =   "Splits(0).Columns(5).Style:id=70,.parent=43,.alignment=0,.valignment=2"
      _StyleDefs(56)  =   "Splits(0).Columns(5).HeadingStyle:id=67,.parent=44"
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
      _StyleDefs(67)  =   "Splits(0).Columns(8).Style:id=82,.parent=43,.alignment=1,.valignment=1"
      _StyleDefs(68)  =   "Splits(0).Columns(8).HeadingStyle:id=79,.parent=44,.alignment=2"
      _StyleDefs(69)  =   "Splits(0).Columns(8).FooterStyle:id=80,.parent=45"
      _StyleDefs(70)  =   "Splits(0).Columns(8).EditorStyle:id=81,.parent=47"
      _StyleDefs(71)  =   "Splits(0).Columns(9).Style:id=86,.parent=43,.alignment=1,.valignment=1"
      _StyleDefs(72)  =   "Splits(0).Columns(9).HeadingStyle:id=83,.parent=44,.alignment=2"
      _StyleDefs(73)  =   "Splits(0).Columns(9).FooterStyle:id=84,.parent=45"
      _StyleDefs(74)  =   "Splits(0).Columns(9).EditorStyle:id=85,.parent=47"
      _StyleDefs(75)  =   "Splits(0).Columns(10).Style:id=90,.parent=43,.alignment=1,.valignment=1"
      _StyleDefs(76)  =   "Splits(0).Columns(10).HeadingStyle:id=87,.parent=44,.alignment=2"
      _StyleDefs(77)  =   "Splits(0).Columns(10).FooterStyle:id=88,.parent=45"
      _StyleDefs(78)  =   "Splits(0).Columns(10).EditorStyle:id=89,.parent=47"
      _StyleDefs(79)  =   "Splits(0).Columns(11).Style:id=94,.parent=43,.alignment=1,.valignment=1"
      _StyleDefs(80)  =   "Splits(0).Columns(11).HeadingStyle:id=91,.parent=44,.alignment=2"
      _StyleDefs(81)  =   "Splits(0).Columns(11).FooterStyle:id=92,.parent=45"
      _StyleDefs(82)  =   "Splits(0).Columns(11).EditorStyle:id=93,.parent=47"
      _StyleDefs(83)  =   "Splits(0).Columns(12).Style:id=98,.parent=43,.alignment=1,.valignment=1"
      _StyleDefs(84)  =   "Splits(0).Columns(12).HeadingStyle:id=95,.parent=44,.alignment=2"
      _StyleDefs(85)  =   "Splits(0).Columns(12).FooterStyle:id=96,.parent=45"
      _StyleDefs(86)  =   "Splits(0).Columns(12).EditorStyle:id=97,.parent=47"
      _StyleDefs(87)  =   "Named:id=33:Normal"
      _StyleDefs(88)  =   ":id=33,.parent=0"
      _StyleDefs(89)  =   "Named:id=34:Heading"
      _StyleDefs(90)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(91)  =   ":id=34,.wraptext=-1"
      _StyleDefs(92)  =   "Named:id=35:Footing"
      _StyleDefs(93)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(94)  =   "Named:id=36:Selected"
      _StyleDefs(95)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(96)  =   "Named:id=37:Caption"
      _StyleDefs(97)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(98)  =   "Named:id=38:HighlightRow"
      _StyleDefs(99)  =   ":id=38,.parent=33,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(100) =   "Named:id=39:EvenRow"
      _StyleDefs(101) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(102) =   "Named:id=40:OddRow"
      _StyleDefs(103) =   ":id=40,.parent=33"
      _StyleDefs(104) =   "Named:id=41:RecordSelector"
      _StyleDefs(105) =   ":id=41,.parent=34"
      _StyleDefs(106) =   "Named:id=42:FilterBar"
      _StyleDefs(107) =   ":id=42,.parent=33"
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   450
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   1260
      _Version        =   65536
      _ExtentX        =   2222
      _ExtentY        =   794
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
      Begin Threed.SSCommand BtnExitO 
         Height          =   330
         Left            =   495
         TabIndex        =   4
         ToolTipText     =   "Salir"
         Top             =   45
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   582
         _StockProps     =   78
         Picture         =   "RREG_COM.frx":0018
      End
      Begin Threed.SSCommand BtnPrint 
         Height          =   330
         Left            =   90
         TabIndex        =   5
         ToolTipText     =   "Imprimir"
         Top             =   45
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   582
         _StockProps     =   78
         Picture         =   "RREG_COM.frx":0172
      End
   End
   Begin VB.Label PnlMes 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   4590
      TabIndex        =   2
      Top             =   540
      Width           =   2655
   End
   Begin VB.Label TxtTitulo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "REGISTRO  DE  COMPRAS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   4272
      TabIndex        =   1
      Top             =   240
      Width           =   3468
   End
End
Attribute VB_Name = "FrmRepReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tbsuma      As DAO.Recordset
Dim Fnummes     As String
Dim Msql        As String
Dim wtipdoc     As String
Dim sw          As Boolean
Dim xnomdoc     As String

Sub SUMA()
    
    Set dbtempcomp = OpenDatabase(wrutatemp & "\" & "temp_com.mdb")
    Set tbsuma = dbtempcomp.OpenRecordset("select sum(f4basimp) as base, sum(f4igv) as igv, sum(f4Monina) as exonerado, sum(f4otrimp) as otrimp, sum(f4total) as total , sum(f4totald) as totald from temp_regis")
     
    txtbase.Caption = Format(tbsuma!Base, "###,###,##0.00")
    txtigv.Caption = Format(tbsuma!IGV, "###,###,##0.00")
    txtexonerado.Caption = Format(tbsuma!exonerado, "###,###,##0.00")
    TxtOtrImp.Caption = Format(tbsuma!otrimp, "###,###,##0.00")
    txttotal.Caption = Format(tbsuma!TOTAL, "###,###,##0.00")
    txttotald.Caption = Format(tbsuma!totald, "###,###,##0.00")
    
End Sub

Private Sub BtnExitO_Click()

    Unload Me

End Sub

Private Sub BtnPrint_Click()

    frmimpre.Show 1
    
End Sub

Private Sub Form_Load()

    Set dbcompras = OpenDatabase(wrutabancos & "\DB_BANCOS.mdb")
    Set TbCabRegis = dbcompras.OpenRecordset("REGISDOC")
    TbCabRegis.Index = "idmesnum"
    Set TbOfiRegis = dbcompras.OpenRecordset("REGISOFI")
    Set dbtempcomp = OpenDatabase(wrutatemp & "\temp_com.mdb")
    Set Tbtemp_regis = dbtempcomp.OpenRecordset("temp_regis")
    Set dbcomtabla = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set TbDocumento = dbcomtabla.OpenRecordset("DOCUMENTOS")
    TbDocumento.Index = "IDCODDOC"

    dbtempcomp.Execute ("Delete From temp_regis")
    
    ' CALCULO DEL CAMPO F4MES
    Fnummes = Val(FrmRepRegistro.txtmes.Text)

    Select Case Fnummes
        Case 1: F4MES = "ENERO"
        Case 2: F4MES = "FEBRERO"
        Case 3: F4MES = "MARZO"
        Case 4: F4MES = "ABRIL"
        Case 5: F4MES = "MAYO"
        Case 6: F4MES = "JUNIO"
        Case 7: F4MES = "JULIO"
        Case 8: F4MES = "AGOSTO"
        Case 9: F4MES = "SETIEMBRE"
        Case 10: F4MES = "OCTUBRE"
        Case 11: F4MES = "NOVIEMBRE"
        Case 12: F4MES = "DICIEMBRE"
    End Select
    F4MES = F4MES & " - " & wanno
    
    PnlMes.Caption = F4MES
    If FrmRepRegistro.optorden(0).Value = True Then
        Proceso
    Else
        If FrmRepRegistro.optorden(1).Value = True Then
            proceso2
        Else
            proceso3
        End If
    End If

    DataAyuda.DatabaseName = wrutatemp & "\" & "temp_com.mdb"
    DataAyuda.RecordSource = "temp_regis"
    DataAyuda.Refresh
    SUMA
    
End Sub

Private Sub Proceso()
Dim TbConsulta                          As DAO.Recordset
Dim wbasimp                             As Double
Dim wtotbasimp                          As Double
Dim wigv                                As Double
Dim wtotigv                             As Double
Dim wmonina                             As Double
Dim wtotmonina                          As Double
Dim wotrimp                             As Double
Dim wtototrimp                          As Double
Dim wmontosol                           As Double
Dim wtotalsol                           As Double
Dim wmontodol                           As Double
Dim wtotaldol                           As Double
Dim nmtodol                             As Double
Dim xpoliza                             As String
Dim i                                   As Integer

    TbDocumento.Index = "idcoddoc"
    If FrmRepRegistro.chkfechas.Value = True Then
        Msql = "Select * From  REGISOFI WHERE cvdate(f4fecha) >= CVDate('" & FrmRepRegistro.TxtDesde.Text & "')" & " and cvdate(f4fecha) <= CVDate('" & FrmRepRegistro.TxtHasta.Text & "')" & " order by F4TIPDOC,F4NUMMOV"
    Else
        Msql = "Select * From  REGISOFI WHERE F4MESMOV = '" + FrmRepRegistro.txtmes.Text + "' order by F4TIPDOC,F4NUMMOV"
    End If
    Set TbConsulta = dbcompras.OpenRecordset(Msql)
    wtotbasimp = 0#: wtotigv = 0#: wtotmonina = 0#: wtototrimp = 0#: wtotalsol = 0#: wtotaldol = 0#
    Do While Not TbConsulta.EOF
        wtipdoc = "" & TbConsulta.Fields("F4TIPDOC")
        
        wbasimp = 0#: wigv = 0#: wmonina = 0#: wotrimp = 0#: wmontosol = 0#: wmontodol = 0#
        '----------------------------------------------------------
        sw = False
        For i = 0 To 35
            If FrmRepRegistro.ChkTipDoc(i).Visible = True And FrmRepRegistro.ChkTipDoc(i).Value = True Then
                If Right(FrmRepRegistro.ChkTipDoc(i).Caption, 2) = wtipdoc Then
                    sw = True
                    xnomdoc = FrmRepRegistro.ChkTipDoc(i).Caption & ""
                End If
            End If
        Next
        '----------------------------------------------------------
        If sw = True Then
            Do While "" & TbConsulta.Fields("F4TIPDOC") = wtipdoc
                '-------------------------------------------------
                Tbtemp_regis.AddNew
                TbDocumento.Seek "=", TbConsulta.Fields("F4TIPDOC")
                If Not TbDocumento.NoMatch = True Then
                    Tbtemp_regis.Fields("F4DESDOC") = TbDocumento.Fields("f2desdoc")
                    xpoliza = TbDocumento.Fields("f2abrev") & ""
                Else
                    Tbtemp_regis.Fields("F4DESDOC") = " "
                    xpoliza = ""
                End If
                Tbtemp_regis.Fields("F4MESMOV") = TbConsulta.Fields("F4MESMOV")
                Tbtemp_regis.Fields("F4NUMMOV") = TbConsulta.Fields("F4NUMMOV")
                Tbtemp_regis.Fields("F4CODPRV") = TbConsulta.Fields("F4CODPRV")
                Tbtemp_regis.Fields("F4NOMPRV") = TbConsulta.Fields("F4NOMPRV") & ""
                Tbtemp_regis.Fields("F4TIPDOC") = TbConsulta.Fields("F4TIPDOC")
                Tbtemp_regis.Fields("F4SERDOC") = TbConsulta.Fields("F4SERDOC")
                Tbtemp_regis.Fields("F4NUMDOC") = TbConsulta.Fields("F4NUMDOC")
                Tbtemp_regis.Fields("F4FECHA") = TbConsulta.Fields("F4FECHA")
                Tbtemp_regis.Fields("F4MONEDA") = TbConsulta.Fields("F4MONEDA")
                Tbtemp_regis.Fields("F4RUCPRV") = TbConsulta.Fields("F4RUCPRV")
                Tbtemp_regis.Fields("F4REFERE") = TbConsulta.Fields("F4REFERE") & ""
                Select Case wf1formato
                    Case "2"
                        Tbtemp_regis.Fields("F4POLIZA") = Trim(TbConsulta.Fields("F4POLIZA") & "")
                        If TbConsulta.Fields("f4codigv") = "001" Then
                            Tbtemp_regis.Fields("F4BASIMP") = Format(TbConsulta.Fields("F4BASIMP"), "###,##0.00")
                        Else
                            Tbtemp_regis.Fields("F4BASIMP2") = Format(TbConsulta.Fields("F4BASIMP"), "###,##0.00")
                        End If
                        If Len(Trim(TbConsulta.Fields("F4POLIZA"))) > 0 Then
                            Tbtemp_regis.Fields("F4IMPORTA") = Format(TbConsulta.Fields("F4BASIMP"), "###,##0.00")
                            Tbtemp_regis.Fields("F4BASIMP") = 0#
                            Tbtemp_regis.Fields("F4BASIMP2") = 0#
                        End If
                    Case "7":
                        Tbtemp_regis.Fields("F4POLIZA") = Trim(TbConsulta.Fields("F4POLIZA") & "")
                        If Len(Trim(TbConsulta.Fields("F4POLIZA") & "")) > 0 Then
                            Tbtemp_regis.Fields("F4BASIMP") = 0#
                            Tbtemp_regis.Fields("F4FACEXT") = Format(TbConsulta.Fields("F4BASIMP"), "###,##0.00")
                        End If
                        If UCase(xpoliza) = "POL" Then '----- xpoliza ="POL"
                            Tbtemp_regis.Fields("F4IMPORTA") = Format(TbConsulta.Fields("F4BASIMP"), "###,##0.00")
                            Tbtemp_regis.Fields("F4BASIMP") = 0#
                            Tbtemp_regis.Fields("F4FACEXT") = 0#
                        End If
                        If UCase(xpoliza) = "HON" Then '----- xpoliza ="HON"
                            Tbtemp_regis.Fields("F4BASIMP") = 0#
                            Tbtemp_regis.Fields("F4FACEXT") = 0#
                            Tbtemp_regis.Fields("F4IMPORTA") = 0#
                            Tbtemp_regis.Fields("F4HONORARIO") = Format(TbConsulta.Fields("F4BASIMP"), "###,##0.00")
                            Tbtemp_regis.Fields("F4_4TA_IES") = TbConsulta.Fields("F4MONTORET") + TbConsulta.Fields("F4FONAVI")
                        End If
                    Case "9":
                        
                        Select Case "" & TbConsulta.Fields("f4codigv")
                            Case "001":
                                Tbtemp_regis.Fields("F4BASIMP") = Format(TbConsulta.Fields("F4BASIMP"), "0.00")
                            Case "004":
                                Tbtemp_regis.Fields("F4IMPORTA") = Format(TbConsulta.Fields("F4BASIMP") + TbConsulta.Fields("F4MONINA"), "0.00")
                            Case Else
                                Tbtemp_regis.Fields("F4BASIMP2") = Format(TbConsulta.Fields("F4BASIMP"), "0.00")
                        End Select
                        
                    Case Else
                        Tbtemp_regis.Fields("F4BASIMP") = Format(TbConsulta.Fields("F4BASIMP"), "###,##0.00")
                End Select
                
                If wf1formato = "4" Or wf1formato = "5" Then
                    If TbConsulta.Fields("f4codigv") & "" = "001" Then
                        Tbtemp_regis.Fields("F4IGV") = Format(TbConsulta.Fields("F4IGV"), "###,##0.00")
                    End If
                    If TbConsulta.Fields("f4codigv") & "" = "002" Then
                        Tbtemp_regis.Fields("F4IGV2") = Format(TbConsulta.Fields("F4IGV"), "###,##0.00")
                    End If
                    If TbConsulta.Fields("f4codigv") & "" = "003" Then
                        Tbtemp_regis.Fields("F4IGV3") = Format(TbConsulta.Fields("F4IGV"), "###,##0.00")
                    End If
                Else
                    Tbtemp_regis.Fields("F4IGV") = Format(TbConsulta.Fields("F4IGV"), "###,##0.00")
                End If
                If wf1formato = "9" Then
                    Tbtemp_regis.Fields("F4TOTAL") = Format(TbConsulta.Fields("F4BASIMP") + TbConsulta.Fields("F4MONINA"), "###,##0.00")
                Else
                    Tbtemp_regis.Fields("F4MONINA") = Format(TbConsulta.Fields("F4MONINA"), "###,##0.00")
                    Tbtemp_regis.Fields("F4OTRIMP") = Format(TbConsulta.Fields("F4OTRIMP"), "###,##0.00")
                    Tbtemp_regis.Fields("F4TOTAL") = Format(TbConsulta.Fields("F4TOTAL"), "###,##0.00")
                End If
                Tbtemp_regis.Fields("F4TIPCAM") = Format(TbConsulta.Fields("F4TIPCAM"), "###,##0.000")
                Tbtemp_regis.Fields("F4EMPRESA") = wnomcia
                Tbtemp_regis.Fields("F4MES") = F4MES
                nmtodol = 0#
                TbCabRegis.Seek "=", TbConsulta.Fields("F4MESMOV"), TbConsulta.Fields("F4NUMMOV")
                If Not TbCabRegis.NoMatch Then
                    If wf1formato = "4" Or wf1formato = "5" Then
                        Tbtemp_regis.Fields("f4moninad") = Val(TbConsulta.Fields("f4redsuma") & "") - Val(TbConsulta.Fields("f4redresta") & "")
                        Tbtemp_regis.Fields("f4moneda") = TbCabRegis.Fields("f4moneda")
                        If TbCabRegis.Fields("f4moneda") = "D" Then
                            Tbtemp_regis.Fields("f4totald") = TbCabRegis.Fields("f4total")
                        Else
                            If TbConsulta.Fields("f4tipcam") > 0 Then
                               Tbtemp_regis.Fields("f4totald") = Format(TbConsulta.Fields("f4total") / TbConsulta.Fields("f4tipcam"), "#0.00")
                            End If
                        End If
                    Else
                        If TbCabRegis.Fields("f4moneda") = "D" Then
                            nmtodol = TbCabRegis.Fields("f4total")
                            Tbtemp_regis.Fields("f4totald") = TbCabRegis.Fields("f4total")
                            Tbtemp_regis.Fields("f4basimpd") = TbCabRegis.Fields("f4basimp")
                            Tbtemp_regis.Fields("f4igvd") = TbCabRegis.Fields("f4igv")
                            Tbtemp_regis.Fields("f4moninad") = TbCabRegis.Fields("f4monina")
                            Tbtemp_regis.Fields("f4otrimpd") = TbCabRegis.Fields("f4otrimp")
                        Else
                            If TbConsulta.Fields("f4tipcam") > 0 Then
                                nmtodol = Format(TbConsulta.Fields("f4total") / TbConsulta.Fields("f4tipcam"), "#0.00")
                                Tbtemp_regis.Fields("f4totald") = Format(TbConsulta.Fields("f4total") / TbConsulta.Fields("f4tipcam"), "#0.00")
                                Tbtemp_regis.Fields("f4basimpd") = Format(TbConsulta.Fields("f4basimp") / TbConsulta.Fields("f4tipcam"), "#0.00")
                                Tbtemp_regis.Fields("f4igvd") = Format(TbConsulta.Fields("f4igv") / TbConsulta.Fields("f4tipcam"), "#0.00")
                                Tbtemp_regis.Fields("f4moninad") = Format(TbConsulta.Fields("f4monina") / TbConsulta.Fields("f4tipcam"), "#0.00")
                                Tbtemp_regis.Fields("f4otrimpd") = Format(TbConsulta.Fields("f4otrimp") / TbConsulta.Fields("f4tipcam"), "#0.00")
                            End If
                        End If
                    End If
                End If
                Tbtemp_regis.Update
                '-------------------------------------------------
                '-------------------------------------------------
                wbasimp = wbasimp + Val(Format(TbConsulta.Fields("F4BASIMP"), "0.00"))
                wigv = wigv + Val(Format(TbConsulta.Fields("F4IGV"), "0.00"))
                wmonina = wmonina + Val(Format(TbConsulta.Fields("F4MONINA"), "0.00")) + Val(TbConsulta.Fields("f4redsuma") & "") - Val(TbConsulta.Fields("f4redresta") & "")
                wotrimp = wotrimp + Val(Format(TbConsulta.Fields("F4OTRIMP"), "0.00"))
                wmontosol = wmontosol + Val(Format(TbConsulta.Fields("F4TOTAL"), "0.00"))
                wmontodol = wmontodol + nmtodol
                wtotbasimp = wtotbasimp + Val(Format(TbConsulta.Fields("F4BASIMP"), "0.00"))
                wtotigv = wtotigv + Val(Format(TbConsulta.Fields("F4IGV"), "0.00"))
                wtotmonina = wtotmonina + Val(Format(TbConsulta.Fields("F4MONINA"), "0.00")) + Val(TbConsulta.Fields("f4redsuma") & "") - Val(TbConsulta.Fields("f4redresta") & "")
                wtototrimp = wtototrimp + Val(Format(TbConsulta.Fields("F4OTRIMP"), "0.00"))
                wtotalsol = wtotalsol + Val(Format(TbConsulta.Fields("F4TOTAL"), "0.00"))
                wtotaldol = wtotaldol + nmtodol
                TbConsulta.MoveNext
                If TbConsulta.EOF Then Exit Do
            Loop
        Else
            Do While "" & TbConsulta.Fields("F4TIPDOC") = wtipdoc
                TbConsulta.MoveNext
                If TbConsulta.EOF Then Exit Do
            Loop
        End If
    Loop

    If TbConsulta.RecordCount = 0 Then
        Tbtemp_regis.AddNew
        Tbtemp_regis.Fields("F4EMPRESA") = wnomcia
        Tbtemp_regis.Fields("F4MES") = F4MES
        Tbtemp_regis.Update
    End If

End Sub

Private Sub proceso2()
Dim TbConsulta                          As DAO.Recordset
Dim wbasimp As Double, wtotbasimp       As Double
Dim wigv As Double, wtotigv             As Double
Dim wmonina As Double, wtotmonina       As Double
Dim wotrimp As Double, wtototrimp       As Double
Dim wmontosol As Double, wtotalsol      As Double
Dim wmontodol As Double, wtotaldol      As Double
Dim nmtodol                             As Double
Dim i                                   As Integer

    TbDocumento.Index = "idcoddoc"
    If FrmRepRegistro.chkfechas.Value = True Then
        Msql = "Select * From  REGISOFI WHERE cvdate(f4fecha) >= CVDate('" & FrmRepRegistro.TxtDesde.Text & "')" & " and cvdate(f4fecha) <= CVDate('" & FrmRepRegistro.TxtHasta.Text & "')" & " order by F4NUMMOV"
    Else
        Msql = "Select * From  REGISOFI WHERE F4MESMOV = '" + FrmRepRegistro.txtmes.Text + "' order by F4NUMMOV"
    End If
    Set TbConsulta = dbcompras.OpenRecordset(Msql)
 '   Grid1.Row = 1
    wtotbasimp = 0#: wtotigv = 0#: wtotmonina = 0#: wtototrimp = 0#: wtotalsol = 0#: wtotaldol = 0#
    Do While Not TbConsulta.EOF
        wtipdoc = "" & TbConsulta.Fields("F4TIPDOC")
        wbasimp = 0#: wigv = 0#: wmonina = 0#: wotrimp = 0#: wmontosol = 0#: wmontodol = 0#
        '----------------------------------------------------------
        sw = False
        For i = 0 To 35
            If FrmRepRegistro.ChkTipDoc(i).Visible = True And FrmRepRegistro.ChkTipDoc(i).Value = True Then
                If Right(FrmRepRegistro.ChkTipDoc(i).Caption, 2) = wtipdoc Then
                    sw = True
                End If
            End If
        Next
        '----------------------------------------------------------
        If sw = True Then
            Tbtemp_regis.AddNew
            Tbtemp_regis.Fields("F4MESMOV") = TbConsulta.Fields("F4MESMOV")
            Tbtemp_regis.Fields("F4NUMMOV") = TbConsulta.Fields("F4NUMMOV")
            Tbtemp_regis.Fields("F4CODPRV") = TbConsulta.Fields("F4CODPRV")
            Tbtemp_regis.Fields("F4NOMPRV") = TbConsulta.Fields("F4NOMPRV")
            Tbtemp_regis.Fields("F4TIPDOC") = TbConsulta.Fields("F4TIPDOC")
            Tbtemp_regis.Fields("F4SERDOC") = TbConsulta.Fields("F4SERDOC")
            Tbtemp_regis.Fields("F4NUMDOC") = TbConsulta.Fields("F4NUMDOC")
            Tbtemp_regis.Fields("F4FECHA") = TbConsulta.Fields("F4FECHA")
            Tbtemp_regis.Fields("F4MONEDA") = TbConsulta.Fields("F4MONEDA")
            Tbtemp_regis.Fields("F4RUCPRV") = TbConsulta.Fields("F4RUCPRV")
            Tbtemp_regis.Fields("F4REFERE") = TbConsulta.Fields("F4REFERE")
            Tbtemp_regis.Fields("F4BASIMP") = Format(TbConsulta.Fields("F4BASIMP"), "###,##0.00")
            Tbtemp_regis.Fields("F4IGV") = Format(TbConsulta.Fields("F4IGV"), "###,##0.00")
            Tbtemp_regis.Fields("F4MONINA") = Format(TbConsulta.Fields("F4MONINA"), "###,##0.00")
            Tbtemp_regis.Fields("F4OTRIMP") = Format(TbConsulta.Fields("F4OTRIMP"), "###,##0.00")
            Tbtemp_regis.Fields("F4TOTAL") = Format(TbConsulta.Fields("F4TOTAL"), "###,##0.00")
            Tbtemp_regis.Fields("F4TIPCAM") = Format(TbConsulta.Fields("F4TIPCAM"), "###,##0.000")
            Tbtemp_regis.Fields("F4EMPRESA") = wnomcia
            Tbtemp_regis.Fields("F4MES") = F4MES
            TbDocumento.Seek "=", TbConsulta.Fields("F4TIPDOC")
            If Not TbDocumento.NoMatch = True Then
                Tbtemp_regis.Fields("F4DESDOC") = TbDocumento.Fields("f2desdoc")
            Else
                Tbtemp_regis.Fields("F4DESDOC") = " "
            End If
            nmtodol = 0#
            TbCabRegis.Seek "=", TbConsulta.Fields("F4MESMOV"), TbConsulta.Fields("F4NUMMOV")
            If Not TbCabRegis.NoMatch Then
                If TbConsulta.Fields("f4moneda") = "D" Then
                    nmtodol = TbConsulta.Fields("f4total")
                    Tbtemp_regis.Fields("f4totald") = TbConsulta.Fields("f4total")
                    Tbtemp_regis.Fields("f4basimpd") = TbConsulta.Fields("f4basimp")
                    Tbtemp_regis.Fields("f4igvd") = TbConsulta.Fields("f4igv")
                    Tbtemp_regis.Fields("f4moninad") = TbConsulta.Fields("f4monina")
                    Tbtemp_regis.Fields("f4otrimpd") = TbConsulta.Fields("f4otrimp")
                Else
                    If TbConsulta.Fields("f4tipcam") = 0 Then
                        nmtodol = Format(TbConsulta.Fields("f4total") / 1, "#0.00")
                        Tbtemp_regis.Fields("f4totald") = Format(TbConsulta.Fields("f4total") / 1, "#0.00")
                        Tbtemp_regis.Fields("f4basimpd") = Format(TbConsulta.Fields("f4basimp") / 1, "#0.00")
                        Tbtemp_regis.Fields("f4igvd") = Format(TbConsulta.Fields("f4igv") / 1, "#0.00")
                        Tbtemp_regis.Fields("f4moninad") = Format(TbConsulta.Fields("f4monina") / 1, "#0.00")
                        Tbtemp_regis.Fields("f4otrimpd") = Format(TbConsulta.Fields("f4otrimp") / 1, "#0.00")
                    Else
                        nmtodol = Format(TbConsulta.Fields("f4total") / TbConsulta.Fields("f4tipcam"), "#0.00")
                        Tbtemp_regis.Fields("f4totald") = Format(TbConsulta.Fields("f4total") / TbConsulta.Fields("f4tipcam"), "#0.00")
                        Tbtemp_regis.Fields("f4basimpd") = Format(TbConsulta.Fields("f4basimp") / TbConsulta.Fields("f4tipcam"), "#0.00")
                        Tbtemp_regis.Fields("f4igvd") = Format(TbConsulta.Fields("f4igv") / TbConsulta.Fields("f4tipcam"), "#0.00")
                        Tbtemp_regis.Fields("f4moninad") = Format(TbConsulta.Fields("f4monina") / TbConsulta.Fields("f4tipcam"), "#0.00")
                        Tbtemp_regis.Fields("f4otrimpd") = Format(TbConsulta.Fields("f4otrimp") / TbConsulta.Fields("f4tipcam"), "#0.00")
                    End If
                End If
            End If
            Tbtemp_regis.Update
            wbasimp = wbasimp + Val(Format(TbConsulta.Fields("F4BASIMP"), "0.00"))
            wigv = wigv + Val(Format(TbConsulta.Fields("F4IGV"), "0.00"))
            wmonina = wmonina + Val(Format(TbConsulta.Fields("F4MONINA"), "0.00"))
            wotrimp = wotrimp + Val(Format(TbConsulta.Fields("F4OTRIMP"), "0.00"))
            wmontosol = wmontosol + Val(Format(TbConsulta.Fields("F4TOTAL"), "0.00"))
            wmontodol = wmontodol + nmtodol
            wtotbasimp = wtotbasimp + Val(Format(TbConsulta.Fields("F4BASIMP"), "0.00"))
            wtotigv = wtotigv + Val(Format(TbConsulta.Fields("F4IGV"), "0.00"))
            wtotmonina = wtotmonina + Val(Format(TbConsulta.Fields("F4MONINA"), "0.00"))
            wtototrimp = wtototrimp + Val(Format(TbConsulta.Fields("F4OTRIMP"), "0.00"))
            wtotalsol = wtotalsol + Val(Format(TbConsulta.Fields("F4TOTAL"), "0.00"))
            wtotaldol = wtotaldol + nmtodol
        End If
        TbConsulta.MoveNext
        If TbConsulta.EOF Then Exit Do
    Loop

    If TbConsulta.RecordCount = 0 Then
        Tbtemp_regis.AddNew
        Tbtemp_regis.Fields("F4EMPRESA") = wnomcia
        Tbtemp_regis.Fields("F4MES") = F4MES
        Tbtemp_regis.Update
    End If
    
End Sub

Private Sub proceso3()
Dim TbConsulta                          As DAO.Recordset
Dim wbasimp As Double, wtotbasimp       As Double
Dim wigv As Double, wtotigv             As Double
Dim wmonina As Double, wtotmonina       As Double
Dim wotrimp As Double, wtototrimp       As Double
Dim wmontosol As Double, wtotalsol      As Double
Dim wmontodol As Double, wtotaldol      As Double
Dim nmtodol                             As Double
Dim i                                   As Integer

    TbDocumento.Index = "idcoddoc"
    If FrmRepRegistro.chkfechas.Value = True Then
        Msql = "Select * From  REGISOFI WHERE cvdate(f4fecha) >= CVDate('" & FrmRepRegistro.TxtDesde.Text & "')" & " and cvdate(f4fecha) <= CVDate('" & FrmRepRegistro.TxtHasta.Text & "')" & " order by f4fecha"
    Else
        Msql = "Select * From  REGISOFI WHERE F4MESMOV = '" + FrmRepRegistro.txtmes.Text + "' order by f4fecha"
    End If
    Set TbConsulta = dbcompras.OpenRecordset(Msql)
    
    wtotbasimp = 0#: wtotigv = 0#: wtotmonina = 0#: wtototrimp = 0#: wtotalsol = 0#: wtotaldol = 0#
    Do While Not TbConsulta.EOF
        wtipdoc = "" & TbConsulta.Fields("F4TIPDOC")
        wbasimp = 0#: wigv = 0#: wmonina = 0#: wotrimp = 0#: wmontosol = 0#: wmontodol = 0#
        '----------------------------------------------------------
        sw = False
        For i = 0 To 35
            If FrmRepRegistro.ChkTipDoc(i).Visible = True And FrmRepRegistro.ChkTipDoc(i).Value = True Then
                If Right(FrmRepRegistro.ChkTipDoc(i).Caption, 2) = wtipdoc Then
                    sw = True
                End If
            End If
        Next
        '----------------------------------------------------------
        If sw = True Then
            Tbtemp_regis.AddNew
            Tbtemp_regis.Fields("F4MESMOV") = TbConsulta.Fields("F4MESMOV")
            Tbtemp_regis.Fields("F4NUMMOV") = TbConsulta.Fields("F4NUMMOV")
            Tbtemp_regis.Fields("F4CODPRV") = TbConsulta.Fields("F4CODPRV")
            Tbtemp_regis.Fields("F4NOMPRV") = TbConsulta.Fields("F4NOMPRV")
            Tbtemp_regis.Fields("F4TIPDOC") = TbConsulta.Fields("F4TIPDOC")
            Tbtemp_regis.Fields("F4SERDOC") = TbConsulta.Fields("F4SERDOC")
            Tbtemp_regis.Fields("F4NUMDOC") = TbConsulta.Fields("F4NUMDOC")
            Tbtemp_regis.Fields("F4FECHA") = TbConsulta.Fields("F4FECHA")
            Tbtemp_regis.Fields("F4MONEDA") = TbConsulta.Fields("F4MONEDA")
            Tbtemp_regis.Fields("F4RUCPRV") = TbConsulta.Fields("F4RUCPRV")
            Tbtemp_regis.Fields("F4REFERE") = TbConsulta.Fields("F4REFERE")
            Tbtemp_regis.Fields("F4BASIMP") = Format(TbConsulta.Fields("F4BASIMP"), "###,##0.00")
            Tbtemp_regis.Fields("F4IGV") = Format(TbConsulta.Fields("F4IGV"), "###,##0.00")
            Tbtemp_regis.Fields("F4MONINA") = Format(TbConsulta.Fields("F4MONINA"), "###,##0.00")
            Tbtemp_regis.Fields("F4OTRIMP") = Format(TbConsulta.Fields("F4OTRIMP"), "###,##0.00")
            Tbtemp_regis.Fields("F4TOTAL") = Format(TbConsulta.Fields("F4TOTAL"), "###,##0.00")
            Tbtemp_regis.Fields("F4TIPCAM") = Format(TbConsulta.Fields("F4TIPCAM"), "###,##0.000")
            Tbtemp_regis.Fields("F4EMPRESA") = wnomcia
            Tbtemp_regis.Fields("F4MES") = F4MES
            TbDocumento.Seek "=", TbConsulta.Fields("F4TIPDOC")
            If Not TbDocumento.NoMatch = True Then
                Tbtemp_regis.Fields("F4DESDOC") = TbDocumento.Fields("f2desdoc")
            Else
                Tbtemp_regis.Fields("F4DESDOC") = " "
            End If
            nmtodol = 0#
            TbCabRegis.Seek "=", TbConsulta.Fields("F4MESMOV"), TbConsulta.Fields("F4NUMMOV")
            If Not TbCabRegis.NoMatch Then
                If TbConsulta.Fields("f4moneda") = "D" Then
                    nmtodol = TbConsulta.Fields("f4total")
                    Tbtemp_regis.Fields("f4totald") = TbConsulta.Fields("f4total")
                    Tbtemp_regis.Fields("f4basimpd") = TbConsulta.Fields("f4basimp")
                    Tbtemp_regis.Fields("f4igvd") = TbConsulta.Fields("f4igv")
                    Tbtemp_regis.Fields("f4moninad") = TbConsulta.Fields("f4monina")
                    Tbtemp_regis.Fields("f4otrimpd") = TbConsulta.Fields("f4otrimp")
                Else
                    If TbConsulta.Fields("f4tipcam") > 0 Then
                        nmtodol = Format(TbConsulta.Fields("f4total") / TbConsulta.Fields("f4tipcam"), "#0.00")
                        Tbtemp_regis.Fields("f4totald") = Format(TbConsulta.Fields("f4total") / TbConsulta.Fields("f4tipcam"), "#0.00")
                        Tbtemp_regis.Fields("f4basimpd") = Format(TbConsulta.Fields("f4basimp") / TbConsulta.Fields("f4tipcam"), "#0.00")
                        Tbtemp_regis.Fields("f4igvd") = Format(TbConsulta.Fields("f4igv") / TbConsulta.Fields("f4tipcam"), "#0.00")
                        Tbtemp_regis.Fields("f4moninad") = Format(TbConsulta.Fields("f4monina") / TbConsulta.Fields("f4tipcam"), "#0.00")
                        Tbtemp_regis.Fields("f4otrimpd") = Format(TbConsulta.Fields("f4otrimp") / TbConsulta.Fields("f4tipcam"), "#0.00")
                    End If
                End If
            End If
            Tbtemp_regis.Update
            wbasimp = wbasimp + Val(Format(TbConsulta.Fields("F4BASIMP"), "0.00"))
            wigv = wigv + Val(Format(TbConsulta.Fields("F4IGV"), "0.00"))
            wmonina = wmonina + Val(Format(TbConsulta.Fields("F4MONINA"), "0.00"))
            wotrimp = wotrimp + Val(Format(TbConsulta.Fields("F4OTRIMP"), "0.00"))
            wmontosol = wmontosol + Val(Format(TbConsulta.Fields("F4TOTAL"), "0.00"))
            wmontodol = wmontodol + nmtodol
            wtotbasimp = wtotbasimp + Val(Format(TbConsulta.Fields("F4BASIMP"), "0.00"))
            wtotigv = wtotigv + Val(Format(TbConsulta.Fields("F4IGV"), "0.00"))
            wtotmonina = wtotmonina + Val(Format(TbConsulta.Fields("F4MONINA"), "0.00"))
            wtototrimp = wtototrimp + Val(Format(TbConsulta.Fields("F4OTRIMP"), "0.00"))
            wtotalsol = wtotalsol + Val(Format(TbConsulta.Fields("F4TOTAL"), "0.00"))
            wtotaldol = wtotaldol + nmtodol
        End If
        TbConsulta.MoveNext
        If TbConsulta.EOF Then Exit Do
    Loop

    If TbConsulta.RecordCount = 0 Then
        Tbtemp_regis.AddNew
        Tbtemp_regis.Fields("F4EMPRESA") = wnomcia
        Tbtemp_regis.Fields("F4MES") = F4MES
        Tbtemp_regis.Update
    End If

End Sub
