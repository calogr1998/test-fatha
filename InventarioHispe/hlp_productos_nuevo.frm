VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form hlp_productos_nuevo 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista de Productos"
   ClientHeight    =   4785
   ClientLeft      =   540
   ClientTop       =   2175
   ClientWidth     =   11895
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
   ScaleHeight     =   4785
   ScaleWidth      =   11895
   Begin Threed.SSPanel SSPanel1 
      Height          =   4650
      Left            =   45
      TabIndex        =   1
      Top             =   90
      Width           =   11775
      _Version        =   65536
      _ExtentX        =   20770
      _ExtentY        =   8202
      _StockProps     =   15
      Caption         =   "SSPanel1"
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
      Begin VB.ComboBox cmbfiltro 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   4140
         Visible         =   0   'False
         Width           =   1635
      End
      Begin Threed.SSCommand cmdnuevo 
         Height          =   420
         Left            =   10380
         TabIndex        =   7
         Top             =   3450
         Width           =   1185
         _Version        =   65536
         _ExtentX        =   2090
         _ExtentY        =   741
         _StockProps     =   78
         Caption         =   "&Nuevo"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin MSAdodcLib.Adodc DataAyuda 
         Height          =   330
         Left            =   2070
         Top             =   1215
         Visible         =   0   'False
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
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
         Height          =   330
         Left            =   1800
         TabIndex        =   0
         ToolTipText     =   "Simplifique la consulta digitando el Simbolo (*)"
         Top             =   4140
         Width           =   2445
      End
      Begin TrueOleDBGrid70.TDBGrid DataGrid 
         Bindings        =   "hlp_productos_nuevo.frx":0000
         Height          =   3210
         Left            =   90
         TabIndex        =   2
         Top             =   90
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   5662
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Codigo"
         Columns(0).DataField=   "f5codpro"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Cod. Fab."
         Columns(1).DataField=   "F5CODFAB"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Marca"
         Columns(2).DataField=   "F2DESMAR"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Descripción"
         Columns(3).DataField=   "F5NOMPRO"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Precio. Vta"
         Columns(4).DataField=   "F5valvta"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Unidad"
         Columns(5).DataField=   "F7SIGMED"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "F5STOCKACT"
         Columns(6).DataField=   "F5STOCKACT"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "F5PRECOS"
         Columns(7).DataField=   "F5FOB"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "F7CODMED"
         Columns(8).DataField=   "F7CODMED"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Partida Arancelaria"
         Columns(9).DataField=   "f5partara"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "F5MARCA"
         Columns(10).DataField=   "F5MARCA"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   11
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   2
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   13154464
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=11"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1879"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1799"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=65808"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=1588"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1508"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).AllowSizing=0"
         Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=65808"
         Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(15)=   "Column(2).Width=4445"
         Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=4366"
         Splits(0)._ColumnProps(18)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(19)=   "Column(2)._ColStyle=65792"
         Splits(0)._ColumnProps(20)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(21)=   "Column(3).Width=7303"
         Splits(0)._ColumnProps(22)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(3)._WidthInPix=7223"
         Splits(0)._ColumnProps(24)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(25)=   "Column(3).AllowSizing=0"
         Splits(0)._ColumnProps(26)=   "Column(3)._ColStyle=65808"
         Splits(0)._ColumnProps(27)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(28)=   "Column(4).Width=1958"
         Splits(0)._ColumnProps(29)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(30)=   "Column(4)._WidthInPix=1879"
         Splits(0)._ColumnProps(31)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(32)=   "Column(4).AllowSizing=0"
         Splits(0)._ColumnProps(33)=   "Column(4)._ColStyle=73730"
         Splits(0)._ColumnProps(34)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(35)=   "Column(5).Width=1217"
         Splits(0)._ColumnProps(36)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(37)=   "Column(5)._WidthInPix=1138"
         Splits(0)._ColumnProps(38)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(39)=   "Column(5).AllowSizing=0"
         Splits(0)._ColumnProps(40)=   "Column(5)._ColStyle=66082"
         Splits(0)._ColumnProps(41)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(42)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(43)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(44)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(45)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(46)=   "Column(6)._ColStyle=65792"
         Splits(0)._ColumnProps(47)=   "Column(6).Visible=0"
         Splits(0)._ColumnProps(48)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(49)=   "Column(7).Width=2725"
         Splits(0)._ColumnProps(50)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(51)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(52)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(53)=   "Column(7)._ColStyle=65792"
         Splits(0)._ColumnProps(54)=   "Column(7).Visible=0"
         Splits(0)._ColumnProps(55)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(56)=   "Column(8).Width=2725"
         Splits(0)._ColumnProps(57)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(58)=   "Column(8)._WidthInPix=2646"
         Splits(0)._ColumnProps(59)=   "Column(8)._EditAlways=0"
         Splits(0)._ColumnProps(60)=   "Column(8)._ColStyle=65792"
         Splits(0)._ColumnProps(61)=   "Column(8).Visible=0"
         Splits(0)._ColumnProps(62)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(63)=   "Column(9).Width=4260"
         Splits(0)._ColumnProps(64)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(65)=   "Column(9)._WidthInPix=4180"
         Splits(0)._ColumnProps(66)=   "Column(9)._EditAlways=0"
         Splits(0)._ColumnProps(67)=   "Column(9)._ColStyle=65792"
         Splits(0)._ColumnProps(68)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(69)=   "Column(10).Width=2725"
         Splits(0)._ColumnProps(70)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(71)=   "Column(10)._WidthInPix=2646"
         Splits(0)._ColumnProps(72)=   "Column(10)._EditAlways=0"
         Splits(0)._ColumnProps(73)=   "Column(10)._ColStyle=65792"
         Splits(0)._ColumnProps(74)=   "Column(10).Visible=0"
         Splits(0)._ColumnProps(75)=   "Column(10).Order=11"
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
         DeadAreaBackColor=   -2147483633
         RowDividerColor =   13154464
         RowSubDividerColor=   13154464
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=0,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=-1,.fontsize=825,.italic=0"
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
         _StyleDefs(35)  =   "Splits(0).Columns(0).Style:id=28,.parent=43,.alignment=0,.valignment=2,.locked=0"
         _StyleDefs(36)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=44"
         _StyleDefs(37)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=45"
         _StyleDefs(38)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=47"
         _StyleDefs(39)  =   "Splits(0).Columns(1).Style:id=32,.parent=43,.alignment=0,.valignment=2,.locked=0"
         _StyleDefs(40)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=44"
         _StyleDefs(41)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=45"
         _StyleDefs(42)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=47"
         _StyleDefs(43)  =   "Splits(0).Columns(2).Style:id=24,.parent=43"
         _StyleDefs(44)  =   "Splits(0).Columns(2).HeadingStyle:id=21,.parent=44"
         _StyleDefs(45)  =   "Splits(0).Columns(2).FooterStyle:id=22,.parent=45"
         _StyleDefs(46)  =   "Splits(0).Columns(2).EditorStyle:id=23,.parent=47"
         _StyleDefs(47)  =   "Splits(0).Columns(3).Style:id=58,.parent=43,.alignment=0,.valignment=2,.locked=0"
         _StyleDefs(48)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=44"
         _StyleDefs(49)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=45"
         _StyleDefs(50)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=47"
         _StyleDefs(51)  =   "Splits(0).Columns(4).Style:id=62,.parent=43,.alignment=1,.locked=-1"
         _StyleDefs(52)  =   "Splits(0).Columns(4).HeadingStyle:id=59,.parent=44,.alignment=3"
         _StyleDefs(53)  =   "Splits(0).Columns(4).FooterStyle:id=60,.parent=45"
         _StyleDefs(54)  =   "Splits(0).Columns(4).EditorStyle:id=61,.parent=47"
         _StyleDefs(55)  =   "Splits(0).Columns(5).Style:id=66,.parent=43,.alignment=1,.valignment=1,.locked=0"
         _StyleDefs(56)  =   "Splits(0).Columns(5).HeadingStyle:id=63,.parent=44,.alignment=2"
         _StyleDefs(57)  =   "Splits(0).Columns(5).FooterStyle:id=64,.parent=45"
         _StyleDefs(58)  =   "Splits(0).Columns(5).EditorStyle:id=65,.parent=47"
         _StyleDefs(59)  =   "Splits(0).Columns(6).Style:id=16,.parent=43"
         _StyleDefs(60)  =   "Splits(0).Columns(6).HeadingStyle:id=13,.parent=44"
         _StyleDefs(61)  =   "Splits(0).Columns(6).FooterStyle:id=14,.parent=45"
         _StyleDefs(62)  =   "Splits(0).Columns(6).EditorStyle:id=15,.parent=47"
         _StyleDefs(63)  =   "Splits(0).Columns(7).Style:id=20,.parent=43"
         _StyleDefs(64)  =   "Splits(0).Columns(7).HeadingStyle:id=17,.parent=44"
         _StyleDefs(65)  =   "Splits(0).Columns(7).FooterStyle:id=18,.parent=45"
         _StyleDefs(66)  =   "Splits(0).Columns(7).EditorStyle:id=19,.parent=47"
         _StyleDefs(67)  =   "Splits(0).Columns(8).Style:id=70,.parent=43"
         _StyleDefs(68)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=44"
         _StyleDefs(69)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=45"
         _StyleDefs(70)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=47"
         _StyleDefs(71)  =   "Splits(0).Columns(9).Style:id=74,.parent=43"
         _StyleDefs(72)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=44"
         _StyleDefs(73)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=45"
         _StyleDefs(74)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=47"
         _StyleDefs(75)  =   "Splits(0).Columns(10).Style:id=78,.parent=43"
         _StyleDefs(76)  =   "Splits(0).Columns(10).HeadingStyle:id=75,.parent=44"
         _StyleDefs(77)  =   "Splits(0).Columns(10).FooterStyle:id=76,.parent=45"
         _StyleDefs(78)  =   "Splits(0).Columns(10).EditorStyle:id=77,.parent=47"
         _StyleDefs(79)  =   "Named:id=33:Normal"
         _StyleDefs(80)  =   ":id=33,.parent=0"
         _StyleDefs(81)  =   "Named:id=34:Heading"
         _StyleDefs(82)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(83)  =   ":id=34,.wraptext=-1"
         _StyleDefs(84)  =   "Named:id=35:Footing"
         _StyleDefs(85)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(86)  =   "Named:id=36:Selected"
         _StyleDefs(87)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(88)  =   "Named:id=37:Caption"
         _StyleDefs(89)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(90)  =   "Named:id=38:HighlightRow"
         _StyleDefs(91)  =   ":id=38,.parent=33,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
         _StyleDefs(92)  =   "Named:id=39:EvenRow"
         _StyleDefs(93)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(94)  =   "Named:id=40:OddRow"
         _StyleDefs(95)  =   ":id=40,.parent=33"
         _StyleDefs(96)  =   "Named:id=41:RecordSelector"
         _StyleDefs(97)  =   ":id=41,.parent=34"
         _StyleDefs(98)  =   "Named:id=42:FilterBar"
         _StyleDefs(99)  =   ":id=42,.parent=33"
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "[F4] --> Todos"
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
         Left            =   3105
         TabIndex        =   9
         Top             =   3420
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "INS - Nuevo Producto"
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
         Left            =   8625
         TabIndex        =   6
         Top             =   3585
         Width           =   1575
      End
      Begin VB.Label lblbusca 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   3735
         Width           =   45
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "[F3] --> Filtro Avanzado : Descripcion"
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
         Left            =   90
         TabIndex        =   4
         Top             =   3420
         Width           =   2610
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
         Height          =   420
         Left            =   9030
         TabIndex        =   3
         Top             =   4035
         Visible         =   0   'False
         Width           =   2490
      End
   End
   Begin VB.Menu MNUPRI 
      Caption         =   ""
      Begin VB.Menu mnufiltro 
         Caption         =   "Filtrar"
      End
      Begin VB.Menu mnufiltroavanz 
         Caption         =   "Filtro Avanzado:"
      End
      Begin VB.Menu mnuordasc 
         Caption         =   "Ord. Asc."
      End
      Begin VB.Menu mnuorddesc 
         Caption         =   "Ord. Desc"
      End
      Begin VB.Menu mnutodos 
         Caption         =   "Mostrar Todos"
      End
   End
End
Attribute VB_Name = "hlp_productos_nuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs  As New ADODB.Recordset
Dim col As TrueOleDBGrid70.Column
Dim cols As TrueOleDBGrid70.Columns
Public cad As String

Private Sub cmbfiltro_Change()

    If cmbfiltro.ListIndex = 0 Then DataGrid.col = 0
    If cmbfiltro.ListIndex = 1 Then DataGrid.col = 1
    If cmbfiltro.ListIndex = 2 Then DataGrid.col = 2
    If cmbfiltro.ListIndex = 3 Then DataGrid.col = 3
    If cmbfiltro.ListIndex = 4 Then DataGrid.col = 4
    mnufiltroavanz_Click
End Sub

Private Sub cmbfiltro_LostFocus()

    cmbfiltro_Change
    mnufiltroavanz_Click

End Sub

Private Sub cmdnuevo_Click()

    sw_load_mant = True
    sw_nuevo_mant = True
    mant_productos.Show 1
    sw_nuevo_mant = False
    sw_load_mant = False
    DataAyuda.Refresh
            
End Sub

Private Sub DataGrid_DblClick()
            
    DataGrid_KeyDown 13, 0

End Sub

Private Sub DataGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim msguse  As String
Dim msgdev  As String
    
    Select Case KeyCode
        Case 13:
            wllamada = 0
            wcodproducto = DataGrid.Columns(0) & ""
            wcodfab = DataGrid.Columns(1) & ""
            wdesproducto = DataGrid.Columns(3).Value & ""
            wmedida = DataGrid.Columns(5) & ""
            wstockact = Val(DataGrid.Columns(6) & "")
            wprecos = Val(DataGrid.Columns(7) & "")
            wf5partara = DataGrid.Columns(9) & ""
            wmarca = DataGrid.Columns(2) & ""
            wcodmarca = DataGrid.Columns(10) & ""
            '-----------PARA LAS FORMULAS---------------
            sw = 1
            codmprima = DataGrid.Columns(0) & ""
            nommprima = DataGrid.Columns(3) & ""
            uniprima = DataGrid.Columns(5) & ""
            '-----------FINAL DEL PROCESO---------------
            If Len(Trim(wcod_alm)) = 0 Then
                'DataAyuda.RecordSource = "Select A.f5codpro,A.F5NOMPRO,A.F5CODFAB,A.F5valvta,A.F5PRECOS,B.F7SIGMED,C.F2DESMAR,A.F7CODMED FROM IF5PLA AS A,EF7MEDIDAS AS B,EF2MARCAS AS C WHERE A.F7CODMED=B.F7CODMED AND A.F5MARCA=C.F2CODMAR ORDER BY A.F5NOMPRO"
                DataAyuda.RecordSource = "Select A.F5MARCA,A.f5codpro,IIF(LEN(TRIM(A.F5TEXTO))>0 ,A.F5TEXTO,A.F5NOMPRO) AS F5NOMPRO,A.F5CODFAB,A.F5valvta,A.F5FOB,B.F7SIGMED,C.F2DESMAR,A.F7CODMED FROM IF5PLA AS A,EF7MEDIDAS AS B,EF2MARCAS AS C WHERE A.F7CODMED=B.F7CODMED AND A.F5MARCA=C.F2CODMAR ORDER BY A.F5NOMPRO"
            Else
                If wtipoguia = "I" Then
                    DataAyuda.RecordSource = "Select A.F5MARCA,A.f5codpro,A.F5CODFAB,A.F5NOMPRO,A.F5valvta,A.F5FOB,C.F7SIGMED,D.F2DESMAR,A.F7CODMED FROM IF5PLA AS A,IF6ALMA AS B,EF7MEDIDAS AS C,EF2MARCAS AS D WHERE A.F5CODPRO=B.F5CODPRO AND B.F2CODALM='" & wcod_alm & "' AND A.F7CODMED=C.F7CODMED AND A.F5MARCA=D.F2CODMAR ORDER BY A.F5CODPRO"
                Else   '----- VALES DE SALIDA
                    DataAyuda.RecordSource = "Select A.F5MARCA,A.f5codpro,A.F5CODFAB,A.F5NOMPRO,A.F5valvta,A.F5FOB,C.F7SIGMED,D.F2DESMAR,A.F7CODMED FROM IF5PLA AS A,IF6ALMA AS B,EF7MEDIDAS AS C,EF2MARCAS AS D WHERE A.F5CODPRO=B.F5CODPRO AND B.F2CODALM='" & wcod_alm & "' AND A.F7CODMED=C.F7CODMED AND A.F5MARCA=D.F2CODMAR ORDER BY A.F5NOMPRO ASC"
                End If
            End If
            DataAyuda.Refresh
            '--------------------------
            For Each col In DataGrid.Columns
                col.FilterText = ""
            Next col
            DataAyuda.Recordset.Filter = adFilterNone
            '--------------------------
            Unload Me
        Case 27:
            wcodproducto = ""
            wcodfab = ""
            wdesproducto = ""
            wmedida = ""
            wstockact = 0#
            wprecos = 0#
            wf5partara = ""
            '-----------PARA LAS FORMULAS---------------
            sw = 0
            codmprima = ""
            nommprima = ""
            uniprima = ""
            wcodmarca = ""
            '-----------FINAL DEL PROCESO---------------
            Unload Me
        Case 45:
            sw_load_mant = True
            sw_nuevo_mant = True
            '------------------ para bloquear y solo poder grabar la persona autorizada ----------------------'
            'If wf1mant_productos = "*" Then
            '    msguse = InputBox("Ingrese su Usuario ...", "Mant. Productos ")
            '    msgdev = InputBox("Ingrese su Contraseña ...", "Mant. Productos ", "*")
            '    If rst.State = adStateOpen Then rst.Close
            '    rst.Open "select * from ef2USERS where f2coduseR= '" & msguse & "' AND f2pass_autoriza_documentos='" & msgdev & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            '    If Not rst.EOF Then
            '        MsgBox "Puede Ud. comenzar a realizar el Mantenimiento de Productos  ...", vbInformation, "Atención "
            '        mant_productos.Show 1
            '        sw_nuevo_mant = False
            '        sw_load_mant = False
            '        DataAyuda.Refresh
            '    Else
            '        MsgBox "Ud. no esta autorizado para realizar el Mantenimiento de Productos  ...", vbCritical, "Atención "
            '        'Unload Me
            '    End If
            '    rst.Close
            'Else
                mant_productos.Show 1
                sw_nuevo_mant = False
                sw_load_mant = False
                DataAyuda.Refresh
            'End If
        Case 114:
            Me.DataGrid.col = 2
            mnufiltroavanz_Click
        Case 115:
            cmbfiltro.Visible = False
            txtope.Visible = False
            mnutodos_Click
    End Select

End Sub

Private Sub datagrid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Label1.Visible = True
    
End Sub

Private Sub datagrid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If txtope.Visible = True Then txtope.Visible = False
    If lblbusca.Visible = True Then lblbusca.Visible = False
    MnuFiltro.Caption = "Filtrar [" + DataGrid.Columns(DataGrid.col).Text + "]"
    Select Case Button
        Case 2
            PopupMenu MnuPri
    End Select
    
End Sub

Private Sub Form_Activate()

    DataGrid.col = 3
    If txtope.Visible = True Then
        lblbusca.Visible = True
        lblbusca.Caption = Trim(DataGrid.Columns(3).Caption)
        cmbfiltro.Visible = True
        txtope.Text = ""
    End If
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case 27:
            wcodproducto = ""
            wcodfab = ""
            wdesproducto = ""
            wmedida = ""
            wstockact = 0#
            wprecos = 0#
            sw = 0
            codmprima = ""
            nommprima = ""
            uniprima = ""
            wcodmarca = ""
            Me.Hide
        Case 45:
            sw_load_mant = True
            sw_nuevo_mant = True
            mant_productos.Show 1
            sw_nuevo_mant = False
            sw_load_mant = False
            DataAyuda.Refresh
    End Select

End Sub

Private Sub Form_Load()

    DataGrid.OddRowStyle.BackColor = &HC0FFFF
    DataGrid.EvenRowStyle.BackColor = &HFFFFFF
    DataGrid.HighlightRowStyle.BackColor = vbActiveTitleBar
    DataGrid.HighlightRowStyle.ForeColor = vbWhite
    DataGrid.AlternatingRowStyle = True

    cmbfiltro.AddItem "Codigo", 0
    cmbfiltro.AddItem "Cod. Fabricante", 1
    cmbfiltro.AddItem "Marca", 2
    cmbfiltro.AddItem "Descripción", 3
    cmbfiltro.AddItem "Unidad", 4
    cmbfiltro.ListIndex = 3

    'If wf1visualiza_precio_hlp = "F" Then
        DataGrid.Columns(4).Visible = False
    'End If

    DataAyuda.ConnectionString = cnn_dbbancos
    If Len(Trim(wcod_alm)) = 0 Then
        DataAyuda.RecordSource = "Select A.F5MARCA,A.f5codpro,IIF(LEN(TRIM(A.F5TEXTO))>0 ,A.F5TEXTO,A.F5NOMPRO) AS F5NOMPRO,A.F5CODFAB,A.F5valvta,A.F5FOB,B.F7SIGMED,C.F2DESMAR,A.F7CODMED FROM IF5PLA AS A,EF7MEDIDAS AS B,EF2MARCAS AS C WHERE A.F7CODMED=B.F7CODMED AND A.F5MARCA=C.F2CODMAR " & cad & " ORDER BY A.F5NOMPRO"
        'DataAyuda.RecordSource = "Select A.f5codpro,A.F5NOMPRO,A.F5CODFAB,A.F5valvta,A.F5PRECOS,B.F7SIGMED,C.F2DESMAR,A.F7CODMED FROM IF5PLA AS A,EF7MEDIDAS AS B,EF2MARCAS AS C WHERE A.F7CODMED=B.F7CODMED AND A.F5MARCA=C.F2CODMAR ORDER BY A.F5NOMPRO"
    Else
        If wtipoguia = "I" Then
            Rem NSE DataAyuda.RecordSource = "Select A.f5codpro,A.F5CODFAB,A.F5NOMPRO,A.F5valvta,C.F7SIGMED,D.F2DESMAR,A.F7CODMED FROM IF5PLA AS A,IF6ALMA AS B,EF7MEDIDAS AS C,EF2MARCAS AS D WHERE A.F5CODPRO=B.F5CODPRO AND B.F2CODALM='" & wcod_alm & "' AND A.F7CODMED=C.F7CODMED AND A.F5MARCA=D.F2CODMAR ORDER BY A.F5CODPRO"
            DataAyuda.RecordSource = "Select A.F5MARCA,A.f5codpro,A.F5NOMPRO,A.F5CODFAB,A.F5valvta,A.F5FOB,B.F7SIGMED,C.F2DESMAR,A.F7CODMED FROM IF5PLA AS A,EF7MEDIDAS AS B,EF2MARCAS AS C WHERE A.F7CODMED=B.F7CODMED AND A.F5MARCA=C.F2CODMAR " & cad & " ORDER BY A.F5NOMPRO"
        Else   '----- VALES DE SALIDA                                                                                                                                                                                       '--- and por or ---------'
            Rem EMB DataAyuda.RecordSource = "Select A.f5codpro,A.F5CODFAB,A.F5NOMPRO,A.F5valvta,A.F5PRECOS,C.F7SIGMED,D.F2DESMAR,A.F7CODMED FROM IF5PLA AS A,IF6ALMA AS B,EF7MEDIDAS AS C,EF2MARCAS AS D WHERE A.F5CODPRO=B.F5CODPRO OR B.F2CODALM='" & wcod_alm & "' AND A.F7CODMED=C.F7CODMED AND A.F5MARCA=D.F2CODMAR ORDER BY A.F5NOMPRO ASC"
            DataAyuda.RecordSource = "SELECT A.F5MARCA,A.F5CODPRO, A.F5CODFAB, A.F5NOMPRO, A.F5VALVTA, A.F5FOB, C.F7SIGMED, D.F2DESMAR, A.F7CODMED FROM ((IF5PLA AS A INNER JOIN IF6ALMA AS B ON A.F5CODPRO = B.F5CODPRO) INNER JOIN EF7MEDIDAS AS C ON A.F7CODMED = C.F7CODMED) INNER JOIN EF2MARCAS AS D ON A.F5MARCA = D.F2CODMAR " & _
                                    " WHERE (((A.F5CODPRO)=[B].[F5CODPRO]) AND ((A.F7CODMED)=[C].[F7CODMED]) OR ((B.F2CODALM)='" & wcod_alm & "') AND ((A.F5MARCA)=[D].[F2CODMAR])) ORDER BY A.F5NOMPRO;"
        End If
    End If
    DataAyuda.Refresh
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = 0 Then
       sw = 0
    End If
    
End Sub

Private Sub mnufiltro_Click()
    
    Set rs = DataAyuda.Recordset
    Select Case DataAyuda.Recordset.Fields(DataGrid.Columns(DataGrid.col).DataField).Type
        Case 10
            rs.Filter = "[" + DataGrid.Columns(DataGrid.col).DataField + "]" + " = '" + Trim("" & DataGrid.Columns(DataGrid.col).Text) + "'"
        Case 4
            rs.Filter = "[" + DataGrid.Columns(DataGrid.col).DataField + "]" + " = " + DataGrid.Columns(DataGrid.col).Text
        Case 8
            If IsDate(DataGrid.Columns(DataGrid.col).Text) Then
                rs.Filter = "[" + DataGrid.Columns(DataGrid.col).DataField + "]" + "=#" + DataGrid.Columns(DataGrid.col).Text + "#"
            Else
                MsgBox "Ingrese una Fecha Valida..!", 32, "Advertencia"
                Exit Sub
            End If
    End Select
    Set DataAyuda.Recordset = rs.DataSource
    Set rs = Nothing
    
End Sub

Private Sub mnufiltroavanz_Click()
    
    Select Case DataGrid.col
        Case 0
            lblbusca.Visible = True
            lblbusca.Caption = Trim(DataGrid.Columns(0).Caption)
            cmbfiltro.Visible = True
            txtope.Visible = True
            txtope.Text = ""
            txtope.SetFocus
        Case 1
            lblbusca.Visible = True
            lblbusca.Caption = Trim(DataGrid.Columns(1).Caption)
            cmbfiltro.Visible = True
            txtope.Visible = True
            txtope.Text = ""
            txtope.SetFocus
         Case 2
            lblbusca.Visible = True
            lblbusca.Caption = Trim(DataGrid.Columns(2).Caption)
            cmbfiltro.Visible = True
            cmbfiltro.ListIndex = 2
            txtope.Visible = True
            txtope.Text = ""
            txtope.SetFocus
        Case 3
            lblbusca.Visible = True
            lblbusca.Caption = Trim(DataGrid.Columns(3).Caption)
            cmbfiltro.Visible = True
            cmbfiltro.ListIndex = 3
            txtope.Visible = True
            txtope.Text = ""
            txtope.SetFocus
    End Select
    
End Sub

Private Sub mnuordasc_Click()
    
    Set rs = DataAyuda.Recordset
    rs.Sort = "[" + DataGrid.Columns(DataGrid.col).DataField + "] Asc"
    Set DataAyuda.Recordset = rs.DataSource
    Set rs = Nothing
    
End Sub

Private Sub mnuorddesc_Click()
    
    Set rs = DataAyuda.Recordset
    rs.Sort = "[" + DataGrid.Columns(DataGrid.col).DataField + "] Desc"
    Set DataAyuda.Recordset = rs.DataSource
    Set rs = Nothing
    
End Sub

Private Sub mnutodos_Click()
   
    If Len(Trim(wcod_alm)) = 0 Then
        DataAyuda.RecordSource = "Select A.F5MARCA,A.f5codpro,A.F5NOMPRO,A.F5CODFAB,A.F5valvta,B.F7SIGMED,C.F2DESMAR,A.F7CODMED FROM IF5PLA AS A,EF7MEDIDAS AS B,EF2MARCAS AS C WHERE A.F7CODMED=B.F7CODMED AND A.F5MARCA=C.F2CODMAR ORDER BY A.F5NOMPRO"
    Else
        If wtipoguia = "I" Then
            Rem NSE DataAyuda.RecordSource = "Select A.f5codpro,A.F5CODFAB,A.F5NOMPRO,A.F5valvta,C.F7SIGMED,D.F2DESMAR,A.F7CODMED FROM IF5PLA AS A,IF6ALMA AS B,EF7MEDIDAS AS C,EF2MARCAS AS D WHERE A.F5CODPRO=B.F5CODPRO AND B.F2CODALM='" & wcod_alm & "' AND A.F7CODMED=C.F7CODMED AND A.F5MARCA=D.F2CODMAR ORDER BY A.F5CODPRO"
            DataAyuda.RecordSource = "Select A.F5MARCA,A.f5codpro,A.F5NOMPRO,A.F5CODFAB,A.F5valvta,B.F7SIGMED,C.F2DESMAR,A.F7CODMED FROM IF5PLA AS A,EF7MEDIDAS AS B,EF2MARCAS AS C WHERE A.F7CODMED=B.F7CODMED AND A.F5MARCA=C.F2CODMAR ORDER BY A.F5NOMPRO"
        Else   '----- VALES DE SALIDA
            DataAyuda.RecordSource = "Select A.F5MARCA,A.f5codpro,A.F5CODFAB,A.F5NOMPRO,A.F5valvta,C.F7SIGMED,D.F2DESMAR,A.F7CODMED FROM IF5PLA AS A,IF6ALMA AS B,EF7MEDIDAS AS C,EF2MARCAS AS D WHERE A.F5CODPRO=B.F5CODPRO AND B.F2CODALM='" & wcod_alm & "' AND A.F7CODMED=C.F7CODMED AND A.F5MARCA=D.F2CODMAR ORDER BY A.F5NOMPRO ASC"
        End If
    End If
    DataAyuda.Refresh

End Sub

Private Sub SSPanel1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Label1.Visible = False
 
End Sub

Private Sub txtope_KeyPress(KeyAscii As Integer)
Dim SQL     As String
Dim DATO    As String

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If Len(Trim(txtope.Text)) > 0 Then
            If Len(txtope.Text) = 0 Then txtope.SetFocus: Exit Sub
            If InStr(txtope, "*") = Len(Trim(txtope)) Then
                DATO = Left(txtope, Len(Trim(txtope)) - 1)
            Else
                DATO = txtope.Text
            End If
            txtope.Text = ""
            Set rs = DataAyuda.Recordset
            'cmbfiltro_Change
            rs.Filter = "[" + DataGrid.Columns(DataGrid.col).DataField + "]" + " Like  '" + DATO + "*'"
            If rs.EOF Then txtope.SetFocus: Exit Sub Else txtope.Visible = False: cmbfiltro.Visible = False: lblbusca.Visible = False
            Set DataAyuda.Recordset = rs.DataSource
            Set rs = Nothing
            DataGrid.SetFocus
        End If
    End If
    
End Sub

Private Sub DataGrid_FilterChange()
On Error GoTo errhandler
Set cols = DataGrid.Columns
Dim c As Integer

    c = DataGrid.col
    DataGrid.HoldFields
    DataAyuda.Recordset.Filter = getFilter()
    DataGrid.col = c
    DataGrid.EditActive = True
    Exit Sub
    
errhandler:
    
    MsgBox Err.Source & ":" & vbCrLf & Err.Description
    For Each col In DataGrid.Columns
        col.FilterText = ""
    Next col
    DataAyuda.Recordset.Filter = adFilterNone

End Sub

Private Function getFilter() As String
Dim cadena As String
Dim n As Integer

    For Each col In cols
        If Trim(col.FilterText) <> "" Then
            n = n + 1
            If n > 1 Then
                cadena = cadena & " AND "
            End If
            cadena = cadena & col.DataField & " LIKE '" & col.FilterText & "*'"
        End If
    Next col
    getFilter = cadena

End Function
