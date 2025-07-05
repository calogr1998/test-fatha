VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmSaldoIni 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Saldos Iniciales"
   ClientHeight    =   6480
   ClientLeft      =   4335
   ClientTop       =   3390
   ClientWidth     =   9885
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6480
   ScaleWidth      =   9885
   Begin MSComDlg.CommonDialog Caja01 
      Left            =   6165
      Top             =   225
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Threed.SSCommand BtnExit 
      Height          =   348
      Left            =   7728
      TabIndex        =   11
      Top             =   6000
      Width           =   1116
      _Version        =   65536
      _ExtentX        =   1968
      _ExtentY        =   614
      _StockProps     =   78
      Caption         =   "&Salir"
      ForeColor       =   8388608
      Font3D          =   3
   End
   Begin Threed.SSCommand BtnPrint 
      Height          =   345
      Left            =   4500
      TabIndex        =   10
      Top             =   6000
      Width           =   1065
      _Version        =   65536
      _ExtentX        =   1884
      _ExtentY        =   614
      _StockProps     =   78
      Caption         =   "&Imprimir"
      ForeColor       =   8388608
      Font3D          =   3
   End
   Begin Threed.SSCommand BtnGraba 
      Height          =   345
      Left            =   3330
      TabIndex        =   9
      Top             =   6000
      Width           =   1065
      _Version        =   65536
      _ExtentX        =   1884
      _ExtentY        =   614
      _StockProps     =   78
      Caption         =   "&Grabar"
      ForeColor       =   8388608
      Font3D          =   3
   End
   Begin Threed.SSCommand btnimpdeta 
      Height          =   348
      Left            =   192
      TabIndex        =   8
      Top             =   6000
      Visible         =   0   'False
      Width           =   1644
      _Version        =   65536
      _ExtentX        =   2900
      _ExtentY        =   614
      _StockProps     =   78
      Caption         =   "Imprimir &Detallado"
      ForeColor       =   8388608
      Font3D          =   3
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   828
      Left            =   192
      TabIndex        =   3
      Top             =   -48
      Width           =   5820
      _Version        =   65536
      _ExtentX        =   10266
      _ExtentY        =   1460
      _StockProps     =   14
      Caption         =   "Buscar"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Font3D          =   3
      Begin VB.TextBox Txtnombre 
         Appearance      =   0  'Flat
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
         Left            =   2436
         MaxLength       =   50
         TabIndex        =   5
         Top             =   336
         Width           =   2940
      End
      Begin VB.TextBox Txtcodigo 
         Appearance      =   0  'Flat
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
         Left            =   732
         MaxLength       =   10
         TabIndex        =   4
         Top             =   336
         Width           =   870
      End
      Begin VB.Label LblMsg 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   0
         Left            =   144
         TabIndex        =   7
         Top             =   336
         Width           =   540
      End
      Begin VB.Label LblMsg 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   1
         Left            =   1764
         TabIndex        =   6
         Top             =   336
         Width           =   600
      End
   End
   Begin Threed.SSCommand BtnMenos 
      Height          =   300
      Left            =   7920
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   1785
      _Version        =   65536
      _ExtentX        =   3154
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "<< &Quitar"
      ForeColor       =   8388608
      Font3D          =   3
   End
   Begin Threed.SSCommand Btnmas 
      Height          =   345
      Left            =   7920
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1785
      _Version        =   65536
      _ExtentX        =   3154
      _ExtentY        =   614
      _StockProps     =   78
      Caption         =   "Más Productos..."
      ForeColor       =   8388608
      Font3D          =   3
   End
   Begin VB.ListBox lista 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1830
      Left            =   2640
      TabIndex        =   0
      Top             =   1536
      Visible         =   0   'False
      Width           =   3888
   End
   Begin VB.Data DataDetalle 
      Appearance      =   0  'Flat
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   7440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2160
      Visible         =   0   'False
      Width           =   1275
   End
   Begin TrueOleDBGrid70.TDBGrid GridDetalle 
      Bindings        =   "t_salini.frx":0000
      Height          =   4815
      Left            =   135
      TabIndex        =   12
      Top             =   855
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   8493
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Código"
      Columns(0).DataField=   "f5codpro"
      Columns(0).DataWidth=   255
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "f5codfab"
      Columns(1).DataField=   "f5codfab"
      Columns(1).DataWidth=   255
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Descripción"
      Columns(2).DataField=   "f5nompro"
      Columns(2).DataWidth=   255
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Saldo Inicial"
      Columns(3).DataField=   "f6canmov"
      Columns(3).DataWidth=   23
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Unidad"
      Columns(4).DataField=   "f7sigmed"
      Columns(4).DataWidth=   255
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Costo Unit. S/."
      Columns(5).DataField=   "f6valpro"
      Columns(5).DataWidth=   23
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "f6total"
      Columns(6).DataField=   "f6total"
      Columns(6).DataWidth=   23
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "f5afecto"
      Columns(7).DataField=   "f5afecto"
      Columns(7).DataWidth=   255
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Costo Unit. $"
      Columns(8).DataField=   "f6valprod"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   9
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=9"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2143"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2064"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=65808"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=556"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=476"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=65808"
      Splits(0)._ColumnProps(12)=   "Column(1).Visible=0"
      Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(14)=   "Column(2).Width=6324"
      Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=6244"
      Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=65808"
      Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(20)=   "Column(3).Width=1826"
      Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=1746"
      Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=65538"
      Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(26)=   "Column(4).Width=1217"
      Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=1138"
      Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=65808"
      Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(32)=   "Column(5).Width=2117"
      Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=2037"
      Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=65538"
      Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(38)=   "Column(6).Width=212"
      Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=132"
      Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=65538"
      Splits(0)._ColumnProps(43)=   "Column(6).Visible=0"
      Splits(0)._ColumnProps(44)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(45)=   "Column(7).Width=238"
      Splits(0)._ColumnProps(46)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(47)=   "Column(7)._WidthInPix=159"
      Splits(0)._ColumnProps(48)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(49)=   "Column(7)._ColStyle=65808"
      Splits(0)._ColumnProps(50)=   "Column(7).Visible=0"
      Splits(0)._ColumnProps(51)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(52)=   "Column(8).Width=1773"
      Splits(0)._ColumnProps(53)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(54)=   "Column(8)._WidthInPix=1693"
      Splits(0)._ColumnProps(55)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(56)=   "Column(8)._ColStyle=65808"
      Splits(0)._ColumnProps(57)=   "Column(8).Order=9"
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=172,.bold=0,.fontsize=780,.italic=0"
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
      _StyleDefs(47)  =   "Splits(0).Columns(3).Style:id=62,.parent=43,.alignment=1"
      _StyleDefs(48)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=44,.alignment=3"
      _StyleDefs(49)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=45"
      _StyleDefs(50)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=47"
      _StyleDefs(51)  =   "Splits(0).Columns(4).Style:id=66,.parent=43,.alignment=0,.valignment=2"
      _StyleDefs(52)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=44"
      _StyleDefs(53)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=45"
      _StyleDefs(54)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=47"
      _StyleDefs(55)  =   "Splits(0).Columns(5).Style:id=70,.parent=43,.alignment=1"
      _StyleDefs(56)  =   "Splits(0).Columns(5).HeadingStyle:id=67,.parent=44,.alignment=3"
      _StyleDefs(57)  =   "Splits(0).Columns(5).FooterStyle:id=68,.parent=45"
      _StyleDefs(58)  =   "Splits(0).Columns(5).EditorStyle:id=69,.parent=47"
      _StyleDefs(59)  =   "Splits(0).Columns(6).Style:id=74,.parent=43,.alignment=1"
      _StyleDefs(60)  =   "Splits(0).Columns(6).HeadingStyle:id=71,.parent=44,.alignment=3"
      _StyleDefs(61)  =   "Splits(0).Columns(6).FooterStyle:id=72,.parent=45"
      _StyleDefs(62)  =   "Splits(0).Columns(6).EditorStyle:id=73,.parent=47"
      _StyleDefs(63)  =   "Splits(0).Columns(7).Style:id=78,.parent=43,.alignment=0,.valignment=2"
      _StyleDefs(64)  =   "Splits(0).Columns(7).HeadingStyle:id=75,.parent=44"
      _StyleDefs(65)  =   "Splits(0).Columns(7).FooterStyle:id=76,.parent=45"
      _StyleDefs(66)  =   "Splits(0).Columns(7).EditorStyle:id=77,.parent=47"
      _StyleDefs(67)  =   "Splits(0).Columns(8).Style:id=82,.parent=43,.alignment=0,.valignment=2"
      _StyleDefs(68)  =   "Splits(0).Columns(8).HeadingStyle:id=79,.parent=44"
      _StyleDefs(69)  =   "Splits(0).Columns(8).FooterStyle:id=80,.parent=45"
      _StyleDefs(70)  =   "Splits(0).Columns(8).EditorStyle:id=81,.parent=47"
      _StyleDefs(71)  =   "Named:id=33:Normal"
      _StyleDefs(72)  =   ":id=33,.parent=0"
      _StyleDefs(73)  =   "Named:id=34:Heading"
      _StyleDefs(74)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(75)  =   ":id=34,.wraptext=-1"
      _StyleDefs(76)  =   "Named:id=35:Footing"
      _StyleDefs(77)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(78)  =   "Named:id=36:Selected"
      _StyleDefs(79)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(80)  =   "Named:id=37:Caption"
      _StyleDefs(81)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(82)  =   "Named:id=38:HighlightRow"
      _StyleDefs(83)  =   ":id=38,.parent=33,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(84)  =   "Named:id=39:EvenRow"
      _StyleDefs(85)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(86)  =   "Named:id=40:OddRow"
      _StyleDefs(87)  =   ":id=40,.parent=33"
      _StyleDefs(88)  =   "Named:id=41:RecordSelector"
      _StyleDefs(89)  =   ":id=41,.parent=34"
      _StyleDefs(90)  =   "Named:id=42:FilterBar"
      _StyleDefs(91)  =   ":id=42,.parent=33"
   End
   Begin VB.Menu menu400 
      Caption         =   "&lngreso"
      Visible         =   0   'False
      Begin VB.Menu menuedit 
         Caption         =   "&Ingreso de Series"
      End
   End
End
Attribute VB_Name = "FrmSaldoIni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbstock             As DAO.Database
Dim DbTempfac           As DAO.Database
Dim Tbl                 As New TableDef
Dim Fld                 As Field
Dim Tabla               As Table
Dim swsaldo             As Integer
Dim swcod               As Integer
Dim swnom               As Integer
Dim TIPO                As Double
Dim wfila               As Integer
Dim tbniveles           As Table
Dim CONSULTA            As DAO.Recordset
Dim parcial             As Double
Dim TOTAL               As Double

Dim tbnivelesR As DAO.Recordset  'FVG
Dim tbstock As DAO.Recordset
Dim TbMovAlmacen As DAO.Recordset
Dim TBPRODUCTO As DAO.Recordset



'''Private Sub Agrega_Productos()
'''
'''           TbMovAlmacen.Seek "=", gcodalm, Trim(FrmHelpProduc.datagrid.Columns(0).Text)
'''            If TbMovAlmacen.NoMatch Then
'''                tbstock.AddNew
'''                tbstock.Fields("f5codpro") = "" & FrmHelpProduc.datagrid.Columns(0).Text
'''                tbstock.Fields("f5nompro") = "" & FrmHelpProduc.datagrid.Columns(2).Text
'''                tbstock.Fields("f6canmov") = 0#
'''                tbstock.Fields("f7sigmed") = "" & FrmHelpProduc.datagrid.Columns(3).Text
'''                tbstock.Fields("f6valpro") = 0#
'''                tbstock.Fields("f6total") = 0#
'''                tbstock.Update
'''           Else
'''            MsgBox "Codigo ya Existe"
'''           End If
'''
'''        DataDetalle.Refresh
'''
'''        GridDetalle.Refresh
'''
'''End Sub

Private Sub BtnExit_Click()

    Unload Me

End Sub

Private Sub BtnGraba_Click()
    
    Me.MousePointer = 11
    tbstock.MoveFirst
    Do While Not tbstock.EOF
        DataDetalle.UpdateRecord
        Graba_Producto
        tbstock.MoveNext
    Loop
    Me.MousePointer = 1
        
End Sub

Private Sub BtnImpdeta_Click()

On Error GoTo seguido

    If MsgBox("Desea imprimir el Saldo Inicial Detallado", 36, "Almacén") = 6 Then
        lista.Visible = True
        lista.ListIndex = 0
        lista.SetFocus
    Else
        If MsgBox("Desea imprimir el Saldo Inicial Totalizado", 36, "Almacén") = 6 Then
            Totalizado
        End If
    End If

    Exit Sub
seguido: Resume Next

End Sub

Private Sub BtnMas_Click()
    
    Me.MousePointer = 11
    Rem NSE FrmHelpProduc.Show 1
    Rem NSE Agrega_Productos
    Me.MousePointer = 1

End Sub

Private Sub BtnMenos_Click()

    Me.MousePointer = 11
    TbMovAlmacen.Seek "=", gcodalm, Trim(GridDetalle.Columns(0).Text) 'Text(0))
    If TbMovAlmacen.NoMatch Then
        Quita_Productos
    End If
    Me.MousePointer = 1

End Sub

Private Sub BtnPrint_Click()
    
On Error GoTo fin
    Caja01.CancelError = True
    Caja01.Action = 5
    Me.MousePointer = 11
    Imprimir_Saldo
    Me.MousePointer = 1
fin: Exit Sub

End Sub

Private Sub Detallado()

    Me.MousePointer = 11

    cabeceras
    Printer.FontSize = 10
    writexy "Almacén:", 5, 1, 0
    writexy Trim(gcodalm & " - " & FrmSaldoAlm.Txtnomalm.Text), 5, 12, 0
    Printer.FontBold = True
    Printer.FontUnderline = True
    Printer.FontSize = 11
    writexy Trim(UCase(Mid(lista.Text, 1, 90))), 7, (100 - Len(Trim(UCase(Mid(lista.Text, 1, 90))))) / 2, 0
    Printer.FontUnderline = False
    writexy "( US$ )", 8, (100 - Len("( US$ )")) / 2, 0
    Printer.FontSize = 8
    Titulo_Detallado
    wfila = 13

    parcial = 0#
    Set CONSULTA = DbInventa.OpenRecordset("select * from if6alma where trim(f2codalm)='" + Trim(gcodalm) + "' and left(f5codpro,2)='" + Trim(Right(lista.Text, 10)) + "' order by f5codpro")
    If CONSULTA.RecordCount > 0 Then
        Do While Not CONSULTA.EOF
            TBPRODUCTO.Index = "IDCODPRO"
            TBPRODUCTO.Seek "=", Trim(CONSULTA.Fields("f5codpro"))
            If Not TBPRODUCTO.NoMatch Then
                writexy Trim("" & CONSULTA.Fields("f5codpro")), wfila, 0, 0
                writexy Trim("" & Left(TBPRODUCTO.Fields("f5nompro"), 50)), wfila, 10, 0
                writexy Trim("" & TBPRODUCTO.Fields("f7codmed")), wfila, 60, 0
                writexy Format(CONSULTA.Fields("f5debm00"), "#0.00"), wfila, 63, 2
                If Trim("" & TBPRODUCTO.Fields("f5moneda")) = "D" Then
                    writexy Format(CONSULTA.Fields("F5COSTOINI"), "#0.00"), wfila, 74, 2
                    If Trim("" & CONSULTA.Fields("F5COSTOINI")) <> "" And Trim("" & CONSULTA.Fields("f5debm00")) <> "" Then
                        If Format(CONSULTA.Fields("F5COSTOINI"), "#0.00") > 0# And Format("" & CONSULTA.Fields("f5debm00"), "#0.00") > 0# Then
                            writexy Format(CONSULTA.Fields("F5COSTOINI"), "#0.00") * Format(CONSULTA.Fields("f5debm00"), "#0.00"), wfila, 87, 2
                            parcial = parcial + Format(CONSULTA.Fields("F5COSTOINI"), "#0.00") * Format(CONSULTA.Fields("f5debm00"), "#0.00")
                        End If
                    End If
                Else
                    writexy Format(CONSULTA.Fields("F5COSTOINI") * gtipcam, "#0.00"), wfila, 74, 2
                    If Trim("" & CONSULTA.Fields("F5COSTOINI")) <> "" And Trim("" & CONSULTA.Fields("f5debm00")) <> "" Then
                        If Format(CONSULTA.Fields("F5COSTOINI"), "#0.00") > 0# And Format(CONSULTA.Fields("f5debm00"), "#0.00") > 0# Then
                            writexy Format(CONSULTA.Fields("F5COSTOINI") * gtipcam * Format(CONSULTA.Fields("f5debm00"), "#0.00"), "#0.00"), wfila, 87, 2
                            parcial = parcial + Format(CONSULTA.Fields("F5COSTOINI") * gtipcam * Format(CONSULTA.Fields("f5debm00"), "#0.00"), "#0.00")
                        End If
                    End If
                End If
                wfila = wfila + 1
                If wfila >= 60 Then
                    Printer.NewPage
                    cabeceras
                    Printer.FontSize = 10
                    writexy "Almacén:", 5, 1, 0
                    writexy Trim(gcodalm & " - " & FrmSaldoAlm.Txtnomalm.Text), 5, 12, 0
                    Printer.FontBold = True
                    Printer.FontUnderline = True
                    Printer.FontSize = 11
                    writexy Trim(UCase(Mid(lista.Text, 1, 90))), 7, (100 - Len(Trim(UCase(Mid(lista.Text, 1, 90))))) / 2, 0
                    Printer.FontUnderline = False
                    writexy "( US$ )", 8, (100 - Len("( US$ )")) / 2, 0
                    Printer.FontSize = 8
                    Titulo_Detallado
                    wfila = 13
                End If
            End If
            CONSULTA.MoveNext
            If CONSULTA.EOF Then
                CONSULTA.Close
                Exit Do
            End If
        Loop
    End If
    wfila = wfila + 1
    Printer.Line (0, wfila)-(100, wfila)
    wfila = wfila + 1
    Printer.FontSize = 9
    writexy "Total ===>  " & " US$ ", wfila, 50, 0
    writexy Format(parcial, "#0.00"), wfila, 87, 2
    wfila = wfila + 2
    Printer.EndDoc
    
    lista.Visible = False
    BtnExit.SetFocus

    Me.MousePointer = 1

End Sub

Private Sub Form_Load()
    
    Set dbempresa = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set tbnivelesR = dbempresa.OpenRecordset("sf7nivel01")
    
    Set DbInventa = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set TBPRODUCTO = DbInventa.OpenRecordset("IF5PLA")
    Set TbMovAlmacen = DbInventa.OpenRecordset("IF6ALMA")
    Set TbStockCab = DbInventa.OpenRecordset("if4vales")
    Set TbStockDet = DbInventa.OpenRecordset("if3vales")
    
    Set DbTempfac = OpenDatabase(App.Path & "\TempFac.mdb")
    
    
    
    Set dbstock = OpenDatabase(App.Path & "\tempfac.mdb")
    dbstock.Execute ("delete * from tmpstock")
    
    
    Set tbstock = dbstock.OpenRecordset("tmpstock")
    
      
    TBPRODUCTO.Index = "IDCODPRO"
    TbMovAlmacen.Index = "IDALMPROTRA"

    Llena_Grid
    Llena_Combo
    
    swsaldo = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error GoTo seguir

    DbTempfac.Execute ("Delete From tmpstock")
    DataDetalle.Refresh
    
    tbnivelesR.Close
    CONSULTA.Close
    TBPRODUCTO.Close
    TbMovAlmacen.Close
    TbStockCab.Close
    TbStockDet.Close

    dbempresa.Close
    DbTempfac.Close
    DbInventa.Close
    tbstock.Close
    dbstock.Close

    Exit Sub
seguir: Resume Next

End Sub

Private Sub Graba_Producto()
Dim i       As Integer

    Dim wstock      As Double
    wstock = 0#
    If tbstock!f6canmov > 0 Then
        TbMovAlmacen.Seek "=", gcodalm, tbstock!F5CODPRO
        If TbMovAlmacen.NoMatch Then
            TbMovAlmacen.AddNew
            TbMovAlmacen.Fields("F2codalm") = gcodalm
            TbMovAlmacen.Fields("F5codpro") = tbstock!F5CODPRO
        Else
            TbMovAlmacen.Edit
        End If

        TBPRODUCTO.Seek "=", tbstock!F5CODPRO
        If Not TBPRODUCTO.NoMatch Then
           TBPRODUCTO.Edit
           'X = Val(Format(TBPRODUCTO.Fields("F5STOCKACT"), "0.000000")) + Val(Format(tbstock!f6canmov, "#0.000")) - Val(Format(TbMovAlmacen.Fields("F5DEBM00"), "#0.000000"))
           TBPRODUCTO.Fields("F5STOCKACT") = Val(Format(TBPRODUCTO.Fields("F5STOCKACT"), "0.000000")) + Val(Format(tbstock!f6canmov, "#0.000")) - Val(Format(TbMovAlmacen.Fields("F5DEBM00"), "#0.000000"))
           TBPRODUCTO.Update
        End If

        TbMovAlmacen.Fields("F5DEBM00") = Val(Format(tbstock!f6canmov, "#0.000000"))
        TbMovAlmacen.Fields("F5COSTOINI") = Val(Format(tbstock!f6valpro, "#0.000000"))

        For i = 0 To 12   ' VER calculo fvg
            wstock = wstock + (Val(Format(TbMovAlmacen.Fields("F5DEBM" & Format(i, "00")), "#0.000000")) - Val(Format(TbMovAlmacen.Fields("F5HABM" & Format(i, "00")), "#0.000")))
        Next i
        TbMovAlmacen.Fields("F6STOCKACT") = wstock
        TbMovAlmacen.Fields("F6STOCKLOG") = wstock
        
        If Trim(UCase(TBPRODUCTO.Fields("f5moneda"))) = "S" Then
            ''A CUATRO DECIMALES FVG  OBS
            TbMovAlmacen.Fields("F5ING00") = Val(Format(tbstock!f6canmov, "#0.000000")) * Val(Format(tbstock!f6valpro, "#0.000000"))
            TbMovAlmacen.Fields("F5INGD00") = Val(Format(tbstock!f6canmov, "#0.000000")) * (Val(Format(tbstock!f6valprod, "#0.000000"))) '/ gtipcam
            TbMovAlmacen.Fields("F5COSPRO") = Val(Format(tbstock!f6valpro, "#0.000000"))
            TbMovAlmacen.Fields("F5COSPROD") = Val(Format(tbstock!f6valpro, "#0.000000")) / gtipcam
        Else
            TbMovAlmacen.Fields("F5ING00") = Val(Format(tbstock!f6canmov, "#0.000000")) * Val(Format(tbstock!f6valpro, "#0.000000")) '* gtipcam
            TbMovAlmacen.Fields("F5INGD00") = Val(Format(tbstock!f6canmov, "#0.000000")) * (Val(Format(tbstock!f6valprod, "#0.000000")))
            TbMovAlmacen.Fields("F5COSPRO") = Val(Format(tbstock!f6valpro, "#0.000000")) * gtipcam
            TbMovAlmacen.Fields("F5COSPROD") = Val(Format(tbstock!f6valpro, "#0.000000"))
        End If
        TbMovAlmacen.Update
        swsaldo = 1
    Else
        DataDetalle.UpdateRecord
        TbMovAlmacen.Seek "=", gcodalm, tbstock!F5CODPRO
        If Not TbMovAlmacen.NoMatch Then
            TBPRODUCTO.Seek "=", tbstock!F5CODPRO
            
            TbMovAlmacen.Edit
            If Trim(UCase(TBPRODUCTO.Fields("f5moneda"))) = "S" Then
                TbMovAlmacen.Fields("F5COSPRO") = Val(Format(tbstock!f6valpro, "#0.000"))
                TbMovAlmacen.Fields("F5COSPROD") = Val(Format(tbstock!f6valpro, "#0.000")) / gtipcam
            Else
                TbMovAlmacen.Fields("F5COSPRO") = Val(Format(tbstock!f6valpro, "#0.000")) * gtipcam
                TbMovAlmacen.Fields("F5COSPROD") = Val(Format(tbstock!f6valpro, "#0.000"))
            End If
    
            TbMovAlmacen.Fields("F5DEBM00") = Val(Format(tbstock!f6canmov, "#0.000")) 'FVG
            TbMovAlmacen.Fields("F5COSTOINI") = Val(Format(tbstock!f6valpro, "#0.000000"))
    
            TbMovAlmacen.Update
        End If
    End If

    'DataDetalle.Recordset.MoveNext
End Sub



Private Sub griddetalle_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim wrow    As Integer

    If KeyCode = 13 Then
       Select Case GridDetalle.col
         Case 0: GridDetalle.col = 2: GridDetalle.SetFocus
         Case 1: GridDetalle.col = 2: GridDetalle.SetFocus
         Case 2: GridDetalle.Columns(2) = Val(Format(GridDetalle.Columns(2), "###,##0.000"))
                 GridDetalle.col = 4: GridDetalle.SetFocus
         Case 3: GridDetalle.col = 4: GridDetalle.SetFocus
         Case 4: GridDetalle.Columns(2) = Val(Format(GridDetalle.Columns(4), "###,##0.000"))
                wrow = GridDetalle.Row
                If wrow = DataDetalle.Recordset.RecordCount Then
                    GridDetalle.Row = 0
                Else
                    GridDetalle.Row = wrow + 1
                End If
                GridDetalle.col = 2: GridDetalle.SetFocus
       End Select
    End If

    If KeyCode = 114 Then
        If swcod Then
            DataDetalle.Recordset.FindNext "F5CODPRO Like '" + Trim(txtcodigo.Text) & "*" + "'"
        Else
            DataDetalle.Recordset.FindNext "F5NOMPRO Like '" + "*" & Trim(TxtNombre.Text) & "*" + "'"
        End If
        'GridDetalle.SelStartRow = GridDetalle.Row
        'GridDetalle.SelEndRow = GridDetalle.Row
        GridDetalle.SetFocus
    End If

End Sub

Private Sub GridDetalle_KeyPress(KeyAscii As Integer)

    Select Case GridDetalle.col
        Case 0, 1, 2: KeyAscii = 7
        'Case 2: KeyAscii = IngCar(GridDetalle, 2, 2, 10, KeyAscii)
        'Case 4: KeyAscii = IngCar(GridDetalle, 4, 2, 10, KeyAscii)
    End Select

End Sub

Private Sub GridDetalle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
       PopupMenu menu400
    End If

End Sub

Private Sub Imprime_Titulo()

    cabeceras
    Printer.FontSize = 14
    Printer.FontBold = True
    Printer.FontUnderline = True
    writexy "SALDOS INICIALES", 5, 32, 0
    Printer.FontUnderline = False
    Printer.FontSize = 10
    writexy "Almacén:", 8, 1, 0
    writexy Trim(gcodalm & " - " & FrmSaldoAlm.Txtnomalm.Text), 8, 12, 0
    Printer.FontSize = 9
    Printer.Line (0, 10)-(100, 10)
    writexy "Código", 11, 1, 0
    writexy "Descripción", 11, 14, 0
    writexy "Saldo Inicial", 11, 43, 0
    writexy "Und.", 11, 58, 0
    writexy "Precio-Costo", 11, 68, 0
    writexy "Total", 11, 86, 0
    Printer.Line (0, 12)-(100, 12)
    Printer.FontBold = False

End Sub

Private Sub Imprimir_Saldo()
Dim i   As Double
Dim TOTAL As Double
    TOTAL = 0
    Printer.ScaleMode = 4
    Printer.FontName = "Courier New" 'Printer.Fonts(TbParametro.Fields("F1FONNAM"))
    
    Imprime_Titulo
    wfila = 13
    For i = 0 To DataDetalle.Recordset.RecordCount ' GridDetalle.Rows - 1
        GridDetalle.Row = i
        If Val(Format("" & GridDetalle.Columns(3).Text, "0.000")) > 0 Then
            
            writexy Trim(GridDetalle.Columns(0).Text), wfila, 1, 0    'codigo
            writexy Left(GridDetalle.Columns(2).Text, 38), wfila, 12, 0 'descripcion
            writexy Format(GridDetalle.Columns(3).Text, "#0.00"), wfila, 43, 2 'SALDO
            writexy Trim(GridDetalle.Columns(4).Text), wfila, 58, 0
            writexy Format(GridDetalle.Columns(5).Text, "#0.00"), wfila, 65, 2
            'mod
            TOTAL = Val(Format(GridDetalle.Columns(3), "#0.00")) * Val(Format(GridDetalle.Columns(5).Text, "#0.00")) + TOTAL
            writexy Val(Format(GridDetalle.Columns(3), "#0.00")) * Val(Format(GridDetalle.Columns(5).Text, "#0.00")), wfila, 81, 2

            wfila = wfila + 1
            If wfila >= 60 Then
               Printer.NewPage
               Imprime_Titulo
               wfila = 13
            End If
        End If
    Next i
    Printer.Line (0, wfila)-(100, wfila)
    writexy "total General==>              " & Format(TOTAL, "#0.00"), wfila, 59, 2
    Printer.EndDoc

End Sub

Private Sub GridDetalle1_Click()

End Sub

Private Sub lista_DblClick()

    lista_keypress 13

End Sub

Private Sub lista_keypress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Detallado
    End If

End Sub

Private Sub Llena_Combo()

    lista.Clear
    tbnivelesR.Index = "IDNIVEL01"
    tbnivelesR.MoveFirst
    Do While Not tbnivelesR.EOF
        lista.AddItem UCase(Trim(tbnivelesR.Fields("f7descon"))) & Space(100) & Trim(tbnivelesR.Fields("f7codcon"))
        tbnivelesR.MoveNext
        If tbnivelesR.EOF Then Exit Do
    Loop

End Sub

Private Sub Llena_Grid()
    TbMovAlmacen.Seek ">=", gcodalm
    If Not TbMovAlmacen.NoMatch Then
        Do While TbMovAlmacen.Fields("F2codalm") = gcodalm
            TBPRODUCTO.Seek "=", TbMovAlmacen.Fields("F5codpro")
            If Not TBPRODUCTO.NoMatch Then
                tbstock.AddNew
                'DataDetalle.Recordset.AddNew
                'DataDetalle.Recordset.Fields("F5codpro") = "" & TbMovAlmacen.Fields("F5CODPRO")
                tbstock!F5CODPRO = "" & TbMovAlmacen.Fields("F5CODPRO")
                'DataDetalle.Recordset.Fields("F5nompro") = "" & Left(TbProducto.Fields("F5NOMPRO"), 50)
                tbstock!F5NOMPRO = "" & Left(TBPRODUCTO.Fields("F5NOMPRO"), 50)
               ' DataDetalle.Recordset.Fields("F6canmov") = Format(TbMovAlmacen.Fields("F5DEBM00"), "###,##0.00")
               tbstock!f6canmov = Format(TbMovAlmacen.Fields("F5DEBM00"), "###,##0.000000")
               ' DataDetalle.Recordset.Fields("F7sigmed") = Format(TbProducto.Fields("F7CODMED"), "###,##0.00")
                tbstock!f7sigmed = Format(TBPRODUCTO.Fields("F7CODMED"), "###,##0.000000")
                'DataDetalle.Recordset.Fields("f6valpro") = Format(TbMovAlmacen.Fields("F5COSTOINI"), "0.00")
               '' DataDetalle.Recordset.Fields("f6valpro") = VAL(0 & TbMovAlmacen.Fields("F5COSTOINI"))
               
               If TbMovAlmacen.Fields("F5DEBM00") = 0 Then
                tbstock!f6valpro = 0
                'tbstock!f6valprod = 0
               Else
                tbstock!f6valpro = Format(Val(0 & TbMovAlmacen.Fields("F5ING00")) / Val(0 & TbMovAlmacen.Fields("F5DEBM00")), "#0.000000")
                'tbstock!f6valprod = Format(Val(0 & TbMovAlmacen.Fields("F5INGD00")) / Val(0 & TbMovAlmacen.Fields("F5DEBM00")), "#0.000000")
               End If
               ' DataDetalle.Recordset.Update
               tbstock.Update
             End If
            TbMovAlmacen.MoveNext
            If TbMovAlmacen.EOF Then Exit Do
        Loop
    End If
    DataDetalle.DatabaseName = App.Path & "\TempFac.Mdb"
    DataDetalle.RecordSource = "TmpStock"
    DataDetalle.Refresh
End Sub
Private Sub menuedit_Click()

    TBPRODUCTO.Index = "IDCODPRO"
    TBPRODUCTO.Seek "=", Trim(GridDetalle.Columns(0).Text) 'GridDetalle.ColText(0))
    If Not TBPRODUCTO.NoMatch Then
      If TBPRODUCTO.Fields("F5SERIES") = "1" And Val(GridDetalle.Columns(2)) > 0 Then 'ColText(2) > 0 Then
         Me.MousePointer = 11
            'FRMINGSERIES6.Show 1
         Me.MousePointer = 1
      End If
    End If


End Sub

Private Sub Quita_Productos()

   ' If GridDetalle.Rows > 0 Then
        DataDetalle.Recordset.Delete
        DataDetalle.Refresh
   ' End If

End Sub

Private Sub Titulo_Detallado()

    Printer.FontSize = 9
    Printer.Line (0, 10)-(100, 10)
    writexy "CODIGO", 11, 1, 0
    writexy "DESCRIPCION", 11, 30, 0
    writexy "UND", 11, 60, 0
    writexy "STOCK", 11, 67, 0
    writexy "COST.UNIT.", 11, 75, 0
    writexy "COST.PARC.", 11, 88, 0
    Printer.Line (0, 12)-(100, 12)
    Printer.FontBold = False

End Sub

Private Sub Totalizado()

    Me.MousePointer = 11

    cabeceras
    Printer.FontSize = 10
    writexy "Almacén:", 5, 1, 0
    writexy Trim(gcodalm & " - " & FrmSaldoAlm.Txtnomalm.Text), 5, 12, 0
    Printer.FontBold = True
    Printer.FontUnderline = True
    Printer.FontSize = 11
    writexy "SALDOS INICIALES", 7, 40, 0
    Printer.FontUnderline = False
    writexy "( US$ )", 8, 44, 0
    Printer.FontSize = 8
    Titulo_Detallado
    wfila = 13

    TOTAL = 0#
    tbnivelesR.Index = "IDNIVEL01"
    tbnivelesR.MoveFirst
    Do While Not tbnivelesR.EOF
        parcial = 0#
        Set CONSULTA = DbInventa.OpenRecordset("select * from if6alma where trim(f2codalm)='" + Trim(gcodalm) + "' and left(f5codpro,2)='" + Trim(tbnivelesR.Fields("f7codcon")) + "' order by f5codpro")
        If CONSULTA.RecordCount > 0 Then
            Do While Not CONSULTA.EOF
                TBPRODUCTO.Index = "IDCODPRO"
                TBPRODUCTO.Seek "=", Trim(CONSULTA.Fields("f5codpro"))
                If Not TBPRODUCTO.NoMatch Then
                    writexy Trim("" & CONSULTA.Fields("f5codpro")), wfila, 0, 0
                    writexy Trim("" & Left(TBPRODUCTO.Fields("f5nompro"), 50)), wfila, 10, 0
                    writexy Trim("" & TBPRODUCTO.Fields("f7codmed")), wfila, 60, 0
                    writexy Format(CONSULTA.Fields("f5debm00"), "#0.00"), wfila, 63, 2
                    If Trim("" & TBPRODUCTO.Fields("f5moneda")) = "D" Then
                        writexy Format(CONSULTA.Fields("F5COSTOINI"), "#0.00"), wfila, 74, 2
                        If Trim("" & CONSULTA.Fields("F5COSTOINI")) <> "" And Trim("" & CONSULTA.Fields("f5debm00")) <> "" Then
                            If Format(CONSULTA.Fields("F5COSTOINI"), "#0.00") > 0# And Format("" & CONSULTA.Fields("f5debm00"), "#0.00") > 0# Then
                                writexy Format(CONSULTA.Fields("F5COSTOINI"), "#0.00") * Format(CONSULTA.Fields("f5debm00"), "#0.00"), wfila, 87, 2
                                parcial = parcial + Format(CONSULTA.Fields("F5COSTOINI"), "#0.00") * Format(CONSULTA.Fields("f5debm00"), "#0.00")
                                TOTAL = TOTAL + parcial
                            End If
                        End If
                    Else
                        writexy Format(CONSULTA.Fields("F5COSTOINI") * gtipcam, "#0.00"), wfila, 74, 2
                        If Trim("" & CONSULTA.Fields("F5COSTOINI")) <> "" And Trim("" & CONSULTA.Fields("f5debm00")) <> "" Then
                            If Format(CONSULTA.Fields("F5COSTOINI"), "#0.00") > 0# And Format(CONSULTA.Fields("f5debm00"), "#0.00") > 0# Then
                                writexy Format(CONSULTA.Fields("F5COSTOINI") * gtipcam * Format(CONSULTA.Fields("f5debm00"), "#0.00"), "#0.00"), wfila, 87, 2
                                parcial = parcial + Format(CONSULTA.Fields("F5COSTOINI") * gtipcam * Format(CONSULTA.Fields("f5debm00"), "#0.00"), "#0.00")
                                TOTAL = TOTAL + parcial
                            End If
                        End If
                    End If
                    wfila = wfila + 1
                    If wfila >= 60 Then
                        Printer.NewPage
                        cabeceras
                        Printer.FontSize = 10
                        writexy "Almacén:", 5, 1, 0
                        writexy Trim(gcodalm & " - " & FrmSaldoAlm.Txtnomalm.Text), 5, 12, 0
                        Printer.FontBold = True
                        Printer.FontUnderline = True
                        Printer.FontSize = 11
                        writexy "SALDOS INICIALES", 7, 40, 0
                        Printer.FontUnderline = False
                        writexy "( US$ )", 8, 44, 0
                        Printer.FontSize = 8
                        Titulo_Detallado
                        wfila = 13
                    End If
                End If
                CONSULTA.MoveNext
                If CONSULTA.EOF Then
                CONSULTA.Close
                Exit Do
                End If
            Loop
        End If
        wfila = wfila + 1
        Printer.Line (0, wfila)-(100, wfila)
        wfila = wfila + 1
        writexy "Total de " & tbnivelesR.Fields("f7descon") & "  ===>  ", wfila, 45, 0
        writexy Format(parcial, "#0.00"), wfila, 85, 2
        wfila = wfila + 2
        tbnivelesR.MoveNext
        If tbnivelesR.EOF Then Exit Do
    Loop
    wfila = wfila + 1
    Printer.Line (0, wfila)-(100, wfila)
    wfila = wfila + 1
    writexy "Total General" & "===>", wfila, 50, 0
    writexy Format(TOTAL, "#0.00"), wfila, 85, 2
    Printer.EndDoc
    
    BtnExit.SetFocus

    Me.MousePointer = 1

End Sub

Private Sub Txtcodigo_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        DataDetalle.Recordset.FindFirst "F5CODPRO Like '" + txtcodigo.Text & "*" + "'"
        If Not DataDetalle.Recordset.NoMatch Then
            swcod = 1
            swnom = 0
            GridDetalle.SetFocus
        Else
            MsgBox ("No se encontró el registro"), 64, "Productos"
        End If
    End If

End Sub

Private Sub txtnombre_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        DataDetalle.Recordset.FindFirst "F5NOMPRO Like '" + "*" & Trim(TxtNombre.Text) & "*" + "'"
        If Not DataDetalle.Recordset.NoMatch Then
            swcod = 0
            swnom = 1
            GridDetalle.SetFocus
        Else
            MsgBox ("No se encontró el registro"), 64, "Productos"
        End If
    End If

End Sub

