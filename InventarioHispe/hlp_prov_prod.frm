VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "SSTBARS2.OCX"
Begin VB.Form hlp_prov_prod 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista de Productos por Proveedor"
   ClientHeight    =   5010
   ClientLeft      =   1965
   ClientTop       =   3015
   ClientWidth     =   11055
   BeginProperty Font 
      Name            =   "Arial"
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
   ScaleHeight     =   5010
   ScaleWidth      =   11055
   Begin Threed.SSPanel SSPanel1 
      Height          =   4425
      Left            =   45
      TabIndex        =   1
      Top             =   90
      Width           =   10950
      _Version        =   65536
      _ExtentX        =   19315
      _ExtentY        =   7805
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
      Begin MSAdodcLib.Adodc DataAyudaPP 
         Height          =   465
         Left            =   3330
         Top             =   3735
         Visible         =   0   'False
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   820
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         TabIndex        =   0
         ToolTipText     =   "Simplifique la consulta digitando el Simbolo (*)"
         Top             =   3975
         Width           =   2445
      End
      Begin TrueOleDBGrid70.TDBGrid DataGrid 
         Bindings        =   "hlp_prov_prod.frx":0000
         Height          =   3210
         Left            =   90
         TabIndex        =   2
         Top             =   90
         Width           =   10770
         _ExtentX        =   18997
         _ExtentY        =   5662
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Codigo"
         Columns(0).DataField=   "F5CODPRO"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Cod. Fab."
         Columns(1).DataField=   "F5CODFAB"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Descripción"
         Columns(2).DataField=   "F5NOMPRO"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Precio. Vta"
         Columns(3).DataField=   "F5valvta"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Unidad"
         Columns(4).DataField=   "f7codmed"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Marca"
         Columns(5).DataField=   "f5marca"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "F5AFECTO"
         Columns(6).DataField=   "F5AFECTO"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   7
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   2
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   13154464
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=7"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1879"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1799"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=74000"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=2514"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2434"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).AllowSizing=0"
         Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=74000"
         Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(15)=   "Column(2).Width=11827"
         Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=11748"
         Splits(0)._ColumnProps(18)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(19)=   "Column(2).AllowSizing=0"
         Splits(0)._ColumnProps(20)=   "Column(2)._ColStyle=74000"
         Splits(0)._ColumnProps(21)=   "Column(2).WrapText=1"
         Splits(0)._ColumnProps(22)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(23)=   "Column(2).AutoCompletion=1"
         Splits(0)._ColumnProps(24)=   "Column(3).Width=1535"
         Splits(0)._ColumnProps(25)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(26)=   "Column(3)._WidthInPix=1455"
         Splits(0)._ColumnProps(27)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(28)=   "Column(3).AllowSizing=0"
         Splits(0)._ColumnProps(29)=   "Column(3)._ColStyle=73730"
         Splits(0)._ColumnProps(30)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(31)=   "Column(4).Width=1058"
         Splits(0)._ColumnProps(32)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(4)._WidthInPix=979"
         Splits(0)._ColumnProps(34)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(35)=   "Column(4).AllowSizing=0"
         Splits(0)._ColumnProps(36)=   "Column(4)._ColStyle=74273"
         Splits(0)._ColumnProps(37)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(38)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(39)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(40)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(41)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(42)=   "Column(5)._ColStyle=65792"
         Splits(0)._ColumnProps(43)=   "Column(5).Visible=0"
         Splits(0)._ColumnProps(44)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(45)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(46)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(47)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(48)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(49)=   "Column(6)._ColStyle=65792"
         Splits(0)._ColumnProps(50)=   "Column(6).Visible=0"
         Splits(0)._ColumnProps(51)=   "Column(6).Order=7"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         MultiSelect     =   2
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
         _StyleDefs(8)   =   ":id=1,.fontname=Arial"
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
         _StyleDefs(36)  =   ":id=28,.locked=-1"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=44"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=45"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=47"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(41)  =   ":id=32,.locked=-1"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=44"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=45"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=47"
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=58,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(46)  =   ":id=58,.wraptext=-1,.locked=-1"
         _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=44"
         _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=45"
         _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=47,.wraptext=-1"
         _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=62,.parent=43,.alignment=1,.locked=-1"
         _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=44,.alignment=3"
         _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=45"
         _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=47"
         _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=66,.parent=43,.alignment=2,.valignment=1"
         _StyleDefs(55)  =   ":id=66,.locked=-1"
         _StyleDefs(56)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=44,.alignment=2"
         _StyleDefs(57)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=45"
         _StyleDefs(58)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=47"
         _StyleDefs(59)  =   "Splits(0).Columns(5).Style:id=16,.parent=43"
         _StyleDefs(60)  =   "Splits(0).Columns(5).HeadingStyle:id=13,.parent=44"
         _StyleDefs(61)  =   "Splits(0).Columns(5).FooterStyle:id=14,.parent=45"
         _StyleDefs(62)  =   "Splits(0).Columns(5).EditorStyle:id=15,.parent=47"
         _StyleDefs(63)  =   "Splits(0).Columns(6).Style:id=20,.parent=43"
         _StyleDefs(64)  =   "Splits(0).Columns(6).HeadingStyle:id=17,.parent=44"
         _StyleDefs(65)  =   "Splits(0).Columns(6).FooterStyle:id=18,.parent=45"
         _StyleDefs(66)  =   "Splits(0).Columns(6).EditorStyle:id=19,.parent=47"
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "INS - Nuevo Producto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   9045
         TabIndex        =   6
         Top             =   3420
         Width           =   1545
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
         Left            =   135
         TabIndex        =   5
         Top             =   3690
         Width           =   45
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "[F3] --> Filtro Avanzado : Descripcion"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   90
         TabIndex        =   4
         Top             =   3420
         Width           =   2700
      End
      Begin VB.Label Label1 
         Caption         =   "Ayuda: Hacer Click Derecho en la Columna que desee buscar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   8325
         TabIndex        =   3
         Top             =   3825
         Visible         =   0   'False
         Width           =   2490
      End
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   180
      Top             =   4590
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Tools           =   "hlp_prov_prod.frx":001A
      ToolBars        =   "hlp_prov_prod.frx":4C0A
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
Attribute VB_Name = "hlp_prov_prod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim csql        As String
Dim rs          As New ADODB.Recordset
Dim rsprov_prod As New ADODB.Recordset
Dim inicio      As Boolean

Private Sub DataGrid_DblClick()
            
    DataGrid_KeyDown 13, 0

End Sub

Private Sub DataGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case 13:
            wcodproducto = DataGrid.Columns(0) & ""
            wcodfab = DataGrid.Columns(1) & ""
            wdesproducto = DataGrid.Columns(2) & ""
            wvv_prod = Format(Val(DataGrid.Columns(3)), "0.00")
            wmedida = DataGrid.Columns(4) & ""
            wmarca = DataGrid.Columns(5) & ""
            wafecto = DataGrid.Columns(6) & ""
            Unload Me
        Case 27:
            wcodproducto = ""
            wcodfab = ""
            wdesproducto = ""
            wmedida = ""
            wmarca = ""
            wafecto = ""
            Unload Me
        Case 45:
            sw_load_mant = True
            sw_nuevo_mant = True
            mant_productos.Show 1
            sw_nuevo_mant = False
            sw_load_mant = False
            DataAyudaPP.Refresh
        Case 114:
            Me.DataGrid.col = 2
            mnufiltroavanz_Click
    End Select

End Sub

Private Sub datagrid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Label1.Visible = True
    
End Sub

Private Sub datagrid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If txtope.Visible = True Then txtope.Visible = False
    If lblbusca.Visible = True Then lblbusca.Visible = False
    mnufiltro.Caption = "Filtrar [" + DataGrid.Columns(DataGrid.col).Text + "]"
    Select Case Button
        Case 2
            PopupMenu mnupri
    End Select
    
End Sub

Private Sub Form_Activate()
If wcodproducto = "" And Not inicio Then Exit Sub
    DataGrid.OddRowStyle.BackColor = &HC0FFFF
    DataGrid.EvenRowStyle.BackColor = &HFFFFFF
    DataGrid.HighlightRowStyle.BackColor = vbActiveTitleBar
    DataGrid.HighlightRowStyle.ForeColor = vbWhite
    DataGrid.AlternatingRowStyle = True

    If wf1visualiza_precio_hlp = "F" Then
        DataGrid.Columns(3).Visible = False
    End If

    DataAyudaPP.ConnectionString = cnn_dbbancos
    'csql = "SELECT A.F5CODPRO,B.F5NOMPRO,B.F5CODFAB,A.F5VALVTA,B.F7CODMED,B.F5MARCA FROM EF2PROD_PROV AS A,IF5PLA AS B WHERE A.F2CODPRV='" & wrucprov & "' AND A.F5CODPRO=B.F5CODPRO ORDER BY B.F5CODPRO"
    Select Case WMONEDAX
    Case "S"
        csql = "SELECT A.F5CODPRO, IIf(Len("" & B.F5NOMPRO)>Len("" & B.F5TEXTO),B.F5NOMPRO,B.F5TEXTO) AS F5NOMPRO,  B.F5AFECTO, B.F5CODFAB, B.F5VALVTA, B.F7CODMED, B.F5MARCA FROM EF2PROD_PROV AS A INNER JOIN IF5PLA AS B ON A.F5CODPRO = B.F5CODPRO WHERE (((A.F5CODPRO)=[B].[F5CODPRO]) AND ((A.F2CODPRV)='" & wrucprov & "')) ORDER BY B.F5CODPRO;"
    Case "D"
        csql = "SELECT A.F5CODPRO, IIf(Len("" & B.F5NOMPRO)>Len("" & B.F5TEXTO),B.F5NOMPRO,B.F5TEXTO) AS F5NOMPRO,  B.F5AFECTO, B.F5CODFAB, B.F5VALVTA, B.F7CODMED, B.F5MARCA FROM EF2PROD_PROV AS A INNER JOIN IF5PLA AS B ON A.F5CODPRO = B.F5CODPRO WHERE (((A.F5CODPRO)=[B].[F5CODPRO]) AND ((A.F2CODPRV)='" & wrucprov & "')) ORDER BY B.F5CODPRO;"
    Case "E"
        csql = "SELECT A.F5CODPRO, IIf(Len("" & B.F5NOMPRO)>Len("" & B.F5TEXTO),B.F5NOMPRO,B.F5TEXTO) AS F5NOMPRO,  B.F5AFECTO, B.F5CODFAB, B.F5VALVTA, B.F7CODMED, B.F5MARCA FROM EF2PROD_PROV AS A INNER JOIN IF5PLA AS B ON A.F5CODPRO = B.F5CODPRO WHERE (((A.F5CODPRO)=[B].[F5CODPRO]) AND ((A.F2CODPRV)='" & wrucprov & "')) ORDER BY B.F5CODPRO;"
    Case Else
        csql = "SELECT A.F5CODPRO, IIf(Len(B.F5NOMPRO)>Len(B.F5TEXTO),B.F5NOMPRO,B.F5TEXTO) AS F5NOMPRO,  B.F5AFECTO, B.F5CODFAB, B.F5VALVTA, B.F7CODMED, B.F5MARCA FROM EF2PROD_PROV AS A INNER JOIN IF5PLA AS B ON A.F5CODPRO = B.F5CODPRO WHERE (((A.F5CODPRO)=[B].[F5CODPRO]) AND ((A.F2CODPRV)='" & wrucprov & "')) ORDER BY B.F5CODPRO;"
    End Select
    
    DataAyudaPP.RecordSource = csql
    DataAyudaPP.Refresh
    
    If txtope.Visible = True Then
        txtope.SetFocus
    End If
    DataGrid.col = 2
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case 27:
            wcodproducto = ""
            wcodfab = ""
            wdesproducto = ""
            wmedida = ""
            wafecto = ""
            Unload Me
        Case 45:
            sw_load_mant = True
            sw_nuevo_mant = True
            mant_productos.Show 1
            sw_nuevo_mant = False
            sw_load_mant = False
            DataAyudaPP.Refresh
    End Select

End Sub

Private Sub Form_Load()
    
    inicio = True
    txtope.Visible = True

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If sw_ayuda_prod = True Then
        Unload ayuda_productos
    End If
    
End Sub

Private Sub mnufiltro_Click()
    
    Set rs = DataAyudaPP.Recordset
    Select Case DataAyudaPP.Recordset.Fields(DataGrid.Columns(DataGrid.col).DataField).Type
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
    Set DataAyudaPP.Recordset = rs.DataSource
    Set rs = Nothing
    
End Sub

Private Sub mnufiltroavanz_Click()
    
    Select Case DataGrid.col
        Case 0
            lblbusca.Visible = True
            lblbusca.Caption = Trim(DataGrid.Columns(0).Caption)
            txtope.Visible = True
            txtope.Text = ""
            txtope.SetFocus
        Case 1
            lblbusca.Visible = True
            lblbusca.Caption = Trim(DataGrid.Columns(1).Caption)
            txtope.Visible = True
            txtope.Text = ""
            txtope.SetFocus
         Case 2
            lblbusca.Visible = True
            lblbusca.Caption = Trim(DataGrid.Columns(2).Caption)
            txtope.Visible = True
            txtope.Text = ""
            txtope.SetFocus
    End Select
    
End Sub

Private Sub mnuordasc_Click()
    
    Set rs = DataAyudaPP.Recordset
    rs.Sort = "[" + DataGrid.Columns(DataGrid.col).DataField + "] Asc"
    Set DataAyudaPP.Recordset = rs.DataSource
    Set rs = Nothing
    
End Sub

Private Sub mnuorddesc_Click()
    
    Set rs = DataAyudaPP.Recordset
    rs.Sort = "[" + DataGrid.Columns(DataGrid.col).DataField + "] Desc"
    Set DataAyudaPP.Recordset = rs.DataSource
    Set rs = Nothing
    
End Sub

Private Sub mnutodos_Click()
   
    DataAyudaPP.RecordSource = csql
    DataAyudaPP.Refresh

End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    
    Select Case Tool.Id
        Case "ID_MásProductos":
            wcod_alm = "": wcodproducto = ""
            sw_ayuda_prod = True
            ayuda_productos.Show 1
            inicio = False
            If Len(Trim(wcodproducto)) > 0 Then
                Dim wdes
                wdes = ""
                If rsprov_prod.State = adStateOpen Then rsprov_prod.Close
                rsprov_prod.Open "SELECT F5CODPRO FROM EF2PROD_PROV WHERE F2CODPRV='" & wrucprov & "' AND F5CODPRO='" & wcodproducto & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If rsprov_prod.EOF Then
                    csql = "INSERT INTO EF2PROD_PROV (F2CODPRV,F2NOMPRV,F5CODPRO,F5NOMPRO,F5VALVTA,F5CODFAB,F7CODMED) " & _
                           " VALUES('" & wrucprov & "','" & wnomprov & "','" & wcodproducto & "','" & wdesproducto & _
                           "',0,'" & wcodfab & "','" & wmedida & "')"
                    cnn_dbbancos.Execute (csql)
                                        
                    Dim RSTa As New ADODB.Recordset
                                        
                    If RSTa.State = adStateOpen Then RSTa.Close
                    csql = "select f5nompro from ef2prod_prov where f2codprv='" & wrucprov & "' and f5codpro='" & wcodproducto & "'"
                    RSTa.Open csql, cnn_dbbancos, adOpenStatic, adLockOptimistic
                    RSTa("f5nompro").Value = wdesproducto
                    RSTa.Update
                    RSTa.Close
                    'cnn_dbbancos.Execute (csql)s
                    DataAyudaPP.Refresh
                Else
                    MsgBox "El código del producto ya fue asignado al proveedor. Verifique.", vbInformation, "Atención"
                End If
                DataAyudaPP.Refresh
            End If
        Case "ID_Salir":
            Unload Me
    End Select
    
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
            txtope.Text = "*" & txtope.Text
            If Len(txtope) = 0 Then txtope.SetFocus: Exit Sub
            If InStr(txtope, "*") = Len(Trim(txtope)) Then
                DATO = Left(txtope, Len(Trim(txtope)) - 1)
            Else
                DATO = txtope.Text
            End If
            txtope.Text = ""
            Set rs = DataAyudaPP.Recordset
            rs.Filter = "[" + DataGrid.Columns(DataGrid.col).DataField + "]" + " Like  '" + DATO + "*'"
            If rs.EOF Then txtope.SetFocus: Exit Sub Else txtope.Visible = False: lblbusca.Visible = False
            Set DataAyudaPP.Recordset = rs.DataSource
            Set rs = Nothing
            DataGrid.SetFocus
        End If
    End If
    
End Sub
