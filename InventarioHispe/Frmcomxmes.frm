VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmcomxmes 
   Caption         =   "Resumen de Compras x Mes"
   ClientHeight    =   5175
   ClientLeft      =   1665
   ClientTop       =   1380
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   11400
   Begin ComctlLib.Toolbar tblbar 
      Height          =   390
      Left            =   45
      TabIndex        =   18
      Top             =   0
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   688
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   2
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Imprimir"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   825
      Left            =   45
      TabIndex        =   21
      Top             =   450
      Width           =   11265
      _Version        =   65536
      _ExtentX        =   19870
      _ExtentY        =   1455
      _StockProps     =   15
      BackColor       =   13160660
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
      Begin VB.TextBox txtdesde 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1620
         MaxLength       =   2
         TabIndex        =   0
         Top             =   270
         Width           =   480
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   600
         Left            =   5940
         TabIndex        =   23
         ToolTipText     =   "Hacer click soles o dolares"
         Top             =   90
         Width           =   3795
         _Version        =   65536
         _ExtentX        =   6694
         _ExtentY        =   1058
         _StockProps     =   14
         Caption         =   "Moneda"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Begin Threed.SSOption opdolares 
            Height          =   240
            Left            =   2430
            TabIndex        =   2
            Top             =   270
            Width           =   960
            _Version        =   65536
            _ExtentX        =   1693
            _ExtentY        =   423
            _StockProps     =   78
            Caption         =   "Dólares"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption opsoles 
            Height          =   240
            Left            =   540
            TabIndex        =   1
            Top             =   270
            Width           =   1185
            _Version        =   65536
            _ExtentX        =   2090
            _ExtentY        =   423
            _StockProps     =   78
            Caption         =   "Soles"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Mes"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   990
         TabIndex        =   22
         Top             =   315
         Width           =   390
      End
   End
   Begin VB.TextBox txtbusca 
      Height          =   285
      Left            =   576
      TabIndex        =   12
      ToolTipText     =   "Simplifique la consulta digitando el Simbolo (*)"
      Top             =   4668
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.TextBox txtplazini 
      Height          =   330
      Left            =   360
      MaxLength       =   3
      TabIndex        =   11
      Top             =   5160
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txtplazfin 
      Height          =   330
      Left            =   1056
      MaxLength       =   3
      TabIndex        =   10
      Top             =   5160
      Visible         =   0   'False
      Width           =   420
   End
   Begin MSMask.MaskEdBox MkEdFin 
      Height          =   300
      Left            =   2340
      TabIndex        =   13
      Top             =   4710
      Visible         =   0   'False
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox MkEdIni 
      Height          =   300
      Left            =   570
      TabIndex        =   14
      Top             =   4710
      Visible         =   0   'False
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin Crystal.CrystalReport reportetotal 
      Left            =   11040
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
   End
   Begin Threed.SSPanel PnlMovPen 
      Height          =   2040
      Left            =   2445
      TabIndex        =   4
      Top             =   2070
      Visible         =   0   'False
      Width           =   6840
      _Version        =   65536
      _ExtentX        =   12065
      _ExtentY        =   3598
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
      BorderWidth     =   1
      BevelInner      =   1
      Begin Threed.SSCommand cmdsalir 
         Height          =   252
         Left            =   5952
         TabIndex        =   20
         Top             =   96
         Width           =   588
         _Version        =   65536
         _ExtentX        =   1037
         _ExtentY        =   444
         _StockProps     =   78
         Caption         =   "Salir"
      End
      Begin TrueOleDBGrid70.TDBGrid Grd_Mvto1 
         Bindings        =   "Frmcomxmes.frx":0000
         Height          =   1590
         Left            =   90
         TabIndex        =   19
         Top             =   390
         Visible         =   0   'False
         Width           =   6630
         _ExtentX        =   11695
         _ExtentY        =   2805
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Banco"
         Columns(0).DataField=   "Banco"
         Columns(0).DataWidth=   255
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Comprobante"
         Columns(1).DataField=   "n_cuenta"
         Columns(1).DataWidth=   255
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "fecha"
         Columns(2).DataField=   "fecha"
         Columns(2).DataWidth=   19
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Moneda"
         Columns(3).DataField=   "MONEDA"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Imputado"
         Columns(4).DataField=   "imputado"
         Columns(4).DataWidth=   255
         Columns(4).NumberFormat=   "Standard"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   5
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=5"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=3598"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3519"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=66080"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=4657"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=4577"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=66080"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=1640"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1561"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=66082"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=1217"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1138"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=66080"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=3043"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2963"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=66082"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
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
         _StyleDefs(31)  =   "Splits(0).Columns(0).Style:id=28,.parent=43,.alignment=0,.valignment=1"
         _StyleDefs(32)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=44,.alignment=2"
         _StyleDefs(33)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=45"
         _StyleDefs(34)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=47"
         _StyleDefs(35)  =   "Splits(0).Columns(1).Style:id=32,.parent=43,.alignment=0,.valignment=1"
         _StyleDefs(36)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=44,.alignment=2"
         _StyleDefs(37)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=45"
         _StyleDefs(38)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=47"
         _StyleDefs(39)  =   "Splits(0).Columns(2).Style:id=58,.parent=43,.alignment=1,.valignment=1"
         _StyleDefs(40)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=44,.alignment=2"
         _StyleDefs(41)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=45"
         _StyleDefs(42)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=47"
         _StyleDefs(43)  =   "Splits(0).Columns(3).Style:id=62,.parent=43,.alignment=0,.valignment=1"
         _StyleDefs(44)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=44,.alignment=2"
         _StyleDefs(45)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=45"
         _StyleDefs(46)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=47"
         _StyleDefs(47)  =   "Splits(0).Columns(4).Style:id=66,.parent=43,.alignment=1,.valignment=1"
         _StyleDefs(48)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=44,.alignment=2"
         _StyleDefs(49)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=45"
         _StyleDefs(50)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=47"
         _StyleDefs(51)  =   "Named:id=33:Normal"
         _StyleDefs(52)  =   ":id=33,.parent=0"
         _StyleDefs(53)  =   "Named:id=34:Heading"
         _StyleDefs(54)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(55)  =   ":id=34,.wraptext=-1"
         _StyleDefs(56)  =   "Named:id=35:Footing"
         _StyleDefs(57)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(58)  =   "Named:id=36:Selected"
         _StyleDefs(59)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(60)  =   "Named:id=37:Caption"
         _StyleDefs(61)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(62)  =   "Named:id=38:HighlightRow"
         _StyleDefs(63)  =   ":id=38,.parent=33,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
         _StyleDefs(64)  =   "Named:id=39:EvenRow"
         _StyleDefs(65)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(66)  =   "Named:id=40:OddRow"
         _StyleDefs(67)  =   ":id=40,.parent=33"
         _StyleDefs(68)  =   "Named:id=41:RecordSelector"
         _StyleDefs(69)  =   ":id=41,.parent=34"
         _StyleDefs(70)  =   "Named:id=42:FilterBar"
         _StyleDefs(71)  =   ":id=42,.parent=33"
      End
      Begin VB.Data DataMvto 
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   324
         Left            =   144
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   45
         Visible         =   0   'False
         Width           =   1770
      End
      Begin TrueOleDBGrid70.TDBGrid Grd_Mvto 
         Bindings        =   "Frmcomxmes.frx":0017
         Height          =   1596
         Left            =   96
         TabIndex        =   5
         Top             =   384
         Width           =   6432
         _ExtentX        =   11351
         _ExtentY        =   2805
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Banco"
         Columns(0).DataField=   "Banco"
         Columns(0).DataWidth=   255
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Cuenta"
         Columns(1).DataField=   "n_cuenta"
         Columns(1).DataWidth=   255
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Fecha"
         Columns(2).DataField=   "fecha"
         Columns(2).DataWidth=   19
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=3545"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3466"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=66080"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=4657"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=4577"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=66080"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=2117"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2037"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=66082"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         Appearance      =   0
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=144,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=44,.alignment=2"
         _StyleDefs(37)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=45"
         _StyleDefs(38)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=47"
         _StyleDefs(39)  =   "Splits(0).Columns(1).Style:id=32,.parent=43,.alignment=0,.valignment=1"
         _StyleDefs(40)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=44,.alignment=2"
         _StyleDefs(41)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=45"
         _StyleDefs(42)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=47"
         _StyleDefs(43)  =   "Splits(0).Columns(2).Style:id=58,.parent=43,.alignment=1,.valignment=1"
         _StyleDefs(44)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=44,.alignment=2"
         _StyleDefs(45)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=45"
         _StyleDefs(46)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=47"
         _StyleDefs(47)  =   "Named:id=33:Normal"
         _StyleDefs(48)  =   ":id=33,.parent=0"
         _StyleDefs(49)  =   "Named:id=34:Heading"
         _StyleDefs(50)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(51)  =   ":id=34,.wraptext=-1"
         _StyleDefs(52)  =   "Named:id=35:Footing"
         _StyleDefs(53)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(54)  =   "Named:id=36:Selected"
         _StyleDefs(55)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(56)  =   "Named:id=37:Caption"
         _StyleDefs(57)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(58)  =   "Named:id=38:HighlightRow"
         _StyleDefs(59)  =   ":id=38,.parent=33,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
         _StyleDefs(60)  =   "Named:id=39:EvenRow"
         _StyleDefs(61)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(62)  =   "Named:id=40:OddRow"
         _StyleDefs(63)  =   ":id=40,.parent=33"
         _StyleDefs(64)  =   "Named:id=41:RecordSelector"
         _StyleDefs(65)  =   ":id=41,.parent=34"
         _StyleDefs(66)  =   "Named:id=42:FilterBar"
         _StyleDefs(67)  =   ":id=42,.parent=33"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         Caption         =   "Movimientos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   2460
         TabIndex        =   6
         Top             =   90
         Width           =   1065
      End
   End
   Begin TrueOleDBGrid70.TDBGrid Grd_Docum 
      Bindings        =   "Frmcomxmes.frx":0035
      Height          =   3000
      Left            =   45
      TabIndex        =   3
      Top             =   1350
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   5292
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Documento"
      Columns(0).DataField=   "F4DOC"
      Columns(0).DataWidth=   255
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Nombre"
      Columns(1).DataField=   "F4NOMPRV"
      Columns(1).DataWidth=   255
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Fecha"
      Columns(2).DataField=   "F4FECHA"
      Columns(2).DataWidth=   255
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Moneda"
      Columns(3).DataField=   "F4MONEDA"
      Columns(3).DataWidth=   255
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Total"
      Columns(4).DataField=   "F4TOTAL"
      Columns(4).DataWidth=   22
      Columns(4).NumberFormat=   "Standard"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "F4REFERE"
      Columns(5).DataField=   "F4REFERE"
      Columns(5).DataWidth=   255
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "F4CORRELA"
      Columns(6).DataField=   "F4CORRELA"
      Columns(6).DataWidth=   23
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "f4mov"
      Columns(7).DataField=   "f4mov"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "T.C."
      Columns(8).DataField=   "F4TIPCAM"
      Columns(8).DataWidth=   22
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Pagado"
      Columns(9).DataField=   "f4pagado"
      Columns(9).NumberFormat=   "Standard"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Saldo"
      Columns(10).DataField=   "f4saldo"
      Columns(10).NumberFormat=   "Standard"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   11
      Splits(0)._UserFlags=   0
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).ScrollBars=   2
      Splits(0).AllowColSelect=   0   'False
      Splits(0).AllowRowSelect=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=11"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2937"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2858"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=74273"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=8176"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=8096"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=66080"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=1879"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1799"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=74273"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=1296"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1217"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=74273"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=1852"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=1773"
      Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=74546"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=1931"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=1852"
      Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=74273"
      Splits(0)._ColumnProps(30)=   "Column(5).Visible=0"
      Splits(0)._ColumnProps(31)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(32)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(33)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(35)=   "Column(6)._ColStyle=73730"
      Splits(0)._ColumnProps(36)=   "Column(6).Visible=0"
      Splits(0)._ColumnProps(37)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(38)=   "Column(7).Width=6006"
      Splits(0)._ColumnProps(39)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(40)=   "Column(7)._WidthInPix=5927"
      Splits(0)._ColumnProps(41)=   "Column(7)._ColStyle=74000"
      Splits(0)._ColumnProps(42)=   "Column(7).Visible=0"
      Splits(0)._ColumnProps(43)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(44)=   "Column(8).Width=2037"
      Splits(0)._ColumnProps(45)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(46)=   "Column(8)._WidthInPix=1958"
      Splits(0)._ColumnProps(47)=   "Column(8)._ColStyle=74274"
      Splits(0)._ColumnProps(48)=   "Column(8).Visible=0"
      Splits(0)._ColumnProps(49)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(50)=   "Column(9).Width=1561"
      Splits(0)._ColumnProps(51)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(52)=   "Column(9)._WidthInPix=1482"
      Splits(0)._ColumnProps(53)=   "Column(9)._ColStyle=74546"
      Splits(0)._ColumnProps(54)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(55)=   "Column(10).Width=1720"
      Splits(0)._ColumnProps(56)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(57)=   "Column(10)._WidthInPix=1640"
      Splits(0)._ColumnProps(58)=   "Column(10)._ColStyle=74546"
      Splits(0)._ColumnProps(59)=   "Column(10).Order=11"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=48,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(31)  =   "Splits(0).Columns(0).Style:id=28,.parent=43,.alignment=2,.valignment=1"
      _StyleDefs(32)  =   ":id=28,.locked=-1"
      _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=44,.alignment=2"
      _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=45"
      _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=47"
      _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=43,.alignment=0,.valignment=1"
      _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=44,.alignment=2"
      _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=45"
      _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=47"
      _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=58,.parent=43,.alignment=2,.valignment=1"
      _StyleDefs(41)  =   ":id=58,.locked=-1"
      _StyleDefs(42)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=44,.alignment=2"
      _StyleDefs(43)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=45"
      _StyleDefs(44)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=47"
      _StyleDefs(45)  =   "Splits(0).Columns(3).Style:id=62,.parent=43,.alignment=2,.valignment=1"
      _StyleDefs(46)  =   ":id=62,.locked=-1"
      _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=44,.alignment=2"
      _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=45"
      _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=47"
      _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=66,.parent=43,.alignment=1,.valignment=3"
      _StyleDefs(51)  =   ":id=66,.locked=-1"
      _StyleDefs(52)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=44,.alignment=1"
      _StyleDefs(53)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=45"
      _StyleDefs(54)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=47"
      _StyleDefs(55)  =   "Splits(0).Columns(5).Style:id=70,.parent=43,.alignment=2,.valignment=1"
      _StyleDefs(56)  =   ":id=70,.locked=-1"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=67,.parent=44,.alignment=2"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=68,.parent=45"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=69,.parent=47"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=74,.parent=43,.alignment=1,.locked=-1"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=71,.parent=44,.alignment=3"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=72,.parent=45"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=73,.parent=47"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=78,.parent=43,.alignment=0,.valignment=2"
      _StyleDefs(65)  =   ":id=78,.locked=-1"
      _StyleDefs(66)  =   "Splits(0).Columns(7).HeadingStyle:id=75,.parent=44"
      _StyleDefs(67)  =   "Splits(0).Columns(7).FooterStyle:id=76,.parent=45"
      _StyleDefs(68)  =   "Splits(0).Columns(7).EditorStyle:id=77,.parent=47"
      _StyleDefs(69)  =   "Splits(0).Columns(8).Style:id=82,.parent=43,.alignment=1,.valignment=1"
      _StyleDefs(70)  =   ":id=82,.locked=-1"
      _StyleDefs(71)  =   "Splits(0).Columns(8).HeadingStyle:id=79,.parent=44,.alignment=2"
      _StyleDefs(72)  =   "Splits(0).Columns(8).FooterStyle:id=80,.parent=45"
      _StyleDefs(73)  =   "Splits(0).Columns(8).EditorStyle:id=81,.parent=47"
      _StyleDefs(74)  =   "Splits(0).Columns(9).Style:id=86,.parent=43,.alignment=1,.valignment=3"
      _StyleDefs(75)  =   ":id=86,.locked=-1"
      _StyleDefs(76)  =   "Splits(0).Columns(9).HeadingStyle:id=83,.parent=44,.alignment=1"
      _StyleDefs(77)  =   "Splits(0).Columns(9).FooterStyle:id=84,.parent=45"
      _StyleDefs(78)  =   "Splits(0).Columns(9).EditorStyle:id=85,.parent=47"
      _StyleDefs(79)  =   "Splits(0).Columns(10).Style:id=90,.parent=43,.alignment=1,.valignment=3"
      _StyleDefs(80)  =   ":id=90,.locked=-1"
      _StyleDefs(81)  =   "Splits(0).Columns(10).HeadingStyle:id=87,.parent=44,.alignment=1"
      _StyleDefs(82)  =   "Splits(0).Columns(10).FooterStyle:id=88,.parent=45"
      _StyleDefs(83)  =   "Splits(0).Columns(10).EditorStyle:id=89,.parent=47"
      _StyleDefs(84)  =   "Named:id=33:Normal"
      _StyleDefs(85)  =   ":id=33,.parent=0"
      _StyleDefs(86)  =   "Named:id=34:Heading"
      _StyleDefs(87)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(88)  =   ":id=34,.wraptext=-1"
      _StyleDefs(89)  =   "Named:id=35:Footing"
      _StyleDefs(90)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(91)  =   "Named:id=36:Selected"
      _StyleDefs(92)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(93)  =   "Named:id=37:Caption"
      _StyleDefs(94)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(95)  =   "Named:id=38:HighlightRow"
      _StyleDefs(96)  =   ":id=38,.parent=33,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(97)  =   "Named:id=39:EvenRow"
      _StyleDefs(98)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(99)  =   "Named:id=40:OddRow"
      _StyleDefs(100) =   ":id=40,.parent=33"
      _StyleDefs(101) =   "Named:id=41:RecordSelector"
      _StyleDefs(102) =   ":id=41,.parent=34"
      _StyleDefs(103) =   "Named:id=42:FilterBar"
      _StyleDefs(104) =   ":id=42,.parent=33"
   End
   Begin VB.Frame Frame1 
      Height          =   705
      Left            =   8400
      TabIndex        =   7
      Top             =   4320
      Width           =   2865
      Begin VB.TextBox TXTTOTAL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         DataSource      =   "DATATOTAL"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1395
         Locked          =   -1  'True
         TabIndex        =   9
         ToolTipText     =   "Importe Total en Dolares"
         Top             =   270
         Width           =   1290
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Total Soles:"
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Top             =   315
         Width           =   870
      End
   End
   Begin VB.Data DataDOCUM 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   180
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4095
      Visible         =   0   'False
      Width           =   1068
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   10368
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   10
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Frmcomxmes.frx":004D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Frmcomxmes.frx":0367
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Frmcomxmes.frx":0681
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Frmcomxmes.frx":078B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Frmcomxmes.frx":0AA5
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Frmcomxmes.frx":0DBF
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Frmcomxmes.frx":0EC9
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Frmcomxmes.frx":11E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Frmcomxmes.frx":14FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Frmcomxmes.frx":1817
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label LbBusca 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   585
      TabIndex        =   17
      Top             =   4470
      Width           =   45
   End
   Begin VB.Label LbDel 
      Caption         =   "Del :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   144
      TabIndex        =   16
      Top             =   4704
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.Label LbAl 
      Caption         =   "Al :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   2064
      TabIndex        =   15
      Top             =   4704
      Visible         =   0   'False
      Width           =   288
   End
   Begin VB.Menu MNUPRI 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnufiltro 
         Caption         =   "Filtrar Por:"
      End
      Begin VB.Menu mnufiltroex 
         Caption         =   "Filtrar excluyendo:"
      End
      Begin VB.Menu mnufiltroavan 
         Caption         =   "Filtro Avanzado:"
      End
      Begin VB.Menu mnuordasc 
         Caption         =   "Ord. Ascendente:"
      End
      Begin VB.Menu mnuorddesc 
         Caption         =   "Ord. Descendente:"
      End
      Begin VB.Menu mnu_todos 
         Caption         =   "Todos:"
      End
   End
End
Attribute VB_Name = "frmcomxmes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TbDetTmp        As DAO.Recordset
Dim RSCTA_DCTO      As DAO.Recordset
Dim RSCTA_MVTO      As DAO.Recordset
Dim TOTAL           As Double
Dim dbcompras       As DAO.Database
Dim tbcompras       As DAO.Recordset
Dim dbtempcob       As DAO.Database
Dim tbtempcompras   As DAO.Recordset
Dim dbtipodoc       As DAO.Database
Dim tbtipodoc       As DAO.Recordset
Dim dbctaspag       As DAO.Database
Dim tbctaspag       As DAO.Recordset
Dim tbtmpcomprasdet As DAO.Recordset
Dim tbbancos        As DAO.Recordset
Dim tbbf5pla        As DAO.Recordset
Dim SQL             As String
Dim SQL0            As String
Dim COLUMNA         As String
Dim correla         As Integer
Dim DBDB_TABLA      As DAO.Database
Dim TBDB_TABLA      As DAO.Recordset
Dim Codigo          As String
Dim MOV             As String
Dim TBBF4TCO        As DAO.Recordset
Dim tbpagdcto       As DAO.Recordset
Dim tipodoc         As String
Dim descrip         As String
Dim SALDO           As Double
Dim pagado          As Double
Dim txttotaldol     As Double
Dim sw              As Boolean
Dim Tabla           As String
Dim rs              As DAO.Recordset
Dim DATO            As String
Dim dato1           As Double
Dim dato2           As Double
Dim dat1            As DAO.Recordset

Private Sub CALCULA1(sql1 As String)
On Error GoTo DEFINE
        
        Set dbtempcob = OpenDatabase(wrutatemp & "\TempCob.mdb")
        Set tbcabtmp = dbtempcob.OpenRecordset(sql1)
        
        TXTTOTAL.Text = "" & tbcabtmp!SUMASOLES
        TXTTOTAL = Format(TXTTOTAL.Text, "###,###,##0.00")
        If TXTTOTAL.Text = "" Then
          TXTTOTAL = Format(0#, "###,###,##0.00")
        End If
        
        TXTTOTAL.Refresh
        
DEFINE: Exit Sub
End Sub

Private Sub Docum_COMPRAS()
Dim nlon    As Integer
     
    Me.MousePointer = 11
    Set tbcompras = dbcompras.OpenRecordset("regisdoc")
    Set tbtempcompras = dbtempcob.OpenRecordset("comprasxmes")
    Set tbtipodoc = dbtipodoc.OpenRecordset("documentos")
    Set tbpagdcto = dbctaspag.OpenRecordset("pag_dcto")
    dbtempcob.Execute ("Delete * From comprasxmes")
    
    tbcompras.Index = "IDMESNUM"
    tbtipodoc.Index = "IDCODDOC"
    tbpagdcto.Index = "nro_corre"
    
     
    tbcompras.Seek ">=", TxtDesde
    If Not tbcompras.NoMatch Then
      Do While tbcompras!f4mesmov >= TxtDesde.Text
         tipodoc = tbcompras!f4tipdoc
         tbtipodoc.Seek "=", tipodoc
         If Not tbtipodoc.NoMatch Then
             descrip = tbtipodoc!f2abrev
         End If
         If opsoles.Value = True Then
            If tbcompras!f4moneda = "S" Then
               TOTAL = Format(tbcompras!F4TOTAL, "###,###,##0.00")
            Else
                TOTAL = Format((tbcompras!F4TOTAL * tbcompras!F4TIPCAM), "###,###,##0.00")
            End If
          Else
             If tbcompras!f4moneda = "D" Then
               TOTAL = Format(tbcompras!F4TOTAL, "###,###,##0.00")
            Else
                TOTAL = Format((tbcompras!F4TOTAL / tbcompras!F4TIPCAM), "###,###,##0.00")
            End If
          End If
         correla = tbcompras!f4correla
         If correla <> 0 Then
            tbpagdcto.Seek "=", correla
            If Not tbpagdcto.NoMatch Then
               If opsoles.Value = True Then
                    If tbpagdcto!moneda = "S" Then
                        SALDO = Format(tbpagdcto!SALDO, "###,###,##0.00")
                     Else
                        SALDO = Format((tbpagdcto!SALDO * tbpagdcto!tcambio), "###,###,##0.00")
                     End If
                Else
                      If tbpagdcto!moneda = "D" Then
                        SALDO = Format(tbpagdcto!SALDO, "###,###,##0.00")
                     Else
                        SALDO = Format((tbpagdcto!SALDO / tbpagdcto!tcambio), "###,###,##0.00")
                     End If
                End If
                pagado = TOTAL - SALDO
'            Else
'                saldo = 0#
'                pagado = TOTAL - saldo
            End If
         Else
            SALDO = 0#
            pagado = TOTAL
         End If
         
         tbtempcompras.AddNew
                tbtempcompras!f4doc = "" & descrip & tbcompras!f4serdoc & "/" & tbcompras!f4numdoc
                tbtempcompras!f4nomprv = tbcompras!f4nomprv & ""
                tbtempcompras!f4fecha = tbcompras!f4fecha
                tbtempcompras!f4moneda = tbcompras!f4moneda & ""
                tbtempcompras!F4TOTAL = TOTAL
                tbtempcompras!f4pagado = pagado
                tbtempcompras!f4saldo = SALDO
                tbtempcompras!F4REFERE = tbcompras!F4REFERE
                tbtempcompras!F4TIPCAM = tbcompras!F4TIPCAM
                tbtempcompras!f4correla = correla
                tbtempcompras!f4mov = tbcompras!f4mesmov & tbcompras!f4nummov
                tbtempcompras!mes = TxtDesde.Text
                If opsoles.Value = True Then
                   tbtempcompras!tipoopc = "S/."
                Else
                    tbtempcompras!tipoopc = "US$."
                End If
         tbtempcompras.Update
          
         tbcompras.MoveNext
         
         If tbcompras.EOF Then Exit Do
    Loop
 Else
        tbcompras.MoveLast
        Rem NSE mes = tbcompras!f4mesmov
        MsgBox "Mes no Registrado " & "  -  " & " El último mes registrado es: " & tbcompras!f4mesmov
        TXTTOTAL = Format(0#, "###,###,##0.00")
        txttotaldol = Format(0#, "###,###,##0.00")
        TxtDesde.Text = ""
        TxtDesde.SetFocus
 End If
    
 Me.MousePointer = 1
        
            tbcompras.Close
            tbtempcompras.Close
            tbtipodoc.Close
            tbpagdcto.Close
          sw = 0
    End Sub
Sub Docum_Detalle(cod As String, MOVI As String)
Dim nlon As Integer

  Me.MousePointer = 11
    Set tbtempcompras = dbtempcob.OpenRecordset("comprasxmes")
    
    dbtempcob.Execute ("DELETE * FROM  comprasxmesdet")
    
    Set tbctaspag = dbctaspag.OpenRecordset("pag_mvto")
    Set tbtmpcomprasdet = dbtempcob.OpenRecordset("comprasxmesdet")
    Set tbbf5pla = DBDB_TABLA.OpenRecordset("bf5pla")
    Set tbbancos = DBDB_TABLA.OpenRecordset("bancos")
    Set tbpagdcto = dbctaspag.OpenRecordset("pag_dcto")
    tbctaspag.Index = "idcorrdcto "
    
    tbbancos.Index = "IDCODIGO"
    tbbf5pla.Index = "IDCODIGO"
    tbpagdcto.Index = "nro_corre"
    If cod <> 0 Then
        tbctaspag.Seek "=", cod
        If Not tbctaspag.NoMatch Then
            tbctaspag.MoveFirst
            Do While Not tbctaspag.EOF
                If tbctaspag!CORR_DCTO = cod Then
                         tbpagdcto.Seek "=", tbctaspag!corr_comp
                         If Not tbpagdcto.NoMatch Then
                             tbbf5pla.Seek "=", tbpagdcto!ctabanc
                             If Not tbbf5pla.NoMatch Then
                                tbbancos.Seek "=", tbbf5pla!codban
                                If Not tbbancos.NoMatch Then
                                   tbtmpcomprasdet.AddNew
                                       tbtmpcomprasdet!FECHA = tbctaspag!fch_mvto
                                       tbtmpcomprasdet!imputado = tbctaspag!imputado
                                       tbtmpcomprasdet!n_cuenta = tbpagdcto!nro_comp
                                       tbtmpcomprasdet!moneda = tbpagdcto!moneda
                                       tbtmpcomprasdet!banco = tbbancos!banco
                                    tbtmpcomprasdet.Update
                                 End If
                                Else
                                    Rem NSE MsgBox "Codigo no encontrado"
                              End If
                          Else
                                Rem NSE MsgBox "Codigo no encontrado"
                                Me.MousePointer = 1
                                Exit Sub
                           End If
                 End If
                tbctaspag.MoveNext
            Loop
            DataMvto.RecordSource = "select * from comprasxmesdet"
            DataMvto.Refresh
            If DataMvto.Recordset.RecordCount > 0 Then
                PnlMovPen.Visible = True
                Grd_Mvto1.Visible = True
                Grd_Mvto.Visible = False
                Grd_Mvto1.SetFocus
            End If
        Else
            Rem NSE MsgBox "Codigo no encontrado"
            Me.MousePointer = 1
            Exit Sub
        End If
    Else
        Tabla = "BF3MOV" & TxtDesde.Text
        Set TBDB_TABLA = DBDB_TABLA.OpenRecordset(Tabla)
        TBDB_TABLA.Index = "BF3MOV" & TxtDesde.Text
        TBDB_TABLA.Seek "=", cod, MOVI
        If Not TBDB_TABLA.NoMatch Then
            tbbf5pla.Seek "=", TBDB_TABLA!codcta
            If Not tbbf5pla.NoMatch Then
                tbbancos.Seek "=", tbbf5pla!codban
                If Not tbbancos.NoMatch Then
                    tbtmpcomprasdet.AddNew
                    tbtmpcomprasdet!banco = tbbancos!banco
                    tbtmpcomprasdet!n_cuenta = tbbf5pla!numcta
                    tbtmpcomprasdet!FECHA = TBDB_TABLA!fecdis
                    tbtmpcomprasdet!nummov = TBDB_TABLA!nummov
                    tbtmpcomprasdet.Update
                    
                    DataMvto.RecordSource = "select * from comprasxmesdet"
                    DataMvto.Refresh
                    If DataMvto.Recordset.RecordCount > 0 Then
                       PnlMovPen.Visible = True
                       Grd_Mvto1.Visible = False
                       Grd_Mvto.Visible = Visible
                       Grd_Mvto.SetFocus
                    End If
                End If
             End If
          Else
                Rem NSE MsgBox "Codigo no encontrado"
                 Me.MousePointer = 1
                Exit Sub
               
          End If
            
            
     End If
      
    Me.MousePointer = 1
End Sub '********************** FIN DE DOCUM_DETALLE

Private Sub cmdsalir_Click()
PnlMovPen.Visible = False

End Sub

Private Sub Form_Activate()
       
    Grd_Docum.OddRowStyle.BackColor = &HC0FFFF
    Grd_Docum.EvenRowStyle.BackColor = &HFFFFFF
    Grd_Docum.HighlightRowStyle.BackColor = vbActiveTitleBar
    Grd_Docum.HighlightRowStyle.ForeColor = vbWhite
    Grd_Docum.AlternatingRowStyle = True
    
    Grd_Mvto.OddRowStyle.BackColor = &HC0FFFF
    Grd_Mvto.EvenRowStyle.BackColor = &HFFFFFF
    Grd_Mvto.HighlightRowStyle.BackColor = vbActiveTitleBar
    Grd_Mvto.HighlightRowStyle.ForeColor = vbWhite
    Grd_Mvto.AlternatingRowStyle = True
    
    
    Grd_Mvto1.OddRowStyle.BackColor = &HC0FFFF
    Grd_Mvto1.EvenRowStyle.BackColor = &HFFFFFF
    Grd_Mvto1.HighlightRowStyle.BackColor = vbActiveTitleBar
    Grd_Mvto1.HighlightRowStyle.ForeColor = vbWhite
    Grd_Mvto1.AlternatingRowStyle = True
        
End Sub

Private Sub Form_Click()
    PnlMovPen.Visible = False
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 27 Then
        Call Form_Click
        PnlMovPen.Visible = False
    End If
    
End Sub
Private Sub Form_Load()

    Set dbcompras = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set dbtempcob = OpenDatabase(wrutatemp & "\TempCob.mdb")
    Set dbtipodoc = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set dbctaspag = OpenDatabase(wrutabancos & "\ctaspag.mdb")
    Set DBDB_TABLA = OpenDatabase(wrutabancos & "\db_tabla.mdb")
    
    TXTTOTAL.Text = Format(TXTTOTAL, "###,###,##0.00")
    
    
    TxtDesde.Text = mes

    DataDOCUM.DatabaseName = (wrutatemp & "\tempcob.mdb")
    DataMvto.DatabaseName = (wrutatemp & "\tempcob.mdb")
    
End Sub
Private Sub Grd_Docum_Dblclick()
   Codigo = Trim(Grd_Docum.Columns(6))
   MOV = Trim(Grd_Docum.Columns(7))
   Call Docum_Detalle(Codigo, MOV)
   
   
End Sub
Private Sub Grd_Docum_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Call Form_Click
    End If
End Sub
Private Sub Grd_Docum_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If txtbusca.Visible = True Then txtbusca.Visible = False
    If LbBusca.Visible = True Then LbBusca.Visible = False
    If LbDel.Visible = True Then LbDel.Visible = False
    If LbAl.Visible = True Then LbAl.Visible = False
    If MkEdIni.Visible = True Then MkEdIni.Visible = False
    If MkEdFin.Visible = True Then MkEdFin.Visible = False
        
    mnufiltro.Caption = "Filtrar (" + Grd_Docum.Columns(Grd_Docum.col).Text + ")"
    mnufiltroex.Caption = "Filtrar excluyendo: (" + Grd_Docum.Columns(Grd_Docum.col).Text + ")"
    If Button = 2 Then
        PopupMenu mnupri
    End If

End Sub

Private Sub Grd_Mvto_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        PnlMovPen.Visible = False
    End If

End Sub
Private Sub Grd_Mvto_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 Then
        Call Form_Click
    End If

End Sub

Private Sub Grd_Mvto1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 27 Then
        PnlMovPen.Visible = False
    End If
End Sub

Private Sub Grd_Mvto1_KeyPress(KeyAscii As Integer)
 If KeyAscii = 27 Then
        Call Form_Click
    End If

End Sub

Private Sub mnu_todos_Click()
 DataDOCUM.RecordSource = "Select * from COMPRASXMES "
 DataDOCUM.Refresh
 
 SQL = "SELECT SUM(f4total) AS SUMASOLES FROM comprasxmes"
 Call CALCULA1(SQL)
 
End Sub

Private Sub mnufiltro_Click()
Dim SQL     As String
Dim sql1    As String
     
    '******* SETEAMOS *****
     Set rs = DataDOCUM.Recordset
      Select Case DataDOCUM.Recordset.Fields(Grd_Docum.Columns(Grd_Docum.col).DataField).Type
             Case 10
                  rs.Filter = "[" + Grd_Docum.Columns(Grd_Docum.col).DataField + "]" + " = '" + Trim("" & Grd_Docum.Columns(Grd_Docum.col).Text) + "'"
                  COLUMNA = "comprasxmes." + Trim(Grd_Docum.Columns(Grd_Docum.col).DataField)
                  DATO = Trim(Grd_Docum.Columns(Grd_Docum.col).Text)
                  
                  SQL = "SELECT DISTINCTROW Sum(f4total) AS SUMASOLES FROM comprasxmes WHERE  " + COLUMNA + "  = '" & DATO & "'  "
                  
             Case 4
                  rs.Filter = "[" + Grd_Docum.Columns(Grd_Docum.col).DataField + "]" + " = " + Grd_Docum.Columns(Grd_Docum.col).Text
                  COLUMNA = "comprasxmes." + Trim(Grd_Docum.Columns(Grd_Docum.col).DataField)
                  DATO = Trim(Grd_Docum.Columns(Grd_Docum.col).Text)
                  SQL = "SELECT DISTINCTROW Sum(f4total) AS SUMASOLES FROM comprasxmes WHERE  " + COLUMNA + "  = '" & DATO & "'  "
             Case 8
                  If IsDate(Grd_Docum.Columns(Grd_Docum.col).Text) Then
                     rs.Filter = "[" + Grd_Docum.Columns(Grd_Docum.col).DataField + "]" + "=#" + Format(Grd_Docum.Columns(Grd_Docum.col).Text, "mm/dd/yyyy") + "#"
                     COLUMNA = "VENTAXMES." + Trim(Grd_Docum.Columns(Grd_Docum.col).DataField)
                     DATO = "#" + Format(Grd_Docum.Columns(Grd_Docum.col).Text, "mm/dd/yyyy") + "#"
                     SQL = "SELECT DISTINCTROW Sum(f4total) AS SumaSOLES FROM comprasxmes WHERE  " + COLUMNA + "  = '" & DATO & "'"
                  Else
                     MsgBox "Ingrese una Fecha Valida..!", 32, "Advertencia"
                     Exit Sub
                  End If
      End Select
      Set DataDOCUM.Recordset = rs.OpenRecordset(rs.Type)
      Call CALCULA1(SQL)
      sw = 1
      Set rs = Nothing
End Sub

Private Sub mnufiltroavan_Click()
 
 Select Case Grd_Docum.col   'Posicion de columna
                
       Case 0 'Nro. Comprobante
            txtbusca = ""
            txtbusca.Visible = True
            LbBusca.Visible = True
            LbBusca.Caption = "Documento: "
            txtbusca.SetFocus

       Case 1 'Nro. Comprobante
            txtbusca = ""
            txtbusca.Visible = True
            LbBusca.Visible = True
            LbBusca.Caption = "Proveedor : "
            txtbusca.SetFocus
        
        Case 2 'Fecha
            
            MkEdIni.Text = "  /  /    "
            MkEdFin.Text = "  /  /    "
            LbDel.Visible = True
            LbAl.Visible = True
            MkEdIni.Visible = True
            MkEdFin.Visible = True
            MkEdIni.SetFocus
    End Select
End Sub
Private Sub mnufiltroex_Click()
    
    Dim SQL As String
     
     Set rs = DataDOCUM.Recordset
      Select Case DataDOCUM.Recordset.Fields(Grd_Docum.Columns(Grd_Docum.col).DataField).Type
             Case 10
                  rs.Filter = "[" + Grd_Docum.Columns(Grd_Docum.col).DataField + "]" + " <> '" + Trim("" & Grd_Docum.Columns(Grd_Docum.col).Text) + "'"
                  COLUMNA = "comprasxmes." + Trim(Grd_Docum.Columns(Grd_Docum.col).DataField)
                  
                  DATO = Trim(Grd_Docum.Columns(Grd_Docum.col).Text)
                  SQL = "SELECT DISTINCTROW Sum(f4total) AS SUMASOLES FROM comprasxmes HAVING " + COLUMNA + "  <> '" & DATO & "' "
             Case 4
                  rs.Filter = "[" + Grd_Docum.Columns(Grd_Docum.col).DataField + "]" + " <> " + Grd_Docum.Columns(Grd_Docum.col).Text
                  COLUMNA = "comprasxmes." + Trim(Grd_Docum.Columns(Grd_Docum.col).DataField)
                  DATO = Trim(Grd_Docum.Columns(Grd_Docum.col).Text)
                  SQL = "SELECT DISTINCTROW Sum(f4total) AS SUMASOLES FROM comprasxmes HAVING  " + COLUMNA + "  <> '" & DATO & "'"
             Case 8
                  If IsDate(Grd_Docum.Columns(Grd_Docum.col).Text) Then
                     rs.Filter = "[" + Grd_Docum.Columns(Grd_Docum.col).DataField + "]" + "<>#" + Format(Grd_Docum.Columns(Grd_Docum.col).Text, "mm/dd/yyyy") + "#"
                     COLUMNA = "comprasxmes." + Trim(Grd_Docum.Columns(Grd_Docum.col).DataField)
                     DATO = "#" + Format(Grd_Docum.Columns(Grd_Docum.col).Text, "mm/dd/yyyy") + "#"
                     SQL = "SELECT DISTINCTROW Sum(f4total) AS SUMASOLES FROM comprasxmes HAVING  " + COLUMNA + " <> '" & DATO & "'"
                  Else
                     MsgBox "Ingrese una Fecha Valida..!", 32, "Advertencia"
                     Exit Sub
                  End If
      End Select
      Set DataDOCUM.Recordset = rs.OpenRecordset(rs.Type)
     Call CALCULA1(SQL)
     sw = 1
     Set rs = Nothing

End Sub
Private Sub mnuordasc_Click()
    
    Set rs = DataDOCUM.Recordset
    rs.Sort = "[" + Grd_Docum.Columns(Grd_Docum.col).DataField + "] Asc"
    Set DataDOCUM.Recordset = rs.OpenRecordset(rs.Type)
    Set rs = Nothing
    
End Sub
Private Sub mnuorddesc_Click()
    
    Set rs = DataDOCUM.Recordset
    rs.Sort = "[" + Grd_Docum.Columns(Grd_Docum.col).DataField + "] Desc"
    Set DataDOCUM.Recordset = rs.OpenRecordset(rs.Type)
    Set rs = Nothing
    
End Sub

Private Sub opdolares_Click(Value As Integer)
   Label4.Caption = "Total US$"
   Docum_COMPRAS
    DataDOCUM.RecordSource = "Select * from comprasxmes"
    DataDOCUM.Refresh
    If DataDOCUM.Recordset.RecordCount > 0 Then
          SQL = "SELECT SUM(f4total) AS SUMASOLES FROM comprasxmes "
          Call CALCULA1(SQL)
     End If
End Sub

Private Sub opdolares_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        opdolares_Click 1
    End If
    
End Sub

Private Sub opsoles_Click(Value As Integer)
  Label4.Caption = "Total S/."
  Docum_COMPRAS
    DataDOCUM.RecordSource = "Select * from comprasxmes"
    DataDOCUM.Refresh
    If DataDOCUM.Recordset.RecordCount > 0 Then
          SQL = "SELECT SUM(f4total) AS SUMASOLES FROM comprasxmes "
          Call CALCULA1(SQL)
     End If
End Sub

Private Sub opsoles_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        opsoles_Click 1
    End If
End Sub

Private Sub tblbar_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Index
    Case 1:
         reportetotal.DataFiles(0) = wrutatemp & "\tempcob.mdb"
         reportetotal.ReportFileName = wrutatemp & "\comprasxmes.rpt"
        reportetotal.Action = 1
    Case 2:
         Unload Me
    End Select
End Sub

Private Sub txtdesde_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        TxtDesde.Text = Format(TxtDesde.Text, "00")
        If IsNumeric(TxtDesde) Then
            If (TxtDesde) >= 1 And (TxtDesde) <= 12 Then
                    opsoles.SetFocus
            Else
                TxtDesde.Text = ""
                TxtDesde.SetFocus
            End If
        Else
            TxtDesde.Text = ""
            TxtDesde.SetFocus
        End If
    End If
   
End Sub
Private Sub txtdesde_LostFocus()
    TxtDesde.Text = Format(TxtDesde.Text, "00")
    If IsNumeric(TxtDesde) Then
            If (TxtDesde) >= 1 And (TxtDesde) <= 12 Then
                   opsoles.SetFocus
            Else
                TxtDesde.Text = ""
                TxtDesde.SetFocus
            End If
    Else
       TxtDesde.Text = ""
       TxtDesde.SetFocus
    End If
End Sub
Private Sub txtplazfin_GotFocus()
    
    txtplazfin.SelStart = 0
    txtplazfin.MaxLength = Len(txtplazfin.Text)

End Sub
Private Sub txtplazfin_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
      If Len(txtplazfin.Text) = 0 Then MsgBox "Dato en blanco no valido..!", 32, "Advertencia": txtplazfin.SetFocus: Exit Sub
          Set rs = DataDOCUM.Recordset
          dato1 = Str(CDbl(txtplazini.Text))
          dato2 = Str(CDbl(txtplazfin.Text))
    
          rs.Filter = "[" + Grd_Docum.Columns(Grd_Docum.col).DataField + "]" + " >= " + dato1 + "  And  " & _
                      "[" + Grd_Docum.Columns(Grd_Docum.col).DataField + "]" + " <= " + dato2 + " "
    
          If rs.EOF Then txtplazini.SetFocus: Exit Sub Else txtplazini.Visible = False: txtplazfin.Visible = False: LbDel.Visible = False: LbAl.Visible = False
    
    
          COLUMNA = "comprasxmes." + Trim(Grd_Docum.Columns(Grd_Docum.col).DataField)
    
          SQL = "SELECT DISTINCTROW *  FROM comprasxmes WHERE  " + COLUMNA + " >= " + dato1 + "  And  " + _
                COLUMNA + " <= " + dato2 + ""
    
          Set DataDOCUM.Recordset = rs.OpenRecordset(rs.Type)
    
          Set rs = Nothing
          'DataDOCUM.SetFocus

    End If
    
End Sub
Private Sub txtplazini_GotFocus()
    txtplazini.SelStart = 0
    txtplazini.MaxLength = Len(txtplazini.Text)
End Sub
Private Sub txtplazini_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
      If Len(txtplazini.Text) = 0 Then MsgBox "Dato en blanco no valido..!", 32, "Advertencia": txtplazini.SetFocus: Exit Sub
      txtplazfin.SetFocus
    End If
End Sub
Private Sub txtbusca_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
       If Len(txtbusca) = 0 Then txtbusca.SetFocus: Exit Sub
          If InStr(txtbusca, "*") = Len(Trim(txtbusca)) Then
             DATO = Left(txtbusca, Len(Trim(txtbusca)) - 1)
          Else
              DATO = txtbusca
          End If
                      Set dat1 = DataDOCUM.Recordset
                      dat1.Filter = "[" + Grd_Docum.Columns(Grd_Docum.col).DataField + "]" + " Like  '" + DATO + "*'"
                      If dat1.EOF Then txtbusca.SetFocus: Exit Sub Else txtbusca.Visible = False: LbBusca.Visible = False
                      
                        COLUMNA = "comprasxmes." + Trim(Grd_Docum.Columns(Grd_Docum.col).DataField)
                        
                        SQL = "SELECT DISTINCTROW Sum(f4total) AS SUMASOLES FROM comprasxmes WHERE " + COLUMNA + " LIKE '" + DATO + "*'  "
                    
                    Set DataDOCUM.Recordset = dat1.OpenRecordset(dat1.Type)
                    
                    Call CALCULA1(SQL)
                    
                    Set dat1 = Nothing
                    
    
    End If

End Sub

Private Sub MkEdIni_GotFocus()
    
    MkEdIni.SelStart = 0
    MkEdIni.SelLength = MkEdIni.MaxLength

End Sub
Private Sub MkEdFin_GotFocus()
    
    MkEdFin.SelStart = 0
    MkEdFin.SelLength = MkEdFin.MaxLength

End Sub
Private Sub MkEdFin_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
      If Not IsDate(MkEdFin.Text) Then MsgBox "Ingrese una Fecha Valida..!", 32, "Advertencia": MkEdFin.SetFocus: Exit Sub
            
          Set rs = DataDOCUM.Recordset
          dato1 = MkEdIni.Text
          dato2 = MkEdFin.Text
                      
          rs.Filter = "[" + Grd_Docum.Columns(Grd_Docum.col).DataField + "]" + " >= #" + Format(dato1, "mm/dd/yyyy") + "#  And  " & _
                      "[" + Grd_Docum.Columns(Grd_Docum.col).DataField + "]" + " <= #" + Format(dato2, "mm/dd/yyyy") + "# "
                      
          If rs.EOF Then MkEdIni.SetFocus: Exit Sub Else MkEdIni.Visible = False: MkEdFin.Visible = False: LbDel.Visible = False: LbAl.Visible = False
                                   
          COLUMNA = "comprasxmes." + Trim(Grd_Docum.Columns(Grd_Docum.col).DataField)
          SQL = "SELECT DISTINCTROW  * FROM comprasxmes WHERE  " + COLUMNA + " >= #" + Format(dato1, "mm/dd/yyyy") + "#  And  " + COLUMNA + " <= #" + Format(dato2, "mm/dd/yyyy") + "# "
          Set DataDOCUM.Recordset = rs.OpenRecordset(rs.Type)
          DataDOCUM.RecordSource = SQL
          DataDOCUM.Refresh
          
          
          SQL = "SELECT DISTINCTROW Sum(f4total) AS SumaSOLES FROM comprasxmes WHERE " + COLUMNA + " >= #" + Format(dato1, "mm/dd/yyyy") + "#  And  " + COLUMNA + " <= #" + Format(dato2, "mm/dd/yyyy") + "#   "
          Call CALCULA1(SQL)
          
          Set rs = Nothing
          
 End If

End Sub
Private Sub MkEdIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not IsDate(MkEdIni.Text) Then MsgBox "Ingrese una Fecha Valida..!", 32, "Advertencia": MkEdIni.SetFocus: Exit Sub
        MkEdFin.SetFocus
    End If
End Sub
