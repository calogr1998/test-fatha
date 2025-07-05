VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form cmovi 
   Appearance      =   0  'Flat
   Caption         =   "Movimientos de Gastos"
   ClientHeight    =   4590
   ClientLeft      =   1110
   ClientTop       =   2295
   ClientWidth     =   9060
   ControlBox      =   0   'False
   FillColor       =   &H00404080&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00004000&
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4590
   ScaleWidth      =   9060
   Begin VB.Data Data 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   7728
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3024
      Visible         =   0   'False
      Width           =   1056
   End
   Begin TrueOleDBGrid70.TDBGrid Grid1 
      Bindings        =   "CMOVI.frx":0000
      Height          =   3150
      Left            =   90
      TabIndex        =   9
      Top             =   810
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   5556
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Nº Mov"
      Columns(0).DataField=   "MOVIMIENTO"
      Columns(0).DataWidth=   255
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Nº Doc"
      Columns(1).DataField=   "NUMDOC"
      Columns(1).DataWidth=   255
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Fecha"
      Columns(2).DataField=   "FECHA"
      Columns(2).DataWidth=   19
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Neto"
      Columns(3).DataField=   "NETO"
      Columns(3).DataWidth=   23
      Columns(3).NumberFormat=   "Standard"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Igv"
      Columns(4).DataField=   "IGV"
      Columns(4).DataWidth=   23
      Columns(4).NumberFormat=   "Standard"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Total"
      Columns(5).DataField=   "TOTAL"
      Columns(5).DataWidth=   23
      Columns(5).NumberFormat=   "Standard"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   6
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).ScrollBars=   2
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=6"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2223"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2143"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=66080"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=2831"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2752"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=66080"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=2328"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2249"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=66082"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=2699"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=2619"
      Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=66082"
      Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(25)=   "Column(4).Width=2302"
      Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=2223"
      Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=66082"
      Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(31)=   "Column(5).Width=2196"
      Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=2117"
      Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=66082"
      Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
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
      _StyleDefs(5)   =   ":id=0,.fontname=Arial"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Arial"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&H8000000F&"
      _StyleDefs(11)  =   ":id=2,.fgcolor=&H80000012&,.bold=0,.fontsize=825,.italic=0,.underline=0"
      _StyleDefs(12)  =   ":id=2,.strikethrough=0,.charset=0"
      _StyleDefs(13)  =   ":id=2,.fontname=Arial"
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
      _StyleDefs(47)  =   "Splits(0).Columns(3).Style:id=62,.parent=43,.alignment=1,.valignment=1"
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
      _StyleDefs(59)  =   "Named:id=33:Normal"
      _StyleDefs(60)  =   ":id=33,.parent=0"
      _StyleDefs(61)  =   "Named:id=34:Heading"
      _StyleDefs(62)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(63)  =   ":id=34,.wraptext=-1"
      _StyleDefs(64)  =   "Named:id=35:Footing"
      _StyleDefs(65)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(66)  =   "Named:id=36:Selected"
      _StyleDefs(67)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(68)  =   "Named:id=37:Caption"
      _StyleDefs(69)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(70)  =   "Named:id=38:HighlightRow"
      _StyleDefs(71)  =   ":id=38,.parent=33,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(72)  =   "Named:id=39:EvenRow"
      _StyleDefs(73)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(74)  =   "Named:id=40:OddRow"
      _StyleDefs(75)  =   ":id=40,.parent=33"
      _StyleDefs(76)  =   "Named:id=41:RecordSelector"
      _StyleDefs(77)  =   ":id=41,.parent=34"
      _StyleDefs(78)  =   "Named:id=42:FilterBar"
      _StyleDefs(79)  =   ":id=42,.parent=33"
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   450
      Left            =   90
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      _Version        =   65536
      _ExtentX        =   15690
      _ExtentY        =   794
      _StockProps     =   15
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSCommand BtnExitO 
         Height          =   336
         Left            =   528
         TabIndex        =   1
         ToolTipText     =   "Salir"
         Top             =   48
         Width           =   372
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   582
         _StockProps     =   78
         ForeColor       =   -2147483640
         Picture         =   "CMOVI.frx":0013
      End
      Begin Threed.SSCommand BtnPrint 
         Height          =   336
         Left            =   48
         TabIndex        =   2
         ToolTipText     =   "Imprimir"
         Top             =   48
         Width           =   372
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   582
         _StockProps     =   78
         ForeColor       =   -2147483640
         Picture         =   "CMOVI.frx":016D
      End
      Begin Crystal.CrystalReport rep_mov 
         Left            =   8064
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
      End
   End
   Begin Threed.SSPanel Panel3D1 
      Height          =   345
      Left            =   90
      TabIndex        =   3
      Top             =   435
      Width           =   8895
      _Version        =   65536
      _ExtentX        =   15690
      _ExtentY        =   609
      _StockProps     =   15
      Caption         =   "Descrip. Grupo"
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelOuter      =   0
      BevelInner      =   1
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   480
      Left            =   90
      TabIndex        =   4
      Top             =   4005
      Width           =   8880
      _Version        =   65536
      _ExtentX        =   15663
      _ExtentY        =   847
      _StockProps     =   15
      ForeColor       =   -2147483640
      BackColor       =   -2147483638
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
      Begin Threed.SSPanel TTOTAL 
         Height          =   345
         Left            =   7425
         TabIndex        =   5
         Top             =   90
         Width           =   1170
         _Version        =   65536
         _ExtentX        =   2064
         _ExtentY        =   609
         _StockProps     =   15
         Caption         =   "0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483638
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelOuter      =   0
         BevelInner      =   1
         Alignment       =   4
      End
      Begin Threed.SSPanel TNETO 
         Height          =   345
         Left            =   4950
         TabIndex        =   6
         Top             =   90
         Width           =   1170
         _Version        =   65536
         _ExtentX        =   2064
         _ExtentY        =   609
         _StockProps     =   15
         Caption         =   "0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483638
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelOuter      =   0
         BevelInner      =   1
         Alignment       =   4
      End
      Begin Threed.SSPanel TIGV 
         Height          =   345
         Left            =   6180
         TabIndex        =   7
         Top             =   90
         Width           =   1170
         _Version        =   65536
         _ExtentX        =   2064
         _ExtentY        =   609
         _StockProps     =   15
         Caption         =   "0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483638
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelOuter      =   0
         BevelInner      =   1
         Alignment       =   4
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   285
         Left            =   225
         TabIndex        =   8
         Top             =   90
         Width           =   2580
      End
   End
End
Attribute VB_Name = "cmovi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TIPMON      As String
Dim F3NETO      As Double, F3IGV As Double, F3TOTAL As Double
Dim TF3NETO     As Double, TF3IGV As Double, TF3TOTAL As Double
Dim PCODGAS     As String
Dim VALORCAMBIO As Double

Private Sub BtnExitO_Click()
Unload Me
End Sub



Private Sub BtnPrint_Click()
    
    rep_mov.DataFiles(0) = wrutatemp & "\temp_com.mdb"
    rep_mov.ReportFileName = wrutatemp & "\Grupmov.rpt"
    rep_mov.Action = 1

End Sub

Private Sub Form_Activate()

    Grid1.EvenRowStyle.BackColor = &HFFFFFF
    Grid1.OddRowStyle.BackColor = &HC0FFFF
    Grid1.HighlightRowStyle.BackColor = vbActiveTitleBar
    Grid1.HighlightRowStyle.ForeColor = vbWhite
    Grid1.AlternatingRowStyle = True
    
End Sub

Private Sub Form_Load()
On Error GoTo error_bd

    Set dbcompra = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set tbcompra = dbcompra.OpenRecordset("regismov")
    tbcompra.Index = "IGASTO"
    Set TBCOMPraDOC = dbcompra.OpenRecordset("regisDOC")
    TBCOMPraDOC.Index = "IDMESNUM"

    Set dbmovim = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set tbmovim = dbmovim.OpenRecordset("bf9gin")
    tbmovim.Index = "ICodigo"

    Set dbgrupos = OpenDatabase(wrutabancos & "\db_tabla.mdb")
    Set tbgrupos = dbgrupos.OpenRecordset("grupgast")
    tbgrupos.Index = "ICodigo"
    If TIPMON = "S" Then Label1.Caption = " TOTALES SOLES"
    If TIPMON = "D" Then Label1.Caption = " TOTALES DOLARES"

    F3NETO = 0
    F3IGV = 0
    F3TOTAL = 0
    TF3NETO = 0
    TF3IGV = 0
    TF3TOTAL = 0

    PCODGAS = CGAST.Grid1.Text
    CGAST.Grid1.col = 1
    Panel3D1.Caption = CGAST.Grid1.Text + "          MES : " & xmes & TIPMON

    Llena_Grid
    Data.DatabaseName = (wrutatemp & "\temp_com.mdb")
    Data.RecordSource = "select  movimiento,numdoc,fecha, neto, igv, total from tempmovgas"
    Data.Refresh
    
    tbgrupos.Close
    dbgrupos.Close

    tbmovim.Close
    dbmovim.Close

    TBCOMPraDOC.Close
    tbcompra.Close
    dbcompra.Close

    Exit Sub

error_bd:
    Resume Next

End Sub

Private Sub Llena_Grid()
Dim dbtemp As DAO.Database
Dim tempmov As DAO.Recordset
Dim CAS As Integer

    Set dbtemp = OpenDatabase(wrutatemp & "\temp_com.mdb")
    Set tempmov = dbtemp.OpenRecordset("tempmovgas")
    dbtemp.Execute ("delete from tempmovgas")

'    Grid1.Row = 0
    tbcompra.Index = "IGASTO"
    tbcompra.Seek "=", PCODGAS
    If Not tbcompra.NoMatch Then
        Do While PCODGAS = tbcompra.Fields("f3gasto")
            If tbcompra.Fields("f4mesmov") = xmes Then
                TBCOMPraDOC.Index = "IDMESNUM"
                TBCOMPraDOC.Seek "=", tbcompra.Fields("f4mesmov"), tbcompra.Fields("f4nummov")
                If Not TBCOMPraDOC.NoMatch Then
                    If TIPMON = "S" And TBCOMPraDOC.Fields("f4moneda") = "S" Then CAS = 1
                    If TIPMON = "D" And TBCOMPraDOC.Fields("f4moneda") = "D" Then CAS = 1
                    If TIPMON = "D" And TBCOMPraDOC.Fields("f4moneda") = "S" Then CAS = 2
                    If TIPMON = "S" And TBCOMPraDOC.Fields("f4moneda") = "D" Then CAS = 3
                    Select Case CAS
                        Case 1:
                            F3NETO = tbcompra.Fields("f3importe")
                        Case 2:
                            VALORCAMBIO = TBCOMPraDOC.Fields("f4TIPCAM")
                            F3NETO = Format(tbcompra.Fields("f3importe") / VALORCAMBIO, "#0.00")
                        Case 3:
                            VALORCAMBIO = TBCOMPraDOC.Fields("f4TIPCAM")
                            F3NETO = Format(tbcompra.Fields("f3importe") * VALORCAMBIO, "#0.00")
                    End Select
                    If tbcompra.Fields("f3AFECTO") = "*" Then
                        F3IGV = F3NETO * 0.19
                    End If
                    F3TOTAL = F3NETO + F3IGV
                    TF3NETO = TF3NETO + F3NETO
                    TF3IGV = TF3IGV + F3IGV
                    TF3TOTAL = TF3TOTAL + F3TOTAL
                    If Not (F3NETO = 0) Then
                      '  Grid1.Rows = Grid1.Rows + 1
                      '  Grid1.Row = Grid1.Rows - 2
                      '  Grid1.Col = 0: Grid1.Text = "" & tbcompra.Fields("f4NUMMOV")
                      '  Grid1.Col = 1: Grid1.Text = "" & TBCOMPraDOC.Fields("f4NUMDOC")
                      '  Grid1.Col = 2: Grid1.Text = "" & TBCOMPraDOC.Fields("f4FECHA")
                      '  Grid1.Col = 3: Grid1.Text = Format(F3NETO, "##,##0.00")
                      '  Grid1.Col = 4: Grid1.Text = Format(F3IGV, "##,##0.00")
                      '  Grid1.Col = 5: Grid1.Text = Format(F3TOTAL, "##,##0.00")
                        tempmov.AddNew
                        tempmov.Fields("MOVIMIENTO") = "" & tbcompra.Fields("f4NUMMOV")
                        tempmov.Fields("NUMDOC") = "" & TBCOMPraDOC.Fields("f4NUMDOC")
                        tempmov.Fields("FECHA") = "" & TBCOMPraDOC.Fields("f4FECHA")
                        tempmov.Fields("NETO") = Format(F3NETO, "##,##0.00")
                        tempmov.Fields("IGV") = Format(F3IGV, "##,##0.00")
                        tempmov.Fields("TOTAL") = Format(F3TOTAL, "##,##0.00")
                        tempmov.Fields("EMPRESA") = UCase(wnomcia)
                        tempmov.Fields("titulo") = Panel3D1.Caption
                        tempmov.Update
                    End If
                End If
            End If
            tbcompra.MoveNext
            F3NETO = 0
            F3IGV = 0
            F3TOTAL = 0
            If tbcompra.EOF = True Then Exit Do 'SALIR movim
        Loop
    End If
    'Grid1.Col = 0
    'Grid1.Row = Grid1.Rows - 1
    'If Trim(Grid1.Text) = "" And Grid1.Rows <> 2 Then
    '    Grid1.Rows = Grid1.Rows - 1
    'End If
    '---------------------------
    TNETO.Caption = Format(TF3NETO, "##,##0.00")
    TIGV.Caption = Format(TF3IGV, "##,##0.00")
    TTOTAL.Caption = Format(TF3TOTAL, "##,##0.00")

    tempmov.Close
    dbtemp.Close

End Sub

