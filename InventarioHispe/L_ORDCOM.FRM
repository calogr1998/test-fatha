VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmAyudaOrdCom 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista de Ordenes de Compra"
   ClientHeight    =   4935
   ClientLeft      =   1545
   ClientTop       =   2190
   ClientWidth     =   8955
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
   ScaleHeight     =   4935
   ScaleWidth      =   8955
   Begin Threed.SSPanel SSPanel1 
      Height          =   4875
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8880
      _Version        =   65536
      _ExtentX        =   15663
      _ExtentY        =   8599
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
      Begin MSMask.MaskEdBox mkfecha 
         Height          =   285
         Left            =   90
         TabIndex        =   6
         Top             =   4410
         Visible         =   0   'False
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Appearance      =   0  'Flat
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
         Left            =   90
         TabIndex        =   2
         ToolTipText     =   "Simplifique la consulta digitando el Simbolo (*)"
         Top             =   4410
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
         Left            =   180
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1035
         Visible         =   0   'False
         Width           =   1155
      End
      Begin TrueOleDBGrid70.TDBGrid datagrid 
         Bindings        =   "L_ORDCOM.frx":0000
         Height          =   3615
         Left            =   45
         TabIndex        =   1
         Top             =   90
         Width           =   8790
         _ExtentX        =   15505
         _ExtentY        =   6376
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "No. Orden"
         Columns(0).DataField=   "F4NUMORD"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Proveedor"
         Columns(1).DataField=   "F2NOMPROV"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Fecha"
         Columns(2).DataField=   "F4FECEMI"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Cod. Prov"
         Columns(3).DataField=   "F4CODPRV"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Observaciones"
         Columns(4).DataField=   "F4OBSERVA"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Total"
         Columns(5).DataField=   "F4MONTO"
         Columns(5).NumberFormat=   "Standard"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   6
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   2
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   2
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=6"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2355"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2275"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=65808"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=9869"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=9790"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1).AllowSizing=0"
         Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=74000"
         Splits(0)._ColumnProps(13)=   "Column(1).Visible=0"
         Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(15)=   "Column(2).Width=2275"
         Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=2196"
         Splits(0)._ColumnProps(18)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(19)=   "Column(2).AllowSizing=0"
         Splits(0)._ColumnProps(20)=   "Column(2)._ColStyle=74000"
         Splits(0)._ColumnProps(21)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(22)=   "Column(3).Width=1667"
         Splits(0)._ColumnProps(23)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(24)=   "Column(3)._WidthInPix=1588"
         Splits(0)._ColumnProps(25)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(26)=   "Column(3).AllowSizing=0"
         Splits(0)._ColumnProps(27)=   "Column(3)._ColStyle=74000"
         Splits(0)._ColumnProps(28)=   "Column(3).Visible=0"
         Splits(0)._ColumnProps(29)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(30)=   "Column(4).Width=7911"
         Splits(0)._ColumnProps(31)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(32)=   "Column(4)._WidthInPix=7832"
         Splits(0)._ColumnProps(33)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(34)=   "Column(4).AllowSizing=0"
         Splits(0)._ColumnProps(35)=   "Column(4)._ColStyle=74000"
         Splits(0)._ColumnProps(36)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(37)=   "Column(5).Width=2434"
         Splits(0)._ColumnProps(38)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(39)=   "Column(5)._WidthInPix=2355"
         Splits(0)._ColumnProps(40)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(41)=   "Column(5).AllowSizing=0"
         Splits(0)._ColumnProps(42)=   "Column(5)._ColStyle=74002"
         Splits(0)._ColumnProps(43)=   "Column(5).Order=6"
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
         _StyleDefs(35)  =   "Splits(0).Columns(0).Style:id=28,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(36)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=44"
         _StyleDefs(37)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=45"
         _StyleDefs(38)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=47"
         _StyleDefs(39)  =   "Splits(0).Columns(1).Style:id=32,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(40)  =   ":id=32,.locked=-1"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=44"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=45"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=47"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=58,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(45)  =   ":id=58,.locked=-1"
         _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=44"
         _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=45"
         _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=47"
         _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=62,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(50)  =   ":id=62,.locked=-1"
         _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=44"
         _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=45"
         _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=47"
         _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=66,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(55)  =   ":id=66,.locked=-1"
         _StyleDefs(56)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=44"
         _StyleDefs(57)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=45"
         _StyleDefs(58)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=47"
         _StyleDefs(59)  =   "Splits(0).Columns(5).Style:id=70,.parent=43,.alignment=1,.valignment=2"
         _StyleDefs(60)  =   ":id=70,.locked=-1"
         _StyleDefs(61)  =   "Splits(0).Columns(5).HeadingStyle:id=67,.parent=44"
         _StyleDefs(62)  =   "Splits(0).Columns(5).FooterStyle:id=68,.parent=45"
         _StyleDefs(63)  =   "Splits(0).Columns(5).EditorStyle:id=69,.parent=47"
         _StyleDefs(64)  =   "Named:id=33:Normal"
         _StyleDefs(65)  =   ":id=33,.parent=0"
         _StyleDefs(66)  =   "Named:id=34:Heading"
         _StyleDefs(67)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(68)  =   ":id=34,.wraptext=-1"
         _StyleDefs(69)  =   "Named:id=35:Footing"
         _StyleDefs(70)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(71)  =   "Named:id=36:Selected"
         _StyleDefs(72)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(73)  =   "Named:id=37:Caption"
         _StyleDefs(74)  =   ":id=37,.parent=34,.alignment=2,.bold=0,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(75)  =   ":id=37,.strikethrough=0,.charset=0"
         _StyleDefs(76)  =   ":id=37,.fontname=MS Sans Serif"
         _StyleDefs(77)  =   "Named:id=38:HighlightRow"
         _StyleDefs(78)  =   ":id=38,.parent=33,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
         _StyleDefs(79)  =   "Named:id=39:EvenRow"
         _StyleDefs(80)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(81)  =   "Named:id=40:OddRow"
         _StyleDefs(82)  =   ":id=40,.parent=33"
         _StyleDefs(83)  =   "Named:id=41:RecordSelector"
         _StyleDefs(84)  =   ":id=41,.parent=34"
         _StyleDefs(85)  =   "Named:id=42:FilterBar"
         _StyleDefs(86)  =   ":id=42,.parent=33"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "[F3] ---> Filtro Avanzado: Fecha"
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
         TabIndex        =   5
         Top             =   3780
         Width           =   2220
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
         Height          =   375
         Left            =   5985
         TabIndex        =   4
         Top             =   4320
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
         Left            =   90
         TabIndex        =   3
         Top             =   4140
         Visible         =   0   'False
         Width           =   60
      End
   End
   Begin VB.Menu mnupri 
      Caption         =   ""
      Begin VB.Menu mnufiltro 
         Caption         =   "Filtar"
      End
      Begin VB.Menu mnufiltroavanz 
         Caption         =   "Filtro Avanzado:"
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
Attribute VB_Name = "FrmAyudaOrdCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs      As DAO.Recordset
Dim xCodigo As Integer
Dim swcod   As Integer
Dim swnom   As Integer
Dim dbbase  As DAO.Database
Dim dbtempo As DAO.Database
Dim tbtempo As DAO.Recordset
Dim dbempre As DAO.Database
Dim TbProv  As DAO.Recordset


Private Sub DataGrid_DblClick()

    DataGrid_KeyPress 13
    'DataGrid_KeyDown 13, 0

End Sub

Private Sub DataGrid_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 114 Then
    DataGrid.col = 2
    mnufiltroavanz_Click
  End If
  DataGrid_KeyPress KeyCode
  
End Sub

Private Sub DataGrid_KeyPress(KeyAscii As Integer)
 Select Case KeyAscii
  Case 13:
           gcodord = Val(Right(DataGrid.Columns(0) & "", 7))   ' Val(Right(DataAyuda.Recordset.Fields("F4NUMORD") & "", 7))
           gorden_cs = DataGrid.Columns(0) & ""
           Unload Me
  Case 27:
            'gcodord = ""
            'gorden_cs = ""
            Unload Me
End Select
End Sub

Private Sub DataGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
        If DataAyuda.Recordset.RecordCount > 0 Then
            gcodord = Val(Right(DataAyuda.Recordset.Fields("F4NUMORD") & "", 7))
            gtipo = Left(DataAyuda.Recordset.Fields("F4NUMORD") & "", 3)
            'PopupMenu mnu100
        End If
    End If

End Sub

Private Sub datagrid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Label1.Visible = True
End Sub

Private Sub datagrid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If txtope.Visible = True Then txtope.Visible = False
    If lblbusca.Visible = True Then lblbusca.Visible = False
    If mkfecha.Visible = True Then mkfecha.Visible = False
    mnufiltro.Caption = "Filtrar [" + DataGrid.Columns(DataGrid.col).Text + "]"
    Select Case Button
        Case 2
            PopupMenu mnupri
    End Select

End Sub

Private Sub Form_Activate()
     DataGrid.EvenRowStyle.BackColor = &HFFFFFF
    DataGrid.OddRowStyle.BackColor = &HC0FFFF
    DataGrid.HighlightRowStyle.BackColor = vbActiveTitleBar
    DataGrid.HighlightRowStyle.ForeColor = vbWhite
    DataGrid.AlternatingRowStyle = True

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub Form_Load()
    Me.Top = 1020
    Me.Left = 2820
       
    DataGrid.OddRowStyle.BackColor = &HC0FFFF
    DataGrid.EvenRowStyle.BackColor = &HFFFFFF
    
    DataGrid.HighlightRowStyle.BackColor = vbActiveTitleBar
    DataGrid.HighlightRowStyle.ForeColor = vbWhite
    DataGrid.AlternatingRowStyle = True
    
    DataAyuda.DatabaseName = wrutatemp & "\Temp_com.mdb"
    DataAyuda.RecordSource = "select * from temp_ocompra"
    DataAyuda.Refresh
    

    Set dbbase = OpenDatabase(wrutabancos & " \DB_BANCOS.mdb")
    Set TbDetOrdenes = dbbase.OpenRecordset("if3orden")
    Set TbCabOrdenes = dbbase.OpenRecordset("if4orden")
    
    Set dbempre = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set TbProv = dbempre.OpenRecordset("EF2PROVEEDORES")
    TbProv.Index = "IDCODPROV"

    Set dbtempo = OpenDatabase(wrutatemp & "\temp_com.mdb")
    dbtempo.Execute ("delete * from temp_ocompra")
    Set tbtempo = dbtempo.OpenRecordset("temp_ocompra")
    
    If TbCabOrdenes.RecordCount > 0 Then
        TbCabOrdenes.MoveFirst
        Do While Not TbCabOrdenes.EOF
            Rem NSE If Len(Trim(TbCabOrdenes.Fields("f4regcom") & "")) = 0 And "" & TbCabOrdenes.Fields("F4CODPRV") = frmregiscom.TxtRucPrv.Text Then
            If "" & TbCabOrdenes.Fields("F4CODPRV") = frmregiscom.TxtRucPrv.Text Then
                If TbCabOrdenes.Fields("F4CODTEMP") & "" <> "1" Then
                    tbtempo.AddNew
                    tbtempo.Fields("F4NUMORD") = "" & "O/C" & Format(TbCabOrdenes.Fields("F4NUMORD"), "0000000")
                    tbtempo.Fields("F4CODPRV") = "" & TbCabOrdenes.Fields("F4CODPRV")
                    TbProv.Seek "=", TbCabOrdenes.Fields("F4CODPRV") & ""
                    If Not TbProv.NoMatch Then
                        tbtempo.Fields("F2NOMPROV") = "" & TbProv.Fields("F2NOMPROV")
                    End If
                    tbtempo.Fields("F4FECEMI") = "" & TbCabOrdenes.Fields("F4FECEMI")
                    tbtempo.Fields("F4OBSERVA") = "" & TbCabOrdenes.Fields("F4OBSERVA")
                    tbtempo.Fields("F4MONTO") = Val("" & TbCabOrdenes.Fields("F4MONTO")) - Val("" & TbCabOrdenes.Fields("F4APLICADO"))
                    tbtempo.Update
                End If
            End If
            TbCabOrdenes.MoveNext
        Loop
        Set TbCabOrdenes = dbbase.OpenRecordset("if4servicio")
        If TbCabOrdenes.RecordCount > 0 Then
            TbCabOrdenes.MoveFirst
            Do While Not TbCabOrdenes.EOF
                If "" & TbCabOrdenes.Fields("F4CODPRV") = frmregiscom.TxtRucPrv.Text Then
                    If "" & TbCabOrdenes.Fields("F4CODTEMP") <> "1" Then
                        tbtempo.AddNew
                        tbtempo.Fields("F4NUMORD") = "" & "O/S" & Format(TbCabOrdenes.Fields("F4NUMORD"), "0000000")
                        tbtempo.Fields("F4CODPRV") = "" & TbCabOrdenes.Fields("F4CODPRV")
                        TbProv.Seek "=", TbCabOrdenes.Fields("F4CODPRV")
                        If Not TbProv.NoMatch Then
                            tbtempo.Fields("F2NOMPROV") = "" & TbProv.Fields("F2NOMPROV")
                        End If
                        tbtempo.Fields("F4FECEMI") = "" & TbCabOrdenes.Fields("F4FECEMI")
                        tbtempo.Fields("F4OBSERVA") = Left("" & TbCabOrdenes.Fields("F4OBSERVA"), 200)
                        tbtempo.Fields("F4MONTO") = Val("" & TbCabOrdenes.Fields("F4MONTO"))
                        tbtempo.Update
                    End If
                End If
                TbCabOrdenes.MoveNext
            Loop
        End If
       '--------------------
        If tbtempo.RecordCount < 0 Then
           MsgBox "No hay Ordenes Registradas", 64, "Lista de Orden de Compra"
           Unload Me
        End If
   End If
   
    TbDetOrdenes.Close
    TbCabOrdenes.Close
    dbbase.Close

    tbtempo.Close
    dbtempo.Close

    TbProv.Close
    dbempre.Close

    DataAyuda.DatabaseName = wrutatemp & "\temp_com.mdb"
    DataAyuda.RecordSource = "select * from temp_ocompra order by f4numord"
    DataAyuda.Refresh
    

End Sub

Private Sub mkfecha_GotFocus()
mkfecha.SelStart = 0
mkfecha.SelLength = Len(mkfecha.Text)
End Sub

Private Sub mkfecha_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  txtope.Text = mkfecha.Text
  txtope_KeyPress 13
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
  Set DataAyuda.Recordset = rs.OpenRecordset(rs.Type)
  Set rs = Nothing

End Sub

Private Sub mnufiltroavanz_Click()
    Select Case DataGrid.col
        Case 0:
            lblbusca.Visible = True
            lblbusca.Caption = DataGrid.Columns(DataGrid.col).Caption
            txtope.Visible = True
            txtope.Text = ""
            txtope.SetFocus
        Case 2
            lblbusca.Visible = True
            lblbusca.Caption = DataGrid.Columns(DataGrid.col).Caption
            mkfecha.Visible = True
            mkfecha.Text = Date
            'txtope.Visible = True
            'txtope.Text = ""
            mkfecha.SetFocus
            'txtope.SetFocus
        'Case 3
         '   lblbusca.Visible = True
          '  lblbusca.Caption = datagrid.Columns(datagrid.col).Caption
          '  txtope.Visible = True
           ' txtope.Text = ""
           ' txtope.SetFocus
    End Select
End Sub

Private Sub mnuordasc_Click()
  Set rs = DataAyuda.Recordset
  rs.Sort = "[" + DataGrid.Columns(DataGrid.col).DataField + "] Asc"
  Set DataAyuda.Recordset = rs.OpenRecordset(rs.Type)
  Set rs = Nothing

End Sub

Private Sub mnuorddesc_Click()
  Set rs = DataAyuda.Recordset
  rs.Sort = "[" + DataGrid.Columns(DataGrid.col).DataField + "] Desc"
  Set DataAyuda.Recordset = rs.OpenRecordset(rs.Type)
  Set rs = Nothing

End Sub


Private Sub MnuTodo_Click()
    
    DataAyuda.DatabaseName = wrutatemp & "\temp_com.mdb"
    DataAyuda.RecordSource = "select * from temp_ocompra order by f4numord"
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
        txtope.Text = "*" & txtope.Text
       If Len(txtope) = 0 Then txtope.SetFocus: Exit Sub
          If InStr(txtope, "*") = Len(Trim(txtope)) Then
             DATO = Left(txtope, Len(Trim(txtope)) - 1)
          Else
              DATO = txtope.Text
          End If
          Set rs = DataAyuda.Recordset
          rs.Filter = "[" + DataGrid.Columns(DataGrid.col).DataField + "]" + " Like  '" + DATO + "*'"
          If rs.EOF Then txtope.SetFocus: Exit Sub Else txtope.Visible = False: lblbusca.Visible = False
          Set DataAyuda.Recordset = rs.OpenRecordset(rs.Type)
          Set rs = Nothing
         DataGrid.SetFocus
    
    End If
End Sub
