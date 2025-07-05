VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form hlp_clientes 
   Appearance      =   0  'Flat
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ayuda de Clientes"
   ClientHeight    =   4500
   ClientLeft      =   1305
   ClientTop       =   2235
   ClientWidth     =   9360
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
   Icon            =   "hlp_clientes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4500
   ScaleWidth      =   9360
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   4425
      Left            =   45
      TabIndex        =   1
      Top             =   0
      Width           =   9240
      _Version        =   65536
      _ExtentX        =   16298
      _ExtentY        =   7805
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
      Begin MSAdodcLib.Adodc DataAyuda 
         Height          =   330
         Left            =   1260
         Top             =   2520
         Visible         =   0   'False
         Width           =   4515
         _ExtentX        =   7964
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
         Top             =   3780
         Visible         =   0   'False
         Width           =   2445
      End
      Begin TrueOleDBGrid70.TDBGrid datagrid 
         Bindings        =   "hlp_clientes.frx":000C
         Height          =   3030
         Left            =   90
         TabIndex        =   2
         Top             =   90
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   5345
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Código"
         Columns(0).DataField=   "F2CODCLI"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Cliente"
         Columns(1).DataField=   "F2NOMCLI"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "RUC"
         Columns(2).DataField=   "F2NEWRUC"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   2
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1799"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1720"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=65808"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=10186"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=10107"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=65808"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=3334"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=3254"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=65808"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
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
         DeadAreaBackColor=   -2147483633
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=124,.bold=-1,.fontsize=825,.italic=0"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=44"
         _StyleDefs(37)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=45"
         _StyleDefs(38)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=47"
         _StyleDefs(39)  =   "Splits(0).Columns(1).Style:id=32,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(40)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=44"
         _StyleDefs(41)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=45"
         _StyleDefs(42)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=47"
         _StyleDefs(43)  =   "Splits(0).Columns(2).Style:id=62,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(44)  =   "Splits(0).Columns(2).HeadingStyle:id=59,.parent=44"
         _StyleDefs(45)  =   "Splits(0).Columns(2).FooterStyle:id=60,.parent=45"
         _StyleDefs(46)  =   "Splits(0).Columns(2).EditorStyle:id=61,.parent=47"
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
         _StyleDefs(57)  =   ":id=37,.parent=34,.alignment=2,.bold=0,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(58)  =   ":id=37,.strikethrough=0,.charset=0"
         _StyleDefs(59)  =   ":id=37,.fontname=MS Sans Serif"
         _StyleDefs(60)  =   "Named:id=38:HighlightRow"
         _StyleDefs(61)  =   ":id=38,.parent=33,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
         _StyleDefs(62)  =   "Named:id=39:EvenRow"
         _StyleDefs(63)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(64)  =   "Named:id=40:OddRow"
         _StyleDefs(65)  =   ":id=40,.parent=33"
         _StyleDefs(66)  =   "Named:id=41:RecordSelector"
         _StyleDefs(67)  =   ":id=41,.parent=34"
         _StyleDefs(68)  =   "Named:id=42:FilterBar"
         _StyleDefs(69)  =   ":id=42,.parent=33"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "INS -> Nuevo"
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
         Left            =   7965
         TabIndex        =   6
         Top             =   3240
         Width           =   945
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
         TabIndex        =   5
         Top             =   3510
         Visible         =   0   'False
         Width           =   60
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
         Left            =   6525
         TabIndex        =   4
         Top             =   3780
         Visible         =   0   'False
         Width           =   2490
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "[F3] ---> Filtro Avanzado: Cliente"
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
         TabIndex        =   3
         Top             =   3150
         Width           =   2340
      End
   End
   Begin VB.Menu mnupri 
      Caption         =   ""
      Begin VB.Menu mnufiltro 
         Caption         =   "Filtrar"
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
Attribute VB_Name = "hlp_clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As adodb.Recordset

Private Sub datagrid_DblClick()
    
    datagrid_KeyDown 13, 0
    
End Sub

Private Sub datagrid_KeyDown(KeyCode As Integer, Shift As Integer)
   
    Select Case KeyCode
        Case 13:
            wcodcliprov = "" & datagrid.Columns(0)
            wnomcliprov = "" & datagrid.Columns(1)
            wcodpag = "" & datagrid.Columns(0)
            wnompag = "" & datagrid.Columns(1)
            Unload Me
        Case 27:
            wcodcliprov = ""
            wnomcliprov = ""
            Unload Me
        Case 114:
            datagrid.Col = 1
            mnufiltroavanz_Click
    End Select
   
End Sub

Private Sub datagrid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   Label1.Visible = True
   
End Sub

Private Sub datagrid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If txtope.Visible = True Then txtope.Visible = False
    If lblbusca.Visible = True Then lblbusca.Visible = False
    mnufiltro.Caption = "Filtrar [" + datagrid.Columns(datagrid.Col).Text + "]"
    Select Case Button
        Case 2
            PopupMenu mnupri
    End Select
    
End Sub

Private Sub Form_Load()

    datagrid.EvenRowStyle.BackColor = &HFFFFFF
    datagrid.OddRowStyle.BackColor = &HC0FFFF
    datagrid.HighlightRowStyle.BackColor = vbActiveTitleBar
    datagrid.HighlightRowStyle.ForeColor = vbWhite
    datagrid.AlternatingRowStyle = True

    DataAyuda.ConnectionString = cnn_dbbancos
    DataAyuda.RecordSource = "select f2codcli,f2nomcli,f2newruc,f2dircli,f2forpag,f2cotiza,F2COTIZA_OBS,F2TIPDOC,F2DOCIDENTIDAD from EF2CLIENTES order by F2CODCLI"
    DataAyuda.Refresh
    
    txtope.Visible = True
    datagrid.Col = 1

End Sub

Private Sub mnufiltro_Click()
    
    Set rs = New adodb.Recordset
    Set rs = DataAyuda.Recordset
    Select Case DataAyuda.Recordset.Fields(datagrid.Columns(datagrid.Col).DataField).Type
        Case 10
            rs.Filter = "[" + datagrid.Columns(datagrid.Col).DataField + "]" + " = '" + Trim("" & datagrid.Columns(datagrid.Col).Text) + "'"
        Case 4
            rs.Filter = "[" + datagrid.Columns(datagrid.Col).DataField + "]" + " = " + datagrid.Columns(datagrid.Col).Text
        Case 8
            If IsDate(datagrid.Columns(datagrid.Col).Text) Then
               rs.Filter = "[" + datagrid.Columns(datagrid.Col).DataField + "]" + "=#" + datagrid.Columns(datagrid.Col).Text + "#"
            Else
               MsgBox "Ingrese una Fecha Valida..!", 32, "Advertencia"
               Exit Sub
            End If
    End Select
    Set DataAyuda.Recordset = rs.DataSource
    Set rs = Nothing
    
End Sub

Private Sub mnufiltroavanz_Click()

    Select Case datagrid.Col
        Case 0:
            lblbusca.Visible = True
            lblbusca.Caption = datagrid.Columns(datagrid.Col).Caption
            txtope.Visible = True
            txtope.Text = ""
            txtope.SetFocus
        Case 1
            lblbusca.Visible = True
            lblbusca.Caption = datagrid.Columns(datagrid.Col).Caption
            txtope.Visible = True
            txtope.Text = ""
            txtope.SetFocus
        Case 3
            lblbusca.Visible = True
            lblbusca.Caption = datagrid.Columns(datagrid.Col).Caption
            txtope.Visible = True
            txtope.Text = ""
            txtope.SetFocus
        Case 5
            lblbusca.Visible = True
            lblbusca.Caption = datagrid.Columns(datagrid.Col).Caption
            txtope.Visible = True
            txtope.Text = ""
            txtope.SetFocus
        Case 2
            lblbusca.Visible = True
            lblbusca.Caption = datagrid.Columns(datagrid.Col).Caption
            txtope.Visible = True
            txtope.Text = ""
            txtope.SetFocus
    End Select
End Sub

Private Sub mnuordasc_Click()
    
    Set rs = New adodb.Recordset
    Set rs = DataAyuda.Recordset
    rs.Sort = "[" + datagrid.Columns(datagrid.Col).DataField + "] Asc"
    Set DataAyuda.Recordset = rs.DataSource
    Set rs = Nothing
    
End Sub

Private Sub mnuorddesc_Click()
    
    Set rs = New adodb.Recordset
    Set rs = DataAyuda.Recordset
    rs.Sort = "[" + datagrid.Columns(datagrid.Col).DataField + "] Desc"
    Set DataAyuda.Recordset = rs.DataSource
    Set rs = Nothing
    
End Sub

Private Sub MnuTodo_Click()

    DataAyuda.RecordSource = "select f2codcli,f2nomcli,f2newruc,f2dircli,f2forpag,f2cotiza,F2COTIZA_OBS,F2TIPDOC,F2DOCIDENTIDAD from EF2CLIENTES order by F2CODCLI"
    DataAyuda.Refresh
    
End Sub

Private Sub SSPanel1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Label1.Visible = False
    
End Sub

Private Sub txtope_KeyPress(KeyAscii As Integer)
Dim sql As String

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If Len(txtope.Text) = 0 Then txtope.SetFocus: Exit Sub
        txtope.Text = "*" & txtope.Text
        If InStr(txtope, "*") = Len(Trim(txtope)) Then
            DATO = left(txtope, Len(Trim(txtope)) - 1)
        Else
            DATO = txtope.Text
        End If
        Set rs = New adodb.Recordset
        Set rs = DataAyuda.Recordset
        rs.Filter = "[" + datagrid.Columns(datagrid.Col).DataField + "]" + " Like  '" + DATO + "*'"
        If rs.EOF Then txtope.SetFocus: Exit Sub Else txtope.Visible = False: lblbusca.Visible = False
        Set DataAyuda.Recordset = rs.DataSource
        Set rs = Nothing
        datagrid.SetFocus
    End If
    
End Sub
