VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmayusoli 
   Caption         =   "Solicitud de Suministro"
   ClientHeight    =   3060
   ClientLeft      =   2085
   ClientTop       =   2370
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   6060
   Begin Threed.SSPanel SSPanel1 
      Height          =   2850
      Left            =   0
      TabIndex        =   0
      Top             =   45
      Width           =   6000
      _Version        =   65536
      _ExtentX        =   10583
      _ExtentY        =   5027
      _StockProps     =   15
      BackColor       =   12632256
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
      Begin MSAdodcLib.Adodc Dt_HlpSolComp 
         Height          =   330
         Left            =   1305
         Top             =   1575
         Visible         =   0   'False
         Width           =   2985
         _ExtentX        =   5265
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin TrueOleDBGrid70.TDBGrid DBG1 
         Bindings        =   "frmayusoli.frx":0000
         Height          =   1725
         Left            =   225
         TabIndex        =   1
         Top             =   180
         Width           =   5550
         _ExtentX        =   9790
         _ExtentY        =   3043
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "# Sol."
         Columns(0).DataField=   "cod_solicitud"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Fecha"
         Columns(1).DataField=   "cs_fecha"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Estado"
         Columns(2).DataField=   "cs_estado"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Solicitante"
         Columns(3).DataField=   "cs_codsolicitante"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Obra"
         Columns(4).DataField=   "cs_codcosto"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   5
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   2
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   2
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=5"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2117"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2037"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=65808"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=2328"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2249"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=65808"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=1799"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1720"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=65808"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=1535"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1455"
         Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=65808"
         Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(25)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=65792"
         Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=40,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(32)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=44"
         _StyleDefs(33)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=45"
         _StyleDefs(34)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=47"
         _StyleDefs(35)  =   "Splits(0).Columns(1).Style:id=32,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(36)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=44"
         _StyleDefs(37)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=45"
         _StyleDefs(38)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=47"
         _StyleDefs(39)  =   "Splits(0).Columns(2).Style:id=58,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(40)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=44"
         _StyleDefs(41)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=45"
         _StyleDefs(42)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=47"
         _StyleDefs(43)  =   "Splits(0).Columns(3).Style:id=62,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(44)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=44"
         _StyleDefs(45)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=45"
         _StyleDefs(46)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=47"
         _StyleDefs(47)  =   "Splits(0).Columns(4).Style:id=16,.parent=43"
         _StyleDefs(48)  =   "Splits(0).Columns(4).HeadingStyle:id=13,.parent=44"
         _StyleDefs(49)  =   "Splits(0).Columns(4).FooterStyle:id=14,.parent=45"
         _StyleDefs(50)  =   "Splits(0).Columns(4).EditorStyle:id=15,.parent=47"
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
      Begin MSMask.MaskEdBox txtfecha 
         Height          =   285
         Left            =   225
         TabIndex        =   6
         Top             =   2430
         Visible         =   0   'False
         Width           =   1230
         _ExtentX        =   2170
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
         Height          =   285
         Left            =   225
         TabIndex        =   2
         ToolTipText     =   "Simplifique la consulta digitando el Simbolo (*)"
         Top             =   2430
         Visible         =   0   'False
         Width           =   2445
      End
      Begin VB.Label Label1 
         Caption         =   "Ayuda: Hacer Click Derecho en la Descripcion que desee buscar"
         ForeColor       =   &H00800000&
         Height          =   780
         Left            =   3285
         TabIndex        =   5
         Top             =   1980
         Visible         =   0   'False
         Width           =   1680
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
         Left            =   225
         TabIndex        =   4
         Top             =   2160
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "[F3] ---> Filtro Avanzado: Descripción"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   225
         TabIndex        =   3
         Top             =   1935
         Width           =   2610
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
         Caption         =   "Ord. Asc."
      End
      Begin VB.Menu mnuorddesc 
         Caption         =   "Ord. Desc."
      End
      Begin VB.Menu mnutodos 
         Caption         =   "Mostrar Todos"
      End
   End
End
Attribute VB_Name = "frmayusoli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQL As String

Private Sub DBG1_KeyPress(KeyAscii As Integer)
    
    num_solcomp = DBG1.Columns(0)
    selecciona = True
    Unload Me
    
End Sub

Private Sub DBG1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   Label1.Visible = True
   
End Sub

Private Sub DBG1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   If txtope.Visible = True Then txtope.Visible = False
   If lblbusca.Visible = True Then lblbusca.Visible = False
   If txtfecha.Visible = True Then txtfecha.Visible = False
    
    mnufiltro.Caption = "Filtrar [" + DBG1.Columns(DBG1.Col).Text + "]"
    Select Case Button
        Case 2
            PopupMenu mnupri
    End Select
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 27 Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    Dt_HlpSolComp.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutabancos & "\db_logi.mdb;Persist Security Info=False"
    SQL = "Select * From Tb_CabSolicitud Where cs_estado ='" & "PEND" & "' and CS_DOCUMENTO='ST' Order By cod_solicitud Desc" 'SOLICITUD NUEVA
    Dt_HlpSolComp.RecordSource = SQL
    Dt_HlpSolComp.Refresh

End Sub

Private Sub DBG1_DblClick()
    
    num_solcomp = DBG1.Columns(0)
    selecciona = True
    Unload Me
    
End Sub

Private Sub DBG1_KeyDown(KeyCode As Integer, Shift As Integer) 'sol. de compra/serv.
   
    If KeyCode = 114 Then
        DBG1.Col = 0
        mnufiltroavanz_Click
    End If
    
    If KeyCode = 13 Then DBG1_DblClick
    
End Sub

Private Sub mnufiltro_Click()
    
    Set rs = Dt_HlpSolComp.Recordset
    Select Case Dt_HlpSolComp.Recordset.Fields(DBG1.Columns(DBG1.Col).DataField).Type
        Case 10
            rs.Filter = "[" + DBG1.Columns(DBG1.Col).DataField + "]" + " = '" + Trim("" & DBG1.Columns(DBG1.Col).Text) + "'"
        Case 4
            rs.Filter = "[" + DBG1.Columns(DBG1.Col).DataField + "]" + " = " + DBG1.Columns(DBG1.Col).Text
        Case 8
            If IsDate(DBG1.Columns(DBG1.Col).Text) Then
               rs.Filter = "[" + DBG1.Columns(DBG1.Col).DataField + "]" + "=#" + DBG1.Columns(DBG1.Col).Text + "#"
            Else
               MsgBox "Ingrese una Fecha Valida..!", 32, "Advertencia"
               Exit Sub
            End If
    End Select
    Set Dt_HlpSolComp.Recordset = rs.OpenRecordset(rs.Type)
    Set rs = Nothing
    
End Sub

Private Sub mnufiltroavanz_Click()

    Select Case DBG1.Col
        Case 0:
            lblbusca.Visible = True
            lblbusca.Caption = DBG1.Columns(DBG1.Col).Caption
            txtope.Visible = True
            txtope.Text = ""
            txtope.SetFocus
        Case 1
            lblbusca.Visible = True
            lblbusca.Caption = DataGrid.Columns(DataGrid.Col).Caption
            txtfecha.Visible = True
            txtfecha.Text = Date
            txtfecha.SetFocus
        Case 2:
            lblbusca.Visible = True
            lblbusca.Caption = DBG1.Columns(DBG1.Col).Caption
            txtope.Visible = True
            txtope.Text = ""
            txtope.SetFocus
        Case 3:
            lblbusca.Visible = True
            lblbusca.Caption = DBG1.Columns(DBG1.Col).Caption
            txtope.Visible = True
            txtope.Text = ""
            txtope.SetFocus
    End Select
    
End Sub

Private Sub mnuordasc_Click()
   
    Set rs = Dt_HlpSolComp.Recordset
    rs.Sort = "[" + DBG1.Columns(DBG1.Col).DataField + "] Asc"
    Set Dt_HlpSolComp.Recordset = rs.OpenRecordset(rs.Type)
    Set rs = Nothing
    
End Sub

Private Sub mnuorddesc_Click()
   
    Set rs = Dt_HlpSolComp.Recordset
    rs.Sort = "[" + DBG1.Columns(DBG1.Col).DataField + "] Desc"
    Set Dt_HlpSolComp.Recordset = rs.OpenRecordset(rs.Type)
    Set rs = Nothing
    
End Sub

Private Sub mnutodos_Click()
   
    Dt_HlpSolComp.RecordSource = SQL
    Dt_HlpSolComp.Refresh

End Sub

Private Sub SSPanel1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
    Label1.Visible = False
    
End Sub

Private Sub txtfecha_GotFocus()

    txtfecha.SelStart = 0
    txtfecha.SelLength = Len(txtfecha.Text)
    
End Sub

Private Sub txtfecha_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtope.Text = txtfecha.Text
        txtope_KeyPress 13
    End If
    
End Sub

Private Sub txtope_KeyPress(KeyAscii As Integer)
Dim SQL As String

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If Len(txtope) = 0 Then txtope.SetFocus: Exit Sub
        txtope.Text = "*" & txtope.Text
        If InStr(txtope, "*") = Len(Trim(txtope)) Then
            DATO = Left(txtope, Len(Trim(txtope)) - 1)
        Else
            DATO = txtope.Text
        End If
        Set rs = Dt_HlpSolComp.Recordset
        rs.Filter = "[" + DBG1.Columns(DBG1.Col).DataField + "]" + " Like  '" + DATO + "*'"
        If rs.EOF Then txtope.SetFocus: Exit Sub Else txtope.Visible = False: lblbusca.Visible = False
        Set Dt_HlpSolComp.Recordset = rs.OpenRecordset(rs.Type)
        Set rs = Nothing
        DBG1.SetFocus
    End If
    
End Sub
