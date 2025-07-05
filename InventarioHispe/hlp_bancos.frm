VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form hlp_bancos 
   Caption         =   "Bancos"
   ClientHeight    =   3750
   ClientLeft      =   3585
   ClientTop       =   5160
   ClientWidth     =   5625
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   5625
   Begin Threed.SSPanel SSPanel1 
      Height          =   3180
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   5550
      _Version        =   65536
      _ExtentX        =   9790
      _ExtentY        =   5609
      _StockProps     =   15
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
      Begin VB.TextBox TxtFiltro 
         Height          =   315
         Index           =   0
         Left            =   105
         MaxLength       =   2
         TabIndex        =   4
         Top             =   2655
         Width           =   870
      End
      Begin VB.TextBox TxtFiltro 
         Height          =   315
         Index           =   1
         Left            =   1140
         MaxLength       =   20
         TabIndex        =   3
         Top             =   2655
         Width           =   4035
      End
      Begin TrueOleDBGrid70.TDBGrid DBGrid1 
         Bindings        =   "hlp_bancos.frx":0000
         Height          =   2445
         Left            =   90
         TabIndex        =   1
         Top             =   90
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   4313
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Código"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Descripción"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   2
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   2
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1614"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1535"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=74272"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=7355"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=7276"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).AllowSizing=0"
         Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=74272"
         Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Arial"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&H8000000F&"
         _StyleDefs(11)  =   ":id=2,.fgcolor=&H80000012&"
         _StyleDefs(12)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(13)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(14)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(15)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(16)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(17)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(18)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(19)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(20)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(21)  =   "Splits(0).Style:id=43,.parent=1"
         _StyleDefs(22)  =   "Splits(0).CaptionStyle:id=52,.parent=4"
         _StyleDefs(23)  =   "Splits(0).HeadingStyle:id=44,.parent=2"
         _StyleDefs(24)  =   "Splits(0).FooterStyle:id=45,.parent=3"
         _StyleDefs(25)  =   "Splits(0).InactiveStyle:id=46,.parent=5"
         _StyleDefs(26)  =   "Splits(0).SelectedStyle:id=48,.parent=6"
         _StyleDefs(27)  =   "Splits(0).EditorStyle:id=47,.parent=7"
         _StyleDefs(28)  =   "Splits(0).HighlightRowStyle:id=49,.parent=8"
         _StyleDefs(29)  =   "Splits(0).EvenRowStyle:id=50,.parent=9"
         _StyleDefs(30)  =   "Splits(0).OddRowStyle:id=51,.parent=10"
         _StyleDefs(31)  =   "Splits(0).RecordSelectorStyle:id=53,.parent=11"
         _StyleDefs(32)  =   "Splits(0).FilterBarStyle:id=54,.parent=12"
         _StyleDefs(33)  =   "Splits(0).Columns(0).Style:id=28,.parent=43,.alignment=0,.valignment=1"
         _StyleDefs(34)  =   ":id=28,.locked=-1"
         _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=44,.alignment=2"
         _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=45"
         _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=47"
         _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=32,.parent=43,.alignment=0,.valignment=1"
         _StyleDefs(39)  =   ":id=32,.locked=-1"
         _StyleDefs(40)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=44,.alignment=2"
         _StyleDefs(41)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=45"
         _StyleDefs(42)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=47"
         _StyleDefs(43)  =   "Named:id=33:Normal"
         _StyleDefs(44)  =   ":id=33,.parent=0"
         _StyleDefs(45)  =   "Named:id=34:Heading"
         _StyleDefs(46)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(47)  =   ":id=34,.wraptext=-1"
         _StyleDefs(48)  =   "Named:id=35:Footing"
         _StyleDefs(49)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(50)  =   "Named:id=36:Selected"
         _StyleDefs(51)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(52)  =   "Named:id=37:Caption"
         _StyleDefs(53)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(54)  =   "Named:id=38:HighlightRow"
         _StyleDefs(55)  =   ":id=38,.parent=33,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
         _StyleDefs(56)  =   "Named:id=39:EvenRow"
         _StyleDefs(57)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(58)  =   "Named:id=40:OddRow"
         _StyleDefs(59)  =   ":id=40,.parent=33"
         _StyleDefs(60)  =   "Named:id=41:RecordSelector"
         _StyleDefs(61)  =   ":id=41,.parent=34"
         _StyleDefs(62)  =   "Named:id=42:FilterBar"
         _StyleDefs(63)  =   ":id=42,.parent=33"
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
         TabIndex        =   2
         Top             =   2745
         Visible         =   0   'False
         Width           =   120
      End
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   75
      Top             =   3330
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   1
      Tools           =   "hlp_bancos.frx":0018
      ToolBars        =   "hlp_bancos.frx":008B
   End
   Begin VB.Menu MnuPri 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu MnuFiltro 
         Caption         =   "Filtrar"
      End
      Begin VB.Menu mnufiltroavaz 
         Caption         =   "Filtro Avanzado"
      End
      Begin VB.Menu MnuOrdAsc 
         Caption         =   "Ord. Asc"
      End
      Begin VB.Menu MnuOrdDesc 
         Caption         =   "Ord. Desc"
      End
      Begin VB.Menu MnuTodos 
         Caption         =   "Mostrar Todos"
      End
   End
End
Attribute VB_Name = "hlp_bancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cnnDatos  As New ADODB.Connection
Dim oColLis   As TrueOleDBGrid70.Column
Dim rstAyuda  As New ADODB.Recordset
Dim sw_filtro As Integer

Private Sub SSActiveToolBars1_ComboCloseUp(ByVal Tool As ActiveToolBars.SSTool)
Select Case Tool.Id
        Case "ID_Filtro"
            Select Case Tool.ComboBox.ListIndex
                Case 0
                    sw_filtro = 0
                Case 1
                    sw_filtro = 1
            End Select
End Select
                    
End Sub

Private Sub Form_Load()
    Inicializar
    If Not Conectar() Then
        MsgBox "Existe un Problema al Conectarse conla Base de Datos.", vbExclamation, "Mensaje de Advertencia"
    Else
        MostrarGrid
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyEscape:
        Unload Me
    Case vbKeyF3:
        If TxtFiltro(0).Enabled Then TxtFiltro(0).SetFocus
    Case vbKeyF4
        If DBGrid1.Enabled Then DBGrid1.SetFocus
    Case vbKeyF5:
        CambiarFiltro
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DBGrid1.Close
    
    If rstAyuda.State = adStateOpen Then rstAyuda.Close
    If cnnDatos.State = adStateOpen Then cnnDatos.Close
    
    Set rstAyuda = Nothing
    Set cnnDatos = Nothing
End Sub

Private Sub dbgrid1_DblClick()
    If DBGrid1.Row <> 0 Then Elegir
End Sub

Private Sub dbgrid1_HeadClick(ByVal ColIndex As Integer)
    If DBGrid1.Row = 0 Then
        rstAyuda.Sort = rstAyuda(ColIndex).Name
        DBGrid1.SetFocus
    End If
End Sub

Private Sub dbgrid1_FilterChange()
    rstAyuda.Filter = Filtrar()
    If rstAyuda.Bof And rstAyuda.EOF Then
        MsgBox "No Existe Infrormacion para el Filtro.", vbInformation, "Mensaje Informativo"
        Desfiltrar
    End If
End Sub

Private Sub dbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyReturn: Elegir
    End Select
End Sub

Private Sub TxtFiltro_Change(Index As Integer)
    Dim cValFil As String
    '
    cValFil = TxtFiltro(Index)
    '
    DBGrid1.Columns(Index).FilterText = cValFil
    dbgrid1_FilterChange
    '
    TxtFiltro(Index).SetFocus
End Sub

Private Sub TxtFiltro_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then TxtFiltro_lostfocus Index
End Sub

Private Sub TxtFiltro_lostfocus(Index As Integer)
    Select Case Index
    Case 1:
        DBGrid1.SetFocus
    Case Else:
        If TxtFiltro(Index + 1).Enabled Then TxtFiltro(Index + 1).SetFocus
    End Select
End Sub

Private Function Conectar() As Boolean
    Dim cValcon As String
    
    cValcon = "Provider = Microsoft.Jet.Oledb.4.0;Data Source = " & wrutabancos & "\DB_BANCOS.MDB"
    Conectar = True
    On Error GoTo ErrAdo
    cnnDatos.ConnectionString = cValcon
    cnnDatos.Open
    On Error GoTo 0
    '
    Exit Function
    
ErrAdo:
    Conectar = False
    On Error GoTo 0
    Exit Function
End Function

Private Sub MostrarGrid()
    Dim cValSql As String
    
    cValSql = "select CODIGO,BANCO from bancos where modalidad ='1'"
    On Error GoTo Errdao
    rstAyuda.CursorLocation = adUseClient
    rstAyuda.Open cValSql, cnnDatos, adOpenDynamic, adLockReadOnly
    On Error GoTo 0
    '
     DBGrid1.EvenRowStyle.BackColor = &HFFFFFF
    DBGrid1.OddRowStyle.BackColor = &HC0FFFF
    DBGrid1.HighlightRowStyle.BackColor = vbActiveTitleBar
    DBGrid1.HighlightRowStyle.ForeColor = vbWhite
    DBGrid1.AlternatingRowStyle = True
    '
    DBGrid1.Columns(0).DataField = "CODIGO"
    DBGrid1.Columns(1).DataField = "BANCO"
    Set DBGrid1.DataSource = rstAyuda
    '
    Exit Sub
    
Errdao:
    On Error GoTo 0
    Exit Sub
End Sub

Private Sub Elegir()
    wcodban = DBGrid1.Columns(0)
    wnomban = DBGrid1.Columns(1)
    Unload Me
End Sub

Private Function Filtrar() As String
    Dim nValFil As Single
    Dim cValFil As String
    '
    nValFil = sw_filtro
    cValFil = ""
    '
    For Each oColLis In DBGrid1.Columns
        If Trim(oColLis.FilterText) <> "" Then
            If cValFil <> "" Then cValFil = cValFil & " AND "
            '
            Select Case nValFil
            Case 0: cValFil = cValFil & oColLis.DataField & " LIKE  '" & oColLis.FilterText & "*'"
            Case 1: cValFil = cValFil & oColLis.DataField & " LIKE '*" & oColLis.FilterText & "*'"
            End Select
        End If
    Next oColLis
    '
    Filtrar = cValFil
End Function

Private Sub Desfiltrar()
    For Each oColLis In DBGrid1.Columns
            oColLis.FilterText = ""
    Next oColLis
    TxtFiltro(0).Text = ""
    TxtFiltro(1).Text = ""
    rstAyuda.Filter = adFilterNone
End Sub

Private Sub CambiarFiltro()
    Dim nValFil As Single
    nValFil = sw_filtro
    Desfiltrar
    SSActiveToolBars1.Tools("ID_Filtro").ComboBox.ListIndex = IIf(nValFil = 0, 1, 0)
    sw_filtro = IIf(nValFil = 0, 1, 0)
End Sub

Private Sub Inicializar()
    wcodban = ""
    wnomban = ""
End Sub

