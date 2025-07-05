VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "SSTBARS2.OCX"
Begin VB.Form hlp_productos 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista de Productos"
   ClientHeight    =   4575
   ClientLeft      =   660
   ClientTop       =   1860
   ClientWidth     =   9225
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
   ScaleHeight     =   4575
   ScaleWidth      =   9225
   Begin VB.Frame FrmDato 
      Height          =   4215
      Left            =   30
      TabIndex        =   0
      Top             =   345
      Width           =   9180
      Begin VB.TextBox TxtDato 
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
         Index           =   2
         Left            =   2070
         TabIndex        =   4
         Top             =   3855
         Width           =   4965
      End
      Begin VB.TextBox TxtDato 
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
         Index           =   1
         Left            =   1155
         TabIndex        =   3
         Top             =   3855
         Width           =   900
      End
      Begin VB.TextBox TxtDato 
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
         Index           =   0
         Left            =   60
         TabIndex        =   2
         Top             =   3855
         Width           =   1080
      End
      Begin TrueOleDBGrid70.TDBGrid TdbLista 
         Height          =   3660
         Left            =   60
         TabIndex        =   1
         Top             =   150
         Width           =   9045
         _ExtentX        =   15954
         _ExtentY        =   6456
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Codigo"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Cod. Fab."
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Descripción"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Precio. Vta"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Unidad"
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "F5STOCKACT"
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "F5PRECOS"
         Columns(6).DataField=   ""
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   7
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   2
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AllowColSelect=   0   'False
         Splits(0).DividerColor=   13154464
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=7"
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
         Splits(0)._ColumnProps(15)=   "Column(2).Width=9340"
         Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=9260"
         Splits(0)._ColumnProps(18)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(19)=   "Column(2).AllowSizing=0"
         Splits(0)._ColumnProps(20)=   "Column(2)._ColStyle=65808"
         Splits(0)._ColumnProps(21)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(22)=   "Column(3).Width=1535"
         Splits(0)._ColumnProps(23)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(24)=   "Column(3)._WidthInPix=1455"
         Splits(0)._ColumnProps(25)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(26)=   "Column(3).AllowSizing=0"
         Splits(0)._ColumnProps(27)=   "Column(3)._ColStyle=73730"
         Splits(0)._ColumnProps(28)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(29)=   "Column(4).Width=1058"
         Splits(0)._ColumnProps(30)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(4)._WidthInPix=979"
         Splits(0)._ColumnProps(32)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(33)=   "Column(4).AllowSizing=0"
         Splits(0)._ColumnProps(34)=   "Column(4)._ColStyle=66082"
         Splits(0)._ColumnProps(35)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(36)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(37)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(39)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(40)=   "Column(5)._ColStyle=65792"
         Splits(0)._ColumnProps(41)=   "Column(5).Visible=0"
         Splits(0)._ColumnProps(42)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(43)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(44)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(45)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(46)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(47)=   "Column(6)._ColStyle=65792"
         Splits(0)._ColumnProps(48)=   "Column(6).Visible=0"
         Splits(0)._ColumnProps(49)=   "Column(6).Order=7"
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
         _StyleDefs(43)  =   "Splits(0).Columns(2).Style:id=58,.parent=43,.alignment=0,.valignment=2,.locked=0"
         _StyleDefs(44)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=44"
         _StyleDefs(45)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=45"
         _StyleDefs(46)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=47"
         _StyleDefs(47)  =   "Splits(0).Columns(3).Style:id=62,.parent=43,.alignment=1,.locked=-1"
         _StyleDefs(48)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=44,.alignment=3"
         _StyleDefs(49)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=45"
         _StyleDefs(50)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=47"
         _StyleDefs(51)  =   "Splits(0).Columns(4).Style:id=66,.parent=43,.alignment=1,.valignment=1,.locked=0"
         _StyleDefs(52)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=44,.alignment=2"
         _StyleDefs(53)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=45"
         _StyleDefs(54)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=47"
         _StyleDefs(55)  =   "Splits(0).Columns(5).Style:id=16,.parent=43"
         _StyleDefs(56)  =   "Splits(0).Columns(5).HeadingStyle:id=13,.parent=44"
         _StyleDefs(57)  =   "Splits(0).Columns(5).FooterStyle:id=14,.parent=45"
         _StyleDefs(58)  =   "Splits(0).Columns(5).EditorStyle:id=15,.parent=47"
         _StyleDefs(59)  =   "Splits(0).Columns(6).Style:id=20,.parent=43"
         _StyleDefs(60)  =   "Splits(0).Columns(6).HeadingStyle:id=17,.parent=44"
         _StyleDefs(61)  =   "Splits(0).Columns(6).FooterStyle:id=18,.parent=45"
         _StyleDefs(62)  =   "Splits(0).Columns(6).EditorStyle:id=19,.parent=47"
         _StyleDefs(63)  =   "Named:id=33:Normal"
         _StyleDefs(64)  =   ":id=33,.parent=0"
         _StyleDefs(65)  =   "Named:id=34:Heading"
         _StyleDefs(66)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(67)  =   ":id=34,.wraptext=-1"
         _StyleDefs(68)  =   "Named:id=35:Footing"
         _StyleDefs(69)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(70)  =   "Named:id=36:Selected"
         _StyleDefs(71)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(72)  =   "Named:id=37:Caption"
         _StyleDefs(73)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(74)  =   "Named:id=38:HighlightRow"
         _StyleDefs(75)  =   ":id=38,.parent=33,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
         _StyleDefs(76)  =   "Named:id=39:EvenRow"
         _StyleDefs(77)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(78)  =   "Named:id=40:OddRow"
         _StyleDefs(79)  =   ":id=40,.parent=33"
         _StyleDefs(80)  =   "Named:id=41:RecordSelector"
         _StyleDefs(81)  =   ":id=41,.parent=34"
         _StyleDefs(82)  =   "Named:id=42:FilterBar"
         _StyleDefs(83)  =   ":id=42,.parent=33"
      End
   End
   Begin ActiveToolBars.SSActiveToolBars TlbBarra 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   3
      Tools           =   "algo.frx":0000
      ToolBars        =   "algo.frx":0D68
   End
   Begin VB.Menu MNUPRI 
      Caption         =   ""
      Visible         =   0   'False
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
Attribute VB_Name = "hlp_productos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RstTemp As New ADODB.Recordset

Private Sub TlbBarra_ComboCloseUp(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.Id
    Case "Filtro": Call DesFiltrar
    End Select
End Sub

Private Sub TlbBarra_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.Id
    Case "Salir": Unload Me
    Case "Filtro":
    End Select
End Sub
    
Private Sub Form_Load()
    Me.Width = 9315
    Me.Height = 4955
    Me.Top = 2000
    Me.Left = 3000
    '
    FrmDato.Top = -80
    '
    wcodproducto = ""
    wcodfab = ""
    wdesproducto = ""
    wmedida = ""
    wstockact = 0#
    wprecos = 0#
    '
    TlbBarra.Tools.ITEM("Filtro").ComboBox.ListIndex = 0
    '
    Call MostrarGrid
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyEscape: Unload Me
    Case vbKeyF3:     TxtDato(2).SetFocus
    Case vbKeyF4:     TdbLista.SetFocus
    Case vbKeyF5:     Call CambiarFiltro
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If RstTemp.State = adStateOpen Then RstTemp.Close
    '
    Set RstTemp = Nothing
End Sub

Private Sub TdbLista_DblClick()
    If TdbLista.Row <> 0 Then Call Seleccionar
 End Sub

Private Sub TdbLista_HeadClick(ByVal ColIndex As Integer)
    RstTemp.Sort = RstTemp(ColIndex).Name
    TdbLista.SetFocus
End Sub

Private Sub Tdblista_FilterChange()
    Dim oColLis As TrueOleDBGrid70.Column
    Dim oClsLis As TrueOleDBGrid70.Columns
    '
    Set oClsLis = TdbLista.Columns
    '
    TdbLista.HoldFields
    RstTemp.Filter = Filtrar()
    '
    If RstTemp.Bof And RstTemp.EOF Then
        MsgBox "No Existe Infrormacion para el Filtro.", vbInformation, "Mensaje Informativo"
        '
        Call DesFiltrar
    End If
    '
    TdbLista.EditActive = True
End Sub

Private Sub TdbLista_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyReturn: Call Seleccionar
    End Select
End Sub

Private Sub TxtDato_Change(Index As Integer)
    Dim cValFil As String
    '
    cValFil = TxtDato(Index)
    '
    TdbLista.Columns(Index).FilterText = cValFil
    Tdblista_FilterChange
    TxtDato(Index).SetFocus
End Sub

Private Sub TxtDato_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then TxtDato_Validate Index, True
End Sub

Private Sub TxtDato_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
    Case 2:
        If TxtDato(0).Enabled Then TxtDato(0).SetFocus
    Case Else:
        If TxtDato(Index + 1).Enabled Then TxtDato(Index + 1).SetFocus
    End Select
End Sub

Private Sub MostrarGrid()
    Dim cValCon As String
    Dim cValSql As String
    '
    cValCon = "provider=microsoft.jet.oledb.4.0;data source=" & wrutabancos & "\db_bancos.mdb"
    '
    If Len(Trim(wcod_alm)) = 0 Then
        cValSql = "SELECT a.f5codpro," & _
                         "a.F5NOMPRO," & _
                         "a.F5CODFAB," & _
                         "a.F5valvta," & _
                         "b.F7SIGMED " & _
                    "FROM IF5PLA AS a,EF7MEDIDAS AS b " & _
                   "WHERE a.F7CODMED = b.F7CODMED " & _
                   "ORDER BY a.F5NOMPRO"
    Else
        If wtipoguia = "I" Then
            cValSql = "SELECT a.f5codpro," & _
                             "a.F5CODFAB," & _
                             "a.F5NOMPRO," & _
                             "a.F5valvta," & _
                             "c.F7SIGMED " & _
                        "FROM IF5PLA AS a,IF6ALMA AS b,EF7MEDIDAS AS c " & _
                       "WHERE a.F5CODPRO = b.F5CODPRO AND " & _
                             "b.F2CODALM = '" & wcod_alm & "' AND " & _
                             "a.F7CODMED = c.F7CODMED " & _
                       "ORDER BY A.F5CODPRO"
        Else   '----- VALES DE SALIDA
            cValSql = "SELECT a.f5codpro," & _
                             "a.F5CODFAB," & _
                             "a.F5NOMPRO," & _
                             "a.F5valvta," & _
                             "C.F7SIGMED " & _
                        "FROM IF5PLA AS a, IF6ALMA AS b,EF7MEDIDAS AS c " & _
                       "WHERE a.F5CODPRO = b.F5CODPRO AND " & _
                             "b.F2CODALM = '" & wcod_alm & "' AND " & _
                             "a.F7CODMED = c.F7CODMED " & _
                       "ORDER BY a.F5CODPRO"
        End If
    End If
    '
    On Error GoTo ErrAdo
    RstTemp.CursorLocation = adUseClient
    RstTemp.Open cValSql, cValCon, adOpenDynamic, adLockReadOnly
    On Error GoTo 0
    '
    TdbLista.EvenRowStyle.BackColor = &HFFFFFF
    TdbLista.OddRowStyle.BackColor = &HC0FFFF
    TdbLista.HighlightRowStyle.BackColor = vbActiveTitleBar
    TdbLista.HighlightRowStyle.ForeColor = vbWhite
    TdbLista.AlternatingRowStyle = True
    '
    TdbLista.Columns(0).DataField = "F5CODPRO"
    TdbLista.Columns(1).DataField = "F5CODFAB"
    TdbLista.Columns(2).DataField = "F5NOMPRO"
    TdbLista.Columns(3).DataField = "F5VALVTA"
    TdbLista.Columns(4).DataField = "F7SIGMED"
    TdbLista.Columns(5).DataField = "F5STOCKACT"
    TdbLista.Columns(6).DataField = "F5PRECOS"
    '
    Set TdbLista.DataSource = RstTemp
    '
    Exit Sub

ErrAdo:
    MsgBox Err.Description
    On Error GoTo 0
End Sub

Private Function Filtrar() As String
    Dim oColLis As TrueOleDBGrid70.Column
    Dim oClsLis As TrueOleDBGrid70.Columns
    Dim nValFil As Single
    Dim cValFil As String
    '
    nValFil = TlbBarra.Tools.ITEM("Filtro").ComboBox.ListIndex
    cValFil = ""
    '
    Set oClsLis = TdbLista.Columns
    '
    For Each oColLis In oClsLis
        If oColLis.ColIndex <= 3 Then
            If Trim(oColLis.FilterText) <> "" Then
                If cValFil <> "" Then cValFil = cValFil & " AND "
                '
                Select Case nValFil
                Case 0: cValFil = cValFil & oColLis.DataField & " LIKE  '" & oColLis.FilterText & "*'"
                Case 1: cValFil = cValFil & oColLis.DataField & " LIKE '*" & oColLis.FilterText & "*'"
                End Select
            End If
        Else
        End If
    Next oColLis
    '
    Filtrar = cValFil
End Function

Private Sub DesFiltrar()
    For Each oColLis In TdbLista.Columns
        oColLis.FilterText = ""
    Next oColLis
    '
    TxtDato(0) = ""
    TxtDato(1) = ""
    TxtDato(2) = ""
    '
    RstTemp.Filter = adFilterNone
    '
    TxtDato(2).SetFocus
End Sub

Private Sub CambiarFiltro()
    If TlbBarra.Tools.ITEM("Filtro").ComboBox.ListIndex = 0 Then
        TlbBarra.Tools.ITEM("Filtro").ComboBox.ListIndex = 1
    Else
        TlbBarra.Tools.ITEM("Filtro").ComboBox.ListIndex = 0
    End If
    '
    'If CmdFiltro.Caption = "B" Then
    '    CmdFiltro.Caption = "A"
    'Else
    '    CmdFiltro.Caption = "B"
    'End If
    '
    Call DesFiltrar
End Sub

Private Sub Seleccionar()
    wcodproducto = TdbLista.Columns(0) & ""
    wcodfab = TdbLista.Columns(1) & ""
    wdesproducto = TdbLista.Columns(2) & ""
    wmedida = TdbLista.Columns(4) & ""
    wstockact = Val(TdbLista.Columns(5) & "")
    wprecos = Val(TdbLista.Columns(6) & "")
    '
    Unload Me
End Sub


'Option Explicit
'Dim rs  As New ADODB.Recordset
'Dim col As TrueOleDBGrid70.Column
'Dim cols As TrueOleDBGrid70.Columns
'
'Private Sub cmbfiltro_Change()
'If cmbfiltro.ListIndex = 0 Then DataGrid.col = 0
'If cmbfiltro.ListIndex = 1 Then DataGrid.col = 1
'If cmbfiltro.ListIndex = 2 Then DataGrid.col = 2
'If cmbfiltro.ListIndex = 3 Then DataGrid.col = 3
'mnufiltroavanz_Click
'End Sub
'
'
'Private Sub cmbfiltro_LostFocus()
'
'cmbfiltro_Change
'mnufiltroavanz_Click
'
'
'End Sub
'
'Private Sub cmdnuevo_Click()
'
'    sw_load_mant = True
'    sw_nuevo_mant = True
'    mant_productos.Show 1
'    sw_nuevo_mant = False
'    sw_load_mant = False
'    DataAyuda.Refresh
'
'End Sub
'
'Private Sub DataGrid_DblClick()
'
'    DataGrid_KeyDown 13, 0
'
'End Sub
'
'Private Sub DataGrid_KeyDown(KeyCode As Integer, Shift As Integer)
'
'    Select Case KeyCode
'        Case 13:
'            wcodproducto = DataGrid.Columns(0) & ""
'            wcodfab = DataGrid.Columns(1) & ""
'            wdesproducto = DataGrid.Columns(2) & ""
'            wmedida = DataGrid.Columns(4) & ""
'            wstockact = Val(DataGrid.Columns(5) & "")
'            wprecos = Val(DataGrid.Columns(6) & "")
'            If Len(Trim(wcod_alm)) = 0 Then
'                DataAyuda.RecordSource = "Select A.f5codpro,A.F5NOMPRO,A.F5CODFAB,A.F5valvta,B.F7SIGMED FROM IF5PLA AS A,EF7MEDIDAS AS B WHERE A.F7CODMED=B.F7CODMED ORDER BY A.F5CODPRO"
'            Else
'                If wtipoguia = "I" Then
'                    DataAyuda.RecordSource = "Select A.f5codpro,A.F5CODFAB,A.F5NOMPRO,A.F5valvta,C.F7SIGMED FROM IF5PLA AS A,IF6ALMA AS B,EF7MEDIDAS AS C WHERE A.F5CODPRO=B.F5CODPRO AND B.F2CODALM='" & wcod_alm & "' AND A.F7CODMED=C.F7CODMED ORDER BY A.F5CODPRO"
'                Else   '----- VALES DE SALIDA
'                    DataAyuda.RecordSource = "Select A.f5codpro,A.F5CODFAB,A.F5NOMPRO,A.F5valvta,C.F7SIGMED FROM IF5PLA AS A,IF6ALMA AS B,EF7MEDIDAS AS C WHERE A.F5CODPRO=B.F5CODPRO AND B.F2CODALM='" & wcod_alm & "' AND A.F7CODMED=C.F7CODMED AND B.F6STOCKACT>0 ORDER BY A.F5CODPRO"
'                End If
'            End If
'            DataAyuda.Refresh
'            --------------------------
'            For Each col In DataGrid.Columns
'                col.FilterText = ""
'            Next col
'            DataAyuda.Recordset.Filter = adFilterNone
'            --------------------------
'            Me.Hide
'        Case 27:
'            wcodproducto = ""
'            wcodfab = ""
'            wdesproducto = ""
'            wmedida = ""
'            wstockact = 0#
'            wprecos = 0#
'            Me.Hide
'        Case 45:
'            sw_load_mant = True
'            sw_nuevo_mant = True
'            mant_productos.Show 1
'            sw_nuevo_mant = False
'            sw_load_mant = False
'            DataAyuda.Refresh
'        Case 114:
'            Me.DataGrid.col = 2
'            mnufiltroavanz_Click
'        Case 115:
'            cmbfiltro.Visible = False
'            txtope.Visible = False
'            mnutodos_Click
'    End Select
'
'End Sub
'
'Private Sub datagrid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'    Label1.Visible = True
'
'End Sub
'
'Private Sub datagrid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'    If txtope.Visible = True Then txtope.Visible = False
'    If lblbusca.Visible = True Then lblbusca.Visible = False
'    mnufiltro.Caption = "Filtrar [" + DataGrid.Columns(DataGrid.col).Text + "]"
'    Select Case Button
'        Case 2
'            PopupMenu MNUPRI
'    End Select
'
'End Sub
'
'Private Sub Form_Activate()
'
'    DataGrid.col = 2
'    If txtope.Visible = True Then
'        lblbusca.Visible = True
'        lblbusca.Caption = Trim(DataGrid.Columns(2).Caption)
'        cmbfiltro.Visible = True
'        txtope.Text = ""
'    End If
'
'End Sub
'
'Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'
'    Select Case KeyCode
'        Case 27:
'            wcodproducto = ""
'            wcodfab = ""
'            wdesproducto = ""
'            wmedida = ""
'            wstockact = 0#
'            wprecos = 0#
'            Me.Hide
'        Case 45:
'            sw_load_mant = True
'            sw_nuevo_mant = True
'            mant_productos.Show 1
'            sw_nuevo_mant = False
'            sw_load_mant = False
'            DataAyuda.Refresh
'    End Select
'
'End Sub
'
'Private Sub Form_Load()
'
'    DataGrid.OddRowStyle.BackColor = &HC0FFFF
'    DataGrid.EvenRowStyle.BackColor = &HFFFFFF
'    DataGrid.HighlightRowStyle.BackColor = vbActiveTitleBar
'    DataGrid.HighlightRowStyle.ForeColor = vbWhite
'    DataGrid.AlternatingRowStyle = True
'
'    cmbfiltro.AddItem "Codigo", 0
'    cmbfiltro.AddItem "Cod. Fabricante", 1
'    cmbfiltro.AddItem "Descripción", 2
'    cmbfiltro.AddItem "Unidad", 3
'    cmbfiltro.ListIndex = 2
'
'    If wf1visualiza_precio_hlp = "F" Then
'        DataGrid.Columns(3).Visible = False
'    End If
'
'    DataAyuda.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source=" & wrutabancos & "\inventa.mdb"
'    If Len(Trim(wcod_alm)) = 0 Then
'        DataAyuda.RecordSource = "Select A.f5codpro,A.F5NOMPRO,A.F5CODFAB,A.F5valvta,B.F7SIGMED FROM IF5PLA AS A,EF7MEDIDAS AS B WHERE A.F7CODMED=B.F7CODMED ORDER BY A.F5NOMPRO"
'    Else
'        If wtipoguia = "I" Then
'            DataAyuda.RecordSource = "Select A.f5codpro,A.F5CODFAB,A.F5NOMPRO,A.F5valvta,C.F7SIGMED FROM IF5PLA AS A,IF6ALMA AS B,EF7MEDIDAS AS C WHERE A.F5CODPRO=B.F5CODPRO AND B.F2CODALM='" & wcod_alm & "' AND A.F7CODMED=C.F7CODMED ORDER BY A.F5CODPRO"
'        Else   '----- VALES DE SALIDA
'            DataAyuda.RecordSource = "Select A.f5codpro,A.F5CODFAB,A.F5NOMPRO,A.F5valvta,C.F7SIGMED FROM IF5PLA AS A,IF6ALMA AS B,EF7MEDIDAS AS C WHERE A.F5CODPRO=B.F5CODPRO AND B.F2CODALM='" & wcod_alm & "' AND A.F7CODMED=C.F7CODMED AND B.F6STOCKACT>0 ORDER BY A.F5CODPRO"
'            DataAyuda.RecordSource = "Select A.f5codpro,A.F5CODFAB,A.F5NOMPRO,A.F5valvta,C.F7SIGMED FROM IF5PLA AS A,IF6ALMA AS B,EF7MEDIDAS AS C WHERE A.F5CODPRO=B.F5CODPRO AND B.F2CODALM='" & wcod_alm & "' AND A.F7CODMED=C.F7CODMED ORDER BY A.F5CODPRO"  'Giannina
'        End If
'    End If
'    DataAyuda.Refresh
'
'End Sub
'
'Private Sub mnufiltro_Click()
'
'    Set rs = DataAyuda.Recordset
'    Select Case DataAyuda.Recordset.Fields(DataGrid.Columns(DataGrid.col).DataField).Type
'        Case 10
'            rs.Filter = "[" + DataGrid.Columns(DataGrid.col).DataField + "]" + " = '" + Trim("" & DataGrid.Columns(DataGrid.col).Text) + "'"
'        Case 4
'            rs.Filter = "[" + DataGrid.Columns(DataGrid.col).DataField + "]" + " = " + DataGrid.Columns(DataGrid.col).Text
'        Case 8
'            If IsDate(DataGrid.Columns(DataGrid.col).Text) Then
'                rs.Filter = "[" + DataGrid.Columns(DataGrid.col).DataField + "]" + "=#" + DataGrid.Columns(DataGrid.col).Text + "#"
'            Else
'                MsgBox "Ingrese una Fecha Valida..!", 32, "Advertencia"
'                Exit Sub
'            End If
'    End Select
'    Set DataAyuda.Recordset = rs.DataSource
'    Set rs = Nothing
'
'End Sub
'
'Private Sub mnufiltroavanz_Click()
'
'    Select Case DataGrid.col
'        Case 0
'            lblbusca.Visible = True
'            lblbusca.Caption = Trim(DataGrid.Columns(0).Caption)
'            cmbfiltro.Visible = True
'            txtope.Visible = True
'            txtope.Text = ""
'            txtope.SetFocus
'        Case 1
'            lblbusca.Visible = True
'            lblbusca.Caption = Trim(DataGrid.Columns(1).Caption)
'            cmbfiltro.Visible = True
'            txtope.Visible = True
'            txtope.Text = ""
'            txtope.SetFocus
'         Case 2
'            lblbusca.Visible = True
'            lblbusca.Caption = Trim(DataGrid.Columns(2).Caption)
'            cmbfiltro.Visible = True
'            cmbfiltro.ListIndex = 2
'            txtope.Visible = True
'            txtope.Text = ""
'            txtope.SetFocus
'    End Select
'
'End Sub
'
'Private Sub mnuordasc_Click()
'
'    Set rs = DataAyuda.Recordset
'    rs.Sort = "[" + DataGrid.Columns(DataGrid.col).DataField + "] Asc"
'    Set DataAyuda.Recordset = rs.DataSource
'    Set rs = Nothing
'
'End Sub
'
'Private Sub mnuorddesc_Click()
'
'    Set rs = DataAyuda.Recordset
'    rs.Sort = "[" + DataGrid.Columns(DataGrid.col).DataField + "] Desc"
'    Set DataAyuda.Recordset = rs.DataSource
'    Set rs = Nothing
'
'End Sub
'
'Private Sub mnutodos_Click()
'
'    If Len(Trim(wcod_alm)) = 0 Then
'        DataAyuda.RecordSource = "Select A.f5codpro,A.F5NOMPRO,A.F5CODFAB,A.F5valvta,B.F7SIGMED FROM IF5PLA AS A,EF7MEDIDAS AS B WHERE A.F7CODMED=B.F7CODMED ORDER BY A.F5NOMPRO"
'    Else
'        If wtipoguia = "I" Then
'            DataAyuda.RecordSource = "Select A.f5codpro,A.F5CODFAB,A.F5NOMPRO,A.F5valvta,C.F7SIGMED FROM IF5PLA AS A,IF6ALMA AS B,EF7MEDIDAS AS C WHERE A.F5CODPRO=B.F5CODPRO AND B.F2CODALM='" & wcod_alm & "' AND A.F7CODMED=C.F7CODMED ORDER BY A.F5CODPRO"
'        Else   '----- VALES DE SALIDA
'            DataAyuda.RecordSource = "Select A.f5codpro,A.F5CODFAB,A.F5NOMPRO,A.F5valvta,C.F7SIGMED FROM IF5PLA AS A,IF6ALMA AS B,EF7MEDIDAS AS C WHERE A.F5CODPRO=B.F5CODPRO AND B.F2CODALM='" & wcod_alm & "' AND A.F7CODMED=C.F7CODMED AND B.F6STOCKACT>0 ORDER BY A.F5CODPRO"
'        End If
'    End If
'    DataAyuda.Refresh
'
'End Sub
'
'Private Sub SSPanel1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'    Label1.Visible = False
'
'End Sub
'
'Private Sub txtope_KeyPress(KeyAscii As Integer)
'Dim SQL     As String
'Dim DATO    As String
'
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'    If KeyAscii = 13 Then
'        If Len(Trim(txtope.Text)) > 0 Then
'            If Len(txtope.Text) = 0 Then txtope.SetFocus: Exit Sub
'            txtope.Text = "*" & txtope.Text
'            If InStr(txtope, "*") = Len(Trim(txtope)) Then
'                DATO = Left(txtope, Len(Trim(txtope)) - 1)
'            Else
'                DATO = txtope.Text
'            End If
'            txtope.Text = ""
'            Set rs = DataAyuda.Recordset
'            cmbfiltro_Change
'            rs.Filter = "[" + DataGrid.Columns(DataGrid.col).DataField + "]" + " Like  '" + DATO + "*'"
'            If rs.EOF Then txtope.SetFocus: Exit Sub Else txtope.Visible = False: cmbfiltro.Visible = False: lblbusca.Visible = False
'            Set DataAyuda.Recordset = rs.DataSource
'            Set rs = Nothing
'            DataGrid.SetFocus
'        End If
'    End If
'
'End Sub
'
'Private Sub DataGrid_FilterChange()
'
'On Error GoTo errhandler
'Set cols = DataGrid.Columns
'Dim c As Integer
'
'    c = DataGrid.col
'    DataGrid.HoldFields
'    DataAyuda.Recordset.Filter = getFilter()
'    DataGrid.col = c
'    DataGrid.EditActive = True
'    Exit Sub
'
'errhandler:
'
'    MsgBox Err.Source & ":" & vbCrLf & Err.Description
'    For Each col In DataGrid.Columns
'        col.FilterText = ""
'    Next col
'    DataAyuda.Recordset.Filter = adFilterNone
'
'End Sub
'
'Private Function getFilter() As String
'Dim cadena As String
'Dim n As Integer
'
'    For Each col In cols
'        If Trim(col.FilterText) <> "" Then
'            n = n + 1
'            If n > 1 Then
'                cadena = cadena & " AND "
'            End If
'            cadena = cadena & col.DataField & " LIKE '" & col.FilterText & "*'"
'        End If
'    Next col
'    getFilter = cadena
'
'End Function
