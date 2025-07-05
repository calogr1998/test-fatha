VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmayucen 
   Appearance      =   0  'Flat
   Caption         =   "Centros de Costo"
   ClientHeight    =   4770
   ClientLeft      =   555
   ClientTop       =   2145
   ClientWidth     =   10470
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
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4770
   ScaleWidth      =   10470
   Begin Threed.SSPanel SSPanel1 
      Height          =   4695
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   10320
      _Version        =   65536
      _ExtentX        =   18203
      _ExtentY        =   8281
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
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Simplifique la consulta digitando el Simbolo (*)"
         Top             =   4230
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
         Left            =   450
         Options         =   0
         ReadOnly        =   -1  'True
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "temp_hlpcentros"
         Top             =   1125
         Visible         =   0   'False
         Width           =   1155
      End
      Begin TrueOleDBGrid70.TDBGrid datagrid 
         Bindings        =   "FAYUCEN.frx":0000
         Height          =   3480
         Left            =   90
         TabIndex        =   1
         Top             =   90
         Width           =   10140
         _ExtentX        =   17886
         _ExtentY        =   6138
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Código"
         Columns(0).DataField=   "F3COSTO"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Descripción"
         Columns(1).DataField=   "F3DESCRIP"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Cliente"
         Columns(2).DataField=   "F3NOMCLI"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   2
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   2
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2355"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2275"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=65808"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=8229"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=8149"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=65808"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=6826"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=6747"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=65808"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
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
         _StyleDefs(40)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=44"
         _StyleDefs(41)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=45"
         _StyleDefs(42)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=47"
         _StyleDefs(43)  =   "Splits(0).Columns(2).Style:id=58,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(44)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=44"
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "[F3] ---> Filtro Avanzado: Descripción"
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
         Left            =   120
         TabIndex        =   5
         Top             =   3600
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
         Height          =   375
         Left            =   7680
         TabIndex        =   4
         Top             =   4140
         Visible         =   0   'False
         Width           =   2625
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
         Left            =   120
         TabIndex        =   3
         Top             =   3960
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
Attribute VB_Name = "frmayucen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbbase          As DAO.Database
Dim tbbase          As DAO.Recordset
Dim tbtempo         As DAO.Recordset
Dim dbtempo         As DAO.Database
Dim tbclie          As DAO.Recordset
Dim dbclie          As DAO.Database
Dim tbresponsable   As DAO.Recordset
Dim rs              As ADODB.Recordset

Private Sub DataGrid_DblClick()

  DataGrid_KeyPress 13

End Sub
Private Sub DataGrid_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 114 Then
         DataGrid.col = 1
         mnufiltroavanz_Click
    End If
    
    DataGrid_KeyPress KeyCode

End Sub

Private Sub DataGrid_KeyPress(KeyAscii As Integer)
   
    Select Case KeyAscii
        Case 13:
            gcodcen = "" & DataAyuda.Recordset.Fields("F3COSTO")
            gnomcen = "" & DataAyuda.Recordset.Fields("F3DESCRIP")
            Unload Me
        Case 27:
            gcodcen = " "
            gnomcen = ""
            Unload Me
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

    DataGrid.EvenRowStyle.BackColor = &HFFFFFF
    DataGrid.OddRowStyle.BackColor = &HC0FFFF
    DataGrid.HighlightRowStyle.BackColor = vbActiveTitleBar
    DataGrid.HighlightRowStyle.ForeColor = vbWhite
    DataGrid.AlternatingRowStyle = True

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then Unload Me

End Sub

Private Sub Form_Load()
        
    Set dbbase = OpenDatabase(wrutaconta & "\db_costo.mdb")
    Set tbbase = dbbase.OpenRecordset("centros")

    Set dbclie = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set tbclie = dbclie.OpenRecordset("ef2clientes")
    Set tbresponsable = dbclie.OpenRecordset("ef2vendedores")

    Set dbtempo = OpenDatabase(wrutatemp & "\temp_com.mdb")
    dbtempo.Execute ("delete * from temp_hlpcentros")
    Set tbtempo = dbtempo.OpenRecordset("temp_hlpcentros")
    
    If tbbase.RecordCount > 0 Then
        tbbase.MoveFirst
        Do While Not tbbase.EOF
            tbtempo.AddNew
            tbtempo.Fields("f3costo") = tbbase.Fields("f3costo") & ""
            tbtempo.Fields("f3descrip") = tbbase.Fields("f3descrip") & ""
            tbtempo.Fields("f3codcli") = tbbase.Fields("f3codcli") & ""
            '-----------------------
            tbclie.Index = "idcodcli"
            tbclie.Seek "=", tbbase.Fields("f3codcli") & ""
            If Not tbclie.NoMatch Then
                tbtempo.Fields("f3nomcli") = tbclie.Fields("f2nomcli") & ""
            End If
            '-----------------------
            tbresponsable.Index = "idcodigo"
            tbresponsable.Seek "=", tbbase.Fields("f3responsable") & ""
            If Not tbresponsable.NoMatch Then
                tbtempo.Fields("f3responsable") = tbresponsable.Fields("f2nomven") & ""
            End If
            '-----------------------
            tbtempo.Update
            tbbase.MoveNext
        Loop

        tbtempo.Close
        tbbase.Close
        tbclie.Close
        tbresponsable.Close
        dbtempo.Close
        dbbase.Close
        dbclie.Close
    Else
        MsgBox "No se ingresaron centros de costos.", 48, "Compras"
        Unload Me
    End If
        
    DataAyuda.DatabaseName = wrutatemp & "\temp_com.mdb"
    DataAyuda.RecordSource = "Select * FROM temp_hlpcentros ORDER BY F3COSTO"
    DataAyuda.Refresh
     

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
        Case 0:
            lblbusca.Visible = True
            lblbusca.Caption = DataGrid.Columns(DataGrid.col).Caption
            txtope.Visible = True
            txtope.Text = ""
            txtope.SetFocus
        Case 1
            lblbusca.Visible = True
            lblbusca.Caption = DataGrid.Columns(DataGrid.col).Caption
            txtope.Visible = True
            txtope.Text = ""
            txtope.SetFocus
        Case 2
            lblbusca.Visible = True
            lblbusca.Caption = DataGrid.Columns(DataGrid.col).Caption
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

Private Sub MnuTodo_Click()
DataAyuda.DatabaseName = wrutatemp & "\temp_com.mdb"
DataAyuda.RecordSource = "Select * FROM temp_hlpcentros "
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
        Set DataAyuda.Recordset = rs.DataSource
        Set rs = Nothing
        DataGrid.SetFocus
    End If

End Sub
