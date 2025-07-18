VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form hlp_uupp 
   Caption         =   "Ayuda de UUPP"
   ClientHeight    =   4335
   ClientLeft      =   2595
   ClientTop       =   2370
   ClientWidth     =   8145
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   8145
   Begin Threed.SSPanel SSPanel1 
      Height          =   4185
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   7980
      _Version        =   65536
      _ExtentX        =   14076
      _ExtentY        =   7382
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
      Begin VB.ComboBox cmbfiltro 
         Height          =   330
         Left            =   105
         TabIndex        =   6
         Top             =   3780
         Visible         =   0   'False
         Width           =   1380
      End
      Begin MSAdodcLib.Adodc Data1 
         Height          =   330
         Left            =   3255
         Top             =   2100
         Visible         =   0   'False
         Width           =   2220
         _ExtentX        =   3916
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
         Caption         =   "Data1"
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
      Begin VB.TextBox txtope 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1575
         TabIndex        =   1
         ToolTipText     =   "Simplifique la consulta digitando el Simbolo (*)"
         Top             =   3780
         Visible         =   0   'False
         Width           =   2445
      End
      Begin TrueOleDBGrid70.TDBGrid Grid1 
         Bindings        =   "hlp_uupp.frx":0000
         Height          =   2985
         Left            =   135
         TabIndex        =   4
         Top             =   90
         Width           =   7710
         _ExtentX        =   13600
         _ExtentY        =   5265
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Código"
         Columns(0).DataField=   "F4LOCALIDAD"
         Columns(0).DataWidth=   255
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Descripción"
         Columns(1).DataField=   "F4DESCLOC"
         Columns(1).DataWidth=   255
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   2
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   2
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1773"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1693"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=66080"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=11165"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=11086"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).AllowSizing=0"
         Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=66080"
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=144,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(33)  =   "Splits(0).Columns(0).Style:id=28,.parent=43,.alignment=0,.valignment=1,.locked=0"
         _StyleDefs(34)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=44,.alignment=2"
         _StyleDefs(35)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=45"
         _StyleDefs(36)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=47"
         _StyleDefs(37)  =   "Splits(0).Columns(1).Style:id=32,.parent=43,.alignment=0,.valignment=1,.locked=0"
         _StyleDefs(38)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=44,.alignment=2"
         _StyleDefs(39)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=45"
         _StyleDefs(40)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=47"
         _StyleDefs(41)  =   "Named:id=33:Normal"
         _StyleDefs(42)  =   ":id=33,.parent=0"
         _StyleDefs(43)  =   "Named:id=34:Heading"
         _StyleDefs(44)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(45)  =   ":id=34,.wraptext=-1"
         _StyleDefs(46)  =   "Named:id=35:Footing"
         _StyleDefs(47)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(48)  =   "Named:id=36:Selected"
         _StyleDefs(49)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(50)  =   "Named:id=37:Caption"
         _StyleDefs(51)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(52)  =   "Named:id=38:HighlightRow"
         _StyleDefs(53)  =   ":id=38,.parent=33,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
         _StyleDefs(54)  =   "Named:id=39:EvenRow"
         _StyleDefs(55)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(56)  =   "Named:id=40:OddRow"
         _StyleDefs(57)  =   ":id=40,.parent=33"
         _StyleDefs(58)  =   "Named:id=41:RecordSelector"
         _StyleDefs(59)  =   ":id=41,.parent=34"
         _StyleDefs(60)  =   "Named:id=42:FilterBar"
         _StyleDefs(61)  =   ":id=42,.parent=33"
      End
      Begin VB.Label Label3 
         Caption         =   "[F4] ---> Todos"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   3570
         TabIndex        =   7
         Top             =   3150
         Width           =   1275
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "[F3] ---> Filtro Avanzado: Descripción"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   135
         TabIndex        =   5
         Top             =   3105
         Width           =   2715
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
         TabIndex        =   3
         Top             =   3330
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Label Label1 
         Caption         =   "Ayuda: Hacer Click Derecho en la descripción que desea buscar"
         ForeColor       =   &H00800000&
         Height          =   510
         Left            =   5220
         TabIndex        =   2
         Top             =   3240
         Visible         =   0   'False
         Width           =   2625
      End
   End
   Begin VB.Menu MnuPri 
      Caption         =   ""
      Begin VB.Menu MnuFiltro 
         Caption         =   "Filtrar"
      End
      Begin VB.Menu MnuFiltroavanz 
         Caption         =   "Filtro Avanzado:"
      End
      Begin VB.Menu MnuOrdAsc 
         Caption         =   "Ord. Asc"
      End
      Begin VB.Menu MnuOrdDesc 
         Caption         =   "Ord. Desc"
      End
      Begin VB.Menu MnuTodo 
         Caption         =   "Mostrar Todos"
      End
   End
End
Attribute VB_Name = "hlp_uupp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs      As New ADODB.Recordset
Dim col     As TrueOleDBGrid70.Column
Dim cols    As TrueOleDBGrid70.Columns

Private Sub cmbfiltro_Change()
If cmbfiltro.ListIndex = 0 Then Grid1.col = 0
If cmbfiltro.ListIndex = 1 Then Grid1.col = 1
mnufiltroavanz_Click

End Sub

Private Sub cmbfiltro_LostFocus()
cmbfiltro_Change
mnufiltroavanz_Click
End Sub

Private Sub Form_Activate()
    
    Grid1.col = 1
    If txtope.Visible = True Then
        lblbusca.Visible = True
        lblbusca.Caption = Trim(Grid1.Columns(1).Caption)
        cmbfiltro.Visible = True
        txtope.Text = ""
        txtope.SetFocus
    End If
    
    
    Grid1.EvenRowStyle.BackColor = &HFFFFFF
    Grid1.OddRowStyle.BackColor = &HC0FFFF
    Grid1.HighlightRowStyle.BackColor = vbActiveTitleBar
    Grid1.HighlightRowStyle.ForeColor = vbWhite
    Grid1.AlternatingRowStyle = True

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    evalua KeyCode
    Select Case KeyCode
        Case 27:
            wcodlocalidad = "": wdeslocalidad = ""
            Unload Me
    End Select
    
End Sub

Private Sub Form_Load()
    
        
    cmbfiltro.AddItem "Codigo", 0
    cmbfiltro.AddItem "Descripción", 1
    cmbfiltro.ListIndex = 0
    
    Data1.ConnectionString = cnn_dbbancos
    Data1.RecordSource = "SELECT * FROM IF6DOCUM ORDER BY F4LOCALIDAD"
    Data1.Refresh
    
End Sub

Private Sub Grid1_DblClick()

   Grid1_KeyDown 13, 0
   
End Sub

Private Sub Grid1_FilterChange()
Set cols = Grid1.Columns
Dim c As Integer

    c = Grid1.col
    Grid1.HoldFields
    Data1.Recordset.Filter = getFilter()
    Grid1.col = c
    Grid1.EditActive = True
    Exit Sub

End Sub

Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case 13:
            wcodlocalidad = Grid1.Columns(0)
            wdeslocalidad = Grid1.Columns(1)
            Unload Me
        Case 27:
            wcodlocalidad = "": wdeslocalidad = ""
            Unload Me
        Case 114:
            Grid1.col = 1
            mnufiltroavanz_Click
    End Select
    
End Sub

Private Sub Grid1_KeyUp(KeyCode As Integer, Shift As Integer)
evalua KeyCode
End Sub

Private Sub Grid1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Label1.Visible = True
 
End Sub

Private Sub Grid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
    If cmbfiltro.Visible = True Then cmbfiltro.Visible = False
    If txtope.Visible = True Then txtope.Visible = False
    If lblbusca.Visible = True Then lblbusca.Visible = False

    mnufiltro.Caption = "Filtrar [" + Grid1.Columns(Grid1.col).Text + "]"
    Select Case Button
        Case 2
            PopupMenu mnupri
    End Select
    
End Sub

Private Sub mnufiltro_Click()

    Set rs = Data1.Recordset
    Select Case Data1.Recordset.Fields(Grid1.Columns(Grid1.col).DataField).Type
        Case 10
            rs.Filter = "[" + Grid1.Columns(Grid1.col).DataField + "]" + " = '" + Trim("" & Grid1.Columns(Grid1.col).Text) + "'"
        Case 4
            rs.Filter = "[" + Grid1.Columns(Grid1.col).DataField + "]" + " = " + Grid1.Columns(Grid1.col).Text
        Case 8
            If IsDate(Grid1.Columns(Grid1.col).Text) Then
                rs.Filter = "[" + Grid1.Columns(Grid1.col).DataField + "]" + "=#" + Grid1.Columns(Grid1.col).Text + "#"
            Else
                MsgBox "Ingrese una Fecha Valida..!", 32, "Advertencia"
                Exit Sub
            End If
    End Select
    Set Data1.Recordset = rs.DataSource
    Set rs = Nothing
    
End Sub

Private Sub mnufiltroavanz_Click()
    
    Select Case Grid1.col
        Case 0:
            lblbusca.Visible = True
            lblbusca.Caption = Grid1.Columns(Grid1.col).Caption
            cmbfiltro.Visible = True
            cmbfiltro.ListIndex = 0
            txtope.Visible = True
            txtope.Text = ""
            txtope.SetFocus
        Case 1
            lblbusca.Visible = True
            lblbusca.Caption = Grid1.Columns(Grid1.col).Caption
            cmbfiltro.Visible = True
            cmbfiltro.ListIndex = 1
            txtope.Visible = True
            txtope.Text = ""
            txtope.SetFocus
        Case 2
            lblbusca.Visible = True
            lblbusca.Caption = Grid1.Columns(Grid1.col).Caption
            cmbfiltro.Visible = True
            cmbfiltro.ListIndex = 2
            txtope.Visible = True
            txtope.Text = ""
            txtope.SetFocus
    End Select
    
End Sub

Private Sub mnuordasc_Click()

    Set rs = Data1.Recordset
    rs.Sort = "[" + Grid1.Columns(Grid1.col).DataField + "] Asc"
    Set Data1.Recordset = rs.DataSource
    Set rs = Nothing
    
End Sub

Private Sub mnuorddesc_Click()

    Set rs = Data1.Recordset
    rs.Sort = "[" + Grid1.Columns(Grid1.col).DataField + "] Desc"
    Set Data1.Recordset = rs.DataSource
    Set rs = Nothing
    
End Sub

Private Sub MnuTodo_Click()
    
    Data1.RecordSource = "SELECT * FROM IF6DOCUM ORDER BY F4LOCALIDAD"
    Data1.Refresh
    
End Sub

Private Sub SSPanel1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Label1.Visible = False
  
End Sub

Private Sub txtope_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        txtope.Text = "*" & txtope.Text
        If Len(txtope) = 0 Then txtope.SetFocus: Exit Sub
        If InStr(txtope, "*") = Len(Trim(txtope)) Then
           DATO = Left(txtope, Len(Trim(txtope)) - 1)
        Else
            DATO = txtope.Text
        End If
        Set rs = Data1.Recordset
        cmbfiltro_Change
        rs.Filter = "[" + Grid1.Columns(Grid1.col).DataField + "]" + " Like  '" + DATO + "*'"
        If rs.EOF Then txtope.SetFocus: Exit Sub Else txtope.Visible = False: cmbfiltro.Visible = False: lblbusca.Visible = False
        Set Data1.Recordset = rs.DataSource
        Set rs = Nothing
        Grid1.SetFocus
    End If
    
End Sub

Private Sub evalua(KeyCode As Integer)
    If KeyCode = 115 Then
        lblbusca.Visible = False
        txtope.Visible = False
        cmbfiltro.Visible = False
        MnuTodo_Click
    End If
End Sub

Private Function getFilter() As String
Dim cadena As String
Dim n As Integer

    For Each col In cols
        If Trim(col.FilterText) <> "" Then
            n = n + 1
            If n > 1 Then
                cadena = cadena & " AND "
            End If
            cadena = cadena & col.DataField & " LIKE '" & col.FilterText & "*'"
        End If
    Next col
    getFilter = cadena

End Function

Private Sub txtope_KeyUp(KeyCode As Integer, Shift As Integer)
evalua KeyCode
End Sub
