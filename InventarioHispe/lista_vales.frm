VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form lista_vales 
   Caption         =   "Vales"
   ClientHeight    =   9015
   ClientLeft      =   480
   ClientTop       =   1725
   ClientWidth     =   14010
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "lista_vales.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9015
   ScaleWidth      =   14010
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkExcluirValeCompromiso 
      Caption         =   "Excluir Vales de Compromiso"
      Height          =   255
      Left            =   3840
      TabIndex        =   9
      Top             =   120
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox txtBusqueda 
      Height          =   315
      Left            =   7560
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   50
      Width           =   6375
   End
   Begin VB.Frame fraProceso 
      Caption         =   " Procesando "
      Height          =   735
      Left            =   4080
      TabIndex        =   5
      Top             =   2880
      Visible         =   0   'False
      Width           =   5535
      Begin ComctlLib.ProgressBar pgbProceso 
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin VB.CheckBox Checkagrupar 
      Caption         =   "Agrupar columnas"
      Height          =   255
      Left            =   1740
      TabIndex        =   3
      Top             =   105
      Width           =   2055
   End
   Begin VB.CheckBox CheckFiltro 
      Caption         =   "Activar Filtro"
      Height          =   255
      Left            =   300
      TabIndex        =   2
      Top             =   105
      Width           =   1455
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   4530
      Left            =   120
      OleObjectBlob   =   "lista_vales.frx":058A
      TabIndex        =   0
      Top             =   405
      Width           =   13785
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid2 
      Height          =   3570
      Left            =   120
      OleObjectBlob   =   "lista_vales.frx":5A20
      TabIndex        =   4
      Top             =   5040
      Width           =   13725
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   9
      ShowShortcutsInToolTips=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Tools           =   "lista_vales.frx":9EA3
      ToolBars        =   "lista_vales.frx":F7E6
   End
   Begin VB.Label Label1 
      Caption         =   "Busqueda"
      Height          =   255
      Left            =   6600
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ctrl+Enter -> Buscar Siguiente  /  Shift+Enter -> Encontrar Nuevo"
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
      Height          =   210
      Left            =   5400
      TabIndex        =   1
      Top             =   90
      Visible         =   0   'False
      Width           =   4650
   End
End
Attribute VB_Name = "lista_vales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Rem SK ADD:
Private strTipoVale As String

Private strFichero As String


Rem Variables para Controlar la Devolucion de Foco del Registro en Grilla señalado antes de alguna Modificacion o Uso
Dim d As Double
Dim nSaveRecNo As Double

Public Property Let TipoVale(ByVal Value As String)
    strTipoVale = Value
End Property

Public Property Get TipoVale() As String
    TipoVale = strTipoVale
End Property




Private Sub listarAnnos()
    objAyudaVale.listarAnnoVale SSActiveToolBars1.Tools.ITEM("ID_Anno").ComboBox, True
    
    If SSActiveToolBars1.Tools.ITEM("ID_Anno").ComboBox.ListCount > 1 Then
        SSActiveToolBars1.Tools.ITEM("ID_Anno").ComboBox.ListIndex = SSActiveToolBars1.Tools.ITEM("ID_Anno").ComboBox.ListCount - 1
    Else
        SSActiveToolBars1.Tools.ITEM("ID_Anno").ComboBox.ListIndex = 0
    End If
End Sub

Private Sub listarMeses()
    ModUtilitario.listarMeses 0, SSActiveToolBars1.Tools.ITEM("ID_Mes").ComboBox, True
    
    If SSActiveToolBars1.Tools.ITEM("ID_Anno").ComboBox.ListIndex = 0 Then
        SSActiveToolBars1.Tools.ITEM("ID_Mes").ComboBox.ListIndex = 0
        SSActiveToolBars1.Tools.ITEM("ID_Mes").Enabled = False
    Else
        SSActiveToolBars1.Tools.ITEM("ID_Mes").ComboBox.ListIndex = ModUtilitario.seleccionarItem(SSActiveToolBars1.Tools.ITEM("ID_Mes").ComboBox, Format(Month(Date), "00"), "DER", 2)
    End If
End Sub

Public Sub listarVale()
    With dxDBGrid1
        .Dataset.Close
        
        .Dataset.ADODataset.ConnectionString = cnn_dbbancos
        
        SqlCad = vbNullString
        SqlCad = SqlCad & "SELECT "
        SqlCad = SqlCad & "A.F2CODALM & A.F4NUMVAL AS ALM_VALE, "
        SqlCad = SqlCad & "A.F2CODALM, "
        SqlCad = SqlCad & "A.F4NUMVAL, "
        SqlCad = SqlCad & "A.NUMENSAM, "
        SqlCad = SqlCad & "A.F4FECVAL, "
        SqlCad = SqlCad & "B.F1NOMORI, "
        SqlCad = SqlCad & "AUXILIARES.NOMBRE, "
        SqlCad = SqlCad & "A.F4SERGUIA, "
        SqlCad = SqlCad & "A.F4NUMGUIA, "
        SqlCad = SqlCad & "DOC.F2ABREV, "
        SqlCad = SqlCad & "A.F4SERDOC, "
        SqlCad = SqlCad & "A.F4NUMDOC, "
        
            Select Case strTipoVale
                Case "I"
                    SqlCad = SqlCad & "A.NUMORDEN AS ORDEN, "
                Case "S"
                    SqlCad = SqlCad & "A.F4NUMORD AS ORDEN, "
            End Select
            
        SqlCad = SqlCad & "A.F4ORDTRA, "
        SqlCad = SqlCad & "A.F4REGCOM "
        
        SqlCad = SqlCad & "FROM "
        SqlCad = SqlCad & "((IF4VALES AS A "
        SqlCad = SqlCad & "LEFT JOIN SF1ORIGENES AS B ON B.F1CODORI = A.F1CODORI) "
        SqlCad = SqlCad & "LEFT JOIN DOCUMENTOS AS DOC ON DOC.F2CODDOC = A.F4TIPDOC) "
        SqlCad = SqlCad & "LEFT JOIN "
        SqlCad = SqlCad & "("
        SqlCad = SqlCad & "SELECT "
        SqlCad = SqlCad & "'C' AS TIPO, "
        SqlCad = SqlCad & "C.F2CODCLI AS CODIGO, "
        SqlCad = SqlCad & "C.F2NOMCLI AS NOMBRE "
        SqlCad = SqlCad & "FROM "
        SqlCad = SqlCad & "EF2CLIENTES AS C "
        SqlCad = SqlCad & "UNION "
        SqlCad = SqlCad & "SELECT "
        SqlCad = SqlCad & "'P' AS TIPO, "
        SqlCad = SqlCad & "P.F2CODPROV AS CODIGO, "
        SqlCad = SqlCad & "P.F2NOMPROV AS NOMBRE "
        SqlCad = SqlCad & "FROM "
        SqlCad = SqlCad & "EF2PROVEEDORES AS P"
        SqlCad = SqlCad & ") AS AUXILIARES "
        SqlCad = SqlCad & "ON AUXILIARES.TIPO = A.F1TIPPRV AND AUXILIARES.CODIGO = A.F2CODPROV "
        
        SqlCad = SqlCad & "WHERE "
        SqlCad = SqlCad & "A.F4TIPOVALE = '" & strTipoVale & "' "
            
            If CBool(chkExcluirValeCompromiso.Value) Then
                SqlCad = SqlCad & "AND A.F1CODORI NOT IN ('XCS') "
            End If
            
            If IsNumeric(SSActiveToolBars1.Tools.ITEM("ID_Anno").ComboBox.Text) Then
                SqlCad = SqlCad & "AND YEAR(A.F4FECVAL) = " & Trim(SSActiveToolBars1.Tools.ITEM("ID_Anno").ComboBox.Text) & " "
            End If
            
            If right(SSActiveToolBars1.Tools.ITEM("ID_Mes").ComboBox.Text, 2) <> "00" Then
                SqlCad = SqlCad & "AND MONTH(A.F4FECVAL) = " & Val(right(SSActiveToolBars1.Tools.ITEM("ID_Mes").ComboBox.Text, 2)) & " "
            End If
            
        SqlCad = SqlCad & "ORDER BY "
        SqlCad = SqlCad & "A.F4FECVAL DESC, A.F2CODALM, A.F4NUMVAL DESC"
        
        '.Columns("0").Visible = False
        .Columns.ColumnByFieldName("ALM_VALE").Visible = False
        
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = SqlCad
        .Dataset.Active = True
        
        .KeyField = "F4NUMVAL"
    End With
    
    SqlCad = vbNullString
    
    dxDBGrid1_RowColChange
End Sub

Public Sub listarValeDetalle(ByVal Codigo As String, almacen As String)
    Dim sql As String
    On Error Resume Next
    With dxDBGrid2
        SqlCad = vbNullString
        SqlCad = SqlCad & "select "
        'SqlCad = SqlCad & "A.F4NUMORD, "
        
        SqlCad = SqlCad & "IIF(CAB.F1CODORI = 'XC0', A.F4NUMORD, '') AS F4NUMORD, "
        
        SqlCad = SqlCad & "IIF(TRIM(A.COD_SOLICITUD) <> '', A.COD_SOLICITUD, 'STOCK LIBRE') AS COD_SOLICITUD, "
        SqlCad = SqlCad & "a.f5codpro, "
        SqlCad = SqlCad & "b.f5nompro, "
        SqlCad = SqlCad & "EF7MEDIDAS.F7SIGMED as F7CODMED, "
        SqlCad = SqlCad & "a.f3canpro, "
        'SqlCad = SqlCad & "a.f3punit, "
        'SqlCad = SqlCad & "a.f3valvta " ',a.partida "
        SqlCad = SqlCad & "IIF(CAB.F4MONEDA = 'S', A.F3VALVTA, A.F3VALDOL) AS COSTO, "
        
        SqlCad = SqlCad & "IIF(CAB.F4MONEDA = 'S', A.F3TOTITE - A.F3IGV, A.F3TOTDOL - A.F3IGVDOL) AS SUBTOTAL, "
        SqlCad = SqlCad & "IIF(CAB.F4MONEDA = 'S', A.F3IGV, A.F3IGVDOL) AS IGV, "
        SqlCad = SqlCad & "IIF(CAB.F4MONEDA = 'S', A.F3TOTITE, A.F3TOTDOL) AS TOTAL "
        
        SqlCad = SqlCad & "FROM "
        SqlCad = SqlCad & "((if3vales AS a "
        SqlCad = SqlCad & "LEFT JOIN IF4VALES AS CAB ON CAB.F2CODALM = A.F2CODALM AND CAB.F4NUMVAL = A.F4NUMVAL) "
        SqlCad = SqlCad & "LEFT JOIN if5pla AS b ON a.F5CODPRO = b.F5CODPRO) "
        SqlCad = SqlCad & "LEFT JOIN EF7MEDIDAS ON b.F7CODMED = EF7MEDIDAS.F7CODMED "
        SqlCad = SqlCad & "where "
        SqlCad = SqlCad & "a.f4numval='" & Codigo & "' AND "
        SqlCad = SqlCad & "a.f2codalm='" & almacen & "' "
        SqlCad = SqlCad & "ORDER BY "
        SqlCad = SqlCad & "b.f5nompro"
        
        .Dataset.Active = False
        .Dataset.ADODataset.ConnectionString = cnn_dbbancos
        .Dataset.ADODataset.CommandText = SqlCad
        .Dataset.Active = True
        .KeyField = "f5codpro"
    End With
End Sub

Private Sub Checkagrupar_Click()
    If Checkagrupar.Value = 1 Then
      dxDBGrid1.Options.Set (egoShowGroupPanel)
    Else
      dxDBGrid1.Options.Unset (egoShowGroupPanel)
    End If
End Sub

Private Sub CheckFiltro_Click()
    If CheckFiltro.Value = 1 Then
      dxDBGrid1.Filter.FilterActive = True
    Else
      dxDBGrid1.Filter.FilterActive = False
    End If
End Sub

Private Sub chkExcluirValeCompromiso_Click()
    listarVale
End Sub

Private Sub dxDBGrid1_OnChangeNode(ByVal OldNode As DXDBGRIDLibCtl.IdxGridNode, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    dxDBGrid1_RowColChange
End Sub

Private Sub dxDBGrid1_RowColChange()
    listarValeDetalle dxDBGrid1.Columns.ColumnByFieldName("f4numval").Value, dxDBGrid1.Columns.ColumnByFieldName("f2codalm").Value
End Sub

Private Sub dxDBGrid1_OnClick()
    dxDBGrid1_RowColChange
End Sub

Private Sub dxDBGrid1_OnDblClick()
    sw_nuevo_documento = False
    
    Me.MousePointer = vbHourglass
    
    Select Case strTipoVale
        Case "I"
'            If ModUtilitario.sGetINI(strFichero, "ConfigCP", "ValeIngresoAbierto", "l") = "1" Then
'                MsgBox "Formulario abierto desde otra Ventana Principal, verifique.", vbInformation + vbOKOnly, App.ProductName
'
'                Me.MousePointer = vbDefault
'
'                Exit Sub
'            End If
            
            If ModUtilitario.validarFormAbierto("vale_ingreso") Then
                Unload vale_ingreso
            End If
            
            For d = 0 To 25
                nSaveRecNo = dxDBGrid1.Dataset.RecNo
            Next
            
            With vale_ingreso
                .Ayuda = False
                .CodigoAlmacen = Trim(dxDBGrid1.Columns.ColumnByFieldName("F2CODALM").Value & "")
                .NumeroVale = Trim(dxDBGrid1.Columns.ColumnByFieldName("F4NUMVAL").Value & "")
                
                .Show 1
                
                'Me.Hide
                listarVale
            End With
            
            If dxDBGrid1.Dataset.RecordCount >= nSaveRecNo Then
                dxDBGrid1.Dataset.RecNo = nSaveRecNo
            End If
        Case "S"
'            If ModUtilitario.sGetINI(strFichero, "ConfigCP", "ValeSalidaAbierto", "l") = "1" Then
'                MsgBox "Formulario abierto desde otra Ventana Principal, verifique.", vbInformation + vbOKOnly, App.ProductName
'
'                Me.MousePointer = vbDefault
'
'                Exit Sub
'            End If
            
            If ModUtilitario.validarFormAbierto("vale_salida") Then
                Unload vale_salida
            End If
            
            For d = 0 To 25
                nSaveRecNo = dxDBGrid1.Dataset.RecNo
            Next
            
            With vale_salida
                .Ayuda = False
                .CodigoAlmacen = Trim(dxDBGrid1.Columns.ColumnByFieldName("F2CODALM").Value & "")
                .NumeroVale = Trim(dxDBGrid1.Columns.ColumnByFieldName("F4NUMVAL").Value & "")
                
                .Show 1
                
                listarVale
            End With
            
            If dxDBGrid1.Dataset.RecordCount >= nSaveRecNo Then
                dxDBGrid1.Dataset.RecNo = nSaveRecNo
            End If
    End Select
    
    Me.MousePointer = vbDefault
End Sub

Private Sub dxDBGrid1_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
    Select Case KeyCode
        Case vbKeyReturn
            dxDBGrid1_OnDblClick
    End Select
End Sub

Private Sub dxDBGrid2_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    Select Case Column.FieldName
        Case "F3CANPRO"
            Select Case Val(Text)
                Case Is < 0
                    FontColor = vbRed
                Case Is > 0
                    FontColor = vbBlue
                Case Is = 0
                    FontColor = vbGreen
            End Select
            
            Text = Format(Text, "#,0.00;(#,0.00)")
    End Select
End Sub

Private Sub Form_Load()
    Me.MousePointer = vbHourglass
    
    Me.left = 1600
    Me.top = 1150
    
    sw_nuevo_documento = True
    
    Select Case strTipoVale
        Case "I"
            Me.Caption = "Vale de Ingreso a Almacen"
        Case "S"
            Me.Caption = "Vale de Salida a Almacen"
    End Select
    
    txtBusqueda.Text = vbNullString
    
    dxDBGrid1.Columns.ColumnByFieldName("F4REGCOM").Visible = IIf(strTipoVale = "I", True, False)
    
    listarAnnos
    
    listarMeses
    
    listarVale
    
    dxDBGrid1.Options.Unset (egoShowGroupPanel)
    dxDBGrid1.Filter.FilterActive = False
    dxDBGrid1_RowColChange
    dxDBGrid2.Dataset.ADODataset.ConnectionString = cnn_dbbancos
    
    'Instanciar ruta de Fichero de Configuracion por Usuario
    strFichero = wrutatemp & strNombreFicheroConfigCPusuario
    
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    On Error Resume Next

    Dim dblPromH As Double

    dblPromH = Me.ScaleHeight / 6
    
    dxDBGrid1.Move 0, dxDBGrid1.top, Me.ScaleWidth, dblPromH * 3.5

    dxDBGrid2.Move 0, dxDBGrid1.top + dxDBGrid1.Height, Me.ScaleWidth, (dblPromH * 2.5) - 400
End Sub

Private Sub Form_Unload(Cancel As Integer)
    wtipoguia = ""
    
    Rem SK ADD:
    strTipoVale = vbNullString
End Sub

Private Sub SSActiveToolBars1_ComboCloseUp(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.ID
        Case "ID_Anno"
            If SSActiveToolBars1.Tools.ITEM("ID_Anno").ComboBox.ListIndex = 0 Then
                SSActiveToolBars1.Tools.ITEM("ID_Mes").ComboBox.ListIndex = 0
                SSActiveToolBars1.Tools.ITEM("ID_Mes").Enabled = False
            Else
                SSActiveToolBars1.Tools.ITEM("ID_Mes").ComboBox.ListIndex = IIf(Val(SSActiveToolBars1.Tools.ITEM("ID_Anno").ComboBox.Text) = Year(Date), Month(Date), 1)
                SSActiveToolBars1.Tools.ITEM("ID_Mes").Enabled = True
            End If
            
            listarVale
        Case "ID_Mes"
            listarVale
    End Select
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    
    Select Case Tool.ID
        Case "ID_Nuevo"
            sw_nuevo_documento = True
            Me.MousePointer = vbHourglass
            
            Select Case strTipoVale
                Case "I"
'                    If ModUtilitario.sGetINI(strFichero, "ConfigCP", "ValeIngresoAbierto", "l") = "1" Then
'                        MsgBox "Formulario abierto desde otra Ventana Principal, verifique.", vbInformation + vbOKOnly, App.ProductName
'
'                        Me.MousePointer = vbDefault
'
'                        Exit Sub
'                    End If
'
                    With vale_ingreso
                        .Ayuda = False
                        .CodigoAlmacen = vbNullString
                        .NumeroVale = vbNullString
                        
                        .Show 1
                    End With
                Case "S"
'                    If ModUtilitario.sGetINI(strFichero, "ConfigCP", "ValeSalidaAbierto", "l") = "1" Then
'                        MsgBox "Formulario abierto desde otra Ventana Principal, verifique.", vbInformation + vbOKOnly, App.ProductName
'
'                        Me.MousePointer = vbDefault
'
'                        Exit Sub
'                    End If
                    
                    With vale_salida
                        .Ayuda = False
                        .CodigoAlmacen = vbNullString
                        .NumeroVale = vbNullString
                        
                        .Show 1
                    End With
            End Select
            
            Me.MousePointer = vbDefault
        Case "Imprimir"
            Dim rpt As New rptValeIngreso
            
            With rpt
                .TipoVale = strTipoVale
                .CodAlmacen = Trim(dxDBGrid1.Columns.ColumnByFieldName("F2CODALM").Value & "")
                .NumeroVale = Trim(dxDBGrid1.Columns.ColumnByFieldName("F4NUMVAL").Value & "")
                
                'ModMilano.abrirCnDBMilano
                
'                .fldCategoria.Text = ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "CT.NOMBRE", "ORDENPRODUCCION AS OP LEFT JOIN CATEGORIATIPO AS CT ON CT.IDCATEGORIATIPO = OP.IDCATEGORIATIPO", "OP.IDORDENPRODUCCION", Val(dxDBGrid1.Columns.ColumnByFieldName("F4ORDTRA").value & ""), "N")
'                .fldOP.Text = ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "OP", "ORDENPRODUCCION", "IDORDENPRODUCCION", Val(dxDBGrid1.Columns.ColumnByFieldName("F4ORDTRA").value & ""), "N")
                
                .Show
            End With
        Case "Importar"
'            Me.MousePointer = vbHourglass
'
'            ModMilano.importarValesServidorExterno strTipoVale, fraProceso, pgbProceso
'
'            listarVale
'
'            dxDBGrid1.Options.Unset (egoShowGroupPanel)
'            dxDBGrid1.Filter.FilterActive = False
'            dxDBGrid1_RowColChange
'
'            Me.MousePointer = vbDefault
            
        Case "ID_Salir"
            Unload Me
    End Select
End Sub

Private Sub dxDBGrid1_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    With dxDBGrid1.Dataset
        If dxDBGrid1.Columns.FocusedColumn.ColumnType = gedLookupEdit Then
            If .State = dsEdit Then
                dxDBGrid1.m.HideEditor
                .Post
                .DisableControls
                .Close
                .Open
                .EnableControls
            End If
        End If
    End With
End Sub

Private Sub txtBusqueda_GotFocus()
    ModUtilitario.seleccionarTextoCaja txtBusqueda
End Sub

Private Sub txtbusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            With dxDBGrid1.Dataset
                .Filtered = True
                .Filter = "F4NUMVAL LIKE '*" & txtBusqueda.Text & "*' OR " & _
                            "NUMENSAM LIKE '*" & txtBusqueda.Text & "*' OR " & _
                            "F4NUMGUIA LIKE '*" & txtBusqueda.Text & "*' OR " & _
                            "F4NUMDOC LIKE '*" & txtBusqueda.Text & "*' OR " & _
                            "NOMBRE LIKE '*" & txtBusqueda.Text & "*' OR " & _
                            "F1NOMORI LIKE '*" & txtBusqueda.Text & "*' OR " & _
                            "ORDEN LIKE '*" & txtBusqueda.Text & "*'"
                
                If Len(Trim(txtBusqueda.Text)) = 0 Then
                    .Filtered = False
                End If
            End With
    End Select
End Sub
