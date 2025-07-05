VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmRptResCompromisoAuto 
   Caption         =   "Reporte de Compromisos Pendientes para su Compra o Reposición"
   ClientHeight    =   8715
   ClientLeft      =   450
   ClientTop       =   1815
   ClientWidth     =   13665
   Icon            =   "frmRptResCompromisoAuto.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   13665
   WindowState     =   2  'Maximized
   Begin VB.Frame fraAgruparPor 
      Caption         =   " Datos de Proceso Automatico "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8520
      TabIndex        =   3
      Top             =   120
      Width           =   5055
      Begin VB.Label lblFechaCorte 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fec. Corte"
         Height          =   255
         Left            =   2520
         TabIndex        =   12
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label lblHoraInicio 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hora de Inicio"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label lblHoraFin 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hora de Fin"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   2295
      End
   End
   Begin MSComDlg.CommonDialog cmdlgReporte 
      Left            =   0
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraPeriodo 
      Caption         =   " Datos de Consulta "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8295
      Begin VB.TextBox txtCodProducto 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "Text1"
         ToolTipText     =   "Seleccione Producto (F2)"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtProducto 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3960
         TabIndex        =   15
         Text            =   "Text1"
         ToolTipText     =   "Ingrese cadena a buscar"
         Top             =   1320
         Width           =   4215
      End
      Begin VB.TextBox txtBusqueda 
         Height          =   285
         Left            =   5160
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   600
         Width           =   3015
      End
      Begin VB.ComboBox cmbCliente 
         Height          =   315
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   240
         Width           =   3015
      End
      Begin VB.ComboBox cmbAnno 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox cmbMes 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox cmbFechaEjecucion 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Producto"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2400
         TabIndex        =   17
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Filtrar"
         Height          =   255
         Left            =   4560
         TabIndex        =   14
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Cliente"
         Height          =   255
         Left            =   4560
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Periodo"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   615
      End
   End
   Begin ActiveToolBars.SSActiveToolBars tlbReporte 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   11
      ShowShortcutsInToolTips=   -1  'True
      Tools           =   "frmRptResCompromisoAuto.frx":058A
      ToolBars        =   "frmRptResCompromisoAuto.frx":A9D1
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dbgReporte 
      Height          =   7440
      Left            =   120
      OleObjectBlob   =   "frmRptResCompromisoAuto.frx":AB27
      TabIndex        =   0
      Top             =   1200
      Width           =   13410
   End
End
Attribute VB_Name = "frmRptResCompromisoAuto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private intOpcionResumen As Integer

Private Sub listarAnnosSQL()
    Dim rstAnnoPA As New ADODB.Recordset
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "YEAR(FECHAEJECUCION) AS ANNO "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "SF4COMPROMISOAUTOMATICO "
    SqlCad = SqlCad & "GROUP BY "
    SqlCad = SqlCad & "YEAR(FECHAEJECUCION)"
    
    If rstAnnoPA.State = 1 Then rstAnnoPA.Close
    
    rstAnnoPA.Open SqlCad, cnBdCPlus, adOpenForwardOnly, adLockOptimistic
    
    cmbAnno.Clear
    
    If Not rstAnnoPA.EOF Then
        Do While Not rstAnnoPA.EOF
            cmbAnno.AddItem Trim(rstAnnoPA!Anno & "")
            
            rstAnnoPA.MoveNext
        Loop
    End If
    
    If cmbAnno.ListCount > 0 Then
        cmbAnno.ListIndex = cmbAnno.ListCount - 1
    End If
End Sub

Private Sub listarAnnos()
    Dim rstAnnoPA As New ADODB.Recordset
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "YEAR(FECHAEJECUCION) AS ANNO "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "SF4COMPROMISOAUTOMATICO "
    SqlCad = SqlCad & "GROUP BY "
    SqlCad = SqlCad & "YEAR(FECHAEJECUCION)"
    
    If rstAnnoPA.State = 1 Then rstAnnoPA.Close
    
    rstAnnoPA.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockOptimistic
    
    cmbAnno.Clear
    
    If Not rstAnnoPA.EOF Then
        Do While Not rstAnnoPA.EOF
            cmbAnno.AddItem Trim(rstAnnoPA!Anno & "")
            
            rstAnnoPA.MoveNext
        Loop
    End If
    
    If cmbAnno.ListCount > 0 Then
        cmbAnno.ListIndex = cmbAnno.ListCount - 1
    End If
End Sub

Private Sub listarMeses()
    Dim rstMesPA As New ADODB.Recordset
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "MONTH(FECHAEJECUCION) AS MES "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "SF4COMPROMISOAUTOMATICO "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "YEAR(FECHAEJECUCION) = '" & cmbAnno.Text & "' "
    SqlCad = SqlCad & "GROUP BY "
    SqlCad = SqlCad & "MONTH(FECHAEJECUCION)"
    
    If rstMesPA.State = 1 Then rstMesPA.Close
    
    rstMesPA.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockOptimistic
    
    cmbMes.Clear
    
    If Not rstMesPA.EOF Then
        Do While Not rstMesPA.EOF
            cmbMes.AddItem UCase(Format("01/" & Format(Trim(rstMesPA!mes & ""), "00") & "/" & cmbAnno.Text, "MMMM")) & Space(100) & Format(Trim(rstMesPA!mes & ""), "00")
            
            rstMesPA.MoveNext
        Loop
    End If
    
    If cmbMes.ListCount > 0 Then
        cmbMes.ListIndex = cmbMes.ListCount - 1
    End If
End Sub

Private Sub listarFechasSQL()
    Dim rstFechaPA As New ADODB.Recordset
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "FECHAEJECUCION "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "SF4COMPROMISOAUTOMATICO "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "YEAR(FECHAEJECUCION) = " & cmbAnno.Text & " AND "
    SqlCad = SqlCad & "MONTH(FECHAEJECUCION) = " & Val(right(cmbMes.Text, 2)) & " "
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "FECHAEJECUCION"
    
    If rstFechaPA.State = 1 Then rstFechaPA.Close
    
    rstFechaPA.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockOptimistic
    
    cmbFechaEjecucion.Clear
    
    If Not rstFechaPA.EOF Then
        Do While Not rstFechaPA.EOF
            cmbFechaEjecucion.AddItem Trim(rstFechaPA!FECHAEJECUCION & "")
            
            rstFechaPA.MoveNext
        Loop
    End If
    
    If cmbFechaEjecucion.ListCount > 0 Then
        cmbFechaEjecucion.ListIndex = cmbFechaEjecucion.ListCount - 1
    End If
End Sub

Private Sub listarFechas()
    Dim rstFechaPA As New ADODB.Recordset
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "FECHAEJECUCION "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "SF4COMPROMISOAUTOMATICO "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "YEAR(FECHAEJECUCION) = " & cmbAnno.Text & " AND "
    SqlCad = SqlCad & "MONTH(FECHAEJECUCION) = " & Val(right(cmbMes.Text, 2)) & " "
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "FECHAEJECUCION"
    
    If rstFechaPA.State = 1 Then rstFechaPA.Close
    
    rstFechaPA.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockOptimistic
    
    cmbFechaEjecucion.Clear
    
    If Not rstFechaPA.EOF Then
        Do While Not rstFechaPA.EOF
            cmbFechaEjecucion.AddItem Trim(rstFechaPA!FECHAEJECUCION & "")
            
            rstFechaPA.MoveNext
        Loop
    End If
    
    If cmbFechaEjecucion.ListCount > 0 Then
        cmbFechaEjecucion.ListIndex = cmbFechaEjecucion.ListCount - 1
    End If
End Sub

Private Sub listarClientes()
    Dim rstClientePA As New ADODB.Recordset
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "CLIENTE "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "SF3COMPROMISOAUTOMATICO "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "CVDATE(FECHAEJECUCION) = CVDATE('" & cmbFechaEjecucion.Text & "') "
    SqlCad = SqlCad & "GROUP BY "
    SqlCad = SqlCad & "CLIENTE"
    
    If rstClientePA.State = 1 Then rstClientePA.Close
    
    rstClientePA.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockOptimistic
    
    cmbCliente.Clear
    
    cmbCliente.AddItem "(*) Todos"
    
    If Not rstClientePA.EOF Then
        Do While Not rstClientePA.EOF
            cmbCliente.AddItem Trim(rstClientePA!CLIENTE & "")
            
            rstClientePA.MoveNext
        Loop
    End If
    
    If cmbCliente.ListCount > 0 Then
        cmbCliente.ListIndex = 0
    End If
End Sub

Private Sub limpiarCajas()
    lblHoraInicio.Caption = "Hora de Inicio"
    lblHoraFin.Caption = "Hora de Fin"
    lblFechaCorte.Caption = "Fec. Corte"
    
    txtbusqueda.Text = vbNullString
    
    txtCodProducto.Text = vbNullString
    txtProducto.Text = vbNullString
End Sub

Private Sub procesarConsulta()
    On Error GoTo errProcesarConsulta
    
    Screen.MousePointer = vbHourglass
    
    lblHoraInicio.Caption = "Hora de Inicio: " & ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "HORAINICIO", "SF4COMPROMISOAUTOMATICO", "FECHAEJECUCION", cmbFechaEjecucion.Text, "F")
    lblHoraFin.Caption = "Hora de Fin: " & ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "HORAFIN", "SF4COMPROMISOAUTOMATICO", "FECHAEJECUCION", cmbFechaEjecucion.Text, "F")
    lblFechaCorte.Caption = "Fec. Corte: " & ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "FECHACORTEVALIDEZPEDIDO", "SF4COMPROMISOAUTOMATICO", "FECHAEJECUCION", cmbFechaEjecucion.Text, "F")
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "FORMAT(DET.FECHAEJECUCION, 'DDMMYYYY') & DET.NROPEDIDO & DET.CODPRODUCTO AS LLAVE, "
    SqlCad = SqlCad & "DET.FECHAEJECUCION, "
    SqlCad = SqlCad & "DET.NROPEDIDO, "
    SqlCad = SqlCad & "DET.CODPRODUCTO, "
    SqlCad = SqlCad & "('[ No. Requerimiento: ' & DET.NROPEDIDO & ' ]                    [ Cliente: ' & DET.CLIENTE & ' ]                    [ Fec. Entrega: ' & DET.FECHAENTREGA & ']                    [ Vendedor: ' & DET.VENDEDOR & ' ]') AS INFO, "
    SqlCad = SqlCad & "TE.DESCRIPCION AS TIPO, "
    SqlCad = SqlCad & "FAM.F7DESCON AS FAMILIA, "
    SqlCad = SqlCad & "SFAM.F7DESCON AS SUBFAMILIA, "
    SqlCad = SqlCad & "PROD.F5NOMPRO AS NOMPRODUCTO, "
    SqlCad = SqlCad & "MED.F7SIGMED AS UM, "
    SqlCad = SqlCad & "DET.SALDOPORCOMPROMETER AS SALDO "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "(((((SF3COMPROMISOAUTOMATICO AS DET "
    SqlCad = SqlCad & "LEFT JOIN SF4COMPROMISOAUTOMATICO AS CAB ON CAB.FECHAEJECUCION = DET.FECHAEJECUCION) "
    SqlCad = SqlCad & "LEFT JOIN IF5PLA AS PROD ON PROD.F5CODPRO = DET.CODPRODUCTO) "
    SqlCad = SqlCad & "LEFT JOIN EF7MEDIDAS AS MED ON MED.F7CODMED = PROD.F7CODMED) "
    SqlCad = SqlCad & "LEFT JOIN EF2TIPOEXISTENCIA AS TE ON TE.CODIGO = PROD.F5TIPO) "
    SqlCad = SqlCad & "LEFT JOIN SF7NIVEL02 AS SFAM ON SFAM.F7CODCON = PROD.F5UBICACIO) "
    SqlCad = SqlCad & "LEFT JOIN SF7NIVEL01 AS FAM ON FAM.F7CODCON = SFAM.F7NIVEL01 "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "CVDATE(DET.FECHAEJECUCION) = CVDATE('" & cmbFechaEjecucion.Text & "') "
    
        If left(cmbCliente.Text, 3) <> "(*)" Then
            SqlCad = SqlCad & "AND DET.CLIENTE = '" & cmbCliente.Text & "' "
        End If
        
        If txtbusqueda.Text <> vbNullString Then
            SqlCad = SqlCad & "AND ("
            'SqlCad = SqlCad & "DET.NROPEDIDO LIKE '%" & txtBusqueda.Text & "%' OR "
            SqlCad = SqlCad & "TE.DESCRIPCION LIKE '%" & txtbusqueda.Text & "%' OR "
            SqlCad = SqlCad & "FAM.F7DESCON LIKE '%" & txtbusqueda.Text & "%' OR "
            SqlCad = SqlCad & "SFAM.F7DESCON LIKE '%" & txtbusqueda.Text & "%' OR "
            SqlCad = SqlCad & "PROD.F5NOMPRO LIKE '%" & txtbusqueda.Text & "%'"
            SqlCad = SqlCad & ") "
        End If
        
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "DET.FECHAENTREGA, "
    SqlCad = SqlCad & "DET.NROPEDIDO, "
    SqlCad = SqlCad & "PROD.F5NOMPRO"
    
    If Not dbgReporte Is Nothing Then
        With dbgReporte
            .Dataset.Close

            .Columns.DestroyColumns
        End With

        Dim gColumn As dxGridColumn

        With dbgReporte
            'Columna Llave
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Llave"
                .DisableEditor = True
                .FieldName = "LLAVE"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColLlave"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 50
                .Visible = False
            End With
            
            'Columna Nro Pedido
            Set gColumn = .Columns.Add(gedTextEdit)

            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "No. Requerimiento"
                .DisableEditor = True
                .FieldName = "NROPEDIDO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColNroPedido"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 50
                .Visible = False
            End With
            
            'Columna Codigo de Producto
            Set gColumn = .Columns.Add(gedTextEdit)

            With gColumn
                .Alignment = taLeftJustify
                .BandIndex = 0
                .Caption = "Codigo"
                .DisableEditor = True
                .FieldName = "CODPRODUCTO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodProducto"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
                .Visible = False
            End With
            
            'Columna Informacion del Requerimiento
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Información de Requerimiento"
                .DisableEditor = True
                .FieldName = "INFO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColInfo"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 50
            End With
            
            'Columna Tipo de Existencia
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Tipo Existencia"
                .Color = vbWhite
                .DisableEditor = True
                .FieldName = "TIPO"
                .HeaderAlignment = taCenter
                .ObjectName = "ColTipoExistencia"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
                .Visible = CBool(tlbReporte.Tools.ITEM("Mostrar").State)
            End With
            
            'Columna Familia de Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Familia"
                .DisableEditor = True
                .FieldName = "FAMILIA"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColFamilia"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
                .Visible = CBool(tlbReporte.Tools.ITEM("Mostrar").State)
            End With
            
            'Columna Sub-Familia de Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Sub-Familia"
                .DisableEditor = True
                .FieldName = "SUBFAMILIA"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColSubFamilia"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
                .Visible = CBool(tlbReporte.Tools.ITEM("Mostrar").State)
            End With

            'Columna Descripcion de Producto
            Set gColumn = .Columns.Add(gedTextEdit)

            With gColumn
                .Alignment = taLeftJustify
                .BandIndex = 0
                .Caption = "Descripción del Producto"
                .Color = vbWhite
                .DisableEditor = True
                .FieldName = "NOMPRODUCTO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColNomProducto"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 250
            End With

            'Columna Unidad de Medida
            Set gColumn = .Columns.Add(gedTextEdit)

            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "U.M."
                .Color = vbWhite
                .DisableEditor = True
                .FieldName = "UM"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColUM"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 50
            End With
            
            'Columna Cantidad
            Set gColumn = .Columns.Add(gedSpinEdit)

            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 5
                .Caption = "Cant. por Reponer"
                .Color = &HFFFFC0
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "SALDO"
                .Font.Charset = 0
                .Font.Bold = True
                .HeaderAlignment = taCenter
                .ObjectName = "ColCantidad"
                .SummaryFooterType = cstSum
                .Width = 70
            End With
            
'
'            'Columna Atendido por Proveedor
'            Set gColumn = .Columns.Add(gedImageEdit)
'
'            With gColumn
'                .Alignment = taCenter
'                .BandIndex = 0
'                .Caption = "Atte."
'                .DisableEditor = True
'                .FieldName = "ATENDIDOPORPROV2"
'                .HeaderAlignment = taCenter
'                .ObjectName = "ColAtendidoPorProv2"
'
'                With .ImageColumn
'                    .Images = imgLstEstado.hImageList
'
'                    .ImageIndexes.Add ("0") 'Atendido por Proveedor
'                    .Values.Add ("1")
'                    .Descriptions.Add ("Atendido por Proveedor")
'
'                    .ShowDescription = False
'                End With
'
'                .SummaryFooterType = cstCount
'                .SummaryFooterFormat = " "
'                .Width = 20
'            End With
'
'            'Columna Compromiso en Almacen a favor del Producto
'            Set gColumn = .Columns.Add(gedSpinEdit)
'
'            With gColumn
'                .Alignment = taRightJustify
'                .BandIndex = 1
'                .Caption = "En Almacen"
'                .Color = &HC0&
'                .DecimalPlaces = 2
'                .DisableEditor = True
'                .FieldName = "COMPROMISOEA"
'                .Font.Bold = True
'                .FontColor = &HFFFFFF
'                .Font.Charset = 0
'                .HeaderAlignment = taCenter
'                .ObjectName = "ColCompromisoEA"
'                .SummaryFooterType = cstCount
'                .SummaryFooterFormat = " "
'                .Width = 70
'            End With
'
'            'Columna Compromiso por Llegar a favor del Producto
'            Set gColumn = .Columns.Add(gedSpinEdit)
'
'            With gColumn
'                .Alignment = taRightJustify
'                .BandIndex = 1
'                .Caption = "Por Llegar"
'                .Color = &H80FFFF
'                .DecimalPlaces = 2
'                .DisableEditor = True
'                .FieldName = "COMPROMISOPL"
'                .Font.Bold = False
'                .FontColor = &H80000012
'                .Font.Charset = 0
'                .HeaderAlignment = taCenter
'                .ObjectName = "ColCompromisoPL"
'                .SummaryFooterType = cstCount
'                .SummaryFooterFormat = " "
'                .Width = 70
'            End With
'
'            'Columna Saldo en Produccion
'            Set gColumn = .Columns.Add(gedSpinEdit)
'
'            With gColumn
'                .Alignment = taRightJustify
'                .BandIndex = 1
'                .Caption = "Saldo en OP's"
'                .Color = vbWhite
'                .DecimalPlaces = 2
'                .DisableEditor = True
'                .FieldName = "SALDO"
'                .Font.Charset = 0
'                .HeaderAlignment = taCenter
'                .ObjectName = "ColSaldo"
'                .SummaryFooterType = cstCount
'                .SummaryFooterFormat = " "
'                .Width = 70
'            End With
'
'            'Columna Compromiso en Almacen General
'            Set gColumn = .Columns.Add(gedSpinEdit)
'
'            With gColumn
'                .Alignment = taRightJustify
'                .BandIndex = 2
'                .Caption = "En Almacen"
'                .Color = &HC0&
'                .DecimalPlaces = 2
'                .DisableEditor = True
'                .FieldName = "COMPROMISOEAG"
'                .Font.Bold = True
'                .FontColor = &HFFFFFF
'                .Font.Charset = 0
'                .HeaderAlignment = taCenter
'                .ObjectName = "ColCompromisoEAG"
'                .SummaryFooterType = cstCount
'                .SummaryFooterFormat = " "
'                .Width = 70
'            End With
'
'            'Columna Compromiso Por Llegar General
'            Set gColumn = .Columns.Add(gedSpinEdit)
'
'            With gColumn
'                .Alignment = taRightJustify
'                .BandIndex = 2
'                .Caption = "Por Llegar"
'                .Color = &H80FFFF
'                .DecimalPlaces = 2
'                .DisableEditor = True
'                .FieldName = "COMPROMISOPLG"
'                .Font.Bold = False
'                .FontColor = &H80000012
'                .Font.Charset = 0
'                .HeaderAlignment = taCenter
'                .ObjectName = "ColCompromisoPLG"
'                .SummaryFooterType = cstCount
'                .SummaryFooterFormat = " "
'                .Width = 70
'            End With
'
'
'            'Columna Libre en Almacen General
'            Set gColumn = .Columns.Add(gedSpinEdit)
'
'            With gColumn
'                .Alignment = taRightJustify
'                .BandIndex = 3
'                .Caption = "En Almacen"
'                .Color = &HC000&
'                .DecimalPlaces = 2
'                .DisableEditor = True
'                .FieldName = "LIBREEAG"
'                .Font.Bold = True
'                .FontColor = &HFFFFFF
'                .Font.Charset = 0
'                .HeaderAlignment = taCenter
'                .ObjectName = "ColLibreEAG"
'                .SummaryFooterType = cstCount
'                .SummaryFooterFormat = " "
'                .Width = 70
'            End With
'
'            'Columna Libre Por Llegar General
'            Set gColumn = .Columns.Add(gedSpinEdit)
'
'            With gColumn
'                .Alignment = taRightJustify
'                .BandIndex = 3
'                .Caption = "Por Llegar"
'                .Color = &H80FFFF
'                .DecimalPlaces = 2
'                .DisableEditor = True
'                .FieldName = "LIBREPLG"
'                .Font.Bold = False
'                .FontColor = &H80000012
'                .Font.Charset = 0
'                .HeaderAlignment = taCenter
'                .ObjectName = "ColLibrePLG"
'                .SummaryFooterType = cstCount
'                .SummaryFooterFormat = " "
'                .Width = 70
'            End With
'
'            'Columna Stock en Almacen General
'            Set gColumn = .Columns.Add(gedSpinEdit)
'
'            With gColumn
'                .Alignment = taRightJustify
'                .BandIndex = 4
'                .Caption = "En Almacen"
'                .Color = &HC00000
'                .DecimalPlaces = 2
'                .DisableEditor = True
'                .FieldName = "STOCKEAG"
'                .Font.Bold = True
'                .FontColor = &HFFFFFF
'                .Font.Charset = 0
'                .HeaderAlignment = taCenter
'                .ObjectName = "ColStockEAG"
'                .SummaryFooterType = cstCount
'                .SummaryFooterFormat = " "
'                .Width = 70
'            End With
'
'            'Columna Stock Por Llegar General
'            Set gColumn = .Columns.Add(gedSpinEdit)
'
'            With gColumn
'                .Alignment = taRightJustify
'                .BandIndex = 4
'                .Caption = "Por Llegar"
'                .Color = &H80FFFF
'                .DecimalPlaces = 2
'                .DisableEditor = True
'                .FieldName = "STOCKPLG"
'                .Font.Charset = 0
'                .HeaderAlignment = taCenter
'                .ObjectName = "ColStockPLG"
'                .SummaryFooterType = cstCount
'                .SummaryFooterFormat = " "
'                .Width = 70
'            End With
'
'            'Columna Cantidad Atendida
'            Set gColumn = .Columns.Add(gedSpinEdit)
'
'            With gColumn
'                .Alignment = taRightJustify
'                .BandIndex = 5
'                .Caption = "Recomendacion"
'                .Color = vbWhite
'                .DecimalPlaces = 2
'                .DisableEditor = True
'                .FieldName = "COMPRAR"
'                .Font.Charset = 0
'                .HeaderAlignment = taCenter
'                .ObjectName = "ColComprar"
'                .SummaryFooterType = cstCount
'                .SummaryFooterFormat = " "
'                .Width = 70
'            End With
'
'            'Columna Cantidad Por Comprar
'            Set gColumn = .Columns.Add(gedSpinEdit)
'
'            With gColumn
'                .Alignment = taRightJustify
'                .BandIndex = 5
'                .Caption = "Cant. a Comprar"
'                .Color = &HFFFFC0
'                .DecimalPlaces = 2
'                .DisableEditor = True
'                .FieldName = "COMPRA"
'                .Font.Charset = 0
'                .HeaderAlignment = taCenter
'                .ObjectName = "ColCompra"
'                .SummaryFooterType = cstSum
'                .Width = 70
'            End With
            
            abrirCnnDbBancos

            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = cnn_dbbancos
            
            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctStatic
            .Dataset.ADODataset.LockType = ltOptimistic
            .Dataset.ADODataset.CommandText = SqlCad
            .Dataset.Active = True
            .Dataset.Refresh

            .KeyField = "LLAVE"
            
            .Columns.ColumnByFieldName("INFO").GroupIndex = 0
            
            .Columns.HeaderFont.Bold = True
            
            .m.FullExpand
            
            .Columns.ColumnByFieldName("NOMPRODUCTO").SummaryFooterType = cstCount
            .Columns.ColumnByFieldName("NOMPRODUCTO").SummaryFooterFormat = "Cantidad de Registros = " & .Dataset.RecordCount
            
            .Filter.FilterActive = True
        End With
    End If
    
    SqlCad = vbNullString
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
errProcesarConsulta:
    Select Case Err.Number
        Case 3704, 3709
            abrirCnnDbBancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ProcesarConsulta"
    End Select
    
    Err.Clear
End Sub

Private Sub cmbAnno_Click()
    listarMeses
End Sub

Private Sub cmbFechaEjecucion_Click()
    listarClientes
End Sub

Private Sub cmbMes_Click()
    listarFechas
End Sub

Private Sub dbgReporte_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    Select Case UCase(Column.FieldName)
        Case "CANTIDAD", "TOTALMN", "TOTALME"
            Text = Format(Text, "#,0.00")
    End Select
End Sub

Private Sub dbgReporte_OnCustomDrawFooter(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    Select Case UCase(Column.FieldName)
        Case "CANTIDAD", "TOTALMN", "TOTALME"
            Text = Format(Text, "#,0.00")
            FontColor = vbBlue
            Color = vbWhite
            Font.Bold = True
    End Select
End Sub

'Private Sub dbgReporte_OnCustomDrawFooterNode(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal FooterIndex As Integer, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
'    Select Case UCase(Column.FieldName)
'        Case "CANTIDAD", "PORCENTAJEDSCTO", "COSTOMN", "COSTOME", "SUBTOTALMN", "SUBTOTALME", "IMPUESTOMN", "IMPUESTOME", "TOTALMN", "TOTALME"
'            Text = Format(Text, "#,0.00")
'            FontColor = vbBlue
'            Color = vbWhite
'            Font.Bold = True
'    End Select
'End Sub

Private Sub dtpDesde_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub dtpHasta_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub Form_Load()
    limpiarCajas
    
    
    
    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
        listarAnnos
    Else
        listarAnnos
    End If
    
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    'fraPeriodo.Move 0, 0, Me.ScaleWidth, 1000
    
    dbgReporte.Move 0, fraPeriodo.Height + 200, Me.ScaleWidth, (Me.ScaleHeight - fraPeriodo.Height) - 200
End Sub

Private Sub optAgruparPor_Click(Index As Integer)
    intOpcionResumen = Index
End Sub

Private Sub tlbReporte_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.ID
        Case "Consultar"
            procesarConsulta
        Case "Mostrar"
            If Not CBool(Tool.State) Then
                tlbReporte.Tools("Mostrar").ChangeAll ssChangeAllName, "Mostr&ar"
            Else
                tlbReporte.Tools("Mostrar").ChangeAll ssChangeAllName, "Ocult&ar"
            End If
            
            procesarConsulta
        Case "Excel"
            Screen.MousePointer = vbHourglass
            
            With cmdlgReporte
                .DialogTitle = "Guardar como..."
                .Filter = "Archivos de MS Excel | *.xls"
                .FileName = vbNullString
                
                .ShowSave
                
                If .FileName <> vbNullString Then
                    dbgReporte.m.ExportToXLS .FileName
                    
                    If Dir(.FileName) <> vbNullString Then
                        MsgBox "Exportación terminada.", vbInformation, App.ProductName
                    Else
                        MsgBox "Exportación fallida.", vbInformation, App.ProductName
                    End If
                End If
            End With
            
            Screen.MousePointer = vbDefault
        Case "Salir"
            Unload Me
    End Select
End Sub

Private Sub txtbusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            procesarConsulta
    End Select
End Sub

Private Sub txtCodProducto_DblClick()
    txtCodProducto_KeyDown vbKeyF2, 0
End Sub

Private Sub txtCodProducto_GotFocus()
    ModUtilitario.seleccionarTextoCaja txtCodProducto
End Sub

Private Sub txtCodProducto_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
            If ModUtilitario.validarFormAbierto("frmListaBien") Then
                Unload frmListaBien
            End If
            
            With frmListaBien
                objAyudaBien.inicializarEntidades
                
                '.Ayuda = True
                '.TieneMovimientoAlmacen = True
                '.InsumoOP = True 'False
                
                .Ayuda = True
                .InsumoOP = True
                .ParaVenta = False
                .TieneMovimientoAlmacen = True
                .CadenaCorte = vbNullString
                .FiltroAdicional = vbNullString
                .TipoBienMostrar = "P"
                
                objAyudaBien.inicializarEntidades
                
                .Show 1
                
                If objAyudaBien.Codigo <> vbNullString Then
                    objAyudaBien.obtenerConfigBien
                    
                    txtCodProducto.Text = objAyudaBien.Codigo
                    txtProducto.Text = objAyudaBien.Descripcion
                    txtProducto.ToolTipText = txtProducto.Text
                Else
                    txtCodProducto.Text = vbNullString
                    txtProducto.ToolTipText = "Ingrese cadena a buscar"
                End If
            End With
        Case vbKeyReturn
            'ModUtilitario.pulsarTecla vbKeyTab
            procesarConsulta
    End Select
End Sub

Private Sub txtCodProducto_LostFocus()
    If Trim(txtCodProducto.Text) <> vbNullString Then
        txtCodProducto.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5CODPRO", "IF5PLA", "F5CODPRO", Trim(txtCodProducto.Text), "T")
        txtProducto.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5NOMPRO", "IF5PLA", "F5CODPRO", Trim(txtCodProducto.Text), "T")
        txtProducto.ToolTipText = txtProducto.Text
        
        If Trim(txtCodProducto.Text) = vbNullString Then
            txtProducto.Text = vbNullString
            txtProducto.ToolTipText = "Ingrese cadena a buscar"
        End If
    Else
        txtProducto.Text = vbNullString
        txtProducto.ToolTipText = "Ingrese cadena a buscar"
    End If
End Sub

Private Sub txtProducto_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            procesarConsulta
    End Select
End Sub
