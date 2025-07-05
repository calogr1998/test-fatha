VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUtilRegularizacionStockCEA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Regularización de Stock Comprometido en EA y PL"
   ClientHeight    =   5520
   ClientLeft      =   1425
   ClientTop       =   1845
   ClientWidth     =   8415
   Icon            =   "frmUtilRegularizacionStockCEA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   8415
   Begin VB.Frame fraDatos 
      Caption         =   " Datos "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   8175
      Begin VB.CheckBox chkProceso 
         Caption         =   "Compromiso Automatico de Pedidos con Fec. Entrega mayor a la Fecha de Hoy."
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   18
         Top             =   2400
         Width           =   7695
      End
      Begin VB.CheckBox chkProceso 
         Caption         =   "Reposición de Compromisos Afectados de Pedidos con Fec. Entrega dentro del Periodo de Validez."
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   17
         Top             =   2160
         Width           =   7695
      End
      Begin VB.CheckBox chkProceso 
         Caption         =   "Descargar Stock Actual de CEA y CPL."
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   16
         Top             =   1920
         Value           =   1  'Checked
         Width           =   7695
      End
      Begin VB.CheckBox chkProceso 
         Caption         =   "Limpieza de Stock CEA y CPL."
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   1680
         Width           =   7695
      End
      Begin VB.CheckBox chkProceso 
         Caption         =   "Resumen de Saldos por Atender en OPs."
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   1440
         Value           =   1  'Checked
         Width           =   7695
      End
      Begin VB.TextBox txtNroPedido 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtCodProducto 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblNroPedido 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label4"
         Height          =   615
         Left            =   2640
         TabIndex        =   13
         Top             =   720
         Width           =   5415
      End
      Begin VB.Label Label1 
         Caption         =   "No. Pedido"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Producto"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblProducto 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label4"
         Height          =   255
         Left            =   2040
         TabIndex        =   10
         Top             =   360
         Width           =   6015
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Proceso en curso:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   8175
      Begin MSComctlLib.ProgressBar pgbProceso2 
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar pgbProceso1 
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblProceso2 
         Caption         =   "Proceso 2"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   7695
      End
      Begin VB.Label lblProceso1 
         Caption         =   "Proceso 1"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   7695
      End
   End
   Begin VB.CommandButton cmdOperacion 
      Caption         =   "Salir"
      Height          =   375
      Index           =   1
      Left            =   4680
      TabIndex        =   3
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdOperacion 
      Caption         =   "Aceptar"
      Height          =   375
      Index           =   0
      Left            =   2640
      TabIndex        =   2
      Top             =   4920
      Width           =   1215
   End
End
Attribute VB_Name = "frmUtilRegularizacionStockCEA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim objValeRegularizacionCEA As ClsVale
Dim objOrdenRegularizacionCPL As ClsOrden

Private Sub inicializarControles()
    txtCodProducto.Text = vbNullString
        lblProducto.Caption = "Todos los Productos (*)"
        
    txtNroPedido.Text = vbNullString
        lblNroPedido.Caption = vbNullString
    
    lblProceso1.Caption = vbNullString
    pgbProceso1.Value = 0
    lblProceso2.Caption = vbNullString
    pgbProceso2.Value = 0
    
    chkProceso(0).BackColor = &H8000000F
    chkProceso(1).BackColor = &H8000000F: chkProceso(1).Value = vbUnchecked
    chkProceso(2).BackColor = &H8000000F
    chkProceso(3).BackColor = &H8000000F: chkProceso(3).Value = vbUnchecked
    chkProceso(4).BackColor = &H8000000F: chkProceso(4).Value = vbUnchecked
    
    cmdOperacion(0).Enabled = True
    cmdOperacion(1).Enabled = True
End Sub

Private Sub cmdOperacion_Click(Index As Integer)
    Select Case Index
        Case 0
            If lblProducto.Caption = "Todos los Productos (*)" Then
                txtCodProducto.Text = vbNullString
            End If
            
            If lblNroPedido.Caption = vbNullString Then
                txtNroPedido.Text = vbNullString
            End If
            
            cmdOperacion(0).Enabled = False
            cmdOperacion(1).Enabled = False
            
            analisisYCorreccionStockComprometidoActual
            
            inicializarControles
        Case 1
            Unload Me
    End Select
End Sub

Private Sub cmdOperacion_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub Form_Load()
    Me.left = (Screen.Width / 2) - (Me.Width / 2)
    Me.top = (Screen.Height / 2) - (Me.Height / 2)
    
    ModUtilitario.deshabilitarBotonCerrarForm frmUtilRegularizacionStockCEA
    
    inicializarControles
End Sub

Private Sub analisisYCorreccionStockComprometidoActual()
    On Error GoTo errAnalisisYCorreccionStockComprometidoActual
    
    Dim rstStockCEA As New ADODB.Recordset
    Dim rstStockCPL As New ADODB.Recordset
    Dim rstProduccion As New ADODB.Recordset
    
    Dim rstResProd As New ADODB.Recordset
    Dim rstStockLEA As New ADODB.Recordset
    Dim rstStockLPL As New ADODB.Recordset
    
    Dim dblStock As Double
    Dim dblCantidadProdObservado As Double
    
    Dim dblLiberarCompromiso As Double
    
    Dim rstDistribuir As New ADODB.Recordset
    Dim rstOCDetalle As New ADODB.Recordset
    Dim dblItem As Double
    Dim dblCantidadDestino As Double
    Dim dblCantidadDestinoSegunOC As Double
    Dim dblCantidadDisponibleSegunOC As Double
    
    Dim dblFactorUM As Double
    
    Dim dblItemLibreEnOC As Double
    
    Dim bolCompromisoEjecutado As Boolean
    Dim dblCantidadComprometer As Double
    
    Dim strFechaCorteValidezCompromiso As String
    Dim strFechaCorteActual As String
    Dim intCantidadMesesDeValidezCompromiso As Integer
    
    Dim strNumValeProceso1 As String
    Dim strNumValeProceso2 As String
    Dim strNumValeProceso3 As String
    Dim strNumValeProceso3_add As String
    
    
    Rem SK: PM = Proceso Manual
    
    If MsgBox("¿Desea ejecutar el siguiente proceso?" & vbNewLine & _
                "RECUERDE: Este proceso puede tardar varios minutos.", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
        
        Exit Sub
    End If
    
    'Descargar en el Temporal los compromisos actuales, segun Ordenes de Produccion
    
    Screen.MousePointer = vbHourglass
    
    Actualiza_Log "INICIO DE PROCESO MANUAL: " & Now, StrConexDbBancos
    
    intCantidadMesesDeValidezCompromiso = Val(ModUtilitario.sGetINI(App.Path & strNombreFicheroConfigSQLCliente, "ConfigServidorSQLCliente", "CantidadMesesDeValidezCompromiso", "l"))
    strFechaCorteValidezCompromiso = Format(Date - (30 * intCantidadMesesDeValidezCompromiso), "Short Date")
    strFechaCorteActual = Format(Date, "Short Date")
    
    '-------- PARA REPORTE DE COMPROMISOS NO COMPLETADOS POR FALTA DE STOCK -------
    
    If ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "FECHAEJECUCION", "SF4COMPROMISOAUTOMATICO", "FECHAEJECUCION", strFechaCorteActual, "F") = vbNullString Then
        SqlCad = vbNullString
        SqlCad = SqlCad & "INSERT INTO SF4COMPROMISOAUTOMATICO("
        SqlCad = SqlCad & "FECHAEJECUCION, "
        SqlCad = SqlCad & "HORAINICIO, "
        SqlCad = SqlCad & "FECHACORTEVALIDEZPEDIDO, "
        SqlCad = SqlCad & "USUREG, "
        SqlCad = SqlCad & "FECREG) "
        SqlCad = SqlCad & "VALUES("
        SqlCad = SqlCad & "CVDATE('" & strFechaCorteActual & "'), "
        SqlCad = SqlCad & "CVDATE('" & Format(Now, "hh:mm:ss AM/PM") & "'), "
        SqlCad = SqlCad & "CVDATE('" & strFechaCorteValidezCompromiso & "'), "
        SqlCad = SqlCad & "'" & wusuario & "', "
        SqlCad = SqlCad & "CVDATE('" & Format(Date, "Short Date") & "')"
        SqlCad = SqlCad & ")"
    Else
        SqlCad = vbNullString
        SqlCad = SqlCad & "UPDATE "
        SqlCad = SqlCad & "SF4COMPROMISOAUTOMATICO "
        SqlCad = SqlCad & "SET "
        SqlCad = SqlCad & "USUMOD = '" & wusuario & "', "
        SqlCad = SqlCad & "FECMOD = CVDATE('" & Format(Date, "Short Date") & "') "
        SqlCad = SqlCad & "WHERE "
        SqlCad = SqlCad & "CVDATE(FECHAEJECUCION) = CVDATE('" & strFechaCorteActual & "')"
    End If
    
    abrirCnnDbBancos
    
    cnn_dbbancos.Execute SqlCad
    
    Actualiza_Log SqlCad, StrConexDbBancos
    
    '------------------------------------------------------------------------------
    
    If CBool(chkProceso(0).Value) Then
        If Not ModMilano.importarResumenRequerimientoProduccionV2(lblProceso1, _
                                                                pgbProceso1, _
                                                                "tmpCPResumenProduccionPM" & wusuario, _
                                                                Trim(txtNroPedido.Text), _
                                                                Trim(txtCodProducto.Text), _
                                                                vbNullString, _
                                                                strFechaCorteValidezCompromiso) Then
            
            chkProceso(0).BackColor = RGB(255, 51, 51)
            
            Screen.MousePointer = vbDefault
            
            Actualiza_Log "Proceso de Regularizacion de Stock Trunco al no poder Importar Resumen de Requerimiento de Produccion.", StrConexDbBancos
            
            Exit Sub
        Else
            Actualiza_Log "PROCESO 1 FINALIZO AL: " & Now, StrConexDbBancos
            
            chkProceso(0).BackColor = RGB(255, 153, 51)
        End If
    End If
    
    Screen.MousePointer = vbDefault
    
    lblProceso1.Visible = True
    
    If CBool(chkProceso(1).Value) Then
        'Seleccionar Stock Comprometido en Almacen Actual para su verificacion (Limpieza de Stock Comprometido En Almacen)
        SqlCad = vbNullString
        SqlCad = SqlCad & "SELECT "
        SqlCad = SqlCad & "F2CODALM, "
        SqlCad = SqlCad & "F5CODPRO, "
        SqlCad = SqlCad & "COD_SOLICITUD, "
        SqlCad = SqlCad & "VAL(FORMAT( SUM(VAL(FORMAT(VAL(F3CANPRO & ''), '#0.00')) * IIF(TIPO = 'S', -1, 1)) , '#0.00')) AS CANTIDAD "
        SqlCad = SqlCad & "FROM "
        SqlCad = SqlCad & "IF3VALES "
        SqlCad = SqlCad & "WHERE "
        SqlCad = SqlCad & "F2CODALM = '01' "
        
            If Trim(txtNroPedido.Text) <> vbNullString Then
                SqlCad = SqlCad & "AND TRIM(COD_SOLICITUD & '') = '" & Trim(txtNroPedido.Text) & "' "
            Else
                SqlCad = SqlCad & "AND TRIM(COD_SOLICITUD & '') <> '' "
            End If
            
            If Trim(txtCodProducto.Text) <> vbNullString Then
                SqlCad = SqlCad & "AND F5CODPRO = '" & Trim(txtCodProducto.Text) & "' "
            End If
            
        SqlCad = SqlCad & "GROUP BY "
        SqlCad = SqlCad & "F5CODPRO, "
        SqlCad = SqlCad & "F2CODALM, "
        SqlCad = SqlCad & "COD_SOLICITUD "
        SqlCad = SqlCad & "HAVING "
        SqlCad = SqlCad & "VAL(FORMAT( SUM(VAL(FORMAT(VAL(F3CANPRO & ''), '#0.00')) * IIF(TIPO = 'S', -1, 1)) , '#0.00')) > 0 "
        SqlCad = SqlCad & "ORDER BY "
        SqlCad = SqlCad & "F2CODALM, "
        SqlCad = SqlCad & "F5CODPRO, "
        SqlCad = SqlCad & "COD_SOLICITUD"
    
        If rstStockCEA.State = 1 Then rstStockCEA.Close
    
        rstStockCEA.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
        If Not rstStockCEA.EOF Then
            pgbProceso2.Max = ModUtilitario.devuelveCantRegistros(rstStockCEA)
            pgbProceso2.Value = 0
            lblProceso2.Caption = "Evaluando Stock Comprometido En Almacen (CEA) Actual..."
    
            dblCantidadProdObservado = 0
            
            strNumValeProceso1 = vbNullString
            
            Do While Not rstStockCEA.EOF
    
                SqlCad = vbNullString
                SqlCad = SqlCad & "SELECT "
                SqlCad = SqlCad & "RES.NROPEDIDO, "
                SqlCad = SqlCad & "RES.CODPRODUCTO, "
                SqlCad = SqlCad & "SUM(RES.SALDO) AS SALDO "
                SqlCad = SqlCad & "FROM "
                SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTOPRODUCCION AS RES "
                SqlCad = SqlCad & "WHERE "
                'SqlCad = SqlCad & "CVDATE(RES.FENTREGA) BETWEEN CVDATE('" & strFechaCorteValidezCompromiso & "') AND CVDATE('" & Format(Date, "Short Date") & "') AND "
                SqlCad = SqlCad & "RES.NROPEDIDO = '" & Trim(rstStockCEA!COD_SOLICITUD & "") & "' AND "
                SqlCad = SqlCad & "RES.CODPRODUCTO = '" & Trim(rstStockCEA!f5codpro & "") & "' "
                SqlCad = SqlCad & "GROUP BY "
                SqlCad = SqlCad & "RES.NROPEDIDO, "
                SqlCad = SqlCad & "RES.CODPRODUCTO "
                SqlCad = SqlCad & "ORDER BY "
                SqlCad = SqlCad & "RES.NROPEDIDO, "
                SqlCad = SqlCad & "RES.CODPRODUCTO"
    
                If rstProduccion.State = 1 Then rstProduccion.Close
    
                rstProduccion.Open SqlCad, cnDBTemp, adOpenForwardOnly, adLockReadOnly
    
                If Not rstProduccion.EOF Then
                    Select Case Val(rstProduccion!SALDO & "")
                        Case Is >= Val(rstStockCEA!Cantidad & "")
                            dblLiberarCompromiso = 0
                        Case Else
                            dblLiberarCompromiso = Val(rstStockCEA!Cantidad & "") - Val(rstProduccion!SALDO & "")
                    End Select
                Else
                    dblLiberarCompromiso = Val(rstStockCEA!Cantidad & "")
                End If
    
                If dblLiberarCompromiso > 0 Then
                    Set objValeRegularizacionCEA = New ClsVale
    
                    With objValeRegularizacionCEA
    
                        .inicializarEntidades
    
                        .CodigoAlmacen = Trim(rstStockCEA!f2codalm & "")
                        .NumeroVale = strNumValeProceso1
                        .TipoVale = "I"
    
                        .Fecha = Format(Date, "dd/mm/yyyy")
                        .CodigoOrigen = "XCS"
                        .TipoCambio = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CAMBIO", "CAMBIOS", "FECHA", .Fecha, "F"))
    
                            If .TipoCambio = 0 Then
                                .TipoCambio = "4.05"
                            End If
    
                        .CodigoMoneda = "S"
    
                        .referencia = wnomcia
                        .observaciones = "PROCESO DE CORRECCION DE STOCK COMPROMETIDO."
    
                        .FecReg = Format(Date, "Short Date")
                        .UsuReg = wusuario
                        .FecMod = Format(Date, "Short Date")
                        .UsuMod = wusuario
    
                        If .guardarVale Then
                            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                            
                            If strNumValeProceso1 = vbNullString Then
                                strNumValeProceso1 = .NumeroVale
                                    
                                'Borrar Detalle de Vale
                                SqlCad = vbNullString
                                SqlCad = "DELETE FROM IF3VALES WHERE F2CODALM = '" & .CodigoAlmacen & "' AND F4NUMVAL = '" & .NumeroVale & "'"
        
                                cnn_dbbancos.Execute SqlCad
                                Actualiza_Log SqlCad, cnn_dbbancos.ConnectionString
                            End If
    
    
                            .inicializarEntidadesDetalle
    
                            .NumeroOrdenCompra = vbNullString
                            .Requerimiento = Trim(rstStockCEA!COD_SOLICITUD & "")
    
                            .CodigoProducto = Trim(rstStockCEA!f5codpro & "")
                            .CodigoProductoOriginal = Trim(rstStockCEA!f5codpro & "")
                            .Cantidad = dblLiberarCompromiso * -1
    
    
                            dblItem = dblItem + 1
    
                            .ITEM = dblItem
    
                            .guardarValeDetalleOneByOne
    
                            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
    
                            .inicializarEntidadesDetalle
    
                            dblItem = dblItem + 1
    
                            .Requerimiento = vbNullString
    
                            .CodigoProducto = Trim(rstStockCEA!f5codpro & "")
                            .CodigoProductoOriginal = Trim(rstStockCEA!f5codpro & "")
                            .Cantidad = dblLiberarCompromiso
                            .ITEM = dblItem
    
                            .guardarValeDetalleOneByOne
    
                            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                        End If
                    End With
    
                    Set objValeRegularizacionCEA = Nothing
                End If
    
                DoEvents
    
                pgbProceso2.Value = pgbProceso2.Value + 1
                lblProceso2.Caption = "Evaluando Stock Comprometido En Almacen (CEA) Actual..." & FormatPercent(pgbProceso2.Value / pgbProceso2.Max, 3)
    
                rstStockCEA.MoveNext
            Loop
                Actualiza_Log "PROCESO 2.1 FINALIZO AL: " & Now, StrConexDbBancos
        End If
        
        
        'Seleccionar Stock Comprometido Por Llegar Actual para su verificacion (Limpieza de Stock Comprometido Por llegar)
        SqlCad = vbNullString
        SqlCad = SqlCad & "SELECT "
        SqlCad = SqlCad & "DET.F3CODPRO, "
        SqlCad = SqlCad & "DET.COD_SOLICITUD, "
        SqlCad = SqlCad & "SUM( VAL(FORMAT( (((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) - VAL(INGRESOS.CANTIDAD & '')) , '#.0000')) ) AS SALDO "
        SqlCad = SqlCad & "FROM "
        SqlCad = SqlCad & "(((((IF3ORDEN AS DET "
        SqlCad = SqlCad & "LEFT JOIN IF5PLA AS PROD ON PROD.F5CODPRO = DET.F3CODPRO) "
        SqlCad = SqlCad & "LEFT JOIN EF7MEDIDAS AS MED ON MED.F7CODMED = PROD.F7CODMED) "
        SqlCad = SqlCad & "LEFT JOIN EF7MEDIDAS AS MED2 ON MED2.F7CODMED = DET.UNIDAD) "
        SqlCad = SqlCad & "LEFT JOIN MEDIVENTAS AS MEDALTER ON MEDALTER.F5CODPRO = DET.F3CODPRO AND MEDALTER.F7CODMED = DET.UNIDAD) "
        SqlCad = SqlCad & "LEFT JOIN "
        SqlCad = SqlCad & "(SELECT "
        SqlCad = SqlCad & "DET.F4NUMORD, "
        SqlCad = SqlCad & "DET.COD_SOLICITUD, "
        SqlCad = SqlCad & "DET.F5CODPROORIGINAL, "
        SqlCad = SqlCad & "VAL(FORMAT( SUM(DET.F3CANPRO * IIF(DET.TIPO = 'S', -1, 1)) , '#.0000')) AS CANTIDAD "
        SqlCad = SqlCad & "FROM "
        SqlCad = SqlCad & "IF3VALES AS DET "
        SqlCad = SqlCad & "LEFT JOIN IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM "
        SqlCad = SqlCad & "WHERE "
        SqlCad = SqlCad & "CAB.F1CODORI IN ('XC0') AND "
        SqlCad = SqlCad & "DET.F4NUMORD <> '' AND "
        SqlCad = SqlCad & "DET.COD_SOLICITUD <> '' "
        SqlCad = SqlCad & "GROUP BY "
        SqlCad = SqlCad & "DET.F4NUMORD, "
        SqlCad = SqlCad & "DET.COD_SOLICITUD, "
        SqlCad = SqlCad & "DET.F5CODPROORIGINAL) AS INGRESOS "
        SqlCad = SqlCad & "ON INGRESOS.F4NUMORD = DET.F4NUMORD AND INGRESOS.COD_SOLICITUD = DET.COD_SOLICITUD AND INGRESOS.F5CODPROORIGINAL = DET.F3CODPRO) "
        SqlCad = SqlCad & "LEFT JOIN TB_CABSOLICITUD AS CABPED ON CABPED.COD_SOLICITUD = DET.COD_SOLICITUD "
        SqlCad = SqlCad & "WHERE "
        SqlCad = SqlCad & "DET.F4LOCAL = 'OC' AND "
        SqlCad = SqlCad & "VAL(FORMAT( (((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) - VAL(INGRESOS.CANTIDAD & '')) , '#.0000')) > 0 "
            
            If Trim(txtNroPedido.Text) <> vbNullString Then
                SqlCad = SqlCad & "AND DET.COD_SOLICITUD = '" & Trim(txtNroPedido.Text) & "' "
            Else
                SqlCad = SqlCad & "AND DET.COD_SOLICITUD <> '' "
            End If
            
            If Trim(txtCodProducto.Text) <> vbNullString Then
                SqlCad = SqlCad & "AND DET.F3CODPRO = '" & Trim(txtCodProducto.Text) & "' "
            End If
        
        SqlCad = SqlCad & "GROUP BY "
        SqlCad = SqlCad & "DET.F3CODPRO, "
        SqlCad = SqlCad & "PROD.F5NOMPRO, "
        SqlCad = SqlCad & "DET.COD_SOLICITUD "
        SqlCad = SqlCad & "ORDER BY "
        SqlCad = SqlCad & "PROD.F5NOMPRO, "
        SqlCad = SqlCad & "DET.COD_SOLICITUD"
    
        If rstStockCPL.State = 1 Then rstStockCPL.Close
    
        rstStockCPL.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
        If Not rstStockCPL.EOF Then
            pgbProceso1.Max = ModUtilitario.devuelveCantRegistros(rstStockCPL)
            pgbProceso1.Value = 0
            lblProceso1.Caption = "Evaluando Stock Comprometido Por Llegar (CPL) Actual..."
    
            dblCantidadProdObservado = 0
    
            Do While Not rstStockCPL.EOF
                SqlCad = vbNullString
                SqlCad = SqlCad & "SELECT "
                SqlCad = SqlCad & "RES.NROPEDIDO, "
                SqlCad = SqlCad & "RES.CODPRODUCTO, "
                SqlCad = SqlCad & "SUM(RES.SALDO) AS SALDO "
                SqlCad = SqlCad & "FROM "
                SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTOPRODUCCION AS RES "
                SqlCad = SqlCad & "WHERE "
                'SqlCad = SqlCad & "CVDATE(RES.FENTREGA) BETWEEN CVDATE('" & strFechaCorteValidezCompromiso & "') AND CVDATE('" & Format(Date, "Short Date") & "') AND "
                SqlCad = SqlCad & "RES.NROPEDIDO = '" & Trim(rstStockCPL!COD_SOLICITUD & "") & "' AND "
                SqlCad = SqlCad & "RES.CODPRODUCTO = '" & Trim(rstStockCPL!F3CODPRO & "") & "' "
                SqlCad = SqlCad & "GROUP BY "
                SqlCad = SqlCad & "RES.NROPEDIDO, "
                SqlCad = SqlCad & "RES.CODPRODUCTO "
                SqlCad = SqlCad & "ORDER BY "
                SqlCad = SqlCad & "RES.NROPEDIDO, "
                SqlCad = SqlCad & "RES.CODPRODUCTO"
    
                abrirCnTemporal
    
                If rstProduccion.State = 1 Then rstProduccion.Close
    
                rstProduccion.Open SqlCad, cnDBTemp, adOpenForwardOnly, adLockReadOnly
    
                If Not rstProduccion.EOF Then
                    Select Case Val(rstProduccion!SALDO & "")
                        Case Is >= Val(rstStockCPL!SALDO & "")
                            dblLiberarCompromiso = 0
                        Case Else
                            dblLiberarCompromiso = Val(rstStockCPL!SALDO & "") - Val(rstProduccion!SALDO & "")
                    End Select
                Else
                    dblLiberarCompromiso = Val(rstStockCPL!SALDO & "")
                End If
    
                If dblLiberarCompromiso > 0 Then
                    SqlCad = vbNullString
                    SqlCad = SqlCad & "SELECT "
                    SqlCad = SqlCad & "DET.ITEM, "
                    SqlCad = SqlCad & "DET.F3CODPRO, "
                    SqlCad = SqlCad & "DET.COD_SOLICITUD, "
                    SqlCad = SqlCad & "DET.F4NUMORD, "
                    SqlCad = SqlCad & "VAL(FORMAT( (((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) - VAL(INGRESOS.CANTIDAD & '')) , '#.0000')) AS SALDO "
                    SqlCad = SqlCad & "FROM "
                    SqlCad = SqlCad & "(((((IF3ORDEN AS DET "
                    SqlCad = SqlCad & "LEFT JOIN IF5PLA AS PROD ON PROD.F5CODPRO = DET.F3CODPRO) "
                    SqlCad = SqlCad & "LEFT JOIN EF7MEDIDAS AS MED ON MED.F7CODMED = PROD.F7CODMED) "
                    SqlCad = SqlCad & "LEFT JOIN EF7MEDIDAS AS MED2 ON MED2.F7CODMED = DET.UNIDAD) "
                    SqlCad = SqlCad & "LEFT JOIN MEDIVENTAS AS MEDALTER ON MEDALTER.F5CODPRO = DET.F3CODPRO AND MEDALTER.F7CODMED = DET.UNIDAD) "
                    SqlCad = SqlCad & "LEFT JOIN "
                    SqlCad = SqlCad & "(SELECT "
                    SqlCad = SqlCad & "DET.F4NUMORD, "
                    SqlCad = SqlCad & "DET.COD_SOLICITUD, "
                    SqlCad = SqlCad & "DET.F5CODPROORIGINAL, "
                    SqlCad = SqlCad & "VAL(FORMAT( SUM(DET.F3CANPRO * IIF(DET.TIPO = 'S', -1, 1)) , '#.0000')) AS CANTIDAD "
                    SqlCad = SqlCad & "FROM "
                    SqlCad = SqlCad & "IF3VALES AS DET "
                    SqlCad = SqlCad & "LEFT JOIN IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM "
                    SqlCad = SqlCad & "WHERE "
                    SqlCad = SqlCad & "CAB.F1CODORI IN ('XC0') AND "
                    SqlCad = SqlCad & "DET.F4NUMORD <> '' AND "
                    SqlCad = SqlCad & "DET.COD_SOLICITUD <> '' "
                    SqlCad = SqlCad & "GROUP BY "
                    SqlCad = SqlCad & "DET.F4NUMORD, "
                    SqlCad = SqlCad & "DET.COD_SOLICITUD, "
                    SqlCad = SqlCad & "DET.F5CODPROORIGINAL) AS INGRESOS "
                    SqlCad = SqlCad & "ON INGRESOS.F4NUMORD = DET.F4NUMORD AND INGRESOS.COD_SOLICITUD = DET.COD_SOLICITUD AND INGRESOS.F5CODPROORIGINAL = DET.F3CODPRO) "
                    SqlCad = SqlCad & "LEFT JOIN TB_CABSOLICITUD AS CABPED ON CABPED.COD_SOLICITUD = DET.COD_SOLICITUD "
                    SqlCad = SqlCad & "WHERE "
                    SqlCad = SqlCad & "DET.F4LOCAL = 'OC' AND "
                    SqlCad = SqlCad & "DET.COD_SOLICITUD = '" & Trim(rstStockCPL!COD_SOLICITUD & "") & "' AND "
                    SqlCad = SqlCad & "DET.F3CODPRO = '" & Trim(rstStockCPL!F3CODPRO & "") & "' AND "
                    SqlCad = SqlCad & "VAL(FORMAT( (((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) - VAL(INGRESOS.CANTIDAD & '')) , '#.0000')) > 0 "
                    SqlCad = SqlCad & "ORDER BY "
                    SqlCad = SqlCad & "PROD.F5NOMPRO, "
                    SqlCad = SqlCad & "DET.COD_SOLICITUD, "
                    SqlCad = SqlCad & "DET.F4NUMORD"
                    
                    If rstDistribuir.State = 1 Then rstDistribuir.Close
    
                    rstDistribuir.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
                    If Not rstDistribuir.EOF Then
                        dblItem = 0
                        
                        pgbProceso2.Max = ModUtilitario.devuelveCantRegistros(rstDistribuir)
                        pgbProceso2.Value = 0
                        lblProceso2.Caption = "Re-distribuyendo Stock Comprometido Por Llegar (CPL) Actual..."
    
                        Do While Not rstDistribuir.EOF
                            SqlCad = vbNullString
                            SqlCad = SqlCad & "SELECT "
                            SqlCad = SqlCad & "* "
                            SqlCad = SqlCad & "FROM "
                            SqlCad = SqlCad & "IF3ORDEN "
                            SqlCad = SqlCad & "WHERE "
                            SqlCad = SqlCad & "F4LOCAL = 'OC' AND "
                            SqlCad = SqlCad & "F4NUMORD = '" & Trim(rstDistribuir!F4NUMORD & "") & "' AND "
                            SqlCad = SqlCad & "COD_SOLICITUD = '" & Trim(rstDistribuir!COD_SOLICITUD & "") & "' AND "
                            SqlCad = SqlCad & "F3CODPRO = '" & Trim(rstDistribuir!F3CODPRO & "") & "' AND "
                            SqlCad = SqlCad & "ITEM = '" & Trim(rstDistribuir!ITEM & "") & "'"
    
                            If rstOCDetalle.State = 1 Then rstOCDetalle.Close
    
                            rstOCDetalle.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
                            If Not rstOCDetalle.EOF Then
                                rstOCDetalle.MoveFirst
    
                                dblCantidadDestino = IIf(Val(rstDistribuir!SALDO & "") <= dblLiberarCompromiso, Val(rstDistribuir!SALDO & ""), dblLiberarCompromiso)
                                
                                Set objOrdenRegularizacionCPL = New ClsOrden
    
                                With objOrdenRegularizacionCPL
                                    .inicializarEntidadesDetalle
    
                                    .TipoOrden = "OC"
                                    .NumeroOrden = Trim(rstOCDetalle!F4NUMORD & "")
                                    .CodigoProducto = Trim(rstOCDetalle!F3CODPRO & "")
                                    .Requerimiento = vbNullString
    
                                    .obtenerConfigOrdenDetalleOnebyOne
    
                                    If .ITEM = 0 Then
                                        'Insertamos el Stock Por Llear Re-Distribuido A LIBRE
                                        .inicializarEntidadesDetalle
    
                                        .TipoOrden = "OC"
                                        .NumeroOrden = Trim(rstOCDetalle!F4NUMORD & "")
    
                                        .ITEM = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "VAL(ITEM & '')", "IF3ORDEN", "F4LOCAL", "OC", "T", "AND F4NUMORD = '" & .NumeroOrden & "' ORDER BY VAL(ITEM & '') DESC")) + 1
    
                                        .Requerimiento = vbNullString
                                        .CodigoProducto = Trim(rstOCDetalle!F3CODPRO & "")
                                        .CodigoFabricante = Trim(rstOCDetalle!F3CODFAB & "")
                                        .NombreProducto = Trim(rstOCDetalle!F5NOMPRO & "")
                                        .NombreProductoInterno = Trim(rstOCDetalle!F5NOMPRO_ING & "")
                                        .CodigoUM = Trim(rstOCDetalle!UNIDAD & "")
                                        .CodigoColor = Trim(rstOCDetalle!CODCOLOR & "")
                                        .ObservacionPorItem = "REDISTRIBUCION DE SCTOCK COMPROMETIDO POR LLEGAR AUTOMATICA A FAVOR DEL STOCK LIBRE POR LLEGAR."
    
                                        dblFactorUM = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5FACTOR", "MEDIVENTAS", "F5CODPRO", .CodigoProducto, "T", "AND F7CODMED = '" & .CodigoUM & "'"))
    
                                        If dblFactorUM = 0 Then dblFactorUM = 1
    
                                        dblCantidadDisponibleSegunOC = Val(Format((Val(rstDistribuir!SALDO & "") / (1 + (Val(rstOCDetalle!F3PORCDEMASIA & "") / 100))) / dblFactorUM, "#.0000"))
                                        dblCantidadDestinoSegunOC = Val(Format((dblCantidadDestino / (1 + (Val(rstOCDetalle!F3PORCDEMASIA & "") / 100))) / dblFactorUM, "#.0000"))
    
                                        .PorcentajeImpuesto = wwigv / 100
                                        .SignoImpuesto = 1
    
                                        .Cantidad = dblCantidadDestinoSegunOC
                                        .CantidadMaxima = dblCantidadDestinoSegunOC
    
                                        .PorcentajeDemasia = Val(rstOCDetalle!F3PORCDEMASIA & "") / 100
                                        .PrecioSinImpuesto = Val(rstOCDetalle!F3PRECOS & "")
                                        .PrecioConImpuesto = 0
                                        .PorcentajeDscto = Val(rstOCDetalle!F3PORDCT & "") / 100
                                        .TotalDscto = 0
                                        
                                        .Afecto = IIf(Trim(rstOCDetalle!F5AFECTO & "") = "*", True, False)
    
                                        .calculosPorItem
    
                                        .PorcentajeDemasia = Val(rstOCDetalle!F3PORCDEMASIA & "")
                                        .PorcentajeDscto = Val(rstOCDetalle!F3PORDCT & "")
    
                                        .CodigoGasto = Trim(rstOCDetalle!F3CUENTA & "")
                                        .CuentaContable = Trim(rstOCDetalle!F3GASTO & "")
    
                                        Select Case dblCantidadDisponibleSegunOC
                                            Case Is = Val(Format(Val(rstOCDetalle!F3CANPRO & ""), "#.0000"))
                                                If dblCantidadDestinoSegunOC < dblCantidadDisponibleSegunOC Then
    
                                                    .guardarOrdenDetalleOneByOne
    
                                                    Actualiza_Log .SQLSelectAlter, StrConexDbBancos
    
                                                    'Actualizamos la Cantidad del Item de Origen
                                                    .inicializarEntidadesDetalle
    
                                                    .PorcentajeImpuesto = wwigv / 100
                                                    .SignoImpuesto = 1
    
                                                    .Cantidad = Val(Format(Val(rstOCDetalle!F3CANPRO & "") - dblCantidadDestinoSegunOC, "#.0000"))
                                                    .CantidadMaxima = .Cantidad
    
                                                    .PorcentajeDemasia = Val(rstOCDetalle!F3PORCDEMASIA & "") / 100
                                                    .PrecioSinImpuesto = Val(rstOCDetalle!F3PRECOS & "")
                                                    .PrecioConImpuesto = 0
                                                    .PorcentajeDscto = Val(rstOCDetalle!F3PORDCT & "") / 100
                                                    .TotalDscto = 0
                                                    
                                                    .Afecto = IIf(Trim(rstOCDetalle!F5AFECTO & "") = "*", True, False)
    
                                                    .calculosPorItem
    
                                                    .PorcentajeDemasia = Val(rstOCDetalle!F3PORCDEMASIA & "")
                                                    .PorcentajeDscto = Val(rstOCDetalle!F3PORDCT & "")
                                                End If
    
                                                SqlCad = vbNullString
                                                SqlCad = SqlCad & "UPDATE "
                                                SqlCad = SqlCad & "IF3ORDEN "
                                                SqlCad = SqlCad & "SET "
    
                                                    If dblCantidadDestinoSegunOC < dblCantidadDisponibleSegunOC Then
                                                        SqlCad = SqlCad & "F5NOMPRO = TRIM(F5NOMPRO & ''), "
                                                        SqlCad = SqlCad & "F5NOMPRO_ING = TRIM(F5NOMPRO_ING & ''), "
                                                        SqlCad = SqlCad & "F3CANPRO = " & .Cantidad & ", "
                                                        SqlCad = SqlCad & "F3CANPRO2 = " & .CantidadMaxima & ", "
                                                        SqlCad = SqlCad & "F3PORCDEMASIA = " & .PorcentajeDemasia & ", "
                                                        SqlCad = SqlCad & "F5VALVTA = " & .BasePorItem & ", "
                                                        SqlCad = SqlCad & "F3IGV = " & .ImpuestoPorItem & ", "
                                                        SqlCad = SqlCad & "F3TOTAL = " & .TotalPorItem & ", "
                                                        SqlCad = SqlCad & "F3PORDCT = " & .PorcentajeDscto & ", "
                                                        SqlCad = SqlCad & "F3TOTDCT = " & .TotalDscto & " "
                                                    Else
                                                        SqlCad = SqlCad & "COD_SOLICITUD = '" & .Requerimiento & "' "
                                                    End If
    
                                                SqlCad = SqlCad & "WHERE "
                                                SqlCad = SqlCad & "F4LOCAL = 'OC' AND "
                                                SqlCad = SqlCad & "TRIM(F4NUMORD & '') = '" & Trim(rstDistribuir!F4NUMORD & "") & "' AND "
                                                SqlCad = SqlCad & "TRIM(COD_SOLICITUD & '') = '" & Trim(rstDistribuir!COD_SOLICITUD & "") & "' AND "
                                                SqlCad = SqlCad & "F3CODPRO = '" & Trim(rstDistribuir!F3CODPRO & "") & "' AND "
                                                SqlCad = SqlCad & "ITEM = '" & Trim(rstDistribuir!ITEM & "") & "'"
    
                                                cnn_dbbancos.Execute SqlCad
    
                                                Actualiza_Log SqlCad, StrConexDbBancos
                                                
                                                dblLiberarCompromiso = dblLiberarCompromiso - dblCantidadDestino
                                            Case Is < Val(Format(Val(rstOCDetalle!F3CANPRO & ""), "#.0000"))
    
                                                .guardarOrdenDetalleOneByOne
    
                                                Actualiza_Log .SQLSelectAlter, StrConexDbBancos
    
                                                'Actualizamos la Cantidad del Item de Origen
                                                .inicializarEntidadesDetalle
    
                                                .PorcentajeImpuesto = wwigv / 100
                                                .SignoImpuesto = 1
    
                                                .Cantidad = Val(Format(Val(rstOCDetalle!F3CANPRO & "") - dblCantidadDestinoSegunOC, "#.0000")) 'Val(Format(((Val(rstDistribuir!Cantidad & "") - dblCantidadDestino) / (1 + (Val(rstOCDetalle!F3PORCDEMASIA & "") / 100))) / dblFactorUM, "#.0000"))
                                                .CantidadMaxima = .Cantidad
    
                                                .PorcentajeDemasia = Val(rstOCDetalle!F3PORCDEMASIA & "") / 100
                                                .PrecioSinImpuesto = Val(rstOCDetalle!F3PRECOS & "")
                                                .PrecioConImpuesto = 0
                                                .PorcentajeDscto = Val(rstOCDetalle!F3PORDCT & "") / 100
                                                .TotalDscto = 0
    
                                                .Afecto = IIf(Trim(rstOCDetalle!F5AFECTO & "") = "*", True, False)
    
                                                .calculosPorItem
    
                                                .PorcentajeDemasia = Val(rstOCDetalle!F3PORCDEMASIA & "")
                                                .PorcentajeDscto = Val(rstOCDetalle!F3PORDCT & "")
    
                                                SqlCad = vbNullString
                                                SqlCad = SqlCad & "UPDATE "
                                                SqlCad = SqlCad & "IF3ORDEN "
                                                SqlCad = SqlCad & "SET "
    
                                                        SqlCad = SqlCad & "F3CANPRO = " & .Cantidad & ", "
                                                        SqlCad = SqlCad & "F3CANPRO2 = " & .CantidadMaxima & ", "
                                                        SqlCad = SqlCad & "F5VALVTA = " & .BasePorItem & ", "
                                                        SqlCad = SqlCad & "F3IGV = " & .ImpuestoPorItem & ", "
                                                        SqlCad = SqlCad & "F3TOTAL = " & .TotalPorItem & ", "
                                                        SqlCad = SqlCad & "F3PORDCT = " & .PorcentajeDscto & ", "
                                                        SqlCad = SqlCad & "F3TOTDCT = " & .TotalDscto & " "
    
                                                SqlCad = SqlCad & "WHERE "
                                                SqlCad = SqlCad & "F4LOCAL = 'OC' AND "
                                                SqlCad = SqlCad & "F4NUMORD = '" & Trim(rstDistribuir!F4NUMORD & "") & "' AND "
                                                SqlCad = SqlCad & "COD_SOLICITUD = '" & Trim(rstDistribuir!COD_SOLICITUD & "") & "' AND "
                                                SqlCad = SqlCad & "F3CODPRO = '" & Trim(rstDistribuir!F3CODPRO & "") & "' AND "
                                                SqlCad = SqlCad & "ITEM = '" & Trim(rstDistribuir!ITEM & "") & "'"
    
                                                cnn_dbbancos.Execute SqlCad
    
                                                Actualiza_Log SqlCad, StrConexDbBancos
    
                                                'bolRedistribucionEjecutada = True
                                                dblLiberarCompromiso = dblLiberarCompromiso - dblCantidadDestino
                                        End Select
                                    Else
                                        'Actualizamos el Stock Por Llegar del Item Existente como Libre en la Orden
                                        
                                        dblFactorUM = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5FACTOR", "MEDIVENTAS", "F5CODPRO", .CodigoProducto, "T", "AND F7CODMED = '" & .CodigoUM & "'"))
    
                                        If dblFactorUM = 0 Then dblFactorUM = 1
    
                                        dblCantidadDisponibleSegunOC = Val(Format((Val(rstDistribuir!SALDO & "") / (1 + (.PorcentajeDemasia / 100))) / dblFactorUM, "#.0000"))
                                        dblCantidadDestinoSegunOC = Val(Format((dblCantidadDestino / (1 + (.PorcentajeDemasia / 100))) / dblFactorUM, "#.0000"))
    
                                        .PorcentajeImpuesto = wwigv / 100
                                        .SignoImpuesto = 1
    
                                        .Cantidad = .Cantidad + dblCantidadDestinoSegunOC
                                        
                                        If .CantidadMaxima + dblCantidadDestinoSegunOC = .Cantidad Then
                                            .CantidadMaxima = .CantidadMaxima + dblCantidadDestinoSegunOC
                                        End If
                                        
                                        .PorcentajeDemasia = .PorcentajeDemasia / 100
                                        .PrecioConImpuesto = 0
                                        .PorcentajeDscto = .PorcentajeDscto / 100
                                        .TotalDscto = 0
                                        
                                        .calculosPorItem
    
                                        If .ObservacionPorItem = vbNullString Then
                                            .ObservacionPorItem = "STOCK LIBRE POR LLEGAR INCREMENTADO EN " & dblCantidadDestinoSegunOC & "."
                                        Else
                                            .ObservacionPorItem = left(.ObservacionPorItem & ", INCREMENTADO EN " & dblCantidadDestinoSegunOC & ".", 255)
                                        End If
                                        
                                        .PorcentajeDemasia = .PorcentajeDemasia * 100
                                        .PorcentajeDscto = .PorcentajeDscto * 100
                                        
                                        Select Case dblCantidadDisponibleSegunOC
                                            Case Is = Val(Format(Val(rstOCDetalle!F3CANPRO & ""), "#.0000"))
                                                
                                                If dblCantidadDestinoSegunOC = dblCantidadDisponibleSegunOC Then
                                                    .actualizarOrdenDetalleOneByOne
                                                    
                                                    Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                                                    
                                                    SqlCad = vbNullString
                                                    SqlCad = SqlCad & "DELETE "
                                                    SqlCad = SqlCad & "FROM "
                                                    SqlCad = SqlCad & "IF3ORDEN "
                                                    SqlCad = SqlCad & "WHERE "
                                                    SqlCad = SqlCad & "F4LOCAL = 'OC' AND "
                                                    SqlCad = SqlCad & "TRIM(F4NUMORD & '') = '" & Trim(rstDistribuir!F4NUMORD & "") & "' AND "
                                                    SqlCad = SqlCad & "TRIM(COD_SOLICITUD & '') = '" & Trim(rstDistribuir!COD_SOLICITUD & "") & "' AND "
                                                    SqlCad = SqlCad & "F3CODPRO = '" & Trim(rstDistribuir!F3CODPRO & "") & "' AND "
                                                    SqlCad = SqlCad & "ITEM = '" & Trim(rstDistribuir!ITEM & "") & "'"
        
                                                    cnn_dbbancos.Execute SqlCad
        
                                                    Actualiza_Log SqlCad, StrConexDbBancos
                                                ElseIf dblCantidadDestinoSegunOC < dblCantidadDisponibleSegunOC Then
                                                    
                                                    .actualizarOrdenDetalleOneByOne
    
                                                    Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                                                    
                                                    'Actualizamos la Cantidad del Item de Origen
                                                    .inicializarEntidadesDetalle
                                                    
                                                    .PorcentajeImpuesto = wwigv / 100
                                                    .SignoImpuesto = 1
                                                    
                                                    .Cantidad = Val(Format(Val(rstOCDetalle!F3CANPRO & "") - dblCantidadDestinoSegunOC, "#.0000")) 'Val(Format(((Val(rstDistribuir!Cantidad & "") - dblCantidadDestino) / (1 + (Val(rstOCDetalle!F3PORCDEMASIA & "") / 100))) / dblFactorUM, "#.0000"))
                                                    .CantidadMaxima = .Cantidad
                                                    
                                                    .PorcentajeDemasia = Val(rstOCDetalle!F3PORCDEMASIA & "") / 100
                                                    .PrecioSinImpuesto = Val(rstOCDetalle!F3PRECOS & "")
                                                    .PrecioConImpuesto = 0
                                                    .PorcentajeDscto = Val(rstOCDetalle!F3PORDCT & "") / 100
                                                    .TotalDscto = 0
                                                    
                                                    .Afecto = IIf(Trim(rstOCDetalle!F5AFECTO & "") = "*", True, False)
                                                    
                                                    .calculosPorItem
                                                    
                                                    .PorcentajeDemasia = Val(rstOCDetalle!F3PORCDEMASIA & "")
                                                    .PorcentajeDscto = Val(rstOCDetalle!F3PORDCT & "")
                                                    
                                                    SqlCad = vbNullString
                                                    SqlCad = SqlCad & "UPDATE "
                                                    SqlCad = SqlCad & "IF3ORDEN "
                                                    SqlCad = SqlCad & "SET "
                                                    
                                                    SqlCad = SqlCad & "F5NOMPRO = TRIM(F5NOMPRO & ''), "
                                                    SqlCad = SqlCad & "F5NOMPRO_ING = TRIM(F5NOMPRO_ING & ''), "
                                                    SqlCad = SqlCad & "F3CANPRO = " & .Cantidad & ", "
                                                    SqlCad = SqlCad & "F3CANPRO2 = " & .CantidadMaxima & ", "
                                                    SqlCad = SqlCad & "F3PORCDEMASIA = " & .PorcentajeDemasia & ", "
                                                    SqlCad = SqlCad & "F5VALVTA = " & .BasePorItem & ", "
                                                    SqlCad = SqlCad & "F3IGV = " & .ImpuestoPorItem & ", "
                                                    SqlCad = SqlCad & "F3TOTAL = " & .TotalPorItem & ", "
                                                    SqlCad = SqlCad & "F3PORDCT = " & .PorcentajeDscto & ", "
                                                    SqlCad = SqlCad & "F3TOTDCT = " & .TotalDscto & " "
                                                    
                                                    SqlCad = SqlCad & "WHERE "
                                                    SqlCad = SqlCad & "F4LOCAL = 'OC' AND "
                                                    SqlCad = SqlCad & "TRIM(F4NUMORD & '') = '" & Trim(rstDistribuir!F4NUMORD & "") & "' AND "
                                                    SqlCad = SqlCad & "TRIM(COD_SOLICITUD & '') = '" & Trim(rstDistribuir!COD_SOLICITUD & "") & "' AND "
                                                    SqlCad = SqlCad & "F3CODPRO = '" & Trim(rstDistribuir!F3CODPRO & "") & "' AND "
                                                    SqlCad = SqlCad & "ITEM = '" & Trim(rstDistribuir!ITEM & "") & "'"
        
                                                    cnn_dbbancos.Execute SqlCad
        
                                                    Actualiza_Log SqlCad, StrConexDbBancos
                                                End If
                                                
                                                dblLiberarCompromiso = dblLiberarCompromiso - dblCantidadDestino
                                            Case Is < Val(Format(Val(rstOCDetalle!F3CANPRO & ""), "#.0000"))
    
                                                .actualizarOrdenDetalleOneByOne
    
                                                Actualiza_Log .SQLSelectAlter, StrConexDbBancos
    
                                                'Actualizamos la Cantidad del Item de Origen
                                                .inicializarEntidadesDetalle
    
                                                .PorcentajeImpuesto = wwigv / 100
                                                .SignoImpuesto = 1
                                                
                                                .Cantidad = Val(Format(Val(rstOCDetalle!F3CANPRO & "") - dblCantidadDestinoSegunOC, "#.0000"))
                                                .CantidadMaxima = .Cantidad
    
                                                .PorcentajeDemasia = Val(rstOCDetalle!F3PORCDEMASIA & "") / 100
                                                .PrecioSinImpuesto = Val(rstOCDetalle!F3PRECOS & "")
                                                .PrecioConImpuesto = 0
                                                .PorcentajeDscto = Val(rstOCDetalle!F3PORDCT & "") / 100
                                                .TotalDscto = 0
                                                
                                                .Afecto = IIf(Trim(rstOCDetalle!F5AFECTO & "") = "*", True, False)
    
                                                .calculosPorItem
    
                                                .PorcentajeDemasia = Val(rstOCDetalle!F3PORCDEMASIA & "")
                                                .PorcentajeDscto = Val(rstOCDetalle!F3PORDCT & "")
                                                
                                                SqlCad = vbNullString
                                                SqlCad = SqlCad & "UPDATE "
                                                SqlCad = SqlCad & "IF3ORDEN "
                                                SqlCad = SqlCad & "SET "
                                                
                                                SqlCad = SqlCad & "F3CANPRO = " & .Cantidad & ", "
                                                SqlCad = SqlCad & "F3CANPRO2 = " & .CantidadMaxima & ", "
                                                SqlCad = SqlCad & "F5VALVTA = " & .BasePorItem & ", "
                                                SqlCad = SqlCad & "F3IGV = " & .ImpuestoPorItem & ", "
                                                SqlCad = SqlCad & "F3TOTAL = " & .TotalPorItem & ", "
                                                SqlCad = SqlCad & "F3PORDCT = " & .PorcentajeDscto & ", "
                                                SqlCad = SqlCad & "F3TOTDCT = " & .TotalDscto & " "
                                                
                                                SqlCad = SqlCad & "WHERE "
                                                SqlCad = SqlCad & "F4LOCAL = 'OC' AND "
                                                SqlCad = SqlCad & "F4NUMORD = '" & Trim(rstDistribuir!F4NUMORD & "") & "' AND "
                                                SqlCad = SqlCad & "COD_SOLICITUD = '" & Trim(rstDistribuir!COD_SOLICITUD & "") & "' AND "
                                                SqlCad = SqlCad & "F3CODPRO = '" & Trim(rstDistribuir!F3CODPRO & "") & "' AND "
                                                SqlCad = SqlCad & "ITEM = '" & Trim(rstDistribuir!ITEM & "") & "'"
    
                                                cnn_dbbancos.Execute SqlCad
    
                                                Actualiza_Log SqlCad, StrConexDbBancos
                                                
                                                dblLiberarCompromiso = dblLiberarCompromiso - dblCantidadDestino
                                        End Select
                                    End If
                                End With
                                
                                Set objOrdenRegularizacionCPL = Nothing
                            End If
                            
                            DoEvents
                            
                            pgbProceso2.Value = pgbProceso2.Value + 1
                            lblProceso2.Caption = "Re-distribuyendo Stock Comprometido Por Llegar (CPL) Actual... " & FormatPercent(pgbProceso2.Value / pgbProceso2.Max, 3)
                            
                            rstDistribuir.MoveNext
                        Loop
                            SqlCad = vbNullString
                    End If
                End If
                
                DoEvents
                
                pgbProceso1.Value = pgbProceso1.Value + 1
                lblProceso1.Caption = "Evaluando Stock Comprometido Por Llegar (CPL) ..." & FormatPercent(pgbProceso1.Value / pgbProceso1.Max, 3)
                
                rstStockCPL.MoveNext
            Loop
                Actualiza_Log "PROCESO 2.2 FINALIZO AL: " & Now, StrConexDbBancos
        End If
        
        chkProceso(1).BackColor = RGB(255, 153, 51)
    End If
    
    
    If CBool(chkProceso(2).Value) Then
        '---------------------------------------------------------------------------------------------
        '---------------------------------------------------------------------------------------------
        'Proceso de Auto-Correccion de Compromisos en Negativos
        
        abrirCnTemporal
        
        SqlCad = "DELETE FROM TMPUTILSTOCKCEA"
        
        cnDBTemp.Execute SqlCad
        
        SqlCad = vbNullString
        
        SqlCad = SqlCad & "INSERT INTO TMPUTILSTOCKCEA "
        
        SqlCad = SqlCad & "IN '" & wrutatemp & "Templus.mdb' "
        
        SqlCad = SqlCad & "SELECT "
        SqlCad = SqlCad & "F2CODALM, "
        SqlCad = SqlCad & "F5CODPRO, "
        SqlCad = SqlCad & "COD_SOLICITUD, "
        SqlCad = SqlCad & "VAL(FORMAT( SUM(VAL(FORMAT(VAL(F3CANPRO & ''), '#0.00')) * IIF(TIPO = 'S', -1, 1)) , '#0.00')) AS CANTIDAD "
        SqlCad = SqlCad & "FROM "
        SqlCad = SqlCad & "IF3VALES "
        SqlCad = SqlCad & "WHERE "
        SqlCad = SqlCad & "F2CODALM = '01' "
        
            If Trim(txtNroPedido.Text) <> vbNullString Then
                SqlCad = SqlCad & "AND TRIM(COD_SOLICITUD & '') = '" & Trim(txtNroPedido.Text) & "' "
            Else
                SqlCad = SqlCad & "AND TRIM(COD_SOLICITUD & '') <> '' "
            End If
            
            If Trim(txtCodProducto.Text) <> vbNullString Then
                SqlCad = SqlCad & "AND F5CODPRO = '" & Trim(txtCodProducto.Text) & "' "
            End If
            
        SqlCad = SqlCad & "GROUP BY "
        SqlCad = SqlCad & "F5CODPRO, "
        SqlCad = SqlCad & "F2CODALM, "
        SqlCad = SqlCad & "COD_SOLICITUD "
        SqlCad = SqlCad & "HAVING "
        SqlCad = SqlCad & "VAL(FORMAT( SUM(VAL(FORMAT(VAL(F3CANPRO & ''), '#0.00')) * IIF(TIPO = 'S', -1, 1)) , '#0.00')) < 0 "
        SqlCad = SqlCad & "ORDER BY "
        SqlCad = SqlCad & "F2CODALM, "
        SqlCad = SqlCad & "F5CODPRO, "
        SqlCad = SqlCad & "COD_SOLICITUD"
        
        abrirCnnDbBancos
        
        cnn_dbbancos.Execute SqlCad
        
        
        SqlCad = vbNullString
        SqlCad = SqlCad & "SELECT * FROM TMPUTILSTOCKCEA WHERE CANTIDAD < 0 ORDER BY CANTIDAD"
        
        abrirCnTemporal
        
        If rstStockCEA.State = 1 Then rstStockCEA.Close
        
        rstStockCEA.Open SqlCad, cnDBTemp, adOpenForwardOnly, adLockReadOnly
        
        If Not rstStockCEA.EOF Then
            pgbProceso2.Max = ModUtilitario.devuelveCantRegistros(rstStockCEA)
            pgbProceso2.Value = 0
            lblProceso2.Caption = "Regularización de Negativos en CEA..."
            
            Do While Not rstStockCEA.EOF
                SqlCad = vbNullString
                SqlCad = SqlCad & "SELECT "
                SqlCad = SqlCad & "F2CODALM, "
                SqlCad = SqlCad & "F5CODPRO, "
                SqlCad = SqlCad & "COD_SOLICITUD, "
                SqlCad = SqlCad & "VAL(FORMAT( SUM(VAL(FORMAT(VAL(F3CANPRO & ''), '#0.00')) * IIF(TIPO = 'S', -1, 1)) , '#0.00')) AS CANTIDAD "
                SqlCad = SqlCad & "FROM "
                SqlCad = SqlCad & "IF3VALES "
                SqlCad = SqlCad & "WHERE "
                SqlCad = SqlCad & "F2CODALM = '01' "
                SqlCad = SqlCad & "AND TRIM(COD_SOLICITUD & '') = '' "
                SqlCad = SqlCad & "AND F5CODPRO = '" & Trim(rstStockCEA!f5codpro & "") & "' "
                SqlCad = SqlCad & "GROUP BY "
                SqlCad = SqlCad & "F5CODPRO, "
                SqlCad = SqlCad & "F2CODALM, "
                SqlCad = SqlCad & "COD_SOLICITUD "
                SqlCad = SqlCad & "HAVING "
                SqlCad = SqlCad & "VAL(FORMAT( SUM(VAL(FORMAT(VAL(F3CANPRO & ''), '#0.00')) * IIF(TIPO = 'S', -1, 1)) , '#0.00')) > 0 "
                SqlCad = SqlCad & "ORDER BY "
                SqlCad = SqlCad & "F2CODALM, "
                SqlCad = SqlCad & "F5CODPRO, "
                SqlCad = SqlCad & "COD_SOLICITUD"
            
                If rstStockLEA.State = 1 Then rstStockLEA.Close
                
                rstStockLEA.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
            
                If Not rstStockLEA.EOF Then
                    If Val(rstStockLEA!Cantidad & "") > 0 And Val(rstStockLEA!Cantidad & "") >= Val(rstStockCEA!Cantidad & "") Then
                        
                        Set objValeRegularizacionCEA = New ClsVale
                        
                        With objValeRegularizacionCEA
                            .inicializarEntidades
                            
                            .CodigoAlmacen = Trim(rstStockLEA!f2codalm & "")
                            .NumeroVale = strNumValeProceso3_add
                            .TipoVale = "I"
                            
                            .Fecha = Format(Date, "dd/mm/yyyy")
                            .CodigoOrigen = "XCS"
                            .TipoCambio = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CAMBIO", "CAMBIOS", "FECHA", .Fecha, "F"))
                            
                                If .TipoCambio = 0 Then
                                    .TipoCambio = "4.05"
                                End If
                            
                            .CodigoMoneda = "S"
                            
                            .referencia = wnomcia
                            .observaciones = "PROCESO DE COMPROMISO AUTOMATICO MASIVO - CORRECION STOCK CEA."
                            
                            .FecReg = Format(Date, "Short Date")
                            .UsuReg = wusuario
                            .FecMod = Format(Date, "Short Date")
                            .UsuMod = wusuario
                            
                            If .guardarVale Then
                                Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                                
                                If strNumValeProceso3_add = vbNullString Then
                                    strNumValeProceso3_add = .NumeroVale
                                    
                                    'Borrar Detalle de Vale
                                    SqlCad = vbNullString
                                    SqlCad = "DELETE FROM IF3VALES WHERE F2CODALM = '" & .CodigoAlmacen & "' AND F4NUMVAL = '" & .NumeroVale & "'"
                                    
                                    cnn_dbbancos.Execute SqlCad
                                    
                                    Actualiza_Log SqlCad, cnn_dbbancos.ConnectionString
                                End If
                                
                                .inicializarEntidadesDetalle
        
                                .NumeroOrdenCompra = vbNullString
                                .Requerimiento = vbNullString
                                
                                .CodigoProducto = Trim(rstStockLEA!f5codpro & "")
                                .CodigoProductoOriginal = Trim(rstStockLEA!f5codpro & "")
                                .Cantidad = Abs(Val(rstStockCEA!Cantidad & "")) * -1
                                
                                
                                dblItem = dblItem + 1
                                
                                .ITEM = dblItem
                                
                                .guardarValeDetalleOneByOne
                                
                                Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                                
                                .inicializarEntidadesDetalle
                                
                                dblItem = dblItem + 1
                                
                                .Requerimiento = Trim(rstStockCEA!COD_SOLICITUD & "")
                                
                                .CodigoProducto = Trim(rstStockLEA!f5codpro & "")
                                .CodigoProductoOriginal = Trim(rstStockLEA!f5codpro & "")
                                .Cantidad = Abs(Val(rstStockCEA!Cantidad & ""))
                                .ITEM = dblItem
                                
                                .guardarValeDetalleOneByOne
                                
                                Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                            End If
                        End With
                        
                        Set objValeRegularizacionCEA = Nothing
                    End If
                End If
                
                DoEvents
                
                pgbProceso2.Value = pgbProceso2.Value + 1
                lblProceso2.Caption = "Regularización de Negativos en CEA..." & FormatPercent(pgbProceso2.Value / pgbProceso2.Max, 3)
                
                rstStockCEA.MoveNext
            Loop
                Actualiza_Log "PROCESO 3.1.1 FINALIZO AL: " & Now, StrConexDbBancos
        End If
        '---------------------------------------------------------------------------------------------
        '---------------------------------------------------------------------------------------------
        
        'Actualizar Stock CEA y CPL en Resumen de Produccion (Posterior a la Limpieza de Stock CEA y CPL)
        
        abrirCnTemporal
        
        SqlCad = "DELETE FROM TMPUTILSTOCKCEA"
        
        cnDBTemp.Execute SqlCad
        
        SqlCad = vbNullString
        
        SqlCad = SqlCad & "INSERT INTO TMPUTILSTOCKCEA "
        
        SqlCad = SqlCad & "IN '" & wrutatemp & "Templus.mdb' "
        
        SqlCad = SqlCad & "SELECT "
        SqlCad = SqlCad & "F2CODALM, "
        SqlCad = SqlCad & "F5CODPRO, "
        SqlCad = SqlCad & "COD_SOLICITUD, "
        SqlCad = SqlCad & "VAL(FORMAT( SUM(VAL(FORMAT(VAL(F3CANPRO & ''), '#0.00')) * IIF(TIPO = 'S', -1, 1)) , '#0.00')) AS CANTIDAD "
        SqlCad = SqlCad & "FROM "
        SqlCad = SqlCad & "IF3VALES "
        SqlCad = SqlCad & "WHERE "
        SqlCad = SqlCad & "F2CODALM = '01' "
        
            If Trim(txtNroPedido.Text) <> vbNullString Then
                SqlCad = SqlCad & "AND TRIM(COD_SOLICITUD & '') = '" & Trim(txtNroPedido.Text) & "' "
            Else
                SqlCad = SqlCad & "AND TRIM(COD_SOLICITUD & '') <> '' "
            End If
            
            If Trim(txtCodProducto.Text) <> vbNullString Then
                SqlCad = SqlCad & "AND F5CODPRO = '" & Trim(txtCodProducto.Text) & "' "
            End If
            
        SqlCad = SqlCad & "GROUP BY "
        SqlCad = SqlCad & "F5CODPRO, "
        SqlCad = SqlCad & "F2CODALM, "
        SqlCad = SqlCad & "COD_SOLICITUD "
        SqlCad = SqlCad & "HAVING "
        SqlCad = SqlCad & "VAL(FORMAT( SUM(VAL(FORMAT(VAL(F3CANPRO & ''), '#0.00')) * IIF(TIPO = 'S', -1, 1)) , '#0.00')) > 0 "
        SqlCad = SqlCad & "ORDER BY "
        SqlCad = SqlCad & "F2CODALM, "
        SqlCad = SqlCad & "F5CODPRO, "
        SqlCad = SqlCad & "COD_SOLICITUD"
        
        abrirCnnDbBancos
        
        cnn_dbbancos.Execute SqlCad
        
        SqlCad = vbNullString
        SqlCad = SqlCad & "SELECT * FROM TMPUTILSTOCKCEA WHERE CANTIDAD > 0"
        
        abrirCnTemporal
        
        If rstStockCEA.State = 1 Then rstStockCEA.Close
        
        rstStockCEA.Open SqlCad, cnDBTemp, adOpenForwardOnly, adLockReadOnly
        
        If Not rstStockCEA.EOF Then
            pgbProceso1.Max = ModUtilitario.devuelveCantRegistros(rstStockCEA)
            pgbProceso1.Value = 0
            lblProceso1.Caption = "Descargando Stock Comprometido En Almacen (CEA) Actual..."
            
            Do While Not rstStockCEA.EOF
                SqlCad = vbNullString
                SqlCad = SqlCad & "UPDATE "
                SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTOPRODUCCION "
                SqlCad = SqlCad & "SET "
                SqlCad = SqlCad & "COMPROMISOEAG = " & Val(rstStockCEA!Cantidad & "") & " "
                SqlCad = SqlCad & "WHERE "
                SqlCad = SqlCad & "NROPEDIDO = '" & Trim(rstStockCEA!COD_SOLICITUD & "") & "' AND "
                SqlCad = SqlCad & "CODPRODUCTO = '" & Trim(rstStockCEA!f5codpro & "") & "'"
                
                'abrirCnTemporal
                
                cnDBTemp.Execute SqlCad
                
                DoEvents
                
                pgbProceso1.Value = pgbProceso1.Value + 1
                lblProceso1.Caption = "Descargando Stock Comprometido En Almacen (CEA)..." & FormatPercent(pgbProceso1.Value / pgbProceso1.Max, 3)
                
                rstStockCEA.MoveNext
            Loop
                Actualiza_Log "PROCESO 3.1.2 FINALIZO AL: " & Now, StrConexDbBancos
        End If
        
        abrirCnTemporal
        
        SqlCad = "DELETE FROM TMPUTILSTOCKCPL"
        
        cnDBTemp.Execute SqlCad
        
        SqlCad = vbNullString
        
        SqlCad = SqlCad & "INSERT INTO TMPUTILSTOCKCPL "
        
        SqlCad = SqlCad & "IN '" & wrutatemp & "Templus.mdb' "
        
        SqlCad = SqlCad & "SELECT "
        SqlCad = SqlCad & "DET.F3CODPRO, "
        SqlCad = SqlCad & "DET.COD_SOLICITUD, "
        SqlCad = SqlCad & "SUM( VAL(FORMAT( (((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) - VAL(INGRESOS.CANTIDAD & '')) , '#.0000')) ) AS SALDO "
        SqlCad = SqlCad & "FROM "
        SqlCad = SqlCad & "(((((IF3ORDEN AS DET "
        SqlCad = SqlCad & "LEFT JOIN IF5PLA AS PROD ON PROD.F5CODPRO = DET.F3CODPRO) "
        SqlCad = SqlCad & "LEFT JOIN EF7MEDIDAS AS MED ON MED.F7CODMED = PROD.F7CODMED) "
        SqlCad = SqlCad & "LEFT JOIN EF7MEDIDAS AS MED2 ON MED2.F7CODMED = DET.UNIDAD) "
        SqlCad = SqlCad & "LEFT JOIN MEDIVENTAS AS MEDALTER ON MEDALTER.F5CODPRO = DET.F3CODPRO AND MEDALTER.F7CODMED = DET.UNIDAD) "
        SqlCad = SqlCad & "LEFT JOIN "
        SqlCad = SqlCad & "(SELECT "
        SqlCad = SqlCad & "DET.F4NUMORD, "
        SqlCad = SqlCad & "DET.COD_SOLICITUD, "
        SqlCad = SqlCad & "DET.F5CODPROORIGINAL, "
        SqlCad = SqlCad & "VAL(FORMAT( SUM(DET.F3CANPRO * IIF(DET.TIPO = 'S', -1, 1)) , '#.0000')) AS CANTIDAD "
        SqlCad = SqlCad & "FROM "
        SqlCad = SqlCad & "IF3VALES AS DET "
        SqlCad = SqlCad & "LEFT JOIN IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM "
        SqlCad = SqlCad & "WHERE "
        SqlCad = SqlCad & "CAB.F1CODORI IN ('XC0') AND "
        SqlCad = SqlCad & "DET.F4NUMORD <> '' AND "
        SqlCad = SqlCad & "DET.COD_SOLICITUD <> '' "
        SqlCad = SqlCad & "GROUP BY "
        SqlCad = SqlCad & "DET.F4NUMORD, "
        SqlCad = SqlCad & "DET.COD_SOLICITUD, "
        SqlCad = SqlCad & "DET.F5CODPROORIGINAL) AS INGRESOS "
        SqlCad = SqlCad & "ON INGRESOS.F4NUMORD = DET.F4NUMORD AND INGRESOS.COD_SOLICITUD = DET.COD_SOLICITUD AND INGRESOS.F5CODPROORIGINAL = DET.F3CODPRO) "
        SqlCad = SqlCad & "LEFT JOIN TB_CABSOLICITUD AS CABPED ON CABPED.COD_SOLICITUD = DET.COD_SOLICITUD "
        SqlCad = SqlCad & "WHERE "
        SqlCad = SqlCad & "DET.F4LOCAL = 'OC' AND "
        SqlCad = SqlCad & "VAL(FORMAT( (((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) - VAL(INGRESOS.CANTIDAD & '')) , '#.0000')) > 0 "
        
            If Trim(txtNroPedido.Text) <> vbNullString Then
                SqlCad = SqlCad & "AND DET.COD_SOLICITUD = '" & Trim(txtNroPedido.Text) & "' "
            Else
                SqlCad = SqlCad & "AND DET.COD_SOLICITUD <> '' "
            End If
            
            If Trim(txtCodProducto.Text) <> vbNullString Then
                SqlCad = SqlCad & "AND DET.F3CODPRO = '" & Trim(txtCodProducto.Text) & "' "
            End If
        
        SqlCad = SqlCad & "GROUP BY "
        SqlCad = SqlCad & "DET.F3CODPRO, "
        SqlCad = SqlCad & "PROD.F5NOMPRO, "
        SqlCad = SqlCad & "DET.COD_SOLICITUD "
        SqlCad = SqlCad & "ORDER BY "
        SqlCad = SqlCad & "PROD.F5NOMPRO, "
        SqlCad = SqlCad & "DET.COD_SOLICITUD"
        
        abrirCnnDbBancos
        
        cnn_dbbancos.Execute SqlCad
        
        SqlCad = vbNullString
        SqlCad = SqlCad & "SELECT * FROM TMPUTILSTOCKCPL"
        
        abrirCnTemporal
        
        If rstStockCPL.State = 1 Then rstStockCPL.Close
    
        rstStockCPL.Open SqlCad, cnDBTemp, adOpenForwardOnly, adLockReadOnly
        
        If Not rstStockCPL.EOF Then
            pgbProceso1.Max = ModUtilitario.devuelveCantRegistros(rstStockCPL)
            pgbProceso1.Value = 0
            lblProceso1.Caption = "Descargando Stock Comprometido Por Llegar (CPL) Actual..."
    
            dblCantidadProdObservado = 0
            
            Do While Not rstStockCPL.EOF
                SqlCad = vbNullString
                SqlCad = SqlCad & "UPDATE "
                SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTOPRODUCCION "
                SqlCad = SqlCad & "SET "
                SqlCad = SqlCad & "COMPROMISOPLG = " & Val(rstStockCPL!SALDO & "") & " "
                SqlCad = SqlCad & "WHERE "
                SqlCad = SqlCad & "NROPEDIDO = '" & Trim(rstStockCPL!COD_SOLICITUD & "") & "' AND "
                SqlCad = SqlCad & "CODPRODUCTO = '" & Trim(rstStockCPL!F3CODPRO & "") & "'"
                
                'abrirCnTemporal
                
                cnDBTemp.Execute SqlCad
                
                DoEvents
                
                pgbProceso1.Value = pgbProceso1.Value + 1
                lblProceso1.Caption = "Descargando Stock Comprometido Por Llegar (CPL) ..." & FormatPercent(pgbProceso1.Value / pgbProceso1.Max, 3)
                
                rstStockCPL.MoveNext
            Loop
                Actualiza_Log "PROCESO 3.2 FINALIZO AL: " & Now, StrConexDbBancos
        End If
        
        chkProceso(2).BackColor = RGB(255, 153, 51)
    End If
    
    
    
    
    If CBool(chkProceso(3).Value) Then
        abrirCnTemporal
        
        'Seleccionar Saldo de Produccion para su Compromiso Automatico - Hasta el Corte de Hoy (Reposicion de Compromisos)
        SqlCad = vbNullString
        SqlCad = SqlCad & "SELECT "
        SqlCad = SqlCad & "NROPEDIDO, "
        SqlCad = SqlCad & "FENTREGA, "
        SqlCad = SqlCad & "CODPRODUCTO, "
        SqlCad = SqlCad & "NOMPRODUCTO, "
        SqlCad = SqlCad & "SUM(SALDO) - VAL(FORMAT(AVG(COMPROMISOEAG) + AVG(COMPROMISOPLG), '#0.0000')) AS SALDOTOTAL "
        SqlCad = SqlCad & "FROM "
        SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTOPRODUCCION "
        SqlCad = SqlCad & "WHERE "
        SqlCad = SqlCad & "SALDO > 0 AND "
        SqlCad = SqlCad & "CVDATE(FENTREGA) BETWEEN CVDATE('" & strFechaCorteValidezCompromiso & "') AND CVDATE('" & strFechaCorteActual & "') "
            
            If Trim(txtNroPedido.Text) <> vbNullString Then
                SqlCad = SqlCad & "AND NROPEDIDO = '" & Trim(txtNroPedido.Text) & "' "
            End If
            
            If Trim(txtCodProducto.Text) <> vbNullString Then
                SqlCad = SqlCad & "AND CODPRODUCTO = '" & Trim(txtCodProducto.Text) & "' "
            End If
            
        SqlCad = SqlCad & "GROUP BY "
        SqlCad = SqlCad & "NROPEDIDO, "
        SqlCad = SqlCad & "FENTREGA, "
        SqlCad = SqlCad & "CODPRODUCTO, "
        SqlCad = SqlCad & "NOMPRODUCTO "
        SqlCad = SqlCad & "ORDER BY "
        SqlCad = SqlCad & "NOMPRODUCTO, "
        SqlCad = SqlCad & "FENTREGA, "
        SqlCad = SqlCad & "NROPEDIDO"
        
        If rstResProd.State = 1 Then rstResProd.Close
        
        rstResProd.Open SqlCad, cnDBTemp, adOpenForwardOnly, adLockReadOnly
    
        If Not rstResProd.EOF Then
            pgbProceso1.Max = ModUtilitario.devuelveCantRegistros(rstResProd)
            pgbProceso1.Value = 0
            lblProceso1.Caption = "Compromiso Automatico 1/2..."
            
            dblCantidadComprometer = 0
            
            strNumValeProceso2 = vbNullString
            
            Do While Not rstResProd.EOF
                dblCantidadComprometer = Val(rstResProd!SALDOTOTAL & "")
                
                If dblCantidadComprometer > 0 Then
                    SqlCad = vbNullString
                    SqlCad = SqlCad & "SELECT "
                    SqlCad = SqlCad & "F2CODALM, "
                    SqlCad = SqlCad & "F5CODPRO, "
                    SqlCad = SqlCad & "COD_SOLICITUD, "
                    SqlCad = SqlCad & "VAL(FORMAT( SUM(VAL(FORMAT(VAL(F3CANPRO & ''), '#0.00')) * IIF(TIPO = 'S', -1, 1)) , '#0.00')) AS CANTIDAD "
                    SqlCad = SqlCad & "FROM "
                    SqlCad = SqlCad & "IF3VALES "
                    SqlCad = SqlCad & "WHERE "
                    SqlCad = SqlCad & "F2CODALM = '01' "
                    SqlCad = SqlCad & "AND TRIM(COD_SOLICITUD & '') = '' "
                    SqlCad = SqlCad & "AND F5CODPRO = '" & Trim(rstResProd!CodProducto & "") & "' "
                    SqlCad = SqlCad & "GROUP BY "
                    SqlCad = SqlCad & "F5CODPRO, "
                    SqlCad = SqlCad & "F2CODALM, "
                    SqlCad = SqlCad & "COD_SOLICITUD "
                    SqlCad = SqlCad & "HAVING "
                    SqlCad = SqlCad & "VAL(FORMAT( SUM(VAL(FORMAT(VAL(F3CANPRO & ''), '#0.00')) * IIF(TIPO = 'S', -1, 1)) , '#0.00')) > 0 "
                    SqlCad = SqlCad & "ORDER BY "
                    SqlCad = SqlCad & "F2CODALM, "
                    SqlCad = SqlCad & "F5CODPRO, "
                    SqlCad = SqlCad & "COD_SOLICITUD"
                
                    If rstStockLEA.State = 1 Then rstStockLEA.Close
                    
                    rstStockLEA.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
                
                    If Not rstStockLEA.EOF Then
                        If Val(rstStockLEA!Cantidad & "") > 0 Then
                            
                            Set objValeRegularizacionCEA = New ClsVale
                            
                            With objValeRegularizacionCEA
                                
                                .inicializarEntidades
                                
                                .CodigoAlmacen = Trim(rstStockLEA!f2codalm & "")
                                .NumeroVale = strNumValeProceso2
                                .TipoVale = "I"
                                
                                .Fecha = Format(Date, "dd/mm/yyyy")
                                .CodigoOrigen = "XCS"
                                .TipoCambio = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CAMBIO", "CAMBIOS", "FECHA", .Fecha, "F"))
                                
                                    If .TipoCambio = 0 Then
                                        .TipoCambio = "4.05"
                                    End If
                                
                                .CodigoMoneda = "S"
                                
                                .referencia = wnomcia
                                .observaciones = "PROCESO DE COMPROMISO AUTOMATICO MASIVO."
                                
                                .FecReg = Format(Date, "Short Date")
                                .UsuReg = wusuario
                                .FecMod = Format(Date, "Short Date")
                                .UsuMod = wusuario
                                
                                If .guardarVale Then
                                    Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                                    
                                    If strNumValeProceso2 = vbNullString Then
                                        strNumValeProceso2 = .NumeroVale
                                        
                                        'Borrar Detalle de Vale
                                        SqlCad = vbNullString
                                        SqlCad = "DELETE FROM IF3VALES WHERE F2CODALM = '" & .CodigoAlmacen & "' AND F4NUMVAL = '" & .NumeroVale & "'"
                
                                        cnn_dbbancos.Execute SqlCad
                                        Actualiza_Log SqlCad, cnn_dbbancos.ConnectionString
                                    End If
                                    
                                    .inicializarEntidadesDetalle
            
                                    .NumeroOrdenCompra = vbNullString
                                    .Requerimiento = vbNullString
                                    
                                    .CodigoProducto = Trim(rstStockLEA!f5codpro & "")
                                    .CodigoProductoOriginal = Trim(rstStockLEA!f5codpro & "")
                                    .Cantidad = IIf(Val(rstStockLEA!Cantidad & "") >= dblCantidadComprometer, dblCantidadComprometer, Val(rstStockLEA!Cantidad & "")) * -1
                                    
                                    
                                    dblItem = dblItem + 1
                                    
                                    .ITEM = dblItem
                                    
                                    .guardarValeDetalleOneByOne
                                    
                                    Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                                    
                                    .inicializarEntidadesDetalle
                                    
                                    dblItem = dblItem + 1
                                    
                                    .Requerimiento = Trim(rstResProd!NroPedido & "")
                                    
                                    .CodigoProducto = Trim(rstStockLEA!f5codpro & "")
                                    .CodigoProductoOriginal = Trim(rstStockLEA!f5codpro & "")
                                    .Cantidad = IIf(Val(rstStockLEA!Cantidad & "") >= dblCantidadComprometer, dblCantidadComprometer, Val(rstStockLEA!Cantidad & ""))
                                    .ITEM = dblItem
                                    
                                    .guardarValeDetalleOneByOne
                                    
                                    Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                                    
                                    dblCantidadComprometer = dblCantidadComprometer - .Cantidad
                                End If
                            End With
                            
                            Set objValeRegularizacionCEA = Nothing
                        End If
                    End If
                    
                    If dblCantidadComprometer > 0 Then
                        SqlCad = vbNullString
                        SqlCad = SqlCad & "SELECT "
                        SqlCad = SqlCad & "DET.ITEM, "
                        SqlCad = SqlCad & "DET.F3CODPRO, "
                        SqlCad = SqlCad & "DET.COD_SOLICITUD, "
                        SqlCad = SqlCad & "DET.F4NUMORD, "
                        SqlCad = SqlCad & "VAL(FORMAT( (((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) - VAL(INGRESOS.CANTIDAD & '')) , '#.0000')) AS SALDO "
                        SqlCad = SqlCad & "FROM "
                        SqlCad = SqlCad & "(((((IF3ORDEN AS DET "
                        SqlCad = SqlCad & "LEFT JOIN IF5PLA AS PROD ON PROD.F5CODPRO = DET.F3CODPRO) "
                        SqlCad = SqlCad & "LEFT JOIN EF7MEDIDAS AS MED ON MED.F7CODMED = PROD.F7CODMED) "
                        SqlCad = SqlCad & "LEFT JOIN EF7MEDIDAS AS MED2 ON MED2.F7CODMED = DET.UNIDAD) "
                        SqlCad = SqlCad & "LEFT JOIN MEDIVENTAS AS MEDALTER ON MEDALTER.F5CODPRO = DET.F3CODPRO AND MEDALTER.F7CODMED = DET.UNIDAD) "
                        SqlCad = SqlCad & "LEFT JOIN "
                        SqlCad = SqlCad & "(SELECT "
                        SqlCad = SqlCad & "DET.F4NUMORD, "
                        SqlCad = SqlCad & "DET.COD_SOLICITUD, "
                        SqlCad = SqlCad & "DET.F5CODPROORIGINAL, "
                        SqlCad = SqlCad & "VAL(FORMAT( SUM(DET.F3CANPRO * IIF(DET.TIPO = 'S', -1, 1)) , '#.0000')) AS CANTIDAD "
                        SqlCad = SqlCad & "FROM "
                        SqlCad = SqlCad & "IF3VALES AS DET "
                        SqlCad = SqlCad & "LEFT JOIN IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM "
                        SqlCad = SqlCad & "WHERE "
                        SqlCad = SqlCad & "CAB.F1CODORI IN ('XC0') AND "
                        SqlCad = SqlCad & "DET.F4NUMORD <> '' AND "
                        SqlCad = SqlCad & "DET.COD_SOLICITUD = '' AND "
                        SqlCad = SqlCad & "DET.F5CODPROORIGINAL = '" & Trim(rstResProd!CodProducto & "") & "' "
                        SqlCad = SqlCad & "GROUP BY "
                        SqlCad = SqlCad & "DET.F4NUMORD, "
                        SqlCad = SqlCad & "DET.COD_SOLICITUD, "
                        SqlCad = SqlCad & "DET.F5CODPROORIGINAL) AS INGRESOS "
                        SqlCad = SqlCad & "ON INGRESOS.F4NUMORD = DET.F4NUMORD AND INGRESOS.COD_SOLICITUD = DET.COD_SOLICITUD AND INGRESOS.F5CODPROORIGINAL = DET.F3CODPRO) "
                        SqlCad = SqlCad & "LEFT JOIN TB_CABSOLICITUD AS CABPED ON CABPED.COD_SOLICITUD = DET.COD_SOLICITUD "
                        SqlCad = SqlCad & "WHERE "
                        SqlCad = SqlCad & "DET.F4LOCAL = 'OC' AND "
                        SqlCad = SqlCad & "DET.COD_SOLICITUD = '' AND "
                        SqlCad = SqlCad & "DET.F3CODPRO = '" & Trim(rstResProd!CodProducto & "") & "' AND "
                        SqlCad = SqlCad & "VAL(FORMAT( (((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) - VAL(INGRESOS.CANTIDAD & '')) , '#.0000')) > 0 "
                        SqlCad = SqlCad & "ORDER BY "
                        SqlCad = SqlCad & "PROD.F5NOMPRO, "
                        SqlCad = SqlCad & "DET.COD_SOLICITUD, "
                        SqlCad = SqlCad & "DET.F4NUMORD"
        
                        If rstStockLPL.State = 1 Then rstStockLPL.Close
        
                        rstStockLPL.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
        
                        If Not rstStockLPL.EOF Then
                            
                            Do While Not rstStockLPL.EOF
                                SqlCad = vbNullString
                                SqlCad = SqlCad & "SELECT "
                                SqlCad = SqlCad & "* "
                                SqlCad = SqlCad & "FROM "
                                SqlCad = SqlCad & "IF3ORDEN "
                                SqlCad = SqlCad & "WHERE "
                                SqlCad = SqlCad & "F4LOCAL = 'OC' AND "
                                SqlCad = SqlCad & "F4NUMORD = '" & Trim(rstStockLPL!F4NUMORD & "") & "' AND "
                                SqlCad = SqlCad & "COD_SOLICITUD = '" & Trim(rstStockLPL!COD_SOLICITUD & "") & "' AND "
                                SqlCad = SqlCad & "F3CODPRO = '" & Trim(rstStockLPL!F3CODPRO & "") & "' AND "
                                SqlCad = SqlCad & "ITEM = '" & Trim(rstStockLPL!ITEM & "") & "'"
                                
                                If rstOCDetalle.State = 1 Then rstOCDetalle.Close
        
                                rstOCDetalle.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
        
                                If Not rstOCDetalle.EOF Then
                                    rstOCDetalle.MoveFirst
        
                                    Set objOrdenRegularizacionCPL = New ClsOrden
        
                                    With objOrdenRegularizacionCPL
                                        'Insertamos el Stock Re-Distribuido
                                        .inicializarEntidadesDetalle
        
                                        .TipoOrden = "OC"
                                        .NumeroOrden = Trim(rstOCDetalle!F4NUMORD & "")
                                        
                                        .ITEM = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "VAL(ITEM & '')", "IF3ORDEN", "F4LOCAL", "OC", "T", "AND F4NUMORD = '" & .NumeroOrden & "' ORDER BY VAL(ITEM & '') DESC")) + 1
                                        
                                        .Requerimiento = Trim(rstResProd!NroPedido & "")
                                        
                                        .CodigoProducto = Trim(rstOCDetalle!F3CODPRO & "")
                                        .CodigoFabricante = Trim(rstOCDetalle!F3CODFAB & "")
                                        .NombreProducto = Trim(rstOCDetalle!F5NOMPRO & "")
                                        .NombreProductoInterno = Trim(rstOCDetalle!F5NOMPRO_ING & "")
                                        .CodigoUM = Trim(rstOCDetalle!UNIDAD & "")
                                        .CodigoColor = Trim(rstOCDetalle!CODCOLOR & "")
                                        .ObservacionPorItem = "REDISTRIBUCION DE SCTOCK AUTOMATICO MASIVO A FAVOR DEL PEDIDO " & .Requerimiento & "."
        
                                        dblFactorUM = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5FACTOR", "MEDIVENTAS", "F5CODPRO", .CodigoProducto, "T", "AND F7CODMED = '" & .CodigoUM & "'"))
        
                                        If dblFactorUM = 0 Then dblFactorUM = 1
        
                                        dblCantidadDisponibleSegunOC = Val(Format((Val(rstStockLPL!SALDO & "") / (1 + (Val(rstOCDetalle!F3PORCDEMASIA & "") / 100))) / dblFactorUM, "#.0000"))
                                        dblCantidadDestinoSegunOC = Val(Format((IIf(Val(rstStockLPL!SALDO & "") >= dblCantidadComprometer, dblCantidadComprometer, Val(rstStockLPL!SALDO & "")) / (1 + (Val(rstOCDetalle!F3PORCDEMASIA & "") / 100))) / dblFactorUM, "#.0000"))
                                        
                                        .PorcentajeImpuesto = wwigv / 100
                                        .SignoImpuesto = 1
        
                                        .Cantidad = dblCantidadDestinoSegunOC 'Val(Format((Val(rstResProd!CANTIDADDESTINO & "") / (1 + (Val(rstOCDetalle!F3PORCDEMASIA & "") / 100))) / dblFactorUM, "#.0000"))
        
                                        If Trim(rstResProd!NroPedido & "") <> vbNullString Then
                                            .CantidadMaxima = dblCantidadDestinoSegunOC '.Cantidad
                                        End If
        
                                        .PorcentajeDemasia = Val(rstOCDetalle!F3PORCDEMASIA & "") / 100
                                        .PrecioSinImpuesto = Val(rstOCDetalle!F3PRECOS & "")
                                        .PrecioConImpuesto = 0 'Val(rstOCDetalle!F3PREUNI & "")
                                        .PorcentajeDscto = Val(rstOCDetalle!F3PORDCT & "") / 100
                                        .TotalDscto = 0 'Val(rstOCDetalle!F3TOTDCT & "")
        
                                        .Afecto = IIf(Trim(rstOCDetalle!F5AFECTO & "") = "*", True, False)
        
                                        .calculosPorItem
        
                                        .PorcentajeDemasia = Val(rstOCDetalle!F3PORCDEMASIA & "")
                                        .PorcentajeDscto = Val(rstOCDetalle!F3PORDCT & "")
        
                                        .CodigoGasto = Trim(rstOCDetalle!F3CUENTA & "")
                                        .CuentaContable = Trim(rstOCDetalle!F3GASTO & "")
        
                                        Select Case dblCantidadDisponibleSegunOC
                                            Case Is = Val(Format(Val(rstOCDetalle!F3CANPRO & ""), "#.0000"))
                                                If dblCantidadDestinoSegunOC < dblCantidadDisponibleSegunOC Then
        
                                                    .guardarOrdenDetalleOneByOne
        
                                                    Actualiza_Log .SQLSelectAlter, StrConexDbBancos
        
                                                    'Actualizamos la Cantidad del Item de Origen
                                                    .inicializarEntidadesDetalle
        
                                                    .PorcentajeImpuesto = wwigv / 100
                                                    .SignoImpuesto = 1
        
                                                    .Cantidad = Val(Format(Val(rstOCDetalle!F3CANPRO & "") - dblCantidadDestinoSegunOC, "#.0000")) 'Val(Format(((Val(rstResProd!Cantidad & "") - Val(rstResProd!CANTIDADDESTINO & "")) / (1 + (Val(rstOCDetalle!F3PORCDEMASIA & "") / 100))) / dblFactorUM, "#.0000"))
        
                                                    If Trim(rstResProd!NroPedido & "") <> vbNullString Then
                                                        .CantidadMaxima = .Cantidad
                                                    End If
        
                                                    .PorcentajeDemasia = Val(rstOCDetalle!F3PORCDEMASIA & "") / 100
                                                    .PrecioSinImpuesto = Val(rstOCDetalle!F3PRECOS & "")
                                                    .PrecioConImpuesto = 0 'Val(rstOCDetalle!F3PREUNI & "")
                                                    .PorcentajeDscto = Val(rstOCDetalle!F3PORDCT & "") / 100
                                                    .TotalDscto = 0 'Val(rstOCDetalle!F3TOTDCT & "")
        
                                                    .Afecto = IIf(Trim(rstOCDetalle!F5AFECTO & "") = "*", True, False)
        
                                                    .calculosPorItem
        
                                                    .PorcentajeDemasia = Val(rstOCDetalle!F3PORCDEMASIA & "")
                                                    .PorcentajeDscto = Val(rstOCDetalle!F3PORDCT & "")
                                                End If
        
                                                SqlCad = vbNullString
                                                SqlCad = SqlCad & "UPDATE "
                                                SqlCad = SqlCad & "IF3ORDEN "
                                                SqlCad = SqlCad & "SET "
        
                                                    If dblCantidadDestinoSegunOC < dblCantidadDisponibleSegunOC Then
                                                        SqlCad = SqlCad & "F5NOMPRO = TRIM(F5NOMPRO & ''), "
                                                        SqlCad = SqlCad & "F5NOMPRO_ING = TRIM(F5NOMPRO_ING & ''), "
                                                        SqlCad = SqlCad & "F3CANPRO = " & .Cantidad & ", "
                                                        SqlCad = SqlCad & "F3CANPRO2 = " & .CantidadMaxima & ", "
                                                        SqlCad = SqlCad & "F5VALVTA = " & .BasePorItem & ", "
                                                        SqlCad = SqlCad & "F3IGV = " & .ImpuestoPorItem & ", "
                                                        SqlCad = SqlCad & "F3TOTAL = " & .TotalPorItem & ", "
                                                        SqlCad = SqlCad & "F3PORDCT = " & .PorcentajeDscto & ", "
                                                        SqlCad = SqlCad & "F3TOTDCT = " & .TotalDscto & " "
                                                    Else
                                                        SqlCad = SqlCad & "COD_SOLICITUD = '" & Trim(rstResProd!NroPedido & "") & "' "
                                                    End If
        
                                                SqlCad = SqlCad & "WHERE "
                                                SqlCad = SqlCad & "F4LOCAL = 'OC' AND "
                                                SqlCad = SqlCad & "TRIM(F4NUMORD & '') = '" & Trim(rstOCDetalle!F4NUMORD & "") & "' AND "
                                                SqlCad = SqlCad & "TRIM(COD_SOLICITUD & '') = '" & Trim(rstOCDetalle!COD_SOLICITUD & "") & "' AND "
                                                SqlCad = SqlCad & "F3CODPRO = '" & Trim(rstOCDetalle!F3CODPRO & "") & "'"
        
                                                cnn_dbbancos.Execute SqlCad
        
                                                Actualiza_Log SqlCad, StrConexDbBancos
                                            Case Is < Val(Format(Val(rstOCDetalle!F3CANPRO & ""), "#.0000"))
        
                                                .guardarOrdenDetalleOneByOne
        
                                                Actualiza_Log .SQLSelectAlter, StrConexDbBancos
        
                                                'Actualizamos la Cantidad del Item de Origen
                                                .inicializarEntidadesDetalle
        
                                                .PorcentajeImpuesto = wwigv / 100
                                                .SignoImpuesto = 1
        
                                                .Cantidad = Val(Format(Val(rstOCDetalle!F3CANPRO & "") - dblCantidadDestinoSegunOC, "#.0000")) 'Val(Format(((Val(rstResProd!Cantidad & "") - Val(rstResProd!CANTIDADDESTINO & "")) / (1 + (Val(rstOCDetalle!F3PORCDEMASIA & "") / 100))) / dblFactorUM, "#.0000"))
                                                .CantidadMaxima = .Cantidad
        
                                                .PorcentajeDemasia = Val(rstOCDetalle!F3PORCDEMASIA & "") / 100
                                                .PrecioSinImpuesto = Val(rstOCDetalle!F3PRECOS & "")
                                                .PrecioConImpuesto = 0 'Val(rstOCDetalle!F3PREUNI & "")
                                                .PorcentajeDscto = Val(rstOCDetalle!F3PORDCT & "") / 100
                                                .TotalDscto = 0 'Val(rstOCDetalle!F3TOTDCT & "")
        
                                                .Afecto = IIf(Trim(rstOCDetalle!F5AFECTO & "") = "*", True, False)
        
                                                .calculosPorItem
        
                                                .PorcentajeDemasia = Val(rstOCDetalle!F3PORCDEMASIA & "")
                                                .PorcentajeDscto = Val(rstOCDetalle!F3PORDCT & "")
        
                                                SqlCad = vbNullString
                                                SqlCad = SqlCad & "UPDATE "
                                                SqlCad = SqlCad & "IF3ORDEN "
                                                SqlCad = SqlCad & "SET "
        
                                                        SqlCad = SqlCad & "F3CANPRO = " & .Cantidad & ", "
                                                        SqlCad = SqlCad & "F3CANPRO2 = " & .CantidadMaxima & ", "
                                                        SqlCad = SqlCad & "F5VALVTA = " & .BasePorItem & ", "
                                                        SqlCad = SqlCad & "F3IGV = " & .ImpuestoPorItem & ", "
                                                        SqlCad = SqlCad & "F3TOTAL = " & .TotalPorItem & ", "
                                                        SqlCad = SqlCad & "F3PORDCT = " & .PorcentajeDscto & ", "
                                                        SqlCad = SqlCad & "F3TOTDCT = " & .TotalDscto & " "
        
                                                SqlCad = SqlCad & "WHERE "
                                                SqlCad = SqlCad & "F4LOCAL = 'OC' AND "
                                                SqlCad = SqlCad & "F4NUMORD = '" & Trim(rstOCDetalle!F4NUMORD & "") & "' AND "
                                                SqlCad = SqlCad & "COD_SOLICITUD = '" & Trim(rstOCDetalle!COD_SOLICITUD & "") & "' AND "
                                                SqlCad = SqlCad & "F3CODPRO = '" & Trim(rstOCDetalle!F3CODPRO & "") & "'"
        
                                                cnn_dbbancos.Execute SqlCad
        
                                                Actualiza_Log SqlCad, StrConexDbBancos
                                        End Select
                                    End With
        
                                    Set objOrdenRegularizacionCPL = Nothing
                                    
                                    dblCantidadComprometer = dblCantidadComprometer - IIf(Val(rstStockLPL!SALDO & "") >= dblCantidadComprometer, dblCantidadComprometer, Val(rstStockLPL!SALDO & ""))
                                End If
                                
                                If dblCantidadComprometer = 0 Then Exit Do
                                
                                rstStockLPL.MoveNext
                            Loop
                        End If
                    End If
                End If
                
                'Si no se compromete toda la cantidad en el Proceso de Re-Distribucion automatica a Favor del Pedido para su Reposicion de Compromiso
                'Adicionar un registro de Alerta con los datos del Pedido y Producto que estan pendientes de reposicion.
                If dblCantidadComprometer > 0 Then
                    SqlCad = vbNullString
                    SqlCad = SqlCad & "INSERT INTO SF3COMPROMISOAUTOMATICO("
                    SqlCad = SqlCad & "FECHAEJECUCION, "
                    SqlCad = SqlCad & "NROPEDIDO, "
                    SqlCad = SqlCad & "CLIENTE, "
                    SqlCad = SqlCad & "FECHAEMISION, "
                    SqlCad = SqlCad & "FECHAENTREGA, "
                    SqlCad = SqlCad & "VENDEDOR, "
                    SqlCad = SqlCad & "DIASPORVENCER, "
                    SqlCad = SqlCad & "CODPRODUCTO, "
                    SqlCad = SqlCad & "NOMPRODUCTO, "
                    SqlCad = SqlCad & "SALDOPORCOMPROMETER) "
                    SqlCad = SqlCad & "VALUES("
                    SqlCad = SqlCad & "CVDATE('" & strFechaCorteActual & "'), "
                    SqlCad = SqlCad & "'" & Trim(rstResProd!NroPedido & "") & "', "
                    SqlCad = SqlCad & "'" & Trim(rstResProd!CLIENTE & "") & "', "
                    SqlCad = SqlCad & "CVDATE('" & Trim(rstResProd!FEMISION & "") & "'), "
                    SqlCad = SqlCad & "CVDATE('" & Trim(rstResProd!FENTREGA & "") & "'), "
                    SqlCad = SqlCad & "'" & Trim(rstResProd!VENDEDOR & "") & "', "
                    SqlCad = SqlCad & (CVDate(rstResProd!FENTREGA & "") - CDate(strFechaCorteValidezCompromiso)) & ", "
                    SqlCad = SqlCad & "'" & Trim(rstResProd!CodProducto & "") & "', "
                    SqlCad = SqlCad & "'" & Trim(rstResProd!NOMPRODUCTOUM & "") & "', "
                    SqlCad = SqlCad & dblCantidadComprometer
                    SqlCad = SqlCad & ")"
                    
                    abrirCnnDbBancos
                    
                    cnn_dbbancos.Execute SqlCad
                    
                    Actualiza_Log SqlCad, StrConexDbBancos
                End If
                
                DoEvents
                
                pgbProceso1.Value = pgbProceso1.Value + 1
                lblProceso1.Caption = "Compromiso Automatico 1/2..." & FormatPercent(pgbProceso1.Value / pgbProceso1.Max, 3)
                
                rstResProd.MoveNext
            Loop
        End If
        
        chkProceso(3).BackColor = RGB(255, 153, 51)
        
        Actualiza_Log "PROCESO 4 FINALIZO AL: " & Now, StrConexDbBancos
    End If
    
    
    
    
    If CBool(chkProceso(4).Value) Then
        abrirCnTemporal
        
        'Seleccionar Saldo de Produccion para su Compromiso Automatico - Pedidos con Fecha de Entrega mayor a la Fecha de Hoy (Compromisos Futuros)
        SqlCad = vbNullString
        SqlCad = SqlCad & "SELECT "
        SqlCad = SqlCad & "NROPEDIDO, "
        SqlCad = SqlCad & "FENTREGA, "
        SqlCad = SqlCad & "CODPRODUCTO, "
        SqlCad = SqlCad & "NOMPRODUCTO, "
        SqlCad = SqlCad & "SUM(SALDO) - VAL(FORMAT(AVG(COMPROMISOEAG) + AVG(COMPROMISOPLG), '#0.0000')) AS SALDOTOTAL "
        SqlCad = SqlCad & "FROM "
        SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTOPRODUCCION "
        SqlCad = SqlCad & "WHERE "
        SqlCad = SqlCad & "SALDO > 0 AND "
        SqlCad = SqlCad & "CVDATE(FENTREGA) > CVDATE('" & strFechaCorteActual & "') "
            
            If Trim(txtNroPedido.Text) <> vbNullString Then
                SqlCad = SqlCad & "AND NROPEDIDO = '" & Trim(txtNroPedido.Text) & "' "
            End If
            
            If Trim(txtCodProducto.Text) <> vbNullString Then
                SqlCad = SqlCad & "AND CODPRODUCTO = '" & Trim(txtCodProducto.Text) & "' "
            End If
            
        SqlCad = SqlCad & "GROUP BY "
        SqlCad = SqlCad & "NROPEDIDO, "
        SqlCad = SqlCad & "FENTREGA, "
        SqlCad = SqlCad & "CODPRODUCTO, "
        SqlCad = SqlCad & "NOMPRODUCTO "
        SqlCad = SqlCad & "ORDER BY "
        SqlCad = SqlCad & "NOMPRODUCTO, "
        SqlCad = SqlCad & "FENTREGA, "
        SqlCad = SqlCad & "NROPEDIDO"
        
        If rstResProd.State = 1 Then rstResProd.Close
        
        rstResProd.Open SqlCad, cnDBTemp, adOpenForwardOnly, adLockReadOnly
        
        If Not rstResProd.EOF Then
            pgbProceso1.Max = ModUtilitario.devuelveCantRegistros(rstResProd)
            pgbProceso1.Value = 0
            lblProceso1.Caption = "Compromiso Automatico 1/2..."
            
            dblCantidadComprometer = 0
            
            strNumValeProceso3 = vbNullString
            
            Do While Not rstResProd.EOF
                dblCantidadComprometer = Val(rstResProd!SALDOTOTAL & "")
                
                If dblCantidadComprometer > 0 Then
                    SqlCad = vbNullString
                    SqlCad = SqlCad & "SELECT "
                    SqlCad = SqlCad & "F2CODALM, "
                    SqlCad = SqlCad & "F5CODPRO, "
                    SqlCad = SqlCad & "COD_SOLICITUD, "
                    SqlCad = SqlCad & "VAL(FORMAT( SUM(VAL(FORMAT(VAL(F3CANPRO & ''), '#0.00')) * IIF(TIPO = 'S', -1, 1)) , '#0.00')) AS CANTIDAD "
                    SqlCad = SqlCad & "FROM "
                    SqlCad = SqlCad & "IF3VALES "
                    SqlCad = SqlCad & "WHERE "
                    SqlCad = SqlCad & "F2CODALM = '01' "
                    SqlCad = SqlCad & "AND TRIM(COD_SOLICITUD & '') = '' "
                    SqlCad = SqlCad & "AND F5CODPRO = '" & Trim(rstResProd!CodProducto & "") & "' "
                    SqlCad = SqlCad & "GROUP BY "
                    SqlCad = SqlCad & "F5CODPRO, "
                    SqlCad = SqlCad & "F2CODALM, "
                    SqlCad = SqlCad & "COD_SOLICITUD "
                    SqlCad = SqlCad & "HAVING "
                    SqlCad = SqlCad & "VAL(FORMAT( SUM(VAL(FORMAT(VAL(F3CANPRO & ''), '#0.00')) * IIF(TIPO = 'S', -1, 1)) , '#0.00')) > 0 "
                    SqlCad = SqlCad & "ORDER BY "
                    SqlCad = SqlCad & "F2CODALM, "
                    SqlCad = SqlCad & "F5CODPRO, "
                    SqlCad = SqlCad & "COD_SOLICITUD"
                
                    If rstStockLEA.State = 1 Then rstStockLEA.Close
                    
                    rstStockLEA.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
                
                    If Not rstStockLEA.EOF Then
                        If Val(rstStockLEA!Cantidad & "") > 0 Then
                            
                            Set objValeRegularizacionCEA = New ClsVale
                            
                            With objValeRegularizacionCEA
                                
                                .inicializarEntidades
                                
                                .CodigoAlmacen = Trim(rstStockLEA!f2codalm & "")
                                .NumeroVale = strNumValeProceso3
                                .TipoVale = "I"
                                
                                .Fecha = Format(Date, "dd/mm/yyyy")
                                .CodigoOrigen = "XCS"
                                .TipoCambio = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CAMBIO", "CAMBIOS", "FECHA", .Fecha, "F"))
                                
                                    If .TipoCambio = 0 Then
                                        .TipoCambio = "4.05"
                                    End If
                                
                                .CodigoMoneda = "S"
                                
                                .referencia = wnomcia
                                .observaciones = "PROCESO DE COMPROMISO AUTOMATICO MASIVO."
                                
                                .FecReg = Format(Date, "Short Date")
                                .UsuReg = wusuario
                                .FecMod = Format(Date, "Short Date")
                                .UsuMod = wusuario
                                
                                If .guardarVale Then
                                    Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                                    
                                    If strNumValeProceso3 = vbNullString Then
                                        strNumValeProceso3 = .NumeroVale
            
                                        'Borrar Detalle de Vale
                                        SqlCad = vbNullString
                                        SqlCad = "DELETE FROM IF3VALES WHERE F2CODALM = '" & .CodigoAlmacen & "' AND F4NUMVAL = '" & .NumeroVale & "'"
                
                                        cnn_dbbancos.Execute SqlCad
                                        Actualiza_Log SqlCad, cnn_dbbancos.ConnectionString
                                    End If
                                    
                                    .inicializarEntidadesDetalle
            
                                    .NumeroOrdenCompra = vbNullString
                                    .Requerimiento = vbNullString
                                    
                                    .CodigoProducto = Trim(rstStockLEA!f5codpro & "")
                                    .CodigoProductoOriginal = Trim(rstStockLEA!f5codpro & "")
                                    .Cantidad = IIf(Val(rstStockLEA!Cantidad & "") >= dblCantidadComprometer, dblCantidadComprometer, Val(rstStockLEA!Cantidad & "")) * -1
                                    
                                    
                                    dblItem = dblItem + 1
                                    
                                    .ITEM = dblItem
                                    
                                    .guardarValeDetalleOneByOne
                                    
                                    Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                                    
                                    .inicializarEntidadesDetalle
                                    
                                    dblItem = dblItem + 1
                                    
                                    .Requerimiento = Trim(rstResProd!NroPedido & "")
                                    
                                    .CodigoProducto = Trim(rstStockLEA!f5codpro & "")
                                    .CodigoProductoOriginal = Trim(rstStockLEA!f5codpro & "")
                                    .Cantidad = IIf(Val(rstStockLEA!Cantidad & "") >= dblCantidadComprometer, dblCantidadComprometer, Val(rstStockLEA!Cantidad & ""))
                                    .ITEM = dblItem
                                    
                                    .guardarValeDetalleOneByOne
                                    
                                    Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                                    
                                    dblCantidadComprometer = dblCantidadComprometer - .Cantidad
                                End If
                            End With
                            
                            Set objValeRegularizacionCEA = Nothing
                        End If
                    End If
                    
                    If dblCantidadComprometer > 0 Then
                        SqlCad = vbNullString
                        SqlCad = SqlCad & "SELECT "
                        SqlCad = SqlCad & "DET.ITEM, "
                        SqlCad = SqlCad & "DET.F3CODPRO, "
                        SqlCad = SqlCad & "DET.COD_SOLICITUD, "
                        SqlCad = SqlCad & "DET.F4NUMORD, "
                        SqlCad = SqlCad & "VAL(FORMAT( (((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) - VAL(INGRESOS.CANTIDAD & '')) , '#.0000')) AS SALDO "
                        SqlCad = SqlCad & "FROM "
                        SqlCad = SqlCad & "(((((IF3ORDEN AS DET "
                        SqlCad = SqlCad & "LEFT JOIN IF5PLA AS PROD ON PROD.F5CODPRO = DET.F3CODPRO) "
                        SqlCad = SqlCad & "LEFT JOIN EF7MEDIDAS AS MED ON MED.F7CODMED = PROD.F7CODMED) "
                        SqlCad = SqlCad & "LEFT JOIN EF7MEDIDAS AS MED2 ON MED2.F7CODMED = DET.UNIDAD) "
                        SqlCad = SqlCad & "LEFT JOIN MEDIVENTAS AS MEDALTER ON MEDALTER.F5CODPRO = DET.F3CODPRO AND MEDALTER.F7CODMED = DET.UNIDAD) "
                        SqlCad = SqlCad & "LEFT JOIN "
                        SqlCad = SqlCad & "(SELECT "
                        SqlCad = SqlCad & "DET.F4NUMORD, "
                        SqlCad = SqlCad & "DET.COD_SOLICITUD, "
                        SqlCad = SqlCad & "DET.F5CODPROORIGINAL, "
                        SqlCad = SqlCad & "VAL(FORMAT( SUM(DET.F3CANPRO * IIF(DET.TIPO = 'S', -1, 1)) , '#.0000')) AS CANTIDAD "
                        SqlCad = SqlCad & "FROM "
                        SqlCad = SqlCad & "IF3VALES AS DET "
                        SqlCad = SqlCad & "LEFT JOIN IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM "
                        SqlCad = SqlCad & "WHERE "
                        SqlCad = SqlCad & "CAB.F1CODORI IN ('XC0') AND "
                        SqlCad = SqlCad & "DET.F4NUMORD <> '' AND "
                        SqlCad = SqlCad & "DET.COD_SOLICITUD = '' AND "
                        SqlCad = SqlCad & "DET.F5CODPROORIGINAL = '" & Trim(rstResProd!CodProducto & "") & "' "
                        SqlCad = SqlCad & "GROUP BY "
                        SqlCad = SqlCad & "DET.F4NUMORD, "
                        SqlCad = SqlCad & "DET.COD_SOLICITUD, "
                        SqlCad = SqlCad & "DET.F5CODPROORIGINAL) AS INGRESOS "
                        SqlCad = SqlCad & "ON INGRESOS.F4NUMORD = DET.F4NUMORD AND INGRESOS.COD_SOLICITUD = DET.COD_SOLICITUD AND INGRESOS.F5CODPROORIGINAL = DET.F3CODPRO) "
                        SqlCad = SqlCad & "LEFT JOIN TB_CABSOLICITUD AS CABPED ON CABPED.COD_SOLICITUD = DET.COD_SOLICITUD "
                        SqlCad = SqlCad & "WHERE "
                        SqlCad = SqlCad & "DET.F4LOCAL = 'OC' AND "
                        SqlCad = SqlCad & "DET.COD_SOLICITUD = '' AND "
                        SqlCad = SqlCad & "DET.F3CODPRO = '" & Trim(rstResProd!CodProducto & "") & "' AND "
                        SqlCad = SqlCad & "VAL(FORMAT( (((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) - VAL(INGRESOS.CANTIDAD & '')) , '#.0000')) > 0 "
                        SqlCad = SqlCad & "ORDER BY "
                        SqlCad = SqlCad & "PROD.F5NOMPRO, "
                        SqlCad = SqlCad & "DET.COD_SOLICITUD, "
                        SqlCad = SqlCad & "DET.F4NUMORD"
        
                        If rstStockLPL.State = 1 Then rstStockLPL.Close
        
                        rstStockLPL.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
        
                        If Not rstStockLPL.EOF Then
                            
                            Do While Not rstStockLPL.EOF
                                SqlCad = vbNullString
                                SqlCad = SqlCad & "SELECT "
                                SqlCad = SqlCad & "* "
                                SqlCad = SqlCad & "FROM "
                                SqlCad = SqlCad & "IF3ORDEN "
                                SqlCad = SqlCad & "WHERE "
                                SqlCad = SqlCad & "F4LOCAL = 'OC' AND "
                                SqlCad = SqlCad & "F4NUMORD = '" & Trim(rstStockLPL!F4NUMORD & "") & "' AND "
                                SqlCad = SqlCad & "COD_SOLICITUD = '" & Trim(rstStockLPL!COD_SOLICITUD & "") & "' AND "
                                SqlCad = SqlCad & "F3CODPRO = '" & Trim(rstStockLPL!F3CODPRO & "") & "' AND "
                                SqlCad = SqlCad & "ITEM = '" & Trim(rstStockLPL!ITEM & "") & "'"
                                
                                If rstOCDetalle.State = 1 Then rstOCDetalle.Close
        
                                rstOCDetalle.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
        
                                If Not rstOCDetalle.EOF Then
                                    rstOCDetalle.MoveFirst
        
                                    Set objOrdenRegularizacionCPL = New ClsOrden
        
                                    With objOrdenRegularizacionCPL
                                        'Insertamos el Stock Re-Distribuido
                                        .inicializarEntidadesDetalle
        
                                        .TipoOrden = "OC"
                                        .NumeroOrden = Trim(rstOCDetalle!F4NUMORD & "")
                                        
                                        .ITEM = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "VAL(ITEM & '')", "IF3ORDEN", "F4LOCAL", "OC", "T", "AND F4NUMORD = '" & .NumeroOrden & "' ORDER BY VAL(ITEM & '') DESC")) + 1
                                        
                                        .Requerimiento = Trim(rstResProd!NroPedido & "")
                                        
                                        .CodigoProducto = Trim(rstOCDetalle!F3CODPRO & "")
                                        .CodigoFabricante = Trim(rstOCDetalle!F3CODFAB & "")
                                        .NombreProducto = Trim(rstOCDetalle!F5NOMPRO & "")
                                        .NombreProductoInterno = Trim(rstOCDetalle!F5NOMPRO_ING & "")
                                        .CodigoUM = Trim(rstOCDetalle!UNIDAD & "")
                                        .CodigoColor = Trim(rstOCDetalle!CODCOLOR & "")
                                        .ObservacionPorItem = "REDISTRIBUCION DE SCTOCK AUTOMATICO MASIVO A FAVOR DEL PEDIDO " & .Requerimiento & "."
        
                                        dblFactorUM = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5FACTOR", "MEDIVENTAS", "F5CODPRO", .CodigoProducto, "T", "AND F7CODMED = '" & .CodigoUM & "'"))
        
                                        If dblFactorUM = 0 Then dblFactorUM = 1
        
                                        dblCantidadDisponibleSegunOC = Val(Format((Val(rstStockLPL!SALDO & "") / (1 + (Val(rstOCDetalle!F3PORCDEMASIA & "") / 100))) / dblFactorUM, "#.0000"))
                                        dblCantidadDestinoSegunOC = Val(Format((IIf(Val(rstStockLPL!SALDO & "") >= dblCantidadComprometer, dblCantidadComprometer, Val(rstStockLPL!SALDO & "")) / (1 + (Val(rstOCDetalle!F3PORCDEMASIA & "") / 100))) / dblFactorUM, "#.0000"))
                                        
                                        .PorcentajeImpuesto = wwigv / 100
                                        .SignoImpuesto = 1
        
                                        .Cantidad = dblCantidadDestinoSegunOC 'Val(Format((Val(rstResProd!CANTIDADDESTINO & "") / (1 + (Val(rstOCDetalle!F3PORCDEMASIA & "") / 100))) / dblFactorUM, "#.0000"))
        
                                        If .Requerimiento <> vbNullString Then
                                            .CantidadMaxima = dblCantidadDestinoSegunOC '.Cantidad
                                        End If
        
                                        .PorcentajeDemasia = Val(rstOCDetalle!F3PORCDEMASIA & "") / 100
                                        .PrecioSinImpuesto = Val(rstOCDetalle!F3PRECOS & "")
                                        .PrecioConImpuesto = 0 'Val(rstOCDetalle!F3PREUNI & "")
                                        .PorcentajeDscto = Val(rstOCDetalle!F3PORDCT & "") / 100
                                        .TotalDscto = 0 'Val(rstOCDetalle!F3TOTDCT & "")
        
                                        .Afecto = IIf(Trim(rstOCDetalle!F5AFECTO & "") = "*", True, False)
        
                                        .calculosPorItem
        
                                        .PorcentajeDemasia = Val(rstOCDetalle!F3PORCDEMASIA & "")
                                        .PorcentajeDscto = Val(rstOCDetalle!F3PORDCT & "")
        
                                        .CodigoGasto = Trim(rstOCDetalle!F3CUENTA & "")
                                        .CuentaContable = Trim(rstOCDetalle!F3GASTO & "")
        
                                        Select Case dblCantidadDisponibleSegunOC
                                            Case Is = Val(Format(Val(rstOCDetalle!F3CANPRO & ""), "#.0000"))
                                                If dblCantidadDestinoSegunOC < dblCantidadDisponibleSegunOC Then
        
                                                    .guardarOrdenDetalleOneByOne
        
                                                    Actualiza_Log .SQLSelectAlter, StrConexDbBancos
        
                                                    'Actualizamos la Cantidad del Item de Origen
                                                    .inicializarEntidadesDetalle
        
                                                    .PorcentajeImpuesto = wwigv / 100
                                                    .SignoImpuesto = 1
        
                                                    .Cantidad = Val(Format(Val(rstOCDetalle!F3CANPRO & "") - dblCantidadDestinoSegunOC, "#.0000")) 'Val(Format(((Val(rstResProd!Cantidad & "") - Val(rstResProd!CANTIDADDESTINO & "")) / (1 + (Val(rstOCDetalle!F3PORCDEMASIA & "") / 100))) / dblFactorUM, "#.0000"))
        
                                                    If Trim(rstResProd!NroPedido & "") <> vbNullString Then
                                                        .CantidadMaxima = .Cantidad
                                                    End If
        
                                                    .PorcentajeDemasia = Val(rstOCDetalle!F3PORCDEMASIA & "") / 100
                                                    .PrecioSinImpuesto = Val(rstOCDetalle!F3PRECOS & "")
                                                    .PrecioConImpuesto = 0 'Val(rstOCDetalle!F3PREUNI & "")
                                                    .PorcentajeDscto = Val(rstOCDetalle!F3PORDCT & "") / 100
                                                    .TotalDscto = 0 'Val(rstOCDetalle!F3TOTDCT & "")
        
                                                    .Afecto = IIf(Trim(rstOCDetalle!F5AFECTO & "") = "*", True, False)
        
                                                    .calculosPorItem
        
                                                    .PorcentajeDemasia = Val(rstOCDetalle!F3PORCDEMASIA & "")
                                                    .PorcentajeDscto = Val(rstOCDetalle!F3PORDCT & "")
                                                End If
        
                                                SqlCad = vbNullString
                                                SqlCad = SqlCad & "UPDATE "
                                                SqlCad = SqlCad & "IF3ORDEN "
                                                SqlCad = SqlCad & "SET "
        
                                                    If dblCantidadDestinoSegunOC < dblCantidadDisponibleSegunOC Then
                                                        SqlCad = SqlCad & "F5NOMPRO = TRIM(F5NOMPRO & ''), "
                                                        SqlCad = SqlCad & "F5NOMPRO_ING = TRIM(F5NOMPRO_ING & ''), "
                                                        SqlCad = SqlCad & "F3CANPRO = " & .Cantidad & ", "
                                                        SqlCad = SqlCad & "F3CANPRO2 = " & .CantidadMaxima & ", "
                                                        SqlCad = SqlCad & "F5VALVTA = " & .BasePorItem & ", "
                                                        SqlCad = SqlCad & "F3IGV = " & .ImpuestoPorItem & ", "
                                                        SqlCad = SqlCad & "F3TOTAL = " & .TotalPorItem & ", "
                                                        SqlCad = SqlCad & "F3PORDCT = " & .PorcentajeDscto & ", "
                                                        SqlCad = SqlCad & "F3TOTDCT = " & .TotalDscto & " "
                                                    Else
                                                        SqlCad = SqlCad & "COD_SOLICITUD = '" & Trim(rstResProd!NroPedido & "") & "' "
                                                    End If
        
                                                SqlCad = SqlCad & "WHERE "
                                                SqlCad = SqlCad & "F4LOCAL = 'OC' AND "
                                                SqlCad = SqlCad & "TRIM(F4NUMORD & '') = '" & Trim(rstOCDetalle!F4NUMORD & "") & "' AND "
                                                SqlCad = SqlCad & "TRIM(COD_SOLICITUD & '') = '" & Trim(rstOCDetalle!COD_SOLICITUD & "") & "' AND "
                                                SqlCad = SqlCad & "F3CODPRO = '" & Trim(rstOCDetalle!F3CODPRO & "") & "'"
        
                                                cnn_dbbancos.Execute SqlCad
        
                                                Actualiza_Log SqlCad, StrConexDbBancos
                                            Case Is < Val(Format(Val(rstOCDetalle!F3CANPRO & ""), "#.0000"))
        
                                                .guardarOrdenDetalleOneByOne
        
                                                Actualiza_Log .SQLSelectAlter, StrConexDbBancos
        
                                                'Actualizamos la Cantidad del Item de Origen
                                                .inicializarEntidadesDetalle
        
                                                .PorcentajeImpuesto = wwigv / 100
                                                .SignoImpuesto = 1
        
                                                .Cantidad = Val(Format(Val(rstOCDetalle!F3CANPRO & "") - dblCantidadDestinoSegunOC, "#.0000")) 'Val(Format(((Val(rstResProd!Cantidad & "") - Val(rstResProd!CANTIDADDESTINO & "")) / (1 + (Val(rstOCDetalle!F3PORCDEMASIA & "") / 100))) / dblFactorUM, "#.0000"))
                                                .CantidadMaxima = .Cantidad
        
                                                .PorcentajeDemasia = Val(rstOCDetalle!F3PORCDEMASIA & "") / 100
                                                .PrecioSinImpuesto = Val(rstOCDetalle!F3PRECOS & "")
                                                .PrecioConImpuesto = 0 'Val(rstOCDetalle!F3PREUNI & "")
                                                .PorcentajeDscto = Val(rstOCDetalle!F3PORDCT & "") / 100
                                                .TotalDscto = 0 'Val(rstOCDetalle!F3TOTDCT & "")
        
                                                .Afecto = IIf(Trim(rstOCDetalle!F5AFECTO & "") = "*", True, False)
        
                                                .calculosPorItem
        
                                                .PorcentajeDemasia = Val(rstOCDetalle!F3PORCDEMASIA & "")
                                                .PorcentajeDscto = Val(rstOCDetalle!F3PORDCT & "")
        
                                                SqlCad = vbNullString
                                                SqlCad = SqlCad & "UPDATE "
                                                SqlCad = SqlCad & "IF3ORDEN "
                                                SqlCad = SqlCad & "SET "
        
                                                        SqlCad = SqlCad & "F3CANPRO = " & .Cantidad & ", "
                                                        SqlCad = SqlCad & "F3CANPRO2 = " & .CantidadMaxima & ", "
                                                        SqlCad = SqlCad & "F5VALVTA = " & .BasePorItem & ", "
                                                        SqlCad = SqlCad & "F3IGV = " & .ImpuestoPorItem & ", "
                                                        SqlCad = SqlCad & "F3TOTAL = " & .TotalPorItem & ", "
                                                        SqlCad = SqlCad & "F3PORDCT = " & .PorcentajeDscto & ", "
                                                        SqlCad = SqlCad & "F3TOTDCT = " & .TotalDscto & " "
        
                                                SqlCad = SqlCad & "WHERE "
                                                SqlCad = SqlCad & "F4LOCAL = 'OC' AND "
                                                SqlCad = SqlCad & "F4NUMORD = '" & Trim(rstOCDetalle!F4NUMORD & "") & "' AND "
                                                SqlCad = SqlCad & "COD_SOLICITUD = '" & Trim(rstOCDetalle!COD_SOLICITUD & "") & "' AND "
                                                SqlCad = SqlCad & "F3CODPRO = '" & Trim(rstOCDetalle!F3CODPRO & "") & "'"
        
                                                cnn_dbbancos.Execute SqlCad
        
                                                Actualiza_Log SqlCad, StrConexDbBancos
                                        End Select
                                    End With
        
                                    Set objOrdenRegularizacionCPL = Nothing
                                    
                                    dblCantidadComprometer = dblCantidadComprometer - IIf(Val(rstStockLPL!SALDO & "") >= dblCantidadComprometer, dblCantidadComprometer, Val(rstStockLPL!SALDO & ""))
                                End If
                                
                                If dblCantidadComprometer = 0 Then Exit Do
                                
                                rstStockLPL.MoveNext
                            Loop
                        End If
                    End If
                End If
                
                DoEvents
                
                pgbProceso1.Value = pgbProceso1.Value + 1
                lblProceso1.Caption = "Compromiso Automatico 1/2..." & FormatPercent(pgbProceso1.Value / pgbProceso1.Max, 3)
                
                rstResProd.MoveNext
            Loop
        End If
        
        chkProceso(4).BackColor = RGB(255, 153, 51)
        
        Actualiza_Log "PROCESO 5 FINALIZO AL: " & Now, StrConexDbBancos
    End If
    
    
    If rstStockCEA.State = 1 Then rstStockCEA.Close
    If rstStockCPL.State = 1 Then rstStockCPL.Close
    If rstProduccion.State = 1 Then rstProduccion.Close
    If rstResProd.State = 1 Then rstResProd.Close
    If rstStockLEA.State = 1 Then rstStockLEA.Close
    If rstStockLPL.State = 1 Then rstStockLPL.Close
    If rstDistribuir.State = 1 Then rstDistribuir.Close
    If rstOCDetalle.State = 1 Then rstOCDetalle.Close
    
    Set rstStockCEA = Nothing
    Set rstStockCPL = Nothing
    Set rstProduccion = Nothing
    
    Set rstResProd = Nothing
    Set rstStockLEA = Nothing
    Set rstStockLPL = Nothing
    
    Set rstDistribuir = Nothing
    Set rstOCDetalle = Nothing
    
    'dblItem As Double
    dblStock = 0
    dblCantidadProdObservado = 0
    
    dblLiberarCompromiso = 0
    
    dblItem = 0
    dblCantidadDestino = 0
    dblCantidadDestinoSegunOC = 0
    dblCantidadDisponibleSegunOC = 0
    
    dblFactorUM = 0
    
    dblItemLibreEnOC = 0
    
    bolCompromisoEjecutado = False
    dblCantidadComprometer = 0
    
    strFechaCorteValidezCompromiso = vbNullString
    intCantidadMesesDeValidezCompromiso = 0
    
    strNumValeProceso1 = vbNullString
    strNumValeProceso2 = vbNullString
    strNumValeProceso3 = vbNullString
    
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "UPDATE "
    SqlCad = SqlCad & "SF4COMPROMISOAUTOMATICO "
    SqlCad = SqlCad & "SET "
    SqlCad = SqlCad & "HORAFIN = CVDATE('" & Format(Now, "hh:mm:ss AM/PM") & "'), "
    SqlCad = SqlCad & "USUMOD = '" & wusuario & "', "
    SqlCad = SqlCad & "FECMOD = CVDATE('" & Format(Date, "Short Date") & "') "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "CVDATE(FECHAEJECUCION) = CVDATE('" & strFechaCorteActual & "')"
    
    abrirCnnDbBancos
    
    cnn_dbbancos.Execute SqlCad
    
    Actualiza_Log SqlCad, StrConexDbBancos
    
    Actualiza_Log "FIN DE PROCESO: " & Now, StrConexDbBancos
    
    MsgBox "Proceso Finalizado.", vbInformation + vbOKOnly, App.ProductName
    
    Exit Sub
errAnalisisYCorreccionStockComprometidoActual:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Actualiza_Log "Proceso de Regularizacion de Stock Trunco por errores. / No. Error: " & Err.Number & " / " & "Descripción: " & Err.Description, StrConexDbBancos
    
    Err.Clear
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
                '.InsumoOP = False
                
                .Ayuda = True
                .InsumoOP = False
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
                    lblProducto.Caption = objAyudaBien.Descripcion
                End If
            End With
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub txtCodProducto_LostFocus()
    If Trim(txtCodProducto.Text) <> vbNullString Then
        txtCodProducto.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5CODPRO", "IF5PLA", "F5CODPRO", Trim(txtCodProducto.Text), "T")
        lblProducto.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5NOMPRO", "IF5PLA", "F5CODPRO", Trim(txtCodProducto.Text), "T")
        
        If Trim(txtCodProducto.Text) = vbNullString Then
            lblProducto.Caption = "Todos los Productos (*)"
        End If
    Else
        lblProducto.Caption = "Todos los Productos (*)"
    End If
End Sub

Private Sub txtNroPedido_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'abrirCnDBMilano
            
            lblNroPedido.Caption = ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "(PER.NOMBRE + ' ( FEC. PEDIDO: ' + CONVERT(CHAR(10), PED.FECHAEMISION, 103) + ' / FEC. ENTREGA: ' + CONVERT(CHAR(10), PED.FECHAENTREGA, 103) + ')') AS RESUMEN", "PEDIDO AS PED LEFT JOIN PERSONA AS PER ON PER.IDPERSONA = PED.IDPERSONA", "PED.IDPEDIDO", Trim(txtNroPedido.Text), "T")
            
            If lblNroPedido.Caption = vbNullString Then
                MsgBox "No. de Pedido no encontrado o inválido.", vbInformation + vbOKOnly, App.ProductName
                
                txtNroPedido.Text = vbNullString
                
                txtNroPedido.SetFocus
            Else
                ModUtilitario.pulsarTecla vbKeyTab
            End If
    End Select
End Sub
