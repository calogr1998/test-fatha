Attribute VB_Name = "ModMilano"
Option Explicit

Public strCadenaConexioBdStudioModa  As String

Public cnBdStudioModa   As New ADODB.Connection

Public Sub abrirCnDBMilano()
    On Error GoTo errAbrirCnDBMilano
    
    'Dim strCadenaConexion As String
    
    strCadenaConexioBdStudioModa = sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "CadenaConexion", "l")
    
    If cnBdStudioModa.State = 1 Then cnBdStudioModa.Close
    
    cnBdStudioModa.Open strCadenaConexioBdStudioModa
    
    Exit Sub
errAbrirCnDBMilano:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName & " - AbrirCnMilano"
    
    Err.Clear
End Sub

Public Sub importarRequerimientosServidorExterno(ByVal frameProgreso As Object, _
                                                    ByVal barraProgreso As Object, _
                                                    Optional ByVal strNoPedido As String)
    
    On Error GoTo errImportarRequerimientosServidorExterno
    
    Dim rstRequerimientoCab As ADODB.Recordset
    Dim rstRequerimientoDet As ADODB.Recordset
    Dim strUltimaLecturaDeRequerimientos As String
    Dim dblItem As Double
    
    Set rstRequerimientoCab = Nothing
    Set rstRequerimientoDet = Nothing
    
    Set rstRequerimientoCab = New ADODB.Recordset
    Set rstRequerimientoDet = New ADODB.Recordset
    
    strUltimaLecturaDeRequerimientos = sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UltimaLecturaDeRequerimientos", "l")
    
    'abrirCnDBMilano
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "CAB.IDPEDIDO, "
    SqlCad = SqlCad & "TPER.NOMBRE AS TIPOPERSONA, "
    SqlCad = SqlCad & "PER.NOMBRE, "
    SqlCad = SqlCad & "PED.FECHAEMISION, "
    SqlCad = SqlCad & "PED.FECHAENTREGA, "
    SqlCad = SqlCad & "PED.FECHAINGRESO, "
    SqlCad = SqlCad & "PED.FECHAACTUALIZACION, "
    SqlCad = SqlCad & "PED.IDUSUARIOACTUALIZACION, "
    SqlCad = SqlCad & "PED.ANULADO "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "REQUERIMIENTO AS CAB "
    SqlCad = SqlCad & "LEFT JOIN PEDIDO AS PED ON PED.IDPEDIDO = CAB.IDPEDIDO "
    SqlCad = SqlCad & "LEFT JOIN PERSONA AS PER ON PER.IDPERSONA = PED.IDPERSONA "
    SqlCad = SqlCad & "LEFT JOIN TIPOPERSONA AS TPER ON TPER.IDTIPOPERSONA = PER.IDTIPOPERSONA "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "CAB.OP = 1 AND "
    SqlCad = SqlCad & "PED.REQUERIMIENTO = 1 AND "
    SqlCad = SqlCad & "LTRIM(CAB.IDPEDIDO) <> '' AND "
        
        If strNoPedido = vbNullString Then
            SqlCad = SqlCad & "dbo.Fecha(CAB.FECHAACTUALIZACION) >= '" & strUltimaLecturaDeRequerimientos & "' "
        Else
            SqlCad = SqlCad & "PED.IDPEDIDO = '" & strNoPedido & "' "
        End If
        
    SqlCad = SqlCad & "GROUP BY "
    SqlCad = SqlCad & "CAB.IDPEDIDO, "
    SqlCad = SqlCad & "TPER.NOMBRE, "
    SqlCad = SqlCad & "PER.NOMBRE, "
    SqlCad = SqlCad & "PED.FECHAEMISION, "
    SqlCad = SqlCad & "PED.FECHAENTREGA, "
    SqlCad = SqlCad & "PED.FECHAINGRESO, "
    SqlCad = SqlCad & "PED.FECHAACTUALIZACION, "
    SqlCad = SqlCad & "PED.IDUSUARIOACTUALIZACION, "
    SqlCad = SqlCad & "PED.ANULADO "
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "CAB.IDPEDIDO"
    
    If rstRequerimientoCab.State = 1 Then rstRequerimientoCab.Close
    
    rstRequerimientoCab.Open SqlCad, cnBdStudioModa, adOpenForwardOnly, adLockReadOnly
    
    If Not rstRequerimientoCab.EOF Then
        'rstRequerimientoCab.MoveFirst
        
        frameProgreso.Visible = True
        barraProgreso.Max = ModUtilitario.devuelveCantRegistros(rstRequerimientoCab)
        barraProgreso.Value = 0
        frameProgreso.Caption = "Actualizando..."
        
        Do While Not rstRequerimientoCab.EOF
            With objAyudaSolicitud
                .inicializarEntidades
                
                .TipoDocumento = "OC"
                .Codigo = Trim(rstRequerimientoCab!IDPEDIDO & "")
                .Fecha = Format(Trim(rstRequerimientoCab!FechaEmision & ""), "dd/mm/yyyy")
                .Estado1 = "2"
                .VBJefe = True
                .VBFecha = Format(Trim(rstRequerimientoCab!FechaEmision & ""), "dd/mm/yyyy")
                .VBUsuario = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODUSER", "EF2USERS", "F2CODUSEREXTERNO", Trim(rstRequerimientoCab!IDUSUARIOACTUALIZACION & ""), "T")
                .Estado2 = "P"
                .Prioridad = "1"
                .observaciones = "EN ATENCION A " & Trim(rstRequerimientoCab!TipoPersona & "") & ": " & Replace(Trim(rstRequerimientoCab!nombre & ""), "'", "' & Chr(39) & '", 1)
                .NombreReferencial = Replace(Trim(rstRequerimientoCab!nombre & ""), "'", "' & Chr(39) & '", 1)
                .Anulado = CBool(rstRequerimientoCab!Anulado)
                
                If .Anulado Then
                    .Estado1 = "5"
                End If
                '.Cerrado = CBool(rstRequerimientoCab!Cerrado)
                
                .CodigoSolicitante = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODUSER", "EF2USERS", "F2CODUSEREXTERNO", Trim(rstRequerimientoCab!IDUSUARIOACTUALIZACION & ""), "T")
                .CodigoAprobadoPor1 = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODUSER", "EF2USERS", "F2CODUSEREXTERNO", Trim(rstRequerimientoCab!IDUSUARIOACTUALIZACION & ""), "T")
                .Empresa = wF1Dir
                .FechaEntrega = Format(Trim(rstRequerimientoCab!FechaEntrega & ""), "dd/mm/yyyy")
                .Usuario = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODUSER", "EF2USERS", "F2CODUSEREXTERNO", Trim(rstRequerimientoCab!IDUSUARIOACTUALIZACION & ""), "T")
                '.TOTAL = Val(rstRequerimientoCab!TotalImporte & "")
                .FechaReg = Trim(rstRequerimientoCab!FechaIngreso & "")
                .FechaMod = Trim(rstRequerimientoCab!FECHAACTUALIZACION & "")
                
                If .guardarSolicitud(True) Then
                    Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                    
                    
                    
                    'Borrar Detalle de Requerimiento
                    SqlCad = vbNullString
                    SqlCad = "DELETE FROM TB_DETSOLICITUD WHERE COD_SOLICITUD = '" & .Codigo & "'"
                    
                    cnn_dbbancos.Execute SqlCad
                    Actualiza_Log SqlCad, cnn_dbbancos.ConnectionString
                    
                    SqlCad = vbNullString
'                    SqlCad = SqlCad & "SELECT "
'                    SqlCad = SqlCad & "DET.IDINSUMO, "
'                    SqlCad = SqlCad & "INS.NOMBRE, "
'                    SqlCad = SqlCad & "(SUM(DET.CANTIDAD * (CASE WHEN LEN(SFAM.BaseTalla) > 0 AND SFAM.BASETALLA <> 'CINTURON' THEN 1 ELSE CAB.TOTALCANTIDAD END))) AS CANTTOTAL "
'                    SqlCad = SqlCad & "FROM "
'                    SqlCad = SqlCad & "((REQUERIMIENTODETALLE AS DET "
'                    SqlCad = SqlCad & "LEFT JOIN REQUERIMIENTO AS CAB ON CAB.IDREQUERIMIENTO = DET.IDREQUERIMIENTO) "
'                    SqlCad = SqlCad & "LEFT JOIN INSUMO AS INS ON INS.IDINSUMO = DET.IDINSUMO) "
'                    SqlCad = SqlCad & "LEFT JOIN SUBFAMILIA AS SFAM ON SFAM.IDSUBFAMILIA = INS.IDSUBFAMILIA "
'                    SqlCad = SqlCad & "WHERE "
'                    SqlCad = SqlCad & "CAB.IDPEDIDO = '" & .Codigo & "' AND "
'                    SqlCad = SqlCad & "CAB.ANULADO = 0 "
'                    SqlCad = SqlCad & "GROUP BY "
'                    SqlCad = SqlCad & "DET.IDINSUMO, INS.NOMBRE"
                    SqlCad = SqlCad & "SELECT "
                    SqlCad = SqlCad & "DET.IDINSUMO, "
                    SqlCad = SqlCad & "INS.NOMBRE, "
                    SqlCad = SqlCad & "(SUM(DET.CANTIDAD * (CASE WHEN LEN(SFAM.BASETALLA) > 0 AND SFAM.BASETALLA <> 'CINTURON' THEN 1 ELSE CAB.TOTALCANTIDAD END))) AS CANTTOTAL "
                    SqlCad = SqlCad & "FROM "
                    SqlCad = SqlCad & "((REQUERIMIENTODETALLE AS DET "
                    SqlCad = SqlCad & "LEFT JOIN REQUERIMIENTO AS CAB "
                    SqlCad = SqlCad & "ON CAB.IDREQUERIMIENTO = DET.IDREQUERIMIENTO) "
                    SqlCad = SqlCad & "LEFT JOIN INSUMO AS INS "
                    SqlCad = SqlCad & "ON INS.IDINSUMO = DET.IDINSUMO) "
                    SqlCad = SqlCad & "LEFT JOIN SUBFAMILIA AS SFAM ON SFAM.IDSUBFAMILIA = INS.IDSUBFAMILIA "
                    SqlCad = SqlCad & "WHERE "
                    SqlCad = SqlCad & "CAB.IDPEDIDO = '" & .Codigo & "' AND "
                    SqlCad = SqlCad & "CAB.ANULADO = 0 "
                    SqlCad = SqlCad & "GROUP BY "
                    SqlCad = SqlCad & "DET.IDINSUMO, INS.NOMBRE "
                    SqlCad = SqlCad & "ORDER BY "
                    SqlCad = SqlCad & "INS.NOMBRE"
                    
                    If rstRequerimientoDet.State = 1 Then rstRequerimientoDet.Close
                    
                    rstRequerimientoDet.Open SqlCad, cnBdStudioModa, adOpenForwardOnly, adLockReadOnly '3, 1
                    
                    If Not rstRequerimientoDet.EOF Then
                        'rstRequerimientoDet.MoveFirst
                        
                        dblItem = 0
                        
                        '/-- Lectura de los datos del detalle hasta que no exista ningun dato
                        Do While Not rstRequerimientoDet.EOF
                            .inicializarEntidadesDetalle
                            
                            dblItem = dblItem + 1
                            
                            .CodProducto = Trim(rstRequerimientoDet!IDINSUMO & "")
                            .Descripcion = Trim(rstRequerimientoDet!nombre & "")
                            .Cantidad = Val(rstRequerimientoDet!CANTTOTAL & "")
                            .Cantidad2 = Val(rstRequerimientoDet!CANTTOTAL & "")
                            .ITEM = dblItem
                            .CodUniMed = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F7CODMED", "IF5PLA", "F5CODPRO", .CodProducto, "T")  'Trim(rstRequerimientoDet!IDBASETALLA & "")
                            
                            .guardarSolicitudDetalleOneByOne
                            
                            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
    
                            rstRequerimientoDet.MoveNext
                        Loop
                    End If
                Else
                    Actualiza_Log "Importación de Requerimiento No. " & .Codigo & " fallido.", StrConexDbBancos
                End If
            End With
            
            
                        
            DoEvents
            
            barraProgreso.Value = barraProgreso.Value + 1
            frameProgreso.Caption = " Actualización al... " & FormatPercent(barraProgreso.Value / barraProgreso.Max, 3)
            
            rstRequerimientoCab.MoveNext
        Loop
            MsgBox "Actualización Finalizada", vbInformation + vbOKOnly, wnomcia
            
            ModUtilitario.sWrtIni App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UltimaLecturaDeRequerimientos", Format(Date, "Short Date") '"04/07/2014"
    Else
        MsgBox "No se ubicaron Requerimientos nuevos y/o modificados.", vbInformation + vbOKOnly, wnomcia
    End If
    
    'cnBdStudioModa.Close
    SqlCad = vbNullString
    strUltimaLecturaDeRequerimientos = vbNullString
    dblItem = 0
    
    frameProgreso.Visible = False
    
    If rstRequerimientoCab.State = 1 Then rstRequerimientoCab.Close
    If rstRequerimientoDet.State = 1 Then rstRequerimientoDet.Close
    
    Set rstRequerimientoCab = Nothing
    Set rstRequerimientoDet = Nothing
    
    Exit Sub
errImportarRequerimientosServidorExterno:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName & " - ImportarRequerimientosServidorExterno"
    
    Err.Clear
End Sub

Public Sub importarValesServidorExterno(ByVal strTipoVale As String, _
                                            ByVal frameProgreso As Object, _
                                            ByVal barraProgreso As Object)
    
    On Error GoTo errImportarValesServidorExterno
    
    Dim rstValeCab As ADODB.Recordset
    Dim rstValeDet As ADODB.Recordset
    
    Dim strFechaCorteInicialDeValesParaCP As String
    Dim strUltimaLecturaDeValesIngreso As String
    Dim strUltimaLecturaDeValesSalida As String
    Dim dblItem As Double
    
    Set rstValeCab = Nothing
    Set rstValeDet = Nothing
    
    Set rstValeCab = New ADODB.Recordset
    Set rstValeDet = New ADODB.Recordset
    
    Rem LECTURA DE INGRESOS Y SALIDAS (Vales)
    
    strFechaCorteInicialDeValesParaCP = sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "FechaCorteInicialDeValesParaCP", "l")
    
    If MsgBox("¿Desea importar los Vales de " & IIf(strTipoVale = "I", "Ingreso", "Salida") & " a partir de la Fecha: " & strFechaCorteInicialDeValesParaCP & "?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
        Exit Sub
    End If
    
    Select Case strTipoVale
        Case "I"
            strUltimaLecturaDeValesIngreso = sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UltimaLecturaDeValesIngreso", "l")
        Case "S"
            strUltimaLecturaDeValesSalida = sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UltimaLecturaDeValesSalida", "l")
    End Select
    
    'abrirCnDBMilano
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "CAB.IDALMACEN, "
    
        Select Case strTipoVale
            Case "I"
                SqlCad = SqlCad & "CAB.IDINGRESO AS ID, "
            Case "S"
                SqlCad = SqlCad & "CAB.IDSALIDA AS ID, "
        End Select
            
    SqlCad = SqlCad & "'" & strTipoVale & "' AS TIPO, "
    SqlCad = SqlCad & "CAB.IDPERSONA, "
    SqlCad = SqlCad & "CAB.[Num-Doc] AS NRODOCUMENTO, "
    SqlCad = SqlCad & "CAB.FECHA, "
        
        Select Case strTipoVale
            Case "I"
                SqlCad = SqlCad & "CAB.IDTIPOINGRESO AS TIPOMOVIMIENTO, "
            Case "S"
                SqlCad = SqlCad & "CAB.IDTIPOSALIDA AS TIPOMOVIMIENTO, "
        End Select
    
    SqlCad = SqlCad & "CAB.PORCIGV, "
    SqlCad = SqlCad & "CAB.TCAMBIO, "
    SqlCad = SqlCad & "MON.NOMBRE AS MONEDA, "
    SqlCad = SqlCad & "CAB.IDORDENPRODUCCION, "
    SqlCad = SqlCad & "CAB.IDDOCUMENTO, "
    SqlCad = SqlCad & "CAB.FECHAINGRESO, "
    SqlCad = SqlCad & "CAB.IDUSUARIOINGRESO, "
    SqlCad = SqlCad & "CAB.FECHAACTUALIZACION, "
    SqlCad = SqlCad & "CAB.IDUSUARIOACTUALIZACION, "
    SqlCad = SqlCad & "CAB.OBSERVACIONDOCUMENTO, "
        
        Select Case strTipoVale
            Case "I"
                SqlCad = SqlCad & "CAB.IDORDENCOMPRA "
            Case "S"
                SqlCad = SqlCad & "NULL AS IDORDENCOMPRA "
        End Select
        
    SqlCad = SqlCad & "FROM "
        
        Select Case strTipoVale
            Case "I"
                SqlCad = SqlCad & "INGRESO AS CAB "
            Case "S"
                SqlCad = SqlCad & "SALIDA AS CAB "
        End Select
        
    SqlCad = SqlCad & "LEFT JOIN MONEDA AS MON ON MON.IDMONEDA = CAB.IDMONEDA "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "CAB.ANULADO = 0 AND "
    SqlCad = SqlCad & "CAB.FECHA >= '" & strFechaCorteInicialDeValesParaCP & "' AND "
    
        Select Case strTipoVale
            Case "I"
                SqlCad = SqlCad & "CAB.FECHAACTUALIZACION >= '" & strUltimaLecturaDeValesIngreso & "' "
            Case "S"
                SqlCad = SqlCad & "CAB.FECHAACTUALIZACION >= '" & strUltimaLecturaDeValesSalida & "' "
        End Select
        
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "CAB.FECHA, "
        
        Select Case strTipoVale
            Case "I"
                SqlCad = SqlCad & "CAB.IDINGRESO"
            Case "S"
                SqlCad = SqlCad & "CAB.IDSALIDA"
        End Select
    
    If rstValeCab.State = 1 Then rstValeCab.Close
    
    rstValeCab.Open SqlCad, cnBdStudioModa, adOpenForwardOnly, adLockReadOnly
    
    If Not rstValeCab.EOF Then
        frameProgreso.Visible = True
        barraProgreso.Max = ModUtilitario.devuelveCantRegistros(rstValeCab)
        barraProgreso.Value = 0
        frameProgreso.Caption = "Actualizando..."
        
        Do While Not rstValeCab.EOF
            With objAyudaVale
                .inicializarEntidades
                
                If ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "NUMENSAM", "IF4VALES", "NUMENSAM", Trim(rstValeCab!ID & ""), "T", "AND F4TIPOVALE = '" & Trim(rstValeCab!Tipo & "") & "'") = vbNullString Then
                    .CodigoAlmacen = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODALM", "EF2ALMACENES", "F2CODALMEXTERNO", Trim(rstValeCab!IDALMACEN & ""), "T")
                    .NumeroVale = vbNullString
                    .TipoVale = Trim(rstValeCab!Tipo & "")
                Else
                    .CodigoAlmacen = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODALM", "IF4VALES", "NUMENSAM", Trim(rstValeCab!ID & ""), "T", "AND F4TIPOVALE = '" & Trim(rstValeCab!Tipo & "") & "'")
                    .NumeroVale = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F4NUMVAL", "IF4VALES", "NUMENSAM", Trim(rstValeCab!ID & ""), "T", "AND F4TIPOVALE = '" & Trim(rstValeCab!Tipo & "") & "'")
                    .TipoVale = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F4TIPOVALE", "IF4VALES", "NUMENSAM", Trim(rstValeCab!ID & ""), "T", "AND F4TIPOVALE = '" & Trim(rstValeCab!Tipo & "") & "'")
                    
                    .obtenerConfigVale
                End If
                
                .CodigoProveedor = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODPROV", "EF2PROVEEDORES", "F2CODPROVEXTERNO", Trim(rstValeCab!IDPERSONA & ""), "T")
                .Fecha = Format(Trim(rstValeCab!Fecha & ""), "dd/mm/yyyy")
                .CodigoOrigen = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F1CODORI", "SF1ORIGENES", "F1CODORIEXTERNO", Trim(rstValeCab!TipoMovimiento & ""), "T", "AND F1TIPMOV = '" & Trim(rstValeCab!Tipo & "") & "'")
                .TipoCambio = Val(rstValeCab!TCambio & "")
                
                    If .TipoCambio = 0 Then
                        .TipoCambio = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CAMBIO", "CAMBIOS", "FECHA", .Fecha, "F")
                    End If
                        
                .CodigoMoneda = left(Trim(rstValeCab!Moneda & ""), 1)
                
                .OrdenTrabajo = Trim(rstValeCab!IdOrdenProduccion & "")
                .referencia = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NOMPROV", "EF2PROVEEDORES", "F2CODPROVEXTERNO", Trim(rstValeCab!IDPERSONA & ""), "T")
                
                .CodTipoComprobante = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODDOC", "DOCUMENTOS", "CODEXT3", Trim(rstValeCab!IDDOCUMENTO & ""), "T")
                
                Select Case .CodTipoComprobante
                    Case vbNullString
                        .CodTipoComprobante = "99"
                        
                        If Trim(rstValeCab!NRODOCUMENTO & "") <> vbNullString Then
                            If InStr(1, Trim(rstValeCab!NRODOCUMENTO & ""), "-") > 0 Then
                                .SerieDocumento = Mid(Trim(rstValeCab!NRODOCUMENTO & ""), 1, InStr(1, Trim(rstValeCab!NRODOCUMENTO & ""), "-") - 1)
                                .NumeroDocumento = Mid(Trim(rstValeCab!NRODOCUMENTO & ""), InStr(1, Trim(rstValeCab!NRODOCUMENTO & ""), "-") + 1)
                            Else
                                .NumeroDocumento = Trim(rstValeCab!NRODOCUMENTO & "")
                            End If
                        End If
                    Case "86"
                        .CodTipoComprobante = vbNullString
                        
                        .SerieGuia = Mid(Trim(rstValeCab!NRODOCUMENTO & ""), 1, InStr(1, Trim(rstValeCab!NRODOCUMENTO & ""), "-") - 1)
                        .NumeroGuia = Mid(Trim(rstValeCab!NRODOCUMENTO & ""), InStr(1, Trim(rstValeCab!NRODOCUMENTO & ""), "-") + 1)
                    Case Else
                        If Trim(rstValeCab!NRODOCUMENTO & "") <> vbNullString Then
                            If InStr(1, Trim(rstValeCab!NRODOCUMENTO & ""), "-") > 0 Then
                                .SerieDocumento = Mid(Trim(rstValeCab!NRODOCUMENTO & ""), 1, InStr(1, Trim(rstValeCab!NRODOCUMENTO & ""), "-") - 1)
                                .NumeroDocumento = Mid(Trim(rstValeCab!NRODOCUMENTO & ""), InStr(1, Trim(rstValeCab!NRODOCUMENTO & ""), "-") + 1)
                            Else
                                .NumeroDocumento = Trim(rstValeCab!NRODOCUMENTO & "")
                            End If
                        End If
                End Select
                
                .observaciones = Trim(rstValeCab!OBSERVACIONDOCUMENTO & "")
                
                .NumeroOrdenCompra = IIf(Trim(rstValeCab!IDORDENCOMPRA & "") = "0", vbNullString, Trim(rstValeCab!IDORDENCOMPRA & ""))
                .NumeroValeExterno = Trim(rstValeCab!ID & "")
                
                .FecReg = Format(Trim(rstValeCab!FechaIngreso & ""), "Short Date")
                .UsuReg = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODUSER", "EF2USERS", "F2CODUSEREXTERNO", Trim(rstValeCab!IDUSUARIOINGRESO & ""), "T")
                .FecMod = Format(Trim(rstValeCab!FECHAACTUALIZACION & ""), "Short Date")
                .UsuMod = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODUSER", "EF2USERS", "F2CODUSEREXTERNO", Trim(rstValeCab!IDUSUARIOACTUALIZACION & ""), "T")
                
                If .guardarVale(True) Then
                    Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                    
                    'Borrar Detalle de Vale
                    SqlCad = vbNullString
                    SqlCad = "DELETE FROM IF3VALES WHERE F2CODALM = '" & .CodigoAlmacen & "' AND F4NUMVAL = '" & .NumeroVale & "'"
                    
                    cnn_dbbancos.Execute SqlCad
                    Actualiza_Log SqlCad, cnn_dbbancos.ConnectionString
                    
                    SqlCad = vbNullString
                    SqlCad = SqlCad & "SELECT "
                    'SqlCad = SqlCad & "RQ.IDPEDIDO, "
                    'SqlCad = SqlCad & "OP.IDREQUERIMIENTO, "
                    'SqlCad = SqlCad & "OPD.IDORDENPRODUCCION, "
                    'SqlCad = SqlCad & "OPD.IDINSUMO AS OPDIDINSUMO, "
                    'SqlCad = SqlCad & "OPD.IDINSUMOV AS OPDIDINSUMOV, "
                    SqlCad = SqlCad & "DET.IDINSUMO, "
                    SqlCad = SqlCad & "DET.CANTIDAD, "
                    SqlCad = SqlCad & "DET.COSTO, "
                    SqlCad = SqlCad & "DET.IMPORTE, "
                    
                        Select Case .TipoVale
                            Case "I"
                                SqlCad = SqlCad & "DET.MONTODESCUENTO, "
                                SqlCad = SqlCad & "DET.PORCDESCUENTO "
                            Case "S"
                                SqlCad = SqlCad & "NULL AS MONTODESCUENTO, "
                                SqlCad = SqlCad & "NULL AS PORCDESCUENTO "
                        End Select
                    
                    'SqlCad = SqlCad & "FROM ((("
                    SqlCad = SqlCad & "FROM "
                    
                        Select Case .TipoVale
                            Case "I"
                                SqlCad = SqlCad & "INGRESODETALLE AS DET "
                                'SqlCad = SqlCad & "LEFT JOIN INGRESO AS CAB ON CAB.IDINGRESO = DET.IDINGRESO "
                            Case "S"
                                SqlCad = SqlCad & "SALIDADETALLE AS DET "
                                'SqlCad = SqlCad & "LEFT JOIN SALIDA AS CAB ON CAB.IDSALIDA = DET.IDSALIDA "
                        End Select
                        
                        'SqlCad = SqlCad & "LEFT JOIN ORDENPRODUCCIONDESCARGO AS OPD ON "
                        'SqlCad = SqlCad & "OPD.IDORDENPRODUCCION = CAB.IDORDENPRODUCCION AND OPD.IDINSUMO = DET.IDINSUMO) "
                        'SqlCad = SqlCad & "LEFT JOIN ORDENPRODUCCION AS OP ON OP.IDORDENPRODUCCION = OPD.IDORDENPRODUCCION) "
                        'SqlCad = SqlCad & "LEFT JOIN REQUERIMIENTO AS RQ ON RQ.IDREQUERIMIENTO = OP.IDREQUERIMIENTO "
                    
                    SqlCad = SqlCad & "WHERE "
                    
                        Select Case .TipoVale
                            Case "I"
                                SqlCad = SqlCad & "DET.IDINGRESO = " & .NumeroValeExterno
                            Case "S"
                                SqlCad = SqlCad & "DET.IDSALIDA = " & .NumeroValeExterno
                        End Select
                    
                    SqlCad = SqlCad & " ORDER BY "
                    SqlCad = SqlCad & "DET.ITEMS"
                    
                    If rstValeDet.State = 1 Then rstValeDet.Close
                    
                    rstValeDet.Open SqlCad, cnBdStudioModa, adOpenForwardOnly, adLockReadOnly
                    
                    If Not rstValeDet.EOF Then
                        dblItem = 0
                        
                        '/-- Lectura de los datos del detalle hasta que no exista ningun dato
                        Do While Not rstValeDet.EOF
                            .inicializarEntidadesDetalle
                            
                            dblItem = dblItem + 1
                            
                            '.Requerimiento = Trim(rstValeDet!IDPEDIDO & "")
                            
                            '.CodigoProducto = IIf(Trim(rstValeDet!OPDIDINSUMO & "") <> vbNullString, Trim(rstValeDet!OPDIDINSUMO & ""), Trim(rstValeDet!IDINSUMO & ""))
                            .CodigoProducto = Trim(rstValeDet!IDINSUMO & "")
                            .CodigoProductoOriginal = Trim(rstValeDet!IDINSUMO & "")
                            .Cantidad = Val(rstValeDet!Cantidad & "")
                            
                            Select Case .CodigoMoneda
                                Case "S"
                                    .ValorVenta = Val(rstValeDet!costo & "")
                                    .IGV = Val(Format((Val(rstValeDet!Importe & "") - Val(rstValeDet!MONTODESCUENTO & "")) * Val(rstValeCab!PORCIGV & "") / 100, "#0.0000"))
                                    .TOTAL = Val(Format((Val(rstValeDet!Importe & "") - Val(rstValeDet!MONTODESCUENTO & "")) + .IGV, "#0.0000"))
                                    .ValorVentaDol = Val(Format(.ValorVenta / .TipoCambio, "#0.0000"))
                                    .IgvDol = Val(Format(.IGV / .TipoCambio, "#0.0000"))
                                    .TotalDol = Val(Format(.TOTAL / .TipoCambio, "#0.0000"))
                                Case "D"
                                    .ValorVentaDol = Val(rstValeDet!costo & "")
                                    .IgvDol = Val(Format((Val(rstValeDet!Importe & "") - Val(rstValeDet!MONTODESCUENTO & "")) * Val(rstValeCab!PORCIGV & "") / 100, "#0.0000"))
                                    .TotalDol = Val(Format((Val(rstValeDet!Importe & "") - Val(rstValeDet!MONTODESCUENTO & "")) + .IgvDol, "#0.0000"))
                                    .ValorVenta = Val(Format(.ValorVentaDol * .TipoCambio, "#0.0000"))
                                    .IGV = Val(Format(.IgvDol * .TipoCambio, "#0.0000"))
                                    .TOTAL = Val(Format(.TotalDol * .TipoCambio, "#0.0000"))
                            End Select
                            
                            .ITEM = dblItem
                            
                            .guardarValeDetalleOneByOne
                            
                            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
    
                            rstValeDet.MoveNext
                        Loop
                    End If
                End If
            End With
            
            DoEvents
            
            barraProgreso.Value = barraProgreso.Value + 1
            frameProgreso.Caption = " Actualización al... " & FormatPercent(barraProgreso.Value / barraProgreso.Max, 3)
            
            rstValeCab.MoveNext
        Loop
            MsgBox "Actualización Finalizada", vbInformation + vbOKOnly, wnomcia
            
            Select Case strTipoVale
                Case "I"
                    ModUtilitario.sWrtIni App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UltimaLecturaDeValesIngreso", Format(Date, "Short Date")
                Case "S"
                    ModUtilitario.sWrtIni App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UltimaLecturaDeValesSalida", Format(Date, "Short Date")
            End Select
    Else
        MsgBox "No se ubicaron registros pendientes de importación.", vbInformation + vbOKOnly, wnomcia
    End If
    
    
    SqlCad = vbNullString
    strUltimaLecturaDeValesIngreso = vbNullString
    strUltimaLecturaDeValesSalida = vbNullString
    dblItem = 0
    
    frameProgreso.Visible = False
    
    If rstValeCab.State = 1 Then rstValeCab.Close
    If rstValeDet.State = 1 Then rstValeDet.Close
    
    Set rstValeCab = Nothing
    Set rstValeDet = Nothing
    
    Exit Sub
errImportarValesServidorExterno:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    frameProgreso.Visible = False
    barraProgreso.Value = 0
    
    Err.Clear
End Sub

Public Function exportarValeAserverSQL(ByVal strCodigoAlmacen As String, _
                                        ByVal strNumeroVale As String, _
                                        ByVal etiquetaIDExterno As Object) As Boolean
    
    On Error GoTo errExportarValeAserverSQL
    
    Dim cmdSP As New ADODB.Command
    
    Dim strTablaSQL As String
    Dim rstValeDet As New ADODB.Recordset
    Dim strNombrePC As String
    Dim bolInicioTransaccion As Boolean
    Dim dblItem As Double
    Dim strTipoComprobanteProvDefault As String
    
    'Variables para Ingreso/Salida
    Dim dblIdVale As Double
    Dim strIdTipoVale As String
    Dim strIdDocumento As String
    Dim strNumDoc As String
    Dim StrFecha As String
    Dim strIdMoneda As String
    Dim strIdAlmacen As String
    Dim strIdPersona As String
    Dim dblSubTotal As Double
    Dim dblIgv As Double
    Dim dblTotal As Double
    Dim dblTCambio As Double
    Dim dblPorcIGV As Double
    Dim dblMontoDscto As Double
    Dim dblPorcDscto As Double
    Dim strObservacion As String
    Dim strObservacionAnulacion As String
    Dim strIdPuntoVenta As String
    Dim strIdUsuarioIngreso As String
    Dim strIdUsuarioActualizacion As String
    Dim strFechaIngreso As String
    Dim strFechaActualizacion As String
    Dim intAnulado As Integer
    Dim dblIdOrdenProduccion As Double
    Dim dblIdOrdenCompra As Double
    
    bolInicioTransaccion = False
    
    'abrirCnDBMilano
    
    strNombrePC = modgeneral.ComputerName
    
    With objAyudaVale
        .inicializarEntidades
        
        .CodigoAlmacen = Trim(strCodigoAlmacen)
        .NumeroVale = Trim(strNumeroVale)
        
        If .verificarExistencia Then
            .obtenerConfigVale
            
            Select Case .TipoVale
                Case "I"
                    strTablaSQL = "INGRESO"
                Case "S"
                    strTablaSQL = "SALIDA"
            End Select
            
            If Not .ExportarVale Then
                MsgBox "Vale no exportardo; no cuenta con marca de Exportación de registro." & vbNewLine & vbNewLine & _
                        "Consulte con su administrador de sistemas.", vbInformation + vbOKOnly, App.ProductName & " - " & wnomcia
                
                exportarValeAserverSQL = False
        
                Exit Function
            End If
            
            strIdPuntoVenta = ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "IDPUNTOVENTA", "PUNTOVENTA", "NOMBREPC", strNombrePC, "T", "AND ELIMINADO = 0")
            
            If Trim(strIdPuntoVenta) = vbNullString Then
                MsgBox "Vale no exportado. Su computador no esta habilitado para generar registros de " & strTablaSQL & "." & vbNewLine & vbNewLine & _
                                "Consulte con su administrador de sistemas.", vbInformation + vbOKOnly, App.ProductName & " - " & wnomcia
                
                exportarValeAserverSQL = False
                
                Exit Function
            End If
            
            dblIdVale = Val(.NumeroValeExterno)
            strIdTipoVale = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F1CODORIEXTERNO", "SF1ORIGENES", "F1CODORI", .CodigoOrigen, "T")
            'strIdDocumento = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CODEXT3", "DOCUMENTOS", "F2CODDOC", .CodTipoComprobante, "T")
            
            Select Case .TipoPersona
                Case "P"
                    strTipoComprobanteProvDefault = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2TIPDOC", "EF2PROVEEDORES", "F2CODPROV", .CodigoProveedor, "T")
                    
                    Select Case strTipoComprobanteProvDefault
                        Case "86"
                            If Trim(.NumeroGuia) <> vbNullString Then
                                strIdDocumento = "DOC0000006"
                                
                                strNumDoc = IIf(.SerieGuia <> vbNullString, Format(.SerieGuia, "000") & "-", vbNullString) & Format(.NumeroGuia, "0000000")
                            ElseIf Trim(.NumeroDocumento) <> vbNullString Then
                                strIdDocumento = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CODEXT3", "DOCUMENTOS", "F2CODDOC", .CodTipoComprobante, "T")
                                
                                strNumDoc = IIf(.SerieDocumento <> vbNullString, Format(.SerieDocumento, "000") & "-", vbNullString) & Format(.NumeroDocumento, "0000000")
                            Else
                                strIdDocumento = "DOC0000004"
                                strNumDoc = vbNullString
                            End If
                        Case Else
                            If Trim(.NumeroDocumento) <> vbNullString Then
                                strIdDocumento = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CODEXT3", "DOCUMENTOS", "F2CODDOC", .CodTipoComprobante, "T")
                                
                                strNumDoc = IIf(.SerieDocumento <> vbNullString, Format(.SerieDocumento, "000") & "-", vbNullString) & Format(.NumeroDocumento, "0000000")
                            ElseIf Trim(.NumeroGuia) <> vbNullString Then
                                strIdDocumento = "DOC0000006"
                                
                                strNumDoc = IIf(.SerieGuia <> vbNullString, Format(.SerieGuia, "000") & "-", vbNullString) & Format(.NumeroGuia, "0000000")
                            Else
                                strIdDocumento = "DOC0000004"
                                strNumDoc = vbNullString
                            End If
                    End Select
                Case Else
                    If Trim(.NumeroGuia) <> vbNullString Then
                        strIdDocumento = "DOC0000006"
                    Else
                        If Trim(strIdDocumento) = vbNullString Then
                            strIdDocumento = "DOC0000004"
                        End If
                    End If
                    
                    Select Case strIdDocumento
                        Case "DOC0000006"
                            strNumDoc = IIf(.SerieGuia <> vbNullString, Format(.SerieGuia, "000") & "-", vbNullString) & Format(.NumeroGuia, "0000000")
                        Case "DOC0000004"
                            strNumDoc = vbNullString
                        Case Else
                            strNumDoc = IIf(.SerieDocumento <> vbNullString, Format(.SerieDocumento, "000") & "-", vbNullString) & Format(.NumeroDocumento, "0000000")
                    End Select
            End Select
            
            'strFecha = Format(.Fecha, "dd-mm-yy hh:mm:ss")
            Select Case .TipoVale
                Case "I"
                    StrFecha = Format(.Fecha, "dd-mm-yy 00:00:00")
                Case "S"
                    StrFecha = Format(.Fecha, "dd-mm-yy 12:00:00")
            End Select
            
            dblTCambio = .TipoCambio
            
            Select Case .CodigoMoneda
                Case "S"
                    strIdMoneda = "MON0000001"
                    
                    dblSubTotal = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "FORMAT(SUM(F3CANPRO * F3VALVTA),'#0.00') AS SUMA", "IF3VALES", "F2CODALM", .CodigoAlmacen, "T", "AND F4NUMVAL = '" & .NumeroVale & "'"))
                    dblIgv = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "FORMAT(SUM(F3IGV),'#0.00') AS SUMA", "IF3VALES", "F2CODALM", .CodigoAlmacen, "T", "AND F4NUMVAL = '" & .NumeroVale & "'"))
                    dblTotal = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "FORMAT(SUM(F3TOTITE),'#0.00') AS SUMA", "IF3VALES", "F2CODALM", .CodigoAlmacen, "T", "AND F4NUMVAL = '" & .NumeroVale & "'"))
                Case "D"
                    strIdMoneda = "MON0000002"
                    
                    dblSubTotal = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "FORMAT(SUM(F3CANPRO * F3VALDOL),'#0.00') AS SUMA", "IF3VALES", "F2CODALM", .CodigoAlmacen, "T", "AND F4NUMVAL = '" & .NumeroVale & "'"))
                    dblIgv = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "FORMAT(SUM(F3IGVDOL),'#0.00') AS SUMA", "IF3VALES", "F2CODALM", .CodigoAlmacen, "T", "AND F4NUMVAL = '" & .NumeroVale & "'"))
                    dblTotal = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "FORMAT(SUM(F3TOTDOL),'#0.00') AS SUMA", "IF3VALES", "F2CODALM", .CodigoAlmacen, "T", "AND F4NUMVAL = '" & .NumeroVale & "'"))
            End Select
            
            strIdAlmacen = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODALMEXTERNO", "EF2ALMACENES", "F2CODALM", .CodigoAlmacen, "T")
            
            Select Case .TipoPersona
                Case "C"
                    strIdPersona = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODCLIEXTERNO", "EF2CLIENTES", "F2CODCLI", .CodigoProveedor, "T")
                Case "P"
                    strIdPersona = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODPROVEXTERNO", "EF2PROVEEDORES", "F2CODPROV", .CodigoProveedor, "T")
                Case Else
                    strIdPersona = vbNullString
            End Select
            
            dblPorcIGV = Val(wIgv)
            dblMontoDscto = 0
            dblPorcDscto = 0
            strObservacion = .observaciones
            strIdUsuarioIngreso = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODUSEREXTERNO", "EF2USERS", "F2CODUSER", .UsuReg, "T")
            strIdUsuarioActualizacion = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODUSEREXTERNO", "EF2USERS", "F2CODUSER", .UsuMod, "T")
            strFechaIngreso = Format(.FecReg, "dd-mm-yy hh:mm:ss")
            strFechaActualizacion = Format(.FecMod, "dd-mm-yy hh:mm:ss")
            dblIdOrdenProduccion = Val(.OrdenTrabajo)
            
            SqlCad = vbNullString
            
            'cnBdStudioModa.BeginTrans
            
            'bolInicioTransaccion = True
            
            'GUARDAR CABECERA DE VALE
            If ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "ID" & strTablaSQL, strTablaSQL, "ID" & strTablaSQL, Val(.NumeroValeExterno), "N") = vbNullString Then
                'GENERAR ID
                dblIdVale = Val(ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "CORRELATIVO", "CORRELATIVOTABLA", "IDPUNTOVENTA", strIdPuntoVenta, "T", _
                                                                    "AND TABLA = '" & strTablaSQL & "'"))
                
                If dblIdVale = 0 Then
                    MsgBox "Vale no exportado. Su computador no esta habilitado para generar registros de [" & strTablaSQL & "]." & vbNewLine & vbNewLine & _
                            "Consulte con su administrador de sistemas.", vbInformation + vbOKOnly, App.ProductName & " - " & wnomcia
                    
                    .inicializarEntidades
                    
                    'cnBdStudioModa.RollbackTrans
                    
                    exportarValeAserverSQL = False
                    
                    Exit Function
                End If
                
                'ACTUALIZAR CORRELATIVO
                SqlCad = vbNullString
                SqlCad = SqlCad & "UPDATE "
                SqlCad = SqlCad & "CORRELATIVOTABLA "
                SqlCad = SqlCad & "SET "
                SqlCad = SqlCad & "CORRELATIVO = " & dblIdVale + 1 & " "
                SqlCad = SqlCad & "WHERE "
                SqlCad = SqlCad & "IDPUNTOVENTA = '" & strIdPuntoVenta & "' AND "
                SqlCad = SqlCad & "TABLA = '" & strTablaSQL & "'"
                
                cnBdStudioModa.Execute SqlCad
        
                Actualiza_Log " < DB Externo > " & SqlCad, StrConexDbBancos
                
                'INSERTAR
                SqlCad = vbNullString
                SqlCad = SqlCad & "INSERT INTO " & strTablaSQL & "("
                SqlCad = SqlCad & "ID" & strTablaSQL & ", IDTIPO" & strTablaSQL & ", IDDOCUMENTO, "
                SqlCad = SqlCad & "[NUM-DOC], FECHA, IDMONEDA, "
                SqlCad = SqlCad & "IDALMACEN, IDPERSONA, SUBTOTAL, "
                SqlCad = SqlCad & "IGV, TOTAL, TCAMBIO, "
                SqlCad = SqlCad & "PORCIGV, "
                    
                    If .TipoVale = "I" Then
                        SqlCad = SqlCad & "MONTODESCUENTO , PORCDESCUENTO, "
                    End If
                    
                SqlCad = SqlCad & "OBSERVACIONDOCUMENTO, IDPUNTOVENTA, IDUSUARIOINGRESO, "
                SqlCad = SqlCad & "IDUSUARIOACTUALIZACION, FECHAINGRESO, FECHAACTUALIZACION, "
                SqlCad = SqlCad & "ANULADO, IDORDENPRODUCCION) "
                
                SqlCad = SqlCad & "VALUES("
                SqlCad = SqlCad & dblIdVale & ", '" & strIdTipoVale & "', '" & strIdDocumento & "', "
                SqlCad = SqlCad & "'" & strNumDoc & "', CONVERT(DATETIME,'" & StrFecha & "', 5), '" & strIdMoneda & "', "
                SqlCad = SqlCad & "'" & strIdAlmacen & "', " & IIf(strIdPersona <> vbNullString, "'" & strIdPersona & "'", "NULL") & ", " & dblSubTotal & ", "
                SqlCad = SqlCad & dblIgv & ", " & dblTotal & ", " & dblTCambio & ","
                SqlCad = SqlCad & dblPorcIGV & ", "
                    
                    If .TipoVale = "I" Then
                        SqlCad = SqlCad & dblMontoDscto & ", " & dblPorcDscto & ", "
                    End If
                    
                SqlCad = SqlCad & "'" & strObservacion & "', '" & strIdPuntoVenta & "', '" & strIdUsuarioIngreso & "', "
                SqlCad = SqlCad & "'" & strIdUsuarioIngreso & "', CONVERT(DATETIME,'" & strFechaIngreso & "', 5), CONVERT(DATETIME,'" & strFechaIngreso & "', 5), "
                SqlCad = SqlCad & "0, " & dblIdOrdenProduccion
                SqlCad = SqlCad & ")"
            Else
                'MODIFICAR
                SqlCad = vbNullString
                SqlCad = SqlCad & "UPDATE "
                SqlCad = SqlCad & strTablaSQL & " "
                SqlCad = SqlCad & "SET "
                SqlCad = SqlCad & "IDTIPO" & strTablaSQL & " = '" & strIdTipoVale & "', "
                SqlCad = SqlCad & "IDDOCUMENTO = '" & strIdDocumento & "', "
                SqlCad = SqlCad & "[NUM-DOC] = '" & strNumDoc & "', "
                SqlCad = SqlCad & "FECHA = CONVERT(DATETIME, '" & StrFecha & "', 5), "
                SqlCad = SqlCad & "IDMONEDA = '" & strIdMoneda & "', "
                SqlCad = SqlCad & "IDALMACEN = '" & strIdAlmacen & "', "
                SqlCad = SqlCad & "IDPERSONA = " & IIf(strIdPersona <> vbNullString, "'" & strIdPersona & "'", "NULL") & ", "
                SqlCad = SqlCad & "SUBTOTAL = " & dblSubTotal & ", "
                SqlCad = SqlCad & "IGV = " & dblIgv & ", "
                SqlCad = SqlCad & "TOTAL = " & dblTotal & ", "
                SqlCad = SqlCad & "TCAMBIO = " & dblTCambio & ", "
                SqlCad = SqlCad & "PORCIGV = " & dblPorcIGV & ", "
                    
                    If .TipoVale = "I" Then
                        SqlCad = SqlCad & "MONTODESCUENTO = " & dblMontoDscto & ", "
                        SqlCad = SqlCad & "PORCDESCUENTO = " & dblPorcDscto & ", "
                    End If
                    
                SqlCad = SqlCad & "OBSERVACIONDOCUMENTO = '" & strObservacion & "', "
                SqlCad = SqlCad & "IDPUNTOVENTA = '" & strIdPuntoVenta & "', "
                SqlCad = SqlCad & "IDUSUARIOACTUALIZACION = '" & strIdUsuarioActualizacion & "', "
                SqlCad = SqlCad & "FECHAACTUALIZACION = CONVERT(DATETIME,'" & strFechaActualizacion & "', 5), "
                SqlCad = SqlCad & "ANULADO = 0, "
                SqlCad = SqlCad & "IDORDENPRODUCCION = " & dblIdOrdenProduccion & " "
                SqlCad = SqlCad & "WHERE "
                SqlCad = SqlCad & "ID" & strTablaSQL & " = " & dblIdVale
            End If
            
            cnBdStudioModa.Execute SqlCad
        
            Actualiza_Log "< DB Externo > " & SqlCad, StrConexDbBancos
            
            'VERIFICAR SI DETALLE ESTA REGISTRADO, PARA REVERTIR STOCK ANTES DE ELIMINAR VALE EXTERNO
            If rstValeDet.State = 1 Then rstValeDet.Close
            
            rstValeDet.Open "SELECT * FROM " & strTablaSQL & "DETALLE WHERE ID" & strTablaSQL & " = '" & dblIdVale & "'", cnBdStudioModa, adOpenForwardOnly, adLockReadOnly
            
            If Not rstValeDet.EOF Then
                rstValeDet.MoveFirst
                
                Do While Not rstValeDet.EOF
                    'Actualizar Tabla InsumoStock
                    SqlCad = vbNullString
                    SqlCad = SqlCad & "UPDATE "
                    SqlCad = SqlCad & "INSUMOSTOCK "
                    SqlCad = SqlCad & "SET "
                        
                        Select Case .TipoVale
                            Case "I"
                                SqlCad = SqlCad & "STKACTUAL = (STKACTUAL - " & Val(rstValeDet!Cantidad & "") & "), "
                                SqlCad = SqlCad & "STKCOMPRA = (STKCOMPRA - " & Val(rstValeDet!Cantidad & "") & ") "
                            Case "S"
                                SqlCad = SqlCad & "STKACTUAL = (STKACTUAL + " & Val(rstValeDet!Cantidad & "") & "), "
                                SqlCad = SqlCad & "STKCOMPRA = (STKCOMPRA + " & Val(rstValeDet!Cantidad & "") & ") "
                        End Select
                    
                    SqlCad = SqlCad & "WHERE "
                    SqlCad = SqlCad & "LTRIM(RTRIM(IDINSUMO)) = '" & Trim(rstValeDet!IDINSUMO & "") & "' AND "
                    SqlCad = SqlCad & "LTRIM(RTRIM(IDALMACEN)) = '" & strIdAlmacen & "'"
                    
                    cnBdStudioModa.Execute SqlCad
                    
                    Actualiza_Log "< DB Externo > " & SqlCad, StrConexDbBancos
                    
                    rstValeDet.MoveNext
                Loop
            End If
            
            'GUARDAR DETALLE DEL VALE
            If rstValeDet.State = 1 Then rstValeDet.Close
            
            rstValeDet.Open "SELECT * FROM IF3VALES WHERE F2CODALM = '" & .CodigoAlmacen & "' AND F4NUMVAL = '" & .NumeroVale & "'", cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
            
            If Not rstValeDet.EOF Then
                rstValeDet.MoveFirst
                
                SqlCad = vbNullString
                SqlCad = "DELETE FROM " & strTablaSQL & "DETALLE WHERE ID" & strTablaSQL & " = " & dblIdVale
                
                cnBdStudioModa.Execute SqlCad
            
                Actualiza_Log "< DB Externo > " & SqlCad, StrConexDbBancos
                
                dblItem = 0
                
                Do While Not rstValeDet.EOF
                    dblItem = dblItem + 1
                    
                    SqlCad = vbNullString
                    SqlCad = SqlCad & "INSERT INTO " & strTablaSQL & "DETALLE("
                    SqlCad = SqlCad & "ID" & strTablaSQL & ", [ITEMS], IDINSUMO, "
                    SqlCad = SqlCad & "CANTIDAD, COSTO, IMPORTE, COSTOPROMEDIO"
                    
                        If .TipoVale = "I" Then
                            SqlCad = SqlCad & ", MONTODESCUENTO, PORCDESCUENTO, OBSERVACION"
                        End If
                        
                    SqlCad = SqlCad & ") "
                    SqlCad = SqlCad & "VALUES("
                    SqlCad = SqlCad & dblIdVale & ", " & dblItem & ", '" & Trim(rstValeDet!f5codpro & "") & "', "
                    SqlCad = SqlCad & Val(rstValeDet!F3CANPRO & "") & ", "
                    
                    Select Case .CodigoMoneda
                        Case "S"
                            SqlCad = SqlCad & Val(rstValeDet!F3VALVTA & "") & ", " & Val(rstValeDet!F3VALVTA & "") * Val(rstValeDet!F3CANPRO & "") & ", "
                            SqlCad = SqlCad & Val(rstValeDet!F3VALVTA & "")
                        Case "D"
                            SqlCad = SqlCad & Val(rstValeDet!F3VALDOL & "") & ", " & Val(rstValeDet!F3VALDOL & "") * Val(rstValeDet!F3CANPRO & "") & ", "
                            SqlCad = SqlCad & Val(rstValeDet!F3VALDOL & "")
                    End Select
                    
                    'SqlCad = SqlCad & Val(rstValeDet!F3VALDOL & "")
                    
                        If .TipoVale = "I" Then
                            'SqlCad = SqlCad & ", 0, 0, ''"
                            SqlCad = SqlCad & ", " & Val(rstValeDet!F3MONTODSCTO & "") & ", "
                            SqlCad = SqlCad & Val(rstValeDet!F3PORCENTAJEDSCTO & "") & ", "
                            SqlCad = SqlCad & "'" & Trim(rstValeDet!observaciones & "") & "'"
                        End If
                        
                    SqlCad = SqlCad & ")"
                    
                    cnBdStudioModa.Execute SqlCad
            
                    Actualiza_Log "< DB Externo > " & SqlCad, StrConexDbBancos
                    
                    'ACTUALIZAR TABLA INSUMOSTOCK
                    SqlCad = vbNullString
                    SqlCad = SqlCad & "UPDATE "
                    SqlCad = SqlCad & "INSUMOSTOCK "
                    SqlCad = SqlCad & "SET "
                        
                        Select Case .TipoVale
                            Case "I"
                                SqlCad = SqlCad & "STKACTUAL = (STKACTUAL + " & Val(rstValeDet!F3CANPRO & "") & "), "
                                SqlCad = SqlCad & "STKCOMPRA = (STKCOMPRA + " & Val(rstValeDet!F3CANPRO & "") & ") "
                            Case "S"
                                SqlCad = SqlCad & "STKACTUAL = (STKACTUAL - " & Val(rstValeDet!F3CANPRO & "") & "), "
                                SqlCad = SqlCad & "STKCOMPRA = (STKCOMPRA - " & Val(rstValeDet!F3CANPRO & "") & ") "
                        End Select
                        
                    SqlCad = SqlCad & "WHERE "
                    SqlCad = SqlCad & "LTRIM(RTRIM(IDINSUMO)) = '" & Trim(rstValeDet!f5codpro & "") & "' AND "
                    SqlCad = SqlCad & "LTRIM(RTRIM(IDALMACEN)) = '" & strIdAlmacen & "'"
                    
                    cnBdStudioModa.Execute SqlCad
                    
                    Actualiza_Log "< DB Externo > " & SqlCad, StrConexDbBancos
                    
                    'ACTUALIZAR TABLA INSUMOPERSONA
                    If strIdPersona <> vbNullString Then
                        If Val(ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "COUNT(*)", "PERSONA", "IDPERSONA", strIdPersona, "T")) > 0 Then
                            If Val(ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "COUNT(*)", "INSUMOPERSONA", "IDPERSONA", strIdPersona, "T", "AND IDINSUMO = '" & Trim(rstValeDet!f5codpro & "") & "'")) = 0 Then
                                SqlCad = vbNullString
                                SqlCad = SqlCad & "INSERT INTO INSUMOPERSONA ("
                                SqlCad = SqlCad & "IDINSUMO, ITEMS, IDPERSONA, IDMONEDA, COSTO, COSTOPROMEDIO, "
                                SqlCad = SqlCad & "IDUSUARIOINGRESO,IDUSUARIOACTUALIZACION,FECHAINGRESO,FECHAACTUALIZACION) "
                                SqlCad = SqlCad & "VALUES ("
                                SqlCad = SqlCad & "'" & Trim(rstValeDet!f5codpro & "") & "', "
                                SqlCad = SqlCad & Val(ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "COUNT(*)", "INSUMOPERSONA", "IDINSUMO", Trim(rstValeDet!f5codpro & ""), "T")) + 1 & ", "
                                SqlCad = SqlCad & "'" & strIdPersona & "', "
                                SqlCad = SqlCad & "'" & strIdMoneda & "', "
                                SqlCad = SqlCad & Val(rstValeDet!F3VALVTA & "") & ", "
                                SqlCad = SqlCad & "0, "
                                SqlCad = SqlCad & "'" & strIdUsuarioIngreso & "', "
                                SqlCad = SqlCad & "'" & IIf(strIdUsuarioActualizacion <> vbNullString, strIdUsuarioActualizacion, strIdUsuarioIngreso) & "', "
                                SqlCad = SqlCad & "GETDATE(), "
                                SqlCad = SqlCad & "GETDATE())"
                            Else
                                SqlCad = SqlCad & "UPDATE "
                                SqlCad = SqlCad & "INSUMOPERSONA "
                                SqlCad = SqlCad & "SET "
                                SqlCad = SqlCad & "IDMONEDA = '" & strIdMoneda & "', "
                                SqlCad = SqlCad & "COSTO = " & Val(rstValeDet!F3VALVTA & "") & ", "
                                SqlCad = SqlCad & "IDUSUARIOACTUALIZACION = '" & IIf(strIdUsuarioActualizacion <> vbNullString, strIdUsuarioActualizacion, strIdUsuarioIngreso) & "', "
                                SqlCad = SqlCad & "FECHAACTUALIZACION = GETDATE() "
                                SqlCad = SqlCad & "WHERE "
                                SqlCad = SqlCad & "IDPERSONA = '" & strIdPersona & "' AND "
                                SqlCad = SqlCad & "IDINSUMO = '" & Trim(rstValeDet!f5codpro & "") & "'"
                            End If
                            
                            cnBdStudioModa.Execute SqlCad
                    
                            Actualiza_Log "< DB Externo > " & SqlCad, StrConexDbBancos
                        End If
                    End If
                    
                    'VERIFICAR SI YA SE REALIZO EL CIERRE
                    If Val(ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "COUNT(*)", "CIERREMENSUAL", "IDALMACEN", strIdAlmacen, "T", "AND (AÑO = " & Year(CDate(StrFecha)) & " AND MES = " & Month(CDate(StrFecha)) & ")")) > 0 Then
                        With cmdSP
                            .CommandType = adCmdStoredProc
                            .CommandText = "USP_ActualizaCPIngreso"
                        End With
                        
                        Set cmdSP.ActiveConnection = cnBdStudioModa
                        
                        With cmdSP.Parameters
                            .Append cmdSP.CreateParameter("@IdPuntoVenta", adVarChar, adParamInput, 10, strIdPuntoVenta)
                            .Append cmdSP.CreateParameter("@IdUsuario", adVarChar, adParamInput, 10, IIf(strIdUsuarioActualizacion <> vbNullString, strIdUsuarioActualizacion, strIdUsuarioIngreso))
                            .Append cmdSP.CreateParameter("@IdAlmacen", adVarChar, adParamInput, 10, strIdAlmacen)
                            .Append cmdSP.CreateParameter("@FechaActualiza", adDBTimeStamp, adParamInput, 10, IIf(strFechaActualizacion <> vbNullString, strFechaActualizacion, strFechaIngreso))
                            .Append cmdSP.CreateParameter("@IdInsumo", adVarChar, adParamInput, 50, Trim(rstValeDet!f5codpro & ""))
                        End With
                        
                        cmdSP.Execute
                    End If
                    
                    rstValeDet.MoveNext
                Loop
            End If
            
            'cnBdStudioModa.CommitTrans
            
            'bolInicioTransaccion = False
            
            SqlCad = vbNullString
            SqlCad = SqlCad & "UPDATE "
            SqlCad = SqlCad & "IF4VALES "
            SqlCad = SqlCad & "SET "
            SqlCad = SqlCad & "NUMENSAM = '" & dblIdVale & "' "
            SqlCad = SqlCad & "WHERE "
            SqlCad = SqlCad & "F2CODALM = '" & .CodigoAlmacen & "' AND "
            SqlCad = SqlCad & "F4NUMVAL = '" & .NumeroVale & "'"
            
            cnn_dbbancos.Execute SqlCad
            
            Actualiza_Log SqlCad, StrConexDbBancos
            
            etiquetaIDExterno.Caption = Trim(dblIdVale & "")
            
            exportarValeAserverSQL = True
        Else
            MsgBox "El Vale no se guardo correctamente, vuelva a intentarlo.", vbInformation + vbOKOnly, App.ProductName
            
            exportarValeAserverSQL = False
            
            .inicializarEntidades
        End If
    End With
    
    Exit Function
errExportarValeAserverSQL:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description & vbNewLine & vbNewLine & _
            "Vale no exportado, intente guardar nuevamente el Vale.", vbInformation + vbOKOnly, App.ProductName & " - ExportarValeAserverSQL"
    'Resume
    'If bolInicioTransaccion Then
        'cnBdStudioModa.RollbackTrans
        
        Actualiza_Log "Exportación de Vale [" & objAyudaVale.CodigoAlmacen & "][" & objAyudaVale.NumeroVale & "] cancelada por el siguiente error: " & _
                        "[Numero Error: " & Err.Number & "] [Descripción: " & Err.Description & "]", StrConexDbBancos
    'Else
    '    Actualiza_Log SqlCad, StrConexDbBancos
    '
    '    Actualiza_Log "Procesamiento de Vale [" & objAyudaVale.CodigoAlmacen & "][" & objAyudaVale.NumeroVale & "] trunco por el siguiente error: " & _
    '                    "[Numero Error: " & Err.Number & "][Descripción: " & Err.Description & "]", StrConexDbBancos
    'End If
    
    Err.Clear
End Function

Public Function exportarValeAserverSQLv2(ByVal strCodigoAlmacen As String, _
                                            ByVal strNumeroVale As String, _
                                            ByVal etiquetaIDExterno As Object, _
                                            ByVal frameProgreso As Object, _
                                            ByVal barraProgreso As Object) As Boolean
    
    On Error GoTo errExportarValeAserverSQLv2
    
    Dim objValeExterno As New ClsVale
    
    Dim strTablaSQL As String
    Dim rstValeDet As New ADODB.Recordset
    Dim strNombrePC As String
    Dim dblItem As Double
    Dim strTipoComprobanteProvDefault As String
    Dim bolNuevoRegistro As Boolean
    
    'Variables para Ingreso/Salida
    Dim dblIdVale As Double
    Dim strIdTipoVale As String
    Dim strIdDocumento As String
    Dim strNumDoc As String
    Dim StrFecha As String
    Dim strIdMoneda As String
    Dim strIdAlmacen As String
    Dim strIdPersona As String
    Dim dblSubTotal As Double
    Dim dblIgv As Double
    Dim dblTotal As Double
    Dim dblTCambio As Double
    Dim dblPorcIGV As Double
    Dim dblMontoDscto As Double
    Dim dblPorcDscto As Double
    Dim strObservacion As String
    Dim strObservacionAnulacion As String
    Dim strIdPuntoVenta As String
    Dim strIdUsuarioIngreso As String
    Dim strIdUsuarioActualizacion As String
    Dim strFechaIngreso As String
    Dim strFechaActualizacion As String
    Dim intAnulado As Integer
    Dim dblIdOrdenProduccion As Double
    Dim dblIdOrdenCompra As Double
    
    Dim dblCostoPromedio As Double
    Dim dblStockActual As Double
    
    'abrirCnDBMilano
    
    strNombrePC = modgeneral.ComputerName
    
    With objValeExterno
        .inicializarEntidades
        
        .CodigoAlmacen = Trim(strCodigoAlmacen)
        .NumeroVale = Trim(strNumeroVale)
        
        If .verificarExistencia Then
            .obtenerConfigVale
            
            Select Case .TipoVale
                Case "I"
                    strTablaSQL = "INGRESO"
                Case "S"
                    strTablaSQL = "SALIDA"
            End Select
            
            If Not .ExportarVale Then
                MsgBox "Vale no exportardo; no cuenta con marca de Exportación de registro." & vbNewLine & vbNewLine & _
                        "Consulte con su administrador de sistemas.", vbInformation + vbOKOnly, App.ProductName & " - " & wnomcia
                
                exportarValeAserverSQLv2 = False
        
                Exit Function
            End If
            
'            strIdPuntoVenta = ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "IDPUNTOVENTA", "PUNTOVENTA", "NOMBREPC", strNombrePC, "T", "AND ELIMINADO = 0")
            
'            If Trim(strIdPuntoVenta) = vbNullString Then
'                MsgBox "Vale no exportado. Su computador no esta habilitado para generar registros de " & strTablaSQL & "." & vbNewLine & vbNewLine & _
'                                "Consulte con su administrador de sistemas.", vbInformation + vbOKOnly, App.ProductName & " - " & wnomcia
'
'                exportarValeAserverSQLv2 = False
'
'                Exit Function
'            End If
            
            dblIdVale = Val(.NumeroValeExterno)
            strIdTipoVale = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F1CODORIEXTERNO", "SF1ORIGENES", "F1CODORI", .CodigoOrigen, "T")
            
            Select Case .TipoPersona
                Case "P"
                    strTipoComprobanteProvDefault = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2TIPDOC", "EF2PROVEEDORES", "F2CODPROV", .CodigoProveedor, "T")
                    
                    Select Case strTipoComprobanteProvDefault
                        Case "86"
                            If Trim(.NumeroGuia) <> vbNullString Then
                                strIdDocumento = "DOC0000006"
                                
                                strNumDoc = IIf(.SerieGuia <> vbNullString, Format(.SerieGuia, "000") & "-", vbNullString) & Format(.NumeroGuia, "0000000")
                            ElseIf Trim(.NumeroDocumento) <> vbNullString Then
                                strIdDocumento = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CODEXT3", "DOCUMENTOS", "F2CODDOC", .CodTipoComprobante, "T")
                                
                                strNumDoc = IIf(.SerieDocumento <> vbNullString, Format(.SerieDocumento, "000") & "-", vbNullString) & Format(.NumeroDocumento, "0000000")
                            Else
                                strIdDocumento = "DOC0000004"
                                strNumDoc = vbNullString
                            End If
                        Case Else
                            If Trim(.NumeroDocumento) <> vbNullString Then
                                strIdDocumento = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CODEXT3", "DOCUMENTOS", "F2CODDOC", .CodTipoComprobante, "T")
                                
                                strNumDoc = IIf(.SerieDocumento <> vbNullString, Format(.SerieDocumento, "000") & "-", vbNullString) & Format(.NumeroDocumento, "0000000")
                            ElseIf Trim(.NumeroGuia) <> vbNullString Then
                                strIdDocumento = "DOC0000006"
                                
                                strNumDoc = IIf(.SerieGuia <> vbNullString, Format(.SerieGuia, "000") & "-", vbNullString) & Format(.NumeroGuia, "0000000")
                            Else
                                strIdDocumento = "DOC0000004"
                                strNumDoc = vbNullString
                            End If
                    End Select
                Case Else
                    If Trim(.NumeroGuia) <> vbNullString Then
                        strIdDocumento = "DOC0000006"
                    Else
                        If Trim(strIdDocumento) = vbNullString Then
                            strIdDocumento = "DOC0000004"
                        End If
                    End If
                    
                    Select Case strIdDocumento
                        Case "DOC0000006"
                            strNumDoc = IIf(.SerieGuia <> vbNullString, Format(.SerieGuia, "000") & "-", vbNullString) & Format(.NumeroGuia, "0000000")
                        Case "DOC0000004"
                            strNumDoc = vbNullString
                        Case Else
                            strNumDoc = IIf(.SerieDocumento <> vbNullString, Format(.SerieDocumento, "000") & "-", vbNullString) & Format(.NumeroDocumento, "0000000")
                    End Select
            End Select
            
            Select Case .TipoVale
                Case "I"
                    StrFecha = Format(.Fecha, "dd-mm-yy 00:00:00")
                Case "S"
                    StrFecha = Format(.Fecha, "dd-mm-yy 12:00:00")
            End Select
            
            dblTCambio = .TipoCambio
            
            Select Case .CodigoMoneda
                Case "S"
                    strIdMoneda = "MON0000001"
                    
                    dblSubTotal = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "FORMAT(SUM(F3CANPRO * F3VALVTA),'#0.00') AS SUMA", "IF3VALES", "F2CODALM", .CodigoAlmacen, "T", "AND F4NUMVAL = '" & .NumeroVale & "'"))
                    dblIgv = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "FORMAT(SUM(F3IGV),'#0.00') AS SUMA", "IF3VALES", "F2CODALM", .CodigoAlmacen, "T", "AND F4NUMVAL = '" & .NumeroVale & "'"))
                    dblTotal = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "FORMAT(SUM(F3TOTITE),'#0.00') AS SUMA", "IF3VALES", "F2CODALM", .CodigoAlmacen, "T", "AND F4NUMVAL = '" & .NumeroVale & "'"))
                Case "D"
                    strIdMoneda = "MON0000002"
                    
                    dblSubTotal = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "FORMAT(SUM(F3CANPRO * F3VALDOL),'#0.00') AS SUMA", "IF3VALES", "F2CODALM", .CodigoAlmacen, "T", "AND F4NUMVAL = '" & .NumeroVale & "'"))
                    dblIgv = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "FORMAT(SUM(F3IGVDOL),'#0.00') AS SUMA", "IF3VALES", "F2CODALM", .CodigoAlmacen, "T", "AND F4NUMVAL = '" & .NumeroVale & "'"))
                    dblTotal = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "FORMAT(SUM(F3TOTDOL),'#0.00') AS SUMA", "IF3VALES", "F2CODALM", .CodigoAlmacen, "T", "AND F4NUMVAL = '" & .NumeroVale & "'"))
            End Select
            
            strIdAlmacen = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODALMEXTERNO", "EF2ALMACENES", "F2CODALM", .CodigoAlmacen, "T")
            
            Select Case .TipoPersona
                Case "C"
                    strIdPersona = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODCLIEXTERNO", "EF2CLIENTES", "F2CODCLI", .CodigoProveedor, "T")
                Case "P"
                    strIdPersona = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODPROVEXTERNO", "EF2PROVEEDORES", "F2CODPROV", .CodigoProveedor, "T")
                Case Else
                    strIdPersona = vbNullString
            End Select
            
            dblPorcIGV = Val(wIgv)
            dblMontoDscto = 0
            dblPorcDscto = 0
            strObservacion = .observaciones
            strIdUsuarioIngreso = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODUSEREXTERNO", "EF2USERS", "F2CODUSER", .UsuReg, "T")
            strIdUsuarioActualizacion = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODUSEREXTERNO", "EF2USERS", "F2CODUSER", .UsuMod, "T")
            strFechaIngreso = Format(.FecReg, "dd-mm-yy hh:mm:ss")
            strFechaActualizacion = Format(.FecMod, "dd-mm-yy hh:mm:ss")
            dblIdOrdenProduccion = Val(.OrdenTrabajo)
            
            SqlCad = vbNullString
            
            'GUARDAR CABECERA DE VALE
            If ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "ID" & strTablaSQL, strTablaSQL, "ID" & strTablaSQL, Val(.NumeroValeExterno), "N") = vbNullString Then
                bolNuevoRegistro = True
                
                'GENERAR ID
                dblIdVale = Val(ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "CORRELATIVO", "CORRELATIVOTABLA", "IDPUNTOVENTA", strIdPuntoVenta, "T", _
                                                                    "AND TABLA = '" & strTablaSQL & "'"))
                
                If dblIdVale = 0 Then
                    MsgBox "Vale no exportado. Su computador no esta habilitado para generar registros de [" & strTablaSQL & "]." & vbNewLine & vbNewLine & _
                            "Consulte con su administrador de sistemas.", vbInformation + vbOKOnly, App.ProductName & " - " & wnomcia
                    
                    .inicializarEntidades
                    
                    exportarValeAserverSQLv2 = False
                    
                    Exit Function
                End If
                
                'INSERTAR
                SqlCad = vbNullString
                SqlCad = SqlCad & "INSERT INTO " & strTablaSQL & "("
                SqlCad = SqlCad & "ID" & strTablaSQL & ", IDTIPO" & strTablaSQL & ", IDDOCUMENTO, "
                SqlCad = SqlCad & "[NUM-DOC], FECHA, IDMONEDA, "
                SqlCad = SqlCad & "IDALMACEN, IDPERSONA, SUBTOTAL, "
                SqlCad = SqlCad & "IGV, TOTAL, TCAMBIO, "
                SqlCad = SqlCad & "PORCIGV, "
                    
                    If .TipoVale = "I" Then
                        SqlCad = SqlCad & "MONTODESCUENTO , PORCDESCUENTO, "
                    End If
                    
                SqlCad = SqlCad & "OBSERVACIONDOCUMENTO, IDPUNTOVENTA, IDUSUARIOINGRESO, "
                SqlCad = SqlCad & "IDUSUARIOACTUALIZACION, FECHAINGRESO, FECHAACTUALIZACION, "
                SqlCad = SqlCad & "ANULADO, IDORDENPRODUCCION) "
                
                SqlCad = SqlCad & "VALUES("
                SqlCad = SqlCad & dblIdVale & ", '" & strIdTipoVale & "', '" & strIdDocumento & "', "
                SqlCad = SqlCad & "'" & strNumDoc & "', CONVERT(DATETIME,'" & StrFecha & "', 5), '" & strIdMoneda & "', "
                SqlCad = SqlCad & "'" & strIdAlmacen & "', " & IIf(strIdPersona <> vbNullString, "'" & strIdPersona & "'", "NULL") & ", " & dblSubTotal & ", "
                SqlCad = SqlCad & dblIgv & ", " & dblTotal & ", " & dblTCambio & ","
                SqlCad = SqlCad & dblPorcIGV & ", "
                    
                    If .TipoVale = "I" Then
                        SqlCad = SqlCad & dblMontoDscto & ", " & dblPorcDscto & ", "
                    End If
                    
                SqlCad = SqlCad & "'" & strObservacion & "', '" & strIdPuntoVenta & "', '" & strIdUsuarioIngreso & "', "
                SqlCad = SqlCad & "'" & strIdUsuarioIngreso & "', GETDATE(), GETDATE(), "
                SqlCad = SqlCad & "0, " & dblIdOrdenProduccion
                SqlCad = SqlCad & ")"
            Else
                bolNuevoRegistro = False
                
                'MODIFICAR
                SqlCad = vbNullString
                SqlCad = SqlCad & "UPDATE "
                SqlCad = SqlCad & strTablaSQL & " "
                SqlCad = SqlCad & "SET "
                SqlCad = SqlCad & "IDTIPO" & strTablaSQL & " = '" & strIdTipoVale & "', "
                SqlCad = SqlCad & "IDDOCUMENTO = '" & strIdDocumento & "', "
                SqlCad = SqlCad & "[NUM-DOC] = '" & strNumDoc & "', "
                SqlCad = SqlCad & "FECHA = CONVERT(DATETIME, '" & StrFecha & "', 5), "
                SqlCad = SqlCad & "IDMONEDA = '" & strIdMoneda & "', "
                SqlCad = SqlCad & "IDALMACEN = '" & strIdAlmacen & "', "
                SqlCad = SqlCad & "IDPERSONA = " & IIf(strIdPersona <> vbNullString, "'" & strIdPersona & "'", "NULL") & ", "
                SqlCad = SqlCad & "SUBTOTAL = " & dblSubTotal & ", "
                SqlCad = SqlCad & "IGV = " & dblIgv & ", "
                SqlCad = SqlCad & "TOTAL = " & dblTotal & ", "
                SqlCad = SqlCad & "TCAMBIO = " & dblTCambio & ", "
                SqlCad = SqlCad & "PORCIGV = " & dblPorcIGV & ", "
                    
                    If .TipoVale = "I" Then
                        SqlCad = SqlCad & "MONTODESCUENTO = " & dblMontoDscto & ", "
                        SqlCad = SqlCad & "PORCDESCUENTO = " & dblPorcDscto & ", "
                    End If
                    
                SqlCad = SqlCad & "OBSERVACIONDOCUMENTO = '" & strObservacion & "', "
                SqlCad = SqlCad & "IDPUNTOVENTA = '" & strIdPuntoVenta & "', "
                SqlCad = SqlCad & "IDUSUARIOACTUALIZACION = '" & strIdUsuarioActualizacion & "', "
                SqlCad = SqlCad & "FECHAACTUALIZACION = GETDATE(), "
                SqlCad = SqlCad & "ANULADO = 0, "
                SqlCad = SqlCad & "IDORDENPRODUCCION = " & dblIdOrdenProduccion & " "
                SqlCad = SqlCad & "WHERE "
                SqlCad = SqlCad & "ID" & strTablaSQL & " = " & dblIdVale
            End If
            
            cnBdStudioModa.Execute SqlCad
            
            Actualiza_Log "< DB Externo > " & SqlCad, StrConexDbBancos
            
            If bolNuevoRegistro Then
                'ACTUALIZAR CORRELATIVO
                SqlCad = vbNullString
                SqlCad = SqlCad & "UPDATE "
                SqlCad = SqlCad & "CORRELATIVOTABLA "
                SqlCad = SqlCad & "SET "
                SqlCad = SqlCad & "CORRELATIVO = " & dblIdVale + 1 & " "
                SqlCad = SqlCad & "WHERE "
                SqlCad = SqlCad & "IDPUNTOVENTA = '" & strIdPuntoVenta & "' AND "
                SqlCad = SqlCad & "TABLA = '" & strTablaSQL & "'"
                
                cnBdStudioModa.Execute SqlCad
                
                Actualiza_Log " < DB Externo > " & SqlCad, StrConexDbBancos
                
                'ACTUALIZAR EN CABECERA DE VALE, EL ID EXTERNO
                SqlCad = vbNullString
                SqlCad = SqlCad & "UPDATE "
                SqlCad = SqlCad & "IF4VALES "
                SqlCad = SqlCad & "SET "
                SqlCad = SqlCad & "NUMENSAM = '" & dblIdVale & "' "
                SqlCad = SqlCad & "WHERE "
                SqlCad = SqlCad & "F2CODALM = '" & .CodigoAlmacen & "' AND "
                SqlCad = SqlCad & "F4NUMVAL = '" & .NumeroVale & "'"
                
                cnn_dbbancos.Execute SqlCad
                
                Actualiza_Log SqlCad, StrConexDbBancos
                
                etiquetaIDExterno.Caption = Trim(dblIdVale & "")
            End If
            
            'GUARDAR DETALLE DEL VALE
            SqlCad = vbNullString
            SqlCad = SqlCad & "SELECT "
            SqlCad = SqlCad & "F5CODPRO AS CODPRODUCTO, "
            SqlCad = SqlCad & "SUM(F3CANPRO) AS CANTIDAD, "
            SqlCad = SqlCad & "VAL(FORMAT(F3VALVTA, '#0.0000')) AS VALSOLES, "
            SqlCad = SqlCad & "VAL(FORMAT(F3VALDOL, '#0.0000')) AS VALDOL, "
            SqlCad = SqlCad & "SUM(F3MONTODSCTO) AS DSCTO, "
            SqlCad = SqlCad & "F3PORCENTAJEDSCTO AS PORCDSCTO "
            SqlCad = SqlCad & "FROM "
            SqlCad = SqlCad & "IF3VALES "
            SqlCad = SqlCad & "WHERE "
            SqlCad = SqlCad & "F2CODALM = '" & .CodigoAlmacen & "' AND "
            SqlCad = SqlCad & "F4NUMVAL = '" & .NumeroVale & "' "
            SqlCad = SqlCad & "GROUP BY "
            SqlCad = SqlCad & "F5CODPRO, "
            SqlCad = SqlCad & "F3VALVTA, "
            SqlCad = SqlCad & "F3VALDOL, "
            SqlCad = SqlCad & "F3PORCENTAJEDSCTO "
            
            If rstValeDet.State = 1 Then rstValeDet.Close
            
            rstValeDet.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
            
            If Not rstValeDet.EOF Then
                rstValeDet.MoveFirst
                
                'If Not bolNuevoRegistro Then
                    SqlCad = vbNullString
                    SqlCad = "DELETE FROM " & strTablaSQL & "DETALLE WHERE ID" & strTablaSQL & " = " & dblIdVale
                    
                    cnBdStudioModa.Execute SqlCad
                    
                    Actualiza_Log "< DB Externo > " & SqlCad, StrConexDbBancos
                'End If
                
                dblItem = 0
                
                frameProgreso.Visible = True
                barraProgreso.Max = ModUtilitario.devuelveCantRegistros(rstValeDet)
                barraProgreso.Value = 0
                frameProgreso.Caption = "Exportando Detalle..."
                
                Do While Not rstValeDet.EOF
                    .inicializarEntidadesDetalle
                    
                    dblItem = dblItem + 1
                    
                    SqlCad = vbNullString
                    SqlCad = SqlCad & "INSERT INTO " & strTablaSQL & "DETALLE("
                    SqlCad = SqlCad & "ID" & strTablaSQL & ", [ITEMS], IDINSUMO, "
                    SqlCad = SqlCad & "CANTIDAD, COSTO, IMPORTE, COSTOPROMEDIO"
                    
                        If .TipoVale = "I" Then
                            SqlCad = SqlCad & ", MONTODESCUENTO, PORCDESCUENTO"
                        End If
                        
                    SqlCad = SqlCad & ") "
                    SqlCad = SqlCad & "VALUES("
                    SqlCad = SqlCad & dblIdVale & ", " & dblItem & ", '" & Trim(rstValeDet!CodProducto & "") & "', "
                    SqlCad = SqlCad & Val(rstValeDet!Cantidad & "") & ", "
                    
                    .CodigoProducto = Trim(rstValeDet!CodProducto & "")
                    
                    Rem SK: COMENTADO TEMPORALMENTE
                    'dblCostoPromedio = .calcularCostoPromedioV3conDAO
                    dblCostoPromedio = 0
                    
                    Select Case .CodigoMoneda
                        Case "S"
                            SqlCad = SqlCad & Val(rstValeDet!VALSOLES & "") & ", " & Val(rstValeDet!VALSOLES & "") * Val(rstValeDet!Cantidad & "") & ", "
                            SqlCad = SqlCad & dblCostoPromedio
                        Case "D"
                            SqlCad = SqlCad & Val(rstValeDet!VALDOL & "") & ", " & Val(rstValeDet!VALDOL & "") * Val(rstValeDet!Cantidad & "") & ", "
                            SqlCad = SqlCad & dblCostoPromedio
                    End Select
                    
                    If .TipoVale = "I" Then
                        SqlCad = SqlCad & ", " & Val(rstValeDet!DSCTO & "") & ", "
                        SqlCad = SqlCad & Val(rstValeDet!PORCDSCTO & "")
                    End If
                    
                    SqlCad = SqlCad & ")"
                    
                    cnBdStudioModa.Execute SqlCad
                    
                    Actualiza_Log "< DB Externo > " & SqlCad, StrConexDbBancos
                    
                    Rem SK: COMENTADO TEMPORALMENTE
'                    'ACTUALIZAR TABLA INSUMOSTOCK
'                    SqlCad = vbNullString
'                    SqlCad = SqlCad & "UPDATE "
'                    SqlCad = SqlCad & "INSUMOSTOCK "
'                    SqlCad = SqlCad & "SET "
'
'                        'dblStockActual = .devuelveStockFisicoDeProducto("C", True) + .devuelveStockFisicoDeProducto("L", True)
'
'                        dblStockActual = .devuelveStockFisicoDeProductoConADO("C", True) + .devuelveStockFisicoDeProductoConADO("L", True)
'
'                        Rem PAUSA
'                        'MsgBox "NOVENA PAUSA: STOCK CALCULADO DE ITEM DE " & strTablaSQL & " EN CP.", vbOKOnly + vbInformation, App.ProductName & "PROCESO DE EXPORTACION DE VALE A INTEGRADO"
'
'                        SqlCad = SqlCad & "STKACTUAL = " & dblStockActual & ", "
'                        SqlCad = SqlCad & "STKCOMPRA = " & dblStockActual & " "
'
'                    SqlCad = SqlCad & "WHERE "
'                    SqlCad = SqlCad & "LTRIM(RTRIM(IDINSUMO)) = '" & Trim(rstValeDet!CodProducto & "") & "' AND "
'                    SqlCad = SqlCad & "LTRIM(RTRIM(IDALMACEN)) = '" & strIdAlmacen & "'"
'
'                    cnBdStudioModa.Execute SqlCad
'
'                    Actualiza_Log "< DB Externo > " & SqlCad, StrConexDbBancos
                    
                    'ACTUALIZAR TABLA INSUMOPERSONA
                    If strIdPersona <> vbNullString Then
                        If Val(ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "COUNT(*)", "PERSONA", "IDPERSONA", strIdPersona, "T")) > 0 Then
                            If Val(ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "COUNT(*)", "INSUMOPERSONA", "IDPERSONA", strIdPersona, "T", "AND IDINSUMO = '" & Trim(rstValeDet!CodProducto & "") & "'")) = 0 Then
                                SqlCad = vbNullString
                                SqlCad = SqlCad & "INSERT INTO INSUMOPERSONA("
                                SqlCad = SqlCad & "IDINSUMO, ITEMS, IDPERSONA, IDMONEDA, COSTO, COSTOPROMEDIO, "
                                SqlCad = SqlCad & "IDUSUARIOINGRESO,IDUSUARIOACTUALIZACION,FECHAINGRESO,FECHAACTUALIZACION) "
                                SqlCad = SqlCad & "VALUES("
                                SqlCad = SqlCad & "'" & Trim(rstValeDet!CodProducto & "") & "', "
                                SqlCad = SqlCad & Val(ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "COUNT(*)", "INSUMOPERSONA", "IDINSUMO", Trim(rstValeDet!CodProducto & ""), "T")) + 1 & ", "
                                SqlCad = SqlCad & "'" & strIdPersona & "', "
                                SqlCad = SqlCad & "'" & strIdMoneda & "', "
                                SqlCad = SqlCad & Val(rstValeDet!VALSOLES & "") & ", "
                                SqlCad = SqlCad & dblCostoPromedio & ", "
                                SqlCad = SqlCad & "'" & strIdUsuarioIngreso & "', "
                                SqlCad = SqlCad & "'" & IIf(strIdUsuarioActualizacion <> vbNullString, strIdUsuarioActualizacion, strIdUsuarioIngreso) & "', "
                                SqlCad = SqlCad & "GETDATE(), "
                                SqlCad = SqlCad & "GETDATE())"
                            Else
                                SqlCad = vbNullString
                                SqlCad = SqlCad & "UPDATE "
                                SqlCad = SqlCad & "INSUMOPERSONA "
                                SqlCad = SqlCad & "SET "
                                SqlCad = SqlCad & "IDMONEDA = '" & strIdMoneda & "', "
                                SqlCad = SqlCad & "COSTO = " & Val(rstValeDet!VALSOLES & "") & ", "
                                SqlCad = SqlCad & "IDUSUARIOACTUALIZACION = '" & IIf(strIdUsuarioActualizacion <> vbNullString, strIdUsuarioActualizacion, strIdUsuarioIngreso) & "', "
                                SqlCad = SqlCad & "FECHAACTUALIZACION = GETDATE() "
                                SqlCad = SqlCad & "WHERE "
                                SqlCad = SqlCad & "IDPERSONA = '" & strIdPersona & "' AND "
                                SqlCad = SqlCad & "IDINSUMO = '" & Trim(rstValeDet!CodProducto & "") & "'"
                            End If
                            
                            cnBdStudioModa.Execute SqlCad
                            
                            Actualiza_Log "< DB Externo > " & SqlCad, StrConexDbBancos
                        End If
                    End If
                    
                    'VERIFICAR SI YA SE REALIZO EL CIERRE
                    If Val(ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "COUNT(*)", "CIERREMENSUAL", "IDALMACEN", strIdAlmacen, "T", "AND (AÑO = " & Year(CDate(StrFecha)) & " AND MES = " & Month(CDate(StrFecha)) & ")")) > 0 Then
                        Dim cmdSpActualizarCPIngreso As ADODB.Command
                        
                        Set cmdSpActualizarCPIngreso = New ADODB.Command
                        
                        With cmdSpActualizarCPIngreso
                            .ActiveConnection = cnBdStudioModa
                            .CommandType = adCmdStoredProc
                            .CommandText = "USP_ActualizaCPIngreso"
                            
                            .Parameters.Append .CreateParameter("@IdPuntoVenta", adVarChar, adParamInput, 10, strIdPuntoVenta)
                            .Parameters.Append .CreateParameter("@IdUsuario", adVarChar, adParamInput, 10, IIf(strIdUsuarioActualizacion <> vbNullString, strIdUsuarioActualizacion, strIdUsuarioIngreso))
                            .Parameters.Append .CreateParameter("@IdAlmacen", adVarChar, adParamInput, 10, strIdAlmacen)
                            .Parameters.Append .CreateParameter("@FechaActualiza", adDBTimeStamp, adParamInput, 10, IIf(strFechaActualizacion <> vbNullString, strFechaActualizacion, strFechaIngreso))
                            .Parameters.Append .CreateParameter("@IdInsumo", adVarChar, adParamInput, 50, Trim(rstValeDet!CodProducto & ""))
                            
                            .Execute
                        End With
                        
                        Set cmdSpActualizarCPIngreso = Nothing
                    End If
                    
                    DoEvents
                    
                    barraProgreso.Value = barraProgreso.Value + 1
                    frameProgreso.Caption = "Exportando Detalle... " & FormatPercent(barraProgreso.Value / barraProgreso.Max, 3)
                    
                    rstValeDet.MoveNext
                Loop
            End If
            
            .inicializarEntidades
            .inicializarEntidadesDetalle
            
            exportarValeAserverSQLv2 = True
        Else
            MsgBox "El Vale no se guardo correctamente, vuelva a intentarlo.", vbInformation + vbOKOnly, App.ProductName
            
            exportarValeAserverSQLv2 = False
            
            .inicializarEntidades
            .inicializarEntidadesDetalle
        End If
    End With
    
    If rstValeDet.State = 1 Then rstValeDet.Close
    
    Set cmdSpActualizarCPIngreso = Nothing
    Set rstValeDet = Nothing
    
    strTablaSQL = vbNullString
    strNombrePC = vbNullString
    dblItem = 0
    strTipoComprobanteProvDefault = vbNullString
    bolNuevoRegistro = False
    
    'Variables para Ingreso/Salida
    dblIdVale = 0
    strIdTipoVale = vbNullString
    strIdDocumento = vbNullString
    strNumDoc = vbNullString
    StrFecha = vbNullString
    strIdMoneda = vbNullString
    strIdAlmacen = vbNullString
    strIdPersona = vbNullString
    dblSubTotal = 0
    dblIgv = 0
    dblTotal = 0
    dblTCambio = 0
    dblPorcIGV = 0
    dblMontoDscto = 0
    dblPorcDscto = 0
    strObservacion = vbNullString
    strObservacionAnulacion = vbNullString
    strIdPuntoVenta = vbNullString
    strIdUsuarioIngreso = vbNullString
    strIdUsuarioActualizacion = vbNullString
    strFechaIngreso = vbNullString
    strFechaActualizacion = vbNullString
    intAnulado = 0
    dblIdOrdenProduccion = 0
    dblIdOrdenCompra = 0
    
    dblCostoPromedio = 0
    dblStockActual = 0
    
    frameProgreso.Visible = False
    
    Set objValeExterno = Nothing
    
    Exit Function
    Resume
errExportarValeAserverSQLv2:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description & vbNewLine & vbNewLine & _
            "Vale no exportado ó exportado con problemas, vuelva a guardar el Vale para completar con la exportación correctamente.", vbInformation + vbOKOnly, App.ProductName & " - ExportarValeAserverSQLv2"
    
    Actualiza_Log "Exportación de Vale [Numero: " & objValeExterno.CodigoAlmacen & " / " & objValeExterno.NumeroVale & "] [ID Externo: " & etiquetaIDExterno.Caption & "] cancelada por el siguiente error: " & _
                    "[Numero Error: " & Err.Number & "] [Descripción: " & Err.Description & "]", StrConexDbBancos
    
    exportarValeAserverSQLv2 = False
    
    frameProgreso.Visible = False
    
    Err.Clear
End Function

Public Function exportarProveedorAserverSQL(ByVal strCodigo As String) As Boolean
    
    On Error GoTo errExportarProveedorAserverSQL
    
    Dim cmdSP As New ADODB.Command
    
    'Variables para Proveedor
    Dim strIdPersona As String
    Dim strIdTipoPersona As String
    Dim strNombre As String
    Dim strRuc As String
    Dim strDireccion As String
    Dim strDNI As String
    Dim strPasaporte As String
    Dim strIDPostal As String
    Dim strTelefono1 As String
    Dim strFax As String
    Dim strEmail As String
    Dim strWeb As String
    Dim bolExtranjero As Boolean
    
    Dim strIdUsuarioIngreso As String
    Dim strIdUsuarioActualizacion As String
    Dim strFechaIngreso As String
    Dim strFechaActualizacion As String
    
    Dim strIDArea As String
    Dim strIDSeccion As String
    Dim dblDescuento As Double
    
    'abrirCnDBMilano
    
    With objAyudaProveedor
        .inicializarEntidades
        
        .Codigo = Trim(strCodigo)
        
        If .verificarExistencia Then
            .obtenerConfigProveedor
            
            strIdPersona = .CodigoExterno
            strIdTipoPersona = "TIP0000001"
            strNombre = .NombreProveedor
            strRuc = .NumeroDocumento
            strDireccion = .DireccionProveedor
            strDNI = vbNullString
            strPasaporte = vbNullString
            strIDPostal = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODUSEREXTERNO1", "EF2ZONAS", "F2CODZON", .CodigoDistrito, "T")
            strTelefono1 = .Telefono
            strFax = .Fax
            strEmail = .Email
            strWeb = .CuentaAbono
            bolExtranjero = IIf(.OrigenProveedor = "E", True, False)
            
            strIdUsuarioIngreso = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODUSEREXTERNO", "EF2USERS", "F2CODUSER", IIf(.UsuarioReg <> vbNullString, .UsuarioReg, wusuario), "T")
            strIdUsuarioActualizacion = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODUSEREXTERNO", "EF2USERS", "F2CODUSER", IIf(.UsuarioMod <> vbNullString, .UsuarioMod, wusuario), "T")
            strFechaIngreso = Format(IIf(.FechaReg <> vbNullString, .FechaReg, Date), "dd-mm-yy hh:mm:ss")
            strFechaActualizacion = Format(IIf(.FechaMod <> vbNullString, .FechaMod, Date), "dd-mm-yy hh:mm:ss")
            
            strIDArea = vbNullString
            strIDSeccion = vbNullString
            dblDescuento = 0
            
            'GUARDAR PROVEEDOR
            If ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "IDPERSONA", "PERSONA", "IDPERSONA", .CodigoExterno, "T") = vbNullString Then
                'GENERAR ID
                strIdPersona = Trim(ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "TOP 1 IDPERSONA", "PERSONA", vbNullString, vbNullString, vbNullString, _
                                                                    "LTRIM(RTRIM(IDPERSONA)) <> '' ORDER BY IDPERSONA DESC"))
                
                strIdPersona = "PER" & Format(Val(Mid(strIdPersona, 4)) + 1, "0000000")
                
                'INSERTAR
                SqlCad = vbNullString
                SqlCad = SqlCad & "INSERT INTO PERSONA("
                SqlCad = SqlCad & "IDPERSONA, IDTIPOPERSONA, NOMBRE, "
                SqlCad = SqlCad & "RUC, DIRECCION, DNI, PASAPORTE, IDPOSTAL, "
                SqlCad = SqlCad & "TELEFONO1, TELEFONO2, CELULAR, FAX, "
                SqlCad = SqlCad & "EMAIL, WEB, EXTRANJERO, "
                SqlCad = SqlCad & "IDUSUARIOINGRESO, "
                SqlCad = SqlCad & "IDUSUARIOACTUALIZACION, "
                SqlCad = SqlCad & "FECHAINGRESO, "
                SqlCad = SqlCad & "FECHAACTUALIZACION, "
                SqlCad = SqlCad & "ELIMINADO, IDAREA, "
                SqlCad = SqlCad & "IDSECCION, DESCUENTO"
                SqlCad = SqlCad & ") "
                
                SqlCad = SqlCad & "VALUES("
                SqlCad = SqlCad & "'" & strIdPersona & "', '" & strIdTipoPersona & "', '" & strNombre & "', "
                SqlCad = SqlCad & "'" & strRuc & "', '" & Replace(strDireccion, "'", " ", 1) & "', '', '', '" & strIDPostal & "', '" & strTelefono1 & "', '', '', "
                SqlCad = SqlCad & "'" & strFax & "', '" & strEmail & "', '" & strWeb & "', "
                SqlCad = SqlCad & IIf(bolExtranjero, 1, 0) & ", "
                SqlCad = SqlCad & "'" & strIdUsuarioIngreso & "', "
                SqlCad = SqlCad & "'" & strIdUsuarioActualizacion & "', "
                SqlCad = SqlCad & "GETDATE(), "
                SqlCad = SqlCad & "GETDATE(), "
                SqlCad = SqlCad & "0, '', '', 0"
                SqlCad = SqlCad & ")"
            Else
                'MODIFICAR
                SqlCad = vbNullString
                SqlCad = SqlCad & "UPDATE "
                SqlCad = SqlCad & "PERSONA "
                SqlCad = SqlCad & "SET "
                'SqlCad = SqlCad & "IDTIPOPERSONA = '" & strIdTipoPersona & "', "
                SqlCad = SqlCad & "NOMBRE = '" & strNombre & "', "
                SqlCad = SqlCad & "RUC = '" & strRuc & "', "
                SqlCad = SqlCad & "DIRECCION = '" & Replace(strDireccion, "'", " ", 1) & "', "
                SqlCad = SqlCad & "IDPOSTAL = '" & strIDPostal & "', "
                SqlCad = SqlCad & "TELEFONO1 = '" & strTelefono1 & "', "
                SqlCad = SqlCad & "FAX = '" & strFax & "', "
                SqlCad = SqlCad & "EMAIL = '" & strEmail & "', "
                SqlCad = SqlCad & "WEB = '" & strWeb & "', "
                SqlCad = SqlCad & "EXTRANJERO = " & IIf(bolExtranjero, 1, 0) & ", "
                SqlCad = SqlCad & "IDUSUARIOACTUALIZACION = '" & strIdUsuarioActualizacion & "', "
                SqlCad = SqlCad & "FECHAACTUALIZACION = GETDATE() "
                SqlCad = SqlCad & "WHERE "
                SqlCad = SqlCad & "IDPERSONA = '" & strIdPersona & "'"
            End If
            
            cnBdStudioModa.Execute SqlCad
            
            Actualiza_Log "< DB Externo > " & SqlCad, StrConexDbBancos
            
            If .CodigoExterno = vbNullString Then
                SqlCad = vbNullString
                SqlCad = SqlCad & "UPDATE "
                SqlCad = SqlCad & "EF2PROVEEDORES "
                SqlCad = SqlCad & "SET "
                SqlCad = SqlCad & "F2CODPROVEXTERNO = '" & strIdPersona & "' "
                SqlCad = SqlCad & "WHERE "
                SqlCad = SqlCad & "F2CODPROV = '" & .Codigo & "'"
                
                cnn_dbbancos.Execute SqlCad
                
                Actualiza_Log SqlCad, StrConexDbBancos
            End If
            
            MsgBox "Proveedor EXPORTADO correctamente.", vbInformation + vbOKOnly, App.ProductName
            
            exportarProveedorAserverSQL = True
        Else
            MsgBox "Proveedor no se EXPORTO correctamente, vuelva a intentarlo.", vbInformation + vbOKOnly, App.ProductName
            
            exportarProveedorAserverSQL = False
            
            .inicializarEntidades
        End If
    End With
    
    Exit Function
    Resume
errExportarProveedorAserverSQL:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description & vbNewLine & vbNewLine & _
            "Proveedor no EXPORTADO, intente guardar nuevamente.", vbInformation + vbOKOnly, App.ProductName & " - ExportarProveedorAserverSQL"
    
    Actualiza_Log "Exportación de Proveedor [" & objAyudaProveedor.Codigo & "] cancelada por el siguiente error: " & _
                        "[Numero Error: " & Err.Number & "] [Descripción: " & Err.Description & "]", StrConexDbBancos
    
    Err.Clear
End Function

Public Sub exportarTomaInventarioServidorExterno(ByVal frameProgreso As Object, _
                                                            ByVal barraProgreso As Object)
    
    On Error GoTo errExportarTomaInventarioServidorExterno
    
    Dim rstTomaInventario As ADODB.Recordset
    Dim dblRegistro As Double
    Dim dblCantRegistroActualizado As Double
    Dim strNombrePC As String
    Dim strIdPuntoVenta As String
    
    Set rstTomaInventario = Nothing
    
    Set rstTomaInventario = New ADODB.Recordset
    
    Rem LECTURA DE  TOMA DE INVENTARIOS
    
    
    If MsgBox("¿Desea exportar la Toma de Inventario: 2014-12?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
        Exit Sub
    End If
    
    'abrirCnDBMilano
    
    strNombrePC = modgeneral.ComputerName
    strIdPuntoVenta = ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "IDPUNTOVENTA", "PUNTOVENTA", "NOMBREPC", strNombrePC, "T", "AND ELIMINADO = 0")
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "ALM.F2CODALMEXTERNO, "
    SqlCad = SqlCad & "TI.F4ANNO, "
    SqlCad = SqlCad & "TI.F4MES, "
    SqlCad = SqlCad & "TI.F5CODPRO, "
    SqlCad = SqlCad & "TI.F3STOCKSISTEMA, "
    SqlCad = SqlCad & "TI.F3STOCKFISICO, "
    SqlCad = SqlCad & "TI.F3DIFERENCIA "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "(H3TOMAINV AS TI "
    SqlCad = SqlCad & "LEFT JOIN IF5PLA AS PROD ON PROD.F5CODPRO = TI.F5CODPRO) "
    SqlCad = SqlCad & "LEFT JOIN EF2ALMACENES AS ALM ON ALM.F2CODALM = TI.F2CODALM "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "TI.F4ANNO = '2014' AND "
    SqlCad = SqlCad & "TI.F4MES = '12' AND "
    SqlCad = SqlCad & "TI.F2CODALM = '01' "
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "PROD.F5NOMPRO"
    
    If rstTomaInventario.State = 1 Then rstTomaInventario.Close
    
    rstTomaInventario.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstTomaInventario.EOF Then
        frameProgreso.Visible = True
        barraProgreso.Max = ModUtilitario.devuelveCantRegistros(rstTomaInventario)
        barraProgreso.Value = 0
        frameProgreso.Caption = "Exportando Toma de Inventarios..."
        
        dblCantRegistroActualizado = 0
        
        Do While Not rstTomaInventario.EOF
            
            If ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "IDINSUMO", "TOMAINVENTARIO", "IDALMACEN", Trim(rstTomaInventario!F2CODALMEXTERNO & ""), "T", "AND [AÑO] = " & Val(rstTomaInventario!F4ANNO & "") & " AND MES = " & Val(rstTomaInventario!F4MES & "") & " AND IDINSUMO = '" & Trim(rstTomaInventario!f5codpro & "") & "'") = vbNullString Then
                SqlCad = vbNullString
                SqlCad = SqlCad & "INSERT INTO TOMAINVENTARIO(IDALMACEN, AÑO, MES, IDINSUMO, TEORICO, FISICO, DIFERENCIA, "
                SqlCad = SqlCad & "ACTIVO, IDPUNTOVENTA, IDUSUARIOINGRESO, IDUSUARIOACTUALIZACION, "
                SqlCad = SqlCad & "FECHAINGRESO, FECHAACTUALIZACION, PROMEDIO) "
                SqlCad = SqlCad & "VALUES('" & Trim(rstTomaInventario!F2CODALMEXTERNO & "") & "', " & Val(rstTomaInventario!F4ANNO & "") & ", "
                SqlCad = SqlCad & Val(rstTomaInventario!F4MES & "") & ", '" & Trim(rstTomaInventario!f5codpro & "") & "', "
                SqlCad = SqlCad & Val(rstTomaInventario!F3STOCKSISTEMA & "") & ", " & Val(rstTomaInventario!F3STOCKFISICO & "") & ", "
                SqlCad = SqlCad & Val(rstTomaInventario!F3DIFERENCIA & "") & ", 1, '" & strIdPuntoVenta & "', "
                SqlCad = SqlCad & "'" & ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODUSEREXTERNO", "EF2USERS", "F2CODUSER", wusuario, "T") & "', "
                SqlCad = SqlCad & "'" & ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODUSEREXTERNO", "EF2USERS", "F2CODUSER", wusuario, "T") & "', "
                SqlCad = SqlCad & "GETDATE(), GETDATE(), NULL)"
            Else
                SqlCad = vbNullString
                SqlCad = SqlCad & "UPDATE "
                SqlCad = SqlCad & "TOMAINVENTARIO "
                SqlCad = SqlCad & "SET "
                SqlCad = SqlCad & "TEORICO = " & Val(rstTomaInventario!F3STOCKSISTEMA & "") & ","
                SqlCad = SqlCad & "FISICO = " & Val(rstTomaInventario!F3STOCKFISICO & "") & ", "
                SqlCad = SqlCad & "DIFERENCIA = " & Val(rstTomaInventario!F3DIFERENCIA & "") * -1 & ", "
                SqlCad = SqlCad & "IDUSUARIOACTUALIZACION = '" & ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODUSEREXTERNO", "EF2USERS", "F2CODUSER", wusuario, "T") & "',"
                SqlCad = SqlCad & "FECHAACTUALIZACION = GETDATE() "
                SqlCad = SqlCad & "WHERE "
                SqlCad = SqlCad & "IDALMACEN = '" & Trim(rstTomaInventario!F2CODALMEXTERNO & "") & "' AND "
                SqlCad = SqlCad & "[AÑO] = " & Val(rstTomaInventario!F4ANNO & "") & " AND "
                SqlCad = SqlCad & "MES = " & Val(rstTomaInventario!F4MES & "") & " AND "
                SqlCad = SqlCad & "IDINSUMO = '" & Trim(rstTomaInventario!f5codpro & "") & "'"
            End If
            
            dblRegistro = 0
            
            'cnn_dbbancos.Execute SqlCad, dblRegistro
            
            cnBdStudioModa.Execute SqlCad, dblRegistro
            
            If dblRegistro > 0 Then
                Actualiza_Log "< DB Externo > " & SqlCad, StrConexDbBancos
                
                dblCantRegistroActualizado = dblCantRegistroActualizado + 1
            End If
            
            DoEvents
            
            barraProgreso.Value = barraProgreso.Value + 1
            frameProgreso.Caption = "Exportando Toma de Inventarios... " & FormatPercent(barraProgreso.Value / barraProgreso.Max, 3)
            
            rstTomaInventario.MoveNext
        Loop
            MsgBox "Toma de Inventarios Fisico exportado." & vbNullString & _
                    "Registros Actualizados = " & dblCantRegistroActualizado & ".", vbInformation + vbOKOnly, wnomcia
    Else
        MsgBox "No se ubico Toma de Inventario para el Periodo = 2014-12.", vbInformation + vbOKOnly, wnomcia
    End If
    
    SqlCad = vbNullString
    
    frameProgreso.Visible = False
    
    If rstTomaInventario.State = 1 Then rstTomaInventario.Close
    
    Set rstTomaInventario = Nothing
    
    Exit Sub
errExportarTomaInventarioServidorExterno:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    'Resume
    frameProgreso.Visible = False
    barraProgreso.Value = 0
    
    Err.Clear
End Sub

'Public Function anularValeExterno(ByVal strTipoVale As String, _
'                                    ByVal strIDExterno As String, _
'                                    ByVal strIdAlmacen As String, _
'                                    ByVal frameProgreso As Object, _
'                                    ByVal barraProgreso As Object) As Boolean
'
'    On Error GoTo errAnularValeExterno
'
'    Dim cmdValeDet As New ADODB.Command
'    Dim rstValeDet As New ADODB.Recordset
'    Dim strTablaSQL As String
'
'    Dim strNombrePC As String
'    Dim strIdPuntoVenta As String
'    'Dim strIDAlmacen As String
'
'    Dim dblStockActual As Double
'
'    Select Case strTipoVale
'        Case "I"
'            strTablaSQL = "INGRESO"
'        Case "S"
'            strTablaSQL = "SALIDA"
'    End Select
'
'    strNombrePC = modgeneral.ComputerName
'
'    'abrirCnDBMilano
'
'    strIdPuntoVenta = ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "IDPUNTOVENTA", "PUNTOVENTA", "NOMBREPC", strNombrePC, "T", "AND ELIMINADO = 0")
'
'    If Trim(strIdPuntoVenta) = vbNullString Then
'        MsgBox "Vale Externo no puede ser Anulado. No se identifico el ID de Punto de Venta para modificar el registro de " & strTablaSQL & "." & vbNewLine & vbNewLine & _
'                        "Consulte con su administrador de sistemas.", vbInformation + vbOKOnly, App.ProductName & " - " & wnomcia
'
'        Exit Function
'    End If
'
'    If Val(ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "CORRELATIVO", "CORRELATIVOTABLA", "IDPUNTOVENTA", strIdPuntoVenta, "T", _
'                                                                    "AND TABLA = '" & strTablaSQL & "'")) = 0 Then
'
'        MsgBox "Vale Externo no puede ser Anulado. El Punto de Venta no cuenta con correlativo habilitado de " & strTablaSQL & "." & vbNewLine & vbNewLine & _
'                "Consulte con su administrador de sistemas.", vbInformation + vbOKOnly, App.ProductName & " - " & wnomcia
'
'        Exit Function
'    End If
'
'
'    'VERIFICAR SI DETALLE ESTA REGISTRADO, PARA REVERTIR STOCK ANTES DE ANULAR VALE EXTERNO
'    If rstValeDet.State = 1 Then rstValeDet.Close
'
'    SqlCad = vbNullString
'    SqlCad = SqlCad & "SELECT "
'    SqlCad = SqlCad & "IDINSUMO "
'    SqlCad = SqlCad & "FROM "
'    SqlCad = SqlCad & strTablaSQL & "DETALLE "
'    SqlCad = SqlCad & "WHERE "
'    SqlCad = SqlCad & "ID" & strTablaSQL & " = " & strIDExterno & " "
'    SqlCad = SqlCad & "GROUP BY "
'    SqlCad = SqlCad & "IDINSUMO"
'
'    rstValeDet.Open SqlCad, cnBdStudioModa, adOpenForwardOnly, adLockReadOnly
'
'    If Not rstValeDet.EOF Then
'        rstValeDet.MoveFirst
'
'        frameProgreso.Visible = True
'        barraProgreso.Max = ModUtilitario.devuelveCantRegistros(rstValeDet)
'        barraProgreso.Value = 0
'        frameProgreso.Caption = "Restaurando Stock..."
'
'        Do While Not rstValeDet.EOF
''            'Actualizar Tabla InsumoStock
''            SqlCad = vbNullString
''            SqlCad = SqlCad & "UPDATE "
''            SqlCad = SqlCad & "INSUMOSTOCK "
''            SqlCad = SqlCad & "SET "
''
''                Select Case strTipoVale
''                    Case "I"
''                        SqlCad = SqlCad & "STKACTUAL = (STKACTUAL - " & Val(rstValeDet!Cantidad & "") & "), "
''                        SqlCad = SqlCad & "STKCOMPRA = (STKCOMPRA - " & Val(rstValeDet!Cantidad & "") & ") "
''                    Case "S"
''                        SqlCad = SqlCad & "STKACTUAL = (STKACTUAL + " & Val(rstValeDet!Cantidad & "") & "), "
''                        SqlCad = SqlCad & "STKCOMPRA = (STKCOMPRA + " & Val(rstValeDet!Cantidad & "") & ") "
''                End Select
''
''            SqlCad = SqlCad & "WHERE "
''            SqlCad = SqlCad & "LTRIM(RTRIM(IDINSUMO)) = '" & Trim(rstValeDet!IDINSUMO & "") & "' AND "
''            SqlCad = SqlCad & "LTRIM(RTRIM(IDALMACEN)) = '" & strIdAlmacen & "'"
''
''            cnBdStudioModa.Execute SqlCad
''
''            Actualiza_Log "< DB Externo > " & SqlCad, StrConexDbBancos
'
'
'
'            'ACTUALIZAR TABLA INSUMOSTOCK
'            SqlCad = vbNullString
'            SqlCad = SqlCad & "UPDATE "
'            SqlCad = SqlCad & "INSUMOSTOCK "
'            SqlCad = SqlCad & "SET "
'
'                objAyudaVale.CodigoProducto = Trim(rstValeDet!IDINSUMO & "")
'
'                dblStockActual = objAyudaVale.devuelveStockFisicoDeProducto("C", True) + objAyudaVale.devuelveStockFisicoDeProducto("L", True)
'
'                SqlCad = SqlCad & "STKACTUAL = " & dblStockActual & ", "
'                SqlCad = SqlCad & "STKCOMPRA = " & dblStockActual & " "
'
'            SqlCad = SqlCad & "WHERE "
'            SqlCad = SqlCad & "LTRIM(RTRIM(IDINSUMO)) = '" & Trim(rstValeDet!IDINSUMO & "") & "' AND "
'            SqlCad = SqlCad & "LTRIM(RTRIM(IDALMACEN)) = '" & strIdAlmacen & "'"
'
'            cnBdStudioModa.Execute SqlCad
'
'            Actualiza_Log "< DB Externo > " & SqlCad, StrConexDbBancos
'
'            DoEvents
'
'            barraProgreso.Value = barraProgreso.Value + 1
'            frameProgreso.Caption = "Restaurando Stock... " & FormatPercent(barraProgreso.Value / barraProgreso.Max, 3)
'
'            rstValeDet.MoveNext
'        Loop
'    End If
'
'    'MODIFICAR
'    SqlCad = vbNullString
'    SqlCad = SqlCad & "UPDATE "
'    SqlCad = SqlCad & strTablaSQL & " "
'    SqlCad = SqlCad & "SET "
'    SqlCad = SqlCad & "OBSERVACIONANULACION = 'ANULACION EXTERNA DESDE CP POR " & wusuario & "', "
'    SqlCad = SqlCad & "IDPUNTOVENTA = '" & strIdPuntoVenta & "', "
'    SqlCad = SqlCad & "IDUSUARIOACTUALIZACION = '" & ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODUSEREXTERNO", "EF2USERS", "F2CODUSER", wusuario, "T") & "', "
'    SqlCad = SqlCad & "FECHAACTUALIZACION = GETDATE(), "
'    SqlCad = SqlCad & "ANULADO = 1 "
'    SqlCad = SqlCad & "WHERE "
'    SqlCad = SqlCad & "ID" & strTablaSQL & " = " & strIDExterno
'
'    cnBdStudioModa.Execute SqlCad
'
'    Actualiza_Log "< DB Externo > " & SqlCad, StrConexDbBancos
'
'    'MsgBox "Vale externo Anulado.", vbInformation + vbOKOnly, App.ProductName
'
'    anularValeExterno = True
'
'    frameProgreso.Visible = False
'
'    Exit Function
'errAnularValeExterno:
'    MsgBox "No.: " & Err.Number & vbNewLine & _
'            "Descripción: " & Err.Description & vbNewLine & vbNewLine & _
'            "Anulación de Vale Externo .", vbInformation + vbOKOnly, App.ProductName & " - ExportarValeAserverSQL"
'
'    anularValeExterno = False
'
'    Err.Clear
'End Function
'
Public Sub actualizarEstadoDescargadoOP(ByVal strIDExterno As String, _
                                        ByVal bolOPDescargada As Boolean)
    
    On Error GoTo errActualizarEstadoDescargadoOP
    
    'abrirCnDBMilano
    
    'MODIFICAR
    SqlCad = vbNullString
    SqlCad = SqlCad & "UPDATE "
    SqlCad = SqlCad & "ORDENPRODUCCION "
    SqlCad = SqlCad & "SET "
    SqlCad = SqlCad & "DESCARGADO = " & IIf(bolOPDescargada, "1", "0") & ", "
    SqlCad = SqlCad & "IDUSUARIOACTUALIZACION = '" & ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODUSEREXTERNO", "EF2USERS", "F2CODUSER", wusuario, "T") & "', "
    SqlCad = SqlCad & "FECHAACTUALIZACION = GETDATE() "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "IDORDENPRODUCCION = " & strIDExterno
    
    cnBdStudioModa.Execute SqlCad
    
    Actualiza_Log "< DB Externo > " & SqlCad, StrConexDbBancos
    
    Exit Sub
errActualizarEstadoDescargadoOP:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description & vbNewLine & vbNewLine & _
            "Actualización de Estado de Descarga de OP fallida.", vbInformation + vbOKOnly, App.ProductName & " - ActualizarEstadoDescargadoOP"
    
    Err.Clear
End Sub

Public Function verificarDatosOP(ByVal dblIdOrdenProduccion As Double, _
                                    ByVal strIdCategoria As String, _
                                    ByVal strOp As String, _
                                    ByVal cajaTextoIdOP As Object) As Boolean
    
    On Error GoTo errVerificarDatosOP
    
    verificarDatosOP = False
    
    Dim cmdOP As ADODB.Command
    Dim rstOp As ADODB.Recordset
    Dim intCantidadMesesDeValidezCompromiso As Integer
    
    intCantidadMesesDeValidezCompromiso = Val(ModUtilitario.sGetINI(App.Path & strNombreFicheroConfigSQLCliente, "ConfigServidorSQLCliente", "CantidadMesesDeValidezCompromiso", "l"))
    
    Set cmdOP = New ADODB.Command
    Set rstOp = New ADODB.Recordset
    
    With cmdOP
        .ActiveConnection = cnBdStudioModa
        .CommandType = adCmdStoredProc
        .CommandText = "usp_ConsultaOrdenProduccionCP"
        
        .Parameters.Append .CreateParameter("@IDOP", adBigInt, adParamInput, , dblIdOrdenProduccion)
        .Parameters.Append .CreateParameter("@CATEGORIA", adVarChar, adParamInput, 10, strIdCategoria)
        .Parameters.Append .CreateParameter("@OP", adVarChar, adParamInput, 30, strOp)
        .Parameters.Append .CreateParameter("@CANTIDADMESESVALIDEZPEDIDO", adBigInt, adParamInput, , intCantidadMesesDeValidezCompromiso)
        
        Set rstOp = .Execute()
    End With
    
    Set cmdOP = Nothing
    
    If Not rstOp.EOF Then
        If Not CBool(rstOp!Anulado & "") Then
            If MsgBox("Datos de O.P.: " & Trim(rstOp!CATEGORIA & "") & " / " & Trim(rstOp!NroOP & "") & vbNewLine & _
                    "Fecha: " & Trim(rstOp!FECOP & "") & vbNewLine & _
                    "No Pedido: " & Trim(rstOp!NroPedido & "") & " / Fec. Entrega: " & Trim(rstOp!FENTREGA & "") & vbNewLine & _
                    "Modelo: " & Trim(rstOp!Modelo & "") & " / Color: " & Trim(rstOp!Color & "") & vbNewLine & vbNewLine & _
                    "ESTADO: " & Trim(rstOp!ESTADOOPLETRAS & "") & vbNewLine & vbNewLine & _
                    "¿Desea descargar la O.P.?", vbQuestion + vbYesNo, wnomcia) = vbYes Then
                
                verificarDatosOP = True
            Else
                verificarDatosOP = False
            End If
            
            If Not cajaTextoIdOP Is Nothing Then
                cajaTextoIdOP.Text = Trim(rstOp!IDOP & "")
            End If
        Else
            MsgBox "O.P. anulada.", vbInformation + vbOKOnly, wnomcia
            
            verificarDatosOP = False
        End If
    Else
        MsgBox "O.P. no existe.", vbInformation + vbOKOnly, wnomcia
        
        verificarDatosOP = False
    End If
    
    SqlCad = vbNullString
    
    If rstOp.State = 1 Then rstOp.Close
    
    Set rstOp = Nothing
    
    Exit Function
errVerificarDatosOP:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    verificarDatosOP = False
    
    Err.Clear
End Function

Public Function importarOPServidorExterno(ByVal strIdOrdenProduccion As String, _
                                        ByVal frameProgreso As Object, _
                                        ByVal barraProgreso As Object) As Boolean
    
    On Error GoTo errImportarOPServidorExterno
    
    Dim rstOPCab As New ADODB.Recordset
    Dim rstOPDet As New ADODB.Recordset
    
    Dim dblItem As Double
    
    importarOPServidorExterno = False
    
    Rem LECTURA DE ORDEN DE PRODUCCION (OP)
    
    'abrirCnDBMilano
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "CAB.IDORDENPRODUCCION, "
    SqlCad = SqlCad & "CAB.IDREQUERIMIENTO, "
    SqlCad = SqlCad & "CAB.FECHA, "
    SqlCad = SqlCad & "CAB.CANTIDADTOTAL, "
    SqlCad = SqlCad & "CAB.DESCRIPCION, "
    SqlCad = SqlCad & "REQ.IDPEDIDO, "
    SqlCad = SqlCad & "REQ.NROMODELO, "
    SqlCad = SqlCad & "COL.NOMBRE AS COLOR "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "ORDENPRODUCCION AS CAB "
    SqlCad = SqlCad & "LEFT JOIN REQUERIMIENTO AS REQ ON REQ.IDREQUERIMIENTO = CAB.IDREQUERIMIENTO "
    SqlCad = SqlCad & "LEFT JOIN COLOR AS COL ON COL.IDCOLOR = REQ.IDCOLOR "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "CAB.ANULADO = 0 AND "
    SqlCad = SqlCad & "CAB.IDORDENPRODUCCION = " & strIdOrdenProduccion & " "
    
    If rstOPCab.State = 1 Then rstOPCab.Close
    
    rstOPCab.Open SqlCad, cnBdStudioModa, adOpenForwardOnly, adLockReadOnly
    
    If Not rstOPCab.EOF Then
        Do While Not rstOPCab.EOF
            With objAyudaOrdenTrabajo
                .inicializarEntidades
                
                .NumeroOrden = Trim(rstOPCab!IdOrdenProduccion & "")
                
                .eliminarOrdenTrabajo
                
                .inicializarEntidades
                
                .NumeroOrden = Trim(rstOPCab!IdOrdenProduccion & "")
                .NumeroPlan = Trim(rstOPCab!IDREQUERIMIENTO & "")
                .Fecha = Trim(rstOPCab!Fecha & "")
                .FechaEntrega = Trim(rstOPCab!Fecha & "")
                '.Observacion = Trim(rstOPCab!Descripcion & "")
                .Observacion = Trim(rstOPCab!Color & "")
                .TOTAL = Val(rstOPCab!CANTIDADTOTAL & "")
                .Facturado = Trim(rstOPCab!IDPEDIDO & "")
                .Estacion = ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "(TIPOMODELO + '-' + IDMODELO)", "MODELO", "NROMODELO", Trim(rstOPCab!NROMODELO & ""), "T")
                
                If .Facturado = vbNullString Then
                    MsgBox "Imposible descargar O.P. externa; no cuenta con No. Pedido asignado. " & vbNewLine & vbNewLine & "Comuniquese con su administrador de sistemas.", vbInformation + vbOKOnly, App.ProductName
                    
                    importarOPServidorExterno = False
                    
                    .inicializarEntidades
                    
                    Exit Function
                End If
                
                If .guardarOrdenTrabajo Then
                    Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                    
                    .Observacion = Trim(rstOPCab!Descripcion & "")
                    
                    'Borrar Insumos de OP
                    SqlCad = vbNullString
                    SqlCad = "DELETE FROM ORDENTRAB_INS WHERE NUMORDEN = '" & .NumeroOrden & "'"
                    
                    cnn_dbbancos.Execute SqlCad
                    Actualiza_Log SqlCad, cnn_dbbancos.ConnectionString
                    
                    SqlCad = vbNullString
                    SqlCad = SqlCad & "SELECT "
                    SqlCad = SqlCad & "DET.IDORDENPRODUCCION, "
                    SqlCad = SqlCad & "DET.ITEMS, "
                    SqlCad = SqlCad & "DET.IDINSUMO, "
                    SqlCad = SqlCad & "INS.NOMBRE, "
                    SqlCad = SqlCad & "INS.IDUNIDADMEDIDA, "
                    SqlCad = SqlCad & "DET.CANTIDAD, "
                    SqlCad = SqlCad & "REQ.IDPEDIDO "
                    SqlCad = SqlCad & "FROM "
                    SqlCad = SqlCad & "ORDENPRODUCCIONDESCARGO AS DET "
                    SqlCad = SqlCad & "LEFT JOIN ORDENPRODUCCION AS CAB ON CAB.IDORDENPRODUCCION = DET.IDORDENPRODUCCION "
                    SqlCad = SqlCad & "LEFT JOIN REQUERIMIENTO AS REQ ON REQ.IDREQUERIMIENTO = CAB.IDREQUERIMIENTO "
                    SqlCad = SqlCad & "LEFT JOIN INSUMO AS INS ON INS.IDINSUMO = DET.IDINSUMO "
                    SqlCad = SqlCad & "WHERE "
                    SqlCad = SqlCad & "DET.IDORDENPRODUCCION = " & .NumeroOrden & " AND "
                    SqlCad = SqlCad & "DET.ANULADO = 0 "
                    SqlCad = SqlCad & "ORDER BY "
                    SqlCad = SqlCad & "DET.ITEMS"
                    
                    If rstOPDet.State = 1 Then rstOPDet.Close
                    
                    rstOPDet.Open SqlCad, cnBdStudioModa, adOpenForwardOnly, adLockReadOnly
                    
                    If Not rstOPDet.EOF Then
                        
                        frameProgreso.Visible = True
                        barraProgreso.Max = ModUtilitario.devuelveCantRegistros(rstOPDet)
                        barraProgreso.Value = 0
                        frameProgreso.Caption = "Recuperando O.P. (1/2)..."
                        
                        Do While Not rstOPDet.EOF
                            .inicializarEntidadesInsumos
                            
                            .ITEM = Val(rstOPDet!items & "")
                            .CodigoProductoInsumo = Trim(rstOPDet!IDINSUMO & "")
                            .DescripcionProductoInsumo = Trim(rstOPDet!nombre & "")
                            .CodigoUM = Trim(rstOPDet!IDUNIDADMEDIDA & "")
                            .Cantidad = Val(rstOPDet!Cantidad & "")
                            .NumeroPedido = Trim(rstOPDet!IDPEDIDO & "")
                            
                            .guardarOTInsumoOneByOne
                            
                            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
        
                            rstOPDet.MoveNext
                            
                            DoEvents
                            
                            barraProgreso.Value = barraProgreso.Value + 1
                            frameProgreso.Caption = "Recuperando O.P. (1/2)... " & FormatPercent(barraProgreso.Value / barraProgreso.Max, 0)
                        Loop
                    End If
                    
                    
                    'Borrar Detalle de OP
                    SqlCad = vbNullString
                    SqlCad = "DELETE FROM ORDENTRAB_DET WHERE NUMORDEN = '" & .NumeroOrden & "'"
                    
                    cnn_dbbancos.Execute SqlCad
                    Actualiza_Log SqlCad, cnn_dbbancos.ConnectionString
                    
                    SqlCad = vbNullString
                    SqlCad = SqlCad & "SELECT "
                    SqlCad = SqlCad & "DET.IDORDENPRODUCCION, "
                    SqlCad = SqlCad & "DET.ITEMS, "
                    SqlCad = SqlCad & "DET.IDARTICULO, "
                    SqlCad = SqlCad & "DET.CANTIDAD, "
                    SqlCad = SqlCad & "REQ.IDPEDIDO "
                    SqlCad = SqlCad & "FROM "
                    SqlCad = SqlCad & "ORDENPRODUCCIONDETALLE AS DET "
                    SqlCad = SqlCad & "LEFT JOIN ORDENPRODUCCION AS CAB ON CAB.IDORDENPRODUCCION = DET.IDORDENPRODUCCION "
                    SqlCad = SqlCad & "LEFT JOIN REQUERIMIENTO AS REQ ON REQ.IDREQUERIMIENTO = CAB.IDREQUERIMIENTO "
                    SqlCad = SqlCad & "WHERE "
                    SqlCad = SqlCad & "DET.IDORDENPRODUCCION = " & .NumeroOrden & " "
                    SqlCad = SqlCad & "ORDER BY "
                    SqlCad = SqlCad & "DET.ITEMS"
                    
                    If rstOPDet.State = 1 Then rstOPDet.Close
                    
                    rstOPDet.Open SqlCad, cnBdStudioModa, adOpenForwardOnly, adLockReadOnly
                    
                    If Not rstOPDet.EOF Then
                        frameProgreso.Visible = True
                        barraProgreso.Max = ModUtilitario.devuelveCantRegistros(rstOPDet)
                        barraProgreso.Value = 0
                        frameProgreso.Caption = "Recuperando O.P. (2/2)..."
                        
                        Do While Not rstOPDet.EOF
                            .inicializarEntidadesDetalle
                            
                            .NumeroOrden = Trim(rstOPDet!IdOrdenProduccion & "")
                            .CodigoProducto = Trim(rstOPDet!IDARTICULO & "")
                            .DescripcionProducto = Trim(rstOPDet!IDARTICULO & "")
                            .Cantidad = Val(rstOPDet!Cantidad & "")
                            .NumeroPedido = Trim(rstOPDet!IDPEDIDO & "")
                            
                            .guardarOTDetalleOneByOne
                            
                            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
    
                            rstOPDet.MoveNext
                            
                            DoEvents
                            
                            barraProgreso.Value = barraProgreso.Value + 1
                            frameProgreso.Caption = "Recuperando O.P. (2/2)... " & FormatPercent(barraProgreso.Value / barraProgreso.Max, 0)
                        Loop
                    End If
                Else
                    importarOPServidorExterno = False
                    
                    Exit Function
                End If
            End With
            
            rstOPCab.MoveNext
        Loop
            'MsgBox "Actualización Finalizada", vbInformation + vbOKOnly, wnomcia
            
            importarOPServidorExterno = True
    Else
        MsgBox "No se ubicaron rastros de O.P.", vbInformation + vbOKOnly, wnomcia
    End If
    
    'cnBdStudioModa.Close
    SqlCad = vbNullString
    
    frameProgreso.Visible = False
    
    'If cnBdStudioModa.State = 1 Then cnBdStudioModa.Close
    If rstOPCab.State = 1 Then rstOPCab.Close
    If rstOPDet.State = 1 Then rstOPDet.Close
    
    'Set cnBdStudioModa = Nothing
    Set rstOPCab = Nothing
    Set rstOPDet = Nothing
    
    Exit Function
errImportarOPServidorExterno:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    importarOPServidorExterno = False
    
    frameProgreso.Visible = False
    barraProgreso.Value = 0
    
    Err.Clear
End Function

Public Function importarOPServidorExternoV2(ByVal strIdOrdenProduccion As String, _
                                            ByVal frameProgreso As Object, _
                                            ByVal barraProgreso As Object) As Boolean
    
    On Error GoTo errImportarOPServidorExternoV2
    
    Dim rstOpInsumoPend As New ADODB.Recordset
    Dim rstOPDet As New ADODB.Recordset
    
    Dim dblItem As Double
    
    importarOPServidorExternoV2 = False
    
    Rem LECTURA DE ORDEN DE PRODUCCION (OP)
    
    'abrirCnDBMilano
    
'    SqlCad = vbNullString
'    SqlCad = SqlCad & "SELECT "
'    SqlCad = SqlCad & "DET.IDORDENPRODUCCION, "
'    SqlCad = SqlCad & "RQ.IDPEDIDO, "
'    SqlCad = SqlCad & "'[ ID Orden de Produccion: ' + CAST(DET.IDORDENPRODUCCION AS VARCHAR(20)) + ' ]' + SPACE(18) + "
'    SqlCad = SqlCad & "'[ No. Pedido: ' + RQ.IDPEDIDO + ' ]' + SPACE(18) + "
'    SqlCad = SqlCad & "'[ No. Modelo: ' + MD.TIPOMODELO + '-' + MD.IDMODELO + ' ]' + SPACE(18) + "
'    SqlCad = SqlCad & "'[ Color: ' + COL.NOMBRE + ' ]' + SPACE(18) + "
'    SqlCad = SqlCad & "'[ Cantidad: ' + CAST(CAB.CANTIDADTOTAL AS VARCHAR(18)) + ' ]' AS LLAVE, "
'    SqlCad = SqlCad & "DET.IDINSUMO AS ORIGEN, "
'    SqlCad = SqlCad & "DET.IDINSUMO AS FINAL, "
'    SqlCad = SqlCad & "INS.NOMBRE AS DESCRIPCIONINSUMO, "
'    SqlCad = SqlCad & "UM.NOMBRE AS UNIDADMEDIDA, "
'    SqlCad = SqlCad & "DET.CANTIDAD AS CANTIDADORIGEN, "
'    SqlCad = SqlCad & "DET.CANTIDAD AS CANTIDADFINAL, "
'    'SqlCad = SqlCad & "(DET.CANTIDAD - ISNULL(MOVOP.CANTIDAD, 0)) AS SALDO "
'    SqlCad = SqlCad & "((DET.CANTIDAD - ISNULL(MOVOP.CANTIDAD, 0)) + ISNULL(MOVOPINGRESO.CANTIDAD, 0)) AS SALDO "
'
'    SqlCad = SqlCad & "FROM "
'    SqlCad = SqlCad & "ORDENPRODUCCIONDESCARGO AS DET "
'    SqlCad = SqlCad & "LEFT JOIN ORDENPRODUCCION AS CAB ON CAB.IDORDENPRODUCCION = DET.IDORDENPRODUCCION "
'    SqlCad = SqlCad & "LEFT JOIN "
'    SqlCad = SqlCad & "(SELECT "
'    SqlCad = SqlCad & "IDREQUERIMIENTO, "
'    SqlCad = SqlCad & "(CASE WHEN IDPEDIDO = 0 THEN NROPEDIDO ELSE IDPEDIDO END) AS IDPEDIDO, "
'    SqlCad = SqlCad & "NROMODELO, "
'    SqlCad = SqlCad & "IDCOLOR "
'    SqlCad = SqlCad & "FROM "
'    SqlCad = SqlCad & "REQUERIMIENTO) AS RQ "
'    SqlCad = SqlCad & "ON RQ.IDREQUERIMIENTO = CAB.IDREQUERIMIENTO "
'    SqlCad = SqlCad & "LEFT JOIN MODELO AS MD ON MD.NROMODELO = RQ.NROMODELO "
'    SqlCad = SqlCad & "LEFT JOIN INSUMO AS INS ON INS.IDINSUMO = DET.IDINSUMO "
'    SqlCad = SqlCad & "LEFT JOIN UNIDADMEDIDA AS UM ON UM.IDUNIDADMEDIDA = INS.IDUNIDADMEDIDA "
'    SqlCad = SqlCad & "LEFT JOIN COLOR AS COL ON COL.IDCOLOR = RQ.IDCOLOR "
'
'    SqlCad = SqlCad & "LEFT JOIN "
'    SqlCad = SqlCad & "(SELECT "
'    SqlCad = SqlCad & "CAB.IDORDENPRODUCCION, "
'    SqlCad = SqlCad & "RQ.IDPEDIDO, "
'    SqlCad = SqlCad & "DET.IDINSUMO, "
'    SqlCad = SqlCad & "SUM(DET.CANTIDAD) AS CANTIDAD "
'    SqlCad = SqlCad & "FROM "
'    SqlCad = SqlCad & "SALIDADETALLE AS DET "
'    SqlCad = SqlCad & "LEFT JOIN SALIDA AS CAB "
'    SqlCad = SqlCad & "ON CAB.IDSALIDA = DET.IDSALIDA "
'    SqlCad = SqlCad & "LEFT JOIN ORDENPRODUCCION AS OP ON OP.IDORDENPRODUCCION = CAB.IDORDENPRODUCCION "
'    SqlCad = SqlCad & "LEFT JOIN "
'    SqlCad = SqlCad & "(SELECT "
'    SqlCad = SqlCad & "IDREQUERIMIENTO, "
'    SqlCad = SqlCad & "(CASE WHEN IDPEDIDO = 0 THEN NROPEDIDO ELSE IDPEDIDO END) AS IDPEDIDO "
'    SqlCad = SqlCad & "FROM REQUERIMIENTO) AS RQ "
'    SqlCad = SqlCad & "ON RQ.IDREQUERIMIENTO = OP.IDREQUERIMIENTO "
'    SqlCad = SqlCad & "WHERE "
'    SqlCad = SqlCad & "CAB.IDTIPOSALIDA IN ('TIP0000001') AND "
'    SqlCad = SqlCad & "CAB.ANULADO = 0 AND "
'    'SqlCad = SqlCad & "CAB.IDORDENPRODUCCION <> 0 AND "
'    SqlCad = SqlCad & "CAB.IDORDENPRODUCCION = " & strIdOrdenProduccion & " AND "
'    SqlCad = SqlCad & "RQ.IDPEDIDO <> 0 "
'    SqlCad = SqlCad & "GROUP BY "
'    SqlCad = SqlCad & "CAB.IDORDENPRODUCCION, "
'    SqlCad = SqlCad & "RQ.IDPEDIDO, DET.IDINSUMO) AS MOVOP "
'    SqlCad = SqlCad & "ON MOVOP.IDORDENPRODUCCION = DET.IDORDENPRODUCCION AND MOVOP.IDPEDIDO = RQ.IDPEDIDO AND MOVOP.IDINSUMO = DET.IDINSUMO "
'
'    SqlCad = SqlCad & "LEFT JOIN "
'    SqlCad = SqlCad & "(SELECT "
'    SqlCad = SqlCad & "CAB.IDORDENPRODUCCION, "
'    SqlCad = SqlCad & "RQ.IDPEDIDO, "
'    SqlCad = SqlCad & "DET.IDINSUMO, "
'    SqlCad = SqlCad & "SUM(DET.CANTIDAD) AS CANTIDAD "
'    SqlCad = SqlCad & "FROM "
'    SqlCad = SqlCad & "INGRESODETALLE AS DET "
'    SqlCad = SqlCad & "LEFT JOIN INGRESO AS CAB "
'    SqlCad = SqlCad & "ON CAB.IDINGRESO = DET.IDINGRESO "
'    SqlCad = SqlCad & "LEFT JOIN ORDENPRODUCCION AS OP ON OP.IDORDENPRODUCCION = CAB.IDORDENPRODUCCION "
'    SqlCad = SqlCad & "LEFT JOIN "
'    SqlCad = SqlCad & "(SELECT "
'    SqlCad = SqlCad & "IDREQUERIMIENTO, "
'    SqlCad = SqlCad & "(CASE WHEN IDPEDIDO = 0 THEN NROPEDIDO ELSE IDPEDIDO END) AS IDPEDIDO "
'    SqlCad = SqlCad & "FROM REQUERIMIENTO) AS RQ "
'    SqlCad = SqlCad & "ON RQ.IDREQUERIMIENTO = OP.IDREQUERIMIENTO "
'    SqlCad = SqlCad & "WHERE "
'    SqlCad = SqlCad & "CAB.IDTIPOINGRESO IN ('TIP0000004') AND "
'    SqlCad = SqlCad & "CAB.ANULADO = 0 AND "
'    'SqlCad = SqlCad & "CAB.IDORDENPRODUCCION <> 0 AND "
'    SqlCad = SqlCad & "CAB.IDORDENPRODUCCION = " & strIdOrdenProduccion & " AND "
'    SqlCad = SqlCad & "RQ.IDPEDIDO <> 0 "
'    SqlCad = SqlCad & "GROUP BY "
'    SqlCad = SqlCad & "CAB.IDORDENPRODUCCION, "
'    SqlCad = SqlCad & "RQ.IDPEDIDO, DET.IDINSUMO) AS MOVOPINGRESO "
'    SqlCad = SqlCad & "ON MOVOPINGRESO.IDORDENPRODUCCION = DET.IDORDENPRODUCCION AND MOVOPINGRESO.IDPEDIDO = RQ.IDPEDIDO AND MOVOPINGRESO.IDINSUMO = DET.IDINSUMO "
'
'    SqlCad = SqlCad & "WHERE "
'    SqlCad = SqlCad & "DET.ANULADO = 0 AND "
'    SqlCad = SqlCad & "RQ.IDPEDIDO <> 0 AND "
'    'SqlCad = SqlCad & "(DET.CANTIDAD - ISNULL(MOVOP.CANTIDAD, 0)) > 0 AND "
'    SqlCad = SqlCad & "((DET.CANTIDAD - ISNULL(MOVOP.CANTIDAD, 0)) + ISNULL(MOVOPINGRESO.CANTIDAD, 0)) > 0 AND "
'    SqlCad = SqlCad & "CAB.CERRADO = 0 AND "
'    SqlCad = SqlCad & "CAB.ANULADO = 0 AND "
'    SqlCad = SqlCad & "INS.OP = 1 AND "
'    SqlCad = SqlCad & "CAB.IDORDENPRODUCCION = " & strIdOrdenProduccion
'
'    If rstOpInsumoPend.State = 1 Then rstOpInsumoPend.Close
'
'    rstOpInsumoPend.Open SqlCad, cnBdStudioModa, adOpenForwardOnly, adLockReadOnly

    Dim cmdOP As ADODB.Command
    
    Set cmdOP = New ADODB.Command
    
    With cmdOP
        .ActiveConnection = cnBdStudioModa
        .CommandType = adCmdStoredProc
        .CommandText = "usp_ConsultaOrdenProduccionDescargaCP"
        
        .Parameters.Append .CreateParameter("@IDOP", adBigInt, adParamInput, , strIdOrdenProduccion)
        .Parameters.Append .CreateParameter("@CANTIDADFILAS", adInteger, adParamOutput, 5)
        
        .Execute
        
        If Val(.Parameters("@CANTIDADFILAS") & "") > 0 Then
            barraProgreso.Max = Val(.Parameters("@CANTIDADFILAS") & "")
        End If
        
        Set rstOpInsumoPend = .Execute()
    End With
    
    Set cmdOP = Nothing
    
    If Not rstOpInsumoPend.EOF Then
        abrirCnTemporal

        cnDBTemp.Execute "DELETE FROM TMPUTILDESCARGAOPPENDIENTE"
                
        frameProgreso.Visible = True
        'barraProgreso.Max = ModUtilitario.devuelveCantRegistros(rstOpInsumoPend)
        barraProgreso.Value = 0
        frameProgreso.Caption = "Recuperando O.P. (1/1)..."
        
        Do While Not rstOpInsumoPend.EOF
            SqlCad = vbNullString
            SqlCad = SqlCad & "INSERT INTO TMPUTILDESCARGAOPPENDIENTE("
            SqlCad = SqlCad & "NROOP, "
            SqlCad = SqlCad & "NROPEDIDO, "
            SqlCad = SqlCad & "LLAVEOP, "
            SqlCad = SqlCad & "CODPRODUCTOORIGEN, "
            SqlCad = SqlCad & "CODPRODUCTOFINAL, "
            SqlCad = SqlCad & "NOMPRODUCTO, "
            SqlCad = SqlCad & "UM, "
            SqlCad = SqlCad & "CANTIDADORIGEN, "
            SqlCad = SqlCad & "CANTIDADFINAL, "
            SqlCad = SqlCad & "SALDO) "
            
            SqlCad = SqlCad & "VALUES("
            SqlCad = SqlCad & "'" & Trim(rstOpInsumoPend!IdOrdenProduccion & "") & "', "
            SqlCad = SqlCad & "'" & Trim(rstOpInsumoPend!IDPEDIDO & "") & "', "
            SqlCad = SqlCad & "'" & Trim(rstOpInsumoPend!LLAVE & "") & "', "
            SqlCad = SqlCad & "'" & Trim(rstOpInsumoPend!Origen & "") & "', "
            SqlCad = SqlCad & "'" & Trim(rstOpInsumoPend!Final & "") & "', "
            SqlCad = SqlCad & "'" & Trim(rstOpInsumoPend!DESCRIPCIONINSUMO & "") & "', "
            SqlCad = SqlCad & "'" & Trim(rstOpInsumoPend!UNIDADMEDIDA & "") & "', "
            SqlCad = SqlCad & Val(Format(Val(rstOpInsumoPend!CANTIDADORIGEN & ""), "#0.00")) & ", "
            SqlCad = SqlCad & Val(Format(Val(rstOpInsumoPend!CantidadFinal & ""), "#0.00")) & ", "
            SqlCad = SqlCad & Val(Format(Val(rstOpInsumoPend!SALDO & ""), "#0.00")) & ")"
            
            abrirCnTemporal
            
            cnDBTemp.Execute SqlCad
            
            DoEvents
                            
            barraProgreso.Value = barraProgreso.Value + 1
            frameProgreso.Caption = "Recuperando O.P. (1/1)... " & FormatPercent(barraProgreso.Value / barraProgreso.Max, 0)
            
            rstOpInsumoPend.MoveNext
        Loop
            importarOPServidorExternoV2 = True
    Else
        MsgBox "O.P. sin insumos pendientes de descarga.", vbInformation + vbOKOnly, wnomcia
    End If
    
    'cnBdStudioModa.Close
    SqlCad = vbNullString
    
    frameProgreso.Visible = False
    
    ''If cnBdStudioModa.State = 1 Then cnBdStudioModa.Close
    If rstOpInsumoPend.State = 1 Then rstOpInsumoPend.Close
    If rstOPDet.State = 1 Then rstOPDet.Close
    
    'Set cnBdStudioModa = Nothing
    Set rstOpInsumoPend = Nothing
    Set rstOPDet = Nothing
    
    Exit Function
errImportarOPServidorExternoV2:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    importarOPServidorExternoV2 = False
    
    frameProgreso.Visible = False
    barraProgreso.Value = 0
    
    Err.Clear
End Function

Public Function importarOPServidorExternoV3(ByVal strIdOrdenProduccion As String, _
                                            ByVal frameProgreso As Object, _
                                            ByVal barraProgreso As Object) As Boolean
    
    On Error GoTo errImportarOPServidorExternoV3
    
    Dim rstOpInsumoPend As New ADODB.Recordset
    Dim rstOPDet As New ADODB.Recordset
    
    Dim dblItem As Double
    
    importarOPServidorExternoV3 = False
    
    Rem LECTURA DE ORDEN DE PRODUCCION (OP)
    
    'abrirCnDBMilano
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "DET.IDORDENPRODUCCION, "
    SqlCad = SqlCad & "RQ.IDPEDIDO, "
    SqlCad = SqlCad & "'[ ID Orden de Produccion: ' + CAST(DET.IDORDENPRODUCCION AS VARCHAR(20)) + ' ]' + SPACE(18) + "
    SqlCad = SqlCad & "'[ No. Pedido: ' + RQ.IDPEDIDO + ' ]' + SPACE(18) + "
    SqlCad = SqlCad & "'[ No. Modelo: ' + MD.TIPOMODELO + '-' + MD.IDMODELO + ' ]' + SPACE(18) + "
    SqlCad = SqlCad & "'[ Color: ' + COL.NOMBRE + ' ]' + SPACE(18) + "
    SqlCad = SqlCad & "'[ Cantidad: ' + CAST(CAB.CANTIDADTOTAL AS VARCHAR(18)) + ' ]' AS LLAVE, "
    SqlCad = SqlCad & "DET.IDINSUMO AS ORIGEN, "
    SqlCad = SqlCad & "DET.IDINSUMO AS FINAL, "
    SqlCad = SqlCad & "INS.NOMBRE AS DESCRIPCIONINSUMO, "
    SqlCad = SqlCad & "UM.NOMBRE AS UNIDADMEDIDA, "
    SqlCad = SqlCad & "DET.CANTIDAD AS CANTIDADORIGEN, "
    SqlCad = SqlCad & "DET.CANTIDAD AS CANTIDADFINAL, "
    'SqlCad = SqlCad & "(DET.CANTIDAD - ISNULL(MOVOP.CANTIDAD, 0)) AS SALDO "
    SqlCad = SqlCad & "((DET.CANTIDAD - ISNULL(MOVOP.CANTIDAD, 0)) + ISNULL(MOVOPINGRESO.CANTIDAD, 0)) AS SALDO "
    
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "ORDENPRODUCCIONDESCARGO AS DET "
    SqlCad = SqlCad & "LEFT JOIN ORDENPRODUCCION AS CAB ON CAB.IDORDENPRODUCCION = DET.IDORDENPRODUCCION "
    SqlCad = SqlCad & "LEFT JOIN "
    SqlCad = SqlCad & "(SELECT "
    SqlCad = SqlCad & "IDREQUERIMIENTO, "
    SqlCad = SqlCad & "(CASE WHEN IDPEDIDO = 0 THEN NROPEDIDO ELSE IDPEDIDO END) AS IDPEDIDO, "
    SqlCad = SqlCad & "NROMODELO, "
    SqlCad = SqlCad & "IDCOLOR "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "REQUERIMIENTO) AS RQ "
    SqlCad = SqlCad & "ON RQ.IDREQUERIMIENTO = CAB.IDREQUERIMIENTO "
    SqlCad = SqlCad & "LEFT JOIN MODELO AS MD ON MD.NROMODELO = RQ.NROMODELO "
    SqlCad = SqlCad & "LEFT JOIN INSUMO AS INS ON INS.IDINSUMO = DET.IDINSUMO "
    SqlCad = SqlCad & "LEFT JOIN UNIDADMEDIDA AS UM ON UM.IDUNIDADMEDIDA = INS.IDUNIDADMEDIDA "
    SqlCad = SqlCad & "LEFT JOIN COLOR AS COL ON COL.IDCOLOR = RQ.IDCOLOR "
    
    SqlCad = SqlCad & "LEFT JOIN "
    SqlCad = SqlCad & "(SELECT "
    SqlCad = SqlCad & "CAB.IDORDENPRODUCCION, "
    SqlCad = SqlCad & "RQ.IDPEDIDO, "
    SqlCad = SqlCad & "DET.IDINSUMO, "
    SqlCad = SqlCad & "SUM(DET.CANTIDAD) AS CANTIDAD "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "SALIDADETALLE AS DET "
    SqlCad = SqlCad & "LEFT JOIN SALIDA AS CAB "
    SqlCad = SqlCad & "ON CAB.IDSALIDA = DET.IDSALIDA "
    SqlCad = SqlCad & "LEFT JOIN ORDENPRODUCCION AS OP ON OP.IDORDENPRODUCCION = CAB.IDORDENPRODUCCION "
    SqlCad = SqlCad & "LEFT JOIN "
    SqlCad = SqlCad & "(SELECT "
    SqlCad = SqlCad & "IDREQUERIMIENTO, "
    SqlCad = SqlCad & "(CASE WHEN IDPEDIDO = 0 THEN NROPEDIDO ELSE IDPEDIDO END) AS IDPEDIDO "
    SqlCad = SqlCad & "FROM REQUERIMIENTO) AS RQ "
    SqlCad = SqlCad & "ON RQ.IDREQUERIMIENTO = OP.IDREQUERIMIENTO "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "CAB.IDTIPOSALIDA IN ('TIP0000001') AND "
    SqlCad = SqlCad & "CAB.ANULADO = 0 AND "
    'SqlCad = SqlCad & "CAB.IDORDENPRODUCCION <> 0 AND "
    SqlCad = SqlCad & "CAB.IDORDENPRODUCCION = " & strIdOrdenProduccion & " AND "
    SqlCad = SqlCad & "RQ.IDPEDIDO <> 0 "
    SqlCad = SqlCad & "GROUP BY "
    SqlCad = SqlCad & "CAB.IDORDENPRODUCCION, "
    SqlCad = SqlCad & "RQ.IDPEDIDO, DET.IDINSUMO) AS MOVOP "
    SqlCad = SqlCad & "ON MOVOP.IDORDENPRODUCCION = DET.IDORDENPRODUCCION AND MOVOP.IDPEDIDO = RQ.IDPEDIDO AND MOVOP.IDINSUMO = DET.IDINSUMO "
    
    SqlCad = SqlCad & "LEFT JOIN "
    SqlCad = SqlCad & "(SELECT "
    SqlCad = SqlCad & "CAB.IDORDENPRODUCCION, "
    SqlCad = SqlCad & "RQ.IDPEDIDO, "
    SqlCad = SqlCad & "DET.IDINSUMO, "
    SqlCad = SqlCad & "SUM(DET.CANTIDAD) AS CANTIDAD "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "INGRESODETALLE AS DET "
    SqlCad = SqlCad & "LEFT JOIN INGRESO AS CAB "
    SqlCad = SqlCad & "ON CAB.IDINGRESO = DET.IDINGRESO "
    SqlCad = SqlCad & "LEFT JOIN ORDENPRODUCCION AS OP ON OP.IDORDENPRODUCCION = CAB.IDORDENPRODUCCION "
    SqlCad = SqlCad & "LEFT JOIN "
    SqlCad = SqlCad & "(SELECT "
    SqlCad = SqlCad & "IDREQUERIMIENTO, "
    SqlCad = SqlCad & "(CASE WHEN IDPEDIDO = 0 THEN NROPEDIDO ELSE IDPEDIDO END) AS IDPEDIDO "
    SqlCad = SqlCad & "FROM REQUERIMIENTO) AS RQ "
    SqlCad = SqlCad & "ON RQ.IDREQUERIMIENTO = OP.IDREQUERIMIENTO "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "CAB.IDTIPOINGRESO IN ('TIP0000004') AND "
    SqlCad = SqlCad & "CAB.ANULADO = 0 AND "
    'SqlCad = SqlCad & "CAB.IDORDENPRODUCCION <> 0 AND "
    SqlCad = SqlCad & "CAB.IDORDENPRODUCCION = " & strIdOrdenProduccion & " AND "
    SqlCad = SqlCad & "RQ.IDPEDIDO <> 0 "
    SqlCad = SqlCad & "GROUP BY "
    SqlCad = SqlCad & "CAB.IDORDENPRODUCCION, "
    SqlCad = SqlCad & "RQ.IDPEDIDO, DET.IDINSUMO) AS MOVOPINGRESO "
    SqlCad = SqlCad & "ON MOVOPINGRESO.IDORDENPRODUCCION = DET.IDORDENPRODUCCION AND MOVOPINGRESO.IDPEDIDO = RQ.IDPEDIDO AND MOVOPINGRESO.IDINSUMO = DET.IDINSUMO "
    
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "DET.ANULADO = 0 AND "
    SqlCad = SqlCad & "RQ.IDPEDIDO <> 0 AND "
    'SqlCad = SqlCad & "(DET.CANTIDAD - ISNULL(MOVOP.CANTIDAD, 0)) > 0 AND "
    SqlCad = SqlCad & "((DET.CANTIDAD - ISNULL(MOVOP.CANTIDAD, 0)) + ISNULL(MOVOPINGRESO.CANTIDAD, 0)) > 0 AND "
    SqlCad = SqlCad & "CAB.CERRADO = 0 AND "
    SqlCad = SqlCad & "CAB.ANULADO = 0 AND "
    SqlCad = SqlCad & "INS.OP = 1 AND "
    SqlCad = SqlCad & "CAB.IDORDENPRODUCCION = " & strIdOrdenProduccion
    
    If rstOpInsumoPend.State = 1 Then rstOpInsumoPend.Close
    
    rstOpInsumoPend.Open SqlCad, cnBdStudioModa, adOpenForwardOnly, adLockReadOnly
    
    If Not rstOpInsumoPend.EOF Then
        abrirCnTemporal

        cnDBTemp.Execute "DELETE FROM TMPUTILDESCARGAOPSENCILLA"
                
        frameProgreso.Visible = True
        barraProgreso.Max = ModUtilitario.devuelveCantRegistros(rstOpInsumoPend)
        barraProgreso.Value = 0
        frameProgreso.Caption = "Recuperando O.P. (1/1)..."
        
        Do While Not rstOpInsumoPend.EOF
            SqlCad = vbNullString
            SqlCad = SqlCad & "INSERT INTO TMPUTILDESCARGAOPSENCILLA("
            SqlCad = SqlCad & "NROOP, "
            SqlCad = SqlCad & "NROPEDIDO, "
            SqlCad = SqlCad & "LLAVEOP, "
            SqlCad = SqlCad & "CODPRODUCTOORIGEN, "
            SqlCad = SqlCad & "CODPRODUCTOFINAL, "
            SqlCad = SqlCad & "NOMPRODUCTO, "
            SqlCad = SqlCad & "UM, "
            SqlCad = SqlCad & "CANTIDADORIGEN, "
            SqlCad = SqlCad & "CANTIDADFINAL, "
            SqlCad = SqlCad & "SALDO, "
            SqlCad = SqlCad & "STOCKCOMPROMETIDO, "
            SqlCad = SqlCad & "STOCKPORLLEGAR, "
            SqlCad = SqlCad & "CANTIDADDESCARGA, "
            SqlCad = SqlCad & "PROCESAR"
            SqlCad = SqlCad & ") "
            
            SqlCad = SqlCad & "VALUES("
            SqlCad = SqlCad & "'" & Trim(rstOpInsumoPend!IdOrdenProduccion & "") & "', "
            SqlCad = SqlCad & "'" & Trim(rstOpInsumoPend!IDPEDIDO & "") & "', "
            SqlCad = SqlCad & "'" & Trim(rstOpInsumoPend!LLAVE & "") & "', "
            SqlCad = SqlCad & "'" & Trim(rstOpInsumoPend!Origen & "") & "', "
            SqlCad = SqlCad & "'" & Trim(rstOpInsumoPend!Final & "") & "', "
            SqlCad = SqlCad & "'" & Trim(rstOpInsumoPend!DESCRIPCIONINSUMO & "") & "', "
            SqlCad = SqlCad & "'" & Trim(rstOpInsumoPend!UNIDADMEDIDA & "") & "', "
            SqlCad = SqlCad & Val(Format(Val(rstOpInsumoPend!CANTIDADORIGEN & ""), "#0.00")) & ", "
            SqlCad = SqlCad & Val(Format(Val(rstOpInsumoPend!CantidadFinal & ""), "#0.00")) & ", "
            SqlCad = SqlCad & Val(Format(Val(rstOpInsumoPend!SALDO & ""), "#0.00")) & ", "
            SqlCad = SqlCad & "0, 0, 0, FALSE"
            SqlCad = SqlCad & ")"
            
            abrirCnTemporal
            
            cnDBTemp.Execute SqlCad
            
            DoEvents
                            
            barraProgreso.Value = barraProgreso.Value + 1
            frameProgreso.Caption = "Recuperando O.P. (1/1)... " & FormatPercent(barraProgreso.Value / barraProgreso.Max, 0)
            
            rstOpInsumoPend.MoveNext
        Loop
            importarOPServidorExternoV3 = True
    Else
        MsgBox "O.P. sin insumos pendientes de descarga.", vbInformation + vbOKOnly, wnomcia
    End If
    
    'cnBdStudioModa.Close
    SqlCad = vbNullString
    
    frameProgreso.Visible = False
    
    ''If cnBdStudioModa.State = 1 Then cnBdStudioModa.Close
    If rstOpInsumoPend.State = 1 Then rstOpInsumoPend.Close
    If rstOPDet.State = 1 Then rstOPDet.Close
    
    'Set cnBdStudioModa = Nothing
    Set rstOpInsumoPend = Nothing
    Set rstOPDet = Nothing
    
    Exit Function
errImportarOPServidorExternoV3:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    'Resume
    importarOPServidorExternoV3 = False
    
    frameProgreso.Visible = False
    barraProgreso.Value = 0
    
    Err.Clear
End Function

Public Function insertarProductoEnOPServidorExterno(ByVal strIdOrdenProduccion As String, _
                                                        ByVal strIdInsumo As String) As Boolean
    
    On Error GoTo errInsertarProductoEnOPServidorExterno
    
    Dim dblItem As Double
    
    insertarProductoEnOPServidorExterno = False
    
    Rem LECTURA DE ORDEN DE PRODUCCION (OP)
    
    'abrirCnDBMilano
    
    dblItem = Val(ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "TOP 1 [ITEMS]", "ORDENPRODUCCIONDESCARGO", "IDORDENPRODUCCION", strIdOrdenProduccion, "N", "ORDER BY [ITEMS] DESC")) + 1
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "INSERT INTO ORDENPRODUCCIONDESCARGO("
    SqlCad = SqlCad & "IDORDENPRODUCCION, "
    SqlCad = SqlCad & "[ITEMS], "
    SqlCad = SqlCad & "IDINSUMO, "
    SqlCad = SqlCad & "CANTIDAD, "
    SqlCad = SqlCad & "IDINSUMOV, "
    SqlCad = SqlCad & "ANULADO"
    SqlCad = SqlCad & ") "
    SqlCad = SqlCad & "VALUES("
    SqlCad = SqlCad & strIdOrdenProduccion & ", "
    SqlCad = SqlCad & dblItem & ", "
    SqlCad = SqlCad & "'" & strIdInsumo & "', "
    SqlCad = SqlCad & "1, "
    SqlCad = SqlCad & "'" & strIdInsumo & "', "
    SqlCad = SqlCad & "0"
    SqlCad = SqlCad & ")"
    
    cnBdStudioModa.Execute SqlCad
    
    Actualiza_Log "< DB Externo > " & SqlCad, StrConexDbBancos
    
    
    Dim lngIdCambio As Long
    Dim strIdUsuario As String
    
    lngIdCambio = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "TOP 1 IDCAMBIO", "SF1ORDENPRODUCCION_LOG", "IDORDENPRODUCCION", strIdOrdenProduccion, "T", "ORDER BY IDCAMBIO DESC") & "") + 1
    strIdUsuario = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODUSEREXTERNO", "EF2USERS", "F2CODUSER", wusuario, "T")
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "INSERT INTO SF1ORDENPRODUCCION_LOG("
    SqlCad = SqlCad & "IDORDENPRODUCCION, IDCAMBIO, IDINSUMO, IDINSUMOFINAL, "
    SqlCad = SqlCad & "CANTIDAD, CANTIDADFINAL, IDUSUARIO, FECHAMODIFICACION, "
    SqlCad = SqlCad & "OBSERVACION"
    SqlCad = SqlCad & ") "
    SqlCad = SqlCad & "VALUES("
    SqlCad = SqlCad & "'" & strIdOrdenProduccion & "', "
    SqlCad = SqlCad & lngIdCambio & ", "
    SqlCad = SqlCad & "'" & strIdInsumo & "', "
    SqlCad = SqlCad & "'" & strIdInsumo & "', "
    SqlCad = SqlCad & "1, "
    SqlCad = SqlCad & "1, "
    SqlCad = SqlCad & "'" & IIf(strIdUsuario <> vbNullString, strIdUsuario, wusuario) & "', "
    SqlCad = SqlCad & "CVDATE('" & Now & "'), "
    SqlCad = SqlCad & "'ADICION DE PRODUCTO EN ORDEN DE PRODUCCION DESDE CP.'"
    SqlCad = SqlCad & ")"
    
    cnn_dbbancos.Execute SqlCad
    
    Actualiza_Log SqlCad, StrConexDbBancos
    
    lngIdCambio = 0
    strIdUsuario = vbNullString
    
    
    SqlCad = vbNullString
    
    insertarProductoEnOPServidorExterno = True
    
    Exit Function
errInsertarProductoEnOPServidorExterno:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    'Resume
    insertarProductoEnOPServidorExterno = False
    
    Err.Clear
End Function

Public Function importarColorServidorExterno(ByVal frameProgreso As Object, _
                                                ByVal barraProgreso As Object) As Boolean
    
    On Error GoTo errImportarColorServidorExterno
    
    Dim rstColor As New ADODB.Recordset
    Dim strUltimaLecturaDeColores As String
    Dim dblItem As Double
    
    importarColorServidorExterno = False
    
    strUltimaLecturaDeColores = sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UltimaLecturaDeColores", "l")
    
    Rem LECTURA DE COLORES
    
    'abrirCnDBMilano
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "C.IDCOLOR, "
    SqlCad = SqlCad & "C.NOMBRE, "
    SqlCad = SqlCad & "C.IDUSUARIOINGRESO, "
    SqlCad = SqlCad & "C.FECHAINGRESO, "
    SqlCad = SqlCad & "C.IDUSUARIOACTUALIZACION, "
    SqlCad = SqlCad & "C.FECHAACTUALIZACION, "
    SqlCad = SqlCad & "C.ELIMINADO "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "COLOR AS C "
    
        If strUltimaLecturaDeColores <> vbNullString Then
            SqlCad = SqlCad & "WHERE "
            SqlCad = SqlCad & "C.FECHAACTUALIZACION >= '" & strUltimaLecturaDeColores & "'"
        End If
        
    If rstColor.State = 1 Then rstColor.Close
    
    rstColor.Open SqlCad, cnBdStudioModa, adOpenForwardOnly, adLockReadOnly
    
    If Not rstColor.EOF Then
        
        frameProgreso.Visible = True
        barraProgreso.Max = ModUtilitario.devuelveCantRegistros(rstColor)
        barraProgreso.Value = 0
        frameProgreso.Caption = "Actualizando Colores..."
        
        Do While Not rstColor.EOF
            With objAyudaBienColor
                .inicializarEntidades
                
                .Codigo = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CODIGO", "EF2BIENCOLOR", "CODEXTERNO", Trim(rstColor!IDCOLOR & ""), "T")
                .CodigoExterno = Trim(rstColor!IDCOLOR & "")
                .Descripcion = Trim(rstColor!nombre & "")
                .Estado = CBool(rstColor!ELIMINADO) 'IIf(Val(rstColor!ELIMINADO & "") = 1, True, False)
                .FechaReg = Trim(rstColor!FechaIngreso & "")
                .UsuarioReg = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODUSER", "EF2USERS", "F2CODUSEREXTERNO", Trim(rstColor!IDUSUARIOINGRESO & ""), "T")  'Trim(rstColor!IDUSUARIOINGRESO & "")
                .FechaMod = Trim(rstColor!FECHAACTUALIZACION & "")
                .UsuarioMod = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODUSER", "EF2USERS", "F2CODUSEREXTERNO", Trim(rstColor!IDUSUARIOACTUALIZACION & ""), "T")  'Trim(rstColor!IDUSUARIOACTUALIZACION & "")
                
                If .guardarBienColor Then
                    Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                End If
            End With
            
            If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                With objSqlAyudaBienColor
                    .inicializarEntidades
                    
                    .Codigo = ModUtilitario.ObtenerCampoV2(cnBdCPlus, "CODIGO", "MAESTROS.EF2BIENCOLOR", "CODEXTERNO", Trim(rstColor!IDCOLOR & ""), "T")
                    .CodigoExterno = Trim(rstColor!IDCOLOR & "")
                    .Descripcion = Trim(rstColor!nombre & "")
                    .Estado = CBool(rstColor!ELIMINADO) 'IIf(Val(rstColor!ELIMINADO & "") = 1, True, False)
                    .FechaReg = IIf(objAyudaBienColor.FechaReg = vbNullString, Format(Date, "Short Date"), Format(Trim(rstColor!FechaIngreso & ""), "Short Date"))
                    .UsuarioReg = ModUtilitario.ObtenerCampoV2(cnBdCPlus, "F2CODUSER", "MAESTROS.EF2USERS", "F2CODUSEREXTERNO", Trim(rstColor!IDUSUARIOINGRESO & ""), "T")  'Trim(rstColor!IDUSUARIOINGRESO & "")
                    .FechaMod = IIf(objAyudaBienColor.FechaReg = vbNullString, Format(Date, "Short Date"), Format(Trim(rstColor!FECHAACTUALIZACION & ""), "Short Date"))
                    .UsuarioMod = ModUtilitario.ObtenerCampoV2(cnBdCPlus, "F2CODUSER", "MAESTROS.EF2USERS", "F2CODUSEREXTERNO", Trim(rstColor!IDUSUARIOACTUALIZACION & ""), "T")  'Trim(rstColor!IDUSUARIOACTUALIZACION & "")
                    
                    If .guardarBienColor Then
                        Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                    End If
                End With
            End If
            
            rstColor.MoveNext
            
            DoEvents
            
            barraProgreso.Value = barraProgreso.Value + 1
            frameProgreso.Caption = "Actualizando Colores... " & FormatPercent(barraProgreso.Value / barraProgreso.Max, 0)
        Loop
            MsgBox "Actualización Finalizada", vbInformation + vbOKOnly, wnomcia
            
            ModUtilitario.sWrtIni App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UltimaLecturaDeColores", Format(Date, "Short Date") '"04/07/2014"
    Else
        MsgBox "No se ubicaron Colores nuevos y/o modificados.", vbInformation + vbOKOnly, wnomcia
    End If
    
    'cnBdStudioModa.Close
    
    frameProgreso.Visible = False
    
    'If cnBdStudioModa.State = 1 Then cnBdStudioModa.Close
    If rstColor.State = 1 Then rstColor.Close
    If rstColor.State = 1 Then rstColor.Close
    
    'Set cnBdStudioModa = Nothing
    Set rstColor = Nothing
    Set rstColor = Nothing
    
    strUltimaLecturaDeColores = vbNullString
    dblItem = 0
    SqlCad = vbNullString
    
    Exit Function
errImportarColorServidorExterno:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    'Resume
    importarColorServidorExterno = False
    
    frameProgreso.Visible = False
    barraProgreso.Value = 0
    
    Err.Clear
End Function

Public Function importarUMServidorExterno(ByVal frameProgreso As Object, _
                                                ByVal barraProgreso As Object) As Boolean
    
    On Error GoTo errImportarUMServidorExterno
    
    Dim rstUM As New ADODB.Recordset
    Dim strUltimaLecturaDeUM As String
    Dim dblItem As Double
    
    importarUMServidorExterno = False
    
    strUltimaLecturaDeUM = sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UltimaLecturaDeUM", "l")
    
    Rem LECTURA DE COLORES
    
    'abrirCnDBMilano
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "IDUNIDADMEDIDA, "
    SqlCad = SqlCad & "NOMBRE, "
    SqlCad = SqlCad & "ABREVIATURA, "
    SqlCad = SqlCad & "ELIMINADO, "
    SqlCad = SqlCad & "IDUSUARIOINGRESO, "
    SqlCad = SqlCad & "FECHAINGRESO, "
    SqlCad = SqlCad & "IDUSUARIOACTUALIZACION, "
    SqlCad = SqlCad & "FECHAACTUALIZACION "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "UNIDADMEDIDA "
    
        If strUltimaLecturaDeUM <> vbNullString Then
            SqlCad = SqlCad & "WHERE "
            SqlCad = SqlCad & "FECHAACTUALIZACION >= '" & strUltimaLecturaDeUM & "'"
        End If
        
    If rstUM.State = 1 Then rstUM.Close
    
    rstUM.Open SqlCad, cnBdStudioModa, adOpenForwardOnly, adLockReadOnly
    
    If Not rstUM.EOF Then
        
        frameProgreso.Visible = True
        barraProgreso.Max = ModUtilitario.devuelveCantRegistros(rstUM)
        barraProgreso.Value = 0
        frameProgreso.Caption = "Actualizando U.M..."
        
        Do While Not rstUM.EOF
            With objAyudaUM
                .inicializarEntidades
                
                .Codigo = Trim(rstUM!IDUNIDADMEDIDA & "")
                .Descripcion = Trim(rstUM!nombre & "")
                .Abreviatura = Trim(rstUM!Abreviatura & "")
                .Estado = Not CBool(rstUM!ELIMINADO) 'IIf(Val(rstUM!ELIMINADO & "") = 1, True, False)
                .FechaReg = Trim(rstUM!FechaIngreso & "")
                .UsuarioReg = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODUSER", "EF2USERS", "F2CODUSEREXTERNO", Trim(rstUM!IDUSUARIOINGRESO & ""), "T")  'Trim(rstUM!IDUSUARIOINGRESO & "")
                .FechaMod = Trim(rstUM!FECHAACTUALIZACION & "")
                .UsuarioMod = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODUSER", "EF2USERS", "F2CODUSEREXTERNO", Trim(rstUM!IDUSUARIOACTUALIZACION & ""), "T")  'Trim(rstUM!IDUSUARIOACTUALIZACION & "")
                
                If .guardarUM Then
                    Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                End If
            End With
            
            If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                With objSqlAyudaUM
                    .inicializarEntidades
                    
                    .Codigo = Trim(rstUM!IDUNIDADMEDIDA & "")
                    .Descripcion = Trim(rstUM!nombre & "")
                    .Abreviatura = Trim(rstUM!Abreviatura & "")
                    .Estado = Not CBool(rstUM!ELIMINADO) 'IIf(Val(rstUM!ELIMINADO & "") = 1, True, False)
                    .FechaReg = IIf(objAyudaFamilia.FechaReg = vbNullString, Format(Date, "Short Date"), Format(Trim(rstUM!FechaIngreso & ""), "Short Date"))
                    .UsuarioReg = ModUtilitario.ObtenerCampoV2(cnBdCPlus, "F2CODUSER", "MAESTROS.EF2USERS", "F2CODUSEREXTERNO", Trim(rstUM!IDUSUARIOINGRESO & ""), "T")   'Trim(rstUM!IDUSUARIOINGRESO & "")
                    .FechaMod = IIf(objAyudaFamilia.FechaReg = vbNullString, Format(Date, "Short Date"), Format(Trim(rstUM!FECHAACTUALIZACION & ""), "Short Date"))
                    .UsuarioMod = ModUtilitario.ObtenerCampoV2(cnBdCPlus, "F2CODUSER", "MAESTROS.EF2USERS", "F2CODUSEREXTERNO", Trim(rstUM!IDUSUARIOACTUALIZACION & ""), "T")  'Trim(rstUM!IDUSUARIOACTUALIZACION & "")
                    
                    If .guardarUM Then
                        Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                    End If
                End With
            End If
            
            rstUM.MoveNext
            
            DoEvents
            
            barraProgreso.Value = barraProgreso.Value + 1
            frameProgreso.Caption = "Actualizando U.M... " & FormatPercent(barraProgreso.Value / barraProgreso.Max, 0)
        Loop
            MsgBox "Actualización Finalizada", vbInformation + vbOKOnly, wnomcia
            
            ModUtilitario.sWrtIni App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UltimaLecturaDeUM", Format(Date, "Short Date") '"04/07/2014"
    Else
        MsgBox "No se ubicaron U.M. nuevos y/o modificados.", vbInformation + vbOKOnly, wnomcia
    End If
    
    'cnBdStudioModa.Close
    
    frameProgreso.Visible = False
    
    'If cnBdStudioModa.State = 1 Then cnBdStudioModa.Close
    If rstUM.State = 1 Then rstUM.Close
    If rstUM.State = 1 Then rstUM.Close
    
    'Set cnBdStudioModa = Nothing
    Set rstUM = Nothing
    Set rstUM = Nothing
    
    strUltimaLecturaDeUM = vbNullString
    dblItem = 0
    SqlCad = vbNullString
    
    Exit Function
errImportarUMServidorExterno:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    'Resume
    importarUMServidorExterno = False
    
    frameProgreso.Visible = False
    barraProgreso.Value = 0
    
    Err.Clear
End Function

Public Function importarFamiliaServidorExterno(ByVal frameProgreso As Object, _
                                                ByVal barraProgreso As Object) As Boolean
    
    On Error GoTo errImportarFamiliaServidorExterno
    
    Dim rstFamilia As New ADODB.Recordset
    Dim strUltimaLecturaDeFamilia As String
    Dim dblItem As Double
    
    importarFamiliaServidorExterno = False
    
    strUltimaLecturaDeFamilia = sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UltimaLecturaDeFamilia", "l")
    
    Rem LECTURA DE FAMILIA
    
    'abrirCnDBMilano
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "IDFAMILIA, "
    SqlCad = SqlCad & "NOMBRE, "
    SqlCad = SqlCad & "ELIMINADO, "
    SqlCad = SqlCad & "FECHAINGRESO, "
    SqlCad = SqlCad & "IDUSUARIOINGRESO, "
    SqlCad = SqlCad & "FECHAACTUALIZACION, "
    SqlCad = SqlCad & "IDUSUARIOACTUALIZACION "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "FAMILIA "
    
        If strUltimaLecturaDeFamilia <> vbNullString Then
            SqlCad = SqlCad & "WHERE "
            SqlCad = SqlCad & "FECHAACTUALIZACION >= '" & strUltimaLecturaDeFamilia & "'"
        End If
        
    If rstFamilia.State = 1 Then rstFamilia.Close
    
    rstFamilia.Open SqlCad, cnBdStudioModa, adOpenForwardOnly, adLockReadOnly
    
    If Not rstFamilia.EOF Then
        
        frameProgreso.Visible = True
        barraProgreso.Max = ModUtilitario.devuelveCantRegistros(rstFamilia)
        barraProgreso.Value = 0
        frameProgreso.Caption = "Actualizando Familias..."
        
        Do While Not rstFamilia.EOF
            With objAyudaFamilia
                .inicializarEntidades
                
                .Codigo = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F7CODCON", "SF7NIVEL01", "CODEXTERNO", Trim(rstFamilia!IDFAMILIA & ""), "T")
                .CodigoExterno = Trim(rstFamilia!IDFAMILIA & "")
                .Descripcion = Trim(rstFamilia!nombre & "")
                .DescripcionCorta = Trim(rstFamilia!nombre & "")
                .Estado = Not CBool(rstFamilia!ELIMINADO) 'IIf(Val(rstFamilia!ELIMINADO & "") = 1, True, False)
                .FechaReg = Trim(rstFamilia!FechaIngreso & "")
                .UsuarioReg = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODUSER", "EF2USERS", "F2CODUSEREXTERNO", Trim(rstFamilia!IDUSUARIOINGRESO & ""), "T")  'Trim(rstFamilia!IDUSUARIOINGRESO & "")
                .FechaMod = Trim(rstFamilia!FECHAACTUALIZACION & "")
                .UsuarioMod = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODUSER", "EF2USERS", "F2CODUSEREXTERNO", Trim(rstFamilia!IDUSUARIOACTUALIZACION & ""), "T")  'Trim(rstFamilia!IDUSUARIOACTUALIZACION & "")
                
                If .guardarFamilia Then
                    Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                End If
            End With
            
            If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                With objSqlAyudaFamilia
                    .inicializarEntidades
                    
                    .Codigo = ModUtilitario.ObtenerCampoV2(cnBdCPlus, "F7CODCON", "MAESTROS.SF7NIVEL01", "CODEXTERNO", Trim(rstFamilia!IDFAMILIA & ""), "T")
                    .CodigoExterno = Trim(rstFamilia!IDFAMILIA & "")
                    .Descripcion = Trim(rstFamilia!nombre & "")
                    .DescripcionCorta = Trim(rstFamilia!nombre & "")
                    .Estado = Not CBool(rstFamilia!ELIMINADO) 'IIf(Val(rstFamilia!ELIMINADO & "") = 1, True, False)
                    .FechaReg = IIf(objAyudaFamilia.FechaReg = vbNullString, Format(Date, "Short Date"), Format(Trim(rstFamilia!FechaIngreso & ""), "Short Date"))
                    .UsuarioReg = ModUtilitario.ObtenerCampoV2(cnBdCPlus, "F2CODUSER", "MAESTROS.EF2USERS", "F2CODUSEREXTERNO", Trim(rstFamilia!IDUSUARIOINGRESO & ""), "T")  'Trim(rstFamilia!IDUSUARIOINGRESO & "")
                    .FechaMod = IIf(objAyudaFamilia.FechaReg = vbNullString, Format(Date, "Short Date"), Format(Trim(rstFamilia!FECHAACTUALIZACION & ""), "Short Date"))
                    .UsuarioMod = ModUtilitario.ObtenerCampoV2(cnBdCPlus, "F2CODUSER", "MAESTROS.EF2USERS", "F2CODUSEREXTERNO", Trim(rstFamilia!IDUSUARIOACTUALIZACION & ""), "T")  'Trim(rstFamilia!IDUSUARIOACTUALIZACION & "")
                    
                    If .guardarFamilia Then
                        Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                    End If
                End With
            End If
            
            rstFamilia.MoveNext
            
            DoEvents
            
            barraProgreso.Value = barraProgreso.Value + 1
            frameProgreso.Caption = "Actualizando Familias... " & FormatPercent(barraProgreso.Value / barraProgreso.Max, 0)
        Loop
            MsgBox "Actualización Finalizada", vbInformation + vbOKOnly, wnomcia
            
            ModUtilitario.sWrtIni App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UltimaLecturaDeFamilia", Format(Date, "Short Date") '"04/07/2014"
    Else
        MsgBox "No se ubicaron Familias nuevas y/o modificadas.", vbInformation + vbOKOnly, wnomcia
    End If
    
    'cnBdStudioModa.Close
    
    frameProgreso.Visible = False
    
    'If cnBdStudioModa.State = 1 Then cnBdStudioModa.Close
    If rstFamilia.State = 1 Then rstFamilia.Close
    
    'Set cnBdStudioModa = Nothing
    Set rstFamilia = Nothing
    
    strUltimaLecturaDeFamilia = vbNullString
    dblItem = 0
    SqlCad = vbNullString
    
    Exit Function
errImportarFamiliaServidorExterno:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    'Resume
    importarFamiliaServidorExterno = False
    
    frameProgreso.Visible = False
    barraProgreso.Value = 0
    
    Err.Clear
End Function

Public Function importarSubFamiliaServidorExterno(ByVal frameProgreso As Object, _
                                                ByVal barraProgreso As Object) As Boolean
    
    On Error GoTo errImportarSubFamiliaServidorExterno
    
    Dim rstSubFamilia As New ADODB.Recordset
    Dim strUltimaLecturaDeSubFamilia As String
    Dim dblItem As Double
    
    importarSubFamiliaServidorExterno = False
    
    strUltimaLecturaDeSubFamilia = sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UltimaLecturaDeSubFamilia", "l")
    
    Rem LECTURA DE SUBFAMILIA
    
    'abrirCnDBMilano
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "IDSUBFAMILIA, "
    SqlCad = SqlCad & "IDFAMILIA, "
    SqlCad = SqlCad & "NOMBRE, "
    SqlCad = SqlCad & "ELIMINADO, "
    SqlCad = SqlCad & "FECHAINGRESO, "
    SqlCad = SqlCad & "IDUSUARIOINGRESO, "
    SqlCad = SqlCad & "FECHAACTUALIZACION, "
    SqlCad = SqlCad & "IDUSUARIOACTUALIZACION "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "SUBFAMILIA "
    
        If strUltimaLecturaDeSubFamilia <> vbNullString Then
            SqlCad = SqlCad & "WHERE "
            SqlCad = SqlCad & "FECHAACTUALIZACION >= '" & strUltimaLecturaDeSubFamilia & "'"
        End If
        
    If rstSubFamilia.State = 1 Then rstSubFamilia.Close
    
    rstSubFamilia.Open SqlCad, cnBdStudioModa, adOpenForwardOnly, adLockReadOnly
    
    If Not rstSubFamilia.EOF Then
        
        frameProgreso.Visible = True
        barraProgreso.Max = ModUtilitario.devuelveCantRegistros(rstSubFamilia)
        barraProgreso.Value = 0
        frameProgreso.Caption = "Actualizando Sub-Familias..."
        
        Do While Not rstSubFamilia.EOF
            With objAyudaSubFamilia
                .inicializarEntidades
                
                .Codigo = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F7CODCON", "SF7NIVEL02", "CODEXTERNO", Trim(rstSubFamilia!IDSUBFAMILIA & ""), "T")
                .CodigoExterno = Trim(rstSubFamilia!IDSUBFAMILIA & "")
                .CodigoFamilia = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F7CODCON", "SF7NIVEL01", "CODEXTERNO", Trim(rstSubFamilia!IDFAMILIA & ""), "T")
                .Descripcion = Trim(rstSubFamilia!nombre & "")
                .Estado = Not CBool(rstSubFamilia!ELIMINADO) 'IIf(Val(rstSubFamilia!ELIMINADO & "") = 1, True, False)
                .FechaReg = Trim(rstSubFamilia!FechaIngreso & "")
                .UsuarioReg = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODUSER", "EF2USERS", "F2CODUSEREXTERNO", Trim(rstSubFamilia!IDUSUARIOINGRESO & ""), "T")  'Trim(rstSubFamilia!IDUSUARIOINGRESO & "")
                .FechaMod = Trim(rstSubFamilia!FECHAACTUALIZACION & "")
                .UsuarioMod = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODUSER", "EF2USERS", "F2CODUSEREXTERNO", Trim(rstSubFamilia!IDUSUARIOACTUALIZACION & ""), "T")  'Trim(rstSubFamilia!IDUSUARIOACTUALIZACION & "")
                
                If .guardarSubFamilia Then
                    Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                End If
            End With
            
            If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                With objSqlAyudaSubFamilia
                .inicializarEntidades
                
                .Codigo = ModUtilitario.ObtenerCampoV2(cnBdCPlus, "F7CODCON", "MAESTROS.SF7NIVEL02", "CODEXTERNO", Trim(rstSubFamilia!IDSUBFAMILIA & ""), "T")
                .CodigoExterno = Trim(rstSubFamilia!IDSUBFAMILIA & "")
                .CodigoFamilia = ModUtilitario.ObtenerCampoV2(cnBdCPlus, "F7CODCON", "MAESTROS.SF7NIVEL01", "CODEXTERNO", Trim(rstSubFamilia!IDFAMILIA & ""), "T")
                .Descripcion = Trim(rstSubFamilia!nombre & "")
                .Estado = Not CBool(rstSubFamilia!ELIMINADO) 'IIf(Val(rstSubFamilia!ELIMINADO & "") = 1, True, False)
                .FechaReg = IIf(objAyudaFamilia.FechaReg = vbNullString, Format(Date, "Short Date"), Format(Trim(rstSubFamilia!FechaIngreso & ""), "Short Date"))
                .UsuarioReg = ModUtilitario.ObtenerCampoV2(cnBdCPlus, "F2CODUSER", "MAESTROS.EF2USERS", "F2CODUSEREXTERNO", Trim(rstSubFamilia!IDUSUARIOINGRESO & ""), "T")  'Trim(rstSubFamilia!IDUSUARIOINGRESO & "")
                .FechaMod = IIf(objAyudaFamilia.FechaReg = vbNullString, Format(Date, "Short Date"), Format(Trim(rstSubFamilia!FECHAACTUALIZACION & ""), "Short Date"))
                .UsuarioMod = ModUtilitario.ObtenerCampoV2(cnBdCPlus, "F2CODUSER", "MAESTROS.EF2USERS", "F2CODUSEREXTERNO", Trim(rstSubFamilia!IDUSUARIOACTUALIZACION & ""), "T")  'Trim(rstSubFamilia!IDUSUARIOACTUALIZACION & "")
                
                If .guardarSubFamilia Then
                    Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                End If
            End With
            End If
            
            rstSubFamilia.MoveNext
            
            DoEvents
            
            barraProgreso.Value = barraProgreso.Value + 1
            frameProgreso.Caption = "Actualizando Sub-Familias... " & FormatPercent(barraProgreso.Value / barraProgreso.Max, 0)
        Loop
            MsgBox "Actualización Finalizada", vbInformation + vbOKOnly, wnomcia
            
            ModUtilitario.sWrtIni App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UltimaLecturaDeSubFamilia", Format(Date, "Short Date") '"04/07/2014"
    Else
        MsgBox "No se ubicaron Sub-Familias nuevas y/o modificadas.", vbInformation + vbOKOnly, wnomcia
    End If
    
    'cnBdStudioModa.Close
    
    frameProgreso.Visible = False
    
    'If cnBdStudioModa.State = 1 Then cnBdStudioModa.Close
    If rstSubFamilia.State = 1 Then rstSubFamilia.Close
    If rstSubFamilia.State = 1 Then rstSubFamilia.Close
    
    'Set cnBdStudioModa = Nothing
    Set rstSubFamilia = Nothing
    Set rstSubFamilia = Nothing
    
    strUltimaLecturaDeSubFamilia = vbNullString
    dblItem = 0
    SqlCad = vbNullString
    
    Exit Function
errImportarSubFamiliaServidorExterno:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    'Resume
    importarSubFamiliaServidorExterno = False
    
    frameProgreso.Visible = False
    barraProgreso.Value = 0
    
    Err.Clear
End Function


Public Function importarPersonasServidorExterno(ByVal frameProgreso As Object, _
                                                ByVal barraProgreso As Object) As Boolean
    
    On Error GoTo errImportarPersonasServidorExterno
    
    Dim rstPersona As New ADODB.Recordset
    Dim strUltimaLecturaDePersonas As String
    Dim strCodAuxiliar As String
    Dim strNroDocumentoID As String
    Dim dblItem As Double
    
    importarPersonasServidorExterno = False
    
    strUltimaLecturaDePersonas = sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UltimaLecturaDePersonas", "l")
    
    Rem LECTURA DE INSUMO
    
    'abrirCnDBMilano
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "PER.IDPERSONA, "
    SqlCad = SqlCad & "TPER.NOMBRE AS TIPOPERSONA, "
    SqlCad = SqlCad & "PER.NOMBRE, "
    SqlCad = SqlCad & "PER.RUC, "
    SqlCad = SqlCad & "PER.DIRECCION, "
    SqlCad = SqlCad & "PER.DNI, "
    SqlCad = SqlCad & "PER.EMAIL, "
    SqlCad = SqlCad & "PER.EXTRANJERO "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "PERSONA AS PER "
    SqlCad = SqlCad & "LEFT JOIN TIPOPERSONA AS TPER ON TPER.IDTIPOPERSONA = PER.IDTIPOPERSONA "
        
        If strUltimaLecturaDePersonas <> vbNullString Then
            SqlCad = SqlCad & "WHERE "
            SqlCad = SqlCad & "PER.FECHAACTUALIZACION >= '" & strUltimaLecturaDePersonas & "'"
        End If
    
    If rstPersona.State = 1 Then rstPersona.Close
    
    rstPersona.Open SqlCad, cnBdStudioModa, adOpenForwardOnly, adLockReadOnly
    
    If Not rstPersona.EOF Then
        
        frameProgreso.Visible = True
        barraProgreso.Max = ModUtilitario.devuelveCantRegistros(rstPersona)
        barraProgreso.Value = 0
        frameProgreso.Caption = "Actualizando Clientes / Proveedores..."
        
        Do While Not rstPersona.EOF
            SqlCad = vbNullString
            
            Select Case Trim(rstPersona!TipoPersona & "")
                Case "PROVEEDOR"
                    strCodAuxiliar = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODPROV", "EF2PROVEEDORES", "F2CODPROVEXTERNO", Trim(rstPersona!IDPERSONA & ""), "T")
                    
                    If strCodAuxiliar = vbNullString Then
                        strCodAuxiliar = Format(Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "TOP 1 F2CODPROV", "EF2PROVEEDORES", vbNullString, vbNullString, vbNullString, "TRIM(F2CODPROV & '') <> '' ORDER BY F2CODPROV DESC")) + 1, "0000")
                        
                        strNroDocumentoID = IIf(Len(Trim(rstPersona!ruc & "")) = 11, Trim(rstPersona!ruc & ""), IIf(Len(Trim(rstPersona!DNI & "")) = 8, Trim(rstPersona!DNI & ""), vbNullString))
                        
                        If strNroDocumentoID <> vbNullString And Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "COUNT(*)", "EF2PROVEEDORES", "F2NEWRUC", strNroDocumentoID, "T")) = 1 Then
                            strNroDocumentoID = vbNullString
                        End If
                        
                        If strNroDocumentoID = vbNullString Then
                            strNroDocumentoID = Format(Val(strCodAuxiliar), "00000000000")
                        End If
                        
                        SqlCad = vbNullString
                        SqlCad = SqlCad & "INSERT INTO EF2PROVEEDORES("
                        SqlCad = SqlCad & "F2CODPROV, F2CODPROVEXTERNO, F2NEWRUC, "
                        SqlCad = SqlCad & "F2NOMPROV, F2DIRPROV, INTCODCATEGORIA, "
                        SqlCad = SqlCad & "F2EMAIL, F2TIPMON, F2FORPAG, F2ORDEN, "
                        SqlCad = SqlCad & "F2TIPPROV, F2TIPDOC) "
                        SqlCad = SqlCad & "VALUES("
                        SqlCad = SqlCad & "'" & strCodAuxiliar & "', "
                        SqlCad = SqlCad & "'" & Trim(rstPersona!IDPERSONA & "") & "', "
                        SqlCad = SqlCad & "'" & strNroDocumentoID & "', "
                        SqlCad = SqlCad & "'" & Replace(Trim(rstPersona!nombre & ""), "'", "' & Chr(39) & '", 1) & "', "
                        SqlCad = SqlCad & "'" & Replace(Trim(rstPersona!Direccion & ""), "'", "' & Chr(39) & '", 1) & "', "
                        SqlCad = SqlCad & "1, "
                        SqlCad = SqlCad & "'" & left(Trim(rstPersona!Email & ""), 50) & "', "
                        SqlCad = SqlCad & "'" & IIf(CBool(rstPersona!EXTRANJERO), "D", "S") & "', "
                        SqlCad = SqlCad & "'001', "
                        SqlCad = SqlCad & "TRUE, "
                        SqlCad = SqlCad & "'" & IIf(CBool(rstPersona!EXTRANJERO), "E", "N") & "', "
                        SqlCad = SqlCad & "'03')"
                    End If
                Case Else
                    strCodAuxiliar = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODCLIEXTERNO", "EF2CLIENTES", "F2CODCLIEXTERNO", Trim(rstPersona!IDPERSONA & ""), "T")
                    
                    If strCodAuxiliar = vbNullString Then
                        strCodAuxiliar = Format(Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "TOP 1 F2CODCLI", "EF2CLIENTES", vbNullString, vbNullString, vbNullString, "TRIM(F2CODCLI & '') <> '' ORDER BY F2CODCLI DESC")) + 1, "0000")
                        
                        strNroDocumentoID = IIf(Len(Trim(rstPersona!ruc & "")) = 11, Trim(rstPersona!ruc & ""), IIf(Len(Trim(rstPersona!DNI & "")) = 8, Trim(rstPersona!DNI & ""), vbNullString))
                        
                        If strNroDocumentoID <> vbNullString And Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "COUNT(*)", "EF2CLIENTES", "F2NEWRUC", strNroDocumentoID, "T")) = 1 Then
                            strNroDocumentoID = vbNullString
                        End If
                        
                        If strNroDocumentoID = vbNullString Then
                            strNroDocumentoID = Format(Val(strCodAuxiliar), "00000000000")
                        End If
                        
                        SqlCad = vbNullString
                        SqlCad = SqlCad & "INSERT INTO EF2CLIENTES("
                        SqlCad = SqlCad & "F2CODCLI, F2CODCLIEXTERNO, F2NEWRUC, "
                        SqlCad = SqlCad & "F2NOMCLI, F2DIRCLI, F2DOCCLI, F2TIPDOC, "
                        SqlCad = SqlCad & "F2EMAIL, F2FORPAG, F2TIPOCLI, X) "
                        SqlCad = SqlCad & "VALUES("
                        SqlCad = SqlCad & "'" & strCodAuxiliar & "', "
                        SqlCad = SqlCad & "'" & Trim(rstPersona!IDPERSONA & "") & "', "
                        SqlCad = SqlCad & "'" & strNroDocumentoID & "', "
                        SqlCad = SqlCad & "'" & Replace(Trim(rstPersona!nombre & ""), "'", "' & Chr(39) & '", 1) & "', "
                        SqlCad = SqlCad & "'" & Replace(Trim(rstPersona!Direccion & ""), "'", "' & Chr(39) & '", 1) & "', "
                        SqlCad = SqlCad & "'" & IIf(Len(strNroDocumentoID) = 11, "6", IIf(Len(strNroDocumentoID) = 8, "1", "0")) & "', "
                        SqlCad = SqlCad & "'" & IIf(Len(strNroDocumentoID) = 11, "J", IIf(Len(strNroDocumentoID) = 8, "N", "J")) & "', "
                        SqlCad = SqlCad & "'" & left(Trim(rstPersona!Email & ""), 50) & "', "
                        SqlCad = SqlCad & "'001', "
                        SqlCad = SqlCad & "'" & IIf(Len(strNroDocumentoID) = 11, "N", IIf(Len(strNroDocumentoID) = 8, "N", "E")) & "', "
                        SqlCad = SqlCad & "'" & Trim(rstPersona!TipoPersona & "") & "')"
                    End If
            End Select
            
            If SqlCad <> vbNullString Then
                cnn_dbbancos.Execute SqlCad
                
                Actualiza_Log SqlCad, StrConexDbBancos
            End If
            
            rstPersona.MoveNext
            
            DoEvents
            
            barraProgreso.Value = barraProgreso.Value + 1
            frameProgreso.Caption = "Actualizando Clientes / Proveedores... " & FormatPercent(barraProgreso.Value / barraProgreso.Max, 0)
        Loop
            MsgBox "Actualización de Personas Finalizada.", vbInformation + vbOKOnly, wnomcia
            
            ModUtilitario.sWrtIni App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UltimaLecturaDePersonas", Format(Date, "Short Date") '"04/07/2014"
    Else
        MsgBox "No se ubicaron Personas nuevas y/o modificadas.", vbInformation + vbOKOnly, wnomcia
    End If
    
    'cnBdStudioModa.Close
    
    frameProgreso.Visible = False
    
    'If cnBdStudioModa.State = 1 Then cnBdStudioModa.Close
    If rstPersona.State = 1 Then rstPersona.Close
    If rstPersona.State = 1 Then rstPersona.Close
    
    'Set cnBdStudioModa = Nothing
    Set rstPersona = Nothing
    Set rstPersona = Nothing
    
    strUltimaLecturaDePersonas = vbNullString
    dblItem = 0
    SqlCad = vbNullString
    
    Exit Function
errImportarPersonasServidorExterno:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    'Resume
    importarPersonasServidorExterno = False
    
    frameProgreso.Visible = False
    barraProgreso.Value = 0
    
    Err.Clear
End Function

Public Function importarProductoPorProveedorServidorExterno(ByVal frameProgreso As Object, _
                                                ByVal barraProgreso As Object) As Boolean
    
    On Error GoTo errImportarProductoPorProveedorServidorExterno
    
    Dim rstHistorial As New ADODB.Recordset
    Dim strUltimaLecturaDeProductoPorProveedor As String
    Dim strCodAuxiliar As String
    Dim strNroDocumentoID As String
    Dim dblItem As Double
    Dim strRucProveedor As String
    
    importarProductoPorProveedorServidorExterno = False
    
    strUltimaLecturaDeProductoPorProveedor = sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UltimaLecturaDeProductoPorProveedor", "l")
    
    Rem LECTURA DE HISTORIAL
    
    'abrirCnDBMilano
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT DISTINCT "
    SqlCad = SqlCad & "CAB.IDPERSONA AS F2CODPRV, "
    SqlCad = SqlCad & "PER.NOMBRE AS F2NOMPRV, "
    SqlCad = SqlCad & "DET.IDINSUMO AS F5CODPRO, "
    SqlCad = SqlCad & "INS.NOMBRE AS F5NOMPRO, "
    SqlCad = SqlCad & "DET.COSTO AS F5VALVTA, "
    SqlCad = SqlCad & "'' AS F5CODFAB, "
    SqlCad = SqlCad & "INS.IDUNIDADMEDIDA AS F7CODMED, "
    SqlCad = SqlCad & "0 AS F5STOCKACT, "
    SqlCad = SqlCad & "'' AS F2FORPAG, "
    SqlCad = SqlCad & "'' AS F2COND_PAGO, "
    SqlCad = SqlCad & "MAX(CONVERT(CHAR(10), CAB.FECHA, 103)) AS F2FECHA, "
    SqlCad = SqlCad & "LEFT(MON.NOMBRE,1) AS F2MONEDA "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "INGRESO AS CAB "
    SqlCad = SqlCad & "LEFT JOIN INGRESODETALLE AS DET ON DET.IDINGRESO = CAB.IDINGRESO "
    SqlCad = SqlCad & "LEFT JOIN PERSONA AS PER ON PER.IDPERSONA = CAB.IDPERSONA "
    SqlCad = SqlCad & "LEFT JOIN INSUMO AS INS ON INS.IDINSUMO = DET.IDINSUMO "
    SqlCad = SqlCad & "LEFT JOIN MONEDA AS MON ON MON.IDMONEDA = CAB.IDMONEDA "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "CAB.ANULADO = 0 AND "
    SqlCad = SqlCad & "PER.ELIMINADO = 0 AND "
    SqlCad = SqlCad & "RTRIM(LTRIM(CAB.IDPERSONA)) <> '' AND "
    SqlCad = SqlCad & "PER.IDTIPOPERSONA NOT IN ('TIP0000002') AND "
    SqlCad = SqlCad & "CAB.IDTIPOINGRESO IN ('TIP0000001') "
        
        If strUltimaLecturaDeProductoPorProveedor <> vbNullString Then
            SqlCad = SqlCad & "AND CAB.FECHAACTUALIZACION >= '" & strUltimaLecturaDeProductoPorProveedor & "' "
        End If
        
    SqlCad = SqlCad & "GROUP BY "
    SqlCad = SqlCad & "CAB.IDPERSONA, "
    SqlCad = SqlCad & "PER.NOMBRE, "
    SqlCad = SqlCad & "DET.IDINSUMO, "
    SqlCad = SqlCad & "INS.NOMBRE, "
    SqlCad = SqlCad & "DET.COSTO, "
    SqlCad = SqlCad & "INS.IDUNIDADMEDIDA, "
    SqlCad = SqlCad & "LEFT(MON.NOMBRE,1) "
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "CAB.IDPERSONA, "
    SqlCad = SqlCad & "DET.IDINSUMO, "
    SqlCad = SqlCad & "MAX(CONVERT(CHAR(10), CAB.FECHA, 103)) ASC"
    
    If rstHistorial.State = 1 Then rstHistorial.Close
    
    rstHistorial.Open SqlCad, cnBdStudioModa, adOpenForwardOnly, adLockReadOnly
    
    If Not rstHistorial.EOF Then
        
        frameProgreso.Visible = True
        barraProgreso.Max = ModUtilitario.devuelveCantRegistros(rstHistorial)
        barraProgreso.Value = 0
        frameProgreso.Caption = "Actualizando Historial de Proveedores..."
        
        Do While Not rstHistorial.EOF
            SqlCad = vbNullString
            
            strRucProveedor = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NEWRUC", "EF2PROVEEDORES", "F2CODPROVEXTERNO", Trim(rstHistorial!F2CODPRV & ""), "T")
            
            If ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5CODPRO", "EF2PROD_PROV", "F2CODPRV", strRucProveedor, "T", "AND F5CODPRO = '" & Trim(rstHistorial!f5codpro & "") & "'") = vbNullString Then
                SqlCad = SqlCad & "INSERT INTO EF2PROD_PROV("
                SqlCad = SqlCad & "F2CODPRV, F2NOMPRV, F5CODPRO, F5NOMPRO, F5VALVTA, "
                SqlCad = SqlCad & "F7CODMED, F2FECHA, F2MONEDA) "
                SqlCad = SqlCad & "VALUES ("
                SqlCad = SqlCad & "'" & strRucProveedor & "', "
                SqlCad = SqlCad & "'" & Trim(rstHistorial!F2NOMPRV & "") & "', "
                SqlCad = SqlCad & "'" & Trim(rstHistorial!f5codpro & "") & "', "
                SqlCad = SqlCad & "'" & Trim(rstHistorial!F5NOMPRO & "") & "', "
                SqlCad = SqlCad & Val(rstHistorial!F5VALVTA & "") & ", "
                SqlCad = SqlCad & "'" & Trim(rstHistorial!f7codmed & "") & "', "
                SqlCad = SqlCad & "CVDATE('" & Format(Trim(rstHistorial!F2FECHA & ""), "Short Date") & "'), "
                SqlCad = SqlCad & "'" & Trim(rstHistorial!F2MONEDA & "") & "')"
            Else
                SqlCad = SqlCad & "UPDATE EF2PROD_PROV "
                SqlCad = SqlCad & "SET "
                SqlCad = SqlCad & "F2NOMPRV = '" & Trim(rstHistorial!F2NOMPRV & "") & "', "
                SqlCad = SqlCad & "F5NOMPRO = '" & Trim(rstHistorial!F5NOMPRO & "") & "', "
                SqlCad = SqlCad & "F5VALVTA = " & Val(rstHistorial!F5VALVTA & "") & ", "
                SqlCad = SqlCad & "F7CODMED = '" & Trim(rstHistorial!f7codmed & "") & "', "
                SqlCad = SqlCad & "F2FECHA = CVDATE('" & Format(Trim(rstHistorial!F2FECHA & ""), "Short Date") & "'), "
                SqlCad = SqlCad & "F2MONEDA = '" & Trim(rstHistorial!F2MONEDA & "") & "' "
                SqlCad = SqlCad & "WHERE "
                SqlCad = SqlCad & "F2CODPRV = '" & strRucProveedor & "' AND "
                SqlCad = SqlCad & "F5CODPRO = '" & Trim(rstHistorial!f5codpro & "") & "' AND "
                SqlCad = SqlCad & "F2FECHA <= CVDATE('" & Format(Trim(rstHistorial!F2FECHA & ""), "Short Date") & "')"
            End If
            
            If SqlCad <> vbNullString Then
                cnn_dbbancos.Execute SqlCad
                
                Actualiza_Log SqlCad, StrConexDbBancos
            End If
            
            rstHistorial.MoveNext
            
            DoEvents
            
            barraProgreso.Value = barraProgreso.Value + 1
            frameProgreso.Caption = "Actualizando Historial de Proveedores... " & FormatPercent(barraProgreso.Value / barraProgreso.Max, 0)
        Loop
            MsgBox "Actualización de Producto Por Proveedor Finalizada.", vbInformation + vbOKOnly, wnomcia
            
            ModUtilitario.sWrtIni App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UltimaLecturaDeProductoPorProveedor", Format(Date, "Short Date") '"04/07/2014"
    Else
        MsgBox "No se ubico Historial nuevo y/o modificado.", vbInformation + vbOKOnly, wnomcia
    End If
    
    'cnBdStudioModa.Close
    
    frameProgreso.Visible = False
    
    'If cnBdStudioModa.State = 1 Then cnBdStudioModa.Close
    If rstHistorial.State = 1 Then rstHistorial.Close
    If rstHistorial.State = 1 Then rstHistorial.Close
    
    'Set cnBdStudioModa = Nothing
    Set rstHistorial = Nothing
    Set rstHistorial = Nothing
    
    strUltimaLecturaDeProductoPorProveedor = vbNullString
    dblItem = 0
    SqlCad = vbNullString
    
    Exit Function
errImportarProductoPorProveedorServidorExterno:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    importarProductoPorProveedorServidorExterno = False
    
    frameProgreso.Visible = False
    barraProgreso.Value = 0
    
    Err.Clear
End Function

Public Sub importarCierreMesValesInicialesServidorExterno(ByVal frameProgreso As Object, _
                                                            ByVal barraProgreso As Object)
    
    On Error GoTo errImportarCierreMesValesInicialesServidorExterno
    
    Dim rstValeCab As ADODB.Recordset
    Dim rstValeDet As ADODB.Recordset
    
    Dim strFechaCorteInicialDeValesParaCP As String
    Dim intAnnoCorte As Integer
    Dim intMesCorte As Integer
    Dim dblItem As Double
    
    Set rstValeCab = Nothing
    Set rstValeDet = Nothing
    
    Set rstValeCab = New ADODB.Recordset
    Set rstValeDet = New ADODB.Recordset
    
    Rem LECTURA DE INGRESOS Y SALIDAS (Vales)
    
    strFechaCorteInicialDeValesParaCP = sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "FechaCorteInicialDeValesParaCP", "l")
    
    intAnnoCorte = Year(CDate(strFechaCorteInicialDeValesParaCP)) - IIf(Month(CDate(strFechaCorteInicialDeValesParaCP)) > 1, 0, 1)
    intMesCorte = IIf(Month(CDate(strFechaCorteInicialDeValesParaCP)) > 1, Month(CDate(strFechaCorteInicialDeValesParaCP)) - 1, 12)
    
    If MsgBox("¿Desea generar los Vales Iniciales a partir del Cierre del Periodo: " & intAnnoCorte & "-" & Format(intMesCorte, "00") & "?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
        Exit Sub
    End If
    
    'abrirCnDBMilano
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "CM.IDALMACEN, "
    SqlCad = SqlCad & "FAM.IDFAMILIA, "
    SqlCad = SqlCad & "FAM.NOMBRE AS FAMILIA "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "((CIERREMENSUAL AS CM "
    SqlCad = SqlCad & "LEFT JOIN INSUMO AS INS ON INS.IDINSUMO = CM.IDINSUMO) "
    SqlCad = SqlCad & "LEFT JOIN SUBFAMILIA AS SF ON SF.IDSUBFAMILIA = INS.IDSUBFAMILIA) "
    SqlCad = SqlCad & "LEFT JOIN FAMILIA AS FAM ON FAM.IDFAMILIA = SF.IDFAMILIA "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "CM.AÑO = " & intAnnoCorte & " AND "
    SqlCad = SqlCad & "CM.MES = " & intMesCorte & " "
    SqlCad = SqlCad & "GROUP BY "
    SqlCad = SqlCad & "CM.IDALMACEN, "
    SqlCad = SqlCad & "FAM.IDFAMILIA, "
    SqlCad = SqlCad & "FAM.NOMBRE "
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "CM.IDALMACEN, "
    SqlCad = SqlCad & "FAM.NOMBRE"
    
    If rstValeCab.State = 1 Then rstValeCab.Close
    
    rstValeCab.Open SqlCad, cnBdStudioModa, adOpenForwardOnly, adLockReadOnly
    
    If Not rstValeCab.EOF Then
        frameProgreso.Visible = True
        barraProgreso.Max = ModUtilitario.devuelveCantRegistros(rstValeCab)
        barraProgreso.Value = 0
        frameProgreso.Caption = "Generando Vales Iniciales..."
        
        Do While Not rstValeCab.EOF
            With objAyudaVale
                .inicializarEntidades
                
                .CodigoAlmacen = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODALM", "EF2ALMACENES", "F2CODALMEXTERNO", Trim(rstValeCab!IDALMACEN & ""), "T")
                .NumeroVale = vbNullString
                .TipoVale = "I"
                
                .Fecha = Format(CDate(strFechaCorteInicialDeValesParaCP), "dd/mm/yyyy")
                .CodigoOrigen = "XJ0"
                .TipoCambio = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CAMBIO", "CAMBIOS", "FECHA", strFechaCorteInicialDeValesParaCP, "F"))
                
                    If .TipoCambio = 0 Then
                        .TipoCambio = 2.8
                    End If
                
                .CodigoMoneda = "S"
                .observaciones = "SALDOS INICIALES DE " & Trim(rstValeCab!FAMILIA & "")
                
                .FecReg = Format(Date, "Short Date")
                .UsuReg = wusuario 'ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODUSER", "EF2USERS", "F2CODUSEREXTERNO", Trim(rstValeCab!IDUSUARIOINGRESO & ""), "T")
                .FecMod = Format(Date, "Short Date")
                .UsuMod = wusuario 'ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODUSER", "EF2USERS", "F2CODUSEREXTERNO", Trim(rstValeCab!IDUSUARIOACTUALIZACION & ""), "T")
                
                If .guardarVale Then
                    Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                    
                    'Borrar Detalle de Vale
                    SqlCad = vbNullString
                    SqlCad = "DELETE FROM IF3VALES WHERE F2CODALM = '" & .CodigoAlmacen & "' AND F4NUMVAL = '" & .NumeroVale & "'"
                    
                    cnn_dbbancos.Execute SqlCad
                    
                    Actualiza_Log SqlCad, StrConexDbBancos
                    
                    SqlCad = vbNullString
                    SqlCad = SqlCad & "SELECT "
                    SqlCad = SqlCad & "CM.IDINSUMO, "
                    SqlCad = SqlCad & "FAM.IDFAMILIA, "
                    SqlCad = SqlCad & "CM.CANTIDAD, "
                    SqlCad = SqlCad & "CM.COSTOPROMEDIO "
                    
                    SqlCad = SqlCad & "FROM "
                    SqlCad = SqlCad & "CIERREMENSUAL AS CM "
                    SqlCad = SqlCad & "LEFT JOIN INSUMO AS INS ON INS.IDINSUMO = CM.IDINSUMO "
                    SqlCad = SqlCad & "LEFT JOIN SUBFAMILIA AS SF ON SF.IDSUBFAMILIA = INS.IDSUBFAMILIA "
                    SqlCad = SqlCad & "LEFT JOIN FAMILIA AS FAM ON FAM.IDFAMILIA = SF.IDFAMILIA "
                    
                    SqlCad = SqlCad & "WHERE "
                    SqlCad = SqlCad & "CM.AÑO = " & intAnnoCorte & " AND "
                    SqlCad = SqlCad & "CM.MES = " & intMesCorte & " AND "
                    SqlCad = SqlCad & "FAM.IDFAMILIA = '" & Trim(rstValeCab!IDFAMILIA & "") & "' "
                    
                    SqlCad = SqlCad & "ORDER BY "
                    SqlCad = SqlCad & "CM.IDINSUMO"
                    
                    If rstValeDet.State = 1 Then rstValeDet.Close
                    
                    rstValeDet.Open SqlCad, cnBdStudioModa, adOpenForwardOnly, adLockReadOnly
                    
                    If Not rstValeDet.EOF Then
                        dblItem = 0
                        
                        '/-- Lectura de los datos del detalle hasta que no exista ningun dato
                        Do While Not rstValeDet.EOF
                            .inicializarEntidadesDetalle
                            
                            dblItem = dblItem + 1
                            
                            .CodigoProducto = Trim(rstValeDet!IDINSUMO & "")
                            .CodigoProductoOriginal = Trim(rstValeDet!IDINSUMO & "")
                            .Cantidad = Val(rstValeDet!Cantidad & "")
                            
                            .ValorVenta = Val(rstValeDet!COSTOPROMEDIO & "")
                            .IGV = 0
                            .TOTAL = .ValorVenta
                            .ValorVentaDol = Val(Format(.ValorVenta / .TipoCambio, "#0.0000"))
                            .IgvDol = 0
                            .TotalDol = Val(Format(.TOTAL / .TipoCambio, "#0.0000"))
                            
                            .ITEM = dblItem
                            
                            .guardarValeDetalleOneByOne
                            
                            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
    
                            rstValeDet.MoveNext
                        Loop
                    End If
                End If
            End With
            
            DoEvents
            
            barraProgreso.Value = barraProgreso.Value + 1
            frameProgreso.Caption = "Generando Vales Iniciales... " & FormatPercent(barraProgreso.Value / barraProgreso.Max, 3)
            
            rstValeCab.MoveNext
        Loop
            MsgBox "Vales Iniciales Generados.", vbInformation + vbOKOnly, wnomcia
    Else
        MsgBox "No se ubico Cierre de Mes para el Periodo = " & intAnnoCorte & "-" & Format(intMesCorte, "00") & ".", vbInformation + vbOKOnly, wnomcia
    End If
    
    SqlCad = vbNullString
    strFechaCorteInicialDeValesParaCP = vbNullString
    intAnnoCorte = 0
    intMesCorte = 0
    dblItem = 0
    
    frameProgreso.Visible = False
    
    If rstValeCab.State = 1 Then rstValeCab.Close
    If rstValeDet.State = 1 Then rstValeDet.Close
    
    Set rstValeCab = Nothing
    Set rstValeDet = Nothing
    
    Exit Sub
errImportarCierreMesValesInicialesServidorExterno:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    'Resume
    frameProgreso.Visible = False
    barraProgreso.Value = 0
    
    Err.Clear
End Sub

Public Sub importarOpPendientesParaCompromisoInicialServidorExterno(ByVal frameProgreso As Object, _
                                                                                ByVal barraProgreso As Object)
    
    On Error GoTo errImportarOpPendientesParaCompromisoInicialServidorExterno
    
    Dim rstResumen As New ADODB.Recordset
    Dim rstValeDet As New ADODB.Recordset
    
    Dim dblItem As Double
    Dim dblCantidadLibre As Double
    Dim dblCantidadCompromiso As Double
    
    'Set rstResumen = Nothing
    'Set rstValeDet = Nothing
    
    'Set rstResumen = New ADODB.Recordset
    'Set rstValeDet = New ADODB.Recordset
    
    If MsgBox("¿Desea generar los Vales de Compromiso Inicial a partir de Insumos Pendientes de Descarga en O/P's?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
        Exit Sub
    End If
    
    'abrirCnDBMilano
    
    SqlCad = vbNullString
'    SqlCad = SqlCad & "SELECT "
'    SqlCad = SqlCad & "RESUMEN.IDPEDIDO, "
'    SqlCad = SqlCad & "RESUMEN.IDINSUMO, "
'    SqlCad = SqlCad & "RESUMEN.DESCRIPCIONINSUMO, "
'    SqlCad = SqlCad & "SUM(RESUMEN.SALDO) AS CANTIDADPENDIENTE "
'    SqlCad = SqlCad & "FROM "
'    SqlCad = SqlCad & "(SELECT "
'    SqlCad = SqlCad & "CAB.IDCATEGORIATIPO, "
'    SqlCad = SqlCad & "CAB.OP, "
'    SqlCad = SqlCad & "DET.IDORDENPRODUCCION, "
'    SqlCad = SqlCad & "RQ.IDPEDIDO, "
'    SqlCad = SqlCad & "DET.IDINSUMO, "
'    SqlCad = SqlCad & "INS.NOMBRE AS DESCRIPCIONINSUMO, "
'    'SqlCad = SqlCad & "UM.NOMBRE AS UMEDIDA, "
'    'SqlCad = SqlCad & "DET.CANTIDAD, "
'    SqlCad = SqlCad & "(DET.CANTIDAD - ISNULL(MOVOP.CANTIDAD, 0)) AS SALDO "
'    SqlCad = SqlCad & "FROM "
'    SqlCad = SqlCad & "ORDENPRODUCCIONDESCARGO AS DET "
'    SqlCad = SqlCad & "LEFT JOIN ORDENPRODUCCION AS CAB ON CAB.IDORDENPRODUCCION = DET.IDORDENPRODUCCION "
'    SqlCad = SqlCad & "LEFT JOIN ORDENPRODUCCION AS OP ON OP.IDORDENPRODUCCION = CAB.IDORDENPRODUCCION "
'    SqlCad = SqlCad & "LEFT JOIN "
'    SqlCad = SqlCad & "(SELECT IDREQUERIMIENTO, (CASE WHEN IDPEDIDO = 0 THEN NROPEDIDO ELSE IDPEDIDO END) AS IDPEDIDO FROM REQUERIMIENTO) AS RQ "
'    SqlCad = SqlCad & "ON RQ.IDREQUERIMIENTO = OP.IDREQUERIMIENTO "
'    SqlCad = SqlCad & "LEFT JOIN INSUMO AS INS ON INS.IDINSUMO = DET.IDINSUMO "
'    'SqlCad = SqlCad & "LEFT JOIN UNIDADMEDIDA AS UM ON UM.IDUNIDADMEDIDA = INS.IDUNIDADMEDIDA "
'    SqlCad = SqlCad & "LEFT JOIN "
'    SqlCad = SqlCad & "(SELECT "
'    SqlCad = SqlCad & "CAB.IDORDENPRODUCCION, "
'    SqlCad = SqlCad & "RQ.IDPEDIDO, "
'    SqlCad = SqlCad & "DET.IDINSUMO, "
'    SqlCad = SqlCad & "SUM(DET.CANTIDAD) AS CANTIDAD "
'    SqlCad = SqlCad & "FROM "
'    SqlCad = SqlCad & "SALIDADETALLE AS DET "
'    SqlCad = SqlCad & "LEFT JOIN SALIDA AS CAB ON CAB.IDSALIDA = DET.IDSALIDA "
'    SqlCad = SqlCad & "LEFT JOIN ORDENPRODUCCION AS OP ON OP.IDORDENPRODUCCION = CAB.IDORDENPRODUCCION "
'    SqlCad = SqlCad & "LEFT JOIN "
'    SqlCad = SqlCad & "(SELECT IDREQUERIMIENTO, (CASE WHEN IDPEDIDO = 0 THEN NROPEDIDO ELSE IDPEDIDO END) AS IDPEDIDO FROM REQUERIMIENTO) AS RQ "
'    SqlCad = SqlCad & "ON RQ.IDREQUERIMIENTO = OP.IDREQUERIMIENTO "
'    SqlCad = SqlCad & "WHERE "
'    SqlCad = SqlCad & "CAB.IDTIPOSALIDA IN ('TIP0000001') AND "
'    SqlCad = SqlCad & "CAB.ANULADO = 0 AND "
'    SqlCad = SqlCad & "CAB.IDORDENPRODUCCION <> 0 AND "
'    SqlCad = SqlCad & "RQ.IDPEDIDO <> 0 "
'    SqlCad = SqlCad & "GROUP BY "
'    SqlCad = SqlCad & "CAB.IDORDENPRODUCCION, "
'    SqlCad = SqlCad & "RQ.IDPEDIDO, "
'    SqlCad = SqlCad & "DET.IDINSUMO) AS MOVOP "
'    SqlCad = SqlCad & "ON MOVOP.IDORDENPRODUCCION = DET.IDORDENPRODUCCION AND MOVOP.IDPEDIDO = RQ.IDPEDIDO AND MOVOP.IDINSUMO = DET.IDINSUMO "
'    SqlCad = SqlCad & "WHERE "
'    SqlCad = SqlCad & "YEAR(CAB.FECHA) = 2014 AND "
'    SqlCad = SqlCad & "DET.ANULADO = 0 AND "
'    SqlCad = SqlCad & "CAB.IDORDENPRODUCCION <> 0 AND "
'    SqlCad = SqlCad & "RQ.IDPEDIDO <> 0 AND "
'    SqlCad = SqlCad & "(DET.CANTIDAD - ISNULL(MOVOP.CANTIDAD, 0)) > 0 AND "
'    SqlCad = SqlCad & "CAB.CERRADO = 0 AND "
'    SqlCad = SqlCad & "CAB.ANULADO = 0 AND "
'    SqlCad = SqlCad & "INS.OP = 1) AS RESUMEN "
'    SqlCad = SqlCad & "GROUP BY "
'    SqlCad = SqlCad & "RESUMEN.IDPEDIDO, "
'    SqlCad = SqlCad & "RESUMEN.IDINSUMO, "
'    SqlCad = SqlCad & "RESUMEN.DESCRIPCIONINSUMO"
    
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "PEDIDO AS IDPEDIDO, "
    SqlCad = SqlCad & "CODPRODUCTO AS IDINSUMO, "
    SqlCad = SqlCad & "DESCRIPCION AS DESCRIPCIONINSUMO, "
    SqlCad = SqlCad & "CANTIDADPENDIENTE "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "TMPUTILINSUMOPENDIENTEOPSSERVER"
    
    abrirCnTemporal
    
    If rstResumen.State = 1 Then rstResumen.Close
    
    'rstResumen.Open SqlCad, cnBdStudioModa, adOpenForwardOnly, adLockReadOnly
    rstResumen.Open SqlCad, cnDBTemp, adOpenForwardOnly, adLockReadOnly
    
    If Not rstResumen.EOF Then
        frameProgreso.Visible = True
        barraProgreso.Max = ModUtilitario.devuelveCantRegistros(rstResumen)
        barraProgreso.Value = 0
        frameProgreso.Caption = "Descargando Ordenes de Produccion Pendientes..."
        
        'abrirCnTemporal
        
        cnDBTemp.Execute "DELETE FROM TMPUTILINSUMOPENDIENTEOPS"
        
        Do While Not rstResumen.EOF
            'If ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "IDPEDIDO", "PEDIDO", "IDPEDIDO", Format(Trim(rstResumen!IDPEDIDO & ""), "0000"), "T") <> vbNullString Then
            If ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "COD_SOLICITUD", "TB_CABSOLICITUD", "COD_SOLICITUD", Format(Trim(rstResumen!IDPEDIDO & ""), "0000"), "T") <> vbNullString Then
                If Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "COUNT(*)", "TMPUTILINSUMOPENDIENTEOPS", "PEDIDO", Format(Trim(rstResumen!IDPEDIDO & ""), "0000"), "T", "AND CODPRODUCTO = '" & Trim(rstResumen!IDINSUMO & "") & "'")) = 0 Then
                    SqlCad = vbNullString
                    SqlCad = SqlCad & "INSERT INTO TMPUTILINSUMOPENDIENTEOPS("
                    SqlCad = SqlCad & "PEDIDO, CODPRODUCTO, DESCRIPCION, CANTIDADPENDIENTE, "
                    SqlCad = SqlCad & "FECHAEMISION, FECHAENTREGA) "
                    SqlCad = SqlCad & "VALUES("
                    SqlCad = SqlCad & "'" & Format(Trim(rstResumen!IDPEDIDO & ""), "0000") & "', "
                    SqlCad = SqlCad & "'" & Trim(rstResumen!IDINSUMO & "") & "', "
                    SqlCad = SqlCad & "'" & Trim(rstResumen!DESCRIPCIONINSUMO & "") & "', "
                    SqlCad = SqlCad & Val(rstResumen!CANTIDADPENDIENTE & "") & ", "
                    'SqlCad = SqlCad & "CVDATE('" & Format(ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "FECHAEMISION", "PEDIDO", "IDPEDIDO", Format(Trim(rstResumen!IDPEDIDO & ""), "0000"), "T"), "dd/mm/yyyy") & "'), "
                    'SqlCad = SqlCad & "CVDATE('" & Format(ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "FECHAENTREGA", "PEDIDO", "IDPEDIDO", Format(Trim(rstResumen!IDPEDIDO & ""), "0000"), "T"), "dd/mm/yyyy") & "'))"
                    SqlCad = SqlCad & "CVDATE('" & Format(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CS_FECHA", "TB_CABSOLICITUD", "COD_SOLICITUD", Format(Trim(rstResumen!IDPEDIDO & ""), "0000"), "T"), "dd/mm/yyyy") & "'), "
                    SqlCad = SqlCad & "CVDATE('" & Format(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CS_FENTREGA", "TB_CABSOLICITUD", "COD_SOLICITUD", Format(Trim(rstResumen!IDPEDIDO & ""), "0000"), "T"), "dd/mm/yyyy") & "'))"
                Else
                    SqlCad = vbNullString
                    SqlCad = SqlCad & "UPDATE "
                    SqlCad = SqlCad & "TMPUTILINSUMOPENDIENTEOPS "
                    SqlCad = SqlCad & "SET "
                    SqlCad = SqlCad & "CANTIDADPENDIENTE = CANTIDADPENDIENTE + " & Val(rstResumen!CANTIDADPENDIENTE & "") & " "
                    SqlCad = SqlCad & "WHERE "
                    SqlCad = SqlCad & "PEDIDO = '" & Format(Trim(rstResumen!IDPEDIDO & ""), "0000") & "' AND  "
                    SqlCad = SqlCad & "CODPRODUCTO = '" & Trim(rstResumen!IDINSUMO & "") & "'"
                End If
                
                'abrirCnTemporal
                
                cnDBTemp.Execute SqlCad
            End If
            
            DoEvents
            
            barraProgreso.Value = barraProgreso.Value + 1
            frameProgreso.Caption = "Descargando Ordenes de Produccion Pendientes... " & FormatPercent(barraProgreso.Value / barraProgreso.Max, 3)
            
            rstResumen.MoveNext
        Loop
    End If
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "PEDIDO, "
    SqlCad = SqlCad & "FECHAEMISION, "
    SqlCad = SqlCad & "FECHAENTREGA, "
    SqlCad = SqlCad & "COUNT(*) AS CANTIDAD "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "TMPUTILINSUMOPENDIENTEOPS "
    SqlCad = SqlCad & "GROUP BY "
    SqlCad = SqlCad & "PEDIDO, "
    SqlCad = SqlCad & "FECHAEMISION, "
    SqlCad = SqlCad & "FECHAENTREGA "
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "FECHAENTREGA, "
    SqlCad = SqlCad & "FECHAEMISION"
        
    If rstResumen.State = 1 Then rstResumen.Close
    
    abrirCnTemporal
    
    rstResumen.Open SqlCad, cnDBTemp, adOpenForwardOnly, adLockReadOnly
    
    If Not rstResumen.EOF Then
        frameProgreso.Visible = True
        barraProgreso.Max = ModUtilitario.devuelveCantRegistros(rstResumen)
        barraProgreso.Value = 0
        frameProgreso.Caption = "Generando Vales de Compromiso..."
        
        'abrirCnTemporal

        'cnDBTemp.Execute "DELETE FROM TMPUTILINSUMOPENDIENTEOPS"

        Do While Not rstResumen.EOF
            With objAyudaVale
                .inicializarEntidades
                
                .CodigoAlmacen = "01"
                .NumeroVale = vbNullString
                .TipoVale = "I"
                
                .Fecha = Format(Date, "dd/mm/yyyy")
                .CodigoOrigen = "XCS"
                .TipoCambio = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CAMBIO", "CAMBIOS", "FECHA", .Fecha, "F"))
                
                    If .TipoCambio = 0 Then
                        .TipoCambio = "2.8"
                    End If

                .CodigoMoneda = "S"

                .referencia = wnomcia
                .observaciones = "COMPROMISO INICIAL AUTOMATICO, EN BASE A INSUMOS PENDIENTES DE DESCARGA DE O/P " & _
                                    Trim(rstResumen!PEDIDO & "") & " (" & Val(rstResumen!Cantidad & "") & ")."
                
                .FecReg = Format(Date, "Short Date")
                .UsuReg = wusuario
                .FecMod = Format(Date, "Short Date")
                .UsuMod = wusuario
                
                If .guardarVale Then
                    Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                    
                    'Borrar Detalle de Vale
                    SqlCad = vbNullString
                    SqlCad = "DELETE FROM IF3VALES WHERE F2CODALM = '" & .CodigoAlmacen & "' AND F4NUMVAL = '" & .NumeroVale & "'"

                    cnn_dbbancos.Execute SqlCad

                    Actualiza_Log SqlCad, cnn_dbbancos.ConnectionString

                    SqlCad = vbNullString
                    SqlCad = SqlCad & "SELECT "
                    SqlCad = SqlCad & "* "
                    SqlCad = SqlCad & "FROM "
                    SqlCad = SqlCad & "TMPUTILINSUMOPENDIENTEOPS "
                    SqlCad = SqlCad & "WHERE "
                    SqlCad = SqlCad & "PEDIDO = '" & Trim(rstResumen!PEDIDO & "") & "' "
                    SqlCad = SqlCad & "ORDER BY "
                    SqlCad = SqlCad & "DESCRIPCION"

                    If rstValeDet.State = 1 Then rstValeDet.Close

                    rstValeDet.Open SqlCad, cnDBTemp, adOpenForwardOnly, adLockReadOnly
                    
                    If Not rstValeDet.EOF Then
                        dblItem = 0
                        
                        Do While Not rstValeDet.EOF
                            .CodigoProducto = Trim(rstValeDet!CodProducto & "")
                            
                            dblCantidadLibre = .devuelveStockFisicoDeProducto("L")
                            
                            If dblCantidadLibre > 0 Then
                                If Val(rstValeDet!CANTIDADPENDIENTE & "") <= dblCantidadLibre Then
                                    dblCantidadCompromiso = Val(rstValeDet!CANTIDADPENDIENTE & "")
                                Else
                                    dblCantidadCompromiso = dblCantidadLibre
                                End If
                                
                                .inicializarEntidadesDetalle
                                
                                .NumeroOrdenCompra = vbNullString
                                
                                dblItem = dblItem + 1
                                
                                .Requerimiento = vbNullString
                                
                                .CodigoProducto = Trim(rstValeDet!CodProducto & "")
                                .CodigoProductoOriginal = Trim(rstValeDet!CodProducto & "")
                                .Cantidad = dblCantidadCompromiso * -1
                                .ITEM = dblItem
    
                                .guardarValeDetalleOneByOne
    
                                Actualiza_Log .SQLSelectAlter, StrConexDbBancos
    
                                .inicializarEntidadesDetalle
    
                                dblItem = dblItem + 1
    
                                .Requerimiento = Trim(rstValeDet!PEDIDO & "")
    
                                .CodigoProducto = Trim(rstValeDet!CodProducto & "")
                                .CodigoProductoOriginal = Trim(rstValeDet!CodProducto & "")
                                .Cantidad = dblCantidadCompromiso
                                .ITEM = dblItem
    
                                .guardarValeDetalleOneByOne
    
                                Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                            End If
                            
                            rstValeDet.MoveNext
                        Loop
                            If dblItem = 0 Then
                                .eliminarVale
                                
                                Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                            End If
                    End If
                End If
            End With
            
            DoEvents
            
            barraProgreso.Value = barraProgreso.Value + 1
            frameProgreso.Caption = "Descargando Ordenes de Produccion Pendientes... " & FormatPercent(barraProgreso.Value / barraProgreso.Max, 3)
            
            rstResumen.MoveNext
        Loop
    End If
    
    SqlCad = vbNullString
    dblItem = 0
    
    frameProgreso.Visible = False
    
    If rstResumen.State = 1 Then rstResumen.Close
    If rstValeDet.State = 1 Then rstValeDet.Close
    
    Set rstResumen = Nothing
    Set rstValeDet = Nothing
    
    Exit Sub
errImportarOpPendientesParaCompromisoInicialServidorExterno:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    'Resume
    frameProgreso.Visible = False
    barraProgreso.Value = 0
    
    Err.Clear
End Sub

Public Sub importarTomaInventarioServidorExterno(ByVal frameProgreso As Object, _
                                                            ByVal barraProgreso As Object)
    
    On Error GoTo errImportarTomaInventarioServidorExterno
    
    Dim rstTomaInventario As ADODB.Recordset
    Dim dblRegistro As Double
    Dim dblCantRegistroActualizado As Double
    
    Set rstTomaInventario = Nothing
    
    Set rstTomaInventario = New ADODB.Recordset
    
    Rem LECTURA DE  TOMA DE INVENTARIOS
    
    
    If MsgBox("¿Desea importar la Toma de Inventario: 2014-12?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
        Exit Sub
    End If
    
    'abrirCnDBMilano
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "TI.IDALMACEN, "
    SqlCad = SqlCad & "TI.AÑO, "
    SqlCad = SqlCad & "TI.MES, "
    SqlCad = SqlCad & "TI.IDINSUMO, "
    SqlCad = SqlCad & "TI.TEORICO, "
    SqlCad = SqlCad & "TI.FISICO, "
    SqlCad = SqlCad & "TI.DIFERENCIA "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "TOMAINVENTARIO AS TI "
    SqlCad = SqlCad & "LEFT JOIN INSUMO AS INS ON INS.IDINSUMO = TI.IDINSUMO "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "TI.AÑO = 2014 AND "
    SqlCad = SqlCad & "TI.MES = 12 AND "
    SqlCad = SqlCad & "TI.IDALMACEN = 'ALM0000001' "
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "INS.NOMBRE"
    
    If rstTomaInventario.State = 1 Then rstTomaInventario.Close
    
    rstTomaInventario.Open SqlCad, cnBdStudioModa, adOpenForwardOnly, adLockReadOnly
    
    If Not rstTomaInventario.EOF Then
        frameProgreso.Visible = True
        barraProgreso.Max = ModUtilitario.devuelveCantRegistros(rstTomaInventario)
        barraProgreso.Value = 0
        frameProgreso.Caption = "Importando Toma de Inventarios..."
        
        dblCantRegistroActualizado = 0
        
        Do While Not rstTomaInventario.EOF
            SqlCad = vbNullString
            SqlCad = SqlCad & "UPDATE "
            SqlCad = SqlCad & "H3TOMAINV "
            SqlCad = SqlCad & "SET "
            SqlCad = SqlCad & "F3STOCKFISICO = " & Val(rstTomaInventario!FISICO & "") & ", "
            SqlCad = SqlCad & "F3DIFERENCIA = " & Val(rstTomaInventario!FISICO & "") & " - F3STOCKSISTEMA "
            SqlCad = SqlCad & "WHERE "
            SqlCad = SqlCad & "F2CODALM = '" & ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODALM", "EF2ALMACENES", "F2CODALMEXTERNO", Trim(rstTomaInventario!IDALMACEN & ""), "T") & "' AND "
            SqlCad = SqlCad & "F4ANNO = '" & Trim(rstTomaInventario![AÑO] & "") & "' AND "
            SqlCad = SqlCad & "F4MES = '" & Format(Trim(rstTomaInventario!mes & ""), "00") & "' AND "
            SqlCad = SqlCad & "F5CODPRO = '" & Trim(rstTomaInventario!IDINSUMO & "") & "'"
            
            dblRegistro = 0
            
            cnn_dbbancos.Execute SqlCad, dblRegistro
            
            If dblRegistro > 0 Then
                Actualiza_Log SqlCad, StrConexDbBancos
                
                dblCantRegistroActualizado = dblCantRegistroActualizado + 1
            End If
            
            DoEvents
            
            barraProgreso.Value = barraProgreso.Value + 1
            frameProgreso.Caption = "Importando Toma de Inventarios... " & FormatPercent(barraProgreso.Value / barraProgreso.Max, 3)
            
            rstTomaInventario.MoveNext
        Loop
            MsgBox "Toma de Inventarios Fisico importado." & vbNullString & _
                    "Registros Actualizados = " & dblCantRegistroActualizado & ".", vbInformation + vbOKOnly, wnomcia
    Else
        MsgBox "No se ubico Toma de Inventario para el Periodo = 2014-12.", vbInformation + vbOKOnly, wnomcia
    End If
    
    SqlCad = vbNullString
    
    frameProgreso.Visible = False
    
    If rstTomaInventario.State = 1 Then rstTomaInventario.Close
    
    Set rstTomaInventario = Nothing
    
    Exit Sub
errImportarTomaInventarioServidorExterno:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    'Resume
    frameProgreso.Visible = False
    barraProgreso.Value = 0
    
    Err.Clear
End Sub

Public Function importarResumenRequerimientoProduccion(ByVal frameProgreso As Object, _
                                                        ByVal barraProgreso As Object, _
                                                        Optional ByVal strNroPedido As String, _
                                                        Optional ByVal strCodigoProducto As String, _
                                                        Optional ByVal strFiltroProducto As String, _
                                                        Optional ByVal strFechaEntregaCorte As String) As Boolean
    
    On Error GoTo errImportarResumenRequerimientoProduccion
    
    
    Dim cmdOP As ADODB.Command
    
    Dim rstResumen As New ADODB.Recordset
    
    Dim dblCantidadRegistro As Double
    
    Dim dblCantidadIntentos As Double
    
    importarResumenRequerimientoProduccion = False
    
    abrirCnTemporal
        
    cnDBTemp.Execute "DELETE FROM TMPUTILRESUMENREQUERIMIENTOPRODUCCION"
    
    Rem LECTURA DE RESUMEN DE REQUERIMIENTO DE PRODUCCION (EN BASE ORDENES DE PRODUCCION)
    
    'abrirCnDBMilano
    
''    SqlCad = vbNullString
''    SqlCad = SqlCad & "SELECT "
''    SqlCad = SqlCad & "(RQ.IDPEDIDO + DET.IDINSUMO + CAST(DET.IDORDENPRODUCCION AS VARCHAR(50))) AS LLAVE, "
''    SqlCad = SqlCad & "RQ.IDPEDIDO AS NROPEDIDO, "
''    SqlCad = SqlCad & "PER.NOMBRE AS CLIENTE, "
''    SqlCad = SqlCad & "CONVERT(VARCHAR,(CONVERT(DATE,PED.FECHAEMISION, 103)), 103) AS FEMISION, "
''    SqlCad = SqlCad & "CONVERT(VARCHAR,(CONVERT(DATE,PED.FECHAENTREGA, 103)), 103) AS FENTREGA, "
''    SqlCad = SqlCad & "USU.NOMBRE AS VENDEDOR, "
''    SqlCad = SqlCad & "DET.IDINSUMO AS CODPRODUCTO, "
''    SqlCad = SqlCad & "INS.NOMBRE AS NOMPRODUCTO, "
''    SqlCad = SqlCad & "UM.NOMBRE AS UM, "
''    SqlCad = SqlCad & "(INS.NOMBRE + ' ( ' + UM.NOMBRE + ' ) ') AS NOMPRODUCTOUM, "
''    SqlCad = SqlCad & "DET.IDORDENPRODUCCION AS IDOP, "
''    SqlCad = SqlCad & "CT.NOMBRE AS CATEGORIA, "
''    SqlCad = SqlCad & "CAB.OP AS NROOP, "
''    SqlCad = SqlCad & "(MD.TIPOMODELO + '-' + MD.IDMODELO) AS MODELO, "
''    SqlCad = SqlCad & "COL.NOMBRE AS COLOR, "
''    SqlCad = SqlCad & "CAB.CANTIDADTOTAL AS CANTIDADPEDIDO, "
''    SqlCad = SqlCad & "CAB.DESCRIPCION AS DESCRIPCIONOP, "
''    SqlCad = SqlCad & "CAB.OBSERVACION AS OBSERVACIONOP, "
''    SqlCad = SqlCad & "SUM( DET.CANTIDAD ) AS CANTIDAD, "
''    SqlCad = SqlCad & "SUM( ((DET.CANTIDAD - ISNULL(MOVOP.CANTIDAD, 0)) + ISNULL(MOVOPINGRESO.CANTIDAD, 0)) ) AS SALDO "
''
''    SqlCad = SqlCad & "FROM "
''    SqlCad = SqlCad & "ORDENPRODUCCIONDESCARGO AS DET "
''    SqlCad = SqlCad & "LEFT JOIN ORDENPRODUCCION AS CAB ON CAB.IDORDENPRODUCCION = DET.IDORDENPRODUCCION "
''    SqlCad = SqlCad & "LEFT JOIN CATEGORIATIPO AS CT ON CT.IDCATEGORIATIPO = CAB.IDCATEGORIATIPO "
''
''    SqlCad = SqlCad & "LEFT JOIN "
''    SqlCad = SqlCad & "(SELECT "
''    SqlCad = SqlCad & "IDREQUERIMIENTO, (CASE WHEN IDPEDIDO = 0 THEN NROPEDIDO ELSE IDPEDIDO END) AS IDPEDIDO, NROMODELO, IDCOLOR "
''    SqlCad = SqlCad & "FROM "
''    SqlCad = SqlCad & "REQUERIMIENTO) AS RQ "
''    SqlCad = SqlCad & "ON RQ.IDREQUERIMIENTO = CAB.IDREQUERIMIENTO "
''
''    SqlCad = SqlCad & "LEFT JOIN PEDIDO AS PED ON PED.IDPEDIDO = RQ.IDPEDIDO "
''    SqlCad = SqlCad & "LEFT JOIN PERSONA AS PER ON PER.IDPERSONA = PED.IDPERSONA "
''    SqlCad = SqlCad & "LEFT JOIN USUARIO AS USU ON USU.IDUSUARIO = PED.IDUSUARIOACTUALIZACION "
''    SqlCad = SqlCad & "LEFT JOIN MODELO AS MD ON MD.NROMODELO = RQ.NROMODELO "
''    SqlCad = SqlCad & "LEFT JOIN INSUMO AS INS ON INS.IDINSUMO = DET.IDINSUMO "
''    SqlCad = SqlCad & "LEFT JOIN UNIDADMEDIDA AS UM ON UM.IDUNIDADMEDIDA = INS.IDUNIDADMEDIDA "
''    SqlCad = SqlCad & "LEFT JOIN COLOR AS COL ON COL.IDCOLOR = RQ.IDCOLOR "
''
''    SqlCad = SqlCad & "LEFT JOIN "
''    SqlCad = SqlCad & "(SELECT "
''    SqlCad = SqlCad & "CAB.IDORDENPRODUCCION, "
''    SqlCad = SqlCad & "RQ.IDPEDIDO, "
''    SqlCad = SqlCad & "DET.IDINSUMO, "
''    SqlCad = SqlCad & "SUM(DET.CANTIDAD) AS CANTIDAD "
''    SqlCad = SqlCad & "FROM "
''    SqlCad = SqlCad & "SALIDADETALLE AS DET "
''    SqlCad = SqlCad & "LEFT JOIN SALIDA AS CAB "
''    SqlCad = SqlCad & "ON CAB.IDSALIDA = DET.IDSALIDA "
''    SqlCad = SqlCad & "LEFT JOIN ORDENPRODUCCION AS OP ON OP.IDORDENPRODUCCION = CAB.IDORDENPRODUCCION "
''    SqlCad = SqlCad & "LEFT JOIN "
''    SqlCad = SqlCad & "(SELECT "
''    SqlCad = SqlCad & "IDREQUERIMIENTO, "
''    SqlCad = SqlCad & "(CASE WHEN IDPEDIDO = 0 THEN NROPEDIDO ELSE IDPEDIDO END) AS IDPEDIDO "
''    SqlCad = SqlCad & "FROM REQUERIMIENTO) AS RQ "
''    SqlCad = SqlCad & "ON RQ.IDREQUERIMIENTO = OP.IDREQUERIMIENTO "
''    SqlCad = SqlCad & "WHERE "
''    SqlCad = SqlCad & "CAB.IDTIPOSALIDA IN ('TIP0000001') AND "
''    SqlCad = SqlCad & "CAB.ANULADO = 0 "
''
''        If strNroPedido <> vbNullString Then
''            SqlCad = SqlCad & "AND RQ.IDPEDIDO = '" & strNroPedido & "' "
''        End If
''
''    SqlCad = SqlCad & "GROUP BY "
''    SqlCad = SqlCad & "CAB.IDORDENPRODUCCION, "
''    SqlCad = SqlCad & "RQ.IDPEDIDO, DET.IDINSUMO) AS MOVOP "
''    SqlCad = SqlCad & "ON MOVOP.IDORDENPRODUCCION = DET.IDORDENPRODUCCION AND MOVOP.IDPEDIDO = RQ.IDPEDIDO AND MOVOP.IDINSUMO = DET.IDINSUMO "
''
''    SqlCad = SqlCad & "LEFT JOIN "
''    SqlCad = SqlCad & "(SELECT "
''    SqlCad = SqlCad & "CAB.IDORDENPRODUCCION, "
''    SqlCad = SqlCad & "RQ.IDPEDIDO, "
''    SqlCad = SqlCad & "DET.IDINSUMO, "
''    SqlCad = SqlCad & "SUM(DET.CANTIDAD) AS CANTIDAD "
''    SqlCad = SqlCad & "FROM "
''    SqlCad = SqlCad & "INGRESODETALLE AS DET "
''    SqlCad = SqlCad & "LEFT JOIN INGRESO AS CAB "
''    SqlCad = SqlCad & "ON CAB.IDINGRESO = DET.IDINGRESO "
''    SqlCad = SqlCad & "LEFT JOIN ORDENPRODUCCION AS OP ON OP.IDORDENPRODUCCION = CAB.IDORDENPRODUCCION "
''    SqlCad = SqlCad & "LEFT JOIN "
''    SqlCad = SqlCad & "(SELECT "
''    SqlCad = SqlCad & "IDREQUERIMIENTO, "
''    SqlCad = SqlCad & "(CASE WHEN IDPEDIDO = 0 THEN NROPEDIDO ELSE IDPEDIDO END) AS IDPEDIDO "
''    SqlCad = SqlCad & "FROM REQUERIMIENTO) AS RQ "
''    SqlCad = SqlCad & "ON RQ.IDREQUERIMIENTO = OP.IDREQUERIMIENTO "
''    SqlCad = SqlCad & "WHERE "
''    SqlCad = SqlCad & "CAB.IDTIPOINGRESO IN ('TIP0000004') AND "
''    SqlCad = SqlCad & "CAB.ANULADO = 0 "
''
''        If strNroPedido <> vbNullString Then
''            SqlCad = SqlCad & "AND RQ.IDPEDIDO = '" & strNroPedido & "' "
''        End If
''
''    SqlCad = SqlCad & "GROUP BY "
''    SqlCad = SqlCad & "CAB.IDORDENPRODUCCION, "
''    SqlCad = SqlCad & "RQ.IDPEDIDO, DET.IDINSUMO) AS MOVOPINGRESO "
''    SqlCad = SqlCad & "ON MOVOPINGRESO.IDORDENPRODUCCION = DET.IDORDENPRODUCCION AND MOVOPINGRESO.IDPEDIDO = RQ.IDPEDIDO AND MOVOPINGRESO.IDINSUMO = DET.IDINSUMO "
''
''    SqlCad = SqlCad & "WHERE "
''    SqlCad = SqlCad & "DET.ANULADO = 0 AND "
''    SqlCad = SqlCad & "CAB.CERRADO = 0 AND "
''    SqlCad = SqlCad & "CAB.ANULADO = 0 AND "
''
''    SqlCad = SqlCad & "RTRIM(LTRIM(PER.NOMBRE + '')) <> '' "
''
''        If strNroPedido <> vbNullString Then
''            SqlCad = SqlCad & "AND RQ.IDPEDIDO = '" & strNroPedido & "' "
''        End If
''
''    SqlCad = SqlCad & "GROUP BY "
''    SqlCad = SqlCad & "(RQ.IDPEDIDO + DET.IDINSUMO + CAST(DET.IDORDENPRODUCCION AS VARCHAR(50))), "
''    SqlCad = SqlCad & "RQ.IDPEDIDO, "
''    SqlCad = SqlCad & "PER.NOMBRE, "
''    SqlCad = SqlCad & "CONVERT(VARCHAR,(CONVERT(DATE,PED.FECHAEMISION, 103)), 103), "
''    SqlCad = SqlCad & "CONVERT(VARCHAR,(CONVERT(DATE,PED.FECHAENTREGA, 103)), 103), "
''    SqlCad = SqlCad & "USU.NOMBRE, "
''    SqlCad = SqlCad & "DET.IDINSUMO, "
''    SqlCad = SqlCad & "INS.NOMBRE, "
''    SqlCad = SqlCad & "UM.NOMBRE, "
''    SqlCad = SqlCad & "(INS.NOMBRE + ' ( ' + UM.NOMBRE + ' ) '), "
''    SqlCad = SqlCad & "DET.IDORDENPRODUCCION, "
''    SqlCad = SqlCad & "CT.NOMBRE, "
''    SqlCad = SqlCad & "CAB.OP, "
''    SqlCad = SqlCad & "(MD.TIPOMODELO + '-' + MD.IDMODELO), "
''    SqlCad = SqlCad & "COL.NOMBRE, "
''    SqlCad = SqlCad & "CAB.CANTIDADTOTAL, "
''    SqlCad = SqlCad & "CAB.DESCRIPCION, "
''    SqlCad = SqlCad & "CAB.OBSERVACION "
''
''    'SqlCad = SqlCad & "HAVING "
''    'SqlCad = SqlCad & "SUM( ((DET.CANTIDAD - ISNULL(MOVOP.CANTIDAD, 0)) + ISNULL(MOVOPINGRESO.CANTIDAD, 0)) ) > 0 "
    
    dblCantidadIntentos = 0
    
    If rstResumen.State = 1 Then rstResumen.Close
    
    Set cmdOP = New ADODB.Command
    
    With cmdOP
        .ActiveConnection = cnBdStudioModa
        .CommandType = adCmdStoredProc
        .CommandText = "usp_ConsultaSaldosOrdenesProduccionCPv2"
        
        .Parameters.Append .CreateParameter("@NROPEDIDO", adVarChar, adParamInput, 10, strNroPedido)
        .Parameters.Append .CreateParameter("@CODIGOPRODUCTO", adVarChar, adParamInput, 50, strCodigoProducto)
        .Parameters.Append .CreateParameter("@FILTROPRODUCTO", adVarChar, adParamInput, 255, strFiltroProducto)
        .Parameters.Append .CreateParameter("@FECHAENTREGACORTE", adVarChar, adParamInput, 20, strFechaEntregaCorte)
        .Parameters.Append .CreateParameter("@CANTIDADFILAS", adInteger, adParamOutput, 5)
        
        .Execute
        
        dblCantidadRegistro = Val(.Parameters("@CANTIDADFILAS") & "")
        
        Set rstResumen = .Execute()
    End With
    
    Set cmdOP = Nothing
    
    DoEvents
    
    frameProgreso.Visible = True
    frameProgreso.Caption = "Ejecutando consulta (1/4)..."
    
    If Not rstResumen.EOF Then
        DoEvents
        
        frameProgreso.Caption = "Contabilizando registros extraidos (1/4)..."
        barraProgreso.Max = dblCantidadRegistro
        barraProgreso.Value = 0
        
        DoEvents
        
        frameProgreso.Caption = "Importando Requerimientos de Producción Pendientes (1/4)..."
        
        Do While Not rstResumen.EOF
            SqlCad = vbNullString
            SqlCad = SqlCad & "INSERT INTO TMPUTILRESUMENREQUERIMIENTOPRODUCCION("
            SqlCad = SqlCad & "LLAVE, "
            SqlCad = SqlCad & "NROPEDIDO, "
            SqlCad = SqlCad & "CLIENTE, "
            SqlCad = SqlCad & "FEMISION, "
            SqlCad = SqlCad & "FENTREGA, "
            SqlCad = SqlCad & "VENDEDOR, "
            SqlCad = SqlCad & "CODPRODUCTO, "
            SqlCad = SqlCad & "NOMPRODUCTO, "
            SqlCad = SqlCad & "UM, "
            SqlCad = SqlCad & "NOMPRODUCTOUM, "
            SqlCad = SqlCad & "IDOP, "
            SqlCad = SqlCad & "CATEGORIA, "
            SqlCad = SqlCad & "NROOP, "
            SqlCad = SqlCad & "MODELO, "
            SqlCad = SqlCad & "COLOR, "
            SqlCad = SqlCad & "CANTIDADPEDIDO, "
            SqlCad = SqlCad & "DESCRIPCIONOP, "
            SqlCad = SqlCad & "OBSERVACIONOP, "
            SqlCad = SqlCad & "CANTIDAD, "
            SqlCad = SqlCad & "SALDO) "
            
            SqlCad = SqlCad & "VALUES("
            SqlCad = SqlCad & "'" & Trim(rstResumen!LLAVE & "") & "', "
            SqlCad = SqlCad & "'" & Trim(rstResumen!NroPedido & "") & "', "
            SqlCad = SqlCad & "'" & Trim(rstResumen!CLIENTE & "") & "', "
            SqlCad = SqlCad & "CVDATE('" & Trim(rstResumen!FEMISION & "") & "'), "
            SqlCad = SqlCad & "CVDATE('" & Trim(rstResumen!FENTREGA & "") & "'), "
            SqlCad = SqlCad & "'" & Trim(rstResumen!VENDEDOR & "") & "', "
            SqlCad = SqlCad & "'" & Trim(rstResumen!CodProducto & "") & "', "
            SqlCad = SqlCad & "'" & Trim(rstResumen!NOMPRODUCTO & "") & "', "
            SqlCad = SqlCad & "'" & Trim(rstResumen!um & "") & "', "
            SqlCad = SqlCad & "'" & Trim(rstResumen!NOMPRODUCTOUM & "") & "', "
            SqlCad = SqlCad & "'" & Trim(rstResumen!IDOP & "") & "', "
            SqlCad = SqlCad & "'" & Trim(rstResumen!CATEGORIA & "") & "', "
            SqlCad = SqlCad & "'" & Trim(rstResumen!NroOP & "") & "', "
            SqlCad = SqlCad & "'" & Trim(rstResumen!Modelo & "") & "', "
            SqlCad = SqlCad & "'" & Trim(rstResumen!Color & "") & "', "
            SqlCad = SqlCad & Val(Format(Val(rstResumen!CANTIDADPEDIDO & ""), "#0.00")) & ", "
            SqlCad = SqlCad & "'" & Trim(rstResumen!DESCRIPCIONOP & "") & "', "
            SqlCad = SqlCad & "'" & Trim(rstResumen!OBSERVACIONOP & "") & "', "
            SqlCad = SqlCad & Val(Format(Val(rstResumen!Cantidad & ""), "#0.00")) & ", "
            SqlCad = SqlCad & Val(Format(Val(rstResumen!SALDO & ""), "#0.00")) & ")"
            
            abrirCnTemporal
            
            cnDBTemp.Execute SqlCad
            
            DoEvents
            
            barraProgreso.Value = barraProgreso.Value + 1
            frameProgreso.Caption = "Importando Requerimientos de Producción Pendientes (1/4)... " & FormatPercent(barraProgreso.Value / barraProgreso.Max, 3)
            
            rstResumen.MoveNext
        Loop
            importarResumenRequerimientoProduccion = True
    Else
        MsgBox "No se ubicaron Requerimientos de Producción Pendientes.", vbInformation + vbOKOnly, wnomcia
    End If
    
    SqlCad = vbNullString
    
    frameProgreso.Visible = False
    
    
    If rstResumen.State = 1 Then rstResumen.Close
    
    Set rstResumen = Nothing
    
    Exit Function
errImportarResumenRequerimientoProduccion:
    Select Case Err.Number
        Case -2147217871
'            If MsgBox("No.: " & Err.Number & vbNewLine & _
'                        "Descripción: " & Err.Description & vbNewLine & vbNewLine & _
'                        "Se ha perdido la conexion al Servidor Externo, ¿Desea intentar volver a conectarse?", vbQuestion + vbYesNo, App.ProductName & " - Importacion de Requerimientos de Producción Pendientes") = vbYes Then
'
'                abrirCnDBMilano
'
                
                If dblCantidadIntentos <= 5 Then
                    Resume
                Else
                    dblCantidadIntentos = dblCantidadIntentos + 1
                End If
                
'            End If
        Case Else
            MsgBox "No.: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description & _
                    "Vuelva a Intentarlo.", vbInformation + vbOKOnly, App.ProductName & " - Importacion de Requerimientos de Producción Pendientes"
    End Select
    
    'Resume
    importarResumenRequerimientoProduccion = False
    
    frameProgreso.Visible = False
    barraProgreso.Value = 0
    
    Err.Clear
End Function

Public Function importarResumenRequerimientoProduccionV2(ByVal frameProgreso As Object, _
                                                        ByVal barraProgreso As Object, _
                                                        ByVal strNombreTablaSQLParaResumen As String, _
                                                        Optional ByVal strNroPedido As String, _
                                                        Optional ByVal strCodigoProducto As String, _
                                                        Optional ByVal strFiltroProducto As String, _
                                                        Optional ByVal strFechaEntregaCorte As String, _
                                                        Optional ByVal bolSoloDescargarEnServidorSQL As Boolean) As Boolean
    
    On Error GoTo errImportarResumenRequerimientoProduccionV2
    
    
    Dim cmdOP As ADODB.Command
    Dim rstResumen As New ADODB.Recordset
    Dim dblCantidadRegistro As Double
    Dim dblCantidadIntentos As Double
    
'    Dim bolProductoNoRegistradoEnCP As Boolean
    
    importarResumenRequerimientoProduccionV2 = False
    
    Rem LECTURA DE RESUMEN DE REQUERIMIENTO DE PRODUCCION (EN BASE ORDENES DE PRODUCCION)
    
    'abrirCnDBMilano
    
    dblCantidadIntentos = 0
    
    Set cmdOP = New ADODB.Command
    
    With cmdOP
        .ActiveConnection = cnBdStudioModa
        .CommandType = adCmdStoredProc
        .CommandTimeout = "180"
        .CommandText = "usp_ConsultaSaldosOrdenesProduccionCPv3"
        
        .Parameters.Append .CreateParameter("@NROPEDIDO", adVarChar, adParamInput, 10, strNroPedido)
        .Parameters.Append .CreateParameter("@CODIGOPRODUCTO", adVarChar, adParamInput, 50, strCodigoProducto)
        .Parameters.Append .CreateParameter("@FILTROPRODUCTO", adVarChar, adParamInput, 255, strFiltroProducto)
        .Parameters.Append .CreateParameter("@FECHAENTREGACORTE", adVarChar, adParamInput, 20, strFechaEntregaCorte)
        .Parameters.Append .CreateParameter("@NOMBRETABLA", adVarChar, adParamInput, 255, strNombreTablaSQLParaResumen)
        .Parameters.Append .CreateParameter("@SOLOINSUMODESCARGADODEOP", adInteger, adParamInput, , 0)
        .Parameters.Append .CreateParameter("@CANTIDADFILAS", adInteger, adParamOutput, 5)
        
        .Execute
        
        dblCantidadRegistro = Val(.Parameters("@CANTIDADFILAS") & "")
    End With
    
    Set cmdOP = Nothing
    
    If bolSoloDescargarEnServidorSQL Then
        
        importarResumenRequerimientoProduccionV2 = True
        
        Exit Function
    End If
    
    abrirCnTemporal
    
    cnDBTemp.Execute "DELETE FROM TMPUTILRESUMENREQUERIMIENTOPRODUCCION"
    
    If dblCantidadRegistro > 0 Then
        If rstResumen.State = 1 Then rstResumen.Close
        
        rstResumen.Open "SELECT * FROM " & strNombreTablaSQLParaResumen, cnBdStudioModa, adOpenForwardOnly, adLockReadOnly
        
        DoEvents
        
        frameProgreso.Visible = True
        frameProgreso.Caption = "Ejecutando consulta (1/4)..."
        
        If Not rstResumen.EOF Then
            DoEvents
            
            frameProgreso.Caption = "Contabilizando registros extraidos (1/4)..."
            barraProgreso.Max = dblCantidadRegistro
            barraProgreso.Value = 0
            
            DoEvents
            
            frameProgreso.Caption = "Importando Requerimientos de Producción Pendientes (1/4)..."
            
            Do While Not rstResumen.EOF
                SqlCad = vbNullString
                SqlCad = SqlCad & "INSERT INTO TMPUTILRESUMENREQUERIMIENTOPRODUCCION("
                SqlCad = SqlCad & "LLAVE, "
                SqlCad = SqlCad & "NROPEDIDO, "
                SqlCad = SqlCad & "CLIENTE, "
                SqlCad = SqlCad & "FEMISION, "
                SqlCad = SqlCad & "FENTREGA, "
                SqlCad = SqlCad & "VENDEDOR, "
                SqlCad = SqlCad & "CODPRODUCTO, "
                SqlCad = SqlCad & "NOMPRODUCTO, "
                SqlCad = SqlCad & "UM, "
                SqlCad = SqlCad & "NOMPRODUCTOUM, "
                SqlCad = SqlCad & "IDOP, "
                SqlCad = SqlCad & "CATEGORIA, "
                SqlCad = SqlCad & "NROOP, "
                SqlCad = SqlCad & "MODELO, "
                SqlCad = SqlCad & "COLOR, "
                SqlCad = SqlCad & "CANTIDADPEDIDO, "
                SqlCad = SqlCad & "DESCRIPCIONOP, "
                SqlCad = SqlCad & "OBSERVACIONOP, "
                SqlCad = SqlCad & "CANTIDAD, "
                SqlCad = SqlCad & "SALDO) "
                
                SqlCad = SqlCad & "VALUES("
                SqlCad = SqlCad & "'" & Trim(rstResumen!LLAVE & "") & "', "
                SqlCad = SqlCad & "'" & Trim(rstResumen!NroPedido & "") & "', "
                SqlCad = SqlCad & "'" & Trim(rstResumen!CLIENTE & "") & "', "
                SqlCad = SqlCad & "CVDATE('" & Trim(rstResumen!FEMISION & "") & "'), "
                SqlCad = SqlCad & "CVDATE('" & Trim(rstResumen!FENTREGA & "") & "'), "
                SqlCad = SqlCad & "'" & Trim(rstResumen!VENDEDOR & "") & "', "
                SqlCad = SqlCad & "'" & Trim(rstResumen!CodProducto & "") & "', "
                SqlCad = SqlCad & "'" & Trim(rstResumen!NOMPRODUCTO & "") & "', "
                SqlCad = SqlCad & "'" & Trim(rstResumen!um & "") & "', "
                SqlCad = SqlCad & "'" & Trim(rstResumen!NOMPRODUCTOUM & "") & "', "
                SqlCad = SqlCad & "'" & Trim(rstResumen!IDOP & "") & "', "
                SqlCad = SqlCad & "'" & Trim(rstResumen!CATEGORIA & "") & "', "
                SqlCad = SqlCad & "'" & Trim(rstResumen!NroOP & "") & "', "
                SqlCad = SqlCad & "'" & Trim(rstResumen!Modelo & "") & "', "
                SqlCad = SqlCad & "'" & Trim(rstResumen!Color & "") & "', "
                SqlCad = SqlCad & Val(Format(Val(rstResumen!CANTIDADPEDIDO & ""), "#0.00")) & ", "
                SqlCad = SqlCad & "'" & Trim(rstResumen!DESCRIPCIONOP & "") & "', "
                SqlCad = SqlCad & "'" & Trim(rstResumen!OBSERVACIONOP & "") & "', "
                SqlCad = SqlCad & Val(Format(Val(rstResumen!Cantidad & ""), "#0.00")) & ", "
                SqlCad = SqlCad & Val(Format(Val(rstResumen!SALDO & ""), "#0.00")) & ")"
                
                abrirCnTemporal
                
                cnDBTemp.Execute SqlCad
                
'                If Not bolProductoNoRegistradoEnCP Then
'                    abrirCnnDbBancos
'
'                    If ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5CODPRO", "IF5PLA", "F5CODPRO", Trim(rstResumen!CodProducto & ""), "T") = vbNullString Then
'                        bolProductoNoRegistradoEnCP = True
'                    End If
'                End If
                
                DoEvents
                
                barraProgreso.Value = barraProgreso.Value + 1
                frameProgreso.Caption = "Importando Requerimientos de Producción Pendientes (1/4)... " & FormatPercent(barraProgreso.Value / barraProgreso.Max, 3)
                
                rstResumen.MoveNext
            Loop
                importarResumenRequerimientoProduccionV2 = True
        Else
            MsgBox "No se ubicaron Requerimientos de Producción Pendientes.", vbInformation + vbOKOnly, wnomcia
        End If
        
'        If bolProductoNoRegistradoEnCP Then
'            importarInsumoServidorExterno frameProgreso, barraProgreso, True
'        End If
    Else
        MsgBox "No se ubicaron Requerimientos de Producción Pendientes.", vbInformation + vbOKOnly, wnomcia
    End If
    
    SqlCad = vbNullString
    
    frameProgreso.Visible = False
    
    If rstResumen.State = 1 Then rstResumen.Close
    
    Set rstResumen = Nothing
    
    Exit Function
errImportarResumenRequerimientoProduccionV2:
    Select Case Err.Number
        Case -2147217871
'            If MsgBox("No.: " & Err.Number & vbNewLine & _
'                        "Descripción: " & Err.Description & vbNewLine & vbNewLine & _
'                        "Se ha perdido la conexion al Servidor Externo, ¿Desea intentar volver a conectarse?", vbQuestion + vbYesNo, App.ProductName & " - Importacion de Requerimientos de Producción Pendientes") = vbYes Then
'
'                abrirCnDBMilano
'
                
                If dblCantidadIntentos <= 5 Then
                    Resume
                Else
                    dblCantidadIntentos = dblCantidadIntentos + 1
                End If
                
'            End If
        Case Else
            MsgBox "No.: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description & _
                    "Vuelva a Intentarlo.", vbInformation + vbOKOnly, App.ProductName & " - Importacion de Requerimientos de Producción Pendientes"
    End Select
    
    'Resume
    importarResumenRequerimientoProduccionV2 = False
    
    If Not frameProgreso Is Nothing Then
        frameProgreso.Visible = False
    End If
    
    If Not barraProgreso Is Nothing Then
        barraProgreso.Value = 0
    End If
    
    Err.Clear
End Function

Public Function importarResumenRequerimientoProduccionV3(ByVal frameProgreso1 As Object, _
                                                        ByVal barraProgreso1 As Object, _
                                                        ByVal frameProgreso2 As Object, _
                                                        ByVal barraProgreso2 As Object, _
                                                        ByVal strNombreTablaSQLParaResumen As String, _
                                                        Optional ByVal strNroPedido As String, _
                                                        Optional ByVal strCodigoProducto As String, _
                                                        Optional ByVal strFiltroProducto As String, _
                                                        Optional ByVal strFechaEntregaCorte As String, _
                                                        Optional ByVal bolSoloDescargarEnServidorSQL As Boolean) As Boolean
    
    On Error GoTo errImportarResumenRequerimientoProduccionV3
    
    Dim cmdOP As ADODB.Command
    Dim rstPedidoValido As New ADODB.Recordset
    Dim rstResumen As New ADODB.Recordset
    Dim dblCantidadRegistro As Double
    Dim dblCantidadIntentos As Double
    
'    Dim bolProductoNoRegistradoEnCP As Boolean
    
    importarResumenRequerimientoProduccionV3 = False
    
    Rem LECTURA DE RESUMEN DE REQUERIMIENTO DE PRODUCCION (EN BASE ORDENES DE PRODUCCION)
    
    abrirCnDBMilano
    
    dblCantidadIntentos = 0
    
    abrirCnTemporal
    
    cnDBTemp.Execute "DELETE FROM TMPUTILRESUMENREQUERIMIENTOPRODUCCION"
    
    If rstPedidoValido.State = 1 Then rstPedidoValido.Close
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "IDPEDIDO, "
    SqlCad = SqlCad & "FECHAENTREGA "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "PEDIDO "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "ANULADO = 0 AND "
    SqlCad = SqlCad & "DBO.FECHA(FECHAENTREGA) >= '" & strFechaEntregaCorte & "' "
        
        If strNroPedido <> vbNullString Then
            SqlCad = SqlCad & "AND IDPEDIDO = '" & strNroPedido & "' "
        End If
        
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "DBO.FECHA(FECHAENTREGA)"
    
    rstPedidoValido.Open SqlCad, cnBdStudioModa, adOpenForwardOnly, adLockReadOnly
    
    If Not rstPedidoValido.EOF Then
        DoEvents
        
        frameProgreso1.Caption = "Contabilizando registros extraidos..."
        barraProgreso1.Max = ModUtilitario.devuelveCantRegistros(rstPedidoValido) 'dblCantidadRegistro
        barraProgreso1.Value = 0
        
        DoEvents
        
        frameProgreso1.Caption = "Leyendo Requerimiento de Producción N° " & Trim(rstPedidoValido!IDPEDIDO & "") & " - Fec. Entrega: " & Format(Trim(rstPedidoValido!FechaEntrega & ""), "Short Date") & " ..."
        
        Do While Not rstPedidoValido.EOF
            Set cmdOP = New ADODB.Command
            
            With cmdOP
                .ActiveConnection = cnBdStudioModa
                .CommandType = adCmdStoredProc
                .CommandTimeout = "180"
                .CommandText = "usp_ConsultaSaldosOrdenesProduccionCPv3"
                
                .Parameters.Append .CreateParameter("@NROPEDIDO", adVarChar, adParamInput, 10, Trim(rstPedidoValido!IDPEDIDO & ""))
                .Parameters.Append .CreateParameter("@CODIGOPRODUCTO", adVarChar, adParamInput, 50, strCodigoProducto)
                .Parameters.Append .CreateParameter("@FILTROPRODUCTO", adVarChar, adParamInput, 255, strFiltroProducto)
                .Parameters.Append .CreateParameter("@FECHAENTREGACORTE", adVarChar, adParamInput, 20, vbNullString)
                .Parameters.Append .CreateParameter("@NOMBRETABLA", adVarChar, adParamInput, 255, strNombreTablaSQLParaResumen)
                .Parameters.Append .CreateParameter("@SOLOINSUMODESCARGADODEOP", adInteger, adParamInput, , 0)
                .Parameters.Append .CreateParameter("@CANTIDADFILAS", adInteger, adParamOutput, 5)
                
                .Execute
                
                dblCantidadRegistro = Val(.Parameters("@CANTIDADFILAS") & "")
            End With
            
            Set cmdOP = Nothing
            
            If dblCantidadRegistro > 0 Then
                If rstResumen.State = 1 Then rstResumen.Close
                
                rstResumen.Open "SELECT * FROM " & strNombreTablaSQLParaResumen & " WHERE SALDO > 0", cnBdStudioModa, adOpenForwardOnly, adLockReadOnly
                
                DoEvents
                
                frameProgreso2.Visible = True
                frameProgreso2.Caption = "Ejecutando consulta..."
                
                If Not rstResumen.EOF Then
                    DoEvents
                    
                    frameProgreso2.Caption = "Contabilizando registros extraidos (1/4)..."
                    barraProgreso2.Max = ModUtilitario.devuelveCantRegistros(rstResumen) 'dblCantidadRegistro
                    barraProgreso2.Value = 0
                    
                    DoEvents
                    
                    frameProgreso2.Caption = "Importando Requerimiento de Producción N° " & Trim(rstResumen!NroPedido & "") & " ..."
                    
                    Do While Not rstResumen.EOF
                        SqlCad = vbNullString
                        SqlCad = SqlCad & "INSERT INTO TMPUTILRESUMENREQUERIMIENTOPRODUCCION("
                        SqlCad = SqlCad & "LLAVE, "
                        SqlCad = SqlCad & "NROPEDIDO, "
                        SqlCad = SqlCad & "CLIENTE, "
                        SqlCad = SqlCad & "FEMISION, "
                        SqlCad = SqlCad & "FENTREGA, "
                        SqlCad = SqlCad & "VENDEDOR, "
                        SqlCad = SqlCad & "CODPRODUCTO, "
                        SqlCad = SqlCad & "NOMPRODUCTO, "
                        SqlCad = SqlCad & "UM, "
                        SqlCad = SqlCad & "NOMPRODUCTOUM, "
                        SqlCad = SqlCad & "IDOP, "
                        SqlCad = SqlCad & "CATEGORIA, "
                        SqlCad = SqlCad & "NROOP, "
                        SqlCad = SqlCad & "MODELO, "
                        SqlCad = SqlCad & "COLOR, "
                        SqlCad = SqlCad & "CANTIDADPEDIDO, "
                        SqlCad = SqlCad & "DESCRIPCIONOP, "
                        SqlCad = SqlCad & "OBSERVACIONOP, "
                        SqlCad = SqlCad & "CANTIDAD, "
                        SqlCad = SqlCad & "SALDO) "
                        
                        SqlCad = SqlCad & "VALUES("
                        SqlCad = SqlCad & "'" & Trim(rstResumen!LLAVE & "") & "', "
                        SqlCad = SqlCad & "'" & Trim(rstResumen!NroPedido & "") & "', "
                        SqlCad = SqlCad & "'" & Trim(rstResumen!CLIENTE & "") & "', "
                        SqlCad = SqlCad & "CVDATE('" & Trim(rstResumen!FEMISION & "") & "'), "
                        SqlCad = SqlCad & "CVDATE('" & Trim(rstResumen!FENTREGA & "") & "'), "
                        SqlCad = SqlCad & "'" & Trim(rstResumen!VENDEDOR & "") & "', "
                        SqlCad = SqlCad & "'" & Trim(rstResumen!CodProducto & "") & "', "
                        SqlCad = SqlCad & "'" & Trim(rstResumen!NOMPRODUCTO & "") & "', "
                        SqlCad = SqlCad & "'" & Trim(rstResumen!um & "") & "', "
                        SqlCad = SqlCad & "'" & Trim(rstResumen!NOMPRODUCTOUM & "") & "', "
                        SqlCad = SqlCad & "'" & Trim(rstResumen!IDOP & "") & "', "
                        SqlCad = SqlCad & "'" & Trim(rstResumen!CATEGORIA & "") & "', "
                        SqlCad = SqlCad & "'" & Trim(rstResumen!NroOP & "") & "', "
                        SqlCad = SqlCad & "'" & Trim(rstResumen!Modelo & "") & "', "
                        SqlCad = SqlCad & "'" & Trim(rstResumen!Color & "") & "', "
                        SqlCad = SqlCad & Val(Format(Val(rstResumen!CANTIDADPEDIDO & ""), "#0.00")) & ", "
                        SqlCad = SqlCad & "'" & Trim(rstResumen!DESCRIPCIONOP & "") & "', "
                        SqlCad = SqlCad & "'" & Trim(rstResumen!OBSERVACIONOP & "") & "', "
                        SqlCad = SqlCad & Val(Format(Val(rstResumen!Cantidad & ""), "#0.00")) & ", "
                        SqlCad = SqlCad & Val(Format(Val(rstResumen!SALDO & ""), "#0.00")) & ")"
                        
                        abrirCnTemporal
                        
                        cnDBTemp.Execute SqlCad
                        
                        DoEvents
                        
                        barraProgreso2.Value = barraProgreso2.Value + 1
                        frameProgreso2.Caption = "Importando Requerimiento de Producción N° " & Trim(rstResumen!NroPedido & "") & " ... " & FormatPercent(barraProgreso2.Value / barraProgreso2.Max, 3)
                        
                        rstResumen.MoveNext
                    Loop
                        importarResumenRequerimientoProduccionV3 = True
                End If
            End If
            
            DoEvents
            
            barraProgreso1.Value = barraProgreso1.Value + 1
            frameProgreso1.Caption = "Leyendo Requerimiento de Producción N° " & Trim(rstPedidoValido!IDPEDIDO & "") & " - Fec. Entrega: " & Format(Trim(rstPedidoValido!FechaEntrega & ""), "Short Date") & " ... " & FormatPercent(barraProgreso1.Value / barraProgreso1.Max, 3)
            
            rstPedidoValido.MoveNext
        Loop
    End If
    
    SqlCad = vbNullString
    
    frameProgreso2.Visible = False
    
    If rstResumen.State = 1 Then rstResumen.Close
    
    Set rstResumen = Nothing
    
    If rstPedidoValido.State = 1 Then rstPedidoValido.Close
    
    Set rstPedidoValido = Nothing
    
    Exit Function
errImportarResumenRequerimientoProduccionV3:
    Select Case Err.Number
        Case -2147217871
                dblCantidadIntentos = dblCantidadIntentos + 1
                
                If dblCantidadIntentos <= 5 Then
                    Resume
                End If
        Case Else
            MsgBox "No.: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description & _
                    "Vuelva a Intentarlo.", vbInformation + vbOKOnly, App.ProductName & " - Importacion de Requerimientos de Producción Pendientes"
    End Select
    
    importarResumenRequerimientoProduccionV3 = False
    
    If Not frameProgreso2 Is Nothing Then
        frameProgreso2.Visible = False
    End If
    
    If Not barraProgreso2 Is Nothing Then
        barraProgreso2.Value = 0
    End If
    
    Err.Clear
End Function

Public Function devolverSaldoRequerimientoProduccion(ByVal strNroPedido As String, _
                                                        ByVal strIdInsumo As String) As Double
    
    On Error GoTo errDevolverSaldoRequerimientoProduccion
    
    Dim rstResumen As New ADODB.Recordset
    
    devolverSaldoRequerimientoProduccion = 0
    
    Rem LECTURA DE RESUMEN DE REQUERIMIENTO DE PRODUCCION (EN BASE ORDENES DE PRODUCCION)
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "RQ.IDPEDIDO AS NROPEDIDO, "
    SqlCad = SqlCad & "DET.IDINSUMO AS CODPRODUCTO, "
    SqlCad = SqlCad & "SUM( DET.CANTIDAD ) AS CANTIDAD, "
    SqlCad = SqlCad & "SUM( ((DET.CANTIDAD - ISNULL(MOVOP.CANTIDAD, 0)) + ISNULL(MOVOPINGRESO.CANTIDAD, 0)) ) AS SALDO "
    
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "ORDENPRODUCCIONDESCARGO AS DET "
    SqlCad = SqlCad & "LEFT JOIN ORDENPRODUCCION AS CAB ON CAB.IDORDENPRODUCCION = DET.IDORDENPRODUCCION "
    
    SqlCad = SqlCad & "LEFT JOIN "
    SqlCad = SqlCad & "(SELECT "
    SqlCad = SqlCad & "IDREQUERIMIENTO, (CASE WHEN IDPEDIDO = 0 THEN NROPEDIDO ELSE IDPEDIDO END) AS IDPEDIDO, NROMODELO, IDCOLOR "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "REQUERIMIENTO) AS RQ "
    SqlCad = SqlCad & "ON RQ.IDREQUERIMIENTO = CAB.IDREQUERIMIENTO "
    
    SqlCad = SqlCad & "LEFT JOIN PEDIDO AS PED ON PED.IDPEDIDO = RQ.IDPEDIDO "
    SqlCad = SqlCad & "LEFT JOIN PERSONA AS PER ON PER.IDPERSONA = PED.IDPERSONA "
    
    SqlCad = SqlCad & "LEFT JOIN "
    SqlCad = SqlCad & "(SELECT "
    SqlCad = SqlCad & "CAB.IDORDENPRODUCCION, "
    SqlCad = SqlCad & "RQ.IDPEDIDO, "
    SqlCad = SqlCad & "DET.IDINSUMO, "
    SqlCad = SqlCad & "SUM(DET.CANTIDAD) AS CANTIDAD "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "SALIDADETALLE AS DET "
    SqlCad = SqlCad & "LEFT JOIN SALIDA AS CAB "
    SqlCad = SqlCad & "ON CAB.IDSALIDA = DET.IDSALIDA "
    SqlCad = SqlCad & "LEFT JOIN ORDENPRODUCCION AS OP ON OP.IDORDENPRODUCCION = CAB.IDORDENPRODUCCION "
    SqlCad = SqlCad & "LEFT JOIN "
    SqlCad = SqlCad & "(SELECT "
    SqlCad = SqlCad & "IDREQUERIMIENTO, "
    SqlCad = SqlCad & "(CASE WHEN IDPEDIDO = 0 THEN NROPEDIDO ELSE IDPEDIDO END) AS IDPEDIDO "
    SqlCad = SqlCad & "FROM REQUERIMIENTO) AS RQ "
    SqlCad = SqlCad & "ON RQ.IDREQUERIMIENTO = OP.IDREQUERIMIENTO "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "CAB.IDTIPOSALIDA IN ('TIP0000001') AND "
    SqlCad = SqlCad & "CAB.ANULADO = 0 AND "
    SqlCad = SqlCad & "RQ.IDPEDIDO = '" & strNroPedido & "' AND "
    SqlCad = SqlCad & "DET.IDINSUMO = '" & strIdInsumo & "' "
    
    SqlCad = SqlCad & "GROUP BY "
    SqlCad = SqlCad & "CAB.IDORDENPRODUCCION, "
    SqlCad = SqlCad & "RQ.IDPEDIDO, DET.IDINSUMO) AS MOVOP "
    SqlCad = SqlCad & "ON MOVOP.IDORDENPRODUCCION = DET.IDORDENPRODUCCION AND MOVOP.IDPEDIDO = RQ.IDPEDIDO AND MOVOP.IDINSUMO = DET.IDINSUMO "
    
    SqlCad = SqlCad & "LEFT JOIN "
    SqlCad = SqlCad & "(SELECT "
    SqlCad = SqlCad & "CAB.IDORDENPRODUCCION, "
    SqlCad = SqlCad & "RQ.IDPEDIDO, "
    SqlCad = SqlCad & "DET.IDINSUMO, "
    SqlCad = SqlCad & "SUM(DET.CANTIDAD) AS CANTIDAD "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "INGRESODETALLE AS DET "
    SqlCad = SqlCad & "LEFT JOIN INGRESO AS CAB "
    SqlCad = SqlCad & "ON CAB.IDINGRESO = DET.IDINGRESO "
    SqlCad = SqlCad & "LEFT JOIN ORDENPRODUCCION AS OP ON OP.IDORDENPRODUCCION = CAB.IDORDENPRODUCCION "
    SqlCad = SqlCad & "LEFT JOIN "
    SqlCad = SqlCad & "(SELECT "
    SqlCad = SqlCad & "IDREQUERIMIENTO, "
    SqlCad = SqlCad & "(CASE WHEN IDPEDIDO = 0 THEN NROPEDIDO ELSE IDPEDIDO END) AS IDPEDIDO "
    SqlCad = SqlCad & "FROM REQUERIMIENTO) AS RQ "
    SqlCad = SqlCad & "ON RQ.IDREQUERIMIENTO = OP.IDREQUERIMIENTO "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "CAB.IDTIPOINGRESO IN ('TIP0000004') AND "
    SqlCad = SqlCad & "CAB.ANULADO = 0 AND "
    SqlCad = SqlCad & "RQ.IDPEDIDO = '" & strNroPedido & "' AND "
    SqlCad = SqlCad & "DET.IDINSUMO = '" & strIdInsumo & "' "
        
    SqlCad = SqlCad & "GROUP BY "
    SqlCad = SqlCad & "CAB.IDORDENPRODUCCION, "
    SqlCad = SqlCad & "RQ.IDPEDIDO, DET.IDINSUMO) AS MOVOPINGRESO "
    SqlCad = SqlCad & "ON MOVOPINGRESO.IDORDENPRODUCCION = DET.IDORDENPRODUCCION AND MOVOPINGRESO.IDPEDIDO = RQ.IDPEDIDO AND MOVOPINGRESO.IDINSUMO = DET.IDINSUMO "
    
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "DET.ANULADO = 0 AND "
    SqlCad = SqlCad & "CAB.CERRADO = 0 AND "
    SqlCad = SqlCad & "CAB.ANULADO = 0 AND "
    
    SqlCad = SqlCad & "RTRIM(LTRIM(PER.NOMBRE + '')) <> '' AND "
    SqlCad = SqlCad & "RQ.IDPEDIDO = '" & strNroPedido & "' AND "
    SqlCad = SqlCad & "DET.IDINSUMO = '" & strIdInsumo & "' "
    
    SqlCad = SqlCad & "GROUP BY "
    SqlCad = SqlCad & "RQ.IDPEDIDO, "
    SqlCad = SqlCad & "DET.IDINSUMO"
    
    'abrirCnDBMilano
    
    If rstResumen.State = 1 Then rstResumen.Close
    
    'rstResumen.Open SqlCad, cnBdStudioModa, adOpenDynamic, adLockOptimistic
    
    Set rstResumen = cnBdStudioModa.Execute(SqlCad, , cmdText)
    
    If Not rstResumen.EOF Then
        'rstResumen.MoveFirst
        
        devolverSaldoRequerimientoProduccion = Val(rstResumen!SALDO & "")
    Else
        devolverSaldoRequerimientoProduccion = 0
    End If
    
    SqlCad = vbNullString
    
    If rstResumen.State = 1 Then rstResumen.Close
    
    Set rstResumen = Nothing
    
    Exit Function
errDevolverSaldoRequerimientoProduccion:
    Select Case Err.Number
        Case -2147217871
            If MsgBox("No.: " & Err.Number & vbNewLine & _
                        "Descripción: " & Err.Description & vbNewLine & vbNewLine & _
                        "Se ha perdido la conexion al Servidor Externo, ¿Desea intentar volver a conectarse?", vbQuestion + vbYesNo, App.ProductName & " - Importacion de Requerimientos de Producción Pendientes") = vbYes Then
                        
                abrirCnDBMilano
                
                Resume
            End If
        Case Else
            MsgBox "No.: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName & " - Importacion de Requerimientos de Producción Pendientes"
    End Select
    
    'Resume
    devolverSaldoRequerimientoProduccion = 0
    
    Err.Clear
End Function

Public Function devolverUltimoCostoCompraEnIntegradoEnSoles(ByVal strIdInsumo As String) As Double
    
    On Error GoTo errDevolverUltimoCostoCompraEnIntegradoEnSoles
    
    Dim rstResumen As New ADODB.Recordset
    
    devolverUltimoCostoCompraEnIntegradoEnSoles = 0
    
    Rem LECTURA DE RESUMEN DE REQUERIMIENTO DE PRODUCCION (EN BASE ORDENES DE PRODUCCION)
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT TOP 1 "
    SqlCad = SqlCad & "CAB.IDINGRESO, "
    SqlCad = SqlCad & "CAB.IDMONEDA, "
    SqlCad = SqlCad & "CAB.TCAMBIO, "
    SqlCad = SqlCad & "CAB.FECHA, "
    SqlCad = SqlCad & "DET.IDINSUMO, "
    SqlCad = SqlCad & "DET.COSTO, "
    SqlCad = SqlCad & "DET.COSTO * IIF(CAB.IDMONEDA = 'MON0000001', 1, CAB.TCAMBIO) AS COSTOFINAL "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "INGRESODETALLE AS DET "
    SqlCad = SqlCad & "LEFT JOIN INGRESO AS CAB ON CAB.IDINGRESO = DET.IDINGRESO "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "CAB.IDTIPOINGRESO = 'TIP0000001' AND "
    SqlCad = SqlCad & "CAB.FECHA <= '31/12/2014' AND "
    SqlCad = SqlCad & "DET.IDINSUMO = '" & strIdInsumo & "' "
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "CAB.FECHA DESC"
    
    'abrirCnDBMilano
    
    If rstResumen.State = 1 Then rstResumen.Close
    
    'rstResumen.Open SqlCad, cnBdStudioModa, adOpenDynamic, adLockOptimistic
    
    Set rstResumen = cnBdStudioModa.Execute(SqlCad, , cmdText)
    
    If Not rstResumen.EOF Then
        'rstResumen.MoveFirst
        
        devolverUltimoCostoCompraEnIntegradoEnSoles = Val(rstResumen!COSTOFINAL & "")
    Else
        devolverUltimoCostoCompraEnIntegradoEnSoles = 0
    End If
    
    SqlCad = vbNullString
    
    If rstResumen.State = 1 Then rstResumen.Close
    
    Set rstResumen = Nothing
    
    Exit Function
errDevolverUltimoCostoCompraEnIntegradoEnSoles:
    Select Case Err.Number
        Case -2147217871
            If MsgBox("No.: " & Err.Number & vbNewLine & _
                        "Descripción: " & Err.Description & vbNewLine & vbNewLine & _
                        "Se ha perdido la conexion al Servidor Externo, ¿Desea intentar volver a conectarse?", vbQuestion + vbYesNo, App.ProductName & " - Importacion de Requerimientos de Producción Pendientes") = vbYes Then
                        
                abrirCnDBMilano
                
                Resume
            End If
        Case Else
            MsgBox "No.: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName & " - Importacion de Requerimientos de Producción Pendientes"
    End Select
    
    'Resume
    devolverUltimoCostoCompraEnIntegradoEnSoles = 0
    
    Err.Clear
End Function

'Public Sub listarCategoriaTipo(ByVal comboList As Object)
'    On Error GoTo errListarCategoriaTipo
'
'    Dim rstCategoriaTipo As New ADODB.Recordset
'
'    sql = vbNullString
'    sql = sql & "SELECT IDCATEGORIATIPO, NOMBRE " & _
'                    "FROM CATEGORIATIPO " & _
'                    "WHERE ELIMINADO = 0 " & _
'                    "ORDER BY NOMBRE"
'    'abrirCnDBMilano
'
'    If rstCategoriaTipo.State = 1 Then rstCategoriaTipo.Close
'
'    rstCategoriaTipo.Open sql, cnBdStudioModa, adOpenForwardOnly, adLockReadOnly
'
'    If Not rstCategoriaTipo.EOF Then
'        comboList.Clear
'
'        Do While Not rstCategoriaTipo.EOF
'            comboList.AddItem Trim(rstCategoriaTipo!nombre & "") '& Space(150) & Trim(rstCategoriaTipo!IDCATEGORIATIPO & "")
'
'            rstCategoriaTipo.MoveNext
'        Loop
'            comboList.ListIndex = -1
'    End If
'
'    rstCategoriaTipo.Close
'
'    Set rstCategoriaTipo = Nothing
'    sql = vbNullString
'
'    Set rstCategoriaTipo = Nothing
'
'    Exit Sub
'    Resume
'errListarCategoriaTipo:
'    MsgBox "No. Error: " & Err.Number & vbNewLine & _
'            "Descripción: " & Err.Description, vbCritical, App.ProductName & " - ModMilano: ListarCategoriaTipo"
'
'    Err.Clear
'End Sub

'VISUALIZAR DETALLE DE CONSOLIDADO DE INSUMOS DE PEDIDO
Public Sub visualizarDetalleConsolidadoPedido(ByVal grilla As dxDBGrid, _
                                                ByVal strNroPedido As String, _
                                                ByVal strCodProducto As String)
    
    On Error GoTo errVisualizarDetalleConsolidadoPedido
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "(INS.NOMBRE + ' (' + UM.NOMBRE + ')') AS  INSUMO, "
    SqlCad = SqlCad & "('       [No. Modelo: ' + MODELO.TIPOMODELO + '-' + CAST(MODELO.IDMODELO AS VARCHAR) + ']           [Categoria: ' + CT.NOMBRE + ']          [Cantidad: ' + CAST(CAB.TOTALCANTIDAD AS VARCHAR) + ']') AS INFO, "
    SqlCad = SqlCad & "DET.IDREQUERIMIENTO, "
    SqlCad = SqlCad & "COLOR.NOMBRE AS COLORCUERO, "
    SqlCad = SqlCad & "(SUM(CAST(DET.CANTIDAD AS DECIMAL(7,2)) * (CASE WHEN LEN(SFAM.BASETALLA) > 0 AND SFAM.BASETALLA <> 'CINTURON' THEN 1 ELSE CAST(CAB.TOTALCANTIDAD AS DECIMAL(7,2)) END))) AS CANTTOTAL, "
    SqlCad = SqlCad & "UM.NOMBRE AS UM, "
    SqlCad = SqlCad & "IIF(CAB.ANULADO = 0, 'ACTIVO', 'ANULADO') AS ESTADO "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "REQUERIMIENTODETALLE AS DET "
    SqlCad = SqlCad & "LEFT JOIN REQUERIMIENTO AS CAB ON CAB.IDREQUERIMIENTO = DET.IDREQUERIMIENTO "
    SqlCad = SqlCad & "LEFT JOIN INSUMO AS INS ON INS.IDINSUMO = DET.IDINSUMO "
    SqlCad = SqlCad & "LEFT JOIN UNIDADMEDIDA AS UM ON UM.IDUNIDADMEDIDA = INS.IDUNIDADMEDIDA "
    SqlCad = SqlCad & "LEFT JOIN SUBFAMILIA AS SFAM ON SFAM.IDSUBFAMILIA = INS.IDSUBFAMILIA "
    SqlCad = SqlCad & "LEFT JOIN COLOR ON COLOR.IDCOLOR = CAB.IDCOLOR "
    SqlCad = SqlCad & "LEFT JOIN MODELO ON MODELO.NROMODELO = CAB.NROMODELO "
    SqlCad = SqlCad & "LEFT JOIN CATEGORIATIPO AS CT ON CT.IDCATEGORIATIPO = MODELO.IDCATEGORIATIPO "
    SqlCad = SqlCad & "WHERE "
    'SqlCad = SqlCad & "CAB.ANULADO = 0 AND "
    SqlCad = SqlCad & "CAB.IDPEDIDO = '" & strNroPedido & "' AND "
    SqlCad = SqlCad & "DET.IDINSUMO = '" & strCodProducto & "' "
    SqlCad = SqlCad & "GROUP BY "
    SqlCad = SqlCad & "DET.IDREQUERIMIENTO, "
    SqlCad = SqlCad & "COLOR.NOMBRE, "
    SqlCad = SqlCad & "('       [No. Modelo: ' + MODELO.TIPOMODELO + '-' + CAST(MODELO.IDMODELO AS VARCHAR) + ']           [Categoria: ' + CT.NOMBRE + ']          [Cantidad: ' + CAST(CAB.TOTALCANTIDAD AS VARCHAR) + ']'), "
    SqlCad = SqlCad & "CT.NOMBRE, "
    SqlCad = SqlCad & "CAB.TOTALCANTIDAD, "
    SqlCad = SqlCad & "DET.IDINSUMO, "
    SqlCad = SqlCad & "INS.NOMBRE, "
    SqlCad = SqlCad & "UM.NOMBRE, "
    SqlCad = SqlCad & "IIF(CAB.ANULADO = 0, 'ACTIVO', 'ANULADO') "
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "DET.IDREQUERIMIENTO, INS.NOMBRE "
    
    If Not grilla Is Nothing Then
        With grilla
            .Dataset.Close
             
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        
        With grilla
            'Columna Insumo
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Producto"
                .DisableEditor = True
                .FieldName = "INSUMO"
                .Font.Bold = True
                .FontColor = vbWhite
                .HeaderAlignment = taCenter
                .ObjectName = "ColInsumo"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
            End With
            
            'Columna Información de Modelo
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Informacion de Modelo"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "INFO"
                .Font.Bold = True
                .FontColor = vbWhite
                .HeaderAlignment = taCenter
                .ObjectName = "ColInfo"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
            End With
            
            'Columna ID Requerimiento
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "ID. Requerimiento"
                .DisableEditor = True
                .FieldName = "IDREQUERIMIENTO"
                '.Font.Bold = True
                .HeaderAlignment = taCenter
                .ObjectName = "ColIDREQUERIMIENTO"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
            End With
            
            'Columna Color de Cuero
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Color de Cuero"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "COLORCUERO"
                '.Font.Bold = True
                .HeaderAlignment = taCenter
                .ObjectName = "ColColorCuero"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 200
            End With
            
            'Columna Cantidad
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 0
                .Caption = "Cantidad"
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "CANTTOTAL"
                '.Font.Bold = True
                .HeaderAlignment = taCenter
                .ObjectName = "ColCantTotal"
                '.SummaryFooterType = cstSum
                .ObjectName = "ColCantidad"
                .SummaryFooterType = cstSum
                '.SummaryFooterFormat = " "
                .Width = 100
            End With
            
            'Columna Unidad de Medida
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "U.M."
                .DisableEditor = True
                .FieldName = "UM"
                '.Font.Bold = True
                .HeaderAlignment = taCenter
                .ObjectName = "ColUM"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 60
            End With
            
            'Columna Estado
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Estado"
                .DisableEditor = True
                .FieldName = "ESTADO"
                '.Font.Bold = True
                .HeaderAlignment = taCenter
                .ObjectName = "ColEstado"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'abrirCnDBMilano
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = strCadenaConexioBdStudioModa
            
            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctStatic
            .Dataset.ADODataset.LockType = ltReadOnly
            .Dataset.ADODataset.CommandText = SqlCad
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "IDREQUERIMIENTO"
            
            .Columns.ColumnByFieldName("INSUMO").GroupIndex = 0
            .Columns.ColumnByFieldName("INFO").GroupIndex = 1
            
            .m.FullExpand
            
            .Bands.BandByName("bndPrincipal").Caption = "No. Pedido: " & strNroPedido & " / " & ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CS_OBSERVACIONES", "TB_CABSOLICITUD", "COD_SOLICITUD", strNroPedido, "T")
            .Columns.HeaderFont.Bold = True
            .GroupNodeColor = RGB(75, 172, 198)
        End With
    End If
    
    SqlCad = vbNullString
    
    Exit Sub
errVisualizarDetalleConsolidadoPedido:
    Select Case Err.Number
        Case 3704, 3709
            'cnn_dbbancos.Open StrConexDbBancos
            abrirCnDBMilano
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ModMilano: VisualizarDetalleConsolidadoPedido"
    End Select
    
    Err.Clear
End Sub

'VISUALIZAR ORDENES DE PRODUCCION DE REQUERIMIENTO
Public Sub visualizarOPdeRequerimiento(ByVal grilla As dxDBGrid, _
                                        ByVal strIdRequerimiento As String)
    
    On Error GoTo errVisualizarOPdeRequerimiento
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "CAB.IDORDENPRODUCCION, "
    SqlCad = SqlCad & "CAT.NOMBRE AS CATEGORIA, "
    SqlCad = SqlCad & "CAB.OP, "
    SqlCad = SqlCad & "CAB.DESCRIPCION, "
    SqlCad = SqlCad & "(MODELO.TIPOMODELO + '-' + CAST(MODELO.IDMODELO AS VARCHAR)) AS ARTICULO, "
    SqlCad = SqlCad & "COL.NOMBRE AS COLOR, "
    SqlCad = SqlCad & "CONVERT(VARCHAR(10), CAB.FECHA, 103) AS FECHA, " 'CAB.FECHA, "
    SqlCad = SqlCad & "CAB.CANTIDADTOTAL, "
    SqlCad = SqlCad & "IIF(CAB.ANULADO = 0, 'ACTIVO', 'ANULADO') AS ESTADO "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "ORDENPRODUCCION AS CAB "
    SqlCad = SqlCad & "LEFT JOIN CATEGORIATIPO AS CAT ON CAT.IDCATEGORIATIPO = CAB.IDCATEGORIATIPO "
    SqlCad = SqlCad & "LEFT JOIN REQUERIMIENTO AS REQ ON REQ.IDREQUERIMIENTO = CAB.IDREQUERIMIENTO "
    SqlCad = SqlCad & "LEFT JOIN MODELO AS MODELO ON MODELO.NROMODELO = REQ.NROMODELO "
    SqlCad = SqlCad & "LEFT JOIN COLOR AS COL ON COL.IDCOLOR = REQ.IDCOLOR "
    SqlCad = SqlCad & "WHERE "
    'SqlCad = SqlCad & "CAB.ANULADO = 0 AND "
    SqlCad = SqlCad & "CAB.IDREQUERIMIENTO = " & strIdRequerimiento & " "
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "CAB.FECHA, CAB.IDORDENPRODUCCION"
    
    If Not grilla Is Nothing Then
        With grilla
            .Dataset.Close
             
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        
        With grilla
            'Columna Id Orden de Produccion
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "ID"
                .DisableEditor = True
                .FieldName = "IDORDENPRODUCCION"
                .HeaderAlignment = taCenter
                .ObjectName = "ColIdOrdenProduccion"
                .Visible = False
            End With
            
            'Columna Categoria de Orden de Produccion
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Categoria"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "CATEGORIA"
                .HeaderAlignment = taCenter
                .ObjectName = "ColCategoria"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 90
            End With
            
            'Columna Numero de Orden de Produccion
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "O/P"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "OP"
                .HeaderAlignment = taCenter
                .ObjectName = "ColOP"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
            End With
            
            'Columna Descripcion
            Set gColumn = .Columns.Add(gedMemoEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Descripcion"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "DESCRIPCION"
                .HeaderAlignment = taCenter
                .ObjectName = "ColDescripcion"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 230
            End With
            
            'Columna Articulo (Modelo)
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Modelo"
                .DisableEditor = True
                .FieldName = "ARTICULO"
                .HeaderAlignment = taCenter
                .ObjectName = "ColModelo"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
            End With
            
            'Columna Color de Cuero
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Color"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "COLOR"
                .HeaderAlignment = taCenter
                .ObjectName = "ColColor"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
            End With
            
            'Columna Fecha de O/P
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Fecha"
                .DisableEditor = True
                .FieldName = "FECHA"
                .HeaderAlignment = taCenter
                .ObjectName = "ColFecha"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 60
            End With
            
            'Columna Cantidad Total
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 0
                .Caption = "Cantidad Total"
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "CANTIDADTOTAL"
                .HeaderAlignment = taCenter
                .ObjectName = "ColCantidadTotal"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
            End With
            
            'Columna Estado
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Estado"
                .DisableEditor = True
                .FieldName = "ESTADO"
                .HeaderAlignment = taCenter
                .ObjectName = "ColEstado"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
            End With
            
            'abrirCnDBMilano
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = strCadenaConexioBdStudioModa
            
            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctStatic
            .Dataset.ADODataset.LockType = ltReadOnly
            .Dataset.ADODataset.CommandText = SqlCad
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "IDORDENPRODUCCION"
            
            .Bands.BandByName("bndPrincipal").Caption = "Ordenes de Producción de Requerimiento: " & strIdRequerimiento
            .Columns.HeaderFont.Bold = True
        End With
    End If
    
    SqlCad = vbNullString
    
    Exit Sub
errVisualizarOPdeRequerimiento:
    Select Case Err.Number
        Case 3704, 3709
            'cnn_dbbancos.Open StrConexDbBancos
            cnBdStudioModa.Open cnBdStudioModa.ConnectionString
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ModMilano: VisualizarOPdeRequerimiento"
    End Select
    
    Err.Clear
End Sub

'VISUALIZAR ORDENES DE PRODUCCION DE REQUERIMIENTO
Public Sub visualizarOPDetalle(ByVal grilla As dxDBGrid, _
                                ByVal strIdOrdenProduccion As String)
    
    On Error GoTo errVisualizarOPDetalle
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "DET.ITEMS, "
    SqlCad = SqlCad & "DET.IDINSUMO, "
    SqlCad = SqlCad & "INS.NOMBRE, "
    SqlCad = SqlCad & "UM.NOMBRE AS UM, "
    SqlCad = SqlCad & "CAST(DET.CANTIDAD AS DECIMAL(7,2)) AS CANTIDAD "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "ORDENPRODUCCIONDESCARGO AS DET "
    SqlCad = SqlCad & "LEFT JOIN ORDENPRODUCCION AS CAB ON CAB.IDORDENPRODUCCION = DET.IDORDENPRODUCCION "
    SqlCad = SqlCad & "LEFT JOIN REQUERIMIENTO AS REQ ON REQ.IDREQUERIMIENTO = CAB.IDREQUERIMIENTO "
    SqlCad = SqlCad & "LEFT JOIN INSUMO AS INS ON INS.IDINSUMO = DET.IDINSUMO "
    SqlCad = SqlCad & "LEFT JOIN UNIDADMEDIDA AS UM ON UM.IDUNIDADMEDIDA = INS.IDUNIDADMEDIDA "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "DET.IDORDENPRODUCCION = " & strIdOrdenProduccion & " AND "
    SqlCad = SqlCad & "DET.ANULADO = 0 "
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "DET.ITEMS"
    
    If Not grilla Is Nothing Then
        With grilla
            .Dataset.Close
             
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        
        With grilla
            'Columna Items
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 0
                .Caption = "Item"
                .DisableEditor = True
                .FieldName = "ITEMS"
                .HeaderAlignment = taCenter
                .ObjectName = "ColItems"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 40
            End With
            
            'Columna Id Insumo
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Codigo"
                .DisableEditor = True
                .FieldName = "IDINSUMO"
                .HeaderAlignment = taCenter
                .ObjectName = "ColIdInsumo"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
            End With
            
            'Columna Descripcion
            Set gColumn = .Columns.Add(gedMemoEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Descripcion de Insumo"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "NOMBRE"
                .HeaderAlignment = taCenter
                .ObjectName = "ColNombre"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 230
            End With
            
            'Columna UM
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "U.M."
                .DisableEditor = True
                .FieldName = "UM"
                .HeaderAlignment = taCenter
                .ObjectName = "ColUM"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 60
            End With
            
            'Columna Cantidad Total
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 0
                .Caption = "Cantidad"
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "CANTIDAD"
                .HeaderAlignment = taCenter
                .ObjectName = "ColCantidadTotal"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
            End With
            
            'abrirCnDBMilano
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = strCadenaConexioBdStudioModa
            
            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctStatic
            .Dataset.ADODataset.LockType = ltReadOnly
            .Dataset.ADODataset.CommandText = SqlCad
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "IDINSUMO"
            
            .Bands.BandByName("bndPrincipal").Caption = "Detalle de Orden de Producción: " & ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "(CAT.NOMBRE + ' / ' + OP)", "ORDENPRODUCCION AS OP LEFT JOIN CATEGORIATIPO AS CAT ON CAT.IDCATEGORIATIPO = OP.IDCATEGORIATIPO", "OP.IDORDENPRODUCCION", strIdOrdenProduccion, "N")
            .Columns.HeaderFont.Bold = True
        End With
    End If
    
    SqlCad = vbNullString
    
    Exit Sub
errVisualizarOPDetalle:
    Select Case Err.Number
        Case 3704, 3709
            'cnn_dbbancos.Open StrConexDbBancos
            abrirCnDBMilano
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ModMilano: VisualizarOPDetalle"
    End Select
    
    Err.Clear
End Sub

Public Function modificarProductoEnOP(ByVal strIdOrdenProduccion As String, _
                                        ByVal strCodProductoOriginal As String, _
                                        ByVal strCodProductoFinal As String, _
                                        ByVal dblCantidadOriginal As Double, _
                                        ByVal dblCantidadFinal As Double, _
                                        ByVal strObservacion As String) As Boolean
    
    On Error GoTo errModificarProductoEnOP
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "UPDATE "
    SqlCad = SqlCad & "ORDENPRODUCCIONDESCARGO "
    SqlCad = SqlCad & "SET "
    
        If strCodProductoOriginal <> strCodProductoFinal Then
            SqlCad = SqlCad & "IDINSUMO = '" & strCodProductoFinal & "', "
        End If
        
    SqlCad = SqlCad & "CANTIDAD = " & dblCantidadFinal & " "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "IDORDENPRODUCCION = " & strIdOrdenProduccion & " AND "
    SqlCad = SqlCad & "IDINSUMO = '" & strCodProductoOriginal & "'"
    
    'abrirCnDBMilano
    
    cnBdStudioModa.Execute SqlCad
    
    modificarProductoEnOP = True
    
    Actualiza_Log " < DB Externo > " & SqlCad, StrConexDbBancos
    
    'If strCodProductoOriginal <> strCodProductoFinal Then
        Dim lngIdCambio As Long
        Dim strIdUsuario As String
        Dim strCategoria As String
        Dim strNumeroOP As String
        
        lngIdCambio = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "TOP 1 IDCAMBIO", "SF1ORDENPRODUCCION_LOG", "IDORDENPRODUCCION", strIdOrdenProduccion, "T", "ORDER BY IDCAMBIO DESC") & "") + 1
        strIdUsuario = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODUSEREXTERNO", "EF2USERS", "F2CODUSER", wusuario, "T")
        
        strCategoria = ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "IDCATEGORIATIPO", "ORDENPRODUCCION", "IDORDENPRODUCCION", strIdOrdenProduccion, "N")
        strCategoria = ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "NOMBRE", "CATEGORIATIPO", "IDCATEGORIATIPO", strCategoria, "T")
        strNumeroOP = ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "OP", "ORDENPRODUCCION", "IDORDENPRODUCCION", strIdOrdenProduccion, "N")
        
        SqlCad = vbNullString
        SqlCad = SqlCad & "INSERT INTO SF1ORDENPRODUCCION_LOG("
        SqlCad = SqlCad & "IDORDENPRODUCCION, IDCAMBIO, IDINSUMO, IDINSUMOFINAL, "
        SqlCad = SqlCad & "CANTIDAD, CANTIDADFINAL, IDUSUARIO, FECHAMODIFICACION, "
        SqlCad = SqlCad & "OBSERVACION, CATEGORIA, NUMEROOP"
        SqlCad = SqlCad & ") "
        SqlCad = SqlCad & "VALUES("
        SqlCad = SqlCad & "'" & strIdOrdenProduccion & "', "
        SqlCad = SqlCad & lngIdCambio & ", "
        SqlCad = SqlCad & "'" & strCodProductoOriginal & "', "
        SqlCad = SqlCad & "'" & strCodProductoFinal & "', "
        SqlCad = SqlCad & dblCantidadOriginal & ", "
        SqlCad = SqlCad & dblCantidadFinal & ", "
        SqlCad = SqlCad & "'" & IIf(strIdUsuario <> vbNullString, strIdUsuario, wusuario) & "', "
        SqlCad = SqlCad & "CVDATE('" & Now & "'), "
        SqlCad = SqlCad & "'" & strObservacion & "', "
        SqlCad = SqlCad & "'" & strCategoria & "', "
        SqlCad = SqlCad & "'" & strNumeroOP & "'"
        SqlCad = SqlCad & ")"
        
        cnn_dbbancos.Execute SqlCad
        
        Actualiza_Log SqlCad, StrConexDbBancos
        
        lngIdCambio = 0
        strIdUsuario = vbNullString
    'End If
    
    SqlCad = vbNullString
    
    Exit Function
errModificarProductoEnOP:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbCritical, App.ProductName & " - ModMilano: ModificarProductoEnOP"
    
    modificarProductoEnOP = False
    
    Err.Clear
End Function

Public Function verificarCierreDeMesEnServidorExterno(ByVal intAnnoCorte As Integer, _
                                                        ByVal intMesCorte As Integer, _
                                                        ByVal strIdAlmacen As String) As Boolean
    
    On Error GoTo errVerificarCierreDeMesEnServidorExterno
    
    Dim rstCierre As ADODB.Recordset
    
    Set rstCierre = Nothing
    
    Set rstCierre = New ADODB.Recordset
    
    'abrirCnDBMilano
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "* "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "CIERREMENSUAL AS CM "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "CM.IDALMACEN = '" & strIdAlmacen & "' AND "
    SqlCad = SqlCad & "CM.AÑO = " & intAnnoCorte & " AND "
    SqlCad = SqlCad & "CM.MES = " & intMesCorte
    
    If rstCierre.State = 1 Then rstCierre.Close
    
    rstCierre.Open SqlCad, cnBdStudioModa, adOpenForwardOnly, adLockReadOnly
    
    If Not rstCierre.EOF Then
        verificarCierreDeMesEnServidorExterno = True
    Else
        verificarCierreDeMesEnServidorExterno = False
    End If
    
    If rstCierre.State = 1 Then rstCierre.Close
    
    Set rstCierre = Nothing
    
    Exit Function
errVerificarCierreDeMesEnServidorExterno:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    verificarCierreDeMesEnServidorExterno = True
    
    Err.Clear
End Function

'PROCEDIMIENTO USADO PARA CARGA DEL HISTORIAL POR PROVEEDOR EXTRAIDO A PARTIR DE LOS INGRESOS POR COMPRAS DE LA BASE DE MILANO
Public Sub cargarProductoProveedor()
    On Error GoTo errCargarProductoProveedor
    
    Dim rstHistorialExterno As New ADODB.Recordset
    Dim strRucProveedor As String
    Dim dblTotal As Double
    Dim dblItem As Double
    
    abrirCnTemporal
    
    SQL1 = vbNullString
    SQL1 = SQL1 & "SELECT F2CODPRV, F2NOMPRV, F5CODPRO, F5NOMPRO, F5VALVTA, F7CODMED, F2FECHA, F2MONEDA FROM EF2PROD_PROV ORDER BY F2CODPRV, F5CODPRO, F2FECHA"
    
    If rstHistorialExterno.State = 1 Then rstHistorialExterno.Close
    
    rstHistorialExterno.Open SQL1, cnDBTemp, adOpenForwardOnly, adLockReadOnly
    
    If Not rstHistorialExterno.EOF Then
        'rstHistorialExterno.MoveFirst
        dblTotal = ModUtilitario.devuelveCantRegistros(rstHistorialExterno)
        dblItem = 0
        
        Do While Not rstHistorialExterno.EOF
            Rem SK ADD: Actualizar Historial de Proveedor (EF2PROD_PROV)
            dblItem = dblItem + 1
            
            SQL1 = vbNullString
            
            strRucProveedor = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NEWRUC", "EF2PROVEEDORES", "F2CODPROVEXTERNO", Trim(rstHistorialExterno!F2CODPRV & ""), "T")
            
            If ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5CODPRO", "EF2PROD_PROV", "F2CODPRV", strRucProveedor, "T", "AND F5CODPRO = '" & Trim(rstHistorialExterno!f5codpro & "") & "'") = vbNullString Then
                SQL1 = SQL1 & "INSERT INTO EF2PROD_PROV("
                SQL1 = SQL1 & "F2CODPRV, F2NOMPRV, F5CODPRO, F5NOMPRO, F5VALVTA, "
                SQL1 = SQL1 & "F7CODMED, F2FECHA, F2MONEDA) "
                SQL1 = SQL1 & "VALUES ("
                SQL1 = SQL1 & "'" & strRucProveedor & "', "
                SQL1 = SQL1 & "'" & Trim(rstHistorialExterno!F2NOMPRV & "") & "', "
                SQL1 = SQL1 & "'" & Trim(rstHistorialExterno!f5codpro & "") & "', "
                SQL1 = SQL1 & "'" & Trim(rstHistorialExterno!F5NOMPRO & "") & "', "
                SQL1 = SQL1 & Val(rstHistorialExterno!F5VALVTA & "") & ", "
                SQL1 = SQL1 & "'" & Trim(rstHistorialExterno!f7codmed & "") & "', "
                SQL1 = SQL1 & "CVDATE('" & Format(Trim(rstHistorialExterno!F2FECHA & ""), "Short Date") & "'), "
                SQL1 = SQL1 & "'" & Trim(rstHistorialExterno!F2MONEDA & "") & "')"
            Else
                SQL1 = SQL1 & "UPDATE EF2PROD_PROV "
                SQL1 = SQL1 & "SET "
                SQL1 = SQL1 & "F2NOMPRV = '" & Trim(rstHistorialExterno!F2NOMPRV & "") & "', "
                SQL1 = SQL1 & "F5NOMPRO = '" & Trim(rstHistorialExterno!F5NOMPRO & "") & "', "
                SQL1 = SQL1 & "F5VALVTA = " & Val(rstHistorialExterno!F5VALVTA & "") & ", "
                'SQL1 = SQL1 & "F5CODFAB = '" & Trim(rstHistorialExterno!f5codpro & "") & "', "
                SQL1 = SQL1 & "F7CODMED = '" & Trim(rstHistorialExterno!f7codmed & "") & "', "
                SQL1 = SQL1 & "F2FECHA = CVDATE('" & Format(Trim(rstHistorialExterno!F2FECHA & ""), "Short Date") & "'), "
                SQL1 = SQL1 & "F2MONEDA = '" & Trim(rstHistorialExterno!F2MONEDA & "") & "' "
                SQL1 = SQL1 & "WHERE "
                SQL1 = SQL1 & "F2CODPRV = '" & strRucProveedor & "' AND "
                SQL1 = SQL1 & "F5CODPRO = '" & Trim(rstHistorialExterno!f5codpro & "") & "' AND "
                SQL1 = SQL1 & "F2FECHA <= CVDATE('" & Format(Trim(rstHistorialExterno!F2FECHA & ""), "Short Date") & "')"
            End If
            
            cnn_dbbancos.Execute SQL1
            AlmacenaQuery_sql SQL1, cnn_dbbancos
            Actualiza_Log SQL1, cnn_dbbancos.ConnectionString
            
            Debug.Print dblItem & " de " & dblTotal
            
            rstHistorialExterno.MoveNext
        Loop
            MsgBox "Termine SK :)", vbInformation + vbOKOnly, App.ProductName
    End If
    
    Exit Sub
errCargarProductoProveedor:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    'Resume
    Err.Clear
End Sub

'Visualizar Estado de Atencion Actual de Orden de Producción
Public Sub vistaEstadoOrdenProduccion(ByVal grilla As dxDBGrid, _
                                        ByVal strTabla As String, _
                                        ByVal strCodProducto As String, _
                                        ByVal strFiltroSensitivo As String, _
                                        ByVal strNroPedidoFiltro As String)
    
    On Error GoTo errVistaEstadoOrdenProduccion
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "IDOP, "
    SqlCad = SqlCad & "CATEGORIA, "
    SqlCad = SqlCad & "NROOP, "
    SqlCad = SqlCad & "MODELO, "
    SqlCad = SqlCad & "COLOR, "
    SqlCad = SqlCad & "NROPEDIDO, "
    SqlCad = SqlCad & "CLIENTE, "
    SqlCad = SqlCad & "FENTREGA, "
    SqlCad = SqlCad & "DESCRIPCIONOP, "
    SqlCad = SqlCad & "OBSERVACIONOP "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & strTabla & " "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "IDOP <> 0 "
        
        If strCodProducto <> vbNullString Then
            SqlCad = SqlCad & "AND CODPRODUCTO = '" & strCodProducto & "' "
        End If
        
        If strFiltroSensitivo <> vbNullString Then
            SqlCad = SqlCad & "AND ("
            SqlCad = SqlCad & "CATEGORIA  LIKE '%" & strFiltroSensitivo & "%' OR "
            SqlCad = SqlCad & "NROOP  LIKE '%" & strFiltroSensitivo & "%' OR "
            SqlCad = SqlCad & "MODELO  LIKE '%" & strFiltroSensitivo & "%' OR "
            SqlCad = SqlCad & "CLIENTE LIKE '%" & strFiltroSensitivo & "%'"
            SqlCad = SqlCad & ") "
        End If
        
        If strNroPedidoFiltro <> vbNullString Then
            SqlCad = SqlCad & "AND NROPEDIDO = '" & strNroPedidoFiltro & "' "
        End If
        
    SqlCad = SqlCad & "GROUP BY "
    SqlCad = SqlCad & "IDOP, "
    SqlCad = SqlCad & "CATEGORIA, "
    SqlCad = SqlCad & "NROOP, "
    SqlCad = SqlCad & "MODELO, "
    SqlCad = SqlCad & "COLOR, "
    SqlCad = SqlCad & "NROPEDIDO, "
    SqlCad = SqlCad & "CLIENTE, "
    SqlCad = SqlCad & "FENTREGA, "
    SqlCad = SqlCad & "DESCRIPCIONOP, "
    SqlCad = SqlCad & "OBSERVACIONOP "
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "FENTREGA DESC"
    
    If Not grilla Is Nothing Then
        With grilla
            .Dataset.Close
             
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        
        With grilla
            'Columna IdOP
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "ID"
                .DisableEditor = True
                .FieldName = "IDOP"
                .Font.Bold = True
                .FontColor = vbWhite
                .HeaderAlignment = taCenter
                .ObjectName = "ColIdOP"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 60
                .Visible = False
            End With
            
            'Columna Categoria
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .BandIndex = 0
                .Caption = "Categoria"
                .Color = vbWhite
                .DisableEditor = True
                .FieldName = "CATEGORIA"
                .Font.Bold = True
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCategoria"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
            End With
            
            'Columna No. OP
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "No. OP"
                .Color = vbWhite
                .DisableEditor = True
                .FieldName = "NROOP"
                .Font.Bold = True
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColNroOP"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
            End With
            
            'Columna Modelo
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Modelo"
                .Color = &HC000C0
                .DisableEditor = True
                .FieldName = "MODELO"
                .Font.Bold = True
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColNroOP"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
            End With
            
            'Columna Color
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .BandIndex = 0
                .Caption = "Color"
                .Color = vbWhite
                .DisableEditor = True
                .FieldName = "COLOR"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColColor"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
            End With
            
            'Columna No Pedido
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "No. Pedido"
                .Color = vbWhite
                .DisableEditor = True
                .FieldName = "NROPEDIDO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColNoPedido"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 60
            End With
            
            'Columna Cliente
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .BandIndex = 0
                .Caption = "Cliente"
                .Color = vbWhite
                .DisableEditor = True
                .FieldName = "CLIENTE"
                .HeaderAlignment = taCenter
                .ObjectName = "ColCliente"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 120
            End With
            
            'Columna Fecha de Entrega
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Fec. Entrega"
                .Color = vbWhite
                .DisableEditor = True
                .FieldName = "FENTREGA"
                .HeaderAlignment = taCenter
                .ObjectName = "ColFecEntrega"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Descripcion de OP
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .BandIndex = 0
                .Caption = "Descripcion"
                .Color = vbWhite
                .DisableEditor = True
                .FieldName = "DESCRIPCIONOP"
                .HeaderAlignment = taCenter
                .ObjectName = "ColDescripcion"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 300
            End With
            
            'Columna Observacion de OP
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .BandIndex = 0
                .Caption = "Observacion"
                .Color = vbWhite
                .DisableEditor = True
                .FieldName = "OBSERVACIONOP"
                .HeaderAlignment = taCenter
                .ObjectName = "ColDescripcion"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
            End With
            
            'abrirCnDBMilano
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = strCadenaConexioBdStudioModa
             
            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctStatic
            .Dataset.ADODataset.LockType = ltReadOnly
            .Dataset.ADODataset.CommandText = SqlCad
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "IDOP"
            
            '.Bands.BandByName("bndPrincipal").Caption = "No. Pedido: " & strNroPedido & " / " & ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CS_OBSERVACIONES", "TB_CABSOLICITUD", "COD_SOLICITUD", strNroPedido, "T")
            '.Columns.HeaderFont.Bold = True
            '.GroupNodeColor = RGB(75, 172, 198)
        End With
    End If
    
    SqlCad = vbNullString
    
    Exit Sub
errVistaEstadoOrdenProduccion:
    Select Case Err.Number
        Case 3704, 3709
            abrirCnDBMilano
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ModMilano: VistaEstadoOrdenProduccion"
    End Select
    
    Err.Clear
End Sub

'Visualizar Estado de Atencion Actual de Orden de Producción - Detalle de OP
Public Sub vistaEstadoOrdenProduccionDetalle(ByVal grilla As dxDBGrid, _
                                                ByVal strTabla As String, _
                                                ByVal strIdOrdenProduccion As String, _
                                                ByVal strCodProducto As String)
    
    On Error GoTo errVistaEstadoOrdenProduccionDetalle
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "IDOP, "
    SqlCad = SqlCad & "CODPRODUCTO, "
    SqlCad = SqlCad & "NOMPRODUCTO, "
    SqlCad = SqlCad & "UM, "
    SqlCad = SqlCad & "CAST(CANTIDAD AS NUMERIC(10,2)) AS CANTIDAD, "
    SqlCad = SqlCad & "CAST(SALDO AS NUMERIC(10,2)) AS SALDO "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & strTabla & " "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "IDOP = " & strIdOrdenProduccion
    
        If strCodProducto <> vbNullString Then
            SqlCad = SqlCad & " AND CODPRODUCTO = '" & strCodProducto & "'"
        End If
        
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "NOMPRODUCTO"
    
    If Not grilla Is Nothing Then
        With grilla
            .Dataset.Close
             
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        
        With grilla
            'Columna IdOP
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "ID"
                .Color = vbWhite
                .DisableEditor = True
                .FieldName = "IDOP"
                .Font.Bold = True
                .FontColor = vbWhite
                .HeaderAlignment = taCenter
                .ObjectName = "ColIdOP"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 60
                .Visible = False
            End With
            
            'Columna Codigo de Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .BandIndex = 0
                .Caption = "Codigo"
                .Color = vbWhite
                .DisableEditor = True
                .FieldName = "CODPRODUCTO"
                .Font.Bold = True
                '.FontColor = vbWhite
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodigoProducto"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Descripcion de Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .BandIndex = 0
                .Caption = "Descripcion de Producto"
                .Color = vbWhite
                .DisableEditor = True
                .FieldName = "NOMPRODUCTO"
                .Font.Bold = True
                '.FontColor = vbWhite
                .HeaderAlignment = taCenter
                .ObjectName = "ColDescripcionProducto"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 120
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
                .HeaderAlignment = taCenter
                .ObjectName = "ColUM"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 60
            End With
            
            'Columna Cantidad
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 0
                .Caption = "Cantidad"
                .Color = vbWhite
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "CANTIDAD"
                .HeaderAlignment = taCenter
                .ObjectName = "ColCantidad"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
            End With
            
            'Columna Saldo
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 0
                .Caption = "Pendiente"
                .Color = vbWhite
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "SALDO"
                .HeaderAlignment = taCenter
                .ObjectName = "ColPendiente"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
            End With
            
            'abrirCnDBMilano
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = strCadenaConexioBdStudioModa
            
            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctStatic
            .Dataset.ADODataset.LockType = ltReadOnly
            .Dataset.ADODataset.CommandText = SqlCad
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "CODPRODUCTO"
        End With
    End If
    
    SqlCad = vbNullString
    
    Exit Sub
errVistaEstadoOrdenProduccionDetalle:
    Select Case Err.Number
        Case 3704, 3709
            abrirCnDBMilano
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ModMilano: VistaEstadoOrdenProduccionDetalle"
    End Select
    
    Err.Clear
End Sub

'Visualizar Estado de Atencion Actual de Orden de Producción - Detalle de OP
Public Sub vistaPedidoDetalle(ByVal grilla As dxDBGrid, _
                                ByVal strAnno As String, _
                                ByVal strFiltroSensitivo As String)
    
    On Error GoTo errVistaPedidoDetalle
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "PED.IDPEDIDO, "
    SqlCad = SqlCad & "PER.NOMBRE AS CLIENTE, "
    SqlCad = SqlCad & "CONVERT(CHAR(10), DBO.FECHA(PED.FECHAEMISION), 103) AS FECEMISION, "
    SqlCad = SqlCad & "CONVERT(CHAR(10), DBO.FECHA(PED.FECHAENTREGA), 103) AS FECENTREGA, "
    SqlCad = SqlCad & "USU.NOMBRE AS VENDEDOR, "
    SqlCad = SqlCad & "(PER.NOMBRE + ' ( FEC. PEDIDO: ' + CONVERT(CHAR(10), PED.FECHAEMISION, 103) + ' / FEC. ENTREGA: ' + CONVERT(CHAR(10), PED.FECHAENTREGA, 103) + ')') AS RESUMEN "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "PEDIDO AS PED "
    SqlCad = SqlCad & "LEFT JOIN PERSONA AS PER ON PER.IDPERSONA = PED.IDPERSONA "
    SqlCad = SqlCad & "LEFT JOIN USUARIO AS USU ON USU.IDUSUARIO = PED.IDUSUARIOINGRESO "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "PED.ANULADO = 0 "
        
        If strAnno <> vbNullString Then
            SqlCad = SqlCad & "AND YEAR(DBO.FECHA(PED.FECHAENTREGA)) = " & strAnno & " "
        End If
        
        If strFiltroSensitivo <> vbNullString Then
            SqlCad = SqlCad & "AND ("
            SqlCad = SqlCad & "PED.IDPEDIDO LIKE '" & strFiltroSensitivo & "%' OR "
            SqlCad = SqlCad & "PER.NOMBRE LIKE '%" & strFiltroSensitivo & "%'"
            SqlCad = SqlCad & ")"
        End If
        
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "DBO.FECHA(PED.FECHAENTREGA)"
    
    If Not grilla Is Nothing Then
        With grilla
            .Dataset.Close
             
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        
        With grilla
            'Columna ID de Pedido
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "No. Pedido"
                .Color = vbWhite
                .DisableEditor = True
                .FieldName = "IDPEDIDO"
                .Font.Bold = True
                '.FontColor = vbWhite
                .HeaderAlignment = taCenter
                .ObjectName = "ColIdPedido"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 60
            End With
            
            'Columna Cliente
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .BandIndex = 0
                .Caption = "Cliente"
                .Color = vbWhite
                .DisableEditor = True
                .FieldName = "CLIENTE"
                .Font.Bold = True
                '.FontColor = vbWhite
                .HeaderAlignment = taCenter
                .ObjectName = "ColCliente"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
            End With
            
            'Columna Fecha Emision
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .BandIndex = 0
                .Caption = "Fec. Emision"
                .Color = vbWhite
                .DisableEditor = True
                .FieldName = "FECEMISION"
                .Font.Bold = True
                '.FontColor = vbWhite
                .HeaderAlignment = taCenter
                .ObjectName = "ColFecEmision"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 60
            End With
            
            'Columna Fecha Entrega
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Fec. Entrega"
                .Color = vbWhite
                .DisableEditor = True
                .FieldName = "FECENTREGA"
                .HeaderAlignment = taCenter
                .ObjectName = "ColFecEntrega"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 60
            End With
            
            'Columna Vendedor
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Colocado Por"
                .Color = vbWhite
                .DisableEditor = True
                .FieldName = "VENDEDOR"
                .HeaderAlignment = taCenter
                .ObjectName = "ColVendedor"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
            End With
            
            'Columna Resumen
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Resumen"
                .Color = vbWhite
                .DisableEditor = True
                .FieldName = "RESUMEN"
                .HeaderAlignment = taCenter
                .ObjectName = "ColResumen"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
                .Visible = False
            End With
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = strCadenaConexioBdStudioModa
            
            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctStatic
            .Dataset.ADODataset.LockType = ltReadOnly
            .Dataset.ADODataset.CommandText = SqlCad
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "IDPEDIDO"
        End With
    End If
    
    SqlCad = vbNullString
    
    Exit Sub
errVistaPedidoDetalle:
    Select Case Err.Number
        Case 3704, 3709
            abrirCnDBMilano
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ModMilano: VistaPedidoDetalle"
    End Select
    
    Err.Clear
End Sub

Public Sub visualizarOrdenProduccionAnulada(ByVal grilla As dxDBGrid, _
                                            ByVal strAnno As String, _
                                            ByVal strMes As String, _
                                            ByVal StrFiltro As String)
    
    On Error GoTo errVisualizarOrdenProduccionAnulada
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "CAB.IDORDENPRODUCCION, "
    SqlCad = SqlCad & "ISNULL(REQ.IDPEDIDO, REQ.NROPEDIDO) AS PEDIDO, "
    SqlCad = SqlCad & "CAT.NOMBRE AS CATEGORIA, "
    SqlCad = SqlCad & "CAB.OP, "
    SqlCad = SqlCad & "CAB.DESCRIPCION, "
    SqlCad = SqlCad & "(MODELO.TIPOMODELO + '-' + CAST(MODELO.IDMODELO AS VARCHAR)) AS ARTICULO, "
    SqlCad = SqlCad & "COL.NOMBRE AS COLOR, "
    SqlCad = SqlCad & "CONVERT(VARCHAR(10), CAB.FECHA, 103) AS FECHA, " 'CAB.FECHA, "
    SqlCad = SqlCad & "CAB.CANTIDADTOTAL, "
    SqlCad = SqlCad & "IIF(CAB.ANULADO = 0, 'ACTIVO', 'ANULADO') AS ESTADO "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "ORDENPRODUCCION AS CAB "
    SqlCad = SqlCad & "LEFT JOIN CATEGORIATIPO AS CAT ON CAT.IDCATEGORIATIPO = CAB.IDCATEGORIATIPO "
    SqlCad = SqlCad & "LEFT JOIN REQUERIMIENTO AS REQ ON REQ.IDREQUERIMIENTO = CAB.IDREQUERIMIENTO "
    SqlCad = SqlCad & "LEFT JOIN MODELO AS MODELO ON MODELO.NROMODELO = REQ.NROMODELO "
    SqlCad = SqlCad & "LEFT JOIN COLOR AS COL ON COL.IDCOLOR = REQ.IDCOLOR "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "CAB.ANULADO = 1 AND "
    SqlCad = SqlCad & "YEAR(DBO.FECHA(FECHA)) = " & strAnno & " AND "
    SqlCad = SqlCad & "MONTH(DBO.FECHA(FECHA)) = " & Val(strMes) & " "
    
        If StrFiltro <> vbNullString Then
            SqlCad = SqlCad & "AND (ISNULL(REQ.IDPEDIDO, REQ.NROPEDIDO) LIKE '%" & StrFiltro & "%' OR CAB.OP LIKE '%" & StrFiltro & "%') "
        End If
        
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "CAB.FECHA, "
    SqlCad = SqlCad & "CAB.IDORDENPRODUCCION"
    
    If Not grilla Is Nothing Then
        With grilla
            .Dataset.Close
             
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        
        With grilla
            'Columna Id Orden de Produccion
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "ID"
                .DisableEditor = True
                .FieldName = "IDORDENPRODUCCION"
                .HeaderAlignment = taCenter
                .ObjectName = "ColIdOrdenProduccion"
                .Visible = False
            End With
            
            'Columna Pedid0
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "No Pedido"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "PEDIDO"
                .HeaderAlignment = taCenter
                .ObjectName = "ColPedido"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 60
            End With
            
            'Columna Categoria de Orden de Produccion
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Categoria"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "CATEGORIA"
                .HeaderAlignment = taCenter
                .ObjectName = "ColCategoria"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 90
            End With
            
            'Columna Numero de Orden de Produccion
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "O/P"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "OP"
                .HeaderAlignment = taCenter
                .ObjectName = "ColOP"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
            End With
            
            'Columna Descripcion
            Set gColumn = .Columns.Add(gedMemoEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Descripcion"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "DESCRIPCION"
                .HeaderAlignment = taCenter
                .ObjectName = "ColDescripcion"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 230
            End With
            
            'Columna Articulo (Modelo)
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Modelo"
                .DisableEditor = True
                .FieldName = "ARTICULO"
                .HeaderAlignment = taCenter
                .ObjectName = "ColModelo"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
            End With
            
            'Columna Color de Cuero
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Color"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "COLOR"
                .HeaderAlignment = taCenter
                .ObjectName = "ColColor"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
            End With
            
            'Columna Fecha de O/P
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Fecha"
                .DisableEditor = True
                .FieldName = "FECHA"
                .HeaderAlignment = taCenter
                .ObjectName = "ColFecha"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 60
            End With
            
            'Columna Cantidad Total
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 0
                .Caption = "Cantidad Total"
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "CANTIDADTOTAL"
                .HeaderAlignment = taCenter
                .ObjectName = "ColCantidadTotal"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
            End With
            
            'Columna Estado
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Estado"
                .DisableEditor = True
                .FieldName = "ESTADO"
                .HeaderAlignment = taCenter
                .ObjectName = "ColEstado"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
            End With
            
            'abrirCnDBMilano
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = strCadenaConexioBdStudioModa
            
            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctStatic
            .Dataset.ADODataset.LockType = ltReadOnly
            .Dataset.ADODataset.CommandText = SqlCad
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "IDORDENPRODUCCION"
            
            '.Bands.BandByName("bndPrincipal").Caption = "Ordenes de Producción Anuladas"
            '.Columns.HeaderFont.Bold = True
        End With
    End If
    
    SqlCad = vbNullString
    
    Exit Sub
errVisualizarOrdenProduccionAnulada:
    Select Case Err.Number
        Case 3704, 3709
            'cnn_dbbancos.Open StrConexDbBancos
            cnBdStudioModa.Open cnBdStudioModa.ConnectionString
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ModMilano: VisualizarOrdenProduccionAnulada"
    End Select
    
    Err.Clear
End Sub
