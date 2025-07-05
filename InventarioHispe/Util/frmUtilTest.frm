VERSION 5.00
Begin VB.Form frmUtilTest 
   Caption         =   "Form1"
   ClientHeight    =   7500
   ClientLeft      =   3180
   ClientTop       =   1860
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   10275
   Begin VB.TextBox txtCodAlmacen 
      Height          =   375
      Left            =   1680
      TabIndex        =   12
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton cmdGeneraCierre 
      Caption         =   "Generar Cierres de Mes"
      Height          =   615
      Left            =   2520
      TabIndex        =   10
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Frame fraDatos 
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
      TabIndex        =   3
      Top             =   120
      Width           =   6135
      Begin VB.TextBox txtNroPedido 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   960
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtCodProducto 
         Enabled         =   0   'False
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtProducto 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label lblNroPedido 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label4"
         Height          =   255
         Left            =   2280
         TabIndex        =   9
         Top             =   600
         Width           =   3735
      End
      Begin VB.Label Label2 
         Caption         =   "No. Pedido"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Producto"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Consultar Pedido"
      Height          =   975
      Left            =   6360
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.ListBox List2 
      Height          =   3180
      Left            =   6360
      TabIndex        =   1
      Top             =   1200
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   3180
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   6135
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo Almacen"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   4680
      Width           =   1455
   End
End
Attribute VB_Name = "frmUtilTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGeneraCierre_Click()
    Dim intAnno As Integer
    Dim intMes As Integer
    
    Dim a As Integer
    Dim m As Integer
    
    If ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NOMALM", "EF2ALMACENES", "F2CODALM", Trim(txtCodAlmacen.Text), "T") = vbNullString Then
        MsgBox "Almacen no existe, verifique.", vbInformation + vbOKOnly, App.ProductName
        
        Exit Sub
    End If
    
    intAnno = Year(CDate(ModUtilitario.sGetINI(App.Path & strNombreFicheroConfigSQLCliente, "ConfigServidorSQLCliente", "FechaCorteInicialDeValesParaCP", "l")))
    intMes = Month(CDate(ModUtilitario.sGetINI(App.Path & strNombreFicheroConfigSQLCliente, "ConfigServidorSQLCliente", "FechaCorteInicialDeValesParaCP", "l")))
    
    Me.MousePointer = vbHourglass
    
    For a = intAnno To Year(Date)
        
        For m = intMes To 12
            'Cierre Mes Stock Fisico
            SqlCad = vbNullString
            SqlCad = SqlCad & "INSERT INTO IF3CIERREMENSUAL(ANNO, MES, TIPO, ALMACEN, CODPRODUCTO, COMPROMISO, CANTIDAD) "
            
            SqlCad = SqlCad & "SELECT "
            SqlCad = SqlCad & a & " AS ANNO, "
            SqlCad = SqlCad & m & " AS MES, "
            SqlCad = SqlCad & "CIERREMES.TIPO, "
            SqlCad = SqlCad & "CIERREMES.ALMACEN, "
            SqlCad = SqlCad & "CIERREMES.CODPRODUCTO, "
            SqlCad = SqlCad & "CIERREMES.COMPROMISO, "
            SqlCad = SqlCad & "VAL(FORMAT(SUM(CIERREMES.CANTIDAD), '#0.00')) AS CANTIDAD "
            
            SqlCad = SqlCad & "FROM "
            
            SqlCad = SqlCad & "(SELECT "
            SqlCad = SqlCad & "TIPO, "
            SqlCad = SqlCad & "CM.ALMACEN, "
            SqlCad = SqlCad & "CM.CODPRODUCTO, "
            SqlCad = SqlCad & "TRIM(CM.COMPROMISO & '') AS COMPROMISO, "
            SqlCad = SqlCad & "VAL(FORMAT(CM.CANTIDAD, '#0.00')) AS CANTIDAD "
            SqlCad = SqlCad & "FROM "
            SqlCad = SqlCad & "IF3CIERREMENSUAL AS CM "
            SqlCad = SqlCad & "WHERE "
            SqlCad = SqlCad & "CM.ANNO = " & IIf(m = 1, a - 1, a) & " AND "
            SqlCad = SqlCad & "CM.MES = " & IIf(m = 1, 12, m - 1) & " AND "
            SqlCad = SqlCad & "CM.TIPO = 'F' AND "
            SqlCad = SqlCad & "CM.ALMACEN = '" & Trim(txtCodAlmacen.Text) & "' "
            
            SqlCad = SqlCad & "UNION ALL "
            
            SqlCad = SqlCad & "SELECT "
            SqlCad = SqlCad & "'F' AS TIPO, "
            SqlCad = SqlCad & "DET.F2CODALM AS ALMACEN, "
            SqlCad = SqlCad & "DET.F5CODPRO AS CODPRODUCTO, "
            SqlCad = SqlCad & "DET.COD_SOLICITUD AS COMPROMISO, "
            SqlCad = SqlCad & "VAL(FORMAT( SUM(DET.F3CANPRO * IIF(DET.TIPO = 'S', -1, 1)) , '#.00')) AS CANTIDAD "
            SqlCad = SqlCad & "FROM "
            SqlCad = SqlCad & "IF3VALES AS DET "
            SqlCad = SqlCad & "LEFT JOIN IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM "
            SqlCad = SqlCad & "WHERE "
            SqlCad = SqlCad & "CVDATE(CAB.F4FECVAL) BETWEEN "
                    SqlCad = SqlCad & "CVDATE('" & DateSerial(a, m + 0, 1) & "') AND "
                    SqlCad = SqlCad & "CVDATE('" & DateSerial(a, m + 1, 0) & "') AND "
            SqlCad = SqlCad & "CAB.F2CODALM = '" & Trim(txtCodAlmacen.Text) & "' "
            SqlCad = SqlCad & "GROUP BY "
            SqlCad = SqlCad & "DET.F2CODALM, "
            SqlCad = SqlCad & "DET.F5CODPRO, "
            SqlCad = SqlCad & "DET.COD_SOLICITUD) AS CIERREMES "
            SqlCad = SqlCad & "GROUP BY "
            SqlCad = SqlCad & "CIERREMES.TIPO, "
            SqlCad = SqlCad & "CIERREMES.ALMACEN, "
            SqlCad = SqlCad & "CIERREMES.CODPRODUCTO, "
            SqlCad = SqlCad & "CIERREMES.COMPROMISO "
            SqlCad = SqlCad & "HAVING "
            SqlCad = SqlCad & "VAL(FORMAT(SUM(CIERREMES.CANTIDAD), '#0.00')) <> 0"
            
            abrirCnnDbBancos
            
            cnn_dbbancos.Execute SqlCad
            
            
            'Cierre Mes Stock Virtual
            SqlCad = vbNullString
            SqlCad = SqlCad & "INSERT INTO IF3CIERREMENSUAL(ANNO, MES, TIPO, ALMACEN, CODPRODUCTO, COMPROMISO, CANTIDAD) "
            
            SqlCad = SqlCad & "SELECT "
            SqlCad = SqlCad & a & " AS ANNO, "
            SqlCad = SqlCad & m & " AS MES, "
            SqlCad = SqlCad & "CIERREMES.TIPO, "
            SqlCad = SqlCad & "CIERREMES.ALMACEN, "
            SqlCad = SqlCad & "CIERREMES.CODPRODUCTO, "
            SqlCad = SqlCad & "CIERREMES.COMPROMISO, "
            SqlCad = SqlCad & "VAL(FORMAT(SUM(CIERREMES.CANTIDAD), '#0.00')) AS CANTIDAD "
            SqlCad = SqlCad & "FROM "
            SqlCad = SqlCad & "(SELECT "
            SqlCad = SqlCad & "CM.TIPO, "
            SqlCad = SqlCad & "CM.ALMACEN, "
            SqlCad = SqlCad & "CM.CODPRODUCTO, "
            SqlCad = SqlCad & "TRIM(CM.COMPROMISO & '') AS COMPROMISO, "
            SqlCad = SqlCad & "VAL(FORMAT(CM.CANTIDAD, '#0.00')) AS CANTIDAD "
            SqlCad = SqlCad & "FROM "
            SqlCad = SqlCad & "IF3CIERREMENSUAL AS CM "
            SqlCad = SqlCad & "WHERE "
            SqlCad = SqlCad & "CM.ANNO = " & IIf(m = 1, a - 1, a) & " AND "
            SqlCad = SqlCad & "CM.MES = " & IIf(m = 1, 12, m - 1) & " AND "
            SqlCad = SqlCad & "CM.TIPO = 'V' "
            
            SqlCad = SqlCad & "UNION ALL "
            
            SqlCad = SqlCad & "SELECT "
            SqlCad = SqlCad & "'V' AS TIPO, "
            SqlCad = SqlCad & "DET.F4NUMORD AS ALMACEN, "
            SqlCad = SqlCad & "DET.F5CODPROORIGINAL AS CODPRODUCTO, "
            SqlCad = SqlCad & "TRIM(DET.COD_SOLICITUD & '') AS COMPROMISO, "
            SqlCad = SqlCad & "VAL(FORMAT( SUM(DET.F3CANPRO * IIF(DET.TIPO = 'S', -1, 1)) , '#.00')) AS CANTIDAD "
            SqlCad = SqlCad & "FROM "
            SqlCad = SqlCad & "IF3VALES AS DET "
            SqlCad = SqlCad & "LEFT JOIN IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM "
            SqlCad = SqlCad & "WHERE "
            SqlCad = SqlCad & "CVDATE(CAB.F4FECVAL) BETWEEN "
            SqlCad = SqlCad & "CVDATE('" & DateSerial(a, m + 0, 1) & "') AND "
            SqlCad = SqlCad & "CVDATE('" & DateSerial(a, m + 1, 0) & "') AND "
            SqlCad = SqlCad & "CAB.F1CODORI IN ('XC0') AND "
            SqlCad = SqlCad & "DET.F4NUMORD <> '' "
            SqlCad = SqlCad & "GROUP BY "
            SqlCad = SqlCad & "DET.F4NUMORD, "
            SqlCad = SqlCad & "DET.F5CODPROORIGINAL, "
            SqlCad = SqlCad & "DET.COD_SOLICITUD) AS CIERREMES "
            SqlCad = SqlCad & "GROUP BY "
            SqlCad = SqlCad & "CIERREMES.TIPO, "
            SqlCad = SqlCad & "CIERREMES.ALMACEN, "
            SqlCad = SqlCad & "CIERREMES.CODPRODUCTO, "
            SqlCad = SqlCad & "CIERREMES.COMPROMISO "
            SqlCad = SqlCad & "HAVING "
            SqlCad = SqlCad & "VAL(FORMAT(SUM(CIERREMES.CANTIDAD), '#0.00')) <> 0"
            
            abrirCnnDbBancos
            
            cnn_dbbancos.Execute SqlCad
            
            
            If a = Year(Date) And (m = Month(Date) - 1) Then
                Exit For
            End If
        Next m
            intMes = 1
    Next a
        MsgBox "Listo!", vbInformation + vbOKOnly, App.ProductName
    
    Me.MousePointer = vbDefault
End Sub

Private Sub Command1_Click()
    On Error GoTo errBoton
    
    Dim cmdOP As ADODB.Command
    
    Dim rstResumen As New ADODB.Recordset
    
    Set cmdOP = New ADODB.Command
    
    List1.Clear
    List2.Clear
    
    Me.MousePointer = vbHourglass
    
    List1.AddItem "Inicio: " & Now & " -------------------------------------------------"
    List2.AddItem "Inicio: " & Now & " --------------------"
    
    'abrirCnDBMilano
    
    With cmdOP
        .ActiveConnection = cnBdStudioModa
        .CommandType = adCmdStoredProc
        .CommandTimeout = "180"
        .CommandText = "usp_ConsultaSaldosOrdenesProduccionCPv3"
        
        .Parameters.Append .CreateParameter("@NROPEDIDO", adVarChar, adParamInput, 10, Trim(txtNroPedido.Text))
        .Parameters.Append .CreateParameter("@CODIGOPRODUCTO", adVarChar, adParamInput, 50, "")
        .Parameters.Append .CreateParameter("@FILTROPRODUCTO", adVarChar, adParamInput, 255, "")
        .Parameters.Append .CreateParameter("@FECHAENTREGACORTE", adVarChar, adParamInput, 20, "")
        .Parameters.Append .CreateParameter("@NOMBRETABLA", adVarChar, adParamInput, 255, "tmpResumenProduccionCP" & wusuario)
        .Parameters.Append .CreateParameter("@SOLOINSUMODESCARGADODEOP", adInteger, adParamInput, , 0)
        .Parameters.Append .CreateParameter("@CANTIDADFILAS", adInteger, adParamOutput, 5)
        
        .Execute
        
        dblCantidadRegistro = Val(.Parameters("@CANTIDADFILAS") & "")
        
        'Set rstResumen = .Execute()
    End With
    
    Set cmdOP = Nothing
    
    If rstResumen.State = 1 Then rstResumen.Close
    
    rstResumen.Open "SELECT * FROM tmpResumenProduccionCP" & wusuario, cnBdStudioModa, adOpenForwardOnly, adLockReadOnly
    
    If Not rstResumen.EOF Then
        List1.AddItem "Descarga: " & Now & " -------------------------------------------------"
        List2.AddItem "Descarga: " & Now & " --------------------"
        
        Do While Not rstResumen.EOF
            List1.AddItem Trim(rstResumen!LLAVE & "")
            List2.AddItem Trim(rstResumen!SALDO & "")
            
            rstResumen.MoveNext
        Loop
    End If
    
    List1.AddItem "Fin: " & Now & " ------------------ " & dblCantidadRegistro & " registro(s) ---------------------------"
    List2.AddItem "Fin: " & Now & " --------------------"
    
    Me.MousePointer = vbDefault
    
    Exit Sub
errBoton:
    MsgBox "Error: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    List1.AddItem "Fin: " & Now & " -------------------------------------------------"
    List2.AddItem "Fin: " & Now & " --------------------"
    
    Me.MousePointer = vbDefault
    
    Err.Clear
End Sub

Private Sub txtCodAlmacen_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            MsgBox "Almacen: " & ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NOMALM", "EF2ALMACENES", "F2CODALM", Trim(txtCodAlmacen.Text), "T"), vbInformation + vbOKOnly, App.ProductName
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
                End If
            End With
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub txtCodProducto_LostFocus()
    If Trim(txtCodProducto.Text) <> vbNullString Then
        txtCodProducto.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5CODPRO", "IF5PLA", "F5CODPRO", Trim(txtCodProducto.Text), "T")
        txtProducto.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5NOMPRO", "IF5PLA", "F5CODPRO", Trim(txtCodProducto.Text), "T")
        txtProducto.ToolTipText = txtProducto.Text
        
        If Trim(txtCodProducto.Text) = vbNullString Then
            txtProducto.Text = "Todos los Productos (*)"
            txtProducto.ToolTipText = vbNullString
        End If
    Else
        txtProducto.Text = "Todos los Productos (*)"
        txtProducto.ToolTipText = vbNullString
    End If
End Sub

Private Sub txtNroPedido_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'abrirCnDBMilano
            
            lblNroPedido.Caption = ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "(PER.NOMBRE + ' ( FEC. PEDIDO: ' + CONVERT(CHAR(10), PED.FECHAEMISION, 103) + ' / FEC. ENTREGA: ' + CONVERT(CHAR(10), PED.FECHAENTREGA, 103) + ')') AS RESUMEN", "PEDIDO AS PED LEFT JOIN PERSONA AS PER ON PER.IDPERSONA = PED.IDPERSONA", "PED.IDPEDIDO", Trim(txtNroPedido.Text), "T")
            lblNroPedido.ToolTipText = lblNroPedido.Caption
            
            If lblNroPedido.Caption = vbNullString Then
                MsgBox "No. de Pedido no encontrado o inválido.", vbInformation + vbOKOnly, App.ProductName
                
                txtNroPedido.SetFocus
            Else
                ModUtilitario.pulsarTecla vbKeyTab
            End If
    End Select
End Sub

Private Sub txtNroPedido_LostFocus()
    If lblNroPedido.Caption = vbNullString Then
        'abrirCnDBMilano
                
        lblNroPedido.Caption = ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "(PER.NOMBRE + ' ( FEC. PEDIDO: ' + CONVERT(CHAR(10), PED.FECHAEMISION, 103) + ' / FEC. ENTREGA: ' + CONVERT(CHAR(10), PED.FECHAENTREGA, 103) + ')') AS RESUMEN", "PEDIDO AS PED LEFT JOIN PERSONA AS PER ON PER.IDPERSONA = PED.IDPERSONA", "PED.IDPEDIDO", Trim(txtNroPedido.Text), "T")
        lblNroPedido.ToolTipText = lblNroPedido.Caption
        
        If lblNroPedido.Caption = vbNullString Then
            txtNroPedido.Text = vbNullString
        End If
    End If
End Sub
