VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRptCambioOrdenProduccion 
   Caption         =   "Reporte de Cambio(s )en Orden(es) de Producción"
   ClientHeight    =   8745
   ClientLeft      =   330
   ClientTop       =   2040
   ClientWidth     =   13650
   Icon            =   "frmRptCambioOrdenProduccion.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8745
   ScaleWidth      =   13650
   WindowState     =   2  'Maximized
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
         TabIndex        =   10
         Text            =   "Text1"
         ToolTipText     =   "Seleccione Producto (F2)"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtProducto 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3960
         TabIndex        =   9
         Text            =   "Text1"
         ToolTipText     =   "Ingrese cadena a buscar"
         Top             =   1320
         Width           =   4215
      End
      Begin VB.TextBox txtBusqueda 
         Height          =   285
         Left            =   4320
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   600
         Width           =   3855
      End
      Begin VB.ComboBox cmbUsuario 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   600
         Width           =   3015
      End
      Begin VB.ComboBox cmbAnno 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox cmbMes 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Producto"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2400
         TabIndex        =   11
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Filtrar"
         Height          =   255
         Left            =   3840
         TabIndex        =   8
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Usuario"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Periodo"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
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
      Tools           =   "frmRptCambioOrdenProduccion.frx":058A
      ToolBars        =   "frmRptCambioOrdenProduccion.frx":A9D1
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dbgReporte 
      Height          =   7440
      Left            =   120
      OleObjectBlob   =   "frmRptCambioOrdenProduccion.frx":AB02
      TabIndex        =   0
      Top             =   1200
      Width           =   13410
   End
End
Attribute VB_Name = "frmRptCambioOrdenProduccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private intOpcionResumen As Integer

Private Sub listarAnnos()
    Dim rstAnnoPA As New ADODB.Recordset
    
    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
        SqlCad = vbNullString
        SqlCad = SqlCad & "SELECT "
        SqlCad = SqlCad & "YEAR(FECHAMODIFICACION) AS ANNO "
        SqlCad = SqlCad & "FROM "
        SqlCad = SqlCad & "SF1ORDENPRODUCCION_LOG "
        SqlCad = SqlCad & "GROUP BY "
        SqlCad = SqlCad & "YEAR(FECHAMODIFICACION)"
        
        If rstAnnoPA.State = 1 Then rstAnnoPA.Close
        
        rstAnnoPA.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockOptimistic
    Else
        SqlCad = vbNullString
        SqlCad = SqlCad & "SELECT "
        SqlCad = SqlCad & "YEAR(FECHAMODIFICACION) AS ANNO "
        SqlCad = SqlCad & "FROM "
        SqlCad = SqlCad & "SF1ORDENPRODUCCION_LOG "
        SqlCad = SqlCad & "GROUP BY "
        SqlCad = SqlCad & "YEAR(FECHAMODIFICACION)"
        
        If rstAnnoPA.State = 1 Then rstAnnoPA.Close
        
        rstAnnoPA.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockOptimistic
    End If
    
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
    
    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
        SqlCad = vbNullString
        SqlCad = SqlCad & "SELECT "
        SqlCad = SqlCad & "MONTH(FECHAMODIFICACION) AS MES "
        SqlCad = SqlCad & "FROM "
        SqlCad = SqlCad & "SF1ORDENPRODUCCION_LOG "
        SqlCad = SqlCad & "WHERE "
        SqlCad = SqlCad & "YEAR(FECHAMODIFICACION) = '" & cmbAnno.Text & "' "
        SqlCad = SqlCad & "GROUP BY "
        SqlCad = SqlCad & "MONTH(FECHAMODIFICACION)"
        
        If rstMesPA.State = 1 Then rstMesPA.Close
        
        rstMesPA.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockOptimistic
    Else
        SqlCad = vbNullString
        SqlCad = SqlCad & "SELECT "
        SqlCad = SqlCad & "MONTH(FECHAMODIFICACION) AS MES "
        SqlCad = SqlCad & "FROM "
        SqlCad = SqlCad & "SF1ORDENPRODUCCION_LOG "
        SqlCad = SqlCad & "WHERE "
        SqlCad = SqlCad & "YEAR(FECHAMODIFICACION) = '" & cmbAnno.Text & "' "
        SqlCad = SqlCad & "GROUP BY "
        SqlCad = SqlCad & "MONTH(FECHAMODIFICACION)"
        
        If rstMesPA.State = 1 Then rstMesPA.Close
        
        rstMesPA.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockOptimistic
    End If
    
    cmbMes.Clear
    
    If Not rstMesPA.EOF Then
        cmbMes.AddItem "(*) - Todos los meses" & Space(100) & "00"
        
        Do While Not rstMesPA.EOF
            cmbMes.AddItem UCase(Format("01/" & Format(Trim(rstMesPA!mes & ""), "00") & "/" & cmbAnno.Text, "MMMM")) & Space(100) & Format(Trim(rstMesPA!mes & ""), "00")
            
            rstMesPA.MoveNext
        Loop
    End If
    
    If cmbMes.ListCount > 0 Then
        cmbMes.ListIndex = cmbMes.ListCount - 1
    End If
End Sub

Private Sub listarUsuarios()
    Dim rstUsuarioLOG As New ADODB.Recordset
    
    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
        SqlCad = vbNullString
        SqlCad = SqlCad & "SELECT "
        SqlCad = SqlCad & "U.F2CODUSER AS USUARIO "
        SqlCad = SqlCad & "FROM "
        SqlCad = SqlCad & "SF1ORDENPRODUCCION_LOG AS OP "
        SqlCad = SqlCad & "LEFT JOIN EF2USERS AS U ON U.F2CODUSEREXTERNO = OP.IDUSUARIO "
        SqlCad = SqlCad & "WHERE "
        SqlCad = SqlCad & "YEAR(OP.FECHAMODIFICACION) = " & cmbAnno.Text & " AND "
        SqlCad = SqlCad & "MONTH(OP.FECHAMODIFICACION) = " & Val(right(cmbMes.Text, 2)) & " "
        SqlCad = SqlCad & "GROUP BY "
        SqlCad = SqlCad & "U.F2CODUSER"
        
        If rstUsuarioLOG.State = 1 Then rstUsuarioLOG.Close
        
        rstUsuarioLOG.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockOptimistic
    Else
        SqlCad = vbNullString
        SqlCad = SqlCad & "SELECT "
        SqlCad = SqlCad & "U.F2CODUSER AS USUARIO "
        SqlCad = SqlCad & "FROM "
        SqlCad = SqlCad & "SF1ORDENPRODUCCION_LOG AS OP "
        SqlCad = SqlCad & "LEFT JOIN EF2USERS AS U ON U.F2CODUSEREXTERNO = OP.IDUSUARIO "
        SqlCad = SqlCad & "WHERE "
        SqlCad = SqlCad & "YEAR(OP.FECHAMODIFICACION) = " & cmbAnno.Text & " AND "
        SqlCad = SqlCad & "MONTH(OP.FECHAMODIFICACION) = " & Val(right(cmbMes.Text, 2)) & " "
        SqlCad = SqlCad & "GROUP BY "
        SqlCad = SqlCad & "U.F2CODUSER"
        
        If rstUsuarioLOG.State = 1 Then rstUsuarioLOG.Close
        
        rstUsuarioLOG.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockOptimistic
    End If
    
    cmbUsuario.Clear
    
    cmbUsuario.AddItem "(*) Todos"
    
    If Not rstUsuarioLOG.EOF Then
        Do While Not rstUsuarioLOG.EOF
            cmbUsuario.AddItem Trim(rstUsuarioLOG!Usuario & "")
            
            rstUsuarioLOG.MoveNext
        Loop
    End If
    
    If cmbUsuario.ListCount > 0 Then
        cmbUsuario.ListIndex = 0
    End If
End Sub

Private Sub limpiarCajas()
    txtBusqueda.Text = vbNullString
    
    txtCodProducto.Text = vbNullString
    txtProducto.Text = vbNullString
End Sub

Private Sub procesarConsulta()
    On Error GoTo errProcesarConsulta
    
    Screen.MousePointer = vbHourglass
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "(DET.IDORDENPRODUCCION & DET.IDCAMBIO) AS LLAVE, "
    SqlCad = SqlCad & "DET.IDORDENPRODUCCION, "
    SqlCad = SqlCad & "DET.CATEGORIA, "
    SqlCad = SqlCad & "DET.NUMEROOP, "
    SqlCad = SqlCad & "DET.IDCAMBIO, "
    SqlCad = SqlCad & "USU.F2CODUSER, "
    SqlCad = SqlCad & "DET.FECHAMODIFICACION, "
    
    SqlCad = SqlCad & "DET.IDINSUMO, "
    SqlCad = SqlCad & "PROD1.F5NOMPRO AS PRODUCTOORIGINAL, "
    SqlCad = SqlCad & "MED1.F7SIGMED AS UMORIGINAL, "
    SqlCad = SqlCad & "DET.IDINSUMOFINAL, "
    SqlCad = SqlCad & "PROD2.F5NOMPRO AS PRODUCTOFINAL, "
    SqlCad = SqlCad & "MED1.F7SIGMED AS UMFINAL, "
    
    SqlCad = SqlCad & "DET.CANTIDAD AS ORIGINAL, "
    SqlCad = SqlCad & "DET.CANTIDADFINAL AS FINAL, "
    
    SqlCad = SqlCad & "DET.OBSERVACION "
    
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "((((SF1ORDENPRODUCCION_LOG AS DET "
    SqlCad = SqlCad & "LEFT JOIN IF5PLA AS PROD1 ON PROD1.F5CODPRO = DET.IDINSUMO) "
    SqlCad = SqlCad & "LEFT JOIN IF5PLA AS PROD2 ON PROD2.F5CODPRO = DET.IDINSUMOFINAL) "
    SqlCad = SqlCad & "LEFT JOIN EF7MEDIDAS AS MED1 ON MED1.F7CODMED = PROD1.F7CODMED) "
    SqlCad = SqlCad & "LEFT JOIN EF7MEDIDAS AS MED2 ON MED2.F7CODMED = PROD2.F7CODMED) "
    SqlCad = SqlCad & "LEFT JOIN EF2USERS AS USU ON USU.F2CODUSEREXTERNO = DET.IDUSUARIO "
    
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "YEAR(DET.FECHAMODIFICACION) = " & cmbAnno.Text & " "
        
        If Val(right(cmbMes.Text, 2)) > 0 Then
            SqlCad = SqlCad & "AND MONTH(DET.FECHAMODIFICACION) = " & Val(right(cmbMes.Text, 2)) & " "
        End If
    
        If left(cmbUsuario.Text, 3) <> "(*)" Then
            SqlCad = SqlCad & "AND USU.F2CODUSER = '" & cmbUsuario.Text & "' "
        End If
        
        If txtBusqueda.Text <> vbNullString Then
            SqlCad = SqlCad & "AND ("
            'SqlCad = SqlCad & "DET.IDORDENPRODUCCION LIKE '%" & txtBusqueda.Text & "%' OR "
            SqlCad = SqlCad & "DET.CATEGORIA LIKE '%" & txtBusqueda.Text & "%' OR "
            SqlCad = SqlCad & "DET.NUMEROOP LIKE '%" & txtBusqueda.Text & "%' OR "
            SqlCad = SqlCad & "PROD1.F5NOMPRO LIKE '%" & txtBusqueda.Text & "%' OR "
            SqlCad = SqlCad & "PROD2.F5NOMPRO LIKE '%" & txtBusqueda.Text & "%'"
            SqlCad = SqlCad & ") "
        End If
        
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "DET.FECHAMODIFICACION, "
    SqlCad = SqlCad & "(DET.IDORDENPRODUCCION & DET.IDCAMBIO)"
    
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
            
            'Columna Id de Orden de Produccion
            Set gColumn = .Columns.Add(gedTextEdit)

            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Id O/P"
                .DisableEditor = True
                .FieldName = "IDORDENPRODUCCION"
                .Font.Charset = 0
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
                .DisableEditor = True
                .FieldName = "CATEGORIA"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCategoria"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
            End With
            
            'Columna Numero de O/P
            Set gColumn = .Columns.Add(gedTextEdit)

            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "O/P"
                .DisableEditor = True
                .FieldName = "NUMEROOP"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColNumeroOP"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 50
            End With
            
            'Columna Id. Cambio
            Set gColumn = .Columns.Add(gedTextEdit)

            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Id. Cambio"
                .DisableEditor = True
                .FieldName = "IDCAMBIO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColIdCambio"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 45
            End With
            
            'Columna Usuario
            Set gColumn = .Columns.Add(gedTextEdit)

            With gColumn
                .Alignment = taLeftJustify
                .BandIndex = 0
                .Caption = "Usuario"
                .DisableEditor = True
                .FieldName = "F2CODUSER"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColUsuario"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
            End With
            
            'Columna Fecha de Modificacion
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .BandIndex = 0
                .Caption = "Modificacion"
                .DisableEditor = True
                .FieldName = "FECHAMODIFICACION"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColModificacion"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
            End With
            
            'Columna Codigo de Producto Original
            Set gColumn = .Columns.Add(gedTextEdit)

            With gColumn
                .Alignment = taLeftJustify
                .BandIndex = 0
                .Caption = "Codigo"
                .DisableEditor = True
                .FieldName = "IDINSUMO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodProducto"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
                .Visible = False
            End With

            'Columna Descripcion de Producto Original
            Set gColumn = .Columns.Add(gedTextEdit)

            With gColumn
                .Alignment = taLeftJustify
                .BandIndex = 0
                .Caption = "Producto Original"
                .Color = vbWhite
                .DisableEditor = True
                .FieldName = "PRODUCTOORIGINAL"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColProductoOriginal"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 250
            End With

            'Columna Unidad de Medida Original
            Set gColumn = .Columns.Add(gedTextEdit)

            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "U.M. Orig."
                .Color = vbWhite
                .DisableEditor = True
                .FieldName = "UMORIGINAL"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColUMOriginal"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 60
            End With
            
            'Columna Codigo de Producto Final
            Set gColumn = .Columns.Add(gedTextEdit)

            With gColumn
                .Alignment = taLeftJustify
                .BandIndex = 0
                .Caption = "Codigo Final"
                .DisableEditor = True
                .FieldName = "IDINSUMOFINAL"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodProductoFinal"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
                .Visible = False
            End With

            'Columna Descripcion de Producto Final
            Set gColumn = .Columns.Add(gedTextEdit)

            With gColumn
                .Alignment = taLeftJustify
                .BandIndex = 0
                .Caption = "Producto Final"
                .Color = vbWhite
                .DisableEditor = True
                .FieldName = "PRODUCTOFINAL"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColProductoFinal"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 250
            End With

            'Columna Unidad de Medida Final
            Set gColumn = .Columns.Add(gedTextEdit)

            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "U.M. Final"
                .Color = vbWhite
                .DisableEditor = True
                .FieldName = "UMFINAL"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColUMFinal"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 60
            End With
            
            'Columna Cantidad Original
            Set gColumn = .Columns.Add(gedSpinEdit)

            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 5
                .Caption = "Cant. Original"
                .Color = &HFFFFC0
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "ORIGINAL"
                .Font.Charset = 0
                .Font.Bold = True
                .HeaderAlignment = taCenter
                .ObjectName = "ColCantidadOriginal"
                .SummaryFooterType = cstSum
                .Width = 70
            End With
            
            'Columna Cantidad Final
            Set gColumn = .Columns.Add(gedSpinEdit)

            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 5
                .Caption = "Cant. Final"
                .Color = &HFFFFC0
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "FINAL"
                .Font.Charset = 0
                .Font.Bold = True
                .HeaderAlignment = taCenter
                .ObjectName = "ColCantidadFinal"
                .SummaryFooterType = cstSum
                .Width = 70
            End With
            
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
            
            '.Columns.ColumnByFieldName("INFO").GroupIndex = 0
            
            .Columns.HeaderFont.Bold = True
            
            '.m.FullExpand
            
            .Columns.ColumnByFieldName("PRODUCTOORIGINAL").SummaryFooterType = cstCount
            .Columns.ColumnByFieldName("PRODUCTOORIGINAL").SummaryFooterFormat = "Cantidad de Registros = " & .Dataset.RecordCount
            
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



Private Sub dbgReporte_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    Select Case UCase(Column.FieldName)
        Case "PRODUCTOFINAL"
            If Trim(Text) <> Node.Values(5) Then
                Font.Bold = True
                FontColor = RGB(255, 255, 255)
                Color = RGB(217, 151, 149)
            Else
                Font.Bold = True
                FontColor = RGB(156, 101, 0)
                Color = RGB(255, 235, 156)
            End If
        Case "ORIGINAL"
            Text = Format(Text, "#,0.00")
        Case "FINAL"
            Text = Format(Text, "#,0.00")
            
            If Val(Text) <> Val(Node.Values(11)) Then
                Font.Bold = True
                FontColor = RGB(255, 255, 255)
                Color = RGB(217, 151, 149)
            Else
                Font.Bold = True
                FontColor = RGB(156, 101, 0)
                Color = RGB(255, 235, 156)
            End If
            
            Text = Format(Text, "#,0.00;(#,0.00)")
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

Private Sub Form_Load()
    limpiarCajas
    
    listarAnnos
    
    listarUsuarios
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    'fraPeriodo.Move 0, 0, Me.ScaleWidth, 1000
    
    dbgReporte.Move 0, fraPeriodo.Height + 200, Me.ScaleWidth, (Me.ScaleHeight - fraPeriodo.Height) - 200
End Sub

Private Sub tlbReporte_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.Id
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

Private Sub txtBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
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
