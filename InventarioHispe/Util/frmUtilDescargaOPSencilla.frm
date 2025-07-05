VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUtilDescargaOPSencilla 
   Caption         =   "Descarga de O/P"
   ClientHeight    =   8670
   ClientLeft      =   1230
   ClientTop       =   1680
   ClientWidth     =   16740
   Icon            =   "frmUtilDescargaOPSencilla.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8670
   ScaleWidth      =   16740
   Begin VB.Frame fraDatoVale 
      Caption         =   " Datos de Vale "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9120
      TabIndex        =   17
      Top             =   120
      Visible         =   0   'False
      Width           =   4095
      Begin VB.TextBox txtNumero 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   360
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   240
         Width           =   1845
      End
      Begin VB.Label lblNumeroValeExterno 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "< ID Externo >"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2280
         TabIndex        =   19
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nº"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   18
         Top             =   285
         Width           =   165
      End
   End
   Begin VB.Frame fraDetalleStock 
      Caption         =   " Datos de Descarga "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   5055
      Begin VB.ComboBox cmbConcepto 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "frmUtilDescargaOPSencilla.frx":058A
         Left            =   1080
         List            =   "frmUtilDescargaOPSencilla.frx":058C
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   3900
      End
      Begin VB.ComboBox cmbAlmacen 
         Height          =   315
         ItemData        =   "frmUtilDescargaOPSencilla.frx":058E
         Left            =   1080
         List            =   "frmUtilDescargaOPSencilla.frx":0598
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   3900
      End
      Begin MSComCtl2.DTPicker dtpFechaDespacho 
         Height          =   300
         Left            =   2160
         TabIndex        =   2
         Top             =   960
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   129302529
         CurrentDate     =   41978
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Concepto"
         Height          =   210
         Left            =   240
         TabIndex        =   20
         Top             =   600
         Width           =   690
      End
      Begin VB.Label Label1 
         Caption         =   "Almacen"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha de Despacho"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   2175
      End
   End
   Begin VB.Frame fraOrdenProduccion 
      Caption         =   " Datos de O/P "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5280
      TabIndex        =   10
      Top             =   120
      Width           =   3735
      Begin VB.TextBox txtIDOrdenProduccion 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   960
         Width           =   2535
      End
      Begin VB.ComboBox cmbCategoriaTipo 
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Text            =   "cmbCategoriaTipo"
         Top             =   240
         Width           =   3495
      End
      Begin VB.TextBox txtNroOrdenProduccion 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   4
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label lblIdCategoriaTipo 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ID CategoriaTipo"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1560
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "ID O.P."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   510
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Numero de O.P."
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   1140
      End
   End
   Begin VB.Frame fraProceso 
      Caption         =   " Procesando "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9120
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   4095
      Begin MSComctlLib.ProgressBar pgbProceso 
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   344
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dbgDescarga 
      Height          =   6690
      Left            =   120
      OleObjectBlob   =   "frmUtilDescargaOPSencilla.frx":05C7
      TabIndex        =   7
      Top             =   1560
      Width           =   16545
   End
   Begin ActiveToolBars.SSActiveToolBars tlbDescarga 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   12
      ShowShortcutsInToolTips=   -1  'True
      Tools           =   "frmUtilDescargaOPSencilla.frx":4207
      ToolBars        =   "frmUtilDescargaOPSencilla.frx":E6A1
   End
End
Attribute VB_Name = "frmUtilDescargaOPSencilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub configuraGrilla()
    With dbgDescarga.Options
        .Set (egoEditing)
        .Set (egoTabs)
        .Set (egoTabThrough)
        .Set (egoCanDelete)
        .Set (egoCanAppend)
        .Set (egoCanInsert)
        .Set (egoImmediateEditor)
        '.Set (egoShowIndicator)
        .Set (egoCanNavigation)
        .Set (egoHorzThrough)
        .Set (egoVertThrough)
        .Set (egoAutoWidth)
        .Set (egoEnterShowEditor)
        .Set (egoEnterThrough)
        .Set (egoShowButtonAlways)
        
        .Set (egoColumnSizing)
        .Set (egoColumnMoving)
        .Set (egoTabThrough)
        .Set (egoConfirmDelete)
        .Set (egoCanNavigation)
        .Set (egoCancelOnExit)
        .Set (egoLoadAllRecords)
        .Set (egoShowHourGlass)
        .Set (egoUseBookmarks)
        .Set (egoUseLocate)
        .Set (egoAutoCalcPreviewLines)
        .Set (egoBandSizing)
        .Set (egoBandMoving)
        .Set (egoDragScroll)
        .Set (egoAutoSort)
        .Set (egoExpandOnDblClick)
        .Set (egoShowFooter)
        .Set (egoShowGrid)
        .Set (egoShowButtons)
        .Set (egoNameCaseInsensitive)
        .Set (egoShowHeader)
        .Set (egoShowPreviewGrid)
        .Set (egoShowBorder)
        .Set (egoDynamicLoad)
        '.Set (egoRowSelect)
    End With
End Sub

Private Sub listarAlmacenEnCombo()
    Dim rstAlmacen As New ADODB.Recordset
    
    If rstAlmacen.State = 1 Then rstAlmacen.Close
    
    rstAlmacen.Open "SELECT F2CODALM, F2NOMALM FROM EF2ALMACENES ORDER BY F2CODALM", cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    cmbAlmacen.Clear
    
    If Not rstAlmacen.EOF Then
        rstAlmacen.MoveFirst
        
        Do While Not rstAlmacen.EOF
            cmbAlmacen.AddItem Trim(rstAlmacen!F2NOMALM & "") & Space(100) & "*" & Trim(rstAlmacen!f2codalm & "")
            
            rstAlmacen.MoveNext
        Loop
            If cmbAlmacen.ListCount > 0 Then
                cmbAlmacen.ListIndex = 0
            End If
    End If
    
    If rstAlmacen.State = 1 Then rstAlmacen.Close
    
    Set rstAlmacen = Nothing
End Sub

Private Sub listarConceptosAlmacen(ByVal strCodAlmacen As String)
    Dim rstConcepto As New ADODB.Recordset
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "ORI.F1CODORI, "
    SqlCad = SqlCad & "ORI.F1NOMORI "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "ALMACEN_CONCEPTO AS AC "
    SqlCad = SqlCad & "LEFT JOIN SF1ORIGENES AS ORI "
    SqlCad = SqlCad & "ON ORI.F1CODORI = AC.F1CODORI "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "AC.F2CODALM = '" & strCodAlmacen & "' AND "
    SqlCad = SqlCad & "ORI.F1TIPMOV = 'S' AND "
    SqlCad = SqlCad & "ORI.F1CODORI NOT IN ('XCS') "
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "ORI.F1NOMORI "
    
    If rstConcepto.State = 1 Then rstConcepto.Close
    
    rstConcepto.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    cmbConcepto.Clear
    
    If Not rstConcepto.EOF Then
        rstConcepto.MoveFirst
        
        Do While Not rstConcepto.EOF
            cmbConcepto.AddItem Trim(rstConcepto!F1NOMORI & "") & Space(100) & "*" & Trim(rstConcepto!F1CODORI & "")
            
            rstConcepto.MoveNext
        Loop
            If cmbConcepto.ListCount > 0 Then
                cmbConcepto.ListIndex = ModUtilitario.seleccionarItem(cmbConcepto, "XDP", "DER", Len("XDP"))
            End If
    End If
    
    If rstConcepto.State = 1 Then rstConcepto.Close
    
    Set rstConcepto = Nothing
End Sub

Private Sub limpiarCajasOP()
    cmbAlmacen.ListIndex = 0
    'cmbConcepto.ListIndex = -1
    cmbConcepto.Enabled = False
    
    dtpFechaDespacho.Value = Null
    
    cmbCategoriaTipo.ListIndex = -1
    txtNroOrdenProduccion.Text = vbNullString
    txtIDOrdenProduccion.Text = vbNullString
    lblIdCategoriaTipo.Caption = vbNullString
    
    txtNumero.Text = vbNullString
    lblNumeroValeExterno.Caption = vbNullString
End Sub
    
Private Sub validarCajasOP()
    If Trim(txtNumero.Text) <> vbNullString Then
        Exit Sub
    End If
    
    If cmbConcepto.ListIndex = -1 Then
        MsgBox "El almacen seleccionado, no cuenta con el Concepto de Despacho de OP, verifique.", vbInformation + vbOKOnly, App.ProductName
        
        cmbAlmacen.SetFocus
        
        Exit Sub
    End If
    
    
    If Trim(lblIdCategoriaTipo.Caption) = vbNullString Then
        MsgBox "Seleccione la Categoria de la O.P.", vbInformation + vbOKOnly, App.ProductName
        
        cmbCategoriaTipo.SetFocus
        
        Exit Sub
    End If
    
    If Trim(txtNroOrdenProduccion.Text) = vbNullString Then
        MsgBox "Ingrese el Número de Orden de Producción.", vbInformation + vbOKOnly, App.ProductName
        
        Exit Sub
    End If
    
    consultarOP
End Sub

Private Sub consultarOP()
    'ModMilano.abrirCnDBMilano
    
    txtIDOrdenProduccion.Text = ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "IDORDENPRODUCCION", "ORDENPRODUCCION", "IDCATEGORIATIPO", Trim(lblIdCategoriaTipo.Caption), "T", "AND OP = '" & Trim(txtNroOrdenProduccion.Text) & "' AND ANULADO = 0")
    
    If Trim(txtIDOrdenProduccion.Text) = vbNullString Then
        MsgBox "O.P. no existe o esta anulada.", vbInformation + vbOKOnly, App.ProductName
    Else
        If ModMilano.importarOPServidorExternoV3(Trim(txtIDOrdenProduccion.Text), fraProceso, pgbProceso) Then
            dbgDescarga.Dataset.Close
            
            cmbConcepto.ListIndex = ModUtilitario.seleccionarItem(cmbConcepto, "XDP", "DER", 3)
        Else
            txtIDOrdenProduccion.Text = vbNullString
            
            txtIDOrdenProduccion.SetFocus
        End If
        
        ModUtilitario.pulsarTecla vbKeyTab
    End If
    
    descargarAtencionPedido
    
    listarOP
End Sub

Private Sub descargarAtencionPedido()
    On Error GoTo errDescargarAtencionPedido
    
    Dim rstAtencionPedido As New ADODB.Recordset
    
    dbgDescarga.Dataset.Close
    
    abrirCnTemporal
        
    cnDBTemp.Execute "DELETE FROM TMPUTILRESUMENSTOCKREQUERIMIENTOOP"
    
    If rstAtencionPedido.State = 1 Then rstAtencionPedido.Close
    
    rstAtencionPedido.Open "SELECT NROPEDIDO, CODPRODUCTOFINAL FROM TMPUTILDESCARGAOPSENCILLA GROUP BY NROPEDIDO, CODPRODUCTOFINAL", cnDBTemp, adOpenForwardOnly, adLockReadOnly
    
    If Not rstAtencionPedido.EOF Then
        DoEvents
        
        fraProceso.Visible = True
        pgbProceso.Value = 0
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstAtencionPedido)
        fraProceso.Caption = "Descargando Stock..."
        
        Do While Not rstAtencionPedido.EOF
            With objAyudaVale
                .inicializarEntidades
                .inicializarEntidadesDetalle
                
                .CodigoProducto = Trim(rstAtencionPedido!CODPRODUCTOFINAL & "")
                .CodigoAlmacen = Mid(Trim(cmbAlmacen.Text), InStr(1, Trim(cmbAlmacen.Text), "*") + 1)
                
                If .devuelveStockFisicoDeProductoV2(vbNullString, True) > 0 Then
                    .descargarStockProductoReqOP "CEA", Trim(rstAtencionPedido!NroPedido & "")
                End If
                
                .descargarStockProductoReqOP "CPL", Trim(rstAtencionPedido!NroPedido & "")
            End With
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Descargando Stock... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 0)
            
            rstAtencionPedido.MoveNext
        Loop
        
        rstAtencionPedido.MoveFirst
        
        DoEvents
        
        fraProceso.Visible = True
        pgbProceso.Value = 0
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstAtencionPedido)
        fraProceso.Caption = "Actualizando Stock..."
        
        Do While Not rstAtencionPedido.EOF
            SqlCad = vbNullString
            SqlCad = SqlCad & "UPDATE "
            SqlCad = SqlCad & "TMPUTILDESCARGAOPSENCILLA "
            SqlCad = SqlCad & "SET "
            SqlCad = SqlCad & "STOCKCOMPROMETIDO = " & Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "CANTIDAD", "TMPUTILRESUMENSTOCKREQUERIMIENTOOP", "TIPO", "CEA", "T", "AND CODPRODUCTO = '" & Trim(rstAtencionPedido!CODPRODUCTOFINAL & "") & "'")) & ", "
            SqlCad = SqlCad & "STOCKPORLLEGAR = " & Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "CANTIDAD", "TMPUTILRESUMENSTOCKREQUERIMIENTOOP", "TIPO", "CPL", "T", "AND CODPRODUCTO = '" & Trim(rstAtencionPedido!CODPRODUCTOFINAL & "") & "'")) & " "
            SqlCad = SqlCad & "WHERE "
            SqlCad = SqlCad & "NROPEDIDO = '" & Trim(rstAtencionPedido!NroPedido & "") & "' AND "
            SqlCad = SqlCad & "CODPRODUCTOFINAL = '" & Trim(rstAtencionPedido!CODPRODUCTOFINAL & "") & "'"
            
            cnDBTemp.Execute SqlCad
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Actualizando Stock... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 0)
            
            rstAtencionPedido.MoveNext
        Loop
    End If
    
    fraProceso.Visible = False
    
    Exit Sub
errDescargarAtencionPedido:
    MsgBox "Nro.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub


Private Sub listarOP()
    On Error GoTo errListarOP
    
    With dbgDescarga
        .Dataset.Close
                
        .Columns.DestroyColumns
    End With
    
    Dim gColumn As dxGridColumn
    
    With dbgDescarga
        'Columna ID de Orden
        Set gColumn = .Columns.Add(gedTextEdit)
        
        With gColumn
            .Alignment = taCenter
            .BandIndex = 0
            .Caption = "O.P."
            .DisableEditor = True
            .FieldName = "NROOP"
            .HeaderAlignment = taCenter
            .ObjectName = "ColNroOP"
            .SummaryFooterType = cstCount
            .SummaryFooterFormat = " "
            .Width = 60
            .Visible = False
        End With
        
        'Columna Numero de Pedido
        Set gColumn = .Columns.Add(gedTextEdit)
        
        With gColumn
            .Alignment = taCenter
            .BandIndex = 0
            .Caption = "No. Pedido"
            .DisableEditor = True
            .FieldName = "NROPEDIDO"
            .HeaderAlignment = taCenter
            .ObjectName = "ColNroPedido"
            .SummaryFooterType = cstCount
            .SummaryFooterFormat = " "
            .Width = 60
            .Visible = False
        End With
        
        'Columna Datos Resumen (Llave)
        Set gColumn = .Columns.Add(gedTextEdit)
        
        With gColumn
            .Alignment = taLeftJustify
            .Caption = "Datos de O.P"
            .BandIndex = 0
            .DisableEditor = True
            .FieldName = "LLAVEOP"
            .Font.Name = "Arial"
            .Font.Charset = 0
            .HeaderAlignment = taCenter
            .ObjectName = "ColLlaveOP"
            .SummaryFooterType = cstCount
            .SummaryFooterFormat = " "
            .Width = 50
        End With
        
        'Columna Codigo de Producto de Origen
        Set gColumn = .Columns.Add(gedTextEdit)
        
        With gColumn
            .Alignment = taLeftJustify
            .Caption = "Codigo Origen"
            .BandIndex = 0
            .DisableEditor = True
            .FieldName = "CODPRODUCTOORIGEN"
            .Font.Name = "Arial"
            .Font.Charset = 0
            .HeaderAlignment = taCenter
            .ObjectName = "ColCodProductoOrigen"
            .SummaryFooterType = cstCount
            .SummaryFooterFormat = " "
            .Width = 80
        End With
        
        'Columna Codigo de Producto Final
        Set gColumn = .Columns.Add(gedButtonEdit)
        
        With gColumn
            .Alignment = taLeftJustify
            .Caption = "Codigo Final"
            .BandIndex = 0
            '.DisableEditor = True
            .FieldName = "CODPRODUCTOFINAL"
            .Font.Name = "Arial"
            .Font.Charset = 0
            .ButtonColumn.EditButtonStyle = ebsEllipsis
            .HeaderAlignment = taCenter
            .ObjectName = "ColCodProductoFinal"
            .SummaryFooterType = cstCount
            .SummaryFooterFormat = " "
            .Width = 80
        End With
        
        'Columna Descripcion del Producto
        Set gColumn = .Columns.Add(gedTextEdit)
        
        With gColumn
            .Alignment = taLeftJustify
            .Caption = "Descripción"
            .BandIndex = 0
            .DisableEditor = True
            .FieldName = "NOMPRODUCTO"
            .Font.Name = "Arial"
            .Font.Charset = 0
            .HeaderAlignment = taCenter
            .ObjectName = "ColDescripcion"
            .SummaryFooterType = cstCount
            .SummaryFooterFormat = " "
            .Width = 300
        End With
        
        'Columna Unidad de Medida
        Set gColumn = .Columns.Add(gedTextEdit)
        
        With gColumn
            .Alignment = taCenter
            .Caption = "U.M."
            .BandIndex = 0
            .DisableEditor = True
            .FieldName = "UM"
            .Font.Name = "Arial"
            .Font.Charset = 0
            .HeaderAlignment = taCenter
            .ObjectName = "ColUM"
            .SummaryFooterType = cstCount
            .SummaryFooterFormat = " "
            .Width = 70
        End With
        
        'Columna Cantidad de Origen
        Set gColumn = .Columns.Add(gedSpinEdit)
        
        With gColumn
            .Alignment = taRightJustify
            .BandIndex = 0
            .Caption = "Cantidad Original"
            '.Color = &HC0&
            .DecimalPlaces = 2
            .DisableEditor = True
            .FieldName = "CANTIDADORIGEN"
            .HeaderAlignment = taCenter
            .ObjectName = "ColCantidad"
            .SummaryFooterType = cstCount
            .SummaryFooterFormat = " "
            .Width = 70
        End With
        
        'Columna Cantidad Final
        Set gColumn = .Columns.Add(gedSpinEdit)
        
        With gColumn
            .Alignment = taRightJustify
            .BandIndex = 0
            .Caption = "Cantidad Final"
            '.Color = &HC0&
            .DecimalPlaces = 2
            .FieldName = "CANTIDADFINAL"
            .HeaderAlignment = taCenter
            .ObjectName = "ColCantidad"
            .SummaryFooterType = cstCount
            .SummaryFooterFormat = " "
            .Width = 70
        End With
        
        'Columna Saldo
        Set gColumn = .Columns.Add(gedSpinEdit)
        
        With gColumn
            .Alignment = taRightJustify
            .BandIndex = 0
            .Caption = "Saldo"
            '.Color = &HC0&
            .DecimalPlaces = 2
            .DisableEditor = True
            .FieldName = "SALDO"
            .HeaderAlignment = taCenter
            .ObjectName = "ColCantidad"
            .SummaryFooterType = cstCount
            .SummaryFooterFormat = " "
            .Width = 70
        End With
        
        'Columna Stock Comprometido
        Set gColumn = .Columns.Add(gedSpinEdit)
        
        With gColumn
            .Alignment = taRightJustify
            .BandIndex = 0
            .Caption = "Stock C.E.A."
            .Color = &HC0&
            .DecimalPlaces = 2
            .DisableEditor = True
            .FieldName = "STOCKCOMPROMETIDO"
            .Font.Bold = True
            .FontColor = &HFFFFFF
            .Font.Charset = 0
            .HeaderAlignment = taCenter
            .ObjectName = "ColStockComprometido"
            .SummaryFooterType = cstCount
            .SummaryFooterFormat = " "
            .Width = 70
        End With
        
        'Columna Stock Por Llegar
        Set gColumn = .Columns.Add(gedSpinEdit)
        
        With gColumn
            .Alignment = taRightJustify
            .BandIndex = 0
            .Caption = "Stock C.P.L."
            .Color = &H80FFFF
            .DecimalPlaces = 2
            .DisableEditor = True
            .FieldName = "STOCKPORLLEGAR"
            .Font.Bold = False
            .FontColor = &H80000012
            .Font.Charset = 0
            .HeaderAlignment = taCenter
            .ObjectName = "ColStockPorLlegar"
            .SummaryFooterType = cstCount
            .SummaryFooterFormat = " "
            .Width = 70
        End With
        
        'Columna Cantidad Descarga
        Set gColumn = .Columns.Add(gedSpinEdit)
        
        With gColumn
            .Alignment = taRightJustify
            .BandIndex = 0
            .Caption = "Descargar"
            .Color = &HFFFFC0
            .DecimalPlaces = 2
            .DisableEditor = True
            .FieldName = "CANTIDADDESCARGA"
            .HeaderAlignment = taCenter
            .ObjectName = "ColCantidad"
            .SummaryFooterType = cstCount
            .SummaryFooterFormat = " "
            .Width = 70
        End With
        
        'Columna Procesar
        Set gColumn = .Columns.Add(gedCheckEdit)
        
        With gColumn
            .Alignment = taCenter
            .Caption = "OK"
            .FieldName = "PROCESAR"
            .Font.Charset = 0
            .HeaderAlignment = taCenter
            .ObjectName = "ColProcesar"
            .SummaryFooterType = cstCount
            .SummaryFooterFormat = " "
            .Width = 70
        End With
        
        abrirCnTemporal
        
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = cnDBTemp.ConnectionString
        
        .Dataset.Active = False
        .Dataset.ADODataset.CommandType = cmdText
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.ADODataset.CursorType = ctKeyset
        .Dataset.ADODataset.LockType = ltOptimistic
        .Dataset.ADODataset.CommandText = "SELECT * FROM TMPUTILDESCARGAOPSENCILLA WHERE SALDO > 0 ORDER BY NOMPRODUCTO"
        .Dataset.Active = True
        .Dataset.Refresh
        .KeyField = "CODPRODUCTOORIGEN"
        
        .Columns.ColumnByFieldName("LLAVEOP").GroupIndex = 0
        
        .m.FullExpand
        
        '.Columns.ColumnByFieldName("CANTIDAD").SummaryFooterType = cstSum
    End With
    
    Exit Sub
errListarOP:
    Select Case Err.Number
        Case 3704, 3709
            'cnn_dbbancos.Open StrConexDbBancos
            abrirCnTemporal
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsOrdenTrabajo: ListarGrillaInsumoPendienteDescarga"
    End Select
    
    Err.Clear
End Sub

Private Sub cmbAlmacen_Click()
    listarConceptosAlmacen Mid(cmbAlmacen.Text, InStr(1, cmbAlmacen.Text, "*") + 1)
End Sub

Private Sub cmbalmacen_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub cmbCategoriaTipo_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub cmbCategoriaTipo_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim LenText As Long, ret As Long
    
    'Si los caracteres presionados están entre el 0 y la Z
    If KeyCode >= vbKey0 And KeyCode <= vbKeyZ Then
        ret = ModUtilitario.SendMessage(cmbCategoriaTipo.HWnd, &H14C&, -1, ByVal cmbCategoriaTipo.Text)
        
        If ret >= 0 Then
            LenText = Len(cmbCategoriaTipo.Text)
            
            cmbCategoriaTipo.ListIndex = ret
            cmbCategoriaTipo.Text = cmbCategoriaTipo.List(ret)
            cmbCategoriaTipo.SelStart = LenText
            cmbCategoriaTipo.SelLength = Len(cmbCategoriaTipo.Text) - LenText
        End If
    End If
End Sub

Private Sub cmbCategoriaTipo_LostFocus()
    If Trim(cmbCategoriaTipo.Text) <> vbNullString Then
        'ModMilano.abrirCnDBMilano
        
        lblIdCategoriaTipo.Caption = ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "IDCATEGORIATIPO", "CATEGORIATIPO", "NOMBRE", Trim(cmbCategoriaTipo.Text), "T")
        
        If Trim(lblIdCategoriaTipo.Caption) = vbNullString Then
            MsgBox "Categoria no identificada.", vbInformation + vbOKOnly, App.ProductName
            
            cmbCategoriaTipo.SetFocus
        End If
    End If
End Sub

Private Sub dbgDescarga_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    Select Case Column.FieldName
        Case "CODPRODUCTOFINAL"
            If Trim(Text) <> Node.Values(3) Then
                Font.Bold = True
                FontColor = RGB(255, 255, 255)
                Color = RGB(217, 151, 149)
            Else
                Font.Bold = True
                FontColor = RGB(156, 101, 0)
                Color = RGB(255, 235, 156)
            End If
        Case "CANTIDADFINAL"
            If Val(Text) <> Val(Node.Values(7)) Then
                Font.Bold = True
                FontColor = RGB(255, 255, 255)
                Color = RGB(217, 151, 149)
            Else
                Font.Bold = True
                FontColor = RGB(156, 101, 0)
                Color = RGB(255, 235, 156)
            End If
            
            Text = Format(Text, "#,0.00;(#,0.00)")
        Case "CANTIDADORIGEN", "SALDO"
            If Val(Text) < 0 Then
                FontColor = vbRed
            ElseIf Val(Text) = 0 Then
                FontColor = vbGreen
            Else
                FontColor = vbBlue
            End If
            
            Text = Format(Text, "#,0.00;(#,0.00)")
        Case "STOCKCOMPROMETIDO", "STOCKPORLLEGAR", "CANTIDADDESCARGA"
            Text = Format(Text, "#,0.00;(#,0.00)")
    End Select
End Sub


Private Sub dbgDescarga_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    If Not IsDate(dtpFechaDespacho.Value) Then
        MsgBox "Ingrese la Fecha de Despacho.", vbInformation + vbOKOnly, App.ProductName
        
        dtpFechaDespacho.SetFocus
        
        Exit Sub
    End If
    
    Select Case Column.FieldName
        Case "CODPRODUCTOFINAL"
            If Val(dbgDescarga.Columns.ColumnByFieldName("CANTIDADFINAL").Value & "") <> Val(dbgDescarga.Columns.ColumnByFieldName("SALDO").Value & "") Then
                MsgBox "Imposible reemplazar el Producto, ya fue descargado.", vbInformation + vbOKOnly, App.ProductName
                
                Exit Sub
            End If
            
            If Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "CANTIDAD", "TMPUTILSTOCKDETALLE", "CODPRODUCTO", _
                                                Trim(dbgDescarga.Columns.ColumnByFieldName("CODPRODUCTOFINAL").Value & ""), "T", _
                                                "AND CODALMACEN = '" & right(cmbAlmacen.Text, 2) & "' AND NROPEDIDO = '" & _
                                                Trim(dbgDescarga.Columns.ColumnByFieldName("NROPEDIDO").Value & "") & "'")) > 0 Then
                                                
                If MsgBox("El Producto cuenta actualmente con Stock Comprometido Disponible, ¿Desea continuar con el cambio?" & vbNewLine & vbNewLine & _
                            "RECOMENDACIÓN: Asegurese de liberar el Stock Comprometido del Producto, antes de proceder con el Cambio.", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
                    
                    Exit Sub
                End If
            End If
            
'            If ModUtilitario.validarFormAbierto("ayuda_productos") Then
'                Unload ayuda_productos
'            End If
'
'            With ayuda_productos
'                .CodigoAuxiliar = vbNullString
'                .CodigoRequerimiento = vbNullString
'                .CodigoProducto = vbNullString
'
'                '.txtBusqueda.Text = Trim(dbgDescarga.Columns.ColumnByFieldName("NOMPRODUCTO").value & "")
'                .CadenaCorte = InputBox("Ingrese cadena de texto a buscar:", App.ProductName, vbNullString)
'
'                .listarProductos
'
'                .Show 1
'            End With
'
'            abrirCnTemporal
'
'            objAyudaBien.Codigo = ModUtilitario.ObtenerCampoV2(cnDBTemp, "F5CODPRO", "TMPPRODUCTOS", "F4PERINT", "-1", "N")
            
            If ModUtilitario.validarFormAbierto("frmListaBien") Then
                Unload frmListaBien
            End If
            
            With frmListaBien
                objAyudaBien.inicializarEntidades
                
                objAyudaBien.Codigo = Trim(dbgDescarga.Columns.ColumnByFieldName("CODPRODUCTOFINAL").Value & "")
                
                objAyudaBien.obtenerConfigBien
                
                '.Ayuda = True
                '.TieneMovimientoAlmacen = True
                '.InsumoOP = True
                '.CadenaCorte = objAyudaBien.Modelo
                
                .Ayuda = True
                .InsumoOP = True
                .ParaVenta = False
                .TieneMovimientoAlmacen = True
                .CadenaCorte = objAyudaBien.Modelo
                .FiltroAdicional = vbNullString
                .TipoBienMostrar = "P"
                
                objAyudaBien.inicializarEntidades
                
                .Show 1
                
                If objAyudaBien.Codigo <> vbNullString Then
                    objAyudaBien.obtenerConfigBien
                    
                    If ModUtilitario.ObtenerCampoV2(cnDBTemp, "CODPRODUCTOFINAL", "TMPUTILDESCARGAOPSENCILLA", "CODPRODUCTOFINAL", objAyudaBien.Codigo, "T") <> vbNullString Then
                        MsgBox "Imposible realizar el cambio; el producto:" & vbNewLine & _
                                objAyudaBien.Descripcion & ", " & vbNewLine & _
                                "Se encuentra consignado en la actual OP. Se sugiere:" & vbNewLine & _
                                "1) Sumar la Cantidad del producto a cambiar, al que ya existe." & vbNewLine & _
                                "2) En caso que el producto a cambiar cuente con Stock Comprometido, proceder a Liberarlo." & vbNewLine & _
                                "3) Anular (Desestimar) el producto a cambiar, llevando su Cantidad Final a cero (0)." & vbNewLine & _
                                "4) Proceder con la Descarga del Producto ya existente, que cuenta con la suma de ambos.", vbInformation + vbOKOnly, App.ProductName
                        
                        objAyudaBien.inicializarEntidades
                        
                        Exit Sub
                    End If
                    
                    If Trim(dbgDescarga.Columns.ColumnByFieldName("UM").Value & "") <> Trim(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F7SIGMED", "EF7MEDIDAS", "F7CODMED", objAyudaBien.CodUM, "T")) Then
                        If MsgBox("El Producto seleccionado para el cambio cuenta con diferente Unidad de Medida (U.M.)." & vbNewLine & _
                                    "¿Desea continuar con la acción?", vbInformation + vbYesNo, App.ProductName) = vbNo Then
                            
                            objAyudaBien.inicializarEntidades
                            
                            Exit Sub
                        End If
                    End If
                    
                    If MsgBox("¿Desea efectuar el reemplazo del Producto?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
                        Exit Sub
                    End If
                    
                    With dbgDescarga
                        If ModMilano.modificarProductoEnOP(Trim(.Columns.ColumnByFieldName("NROOP").Value), _
                                                            Trim(.Columns.ColumnByFieldName("CODPRODUCTOFINAL").Value), _
                                                            objAyudaBien.Codigo, _
                                                            Val(.Columns.ColumnByFieldName("CANTIDADORIGEN").Value), _
                                                            Val(.Columns.ColumnByFieldName("CANTIDADFINAL").Value), "DESCARGA DE OP - CAMBIO DE PRODUCTO") Then
                               
                            MsgBox "Producto reemplazado en OP: " & Trim(.Columns.ColumnByFieldName("NROOP").Value), vbInformation + vbOKOnly, App.ProductName
                            
                            .Dataset.Edit
                            
                            .Columns.ColumnByFieldName("CODPRODUCTOFINAL").Value = objAyudaBien.Codigo
                            .Columns.ColumnByFieldName("NOMPRODUCTO").Value = objAyudaBien.Descripcion
                            .Columns.ColumnByFieldName("UM").Value = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F7SIGMED", "EF7MEDIDAS", "F7CODMED", objAyudaBien.CodUM, "T")
                            
                            .Dataset.Post
                        End If
                        
                        'dbgDescarga_OnChangeNodeEx
                    End With
                End If
            End With
    End Select
End Sub


Private Sub dbgDescarga_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    If Not IsDate(dtpFechaDespacho.Value) Then
        MsgBox "Ingrese la Fecha de Despacho.", vbInformation + vbOKOnly, App.ProductName
        
        dtpFechaDespacho.SetFocus
        
        Exit Sub
    End If
    
    Select Case dbgDescarga.Columns.FocusedColumn.FieldName
        Case "CANTIDADFINAL"
            With dbgDescarga
                If .Dataset.State = dsEdit Then
                    Dim dblCantidadOrigen As Double
                    
                    dblCantidadOrigen = Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "CANTIDADFINAL", "TMPUTILDESCARGAOPSENCILLA", "CODPRODUCTOFINAL", Trim(.Columns.ColumnByFieldName("CODPRODUCTOFINAL").Value & ""), "T"))
                    
                    If Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "CANTIDAD", "TMPUTILSTOCKDETALLE", "CODPRODUCTO", _
                                                        Trim(dbgDescarga.Columns.ColumnByFieldName("CODPRODUCTOFINAL").Value & ""), "T", _
                                                        "AND CODALMACEN = '" & right(cmbAlmacen.Text, 2) & "' AND NROPEDIDO = '" & _
                                                        Trim(dbgDescarga.Columns.ColumnByFieldName("NROPEDIDO").Value & "") & "'")) > 0 Then
                                                        
                        If MsgBox("El Producto cuenta actualmente con Stock Comprometido Disponible, ¿Desea continuar con el cambio?" & vbNewLine & vbNewLine & _
                                    "RECOMENDACIÓN: Asegurese de liberar el Stock Comprometido del Producto, antes de proceder con el Cambio.", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
                            
                            .Dataset.Cancel
                            
                            Exit Sub
                        End If
                    End If
                    
                    'If Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "CANTIDADFINAL", "TMPUTILDESCARGAOPSENCILLA", "CODPRODUCTOFINAL", Trim(.Columns.ColumnByFieldName("CODPRODUCTOFINAL").value & ""), "T"))  Then
                    If MsgBox("¿Desea aplicar el Ajuste de Cantidad?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
                        .Dataset.Cancel
                        
                        Exit Sub
                    Else
                        .Dataset.Post
                    End If
                    
                    If ModMilano.modificarProductoEnOP(Trim(.Columns.ColumnByFieldName("NROOP").Value), _
                                                        Trim(.Columns.ColumnByFieldName("CODPRODUCTOFINAL").Value), _
                                                        Trim(.Columns.ColumnByFieldName("CODPRODUCTOFINAL").Value), _
                                                        Val(.Columns.ColumnByFieldName("SALDO").Value), _
                                                        Val(.Columns.ColumnByFieldName("CANTIDADFINAL").Value), "DESCARGA DE OP - AJUSTE DE CANTIDAD DE PRODUCTO") Then
                        
                        MsgBox "Efectuado ajuste de cantidad de Producto en OP: " & Trim(.Columns.ColumnByFieldName("NROOP").Value), vbInformation + vbOKOnly, App.ProductName
                        
                        .Dataset.Edit
                        
                        .Columns.ColumnByFieldName("SALDO").Value = Val(.Columns.ColumnByFieldName("CANTIDADFINAL").Value & "") - (dblCantidadOrigen - Val(.Columns.ColumnByFieldName("SALDO").Value & ""))
                        
                        .Dataset.Post
                    Else
                        .Dataset.Edit
                        
                        .Columns.ColumnByFieldName("CANTIDADFINAL").Value = Val(.Columns.ColumnByFieldName("SALDO").Value & "")
                        
                        .Dataset.Post
                    End If
                    
                    dblCantidadOrigen = 0
                    
                    'dbgDescarga_OnChangeNodeEx
                End If
            End With
    End Select
End Sub


Private Sub dtpFechaDespacho_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub Form_Load()
    Me.top = (Screen.Height / 2) - (Me.Height / 2)
    Me.left = (Screen.Width / 2) - (Me.Width / 2)
    
    abrirCnTemporal
    
    cnDBTemp.Execute "DELETE FROM TMPUTILDESCARGAOPSENCILLA"
    
    limpiarCajasOP
    
    listarAlmacenEnCombo
    
    'ModMilano.listarCategoriaTipo cmbCategoriaTipo
    
    listarOP
    
    ModUtilitario.deshabilitarBotonCerrarForm frmUtilDescargaOPSencilla
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    dbgDescarga.Dataset.Close
End Sub

Private Sub tlbDescarga_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.ID
        Case "ID_Nuevo":
            If MsgBox("¿Desea realizar una nueva descarga?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
                Exit Sub
            End If
            
            Me.MousePointer = vbHourglass
            
            dbgDescarga.Dataset.Close
            
            abrirCnTemporal
            
            cnDBTemp.Execute "DELETE FROM TMPUTILDESCARGAOPSENCILLA"
            
            limpiarCajasOP
            
            listarOP
            
            cmbAlmacen.SetFocus
            
            Me.MousePointer = vbDefault
        Case "ID_Grabar":
            Me.MousePointer = vbHourglass
            
            validarCajasOP
            
            Me.MousePointer = vbDefault
        Case "ID_ImprimirA4":
            With objAyudaVale
                .TipoVale = "S"
                .CodigoAlmacen = Mid(cmbAlmacen.Text, InStr(1, cmbAlmacen.Text, "*") + 1)
                .NumeroVale = Trim(txtNumero.Text)
                
                If Not .verificarExistencia Then
                    MsgBox "Vale no registrado, verifique.", vbInformation + vbOKOnly, App.ProductName
                    
                    Exit Sub
                End If
            End With
            
            With rptValeIngreso
                .TipoVale = objAyudaVale.TipoVale
                .CodAlmacen = objAyudaVale.CodigoAlmacen
                .NumeroVale = objAyudaVale.NumeroVale
                
                'ModMilano.abrirCnDBMilano
                
                .fldCategoria.Text = ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "CT.NOMBRE", "ORDENPRODUCCION AS OP LEFT JOIN CATEGORIATIPO AS CT ON CT.IDCATEGORIATIPO = OP.IDCATEGORIATIPO", "OP.IDORDENPRODUCCION", Val(txtIDOrdenProduccion.Text), "N")
                .fldOP.Text = ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "OP", "ORDENPRODUCCION", "IDORDENPRODUCCION", Val(txtIDOrdenProduccion.Text), "N")
                
                .Show 1
            End With
        Case "ID_Lista":
            Unload Me
        Case "ID_Calculadora":
            Dim lngCalculadora As Long
            
            lngCalculadora = Shell("calc.exe", vbNormalFocus)
        Case "ID_Salir":
            Unload Me
    End Select
End Sub

Private Sub txtNroOrdenProduccion_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Me.MousePointer = vbHourglass
            
            validarCajasOP
            
            Me.MousePointer = vbDefault
    End Select
End Sub

Private Sub txtNroOrdenProduccion_LostFocus()
    If Trim(txtNroOrdenProduccion.Text) = vbNullString Then
        txtIDOrdenProduccion.Text = vbNullString
    End If
End Sub

